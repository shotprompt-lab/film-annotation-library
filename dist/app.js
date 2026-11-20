"use strict";
const $ = id => document.getElementById(id);
const normalize = value => String(value).normalize("NFKC").toLowerCase().trim();
const views = {terms:"专业术语", samples:"对照样本", guide:"标注方法与参考"};
const pageSize = 18;
const poolKey = "film-annotation-library:project-pools:v1";
let library, byId, poolState, noticeTimer;
let view = "terms", main = "", sub = "", visible = pageSize, poolEditMode = "";
const compared = new Set();
const dialog = $("dialog");

function groupEntries(entries) {
  const groups = new Map();
  for (const e of entries) {
    const name = e.stage + " / " + e.sub;
    if (!groups.has(name)) groups.set(name,{name,entries:[]});
    groups.get(name).entries.push(e);
  }
  return [...groups.values()];
}
function markdownPool(entries,today,name) {
  const title = name.replace(/[\\`*_[\]<>#]/g, char => "\\" + char);
  return ["# "+title,today, ...groupEntries(entries).map(group => "## "+group.name+"\n\n"+group.entries.map(e =>
    ["### "+e.cn+" · "+e.en, e.meaning, "**提示词**　"+e.prompt].join("\n\n")).join("\n\n"))].join("\n\n");
}
function workbookPool(entries) {
  const headers = ["环节","子类","中文术语／现场说法","英文／常用缩写","专业含义","可直接使用的提示词"];
  const rows = groupEntries(entries).flatMap(group => group.entries.map(e => [e.stage,e.sub,e.cn,e.en,e.meaning,e.prompt]));
  const sheet = xlsxSheet(headers,rows,[18,22,27,35,58,65],1);
  const types = xmlHeader+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';
  return zipStored([["[Content_Types].xml",types],["_rels/.rels",xlsxRelationships([{id:"rId1",type:"officeDocument",target:"xl/workbook.xml"}])],
    ["xl/workbook.xml",xmlHeader+'<workbook xmlns="'+spreadsheetNs+'" xmlns:r="'+officeRelNs+'"><sheets><sheet name="专业术语" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ["xl/_rels/workbook.xml.rels",xlsxRelationships([{id:"rId1",type:"worksheet",target:"worksheets/sheet1.xml"},{id:"rId2",type:"styles",target:"styles.xml"}])],
    ["xl/styles.xml",xlsxStyle()],["xl/worksheets/sheet1.xml",sheet.xml]]);
}
function notice(text) {
  clearTimeout(noticeTimer);
  $("toast").textContent=text; $("toast").hidden=false;
  noticeTimer=setTimeout(()=>{$("toast").hidden=true},2300);
}
async function copy(text) {
  try { await navigator.clipboard.writeText(text); notice("提示词已复制"); }
  catch { notice("请选中提示词并复制"); }
}
function setPressed(node,active){node.setAttribute("aria-pressed",String(active));return node;}
function block(title,text,style="") {
  const b=element("section","block "+style); b.append(element("h4","",title),element("p","",text)); return b;
}
function termContents(e) {
  const div=element("div","detail-stack");
  div.append(element("p","en",e.en),block("专业含义",e.meaning),block("可直接使用的提示词",e.prompt,"prompt-text"));
  const actions=element("div","buttons");actions.append(button("复制提示词",()=>copy(e.prompt)),poolCheckbox(e));div.append(actions);
  return div;
}
function show(title,body,wide=false) {
  $("dialogTitle").textContent=title; $("dialogBody").replaceChildren(body);
  dialog.classList.toggle("compare",wide); if (!dialog.open) dialog.showModal();
}
function sampleDetail(e) {
  const div=element("div","detail-stack"),pair=element("div","ab-grid");
  pair.append(block("A 版指令",e.a),block("B 版指令",e.b));
  div.append(block("共同场景与约束",e.scene),pair,block("比较哪一处",e.comparison),block("判定边界",e.boundary));
  show(e.cn,div);
}
function safeReference(value) {
  try {const u=new URL(value);return u.protocol==="https:"&&!u.username&&!u.password?u:null;} catch{return null;}
}
function toggleComparison(e) {
  if(compared.has(e.id))compared.delete(e.id);
  else if(compared.size<3)compared.add(e.id);
  else {notice("一次最多并排对比 3 条术语");return;}
  syncComparison();
}
function syncComparison() {
  for(const b of document.querySelectorAll("[data-compare]")) {
    const selected=compared.has(b.dataset.compare);
    setPressed(b,selected);b.textContent=selected?"已选对比":"选作对比";b.closest("article").classList.toggle("selected",selected);
  }
  $("tray").hidden=!compared.size;$("trayCount").textContent="待对比 "+compared.size+" / 3";$("openCompare").disabled=compared.size<2;
  $("trayItems").replaceChildren(...[...compared].map(id=>button(byId.get(id).cn+" ×",()=>toggleComparison(byId.get(id)))));
}
function compare() {
  const grid=element("div","compare-grid"+(compared.size===2?" two-col":""));
  for(const id of compared){const e=byId.get(id),c=element("article","compare-card");c.append(element("h3","",e.cn),termContents(e));grid.append(c);}
  show("术语并排对比",grid,true);
}
function card(e) {
  const c=element("article","card "+view);
  if(view==="terms") {
    const meta=element("div","meta",e.stage+" / "+e.sub);
    const h=element("div");h.append(element("h3","",e.cn),element("p","en",e.en));
    const actions=element("div","buttons");const toggle=button("选作对比",()=>toggleComparison(e),"compare-toggle");
    toggle.dataset.compare=e.id;
    actions.append(button("查看提示词",()=>show(e.cn,termContents(e)),"main-action"),toggle,poolCheckbox(e));
    c.append(meta,h,element("p","meaning",e.meaning),actions);
  } else if(view==="samples") {
    const actions=element("div","buttons");actions.append(button("对照 A / B 指令",()=>sampleDetail(e)));
    c.append(element("h3","",e.cn),block("共同场景",e.scene),actions);
  } else {
    c.append(element("span","meta",e.kind),element("h3","",e.cn),element("p","meaning",e.description));
    const url=safeReference(e.example);
    if(url){
      const field=element("section","block");
      const link=element("a","reference-link",e.example);
      link.href=url.href;
      link.target="_blank";
      link.rel="noopener noreferrer";
      field.append(element("h4","","样例／资料地址"),link);
      c.append(field);
    }
    else c.append(block("样例",e.example,"prompt-text"));
  } return c;
}
function startFilter(){visible=pageSize;render();}
function selectCategory(parent,child="") {
  main=parent;sub=child;startFilter();
  const selected=$("categories").querySelector('button[aria-pressed="true"]');selected?.focus({preventScroll:true});
}
function categoryButton(title,total,parent,child="") {
  const b=setPressed(button(title,()=>selectCategory(parent,child)),parent===main&&child===sub);
  b.append(element("span","n",String(total)));return b;
}
function categoryValue(e){return view==="terms"?e.stage:e.kind;}
function nav() {
  const box=$("categories");
  box.hidden=view==="samples";
  if(box.hidden){box.replaceChildren();return;}
  const items=library[view],list=element("div","catlist");
  list.append(categoryButton("全部"+(view==="terms"?"术语":"条目"),items.length,""));
  for(const name of new Set(items.map(categoryValue))){
    const rows=items.filter(e=>categoryValue(e)===name);
    list.append(categoryButton(name,rows.length,name));
    if(view==="terms"&&main===name){
      const children=element("div","subnav catlist");children.setAttribute("aria-label",name+"的子类");
      for(const kind of new Set(rows.map(e=>e.sub)))children.append(categoryButton(kind,rows.filter(e=>e.sub===kind).length,name,kind));
      list.append(children);
    }
  }
  box.replaceChildren(element("h2","",view==="terms"?"环节与子类":"用途"),list);
}
function render() {
  nav();const query=normalize($("search").value).split(/\s+/).filter(Boolean);
  const evidence=$("evidence").value;
  const filtered=library[view].filter(e=>{
    if(main&&categoryValue(e)!==main || sub&&e.sub!==sub || evidence&&e.evidence!==evidence)return false;
    const fields=view==="terms"?[e.id,e.stage,e.sub,e.cn,e.en,e.meaning,e.prompt]:Object.values(e);
    const text=normalize(fields.join(" "));
    return query.every(token=>text.includes(token));
  });
  $("cards").replaceChildren(...filtered.slice(0,visible).map(card));
  if(!filtered.length)$("cards").append(element("p","empty","没有匹配条目，请调整搜索词或清空筛选。"));
  $("count").textContent=[main||views[view],sub].filter(Boolean).join(" / ")+"　·　"+filtered.length+" 条";
  $("more").hidden=visible>=filtered.length;$("more").textContent="显示更多（剩余 "+Math.max(0,filtered.length-visible)+" 条）";
  $("reset").hidden=!main&&!sub&&!query.length&&!evidence;syncComparison();
}
function switchView(next, category="") {
  view=next;main=category;sub="";$("search").value="";$("search").placeholder=next==="terms"?"搜索术语、英文、含义或提示词":next==="samples"?"搜索术语、场景或对照指令":"搜索标注方法、项目或参考资料";
  const select=$("evidence");
  select.replaceChildren(new Option("全部核验材料",""));
  select.hidden=next!=="terms";
  if(next==="terms"){
    for(const material of new Set(library.terms.map(e=>e.evidence))){
      if(material!=="单帧")select.add(new Option(material,material));
    }
  }
  for(const b of $("sections").children)setPressed(b,b.dataset.view===view);
  startFilter();
}
async function init() {
  const response=await fetch("./library.json");if(!response.ok)throw Error("data");
  library=await response.json();byId=new Map(library.terms.map(e=>[e.id,e]));poolState=readPoolState();
  for(const [key,name] of Object.entries(views)){const b=button(name,()=>switchView(key));b.dataset.view=key;b.append(element("span","total",library[key].length));$("sections").append(b);}
  for(const e of library.samples.slice(0,3))$("quickPairs").append(button(e.cn,()=>sampleDetail(e)));
  $("search").oninput=startFilter;$("evidence").onchange=startFilter;$("reset").onclick=()=>switchView(view);
  $("more").onclick=()=>{visible+=pageSize;render();};$("clearCompare").onclick=()=>{compared.clear();syncComparison();};
  $("openCompare").onclick=compare;$("closeDialog").onclick=()=>dialog.close();
  $("openPool").onclick=()=>{closePoolEditor();renderPool();$("poolDialog").showModal();};
  $("closePool").onclick=()=>$("poolDialog").close();$("poolSelect").onchange=e=>selectPool(e.target.value);
  $("createPool").onclick=()=>showPoolEditor("create");$("renamePool").onclick=()=>showPoolEditor("rename");
  $("poolEditor").onsubmit=submitPoolName;$("cancelPoolName").onclick=closePoolEditor;
  $("exportPoolExcel").onclick=()=>exportPool("xlsx");$("exportPoolMarkdown").onclick=()=>exportPool("md");
  const requestedView=new URLSearchParams(location.search).get("view");
  if(requestedView==="references")switchView("guide","参考");
  else switchView(requestedView==="samples"?"samples":"terms");
  syncPool();
}
init().catch(error=>{console.error(error);$("cards").textContent="词表加载失败，请刷新本地预览页面。";});
