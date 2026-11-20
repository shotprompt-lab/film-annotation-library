function element(tag, className, content) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (content !== undefined) node.textContent = content;
  return node;
}
function button(label, action, className) {
  const node = element("button", className, label);
  node.type = "button";
  node.addEventListener("click", action);
  return node;
}
function termIds(value) {
  return Array.isArray(value) ? [...new Set(value.filter(id => byId.has(id)))] : [];
}
function poolName(value) {
  return typeof value === "string" ? value.replace(/[\u0000-\u001f\u007f]/g," ").trim().slice(0,64) : "";
}
function readPoolState() {
  try {
    const saved = JSON.parse(localStorage.getItem(poolKey) || "null");
    if (Array.isArray(saved?.projects)) {
      const ids = new Set();
      const projects = saved.projects.flatMap(item => {
        if (typeof item?.id !== "string" || !/^pool-[\w-]{1,70}$/.test(item.id) || ids.has(item.id)) return [];
        const name = poolName(item.name);
        if (!name) return [];
        ids.add(item.id);
        return [{id:item.id,name,terms:termIds(item.terms)}];
      });
      if (projects.length) return {active:ids.has(saved.active) ? saved.active : projects[0].id,projects};
    }
  } catch {}
  return {active:"pool-default",projects:[{id:"pool-default",name:"默认关键词池",terms:[]}]};
}
function currentPool() {
  return poolState.projects.find(project => project.id === poolState.active);
}
function savePoolState() {
  try {
    localStorage.setItem(poolKey, JSON.stringify(poolState));
  } catch {
    notice("浏览器未能保存选择，本次仍可导出。");
  }
  syncPool();
}
function setInPool(id, checked) {
  if (!byId.has(id)) return;
  const project = currentPool();
  if (checked) {
    if (!project.terms.includes(id)) project.terms.push(id);
  } else project.terms = project.terms.filter(term => term !== id);
  savePoolState();
}
function selectPool(id) {
  if (!poolState.projects.some(project => project.id === id)) return;
  poolState.active = id;
  closePoolEditor();
  savePoolState();
}
function showPoolEditor(mode) {
  poolEditMode = mode;
  $("poolNameLabel").textContent = mode === "create" ? "新项目池名称" : "修改项目池名称";
  $("savePoolName").textContent = mode === "create" ? "新建并切换" : "保存名称";
  $("poolName").value = mode === "create" ? "" : currentPool().name;
  $("poolNameError").hidden = true;
  $("poolEditor").hidden = false;
  $("poolName").focus();
  $("poolName").select();
}
function closePoolEditor() {
  poolEditMode = "";
  $("poolEditor").hidden = true;
  $("poolNameError").hidden = true;
}
function submitPoolName(event) {
  event.preventDefault();
  if (!poolEditMode) return;
  const name = poolName($("poolName").value);
  const current = currentPool();
  const taken = poolState.projects.some(project =>
    (poolEditMode === "create" || project.id !== current.id) &&
    normalize(project.name) === normalize(name));
  if (!name || taken) {
    $("poolNameError").textContent = taken ? "这个名称已经在用。" : "请输入项目池名称。";
    $("poolNameError").hidden = false;
    $("poolName").focus();
    return;
  }
  if (poolEditMode === "create") {
    const id = "pool-" + Array.from(crypto.getRandomValues(new Uint32Array(3)),n => n.toString(36)).join("-");
    poolState.projects.push({id,name,terms:[]});
    poolState.active = id;
  } else current.name = name;
  closePoolEditor();
  savePoolState();
  $("poolSelect").focus();
}
function poolCheckbox(entry) {
  const label = element("label", "pool-check");
  const checkbox = element("input");
  checkbox.type = "checkbox";
  checkbox.checked = currentPool().terms.includes(entry.id);
  checkbox.dataset.pool = entry.id;
  checkbox.setAttribute("aria-label", "加入当前项目池：" + entry.cn);
  checkbox.addEventListener("change", () => setInPool(entry.id, checkbox.checked));
  label.append(checkbox, element("span", "", "加入池"));
  return label;
}
function syncPool() {
  const project = currentPool();
  const current = new Set(project.terms);
  const text = project.name + " · " + current.size;
  $("openPool").textContent = text;
  $("openPool").title = "关键词池：" + text;
  for (const field of document.querySelectorAll("[data-pool]")) {
    field.checked = current.has(field.dataset.pool);
    field.title = "当前项目池：" + project.name;
  }
  if ($("poolDialog").open) renderPool();
}
function renderPool() {
  const active = currentPool();
  const select = $("poolSelect");
  select.replaceChildren(...poolState.projects.map(project => {
    const option = element("option","",project.name + "（" + project.terms.length + "）");
    option.value = project.id;
    return option;
  }));
  select.value = active.id;
  const groups = groupEntries(selectedEntries());
  $("poolCount").textContent = groups.length
    ? active.terms.length + " 条词目 · " + groups.length + " 个分类 · 仅导出当前项目池"
    : "当前池为空。在词目卡片上勾选「加入池」即可收集。";
  $("exportPoolMarkdown").disabled = !groups.length;
  $("exportPoolExcel").disabled = !groups.length;
  const content = $("poolList");
  content.replaceChildren(...groups.map(group => {
    const section = element("section","pool-group");
    const rows = element("ul","pool-list");
    for (const entry of group.entries) {
      const row = element("li","pool-row");
      const text = element("div");
      text.append(element("strong","",entry.cn),element("p","en",entry.en));
      const remove = button("移除", () => {
        const buttons = [...content.querySelectorAll("[data-remove]")];
        const index = buttons.indexOf(remove);
        const next = (buttons[index+1] || buttons[index-1])?.dataset.remove;
        setInPool(entry.id,false);
        (next ? content.querySelector('[data-remove="' + next + '"]') : select)?.focus();
      });
      remove.dataset.remove = entry.id;
      remove.setAttribute("aria-label","从当前项目池移除 " + entry.cn);
      row.append(text,remove);
      rows.append(row);
    }
    section.append(element("h3","",group.name + " · " + group.entries.length + " 条"),rows);
    return section;
  }));
}
function selectedEntries() {
  return currentPool().terms.map(id => byId.get(id));
}
function downloadPoolFile(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = element("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url),30000);
}
function exportPool(format) {
  const entries = selectedEntries();
  if (!entries.length) return;
  const projectName = currentPool().name;
  const fileName = projectName.replace(/[<>:"/\\|?*]/g,"_").replace(/[. ]+$/g,"") || "项目";
  const now = new Date();
  const today = [now.getFullYear(),String(now.getMonth()+1).padStart(2,"0"),String(now.getDate()).padStart(2,"0")].join("-");
  const blob = format === "md"
    ? new Blob([markdownPool(entries,today,projectName)],{type:"text/markdown;charset=utf-8"})
    : workbookPool(entries);
  downloadPoolFile(blob,"影视术语池_" + fileName + "_" + today + "." + format);
}

