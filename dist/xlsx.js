// Build a small OOXML workbook locally so both downloads work without a server API.
const spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const officeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function xmlEscape(value) {
  return String(value ?? "").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g,"")
    .replace(/_x[0-9a-f]{4}_/gi,match => "_x005F_" + match.slice(1))
    .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
function xlsxRelationships(items) {
  return xmlHeader + '<Relationships xmlns="' + packageRelNs + '">' + items.map(item =>
    '<Relationship Id="' + item.id + '" Type="' + officeRelNs + '/' + item.type + '" Target="' + xmlEscape(item.target) + '"' +
    (item.type === "hyperlink" ? ' TargetMode="External"' : "") + '/>').join("") + '</Relationships>';
}
function xlsxStyle() {
  return xmlHeader + '<styleSheet xmlns="' + spreadsheetNs + '">' +
    '<fonts count="3"><font><sz val="11"/><name val="Microsoft YaHei"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>' +
    '<font><u/><sz val="11"/><color rgb="FF225A9B"/><name val="Microsoft YaHei"/></font></fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF27354B"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF6E5"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left/><right/><top style="medium"><color rgb="FFBD8B3D"/></top><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
}
function xlsxSheet(headers, rows, widths, categoryIndex, urlIndex = -1) {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const links = [];
  const table = [headers,...rows].map((row,index) => {
    let lineCount = 1;
    const firstInCategory = index > 0 && (index === 1 || row[categoryIndex] !== rows[index-2][categoryIndex]);
    const cells = row.map((value,col) => {
      const ref = letters[col] + (index+1);
      let style = index === 0 ? 1 : firstInCategory ? 4 : 2;
      if (index && col === urlIndex) {
        links.push({id:"rId" + (links.length+1),type:"hyperlink",target:value,ref});
        style = firstInCategory ? 5 : 3;
      }
      const wrapped = String(value).split("\n").reduce((sum,line) => {
        const units = [...line].reduce((len,symbol) => len + (symbol.charCodeAt(0) > 255 ? 2 : 1),0);
        return sum + Math.max(1,Math.ceil(units / Math.max(1,widths[col]-3)));
      },0);
      lineCount = Math.max(lineCount,wrapped);
      return '<c r="' + ref + '" t="inlineStr" s="' + style + '"><is><t xml:space="preserve">' + xmlEscape(value) + '</t></is></c>';
    });
    return '<row r="' + (index+1) + '" ht="' + (index ? Math.min(390,Math.max(22,lineCount*17+5)) : 26) + '" customHeight="1">' + cells.join("") + '</row>';
  }).join("");
  const bounds = "A1:" + letters[headers.length-1] + (rows.length+1);
  const columns = widths.map((width,i) => '<col min="' + (i+1) + '" max="' + (i+1) + '" width="' + width + '" customWidth="1"/>').join("");
  return {
    xml:xmlHeader + '<worksheet xmlns="' + spreadsheetNs + '" xmlns:r="' + officeRelNs + '">' +
      '<dimension ref="' + bounds + '"/><sheetViews><sheetView showGridLines="0" workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>' + columns + '</cols><sheetData>' + table +
      '</sheetData><autoFilter ref="' + bounds + '"/>' + (links.length ? '<hyperlinks>' +
      links.map(link => '<hyperlink ref="' + link.ref + '" r:id="' + link.id + '"/>').join("") + '</hyperlinks>' : "") +
      '</worksheet>',
    links
  };
}
function zipStored(files) {
  const encoder = new TextEncoder();
  const table = Uint32Array.from({length:256},(_,i) => {
    for (let bit=0;bit<8;bit++) i = (i >>> 1) ^ ((i & 1) ? 0xedb88320 : 0);
    return i;
  });
  const dataParts = [];
  const indexParts = [];
  let offset = 0;
  let indexSize = 0;
  for (const [name,content] of files) {
    const path = encoder.encode(name);
    const bytes = encoder.encode(content);
    let crc = 0xffffffff;
    for (const byte of bytes) crc = (crc >>> 8) ^ table[(crc ^ byte) & 255];
    crc = (crc ^ 0xffffffff) >>> 0;
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0,0x04034b50,true);
    local.setUint16(4,20,true);
    local.setUint16(12,0x21,true);
    local.setUint32(14,crc,true);
    local.setUint32(18,bytes.length,true);
    local.setUint32(22,bytes.length,true);
    local.setUint16(26,path.length,true);
    dataParts.push(local.buffer,path,bytes);
    const record = new DataView(new ArrayBuffer(46));
    record.setUint32(0,0x02014b50,true);
    record.setUint16(4,20,true);
    record.setUint16(6,20,true);
    record.setUint16(14,0x21,true);
    record.setUint32(16,crc,true);
    record.setUint32(20,bytes.length,true);
    record.setUint32(24,bytes.length,true);
    record.setUint16(28,path.length,true);
    record.setUint32(42,offset,true);
    indexParts.push(record.buffer,path);
    offset += 30 + path.length + bytes.length;
    indexSize += 46 + path.length;
  }
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0,0x06054b50,true);
  end.setUint16(8,files.length,true);
  end.setUint16(10,files.length,true);
  end.setUint32(12,indexSize,true);
  end.setUint32(16,offset,true);
  return new Blob([...dataParts,...indexParts,end.buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
