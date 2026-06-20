import fs from "node:fs/promises";
import JSZip from "/Users/tei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/.pnpm/jszip@3.10.1/node_modules/jszip/lib/index.js";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function indexToColumnLetters(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    letters = String.fromCharCode(65 + mod) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function buildSheetXml(headers, rows) {
  const allRows = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  const rowXml = allRows.map((row, rowIndex) => {
    const cells = row.map((value, colIndex) => {
      const ref = `${indexToColumnLetters(colIndex)}${rowIndex + 1}`;
      if (typeof value === "number" && Number.isFinite(value)) {
        return `<c r="${ref}"><v>${value}</v></c>`;
      }
      return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
    }).join("");
    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");
  const lastCell = `${indexToColumnLetters(headers.length - 1)}${allRows.length}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${lastCell}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <sheetData>${rowXml}</sheetData>
  <autoFilter ref="A1:${lastCell}"/>
</worksheet>`;
}

async function exportResultWorkbook(headers, rows, out) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="分组结果" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Arial"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`);
  zip.file("xl/worksheets/sheet1.xml", buildSheetXml(headers, rows));
  const now = new Date().toISOString();
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>分组结果</dc:title><dc:creator>均衡分组小程序</dc:creator><cp:lastModifiedBy>均衡分组小程序</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>均衡分组小程序</Application>
</Properties>`);
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(out, buffer);
}

const headers = ["分组编号", "组内序号", "姓名", "所属二级子公司", "性别", "年龄", "部门", "岗位"];
const rows = Array.from({ length: 23 }, (_, index) => ({
  "分组编号": `第${(index % 5) + 1}组`,
  "组内序号": Math.floor(index / 5) + 1,
  "姓名": `员工${String(index + 1).padStart(2, "0")}`,
  "所属二级子公司": ["华东公司", "华南公司", "西北公司"][index % 3],
  "性别": ["男", "女"][index % 2],
  "年龄": 24 + ((index * 7) % 23),
  "部门": ["财务部", "市场部", "运营部", "技术部"][index % 4],
  "岗位": ["经理", "主管", "专员"][index % 3]
}));
const out = "/Users/tei/Documents/Codex/2026-06-20/bang-2/work/verify-export.xlsx";
await exportResultWorkbook(headers, rows, out);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
const inspect = await workbook.inspect({
  kind: "table",
  range: "分组结果!A1:H6",
  include: "values",
  tableMaxRows: 6,
  tableMaxCols: 8
});
console.log(inspect.ndjson);
