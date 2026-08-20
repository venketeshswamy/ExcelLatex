import fs from 'fs';
import path from 'path';
import os from 'os';

function crc32(buf) {
  let crc = ~0;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ ~0) >>> 0;
}

function createZip(files) {
  const entries = [];
  let offset = 0;
  for (const file of files) {
    const data = Buffer.from(file.content, 'utf-8');
    const name = Buffer.from(file.name, 'utf-8');
    const crc = crc32(data);

    const localHeader = Buffer.alloc(30 + name.length);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(name.length, 26);
    localHeader.writeUInt16LE(0, 28);
    name.copy(localHeader, 30);

    entries.push({ name, crc, size: data.length, offset, localHeader, data });
    offset += localHeader.length + data.length;
  }

  const cdEntries = [];
  let cdSize = 0;
  for (const entry of entries) {
    const cdHeader = Buffer.alloc(46 + entry.name.length);
    cdHeader.writeUInt32LE(0x02014b50, 0);
    cdHeader.writeUInt16LE(20, 4);
    cdHeader.writeUInt16LE(20, 6);
    cdHeader.writeUInt16LE(0, 8);
    cdHeader.writeUInt16LE(0, 10);
    cdHeader.writeUInt16LE(0, 12);
    cdHeader.writeUInt32LE(entry.crc, 16);
    cdHeader.writeUInt32LE(entry.size, 20);
    cdHeader.writeUInt32LE(entry.size, 24);
    cdHeader.writeUInt16LE(entry.name.length, 28);
    cdHeader.writeUInt16LE(0, 30);
    cdHeader.writeUInt16LE(0, 32);
    cdHeader.writeUInt16LE(0, 34);
    cdHeader.writeUInt16LE(0, 36);
    cdHeader.writeUInt32LE(0, 38);
    cdHeader.writeUInt32LE(entry.offset, 42);
    entry.name.copy(cdHeader, 46);

    cdEntries.push(cdHeader);
    cdSize += cdHeader.length;
  }

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  const chunks = [];
  for (const entry of entries) {
    chunks.push(entry.localHeader, entry.data);
  }
  chunks.push(...cdEntries, eocd);
  return Buffer.concat(chunks);
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
</Types>`;

const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="LaTeX Math Testing" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="FF0F6CBD"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`;

const webextension1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2015/01" id="d71c1b7e-8c38-4e12-b13c-74a621e25e1a">
  <we:reference id="d71c1b7e-8c38-4e12-b13c-74a621e25e1a" version="1.0.0.0" store="developer" storeType="developer"/>
  <we:alternateReferences/>
  <we:properties/>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>`;

const taskpanesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="1">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>`;

const taskpanesRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>`;

const rows = [
  { a: 'LaTeX Expression', b: 'Equation Name', c: 'In-Cell Formula (=MATH.KATEX)', d: 'Background Code' },
  { a: '\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', b: 'Quadratic Formula', f: '_xlfn.MATH.KATEX(A2, 0)', d: '0: Transparent' },
  { a: '\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}', b: 'Gaussian Integral', f: '_xlfn.MATH.KATEX(A3, 1)', d: '1: White' },
  { a: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', b: '2x2 Matrix', f: '_xlfn.MATH.KATEX(A4, 2)', d: '2: Black' },
  { a: '\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}', b: 'Basel Problem', f: '_xlfn.MATH.KATEX(A5, 0)', d: '0: Transparent' },
  { a: '\\nabla \\times \\mathbf{E} = -\\frac{\\partial \\mathbf{B}}{\\partial t}', b: 'Maxwell-Faraday Equation', f: '_xlfn.MATH.KATEX(A6, 1)', d: '1: White' },
  { a: 'e^{i\\pi} + 1 = 0', b: "Euler's Identity", f: '_xlfn.MATH.KATEX(A7, 0)', d: '0: Transparent' }
];

let sheetData = '<sheetData>';
rows.forEach((row, idx) => {
  const r = idx + 1;
  const isHeader = r === 1;
  const sAttr = isHeader ? ' s="1"' : '';
  sheetData += `<row r="${r}">`;
  if (row.a) sheetData += `<c r="A${r}" t="inlineStr"${sAttr}><is><t>${escapeXml(row.a)}</t></is></c>`;
  if (row.b) sheetData += `<c r="B${r}" t="inlineStr"${sAttr}><is><t>${escapeXml(row.b)}</t></is></c>`;
  if (row.f) {
    sheetData += `<c r="C${r}"><f>${escapeXml(row.f)}</f><v></v></c>`;
  } else if (row.c) {
    sheetData += `<c r="C${r}" t="inlineStr"${sAttr}><is><t>${escapeXml(row.c)}</t></is></c>`;
  }
  if (row.d) sheetData += `<c r="D${r}" t="inlineStr"${sAttr}><is><t>${escapeXml(row.d)}</t></is></c>`;
  sheetData += '</row>';
});
sheetData += '</sheetData>';

const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="40" customWidth="1"/>
    <col min="2" max="2" width="28" customWidth="1"/>
    <col min="3" max="3" width="34" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
  </cols>
  ${sheetData}
</worksheet>`;

const files = [
  { name: '[Content_Types].xml', content: contentTypes },
  { name: '_rels/.rels', content: packageRels },
  { name: 'xl/workbook.xml', content: workbookXml },
  { name: 'xl/_rels/workbook.xml.rels', content: workbookRels },
  { name: 'xl/styles.xml', content: stylesXml },
  { name: 'xl/worksheets/sheet1.xml', content: sheet1Xml },
  { name: 'xl/webextensions/webextension1.xml', content: webextension1Xml },
  { name: 'xl/webextensions/taskpanes.xml', content: taskpanesXml },
  { name: 'xl/webextensions/_rels/taskpanes.xml.rels', content: taskpanesRels }
];

const buf = createZip(files);
const localPath = path.resolve('ExcelKatexFile.xlsx');
const dlPath = path.join(os.homedir(), 'Downloads', 'ExcelKatexFile.xlsx');

fs.writeFileSync(localPath, buf);
try {
  fs.writeFileSync(dlPath, buf);
} catch { /* ignore */ }

console.log(`[ExcelKatexFile Created]: ${localPath}`);
