import type { AppState, ExpenseLine } from "./types";
import { copyPages } from "./domain";
import { unzipSync } from "fflate";

const encoder = new TextEncoder();

function xml(value: unknown): string {
  return String(value ?? "").replace(/[<>&'"]/g, (character) => ({
    "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;",
  })[character] ?? character);
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const uint16 = (value: number) => new Uint8Array([value & 255, (value >>> 8) & 255]);
const uint32 = (value: number) => new Uint8Array([value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255]);

function join(parts: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) { result.set(part, offset); offset += part.length; }
  return result;
}

function zipStore(files: Array<{ name: string; data: Uint8Array }>): Uint8Array {
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name);
    const checksum = crc32(file.data);
    const local = join([
      uint32(0x04034b50), uint16(20), uint16(0), uint16(0), uint16(0), uint16(0),
      uint32(checksum), uint32(file.data.length), uint32(file.data.length), uint16(name.length), uint16(0), name, file.data,
    ]);
    localParts.push(local);
    centralParts.push(join([
      uint32(0x02014b50), uint16(20), uint16(20), uint16(0), uint16(0), uint16(0), uint16(0),
      uint32(checksum), uint32(file.data.length), uint32(file.data.length), uint16(name.length), uint16(0), uint16(0),
      uint16(0), uint16(0), uint32(0), uint32(offset), name,
    ]));
    offset += local.length;
  }
  const locals = join(localParts);
  const central = join(centralParts);
  return join([
    locals, central, uint32(0x06054b50), uint16(0), uint16(0), uint16(files.length), uint16(files.length),
    uint32(central.length), uint32(locals.length), uint16(0),
  ]);
}

function textCell(value: unknown, style = "cell"): string {
  return `<table:table-cell table:style-name="${style}" office:value-type="string"><text:p>${xml(value)}</text:p></table:table-cell>`;
}

function numberCell(value: number, style = "cell"): string {
  return `<table:table-cell table:style-name="${style}" office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`;
}

function contentXml(state: AppState, lines: ExpenseLine[]): string {
  const pages = copyPages(lines);
  const tables = (pages.length ? pages : [[]]).map((page, pageIndex) => {
    const rows = Array.from({ length: 20 }, (_, index) => {
      const line = page[index];
      if (!line) return `<table:table-row>${Array.from({ length: 6 }, () => textCell("")).join("")}</table:table-row>`;
      return `<table:table-row>${numberCell(Number(line.date.slice(5, 7)))}${numberCell(Number(line.date.slice(8, 10)))}${textCell(line.destination)}${textCell(line.paidSection)}${numberCell(line.claimAmount, "money")}${textCell(line.reason)}</table:table-row>`;
    }).join("");
    return `<table:table table:name="出張旅費精算_${pageIndex + 1}">
      <table:table-column table:number-columns-repeated="6"/>
      <table:table-row><table:table-cell table:style-name="title" table:number-columns-spanned="6" office:value-type="string"><text:p>${xml(Number(state.selectedMonth.slice(5)))}月度 出張旅費精算書</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="5"/></table:table-row>
      <table:table-row>${textCell(`所属：${state.profile.department}`)}${textCell(`氏名：${state.profile.employeeName}`)}<table:table-cell table:number-columns-repeated="4"/></table:table-row>
      <table:table-row>${textCell(`（ ${pageIndex + 1} ）枚目／（ ${Math.max(1, pages.length)} ）枚中`)}<table:table-cell table:number-columns-repeated="5"/></table:table-row>
      <table:table-row>${["月", "日", "目的地", "区間", "金額", "理由"].map((label) => textCell(label, "header")).join("")}</table:table-row>
      ${rows}
      <table:table-row>${textCell("合計", "header")}<table:table-cell table:number-columns-repeated="3"/><table:table-cell table:style-name="money" office:value-type="float" table:formula="of:=SUM([.E5:.E24])"><text:p>${page.reduce((sum, line) => sum + line.claimAmount, 0)}</text:p></table:table-cell>${textCell("")}</table:table-row>
    </table:table>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.3">
<office:automatic-styles>
  <style:style style:name="cell" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #777777" fo:padding="0.08in" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"/></style:style>
  <style:style style:name="header" style:family="table-cell" style:parent-style-name="cell"><style:text-properties fo:font-weight="bold" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"/></style:style>
  <style:style style:name="title" style:family="table-cell"><style:text-properties fo:font-size="16pt" fo:font-weight="bold" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"/></style:style>
  <style:style style:name="money" style:family="table-cell" style:parent-style-name="cell"/>
</office:automatic-styles><office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
}

export function createOds(state: AppState, lines: ExpenseLine[]): Blob {
  const mimetype = "application/vnd.oasis.opendocument.spreadsheet";
  const files = [
    { name: "mimetype", data: encoder.encode(mimetype) },
    { name: "content.xml", data: encoder.encode(contentXml(state, lines)) },
    { name: "styles.xml", data: encoder.encode(`<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:styles/></office:document-styles>`) },
    { name: "meta.xml", data: encoder.encode(`<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:meta/></office:document-meta>`) },
    { name: "META-INF/manifest.xml", data: encoder.encode(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>`) },
  ];
  return new Blob([zipStore(files) as BlobPart], { type: mimetype });
}

export function parseOdsTableRows(buffer: ArrayBuffer): unknown[][] {
  const archive = unzipSync(new Uint8Array(buffer));
  const content = archive["content.xml"];
  if (!content) throw new Error("ODSにcontent.xmlがありません。");
  const document = new DOMParser().parseFromString(new TextDecoder().decode(content), "application/xml");
  if (document.querySelector("parsererror")) throw new Error("ODSのXMLを解析できません。");
  return [...document.getElementsByTagName("table:table-row")].map((row) => {
    const values: unknown[] = [];
    for (const cell of [...row.children]) {
      if (!["table:table-cell", "table:covered-table-cell"].includes(cell.tagName)) continue;
      const repeated = Math.min(20, Math.max(1, Number(cell.getAttribute("table:number-columns-repeated")) || 1));
      const raw = cell.getAttribute("office:value") ?? cell.textContent?.trim() ?? "";
      for (let index = 0; index < repeated; index += 1) values.push(raw);
    }
    return values;
  });
}
