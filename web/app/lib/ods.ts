import type { AppState, ExpenseLine } from "./types";
import { copyPages } from "./domain";
import { unzipSync } from "fflate";
import { DOMParser, XMLSerializer, type Element as XmlElement } from "@xmldom/xmldom";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
};
const TEMPLATE_SHEET = "【原本】出張旅費精算";

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

function elementChildren(parent: XmlElement): XmlElement[] {
  return Array.from(parent.childNodes).filter((node): node is XmlElement => node.nodeType === 1);
}

function rowElements(table: XmlElement): XmlElement[] {
  return elementChildren(table).filter((element) => element.namespaceURI === NS.table && element.localName === "table-row");
}

function expandFirstColumns(row: XmlElement, limit = 8): void {
  let column = 1;
  for (const cell of elementChildren(row)) {
    if (cell.namespaceURI !== NS.table || !["table-cell", "covered-table-cell"].includes(cell.localName ?? "")) continue;
    const repeated = Math.max(1, Number(cell.getAttributeNS(NS.table, "number-columns-repeated")) || 1);
    if (repeated > 1 && column <= limit) {
      const expanded = Math.min(repeated, limit - column + 1);
      for (let index = 0; index < expanded; index += 1) {
        const clone = cell.cloneNode(true) as XmlElement;
        clone.removeAttributeNS(NS.table, "number-columns-repeated");
        row.insertBefore(clone, cell);
      }
      const remainder = repeated - expanded;
      if (remainder > 0) cell.setAttributeNS(NS.table, "table:number-columns-repeated", String(remainder));
      else row.removeChild(cell);
    }
    column += repeated;
  }
}

function cellAt(table: XmlElement, rowNumber: number, columnNumber: number): XmlElement {
  const row = rowElements(table)[rowNumber - 1];
  if (!row) throw new Error(`原本の${rowNumber}行目がありません。`);
  expandFirstColumns(row);
  let column = 1;
  for (const cell of elementChildren(row)) {
    if (cell.namespaceURI !== NS.table || !["table-cell", "covered-table-cell"].includes(cell.localName ?? "")) continue;
    const repeated = Math.max(1, Number(cell.getAttributeNS(NS.table, "number-columns-repeated")) || 1);
    if (columnNumber >= column && columnNumber < column + repeated) return cell;
    column += repeated;
  }
  throw new Error(`原本の${rowNumber}行${columnNumber}列目がありません。`);
}

function clearCellValue(cell: XmlElement): void {
  for (const attribute of ["value-type", "value", "currency", "date-value"]) cell.removeAttributeNS(NS.office, attribute);
  cell.removeAttributeNS(NS.table, "formula");
  while (cell.firstChild) cell.removeChild(cell.firstChild);
}

function setText(cell: XmlElement, value: string): void {
  clearCellValue(cell);
  cell.setAttributeNS(NS.office, "office:value-type", "string");
  const document = cell.ownerDocument;
  if (!document) throw new Error("ODS原本のXML文書を参照できません。");
  const paragraph = document.createElementNS(NS.text, "text:p");
  paragraph.appendChild(document.createTextNode(value));
  cell.appendChild(paragraph);
}

function setNumber(cell: XmlElement, value: number, currency = false): void {
  clearCellValue(cell);
  cell.setAttributeNS(NS.office, "office:value-type", currency ? "currency" : "float");
  cell.setAttributeNS(NS.office, "office:value", String(value));
  if (currency) cell.setAttributeNS(NS.office, "office:currency", "JPY");
  const document = cell.ownerDocument;
  if (!document) throw new Error("ODS原本のXML文書を参照できません。");
  const paragraph = document.createElementNS(NS.text, "text:p");
  paragraph.appendChild(document.createTextNode(currency ? `¥${value.toLocaleString("ja-JP")}` : String(value)));
  cell.appendChild(paragraph);
}

function populateSheet(table: XmlElement, state: AppState, page: ExpenseLine[], pageIndex: number, pageCount: number): void {
  table.setAttributeNS(NS.table, "table:name", `出張旅費精算_${pageIndex + 1}`);
  setText(cellAt(table, 1, 3), state.profile.department);
  setText(cellAt(table, 1, 7), `（　${pageIndex + 1}　）枚目/（　${pageCount}　）枚中`);
  setText(cellAt(table, 3, 1), `${Number(state.selectedMonth.slice(5, 7))}月度　出張旅費代精算書（電車・バス用）`);
  setText(cellAt(table, 5, 1), `氏名　${state.profile.employeeName}`);

  for (let index = 0; index < 20; index += 1) {
    const row = 11 + index;
    for (const column of [1, 2, 3, 4, 6, 7]) clearCellValue(cellAt(table, row, column));
    const line = page[index];
    if (!line) continue;
    setNumber(cellAt(table, row, 1), Number(line.date.slice(5, 7)));
    setNumber(cellAt(table, row, 2), Number(line.date.slice(8, 10)));
    setText(cellAt(table, row, 3), line.destination);
    setText(cellAt(table, row, 4), line.paidSection);
    setNumber(cellAt(table, row, 6), line.claimAmount, true);
    setText(cellAt(table, row, 7), line.reason);
  }

  const total = page.reduce((sum, line) => sum + line.claimAmount, 0);
  const totalCell = cellAt(table, 31, 6);
  setNumber(totalCell, total, true);
  totalCell.setAttributeNS(NS.table, "table:formula", "of:=SUM([.F11:.F30])");
}

export function createOdsFromTemplate(template: ArrayBuffer, state: AppState, lines: ExpenseLine[]): Blob {
  const archive = unzipSync(new Uint8Array(template));
  const content = archive["content.xml"];
  if (!content) throw new Error("ODS原本にcontent.xmlがありません。");
  const document = new DOMParser().parseFromString(decoder.decode(content), "application/xml");
  const tables = Array.from(document.getElementsByTagNameNS(NS.table, "table"));
  const source = tables.find((table) => table.getAttributeNS(NS.table, "name") === TEMPLATE_SHEET);
  if (!source || !source.parentNode) throw new Error(`ODS原本に「${TEMPLATE_SHEET}」シートがありません。`);
  if (cellAt(source, 31, 6).getAttributeNS(NS.table, "formula") !== "of:=SUM([.F11:.F30])") throw new Error("ODS原本のF31合計式が異なります。");

  const pages = copyPages(lines);
  const parent = source.parentNode;
  for (const table of tables) parent.removeChild(table);
  pages.forEach((page, pageIndex) => {
    const sheet = source.cloneNode(true) as XmlElement;
    populateSheet(sheet, state, page, pageIndex, pages.length);
    parent.appendChild(sheet);
  });
  archive["content.xml"] = encoder.encode(new XMLSerializer().serializeToString(document));
  const orderedFiles = Object.entries(archive).sort(([left], [right]) => left === "mimetype" ? -1 : right === "mimetype" ? 1 : left.localeCompare(right));
  return new Blob([zipStore(orderedFiles.map(([name, data]) => ({ name, data }))) as BlobPart], { type: "application/vnd.oasis.opendocument.spreadsheet" });
}

export function parseOdsTableRows(buffer: ArrayBuffer): unknown[][] {
  const archive = unzipSync(new Uint8Array(buffer));
  const content = archive["content.xml"];
  if (!content) throw new Error("ODSにcontent.xmlがありません。");
  const document = new DOMParser().parseFromString(decoder.decode(content), "application/xml");
  const source = Array.from(document.getElementsByTagNameNS(NS.table, "table"))
    .find((table) => table.getAttributeNS(NS.table, "name") === TEMPLATE_SHEET);
  if (!source) throw new Error(`ODSに「${TEMPLATE_SHEET}」シートがありません。`);
  const valueAt = (row: number, column: number) => {
    const cell = cellAt(source, row, column);
    return cell.getAttributeNS(NS.office, "value") ?? cell.textContent?.trim() ?? "";
  };
  return Array.from({ length: 20 }, (_, index) => {
    const row = 11 + index;
    return [valueAt(row, 1), valueAt(row, 2), valueAt(row, 3), valueAt(row, 4), valueAt(row, 6), valueAt(row, 7)];
  });
}
