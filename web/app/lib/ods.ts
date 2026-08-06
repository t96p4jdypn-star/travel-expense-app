import type { AppState, ExpenseLine } from "./types";
import { copyPages } from "./domain";
import { unzipSync, zipSync, type Zippable } from "fflate";
import { DOMParser, XMLSerializer, type Element as XmlElement, type Node as XmlNode } from "@xmldom/xmldom";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  config: "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
};
const TEMPLATE_SHEET = "【原本】出張旅費精算";
const ZIP_MTIME = new Date(2026, 0, 1, 0, 0, 0);

function elementChildren(parent: XmlElement): XmlElement[] {
  return Array.from(parent.childNodes).filter((node): node is XmlElement => node.nodeType === 1);
}

function visibleCellText(cell: XmlElement): string {
  const collect = (node: XmlNode): string => {
    if (node.nodeType === 3) return node.nodeValue ?? "";
    if (node.nodeType !== 1) return "";
    const element = node as XmlElement;
    if (element.namespaceURI === NS.text && element.localName === "ruby-text") return "";
    return Array.from(element.childNodes).map(collect).join("");
  };
  return collect(cell).trim();
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

function synchronizeSettings(settings: Uint8Array, sheetNames: string[]): Uint8Array<ArrayBuffer> {
  const document = new DOMParser().parseFromString(decoder.decode(settings), "application/xml");
  const configItems = Array.from(document.getElementsByTagNameNS(NS.config, "config-item"));
  configItems
    .filter((item) => item.getAttributeNS(NS.config, "name") === "ActiveTable")
    .forEach((item) => { item.textContent = sheetNames[0]; });

  const namedMaps = Array.from(document.getElementsByTagNameNS(NS.config, "config-item-map-named"));
  namedMaps
    .filter((map) => ["Tables", "ScriptConfiguration"].includes(map.getAttributeNS(NS.config, "name") ?? ""))
    .forEach((map) => {
      const entries = elementChildren(map).filter((element) => element.namespaceURI === NS.config && element.localName === "config-item-map-entry");
      const template = entries.find((entry) => entry.getAttributeNS(NS.config, "name") === TEMPLATE_SHEET) ?? entries[0];
      if (!template) return;
      entries.forEach((entry) => map.removeChild(entry));
      sheetNames.forEach((sheetName, index) => {
        const entry = template.cloneNode(true) as XmlElement;
        entry.setAttributeNS(NS.config, "config:name", sheetName);
        Array.from(entry.getElementsByTagNameNS(NS.config, "config-item"))
          .filter((item) => item.getAttributeNS(NS.config, "name") === "CodeName")
          .forEach((item) => { item.textContent = `Sheet${index + 1}`; });
        map.appendChild(entry);
      });
    });
  const encoded = encoder.encode(new XMLSerializer().serializeToString(document));
  const result = new Uint8Array(encoded.length);
  result.set(encoded);
  return result;
}

function packageOds(archive: Record<string, Uint8Array>): Uint8Array {
  const mimetype = archive.mimetype;
  if (!mimetype) throw new Error("ODS原本にmimetypeがありません。");
  const files: Zippable = { mimetype: [mimetype, { level: 0, mtime: ZIP_MTIME }] };
  Object.entries(archive).forEach(([name, data]) => {
    if (name !== "mimetype") files[name] = [data, { level: name.endsWith("/") ? 0 : 6, mtime: ZIP_MTIME }];
  });
  return zipSync(files);
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
  if (archive["settings.xml"]) {
    archive["settings.xml"] = synchronizeSettings(archive["settings.xml"], pages.map((_, index) => `出張旅費精算_${index + 1}`));
  }
  return new Blob([packageOds(archive) as BlobPart], { type: "application/vnd.oasis.opendocument.spreadsheet" });
}

export function parseOdsTableRows(buffer: ArrayBuffer): unknown[][] {
  const archive = unzipSync(new Uint8Array(buffer));
  const content = archive["content.xml"];
  if (!content) throw new Error("ODSにcontent.xmlがありません。");
  const document = new DOMParser().parseFromString(decoder.decode(content), "application/xml");
  const sources = Array.from(document.getElementsByTagNameNS(NS.table, "table"))
    .filter((table) => {
      const name = table.getAttributeNS(NS.table, "name") ?? "";
      return name === TEMPLATE_SHEET || name.startsWith(`${TEMPLATE_SHEET}_`);
    });
  if (!sources.length) throw new Error(`ODSに「${TEMPLATE_SHEET}」シートがありません。`);
  const valueAt = (source: XmlElement, row: number, column: number) => {
    const cell = cellAt(source, row, column);
    return cell.getAttributeNS(NS.office, "value") ?? visibleCellText(cell);
  };
  return sources.flatMap((source) => Array.from({ length: 20 }, (_, index) => {
      const row = 11 + index;
      return [valueAt(source, row, 1), valueAt(source, row, 2), valueAt(source, row, 3), valueAt(source, row, 4), valueAt(source, row, 6), valueAt(source, row, 7)];
    }));
}
