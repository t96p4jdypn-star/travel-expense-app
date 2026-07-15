import type { AppState, ExpenseLine } from "./types";
import { copyPages } from "./domain";

const TEMPLATE_SHEET = "【原本】出張旅費精算";

export async function createExcel(template: ArrayBuffer, state: AppState, lines: ExpenseLine[], submissionDate: string): Promise<Blob> {
  const ExcelJS = await import("exceljs");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(template);
  const base = workbook.getWorksheet(TEMPLATE_SHEET);
  if (!base) throw new Error(`原本シート「${TEMPLATE_SHEET}」がありません。`);
  if (base.getCell("F31").type !== ExcelJS.ValueType.Formula) throw new Error("F31の合計数式がありません。");
  for (let row = 11; row <= 30; row += 1) {
    if (!base.getCell(`D${row}`).isMerged || !base.getCell(`G${row}`).isMerged) throw new Error("原本の結合セル構造が異なります。");
  }
  const pages = copyPages(lines);
  const prepared = pages.map((_, index) => {
    if (index === 0) { base.name = "出張旅費精算_1"; return base; }
    const sheet = workbook.addWorksheet(`出張旅費精算_${index + 1}`);
    const cloned = structuredClone(base.model);
    cloned.name = `出張旅費精算_${index + 1}`;
    cloned.id = sheet.id;
    sheet.model = cloned;
    return sheet;
  });
  workbook.worksheets.slice().forEach((sheet) => {
    if (!prepared.includes(sheet) && sheet.name !== "【見本】出張旅費精算") workbook.removeWorksheet(sheet.id);
  });
  prepared.forEach((sheet, pageIndex) => {
    const page = pages[pageIndex];
    sheet.getCell("C1").value = state.profile.department;
    sheet.getCell("G1").value = `（ ${pageIndex + 1} ）枚目/（ ${pages.length} ）枚中`;
    sheet.getCell("A3").value = `${Number(state.selectedMonth.slice(5))}月度 出張旅費代精算書（電車・バス用）`;
    sheet.getCell("A5").value = `氏名　${state.profile.employeeName}`;
    const validPasses = state.commuterPasses.map((pass) => `${pass.lines ? `${pass.lines} ` : ""}${pass.startStation} ～ ${pass.endStation}`);
    sheet.getCell("C7").value = validPasses.join("\n") || "定期券なし";
    sheet.getCell("E44").value = submissionDate ? `${Number(submissionDate.slice(0, 4))}年 ${Number(submissionDate.slice(5, 7))}月 ${Number(submissionDate.slice(8, 10))}日` : "";
    for (let row = 11; row <= 30; row += 1) ["A", "B", "C", "D", "F", "G"].forEach((col) => { sheet.getCell(`${col}${row}`).value = null; });
    page.forEach((line, index) => {
      const row = 11 + index; const d = new Date(`${line.date}T00:00:00`);
      sheet.getCell(`A${row}`).value = d.getMonth() + 1;
      sheet.getCell(`B${row}`).value = d.getDate();
      sheet.getCell(`C${row}`).value = line.destination;
      sheet.getCell(`D${row}`).value = line.paidSection;
      sheet.getCell(`F${row}`).value = line.claimAmount;
      sheet.getCell(`G${row}`).value = line.reason;
    });
    sheet.getCell("F31").value = { formula: "SUM(F11:F30)" };
    sheet.pageSetup.printArea = "A1:H44";
    sheet.pageSetup.fitToPage = true; sheet.pageSetup.fitToWidth = 1; sheet.pageSetup.fitToHeight = 1;
  });
  return new Blob([await workbook.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
