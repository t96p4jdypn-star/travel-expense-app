import assert from "node:assert/strict";
import test from "node:test";
import { buildDayRoute, claimDestinationCandidates, claimRouteCandidates, copyPages, findFareRule, isPassCovered, mergeClaimMasters, outputLines, parseClaimRows, parseIcsSchedules, parseOcrSchedules, prepareClaimRowsForRegistration, prioritizeClaimRouteCandidates, recalculateExpenseLine, stationsFromSection, suggestExpenseFromDestination, tabSeparated } from "../app/lib/domain";
import { createInitialState, EMPTY_STATE, normalizeNumericText, normalizeState, safeAmount, type AppState, type ExpenseLine, type ScheduleItem } from "../app/lib/types";
import { readFile } from "node:fs/promises";
import { createOdsFromTemplate, parseOdsTableRows } from "../app/lib/ods";
import { unzipSync, zipSync } from "fflate";

const expense = (id: string, patch: Partial<ExpenseLine> = {}): ExpenseLine => ({
  id, date: "2026-07-15", startTime: "09:00", destination: "浦和高校", origin: "池袋", arrival: "浦和",
  paidSection: "池袋→浦和", icFare: 406, claimAmount: 406, reason: "学校訪問", state: "確認済み",
  routeOrder: 0, duplicateWarning: false, passCovered: false, hiddenZero: false, createdAt: "2026-07-01T00:00:00Z", ...patch,
});

const state = (): AppState => structuredClone({ ...EMPTY_STATE, selectedMonth: "2026-07" });

test("新規環境は入力明細を作らず、初期実績6件だけを候補用に登録する", () => {
  const value = createInitialState("2026-07-26");
  assert.equal(value.expenses.length, 0);
  assert.deepEqual(value.claimMasters.map(({ destination, paidSection, icFare, reason }) => ({ destination, paidSection, icFare, reason })), [
    { destination: "本社", paidSection: "武蔵浦和→北与野", icFare: 199, reason: "本社業務" },
    { destination: "南越谷", paidSection: "北与野→武蔵浦和", icFare: 199, reason: "教室管理" },
    { destination: "越谷レイクタウン", paidSection: "北与野→武蔵浦和", icFare: 199, reason: "教室管理" },
    { destination: "川越教室", paidSection: "ふじみ野→川越", icFare: 178, reason: "巡回" },
    { destination: "本社", paidSection: "川越→北与野", icFare: 341, reason: "本社業務" },
    { destination: "自宅", paidSection: "北与野→武蔵浦和", icFare: 199, reason: "帰宅" },
  ]);
  assert.ok(value.claimMasters.every((master) => master.useCount === 1 && master.lastUsedDate === "2026-07-26" && master.sourceName === "初期実績"));
});

test("目的地候補は重複を除き、選択後の区間候補は目的地で絞り込む", () => {
  const masters = createInitialState("2026-07-26").claimMasters;
  assert.deepEqual(claimDestinationCandidates(masters), ["本社", "南越谷", "越谷レイクタウン", "川越教室", "自宅"]);
  assert.deepEqual(claimDestinationCandidates(masters, "越谷"), ["南越谷", "越谷レイクタウン"]);
  assert.deepEqual(claimDestinationCandidates(masters, "越谷　レイクタウン"), ["越谷レイクタウン"]);
  assert.equal(claimRouteCandidates(masters, " 越谷　レイクタウン ").length, 1);
  assert.deepEqual(claimRouteCandidates(masters, "本社").map((master) => [master.paidSection, master.icFare]), [
    ["武蔵浦和→北与野", 199],
    ["川越→北与野", 341],
  ]);
});

test("金額または理由が不足する区間候補は完全な候補より後方に表示する", () => {
  const incomplete = { id: "incomplete", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和→北与野", icFare: 0, reason: "", useCount: 8, lastUsedDate: "2026-07-29", sourceName: "既存" };
  const complete = { id: "complete", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和→北与野", icFare: 199, reason: "本社業務", useCount: 1, lastUsedDate: "2026-07-27", sourceName: "初期実績" };
  assert.deepEqual(prioritizeClaimRouteCandidates([incomplete, complete]).map((master) => master.id), ["complete", "incomplete"]);
  assert.deepEqual(claimRouteCandidates([incomplete, complete], "本社").map((master) => master.id), ["incomplete", "complete"]);
});

test("旧保存環境は既存実績を保護しながら不足する初期実績だけを一回追加する", () => {
  const existing = state() as unknown as Omit<AppState, "version"> & { version: number };
  existing.version = 2;
  existing.claimMasters = [
    { id: "kept", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和 → 北与野", icFare: 199, reason: "本社業務", useCount: 8, lastUsedDate: "2026-07-20", sourceName: "既存データ" },
    { id: "custom", destination: "利用者追加先", origin: "A", arrival: "B", paidSection: "A→B", icFare: 500, reason: "既存理由", useCount: 4, lastUsedDate: "2026-07-21", sourceName: "画面入力" },
  ];
  const migrated = normalizeState(existing as AppState);
  assert.equal(migrated.version, 3);
  assert.equal(migrated.claimMasters.length, 7);
  assert.deepEqual(migrated.claimMasters.find((master) => master.id === "kept"), existing.claimMasters[0]);
  assert.deepEqual(migrated.claimMasters.find((master) => master.id === "custom"), existing.claimMasters[1]);
  assert.equal(migrated.claimMasters.filter((master) => master.destination === "越谷レイクタウン").length, 1);

  const afterUserDelete = structuredClone(migrated);
  afterUserDelete.claimMasters = afterUserDelete.claimMasters.filter((master) => master.destination !== "越谷レイクタウン");
  assert.equal(normalizeState(afterUserDelete).claimMasters.some((master) => master.destination === "越谷レイクタウン"), false);

  const existingEmpty = state() as unknown as Omit<AppState, "version"> & { version: number };
  existingEmpty.version = 2;
  existingEmpty.claimMasters = [];
  assert.equal(normalizeState(existingEmpty as AppState).claimMasters.length, 6);
});

test("旧保存行は入力値を維持してVer3の行状態を補完する", () => {
  const legacy = state() as unknown as Omit<AppState, "version"> & { version: number };
  legacy.version = 1;
  legacy.expenses = [expense("legacy", {
    date: "2026-07-01", destination: "", paidSection: "旧自動区間", icFare: 0, claimAmount: 0,
    reason: "", state: undefined as never, createdAt: "",
  })];
  const normalized = normalizeState(legacy as AppState);
  assert.equal(normalized.version, 3);
  assert.equal(normalized.expenses[0].date, "2026-07-");
  assert.equal(normalized.expenses[0].paidSection, "");
  assert.equal(normalized.expenses[0].state, "未確認");
  assert.ok(normalized.expenses[0].createdAt);
});

test("出力対象は確認済み・1円以上だけになる", () => {
  const value = state();
  value.expenses = [expense("ok"), expense("zero", { claimAmount: 0 }), expense("hold", { state: "保留" }), expense("dupe", { duplicateWarning: true })];
  assert.deepEqual(outputLines(value).map((line) => line.id), ["ok"]);
});

test("45行は20・20・5行に分かれ、コピーは6列タブ区切り", () => {
  const lines = Array.from({ length: 45 }, (_, index) => expense(String(index)));
  assert.deepEqual(copyPages(lines).map((page) => page.length), [20, 20, 5]);
  assert.equal(tabSeparated(lines.slice(0, 1)), "7\t15\t浦和高校\t池袋→浦和\t406\t学校訪問");
});

test("定期券の経由駅を含む区間は内部0円経路にできる", () => {
  const value = state();
  value.commuterPasses = [{ id: "p", startStation: "ふじみ野", endStation: "池袋", viaStations: "川越", lines: "東上線", validFrom: "2026-07-01", validTo: "2026-07-31" }];
  assert.equal(isPassCovered("川越", "池袋", "2026-07-15", value), true);
  assert.equal(isPassCovered("川越", "浦和", "2026-07-15", value), false);
});

test("1日の予定を時刻順につないで最終戻り先まで経路化", () => {
  const value = state(); value.profile.homeStation = "ふじみ野";
  value.places = [
    { id: "a", name: "浦和高校", nearestStation: "浦和", route: "", reason: "学校訪問", visitCount: 0, lastUsedAt: "" },
    { id: "b", name: "大宮高校", nearestStation: "大宮", route: "", reason: "学校訪問", visitCount: 0, lastUsedAt: "" },
  ];
  const schedules: ScheduleItem[] = [
    { id: "b", date: "2026-07-15", startTime: "13:00", endTime: "14:00", title: "訪問", location: "大宮高校", isBusiness: true, hasTravel: true, confirmed: true, source: "手入力" },
    { id: "a", date: "2026-07-15", startTime: "10:00", endTime: "11:00", title: "訪問", location: "浦和高校", isBusiness: true, hasTravel: true, confirmed: true, source: "手入力" },
  ];
  assert.deepEqual(buildDayRoute(schedules, value).map((line) => line.paidSection), ["ふじみ野→浦和", "浦和→大宮", "大宮→ふじみ野"]);
});

test("iPhoneカレンダーICSから予定候補を作る", () => {
  const items = parseIcsSchedules("BEGIN:VCALENDAR\nBEGIN:VEVENT\nDTSTART:20260715T100000\nDTEND:20260715T110000\nSUMMARY:学校訪問\nLOCATION:浦和高校\nEND:VEVENT\nEND:VCALENDAR");
  assert.equal(items[0].date, "2026-07-15"); assert.equal(items[0].startTime, "10:00"); assert.equal(items[0].location, "浦和高校");
});

test("OCRは日付・時刻・予定名を組み合わせ、複数予定を正しく分ける", () => {
  const items = parseOcrSchedules(`2026年7月15日（水）
10:00 - 11:00
学校訪問
場所：浦和高校
13:00 - 14:00 会議
場所：大宮本部`, "2026-07", "画像OCR");
  assert.deepEqual(items.map(({ date, startTime, endTime, title, location }) => ({ date, startTime, endTime, title, location })), [
    { date: "2026-07-15", startTime: "10:00", endTime: "11:00", title: "学校訪問", location: "浦和高校" },
    { date: "2026-07-15", startTime: "13:00", endTime: "14:00", title: "会議", location: "大宮本部" },
  ]);
});

test("OCRはブラウザやアプリの操作文字を予定として取り込まない", () => {
  const items = parseOcrSchedules(`7月16日（木） 9:12
ChatGPT File Edit View Window Help
出張旅費申請書作成アプリ
09:00 - 10:00
確認する
画像OCR
業務
移動あり`, "2026-07", "画像OCR");
  assert.deepEqual(items, []);
});

test("既知の行き先は最寄駅・理由・履歴運賃を自動補完する", () => {
  const value = state(); value.profile.homeStation = "ふじみ野";
  value.places = [{ id: "p", name: "浦和高校", nearestStation: "浦和", route: "池袋経由", reason: "学校訪問", visitCount: 2, lastUsedAt: "" }];
  value.history = [{ id: "h", destination: "浦和高校", origin: "ふじみ野", arrival: "浦和", paidSection: "ふじみ野→浦和", reason: "学校訪問", usedAt: "2026-07-01T00:00:00Z", count: 3, icFare: 721, fareCheckedAt: "2026-07-01T00:00:00Z" }];
  const suggestion = suggestExpenseFromDestination(value, { date: "2026-07-15", startTime: "10:00", destination: "浦和高校" });
  assert.equal(suggestion.origin, "ふじみ野"); assert.equal(suggestion.arrival, "浦和"); assert.equal(suggestion.reason, "学校訪問"); assert.equal(suggestion.icFare, 721); assert.equal(suggestion.fareSource, "履歴・要確認");
});

test("同日の2件目は直前行の到着駅から自動的につなぐ", () => {
  const value = state(); value.profile.homeStation = "ふじみ野"; value.expenses = [expense("first", { arrival: "浦和", startTime: "10:00" })];
  value.places = [{ id: "p", name: "大宮高校", nearestStation: "大宮", route: "", reason: "学校訪問", visitCount: 0, lastUsedAt: "" }];
  const suggestion = suggestExpenseFromDestination(value, { date: "2026-07-15", startTime: "13:00", destination: "大宮高校" });
  assert.equal(suggestion.origin, "浦和"); assert.equal(suggestion.arrival, "大宮"); assert.equal(suggestion.paidSection, "浦和→大宮");
});

test("確定済みの区間はブラウザ内の運賃台帳から自動計算する", () => {
  const value = state(); value.fareRules = [{ id: "f", origin: "浦和駅", arrival: "大宮", paidSection: "浦和→大宮", icFare: 178, routeDetails: "JR", registeredAt: "2026-07-01T00:00:00Z", lastUsedAt: "2026-07-01T00:00:00Z", useCount: 3 }];
  const calculated = recalculateExpenseLine(expense("x", { origin: "浦和", arrival: "大宮", paidSection: "", icFare: 0, claimAmount: 0, state: "未確認" }), value);
  assert.equal(calculated.icFare, 178); assert.equal(calculated.claimAmount, 178); assert.equal(calculated.fareSource, "登録運賃");
});

test("登録運賃は逆方向でも使えるが表示区間は移動方向になる", () => {
  const value = state(); value.fareRules = [{ id: "f", origin: "浦和", arrival: "大宮", paidSection: "浦和→大宮", icFare: 178, routeDetails: "JR", registeredAt: "2026-07-01T00:00:00Z", lastUsedAt: "2026-07-01T00:00:00Z", useCount: 1 }];
  assert.equal(findFareRule(value, "大宮", "浦和")?.reversed, true);
  const calculated = recalculateExpenseLine(expense("x", { origin: "大宮", arrival: "浦和", icFare: 0 }), value);
  assert.equal(calculated.paidSection, "大宮→浦和"); assert.equal(calculated.icFare, 178);
});

test("旧バックアップの確定履歴を運賃台帳へ移行する", () => {
  const legacy = state(); legacy.history = [{ id: "h", destination: "大宮高校", origin: "浦和", arrival: "大宮", paidSection: "浦和→大宮", reason: "学校訪問", usedAt: "2026-07-01T00:00:00Z", count: 2, icFare: 178 }];
  delete (legacy as Partial<AppState>).fareRules;
  const migrated = normalizeState(legacy);
  assert.equal(migrated.fareRules.length, 1); assert.equal(migrated.fareRules[0].icFare, 178); assert.equal(migrated.fareRules[0].useCount, 2);
});

test("過去申請書の6列を読み取り、不完全行を除外する", () => {
  const rows = parseClaimRows([[7, 15, "浦和高校", "池袋→浦和", "406円", "学校訪問"], ["月", "日", "目的地", "区間", "料金", "理由"], [7, 16, "", "池袋→大宮", 483, "訪問"]], 2026);
  assert.deepEqual(rows, [{ date: "2026-07-15", destination: "浦和高校", paidSection: "池袋→浦和", icFare: 406, reason: "学校訪問" }]);
});

test("過去申請で省略された月日は直前明細から引き継ぐ", () => {
  const rows = parseClaimRows([
    [4, 22, "大和田", "武蔵浦和→大和田", 366, "入試総括"],
    ["", "", "南越谷", "大和田→武蔵浦和", 366, "教室管理"],
    ["", 23, "本社", "武蔵浦和→北与野", 199, "本社業務"],
  ], 2026);
  assert.deepEqual(rows.map((row) => row.date), ["2026-04-22", "2026-04-22", "2026-04-23"]);
});

test("不正な金額は0円へ安全補正する", () => {
  assert.equal(safeAmount(Number.NaN), 0);
  assert.equal(safeAmount(Number.POSITIVE_INFINITY), 0);
  assert.equal(safeAmount("金額なし"), 0);
  assert.equal(safeAmount("199"), 199);
  assert.equal(safeAmount("１９９"), 199);
  assert.equal(safeAmount("１，２３４"), 1234);
  assert.equal(normalizeNumericText("２５"), "25");
  const value = state();
  value.claimMasters = [{ id: "m", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和→北与野", icFare: Number.NaN, reason: "本社業務", useCount: 1, lastUsedDate: "2026-07-01", sourceName: "旧データ" }];
  assert.equal(normalizeState(value).claimMasters[0].icFare, 0);
});

test("空または不正な対象月は現在月へ安全補正する", () => {
  const value = state();
  value.selectedMonth = "";
  assert.match(normalizeState(value).selectedMonth, /^\d{4}-(0[1-9]|1[0-2])$/);
});

test("同じ過去実績は利用回数を集約し、別運賃は別候補にする", () => {
  const rows = parseClaimRows([[7, 15, "浦和高校", "池袋→浦和", 406, "学校訪問"], [7, 20, "浦和高校", "池袋→浦和", 406, "学校訪問"], [8, 1, "浦和高校", "池袋→浦和", 420, "学校訪問"]], 2026);
  const masters = mergeClaimMasters([], rows, "過去申請.xlsx");
  assert.equal(masters.length, 2); assert.equal(masters.find((item) => item.icFare === 406)?.useCount, 2); assert.equal(masters[0].sourceName, "過去申請.xlsx");
  assert.deepEqual(stationsFromSection("池袋 → 浦和"), { origin: "池袋", arrival: "浦和" });
});

test("取込確認中は既存マスターを変更せず、対象行だけを安全な数値で登録準備する", () => {
  const existing = [{ id: "existing", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和→北与野", icFare: 199, reason: "本社業務", useCount: 1, lastUsedDate: "2026-07-01", sourceName: "既存" }];
  const before = structuredClone(existing);
  const targets = prepareClaimRowsForRegistration([
    { id: "edit", excluded: false, date: " 2026-07-20 ", destination: " 本社 ", paidSection: " 武蔵浦和→北与野 ", icFare: Number.NaN, reason: " 本社業務 " },
    { id: "exclude", excluded: true, date: "2026-07-21", destination: "削除対象", paidSection: "A→B", icFare: Number.POSITIVE_INFINITY, reason: "対象外" },
  ]);
  assert.deepEqual(existing, before);
  assert.deepEqual(targets, [{ date: "2026-07-20", destination: "本社", paidSection: "武蔵浦和→北与野", icFare: 0, reason: "本社業務" }]);
  assert.equal(Number.isFinite(targets[0].icFare), true);
});

test("編集後の完全一致は利用回数と最終利用日を更新し、新規内容は追加する", () => {
  const existing = [{ id: "existing", destination: "本社", origin: "武蔵浦和", arrival: "北与野", paidSection: "武蔵浦和→北与野", icFare: 199, reason: "本社業務", useCount: 1, lastUsedDate: "2026-07-01", sourceName: "既存" }];
  const merged = mergeClaimMasters(existing, [
    { date: "2026-07-20", destination: "本社", paidSection: "武蔵浦和→北与野", icFare: 199, reason: "本社業務" },
    { date: "2026-07-22", destination: "編集後訪問先", paidSection: "A→B", icFare: 250, reason: "編集後理由" },
  ], "取込確認.ods");
  assert.equal(merged.find((item) => item.id === "existing")?.useCount, 2);
  assert.equal(merged.find((item) => item.id === "existing")?.lastUsedDate, "2026-07-20");
  assert.equal(merged.find((item) => item.destination === "編集後訪問先")?.icFare, 250);
});

test("ODSは20行単位のシートと合計計算式を持つ", async () => {
  const value = state();
  value.profile.department = "営業部";
  value.profile.employeeName = "山田 太郎";
  const lines = Array.from({ length: 21 }, (_, index) => expense(String(index), { date: `2026-07-${String((index % 28) + 1).padStart(2, "0")}` }));
  const template = await readFile(new URL("../public/2026年度版出張旅費代精算書原本.ods", import.meta.url));
  const blob = createOdsFromTemplate(template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength), value, lines);
  const binary = new Uint8Array(await blob.arrayBuffer());
  const storedText = new TextDecoder().decode(unzipSync(binary)["content.xml"]);
  assert.equal(blob.type, "application/vnd.oasis.opendocument.spreadsheet");
  assert.equal(String.fromCharCode(...binary.slice(0, 2)), "PK");
  assert.match(storedText, /出張旅費精算_1/);
  assert.match(storedText, /出張旅費精算_2/);
  assert.doesNotMatch(storedText, /【見本】出張旅費精算/);
  assert.match(storedText, /table:formula="of:=SUM\(\[\.F11:\.F30\]\)"/);
  assert.match(storedText, /営業部/);
  assert.match(storedText, /山田 太郎/);
});

test("過去ODSは原本シートの11～30行だけを解析する", async () => {
  const template = await readFile(new URL("../public/2026年度版出張旅費代精算書原本.ods", import.meta.url));
  const rows = parseOdsTableRows(template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength));
  assert.equal(rows.length, 20);
  assert.ok(rows.every((row) => row.length === 6));
  assert.ok(rows.every((row) => !row.some((value) => String(value).includes("記入例"))));
});

test("原本形式ODSから目的地・区間・金額・理由を取り込む", async () => {
  const value = state();
  value.profile.department = "営業部";
  value.profile.employeeName = "山田 太郎";
  const template = await readFile(new URL("../public/2026年度版出張旅費代精算書原本.ods", import.meta.url));
  const generated = createOdsFromTemplate(
    template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength),
    value,
    [expense("import", { destination: "川越教室", paidSection: "ふじみ野→川越", claimAmount: 178, icFare: 178, reason: "巡回" })],
  );
  const archive = unzipSync(new Uint8Array(await generated.arrayBuffer()));
  archive["content.xml"] = new TextEncoder().encode(
    new TextDecoder().decode(archive["content.xml"]).replace("出張旅費精算_1", "【原本】出張旅費精算"),
  );
  const rows = parseOdsTableRows(zipSync(archive).buffer);
  const parsed = parseClaimRows(rows, 2026);
  assert.deepEqual(parsed, [{ date: "2026-07-15", destination: "川越教室", paidSection: "ふじみ野→川越", icFare: 178, reason: "巡回" }]);
});

test("過去ODSは追加原本ページを読み、ふりがな注釈を除外する", async () => {
  const value = state();
  const template = await readFile(new URL("../public/2026年度版出張旅費代精算書原本.ods", import.meta.url));
  const lines = Array.from({ length: 21 }, (_, index) => expense(`multi-${index}`, {
    destination: index === 0 ? "浦和高校" : `目的地${index + 1}`,
    date: `2026-07-${String((index % 28) + 1).padStart(2, "0")}`,
  }));
  const generated = createOdsFromTemplate(template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength), value, lines);
  const archive = unzipSync(new Uint8Array(await generated.arrayBuffer()));
  const content = new TextDecoder().decode(archive["content.xml"])
    .replace("出張旅費精算_1", "【原本】出張旅費精算")
    .replace("出張旅費精算_2", "【原本】出張旅費精算_(2)")
    .replace("<text:p>浦和高校</text:p>", "<text:p><text:ruby><text:ruby-base>浦和高校</text:ruby-base><text:ruby-text>ウラワコウコウ</text:ruby-text></text:ruby></text:p>");
  archive["content.xml"] = new TextEncoder().encode(content);
  const rows = parseOdsTableRows(zipSync(archive).buffer);
  const parsed = parseClaimRows(rows, 2026);
  assert.equal(rows.length, 40);
  assert.equal(parsed.length, 21);
  assert.equal(parsed[0].destination, "浦和高校");
});
