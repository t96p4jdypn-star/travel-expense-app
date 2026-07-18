import assert from "node:assert/strict";
import test from "node:test";
import { buildDayRoute, copyPages, findFareRule, isPassCovered, mergeClaimMasters, outputLines, parseClaimRows, parseIcsSchedules, parseOcrSchedules, recalculateExpenseLine, stationsFromSection, suggestExpenseFromDestination, tabSeparated } from "../app/lib/domain";
import { EMPTY_STATE, normalizeState, type AppState, type ExpenseLine, type ScheduleItem } from "../app/lib/types";

const expense = (id: string, patch: Partial<ExpenseLine> = {}): ExpenseLine => ({
  id, date: "2026-07-15", startTime: "09:00", destination: "浦和高校", origin: "池袋", arrival: "浦和",
  paidSection: "池袋→浦和", icFare: 406, claimAmount: 406, reason: "学校訪問", state: "確認済み",
  routeOrder: 0, duplicateWarning: false, passCovered: false, hiddenZero: false, createdAt: "2026-07-01T00:00:00Z", ...patch,
});

const state = (): AppState => structuredClone({ ...EMPTY_STATE, selectedMonth: "2026-07" });

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

test("同じ過去実績は利用回数を集約し、別運賃は別候補にする", () => {
  const rows = parseClaimRows([[7, 15, "浦和高校", "池袋→浦和", 406, "学校訪問"], [7, 20, "浦和高校", "池袋→浦和", 406, "学校訪問"], [8, 1, "浦和高校", "池袋→浦和", 420, "学校訪問"]], 2026);
  const masters = mergeClaimMasters([], rows, "過去申請.xlsx");
  assert.equal(masters.length, 2); assert.equal(masters.find((item) => item.icFare === 406)?.useCount, 2); assert.equal(masters[0].sourceName, "過去申請.xlsx");
  assert.deepEqual(stationsFromSection("池袋 → 浦和"), { origin: "池袋", arrival: "浦和" });
});
