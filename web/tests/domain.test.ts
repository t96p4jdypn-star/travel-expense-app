import assert from "node:assert/strict";
import test from "node:test";
import { buildDayRoute, copyPages, isPassCovered, outputLines, parseIcsSchedules, tabSeparated } from "../app/lib/domain";
import { EMPTY_STATE, type AppState, type ExpenseLine, type ScheduleItem } from "../app/lib/types";

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
