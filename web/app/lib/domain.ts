import type { AppState, ExpenseLine, ScheduleItem } from "./types";

export const uid = () => crypto.randomUUID();
export const monthOf = (date: string) => date.slice(0, 7);
export const yen = (value: number) => `${value.toLocaleString("ja-JP")}円`;

export function outputLines(state: AppState): ExpenseLine[] {
  return state.expenses
    .filter((line) => monthOf(line.date) === state.selectedMonth)
    .filter((line) => ["確認済み", "修正済み"].includes(line.state))
    .filter((line) => line.claimAmount > 0 && !line.duplicateWarning)
    .sort((a, b) => a.date.localeCompare(b.date) || a.startTime.localeCompare(b.startTime) || a.routeOrder - b.routeOrder);
}

export function duplicateKeys(lines: ExpenseLine[]): Set<string> {
  const counts = new Map<string, number>();
  lines.forEach((line) => {
    if (line.state === "申請済み" || line.claimAmount <= 0) return;
    const key = `${line.date}|${line.paidSection}|${line.claimAmount}`;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  });
  return new Set([...counts].filter(([, count]) => count > 1).map(([key]) => key));
}

export function isPassCovered(origin: string, arrival: string, date: string, state: AppState): boolean {
  return state.commuterPasses.some((pass) => {
    if (pass.validFrom && date < pass.validFrom) return false;
    if (pass.validTo && date > pass.validTo) return false;
    const stations = [pass.startStation, ...pass.viaStations.split(/[、,\s]+/), pass.endStation].filter(Boolean);
    return stations.includes(origin) && stations.includes(arrival);
  });
}

function placeStation(name: string, state: AppState): string {
  if (name === state.profile.homeName || name === "自宅") return state.profile.homeStation;
  return state.workBases.find((base) => base.name === name)?.station
    ?? state.places.find((place) => place.name === name)?.nearestStation ?? name;
}

export function buildDayRoute(items: ScheduleItem[], state: AppState, returnOverride = ""): ExpenseLine[] {
  if (!items.length) return [];
  const sorted = [...items].filter((item) => item.isBusiness && item.hasTravel).sort((a, b) => a.startTime.localeCompare(b.startTime));
  if (!sorted.length) return [];
  const day = new Date(`${sorted[0].date}T00:00:00`).getDay();
  const rule = state.dayRules.find((r) => r.weekday === day);
  const startName = rule?.startPlace || state.profile.homeName || "自宅";
  const returnName = returnOverride || rule?.returnPlace || state.profile.homeName || "自宅";
  const legs: ExpenseLine[] = [];
  let originName = startName;
  sorted.forEach((item, index) => {
    const origin = placeStation(originName, state);
    const arrival = placeStation(item.location || item.title, state);
    const covered = isPassCovered(origin, arrival, item.date, state);
    legs.push({
      id: uid(), date: item.date, startTime: item.startTime, destination: item.location || item.title,
      origin, arrival, paidSection: `${origin}→${arrival}`, icFare: 0, claimAmount: 0,
      reason: state.places.find((p) => p.name === item.location)?.reason || "",
      state: "未確認", routeOrder: index, duplicateWarning: false, passCovered: covered,
      hiddenZero: true, createdAt: new Date().toISOString(), sourceScheduleId: item.id,
    });
    originName = item.location || item.title;
  });
  const finalOrigin = placeStation(originName, state);
  const finalArrival = placeStation(returnName, state);
  if (finalOrigin && finalArrival && finalOrigin !== finalArrival) {
    const covered = isPassCovered(finalOrigin, finalArrival, sorted[0].date, state);
    legs.push({
      id: uid(), date: sorted[0].date, startTime: sorted.at(-1)?.endTime || "23:59", destination: returnName,
      origin: finalOrigin, arrival: finalArrival, paidSection: `${finalOrigin}→${finalArrival}`,
      icFare: 0, claimAmount: 0, reason: "帰着", state: "未確認", routeOrder: legs.length,
      duplicateWarning: false, passCovered: covered, hiddenZero: true, createdAt: new Date().toISOString(),
    });
  }
  return legs;
}

export function parseTextSchedules(text: string, month: string, source: ScheduleItem["source"]): ScheduleItem[] {
  const year = Number(month.slice(0, 4));
  const rows = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  return rows.map((row, index) => {
    const cols = row.split(/\t|,/).map((value) => value.trim());
    const dateMatch = row.match(/(?:(\d{4})[\/.-])?(\d{1,2})[\/.-](\d{1,2})/);
    const timeMatch = row.match(/(\d{1,2}):(\d{2})(?:\s*[-～〜]\s*(\d{1,2}):(\d{2}))?/);
    const date = dateMatch
      ? `${dateMatch[1] || year}-${String(dateMatch[2]).padStart(2, "0")}-${String(dateMatch[3]).padStart(2, "0")}`
      : `${month}-01`;
    return {
      id: uid(), date, startTime: timeMatch ? `${timeMatch[1].padStart(2, "0")}:${timeMatch[2]}` : "09:00",
      endTime: timeMatch?.[3] ? `${timeMatch[3].padStart(2, "0")}:${timeMatch[4]}` : "10:00",
      title: cols[2] || cols[1] || row, location: cols[3] || "", isBusiness: true,
      hasTravel: true, confirmed: false, source,
    };
  });
}

export function parseIcsSchedules(text: string): ScheduleItem[] {
  const unfolded = text.replace(/\r?\n[ \t]/g, "");
  return unfolded.split("BEGIN:VEVENT").slice(1).map((block) => {
    const field = (name: string) => block.match(new RegExp(`^${name}(?:;[^:]*)?:(.*)$`, "m"))?.[1]?.trim() ?? "";
    const start = field("DTSTART"); const end = field("DTEND");
    const compact = (value: string) => value.replace(/[^0-9]/g, "");
    const s = compact(start); const e = compact(end);
    const date = s.length >= 8 ? `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}` : new Date().toISOString().slice(0, 10);
    const clock = (value: string, fallback: string) => value.length >= 12 ? `${value.slice(8, 10)}:${value.slice(10, 12)}` : fallback;
    return {
      id: uid(), date, startTime: clock(s, "09:00"), endTime: clock(e, "10:00"),
      title: field("SUMMARY").replace(/\\,/g, ","), location: field("LOCATION").replace(/\\,/g, ","),
      isBusiness: true, hasTravel: true, confirmed: false, source: "テキスト" as const,
    };
  });
}

export function copyPages(lines: ExpenseLine[]) {
  const pages: ExpenseLine[][] = [];
  for (let index = 0; index < lines.length; index += 20) pages.push(lines.slice(index, index + 20));
  return pages.length ? pages : [[]];
}

export function tabSeparated(lines: ExpenseLine[]): string {
  return lines.map((line) => {
    const date = new Date(`${line.date}T00:00:00`);
    return [date.getMonth() + 1, date.getDate(), line.destination, line.paidSection, line.claimAmount, line.reason].join("\t");
  }).join("\n");
}
