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

const OCR_NOISE = /(?:ChatGPT|\bFile\b|\bEdit\b|\bView\b|\bWindow\b|\bHelp\b|共有|確認する|確認済み|画像OCR|移動あり|出張旅費申請書作成アプリ|対象月|バックアップ|復元|予定取込|経路確認|コピー出力|Excel出力)/i;

function isOcrNoise(value: string): boolean {
  return OCR_NOISE.test(value) || /^(?:業務|場所|設定|月|日|時刻|予定名)[：:]?$/.test(value);
}

function dateFromOcr(line: string, year: number): string | null {
  const match = line.match(/(?:(\d{4})\s*[年\/.-]\s*)?(\d{1,2})\s*(?:月|[\/.-])\s*(\d{1,2})\s*日?/);
  if (!match) return null;
  const y = Number(match[1] || year); const month = Number(match[2]); const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return `${y}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function timesFromOcr(line: string): string[] {
  return [...line.matchAll(/(?:^|[^0-9])(\d{1,2})\s*(?::|時)\s*(\d{2})(?:\s*分)?/g)]
    .map((match) => [Number(match[1]), Number(match[2])])
    .filter(([hour, minute]) => hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59)
    .map(([hour, minute]) => `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`);
}

function cleanOcrText(line: string): string {
  return line
    .replace(/(?:\d{4}\s*[年\/.-]\s*\d{1,2}\s*(?:月|[\/.-])\s*\d{1,2}\s*日?|\d{1,2}\s*(?:月|[\/.])\s*\d{1,2}\s*日)(?:\s*[（(][^）)]*[）)])?/g, " ")
    .replace(/\d{1,2}\s*(?::|時)\s*\d{2}(?:\s*分)?/g, " ")
    .replace(/^[\s|｜:：\-–—〜～]+|[\s|｜:：\-–—〜～]+$/g, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

/** OCR/PDF向け。日付と時刻を検出できた予定だけを候補化し、画面UI文字を予定にしない。 */
export function parseOcrSchedules(text: string, selectedMonth: string, source: "画像OCR" | "PDF"): ScheduleItem[] {
  const year = Number(selectedMonth.slice(0, 4));
  const lines = text.replace(/\r/g, "").split("\n").map((line) => line.replace(/\s+/g, " ").trim()).filter(Boolean);
  let currentDate: string | null = null;
  let pending: ScheduleItem | null = null;
  const result: ScheduleItem[] = [];
  const flush = () => {
    if (pending && pending.title.length >= 2 && !isOcrNoise(pending.title)) result.push(pending);
    pending = null;
  };
  for (const line of lines) {
    const foundDate = dateFromOcr(line, year);
    if (foundDate) currentDate = foundDate;
    const times = timesFromOcr(line);
    const cleaned = cleanOcrText(line);
    if (times.length && currentDate) {
      flush();
      pending = {
        id: uid(), date: currentDate, startTime: times[0], endTime: times[1] ?? times[0],
        title: isOcrNoise(cleaned) ? "" : cleaned, location: "", isBusiness: true,
        hasTravel: true, confirmed: false, source,
      };
      continue;
    }
    if (!pending || !cleaned || isOcrNoise(cleaned) || cleaned.length < 2) continue;
    const location = cleaned.match(/^(?:場所|会場|訪問先)\s*[:：]\s*(.+)$/)?.[1];
    if (location) pending.location = location;
    else if (!pending.title) pending.title = cleaned;
    else if (!pending.location && /(?:校|教室|本部|支社|会議室|センター|駅|市|区|町|ビル)$/.test(cleaned)) pending.location = cleaned;
  }
  flush();
  return result.filter((item) => item.date.startsWith(selectedMonth));
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
