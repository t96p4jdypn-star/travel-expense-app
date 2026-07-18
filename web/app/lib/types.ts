export type ReviewState = "未確認" | "確認済み" | "修正済み" | "保留" | "除外" | "申請済み";

export type ScheduleItem = {
  id: string; date: string; startTime: string; endTime: string; title: string;
  location: string; isBusiness: boolean; hasTravel: boolean; confirmed: boolean;
  source: "手入力" | "テキスト" | "CSV" | "画像OCR" | "PDF";
};

export type ExpenseLine = {
  id: string; date: string; startTime: string; destination: string; origin: string;
  arrival: string; paidSection: string; icFare: number; claimAmount: number;
  reason: string; state: ReviewState; routeOrder: number; duplicateWarning: boolean;
  passCovered: boolean; hiddenZero: boolean; createdAt: string; sourceScheduleId?: string;
  fareSource?: "自動計算" | "履歴・要確認" | "手入力";
  fareCheckedAt?: string; routeDetails?: string; fareRevisionStatus?: string;
};

export type Place = { id: string; name: string; nearestStation: string; route: string; reason: string; visitCount: number; lastUsedAt: string };
export type WorkBase = { id: string; name: string; station: string };
export type CommuterPass = { id: string; startStation: string; endStation: string; viaStations: string; lines: string; validFrom: string; validTo: string };
export type DayRule = { weekday: number; startPlace: string; returnPlace: string };
export type HistoryItem = { id: string; destination: string; origin: string; arrival: string; paidSection: string; reason: string; usedAt: string; count: number; icFare?: number; fareCheckedAt?: string; routeDetails?: string };
export type ScheduleCapture = { id: string; month: string; name: string; dataUrl: string; createdAt: string };

export type AppState = {
  version: 1; selectedMonth: string;
  profile: { department: string; employeeName: string; homeName: string; homeStation: string };
  workBases: WorkBase[]; dayRules: DayRule[]; commuterPasses: CommuterPass[];
  places: Place[]; schedules: ScheduleItem[]; expenses: ExpenseLine[]; history: HistoryItem[]; captures: ScheduleCapture[];
  lastSavedAt: string;
};

export const EMPTY_STATE: AppState = {
  version: 1,
  selectedMonth: new Date().toISOString().slice(0, 7),
  profile: { department: "", employeeName: "", homeName: "自宅", homeStation: "" },
  workBases: [],
  dayRules: Array.from({ length: 7 }, (_, weekday) => ({ weekday, startPlace: "自宅", returnPlace: "自宅" })),
  commuterPasses: [], places: [], schedules: [], expenses: [], history: [], captures: [], lastSavedAt: new Date().toISOString(),
};

export function normalizeState(value: AppState): AppState {
  return {
    ...structuredClone(EMPTY_STATE), ...value,
    profile: { ...EMPTY_STATE.profile, ...(value.profile ?? {}) },
    workBases: value.workBases ?? [], dayRules: value.dayRules?.length ? value.dayRules : EMPTY_STATE.dayRules,
    commuterPasses: value.commuterPasses ?? [], places: value.places ?? [], schedules: value.schedules ?? [],
    expenses: value.expenses ?? [], history: value.history ?? [], captures: value.captures ?? [],
  };
}
