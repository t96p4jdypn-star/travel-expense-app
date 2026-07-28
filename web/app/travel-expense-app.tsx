"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { loadState, saveState } from "./lib/db";
import { buildDayRoute, claimDestinationCandidates, claimRouteCandidates, copyPages, duplicateKeys, findFareRule, isPassCovered, mergeClaimMasters, outputLines, parseClaimRows, parseIcsSchedules, parseOcrSchedules, parseTextSchedules, prepareClaimRowsForRegistration, prioritizeClaimRouteCandidates, recalculateExpenseLine, stationsFromSection, suggestExpenseFromDestination, tabSeparated, uid, yen, type ClaimImportPreviewRow } from "./lib/domain";
import { createExcel } from "./lib/excel";
import { createOdsFromTemplate, parseOdsTableRows } from "./lib/ods";
import { EMPTY_STATE, normalizeNumericText, normalizeState, resolveStartupState, safeAmount, type AppState, type ClaimMaster, type CommuterPass, type ExpenseLine, type ScheduleCapture, type ScheduleItem } from "./lib/types";

type Tab = "入力" | "実績マスター" | "ODS出力" | "実績から作成" | "過去データ読込" | "月間" | "予定取込" | "経路確認" | "登録状況" | "コピー出力" | "Excel出力" | "設定";
type MainTab = "旅費入力" | "設定";
type SettingsSection = "基本情報" | "過去データ取込" | "実績マスター管理" | "データ管理";
const MAIN_TABS: MainTab[] = ["旅費入力", "設定"];
const SETTINGS_SECTIONS: SettingsSection[] = ["基本情報", "過去データ取込", "実績マスター管理", "データ管理"];
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob); const anchor = document.createElement("a");
  anchor.href = url; anchor.download = filename; anchor.click(); URL.revokeObjectURL(url);
}

function fileDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = reject; reader.readAsDataURL(file); });
}

function Field({ label, children, hint }: { label: string; children: React.ReactNode; hint?: string }) {
  return <label className="field"><span>{label}</span>{children}{hint && <small>{hint}</small>}</label>;
}

function StatusBadge({ value }: { value: string }) { return <span className={`status status-${value}`}>{value}</span>; }

export function TravelExpenseApp() {
  const [state, setState] = useState<AppState>(EMPTY_STATE);
  const [ready, setReady] = useState(false);
  const [mainTab, setMainTab] = useState<MainTab>("旅費入力");
  const [showZero, setShowZero] = useState(false);
  const [notice, setNotice] = useState("データはこの端末内だけに保存されます");
  const [submissionDate, setSubmissionDate] = useState(new Date().toISOString().slice(0, 10));
  const [requestedMonth, setRequestedMonth] = useState("");

  useEffect(() => { loadState().then((saved) => { setState(resolveStartupState(saved)); setReady(true); }); }, []);
  useEffect(() => { if (!ready) return; const timer = setTimeout(() => saveState({ ...state, lastSavedAt: new Date().toISOString() }), 250); return () => clearTimeout(timer); }, [state, ready]);

  const monthExpenses = useMemo(() => state.expenses.filter((line) => line.date.startsWith(state.selectedMonth)), [state]);
  const monthOptions = useMemo(() => {
    const [year, month] = state.selectedMonth.split("-").map(Number);
    return Array.from({ length: 25 }, (_, index) => {
      const date = new Date(year, month - 1 + index - 12, 1);
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    });
  }, [state.selectedMonth]);
  const visibleExpenses = useMemo(() => monthExpenses.filter((line) => showZero || line.claimAmount > 0 || line.state === "未確認"), [monthExpenses, showZero]);
  const output = useMemo(() => outputLines(state), [state]);
  const total = output.reduce((sum, line) => sum + line.claimAmount, 0);
  const warnings = monthExpenses.filter((line) => line.duplicateWarning).length;

  function mutate(updater: (draft: AppState) => AppState) { setState((current) => updater(structuredClone(current))); }
  function switchMonth(nextMonth: string) {
    mutate((draft) => ({ ...draft, selectedMonth: nextMonth }));
    setRequestedMonth("");
    setNotice(`${Number(nextMonth.slice(0, 4))}年${Number(nextMonth.slice(5, 7))}月へ切り替えました。月別データは保持されています。`);
  }
  function requestMonthChange(nextMonth: string) {
    if (nextMonth === state.selectedMonth) return;
    const hasEnteredData = monthExpenses.some((line) =>
      Boolean(line.destination.trim() || line.paidSection.trim() || line.reason.trim() || safeAmount(line.icFare) > 0)
    );
    if (hasEnteredData) {
      setRequestedMonth(nextMonth);
      return;
    }
    switchMonth(nextMonth);
  }
  function recomputeDuplicates(draft: AppState) {
    const duplicates = duplicateKeys(draft.expenses);
    draft.expenses.forEach((line) => { line.duplicateWarning = duplicates.has(`${line.date}|${line.paidSection}|${line.claimAmount}`); });
    return draft;
  }
  function addExpense(line?: Partial<ExpenseLine>) {
    mutate((draft) => recomputeDuplicates({ ...draft, expenses: [...draft.expenses, {
      id: uid(), date: `${draft.selectedMonth}-01`, startTime: "09:00", destination: "", origin: "", arrival: "",
      paidSection: "", icFare: 0, claimAmount: 0, reason: "", state: "未確認", routeOrder: 0,
      duplicateWarning: false, passCovered: false, hiddenZero: true, createdAt: new Date().toISOString(), ...line,
    }] }));
  }
  function updateExpense(id: string, patch: Partial<ExpenseLine>) {
    mutate((draft) => {
      const line = draft.expenses.find((item) => item.id === id); if (!line) return draft;
      Object.assign(line, patch);
      if (patch.icFare !== undefined && patch.fareSource === undefined) { line.fareSource = "手入力"; line.fareCheckedAt = new Date().toISOString(); }
      line.passCovered = isPassCovered(line.origin, line.arrival, line.date, draft);
      line.claimAmount = line.passCovered ? 0 : Math.max(0, Number(line.icFare || 0));
      line.hiddenZero = line.claimAmount === 0;
      return recomputeDuplicates(draft);
    });
  }
  function recalculateExpense(id: string) {
    let result = "";
    mutate((draft) => {
      const index = draft.expenses.findIndex((item) => item.id === id); if (index < 0) return draft;
      const original = draft.expenses[index]; const recalculated = recalculateExpenseLine(original, draft);
      recalculated.state = "未確認"; draft.expenses[index] = recalculated;
      result = recalculated.passCovered ? "定期券内のため申請額を0円にしました。" : recalculated.fareSource === "登録運賃" ? `登録済み運賃 ${yen(recalculated.icFare)} で再計算しました。` : "登録済み運賃がないため、入力中のIC料金を使いました。内容を確認してください。";
      return recomputeDuplicates(draft);
    });
    setNotice(result);
  }
  function confirmExpense(id: string) {
    let message = "";
    mutate((draft) => {
      const index = draft.expenses.findIndex((item) => item.id === id); if (index < 0) return draft;
      const current = draft.expenses[index];
      let line = current.fareSource === "手入力"
        ? { ...current, passCovered: isPassCovered(current.origin, current.arrival, current.date, draft), claimAmount: isPassCovered(current.origin, current.arrival, current.date, draft) ? 0 : Math.max(0, Number(current.icFare || 0)), hiddenZero: isPassCovered(current.origin, current.arrival, current.date, draft) || Number(current.icFare || 0) === 0 }
        : recalculateExpenseLine(current, draft);
      if (!line.destination.trim() || !line.origin.trim() || !line.arrival.trim()) { message = "目的地・出発駅・到着駅を入力してください。"; return draft; }
      if (!line.passCovered && line.icFare <= 0) { message = "IC料金を入力してから確定してください。"; return draft; }
      const now = new Date().toISOString(); line = { ...line, state: "確認済み", fareCheckedAt: now, fareSource: line.passCovered ? "登録運賃" : "登録運賃" }; draft.expenses[index] = line;
      if (!line.passCovered) {
        const found = findFareRule(draft, line.origin, line.arrival)?.rule;
        if (found) { found.origin = line.origin; found.arrival = line.arrival; found.paidSection = line.paidSection || `${line.origin}→${line.arrival}`; found.icFare = line.icFare; found.routeDetails = line.routeDetails || ""; found.lastUsedAt = now; found.useCount += 1; }
        else draft.fareRules.push({ id: uid(), origin: line.origin, arrival: line.arrival, paidSection: line.paidSection || `${line.origin}→${line.arrival}`, icFare: line.icFare, routeDetails: line.routeDetails || "", registeredAt: now, lastUsedAt: now, useCount: 1 });
      }
      const history = draft.history.find((item) => item.destination === line.destination && item.origin === line.origin && item.arrival === line.arrival);
      if (history) { history.count += 1; history.usedAt = now; history.paidSection = line.paidSection; history.reason = line.reason; history.icFare = line.icFare; history.fareCheckedAt = now; history.routeDetails = line.routeDetails; }
      else draft.history.push({ id: uid(), destination: line.destination, origin: line.origin, arrival: line.arrival, paidSection: line.paidSection, reason: line.reason, usedAt: now, count: 1, icFare: line.icFare, fareCheckedAt: now, routeDetails: line.routeDetails });
      let place = draft.places.find((item) => item.name === line.destination);
      if (!place) { place = { id: uid(), name: line.destination, nearestStation: line.arrival, route: line.routeDetails || line.paidSection, reason: line.reason, visitCount: 0, lastUsedAt: "" }; draft.places.push(place); }
      place.visitCount += 1; place.lastUsedAt = now; place.nearestStation ||= line.arrival; place.reason ||= line.reason; place.route ||= line.routeDetails || line.paidSection;
      message = line.passCovered ? "定期券内0円の経路として確定しました。" : `この区間のIC料金 ${yen(line.icFare)} を端末内に登録しました。次回から自動計算します。`;
      return recomputeDuplicates(draft);
    });
    setNotice(message);
  }
  function removeExpense(id: string) { mutate((draft) => ({ ...draft, expenses: draft.expenses.filter((line) => line.id !== id) })); }

  async function quickAdd(input: { date: string; startTime: string; destination: string; nearestStation?: string; reason?: string; icFare?: number }) {
    let line = suggestExpenseFromDestination(state, input);
    if (!line.passCovered && Number(input.icFare) > 0 && Number(input.icFare) !== Number(line.icFare)) line = { ...line, icFare: Number(input.icFare), claimAmount: Number(input.icFare), hiddenZero: false, fareSource: "手入力", fareCheckedAt: undefined };
    const message = line.fareSource === "登録運賃" ? `端末内の登録運賃 ${yen(Number(line.icFare))} で自動計算しました。` : line.fareSource === "履歴・要確認" ? "以前の金額を候補にしました。現在のIC料金を確認し、確定してください。" : "初めての区間です。IC料金を入力して確定すると、次回から自動計算します。";
    mutate((draft) => {
      if (input.nearestStation && !draft.places.some((item) => item.name === input.destination.trim())) draft.places.push({ id: uid(), name: input.destination.trim(), nearestStation: input.nearestStation.trim(), route: line.routeDetails || "", reason: input.reason?.trim() || "", visitCount: 0, lastUsedAt: "" });
      const created: ExpenseLine = { id: uid(), date: input.date, startTime: input.startTime, destination: input.destination.trim(), origin: "", arrival: "", paidSection: "", icFare: 0, claimAmount: 0, reason: "", state: "未確認", routeOrder: 0, duplicateWarning: false, passCovered: false, hiddenZero: true, createdAt: new Date().toISOString(), ...line };
      return recomputeDuplicates({ ...draft, expenses: [...draft.expenses, created] });
    });
    setNotice(message);
  }

  function backup() {
    download(new Blob([JSON.stringify(state, null, 2)], { type: "application/json" }), `旅費申請バックアップ_${new Date().toISOString().slice(0, 10)}.json`);
    setNotice("JSONバックアップを保存しました");
  }
  async function restore(file: File) {
    try {
      const value = JSON.parse(await file.text()) as AppState;
      if (![1, 2, 3].includes(Number(value.version)) || !Array.isArray(value.expenses)) throw new Error();
      setState(normalizeState(value)); setNotice("バックアップを復元しました");
    } catch { setNotice("バックアップを復元できませんでした。ファイルを確認してください。"); }
  }

  if (!ready) return <main className="loading"><div className="brand-mark">旅</div><p>端末内の旅費データを準備しています…</p></main>;

  return <div className="app-shell">
    <header className="topbar">
      <div className="brand"><span className="brand-mark">旅</span><div><h1>出張旅費申請書作成アプリ</h1><p>毎月の旅費をすばやく入力し、ODS申請書を作成</p></div></div>
      <div className="privacy-pill"><span className="privacy-dot" />外部AIへ送信しません</div>
    </header>
    <nav className="tabs" aria-label="主な機能">{MAIN_TABS.map((name) => <button key={name} className={mainTab === name ? "active" : ""} onClick={() => setMainTab(name)}>{name}</button>)}</nav>
    <main>
      {mainTab === "旅費入力" && <section className="month-strip">
        <div className="month-title">
          <span className="eyebrow">出張旅費入力</span>
          <h2>{Number(state.selectedMonth.slice(0, 4))}年{Number(state.selectedMonth.slice(5, 7))}月 出張旅費入力</h2>
          <label className="month-select-field">
            <span>対象月</span>
            <select aria-label="対象月" value={state.selectedMonth} onChange={(event) => requestMonthChange(event.target.value)}>
              {monthOptions.map((month) => <option key={month} value={month}>{Number(month.slice(0, 4))}年{Number(month.slice(5, 7))}月</option>)}
            </select>
          </label>
          {requestedMonth && <div className="month-change-confirmation" role="alert">
            <strong>入力済みデータがあります。対象月を変更しますか？</strong>
            <span>{Number(requestedMonth.slice(0, 4))}年{Number(requestedMonth.slice(5, 7))}月へ切り替えます。現在の月のデータは削除されません。</span>
            <div>
              <button className="primary" onClick={() => switchMonth(requestedMonth)}>変更する</button>
              <button className="secondary" onClick={() => setRequestedMonth("")}>キャンセル</button>
            </div>
          </div>}
        </div>
        <div className="month-metrics"><div><strong>{monthExpenses.filter((line) => line.destination).length}</strong><span>入力件数</span></div><div><strong>{output.length}</strong><span>出力明細</span></div><div><strong>{yen(total)}</strong><span>申請合計</span></div></div>
      </section>}
      {notice && <div className="notice" role="status"><span>✓</span>{notice}<button aria-label="通知を閉じる" onClick={() => setNotice("")}>×</button></div>}
      {mainTab === "旅費入力" && <div className="travel-workflow">
        <TableEntryView state={state} mutate={mutate} setNotice={setNotice} />
        <OdsView state={state} lines={output} total={total} setNotice={setNotice} />
      </div>}
      {mainTab === "設定" && <Ver3SettingsView state={state} mutate={mutate} setNotice={setNotice} onBackup={backup} onRestore={restore} />}
    </main>
    <footer><span>保存先：このブラウザ内（IndexedDB）</span><span>最終保存：{new Date(state.lastSavedAt).toLocaleString("ja-JP")}</span></footer>
  </div>;

  function markSubmitted(lines: ExpenseLine[]) {
    if (!lines.length) return;
    mutate((draft) => { lines.forEach((out) => { const line = draft.expenses.find((x) => x.id === out.id); if (line) line.state = "申請済み"; }); return draft; });
    setNotice("今回の出力分を申請済みにしました");
  }
}

const TABLE_FIELDS = ["date", "destination", "paidSection", "icFare", "reason"] as const;

function emptyExpense(month: string, day = ""): ExpenseLine {
  return {
    id: uid(), date: day ? `${month}-${day}` : `${month}-`, startTime: "09:00", destination: "", origin: "", arrival: "",
    paidSection: "", icFare: 0, claimAmount: 0, reason: "", state: "未確認", routeOrder: 0,
    duplicateWarning: false, passCovered: false, hiddenZero: true, createdAt: new Date().toISOString(),
  };
}

function Ver3SettingsView({ state, mutate, setNotice, onBackup, onRestore }: {
  state: AppState;
  mutate: (updater: (draft: AppState) => AppState) => void;
  setNotice: (value: string) => void;
  onBackup: () => void;
  onRestore: (file: File) => Promise<void>;
}) {
  const [section, setSection] = useState<SettingsSection>("基本情報");
  return <div className="settings-workspace">
    <nav className="settings-sections" aria-label="設定メニュー">
      {SETTINGS_SECTIONS.map((name) => <button key={name} className={section === name ? "active" : ""} onClick={() => setSection(name)}>{name}</button>)}
    </nav>
    {section === "基本情報" && <BasicInfoView state={state} mutate={mutate} setNotice={setNotice} />}
    {section === "過去データ取込" && <PastClaimsImportView state={state} mutate={mutate} setNotice={setNotice} />}
    {section === "実績マスター管理" && <ClaimMasterView state={state} mutate={mutate} setNotice={setNotice} />}
    {section === "データ管理" && <DataManagementView onBackup={onBackup} onRestore={onRestore} />}
  </div>;
}

function BasicInfoView({ state, mutate, setNotice }: {
  state: AppState;
  mutate: (updater: (draft: AppState) => AppState) => void;
  setNotice: (value: string) => void;
}) {
  function updateProfile(patch: Partial<AppState["profile"]>) {
    mutate((draft) => ({ ...draft, profile: { ...draft.profile, ...patch } }));
  }
  function updatePass(id: string, patch: Partial<CommuterPass>) {
    mutate((draft) => {
      const pass = draft.commuterPasses.find((item) => item.id === id);
      if (pass) Object.assign(pass, patch);
      return draft;
    });
  }
  return <section className="panel basic-info-panel">
    <div className="panel-heading"><div><span className="eyebrow">設定</span><h2>基本情報</h2><p>申請書へ出力する所属・氏名と、保持する定期区間情報を登録します。</p></div></div>
    <div className="basic-profile-grid">
      <Field label="所属"><input aria-label="所属" value={state.profile.department} onChange={(event) => updateProfile({ department: event.target.value })} /></Field>
      <Field label="氏名"><input aria-label="氏名" value={state.profile.employeeName} onChange={(event) => updateProfile({ employeeName: event.target.value })} /></Field>
    </div>
    <div className="basic-section-heading"><div><h3>定期区間情報</h3><p>情報保持のみです。Ver3の旅費判定には使用しません。</p></div><button className="secondary" onClick={() => mutate((draft) => ({ ...draft, commuterPasses: [...draft.commuterPasses, { id: uid(), startStation: "", endStation: "", viaStations: "", lines: "", validFrom: "", validTo: "" }] }))}>＋ 定期区間を追加</button></div>
    <div className="basic-pass-list">
      {state.commuterPasses.map((pass, index) => <div className="basic-pass-row" key={pass.id}>
        <Field label={`定期区間 ${index + 1}・開始駅`}><input aria-label={`定期区間${index + 1} 開始駅`} value={pass.startStation} onChange={(event) => updatePass(pass.id, { startStation: event.target.value })} /></Field>
        <Field label="終了駅"><input aria-label={`定期区間${index + 1} 終了駅`} value={pass.endStation} onChange={(event) => updatePass(pass.id, { endStation: event.target.value })} /></Field>
        <Field label="経由駅"><input aria-label={`定期区間${index + 1} 経由駅`} value={pass.viaStations} onChange={(event) => updatePass(pass.id, { viaStations: event.target.value })} /></Field>
        <Field label="路線"><input aria-label={`定期区間${index + 1} 路線`} value={pass.lines} onChange={(event) => updatePass(pass.id, { lines: event.target.value })} /></Field>
        <button className="icon-button" onClick={() => mutate((draft) => ({ ...draft, commuterPasses: draft.commuterPasses.filter((item) => item.id !== pass.id) }))}>削除</button>
      </div>)}
      {!state.commuterPasses.length && <p className="muted">定期区間情報はまだ登録されていません。</p>}
    </div>
    <button className="primary settings-confirm" onClick={() => setNotice("基本情報を保存しました。")}>基本情報を保存</button>
  </section>;
}

function DataManagementView({ onBackup, onRestore }: { onBackup: () => void; onRestore: (file: File) => Promise<void> }) {
  const restoreRef = useRef<HTMLInputElement>(null);
  return <section className="panel data-management-panel">
    <div className="panel-heading"><div><span className="eyebrow">設定</span><h2>データ管理</h2><p>このブラウザ内の設定・月データ・実績マスターをバックアップまたは復元します。</p></div></div>
    <div className="data-management-actions">
      <article><h3>バックアップ</h3><p>現在のデータをJSONファイルとして保存します。</p><button className="primary" onClick={onBackup}>バックアップを保存</button></article>
      <article><h3>復元</h3><p>保存済みのJSONバックアップからデータを復元します。</p><button className="secondary" onClick={() => restoreRef.current?.click()}>バックアップから復元</button><input ref={restoreRef} hidden type="file" accept="application/json" onChange={(event) => {
        const file = event.target.files?.[0];
        if (file) void onRestore(file);
        event.currentTarget.value = "";
      }} /></article>
    </div>
  </section>;
}

function TableEntryView({ state, mutate, setNotice }: { state: AppState; mutate: (updater: (draft: AppState) => AppState) => void; setNotice: (value: string) => void }) {
  const rows = state.expenses.filter((line) => line.date.startsWith(state.selectedMonth)).sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  const masters = [...state.claimMasters].sort((a, b) => b.useCount - a.useCount || b.lastUsedDate.localeCompare(a.lastUsedDate));
  const pendingFocus = useRef<{ row: number; column: number } | null>(null);
  const [candidateRowId, setCandidateRowId] = useState("");
  const [candidateIndex, setCandidateIndex] = useState(-1);
  const [routeCandidateRowId, setRouteCandidateRowId] = useState("");
  const [routeCandidateIndex, setRouteCandidateIndex] = useState(-1);
  const [entryDrafts, setEntryDrafts] = useState<Record<string, { day?: string; amount?: string }>>({});

  useEffect(() => {
    if (rows.length) return;
    mutate((draft) => ({ ...draft, expenses: [...draft.expenses, emptyExpense(draft.selectedMonth)] }));
  }, [state.selectedMonth, rows.length]);

  useEffect(() => {
    const target = pendingFocus.current;
    if (!target) return;
    const input = document.querySelector<HTMLInputElement>(`[data-entry="${target.row}-${target.column}"]`);
    if (!input) return;
    input.focus();
    if (target.column === 0 || target.column === 3) input.select();
    pendingFocus.current = null;
  });

  function focusCell(row: number, column: number) {
    pendingFocus.current = { row, column };
    requestAnimationFrame(() => {
      const input = document.querySelector<HTMLInputElement>(`[data-entry="${row}-${column}"]`);
      if (!input) return;
      input.focus();
      if (column === 0 || column === 3) input.select();
      pendingFocus.current = null;
    });
  }

  function updateRow(id: string, patch: Partial<ExpenseLine>) {
    mutate((draft) => {
      const line = draft.expenses.find((item) => item.id === id);
      if (!line) return draft;
      Object.assign(line, patch);
      if (patch.icFare !== undefined) line.claimAmount = safeAmount(patch.icFare);
      line.hiddenZero = line.claimAmount === 0;
      return draft;
    });
  }

  function setEntryDraft(id: string, field: "day" | "amount", value: string) {
    setEntryDrafts((current) => ({ ...current, [id]: { ...current[id], [field]: value } }));
  }

  function clearEntryDraft(id: string, field: "day" | "amount") {
    setEntryDrafts((current) => {
      if (current[id]?.[field] === undefined) return current;
      const next = { ...current, [id]: { ...current[id] } };
      delete next[id][field];
      if (next[id].day === undefined && next[id].amount === undefined) delete next[id];
      return next;
    });
  }

  function commitDay(line: ExpenseLine): boolean {
    const draft = entryDrafts[line.id]?.day;
    if (draft === undefined) return true;
    const normalized = normalizeNumericText(draft);
    if (normalized && !/^\d{1,2}$/.test(normalized)) {
      setNotice("日は1～31の数字で入力してください。");
      return false;
    }
    const day = normalized ? Number(normalized) : 0;
    if (day > 31) {
      setNotice("日は1～31の数字で入力してください。");
      return false;
    }
    updateRow(line.id, { date: day ? `${state.selectedMonth}-${String(day).padStart(2, "0")}` : `${state.selectedMonth}-`, state: "未確認" });
    clearEntryDraft(line.id, "day");
    return true;
  }

  function commitAmount(line: ExpenseLine) {
    const draft = entryDrafts[line.id]?.amount;
    if (draft === undefined) return;
    updateRow(line.id, { icFare: safeAmount(draft), state: "未確認" });
    clearEntryDraft(line.id, "amount");
  }

  function applyMaster(id: string, master: ClaimMaster) {
    const stations = stationsFromSection(master.paidSection);
    updateRow(id, {
      destination: master.destination, paidSection: master.paidSection, icFare: master.icFare,
      claimAmount: master.icFare, reason: master.reason, origin: stations.origin, arrival: stations.arrival, state: "未確認",
    });
    clearEntryDraft(id, "amount");
    setCandidateRowId("");
    setCandidateIndex(-1);
    setRouteCandidateRowId("");
    setRouteCandidateIndex(-1);
  }

  function applyDestination(id: string, destination: string) {
    updateRow(id, {
      destination, paidSection: "", icFare: 0, claimAmount: 0, reason: "",
      origin: "", arrival: "", state: "未確認",
    });
    clearEntryDraft(id, "amount");
    setCandidateRowId("");
    setCandidateIndex(-1);
    setRouteCandidateRowId(id);
    setRouteCandidateIndex(-1);
  }

  function finishRow(id: string, rowIndex: number) {
    const current = rows.find((line) => line.id === id);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(current?.date ?? "") || !current?.destination.trim() || !current.paidSection.trim()) {
      setNotice("日付、目的地、区間を入力してください。");
      return;
    }
    mutate((draft) => {
      const line = draft.expenses.find((item) => item.id === id);
      if (!line) return draft;
      const firstCompletion = line.state === "未確認";
      line.icFare = safeAmount(line.icFare);
      line.claimAmount = line.icFare;
      line.hiddenZero = line.icFare === 0;
      line.state = "確認済み";
      const stations = stationsFromSection(line.paidSection);
      line.origin = stations.origin; line.arrival = stations.arrival;
      if (firstCompletion) {
        draft.claimMasters = mergeClaimMasters(draft.claimMasters, [{
          date: line.date, destination: line.destination.trim(), paidSection: line.paidSection.trim(),
          icFare: line.icFare, reason: line.reason.trim(),
        }], "画面入力");
      }
      const ordered = draft.expenses.filter((item) => item.date.startsWith(draft.selectedMonth)).sort((a, b) => a.createdAt.localeCompare(b.createdAt));
      if (rowIndex === ordered.length - 1) draft.expenses.push(emptyExpense(draft.selectedMonth, line.date.slice(8, 10)));
      return draft;
    });
    setNotice("1行を保存し、実績マスターを更新しました。");
    focusCell(rowIndex + 1, 0);
  }

  function handleEnter(event: React.KeyboardEvent<HTMLInputElement>, line: ExpenseLine, rowIndex: number, columnIndex: number) {
    if (event.key !== "Enter") return;
    event.preventDefault();
    if (columnIndex === TABLE_FIELDS.length - 1) finishRow(line.id, rowIndex);
    else focusCell(rowIndex, columnIndex + 1);
  }

  function matchingDestinations(value: string) {
    return claimDestinationCandidates(masters, value).slice(0, 8);
  }

  function matchingRoutes(destination: string) {
    return prioritizeClaimRouteCandidates(claimRouteCandidates(masters, destination));
  }

  function handleDestinationKey(event: React.KeyboardEvent<HTMLInputElement>, line: ExpenseLine, rowIndex: number) {
    const candidates = matchingDestinations(line.destination);
    if (event.key === "ArrowDown" && candidates.length) {
      event.preventDefault();
      setCandidateRowId(line.id);
      setCandidateIndex((current) => current < 0 ? 0 : (current + 1) % candidates.length);
      return;
    }
    if (event.key === "ArrowUp" && candidates.length) {
      event.preventDefault();
      setCandidateRowId(line.id);
      setCandidateIndex((current) => current < 0 ? candidates.length - 1 : (current - 1 + candidates.length) % candidates.length);
      return;
    }
    if (event.key === "Escape") {
      setCandidateRowId("");
      setCandidateIndex(-1);
      return;
    }
    if (event.key !== "Enter") return;
    event.preventDefault();
    if (candidateRowId === line.id && candidateIndex >= 0 && candidates[candidateIndex]) applyDestination(line.id, candidates[candidateIndex]);
    else {
      setCandidateRowId("");
      setCandidateIndex(-1);
    }
    focusCell(rowIndex, 2);
  }

  function handleRouteKey(event: React.KeyboardEvent<HTMLInputElement>, line: ExpenseLine, rowIndex: number) {
    const candidates = matchingRoutes(line.destination);
    if (event.key === "ArrowDown" && candidates.length) {
      event.preventDefault();
      setRouteCandidateRowId(line.id);
      setRouteCandidateIndex((current) => current < 0 ? 0 : (current + 1) % candidates.length);
      return;
    }
    if (event.key === "ArrowUp" && candidates.length) {
      event.preventDefault();
      setRouteCandidateRowId(line.id);
      setRouteCandidateIndex((current) => current < 0 ? candidates.length - 1 : (current - 1 + candidates.length) % candidates.length);
      return;
    }
    if (event.key === "Escape") {
      setRouteCandidateRowId("");
      setRouteCandidateIndex(-1);
      return;
    }
    if (event.key !== "Enter") return;
    event.preventDefault();
    if (routeCandidateRowId === line.id && routeCandidateIndex >= 0 && candidates[routeCandidateIndex]) applyMaster(line.id, candidates[routeCandidateIndex]);
    else {
      setRouteCandidateRowId("");
      setRouteCandidateIndex(-1);
    }
    focusCell(rowIndex, 3);
  }

  return <section className="panel entry-panel">
    <div className="panel-heading"><div><span className="eyebrow">ver3 メイン入力</span><h2>今月の旅費を入力</h2><p>1項目ごとにEnterで次へ進みます。目的地は過去実績から候補を選べます。</p></div><div className="summary-card"><span>今月の入力</span><strong>{rows.filter((line) => line.destination).length}</strong><small>保存済み {rows.filter((line) => line.state === "確認済み").length}行</small></div></div>
    <div className="entry-table">
      <div className="entry-head"><span>日</span><span>目的地</span><span>区間</span><span>金額</span><span>理由</span><span /></div>
      {rows.map((line, rowIndex) => {
        const candidates = candidateRowId === line.id ? matchingDestinations(line.destination) : [];
        const routeCandidates = matchingRoutes(line.destination);
        const showRouteCandidates = routeCandidateRowId === line.id;
        return <div className={`entry-row ${line.state === "確認済み" ? "complete" : ""}`} key={line.id}>
        <input data-entry={`${rowIndex}-0`} aria-label={`${rowIndex + 1}行目 日`} inputMode="numeric" value={entryDrafts[line.id]?.day ?? (Number(line.date.slice(8, 10)) || "")} onChange={(event) => {
          setEntryDraft(line.id, "day", event.target.value);
        }} onBlur={() => { commitDay(line); }} onFocus={(event) => event.currentTarget.select()} onKeyDown={(event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          if (commitDay(line)) focusCell(rowIndex, 1);
        }} />
        <div className="destination-cell">
          <input data-entry={`${rowIndex}-1`} role="combobox" aria-autocomplete="list" aria-label={`${rowIndex + 1}行目 目的地`} aria-expanded={candidateRowId === line.id && candidates.length > 0} aria-controls={`candidate-list-${line.id}`} value={line.destination} onFocus={() => {
            setCandidateRowId(line.id);
            setCandidateIndex(-1);
          }} onChange={(event) => {
            const value = event.target.value;
            updateRow(line.id, { destination: value, state: "未確認" });
            setCandidateRowId(line.id);
            setCandidateIndex(-1);
          }} onKeyDown={(event) => handleDestinationKey(event, line, rowIndex)} />
          {candidates.length > 0 && <div className="candidate-list destination-candidates" id={`candidate-list-${line.id}`} role="listbox" aria-label={`${rowIndex + 1}行目 目的地候補`}>
            {candidates.map((destination, index) => <button type="button" role="option" aria-selected={candidateIndex === index} className={candidateIndex === index ? "selected" : ""} key={destination} onMouseDown={(event) => event.preventDefault()} onClick={() => {
              applyDestination(line.id, destination);
              focusCell(rowIndex, 2);
            }}>
              <b>{destination}</b>
            </button>)}
          </div>}
        </div>
        <div className="route-cell">
          <input data-entry={`${rowIndex}-2`} role="combobox" aria-autocomplete="list" aria-label={`${rowIndex + 1}行目 区間`} aria-expanded={showRouteCandidates && routeCandidates.length > 0} aria-controls={`route-list-${line.id}`} value={line.paidSection} onFocus={() => {
            setRouteCandidateRowId(line.id);
            setRouteCandidateIndex(-1);
          }} onChange={(event) => {
            updateRow(line.id, { paidSection: event.target.value, state: "未確認" });
            setRouteCandidateRowId("");
            setRouteCandidateIndex(-1);
          }} onKeyDown={(event) => handleRouteKey(event, line, rowIndex)} />
          {showRouteCandidates && routeCandidates.length > 0 && <div className="candidate-list route-candidates" id={`route-list-${line.id}`} role="listbox" aria-label={`${rowIndex + 1}行目 区間候補`}>
            {routeCandidates.map((master, index) => <button type="button" role="option" aria-selected={routeCandidateIndex === index} className={routeCandidateIndex === index ? "selected" : ""} key={master.id} onMouseDown={(event) => event.preventDefault()} onClick={() => {
              applyMaster(line.id, master);
              focusCell(rowIndex, 3);
            }}>
              <b>{master.paidSection}</b><strong>{safeAmount(master.icFare).toLocaleString("ja-JP")}円</strong><small>{master.reason || "理由なし"}</small>
            </button>)}
          </div>}
          {showRouteCandidates && line.destination.trim() && routeCandidates.length === 0 && <div className="candidate-empty" role="status">一致する実績がありません</div>}
        </div>
        <input data-entry={`${rowIndex}-3`} aria-label={`${rowIndex + 1}行目 金額`} inputMode="numeric" value={entryDrafts[line.id]?.amount ?? (safeAmount(line.icFare) || "")} onFocus={(event) => event.currentTarget.select()} onChange={(event) => setEntryDraft(line.id, "amount", event.target.value)} onBlur={() => { commitAmount(line); }} onKeyDown={(event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          commitAmount(line);
          focusCell(rowIndex, 4);
        }} />
        <input data-entry={`${rowIndex}-4`} aria-label={`${rowIndex + 1}行目 理由`} value={line.reason} onChange={(event) => updateRow(line.id, { reason: event.target.value, state: "未確認" })} onKeyDown={(event) => handleEnter(event, line, rowIndex, 4)} />
        <button className="icon-button" aria-label={`${rowIndex + 1}行目を削除`} onClick={() => mutate((draft) => ({ ...draft, expenses: draft.expenses.filter((item) => item.id !== line.id) }))}>削除</button>
      </div>;
      })}
    </div>
    <button className="secondary add-entry-row" onClick={() => mutate((draft) => {
      const last = rows.at(-1); return { ...draft, expenses: [...draft.expenses, emptyExpense(draft.selectedMonth, last?.date.slice(8, 10) || "")] };
    })}>＋ 行を追加</button>
  </section>;
}

function ClaimMasterView({ state, mutate, setNotice }: { state: AppState; mutate: (updater: (draft: AppState) => AppState) => void; setNotice: (value: string) => void }) {
  const masters = [...state.claimMasters].sort((a, b) => b.useCount - a.useCount || b.lastUsedDate.localeCompare(a.lastUsedDate));
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">設定</span><h2>実績マスター管理</h2><p>入力候補の内容を編集・削除できます。</p></div><div className="summary-card"><span>登録済み</span><strong>{masters.length}</strong><small>実績候補</small></div></div>
    <div className="master-admin"><div className="master-admin-head"><span>目的地</span><span>区間</span><span>金額</span><span>理由</span><span>利用</span><span /></div>
      {masters.map((master) => <div key={master.id}>
        <input value={master.destination} onChange={(event) => mutate((draft) => { draft.claimMasters.find((item) => item.id === master.id)!.destination = event.target.value; return draft; })} />
        <input value={master.paidSection} onChange={(event) => mutate((draft) => { draft.claimMasters.find((item) => item.id === master.id)!.paidSection = event.target.value; return draft; })} />
        <input inputMode="numeric" value={safeAmount(master.icFare) || ""} onFocus={(event) => event.currentTarget.select()} onChange={(event) => mutate((draft) => { draft.claimMasters.find((item) => item.id === master.id)!.icFare = safeAmount(event.target.value); return draft; })} />
        <input value={master.reason} onChange={(event) => mutate((draft) => { draft.claimMasters.find((item) => item.id === master.id)!.reason = event.target.value; return draft; })} />
        <span>{master.useCount}回<br/><small>{master.lastUsedDate || "—"}</small></span>
        <button className="icon-button" onClick={() => { mutate((draft) => ({ ...draft, claimMasters: draft.claimMasters.filter((item) => item.id !== master.id) })); setNotice("実績マスターから削除しました。"); }}>削除</button>
      </div>)}
    </div>
    {!masters.length && <Empty title="実績マスターはまだありません" body="入力画面で1行を完成させると自動登録されます。" />}
  </section>;
}

function OdsView({ state, lines, total, setNotice }: { state: AppState; lines: ExpenseLine[]; total: number; setNotice: (value: string) => void }) {
  const [busy, setBusy] = useState(false);
  const hasValidMonth = /^\d{4}-(0[1-9]|1[0-2])$/.test(state.selectedMonth);
  async function generate() {
    setBusy(true);
    try {
      const response = await fetch("/2026年度版出張旅費代精算書原本.ods");
      if (!response.ok) throw new Error("ODS原本を読み込めません。");
      download(createOdsFromTemplate(await response.arrayBuffer(), state, lines), `${state.selectedMonth.replace("-", "")}_出張旅費精算書.ods`);
      setNotice("現行社内フォーマットのODSを作成しました。LibreOfficeで内容と印刷結果を確認してください。");
    } catch {
      setNotice("ODS申請書を作成できませんでした。入力データは失われていません。");
    } finally {
      setBusy(false);
    }
  }
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">ver3 ODS出力</span><h2>出張旅費精算書を作成</h2><p>空白行と0円行を除き、20件ごとにシートを分けます。</p></div><div className="summary-card"><span>出力対象</span><strong>{yen(total)}</strong><small>{lines.length}行・{Math.max(1, Math.ceil(lines.length / 20))}枚</small></div></div>
    {!hasValidMonth ? <div className="warning">対象月を選択してください。対象月なしではODSを作成できません。</div> : null}
    {!state.profile.department || !state.profile.employeeName ? <div className="warning">設定画面で所属と氏名を登録してください。</div> : null}
    <div className="ods-summary"><p>ファイル名</p><strong>{hasValidMonth ? `${state.selectedMonth.replace("-", "")}_出張旅費精算書.ods` : "対象月を選択してください"}</strong><span>現行社内原本の書式、罫線、結合セル、注意書き、印刷設定を維持します。合計はODS内の計算式で算出します。</span><button className="primary large" disabled={busy || !hasValidMonth || !lines.length || !state.profile.department || !state.profile.employeeName} onClick={generate}>{busy ? "作成中…" : "ODS申請書を作成"}</button></div>
  </section>;
}

function PastClaimsImportView({ state, mutate, setNotice }: { state: AppState; mutate: (updater: (draft: AppState) => AppState) => void; setNotice: (value: string) => void }) {
  const [busy, setBusy] = useState(false); const [preview, setPreview] = useState<ClaimImportPreviewRow[]>([]); const [fileName, setFileName] = useState("");
  async function readFile(file: File) {
    setBusy(true); setFileName(file.name);
    try {
      let rows: unknown[][] = [];
      if (/\.ods$/i.test(file.name)) {
        rows = parseOdsTableRows(await file.arrayBuffer());
      } else if (/\.xlsx$/i.test(file.name)) {
        const ExcelJS = await import("exceljs"); const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(await file.arrayBuffer());
        workbook.worksheets.forEach((sheet) => { for (let row = 11; row <= 30; row += 1) rows.push([sheet.getCell(`A${row}`).value, sheet.getCell(`B${row}`).value, sheet.getCell(`C${row}`).value, sheet.getCell(`D${row}`).value, sheet.getCell(`F${row}`).value, sheet.getCell(`G${row}`).value]); });
      } else {
        rows = (await file.text()).split(/\r?\n/).filter(Boolean).map((line) => line.split(file.name.toLowerCase().endsWith(".tsv") || line.includes("\t") ? "\t" : ","));
      }
      const parsed = parseClaimRows(rows, Number(state.selectedMonth.slice(0, 4)));
      setPreview(parsed.map((row) => ({ ...row, id: uid(), excluded: false })));
      setNotice(parsed.length ? `${parsed.length}行を読み取りました。内容を確認してマスタへ登録してください。` : "申請明細を検出できませんでした。原本形式または6列のCSVを確認してください。");
    } catch { setPreview([]); setNotice("過去請求データを読み込めませんでした。ODS・XLSX・CSV・TSVファイルを確認してください。"); }
    setBusy(false);
  }
  function updatePreview(id: string, patch: Partial<ClaimImportPreviewRow>) {
    setPreview((current) => current.map((row) => row.id === id ? { ...row, ...patch, icFare: patch.icFare === undefined ? row.icFare : safeAmount(patch.icFare) } : row));
  }
  function register() {
    const targets = prepareClaimRowsForRegistration(preview);
    if (!targets.length) return setNotice("登録対象の行がありません。");
    if (targets.some((row) => !/^\d{4}-\d{2}-\d{2}$/.test(row.date) || !row.destination || !row.paidSection)) return setNotice("登録対象行の日付・目的地・区間を確認してください。");
    mutate((draft) => {
      const before = draft.claimMasters.length; draft.claimMasters = mergeClaimMasters(draft.claimMasters, targets, fileName);
      draft.claimImports.push({ id: uid(), fileName, importedAt: new Date().toISOString(), rowCount: targets.length, addedCount: draft.claimMasters.length - before }); return draft;
    });
    setNotice(`${targets.length}行を過去実績マスタへ登録しました。`); setPreview([]); setFileName("");
  }
  const targetCount = preview.filter((row) => !row.excluded).length;
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">ver3 初期設定</span><h2>過去の出張旅費申請書を読み込む</h2><p>ODS・XLSX申請書、または「月・日・目的地・区間・金額・理由」の6列データから実績マスターを作ります。</p></div><div className="summary-card"><span>登録済み実績</span><strong>{state.claimMasters.length}</strong><small>{state.claimImports.length}ファイル取込済み</small></div></div>
    <label className="past-claim-drop"><input type="file" accept=".ods,.xlsx,.csv,.tsv,text/csv" onChange={(event) => event.target.files?.[0] && void readFile(event.target.files[0])} /><b>{busy ? "読み込み中…" : "過去の申請書を選択"}</b><span>ODS・XLSX・CSV・TSV／データは外部へ送信しません</span></label>
    {preview.length > 0 && <><div className="import-preview editable"><div className="import-preview-head"><span>日付</span><span>目的地</span><span>区間</span><span>金額</span><span>理由</span><span>登録対象</span><span>行操作</span></div>{preview.map((row, index) => <div className={row.excluded ? "excluded" : ""} key={row.id}>
      <input inputMode="numeric" placeholder="YYYY-MM-DD" aria-label={`${index + 1}行目 日付`} value={row.date} disabled={row.excluded} onChange={(event) => updatePreview(row.id, { date: event.target.value })} />
      <input aria-label={`${index + 1}行目 目的地`} value={row.destination} disabled={row.excluded} onChange={(event) => updatePreview(row.id, { destination: event.target.value })} />
      <input aria-label={`${index + 1}行目 区間`} value={row.paidSection} disabled={row.excluded} onChange={(event) => updatePreview(row.id, { paidSection: event.target.value })} />
      <input aria-label={`${index + 1}行目 金額`} inputMode="numeric" value={safeAmount(row.icFare) || ""} disabled={row.excluded} onChange={(event) => updatePreview(row.id, { icFare: safeAmount(event.target.value) })} />
      <input aria-label={`${index + 1}行目 理由`} value={row.reason} disabled={row.excluded} onChange={(event) => updatePreview(row.id, { reason: event.target.value })} />
      <button className={row.excluded ? "secondary" : "target-toggle"} aria-pressed={row.excluded} onClick={() => updatePreview(row.id, { excluded: !row.excluded })}>{row.excluded ? "対象に戻す" : "登録対象"}</button>
      <button className="icon-button" aria-label={`${index + 1}行目を削除`} onClick={() => setPreview((current) => current.filter((item) => item.id !== row.id))}>削除</button>
    </div>)}</div><div className="import-confirm"><span>{fileName}：読取 {preview.length}行／登録対象 {targetCount}行</span><button className="primary" disabled={!targetCount} onClick={register}>実績マスターへ登録</button></div></>}
    {state.claimImports.length > 0 && <div className="import-history"><h3>読込履歴</h3>{[...state.claimImports].reverse().map((item) => <div key={item.id}><b>{item.fileName}</b><span>{item.rowCount}行（新規 {item.addedCount}件）</span><small>{new Date(item.importedAt).toLocaleString("ja-JP")}</small></div>)}</div>}
  </section>;
}

function MasterEntryView({ state, mutate, setNotice, setTab }: { state: AppState; mutate: (updater: (draft: AppState) => AppState) => void; setNotice: (value: string) => void; setTab: (tab: Tab) => void }) {
  const initialDate = `${state.selectedMonth}-01`; const [date, setDate] = useState(initialDate); const [selectedId, setSelectedId] = useState("");
  const [draft, setDraft] = useState({ destination: "", paidSection: "", icFare: 0, reason: "", startTime: "09:00" });
  useEffect(() => { setDate(`${state.selectedMonth}-01`); }, [state.selectedMonth]);
  const masters = [...state.claimMasters].sort((a, b) => b.useCount - a.useCount || b.lastUsedDate.localeCompare(a.lastUsedDate));
  function select(master: ClaimMaster) { setSelectedId(master.id); setDraft({ destination: master.destination, paidSection: master.paidSection, icFare: master.icFare, reason: master.reason, startTime: "09:00" }); }
  function confirm() {
    if (!date || !draft.destination.trim() || !draft.paidSection.trim() || draft.icFare <= 0) return setNotice("日付・目的地・有料区間・IC料金を入力してください。");
    const stations = stationsFromSection(draft.paidSection); const now = new Date().toISOString();
    mutate((stateDraft) => {
      const routeOrder = stateDraft.expenses.filter((line) => line.date === date).length;
      stateDraft.expenses.push({ id: uid(), date, startTime: draft.startTime, destination: draft.destination.trim(), origin: stations.origin, arrival: stations.arrival, paidSection: draft.paidSection.trim(), icFare: draft.icFare, claimAmount: draft.icFare, reason: draft.reason.trim(), state: "確認済み", routeOrder, duplicateWarning: false, passCovered: false, hiddenZero: false, createdAt: now, fareSource: "登録運賃", fareCheckedAt: now });
      stateDraft.claimMasters = mergeClaimMasters(stateDraft.claimMasters, [{ date, destination: draft.destination.trim(), paidSection: draft.paidSection.trim(), icFare: draft.icFare, reason: draft.reason.trim() }], selectedId ? stateDraft.claimMasters.find((item) => item.id === selectedId)?.sourceName || "画面修正" : "新規入力");
      return stateDraft;
    });
    setNotice(`${date} の申請行を確定しました。続けて次の日付を選択できます。`); setSelectedId(""); setDraft({ destination: "", paidSection: "", icFare: 0, reason: "", startTime: "09:00" });
  }
  const confirmedDates = [...new Set(state.expenses.filter((line) => line.date.startsWith(state.selectedMonth) && ["確認済み", "修正済み"].includes(line.state)).map((line) => line.date))].sort();
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">ver2 主な入力</span><h2>日付を選び、過去実績から確定</h2><p>日付を変えながら実績を呼び出し、必要な箇所だけ修正して確定します。新しい訪問先も同じ欄へ直接入力できます。</p></div><div className="summary-card"><span>今月の確定日</span><strong>{confirmedDates.length}</strong><small>{state.expenses.filter((line) => line.date.startsWith(state.selectedMonth) && ["確認済み", "修正済み"].includes(line.state)).length}行</small></div></div>
    {!masters.length && <div className="start-import"><div><b>過去実績マスタがまだありません</b><span>先に過去の申請書を読み込むか、新規訪問先として直接入力できます。</span></div><button className="primary" onClick={() => setTab("過去データ読込")}>過去データ読込へ</button></div>}
    {masters.length > 0 && <div className="master-picker"><span>よく使う実績</span>{masters.slice(0, 12).map((master) => <button key={master.id} className={selectedId === master.id ? "selected" : ""} onClick={() => select(master)}><b>{master.destination}</b><small>{master.paidSection}・{yen(master.icFare)}・{master.useCount}回</small></button>)}</div>}
    <div className="claim-form"><Field label="月"><input value={Number(date.slice(5, 7)) || ""} readOnly /></Field><Field label="日"><input aria-label="申請日" type="date" value={date} onChange={(event) => setDate(event.target.value)} /></Field><Field label="教室名または目的地"><input value={draft.destination} onChange={(event) => setDraft((value) => ({ ...value, destination: event.target.value }))} /></Field><Field label="切符代がかかった区間"><input value={draft.paidSection} placeholder="池袋→浦和" onChange={(event) => setDraft((value) => ({ ...value, paidSection: event.target.value }))} /></Field><Field label="IC料金"><div className="money-input"><span>¥</span><input inputMode="numeric" value={draft.icFare || ""} onChange={(event) => setDraft((value) => ({ ...value, icFare: Math.max(0, Number(event.target.value)) }))} /></div></Field><Field label="移動の理由"><input value={draft.reason} onChange={(event) => setDraft((value) => ({ ...value, reason: event.target.value }))} /></Field><button className="primary claim-confirm" onClick={confirm}>この日を確定して次へ</button></div>
    {confirmedDates.length > 0 && <div className="confirmed-days"><span>確定済みの日付</span>{confirmedDates.map((value) => <button key={value} onClick={() => setDate(value)}>{Number(value.slice(5, 7))}/{Number(value.slice(8, 10))}</button>)}<button className="secondary" onClick={() => setTab("登録状況")}>一覧と書き出し対象を確認</button></div>}
  </section>;
}

function MonthlyView({ state, lines, total, warnings, showZero, setShowZero, onAdd, onQuickAdd, onUpdate, onRecalculate, onConfirm, onRemove, history }: any) {
  const suggestions = [...history].sort((a: any, b: any) => b.count - a.count).slice(0, 5);
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">旅費明細</span><h2>今月の移動</h2><p>0円区間は経路として保存され、通常一覧と出力から隠れます。</p></div><div className="heading-actions"><label className="switch"><input type="checkbox" checked={showZero} onChange={(e) => setShowZero(e.target.checked)} /><span />0円経路も表示</label><button className="primary" onClick={() => onAdd()}>＋ 旅費行を追加</button></div></div>
    <QuickAdd state={state} onAdd={onQuickAdd} />
    {warnings > 0 && <div className="warning">同じ日・区間・金額の重複候補が {warnings} 件あります。確認済みにする前に修正してください。</div>}
    {suggestions.length > 0 && <div className="suggestions"><b>よく使う確定経路</b>{suggestions.map((h: any) => <button key={h.id} onClick={() => onAdd({ destination: h.destination, origin: h.origin, arrival: h.arrival, paidSection: h.paidSection, reason: h.reason })}>{h.destination} · {h.paidSection}</button>)}</div>}
    <div className="expense-list">{lines.length ? lines.map((line: ExpenseLine) => <ExpenseCard key={line.id} line={line} onUpdate={onUpdate} onRecalculate={onRecalculate} onConfirm={onConfirm} onRemove={onRemove} />) : <Empty title="出力できる旅費行はまだありません" body="行き先を入力するか、旅費行を手入力してください。" />}</div>
    <div className="total-bar"><span>確認済み・出力対象</span><strong>{yen(total)}</strong></div>
  </section>;
}

function QuickAdd({ state, onAdd }: { state: AppState; onAdd: (input: { date: string; startTime: string; destination: string; nearestStation?: string; reason?: string; icFare?: number }) => Promise<void> }) {
  const today = new Date().toISOString().slice(0, 10); const initialDate = today.startsWith(state.selectedMonth) ? today : `${state.selectedMonth}-01`;
  const [date, setDate] = useState(initialDate); const [startTime, setStartTime] = useState("09:00"); const [destination, setDestination] = useState("");
  const [nearestStation, setNearestStation] = useState(""); const [reason, setReason] = useState(""); const [icFare, setIcFare] = useState(0); const [fareRegistered, setFareRegistered] = useState(false); const [busy, setBusy] = useState(false);
  useEffect(() => { const current = new Date().toISOString().slice(0, 10); setDate(current.startsWith(state.selectedMonth) ? current : `${state.selectedMonth}-01`); }, [state.selectedMonth]);
  const names = [...new Set([...state.places.map((item) => item.name), ...state.history.map((item) => item.destination)])];
  const ranked = names.map((name) => ({ name, score: (state.places.find((item) => item.name === name)?.visitCount || 0) + (state.history.filter((item) => item.destination === name).reduce((sum, item) => sum + item.count, 0) * 2), recent: state.history.filter((item) => item.destination === name).sort((a, b) => b.usedAt.localeCompare(a.usedAt))[0]?.usedAt || "" })).sort((a, b) => b.recent.localeCompare(a.recent) || b.score - a.score);
  function select(name: string) {
    setDestination(name); const place = state.places.find((item) => item.name === name); const history = [...state.history].filter((item) => item.destination === name).sort((a, b) => b.count - a.count || b.usedAt.localeCompare(a.usedAt))[0];
    const station = place?.nearestStation || history?.arrival || ""; setNearestStation(station); setReason(place?.reason || history?.reason || "");
    const suggestion = name.trim() ? suggestExpenseFromDestination(state, { date, startTime, destination: name, nearestStation: station }) : {};
    setIcFare(Number(suggestion.icFare || 0)); setFareRegistered(suggestion.fareSource === "登録運賃" || Boolean(suggestion.passCovered));
  }
  async function submit() {
    if (!date || !destination.trim()) return; setBusy(true); await onAdd({ date, startTime, destination, nearestStation, reason, icFare }); setBusy(false); setDestination(""); setNearestStation(""); setReason(""); setIcFare(0); setFareRegistered(false);
  }
  return <div className="quick-add"><div className="quick-add-title"><div><span className="eyebrow">最短入力</span><h3>日付と行き先から自動で作成</h3></div><span className="auto-hint">登録した区間はブラウザ内で自動計算</span></div>
    {ranked.length > 0 && <div className="destination-chips"><span>最近・よく使う</span>{ranked.slice(0, 6).map((item) => <button key={item.name} onClick={() => select(item.name)}>{item.name}</button>)}</div>}
    <div className="quick-add-grid"><Field label="日付"><input type="date" value={date} onChange={(event) => setDate(event.target.value)} /></Field><Field label="時刻"><input type="time" value={startTime} onChange={(event) => setStartTime(event.target.value)} /></Field><Field label="行き先"><input list="quick-destinations" placeholder="例：浦和高校" value={destination} onChange={(event) => select(event.target.value)} /><datalist id="quick-destinations">{ranked.map((item) => <option key={item.name} value={item.name} />)}</datalist></Field><Field label="最寄駅" hint="初回だけ。次回から自動"><input placeholder="例：浦和" value={nearestStation} onChange={(event) => { setNearestStation(event.target.value); setFareRegistered(false); }} /></Field><Field label="IC料金" hint={fareRegistered ? "登録運賃を自動表示" : "初回は入力してください"}><div className="money-input"><span>¥</span><input aria-label="簡単入力のIC料金" inputMode="numeric" placeholder="例：406" value={icFare || ""} onChange={(event) => { setIcFare(Math.max(0, Number(event.target.value))); setFareRegistered(false); }} /></div></Field><Field label="移動理由" hint="初回だけ。次回から自動"><input placeholder="例：学校訪問" value={reason} onChange={(event) => setReason(event.target.value)} /></Field><button className="primary quick-submit" disabled={busy || !destination.trim()} onClick={submit}>{busy ? "計算中…" : "追加して確認"}</button></div>
  </div>;
}

function ExpenseCard({ line, onUpdate, onRecalculate, onConfirm, onRemove }: { line: ExpenseLine; onUpdate: (id: string, patch: Partial<ExpenseLine>) => void; onRecalculate: (id: string) => void; onConfirm: (id: string) => void; onRemove: (id: string) => void }) {
  const edit = (patch: Partial<ExpenseLine>) => onUpdate(line.id, { ...patch, state: line.state === "申請済み" ? "申請済み" : "未確認" });
  return <article className={`expense-card ${line.duplicateWarning ? "duplicate" : ""}`}>
    <div className="expense-date"><input aria-label="移動日" type="date" value={line.date} onChange={(e) => onUpdate(line.id, { date: e.target.value })} /><input aria-label="開始時刻" type="time" value={line.startTime} onChange={(e) => onUpdate(line.id, { startTime: e.target.value })} /></div>
    <div className="expense-main"><input aria-label="目的地" placeholder="目的地" value={line.destination} onChange={(e) => edit({ destination: e.target.value })} /><div className="route-inputs"><input aria-label="出発駅" placeholder="出発駅" value={line.origin} onChange={(e) => edit({ origin: e.target.value, paidSection: `${e.target.value}→${line.arrival}` })} /><span>→</span><input aria-label="到着駅" placeholder="到着駅" value={line.arrival} onChange={(e) => edit({ arrival: e.target.value, paidSection: `${line.origin}→${e.target.value}` })} /></div><input aria-label="有料区間" placeholder="有料区間" value={line.paidSection} onChange={(e) => edit({ paidSection: e.target.value })} /></div>
    <div className="expense-meta"><Field label="IC料金"><div className="money-input"><span>¥</span><input aria-label="IC料金" inputMode="numeric" value={line.icFare || ""} onChange={(e) => edit({ icFare: Number(e.target.value), fareSource: "手入力" })} /></div></Field><Field label="移動理由"><input aria-label="移動理由" placeholder="学校訪問" value={line.reason} onChange={(e) => edit({ reason: e.target.value })} /></Field></div>
    <div className="expense-actions"><StatusBadge value={line.state} />{line.fareSource && <span className={`fare-source fare-${line.fareSource}`}>{line.fareSource}{line.fareCheckedAt ? ` ${new Date(line.fareCheckedAt).toLocaleDateString("ja-JP")}` : ""}</span>}{line.passCovered && <span className="pass-badge">定期券内 0円</span>}{line.duplicateWarning && <span className="duplicate-badge">重複候補</span>}<select aria-label="確認状態" value={line.state} onChange={(e) => e.target.value === "確認済み" ? onConfirm(line.id) : onUpdate(line.id, { state: e.target.value as ExpenseLine["state"] })}>{["未確認", "確認済み", "修正済み", "保留", "除外", "申請済み"].map((s) => <option key={s}>{s}</option>)}</select><div className="fare-buttons"><button className="secondary" onClick={() => onRecalculate(line.id)}>再計算</button><button className="primary" onClick={() => onConfirm(line.id)}>確定して登録</button></div><button className="icon-button" aria-label="旅費行を削除" onClick={() => onRemove(line.id)}>削除</button></div>
  </article>;
}

function ImportView({ state, mutate, setNotice }: any) {
  const [text, setText] = useState(""); const [processing, setProcessing] = useState(""); const [ocrText, setOcrText] = useState("");
  const monthCaptures = (state.captures || []).filter((item: ScheduleCapture) => item.month === state.selectedMonth);
  async function saveCaptures(files: File[]) {
    const images = files.filter((file) => file.type.startsWith("image/")); if (!images.length) return;
    const captures: ScheduleCapture[] = await Promise.all(images.map(async (file) => ({ id: uid(), month: state.selectedMonth, name: file.name || "貼り付け画像", dataUrl: await fileDataUrl(file), createdAt: new Date().toISOString() })));
    mutate((draft: AppState) => ({ ...draft, captures: [...(draft.captures || []), ...captures] })); setNotice(`${captures.length}枚のスクリーンショットを端末内に保存しました`);
  }
  function pasteCapture(event: React.ClipboardEvent<HTMLDivElement>) {
    const images = Array.from(event.clipboardData.items).filter((item) => item.type.startsWith("image/")).map((item) => item.getAsFile()).filter((file): file is File => Boolean(file));
    if (images.length) { event.preventDefault(); void saveCaptures(images); }
  }
  const addParsed = (items: ScheduleItem[]) => {
    if (!items.length) { setNotice("日付・時刻を含む予定を検出できませんでした。予定部分だけが見える画像を使うか、手入力してください。"); return; }
    mutate((d: AppState) => ({ ...d, schedules: [...d.schedules, ...items] })); setNotice(`${items.length}件を候補として取り込みました。内容を確認してください。`);
  };
  async function prepareImage(file: File): Promise<HTMLCanvasElement> {
    const bitmap = await createImageBitmap(file); const scale = Math.min(2, 2600 / bitmap.width);
    const canvas = document.createElement("canvas"); canvas.width = Math.round(bitmap.width * scale); canvas.height = Math.round(bitmap.height * scale);
    const context = canvas.getContext("2d", { willReadFrequently: true })!; context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    const image = context.getImageData(0, 0, canvas.width, canvas.height);
    for (let index = 0; index < image.data.length; index += 4) {
      const gray = image.data[index] * .299 + image.data[index + 1] * .587 + image.data[index + 2] * .114;
      const contrast = Math.max(0, Math.min(255, (gray - 128) * 1.35 + 128));
      image.data[index] = contrast; image.data[index + 1] = contrast; image.data[index + 2] = contrast;
    }
    context.putImageData(image, 0, 0); bitmap.close(); return canvas;
  }
  async function files(files: FileList | null) {
    if (!files?.length) return; setProcessing(`${files.length}ファイルを端末内で解析中…`);
    const items: ScheduleItem[] = [];
    try {
      for (const file of Array.from(files)) {
        if (file.type.includes("pdf")) {
          const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
          const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
          let content = ""; for (let p = 1; p <= pdf.numPages; p += 1) { const page = await pdf.getPage(p); const tc = await page.getTextContent(); content += tc.items.map((i: any) => i.str).join(" ") + "\n"; }
          setOcrText((current) => `${current}${current ? "\n\n" : ""}${content}`); items.push(...parseOcrSchedules(content, state.selectedMonth, "PDF"));
        } else if (file.type.startsWith("image/")) {
          const { createWorker, PSM } = await import("tesseract.js");
          const worker = await createWorker("jpn", 1, { langPath: "/tessdata" });
          await worker.setParameters({ tessedit_pageseg_mode: PSM.AUTO, preserve_interword_spaces: "1" });
          const result = await worker.recognize(await prepareImage(file)); await worker.terminate();
          setOcrText((current) => `${current}${current ? "\n\n" : ""}${result.data.text}`); items.push(...parseOcrSchedules(result.data.text, state.selectedMonth, "画像OCR"));
        } else if (file.name.toLowerCase().endsWith(".ics")) items.push(...parseIcsSchedules(await file.text()));
        else items.push(...parseTextSchedules(await file.text(), state.selectedMonth, file.name.endsWith(".csv") ? "CSV" : "テキスト"));
      }
      addParsed(items);
    } catch { setNotice("読み取りに失敗しました。画像・PDFは送信されていません。手入力またはテキスト貼り付けを利用できます。"); }
    setProcessing("");
  }
  function update(id: string, patch: Partial<ScheduleItem>) { mutate((d: AppState) => { Object.assign(d.schedules.find((x) => x.id === id)!, patch); return d; }); }
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">参考画像を端末内に保存</span><h2>予定スクリーンショット</h2><p>画像は読み取りに頼らず、確認用として貼り付け保存できます。外部AIには送信しません。</p></div></div>
    <div className="capture-paste" tabIndex={0} onPaste={pasteCapture}><div><b>ここを選んでスクリーンショットを貼り付け</b><span>Macは ⌘V、Windowsは Ctrl+V</span></div><label className="secondary capture-file">画像を選択<input hidden type="file" multiple accept="image/*" onChange={(event) => void saveCaptures(Array.from(event.target.files || []))} /></label></div>
    {monthCaptures.length > 0 && <div className="capture-grid">{monthCaptures.map((capture: ScheduleCapture) => <article key={capture.id}><a href={capture.dataUrl} target="_blank" rel="noreferrer"><img src={capture.dataUrl} alt={capture.name} /></a><div><span>{new Date(capture.createdAt).toLocaleString("ja-JP")}</span><button className="icon-button" onClick={() => mutate((draft: AppState) => ({ ...draft, captures: draft.captures.filter((item) => item.id !== capture.id) }))}>削除</button></div></article>)}</div>}
    <details className="ocr-lab"><summary><b>試験機能：画像・PDFから予定を読み取る</b><span>精度にばらつきがあるため補助機能として利用</span></summary><div className="import-grid"><label className="dropzone"><input type="file" multiple accept="image/*,.pdf,.csv,.txt,.ics" onChange={(e) => files(e.target.files)} /><b>画像・PDF・CSVを選択</b><span>OCR結果は自動確定しません</span></label><div className="paste-box"><textarea aria-label="予定テキスト" placeholder={'例：7/15\t10:00-11:00\t学校訪問\t浦和高校'} value={text} onChange={(e) => setText(e.target.value)} /><div className="button-row"><button className="secondary" onClick={() => addParsed([{ id: uid(), date: `${state.selectedMonth}-01`, startTime: "09:00", endTime: "10:00", title: "", location: "", isBusiness: true, hasTravel: true, confirmed: false, source: "手入力" }])}>手入力で予定を追加</button><button className="primary" onClick={() => { addParsed(parseTextSchedules(text, state.selectedMonth, text.includes(",") ? "CSV" : "テキスト")); setText(""); }}>テキストから候補作成</button></div></div></div></details>
    {processing && <div className="processing">{processing}</div>}{ocrText && <details className="ocr-raw"><summary>読み取った文字を確認</summary><pre>{ocrText}</pre></details>}
    <div className="section-title-row"><div><h3>OCR・取込結果の確認</h3><p className="section-note">自動確定はしません。日付・時刻・予定名・場所・業務区分・移動有無を確認してください。</p></div><button className="secondary" onClick={() => { mutate((d: AppState) => ({ ...d, schedules: d.schedules.filter((item) => item.confirmed || !item.date.startsWith(d.selectedMonth)) })); setNotice("今月の未確認候補を削除しました"); }}>今月の未確認候補をすべて削除</button></div>
    <div className="schedule-table"><div className="schedule-head"><span>日時</span><span>予定名・場所</span><span>区分</span><span>確認</span></div>{state.schedules.filter((s: ScheduleItem) => s.date.startsWith(state.selectedMonth)).map((item: ScheduleItem) => <div className="schedule-row" key={item.id}><div><input type="date" value={item.date} onChange={(e) => update(item.id, { date: e.target.value })} /><div className="inline"><input type="time" value={item.startTime} onChange={(e) => update(item.id, { startTime: e.target.value })} /><span>–</span><input type="time" value={item.endTime} onChange={(e) => update(item.id, { endTime: e.target.value })} /></div></div><div><input placeholder="予定名" value={item.title} onChange={(e) => update(item.id, { title: e.target.value })} /><input placeholder="場所" value={item.location} onChange={(e) => update(item.id, { location: e.target.value })} /></div><div><label><input type="checkbox" checked={item.isBusiness} onChange={(e) => update(item.id, { isBusiness: e.target.checked })} />業務</label><label><input type="checkbox" checked={item.hasTravel} onChange={(e) => update(item.id, { hasTravel: e.target.checked })} />移動あり</label></div><div><button className={item.confirmed ? "confirmed" : "secondary"} onClick={() => update(item.id, { confirmed: !item.confirmed })}>{item.confirmed ? "確認済み" : "確認する"}</button><button className="icon-button" onClick={() => mutate((d: AppState) => ({ ...d, schedules: d.schedules.filter((candidate) => candidate.id !== item.id) }))}>削除</button><small>{item.source}</small></div></div>)}</div>
  </section>;
}

function RouteView({ state, mutate, setNotice }: any) {
  const confirmed = state.schedules.filter((s: ScheduleItem) => s.confirmed && s.date.startsWith(state.selectedMonth));
  const dates = [...new Set(confirmed.map((s: ScheduleItem) => s.date))].sort() as string[];
  const [returns, setReturns] = useState<Record<string, string>>({});
  function build(date: string) {
    const items = confirmed.filter((s: ScheduleItem) => s.date === date);
    const route = buildDayRoute(items, state, returns[date] || "");
    mutate((d: AppState) => ({ ...d, expenses: [...d.expenses.filter((line) => !route.some((r) => r.sourceScheduleId && line.sourceScheduleId === r.sourceScheduleId)), ...route] }));
    setNotice(`${date} の移動経路を ${route.length}区間作成しました。IC料金を確認してください。`);
  }
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">予定を時刻順につなぐ</span><h2>1日の移動経路</h2><p>予定ごとの単純往復ではなく、前の訪問先から次の訪問先へ移動します。</p></div></div>{dates.length ? dates.map((date) => { const items = confirmed.filter((s: ScheduleItem) => s.date === date).sort((a: ScheduleItem, b: ScheduleItem) => a.startTime.localeCompare(b.startTime)); const weekday = new Date(`${date}T00:00:00`).getDay(); const rule = state.dayRules.find((r: any) => r.weekday === weekday); const historyReturn = state.history.find((h: any) => h.destination === items.at(-1)?.location)?.arrival; const candidates = [items.at(-1)?.location, historyReturn, rule?.returnPlace].filter(Boolean); return <article className="day-route" key={date}><header><div><strong>{new Date(`${date}T00:00:00`).toLocaleDateString("ja-JP", { month: "long", day: "numeric", weekday: "short" })}</strong><span>{items.length}件の予定</span></div><button className="primary" onClick={() => build(date)}>この日の経路を作成</button></header><div className="timeline">{items.map((item: ScheduleItem, index: number) => <div key={item.id}><span className="time">{item.startTime}</span><i /><div><b>{item.title}</b><small>{item.location}</small></div>{index < items.length - 1 && <span className="connector">次の予定へ</span>}</div>)}</div><Field label="戻り先（次の予定 → 当日指定 → 過去履歴 → 曜日設定の順）"><input list={`return-${date}`} value={returns[date] || ""} placeholder={rule?.returnPlace || "自宅"} onChange={(e) => setReturns((r) => ({ ...r, [date]: e.target.value }))} /><datalist id={`return-${date}`}>{candidates.map((c: string) => <option key={c} value={c} />)}</datalist></Field></article>; }) : <Empty title="確認済みの予定がありません" body="予定取込で内容を確認し、「確認済み」にしてください。" />}</section>;
}

function RegistrationView({ state, lines, output, total, mutate, setTab }: { state: AppState; lines: ExpenseLine[]; output: ExpenseLine[]; total: number; mutate: (updater: (draft: AppState) => AppState) => void; setTab: (tab: Tab) => void }) {
  const counts = { unconfirmed: lines.filter((line) => line.state === "未確認").length, confirmed: lines.filter((line) => ["確認済み", "修正済み"].includes(line.state)).length, hold: lines.filter((line) => line.state === "保留").length, zero: lines.filter((line) => line.claimAmount === 0).length };
  const sorted = [...lines].sort((a, b) => a.date.localeCompare(b.date) || a.startTime.localeCompare(b.startTime));
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">確認から出力まで</span><h2>登録状況</h2><p>未確認や保留を確認し、出力できる明細を一覧で確認できます。</p></div><div className="summary-card"><span>出力対象</span><strong>{yen(total)}</strong><small>{output.length}行</small></div></div>
    <div className="registration-metrics"><div><strong>{lines.length}</strong><span>全経路</span></div><div><strong>{counts.unconfirmed}</strong><span>未確認</span></div><div><strong>{counts.confirmed}</strong><span>確認済み</span></div><div><strong>{counts.hold}</strong><span>保留</span></div><div><strong>{counts.zero}</strong><span>0円</span></div></div>
    <div className="registration-actions"><button className="primary" onClick={() => setTab("コピー出力")}>コピー出力へ</button><button className="secondary" onClick={() => setTab("Excel出力")}>Excel出力へ</button></div>
    <h3>今月の明細</h3>{sorted.length ? <div className="registration-list">{sorted.map((line) => <div key={line.id}><span>{new Date(`${line.date}T00:00:00`).toLocaleDateString("ja-JP", { month: "numeric", day: "numeric" })}</span><b>{line.destination || "目的地未入力"}</b><span>{line.paidSection || `${line.origin}→${line.arrival}`}</span><strong>{yen(line.claimAmount)}</strong><StatusBadge value={line.state} /></div>)}</div> : <Empty title="今月の登録はありません" body="月間画面で行き先を入力すると、ここに表示されます。" />}
    <div className="fare-ledger-heading"><div><h3>ブラウザ内の運賃台帳</h3><p>確定した区間だけをこの端末に保存し、次回の自動計算に使います。</p></div><span>{state.fareRules.length}区間</span></div>
    {state.fareRules.length ? <div className="fare-ledger">{[...state.fareRules].sort((a, b) => b.lastUsedAt.localeCompare(a.lastUsedAt)).map((rule) => <div key={rule.id}><div><b>{rule.origin} → {rule.arrival}</b><small>利用 {rule.useCount}回・登録 {new Date(rule.registeredAt).toLocaleDateString("ja-JP")}</small></div><input aria-label={`${rule.origin}から${rule.arrival}の有料区間`} value={rule.paidSection} onChange={(event) => mutate((draft) => { draft.fareRules.find((item) => item.id === rule.id)!.paidSection = event.target.value; return draft; })} /><div className="money-input"><span>¥</span><input aria-label={`${rule.origin}から${rule.arrival}のIC料金`} inputMode="numeric" value={rule.icFare} onChange={(event) => mutate((draft) => { draft.fareRules.find((item) => item.id === rule.id)!.icFare = Math.max(0, Number(event.target.value)); return draft; })} /></div><button className="icon-button" onClick={() => mutate((draft) => ({ ...draft, fareRules: draft.fareRules.filter((item) => item.id !== rule.id) }))}>削除</button></div>)}</div> : <p className="muted">まだ登録運賃はありません。旅費行でIC料金を確認し、「確定して登録」を押すと追加されます。</p>}
  </section>;
}

function CopyView({ lines, total, setNotice, onSubmitted }: any) {
  const pages = copyPages(lines); const [page, setPage] = useState(0); const current = pages[Math.min(page, pages.length - 1)];
  const copy = async (text: string) => { await navigator.clipboard.writeText(text); setNotice("クリップボードへコピーしました"); };
  const columns = [{ label: "月", cell: "A11", get: (l: ExpenseLine) => Number(l.date.slice(5, 7)) }, { label: "日", cell: "B11", get: (l: ExpenseLine) => Number(l.date.slice(8, 10)) }, { label: "目的地", cell: "C11", get: (l: ExpenseLine) => l.destination }, { label: "有料区間", cell: "D11", get: (l: ExpenseLine) => l.paidSection }, { label: "IC料金", cell: "F11", get: (l: ExpenseLine) => l.claimAmount }, { label: "移動理由", cell: "G11", get: (l: ExpenseLine) => l.reason }];
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">Excel自動作成が使えなくても安心</span><h2>Excel貼り付け用コピー</h2><p>結合セルに合わせて、全列または列ごとにコピーできます。</p></div><div className="summary-card"><span>総合計</span><strong>{yen(total)}</strong><small>{lines.length}行・{pages.length}枚</small></div></div>
    <div className="page-tabs">{pages.map((_: any, index: number) => <button key={index} className={page === index ? "active" : ""} onClick={() => setPage(index)}>{index + 1}枚目</button>)}</div>
    <div className="copy-actions"><button className="primary" disabled={!current.length} onClick={() => copy(tabSeparated(current))}>全6列をタブ区切りでコピー</button><span>1枚目はExcelのA11を選んで貼り付け</span></div>
    <div className="copy-table"><div className="copy-head">{columns.map((c) => <span key={c.label}>{c.label}</span>)}</div>{current.map((line: ExpenseLine) => <div className="copy-row" key={line.id}>{columns.map((c) => <span key={c.label}>{c.get(line)}</span>)}</div>)}</div>
    <div className="column-copy-grid">{columns.map((column) => <button key={column.label} onClick={() => copy(current.map((line: ExpenseLine) => String(column.get(line))).join("\n"))}><span>{column.label}をコピー</span><small>貼り付け先 {column.cell}</small></button>)}</div>
    <div className="submit-panel"><div><b>内容をExcelで確認できましたか？</b><p>出力しただけでは申請済みになりません。</p></div><button className="danger-safe" disabled={!lines.length} onClick={onSubmitted}>今回の出力分を申請済みにする</button></div>
  </section>;
}

function ExcelView({ state, lines, total, submissionDate, setSubmissionDate, setNotice, onSubmitted }: any) {
  const [template, setTemplate] = useState<File | null>(null); const [busy, setBusy] = useState(false); const [failed, setFailed] = useState(false);
  async function generate() {
    if (!template) return setNotice("2026年度版出張旅費代精算書原本（XLSX）を選択してください。");
    if (!state.profile.department || !state.profile.employeeName) return setNotice("設定で所属と氏名を登録してください。");
    setBusy(true); setFailed(false);
    try {
      const blob = await createExcel(await template.arrayBuffer(), state, lines, submissionDate);
      download(blob, `${state.selectedMonth.replace("-", "年")}月_出張旅費代精算書_${state.profile.employeeName.replace(/\s/g, "")}.xlsx`);
      setNotice("Excel申請書を作成しました。内容を確認してから申請済みにしてください。");
    } catch { setFailed(true); setNotice("Excel申請書を自動作成できませんでした。旅費データは失われていません。Excel貼り付け用コピーを使用できます。"); }
    setBusy(false);
  }
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">2026年度版原本をコピーして作成</span><h2>Excel申請書を自動作成</h2><p>原本ファイルは上書きしません。書式とF31の数式を保持した新しいXLSXを保存します。</p></div><div className="summary-card"><span>申請合計</span><strong>{yen(total)}</strong><small>{lines.length}行・{Math.max(1, Math.ceil(lines.length / 20))}枚</small></div></div>
    {failed && <div className="error-banner"><b>Excel申請書を自動作成できませんでした。</b><span>旅費データは失われていません。Excel貼り付け用コピーを使用できます。</span></div>}
    <div className="excel-form"><Field label="2026年度版出張旅費代精算書原本"><label className="file-picker"><input type="file" accept=".xlsx" onChange={(e) => setTemplate(e.target.files?.[0] ?? null)} /><span>{template?.name || "XLSX原本を選択"}</span></label></Field><Field label="提出日"><input type="date" value={submissionDate} onChange={(e) => setSubmissionDate(e.target.value)} /></Field><div className="excel-check"><span>✓ 原本を直接上書きしません</span><span>✓ 20行ごとにシートを複製</span><span>✓ F31は各用紙の小計</span></div><button className="primary large" disabled={busy || !lines.length} onClick={generate}>{busy ? "作成中…" : "Excel申請書を作成"}</button></div>
    <div className="submit-panel"><div><b>出力しただけでは申請済みになりません</b><p>Excelの内容を確認した後に押してください。</p></div><button className="danger-safe" disabled={!lines.length} onClick={onSubmitted}>今回の出力分を申請済みにする</button></div>
  </section>;
}

function SettingsView({ state, mutate, setNotice }: any) {
  const addBase = () => mutate((d: AppState) => ({ ...d, workBases: [...d.workBases, { id: uid(), name: "", station: "" }] }));
  const addPass = () => mutate((d: AppState) => ({ ...d, commuterPasses: [...d.commuterPasses, { id: uid(), startStation: "", endStation: "", viaStations: "", lines: "", validFrom: "", validTo: "" }] }));
  const addPlace = () => mutate((d: AppState) => ({ ...d, places: [...d.places, { id: uid(), name: "", nearestStation: "", route: "", reason: "", visitCount: 0, lastUsedAt: "" }] }));
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">最初に登録</span><h2>利用者・移動ルール設定</h2><p>確定履歴と組み合わせ、次回から候補を優先表示します。</p></div></div>
    <div className="settings-grid"><div className="settings-card"><h3>利用者</h3><Field label="所属"><input value={state.profile.department} onChange={(e) => mutate((d: AppState) => { d.profile.department = e.target.value; return d; })} /></Field><Field label="氏名"><input value={state.profile.employeeName} onChange={(e) => mutate((d: AppState) => { d.profile.employeeName = e.target.value; return d; })} /></Field><Field label="自宅の表示名"><input value={state.profile.homeName} onChange={(e) => mutate((d: AppState) => { d.profile.homeName = e.target.value; return d; })} /></Field><Field label="自宅最寄駅"><input value={state.profile.homeStation} onChange={(e) => mutate((d: AppState) => { d.profile.homeStation = e.target.value; return d; })} /></Field></div>
      <div className="settings-card"><div className="card-title"><h3>出勤先</h3><button onClick={addBase}>＋追加</button></div>{state.workBases.map((base: any) => <div className="paired" key={base.id}><input placeholder="本部" value={base.name} onChange={(e) => mutate((d: AppState) => { d.workBases.find((x) => x.id === base.id)!.name = e.target.value; return d; })} /><input placeholder="池袋駅" value={base.station} onChange={(e) => mutate((d: AppState) => { d.workBases.find((x) => x.id === base.id)!.station = e.target.value; return d; })} /></div>)}</div>
      <div className="settings-card wide"><div className="card-title"><h3>定期券</h3><button onClick={addPass}>＋追加</button></div>{state.commuterPasses.length ? state.commuterPasses.map((pass: CommuterPass) => <div className="pass-row" key={pass.id}>{(["startStation", "endStation", "viaStations", "lines", "validFrom", "validTo"] as const).map((key) => <input key={key} type={key.startsWith("valid") ? "date" : "text"} aria-label={key} placeholder={{ startStation: "開始駅", endStation: "終了駅", viaStations: "経由駅", lines: "路線", validFrom: "", validTo: "" }[key]} value={pass[key]} onChange={(e) => mutate((d: AppState) => { (d.commuterPasses.find((x) => x.id === pass.id)![key] as string) = e.target.value; return d; })} />)}</div>) : <p className="muted">定期券がない場合は未登録のままで構いません。出力時は「定期券なし」と表示します。</p>}</div>
      <div className="settings-card wide"><div className="card-title"><h3>訪問先・過去確定候補</h3><button onClick={addPlace}>＋追加</button></div>{state.places.length ? state.places.map((place: any) => <div className="place-row" key={place.id}><input placeholder="申請書表示名" value={place.name} onChange={(e) => mutate((d: AppState) => { d.places.find((x) => x.id === place.id)!.name = e.target.value; return d; })} /><input placeholder="最寄駅" value={place.nearestStation} onChange={(e) => mutate((d: AppState) => { d.places.find((x) => x.id === place.id)!.nearestStation = e.target.value; return d; })} /><input placeholder="よく使う経路" value={place.route} onChange={(e) => mutate((d: AppState) => { d.places.find((x) => x.id === place.id)!.route = e.target.value; return d; })} /><input placeholder="移動理由" value={place.reason} onChange={(e) => mutate((d: AppState) => { d.places.find((x) => x.id === place.id)!.reason = e.target.value; return d; })} /></div>) : <p className="muted">確定した経路は自動で履歴に蓄積されます。ここでは訪問先を先に登録できます。</p>}</div>
      <div className="settings-card wide"><h3>曜日別の標準出発地・戻り先</h3><div className="weekday-grid">{state.dayRules.map((rule: any) => <div key={rule.weekday}><b>{WEEKDAYS[rule.weekday]}曜</b><input value={rule.startPlace} onChange={(e) => mutate((d: AppState) => { d.dayRules[rule.weekday].startPlace = e.target.value; return d; })} /><span>→</span><input value={rule.returnPlace} onChange={(e) => mutate((d: AppState) => { d.dayRules[rule.weekday].returnPlace = e.target.value; return d; })} /></div>)}</div></div>
    </div><button className="primary settings-save" onClick={() => setNotice("設定を端末内に保存しました")}>設定を保存</button>
  </section>;
}

function Empty({ title, body }: { title: string; body: string }) { return <div className="empty"><span>↗</span><b>{title}</b><p>{body}</p></div>; }
