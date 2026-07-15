"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { loadState, saveState } from "./lib/db";
import { buildDayRoute, copyPages, duplicateKeys, isPassCovered, outputLines, parseIcsSchedules, parseTextSchedules, tabSeparated, uid, yen } from "./lib/domain";
import { createExcel } from "./lib/excel";
import { EMPTY_STATE, type AppState, type CommuterPass, type ExpenseLine, type ScheduleItem } from "./lib/types";

type Tab = "月間" | "予定取込" | "経路確認" | "コピー出力" | "Excel出力" | "設定";
const TABS: Tab[] = ["月間", "予定取込", "経路確認", "コピー出力", "Excel出力", "設定"];
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob); const anchor = document.createElement("a");
  anchor.href = url; anchor.download = filename; anchor.click(); URL.revokeObjectURL(url);
}

function Field({ label, children, hint }: { label: string; children: React.ReactNode; hint?: string }) {
  return <label className="field"><span>{label}</span>{children}{hint && <small>{hint}</small>}</label>;
}

function StatusBadge({ value }: { value: string }) { return <span className={`status status-${value}`}>{value}</span>; }

export function TravelExpenseApp() {
  const [state, setState] = useState<AppState>(EMPTY_STATE);
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState<Tab>("月間");
  const [showZero, setShowZero] = useState(false);
  const [notice, setNotice] = useState("データはこの端末内だけに保存されます");
  const [submissionDate, setSubmissionDate] = useState(new Date().toISOString().slice(0, 10));
  const importRef = useRef<HTMLInputElement>(null);

  useEffect(() => { loadState().then((saved) => { if (saved) setState(saved); setReady(true); }); }, []);
  useEffect(() => { if (!ready) return; const timer = setTimeout(() => saveState({ ...state, lastSavedAt: new Date().toISOString() }), 250); return () => clearTimeout(timer); }, [state, ready]);

  const monthExpenses = useMemo(() => state.expenses.filter((line) => line.date.startsWith(state.selectedMonth)), [state]);
  const visibleExpenses = useMemo(() => monthExpenses.filter((line) => showZero || line.claimAmount > 0), [monthExpenses, showZero]);
  const output = useMemo(() => outputLines(state), [state]);
  const total = output.reduce((sum, line) => sum + line.claimAmount, 0);
  const warnings = monthExpenses.filter((line) => line.duplicateWarning).length;

  function mutate(updater: (draft: AppState) => AppState) { setState((current) => updater(structuredClone(current))); }
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
      line.passCovered = isPassCovered(line.origin, line.arrival, line.date, draft);
      line.claimAmount = line.passCovered ? 0 : Math.max(0, Number(line.icFare || 0));
      line.hiddenZero = line.claimAmount === 0;
      if (patch.state === "確認済み" || patch.state === "修正済み") {
        const existing = draft.history.find((h) => h.destination === line.destination && h.origin === line.origin && h.arrival === line.arrival);
        if (existing) { existing.count += 1; existing.usedAt = new Date().toISOString(); existing.paidSection = line.paidSection; existing.reason = line.reason; }
        else draft.history.push({ id: uid(), destination: line.destination, origin: line.origin, arrival: line.arrival, paidSection: line.paidSection, reason: line.reason, usedAt: new Date().toISOString(), count: 1 });
      }
      return recomputeDuplicates(draft);
    });
  }
  function removeExpense(id: string) { mutate((draft) => ({ ...draft, expenses: draft.expenses.filter((line) => line.id !== id) })); }

  function backup() {
    download(new Blob([JSON.stringify(state, null, 2)], { type: "application/json" }), `旅費申請バックアップ_${new Date().toISOString().slice(0, 10)}.json`);
    setNotice("JSONバックアップを保存しました");
  }
  async function restore(file: File) {
    try {
      const value = JSON.parse(await file.text()) as AppState;
      if (value.version !== 1 || !Array.isArray(value.expenses)) throw new Error();
      setState(value); setNotice("バックアップを復元しました");
    } catch { setNotice("バックアップを復元できませんでした。ファイルを確認してください。"); }
  }

  if (!ready) return <main className="loading"><div className="brand-mark">旅</div><p>端末内の旅費データを準備しています…</p></main>;

  return <div className="app-shell">
    <header className="topbar">
      <div className="brand"><span className="brand-mark">旅</span><div><h1>出張旅費申請書作成アプリ</h1><p>1日の移動をつなげて、申請できる区間だけを整理</p></div></div>
      <div className="privacy-pill"><span className="privacy-dot" />外部AIへ送信しません</div>
    </header>
    <nav className="tabs" aria-label="主な機能">{TABS.map((name) => <button key={name} className={tab === name ? "active" : ""} onClick={() => setTab(name)}>{name}</button>)}</nav>
    <main>
      <section className="month-strip">
        <div><span className="eyebrow">対象月</span><input aria-label="対象月" type="month" value={state.selectedMonth} onChange={(e) => mutate((d) => ({ ...d, selectedMonth: e.target.value }))} /></div>
        <div className="month-metrics"><div><strong>{monthExpenses.length}</strong><span>内部経路</span></div><div><strong>{output.length}</strong><span>出力明細</span></div><div><strong>{yen(total)}</strong><span>申請合計</span></div></div>
        <div className="header-actions"><button className="secondary" onClick={backup}>バックアップ</button><button className="secondary" onClick={() => importRef.current?.click()}>復元</button><input ref={importRef} hidden type="file" accept="application/json" onChange={(e) => e.target.files?.[0] && restore(e.target.files[0])} /></div>
      </section>
      {notice && <div className="notice" role="status"><span>✓</span>{notice}<button aria-label="通知を閉じる" onClick={() => setNotice("")}>×</button></div>}
      {tab === "月間" && <MonthlyView lines={visibleExpenses} total={total} warnings={warnings} showZero={showZero} setShowZero={setShowZero} onAdd={addExpense} onUpdate={updateExpense} onRemove={removeExpense} history={state.history} />}
      {tab === "予定取込" && <ImportView state={state} mutate={mutate} setNotice={setNotice} />}
      {tab === "経路確認" && <RouteView state={state} mutate={mutate} setNotice={setNotice} onUpdate={updateExpense} />}
      {tab === "コピー出力" && <CopyView lines={output} total={total} setNotice={setNotice} onSubmitted={() => markSubmitted(output)} />}
      {tab === "Excel出力" && <ExcelView state={state} lines={output} total={total} submissionDate={submissionDate} setSubmissionDate={setSubmissionDate} setNotice={setNotice} onSubmitted={() => markSubmitted(output)} />}
      {tab === "設定" && <SettingsView state={state} mutate={mutate} setNotice={setNotice} />}
    </main>
    <footer><span>保存先：このブラウザ内（IndexedDB）</span><span>最終保存：{new Date(state.lastSavedAt).toLocaleString("ja-JP")}</span></footer>
  </div>;

  function markSubmitted(lines: ExpenseLine[]) {
    if (!lines.length) return;
    mutate((draft) => { lines.forEach((out) => { const line = draft.expenses.find((x) => x.id === out.id); if (line) line.state = "申請済み"; }); return draft; });
    setNotice("今回の出力分を申請済みにしました");
  }
}

function MonthlyView({ lines, total, warnings, showZero, setShowZero, onAdd, onUpdate, onRemove, history }: any) {
  const suggestions = [...history].sort((a: any, b: any) => b.count - a.count).slice(0, 5);
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">旅費明細</span><h2>今月の移動</h2><p>0円区間は経路として保存され、通常一覧と出力から隠れます。</p></div><div className="heading-actions"><label className="switch"><input type="checkbox" checked={showZero} onChange={(e) => setShowZero(e.target.checked)} /><span />0円経路も表示</label><button className="primary" onClick={() => onAdd()}>＋ 旅費行を追加</button></div></div>
    {warnings > 0 && <div className="warning">同じ日・区間・金額の重複候補が {warnings} 件あります。確認済みにする前に修正してください。</div>}
    {suggestions.length > 0 && <div className="suggestions"><b>よく使う確定経路</b>{suggestions.map((h: any) => <button key={h.id} onClick={() => onAdd({ destination: h.destination, origin: h.origin, arrival: h.arrival, paidSection: h.paidSection, reason: h.reason })}>{h.destination} · {h.paidSection}</button>)}</div>}
    <div className="expense-list">{lines.length ? lines.map((line: ExpenseLine) => <ExpenseCard key={line.id} line={line} onUpdate={onUpdate} onRemove={onRemove} />) : <Empty title="出力できる旅費行はまだありません" body="予定を取り込んで1日経路を作るか、旅費行を手入力してください。" />}</div>
    <div className="total-bar"><span>確認済み・出力対象</span><strong>{yen(total)}</strong></div>
  </section>;
}

function ExpenseCard({ line, onUpdate, onRemove }: { line: ExpenseLine; onUpdate: (id: string, patch: Partial<ExpenseLine>) => void; onRemove: (id: string) => void }) {
  return <article className={`expense-card ${line.duplicateWarning ? "duplicate" : ""}`}>
    <div className="expense-date"><input aria-label="移動日" type="date" value={line.date} onChange={(e) => onUpdate(line.id, { date: e.target.value })} /><input aria-label="開始時刻" type="time" value={line.startTime} onChange={(e) => onUpdate(line.id, { startTime: e.target.value })} /></div>
    <div className="expense-main"><input aria-label="目的地" placeholder="目的地" value={line.destination} onChange={(e) => onUpdate(line.id, { destination: e.target.value, state: line.state === "確認済み" ? "修正済み" : line.state })} /><div className="route-inputs"><input aria-label="出発駅" placeholder="出発駅" value={line.origin} onChange={(e) => onUpdate(line.id, { origin: e.target.value, paidSection: `${e.target.value}→${line.arrival}` })} /><span>→</span><input aria-label="到着駅" placeholder="到着駅" value={line.arrival} onChange={(e) => onUpdate(line.id, { arrival: e.target.value, paidSection: `${line.origin}→${e.target.value}` })} /></div><input aria-label="有料区間" placeholder="有料区間" value={line.paidSection} onChange={(e) => onUpdate(line.id, { paidSection: e.target.value })} /></div>
    <div className="expense-meta"><Field label="IC料金"><div className="money-input"><span>¥</span><input aria-label="IC料金" inputMode="numeric" value={line.icFare || ""} onChange={(e) => onUpdate(line.id, { icFare: Number(e.target.value) })} /></div></Field><Field label="移動理由"><input aria-label="移動理由" placeholder="学校訪問" value={line.reason} onChange={(e) => onUpdate(line.id, { reason: e.target.value })} /></Field></div>
    <div className="expense-actions"><StatusBadge value={line.state} />{line.passCovered && <span className="pass-badge">定期券内 0円</span>}{line.duplicateWarning && <span className="duplicate-badge">重複候補</span>}<select aria-label="確認状態" value={line.state} onChange={(e) => onUpdate(line.id, { state: e.target.value as ExpenseLine["state"] })}>{["未確認", "確認済み", "修正済み", "保留", "除外", "申請済み"].map((s) => <option key={s}>{s}</option>)}</select><button className="icon-button" aria-label="旅費行を削除" onClick={() => onRemove(line.id)}>削除</button></div>
  </article>;
}

function ImportView({ state, mutate, setNotice }: any) {
  const [text, setText] = useState(""); const [processing, setProcessing] = useState("");
  const addParsed = (items: ScheduleItem[]) => { mutate((d: AppState) => ({ ...d, schedules: [...d.schedules, ...items] })); setNotice(`${items.length}件を候補として取り込みました。内容を確認してください。`); };
  async function files(files: FileList | null) {
    if (!files?.length) return; setProcessing(`${files.length}ファイルを端末内で解析中…`);
    const items: ScheduleItem[] = [];
    try {
      for (const file of Array.from(files)) {
        if (file.type.includes("pdf")) {
          const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
          const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
          let content = ""; for (let p = 1; p <= pdf.numPages; p += 1) { const page = await pdf.getPage(p); const tc = await page.getTextContent(); content += tc.items.map((i: any) => i.str).join(" ") + "\n"; }
          items.push(...parseTextSchedules(content, state.selectedMonth, "PDF"));
        } else if (file.type.startsWith("image/")) {
          const { createWorker } = await import("tesseract.js");
          const worker = await createWorker("jpn", 1, { langPath: "/tessdata" });
          const result = await worker.recognize(file); await worker.terminate();
          items.push(...parseTextSchedules(result.data.text, state.selectedMonth, "画像OCR"));
        } else if (file.name.toLowerCase().endsWith(".ics")) items.push(...parseIcsSchedules(await file.text()));
        else items.push(...parseTextSchedules(await file.text(), state.selectedMonth, file.name.endsWith(".csv") ? "CSV" : "テキスト"));
      }
      addParsed(items);
    } catch { setNotice("読み取りに失敗しました。画像・PDFは送信されていません。手入力またはテキスト貼り付けを利用できます。"); }
    setProcessing("");
  }
  function update(id: string, patch: Partial<ScheduleItem>) { mutate((d: AppState) => { Object.assign(d.schedules.find((x) => x.id === id)!, patch); return d; }); }
  return <section className="panel"><div className="panel-heading"><div><span className="eyebrow">端末内で読み取り</span><h2>予定を取り込む</h2><p>スクリーンショットやPDFは外部AIへ送信せず、このブラウザ内で解析します。</p></div></div>
    <div className="import-grid"><label className="dropzone"><input type="file" multiple accept="image/*,.pdf,.csv,.txt,.ics" onChange={(e) => files(e.target.files)} /><b>画像・PDF・CSVを選択</b><span>複数画像・iPhoneカレンダー書き出しにも対応</span></label><div className="paste-box"><textarea aria-label="予定テキスト" placeholder={'例：7/15\t10:00-11:00\t学校訪問\t浦和高校'} value={text} onChange={(e) => setText(e.target.value)} /><div className="button-row"><button className="secondary" onClick={() => addParsed([{ id: uid(), date: `${state.selectedMonth}-01`, startTime: "09:00", endTime: "10:00", title: "", location: "", isBusiness: true, hasTravel: true, confirmed: false, source: "手入力" }])}>手入力で予定を追加</button><button className="primary" onClick={() => { addParsed(parseTextSchedules(text, state.selectedMonth, text.includes(",") ? "CSV" : "テキスト")); setText(""); }}>テキストから候補作成</button></div></div></div>
    {processing && <div className="processing">{processing}</div>}
    <h3>OCR・取込結果の確認</h3><p className="section-note">自動確定はしません。日付・時刻・予定名・場所・業務区分・移動有無を確認してください。</p>
    <div className="schedule-table"><div className="schedule-head"><span>日時</span><span>予定名・場所</span><span>区分</span><span>確認</span></div>{state.schedules.filter((s: ScheduleItem) => s.date.startsWith(state.selectedMonth)).map((item: ScheduleItem) => <div className="schedule-row" key={item.id}><div><input type="date" value={item.date} onChange={(e) => update(item.id, { date: e.target.value })} /><div className="inline"><input type="time" value={item.startTime} onChange={(e) => update(item.id, { startTime: e.target.value })} /><span>–</span><input type="time" value={item.endTime} onChange={(e) => update(item.id, { endTime: e.target.value })} /></div></div><div><input placeholder="予定名" value={item.title} onChange={(e) => update(item.id, { title: e.target.value })} /><input placeholder="場所" value={item.location} onChange={(e) => update(item.id, { location: e.target.value })} /></div><div><label><input type="checkbox" checked={item.isBusiness} onChange={(e) => update(item.id, { isBusiness: e.target.checked })} />業務</label><label><input type="checkbox" checked={item.hasTravel} onChange={(e) => update(item.id, { hasTravel: e.target.checked })} />移動あり</label></div><div><button className={item.confirmed ? "confirmed" : "secondary"} onClick={() => update(item.id, { confirmed: !item.confirmed })}>{item.confirmed ? "確認済み" : "確認する"}</button><small>{item.source}</small></div></div>)}</div>
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
