import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("旅費申請アプリを日本語メタデータ付きで描画する", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<html lang="ja">/);
  assert.match(html, /<title>出張旅費申請書作成アプリ<\/title>/);
  assert.match(html, /端末内の旅費データを準備しています/);
  assert.match(html, /manifest\.webmanifest/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("公開物にPWA定義と端末内OCRモデルを同梱する", async () => {
  const [manifest, page, app, styles] = await Promise.all([
    readFile(new URL("../public/manifest.webmanifest", import.meta.url), "utf8"),
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/travel-expense-app.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(manifest, /出張旅費申請書作成アプリ/);
  assert.match(page, /TravelExpenseApp/);
  assert.match(app, /IndexedDB|端末内/);
  assert.match(app, /日付と行き先から自動で作成/);
  assert.match(app, /ブラウザ内の運賃台帳/);
  assert.match(app, /再計算/);
  assert.match(app, /確定して登録/);
  assert.match(app, /簡単入力のIC料金/);
  assert.match(app, /初回は入力してください/);
  assert.match(app, /登録状況/);
  assert.match(app, /過去の出張旅費申請書を読み込む/);
  assert.match(app, /今月の旅費を入力/);
  assert.match(app, /実績マスター管理/);
  assert.match(app, /ODS申請書を作成/);
  assert.match(app, /日付を選び、過去実績から確定/);
  assert.match(app, /この日を確定して次へ/);
  assert.doesNotMatch(app, /fetch\("\/api\/fare"/);
  assert.match(app, /スクリーンショットを貼り付け/);
  assert.match(app, /試験機能：画像・PDFから予定を読み取る/);
  assert.match(app, /createPortal/);
  assert.match(app, /position: "fixed"/);
  assert.match(app, /onPointerDown/);
  assert.match(app, /layer\.scrollTop/);
  assert.match(app, /data-settings-enter/);
  assert.match(app, /nativeEvent\.isComposing/);
  assert.match(app, /Math\.min\(current \+ 1, candidates\.length - 1\)/);
  assert.match(styles, /\.entry-table[^}]*overflow:auto/);
  await access(new URL("../public/tessdata/jpn.traineddata.gz", import.meta.url));
  await access(new URL("../public/2026年度版出張旅費代精算書原本.ods", import.meta.url));
});
