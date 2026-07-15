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
  const [manifest, page, app] = await Promise.all([
    readFile(new URL("../public/manifest.webmanifest", import.meta.url), "utf8"),
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/travel-expense-app.tsx", import.meta.url), "utf8"),
  ]);
  assert.match(manifest, /出張旅費申請書作成アプリ/);
  assert.match(page, /TravelExpenseApp/);
  assert.match(app, /IndexedDB|端末内/);
  await access(new URL("../public/tessdata/jpn.traineddata.gz", import.meta.url));
});
