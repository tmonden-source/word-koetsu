// evidence-process-background.js — バックグラウンド関数（最大15分）。
// ファイル名が -background で終わると、Netlifyは非同期のバックグラウンド関数として扱う。
// 時間制限が長いので、Sonnetでじっくり処理しても途中で打ち切られない。

import { getStore } from "@netlify/blobs";
import { runEvidenceExtraction } from "./_evidence-core.js";

export default async (request) => {
  let jobId = null;
  const store = getStore("evidence-jobs");
  try {
    const body = await request.json();
    jobId = body.jobId;
    const complaint = body.complaint;
    const side = body.side === "乙" ? "乙" : "甲";

    if (!jobId) return new Response("no jobId", { status: 400 });

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      await store.setJSON(jobId, { status: "error", error: "サーバー側でAPIキーが設定されていません", finishedAt: Date.now() });
      return new Response("no api key", { status: 200 });
    }

    // 重い処理（Sonnetで高精度抽出）
    const result = await runEvidenceExtraction(complaint, side, apiKey);

    // 結果を保存
    await store.setJSON(jobId, { status: "done", result, finishedAt: Date.now() });
    return new Response("done", { status: 200 });
  } catch (err) {
    try {
      if (jobId) await store.setJSON(jobId, { status: "error", error: String(err && err.message ? err.message : err), finishedAt: Date.now() });
    } catch (_) { /* 保存失敗は無視 */ }
    return new Response("error", { status: 200 });
  }
};
