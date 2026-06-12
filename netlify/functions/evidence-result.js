// evidence-result.js — 受付番号(jobId)で処理状況・結果を返す。
// アドインが数秒おきにこれを呼んで完了を待つ（ポーリング）。

import { getStore } from "@netlify/blobs";

export default async (request) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (request.method === "OPTIONS") return new Response("", { status: 204, headers: corsHeaders });

  try {
    let jobId = null;
    if (request.method === "GET") {
      jobId = new URL(request.url).searchParams.get("jobId");
    } else {
      const body = await request.json();
      jobId = body.jobId;
    }
    if (!jobId) return new Response(JSON.stringify({ error: "jobIdがありません" }), { status: 400, headers: corsHeaders });

    const store = getStore("evidence-jobs");
    let rec;
    try {
      rec = await store.get(jobId, { type: "json" });
    } catch {
      rec = null;
    }

    if (!rec) {
      // まだ書き込まれていない、または期限切れ
      return new Response(JSON.stringify({ status: "pending" }), { status: 200, headers: corsHeaders });
    }

    if (rec.status === "done") {
      // 取得後に掃除（容量節約）。失敗しても問題ない。
      try { await store.delete(jobId); } catch (_) {}
      return new Response(JSON.stringify({ status: "done", result: rec.result }), { status: 200, headers: corsHeaders });
    }
    if (rec.status === "error") {
      try { await store.delete(jobId); } catch (_) {}
      return new Response(JSON.stringify({ status: "error", error: rec.error || "処理中にエラーが発生しました" }), { status: 200, headers: corsHeaders });
    }
    // pending
    return new Response(JSON.stringify({ status: "pending" }), { status: 200, headers: corsHeaders });
  } catch (err) {
    return new Response(JSON.stringify({ error: "結果取得でエラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
