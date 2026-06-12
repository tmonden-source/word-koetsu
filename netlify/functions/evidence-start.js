// evidence-start.js — 証拠抽出ジョブを受け付け、すぐに受付番号(jobId)を返す。
// 実際の重い処理はバックグラウンド関数 evidence-process-background に任せる。

import { getStore } from "@netlify/blobs";

export default async (request) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (request.method === "OPTIONS") return new Response("", { status: 204, headers: corsHeaders });
  if (request.method !== "POST") return new Response(JSON.stringify({ error: "POSTのみ対応しています" }), { status: 405, headers: corsHeaders });

  try {
    const body = await request.json();
    const complaint = body.complaint;
    const side = body.side === "乙" ? "乙" : "甲";

    if (!complaint || typeof complaint !== "string" || complaint.trim() === "") {
      return new Response(JSON.stringify({ error: "本文が空です" }), { status: 400, headers: corsHeaders });
    }

    // 受付番号を発行
    const jobId = "job_" + Date.now() + "_" + Math.random().toString(36).slice(2, 10);

    // 結果置き場（Blobs）に「処理中」を記録
    const store = getStore("evidence-jobs");
    await store.setJSON(jobId, { status: "pending", createdAt: Date.now() });

    // バックグラウンド関数を起動（応答は待たない）。
    // request.url からオリジンを得るが、取得できない場合は既知のサイトURLを使う。
    let origin;
    try {
      origin = new URL(request.url).origin;
      if (!/^https?:\/\//.test(origin)) origin = "https://word-kouetsu.netlify.app";
    } catch {
      origin = "https://word-kouetsu.netlify.app";
    }
    const bgUrl = origin + "/.netlify/functions/evidence-process-background";

    // 待たずに投げる（fire-and-forget）。awaitしないことで即座に応答を返す。
    fetch(bgUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jobId, complaint, side }),
    }).catch(() => { /* 起動失敗時もjobId側のpendingが残るので、結果取得側でタイムアウト判定する */ });

    return new Response(JSON.stringify({ jobId }), { status: 202, headers: corsHeaders });
  } catch (err) {
    return new Response(JSON.stringify({ error: "受付処理でエラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
