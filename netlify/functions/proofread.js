// netlify/functions/proofread.js
// Wordアドインから本文を受け取り、Claude APIで日本語校閲を行い、
// 「修正前→修正後」の差分リストをJSONで返す中継サーバー。
// APIキーはNetlifyの環境変数 ANTHROPIC_API_KEY から読み込む（クライアントには絶対に渡さない）。

export default async (request) => {
  // --- CORS（アドインの実行元から呼べるように許可）---
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  // ブラウザのプリフライト（OPTIONS）に応答
  if (request.method === "OPTIONS") {
    return new Response("", { status: 204, headers: corsHeaders });
  }

  if (request.method !== "POST") {
    return new Response(JSON.stringify({ error: "POSTのみ対応しています" }), {
      status: 405,
      headers: corsHeaders,
    });
  }

  try {
    const { text } = await request.json();

    if (!text || typeof text !== "string" || text.trim() === "") {
      return new Response(JSON.stringify({ error: "校閲対象のテキストが空です" }), {
        status: 400,
        headers: corsHeaders,
      });
    }

    // 長すぎる入力を防ぐ簡易ガード（必要に応じて調整）
    const MAX_CHARS = 20000;
    const target = text.slice(0, MAX_CHARS);

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return new Response(
        JSON.stringify({ error: "サーバー側でAPIキーが設定されていません" }),
        { status: 500, headers: corsHeaders }
      );
    }

    // Claudeに「誤字脱字の修正箇所だけ」を厳密なJSONで返させる
    const systemPrompt =
      "あなたは日本語の校正者です。渡された文章の誤字・脱字・誤変換・送り仮名の誤り・" +
      "明らかな文法ミスのみを検出し、修正してください。文体や言い回しの好みによる書き換えは行わないでください。" +
      "出力は必ず次のJSON形式のみとし、前後に説明やMarkdownのコードフェンスを一切付けないでください。\n" +
      '{"corrections":[{"before":"修正前の文字列","after":"修正後の文字列","reason":"簡潔な理由"}]}\n' +
      "beforeは元の文章中に実際に存在する文字列を、置換できる十分な長さ（できれば文節〜短文単位）で抜き出してください。" +
      "重複しない一意な文字列にすること。修正が無ければ corrections は空配列にしてください。";

    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 4000,
        system: systemPrompt,
        messages: [{ role: "user", content: target }],
      }),
    });

    if (!anthropicRes.ok) {
      const detail = await anthropicRes.text();
      return new Response(
        JSON.stringify({ error: "校閲APIの呼び出しに失敗しました", detail }),
        { status: 502, headers: corsHeaders }
      );
    }

    const data = await anthropicRes.json();

    // テキストブロックを結合
    const raw = (data.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("");

    // 念のためコードフェンスを除去してからパース
    const cleaned = raw.replace(/```json/g, "").replace(/```/g, "").trim();

    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      return new Response(
        JSON.stringify({ error: "校閲結果の解析に失敗しました", raw }),
        { status: 502, headers: corsHeaders }
      );
    }

    const corrections = Array.isArray(parsed.corrections) ? parsed.corrections : [];

    return new Response(JSON.stringify({ corrections }), {
      status: 200,
      headers: corsHeaders,
    });
  } catch (err) {
    return new Response(
      JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }),
      { status: 500, headers: corsHeaders }
    );
  }
};
