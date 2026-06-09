// netlify/functions/evidence-list.js
// Wordアドインから訴状本文を受け取り、各証拠（甲号証）について
// 証拠番号・標目・原本写し・作成日・作成名義人・立証趣旨を仮作成してJSONで返す。
//
// 入力: { complaint: "訴状の本文テキスト", side: "甲" | "乙" }
// 出力: { evidences: [{number, title, originalOrCopy, date, author, purpose, note}], notes: [...] }

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
      return new Response(JSON.stringify({ error: "訴状の本文が空です" }), { status: 400, headers: corsHeaders });
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return new Response(JSON.stringify({ error: "サーバー側でAPIキーが設定されていません" }), { status: 500, headers: corsHeaders });

    const MAX = 40000;
    const target = complaint.slice(0, MAX);

    const systemPrompt =
"あなたは日本の民事訴訟実務に精通した法律事務所の事務職員を補助するアシスタントです。" +
"与えられた訴状の本文を読み、" + side + "号証の証拠説明書の下書きを作成します。\n\n" +
"【抽出・作成のルール】\n" +
"1. 訴状中で引用されている証拠（『" + side + "第1号証』『" + side + "1』『（" + side + "1）』等の表記）を、" +
"番号順にすべて拾い出す。番号は『" + side + "1』『" + side + "2』…の形式に統一する（枝番は『" + side + "1の1』）。\n" +
"2. 各証拠について、訴状の文脈から次を可能な範囲で推定する：\n" +
"   - title（標目）：書証の名称（例：不動産登記事項証明書、賃貸借契約書、内容証明郵便 等）。\n" +
"   - originalOrCopy（原本／写し）：訴状から判断できなければ『写し』を仮置きする。\n" +
"   - date（作成日）：文書の作成年月日。訴状に明示があれば記載、なければ空欄。\n" +
"   - author（作成名義人）：その文書を作成した者（例：法務局、契約当事者双方、通知人 等）。\n" +
"   - purpose（立証趣旨）：その証拠で何を立証しようとしているかを簡潔に記す。\n" +
"3. 訴状本文だけでは確定できない項目（標目の正確な名称、作成者、作成日、原本写しの別）は、" +
"推定値を入れたうえで note に『要確認』と明記する。推測が全くできない項目は空文字にし note に理由を書く。\n" +
"4. 立証趣旨は訴状の主張と証拠の対応関係から作成してよいが、事実を創作しないこと。\n" +
"5. 標目・作成者等は実際の証拠物を見て確定すべきものであり、本下書きはあくまで仮であることを前提とする。\n\n" +
"【重要】証拠として引用されていないものを創作してはならない。訴状に出てくる証拠だけを対象とする。\n\n" +
"【出力形式】必ず次のJSONのみ。説明やコードフェンスは付けない。\n" +
'{"evidences":[{"number":"' + side + '1","title":"","originalOrCopy":"写し","date":"","author":"","purpose":"","note":""}],"notes":["全体的な注意があれば"]}';

    const userContent =
"次の訴状本文から、" + side + "号証の証拠説明書の下書きをJSONで作成してください。\n\n――― 訴状本文 ここから ―――\n" +
target + "\n――― 訴状本文 ここまで ―――";

    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 6000,
        temperature: 0,
        system: systemPrompt,
        messages: [{ role: "user", content: userContent }],
      }),
    });

    if (!anthropicRes.ok) {
      const detail = await anthropicRes.text();
      return new Response(JSON.stringify({ error: "解析APIの呼び出しに失敗しました", detail }), { status: 502, headers: corsHeaders });
    }

    const data = await anthropicRes.json();
    const raw = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
    const cleaned = raw.replace(/```json/g, "").replace(/```/g, "").trim();

    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      return new Response(JSON.stringify({ error: "解析結果の読み取りに失敗しました", raw }), { status: 502, headers: corsHeaders });
    }

    return new Response(
      JSON.stringify({
        evidences: Array.isArray(parsed.evidences) ? parsed.evidences : [],
        notes: Array.isArray(parsed.notes) ? parsed.notes : [],
      }),
      { status: 200, headers: corsHeaders }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
