// netlify/functions/predict-budget.js
// Excelアドインから実績家計表データを受け取り、Claude APIで
// 「認可後1か月分の予測家計表」を算出してJSONで返す中継サーバー。
//
// 入力: { actuals: [{month, income:[{item,amount}], expense:[{item,amount}]}...],
//         repayment: 33703 }   ← 再生計画に基づく返済額（パネルで手入力）
// 出力: { income:[{item,amount,note}], expense:[{item,amount,note}], notes:[...] }

export default async (request) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (request.method === "OPTIONS") {
    return new Response("", { status: 204, headers: corsHeaders });
  }
  if (request.method !== "POST") {
    return new Response(JSON.stringify({ error: "POSTのみ対応しています" }), { status: 405, headers: corsHeaders });
  }

  try {
    const { actuals, repayment } = await request.json();

    if (!actuals || !Array.isArray(actuals) || actuals.length === 0) {
      return new Response(JSON.stringify({ error: "実績家計表のデータがありません" }), { status: 400, headers: corsHeaders });
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return new Response(JSON.stringify({ error: "サーバー側でAPIキーが設定されていません" }), { status: 500, headers: corsHeaders });
    }

    const repaymentNum = Number(repayment) || 0;

    const systemPrompt =
`あなたは日本の小規模個人再生事件を扱う法律事務所の事務職員を補助するアシスタントです。
再生債務者の「申立前の家計実績（複数月）」をもとに、再生計画認可後の通常月（ボーナス月を除く）に
予測される家計収支表を1か月分作成します。以下のルールを厳格に守ってください。

【基本方針】
- 各収入・支出項目は、与えられた実績の傾向（平均、季節変動、明らかな一時的支出の除外）を踏まえ、
  認可後の通常月として妥当な概算額を算出する。
- 1円単位の細かさは不要。実態に即した概算でよい。
- 一時的・特殊な支出（例：申立てに伴う弁護士費用、引越し費用など、認可後は発生しない費目）は
  予測には計上しない。

【認可後に必ず反映する調整】
- 「再生計画に基づく返済」を支出に必ず1行加える。金額は ${repaymentNum} 円。
  （0 の場合も項目だけは「再生計画に基づく返済」を金額未記入で残す）

【支出項目から推定される、計上漏れの定期支出を補う（重要）】
実績の支出項目から、以下のような「所有・契約が推定されるのに、定期的な関連支出が
予測表に無い」ものを検出し、概算を仮入力したうえで必ず note に「要確認」と記載する。
- ガソリン代・駐車場代がある → 自動車を所有 → 「自動車税」「車検代（2年に1度を月割等）」
  「自動車保険」がなければ補う
- 住宅ローンがある → 持ち家 → 「固定資産税」「火災保険」がなければ補う
- ペット関連費がある → 「ペット保険」「予防接種代」などの定期費を確認
これらの補完項目は、金額を実勢相場の概算で仮入力しつつ、必ず note に「要確認：実額を確認してください」と書く。
金額が全く推定できないものは amount を null にし、note に「要確認」と書く。

【出力形式】
必ず次のJSONのみを出力。前後の説明やMarkdownコードフェンスは一切付けない。
{
  "income": [{"item":"費目名","amount":数値またはnull,"note":""}],
  "expense": [{"item":"費目名","amount":数値またはnull,"note":""}],
  "notes": ["全体に関する補足があれば短く"]
}
noteは通常は空文字。補完・要確認項目にのみ「要確認：…」を記載する。
費目名は実績家計表で使われている表記に合わせる。`;

    const userContent =
`以下が再生債務者の家計実績です（JSON）。これをもとに認可後の予測家計表を作成してください。

${JSON.stringify(actuals, null, 2)}

再生計画に基づく月々の返済額: ${repaymentNum} 円`;

    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 4000,
        system: systemPrompt,
        messages: [{ role: "user", content: userContent }],
      }),
    });

    if (!anthropicRes.ok) {
      const detail = await anthropicRes.text();
      return new Response(JSON.stringify({ error: "予測APIの呼び出しに失敗しました", detail }), { status: 502, headers: corsHeaders });
    }

    const data = await anthropicRes.json();
    const raw = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
    const cleaned = raw.replace(/```json/g, "").replace(/```/g, "").trim();

    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      return new Response(JSON.stringify({ error: "予測結果の解析に失敗しました", raw }), { status: 502, headers: corsHeaders });
    }

    return new Response(
      JSON.stringify({
        income: Array.isArray(parsed.income) ? parsed.income : [],
        expense: Array.isArray(parsed.expense) ? parsed.expense : [],
        notes: Array.isArray(parsed.notes) ? parsed.notes : [],
      }),
      { status: 200, headers: corsHeaders }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
