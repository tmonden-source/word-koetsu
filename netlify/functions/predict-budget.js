// netlify/functions/predict-budget.js
// Excelアドインから「実績家計表データ」と「予測家計表ひな形の費目リスト」を受け取り、
// ひな形の各費目に当てはめる金額（と要確認注記）をClaude APIで算出して返す。
//
// 入力: {
//   actuals: [{month, income:[{item,amount}], expense:[{item,amount}]}...],
//   templateIncomeItems: ["給与(申立人)", ...],   // ひな形のA列費目
//   templateExpenseItems: ["住居費...", "食費", ...], // ひな形のC列費目
//   repayment: 33703
// }
// 出力: {
//   incomeAmounts: { "給与(申立人)": 246900, ... },
//   expenseAmounts: { "食費": 120000, ... },
//   reviewItems: [ { item, where, note } ]
// }

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
    const { actuals, templateIncomeItems, templateExpenseItems, repayment } = await request.json();

    if (!actuals || !Array.isArray(actuals) || actuals.length === 0) {
      return new Response(JSON.stringify({ error: "実績家計表のデータがありません" }), { status: 400, headers: corsHeaders });
    }
    if (!Array.isArray(templateIncomeItems) || !Array.isArray(templateExpenseItems)) {
      return new Response(JSON.stringify({ error: "ひな形の費目リストがありません" }), { status: 400, headers: corsHeaders });
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return new Response(JSON.stringify({ error: "サーバー側でAPIキーが設定されていません" }), { status: 500, headers: corsHeaders });

    const repaymentNum = Number(repayment) || 0;

    const systemPrompt =
"あなたは日本の小規模個人再生事件を扱う法律事務所の事務職員を補助するアシスタントです。" +
"再生債務者の申立前の家計実績（複数月）をもとに、福岡地裁の法定様式である予測家計表（ひな形）の" +
"各費目に当てはめる金額を1か月分（ボーナス月を除く通常月）算出します。\n\n" +
"【最重要・様式の制約】\n" +
"- 出力できる費目は、後述する『ひな形の費目リスト』に載っているものだけです。新しい費目名を作ってはいけません。\n" +
"- リストの費目名は一字一句そのまま使ってください（金額の対応キーになります）。\n\n" +
"【金額算出の方針】\n" +
"- 各費目は実績の傾向（平均・季節変動）から認可後の通常月として妥当な概算額を出す。1円単位は不要。\n" +
"- 申立てに伴う弁護士費用など、認可後に発生しない一時的支出は計上しない。\n" +
"- 実績に無い費目には金額を割り当てない（空欄のまま）。ただし下記の計上漏れ補完を除く。\n\n" +
"【再生計画に基づく返済】\n" +
"- ひな形に『再生計画に基づく返済』費目があれば、そこに " + repaymentNum + " 円を割り当てる（0なら割り当てない）。\n\n" +
"【計上漏れの定期支出の補完（要確認として扱う）】\n" +
"実績の支出から所有・契約が推定されるのに、対応する定期支出費目がひな形にあって実績に無い場合、" +
"概算を割り当てたうえで必ず reviewItems に記録する：\n" +
"- ガソリン代/駐車場代がある→自動車所有→『自動車税』『自動車保険』等の費目があれば概算を入れる\n" +
"- 住宅ローンがある→持ち家→『固定資産税』『火災保険』等の費目があれば概算を入れる\n" +
"これらは expenseAmounts に概算を入れつつ、reviewItems に {item:費目名, where:'expense', note:'要確認：実額を確認してください'} を必ず追加する。\n" +
"推定根拠があるが金額が読めない場合は金額を入れず reviewItems にのみ記録する。\n\n" +
"【出力形式】必ず次のJSONのみ。説明やコードフェンスは付けない。\n" +
'{"incomeAmounts":{"費目名":数値},"expenseAmounts":{"費目名":数値},"reviewItems":[{"item":"費目名","where":"income|expense","note":"要確認：..."}]}\n' +
"金額を割り当てない費目はキー自体を含めない。";

    const userContent =
"【ひな形の収入費目リスト】（この中の費目名のみ使用可）\n" +
JSON.stringify(templateIncomeItems, null, 2) + "\n\n" +
"【ひな形の支出費目リスト】（この中の費目名のみ使用可）\n" +
JSON.stringify(templateExpenseItems, null, 2) + "\n\n" +
"【家計実績（複数月）】\n" +
JSON.stringify(actuals, null, 2) + "\n\n" +
"【再生計画に基づく月々の返済額】" + repaymentNum + " 円\n\n" +
"上記をもとに、ひな形の費目に当てはめる金額をJSONで出力してください。";

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
        incomeAmounts: parsed.incomeAmounts && typeof parsed.incomeAmounts === "object" ? parsed.incomeAmounts : {},
        expenseAmounts: parsed.expenseAmounts && typeof parsed.expenseAmounts === "object" ? parsed.expenseAmounts : {},
        reviewItems: Array.isArray(parsed.reviewItems) ? parsed.reviewItems : [],
      }),
      { status: 200, headers: corsHeaders }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
