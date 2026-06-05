// netlify/functions/predict-budget.js
// Excelアドインから「実績家計表データ」と「予測家計表ひな形の費目リスト」を受け取り、
// (1) ひな形の既存費目に当てはめる金額、(2) ひな形に無い実績費目の追記分、をClaude APIで算出して返す。
//
// 入力: {
//   actuals: [{month, income:[{item,amount}], expense:[{item,amount}]}...],
//   templateIncomeItems: [...], templateExpenseItems: [...],
//   incomeBlankCount, expenseBlankCount,   // 追記できる空白行の数
//   repayment
// }
// 出力: {
//   incomeAmounts: { 費目名: 金額 },         // 既存費目への割当
//   expenseAmounts: { 費目名: 金額 },
//   incomeAdditions: [ {item, amount, note} ], // ひな形に無い→空白行へ追記
//   expenseAdditions: [ {item, amount, note} ],
//   reviewItems: [ {item, where, note} ]
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
    const body = await request.json();
    const { actuals, templateIncomeItems, templateExpenseItems, repayment } = body;
    const incomeBlankCount = Number(body.incomeBlankCount) || 0;
    const expenseBlankCount = Number(body.expenseBlankCount) || 0;

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
"再生債務者の申立前の家計実績（複数月）をもとに、福岡地裁の法定様式である予測家計表（ひな形）を" +
"1か月分（ボーナス月を除く通常月）作成します。\n\n" +
"【出力は2系統に分ける】\n" +
"(A) ひな形に既にある費目への金額割当 … incomeAmounts / expenseAmounts\n" +
"(B) ひな形に無いが実績にあった費目の追記 … incomeAdditions / expenseAdditions\n\n" +
"【(A) 既存費目への割当】\n" +
"- キーは『ひな形の費目リスト』の費目名を一字一句そのまま使う（新費目を作らない）。\n" +
"- 実績の傾向（平均・季節変動）から認可後の通常月として妥当な概算額を出す。1円単位は不要。\n" +
"- 実績に無い費目には割り当てない（キーを含めない）。ただし下記の計上漏れ補完を除く。\n" +
"- ひな形に『再生計画に基づく返済』があれば " + repaymentNum + " 円を割り当てる（0なら割り当てない）。\n\n" +
"【(B) ひな形に無い実績費目の追記（重要）】\n" +
"- 実績の支出費目のうち、ひな形の支出費目リストに相当する費目が無いものは、認可後も継続する費目であれば" +
" expenseAdditions に {item:実績の費目名, amount:概算, note:''} として出す。\n" +
"- 同様に収入側は incomeAdditions に出す。\n" +
"- ただし申立てに伴う弁護士費用など認可後に発生しない一時的費目は追記しない（除外する）。\n" +
"- 追記できる空白行数の上限：収入 " + incomeBlankCount + " 行、支出 " + expenseBlankCount + " 行。" +
"重要なものから順に、この行数を超えないように厳選する。超える場合は金額の大きい定期支出を優先。\n\n" +
"【計上漏れの定期支出の補完（要確認）】\n" +
"実績から所有・契約が推定されるのに対応する定期支出が無い場合、概算を入れて reviewItems に記録する：\n" +
"- ガソリン代/駐車場代→自動車→『自動車税』『自動車保険』等。ひな形に費目があれば expenseAmounts、無ければ expenseAdditions に入れる。\n" +
"- 住宅ローン→持ち家→『固定資産税』『火災保険』等。同上。\n" +
"これらは必ず note に『要確認：実額を確認してください』と書き、reviewItems にも {item,where,note} を追加する。\n\n" +
"【出力形式】必ず次のJSONのみ。説明やコードフェンスは付けない。\n" +
'{"incomeAmounts":{"費目名":数値},"expenseAmounts":{"費目名":数値},' +
'"incomeAdditions":[{"item":"費目名","amount":数値,"note":""}],' +
'"expenseAdditions":[{"item":"費目名","amount":数値,"note":""}],' +
'"reviewItems":[{"item":"費目名","where":"income|expense","note":"要確認：..."}]}';

    const userContent =
"【ひな形の収入費目リスト】（(A)で使う費目名）\n" + JSON.stringify(templateIncomeItems, null, 2) + "\n\n" +
"【ひな形の支出費目リスト】（(A)で使う費目名）\n" + JSON.stringify(templateExpenseItems, null, 2) + "\n\n" +
"【追記できる空白行数】収入 " + incomeBlankCount + " 行 / 支出 " + expenseBlankCount + " 行\n\n" +
"【家計実績（複数月）】\n" + JSON.stringify(actuals, null, 2) + "\n\n" +
"【再生計画に基づく月々の返済額】" + repaymentNum + " 円\n\n" +
"上記をもとにJSONを出力してください。実績にあってひな形に無い費目は追記分(B)に回してください。";

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

    const arr = (x) => (Array.isArray(x) ? x : []);
    const obj = (x) => (x && typeof x === "object" ? x : {});

    return new Response(
      JSON.stringify({
        incomeAmounts: obj(parsed.incomeAmounts),
        expenseAmounts: obj(parsed.expenseAmounts),
        incomeAdditions: arr(parsed.incomeAdditions),
        expenseAdditions: arr(parsed.expenseAdditions),
        reviewItems: arr(parsed.reviewItems),
      }),
      { status: 200, headers: corsHeaders }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
