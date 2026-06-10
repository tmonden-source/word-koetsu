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
"与えられた書面（訴状・準備書面等）の本文を読み、証拠説明書の下書きを作成します。\n\n" +
"【抽出・作成のルール】\n" +
"1. 本文中で実際に引用されている証拠を、その表記に従って拾い出す。" +
"利用者は『" + side + "号証』を想定しているが、本文に『甲』『乙』のいずれが実際に出てくるかを優先する。" +
"本文に乙号証しか出てこなければ乙号証として、甲号証しか出てこなければ甲号証として抽出する。" +
"『甲第1号証』『甲1』『（乙1）』『乙１』等の表記を番号順にすべて拾い、番号は『甲1』『乙2』等の形式に統一する（枝番は『甲1の1』）。\n" +
"2. 各証拠について、本文の文脈から次を可能な範囲で推定する：\n" +
"   - title（標目）：書証の名称（例：不動産登記事項証明書、LINEのトーク履歴、陳述書、内容証明郵便 等）。\n" +
"   - originalOrCopy（原本／写し）：本文から判断できなければ『写し』を仮置きする。\n" +
"   - date（作成日）：文書の作成年月日。本文に明示があれば記載、なければ空欄。\n" +
"   - author（作成名義人）：その文書を作成した者。\n" +
"   - purpose（立証趣旨）：その証拠で立証しようとする事実を簡潔に記す。" +
"文末は必ず『〜事実。』の形で統一する（体言止め＋句点）。" +
"例：『原告が被告に対し暴言を繰り返した事実。』。" +
"『〜を立証する』『〜したこと』などの語尾は使わず、必ず『〜事実。』で終える。" +
"立証趣旨は1文・60字程度までを目安に簡潔にまとめる。\n" +
"3. 訴状本文だけでは確定できない項目（標目の正確な名称、作成者、作成日、原本写しの別）は、" +
"推定値を入れたうえで note に『要確認』と明記する。推測が全くできない項目は空文字にし note に理由を書く。\n" +
"4. 立証趣旨は訴状の主張と証拠の対応関係から作成してよいが、事実を創作しないこと。\n" +
"5. 標目・作成者等は実際の証拠物を見て確定すべきものであり、本下書きはあくまで仮であることを前提とする。\n\n" +
"【重要】証拠として引用されていないものを創作してはならない。訴状に出てくる証拠だけを対象とする。\n\n" +
"【事件情報の抽出】訴状の冒頭・末尾等から、次を抽出する：\n" +
"・caseName … 事件名（事件番号を含む全体。例：『令和5年（ワ）第1326号　建物明渡請求事件』）。\n" +
"・court … 裁判所名（提出先。例：『福岡地方裁判所第1民事部2係』。『御中』は付けない）。\n" +
"・plaintiff … 原告名、defendant … 被告名。\n" +
"・attorneys … 当方（" + side + "号証を提出する側。甲＝原告側、乙＝被告側）の訴訟代理人弁護士の氏名（複数可）。\n" +
"いずれも見つからない項目は空にする。創作しない。\n\n" +
"【出力形式】必ず次のJSONのみ。説明やコードフェンスは付けない。\n" +
'{"caseName":"事件名（事件番号含む。無ければ空）","court":"裁判所名（御中なし。無ければ空）",' +
'"plaintiff":"原告名（無ければ空）","defendant":"被告名（無ければ空）",' +
'"attorneys":["訴訟代理人弁護士の氏名（複数可。無ければ空配列）"],' +
'"side":"実際に本文に出てきた区分。甲号証なら甲、乙号証なら乙",' +
'"evidences":[{"number":"甲1または乙1（実際の区分）","title":"","originalOrCopy":"写し","date":"","author":"","purpose":"","note":""}],"notes":["全体的な注意があれば"]}';

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
        model: "claude-haiku-4-5",
        max_tokens: 4000,
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

    // 立証趣旨の文末を「〜事実。」に統一する保険処理
    function normalizePurpose(p) {
      if (typeof p !== "string") return p;
      var t = p.trim();
      if (t === "") return t;
      // 末尾の句点・空白を一旦除去
      t = t.replace(/[。\.\s　]+$/, "");
      // よくある語尾を「事実」に寄せる
      t = t.replace(/ことを立証する$/, "事実");
      t = t.replace(/を立証する$/, "事実");
      t = t.replace(/を立証$/, "事実");
      t = t.replace(/したこと$/, "した事実");
      t = t.replace(/であること$/, "である事実");
      t = t.replace(/こと$/, "事実");
      // 既に「事実」で終わっていなければ「事実」を補う
      if (!/事実$/.test(t)) t = t + "事実";
      return t + "。";
    }
    var evList = Array.isArray(parsed.evidences) ? parsed.evidences : [];
    evList.forEach(function (e) {
      if (e && typeof e === "object") e.purpose = normalizePurpose(e.purpose);
    });

    return new Response(
      JSON.stringify({
        side: (parsed.side === "甲" || parsed.side === "乙") ? parsed.side : side,
        caseName: typeof parsed.caseName === "string" ? parsed.caseName : "",
        court: typeof parsed.court === "string" ? parsed.court : "",
        plaintiff: typeof parsed.plaintiff === "string" ? parsed.plaintiff : "",
        defendant: typeof parsed.defendant === "string" ? parsed.defendant : "",
        attorneys: Array.isArray(parsed.attorneys) ? parsed.attorneys : [],
        evidences: evList,
        notes: Array.isArray(parsed.notes) ? parsed.notes : [],
      }),
      { status: 200, headers: corsHeaders }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: "サーバー内部エラー", detail: String(err) }), { status: 500, headers: corsHeaders });
  }
};
