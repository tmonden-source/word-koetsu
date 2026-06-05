// netlify/functions/proofread.js

// ===== 金額表記の機械的正規化（裁判文書ルール：○○万○○○○円・全角・万以上は単位付き／,なし）=====
function _toFull(s) {
  return String(s).replace(/[0-9]/g, (d) => String.fromCharCode(d.charCodeAt(0) + 0xFEE0));
}
function _toHalf(s) {
  return s.replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xFEE0));
}
// 数値→「○億○○○○万○○○○円」全角。万・億の後ろは4桁ゼロ詰め。1万未満はそのまま全角＋円。
function _formatYen(n) {
  if (n < 10000) return _toFull(n) + "円";
  const oku = Math.floor(n / 100000000);
  const man = Math.floor((n % 100000000) / 10000);
  const rest = n % 10000;
  let out = "";
  if (oku > 0) {
    out += _toFull(oku) + "億";
    out += _toFull(String(man).padStart(4, "0")) + "万";
    out += _toFull(String(rest).padStart(4, "0"));
  } else {
    out += _toFull(man) + "万";
    out += _toFull(String(rest).padStart(4, "0"));
  }
  return out + "円";
}
// テキストから金額表現を拾い、裁判文書形式と異なるものを {before, after, reason} で返す
function detectMoneyIssues(text) {
  const issues = [];
  const seen = new Set();
  const re = /([0-9０-９][0-9０-９,，億万千百十]*円)/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const orig = m[1];
    if (seen.has(orig)) continue;
    const half = _toHalf(orig).replace(/[,，]/g, "");
    const mm = half.match(/^(?:(\d+)億)?(?:(\d+)万)?(\d+)?円$/);
    if (!mm || (!mm[1] && !mm[2] && !mm[3])) continue;
    const value =
      (mm[1] ? parseInt(mm[1], 10) * 100000000 : 0) +
      (mm[2] ? parseInt(mm[2], 10) * 10000 : 0) +
      (mm[3] ? parseInt(mm[3], 10) : 0);
    if (isNaN(value)) continue;
    const correct = _formatYen(value);
    if (orig !== correct) {
      issues.push({ before: orig, after: correct, reason: "裁判文書の金額表記（○○万○○○○円・全角）" });
      seen.add(orig);
    }
  }
  return issues;
}

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
      "明らかな文法ミスを検出し、修正してください。文体や言い回しの好みによる書き換えは行わないでください。\n" +
      "【送り仮名・表記は公用文の規則に従う】\n" +
      "送り仮名は内閣告示『送り仮名の付け方』の本則に従って統一する。特に次の公用文ルールを適用する：\n" +
      "1) 接続詞・副詞・接頭辞の一部は仮名で書く。例：及び→および、又は→または、並びに→ならびに、" +
      "若しくは→もしくは、飽くまで→あくまで、余り→あまり、既に→すでに、直ちに→ただちに、" +
      "正に→まさに、御指導→ご指導、御参加→ご参加。\n" +
      "2) 活用語は本則どおり送る。例：行なう→行う、表わす→表す、現われる→現れる、明か→明らか、" +
      "短かい→短い、少い→少ない、新らしい→新しい、危い→危ない、起る→起こる、断わる→断る、" +
      "捕える→捕らえる、聞える→聞こえる。\n" +
      "3) 複合の語で活用がなく読み間違えるおそれのない語は、通則6の許容により送り仮名を省く。" +
      "例：申込み→申込、取扱い→取扱、引渡し→引渡、明渡し→明渡、預り金→預り金、言渡し→言渡、" +
      "入替え→入替、打合せ→打合せ（公用文では『打合せ』）。" +
      "ただし読み誤るおそれがある語は省略しない。\n" +
      "4) 常用漢字表にない漢字・音訓で書かれた語は仮名書きにする。\n" +
      "【裁判文書の表記ルールも適用する】\n" +
      "5) 句読点は、読点に『、』（テン）、句点に『。』（マル）を用いる。読点に『，』（コンマ）、" +
      "句点に『．』（ピリオド）が使われていれば『、』『。』に修正する（令和4年4月以降、裁判文書の読点はコンマからテンに改められた）。\n" +
      "6) 金額は『○○万○○○○円』の形式で書く。すなわち『万』の位の後ろは、千・百・十・一の位を" +
      "必ず4桁にそろえ、空位は0で埋める。三桁区切りのコンマ表記（例：5,470,320円）は用いない。" +
      "また『万』の次にさらに『千』などの位取りの語を入れない。" +
      "正しい例：『547万0320円』（誤り：547万320円、547万5320円、5,470,320円、5,475,320円）。" +
      "『100万0000円』のように端数が0でも4桁にそろえる。数字は全角を用いる。\n" +
      "7) 金額以外の数字も原則として算用数字を用いる。『1名』『2名』のように書く。\n" +
      "8) 日時は『午後3時15分』のように書く。\n" +
      "9) 法令の条文番号は、条に『第』を付けず『民法94条2項』のように書く。" +
      "ただし枝番号が付く場合は『民法94条の2第2項』のように『第』を用いる。\n" +
      "10) 都道府県・政令指定都市など裁判所所在地が明らかな場合の都道府県名は文脈に応じて省く（過剰補充はしない）。\n" +
      "【主語・述語・目的語の明確化】\n" +
      "11) 日本語の主語省略により『誰が・誰に・何を・どうした』が一意に読み取れない箇所は、" +
      "文脈から確実に補える場合に限り主体・対象を補う。推測の域を出ない補充は行わず、" +
      "補った場合は reason に『主語（又は目的語）の補充』と記す。事実関係を創作・変更してはならない。\n" +
      "上記に反する表記・不明確な箇所があれば修正対象とし、reason に該当ルール名（『公用文の送り仮名規則』" +
      "『裁判文書の句読点』『法令条文の表記』『主語の補充』等）を簡潔に記す。\n" +
      "なお、文章全体の段落再構成や大幅な書き換えは行わず、語句・文単位の表記是正と必要最小限の主述補充にとどめる。\n" +
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
        model: "claude-sonnet-4-6",
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

    // 金額表記は機械的に検出して確実に追加（LLMの見落としを防ぐ）。
    // 既にLLMが同じbeforeを挙げていれば重複させない。
    const moneyIssues = detectMoneyIssues(target);
    const existingBefores = new Set(corrections.map((c) => c && c.before));
    for (const mi of moneyIssues) {
      if (!existingBefores.has(mi.before)) {
        corrections.push(mi);
        existingBefores.add(mi.before);
      }
    }

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
