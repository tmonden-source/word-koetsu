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
      "【送り仮名・表記は公用文の規則に従う（内閣告示・内閣訓令準拠）】\n" +
      "送り仮名は内閣告示『送り仮名の付け方』の本則・例外・通則に従う。次の公用文ルールを適用する：\n" +
      "1) 接続詞は原則として仮名で書く。例：おって、かつ（且つ→かつ）、したがって（従って→したがって）、" +
      "ただし（但し→ただし）、ついては、ところが、ところで、また（接続詞の又→また）、ゆえに。" +
      "ただし『及び・並びに・又は・若しくは』の4語は原則として漢字で書く。\n" +
      "2) 副詞・連体詞は原則として漢字で書く（公用文の原則）。例：余り、至って、必ず、必ずしも、" +
      "極めて、殊に、更に、既に、全て、直ちに、努めて、常に、特に、再び、全く、最も、明くる、大きな、来る、去る、我が。" +
      "ただし次の副詞は仮名で書く：かなり、ふと、やはり、よほど。" +
      "（注意：副詞・連体詞をむやみに仮名書きへ修正しないこと。漢字が原則である。）\n" +
      "3) 接頭辞『御』は、付く語を漢字で書くなら漢字（御挨拶・御意見）、仮名で書くならや常用漢字表にない漢字を含む語は仮名（ごもっとも、ごちそう）。" +
      "また接尾辞（…げ、…ども、…ぶる、…み、…め）は原則仮名で書く。\n" +
      "4) 助詞・助動詞は仮名で書く。例：ぐらい（位→ぐらい）、ほど（程→ほど）、など（等→など、『とう』と読む場合は『等』）、" +
      "だけ、ようだ（様だ→ようだ）。\n" +
      "5) 補助的な用法の動詞・形式名詞などは仮名で書く。" +
      "例：ある（在る/有る→ある）、いる（居る→いる）、できる（出来る→できる）、なる（成る→なる、『1万円になる』）、" +
      "こと（事→こと）、とき（時→とき、『事故のときは』）、ところ（所→ところ）、もの（物/者→もの）、" +
      "とおり（通り→とおり、『次のとおり』）、ため（為→ため）、ほか（外/他→ほか）、ゆえ（故→ゆえ）、わけ（訳→わけ）、" +
      "…ていく、…ていただく、…ておく、…てください、…てくる、…てしまう、…てみる、…について。" +
      "（ただし実際の動作・具体的対象を表す場合は漢字：『東から来る』『賞状を頂く』『家を建てる所』等。）\n" +
      "6) 活用語の送り仮名は本則どおり。例：行なう→行う、表わす→表す、現われる→現れる、" +
      "短かい→短い、少い→少ない、新らしい→新しい、危い→危ない、起る→起こる、断わる→断る、" +
      "捕える→捕らえる、聞える→聞こえる、明か→明らか。\n" +
      "7) 常用漢字表にない漢字・音訓の語は仮名書き。例：宜しく→よろしく、概ね→おおむね、予め→あらかじめ、" +
      "虞→おそれ、叶う→かなう、嬉しい→うれしい、有難う→ありがとう。\n" +
      "8) 活用のない複合の語のうち読み間違えるおそれのないものは、通則6の許容により送り仮名を省く（内閣訓令の186語）。" +
      "例：申込み→申込、取扱い→取扱、引渡し→引渡、明渡し→明渡、言渡し→言渡、打合せ、見積り→見積、" +
      "取締り→取締、立替え→立替、貸付け→貸付、差押え→差押、繰越し→繰越、取消し→取消、申立て→申立、" +
      "売上げ→売上、買受け→買受、引受け→引受、支払→支払、未払、内払。" +
      "ただし同じ漢字でも動詞は本則どおり送る（取り扱う、申し込む、引き渡す、打ち合わせる等）。\n" +
      "9) 慣用が固定した複合名詞（通則7）は送り仮名を付けない。" +
      "例：取締役、頭取、関取、書留、気付、切手、消印、小包、振替、踏切、請負、組合、手当、" +
      "売値、買値、両替、割引、子守、献立、座敷、試合、字引、場合、羽織、番組、番付、日付、" +
      "物語、役割、屋敷、夕立、割合、植木、置物、織物、貸家、敷地、建物、並木、受付、受取。" +
      "また付表の語で送り仮名を付けない語：息吹、桟敷、時雨、築山、名残、雪崩、吹雪、迷子、行方。\n" +
      "【外来語（片仮名）の表記は内閣告示『外来語の表記』に従う】\n" +
      "9-2) 長音は原則として長音符号『ー』で書く（エネルギー、オーバーコート、グループ、テーブル、パーティー）。" +
      "英語の語末 -er・-or・-ar 等はア列の長音とし『ー』を付ける（コンピューター、エレベーター、マフラー）。" +
      "『エネルギ』『コンピュータ』のように末尾の長音符号を脱落させた表記は、慣用上許容される場合もあるが、" +
      "公用文では長音符号を付ける形に統一する。\n" +
      "9-3) 撥音は『ン』、促音は小書きの『ッ』で書く（シャッター、リュックサック）。" +
      "拗音の『ャ・ュ・ョ』および『ァ・ィ・ェ・ォ』は小書きにする。\n" +
      "9-4) 国語として定着した外来語は第1表の仮名で書くのを基本とし、『ヴ』をむやみに使わず原則として" +
      "『バ・ビ・ブ・ベ・ボ』で書く（ヴァイオリン→バイオリン、ヴィーナス→ビーナス、ヴェール→ベール）。" +
      "ただし固有名詞（人名・地名・会社名・商品名）はこの限りでなく、原音に近い表記の慣用を尊重する。\n" +
      "9-5) 二重表記の語形のゆれ（ハンカチ/ハンケチ等）は一方に強制しないが、" +
      "一つの文書内では表記を統一する（同一語が混在していれば揃える）。\n" +
      "【裁判文書の表記ルールも適用する】\n" +
      "10) 句読点は、読点に『、』（テン）、句点に『。』（マル）を用いる。読点に『，』（コンマ）、" +
      "句点に『．』（ピリオド）が使われていれば『、』『。』に修正する（令和4年以降、裁判文書・公用文の読点はテンを基本とする）。\n" +
      "11) 金額は裁判文書の慣行に従い『○○万○○○○円』の形式で、全角数字を用い、コンマ（，）は付けない。" +
      "『万』『億』の後ろは千・百・十・一の位を必ず4桁にそろえ、空位は0で埋める。" +
      "正しい例：『５４７万０３２０円』（誤り：547万320円、547万5320円、5,470,320円、5,475,320円）。" +
      "『１００万００００円』のように端数が0でも4桁にそろえる。" +
      "（注：この金額形式はサーバー側でも自動的に検出・修正される。）\n" +
      "12) 金額以外の数字も原則として算用数字を用いる。『1名』『2名』のように書く。\n" +
      "13) 日時は『午後3時15分』のように書く。\n" +
      "14) 法令の条文番号は、条に『第』を付けず『民法94条2項』のように書く。" +
      "ただし枝番号が付く場合は『民法94条の2第2項』のように『第』を用いる。\n" +
      "15) 政令指定都市・地方裁判所本庁所在地の都市は道府県名の記載を省く（ただし東京都は常に記載）。過剰な補充はしない。\n" +
      "【主語・述語・目的語の明確化】\n" +
      "16) 日本語の主語省略により『誰が・誰に・何を・どうした』が一意に読み取れない箇所は、" +
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
