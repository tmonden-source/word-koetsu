// taskpane.js — アドインの本体ロジック
//
// 流れ:
//  1. Wordの本文テキストを取得
//  2. Netlifyの中継関数(/proofread)へ送信し、修正リストを受け取る
//  3. 変更履歴(Track Changes)をオンにした状態で、各修正を検索・置換
//     → Wordに赤字の変更履歴として残る

// ▼▼▼ デプロイ後、自分のNetlifyサイトのURLに書き換えてください ▼▼▼
const PROOFREAD_ENDPOINT = "https://YOUR-SITE-NAME.netlify.app/.netlify/functions/proofread";
// ▲▲▲ 例: https://word-koetsu.netlify.app/.netlify/functions/proofread ▲▲▲

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").addEventListener("click", runProofread);
  }
});

const $status = () => document.getElementById("status");
const $count = () => document.getElementById("count");
const $results = () => document.getElementById("results");

function setStatus(msg, kind) {
  const el = $status();
  el.className = "status" + (kind ? " " + kind : "");
  el.innerHTML = msg;
}

async function runProofread() {
  const btn = document.getElementById("run");
  btn.disabled = true;
  $results().innerHTML = "";
  $count().className = "count";
  setStatus('<span class="spinner"></span>本文を読み取っています…');

  try {
    // --- 1. 本文テキストを取得 ---
    const bodyText = await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text;
    });

    if (!bodyText || bodyText.trim() === "") {
      setStatus("文書が空です。校閲対象のテキストがありません。", "err");
      btn.disabled = false;
      return;
    }

    // --- 2. 中継サーバーへ送信して校閲 ---
    setStatus('<span class="spinner"></span>校閲しています…（数十秒かかる場合があります）');

    const res = await fetch(PROOFREAD_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: bodyText }),
    });

    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error || `サーバーエラー (${res.status})`);
    }

    const { corrections } = await res.json();

    if (!corrections || corrections.length === 0) {
      setStatus("誤字脱字は見つかりませんでした。", "ok");
      btn.disabled = false;
      return;
    }

    // --- 3. 変更履歴をオンにして置換 ---
    setStatus('<span class="spinner"></span>変更履歴として反映しています…');
    const applied = await applyCorrections(corrections);

    // --- 結果を画面に表示 ---
    renderResults(corrections, applied);
    $count().className = "count show";
    $count().textContent = `${applied} 件を変更履歴として反映しました（全 ${corrections.length} 件中）`;
    setStatus("完了しました。Wordの「校閲」タブで内容を確認してください。", "ok");
  } catch (err) {
    setStatus("エラー: " + err.message, "err");
  } finally {
    btn.disabled = false;
  }
}

// 各修正を、変更履歴をオンにした状態で検索・置換する
async function applyCorrections(corrections) {
  return await Word.run(async (context) => {
    // 変更履歴を有効化（これ以降の編集がすべて履歴として記録される）
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    let appliedCount = 0;

    for (const c of corrections) {
      if (!c.before || c.before === c.after) continue;

      // 本文から該当文字列を検索
      const searchResults = context.document.body.search(c.before, {
        matchCase: true,
        ignorePunct: false,
      });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length === 0) continue;

      // 最初の1件のみ置換（同じ文字列の過剰な一括置換を避ける）
      searchResults.items[0].insertText(c.after, Word.InsertLocation.replace);
      appliedCount++;
      await context.sync();
    }

    return appliedCount;
  });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderResults(corrections, appliedCount) {
  const html = corrections
    .map(
      (c) => `
      <div class="item">
        <div class="pair">
          <span class="before">${escapeHtml(c.before)}</span>
          <span class="arrow">▶</span>
          <span class="after">${escapeHtml(c.after)}</span>
        </div>
        ${c.reason ? `<div class="reason">${escapeHtml(c.reason)}</div>` : ""}
      </div>`
    )
    .join("");
  $results().innerHTML = html;
}
