// predict.js — 予測家計表アドインの本体ロジック
//
// 流れ:
//  1. ブックの全シート名を一覧表示（実績シートを選ばせる）
//  2. 選ばれたシートから 収入(費目・金額) / 支出(費目・金額) を読み取る
//  3. Netlifyの中継関数(/predict-budget)へ送り、予測結果を受け取る
//  4. 新しいシートに、提出済み予測家計表と同じ2列体裁で書き出す

const PREDICT_ENDPOINT = "https://word-kouetsu.netlify.app/.netlify/functions/predict-budget";

// 提出書式の固定文言
const TITLE_NAME_PREFIX = "再生債務者　";
const HEADER_TITLE = "再生計画認可後の予測家計表";
const NOTE_TEXT = "※　家計表の収支状況等を参考に、これまでの生活状況を見直した上で、再生計画認可後に予測される家計の収支状況を記載してください。金額は概算で結構です。";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadSheetList();
    document.getElementById("run").addEventListener("click", runPredict);
  }
});

const $status = () => document.getElementById("status");
const $results = () => document.getElementById("results");
function setStatus(msg, kind) {
  const el = $status();
  el.className = "status" + (kind ? " " + kind : "");
  el.innerHTML = msg;
}

// シート一覧を読み込んでチェックボックス表示
async function loadSheetList() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const box = document.getElementById("sheets");
      box.innerHTML = "";
      sheets.items.forEach((s) => {
        const name = s.name;
        const id = "sheet_" + Math.random().toString(36).slice(2);
        const label = document.createElement("label");
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = name;
        cb.id = id;
        // 「予測」を含むシートは既存の予測表なので初期は未チェック
        cb.checked = !name.includes("予測");
        label.appendChild(cb);
        label.appendChild(document.createTextNode(name));
        box.appendChild(label);
      });
    });
  } catch (err) {
    document.getElementById("sheets").textContent = "シート一覧の取得に失敗しました";
  }
}

function selectedSheets() {
  return Array.from(document.querySelectorAll("#sheets input:checked")).map((c) => c.value);
}

// 1シートから収入・支出を読み取る
// 想定レイアウト: A列=収入費目 B列=収入額 C列=支出費目 D列=支出額
async function readSheet(context, name) {
  const ws = context.workbook.worksheets.getItem(name);
  const used = ws.getUsedRange();
  used.load("values");
  await context.sync();

  const rows = used.values;
  const income = [];
  const expense = [];
  const stopWords = ["当月収入合計", "当月支出合計", "前月からの繰越", "翌月への繰越", "総合計", "総　合　計", "翌月への繰越予測"];

  for (const row of rows) {
    const aItem = String(row[0] ?? "").trim();
    const bAmt = row[1];
    const cItem = String(row[2] ?? "").trim();
    const dAmt = row[3];

    const isStop = (t) => stopWords.some((w) => t.replace(/\s/g, "").includes(w.replace(/\s/g, "")));

    if (aItem && typeof bAmt === "number" && bAmt !== 0 && !isStop(aItem)) {
      income.push({ item: aItem, amount: bAmt });
    }
    if (cItem && typeof dAmt === "number" && dAmt !== 0 && !isStop(cItem)) {
      expense.push({ item: cItem, amount: dAmt });
    }
  }
  return { month: name, income, expense };
}

async function runPredict() {
  const btn = document.getElementById("run");
  const sheets = selectedSheets();
  if (sheets.length === 0) {
    setStatus("実績シートを1つ以上選んでください。", "err");
    return;
  }
  btn.disabled = true;
  $results().innerHTML = "";
  setStatus('<span class="spinner"></span>実績シートを読み取っています…');

  try {
    const actuals = await Excel.run(async (context) => {
      const out = [];
      for (const name of sheets) out.push(await readSheet(context, name));
      return out;
    });

    const repayment = document.getElementById("repayment").value;

    setStatus('<span class="spinner"></span>予測を作成しています…（数十秒かかる場合があります）');
    const res = await fetch(PREDICT_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ actuals, repayment }),
    });

    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error || `サーバーエラー (${res.status})`);
    }
    const prediction = await res.json();

    setStatus('<span class="spinner"></span>予測家計表シートを作成しています…');
    const sheetName = await writePredictionSheet(prediction);

    renderNotes(prediction);
    setStatus(`「${sheetName}」シートを作成しました。金額は概算です。「要確認」項目は必ず実額をご確認ください。`, "ok");
  } catch (err) {
    setStatus("エラー: " + err.message, "err");
  } finally {
    btn.disabled = false;
  }
}

// 予測結果を新しいシートに2列体裁で書き出す
async function writePredictionSheet(prediction) {
  return await Excel.run(async (context) => {
    // 一意なシート名
    const base = "予測家計表(AI作成)";
    let name = base;
    const all = context.workbook.worksheets;
    all.load("items/name");
    await context.sync();
    let n = 2;
    const existing = all.items.map((s) => s.name);
    while (existing.includes(name)) { name = `${base}${n++}`; }

    const ws = context.workbook.worksheets.add(name);

    const income = prediction.income || [];
    const expense = prediction.expense || [];
    const maxLen = Math.max(income.length, expense.length);

    // ヘッダー部
    ws.getRange("A1").values = [[HEADER_TITLE]];
    ws.getRange("A2").values = [[NOTE_TEXT]];
    ws.getRange("A4").values = [["収　　入"]];
    ws.getRange("C4").values = [["支　　出"]];
    ws.getRange("A5:D5").values = [["費目", "金額（円）", "費目", "金額（円）"]];

    // 明細部
    const startRow = 6;
    const grid = [];
    for (let i = 0; i < maxLen; i++) {
      const inc = income[i] || { item: "", amount: "" };
      const exp = expense[i] || { item: "", amount: "" };
      const incItem = inc.item + (inc.note ? `（${inc.note}）` : "");
      const expItem = exp.item + (exp.note ? `（${exp.note}）` : "");
      grid.push([
        incItem,
        inc.amount == null ? "" : inc.amount,
        expItem,
        exp.amount == null ? "" : exp.amount,
      ]);
    }
    if (grid.length > 0) {
      ws.getRange(`A${startRow}:D${startRow + grid.length - 1}`).values = grid;
    }

    // 合計行
    const sumRow = startRow + maxLen + 1;
    const incEnd = startRow + income.length - 1;
    const expEnd = startRow + expense.length - 1;
    ws.getRange(`A${sumRow}`).values = [["当月収入合計"]];
    ws.getRange(`B${sumRow}`).formulas = [[income.length ? `=SUM(B${startRow}:B${incEnd})` : "0"]];
    ws.getRange(`C${sumRow}`).values = [["当月支出合計"]];
    ws.getRange(`D${sumRow}`).formulas = [[expense.length ? `=SUM(D${startRow}:D${expEnd})` : "0"]];

    const balRow = sumRow + 1;
    ws.getRange(`C${balRow}`).values = [["翌月への繰越予測"]];
    ws.getRange(`D${balRow}`).formulas = [[`=B${sumRow}-D${sumRow}`]];

    // --- 書式 ---
    const title = ws.getRange("A1");
    title.format.font.bold = true;
    title.format.font.size = 14;

    const note = ws.getRange("A2");
    note.format.font.size = 9;
    note.format.font.color = "#666666";

    const sectionHdr = ws.getRange(`A4:D5`);
    sectionHdr.format.font.bold = true;
    ws.getRange("A4").format.fill.color = "#D4E4DA";
    ws.getRange("C4").format.fill.color = "#D4E4DA";
    ws.getRange("A5:D5").format.fill.color = "#EDF3EF";

    // 罫線（明細〜合計）
    const tableEnd = balRow;
    const table = ws.getRange(`A5:D${tableEnd}`);
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"].forEach((edge) => {
      table.format.borders.getItem(edge).style = "Continuous";
      table.format.borders.getItem(edge).color = "#BBC7C0";
    });

    ws.getRange(`A${sumRow}:D${sumRow}`).format.font.bold = true;

    // 金額列の表示形式（カンマ区切り）
    ws.getRange(`B${startRow}:B${balRow}`).numberFormat = [["#,##0"]];
    ws.getRange(`D${startRow}:D${balRow}`).numberFormat = [["#,##0"]];

    // 列幅
    ws.getRange("A1").format.columnWidth = 240;
    ws.getRange("B1").format.columnWidth = 90;
    ws.getRange("C1").format.columnWidth = 240;
    ws.getRange("D1").format.columnWidth = 90;

    // フォント統一
    ws.getUsedRange().format.font.name = "Yu Gothic";

    ws.activate();
    await context.sync();
    return name;
  });
}

function renderNotes(prediction) {
  const warnIncome = (prediction.income || []).filter((x) => x.note);
  const warnExpense = (prediction.expense || []).filter((x) => x.note);
  const extra = prediction.notes || [];
  if (warnIncome.length + warnExpense.length + extra.length === 0) return;

  let html = '<div class="result-note"><b>要確認の項目があります：</b><br>';
  [...warnIncome, ...warnExpense].forEach((x) => {
    html += `・${x.item}：${x.note}<br>`;
  });
  extra.forEach((n) => { html += `・${n}<br>`; });
  html += "</div>";
  $results().innerHTML = html;
}
