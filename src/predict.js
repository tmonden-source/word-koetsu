// predict.js — 予測家計表アドイン本体（ひな形を崩さず金額だけ書き込む方式）
//
// 流れ:
//  1. ブックの全シート名を一覧表示
//     - 実績シート（複数選択可）
//     - 「予測家計表ひな形」シート（書き込み先。1つ選択）
//  2. 実績シートから 収入/支出 を読み取る
//  3. ひな形シートの費目（A列=収入費目, C列=支出費目）と金額セル番地を読み取る
//  4. 中継関数へ「実績＋ひな形の費目リスト」を送り、費目ごとの金額を受け取る
//  5. ひな形シートの金額セル（B列/D列）にだけ値を書き込む
//     → 結合セル・数式セル・罫線・列幅はすべて維持される

const PREDICT_ENDPOINT = "https://word-kouetsu.netlify.app/.netlify/functions/predict-budget";

// 金額を書き込んではいけない費目（合計・繰越など。部分一致で除外）
const STOP_ITEMS = [
  "当月収入合計", "当月支出合計", "前月からの繰越", "翌月への繰越",
  "翌月への繰越予測", "総合計", "総　合　計",
];
// 表のヘッダー見出し（完全一致でのみ除外。部分一致だと「自営収入」等を誤除外するため）
const HEADER_LABELS = [
  "収入", "支出", "費目", "金額（円）", "金額(円)",
];

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

function norm(s) { return String(s == null ? "" : s).replace(/\s/g, "").replace(/\u3000/g, ""); }
function isStop(item) {
  const n = norm(item);
  if (HEADER_LABELS.some((w) => n === norm(w))) return true;       // 見出しは完全一致で除外
  return STOP_ITEMS.some((w) => n.includes(norm(w)));               // 合計類は部分一致で除外
}

// シート一覧を読み込み、実績用チェックボックスとひな形選択を作る
async function loadSheetList() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const names = sheets.items.map((s) => s.name);

      // 実績チェックボックス
      const box = document.getElementById("sheets");
      box.innerHTML = "";
      names.forEach((name) => {
        const label = document.createElement("label");
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = name;
        cb.checked = !name.includes("予測"); // 予測シートは実績ではないので初期OFF
        label.appendChild(cb);
        label.appendChild(document.createTextNode(name));
        box.appendChild(label);
      });

      // ひな形シート選択（プルダウン）
      const sel = document.getElementById("templateSheet");
      sel.innerHTML = "";
      names.forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        if (name.includes("予測")) opt.selected = true; // 予測シートを既定に
        sel.appendChild(opt);
      });
    });
  } catch (err) {
    document.getElementById("sheets").textContent = "シート一覧の取得に失敗しました";
  }
}

function selectedActualSheets() {
  return Array.from(document.querySelectorAll("#sheets input:checked")).map((c) => c.value);
}

// 実績シートから収入・支出を読む（A列費目/B列金額, C列費目/D列金額）
async function readActual(context, name) {
  const ws = context.workbook.worksheets.getItem(name);
  const used = ws.getUsedRange();
  used.load("values");
  await context.sync();

  const income = [], expense = [];
  for (const row of used.values) {
    const aItem = String(row[0] == null ? "" : row[0]).trim();
    const bAmt = row[1];
    const cItem = String(row[2] == null ? "" : row[2]).trim();
    const dAmt = row[3];
    if (aItem && typeof bAmt === "number" && bAmt !== 0 && !isStop(aItem)) income.push({ item: aItem, amount: bAmt });
    if (cItem && typeof dAmt === "number" && dAmt !== 0 && !isStop(cItem)) expense.push({ item: cItem, amount: dAmt });
  }
  return { month: name, income, expense };
}

// ひな形シートを読み、費目→金額セル番地の対応を作る
// 収入: A列に費目, 金額はB列同じ行 / 支出: C列に費目, 金額はD列同じ行
async function readTemplate(context, name) {
  const ws = context.workbook.worksheets.getItem(name);
  const used = ws.getUsedRange();
  used.load("values, rowIndex, columnIndex, formulas");
  await context.sync();

  const baseRow = used.rowIndex; // 0始まり
  const values = used.values;
  const formulas = used.formulas;

  const incomeMap = {};  // 費目名 -> セル番地(B列)
  const expenseMap = {}; // 費目名 -> セル番地(D列)
  const incomeItems = [];
  const expenseItems = [];

  for (let i = 0; i < values.length; i++) {
    const r = baseRow + i + 1; // Excelの実行番号(1始まり)
    const aItem = String(values[i][0] == null ? "" : values[i][0]).trim();
    const cItem = String(values[i][2] == null ? "" : values[i][2]).trim();

    // 収入費目: B列が数式でない（合計行を除外）かつ費目名がある
    if (aItem && !isStop(aItem)) {
      const bFormula = String(formulas[i][1] == null ? "" : formulas[i][1]);
      if (!bFormula.startsWith("=")) {
        incomeMap[aItem] = "B" + r;
        incomeItems.push(aItem);
      }
    }
    // 支出費目: D列が数式でない
    if (cItem && !isStop(cItem)) {
      const dFormula = String(formulas[i][3] == null ? "" : formulas[i][3]);
      if (!dFormula.startsWith("=")) {
        expenseMap[cItem] = "D" + r;
        expenseItems.push(cItem);
      }
    }
  }
  return { incomeMap, expenseMap, incomeItems, expenseItems };
}

async function runPredict() {
  const btn = document.getElementById("run");
  const actualSheets = selectedActualSheets();
  const templateName = document.getElementById("templateSheet").value;

  if (actualSheets.length === 0) { setStatus("実績シートを1つ以上選んでください。", "err"); return; }
  if (!templateName) { setStatus("ひな形シートを選んでください。", "err"); return; }
  if (actualSheets.includes(templateName)) {
    setStatus("ひな形シートは実績シートから外してください（同じシートは選べません）。", "err"); return;
  }

  btn.disabled = true;
  $results().innerHTML = "";
  setStatus('<span class="spinner"></span>実績とひな形を読み取っています…');

  try {
    const { actuals, template } = await Excel.run(async (context) => {
      const acts = [];
      for (const n of actualSheets) acts.push(await readActual(context, n));
      const tpl = await readTemplate(context, templateName);
      return { actuals: acts, template: tpl };
    });

    if (template.incomeItems.length === 0 && template.expenseItems.length === 0) {
      throw new Error("ひな形シートから費目を読み取れませんでした。シート選択が正しいか確認してください。");
    }

    const repayment = document.getElementById("repayment").value;

    setStatus('<span class="spinner"></span>予測を作成しています…（数十秒かかる場合があります）');
    const res = await fetch(PREDICT_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        actuals,
        templateIncomeItems: template.incomeItems,
        templateExpenseItems: template.expenseItems,
        repayment,
      }),
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error || ("サーバーエラー (" + res.status + ")"));
    }
    const prediction = await res.json();

    setStatus('<span class="spinner"></span>ひな形に金額を書き込んでいます…');
    const count = await writeAmounts(templateName, template, prediction);

    renderReview(prediction);
    setStatus("「" + templateName + "」に " + count + " 件の金額を書き込みました。様式は維持されています。「要確認」項目は実額をご確認ください。", "ok");
  } catch (err) {
    setStatus("エラー: " + err.message, "err");
  } finally {
    btn.disabled = false;
  }
}

// ひな形の金額セルにだけ値を書き込む（書式・結合・数式は触らない）
async function writeAmounts(templateName, template, prediction) {
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(templateName);
    let count = 0;

    const inc = prediction.incomeAmounts || {};
    const exp = prediction.expenseAmounts || {};

    for (const [item, addr] of Object.entries(template.incomeMap)) {
      if (item in inc && typeof inc[item] === "number") {
        ws.getRange(addr).values = [[inc[item]]];
        count++;
      }
    }
    for (const [item, addr] of Object.entries(template.expenseMap)) {
      if (item in exp && typeof exp[item] === "number") {
        ws.getRange(addr).values = [[exp[item]]];
        count++;
      }
    }

    ws.activate();
    await context.sync();
    return count;
  });
}

function renderReview(prediction) {
  const items = prediction.reviewItems || [];
  if (items.length === 0) return;
  let html = '<div class="result-note"><b>要確認の項目があります：</b><br>';
  items.forEach((x) => {
    const where = x.where === "income" ? "収入" : "支出";
    html += "・[" + where + "] " + x.item + "：" + (x.note || "要確認") + "<br>";
  });
  html += "</div>";
  $results().innerHTML = html;
}
