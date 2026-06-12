// contract.js — 委任契約書（法人破産）の穴埋めアドイン
// 入力 → 合計・分割を自動計算 → Word本文の表・署名欄に差し込む。AI/サーバー不使用。

Office.onReady(function () {
  document.getElementById("applyBtn").addEventListener("click", apply);
  // 支払方法ラジオの切替
  var radios = document.getElementsByName("pay");
  for (var i = 0; i < radios.length; i++) radios[i].addEventListener("change", togglePay);
  // 金額入力で自動計算
  ["fee", "admin", "deposit", "firstLump", "months", "round"].forEach(function (id) {
    document.getElementById(id).addEventListener("input", recalc);
    document.getElementById(id).addEventListener("change", recalc);
  });
  togglePay();
  recalc();
});

function yen(n) { return Number(n) || 0; }

// 全角数字化（金額表記を全角に）
function toFull(s) { return String(s).replace(/[0-9]/g, function (d) { return String.fromCharCode(d.charCodeAt(0) + 0xFEE0); }); }

// 数値を「○○万○○○○円」形式（端数なしなら○○万円、全角）
function formatYen(n) {
  n = Math.round(n);
  if (n < 10000) return toFull(n) + "円";
  var oku = Math.floor(n / 100000000);
  var man = Math.floor((n % 100000000) / 10000);
  var rest = n % 10000;
  var out = "";
  if (oku > 0) {
    out += toFull(oku) + "億";
    if (man > 0 || rest > 0) out += toFull(String(man)) + "万";
  } else if (man > 0) {
    out += toFull(man) + "万";
  }
  if (rest > 0) out += toFull(String(rest));
  if (out === "") out = toFull(0);
  return out + "円";
}

function getSum() { return yen(document.getElementById("fee").value) + yen(document.getElementById("admin").value) + yen(document.getElementById("deposit").value); }

function togglePay() {
  var v = document.querySelector('input[name="pay"]:checked').value;
  document.getElementById("splitFields").style.display = (v === "分割") ? "block" : "none";
  document.getElementById("lumpFields").style.display = (v === "一括") ? "block" : "none";
}

// 分割計算: 合計total, 初回一括firstLump, 回数months, 丸めround → 内訳
function calcSplit(total, firstLump, months, round) {
  firstLump = yen(firstLump);
  var remain = total - firstLump;
  if (months <= 0 || remain < 0) return null;
  var per = Math.floor(remain / months);
  if (round > 1) per = Math.floor(per / round) * round;
  var last = remain - per * (months - 1);
  return { firstLump: firstLump, per: per, last: last, months: months, remain: remain };
}

function recalc() {
  var sum = getSum();
  var dep = yen(document.getElementById("deposit").value);
  document.getElementById("sumView").textContent = sum > 0 ? formatYen(sum) : "—";
  document.getElementById("depView").textContent = dep > 0 ? formatYen(dep) : "—";

  // 分割プレビュー
  var v = document.querySelector('input[name="pay"]:checked').value;
  if (v === "分割") {
    var months = yen(document.getElementById("months").value);
    var firstLump = yen(document.getElementById("firstLump").value);
    var round = yen(document.getElementById("round").value) || 1;
    var box = document.getElementById("splitPreview");
    if (sum > 0 && months > 0) {
      var r = calcSplit(sum, firstLump, months, round);
      if (r) {
        var txt = "";
        if (r.firstLump > 0) txt += "初回一括：" + formatYen(r.firstLump) + " ／ ";
        if (r.months === 1) txt += "1回：" + formatYen(r.last);
        else txt += "毎月：" + formatYen(r.per) + "（×" + (r.months - 1) + "回）、最終回：" + formatYen(r.last);
        box.textContent = "分割内訳： " + txt;
      }
    } else {
      box.textContent = "分割内訳：合計と回数を入力すると表示されます";
    }
  }
}

function setStatus(msg, kind) {
  var el = document.getElementById("status");
  el.className = "status show " + (kind || "info");
  el.textContent = msg;
}

// 支払条件の文面を組み立てる
function buildSplitText(sum, firstLump, months, round, startY, startM) {
  var r = calcSplit(sum, firstLump, months, round);
  if (!r) return null;
  var lines = [];

  // 月を進めるヘルパー（{y,m}を返す）
  function addMonth(y, m, n) {
    var total = (y * 12 + (m - 1)) + n;
    return { y: Math.floor(total / 12), m: (total % 12) + 1 };
  }
  function ymd(y, m) { return "令和" + toFull(y) + "年" + toFull(m) + "月"; }

  var hasDate = (startY && startM);

  // 初回一括
  if (r.firstLump > 0) {
    if (hasDate) lines.push(ymd(startY, startM) + "末日限り　" + formatYen(r.firstLump) + "（初回一括）");
    else lines.push("初回一括　" + formatYen(r.firstLump));
  }

  if (hasDate) {
    // 月割りの開始月：初回一括があればその翌月から、なければ初回支払月から
    var off = (r.firstLump > 0) ? 1 : 0;
    var s = addMonth(startY, startM, off);            // 月割り開始
    var lastPos = addMonth(s.y, s.m, r.months - 1);   // 最終回
    // 期間表記
    lines.push(ymd(s.y, s.m) + "から" + ymd(lastPos.y, lastPos.m) + "まで");
    if (r.months === 1) {
      lines.push(ymd(s.y, s.m) + "末日限り　" + formatYen(r.last));
    } else {
      lines.push("毎月末日限り　" + formatYen(r.per));
      lines.push(ymd(lastPos.y, lastPos.m) + "末日限り　" + formatYen(r.last) + "（最終回）");
    }
  } else {
    if (r.months === 1) lines.push("1回　" + formatYen(r.last));
    else lines.push("毎月末日限り　" + formatYen(r.per) + "（全" + r.months + "回、最終回 " + formatYen(r.last) + "）");
  }
  return lines.join("\n");
}

function apply() {
  var sum = getSum();
  if (sum <= 0) { setStatus("金額を入力してください。", "err"); return; }
  setStatus("契約書に反映しています…", "info");

  var fee = yen(document.getElementById("fee").value);
  var admin = yen(document.getElementById("admin").value);
  var deposit = yen(document.getElementById("deposit").value);
  var payType = document.querySelector('input[name="pay"]:checked').value;
  var addr = document.getElementById("addr").value.trim();
  var company = document.getElementById("company").value.trim();
  var rep = document.getElementById("rep").value.trim();

  Word.run(function (context) {
    var body = context.document.body;
    var tables = body.tables;
    tables.load("items");
    return context.sync().then(function () {
      // 表0=金額表, 表1=支払方法（本文中の最初の2つの表）
      var t0 = tables.items[0];
      var t1 = tables.items[1];
      t0.load("values"); t1.load("values");
      return context.sync().then(function () {
        // --- 金額表（表0） ---
        // 行1:法人破産, 行2:事務手数料, 行3:預り金, 行4:合計（いずれも3列目=index2）
        setCellText(t0, 1, 2, formatYen(fee));
        setCellText(t0, 2, 2, formatYen(admin));
        setCellText(t0, 3, 2, formatYen(deposit));
        setCellText(t0, 4, 2, "合計" + formatYen(sum) + "\n　　　　　うち、預り金 " + formatYen(deposit));

        // --- 支払方法（表1） ---
        if (payType === "一括") {
          setCellText(t1, 1, 1, "一括払い");
          var ly = yen(document.getElementById("lumpY").value);
          var lm = yen(document.getElementById("lumpM").value);
          var ld = yen(document.getElementById("lumpD").value);
          var cond;
          if (ly && lm && ld) cond = "令和" + toFull(ly) + "年" + toFull(lm) + "月" + toFull(ld) + "日限り　" + formatYen(sum);
          else cond = "令和〇年〇月〇日限り　" + formatYen(sum);
          setCellText(t1, 1, 2, cond);
        } else {
          setCellText(t1, 1, 1, "分割払い");
          var months = yen(document.getElementById("months").value);
          var firstLump = yen(document.getElementById("firstLump").value);
          var round = yen(document.getElementById("round").value) || 1;
          var startY = yen(document.getElementById("startY").value);
          var startM = yen(document.getElementById("startM").value);
          var txt = buildSplitText(sum, firstLump, months, round, startY, startM);
          if (txt) setCellText(t1, 1, 2, txt);
        }

        // --- 署名欄（住所・お名前） ---
        // 段落から「ご住所」「お名前」を含む行を探して追記する
        return fillSignature(context, body, addr, company, rep);
      });
    });
  }).then(function () {
    setStatus("契約書に反映しました。内容をご確認ください。", "ok");
  }).catch(function (e) {
    setStatus("エラー: " + (e.message || String(e)), "err");
  });
}

// 表のセルにテキストを設定（行・列は0始まり）
function setCellText(table, rowIdx, colIdx, text) {
  var cell = table.getCell(rowIdx, colIdx);
  cell.body.clear();
  cell.body.insertText(text, Word.InsertLocation.start);
}

// 署名欄を埋める：「ご住所」「お名前」を含む段落の後ろに値を入れる
function fillSignature(context, body, addr, company, rep) {
  var paras = body.paragraphs;
  paras.load("items/text");
  return context.sync().then(function () {
    var items = paras.items;
    for (var i = 0; i < items.length; i++) {
      var t = items[i].text;
      if (addr && t.indexOf("ご住所") >= 0) {
        items[i].insertText("ご住所　 " + addr, Word.InsertLocation.replace);
      }
      if (company && rep && t.indexOf("お名前") >= 0) {
        items[i].insertText("お名前　 " + company + "　" + rep, Word.InsertLocation.replace);
      }
    }
    return context.sync();
  });
}
