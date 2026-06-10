// evidence.js — 訴状(Word)を読み込み、証拠説明書(Excel)を生成するアドイン
// 仕組み: Word本文 → Netlify中継(evidence-list) → Claude → 証拠一覧 → 既存テンプレートに差し込んでExcelダウンロード

var ENDPOINT = "https://word-kouetsu.netlify.app/.netlify/functions/evidence-list";

// 証拠説明書テンプレート(.xlsx)をbase64で埋め込み
var TEMPLATE_B64 = "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";

var lastEvidences = [];
var lastSide = "甲";
var lastHeader = { caseName: "", court: "", plaintiff: "", defendant: "", attorneys: [] };

Office.onReady(function () {
  document.getElementById("analyzeBtn").addEventListener("click", analyze);
  document.getElementById("downloadBtn").addEventListener("click", download);
});

function setStatus(msg, kind) {
  var el = document.getElementById("status");
  el.className = "status show " + (kind || "info");
  el.innerHTML = msg;
}

function escapeHtml(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// 本文を読み取り証拠を抽出
function analyze() {
  lastSide = document.getElementById("side").value === "乙" ? "乙" : "甲";
  document.getElementById("downloadBtn").style.display = "none";
  document.getElementById("results").innerHTML = "";
  document.getElementById("count").textContent = "";
  setStatus('<span class="spinner"></span>訴状の本文を読み取っています…', "info");

  Word.run(function (context) {
    var body = context.document.body;
    body.load("text");
    return context.sync().then(function () {
      var text = body.text || "";
      if (text.trim().length < 20) {
        setStatus("訴状の本文が見つかりません。訴状を開いた状態で実行してください。", "err");
        return;
      }
      return callApi(text);
    });
  }).catch(function (e) {
    setStatus("エラー: " + escapeHtml(e.message || String(e)), "err");
  });
}

function callApi(text) {
  setStatus('<span class="spinner"></span>証拠を抽出しています…（10〜30秒程度）', "info");
  return fetch(ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ complaint: text, side: lastSide }),
  }).then(function (res) {
    return res.text().then(function (bodyText) {
      var data;
      try {
        data = JSON.parse(bodyText);
      } catch (e) {
        // JSONでない（HTMLのエラー画面等）。多くは処理時間切れ。
        if (res.status === 502 || res.status === 504) {
          throw new Error("処理に時間がかかりすぎたため中断されました。訴状が長い場合に起こることがあります。少し時間をおいて再度お試しください。");
        }
        throw new Error("サーバーから想定外の応答が返りました（状態コード " + res.status + "）。少し待って再度お試しください。");
      }
      if (!res.ok) throw new Error(data.error || "サーバーエラー");
      if (data.side === "甲" || data.side === "乙") lastSide = data.side; // 本文から判定された区分を優先
      lastEvidences = Array.isArray(data.evidences) ? data.evidences : [];
      lastHeader = {
        caseName: data.caseName || "",
        court: data.court || "",
        plaintiff: data.plaintiff || "",
        defendant: data.defendant || "",
        attorneys: Array.isArray(data.attorneys) ? data.attorneys : [],
      };
      render(lastEvidences, data.notes || []);
    });
  });
}

function render(evidences, notes) {
  var box = document.getElementById("results");
  box.innerHTML = "";
  if (evidences.length === 0) {
    setStatus("訴状から証拠の引用が見つかりませんでした。", "err");
    return;
  }
  document.getElementById("count").textContent = lastSide + "号証 " + evidences.length + " 件を抽出しました。内容をご確認ください。";
  evidences.forEach(function (e) {
    var div = document.createElement("div");
    div.className = "ev";
    var html = '<div class="num">' + escapeHtml(e.number) + "</div>";
    html += '<div class="title">' + escapeHtml(e.title || "（標目 要確認）") + "</div>";
    html += '<div class="row">原本／写し: ' + escapeHtml(e.originalOrCopy || "") +
            "　作成日: " + escapeHtml(e.date || "（要確認）") + "</div>";
    html += '<div class="row">作成名義人: ' + escapeHtml(e.author || "（要確認）") + "</div>";
    html += '<div class="row">立証趣旨: ' + escapeHtml(e.purpose || "") + "</div>";
    if (e.note) html += '<div class="note">⚠ ' + escapeHtml(e.note) + "</div>";
    div.innerHTML = html;
    box.appendChild(div);
  });
  setStatus("下書きが完成しました。ボタンからExcelをダウンロードできます。", "ok");
  document.getElementById("downloadBtn").style.display = "block";
}

// base64 → Uint8Array
function b64ToBytes(b64) {
  var bin = atob(b64);
  var len = bin.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function xmlEscape(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// 1行分のXML（A〜F列）。スタイルは折り返し有りのものに統一して、隣の列に文字があふれないようにする。
// A=証拠番号(中央折返12) B=標目(左上折返21) C=原本写し(中央折返14) D=作成日(中央折返22) E=作成名義人(左上折返21) F=立証趣旨(左上折返21)
var COLS = [["A", "12"], ["B", "20"], ["C", "14"], ["D", "15"], ["E", "20"], ["F", "20"]];
function cellXml(col, style, val) {
  // 値が空でも、スタイル（罫線・折り返し）を効かせるため inlineStr の空文字で書く
  if (val == null || val === "") return '<c r="' + col + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve"></t></is></c>';
  return '<c r="' + col + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(val) + "</t></is></c>";
}
function rowXml(rnum, vals) {
  var cs = "";
  COLS.forEach(function (c, i) { cs += cellXml(c[0] + rnum, c[1], vals[i]); });
  return '<row r="' + rnum + '" spans="1:6" ht="70.05" customHeight="1" x14ac:dyDescent="0.3">' + cs + "</row>";
}

function download() {
  try {
    var fflate = window.fflate;
    if (!fflate || typeof fflate.unzipSync !== "function") {
      setStatus("Excel生成に必要な部品の読み込みに失敗しました。ページを再読み込みして再度お試しください。", "err");
      return;
    }
    var zipObj = fflate.unzipSync(b64ToBytes(TEMPLATE_B64));
    var dec = new TextDecoder("utf-8");
    var xml = dec.decode(zipObj["xl/worksheets/sheet1.xml"]);

    // 指定セルを inlineStr で置き換える（スタイルsは指定があれば差し替え、なければ既存を維持）
    function setCell(ref, val, styleOverride) {
      var esc = xmlEscape(val);
      // 既存セル <c r="ref" ...>...</c> または自己終端 <c r="ref" .../> を探す
      var reFull = new RegExp('<c r="' + ref + '"([^>]*?)(/>|>[\\s\\S]*?</c>)');
      var m = xml.match(reFull);
      var style = styleOverride;
      if (m && !style) {
        var sm = m[1].match(/s="(\d+)"/);
        style = sm ? sm[1] : null;
      }
      var sAttr = style ? ' s="' + style + '"' : "";
      var cell = '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc + "</t></is></c>";
      if (m) {
        xml = xml.replace(m[0], cell);
      } else {
        // セルが無い場合は該当行に追加（簡易: 行末尾に挿入）
        var rownum = ref.replace(/[A-Z]+/, "");
        var reRow = new RegExp('(<row r="' + rownum + '"[^>]*>)([\\s\\S]*?)(</row>)');
        if (reRow.test(xml)) xml = xml.replace(reRow, "$1$2" + cell + "$3");
      }
    }

    // 今日の日付を「令和○年○月○日」で
    function todayWareki() {
      var d = new Date();
      var y = d.getFullYear();
      var reiwaY = y - 2018; // 2019=令和1
      var yStr = reiwaY === 1 ? "元" : String(reiwaY);
      return "令和" + yStr + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日";
    }

    // 事件名(A1)・原告(B3)・被告(B5)・作成日(F6=出力日)・裁判所名(A14)・代理人(E10肩書き,F10,F12)を埋める。
    // いずれも既存セルのスタイルsを維持したまま値だけ差し替える。
    if (lastHeader.caseName) setCell("A1", lastHeader.caseName);
    if (lastHeader.plaintiff) setCell("B3", lastHeader.plaintiff);
    if (lastHeader.defendant) setCell("B5", lastHeader.defendant);
    setCell("F6", todayWareki()); // 作成日＝Word出力日
    if (lastHeader.court) setCell("A14", lastHeader.court + "　御中");
    setCell("E10", (lastSide === "乙" ? "被告" : "原告") + "訴訟代理人");
    if (lastHeader.attorneys[0]) setCell("F10", "弁護士　　" + lastHeader.attorneys[0]);
    if (lastHeader.attorneys[1]) setCell("F12", "弁護士　　" + lastHeader.attorneys[1]);
    else setCell("F12", ""); // 2人目がいなければ空に（プレースホルダー消去）
    if (!lastHeader.attorneys[1]) setCell("E12", ""); // 「同」も消す

    var rnum = 19;
    lastEvidences.forEach(function (e) {
      var vals = [e.number || "", e.title || "", e.originalOrCopy || "", e.date || "", e.author || "", e.purpose || ""];
      var re = new RegExp('<row r="' + rnum + '"[^>]*>[\\s\\S]*?</row>');
      var newRow = rowXml(rnum, vals);
      if (re.test(xml)) xml = xml.replace(re, newRow);
      else xml = xml.replace("</sheetData>", newRow + "</sheetData>");
      rnum++;
    });

    zipObj["xl/worksheets/sheet1.xml"] = fflate.strToU8(xml);
    var out = fflate.zipSync(zipObj);
    var blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "証拠説明書_" + lastSide + "号証.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
    setStatus("証拠説明書をダウンロードしました。内容をご確認ください。", "ok");
  } catch (e) {
    setStatus("Excel生成エラー: " + escapeHtml(e.message || String(e)), "err");
  }
}
