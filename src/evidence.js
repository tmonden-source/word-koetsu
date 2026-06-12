// evidence.js — 訴状(Word)を読み込み、証拠説明書(Excel)を生成するアドイン
// 仕組み: Word本文 → evidence-start(受付) → バックグラウンドでSonnet抽出 → evidence-resultで結果取得 → テンプレートに差し込んでExcelダウンロード

var BASE = "https://word-kouetsu.netlify.app/.netlify/functions";
var START_ENDPOINT = BASE + "/evidence-start";
var RESULT_ENDPOINT = BASE + "/evidence-result";

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

function applyResult(data) {
  if (data.side === "甲" || data.side === "乙") lastSide = data.side;
  lastEvidences = Array.isArray(data.evidences) ? data.evidences : [];
  lastHeader = {
    caseName: data.caseName || "",
    court: data.court || "",
    plaintiff: data.plaintiff || "",
    defendant: data.defendant || "",
    attorneys: Array.isArray(data.attorneys) ? data.attorneys : [],
  };
  render(lastEvidences, data.notes || []);
}

// 受付 → 結果待ち（ポーリング）方式。長い書面でも時間切れにならない。
function callApi(text) {
  setStatus('<span class="spinner"></span>証拠抽出を受け付けています…', "info");
  return fetch(START_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ complaint: text, side: lastSide }),
  }).then(function (res) {
    return res.json().then(function (data) {
      if (!res.ok && res.status !== 202) throw new Error(data.error || "受付に失敗しました");
      if (!data.jobId) throw new Error("受付番号を取得できませんでした");
      return pollResult(data.jobId);
    });
  });
}

// 結果が出るまで数秒おきに確認する
function pollResult(jobId) {
  var attempts = 0;
  var maxAttempts = 90; // 約3分（2秒間隔）まで待つ
  var elapsed = 0;
  return new Promise(function (resolve, reject) {
    function check() {
      attempts++;
      elapsed += 2;
      setStatus('<span class="spinner"></span>証拠を抽出しています…（経過 ' + elapsed + ' 秒）しばらくお待ちください。', "info");
      fetch(RESULT_ENDPOINT + "?jobId=" + encodeURIComponent(jobId), { method: "GET" })
        .then(function (res) { return res.json(); })
        .then(function (data) {
          if (data.status === "done") {
            applyResult(data.result || {});
            resolve();
          } else if (data.status === "error") {
            reject(new Error(data.error || "処理中にエラーが発生しました"));
          } else {
            // pending
            if (attempts >= maxAttempts) {
              reject(new Error("時間内に処理が完了しませんでした。証拠の件数が非常に多い場合に起こることがあります。少し時間をおいて再度お試しください。"));
            } else {
              setTimeout(check, 2000);
            }
          }
        })
        .catch(function () {
          // 一時的な通信エラーはリトライ
          if (attempts >= maxAttempts) reject(new Error("結果の取得に失敗しました。少し時間をおいて再度お試しください。"));
          else setTimeout(check, 2000);
        });
    }
    // 受付直後は処理開始まで少し待ってから確認を始める
    setTimeout(check, 2000);
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
var COLS = [["A", "12"], ["B", "18"], ["C", "14"], ["D", "15"], ["E", "18"], ["F", "18"]];
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
