/**
 * フリマ売上ランキング：保存先のスプレッドシート用プログラム
 * スプレッドシートの「拡張機能 → Apps Script」に貼り付けて、ウェブアプリとしてデプロイする。
 */
const SALES_SHEET = "売上";
const SALES_HEAD = ["id", "人番号", "売った人", "ジャンル", "品名", "単価", "個数", "金額", "日時"];
const NAMES_SHEET = "メンバー";
const CATS = ["服", "皿", "靴", "その他雑貨"];
const PEOPLE = 4;

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function names_() {
  const sh = sheet_(NAMES_SHEET, ["名前"]);
  return sh.getRange(2, 1, PEOPLE, 1).getValues().map(r => String(r[0] || "").replace(/^'/, ""));
}

function state_() {
  const sh = sheet_(SALES_SHEET, SALES_HEAD);
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, SALES_HEAD.length).getValues() : [];
  const sales = rows.filter(r => r[0] !== "").map(r => ({
    id: String(r[0]),
    who: Number(r[1]),
    cat: String(r[3]),
    item: String(r[4]).replace(/^'/, ""),
    price: Number(r[5]),
    qty: Number(r[6]),
    at: r[8] instanceof Date ? r[8].getTime() : 0,
  }));
  return { sales: sales, names: names_() };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 品名などが数式として扱われないようにする
function safeText_(s, max) {
  s = String(s || "").slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  return json_(Object.assign({ ok: true }, state_()));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const sh = sheet_(SALES_SHEET, SALES_HEAD);

    if (req.action === "add") {
      const who = Math.max(0, Math.min(PEOPLE - 1, Math.floor(Number(req.who) || 0)));
      const price = Math.round(Number(req.price));
      const qty = Math.max(1, Math.round(Number(req.qty) || 1));
      const cat = CATS.indexOf(req.cat) >= 0 ? req.cat : "その他雑貨";
      if (!(price >= 1 && price <= 10000000) || qty > 1000) return json_({ ok: false, error: "bad_value" });
      const name = names_()[who] || (who + 1) + "人目";
      sh.appendRow([Utilities.getUuid(), who, safeText_(name, 30), cat, safeText_(req.item, 100), price, qty, price * qty, new Date()]);

    } else if (req.action === "delete") {
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
        const i = ids.indexOf(String(req.id));
        if (i >= 0) sh.deleteRow(i + 2);
      }

    } else if (req.action === "clear") {
      const last = sh.getLastRow();
      if (last > 1) sh.deleteRows(2, last - 1);

    } else if (req.action === "names") {
      const list = Array.isArray(req.names) ? req.names : [];
      const values = [];
      for (let i = 0; i < PEOPLE; i++) values.push([safeText_(list[i], 30)]);
      sheet_(NAMES_SHEET, ["名前"]).getRange(2, 1, PEOPLE, 1).setValues(values);

    } else {
      return json_({ ok: false, error: "unknown_action" });
    }
    return json_(Object.assign({ ok: true }, state_()));
  } finally {
    lock.releaseLock();
  }
}
