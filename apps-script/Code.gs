/**
 * ED Shirt Order — Google Apps Script backend.
 *
 * Paste this into Extensions > Apps Script on a PRIVATE Google Sheet.
 * Orders are written to the "Orders" tab; the "Totals" tab tallies sizes.
 * The web app only accepts new orders. It never returns anyone's data,
 * so the sheet stays visible only to people you share it with.
 *
 * Optional: set a team code so only staff who know it can order.
 *   Project Settings > Script Properties > add TEAM_CODE = <your code>
 */

const SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"];
const MAX_PER_SIZE = 25;
const ORDERS = "Orders";
const TOTALS = "Totals";
const HEADERS = ["Order #", "Submitted", "Last updated", "Name", "Email"]
  .concat(SIZES, ["Total"]);
const EMAIL_SUBJECT = "Your ED shirt order";

/** Run once from the editor (Run > setup) to build the tabs. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ORDERS);
  if (!sh) sh = ss.insertSheet(ORDERS);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);

  let t = ss.getSheetByName(TOTALS);
  if (!t) t = ss.insertSheet(TOTALS);
  t.clear();
  t.getRange(1, 1, 1, 2).setValues([["Size", "Shirts"]]).setFontWeight("bold");
  SIZES.forEach(function (s, i) {
    const col = columnLetter(HEADERS.indexOf(s) + 1);
    t.getRange(i + 2, 1, 1, 2).setValues([[s, "=SUM(" + ORDERS + "!" + col + "2:" + col + ")"]]);
  });
  const last = SIZES.length + 2;
  t.getRange(last, 1, 1, 2).setValues([["Total", "=SUM(B2:B" + (last - 1) + ")"]]).setFontWeight("bold");
  t.getRange(last + 1, 1, 1, 2).setValues([["Orders", "=COUNTA(" + ORDERS + "!A2:A)"]]);

  const def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 2 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function doGet() {
  return json({ ok: true, service: "ed-shirt-order" });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // Bots fill hidden fields; people don't.
    if (body.website) return json({ ok: true, number: "000000" });

    const code = PropertiesService.getScriptProperties().getProperty("TEAM_CODE");
    if (code && String(body.code || "").trim().toLowerCase() !== code.trim().toLowerCase()) {
      return json({ ok: false, error: "bad_code" });
    }

    const name = clean(body.name, 80);
    const email = clean(body.email, 120).toLowerCase();
    if (!name) return json({ ok: false, error: "missing_name" });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ ok: false, error: "bad_email" });

    const sizes = {};
    let total = 0;
    SIZES.forEach(function (s) {
      const n = Math.floor(Number(body.sizes && body.sizes[s]) || 0);
      const q = Math.max(0, Math.min(MAX_PER_SIZE, n));
      if (q) { sizes[s] = q; total += q; }
    });
    if (!total) return json({ ok: false, error: "no_shirts" });

    // Slow down repeat submissions from the same email.
    const cache = CacheService.getScriptCache();
    const key = "rl_" + Utilities.base64EncodeWebSafe(email);
    if (cache.get(key)) return json({ ok: false, error: "too_fast" });
    cache.put(key, "1", 20);

    lock.waitLock(10000);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
    if (!sh) return json({ ok: false, error: "not_set_up" });

    const now = new Date();
    const emailCol = HEADERS.indexOf("Email") + 1;
    const lastRow = sh.getLastRow();
    let row = 0, number = "", submitted = now, updated = false;
    if (lastRow > 1) {
      const emails = sh.getRange(2, emailCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (String(emails[i][0]).toLowerCase() === email) { row = i + 2; break; }
      }
    }
    if (row) {
      const existing = sh.getRange(row, 1, 1, 2).getValues()[0];
      number = String(existing[0]);
      submitted = existing[1] || now;
      updated = true;
    } else {
      row = lastRow + 1;
      number = String(Math.floor(100000 + Math.random() * 900000));
    }

    const values = [number, submitted, now, name, email]
      .concat(SIZES.map(function (s) { return sizes[s] || 0; }), [total]);
    sh.getRange(row, 1, 1, values.length).setValues([values]);
    lock.releaseLock();

    sendConfirmation(email, name, number, sizes, total, updated);
    return json({ ok: true, number: number, updated: updated });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: "server" });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function sendConfirmation(email, name, number, sizes, total, updated) {
  const lines = SIZES.filter(function (s) { return sizes[s]; })
    .map(function (s) { return sizes[s] + " × " + s; });
  const text = [
    "Hi " + name + ",",
    "",
    (updated ? "Your shirt order was updated." : "Your shirt order was received."),
    "",
    "Order #" + number,
    lines.join("\n"),
    "Total: " + total + " shirt" + (total === 1 ? "" : "s"),
    "",
    "To change your order, submit the form again with this same email address.",
  ].join("\n");
  try {
    MailApp.sendEmail({
      to: email,
      subject: EMAIL_SUBJECT + (updated ? " (updated)" : ""),
      body: text,
      htmlBody: confirmationHtml(name, number, sizes, total, updated),
      name: "ED Shirt Order",
    });
  } catch (err) {
    console.error("Email failed: " + err);
  }
}

/** Email version of the form's look: grey paper, navy ink, mustard tag. */
function confirmationHtml(name, number, sizes, total, updated) {
  const INK = "#1d2a44", PAPER = "#e7e8eb", CARD = "#f7f7f8", LINE = "#c6c9d1",
        MUTED = "#5d6678", ACCENT = "#e0a526", OK = "#2f7a4f";
  const DISPLAY = "'Archivo Black','Arial Black',Impact,Helvetica,Arial,sans-serif";
  const BODY = "Archivo,Helvetica,Arial,sans-serif";
  const MONO = "'JetBrains Mono',Menlo,Consolas,monospace";
  const esc = function (s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  };
  const label = function (t) {
    return '<td style="padding:8px 0;font-family:' + MONO + ';font-size:12px;letter-spacing:1px;' +
      'text-transform:uppercase;color:' + MUTED + ';vertical-align:top;width:110px">' + t + '</td>';
  };
  const rows = SIZES.filter(function (s) { return sizes[s]; }).map(function (s) {
    return '<tr><td style="padding:10px 0;border-bottom:1px solid ' + LINE + ';font-family:' + DISPLAY +
      ';font-size:18px;color:' + INK + '">' + s + '</td><td align="right" style="padding:10px 0;border-bottom:1px solid ' +
      LINE + ';font-family:' + MONO + ';font-size:15px;color:' + INK + '">× ' + sizes[s] + '</td></tr>';
  }).join("");

  return '' +
  '<div style="margin:0;padding:0;background:' + PAPER + '">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + PAPER + '">' +
  '<tr><td align="center" style="padding:32px 16px">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px">' +
    '<tr><td style="padding-bottom:6px;font-family:' + MONO + ';font-size:12px;letter-spacing:2px;text-transform:uppercase;color:' + MUTED + '">Emergency Department · West Orange</td></tr>' +
    '<tr><td style="padding-bottom:22px;font-family:' + DISPLAY + ';font-size:40px;line-height:44px;text-transform:uppercase;color:' + INK + '">ED team <span style="background:' + ACCENT + ';color:' + INK + ';padding:0 6px">shirt</span> order</td></tr>' +
    '<tr><td style="background:' + CARD + ';border:1px solid ' + LINE + ';border-top:4px solid ' + OK + ';border-radius:10px;padding:24px">' +
      '<div style="font-family:' + DISPLAY + ';font-size:26px;text-transform:uppercase;color:' + OK + ';margin:0 0 8px">' + (updated ? "Order updated" : "Order placed") + '</div>' +
      '<div style="font-family:' + BODY + ';font-size:16px;line-height:24px;color:' + INK + ';margin:0 0 18px">Thanks, ' + esc(name) + '. ' +
        (updated ? "Here's your updated order." : "We've got your order.") + '</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
        '<tr>' + label("Order #") + '<td style="padding:8px 0;font-family:' + MONO + ';font-size:15px;font-weight:bold;color:' + INK + '">' + esc(number) + '</td></tr>' +
        '<tr>' + label("Shirts") + '<td style="padding:8px 0;font-family:' + BODY + ';font-size:15px;font-weight:bold;color:' + INK + '">' + total + ' total</td></tr>' +
      '</table>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:12px;border-top:1.5px solid ' + LINE + '">' + rows + '</table>' +
      '<div style="font-family:' + BODY + ';font-size:13px;line-height:20px;color:' + MUTED + ';margin-top:20px">To change your order, submit the form again with this same email address.</div>' +
    '</td></tr>' +
  '</table></td></tr></table></div>';
}

/** Trim, cap length, and stop spreadsheet formula injection. */
function clean(v, max) {
  let s = String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function columnLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
