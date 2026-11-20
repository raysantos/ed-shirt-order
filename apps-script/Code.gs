/**
 * ED Shirt Order — Google Apps Script backend.
 *
 * Paste this into Extensions > Apps Script on a PRIVATE Google Sheet.
 * Each shirt line of an order is one row on the "Orders" tab (sleeve,
 * back design, title on the back, size, quantity). The "Totals" tab
 * tallies shirts by size and by back title so the printer knows what to
 * make. The web app only accepts new orders. It never returns anyone's
 * data, so the sheet stays visible only to people you share it with.
 *
 * Optional: set a team code so only staff who know it can order.
 *   Project Settings > Script Properties > add TEAM_CODE = <your code>
 */

const SIZES = ["XS", "S", "M", "L", "XL", "2XL"];
const SLEEVES = { SS: "Short sleeve", LS: "Long sleeve" };
const DESIGNS = { B: "Red cross badge", C: "Flag badge" };
const ROLES = { RN: "Registered Nurse", HT: "Health Technician", PH: "Physician", SW: "Social Worker", NM: "Nurse Manager" };
const SLEEVE_KEYS = Object.keys(SLEEVES);
const DESIGN_KEYS = Object.keys(DESIGNS);
const ROLE_KEYS = Object.keys(ROLES);
// Columns on the Totals tab: one per sleeve + back design.
const GROUPS = [];
SLEEVE_KEYS.forEach(function (s) { DESIGN_KEYS.forEach(function (d) { GROUPS.push({ sleeve: SLEEVES[s], design: DESIGNS[d] }); }); });

const MAX_QTY = 25;
const MAX_LINES = 30;
const ORDERS = "Orders";
const TOTALS = "Totals";
const HEADERS = ["Order #", "Submitted", "Last updated", "Name", "Email", "Sleeve", "Back design", "Title on back", "Size", "Qty"];
const COL = {};
HEADERS.forEach(function (h, i) { COL[h] = i + 1; });
const EMAIL_SUBJECT = "Your ED shirt order";
const FORM_URL = "https://raysantos.github.io/ed-shirt-order/";

/**
 * Run from the editor (Run > setup) to build or rebuild the tabs.
 * If the Orders tab has an older column layout and holds orders, it is
 * renamed (never deleted) and a fresh Orders tab is created.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ORDERS);
  if (sh) {
    const width = Math.max(sh.getLastColumn(), 1);
    const current = sh.getRange(1, 1, 1, width).getValues()[0].join("|");
    if (current !== HEADERS.join("|")) {
      if (sh.getLastRow() > 1) {
        sh.setName(ORDERS + " (old " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d h.mm a") + ")");
        sh = null;
      } else {
        sh.clear();
        sh.getBandings().forEach(function (b) { b.remove(); });
      }
    }
  }
  if (!sh) sh = ss.insertSheet(ORDERS, 0);
  if (sh.getMaxColumns() < HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);

  let t = ss.getSheetByName(TOTALS);
  if (!t) t = ss.insertSheet(TOTALS);
  t.clear();
  t.getBandings().forEach(function (b) { b.remove(); });
  const cols = GROUPS.length + 2;
  if (t.getMaxColumns() < cols) t.insertColumnsAfter(t.getMaxColumns(), cols - t.getMaxColumns());
  const O = function (h) { const L = columnLetter(COL[h]); return ORDERS + "!$" + L + "$2:$" + L; };
  const groupHead = GROUPS.map(function (g) { return g.sleeve + " · " + g.design; });

  // Section 1: sizes down, sleeve + back design across.
  t.getRange(1, 1, 1, cols).setValues([["Size"].concat(groupHead, ["Total"])]);
  SIZES.forEach(function (sz, i) {
    const r = i + 2;
    const row = [sz].concat(GROUPS.map(function (g) {
      return "=SUMIFS(" + O("Qty") + "," + O("Sleeve") + ",\"" + g.sleeve + "\"," + O("Back design") + ",\"" + g.design + "\"," + O("Size") + ",$A" + r + ")";
    }));
    row.push("=SUM(B" + r + ":" + columnLetter(cols - 1) + r + ")");
    t.getRange(r, 1, 1, cols).setValues([row]);
  });
  const sizeTotal = SIZES.length + 2;
  const tr = ["Total"];
  for (let c = 2; c <= cols; c++) { const L = columnLetter(c); tr.push("=SUM(" + L + "2:" + L + (sizeTotal - 1) + ")"); }
  t.getRange(sizeTotal, 1, 1, cols).setValues([tr]);
  t.getRange(sizeTotal + 1, 1, 1, 2).setValues([["Orders", "=COUNTUNIQUE(" + O("Order #") + ")"]]);

  // Section 2: back titles down, sleeve + back design across.
  const roleHead = sizeTotal + 3;
  t.getRange(roleHead, 1, 1, cols).setValues([["Title on back"].concat(groupHead, ["Total"])]);
  ROLE_KEYS.forEach(function (k, i) {
    const r = roleHead + 1 + i;
    const row = [ROLES[k]].concat(GROUPS.map(function (g) {
      return "=SUMIFS(" + O("Qty") + "," + O("Sleeve") + ",\"" + g.sleeve + "\"," + O("Back design") + ",\"" + g.design + "\"," + O("Title on back") + ",$A" + r + ")";
    }));
    row.push("=SUM(B" + r + ":" + columnLetter(cols - 1) + r + ")");
    t.getRange(r, 1, 1, cols).setValues([row]);
  });

  const def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 2 && def.getLastRow() === 0) ss.deleteSheet(def);

  styleSheets();
}

/** Formats both tabs to match the form. Safe to run again any time. */
function styleSheets() {
  const INK = "#1d2a44", PAPER = "#e7e8eb", ACCENT = "#e0a526", LINE = "#c6c9d1";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ORDERS);
  const t = ss.getSheetByName(TOTALS);
  const n = HEADERS.length;

  // Orders tab
  const maxRows = sh.getMaxRows();
  sh.getRange(1, 1, maxRows, n).setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle");
  sh.getRange(1, 1, 1, n).setBackground(INK).setFontColor("#ffffff").setFontWeight("bold");
  sh.setRowHeight(1, 32);
  sh.getRange(2, COL["Order #"], maxRows - 1, 1).setNumberFormat("@").setHorizontalAlignment("center");
  sh.getRange(2, COL["Submitted"], maxRows - 1, 2).setNumberFormat("mmm d, yyyy h:mm am/pm");
  sh.getRange(1, COL["Size"], maxRows, 2).setHorizontalAlignment("center");
  sh.getRange(2, COL["Qty"], maxRows - 1, 1).setNumberFormat("0").setFontWeight("bold");
  [90, 150, 150, 170, 220, 110, 130, 150, 60, 50].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, maxRows, n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(INK).setFirstRowColor("#ffffff").setSecondRowColor("#f3f4f6");
  sh.getRange(1, COL["Qty"]).setBackground(ACCENT).setFontColor(INK);
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);
  sh.setTabColor(INK);

  // Totals tab
  const cols = GROUPS.length + 2;
  const sizeTotal = SIZES.length + 2;
  const roleHead = sizeTotal + 3;
  const tRows = roleHead + ROLE_KEYS.length;
  t.getRange(1, 1, tRows, cols).setFontFamily("Arial").setFontSize(11).setVerticalAlignment("middle")
    .setFontColor(INK).setBackground(null).setFontWeight("normal");
  [1, roleHead].forEach(function (r) {
    t.getRange(r, 1, 1, cols).setBackground(INK).setFontColor("#ffffff").setFontWeight("bold").setWrap(true);
    t.setRowHeight(r, 40);
  });
  t.getRange(2, 1, SIZES.length, 1).setFontWeight("bold");
  t.getRange(2, cols, SIZES.length, 1).setFontWeight("bold");
  t.getRange(2, 1, SIZES.length, cols).setBorder(null, null, true, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  t.getRange(sizeTotal, 1, 1, cols).setBackground(ACCENT).setFontWeight("bold");
  t.getRange(sizeTotal + 1, 1, 1, 2).setBackground(PAPER);
  t.getRange(roleHead + 1, 1, ROLE_KEYS.length, 1).setFontWeight("bold");
  t.getRange(roleHead + 1, cols, ROLE_KEYS.length, 1).setFontWeight("bold");
  t.getRange(roleHead + 1, 1, ROLE_KEYS.length, cols).setBorder(null, null, true, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  t.getRange(1, 2, tRows, cols - 1).setHorizontalAlignment("center").setNumberFormat("0");
  t.setColumnWidth(1, 150);
  for (let c = 2; c <= cols; c++) t.setColumnWidth(c, 130);
  for (let r = 2; r <= tRows; r++) if (r !== roleHead) t.setRowHeight(r, 28);
  t.setFrozenRows(1);
  if (t.getMaxColumns() > cols) t.deleteColumns(cols + 1, t.getMaxColumns() - cols);
  t.setTabColor(ACCENT);
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

    // Combine identical shirts; drop anything that isn't a real option.
    const merged = {};
    const order = [];
    (Array.isArray(body.lines) ? body.lines : []).slice(0, MAX_LINES).forEach(function (l) {
      if (!l || !SLEEVES[l.sleeve] || !DESIGNS[l.design] || !ROLES[l.role] || SIZES.indexOf(l.size) < 0) return;
      const q = Math.max(0, Math.min(MAX_QTY, Math.floor(Number(l.qty) || 0)));
      if (!q) return;
      const k = [l.sleeve, l.design, l.role, l.size].join("|");
      if (!merged[k]) { merged[k] = { sleeve: l.sleeve, design: l.design, role: l.role, size: l.size, qty: 0 }; order.push(k); }
      merged[k].qty = Math.min(MAX_QTY, merged[k].qty + q);
    });
    const lines = order.map(function (k) { return merged[k]; });
    const total = lines.reduce(function (a, l) { return a + l.qty; }, 0);
    if (!total) return json({ ok: false, error: "no_shirts" });

    // Slow down repeat submissions from the same email.
    const cache = CacheService.getScriptCache();
    const key = "rl_" + Utilities.base64EncodeWebSafe(email);
    if (cache.get(key)) return json({ ok: false, error: "too_fast" });
    cache.put(key, "1", 20);

    lock.waitLock(10000);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
    if (!sh) return json({ ok: false, error: "not_set_up" });
    const header = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].join("|");
    if (header !== HEADERS.join("|")) return json({ ok: false, error: "not_set_up" });

    // One order per email: replace any earlier rows for this email.
    const now = new Date();
    let number = "", submitted = now, updated = false;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const vals = sh.getRange(2, 1, lastRow - 1, COL["Email"]).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][COL["Email"] - 1]).toLowerCase() === email) {
          number = String(vals[i][0]);
          submitted = vals[i][1] || now;
          updated = true;
          sh.deleteRow(i + 2);
        }
      }
    }
    if (!number) number = String(Math.floor(100000 + Math.random() * 900000));

    const rows = lines.map(function (l) {
      return [number, submitted, now, name, email, SLEEVES[l.sleeve], DESIGNS[l.design], ROLES[l.role], l.size, l.qty];
    });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    lock.releaseLock();

    sendConfirmation(email, name, number, lines, total, updated);
    return json({ ok: true, number: number, updated: updated });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: "server" });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function lineLabel(l) {
  return SLEEVES[l.sleeve] + " · " + DESIGNS[l.design] + " · " + ROLES[l.role] + " · " + l.size;
}

function sendConfirmation(email, name, number, lines, total, updated) {
  const text = [
    "Hi " + name + ",",
    "",
    (updated ? "Your shirt order was updated." : "Your shirt order was received."),
    "",
    "Order #" + number,
    lines.map(function (l) { return lineLabel(l) + " × " + l.qty; }).join("\n"),
    "Total: " + total + " shirt" + (total === 1 ? "" : "s"),
    "",
    "To change your order, submit the form again with this same email address:",
    FORM_URL,
  ].join("\n");
  try {
    MailApp.sendEmail({
      to: email,
      subject: EMAIL_SUBJECT + (updated ? " (updated)" : ""),
      body: text,
      htmlBody: confirmationHtml(name, number, lines, total, updated),
      name: "ED Shirt Order",
    });
  } catch (err) {
    console.error("Email failed: " + err);
  }
}

/** Email version of the form's look: grey paper, navy ink, mustard button. */
function confirmationHtml(name, number, lines, total, updated) {
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
  const rows = lines.map(function (l) {
    return '<tr><td style="padding:10px 0;border-bottom:1px solid ' + LINE + ';font-family:' + BODY +
      ';font-size:14px;line-height:20px;color:' + MUTED + '"><b style="color:' + INK + '">' + ROLES[l.role] + '</b><br>' +
      SLEEVES[l.sleeve] + ' · ' + DESIGNS[l.design] + '</td>' +
      '<td style="padding:10px 0;border-bottom:1px solid ' + LINE + ';font-family:' + DISPLAY +
      ';font-size:18px;color:' + INK + '">' + l.size + '</td><td align="right" style="padding:10px 0;border-bottom:1px solid ' +
      LINE + ';font-family:' + MONO + ';font-size:15px;color:' + INK + '">× ' + l.qty + '</td></tr>';
  }).join("");

  return '' +
  '<div style="margin:0;padding:0;background:' + PAPER + '">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + PAPER + '">' +
  '<tr><td align="center" style="padding:32px 16px">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px">' +
    '<tr><td style="padding-bottom:6px;font-family:' + MONO + ';font-size:12px;letter-spacing:2px;text-transform:uppercase;color:' + MUTED + '">Emergency Department · West Orange</td></tr>' +
    '<tr><td style="padding-bottom:22px;font-family:' + DISPLAY + ';font-size:30px;line-height:36px;white-space:nowrap;text-transform:uppercase;color:' + INK + '">Shirt Order Form</td></tr>' +
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
      '<table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:14px"><tr><td style="background:' + ACCENT + ';border-radius:6px">' +
        '<a href="' + FORM_URL + '" style="display:inline-block;padding:12px 20px;font-family:' + DISPLAY + ';font-size:14px;text-transform:uppercase;letter-spacing:0.5px;color:' + INK + ';text-decoration:none">Open the order form</a>' +
      '</td></tr></table>' +
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
