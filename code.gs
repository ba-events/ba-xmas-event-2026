// ============================================================
//  Community Event Ticketing — Google Apps Script
//  Attach to your Google Sheet (Extensions > Apps Script)
//  Run setup() once, then set the onFormSubmit trigger.
//  Deploy as Web App → open the /exec URL as your portal.
// ============================================================

// ── CONFIG — edit before running ────────────────────────
const CFG = {
  EVENT_NAME:  "BA Christmas 2026",
  EVENT_DATE:  "Sat & Sun, 19–20 Dec 2026",
  EVENT_VENUE: "Birla Alokya, Whitefield, Bangalore",

  SCANNER_PIN: "1912",   // ← change this to your 4-digit PIN

  SHEET_REG: "Registrations",
  SHEET_ATT: "Attendance",

  // Column positions in Registrations sheet (1-indexed)
  C_TIMESTAMP: 1,
  C_NAME:      2,
  C_FLAT:      3,
  C_PHONE:     4,
  C_ADULTS:    5,
  C_KIDS:      6,
  C_PAY_URL:   7,
  C_EMAIL:     8,   // optional email
  C_TICKET_ID: 9,
  C_STATUS:      10,  // PENDING | APPROVED | PARTIAL | CHECKED_IN | REJECTED
  C_APPR_BY:     11,
  C_APPR_AT:     12,
  C_QR_URL:      13,
  C_CHECKED_IN:  14,  // number of people checked in so far
  C_REMAINING:   15,  // number still to arrive
  C_EXPERIENCES: 16,  // comma-separated experience choices from registration form

  WA_TEMPLATE: `Hi {NAME} 👋\n\nThank you for registering for *{EVENT_NAME}*! 🎄\n\nYour family is officially part of *BA Christmas 2026* 🎉\n\n📍 {EVENT_VENUE}\n📅 {EVENT_DATE}\n\n👥 Adults: {ADULTS}  |  Kids: {KIDS}\n🏠 Villament: {FLAT}\n🎟 Ticket ID: *{TICKET_ID}*\n\n🎁 *This ticket is your Food Coupon!*\nShow the QR code or Ticket ID at the food counter to be served.\n\n📲 Your QR code (screenshot & keep handy):\n{QR_URL}\n\n🎅 *Merry Christmas!* 🎅\nWishing your family joy, togetherness and a wonderful Christmas!\n\nSee you on 19 Dec 🎊`,
};
// ─────────────────────────────────────────────────────────


// ── WEB APP ENTRY POINT ──────────────────────────────────
function doGet(e) {
  const action = (e.parameter.action || "").toLowerCase();
  const pin    = e.parameter.pin || "";
  const cb     = e.parameter.callback || "";

  let result;
  if (action === "stats")           result = getStats();
  else if (action === "list")       result = { ok: true, rows: getSheetData() };
  else if (action === "listandstats") result = getListAndStats();
  else if (action === "register")   result = registerAttendee(e.parameter);
  else if (action === "like")     result = addLike(e.parameter.tile);
  else if (action === "getlikes") result = getLikeCounts();
  else if (action === "experiences") result = getExperienceList();
  else if (action === "getticket")      result = getTicket(e.parameter.phone, e.parameter.flat);
  else if (action === "hdcsubmit")      result = submitHomeDecorEntry(e.parameter);
  else if (action === "santaregister")  result = submitSantaRegistration(e.parameter);
  else if (action === "santacount")     result = getSantaCount();
  else if (pin !== CFG.SCANNER_PIN) result = { ok: false, message: "Invalid PIN" };
  else if (action === "lookup")  result = checkinTicket(e.parameter.ticket, null);
  else if (action === "checkin") result = checkinTicket(e.parameter.ticket, e.parameter.count);
  else if (action === "approve") result = approveRow(e.parameter.row);
  else if (action === "reject")  result = rejectRow(e.parameter.row);
  else result = { error: "Unknown action" };

  const json = JSON.stringify(result);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}


// ── FORM SUBMIT TRIGGER ──────────────────────────────────
function onFormSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CFG.SHEET_REG);
  const v     = e.namedValues || {};

  const name   = clean(v["Full Name"]                              || "");
  const flat   = clean(v["BA Unit #"]                              || "");
  const phone  = clean(v["Contact number"]                         || "").replace(/\D/g, "");
  const adults = parseInt(v["Number of adults attending"]   || "1", 10) || 1;
  const kids   = parseInt(v["Number of kids attending"]     || "0", 10) || 0;
  const paid   = clean(v["Total Amount Paid"]                      || "");

  if (!name || !flat) return;

  sheet.appendRow([
    new Date(), name, flat, phone, adults, kids,
    paid, "", "PENDING", "", "", "", 0, adults + kids,
  ]);
}


// ── FUNCTIONS CALLED BY google.script.run ────────────────
// All must return plain objects (no ContentService).

function getList() {
  return getSheetData();
}

function getStats() {
  const rows = getSheetData();
  const counts = { PENDING: 0, APPROVED: 0, REJECTED: 0, CHECKED_IN: 0 };
  let totalAdults = 0, totalKids = 0, checkedAdults = 0, checkedKids = 0;
  rows.forEach(r => {
    counts[r.status] = (counts[r.status] || 0) + 1;
    totalAdults += Number(r.adults) || 0;
    totalKids   += Number(r.kids)   || 0;
    if (r.status === "CHECKED_IN") {
      checkedAdults += Number(r.adults) || 0;
      checkedKids   += Number(r.kids)   || 0;
    }
  });
  return { ok: true, counts, totalAdults, totalKids, checkedAdults, checkedKids, total: rows.length };
}

function getListAndStats() {
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get("las_v1");
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  const rows = getSheetData();
  const counts = { PENDING: 0, APPROVED: 0, REJECTED: 0, CHECKED_IN: 0 };
  let totalAdults = 0, totalKids = 0, checkedAdults = 0, checkedKids = 0;
  rows.forEach(r => {
    counts[r.status] = (counts[r.status] || 0) + 1;
    totalAdults += Number(r.adults) || 0;
    totalKids   += Number(r.kids)   || 0;
    if (r.status === "CHECKED_IN") {
      checkedAdults += Number(r.adults) || 0;
      checkedKids   += Number(r.kids)   || 0;
    }
  });
  const result = { ok: true, rows, counts, totalAdults, totalKids, checkedAdults, checkedKids, total: rows.length };
  try { cache.put("las_v1", JSON.stringify(result), 20); } catch(e) {}
  return result;
}

function clearListCache() {
  try { CacheService.getScriptCache().remove("las_v1"); } catch(e) {}
}

function approveRow(rowNum) {
  rowNum = parseInt(rowNum, 10);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_REG);
  if (!sheet) return { ok: false, message: "Sheet not found" };

  const status = String(sheet.getRange(rowNum, CFG.C_STATUS).getValue()).trim();
  const name   = sheet.getRange(rowNum, CFG.C_NAME).getValue();
  const flat   = sheet.getRange(rowNum, CFG.C_FLAT).getValue();
  const phone  = sheet.getRange(rowNum, CFG.C_PHONE).getValue();
  const adults = sheet.getRange(rowNum, CFG.C_ADULTS).getValue();
  const kids   = sheet.getRange(rowNum, CFG.C_KIDS).getValue();
  const email  = String(sheet.getRange(rowNum, CFG.C_EMAIL).getValue() || "").trim();

  // Already approved — return existing ticket for resend
  if (status !== "PENDING") {
    const ticketId = String(sheet.getRange(rowNum, CFG.C_TICKET_ID).getValue() || "");
    const qrUrl    = String(sheet.getRange(rowNum, CFG.C_QR_URL).getValue() || "");
    if (!ticketId) return { ok: false, message: "No ticket found for this registration" };
    const waMsg = buildWaMessage(name, flat, adults, kids, ticketId, qrUrl);
    const waUrl = "https://wa.me/" + internationalPhone(phone) + "?text=" + encodeURIComponent(waMsg);
    return { ok: true, ticketId, qrUrl, name, flat, phone, waUrl };
  }

  const ticketId = generateUniqueTicketId(sheet);
  const qrUrl    = buildQrUrl(ticketId);
  const now      = new Date();

  sheet.getRange(rowNum, CFG.C_TICKET_ID).setValue(ticketId);
  sheet.getRange(rowNum, CFG.C_STATUS).setValue("APPROVED");
  sheet.getRange(rowNum, CFG.C_APPR_BY).setValue("admin");
  sheet.getRange(rowNum, CFG.C_APPR_AT).setValue(now);
  sheet.getRange(rowNum, CFG.C_QR_URL).setValue(qrUrl);

  const waMsg = buildWaMessage(name, flat, adults, kids, ticketId, qrUrl);
  const waUrl = "https://wa.me/" + internationalPhone(phone) + "?text=" + encodeURIComponent(waMsg);

  if (email) sendTicketEmail(email, name, flat, adults, kids, ticketId, qrUrl);
  clearListCache();
  return { ok: true, ticketId, qrUrl, name, flat, phone, waUrl };
}

function registerAttendee(p) {
  const name   = clean(p.name   || "");
  const flat   = normalizeFlat(p.flat || "");
  const phone  = clean(p.phone  || "").replace(/\D/g, "");
  const email  = clean(p.email  || "");
  const adults = parseInt(p.adults || "1", 10) || 1;
  const kids   = parseInt(p.kids   || "0", 10) || 0;
  const utr    = clean(p.utr    || "");
  const experiences = clean(p.experiences || "");
  const amount = (adults * 600) + (kids * 400);

  if (!name)  return { ok: false, message: "Name is required" };
  if (!flat)  return { ok: false, message: "Unit number is required" };
  if (!phone) return { ok: false, message: "Contact number is required" };
  if (!utr)   return { ok: false, message: "Payment reference (UTR) is required" };

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, CFG.SHEET_REG);

    // Idempotency: if this UTR was already registered, return existing data — no new row
    const rows = sheet.getDataRange().getValues().slice(1);
    for (const row of rows) {
      if (String(row[CFG.C_PAY_URL - 1]).includes(utr)) {
        return { ok: true, name: String(row[CFG.C_NAME - 1]), flat: String(row[CFG.C_FLAT - 1]),
                 adults: Number(row[CFG.C_ADULTS - 1]), kids: Number(row[CFG.C_KIDS - 1]),
                 amount, email };
      }
    }

    sheet.appendRow([
      new Date(), name, flat, phone, adults, kids,
      "UTR: " + utr + " | ₹" + amount, email, "", "PENDING", "", "", "", 0, adults + kids, experiences,
    ]);
    SpreadsheetApp.flush();
    clearListCache();
  } catch(e) {
    Logger.log("registerAttendee error: " + e.message);
    return { ok: false, message: "Server error — please try again" };
  }

  return { ok: true, name, flat, adults, kids, amount, email };
}


function rejectRow(rowNum) {
  rowNum = parseInt(rowNum, 10);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_REG);
  if (!sheet) return { ok: false, message: "Sheet not found" };
  sheet.getRange(rowNum, CFG.C_STATUS).setValue("REJECTED");
  sheet.getRange(rowNum, CFG.C_APPR_BY).setValue("admin");
  sheet.getRange(rowNum, CFG.C_APPR_AT).setValue(new Date());
  clearListCache();
  return { ok: true };
}

function checkinTicket(ticketId, countNow) {
  if (!ticketId) return { ok: false, message: "No ticket ID" };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CFG.SHEET_REG);
    if (!sheet) return { ok: false, message: "Sheet not found" };

    const row = findRowBy(sheet, CFG.C_TICKET_ID, ticketId.trim().toUpperCase());
    if (!row) return { ok: false, message: "Ticket not found", ticketId };

    const status    = String(sheet.getRange(row, CFG.C_STATUS).getValue()).trim();
    const name      = sheet.getRange(row, CFG.C_NAME).getValue();
    const flat      = sheet.getRange(row, CFG.C_FLAT).getValue();
    const adults    = Number(sheet.getRange(row, CFG.C_ADULTS).getValue()) || 0;
    const kids      = Number(sheet.getRange(row, CFG.C_KIDS).getValue()) || 0;
    const checkedIn = Number(sheet.getRange(row, CFG.C_CHECKED_IN).getValue()) || 0;
    const total     = adults + kids;
    const remaining = Number(sheet.getRange(row, CFG.C_REMAINING).getValue()) || (total - checkedIn);

    if (status === "PENDING")    return { ok: false, message: "Payment not yet approved", name, flat };
    if (status === "REJECTED")   return { ok: false, message: "Registration was rejected", name, flat };
    if (status === "CHECKED_IN") return { ok: false, duplicate: true, message: "All members already checked in", name, flat, adults, kids, checkedIn, remaining: 0 };

    if (countNow === undefined || countNow === null) {
      return { ok: true, lookup: true, name, flat, adults, kids, total, checkedIn, remaining, ticketId, status };
    }

    countNow = parseInt(countNow, 10) || 0;
    if (countNow <= 0)        return { ok: false, message: "Count must be at least 1" };
    if (countNow > remaining) return { ok: false, message: "Only " + remaining + " still to arrive", name, flat, remaining };

    const newCheckedIn = checkedIn + countNow;
    const newRemaining = total - newCheckedIn;
    const newStatus    = newRemaining <= 0 ? "CHECKED_IN" : "PARTIAL";
    const now          = new Date();

    sheet.getRange(row, CFG.C_STATUS).setValue(newStatus);
    sheet.getRange(row, CFG.C_CHECKED_IN).setValue(newCheckedIn);
    sheet.getRange(row, CFG.C_REMAINING).setValue(newRemaining);

    const attSheet = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), CFG.SHEET_ATT);
    attSheet.appendRow([now, ticketId, name, flat, countNow, newCheckedIn, total, newRemaining]);
    clearListCache();
    return { ok: true, name, flat, adults, kids, total, checkedIn: newCheckedIn, remaining: newRemaining, countNow, ticketId, status: newStatus };

  } finally {
    lock.releaseLock();
  }
}


// ── HELPERS ───────────────────────────────────────────────
function getSheetData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_REG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const ts  = data[i][CFG.C_TIMESTAMP - 1];
    const aat = data[i][CFG.C_APPR_AT - 1];
    rows.push({
      rowNum:    i + 1,
      timestamp: ts  instanceof Date ? ts.toLocaleString()  : String(ts  || ""),
      name:      String(data[i][CFG.C_NAME      - 1] || ""),
      flat:      String(data[i][CFG.C_FLAT      - 1] || ""),
      phone:     String(data[i][CFG.C_PHONE     - 1] || ""),
      adults:    Number(data[i][CFG.C_ADULTS    - 1] || 0),
      kids:      Number(data[i][CFG.C_KIDS      - 1] || 0),
      payUrl:    String(data[i][CFG.C_PAY_URL   - 1] || ""),
      email:     String(data[i][CFG.C_EMAIL     - 1] || ""),
      ticketId:  String(data[i][CFG.C_TICKET_ID - 1] || ""),
      status:    String(data[i][CFG.C_STATUS    - 1] || "").trim(),
      approvedAt:aat instanceof Date ? aat.toLocaleString() : String(aat || ""),
      qrUrl:     String(data[i][CFG.C_QR_URL    - 1] || ""),
      checkedIn:   Number(data[i][CFG.C_CHECKED_IN  - 1] || 0),
      remaining:   Number(data[i][CFG.C_REMAINING   - 1] || 0),
      experiences: String(data[i][CFG.C_EXPERIENCES - 1] || ""),
    });
  }
  return rows;
}

function sendTicketEmail(email, name, flat, adults, kids, ticketId, qrUrl) {
  const firstName = name.split(" ")[0];
  const subject = "Your Ticket — " + CFG.EVENT_NAME;
  const body =
    "Hi " + firstName + ",\n\n" +
    "Your registration for " + CFG.EVENT_NAME + " is confirmed! 🎉\n\n" +
    "📍 " + CFG.EVENT_VENUE + "\n" +
    "📅 " + CFG.EVENT_DATE  + "\n\n" +
    "👥 Adults: " + adults + "  |  Kids: " + kids + "\n" +
    "🏠 Unit: " + flat + "\n" +
    "🎟 Ticket ID: " + ticketId + "\n\n" +
    "Show this QR code at the entrance:\n" + qrUrl + "\n\n" +
    "See you there! 🎄";
  const html =
    "<div style='font-family:sans-serif;max-width:480px;margin:0 auto'>" +
    "<div style='background:#165b33;padding:24px;text-align:center;border-radius:10px 10px 0 0'>" +
    "<div style='font-size:26px;margin-bottom:6px'>🎄</div>" +
    "<h1 style='color:#fff;margin:0 0 8px;font-size:20px;font-weight:800'>BA Christmas 2026</h1>" +
    "<span style='display:inline-block;background:rgba(255,255,255,.22);color:#fff;font-size:10px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;padding:3px 12px;border-radius:99px;border:1px solid rgba(255,255,255,.5)'>🎁 Food Coupon</span></div>" +
    "<div style='background:#faf9f5;padding:24px;border:1px solid #e8e0d4;border-top:none;border-radius:0 0 10px 10px'>" +
    "<p style='margin:0 0 16px;font-size:14px'>Hi <strong>" + firstName + "</strong>,<br>Your registration is confirmed! 🎉</p>" +
    "<table style='width:100%;border-collapse:collapse;font-size:14px'>" +
    "<tr><td style='padding:9px 7px;color:#9ca3af;width:80px;border-bottom:1px solid #f0ebe2'>📍 Venue</td><td style='padding:9px 7px;border-bottom:1px solid #f0ebe2;font-weight:700;color:#111827'>" + CFG.EVENT_VENUE + "</td></tr>" +
    "<tr><td style='padding:9px 7px;color:#9ca3af;border-bottom:1px solid #f0ebe2'>📅 Date</td><td style='padding:9px 7px;border-bottom:1px solid #f0ebe2;font-weight:700;color:#111827'>" + CFG.EVENT_DATE + "</td></tr>" +
    "<tr><td style='padding:9px 7px;color:#9ca3af;border-bottom:1px solid #f0ebe2'>🏠 Villament</td><td style='padding:9px 7px;border-bottom:1px solid #f0ebe2;font-weight:700;color:#111827'>" + flat + "</td></tr>" +
    "<tr><td style='padding:9px 7px;color:#9ca3af;border-bottom:1px solid #f0ebe2'>👥 Guests</td><td style='padding:9px 7px;border-bottom:1px solid #f0ebe2;font-weight:700;color:#111827'>" + adults + " adult(s), " + kids + " kid(s)</td></tr>" +
    "<tr><td style='padding:9px 7px;color:#9ca3af'>🎫 Ticket ID</td><td style='padding:9px 7px;font-weight:700;font-size:18px;color:#165b33;letter-spacing:2px'>" + ticketId + "</td></tr>" +
    "</table>" +
    "<div style='text-align:center;margin:20px 0;padding-top:16px;border-top:1px solid #e8e0d4'>" +
    "<div style='background:#dcfce7;border:1px solid #86efac;border-radius:10px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#14532d;font-weight:600'>🎁 This ticket is your <strong>Food Coupon</strong> — show the QR code or Ticket ID at the food counter to be served</div>" +
    "<img src='" + qrUrl + "' width='200' height='200' style='border:3px solid #165b33;border-radius:12px'/>" +
    "<p style='font-size:12px;color:#6b7280;margin-top:8px'>Show this QR code at the food counter</p></div>" +
    "<p style='font-size:13px;color:#6b7280;text-align:center;margin:0;padding-top:12px;border-top:1px solid #f0ebe2'>See you there! 🎄 Merry Christmas!</p>" +
    "<div style='margin:20px 0 0;padding-top:16px;border-top:1px solid #e8e0d4'>" +
    "<p style='font-size:13px;font-weight:700;color:#374151;margin:0 0 10px'>🎄 Event Programme</p>" +
    "<table style='width:100%;border-collapse:collapse;font-size:12px'>" +
    "<tr style='background:#f0fdf4'><td style='padding:6px 8px;font-weight:700;color:#165b33;white-space:nowrap'>19 Dec Eve</td><td style='padding:6px 8px;color:#374151'>☕ Tea Stall · 🛍️ Christmas Items Stall</td></tr>" +
    "<tr><td style='padding:6px 8px;font-weight:700;color:#165b33;white-space:nowrap'>20 Dec AM</td><td style='padding:6px 8px;color:#374151'>🎭 Cultural Performances · 🎵 Carols &amp; Live Music</td></tr>" +
    "<tr style='background:#f0fdf4'><td style='padding:6px 8px;font-weight:700;color:#165b33;white-space:nowrap'>20 Dec AM</td><td style='padding:6px 8px;color:#374151'>🎅 Meet Santa Claus · 🌟 Tree Lighting</td></tr>" +
    "<tr><td style='padding:6px 8px;font-weight:700;color:#165b33;white-space:nowrap'>20 Dec PM</td><td style='padding:6px 8px;color:#374151'>🎮 Games &amp; Family Activities</td></tr>" +
    "<tr style='background:#f0fdf4'><td style='padding:6px 8px;font-weight:800;color:#165b33;white-space:nowrap'>20 Dec Eve</td><td style='padding:6px 8px;color:#165b33;font-weight:600'>🍔 Food Trucks &amp; Flea Market</td></tr>" +
    "<tr><td style='padding:6px 8px;font-weight:700;color:#165b33;white-space:nowrap'>20 Dec PM</td><td style='padding:6px 8px;color:#374151'>🏆 Prize Distribution &amp; Felicitation</td></tr>" +
    "</table></div>" +
    "</div></div>";
  try {
    MailApp.sendEmail({ to: email, subject: subject, body: body, htmlBody: html });
  } catch(e) {
    Logger.log("Email send failed: " + e.message);
  }
}


function generateUniqueTicketId(sheet) {
  const existing = new Set(
    sheet.getDataRange().getValues().slice(1)
      .map(r => String(r[CFG.C_TICKET_ID - 1]).trim())
      .filter(Boolean)
  );
  let id, attempts = 0;
  do {
    id = generateTicketId();
    attempts++;
  } while (existing.has(id) && attempts < 200);
  return id;
}

function generateTicketId() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let id = "";
  for (let i = 0; i < 4; i++) id += chars[Math.floor(Math.random() * chars.length)];
  return id;
}

function buildQrUrl(ticketId) {
  return "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(ticketId) + "&ecc=M";
}

function buildWaMessage(name, flat, adults, kids, ticketId, qrUrl) {
  return CFG.WA_TEMPLATE
    .replace("{NAME}", name).replace("{FLAT}", flat)
    .replace("{ADULTS}", adults).replace("{KIDS}", kids)
    .replace("{TICKET_ID}", ticketId).replace("{QR_URL}", qrUrl)
    .replace("{EVENT_NAME}", CFG.EVENT_NAME)
    .replace("{EVENT_DATE}", CFG.EVENT_DATE)
    .replace("{EVENT_VENUE}", CFG.EVENT_VENUE);
}

function internationalPhone(phone) {
  const digits = String(phone).replace(/\D/g, "");
  return digits.length === 10 ? "91" + digits : digits;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function findRowBy(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex - 1]).trim() === String(value).trim()) return i + 1;
  }
  return null;
}

function clean(v) {
  return String(Array.isArray(v) ? v[0] : v || "").trim();
}

function normalizeFlat(raw) {
  if (!raw) return "";
  // Strip all whitespace and invisible chars, uppercase
  let s = String(raw).replace(/[\s ​‌‍﻿]/g, "").toUpperCase();
  // Extract leading letter prefix (e.g. "B", "BA"); default to "B" if missing
  const prefixMatch = s.match(/^([A-Z]+)/);
  const prefix = prefixMatch ? prefixMatch[1][0] : "B";
  let rest = s.replace(/^[A-Z]+/, "");
  // If hyphen already present, split on it
  if (rest.includes("-")) {
    const parts = rest.split("-").map(p => p.replace(/\D/g, "")).filter(Boolean);
    if (parts.length >= 2) {
      const tower = parts[0];
      const unit  = parts[parts.length - 1].padStart(2, "0");
      return prefix + tower + "-" + unit;
    }
  }
  // No hyphen — pure digits: last 2 = unit, rest = tower
  const digits = rest.replace(/\D/g, "");
  if (digits.length >= 3) {
    return prefix + digits.slice(0, -2) + "-" + digits.slice(-2).padStart(2, "0");
  }
  if (digits.length === 2) {
    return prefix + digits[0] + "-" + digits[1].padStart(2, "0");
  }
  return prefix + "-" + digits.padStart(2, "0");
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


function getTicket(phone, flat) {
  if (!phone || !flat) return { ok: false, message: "Phone and Villament number are required" };
  const digits    = String(phone).replace(/\D/g, "").slice(-10);
  const flatUpper = normalizeFlat(flat);

  const rows = getSheetData();
  const matches = rows.filter(r =>
    String(r.phone).replace(/\D/g, "").slice(-10) === digits &&
    normalizeFlat(r.flat) === flatUpper &&
    (r.status === "APPROVED" || r.status === "CHECKED_IN" || r.status === "PARTIAL")
  );

  if (!matches.length) {
    // Check if there's a pending/rejected entry so we can give a helpful message
    const anyMatch = rows.find(r =>
      String(r.phone).replace(/\D/g, "").slice(-10) === digits &&
      normalizeFlat(r.flat) === flatUpper
    );
    if (!anyMatch) return { ok: false, message: "No registration found for these details" };
    if (anyMatch.status === "REJECTED") return { ok: false, message: "This registration was not approved" };
    if (anyMatch.status === "PENDING")  return { ok: false, status: "PENDING", message: "Payment verification is pending" };
    return { ok: false, message: "No approved tickets found" };
  }

  const tickets = matches.map(r => ({
    ok:       true,
    name:     r.name,
    flat:     r.flat,
    adults:   r.adults,
    kids:     r.kids,
    amount:   (r.adults * 600) + (r.kids * 400),
    ticketId: r.ticketId,
    qrUrl:    r.qrUrl,
    status:   r.status,
  }));

  return { ok: true, tickets };
}

// Return unique approved villament numbers for the searchable dropdown
function getApprovedFlats() {
  const rows = getSheetData();
  const flats = [...new Set(
    rows
      .filter(r => r.status === "APPROVED" || r.status === "CHECKED_IN" || r.status === "PARTIAL")
      .map(r => String(r.flat).trim().toUpperCase())
      .filter(Boolean)
  )].sort();
  return { ok: true, flats };
}


function getExperienceList() {
  const rows = getSheetData().filter(r => r.status !== "REJECTED" && r.experiences);
  const byExp = {};
  rows.forEach(r => {
    r.experiences.split(",").map(e => e.trim()).filter(Boolean).forEach(exp => {
      if (!byExp[exp]) byExp[exp] = [];
      byExp[exp].push({ name: r.name, flat: r.flat, phone: r.phone });
    });
  });
  return { ok: true, byExperience: byExp, rows: rows.map(r => ({ name: r.name, flat: r.flat, phone: r.phone, experiences: r.experiences })) };
}


function addLike(tile) {
  if (!tile) return { ok: false };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Reactions");
  if (sheet.getLastRow() === 0) sheet.appendRow(["Timestamp", "Tile"]);
  sheet.appendRow([new Date(), String(tile).trim()]);
  return { ok: true, counts: getLikeCounts().counts };
}

function getLikeCounts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reactions");
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, counts: {} };
  const data   = sheet.getDataRange().getValues().slice(1);
  const counts = {};
  data.forEach(r => {
    const tile = String(r[1] || "").trim();
    if (tile) counts[tile] = (counts[tile] || 0) + 1;
  });
  return { ok: true, counts };
}

function submitHomeDecorEntry(p) {
  const name  = clean(p.name  || "");
  const unit  = normalizeFlat(p.unit || "");
  const phone = clean(p.phone || "").replace(/\D/g, "");
  const link  = clean(p.link  || "");

  if (!name)               return { ok: false, message: "Name is required" };
  if (!unit)               return { ok: false, message: "Villament unit is required" };
  if (phone.length !== 10) return { ok: false, message: "Valid 10-digit phone required" };
  if (!link)               return { ok: false, message: "Video link is required" };

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, "HomeDecorEntries");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Unit #", "Phone", "Video Link"]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), name, unit, phone, link]);
  } catch(e) {
    Logger.log("submitHomeDecorEntry error: " + e.message);
    return { ok: false, message: "Server error — please try again" };
  }

  return { ok: true };
}


function submitSantaRegistration(p) {
  const parentname  = clean(p.parentname  || "");
  const unit        = normalizeFlat(p.unit || "");
  const phone       = clean(p.phone       || "").replace(/\D/g, "");
  const kidname     = clean(p.kidname     || "");
  const codename    = clean(p.codename    || "");
  const kidage      = clean(p.kidage      || "");
  const wish        = clean(p.wish        || "");
  const description = clean(p.description || "");

  if (!parentname)         return { ok: false, message: "Parent name is required" };
  if (!unit)               return { ok: false, message: "Villament unit is required" };
  if (phone.length !== 10) return { ok: false, message: "Valid 10-digit phone required" };
  if (!kidname)            return { ok: false, message: "Child's name is required" };
  if (!codename)           return { ok: false, message: "Secret Name (Code Name) is required" };
  if (!kidage)             return { ok: false, message: "Child's age is required" };
  if (!wish)               return { ok: false, message: "Wish is required" };

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, "SantaWishes");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Parent Name", "Unit #", "Phone", "Child Name", "Secret Name", "Age", "Wish", "Details"]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), parentname, unit, phone, kidname, codename, kidage, wish, description]);
  } catch(e) {
    Logger.log("submitSantaRegistration error: " + e.message);
    return { ok: false, message: "Server error — please try again" };
  }

  return { ok: true };
}


// ── ONE-TIME SETUP ────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const reg = getOrCreateSheet(ss, CFG.SHEET_REG);
  if (reg.getLastRow() === 0) {
    reg.appendRow(["Timestamp","Name","Flat","Phone","Adults","Kids","Payment","Email","Ticket ID","Status","Approved By","Approved At","QR URL","Checked In","Remaining","Experiences"]);
    reg.setFrozenRows(1);
  }

  const att = getOrCreateSheet(ss, CFG.SHEET_ATT);
  if (att.getLastRow() === 0) {
    att.appendRow(["Checked-In At","Ticket ID","Name","Flat No.","Count Now","Total Checked In","Total","Remaining"]);
    att.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert("✅ Setup done!\n\nNext steps:\n1. Deploy as Web App\n2. Add onFormSubmit trigger");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎟 Event Tools")
    .addItem("Run Setup", "setup")
    .addSeparator()
    .addItem("Export Attendance CSV", "exportAttendanceCsv")
    .addToUi();
}

function exportAttendanceCsv() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_ATT);
  if (!sheet) { SpreadsheetApp.getUi().alert("Attendance sheet not found."); return; }
  const data = sheet.getDataRange().getValues();
  const csv  = data.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const file = DriveApp.createFile("attendance_" + ts + ".csv", csv, MimeType.CSV);
  SpreadsheetApp.getUi().alert("Saved to Drive: " + file.getName());
}
