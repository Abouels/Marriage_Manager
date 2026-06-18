const SUPPORT_RECIPIENT_EMAIL = "your-support-email@example.com";
const SHEET_NAME = "Support Requests";
const SPREADSHEET_NAME = "Marriage Manager Support";
const TICKET_PREFIX = "MM";
const SUPPORT_MESSAGE_TYPES = ["مشكلة تقنية", "طلب ميزة", "اقتراح تحسين", "ملاحظة عامة"];
const MAX_LENGTHS = {
  name: 120,
  email: 254,
  subject: 160,
  details: 4000,
};

function doPost(e) {
  try {
    const payload = normalizePayload_(JSON.parse((e && e.postData && e.postData.contents) || "{}"));
    validatePayload_(payload);
    validateRecipient_();

    const ticketId = nextTicketId_();
    const createdAt = new Date();
    const sheet = getOrCreateSheet_();

    sheet.appendRow([
      ticketId,
      createdAt,
      safeCell_(payload.type),
      safeCell_(payload.subject),
      safeCell_(payload.name),
      safeCell_(payload.email),
      safeCell_(payload.details),
      safeCell_(payload.appName),
      safeCell_(payload.appVersion),
      safeCell_(payload.projectName),
      safeCell_(payload.sentAt),
    ]);

    MailApp.sendEmail({
      to: SUPPORT_RECIPIENT_EMAIL,
      subject: `[${ticketId}] ${payload.type}: ${payload.subject}`,
      htmlBody: buildOwnerEmail_(ticketId, payload, createdAt),
      replyTo: payload.email,
      name: "Marriage Manager Support",
    });

    MailApp.sendEmail({
      to: payload.email,
      subject: `تم استلام رسالتك رقم ${ticketId}`,
      htmlBody: buildUserEmail_(ticketId, payload),
      name: "Marriage Manager Support",
    });

    return json_({ ok: true, ticketId });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function normalizePayload_(payload) {
  return {
    name: String(payload.name || "").trim(),
    email: String(payload.email || "").trim(),
    type: String(payload.type || "").trim(),
    subject: String(payload.subject || "").trim(),
    details: String(payload.details || "").trim(),
    appName: String(payload.appName || "").trim(),
    appVersion: String(payload.appVersion || "").trim(),
    projectName: String(payload.projectName || "").trim(),
    sentAt: String(payload.sentAt || "").trim(),
  };
}

function validatePayload_(payload) {
  const required = ["name", "email", "type", "subject", "details"];
  required.forEach((key) => {
    if (!payload[key]) {
      throw new Error(`Missing field: ${key}`);
    }
  });
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(payload.email)) {
    throw new Error("Invalid email address");
  }
  if (!SUPPORT_MESSAGE_TYPES.includes(payload.type)) {
    throw new Error("Invalid message type");
  }
  Object.keys(MAX_LENGTHS).forEach((key) => {
    if (payload[key].length > MAX_LENGTHS[key]) {
      throw new Error(`Field is too long: ${key}`);
    }
  });
}

function validateRecipient_() {
  if (
    SUPPORT_RECIPIENT_EMAIL === "your-support-email@example.com" ||
    !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(SUPPORT_RECIPIENT_EMAIL)
  ) {
    throw new Error("Support recipient email is not configured");
  }
}

function nextTicketId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = Number(props.getProperty("ticketCounter") || "0") + 1;
    props.setProperty("ticketCounter", String(current));
    return `${TICKET_PREFIX}-${String(current).padStart(6, "0")}`;
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet_() {
  const ss = getOrCreateSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Ticket ID",
      "Created At",
      "Type",
      "Subject",
      "Name",
      "Email",
      "Details",
      "App Name",
      "App Version",
      "Project Name",
      "Client Sent At",
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("supportSpreadsheetId");
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (err) {
      props.deleteProperty("supportSpreadsheetId");
    }
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty("supportSpreadsheetId", ss.getId());
  return ss;
}

function buildOwnerEmail_(ticketId, payload, createdAt) {
  return `
    <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.7">
      <h2>رسالة دعم جديدة</h2>
      <p><b>رقم المتابعة:</b> ${escapeHtml_(ticketId)}</p>
      <p><b>وقت التسجيل:</b> ${escapeHtml_(createdAt.toLocaleString())}</p>
      <p><b>النوع:</b> ${escapeHtml_(payload.type)}</p>
      <p><b>العنوان:</b> ${escapeHtml_(payload.subject)}</p>
      <p><b>الاسم:</b> ${escapeHtml_(payload.name)}</p>
      <p><b>الإيميل:</b> ${escapeHtml_(payload.email)}</p>
      <p><b>إصدار البرنامج:</b> ${escapeHtml_(payload.appVersion)}</p>
      <p><b>اسم المشروع:</b> ${escapeHtml_(payload.projectName)}</p>
      <hr>
      <p style="white-space:pre-wrap">${escapeHtml_(payload.details)}</p>
    </div>
  `;
}

function buildUserEmail_(ticketId, payload) {
  return `
    <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.7">
      <p>مرحبًا ${escapeHtml_(payload.name)},</p>
      <p>تم استلام رسالتك بنجاح، ورقم المتابعة هو <b>${escapeHtml_(ticketId)}</b>.</p>
      <p><b>نوع الرسالة:</b> ${escapeHtml_(payload.type)}</p>
      <p><b>العنوان:</b> ${escapeHtml_(payload.subject)}</p>
      <p>شكرًا لمساعدتك في تحسين البرنامج.</p>
    </div>
  `;
}

function safeCell_(value) {
  const text = String(value || "");
  return /^[=+\-@\t\r]/.test(text) ? `'${text}` : text;
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
