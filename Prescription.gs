/**
 * AUMATIQ — Doctor & Clinic Automation System
 * Part 7: Prescription Module (Prescription.gs) — v1.0 (Phase 3a)
 * ─────────────────────────────────────────────
 * - Doctor প্রেসক্রিপশন লেখে → PDF তৈরি হয় → Drive-এ patient folder-এ সেভ হয়
 * - Prescriptions শীটে রেকর্ড থাকে
 * - Email (GmailApp) স্বয়ংক্রিয়ভাবে পাঠানো হয় (patient-এর email থাকলে)
 * - WhatsApp-এর জন্য wa.me লিংক তৈরি করে ফেরত পাঠানো হয় — frontend সেটা
 *   auto-open করবে, Doctor/Assistant শুধু Send বাটনে ট্যাপ করবে (WhatsApp-এর
 *   নিজস্ব নিয়মে সম্পূর্ণ silent auto-send সম্ভব না, এটাই সবচেয়ে কাছের ফ্রি সমাধান)।
 */

// ───────────────────────── প্রেসক্রিপশন সেভ করা (Doctor-only) ─────────────────────────
/**
 * data = {
 *   patientId, patientName, patientPhone, patientEmail,
 *   diagnosis, advice, nextVisit,
 *   medicines: [ { name, dosage, duration, instructions }, ... ]
 * }
 */
function savePrescription(token, data) {
  const session = requireDoctor(token); // শুধু Doctor prescription লিখতে পারবে

  if (!data || !data.patientId || !data.diagnosis) {
    return { success: false, message: "Patient ID এবং Diagnosis আবশ্যক।" };
  }
  if (!data.medicines || !data.medicines.length) {
    return { success: false, message: "অন্তত একটা Medicine যোগ করো।" };
  }

  const profile = getProfileForEmail_();
  const dn = "Dr. " + (profile.DoctorName || "Asma");
  const now = new Date();

  // ── ১. PrescriptionID জেনারেট করা ──
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prescriptions");
  const existing = sheet.getDataRange().getValues();
  let maxNumber = 0;
  for (let i = 1; i < existing.length; i++) {
    const id = String(existing[i][0]);
    const match = id.match(/RX-(\d+)/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNumber) maxNumber = num;
    }
  }
  const prescriptionId = "RX-" + String(maxNumber + 1).padStart(4, "0");

  // ── ২. Medicines-কে readable টেক্সটে রূপান্তর (শীটের জন্য) ──
  const medicinesText = data.medicines.map(function (m, idx) {
    return (idx + 1) + ". " + m.name + " — " + (m.dosage || "") + " — " + (m.duration || "") +
      (m.instructions ? " (" + m.instructions + ")" : "");
  }).join("\n");

  // ── ৩. PDF তৈরি করা ──
  let fileLink = "";
  try {
    const pdfBlob = generatePrescriptionPdf_(data, profile, prescriptionId, dn, now);
    const folder = getPatientFolder(data.patientId);
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileLink = file.getUrl();
  } catch (e) {
    Logger.log("PDF generation error: " + e.message);
    // PDF ফেল করলেও prescription record সেভ হবে, শুধু fileLink খালি থাকবে
  }

  // ── ৪. Prescriptions শীটে সেভ করা ──
  sheet.appendRow([
    prescriptionId,
    data.patientId,
    data.patientName || "",
    now,
    data.diagnosis,
    medicinesText,
    data.advice || "",
    data.nextVisit || "",
    dn,
    fileLink,
  ]);

  // ── ৫. Email পাঠানো (patient-এর email থাকলে) ──
  if (data.patientEmail) {
    try {
      sendPrescriptionEmail_(data.patientEmail, data.patientName, data.patientId, data, fileLink, dn, profile);
    } catch (e) {
      Logger.log("Prescription email error: " + e.message);
    }
  }

  // ── ৬. WhatsApp লিংক তৈরি করে ফেরত পাঠানো (frontend auto-open করবে) ──
  let whatsappLink = "";
  if (data.patientPhone) {
    whatsappLink = buildPrescriptionWhatsAppLink_(data.patientPhone, data.patientName, prescriptionId, data.diagnosis, fileLink, dn);
  }

  return {
    success: true,
    message: "Prescription সেভ হয়েছে।",
    prescriptionId: prescriptionId,
    fileLink: fileLink,
    whatsappLink: whatsappLink,
    emailSent: !!data.patientEmail,
  };
}

// ───────────────────────── PDF তৈরি করার হেল্পার ─────────────────────────
function generatePrescriptionPdf_(data, profile, prescriptionId, doctorName, dateObj) {
  const clinicName = profile.ClinicName || "Dr. Asma Clinic";
  const specialty = profile.Specialty || "Medical Clinic";
  const address = profile.ClinicAddress || "Dhaka, Bangladesh";
  const phone = profile.ContactPhone || "";
  const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone() || "GMT+6", "dd MMMM, yyyy");

  const medRows = data.medicines.map(function (m, idx) {
    return (
      '<tr>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #ddd;font-size:13px;">' + (idx + 1) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #ddd;font-size:13px;font-weight:700;">' + (m.name || "") + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #ddd;font-size:13px;">' + (m.dosage || "") + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #ddd;font-size:13px;">' + (m.duration || "") + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #ddd;font-size:12px;color:#555;">' + (m.instructions || "—") + '</td>' +
      '</tr>'
    );
  }).join("");

  const html =
    '<html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;padding:0;margin:0;color:#1A1A2E;">' +
    '<div style="padding:36px 44px;">' +
      '<table width="100%" style="border-bottom:3px solid #0EA5E9;padding-bottom:16px;margin-bottom:24px;">' +
        '<tr>' +
        '<td><div style="font-size:20px;font-weight:800;">' + clinicName + '</div>' +
        '<div style="font-size:12px;color:#0EA5E9;font-weight:600;">' + specialty + '</div></td>' +
        '<td align="right"><div style="font-size:11px;color:#666;">' + address + '</div>' +
        '<div style="font-size:11px;color:#666;">' + phone + '</div></td>' +
        '</tr>' +
      '</table>' +
      '<table width="100%" style="margin-bottom:20px;">' +
        '<tr>' +
        '<td><div style="font-size:11px;color:#888;text-transform:uppercase;">Prescription ID</div>' +
        '<div style="font-size:14px;font-weight:700;color:#0EA5E9;">' + prescriptionId + '</div></td>' +
        '<td><div style="font-size:11px;color:#888;text-transform:uppercase;">Date</div>' +
        '<div style="font-size:14px;font-weight:700;">' + dateStr + '</div></td>' +
        '<td><div style="font-size:11px;color:#888;text-transform:uppercase;">Patient</div>' +
        '<div style="font-size:14px;font-weight:700;">' + (data.patientName || "") + ' (' + data.patientId + ')</div></td>' +
        '</tr>' +
      '</table>' +
      '<div style="margin-bottom:20px;">' +
        '<div style="font-size:12px;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:6px;">Diagnosis</div>' +
        '<div style="font-size:14px;background:#F5F9FC;padding:12px 14px;border-radius:8px;">' + data.diagnosis + '</div>' +
      '</div>' +
      '<div style="margin-bottom:20px;">' +
        '<div style="font-size:12px;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:8px;">℞ Medicines</div>' +
        '<table width="100%" style="border-collapse:collapse;">' +
        '<tr style="background:#0EA5E9;color:#fff;"><td style="padding:8px 10px;font-size:12px;">#</td><td style="padding:8px 10px;font-size:12px;">Medicine</td><td style="padding:8px 10px;font-size:12px;">Dosage</td><td style="padding:8px 10px;font-size:12px;">Duration</td><td style="padding:8px 10px;font-size:12px;">Instructions</td></tr>' +
        medRows +
        '</table>' +
      '</div>' +
      (data.advice ? '<div style="margin-bottom:24px;"><div style="font-size:12px;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:6px;">Advice</div><div style="font-size:13px;">' + data.advice + '</div></div>' : '') +
      (data.nextVisit ? '<div style="margin-bottom:24px;"><div style="font-size:12px;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:6px;">Next Visit</div><div style="font-size:13px;font-weight:700;color:#0EA5E9;">' + data.nextVisit + '</div></div>' : '') +
      '<div style="margin-top:50px;border-top:1px solid #ddd;padding-top:16px;text-align:right;">' +
        '<div style="font-size:14px;font-weight:700;">' + doctorName + '</div>' +
        '<div style="font-size:11px;color:#888;">' + specialty + '</div>' +
      '</div>' +
      '<div style="margin-top:30px;text-align:center;font-size:10px;color:#aaa;">Generated via AUMATIQ — Automate. Scale. Dominate.</div>' +
    '</div>' +
    '</body></html>';

  const blob = Utilities.newBlob(html, "text/html", "Prescription_" + prescriptionId + ".html");
  const pdfBlob = blob.getAs("application/pdf");
  pdfBlob.setName("Prescription_" + prescriptionId + ".pdf");
  return pdfBlob;
}

// ───────────────────────── WhatsApp লিংক তৈরি করা ─────────────────────────
function buildPrescriptionWhatsAppLink_(phone, patientName, prescriptionId, diagnosis, fileLink, doctorName) {
  let cleanPhone = String(phone).replace(/[^0-9]/g, "");
  if (cleanPhone.length === 11 && cleanPhone.indexOf("0") === 0) {
    cleanPhone = "88" + cleanPhone; // বাংলাদেশি নম্বরে কান্ট্রি কোড যোগ
  }
  const msg =
    "প্রিয় " + (patientName || "") + ",\n\n" +
    doctorName + " আপনার জন্য একটা নতুন Prescription লিখেছেন।\n" +
    "Prescription ID: " + prescriptionId + "\n" +
    "Diagnosis: " + diagnosis + "\n\n" +
    (fileLink ? "সম্পূর্ণ Prescription (PDF): " + fileLink + "\n\n" : "") +
    "ধন্যবাদান্তে,\n" + doctorName;

  return "https://wa.me/" + cleanPhone + "?text=" + encodeURIComponent(msg);
}

// ───────────────────────── Prescription Email পাঠানো ─────────────────────────
function sendPrescriptionEmail_(patientEmail, patientName, patientId, data, fileLink, doctorName, profile) {
  const t = EMAIL_THEME;
  const cn = profile.ClinicName || t.clinicName;

  const medRowsHtml = data.medicines.map(function (m) {
    return infoRow_("💊", m.name, (m.dosage || "") + " — " + (m.duration || ""));
  }).join("");

  const html =
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:' + t.bgOuter + ';font-family:Arial,Helvetica,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" style="padding:32px 20px;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;border-radius:16px;overflow:hidden;box-shadow:0 4px 40px rgba(15,10,26,0.12);">' +
    '<tr><td style="padding:0;">' + emailHeader_(cn, profile.Specialty || "Medical Clinic") + '</td></tr>' +
    '<tr><td style="background:linear-gradient(135deg,#0A2A3A,#0F0A1A);padding:40px 40px 32px;text-align:center;">' +
    '<div style="display:inline-block;width:68px;height:68px;background:rgba(14,165,233,0.18);border:2px solid rgba(14,165,233,0.5);border-radius:50%;font-size:32px;line-height:68px;text-align:center;margin-bottom:18px;">℞</div>' +
    '<h1 style="font-family:Arial,sans-serif;font-size:24px;font-weight:700;color:#FFFFFF;margin:0 0 10px;">Your Prescription is Ready</h1>' +
    '<p style="font-family:Arial,sans-serif;font-size:15px;color:#BAE6FD;margin:0;">Written by <strong style="color:#FFFFFF;">' + doctorName + '</strong></p>' +
    '</td></tr>' +
    '<tr><td style="background:#FFFFFF;padding:32px 40px;">' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">🩺 Diagnosis</h2>' +
    '<p style="font-family:Arial,sans-serif;font-size:14px;color:' + t.textBody + ';">' + data.diagnosis + '</p>' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:20px 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">💊 Medicines</h2>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' + medRowsHtml + '</table>' +
    '</td></tr>' +
    '<tr><td style="background:#FFFFFF;padding:0 40px 32px;text-align:center;">' +
    '<a href="' + (fileLink || "#") + '" style="display:inline-block;background:linear-gradient(135deg,#0EA5E9,#7C3AED);color:#FFFFFF;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:14px 32px;border-radius:10px;text-decoration:none;">℞ View Full Prescription (PDF) →</a>' +
    '</td></tr>' +
    '<tr><td style="padding:0;">' + emailFooter_(profile.ContactEmail || t.clinicEmail, profile.ContactPhone, profile.ClinicAddress) + '</td></tr>' +
    '</table></td></tr></table></body></html>';

  GmailApp.sendEmail(
    patientEmail,
    "℞ Your Prescription — " + cn,
    "Your prescription from " + cn + " (Dr. " + doctorName + ") is ready. Diagnosis: " + data.diagnosis,
    { htmlBody: html, name: cn, replyTo: profile.ContactEmail || t.clinicEmail }
  );
}

// ───────────────────────── পেশেন্টের সব Prescription History আনা (Doctor/Assistant) ─────────────────────────
function getPrescriptionsForPatient(token, patientId) {
  requireDoctorOrReceptionist(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prescriptions");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === patientId) {
      const obj = {};
      headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
      list.push(obj);
    }
  }
  list.reverse(); // সবচেয়ে নতুনটা আগে
  return { success: true, prescriptions: list };
}

// ───────────────────────── পেশেন্ট নিজের Prescription দেখবে (Patient Portal-এর জন্য, Phase 3b-তে ব্যবহার হবে) ─────────────────────────
function getMyPrescriptions(token) {
  const session = requirePatient(token);
  return getPrescriptionsForPatientInternal_(session.identifier);
}
function getPrescriptionsForPatientInternal_(patientId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prescriptions");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === patientId) {
      const obj = {};
      headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
      list.push(obj);
    }
  }
  list.reverse();
  return { success: true, prescriptions: list };
}

// ───────────────────────── AUTO-SETUP: প্রয়োজনীয় Sheet/Tab নিজে থেকে তৈরি করা ─────────────────────────
/**
 * এই ফাংশনটা একবার Apps Script editor থেকে ম্যানুয়ালি Run করলেই হবে।
 * এটা চেক করে দেখে কোন Tab আগে থেকে আছে — থাকলে কিছুই করে না (পুরনো ডেটা অক্ষত থাকে)।
 * না থাকলে headers + styling সহ নতুন Tab বানিয়ে দেয়।
 * নতুন কোনো ফিচারের জন্য নতুন Sheet লাগলে এখানেই যোগ করা হবে — সবসময় এই একটাই
 * ফাংশন Run করলেই সব আপডেট হয়ে যাবে।
 */
function setupAumatiqSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requiredSheets = {
    "Prescriptions": ["PrescriptionID", "PatientID", "PatientName", "Date", "Diagnosis", "Medicines", "Advice", "NextVisit", "CreatedBy", "FileLink"]
  };

  const createdList = [];
  const skippedList = [];

  Object.keys(requiredSheets).forEach(function (sheetName) {
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      skippedList.push(sheetName); // আগে থেকেই আছে — কিছুই বদলানো হয়নি
      return;
    }

    const headers = requiredSheets[sheetName];
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#0D1117")
      .setFontColor("#F8F9FF");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    createdList.push(sheetName);
  });

  const message =
    "✅ Setup সম্পন্ন।\n\n" +
    "নতুন তৈরি হলো: " + (createdList.length ? createdList.join(", ") : "কোনোটাই না — সব আগে থেকেই ছিল।") + "\n" +
    "আগে থেকেই ছিল (স্পর্শ করা হয়নি): " + (skippedList.length ? skippedList.join(", ") : "কোনোটাই না।");

  Logger.log(message);
  return message;
}

