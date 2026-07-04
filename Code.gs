/**
 * AUMATIQ — Doctor & Clinic Automation System
 * Part 1: Database Auto-Setup Script
 * ─────────────────────────────────────────────
 * এই script রান করলে পুরো ডেটাবেস স্ট্রাকচার
 * (৭টা Tab, কলাম হেডার, sample category data,
 * dropdown validation, frozen header row)
 * স্বয়ংক্রিয়ভাবে তৈরি হয়ে যাবে।
 *
 * চালানোর নিয়ম: নিচে setupDatabase() ফাংশনটা রান করো।
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ১. সব Tab তৈরি করা (যদি আগে থেকে থাকে, পুরনোটা ক্লিয়ার করে আবার বানানো হবে)
  const categoriesSheet = getOrCreateSheet(ss, "Categories");
  const patientsSheet = getOrCreateSheet(ss, "Patients");
  const appointmentsSheet = getOrCreateSheet(ss, "Appointments");
  const testRecordsSheet = getOrCreateSheet(ss, "TestRecords");
  const financeSheet = getOrCreateSheet(ss, "Finance");
  const doctorProfileSheet = getOrCreateSheet(ss, "DoctorProfile");
  const settingsSheet = getOrCreateSheet(ss, "Settings");

  // ২. প্রতিটা Tab সেটআপ করা — Categories সবার আগে (অন্য Tab এর dropdown এর উপর নির্ভর করে)
  setupCategoriesTab(categoriesSheet);
  setupPatientsTab(patientsSheet, categoriesSheet);
  setupAppointmentsTab(appointmentsSheet, categoriesSheet);
  setupTestRecordsTab(testRecordsSheet, categoriesSheet);
  setupFinanceTab(financeSheet, categoriesSheet);
  setupDoctorProfileTab(doctorProfileSheet);
  setupSettingsTab(settingsSheet);

  // ৩. ডিফল্ট "Sheet1" থাকলে ডিলিট করা (যদি থাকে এবং আমাদের বানানো ৭টার বাইরে হয়)
  removeDefaultEmptySheet(ss);

  try {

  SpreadsheetApp.getUi().alert(
    "✅ ডেটাবেস সেটআপ সম্পন্ন!\n\n" +
    "সব Tab, কলাম এবং Dropdown তৈরি হয়েছে।"
  );

} catch (err) {

  Logger.log("Database setup completed successfully.");

}
}

// ───────────────────────── HELPER: Tab তৈরি/পরিষ্কার করা ─────────────────────────
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear(); // আগে থেকে কিছু থাকলে পরিষ্কার করে নতুন করে বানানো হচ্ছে
    sheet.clearFormats();
  }
  return sheet;
}

// ───────────────────────── ডিফল্ট ফাঁকা Sheet ডিলিট করা ─────────────────────────
function removeDefaultEmptySheet(ss) {
  const defaultNames = ["Sheet1", "Sheet 1"];
  defaultNames.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

// ───────────────────────── ১. CATEGORIES TAB (v1.3 — IsActive কলাম যুক্ত, Soft Delete) ─────────────────────────
function setupCategoriesTab(sheet) {
  const headers = ["CategoryGroup", "Value", "IsSystemDefault", "IsActive", "AddedBy", "DateAdded"];
  const today = new Date();

  const data = [
    ["AppointmentStatus", "Pending", "TRUE", "TRUE", "System Default", today],
    ["AppointmentStatus", "Confirmed", "TRUE", "TRUE", "System Default", today],
    ["AppointmentStatus", "Completed", "TRUE", "TRUE", "System Default", today],
    ["AppointmentStatus", "Cancelled", "TRUE", "TRUE", "System Default", today],
    ["Gender", "Male", "TRUE", "TRUE", "System Default", today],
    ["Gender", "Female", "TRUE", "TRUE", "System Default", today],
    ["Gender", "Other", "TRUE", "TRUE", "System Default", today],
    ["FinanceType", "Income", "FALSE", "TRUE", "System Default", today],
    ["FinanceType", "Expense", "FALSE", "TRUE", "System Default", today],
    ["FinanceCategory", "Consultation Fee", "FALSE", "TRUE", "System Default", today],
    ["FinanceCategory", "Test Fee", "FALSE", "TRUE", "System Default", today],
    ["FinanceCategory", "Rent", "FALSE", "TRUE", "System Default", today],
    ["FinanceCategory", "Supplies", "FALSE", "TRUE", "System Default", today],
    ["TestName", "Blood Test", "FALSE", "TRUE", "System Default", today],
    ["TestName", "X-Ray", "FALSE", "TRUE", "System Default", today],
    ["TestName", "Ultrasound", "FALSE", "TRUE", "System Default", today],
    ["TestName", "ECG", "FALSE", "TRUE", "System Default", today],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

// ───────────────────────── হেল্পার: Categories Tab থেকে একটা গ্রুপের ACTIVE ভ্যালুগুলোর লিস্ট বের করা (v1.3) ─────────────────────────
function getActiveCategoryValues(categoriesSheet, groupName) {
  const data = categoriesSheet.getDataRange().getValues();
  const headers = data[0];
  const groupCol = headers.indexOf("CategoryGroup");
  const valueCol = headers.indexOf("Value");
  const activeCol = headers.indexOf("IsActive");

  const values = [];
  for (let i = 1; i < data.length; i++) {
    const isActive = String(data[i][activeCol]).toUpperCase() === "TRUE";
    if (data[i][groupCol] === groupName && isActive) {
      values.push(data[i][valueCol]);
    }
  }

  if (values.length === 0) {
    throw new Error("Category group-এ কোনো active ভ্যালু পাওয়া যায়নি: " + groupName);
  }

  return values;
}

// ───────────────────────── ২. PATIENTS TAB ─────────────────────────
function setupPatientsTab(sheet, categoriesSheet) {
  const headers = ["PatientID", "FullName", "Phone", "Age", "Gender", "Address", "RegisteredDate", "LastVisit", "Notes", "Email"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  const genderValues = getActiveCategoryValues(categoriesSheet, "Gender");
  const genderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(genderValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("E2:E1000").setDataValidation(genderRule);

  sheet.autoResizeColumns(1, headers.length);
}

// ───────────────────────── ৩. APPOINTMENTS TAB ─────────────────────────
function setupAppointmentsTab(sheet, categoriesSheet) {
  const headers = ["AppointmentID", "PatientID", "Date", "TimeSlot", "Reason", "Status", "CreatedAt"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  // Status কলাম (F) এ Dropdown — Categories Tab থেকে AppointmentStatus
  const statusValues = getActiveCategoryValues(categoriesSheet, "AppointmentStatus");
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("F2:F1000").setDataValidation(statusRule);

  sheet.autoResizeColumns(1, headers.length);
}

// ───────────────────────── ৪. TESTRECORDS TAB ─────────────────────────
function setupTestRecordsTab(sheet, categoriesSheet) {
  const headers = ["RecordID", "PatientID", "TestName", "TestDate", "FileLink", "UploadedBy", "Status", "ApprovedBy", "ApprovedDate", "Remarks"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  // TestName কলাম (C) এ Dropdown — Categories Tab থেকে
  const testNameValues = getActiveCategoryValues(categoriesSheet, "TestName");
  const testNameRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(testNameValues, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("C2:C1000").setDataValidation(testNameRule);

  // Status কলাম (G) — এটা workflow-specific, Categories Tab-এ রাখা হয়নি (এডিট করার প্রয়োজন নেই)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending Approval", "Approved", "Rejected"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("G2:G1000").setDataValidation(statusRule);

  sheet.autoResizeColumns(1, headers.length);
}

// ───────────────────────── ৫. FINANCE TAB ─────────────────────────
function setupFinanceTab(sheet, categoriesSheet) {
  const headers = ["EntryID", "Date", "Type", "Category", "PatientID", "Amount", "Notes"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  // Type কলাম (C) — FinanceType
  const typeValues = getActiveCategoryValues(categoriesSheet, "FinanceType");
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("C2:C1000").setDataValidation(typeRule);

  const categoryValues = getActiveCategoryValues(categoriesSheet, "FinanceCategory");
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryValues, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("D2:D1000").setDataValidation(categoryRule);

  sheet.autoResizeColumns(1, headers.length);
}

// ───────────────────────── ৬. DOCTORPROFILE TAB ─────────────────────────
function setupDoctorProfileTab(sheet) {
  const fields = [
    ["DoctorName", ""],
    ["Specialty", ""],
    ["Bio", ""],
    ["PhotoURL", ""],
    ["ClinicName", ""],
    ["ClinicAddress", ""],
    ["WorkingHours", ""],
    ["ServicesList", ""],
    ["ContactPhone", ""],
    ["ContactEmail", ""],
    ["Currency", "BDT"],
  ];

  sheet.getRange(1, 1, fields.length, 2).setValues(fields);
  sheet.getRange(1, 1, fields.length, 1).setFontWeight("bold").setBackground("#F3F1FB");
  sheet.autoResizeColumns(1, 2);
}

// ───────────────────────── ৭. SETTINGS TAB (v1.3 — Clinic Hours যুক্ত) ─────────────────────────
function setupSettingsTab(sheet) {
  const fields = [
    ["ClinicGmailAccount", ""],
    // v2.0: username আর লাগবে না — Login screen থেকে শুধু Role (Doctor/Assistant) + Password।
    // ⚠️ প্রথম লগইনের পরপরই Settings & Categories ট্যাব থেকে এই default password দুটো বদলে ফেলো।
    ["DoctorPassword", "doctor123"],
    ["ReceptionistPassword", "assist123"],
    ["ClinicTimezone", "Asia/Dhaka"],
    ["SlotDuration", "20"],
    ["WorkingDays", "Sat,Sun,Mon,Tue,Wed"],
    ["OpeningTime", "09:00"],
    ["ClosingTime", "20:00"],
  ];

  sheet.getRange(1, 1, fields.length, 2).setValues(fields);
  sheet.getRange(1, 1, fields.length, 1).setFontWeight("bold").setBackground("#F3F1FB");
  sheet.autoResizeColumns(1, 2);
}


// ============================================================
// AUMATIQ — Doctor Automation System
// Part 5 FINAL: Public Website Backend
// clinic.drasma@gmail.com | Female Doctor Theme
// ============================================================
// PASTE METHOD:
// Code.gs খুলুন → Part 1-4 এর সব code এর নিচে
// এই পুরো block paste করুন। কিছু মুছবেন না।
// ============================================================


// ─────────────────────────────────────────────────────────────
// 5.1  doGet() — Web App URL Router
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';

  // ── FIX (v2.0): doctor dashboard আগে কখনো serve হতো না ──
  // ?page=dashboard দিলে এখন DoctorDashboard.html serve হবে।
  // Apps Script Editor-এ HTML file-এর নাম অবশ্যই ঠিক "DoctorDashboard" (case-sensitive) হতে হবে।
  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('DoctorDashboard')
      .setTitle('Clinic Dashboard — Sign In')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  var template = HtmlService.createTemplateFromFile('PublicWebsite');
  template.currentPage     = page;
  template.patientIdParam  = (e && e.parameter && e.parameter.pid)  ? e.parameter.pid  : '';

  return template.evaluate()
    .setTitle('Dr. Asma — Clinic & Women\'s Health')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// ─────────────────────────────────────────────────────────────
// 5.2  getDoctorProfile()
//      DoctorProfile Tab → Field/Value pairs → public website
// ─────────────────────────────────────────────────────────────
function getDoctorProfile() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DoctorProfile');
    if (!sheet) return { success: false, error: 'DoctorProfile sheet not found.' };

    var data    = sheet.getDataRange().getValues();
    var profile = {};
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0] !== 'Field') {
        profile[String(data[i][0]).trim()] = data[i][1] !== undefined ? String(data[i][1]) : '';
      }
    }
    return { success: true, profile: profile };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.3  getPublicBookingSlots(dateStr)
//      Settings Tab → working days, hours, slot duration
//      Appointments Tab → already booked slots (exclude)
//      Returns available slot list for chosen date
// ─────────────────────────────────────────────────────────────
function getPublicBookingSlots(dateStr) {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    var apptSheet     = ss.getSheetByName('Appointments');
    if (!settingsSheet) return { success: false, error: 'Settings sheet not found.' };

    // Read Settings
    var sData    = settingsSheet.getDataRange().getValues();
    var settings = {};
    for (var i = 0; i < sData.length; i++) {
      if (sData[i][0]) settings[String(sData[i][0]).trim()] = sData[i][1];
    }

    var workingDays  = String(settings['WorkingDays']  || 'Sat,Sun,Mon,Tue,Wed').split(',').map(function(d){ return d.trim(); });
    var openingTime  = String(settings['OpeningTime']  || '09:00');
    var closingTime  = String(settings['ClosingTime']  || '20:00');
    var slotDuration = parseInt(settings['SlotDuration'] || '20', 10);

    // Check if selected date is a working day
    var targetDate = new Date(dateStr);
    var dayNames   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    var dayName    = dayNames[targetDate.getDay()];
    if (workingDays.indexOf(dayName) === -1) {
      return { success: true, slots: [], message: 'Clinic is closed on this day.' };
    }

    // Generate all possible slots
    var openParts  = openingTime.split(':').map(Number);
    var closeParts = closingTime.split(':').map(Number);
    var openMin    = openParts[0] * 60 + openParts[1];
    var closeMin   = closeParts[0] * 60 + closeParts[1];
    var allSlots   = [];

    for (var m = openMin; m + slotDuration <= closeMin; m += slotDuration) {
      var hh     = Math.floor(m / 60);
      var mm     = m % 60;
      var period = hh < 12 ? 'AM' : 'PM';
      var hh12   = hh === 0 ? 12 : (hh > 12 ? hh - 12 : hh);
      allSlots.push(
        (hh12 < 10 ? '0' : '') + hh12 + ':' + (mm < 10 ? '0' : '') + mm + ' ' + period
      );
    }

    // Get already booked slots for this date
    var bookedSlots = [];
    if (apptSheet) {
      var apptData = apptSheet.getDataRange().getValues();
      var tz = Session.getScriptTimeZone();
      for (var r = 1; r < apptData.length; r++) {
        if (!apptData[r][2]) continue;
        var rowDateStr = Utilities.formatDate(new Date(apptData[r][2]), tz, 'yyyy-MM-dd');
        var status     = String(apptData[r][5] || '');
        if (rowDateStr === dateStr && status !== 'Cancelled') {
          bookedSlots.push(String(apptData[r][3] || '').trim());
        }
      }
    }

    // Available = all minus booked
    var available = allSlots.filter(function(s) { return bookedSlots.indexOf(s) === -1; });
    return { success: true, slots: available };

  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.4  submitPublicBooking(formData)
//      New patient  → create Patient ID → add to Patients tab
//      Return visit → match by phone → reuse existing ID
//      Always       → create Appointment row
//      Auto-trigger → sendConfirmationEmail() if email provided
// ─────────────────────────────────────────────────────────────
function submitPublicBooking(formData) {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var patientsSheet = ss.getSheetByName('Patients');
    var apptSheet     = ss.getSheetByName('Appointments');
    if (!patientsSheet || !apptSheet) return { success: false, error: 'Database sheet not found.' };

    // Sanitise inputs
    var fullName = String(formData.fullName || '').trim();
    var phone    = String(formData.phone    || '').trim();
    var age      = String(formData.age      || '').trim();
    var gender   = String(formData.gender   || '').trim();
    var reason   = String(formData.reason   || '').trim();
    var date     = String(formData.date     || '').trim();
    var timeSlot = String(formData.timeSlot || '').trim();
    var email    = String(formData.email    || '').trim(); // optional

    if (!fullName || !phone || !date || !timeSlot) {
      return { success: false, error: 'Name, phone number, date, and time are required.' };
    }

    // ── Check if patient exists by phone ──
    var patientsData     = patientsSheet.getDataRange().getValues();
    var existingPatientId = null;
    // Patients headers:
    // A=PatientID B=FullName C=Phone D=Age E=Gender F=Address
    // G=RegisteredDate H=LastVisit I=Notes J=Email
    for (var i = 1; i < patientsData.length; i++) {
      if (String(patientsData[i][2] || '').trim() === phone) {
        existingPatientId = patientsData[i][0];
        // Update Email column (J = index 9) if new email given and was empty
        if (email && !patientsData[i][9]) {
          patientsSheet.getRange(i + 1, 10).setValue(email);
        }
        break;
      }
    }

    var patientId    = '';
    var isNewPatient = false;
    var tz           = Session.getScriptTimeZone();

    if (existingPatientId) {
      patientId = existingPatientId;
    } else {
      // New patient → generate ID
      isNewPatient = true;
      var year      = new Date().getFullYear();
      var serial    = patientsData.length; // rows including header; gives unique serial
      patientId     = 'PT-' + year + '-' + ('000' + serial).slice(-4);

      patientsSheet.appendRow([
        patientId,       // A PatientID
        fullName,        // B FullName
        phone,           // C Phone
        age,             // D Age
        gender,          // E Gender
        '',              // F Address
        Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), // G RegisteredDate
        date,            // H LastVisit (= booking date)
        '',              // I Notes
        email            // J Email
      ]);
    }

    // ── Create Appointment row ──
    var apptData      = apptSheet.getDataRange().getValues();
    var apptSerial    = apptData.length;
    var appointmentId = 'APT-' + ('000' + apptSerial).slice(-4);
    var now           = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    // Appointments headers:
    // A=AppointmentID B=PatientID C=Date D=TimeSlot E=Reason F=Status G=CreatedAt
    apptSheet.appendRow([
      appointmentId,
      patientId,
      date,
      timeSlot,
      reason,
      'Pending',
      now
    ]);

    // ── Send confirmation email if email provided ──
    if (email) {
      try {
        sendConfirmationEmail(email, fullName, patientId, appointmentId, date, timeSlot, reason);
      } catch (mailErr) {
        Logger.log('Confirmation email failed: ' + mailErr.message);
        // Don't fail the booking — email is a bonus
      }
    }

    return {
      success:      true,
      patientId:    patientId,
      appointmentId: appointmentId,
      isNewPatient: isNewPatient,
      message:      isNewPatient
        ? 'Booking confirmed! Your Patient ID: ' + patientId + '. Save this to access your records.'
        : 'Booking confirmed! Your Patient ID: ' + patientId
    };

  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.5  patientPortalLogin(patientId, phone)
//      Verify Patient ID + phone → return patient object
// ─────────────────────────────────────────────────────────────
function patientPortalLogin(patientId, phone) {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var patientsSheet = ss.getSheetByName('Patients');
    if (!patientsSheet) return { success: false, error: 'Database not found.' };

    var data = patientsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var pid    = String(data[i][0] || '').trim().toUpperCase();
      var pPhone = String(data[i][2] || '').trim();
      if (pid === String(patientId).trim().toUpperCase() && pPhone === String(phone).trim()) {
        return {
          success: true,
          patient: {
            patientId:      data[i][0],
            fullName:       data[i][1],
            phone:          data[i][2],
            age:            data[i][3],
            gender:         data[i][4],
            address:        data[i][5],
            registeredDate: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'dd MMM yyyy') : '',
            lastVisit:      data[i][7] ? Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), 'dd MMM yyyy') : '',
            email:          data[i][9] || ''
          }
        };
      }
    }
    return { success: false, error: 'Patient ID or phone number does not match. Please check and try again.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.6  getPatientRecords(patientId)
//      Returns appointments (all) + approved test records
// ─────────────────────────────────────────────────────────────
function getPatientRecords(patientId) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var apptSheet = ss.getSheetByName('Appointments');
    var testSheet = ss.getSheetByName('TestRecords');
    var tz        = Session.getScriptTimeZone();
    var pid       = String(patientId || '').trim().toUpperCase();

    // ── Appointments ──
    var appointments = [];
    if (apptSheet) {
      var apptData = apptSheet.getDataRange().getValues();
      for (var i = 1; i < apptData.length; i++) {
        if (String(apptData[i][1] || '').trim().toUpperCase() === pid) {
          appointments.push({
            appointmentId: apptData[i][0],
            date:          apptData[i][2] ? Utilities.formatDate(new Date(apptData[i][2]), tz, 'dd MMM yyyy') : '',
            timeSlot:      apptData[i][3],
            reason:        apptData[i][4],
            status:        apptData[i][5]
          });
        }
      }
      appointments.reverse(); // newest first
    }

    // ── Test Records (Approved only) ──
    var testRecords = [];
    if (testSheet) {
      var testData = testSheet.getDataRange().getValues();
      // TestRecords headers:
      // A=RecordID B=PatientID C=TestName D=TestDate E=FileLink
      // F=UploadedBy G=Status H=ApprovedBy I=ApprovedDate J=Remarks
      for (var j = 1; j < testData.length; j++) {
        if (
          String(testData[j][1] || '').trim().toUpperCase() === pid &&
          String(testData[j][6] || '').trim() === 'Approved'
        ) {
          testRecords.push({
            recordId:   testData[j][0],
            testName:   testData[j][2],
            testDate:   testData[j][3] ? Utilities.formatDate(new Date(testData[j][3]), tz, 'dd MMM yyyy') : '',
            fileLink:   testData[j][4],
            uploadedBy: testData[j][5]
          });
        }
      }
      testRecords.reverse();
    }

    return { success: true, appointments: appointments, testRecords: testRecords };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.7  patientSelfUpload(patientId, base64Data, mimeType, fileName, testName)
//      Patient uploads their own test scan
//      → saves to Google Drive ClinicPatientFiles folder
//      → adds to TestRecords with Status = "Pending Approval"
// ─────────────────────────────────────────────────────────────
function patientSelfUpload(patientId, base64Data, mimeType, fileName, testName) {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var patientsSheet = ss.getSheetByName('Patients');
    var testSheet     = ss.getSheetByName('TestRecords');
    if (!patientsSheet || !testSheet) return { success: false, error: 'Database sheet not found.' };

    var pid = String(patientId || '').trim().toUpperCase();

    // Verify patient exists
    var pData       = patientsSheet.getDataRange().getValues();
    var patientName = null;
    for (var i = 1; i < pData.length; i++) {
      if (String(pData[i][0] || '').trim().toUpperCase() === pid) {
        patientName = pData[i][1];
        break;
      }
    }
    if (!patientName) return { success: false, error: 'Patient ID not found.' };

    // Get or create Drive folder: ClinicPatientFiles / PT-XXXX_Name
    var rootName    = 'ClinicPatientFiles';
    var rootFolders = DriveApp.getFoldersByName(rootName);
    var rootFolder  = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootName);

    var ptFolderName = pid + '_' + String(patientName).replace(/\s+/g, '_');
    var ptFolders    = rootFolder.getFoldersByName(ptFolderName);
    var ptFolder     = ptFolders.hasNext() ? ptFolders.next() : rootFolder.createFolder(ptFolderName);

    // Create file in Drive
    var cleanBase64 = base64Data.indexOf(',') !== -1 ? base64Data.split(',').pop() : base64Data;
    var blob        = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
    var file        = ptFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileLink    = file.getUrl();

    // Add to TestRecords tab
    var testData    = testSheet.getDataRange().getValues();
    var recordId    = 'TR-' + ('000' + testData.length).slice(-4);
    var tz          = Session.getScriptTimeZone();
    var today       = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    testSheet.appendRow([
      recordId,
      pid,
      testName || 'Other',
      today,
      fileLink,
      'Patient (Self-Upload)',
      'Pending Approval',
      '', '', ''
    ]);

    return {
      success:  true,
      recordId: recordId,
      message:  'Uploaded successfully! The doctor will review it before adding it to your records.'
    };

  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.8  getTestNameCategories()
//      Categories Tab → TestName group → active values for dropdown
// ─────────────────────────────────────────────────────────────
function getTestNameCategories() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) return { success: true, values: ['Blood Test','X-Ray','ECG','Ultrasound','MRI','Urine Test','Other'] };
    var data = sheet.getDataRange().getValues();
    var vals = [];
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][0] || '').trim() === 'TestName' &&
        data[i][3] === true
      ) vals.push(data[i][1]);
    }
    return { success: true, values: vals.length ? vals : ['Blood Test','X-Ray','ECG','Ultrasound','MRI','Urine Test','Other'] };
  } catch (err) {
    return { success: true, values: ['Blood Test','X-Ray','ECG','Ultrasound','Other'] };
  }
}


// ─────────────────────────────────────────────────────────────
// 5.9  getGenderCategories()
//      Categories Tab → Gender group → active values for dropdown
// ─────────────────────────────────────────────────────────────
function getGenderCategories() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) return { success: true, values: ['Female','Male','Other','Prefer not to say'] };
    var data = sheet.getDataRange().getValues();
    var vals = [];
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][0] || '').trim() === 'Gender' &&
        data[i][3] === true
      ) vals.push(data[i][1]);
    }
    return { success: true, values: vals.length ? vals : ['Female','Male','Other','Prefer not to say'] };
  } catch (err) {
    return { success: true, values: ['Female','Male','Other','Prefer not to say'] };
  }
}

// ============================================================
// END OF PART 5 — Code.gs Backend
// ============================================================

// ============================================================
// AUMATIQ — Doctor Automation System
// Part 5B: Professional HTML Email Templates
// clinic.drasma@gmail.com | Female Doctor Theme
// Rose Gold + Teal + Deep Plum
// ============================================================
// এই পুরো block টা Code.gs-এর একেবারে শেষে paste করতে হবে।
// ============================================================


// ─────────────────────────────────────────────────────────────
// EMAIL COLOR TOKENS (Female Doctor — Rose Gold + Teal)
// ─────────────────────────────────────────────────────────────
var EMAIL_THEME = {
  bgOuter:    '#F4EEF8',   // Soft lavender outer background
  bgCard:     '#FFFFFF',   // White card
  bgHeader:   '#0F0A1A',   // Deep plum header
  rose:       '#E8608A',   // Rose pink — primary
  teal:       '#14B8A6',   // Teal — secondary
  indigo:     '#7C3AED',   // Indigo/violet
  gold:       '#F5A623',   // Gold accent
  textDark:   '#1A0A2E',   // Deep plum text
  textBody:   '#4A3860',   // Body text
  textMuted:  '#8B7A9E',   // Muted text
  border:     '#E8D5F5',   // Light lavender border
  success:    '#10B981',   // Green
  clinicName: 'Dr. Asma Clinic',
  clinicEmail:'clinic.drasma@gmail.com',
  footer:     'Powered by AUMATIQ &bull; AI Automation Agency &bull; aumatiq.com'
};


// ─────────────────────────────────────────────────────────────
// HELPER: Get Doctor Profile for emails
// ─────────────────────────────────────────────────────────────
function getProfileForEmail_() {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = ss.getSheetByName('DoctorProfile');
    if (!sheet) return {};
    var data   = sheet.getDataRange().getValues();
    var p      = {};
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0] !== 'Field') p[data[i][0]] = data[i][1] || '';
    }
    return p;
  } catch (e) { return {}; }
}


// ─────────────────────────────────────────────────────────────
// HELPER: Shared Email Header HTML
// ─────────────────────────────────────────────────────────────
function emailHeader_(clinicName, specialty) {
  var t = EMAIL_THEME;
  return (
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + t.bgHeader + ';">' +
    '<tr><td align="center" style="padding:0;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;">' +
    '<tr>' +
    // Left: Logo / Clinic name
    '<td style="padding:24px 32px 24px 32px;vertical-align:middle;">' +
    '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td style="vertical-align:middle;padding-right:12px;">' +
    '<div style="width:40px;height:40px;background:linear-gradient(135deg,#E8608A,#7C3AED);border-radius:10px;display:inline-flex;align-items:center;justify-content:center;font-size:20px;line-height:40px;text-align:center;">🌸</div>' +
    '</td>' +
    '<td style="vertical-align:middle;">' +
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:18px;font-weight:700;color:#FFFFFF;letter-spacing:0.5px;">' + (clinicName || t.clinicName) + '</div>' +
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#E8608A;margin-top:2px;font-weight:600;">' + (specialty || 'Medical Clinic') + '</div>' +
    '</td>' +
    '</tr></table>' +
    '</td>' +
    // Right: Tagline pill
    '<td align="right" style="padding:24px 32px 24px 0;vertical-align:middle;">' +
    '<span style="display:inline-block;background:rgba(232,96,138,0.18);border:1px solid rgba(232,96,138,0.4);color:#E8608A;font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;padding:5px 14px;border-radius:20px;">CLINIC NOTIFICATION</span>' +
    '</td>' +
    '</tr>' +
    // Gradient separator line
    '<tr><td colspan="2" style="padding:0;height:3px;background:linear-gradient(90deg,#E8608A,#14B8A6,#7C3AED);"></td></tr>' +
    '</table>' +
    '</td></tr></table>'
  );
}


// ─────────────────────────────────────────────────────────────
// HELPER: Shared Email Footer HTML
// ─────────────────────────────────────────────────────────────
function emailFooter_(clinicEmail, clinicPhone, clinicAddress) {
  var t = EMAIL_THEME;
  return (
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + t.bgOuter + ';border-top:1px solid ' + t.border + ';">' +
    '<tr><td align="center" style="padding:32px 20px;">' +
    '<table width="560" cellpadding="0" cellspacing="0" border="0" style="max-width:560px;width:100%;">' +

    // Contact row
    '<tr><td align="center" style="padding-bottom:20px;">' +
    '<table cellpadding="0" cellspacing="0" border="0">' +
    '<tr>' +
    '<td style="padding:0 16px;font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';border-right:1px solid ' + t.border + ';">📞 ' + (clinicPhone || '') + '</td>' +
    '<td style="padding:0 16px;font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';border-right:1px solid ' + t.border + ';">✉️ <a href="mailto:' + (clinicEmail || t.clinicEmail) + '" style="color:' + t.rose + ';text-decoration:none;">' + (clinicEmail || t.clinicEmail) + '</a></td>' +
    '<td style="padding:0 16px;font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';">📍 ' + (clinicAddress || 'Dhaka, Bangladesh') + '</td>' +
    '</tr></table>' +
    '</td></tr>' +

    // Divider
    '<tr><td style="height:1px;background:' + t.border + ';padding:0;"></td></tr>' +

    // AUMATIQ credit
    '<tr><td align="center" style="padding-top:18px;">' +
    '<p style="font-family:Arial,sans-serif;font-size:11px;color:' + t.textMuted + ';margin:0;line-height:1.6;">' +
    'This is an automated message. Please do not reply directly to this email.<br>' +
    '<span style="color:' + t.rose + ';font-weight:600;">clinic.drasma@gmail.com</span> | ' +
    t.footer +
    '</p>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>'
  );
}


// ─────────────────────────────────────────────────────────────
// HELPER: Info Row (icon + label + value) for detail blocks
// ─────────────────────────────────────────────────────────────
function infoRow_(icon, label, value) {
  var t = EMAIL_THEME;
  return (
    '<tr>' +
    '<td style="padding:10px 16px;vertical-align:top;width:36px;">' +
    '<span style="display:inline-block;width:32px;height:32px;background:rgba(232,96,138,0.1);border:1px solid rgba(232,96,138,0.2);border-radius:8px;font-size:15px;line-height:32px;text-align:center;">' + icon + '</span>' +
    '</td>' +
    '<td style="padding:10px 16px 10px 0;vertical-align:top;border-bottom:1px solid ' + t.border + ';">' +
    '<div style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:' + t.textMuted + ';text-transform:uppercase;letter-spacing:0.5px;">' + label + '</div>' +
    '<div style="font-family:Arial,sans-serif;font-size:14px;font-weight:600;color:' + t.textDark + ';margin-top:3px;">' + value + '</div>' +
    '</td>' +
    '</tr>'
  );
}


// ─────────────────────────────────────────────────────────────
// EMAIL 1: APPOINTMENT CONFIRMATION
// Trigger: submitPublicBooking() success → call sendConfirmationEmail()
// ─────────────────────────────────────────────────────────────
function sendConfirmationEmail(patientEmail, patientName, patientId, appointmentId, date, timeSlot, reason) {
  if (!patientEmail) return; // Email optional in booking

  var p  = getProfileForEmail_();
  var t  = EMAIL_THEME;
  var cn = p.ClinicName  || t.clinicName;
  var dn = 'Dr. ' + (p.DoctorName || 'Asma');
  var sp = p.Specialty   || 'General Physician';

  var fmtDate = formatDateForEmail_(date);

  var html = (
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:' + t.bgOuter + ';font-family:Arial,Helvetica,sans-serif;">' +

    // Outer wrapper
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td align="center" style="padding:32px 20px;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;border-radius:16px;overflow:hidden;box-shadow:0 4px 40px rgba(15,10,26,0.12);">' +

    // ── HEADER ──
    '<tr><td style="padding:0;">' + emailHeader_(cn, sp) + '</td></tr>' +

    // ── HERO BANNER ──
    '<tr><td style="background:linear-gradient(135deg,#2D1A4A,#0F0A1A);padding:44px 40px 36px;text-align:center;">' +
    '<div style="display:inline-block;width:68px;height:68px;background:rgba(20,184,166,0.15);border:2px solid rgba(20,184,166,0.4);border-radius:50%;font-size:32px;line-height:68px;text-align:center;margin-bottom:20px;">✅</div>' +
    '<h1 style="font-family:Arial,Helvetica,sans-serif;font-size:26px;font-weight:700;color:#FFFFFF;margin:0 0 10px;line-height:1.3;">Appointment Confirmed!</h1>' +
    '<p style="font-family:Arial,sans-serif;font-size:15px;color:#C4B5D8;margin:0;line-height:1.6;">Your booking is confirmed with <strong style="color:#E8608A;">' + dn + '</strong>.<br>We look forward to seeing you.</p>' +
    '</td></tr>' +

    // ── PATIENT ID HIGHLIGHT BOX ──
    '<tr><td style="background:#FFFFFF;padding:32px 40px 0;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="background:linear-gradient(135deg,rgba(232,96,138,0.08),rgba(20,184,166,0.06));border:1px solid rgba(232,96,138,0.2);border-radius:14px;padding:20px 24px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td style="vertical-align:middle;">' +
    '<div style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:' + t.textMuted + ';text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">Your Patient ID</div>' +
    '<div style="font-family:\'Courier New\',monospace;font-size:22px;font-weight:700;color:' + t.rose + ';letter-spacing:2px;">' + patientId + '</div>' +
    '<div style="font-family:Arial,sans-serif;font-size:12px;color:' + t.textMuted + ';margin-top:4px;">Save this — you\'ll need it to access your records &amp; portal.</div>' +
    '</td>' +
    '<td align="right" style="vertical-align:middle;padding-left:16px;">' +
    '<span style="display:inline-block;font-family:Arial,sans-serif;font-size:11px;font-weight:700;padding:6px 14px;border-radius:20px;background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.3);color:#10B981;letter-spacing:0.5px;">CONFIRMED</span>' +
    '</td>' +
    '</tr></table>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // ── APPOINTMENT DETAILS ──
    '<tr><td style="background:#FFFFFF;padding:24px 40px;">' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">📋 Appointment Details</h2>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    infoRow_('📅', 'Date', fmtDate) +
    infoRow_('⏰', 'Time', timeSlot || '—') +
    infoRow_('👩‍⚕️', 'Doctor', dn + ' — ' + sp) +
    infoRow_('📋', 'Appointment ID', appointmentId || '—') +
    infoRow_('📝', 'Reason for Visit', reason || 'General Consultation') +
    '</table>' +
    '</td></tr>' +

    // ── WHAT TO BRING BOX ──
    '<tr><td style="background:#FFFFFF;padding:0 40px 28px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="background:rgba(232,96,138,0.05);border:1px solid rgba(232,96,138,0.15);border-radius:14px;padding:20px 24px;">' +
    '<h3 style="font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:' + t.rose + ';margin:0 0 12px;">💡 What to Bring</h3>' +
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="vertical-align:top;padding-right:12px;font-size:16px;">•</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:6px;">National ID card</td></tr><tr><td style="vertical-align:top;padding-right:12px;font-size:16px;">•</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:6px;">Any previous test results or medical records</td></tr><tr><td style="vertical-align:top;padding-right:12px;font-size:16px;">•</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:6px;">Current medication list</td></tr><tr><td style="vertical-align:top;padding-right:12px;font-size:16px;">•</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:6px;">Arrive 10 minutes early for smooth check-in</td></tr></table>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // ── CTA BUTTON ──
    '<tr><td style="background:#FFFFFF;padding:8px 40px 36px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';margin:0 0 16px;">Access your Patient Portal to view records, test results, and future appointments.</p>' +
    '<a href="' + (p.WebsiteURL || '#') + '#portal" style="display:inline-block;background:linear-gradient(135deg,#E8608A,#7C3AED);color:#FFFFFF;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:14px 32px;border-radius:10px;text-decoration:none;letter-spacing:0.3px;">Access My Patient Portal →</a>' +
    '</td></tr>' +

    // ── CANCELLATION NOTICE ──
    '<tr><td style="background:' + t.bgOuter + ';border-top:1px solid ' + t.border + ';padding:20px 40px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:12px;color:' + t.textMuted + ';margin:0;line-height:1.6;">Need to cancel or reschedule? Contact us at least 2 hours before your appointment.<br>Call: <strong>' + (p.ContactPhone || '—') + '</strong> | WhatsApp: <strong>' + (p.ContactPhone || '—') + '</strong></p>' +
    '</td></tr>' +

    // ── FOOTER ──
    '<tr><td style="padding:0;">' + emailFooter_(p.ContactEmail || t.clinicEmail, p.ContactPhone, p.ClinicAddress) + '</td></tr>' +

    '</table>' +
    '</td></tr>' +
    '</table>' +
    '</body></html>'
  );

  try {
    GmailApp.sendEmail(
      patientEmail,
      '✅ Appointment Confirmed — ' + cn + ' | ' + fmtDate + ' at ' + timeSlot,
      'Your appointment with ' + dn + ' is confirmed for ' + fmtDate + ' at ' + timeSlot + '. Patient ID: ' + patientId,
      {
        htmlBody: html,
        name: cn,
        replyTo: p.ContactEmail || t.clinicEmail
      }
    );
    Logger.log('Confirmation email sent to: ' + patientEmail);
  } catch (err) {
    Logger.log('Email error (confirmation): ' + err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// EMAIL 2: APPOINTMENT REMINDER (24-hour before)
// Call this from a time-based trigger the day before appointments
// ─────────────────────────────────────────────────────────────
function sendReminderEmail(patientEmail, patientName, patientId, appointmentId, date, timeSlot) {
  if (!patientEmail) return;

  var p  = getProfileForEmail_();
  var t  = EMAIL_THEME;
  var cn = p.ClinicName  || t.clinicName;
  var dn = 'Dr. ' + (p.DoctorName || 'Asma');
  var sp = p.Specialty   || 'General Physician';
  var fmtDate = formatDateForEmail_(date);

  var html = (
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:' + t.bgOuter + ';font-family:Arial,Helvetica,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td align="center" style="padding:32px 20px;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;border-radius:16px;overflow:hidden;box-shadow:0 4px 40px rgba(15,10,26,0.12);">' +

    '<tr><td style="padding:0;">' + emailHeader_(cn, sp) + '</td></tr>' +

    // Hero — Teal theme for reminder
    '<tr><td style="background:linear-gradient(135deg,#0D4A4A,#0F0A1A);padding:40px 40px 32px;text-align:center;">' +
    '<div style="display:inline-block;width:68px;height:68px;background:rgba(20,184,166,0.18);border:2px solid rgba(20,184,166,0.5);border-radius:50%;font-size:32px;line-height:68px;text-align:center;margin-bottom:18px;">🔔</div>' +
    '<h1 style="font-family:Arial,sans-serif;font-size:24px;font-weight:700;color:#FFFFFF;margin:0 0 10px;line-height:1.3;">Appointment Tomorrow!</h1>' +
    '<p style="font-family:Arial,sans-serif;font-size:15px;color:#A5F3FC;margin:0;line-height:1.6;">Just a reminder — your appointment with <strong style="color:#2DD4BF;">' + dn + '</strong><br>is scheduled for <strong style="color:#FFFFFF;">' + fmtDate + ' at ' + timeSlot + '</strong>.</p>' +
    '</td></tr>' +

    // Appointment Summary
    '<tr><td style="background:#FFFFFF;padding:32px 40px 24px;">' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">📋 Your Appointment</h2>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    infoRow_('📅', 'Date', fmtDate) +
    infoRow_('⏰', 'Time', timeSlot || '—') +
    infoRow_('👩‍⚕️', 'Doctor', dn) +
    infoRow_('🪪', 'Patient ID', patientId || '—') +
    '</table>' +
    '</td></tr>' +

    // Checklist
    '<tr><td style="background:#FFFFFF;padding:0 40px 28px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="background:rgba(20,184,166,0.06);border:1px solid rgba(20,184,166,0.2);border-radius:14px;padding:20px 24px;">' +
    '<h3 style="font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:' + t.teal + ';margin:0 0 12px;">✅ Pre-visit Checklist</h3>' +
    '<table cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="font-size:14px;padding-right:10px;padding-bottom:8px;">✓</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Bring your National ID &amp; any previous test results</td></tr>' +
    '<tr><td style="font-size:14px;padding-right:10px;padding-bottom:8px;">✓</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Carry a list of your current medications</td></tr>' +
    '<tr><td style="font-size:14px;padding-right:10px;padding-bottom:8px;">✓</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Arrive 10 minutes before your scheduled time</td></tr>' +
    '<tr><td style="font-size:14px;padding-right:10px;">✓</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';">Your Patient ID: <strong style="color:' + t.rose + ';">' + patientId + '</strong></td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // Cancel/Reschedule note
    '<tr><td style="background:#FFFFFF;padding:0 40px 32px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';margin:0 0 14px;line-height:1.6;">Need to cancel or reschedule?<br>Contact us at least <strong>2 hours before</strong> your appointment.</p>' +
    '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
    '<td style="padding-right:12px;"><a href="tel:' + (p.ContactPhone || '') + '" style="display:inline-block;background:rgba(232,96,138,0.1);border:1px solid rgba(232,96,138,0.3);color:' + t.rose + ';font-family:Arial,sans-serif;font-size:13px;font-weight:700;padding:11px 22px;border-radius:8px;text-decoration:none;">📞 Call Us</a></td>' +
    '<td><a href="https://wa.me/' + (p.ContactPhone || '').replace(/\D/g,'') + '" style="display:inline-block;background:rgba(20,184,166,0.1);border:1px solid rgba(20,184,166,0.3);color:' + t.teal + ';font-family:Arial,sans-serif;font-size:13px;font-weight:700;padding:11px 22px;border-radius:8px;text-decoration:none;">💬 WhatsApp</a></td>' +
    '</tr></table>' +
    '</td></tr>' +

    '<tr><td style="padding:0;">' + emailFooter_(p.ContactEmail || t.clinicEmail, p.ContactPhone, p.ClinicAddress) + '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</body></html>'
  );

  try {
    GmailApp.sendEmail(
      patientEmail,
      '🔔 Reminder: Appointment Tomorrow — ' + cn + ' | ' + fmtDate + ' at ' + timeSlot,
      'Reminder: Your appointment with ' + dn + ' is tomorrow, ' + fmtDate + ' at ' + timeSlot + '. Patient ID: ' + patientId,
      {
        htmlBody: html,
        name: cn,
        replyTo: p.ContactEmail || t.clinicEmail
      }
    );
  } catch (err) {
    Logger.log('Email error (reminder): ' + err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// EMAIL 3: POST-VISIT FOLLOW-UP
// Send 1–2 days after Appointment status = "Completed"
// ─────────────────────────────────────────────────────────────
function sendFollowUpEmail(patientEmail, patientName, patientId, visitDate, diagnosis, nextVisit) {
  if (!patientEmail) return;

  var p  = getProfileForEmail_();
  var t  = EMAIL_THEME;
  var cn = p.ClinicName  || t.clinicName;
  var dn = 'Dr. ' + (p.DoctorName || 'Asma');
  var sp = p.Specialty   || 'General Physician';
  var fmtDate = formatDateForEmail_(visitDate);

  var html = (
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:' + t.bgOuter + ';font-family:Arial,Helvetica,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td align="center" style="padding:32px 20px;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;border-radius:16px;overflow:hidden;box-shadow:0 4px 40px rgba(15,10,26,0.12);">' +

    '<tr><td style="padding:0;">' + emailHeader_(cn, sp) + '</td></tr>' +

    // Hero — Warm gold/rose theme for follow-up
    '<tr><td style="background:linear-gradient(135deg,#3A1060,#0F0A1A);padding:40px 40px 32px;text-align:center;">' +
    '<div style="display:inline-block;width:68px;height:68px;background:rgba(245,166,35,0.15);border:2px solid rgba(245,166,35,0.4);border-radius:50%;font-size:32px;line-height:68px;text-align:center;margin-bottom:18px;">💗</div>' +
    '<h1 style="font-family:Arial,sans-serif;font-size:24px;font-weight:700;color:#FFFFFF;margin:0 0 10px;line-height:1.3;">Thank You for Visiting!</h1>' +
    '<p style="font-family:Arial,sans-serif;font-size:15px;color:#E8D5F5;margin:0;line-height:1.6;">Dear <strong style="color:#F5A623;">' + (patientName || 'Patient') + '</strong>, thank you for trusting us with your health.<br>We hope you\'re feeling well after your visit on <strong style="color:#FFFFFF;">' + fmtDate + '</strong>.</p>' +
    '</td></tr>' +

    // Visit Summary
    '<tr><td style="background:#FFFFFF;padding:32px 40px 24px;">' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">📋 Visit Summary</h2>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    infoRow_('📅', 'Visit Date', fmtDate) +
    infoRow_('👩‍⚕️', 'Attended By', dn + ' — ' + sp) +
    infoRow_('🪪', 'Patient ID', patientId || '—') +
    (diagnosis ? infoRow_('📝', 'Diagnosis / Notes', diagnosis) : '') +
    (nextVisit ? infoRow_('📅', 'Next Appointment', nextVisit) : '') +
    '</table>' +
    '</td></tr>' +

    // Health tips block
    '<tr><td style="background:#FFFFFF;padding:0 40px 24px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="background:linear-gradient(135deg,rgba(232,96,138,0.06),rgba(124,58,237,0.04));border:1px solid rgba(232,96,138,0.15);border-radius:14px;padding:22px 24px;">' +
    '<h3 style="font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:' + t.rose + ';margin:0 0 12px;">🌸 Take Care of Yourself</h3>' +
    '<table cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="font-size:16px;padding-right:10px;padding-bottom:8px;vertical-align:top;">💊</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Take all prescribed medications as directed</td></tr>' +
    '<tr><td style="font-size:16px;padding-right:10px;padding-bottom:8px;vertical-align:top;">💧</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Stay hydrated — drink at least 8 glasses of water daily</td></tr>' +
    '<tr><td style="font-size:16px;padding-right:10px;padding-bottom:8px;vertical-align:top;">😴</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';padding-bottom:8px;">Rest adequately and avoid overexertion</td></tr>' +
    '<tr><td style="font-size:16px;padding-right:10px;vertical-align:top;">📞</td><td style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';">Contact us immediately if symptoms worsen</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // Test results notice + portal CTA
    '<tr><td style="background:#FFFFFF;padding:0 40px 28px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td style="background:rgba(20,184,166,0.06);border:1px solid rgba(20,184,166,0.2);border-radius:14px;padding:18px 24px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textBody + ';margin:0 0 14px;line-height:1.6;">🧪 <strong>Test results</strong> will appear in your Patient Portal once approved.<br>Log in anytime with your Patient ID: <strong style="color:' + t.rose + ';">' + patientId + '</strong></p>' +
    '<a href="' + (p.WebsiteURL || '#') + '#portal" style="display:inline-block;background:linear-gradient(135deg,#14B8A6,#0D9488);color:#FFFFFF;font-family:Arial,sans-serif;font-size:13px;font-weight:700;padding:12px 28px;border-radius:9px;text-decoration:none;">View My Records →</a>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // Review / Feedback
    '<tr><td style="background:#FFFFFF;padding:0 40px 36px;text-align:center;">' +
    '<div style="border-top:2px solid ' + t.border + ';padding-top:24px;">' +
    '<p style="font-family:Arial,sans-serif;font-size:14px;font-weight:600;color:' + t.textDark + ';margin:0 0 6px;">How was your experience?</p>' +
    '<p style="font-family:Arial,sans-serif;font-size:13px;color:' + t.textMuted + ';margin:0 0 14px;">Your feedback helps us serve you better.</p>' +
    '<div style="font-size:28px;letter-spacing:6px;">⭐⭐⭐⭐⭐</div>' +
    '<p style="font-family:Arial,sans-serif;font-size:12px;color:' + t.textMuted + ';margin:10px 0 0;">Reply to this email with your feedback — we read every message.</p>' +
    '</div>' +
    '</td></tr>' +

    '<tr><td style="padding:0;">' + emailFooter_(p.ContactEmail || t.clinicEmail, p.ContactPhone, p.ClinicAddress) + '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</body></html>'
  );

  try {
    GmailApp.sendEmail(
      patientEmail,
      '💗 Thank You for Your Visit — ' + cn + ' | ' + fmtDate,
      'Dear ' + (patientName || 'Patient') + ', thank you for visiting ' + cn + ' on ' + fmtDate + '. We hope you are feeling well. Patient ID: ' + patientId,
      {
        htmlBody: html,
        name: cn,
        replyTo: p.ContactEmail || t.clinicEmail
      }
    );
  } catch (err) {
    Logger.log('Email error (follow-up): ' + err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// EMAIL 4: TEST RESULT APPROVED NOTIFICATION
// Call from inside approveTestRecord() when doctor approves a test
// ─────────────────────────────────────────────────────────────
function sendTestResultEmail(patientEmail, patientName, patientId, testName, testDate, fileLink) {
  if (!patientEmail) return;

  var p  = getProfileForEmail_();
  var t  = EMAIL_THEME;
  var cn = p.ClinicName  || t.clinicName;
  var dn = 'Dr. ' + (p.DoctorName || 'Asma');
  var fmtDate = formatDateForEmail_(testDate);

  var html = (
    '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:' + t.bgOuter + ';font-family:Arial,Helvetica,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td align="center" style="padding:32px 20px;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;border-radius:16px;overflow:hidden;box-shadow:0 4px 40px rgba(15,10,26,0.12);">' +

    '<tr><td style="padding:0;">' + emailHeader_(cn, p.Specialty || 'Medical Clinic') + '</td></tr>' +

    // Hero — Teal/science theme
    '<tr><td style="background:linear-gradient(135deg,#0A2A3A,#0F0A1A);padding:40px 40px 32px;text-align:center;">' +
    '<div style="display:inline-block;width:68px;height:68px;background:rgba(20,184,166,0.18);border:2px solid rgba(20,184,166,0.5);border-radius:50%;font-size:32px;line-height:68px;text-align:center;margin-bottom:18px;">🧪</div>' +
    '<h1 style="font-family:Arial,sans-serif;font-size:24px;font-weight:700;color:#FFFFFF;margin:0 0 10px;line-height:1.3;">Your Test Result is Ready</h1>' +
    '<p style="font-family:Arial,sans-serif;font-size:15px;color:#A5F3FC;margin:0;line-height:1.6;">Your <strong style="color:#2DD4BF;">' + testName + '</strong> result has been<br>reviewed and approved by <strong style="color:#FFFFFF;">' + dn + '</strong>.</p>' +
    '</td></tr>' +

    // Test details
    '<tr><td style="background:#FFFFFF;padding:32px 40px;">' +
    '<h2 style="font-family:Arial,sans-serif;font-size:16px;font-weight:700;color:' + t.textDark + ';margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid ' + t.border + ';">🔬 Test Details</h2>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    infoRow_('🧪', 'Test Name', testName || '—') +
    infoRow_('📅', 'Test Date', fmtDate) +
    infoRow_('🪪', 'Patient ID', patientId || '—') +
    infoRow_('✅', 'Status', 'Approved by ' + dn) +
    '</table>' +
    '</td></tr>' +

    // Download CTA
    '<tr><td style="background:#FFFFFF;padding:0 40px 32px;text-align:center;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="background:rgba(20,184,166,0.06);border:1px solid rgba(20,184,166,0.2);border-radius:14px;padding:24px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:14px;color:' + t.textBody + ';margin:0 0 16px;line-height:1.6;">Your test report is now available in your <strong>Patient Portal</strong>.<br>Click below to view or download your result.</p>' +
    '<a href="' + (fileLink || '#') + '" style="display:inline-block;background:linear-gradient(135deg,#14B8A6,#7C3AED);color:#FFFFFF;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:14px 32px;border-radius:10px;text-decoration:none;margin-bottom:12px;">🔬 View Test Report →</a><br>' +
    '<a href="' + (p.WebsiteURL || '#') + '#portal" style="display:inline-block;font-family:Arial,sans-serif;font-size:12px;color:' + t.teal + ';text-decoration:none;">or access via Patient Portal</a>' +
    '</td></tr></table>' +
    '</td></tr>' +

    // Privacy note
    '<tr><td style="background:' + t.bgOuter + ';border-top:1px solid ' + t.border + ';padding:18px 40px;text-align:center;">' +
    '<p style="font-family:Arial,sans-serif;font-size:12px;color:' + t.textMuted + ';margin:0;line-height:1.6;">🔒 Your medical data is private and secure. Only you can access your records using your Patient ID and registered phone number.</p>' +
    '</td></tr>' +

    '<tr><td style="padding:0;">' + emailFooter_(p.ContactEmail || t.clinicEmail, p.ContactPhone, p.ClinicAddress) + '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</body></html>'
  );

  try {
    GmailApp.sendEmail(
      patientEmail,
      '🧪 Test Result Ready — ' + testName + ' | ' + cn,
      'Your ' + testName + ' result from ' + cn + ' is now available. Log into your Patient Portal to view it. Patient ID: ' + patientId,
      {
        htmlBody: html,
        name: cn,
        replyTo: p.ContactEmail || t.clinicEmail
      }
    );
  } catch (err) {
    Logger.log('Email error (test result): ' + err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// AUTOMATED DAILY REMINDER TRIGGER
// Set up a time-based trigger to run this every day at 9 AM
// It finds all appointments for TOMORROW and sends reminder emails
// ─────────────────────────────────────────────────────────────
function runDailyReminderEmails() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var apptSheet = ss.getSheetByName('Appointments');
  var ptSheet   = ss.getSheetByName('Patients');
  if (!apptSheet || !ptSheet) return;

  // Tomorrow's date string
  var tom = new Date();
  tom.setDate(tom.getDate() + 1);
  var tomStr = Utilities.formatDate(tom, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Build patient lookup (patientId → {email, name})
  var ptData = ptSheet.getDataRange().getValues();
  var ptMap  = {};
  // Patients headers: PatientID, FullName, Phone, Age, Gender, Address, RegisteredDate, LastVisit, Notes, Email
  for (var i = 1; i < ptData.length; i++) {
    ptMap[ptData[i][0]] = {
      name:  ptData[i][1],
      phone: ptData[i][2],
      email: ptData[i][9] || '' // Email in column J (index 9)
    };
  }

  // Scan appointments for tomorrow, status != Cancelled
  var apptData = apptSheet.getDataRange().getValues();
  // Headers: AppointmentID, PatientID, Date, TimeSlot, Reason, Status, CreatedAt
  for (var r = 1; r < apptData.length; r++) {
    var rowDate = apptData[r][2];
    if (!rowDate) continue;
    var rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var status     = (apptData[r][5] || '').toString();
    if (rowDateStr === tomStr && status !== 'Cancelled' && status !== 'Completed') {
      var pid    = apptData[r][1];
      var pt     = ptMap[pid];
      if (pt && pt.email) {
        sendReminderEmail(
          pt.email,
          pt.name,
          pid,
          apptData[r][0], // appointmentId
          rowDateStr,
          (apptData[r][3] || '').toString()
        );
        Utilities.sleep(2000); // Avoid Gmail rate limit
      }
    }
  }
  Logger.log('Daily reminders sent for: ' + tomStr);
}


// ─────────────────────────────────────────────────────────────
// SETUP HELPER: Creates the daily reminder trigger
// Run this ONCE from Apps Script editor → Run → setupDailyReminderTrigger
// ─────────────────────────────────────────────────────────────
function setupDailyReminderTrigger() {
  // Delete existing reminder triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyReminderEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new trigger — every day at 9 AM
  ScriptApp.newTrigger('runDailyReminderEmails')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('✅ Daily reminder trigger set for 9 AM every day.');
}


// ─────────────────────────────────────────────────────────────
// DATE FORMATTER for email display
// ─────────────────────────────────────────────────────────────
function formatDateForEmail_(dateStr) {
  if (!dateStr) return '—';
  try {
    var d = new Date(dateStr);
    var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var days   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    return days[d.getDay()] + ', ' + d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
  } catch (e) { return dateStr.toString(); }
}

// ============================================================
// END OF PART 5B — Email Templates
// ============================================================


// ═══════════════════════════════════════════════════════════════
// v2.0 PATCH — আগে missing ছিল এমন ৩টা Backend Function
// (doctordashboard.html এগুলো call করত, কিন্তু কোথাও define ছিল না — bug)
// ═══════════════════════════════════════════════════════════════

// ───────────────────────── 6.1 getDashboardOverview() — Dashboard Tab stats ─────────────────────────
function getDashboardOverview(token) {
  requireDoctorOrReceptionist(token);

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Total Patients ──
  var patientsSheet = ss.getSheetByName("Patients");
  var patientsData  = patientsSheet.getDataRange().getValues();
  var totalPatients = Math.max(0, patientsData.length - 1);

  // ── আজকের Appointments ──
  var apptSheet   = ss.getSheetByName("Appointments");
  var apptData    = apptSheet.getDataRange().getValues();
  var apptHeaders = apptData[0];
  var dateCol     = apptHeaders.indexOf("Date");
  var statusCol   = apptHeaders.indexOf("Status");

  var todayStr = new Date().toDateString();
  var todayAppointments = [];
  var todayCompleted = 0;

  for (var i = 1; i < apptData.length; i++) {
    var rowDate = new Date(apptData[i][dateCol]);
    if (rowDate.toDateString() === todayStr) {
      var apptObj = {};
      apptHeaders.forEach(function (h, idx) { apptObj[h] = apptData[i][idx]; });
      todayAppointments.push(apptObj);
      if (String(apptData[i][statusCol]) === "Completed") todayCompleted++;
    }
  }

  // TimeSlot অনুযায়ী sort
  todayAppointments.sort(function (a, b) {
    return String(a.TimeSlot || '').localeCompare(String(b.TimeSlot || ''));
  });

  // ── Pending Test Approvals ──
  var testSheet     = ss.getSheetByName("TestRecords");
  var testData      = testSheet.getDataRange().getValues();
  var testStatusCol = testData[0].indexOf("Status");
  var pendingApprovals = 0;
  for (var j = 1; j < testData.length; j++) {
    if (testData[j][testStatusCol] === "Pending Approval") pendingApprovals++;
  }

  return {
    success: true,
    stats: {
      totalPatients:    totalPatients,
      todayTotal:       todayAppointments.length,
      todayCompleted:   todayCompleted,
      pendingApprovals: pendingApprovals
    },
    todayAppointments: todayAppointments
  };
}

// ───────────────────────── 6.2 getAllTestRecords() — Test Records Tab (full list) ─────────────────────────
function getAllTestRecords(token) {
  requireDoctorOrReceptionist(token);

  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TestRecords");
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  var records = [];
  for (var i = 1; i < data.length; i++) {
    var recordObj = {};
    headers.forEach(function (h, idx) { recordObj[h] = data[i][idx]; });
    recordObj["_rowIndex"] = i + 1;
    records.push(recordObj);
  }

  records.reverse(); // সর্বশেষ আপলোড করা রেকর্ড সবার উপরে

  return { success: true, records: records, count: records.length };
}

// ───────────────────────── 6.3 getActiveCategoryValuesForGroup() — client-callable category dropdown ─────────────────────────
// (getActiveCategoryValues() ভেঙে পড়বে যদি সরাসরি frontend থেকে call করা হয়,
//  কারণ সেটা Sheet object আশা করে, যেটা google.script.run দিয়ে পাঠানো যায় না)
function getActiveCategoryValuesForGroup(token, groupName) {
  requireDoctorOrReceptionist(token);

  var sheet     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  var data      = sheet.getDataRange().getValues();
  var headers   = data[0];
  var groupCol  = headers.indexOf("CategoryGroup");
  var valueCol  = headers.indexOf("Value");
  var activeCol = headers.indexOf("IsActive");

  var values = [];
  for (var i = 1; i < data.length; i++) {
    var isActive = String(data[i][activeCol]).toUpperCase() === "TRUE";
    if (data[i][groupCol] === groupName && isActive) {
      values.push(data[i][valueCol]);
    }
  }

  return { success: true, values: values };
}
