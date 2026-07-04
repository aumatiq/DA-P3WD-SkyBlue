/**
 * AUMATIQ — Doctor & Clinic Automation System
 * Part 4: Appointment + Finance + Categories Manager (AppointmentFinance.gs)
 * ─────────────────────────────────────────────
 */

// ═══════════════════════════════════════════════
// SECTION 1 — APPOINTMENT BOOKING + CALENDAR
// ═══════════════════════════════════════════════

// ───────────────────────── নির্দিষ্ট তারিখে available time slot বের করা ─────────────────────────
/**
 * Settings Tab-এর SlotDuration ও WorkingDays ব্যবহার করে
 * সকাল ৯টা থেকে রাত ৮টা পর্যন্ত স্লট জেনারেট করে, তারপর
 * Appointments Tab-এ যা বুকড আছে তা বাদ দিয়ে available স্লট রিটার্ন করে।
 */
function getAvailableSlots(dateString) {
  const slotDuration = parseInt(getSettingValue("SlotDuration") || 20, 10);
  const workingDaysRaw = String(getSettingValue("WorkingDays") || "Sat,Sun,Mon,Tue,Wed");
  const workingDays = workingDaysRaw.split(",").map(function(d) { return d.trim(); });

  // ── FIX: Timezone-safe date parsing ──
  // "2026-07-05" কে UTC দিয়ে parse না করে local date হিসেবে treat করা হচ্ছে
  const parts = String(dateString).split("-");
  const requestedDate = new Date(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10)
  );

  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const requestedDayName = dayNames[requestedDate.getDay()];

  if (workingDays.indexOf(requestedDayName) === -1) {
    return {
      success: true,
      slots: [],
      message: "এই দিনে ক্লিনিক বন্ধ থাকে। (" + requestedDayName + ")"
    };
  }

  const openingTimeStr = String(getSettingValue("OpeningTime") || "09:00");
  const closingTimeStr = String(getSettingValue("ClosingTime") || "20:00");
  const openingMinutes = timeStringToMinutes(openingTimeStr);
  const closingMinutes = timeStringToMinutes(closingTimeStr);

  const allSlots = [];
  let startMinutes = openingMinutes;
  while (startMinutes < closingMinutes) {
    allSlots.push(minutesToDisplayTime(startMinutes));
    startMinutes += slotDuration;
  }

  // Booked slots check — same timezone-safe comparison
  const apptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  const data = apptSheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol   = headers.indexOf("Date");
  const timeCol   = headers.indexOf("TimeSlot");
  const statusCol = headers.indexOf("Status");

  const bookedSlots = [];
  for (let i = 1; i < data.length; i++) {
    const cellDate = new Date(data[i][dateCol]);
    const sameDate =
      cellDate.getFullYear() === requestedDate.getFullYear() &&
      cellDate.getMonth()    === requestedDate.getMonth() &&
      cellDate.getDate()     === requestedDate.getDate();
    const notCancelled = data[i][statusCol] !== "Cancelled";
    if (sameDate && notCancelled) {
      bookedSlots.push(data[i][timeCol]);
    }
  }

  const availableSlots = allSlots.filter(function(slot) {
    return bookedSlots.indexOf(slot) === -1;
  });

  return { success: true, slots: availableSlots };
}

// ───────────────────────── হেল্পার: "09:00" স্ট্রিং থেকে মিনিট (540) ─────────────────────────
function timeStringToMinutes(timeStr) {
  const parts = timeStr.split(":");
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

// ───────────────────────── হেল্পার: মিনিট থেকে "9:00 AM" ফরম্যাট ─────────────────────────
function minutesToDisplayTime(totalMinutes) {
  const hour = Math.floor(totalMinutes / 60);
  const minute = totalMinutes % 60;
  const displayHour = hour > 12 ? hour - 12 : (hour === 0 ? 12 : hour);
  const ampm = hour >= 12 ? "PM" : "AM";
  return displayHour + ":" + String(minute).padStart(2, "0") + " " + ampm;
}

// ───────────────────────── BOOK APPOINTMENT (Flow A — Patient তৈরি + Appointment তৈরি একসাথে) ─────────────────────────
/**
 * bookingData = { fullName, phone, age, gender, address, date, timeSlot, reason }
 * এটা public website থেকে কল হবে (Part 5), কোনো login ছাড়াই — তাই গার্ড নেই।
 */
function bookAppointment(bookingData) {
  if (!bookingData.phone || !bookingData.date || !bookingData.timeSlot) {
    return { success: false, message: "ফোন নম্বর, তারিখ এবং সময় আবশ্যক।" };
  }

  // স্টেপ ১ — Patient আছে কিনা চেক করা, না থাকলে তৈরি করা (Part 3-এর ফাংশন reuse)
  const patientResult = createPatient({
    fullName: bookingData.fullName,
    phone: bookingData.phone,
    age: bookingData.age,
    gender: bookingData.gender,
    address: bookingData.address,
  });

  if (!patientResult.success) {
    return { success: false, message: "Patient তৈরি করতে সমস্যা হয়েছে।" };
  }

  const patientId = patientResult.patientId;

  // স্লট এখনও available কিনা শেষবার যাচাই করা (race condition এড়াতে)
  // ── বাগ ফিক্স: "ক্লিনিক বন্ধ" বনাম "স্লট আগে বুকড" — দুটো ভিন্ন কারণের জন্য ভিন্ন মেসেজ ──
  const slotCheck = getAvailableSlots(bookingData.date);

  if (slotCheck.slots.length === 0 && slotCheck.message) {
    // getAvailableSlots থেকে নির্দিষ্ট মেসেজ এসেছে — মানে এই দিনে ক্লিনিক বন্ধ
    return { success: false, message: slotCheck.message };
  }

  if (slotCheck.slots.indexOf(bookingData.timeSlot) === -1) {
    return { success: false, message: "এই স্লট সম্প্রতি বুক হয়ে গেছে। অন্য একটা সময় বেছে নাও।" };
  }

  // স্টেপ ৩ — Appointment তৈরি করা
  const apptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  const data = apptSheet.getDataRange().getValues();

  let maxNumber = 0;
  for (let i = 1; i < data.length; i++) {
    const match = String(data[i][0]).match(/APT-(\d+)/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNumber) maxNumber = num;
    }
  }
  const newApptId = "APT-" + String(maxNumber + 1).padStart(4, "0");

  apptSheet.appendRow([
    newApptId,
    patientId,
    new Date(bookingData.date),
    bookingData.timeSlot,
    bookingData.reason || "",
    "Pending",
    new Date(),
  ]);

  return {
    success: true,
    patientId: patientId,
    appointmentId: newApptId,
    message: "বুকিং সফল হয়েছে। আপনার Patient ID: " + patientId + " — এটা সেভ করে রাখুন, ভবিষ্যতে 'My Records' লগইন করতে লাগবে।",
  };
}

// ───────────────────────── অ্যাপয়েন্টমেন্ট লিস্ট (Admin Dashboard-এর ক্যালেন্ডার ভিউ-এর জন্য) ─────────────────────────
function getAppointments(token, dateString) {
  requireDoctorOrReceptionist(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf("Date");

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][dateCol]);
    const matches = !dateString || rowDate.toDateString() === new Date(dateString).toDateString();
    if (matches) {
      const apptObj = {};
      headers.forEach(function (h, idx) { apptObj[h] = data[i][idx]; });
      apptObj["_rowIndex"] = i + 1;
      results.push(apptObj);
    }
  }

  return { success: true, appointments: results, count: results.length };
}

// ───────────────────────── অ্যাপয়েন্টমেন্ট স্ট্যাটাস আপডেট (Confirm/Complete/Cancel) ─────────────────────────
function updateAppointmentStatus(token, appointmentId, newStatus) {
  requireDoctorOrReceptionist(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("AppointmentID");
  const statusCol = headers.indexOf("Status");
  const patientIdCol = headers.indexOf("PatientID");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(appointmentId).trim()) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);

      // যদি Completed করা হয়, Patient-এর LastVisit আপডেট করা হবে
      if (newStatus === "Completed") {
        updatePatientLastVisit(data[i][patientIdCol]);
      }

      return { success: true, message: "অ্যাপয়েন্টমেন্ট স্ট্যাটাস আপডেট হয়েছে।" };
    }
  }

  return { success: false, message: "Appointment ID পাওয়া যায়নি।" };
}

// ───────────────────────── হেল্পার: Patient-এর LastVisit আপডেট করা ─────────────────────────
function updatePatientLastVisit(patientId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("PatientID");
  const lastVisitCol = headers.indexOf("LastVisit");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(patientId).trim()) {
      sheet.getRange(i + 1, lastVisitCol + 1).setValue(new Date());
      return;
    }
  }
}

// ───────────────────────── CLINIC HOURS আপডেট (Dashboard থেকে, Doctor-only) ─────────────────────────
/**
 * Dashboard-এর Settings পেজ থেকে কল হবে — ডাক্তার নিজে Opening/Closing Time,
 * SlotDuration, এবং WorkingDays পরিবর্তন করতে পারবে।
 */
function updateClinicHours(token, settingsData) {
  requireDoctor(token);

  if (settingsData.openingTime) {
    setSettingValue("OpeningTime", settingsData.openingTime);
  }
  if (settingsData.closingTime) {
    setSettingValue("ClosingTime", settingsData.closingTime);
  }
  if (settingsData.slotDuration) {
    setSettingValue("SlotDuration", settingsData.slotDuration);
  }
  if (settingsData.workingDays) {
    setSettingValue("WorkingDays", settingsData.workingDays);
  }

  return { success: true, message: "ক্লিনিকের সময়সূচী আপডেট হয়েছে।" };
}

// ───────────────────────── হেল্পার: Settings Tab-এ ভ্যালু লেখা/আপডেট করা ─────────────────────────
function setSettingValue(fieldName, newValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === fieldName) {
      sheet.getRange(i + 1, 2).setValue(newValue);
      return true;
    }
  }
  return false; // ফিল্ড পাওয়া যায়নি
}

// ───────────────────────── বর্তমান Clinic Hours দেখা (Dashboard লোড হওয়ার সময়) ─────────────────────────
function getClinicHours(token) {
  requireDoctorOrReceptionist(token); // দুজনেই দেখতে পারবে, শুধু Doctor এডিট করতে পারবে

  return {
    success: true,
    openingTime: getSettingValue("OpeningTime"),
    closingTime: getSettingValue("ClosingTime"),
    slotDuration: getSettingValue("SlotDuration"),
    workingDays: getSettingValue("WorkingDays"),
  };
}

// ═══════════════════════════════════════════════
// SECTION 2 — FINANCE MODULE (শুধু DOCTOR অ্যাক্সেস করতে পারবে)
// ═══════════════════════════════════════════════

// ───────────────────────── নতুন ফাইন্যান্স এন্ট্রি অ্যাড করা ─────────────────────────
function addFinanceEntry(token, entryData) {
  requireDoctor(token); // গুরুত্বপূর্ণ — শুধু Doctor, Receptionist না

  if (!entryData.type || !entryData.category || !entryData.amount) {
    return { success: false, message: "Type, Category, এবং Amount আবশ্যক।" };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Finance");
  const data = sheet.getDataRange().getValues();

  let maxNumber = 0;
  for (let i = 1; i < data.length; i++) {
    const match = String(data[i][0]).match(/FN-(\d+)/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNumber) maxNumber = num;
    }
  }
  const newEntryId = "FN-" + String(maxNumber + 1).padStart(4, "0");

  sheet.appendRow([
    newEntryId,
    entryData.date ? new Date(entryData.date) : new Date(),
    entryData.type,
    entryData.category,
    entryData.patientId || "",
    entryData.amount,
    entryData.notes || "",
  ]);

  return { success: true, entryId: newEntryId, message: "ফাইন্যান্স এন্ট্রি যুক্ত হয়েছে।" };
}

// ───────────────────────── ফাইন্যান্স এন্ট্রি লিস্ট (তারিখ রেঞ্জ অনুযায়ী ফিল্টার) ─────────────────────────
function getFinanceEntries(token, startDate, endDate) {
  requireDoctor(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Finance");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf("Date");

  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][dateCol]);
    const afterStart = !start || rowDate >= start;
    const beforeEnd = !end || rowDate <= end;

    if (afterStart && beforeEnd) {
      const entryObj = {};
      headers.forEach(function (h, idx) { entryObj[h] = data[i][idx]; });
      results.push(entryObj);
    }
  }

  return { success: true, entries: results, count: results.length };
}

// ───────────────────────── মাসিক ফাইন্যান্স সামারি (Income/Expense টোটাল) ─────────────────────────
function getMonthlyFinanceSummary(token, year, month) {
  requireDoctor(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Finance");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf("Date");
  const typeCol = headers.indexOf("Type");
  const amountCol = headers.indexOf("Amount");
  const categoryCol = headers.indexOf("Category");

  let totalIncome = 0;
  let totalExpense = 0;
  const categoryBreakdown = {};

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][dateCol]);
    if (rowDate.getFullYear() === year && rowDate.getMonth() + 1 === month) {
      const amount = parseFloat(data[i][amountCol]) || 0;
      const type = data[i][typeCol];
      const category = data[i][categoryCol];

      if (type === "Income") totalIncome += amount;
      if (type === "Expense") totalExpense += amount;

      categoryBreakdown[category] = (categoryBreakdown[category] || 0) + amount;
    }
  }

  return {
    success: true,
    totalIncome: totalIncome,
    totalExpense: totalExpense,
    netProfit: totalIncome - totalExpense,
    categoryBreakdown: categoryBreakdown,
  };
}

// ═══════════════════════════════════════════════
// SECTION 3 — CATEGORIES MANAGER (Settings & Categories Tab)
// ═══════════════════════════════════════════════

// ───────────────────────── সব ক্যাটেগরি দেখা (Active + Deleted আলাদাভাবে, v1.3) ─────────────────────────
function getAllCategories(token) {
  requireDoctor(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const activeCol = headers.indexOf("IsActive");

  const grouped = {};
  for (let i = 1; i < data.length; i++) {
    const group = data[i][0];
    if (!grouped[group]) grouped[group] = { active: [], deleted: [] };

    const catObj = {};
    headers.forEach(function (h, idx) { catObj[h] = data[i][idx]; });
    catObj["_rowIndex"] = i + 1;

    const isActive = String(data[i][activeCol]).toUpperCase() === "TRUE";
    if (isActive) {
      grouped[group].active.push(catObj);
    } else {
      grouped[group].deleted.push(catObj);
    }
  }

  return { success: true, categories: grouped };
}

// ───────────────────────── নতুন ক্যাটেগরি ভ্যালু অ্যাড করা (v1.3) ─────────────────────────
function addCategoryValue(token, categoryGroup, value) {
  const session = requireDoctor(token);

  if (!categoryGroup || !value) {
    return { success: false, message: "Category Group এবং Value দুটোই দিতে হবে।" };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === categoryGroup && String(data[i][1]).toLowerCase() === String(value).toLowerCase()) {
      const isActive = String(data[i][3]).toUpperCase() === "TRUE"; // কলাম index 3 = IsActive
      if (isActive) {
        return { success: false, message: "এই ভ্যালু আগে থেকেই এই ক্যাটেগরিতে আছে।" };
      } else {
        // আগে ডিলিট করা ছিল, আবার সক্রিয় করে দেওয়া হচ্ছে (duplicate row তৈরি না করে)
        sheet.getRange(i + 1, 4).setValue("TRUE");
        return { success: true, message: "এই ভ্যালু আগে ডিলিট করা ছিল, পুনরায় চালু করা হলো।" };
      }
    }
  }

  sheet.appendRow([categoryGroup, value, "FALSE", "TRUE", session.identifier, new Date()]);
  return { success: true, message: "নতুন ক্যাটেগরি ভ্যালু যুক্ত হয়েছে।" };
}

// ───────────────────────── ক্যাটেগরি ভ্যালু SOFT DELETE করা (v1.3) ─────────────────────────
function deleteCategoryValue(token, categoryGroup, value) {
  requireDoctor(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const groupCol = headers.indexOf("CategoryGroup");
  const valueCol = headers.indexOf("Value");
  const defaultCol = headers.indexOf("IsSystemDefault");
  const activeCol = headers.indexOf("IsActive");

  for (let i = 1; i < data.length; i++) {
    if (data[i][groupCol] === categoryGroup && String(data[i][valueCol]) === String(value)) {
      const isProtected = String(data[i][defaultCol]).toUpperCase() === "TRUE";

      if (isProtected) {
        return { success: false, message: "এটা একটা সিস্টেম-প্রোটেক্টেড ভ্যালু, ডিলিট করা যাবে না।" };
      }

      sheet.getRange(i + 1, activeCol + 1).setValue("FALSE"); // hard delete না, শুধু IsActive = FALSE
      return { success: true, message: "ক্যাটেগরি ভ্যালু ডিলিট হয়েছে (Restore করা সম্ভব)।" };
    }
  }

  return { success: false, message: "ক্যাটেগরি ভ্যালু পাওয়া যায়নি।" };
}

// ───────────────────────── ক্যাটেগরি ভ্যালু RESTORE করা (নতুন, v1.3) ─────────────────────────
function restoreCategoryValue(token, categoryGroup, value) {
  requireDoctor(token);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const groupCol = headers.indexOf("CategoryGroup");
  const valueCol = headers.indexOf("Value");
  const activeCol = headers.indexOf("IsActive");

  for (let i = 1; i < data.length; i++) {
    if (data[i][groupCol] === categoryGroup && String(data[i][valueCol]) === String(value)) {
      sheet.getRange(i + 1, activeCol + 1).setValue("TRUE");
      return { success: true, message: "ক্যাটেগরি ভ্যালু পুনরায় চালু করা হলো।" };
    }
  }

  return { success: false, message: "ক্যাটেগরি ভ্যালু পাওয়া যায়নি।" };
}
