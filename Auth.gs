/**
 * AUMATIQ — Doctor & Clinic Automation System
 * Part 2: Login + Security System (Auth.gs) — v2.0
 * ─────────────────────────────────────────────
 * v2.0 বদল:
 *  - Doctor/Assistant Login এখন Username লাগে না — শুধু Role বাছাই (বাটনে ক্লিক) + Password।
 *  - changePassword() যোগ করা হলো (আগে এই function ছিলই না — Password Management UI কাজ করত না)।
 *  - পুরনো ডুপ্লিকেট patientLogin() ফাংশন সরানো হলো (Code.gs-এর patientPortalLogin()
 *    ব্যবহার হয়, patientLogin() কোথাও call হতো না — dead code ছিল)।
 * Patient login, এবং session token validation অপরিবর্তিত।
 */

// ───────────────────────── কনফিগ ─────────────────────────
const SESSION_DURATION_MINUTES = 120; // লগইন token কতক্ষণ valid থাকবে (২ ঘণ্টা)

// ───────────────────────── হেল্পার: Settings Tab থেকে ভ্যালু পড়া ─────────────────────────
function getSettingValue(fieldName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === fieldName) {
      return data[i][1];
    }
  }
  return null;
}

// ───────────────────────── ROLE LOGIN (Doctor বা Assistant — Password-only) ─────────────────────────
/**
 * v2.0: Username লাগে না। Login screen-এ Doctor / Assistant বাটনে ক্লিক করে
 * role বাছাই করা হয়, তারপর শুধু সেই role-এর Password দিলেই লগইন হয়ে যায়।
 * role আভ্যন্তরীণভাবে এখনো "DOCTOR" / "RECEPTIONIST" — বাকি সব guard function
 * (requireDoctor, requireDoctorOrReceptionist ইত্যাদি) অপরিবর্তিত রাখার জন্য।
 * UI-তে "RECEPTIONIST"-কে "Assistant" হিসেবে দেখানো হয়।
 */
function roleLogin(role, password) {
  const validRoles = ["DOCTOR", "RECEPTIONIST"];

  if (!role || validRoles.indexOf(role) === -1) {
    return { success: false, message: "সঠিক Role বাছাই করো (Doctor / Assistant)।" };
  }
  if (!password || String(password).trim() === "") {
    return { success: false, message: "Password দিতে হবে।" };
  }

  const inputPass = String(password).trim();

  const storedPassword = role === "DOCTOR"
    ? String(getSettingValue("DoctorPassword") || "").trim()
    : String(getSettingValue("ReceptionistPassword") || "").trim();

  if (!storedPassword) {
    return {
      success: false,
      message: "এই Role-এর জন্য এখনো কোনো Password সেট করা নেই। Doctor-কে Settings ট্যাব থেকে Password সেট করতে বলো।"
    };
  }

  if (inputPass !== storedPassword) {
    return { success: false, message: "ভুল Password। আবার চেষ্টা করো।" };
  }

  const token = createSession(role, role === "DOCTOR" ? "doctor" : "assistant");
  return {
    success: true,
    token: token,
    role: role,
    message: role === "DOCTOR"
      ? "লগইন সফল হয়েছে। (Doctor — Full Access)"
      : "লগইন সফল হয়েছে। (Assistant — Limited Access, Finance ও Settings দেখা যাবে না)"
  };
}

// ───────────────────────── PASSWORD পরিবর্তন (শুধু Doctor করতে পারবে) ─────────────────────────
/**
 * roleToChange: "DOCTOR" অথবা "RECEPTIONIST" — কার password বদলানো হচ্ছে।
 * newPassword : 2 থেকে 15 characters — সংখ্যা/অক্ষর/সিম্বল/মিশ্র, যেকোনো কম্বিনেশন চলবে।
 * শুধুমাত্র Doctor role-এ লগইন করা থাকলেই এই ফাংশন কাজ করবে (Assistant পারবে না —
 * এমনকি browser console থেকে সরাসরি call করলেও না, কারণ guard সার্ভার সাইডে বসানো)।
 */
function changePassword(token, roleToChange, newPassword) {
  requireDoctor(token);

  const validRoles = ["DOCTOR", "RECEPTIONIST"];
  if (!roleToChange || validRoles.indexOf(roleToChange) === -1) {
    return { success: false, message: "সঠিক Role নির্বাচন করো।" };
  }

  const pass = String(newPassword || "").trim();
  if (pass.length < 2 || pass.length > 15) {
    return { success: false, message: "Password অবশ্যই 2 থেকে 15 characters-এর মধ্যে হতে হবে।" };
  }

  const fieldName = roleToChange === "DOCTOR" ? "DoctorPassword" : "ReceptionistPassword";
  setSettingValue(fieldName, pass);

  return {
    success: true,
    message: (roleToChange === "DOCTOR" ? "Doctor" : "Assistant") + "-এর Password সফলভাবে আপডেট হয়েছে।"
  };
}

// ───────────────────────── SESSION তৈরি করা ─────────────────────────
function createSession(role, identifier) {
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  const sessionData = JSON.stringify({ role: role, identifier: identifier });

  cache.put(token, sessionData, SESSION_DURATION_MINUTES * 60);
  return token;
}

// ───────────────────────── SESSION যাচাই করা ─────────────────────────
function validateSession(token) {
  if (!token) {
    return { valid: false, message: "Session token পাওয়া যায়নি, আবার লগইন করো।" };
  }

  const cache = CacheService.getScriptCache();
  const sessionData = cache.get(token);

  if (!sessionData) {
    return { valid: false, message: "Session expire হয়ে গেছে, আবার লগইন করো।" };
  }

  const parsed = JSON.parse(sessionData);
  return { valid: true, role: parsed.role, identifier: parsed.identifier };
}

// ───────────────────────── LOGOUT ─────────────────────────
function logout(token) {
  if (token) {
    const cache = CacheService.getScriptCache();
    cache.remove(token);
  }
  return { success: true, message: "লগআউট হয়ে গেছে।" };
}

// ───────────────────────── গার্ড: শুধু DOCTOR-এর জন্য (যেমন Finance, Categories Manage, Password Change) ─────────────────────────
function requireDoctor(token) {
  const session = validateSession(token);
  if (!session.valid || session.role !== "DOCTOR") {
    throw new Error("অনুমতি নেই — এই অ্যাকশন শুধুমাত্র Doctor করতে পারবে।");
  }
  return session;
}

// ───────────────────────── গার্ড: DOCTOR অথবা RECEPTIONIST/Assistant (যেমন Patient/Appointment/Test Upload) ─────────────────────────
function requireDoctorOrReceptionist(token) {
  const session = validateSession(token);
  if (!session.valid || (session.role !== "DOCTOR" && session.role !== "RECEPTIONIST")) {
    throw new Error("অনুমতি নেই — Doctor বা Assistant লগইন প্রয়োজন।");
  }
  return session;
}

// ───────────────────────── গার্ড: শুধু PATIENT-এর জন্য ─────────────────────────
function requirePatient(token) {
  const session = validateSession(token);
  if (!session.valid || session.role !== "PATIENT") {
    throw new Error("অনুমতি নেই — Patient লগইন প্রয়োজন।");
  }
  return session;
}
