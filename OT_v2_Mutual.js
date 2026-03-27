// ============================================================
// ELECTRICAL OT AUTOMATION SYSTEM
// FAST BASE + MUTUAL + OT HISTORY + OT TAG + TOMORROW DUTY
// ============================================================

// ============================================================
// SECTION 1: GLOBAL CONFIG
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

var SHEETS = {
  EMP:    "Employees",
  SEQ:    "OT_Sequence",
  TODAY:  "Today_Shift",
  REQ:    "Shift_Requirement",
  LEAVE:  "Leave_List",
  LOG:    "OT_Log",
  STATUS: "OT_Status",
  BASE:   "Shift_Base",
  DUTY:   "Duty_Auto",
  MUTUAL: "Mutual_List"
};

var TOKEN         = "8720679948:AAGGU6g_dALWdJvqOfIrLu806h3jbc0W73k";
var BOT_URL       = "https://api.telegram.org/bot" + TOKEN;
var GROUP_CHAT_ID = "-5107316787";


// ============================================================
// SECTION 2: DATE HELPERS
// ============================================================

function parseDate(str) {
  if (!str) return null;
  str = str.toString().trim();
  var parts = str.split("-");
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
}

function isValidDate(str) {
  if (!str) return false;
  var parts = str.toString().trim().split("-");
  if (parts.length !== 3) return false;
  var d = parseInt(parts[0]);
  var m = parseInt(parts[1]);
  var y = parseInt(parts[2]);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return false;
  if (d < 1 || d > 31)      return false;
  if (m < 1 || m > 12)      return false;
  if (y < 2020 || y > 2099) return false;
  return true;
}

function toDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    var d = new Date(val);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  var str = val.toString().trim();
  if (str.match(/^\d{2}-\d{2}-\d{4}$/)) return parseDate(str);
  var d = new Date(str);
  if (!isNaN(d.getTime())) { d.setHours(0, 0, 0, 0); return d; }
  return null;
}

function today0() {
  var d = new Date(); d.setHours(0, 0, 0, 0); return d;
}


// ============================================================
// SECTION 3: SHEET READ HELPER (in-memory cache)
// ============================================================

var _cache = {};

function readSheet(name) {
  if (_cache[name]) return _cache[name];
  var sheet = SS.getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  _cache[name] = sheet.getDataRange().getValues();
  return _cache[name];
}

function clearCache() { _cache = {}; }


// ============================================================
// SECTION 4: TELEGRAM FUNCTIONS
// ============================================================

function sendTelegram(chatId, msg) {
  try {
    UrlFetchApp.fetch(BOT_URL + "/sendMessage", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: msg }),
      muteHttpExceptions: true
    });
  } catch(err) { Logger.log("sendTelegram Error: " + err); }
}

function sendKB(chatId, msg, kb) {
  try {
    UrlFetchApp.fetch(BOT_URL + "/sendMessage", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId, text: msg,
        reply_markup: JSON.stringify(kb)
      }),
      muteHttpExceptions: true
    });
  } catch(err) { Logger.log("sendKB Error: " + err); }
}

function sendTelegramWithButtons(chatId, msg) {
  if (!chatId) chatId = GROUP_CHAT_ID;
  sendKB(chatId, msg, { inline_keyboard: [
    [
      { text: "✅ Accept", callback_data: "accept_confirm" },
      { text: "❌ Refuse", callback_data: "refuse_confirm" },
      { text: "⏭ Skip",   callback_data: "skip_confirm"   }
    ],
    [
      { text: "🔄 Refresh",  callback_data: "refresh" },
      { text: "📋 Options",  callback_data: "menu"    }
    ]
  ]});
}

function sendMenu(chatId) {
  sendKB(chatId, "📋 Options", { inline_keyboard: [
    [{ text: "📅 Today Shift",    callback_data: "today"       }],
    [{ text: "📆 Tomorrow Shift", callback_data: "tomorrow"    }],
    [{ text: "📝 Leave Entry",    callback_data: "leave"       }],
    [{ text: "🔄 Mutual Entry",   callback_data: "mutual_menu" }],
    [{ text: "📊 OT History",     callback_data: "ot_history"  }],
    [{ text: "🔙 Dashboard",      callback_data: "dashboard"   }]
  ]});
}

function answerCallback(callbackId) {
  try {
    UrlFetchApp.fetch(BOT_URL + "/answerCallbackQuery", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ callback_query_id: callbackId }),
      muteHttpExceptions: true
    });
  } catch(err) {}
}


// ============================================================
// SECTION 5: USER STATE
// ============================================================

function getUserState(chatId) {
  var raw = PropertiesService.getScriptProperties().getProperty("st_" + chatId);
  return raw ? JSON.parse(raw) : null;
}

function setUserState(chatId, state) {
  PropertiesService.getScriptProperties().setProperty("st_" + chatId, JSON.stringify(state));
}

function clearUserState(chatId) {
  PropertiesService.getScriptProperties().deleteProperty("st_" + chatId);
}


// ============================================================
// SECTION 6: doPost — TELEGRAM + APPSHEET COMBINED
// ============================================================

function doPost(e) {
  var output = HtmlService.createHtmlOutput("OK");
  try {
    if (!e || !e.postData) return output;
    var body = JSON.parse(e.postData.contents);

    // ── TELEGRAM ──
    if (body.update_id) {
      var uid   = body.update_id;
      var props = PropertiesService.getScriptProperties();
      var last  = parseInt(props.getProperty("lu") || "0");
      if (uid <= last) return output;
      props.setProperty("lu", uid.toString());

      if (body.message) {
        var chatId = body.message.chat.id;
        var text   = (body.message.text || "").trim();
        var state  = getUserState(chatId);

        if (state && state.flow === "leave_apply") {
          if      (state.step === "uec")       { askLeaveFromDate(chatId, text); }
          else if (state.step === "from_date") { askLeaveToDate(chatId, text);   }
          else if (state.step === "to_date")   { confirmLeaveApply(chatId, text);}
          return output;
        }

        if      (text === "/start" || text === "/refresh") { sendTelegramWithButtons(chatId, buildOTMessage()); }
        else if (text === "/menu")     { sendMenu(chatId); }
        else if (text === "/status")   { sendTelegram(chatId, buildOTMessage()); }
        else if (text === "/tomorrow") { sendTomorrowShiftMessage(chatId); }
      }

      if (body.callback_query) {
        var action     = body.callback_query.data;
        var chatId     = body.callback_query.message.chat.id;
        var callbackId = body.callback_query.id;
        answerCallback(callbackId);

        if      (action === "accept_confirm")              { showAcceptConfirm(chatId); }
        else if (action === "refuse_confirm")              { showRefuseConfirm(chatId); }
        else if (action === "skip_confirm")                { showSkipConfirm(chatId);   }
        else if (action === "accept_do")                   { handleAccept(chatId); }
        else if (action === "refuse_do")                   { handleRefuse(chatId); }
        else if (action === "skip_do")                     { handleSkip(chatId);   }
        else if (action === "undo_ot")                     { handleUndoOT(chatId); }
        else if (action === "refresh" || action === "dashboard") { sendTelegramWithButtons(chatId, buildOTMessage()); }
        else if (action === "menu")                        { sendMenu(chatId); }
        else if (action === "today")                       { sendTodayShiftMessage(chatId); }
        else if (action === "tomorrow")                    { sendTomorrowShiftMessage(chatId); }
        else if (action === "leave")                       { sendLeaveMenu(chatId); }
        else if (action === "leave_apply")                 { startLeaveApply(chatId); }
        else if (action === "leave_submit")                { submitLeave(chatId); }
        else if (action === "leave_cancel")                { cancelLeaveFlow(chatId); }
        else if (action === "mutual_menu")                 { sendMutualMenu(chatId); }
        else if (action === "mutual_start")                { mutualStep1(chatId); }
        else if (action.indexOf("me1_") === 0)             { mutualStep2(chatId, action); }
        else if (action.indexOf("me2_") === 0)             { mutualStep3(chatId, action); }
        else if (action.indexOf("md_")  === 0)             { mutualStep4(chatId, action); }
        else if (action === "mutual_submit")               { mutualSubmit(chatId); }
        else if (action === "mutual_cancel")               { mutualCancel(chatId); }
        else if (action === "ot_history")                  { sendOTHistory(chatId); }
      }
    }

    // ── APPSHEET ──
    else if (body.action) {
      var lock = LockService.getScriptLock();
      if (!lock.tryLock(3000)) {
        Logger.log("AppSheet: Lock nahi mila");
        return output;
      }
      try {
        var seqSheet   = SS.getSheetByName(SHEETS.SEQ);
        var seqData    = seqSheet.getDataRange().getValues();
        var nextName   = seqData.length < 2 ? null : seqData[1][1];
        var nextStatus = seqData.length < 2 ? "" : (seqData[1][2] || "").toString().trim();

        Logger.log("AppSheet | action:" + body.action + " | next:" + nextName + " | status:'" + nextStatus + "'");

        // Sirf NEXT wale ka status match kare tab hi chalao
        if (!nextName || nextStatus !== body.action) {
          Logger.log("AppSheet: Skip — mismatch");
          return output;
        }

        // TURANT status clear — dobara trigger nahi hoga
        seqSheet.getRange(2, 3).setValue("");
        clearCache();

        if      (body.action === "accept") { doAcceptAS(nextName); }
        else if (body.action === "refuse") { doRefuseAS(nextName); }
        else if (body.action === "skip")   { doSkipAS(nextName);   }

      } finally {
        lock.releaseLock();
      }
    }

    // ── APPSHEET LEGACY ──
    else if (body.type) {
      if      (body.type === "leaveUpdate") { updateOTStatus(); notifyIfShortage(); }
      else if (body.type === "otAction")    { processOT(); }
      else if (body.type === "mutual")      { applyMutualDuty(); }
    }

  } catch(err) { Logger.log("doPost Error: " + err); }
  return output;
}


// ============================================================
// SECTION 6B: APPSHEET ACTION FUNCTIONS
// ============================================================

function doAcceptAS(name) {
  var info = getShiftShortage();
  if (!info) {
    sendTelegram(GROUP_CHAT_ID, "⚠️ AppSheet: Koi shortage nahi hai.");
    return;
  }

  // Today_Shift mein add
  SS.getSheetByName(SHEETS.TODAY).appendRow([name, info.shift]);

  // OT Tag
  addOTTag(name, info.shift);

  // Queue rotate
  rotateQueue(name);

  // OT Date
  var otDate = new Date();
  if (info.isTmr) otDate.setDate(otDate.getDate() + 1);

  // Log
  SS.getSheetByName(SHEETS.LOG).appendRow([
    new Date(), name, info.shift, "OT Assigned (AppSheet)", otDate
  ]);

  clearCache();
  updateOTStatus();

  sendTelegram(GROUP_CHAT_ID,
    "✅ OT Assigned (AppSheet)\n" +
    "👤 " + name + "\n" +
    "⚡ " + info.shift + (info.isTmr ? " (Kal)" : " (Aaj)")
  );
}

function doRefuseAS(name) {
  SS.getSheetByName(SHEETS.LOG).appendRow([
    new Date(), name, "", "Refuse (AppSheet)", ""
  ]);
  rotateQueue(name);
  clearCache();
  sendTelegram(GROUP_CHAT_ID, "❌ Refused (AppSheet)\n👤 " + name);
}

function doSkipAS(name) {
  moveDownByRequirement(name, 1);
  clearCache();
  sendTelegram(GROUP_CHAT_ID, "⏭ Skip (AppSheet)\n👤 " + name);
}


// ============================================================
// SECTION 7: OT MESSAGE
// ============================================================

function buildOTMessage() {
  try {
    var shortages  = getAllShortages();
    var candidates = getOTCandidates();
    var info       = getShiftShortage();
    var timeSlot   = getCurrentTimeSlot();
    var anyShortage = false;

    var msg = "⚡ Electrical OT Status\n";
    msg += timeSlot + "\n";
    msg += "─────────────────────\n";

    shortages.forEach(function(s) {
      var sh    = s.shift;
      var isTmr = s.isTmr;
      var req   = isTmr ? s.data.reqM : s.data.reqT;
      var pre   = isTmr ? s.data.preM : s.data.preT;
      var sho   = isTmr ? s.data.shoM : s.data.shoT;
      var day   = isTmr ? " (Kal)" : " (Aaj)";
      var tag   = sho > 0 ? " ⚠️" : " ✅";

      msg += "\n🔹 " + sh + day + tag + "\n";
      msg += "   Req:" + req + "  Pre:" + pre + "  Sho:" + sho + "\n";

      if (sho > 0) anyShortage = true;
    });

    msg += "─────────────────────\n";

    if (!anyShortage) {
      msg += "\n✅ Koi shortage nahi hai.";
      return msg;
    }

    if (info) {
      var dayLabel = info.isTmr ? " (Kal)" : " (Aaj)";
      msg += "\n🎯 OT: " + info.shift + dayLabel + "\n";
    }

    if (candidates.length > 0) {
      msg += "\n👤 Next: " + candidates[0].name + " (" + candidates[0].shift + ")\n";
      if (candidates.length > 1) {
        msg += "\n📋 Waiting:\n";
        for (var j = 1; j < candidates.length; j++) {
          msg += (j+1) + ". " + candidates[j].name + " (" + candidates[j].shift + ")\n";
        }
      }
    } else {
      msg += "\n⚠️ Koi candidate nahi.";
    }
    return msg;
  } catch(err) { return "⚠️ Error: " + err; }
}


// ============================================================
// SECTION 8: TODAY SHIFT
// ============================================================

function buildTodayShiftMessage() {
  try {
    var data   = readSheet(SHEETS.TODAY);
    var groups = { "1st":[], "2nd":[], "3rd":[], "REST":[] };
    for (var i = 1; i < data.length; i++) {
      var n = data[i][0], s = data[i][1];
      if (n && s && groups[s]) groups[s].push(n);
    }
    var date = Utilities.formatDate(new Date(), "GMT+5:30", "dd-MM-yyyy");
    var msg  = "📅 Today Shift — " + date + "\n─────────────────────\n";
    var em   = { "1st":"🔵","2nd":"🟢","3rd":"🟠","REST":"🔴" };
    ["1st","2nd","3rd","REST"].forEach(function(s) {
      msg += "\n" + em[s] + " " + s + " (" + groups[s].length + ")\n";
      if (!groups[s].length) { msg += "  — Koi nahi\n"; }
      else groups[s].forEach(function(n, i) { msg += "  " + (i+1) + ". " + n + "\n"; });
    });
    msg += "\n─────────────────────\n👥 Total: " + (data.length - 1);
    return msg;
  } catch(err) { return "⚠️ Error: " + err; }
}

function sendTodayShiftMessage(chatId) {
  sendKB(chatId, buildTodayShiftMessage(), {
    inline_keyboard: [
      [{ text: "📆 Tomorrow Shift", callback_data: "tomorrow" }],
      [{ text: "🔙 Options",        callback_data: "menu"     }],
      [{ text: "🏠 Dashboard",      callback_data: "dashboard"}]
    ]
  });
}


// ============================================================
// SECTION 8B: TOMORROW DUTY
// ============================================================

function buildTomorrowDuty() {
  try {
    // Kal ki date
    var tmr = new Date();
    tmr.setDate(tmr.getDate() + 1);
    tmr.setHours(0, 0, 0, 0);
    var tmrStr = Utilities.formatDate(tmr, "GMT+5:30", "dd-MM-yyyy");

    // Step 1: Duty_Auto se kal ka base shift lo
    var duty   = readSheet(SHEETS.DUTY);
    var shifts = {}; // name → shift

    for (var i = 1; i < duty.length; i++) {
      var dt = Utilities.formatDate(new Date(duty[i][0]), "GMT+5:30", "dd-MM-yyyy");
      if (dt === tmrStr) {
        shifts[duty[i][1]] = duty[i][2]; // name → shift
      }
    }

    if (Object.keys(shifts).length === 0) {
      return "⚠️ Kal ka duty chart nahi mila.\n(" + tmrStr + ")";
    }

    // Step 2: Leave check — naam ke aage (Leave) lagao
    var leave = readSheet(SHEETS.LEAVE);
    for (var i = 1; i < leave.length; i++) {
      var emp  = leave[i][2];
      var from = toDate(leave[i][4]);
      var to   = toDate(leave[i][5]);
      if (!emp || !from || !to) continue;
      if (tmr >= from && tmr <= to) {
        if (shifts[emp] !== undefined) {
          shifts[emp + " (Leave)"] = shifts[emp];
          delete shifts[emp];
        }
      }
    }

    // Step 3: Mutual check — shift swap + (Mutual) tag
    var mutual = readSheet(SHEETS.MUTUAL);
    for (var i = 1; i < mutual.length; i++) {
      var e1   = mutual[i][1];
      var e2   = mutual[i][2];
      var dt   = toDate(mutual[i][5]);
      var st   = mutual[i][6];
      if (!dt || st !== "Approved") continue;
      if (dt.getTime() !== tmr.getTime()) continue;

      // Find current shifts (clean name)
      var s1 = null, s2 = null;
      var k1 = null, k2 = null;

      Object.keys(shifts).forEach(function(k) {
        var clean = k.replace(/ \(Leave\)| \(OT\)| \(Mutual\)/g, "").trim();
        if (clean === e1) { s1 = shifts[k]; k1 = k; }
        if (clean === e2) { s2 = shifts[k]; k2 = k; }
      });

      if (k1 && k2 && s1 && s2 && s1 !== s2) {
        // Swap karo
        delete shifts[k1];
        delete shifts[k2];
        shifts[e1 + " (Mutual)"] = s2;
        shifts[e2 + " (Mutual)"] = s1;
      }
    }

    // Step 4: OT Log se kal ke OT check karo — Column E = OT_Date
    var log = readSheet(SHEETS.LOG);
    for (var i = 1; i < log.length; i++) {
      var otDate = toDate(log[i][4]); // Column E = OT_Date
      var emp    = log[i][1];
      var sh     = log[i][2];
      var act    = log[i][3];
      if (!otDate || !emp) continue;
      if (otDate.getTime() !== tmr.getTime()) continue;
      if (act !== "OT Assigned") continue;

      // OT tag lagao
      var found = false;
      Object.keys(shifts).forEach(function(k) {
        var clean = k.replace(/ \(Leave\)| \(OT\)| \(Mutual\)/g, "").trim();
        if (clean === emp) {
          shifts[emp + " (OT)"] = sh;
          delete shifts[k];
          found = true;
        }
      });
      // Nahi mila toh add karo
      if (!found) {
        shifts[emp + " (OT)"] = sh;
      }
    }

    // Step 5: Group by shift
    var groups = { "1st":[], "2nd":[], "3rd":[], "REST":[] };
    Object.keys(shifts).forEach(function(name) {
      var sh = shifts[name];
      if (groups[sh]) groups[sh].push(name);
      else groups["REST"].push(name); // unknown shift
    });

    // Step 6: Message build karo
    var msg  = "📆 Tomorrow Duty — " + tmrStr + "\n─────────────────────\n";
    var em   = { "1st":"🔵","2nd":"🟢","3rd":"🟠","REST":"🔴" };
    var tags = {
      "OT":    "⚡",
      "Leave": "🏥",
      "Mutual":"🔄"
    };

    ["1st","2nd","3rd","REST"].forEach(function(sh) {
      var emps = groups[sh];
      msg += "\n" + em[sh] + " " + sh + " Shift (" + emps.length + ")\n";
      if (!emps.length) {
        msg += "  — Koi nahi\n";
      } else {
        emps.sort().forEach(function(n, i) {
          // Tag emoji add karo
          var tag = "";
          if (n.indexOf("(OT)")    > -1) tag = " ⚡";
          if (n.indexOf("(Leave)") > -1) tag = " 🏥";
          if (n.indexOf("(Mutual)")> -1) tag = " 🔄";
          msg += "  " + (i+1) + ". " + n + tag + "\n";
        });
      }
    });

    msg += "\n─────────────────────";
    msg += "\n👥 Total: " + Object.keys(shifts).length;
    msg += "\n\n⚡ OT  🏥 Leave  🔄 Mutual";

    return msg;

  } catch(err) {
    Logger.log("buildTomorrowDuty: " + err);
    return "⚠️ Error: " + err.toString();
  }
}

function sendTomorrowShiftMessage(chatId) {
  sendKB(chatId, buildTomorrowDuty(), {
    inline_keyboard: [
      [{ text: "📅 Today Shift", callback_data: "today"    }],
      [{ text: "🔙 Options",     callback_data: "menu"     }],
      [{ text: "🏠 Dashboard",   callback_data: "dashboard"}]
    ]
  });
}


// ============================================================
// SECTION 9: LEAVE FLOW
// ============================================================

function buildTodayLeaveMessage() {
  try {
    var leave = readSheet(SHEETS.LEAVE);
    var td    = today0();
    var list  = [];
    for (var i = 1; i < leave.length; i++) {
      var emp = leave[i][2], fr = toDate(leave[i][4]), to = toDate(leave[i][5]);
      if (!emp || !fr || !to) continue;
      if (td >= fr && td <= to) {
        list.push({
          emp:  emp,
          from: Utilities.formatDate(fr, "GMT+5:30", "dd-MM-yyyy"),
          to:   Utilities.formatDate(to, "GMT+5:30", "dd-MM-yyyy")
        });
      }
    }
    var msg = "📝 Leave — " + Utilities.formatDate(td, "GMT+5:30", "dd-MM-yyyy") + "\n─────────────────────\n";
    if (!list.length) {
      msg += "\n✅ Aaj koi leave nahi hai.";
    } else {
      msg += "\n🔴 Leave Par:\n";
      list.forEach(function(l, i) {
        msg += (i+1) + ". " + l.emp + "\n   📅 " + l.from + " → " + l.to + "\n";
      });
      msg += "\n👥 Total: " + list.length;
    }
    return msg;
  } catch(err) { return "⚠️ Error: " + err; }
}

function sendLeaveMenu(chatId) {
  sendKB(chatId, buildTodayLeaveMessage(), { inline_keyboard: [
    [{ text: "📝 Leave Apply", callback_data: "leave_apply" }],
    [{ text: "🔙 Options",     callback_data: "menu"        }],
    [{ text: "🏠 Dashboard",   callback_data: "dashboard"   }]
  ]});
}

function startLeaveApply(chatId) {
  setUserState(chatId, { flow:"leave_apply", step:"uec", uec:null, name:null, from:null, to:null });
  sendKB(chatId, "📝 Leave Apply\n\n🔢 Step 1/3\nUEC Number enter karo:", {
    inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
  });
}

function askLeaveFromDate(chatId, uec) {
  var empName = getEmployeeByUEC(uec);
  if (!empName) {
    sendKB(chatId, "⚠️ UEC '" + uec + "' nahi mila!\nDobara enter karo:", {
      inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
    });
    return;
  }
  setUserState(chatId, { flow:"leave_apply", step:"from_date", uec:uec, name:empName, from:null, to:null });
  sendKB(chatId, "✅ " + empName + "\n\n📅 Step 2/3\nFrom Date (DD-MM-YYYY):", {
    inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
  });
}

function askLeaveToDate(chatId, fromDate) {
  var state = getUserState(chatId);
  if (!isValidDate(fromDate)) {
    sendKB(chatId, "⚠️ Galat format!\nDD-MM-YYYY mein likho:", {
      inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
    });
    return;
  }
  state.step = "to_date"; state.from = fromDate;
  setUserState(chatId, state);
  sendKB(chatId, "📅 From: " + fromDate + "\n\n📅 Step 3/3\nTo Date (DD-MM-YYYY):", {
    inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
  });
}

function confirmLeaveApply(chatId, toDate) {
  var state = getUserState(chatId);
  if (!isValidDate(toDate)) {
    sendKB(chatId, "⚠️ Galat format!\nDD-MM-YYYY mein likho:", {
      inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
    });
    return;
  }
  var fp = parseDate(state.from), tp = parseDate(toDate);
  if (tp < fp) {
    sendKB(chatId, "⚠️ To Date, From se pehle nahi!\nDobara To Date enter karo:", {
      inline_keyboard: [[{ text:"❌ Cancel", callback_data:"leave_cancel" }]]
    });
    return;
  }
  state.step = "confirm"; state.to = toDate;
  setUserState(chatId, state);
  var days = Math.floor((tp - fp) / 86400000) + 1;
  var msg  = "📋 Confirm\n─────────────────────\n";
  msg += "👤 " + state.name + "\n🔢 " + state.uec + "\n";
  msg += "📅 " + state.from + " → " + state.to + "\n📆 " + days + " day(s)";
  sendKB(chatId, msg, { inline_keyboard: [[
    { text:"✅ Submit", callback_data:"leave_submit" },
    { text:"❌ Cancel", callback_data:"leave_cancel" }
  ]]});
}

function submitLeave(chatId) {
  var state = getUserState(chatId);
  if (!state || state.flow !== "leave_apply") { sendTelegram(chatId, "⚠️ Koi active request nahi."); return; }
  try {
    var fr   = parseDate(state.from), to = parseDate(state.to);
    var days = Math.floor((to - fr) / 86400000) + 1;
    SS.getSheetByName(SHEETS.LEAVE).appendRow([Utilities.getUuid(), state.uec, state.name, "", fr, to, ""]);
    clearCache();
    clearUserState(chatId);
    updateOTStatus();
    var msg = "✅ Leave Ho Gayi!\n👤 " + state.name + "\n📅 " + state.from + " → " + state.to + "\n📆 " + days + " day(s)\n\n" + buildTodayLeaveMessage();
    sendKB(chatId, msg, { inline_keyboard: [
      [{ text:"📝 Aur Apply", callback_data:"leave_apply" }],
      [{ text:"🔙 Menu",     callback_data:"menu"        }]
    ]});
    notifyIfShortage();
  } catch(err) { Logger.log("submitLeave: " + err); sendTelegram(chatId, "⚠️ Error: " + err); clearUserState(chatId); }
}

function cancelLeaveFlow(chatId) {
  clearUserState(chatId);
  sendKB(chatId, "❌ Cancel.", { inline_keyboard: [
    [{ text:"📝 Leave Menu", callback_data:"leave"    }],
    [{ text:"🔙 Options",    callback_data:"menu"     }],
    [{ text:"🏠 Dashboard",  callback_data:"dashboard"}]
  ]});
}


// ============================================================
// SECTION 10: MUTUAL ENTRY FLOW
// ============================================================

function buildMutualStatusMsg() {
  try {
    var data = readSheet(SHEETS.MUTUAL);
    var td   = today0();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var dt = toDate(data[i][5]), st = data[i][6] || "Pending";
      if (!dt || dt < td) continue;
      list.push({
        e1: data[i][1], e2: data[i][2],
        d:  Utilities.formatDate(dt, "GMT+5:30", "dd-MM-yyyy"),
        s:  st
      });
    }
    var msg = "🔄 Mutual Status\n─────────────────────\n";
    if (!list.length) {
      msg += "\n✅ Koi pending mutual nahi.";
    } else {
      list.forEach(function(l) {
        var e = l.s === "Applied" ? "✅" : l.s.indexOf("Blocked") === 0 ? "❌" : "⏳";
        msg += "\n" + e + " " + l.e1 + " ↔ " + l.e2 + "\n   📅 " + l.d + " | " + l.s + "\n";
      });
    }
    return msg;
  } catch(err) { return "⚠️ Error: " + err; }
}

function sendMutualMenu(chatId) {
  sendKB(chatId, buildMutualStatusMsg(), { inline_keyboard: [
    [{ text:"🔄 Mutual Apply", callback_data:"mutual_start" }],
    [{ text:"🔙 Options",      callback_data:"menu"         }],
    [{ text:"🏠 Dashboard",    callback_data:"dashboard"    }]
  ]});
}

// Step 1: Emp1 select karo
function mutualStep1(chatId) {
  var list = getEmployeeList();
  if (!list.length) { sendTelegram(chatId, "⚠️ Employee list empty."); return; }
  setUserState(chatId, { flow:"mutual", step:"e1", e1:null, e2:null, date:null });
  var rows = makeEmpRows(list, "me1_", null);
  rows.push([{ text:"❌ Cancel", callback_data:"mutual_cancel" }]);
  sendKB(chatId, "🔄 Mutual\n\n👤 Step 1/3\nPehla Employee select karo:", { inline_keyboard: rows });
}

// Step 2: Emp2 select karo
function mutualStep2(chatId, action) {
  var uec1  = action.replace("me1_", "");
  var name1 = getEmployeeByUEC(uec1);
  var state = getUserState(chatId);
  state.step = "e2"; state.e1 = { uec:uec1, name:name1 };
  setUserState(chatId, state);
  var rows = makeEmpRows(getEmployeeList(), "me2_", uec1);
  rows.push([{ text:"❌ Cancel", callback_data:"mutual_cancel" }]);
  sendKB(chatId, "✅ Emp1: " + name1 + "\n\n👤 Step 2/3\nDusra Employee select karo:", { inline_keyboard: rows });
}

// Step 3: Date select karo
function mutualStep3(chatId, action) {
  var uec2  = action.replace("me2_", "");
  var name2 = getEmployeeByUEC(uec2);
  var state = getUserState(chatId);
  state.step = "date"; state.e2 = { uec:uec2, name:name2 };
  setUserState(chatId, state);

  var rows = [], row = [], td = new Date();
  for (var d = 0; d < 7; d++) {
    var dt  = new Date(td); dt.setDate(td.getDate() + d);
    var lbl = d===0 ? "📌 Today" : d===1 ? "➡️ Tomorrow" : Utilities.formatDate(dt, "GMT+5:30", "EEE dd-MM");
    var val = Utilities.formatDate(dt, "GMT+5:30", "dd-MM-yyyy");
    row.push({ text:lbl, callback_data:"md_"+val });
    if (row.length === 2) { rows.push(row); row = []; }
  }
  if (row.length) rows.push(row);
  rows.push([{ text:"❌ Cancel", callback_data:"mutual_cancel" }]);
  sendKB(chatId,
    "✅ Emp1: " + state.e1.name + "\n✅ Emp2: " + name2 + "\n\n📅 Step 3/3\nDate select karo:",
    { inline_keyboard: rows }
  );
}

// Step 4: Confirm
function mutualStep4(chatId, action) {
  var date  = action.replace("md_", "");
  var state = getUserState(chatId);
  state.step = "confirm"; state.date = date;
  setUserState(chatId, state);
  var msg = "📋 Mutual Confirm\n─────────────────────\n";
  msg += "👤 Emp1 : " + state.e1.name + "\n";
  msg += "👤 Emp2 : " + state.e2.name + "\n";
  msg += "📅 Date : " + date + "\n─────────────────────\nShift swap hoga";
  sendKB(chatId, msg, { inline_keyboard: [[
    { text:"✅ Submit", callback_data:"mutual_submit" },
    { text:"❌ Cancel", callback_data:"mutual_cancel" }
  ]]});
}

// Submit
function mutualSubmit(chatId) {
  var state = getUserState(chatId);
  if (!state || state.flow !== "mutual") { sendTelegram(chatId, "⚠️ Koi active request nahi."); return; }
  try {
    var mDate   = parseDate(state.date);
    var td      = today0();
    var isToday = mDate.getTime() === td.getTime();

    // Sheet: ID | Emp1 | Emp2 | UEC1 | UEC2 | Date | Status
    SS.getSheetByName(SHEETS.MUTUAL).appendRow([
      Utilities.getUuid(),
      state.e1.name, state.e2.name,
      state.e1.uec,  state.e2.uec,
      mDate, "Approved"
    ]);

    clearCache();
    if (isToday) swapTodayShift(state.e1.name, state.e2.name);
    clearUserState(chatId);

    var msg = "✅ Mutual Ho Gaya!\n👤 " + state.e1.name + " ↔ " + state.e2.name + "\n📅 " + state.date;
    if (isToday) {
      clearCache();
      msg += "\n\n🔄 Today Shift Updated!\n\n" + buildTodayShiftMessage();
    } else {
      msg += "\n\n⏳ Future date — Us din auto update hoga.";
    }
    sendKB(chatId, msg, { inline_keyboard: [
      [{ text:"🔄 Aur Mutual", callback_data:"mutual_start" }],
      [{ text:"🔙 Options",    callback_data:"menu"         }],
      [{ text:"🏠 Dashboard",  callback_data:"dashboard"    }]
    ]});
  } catch(err) { Logger.log("mutualSubmit: " + err); sendTelegram(chatId, "⚠️ Error: " + err); clearUserState(chatId); }
}

function mutualCancel(chatId) {
  clearUserState(chatId);
  sendKB(chatId, "❌ Mutual cancel.", { inline_keyboard: [
    [{ text:"🔄 Mutual Menu", callback_data:"mutual_menu" }],
    [{ text:"🔙 Options",     callback_data:"menu"        }],
    [{ text:"🏠 Dashboard",   callback_data:"dashboard"   }]
  ]});
}

// Aaj ka swap — batch write (2 calls only)
function swapTodayShift(n1, n2) {
  var sheet = SS.getSheetByName(SHEETS.TODAY);
  var data  = sheet.getDataRange().getValues();
  var r1 = -1, r2 = -1;
  for (var i = 1; i < data.length; i++) {
    var c = data[i][0].toString().replace(/ \(Mutual\)| \(OT\)/g, "").trim();
    if (c === n1) r1 = i;
    if (c === n2) r2 = i;
  }
  if (r1 === -1 || r2 === -1) return;
  var s1 = data[r1][1], s2 = data[r2][1];
  sheet.getRange(r1+1, 1, 1, 2).setValues([[n1 + " (Mutual)", s2]]);
  sheet.getRange(r2+1, 1, 1, 2).setValues([[n2 + " (Mutual)", s1]]);
  clearCache();
}

// Future mutuals auto apply (scheduled trigger mein call karo)
function applyFutureMutuals() {
  var sheet = SS.getSheetByName(SHEETS.MUTUAL);
  var data  = sheet.getDataRange().getValues();
  var td    = today0();
  for (var i = 1; i < data.length; i++) {
    var dt = toDate(data[i][5]), st = data[i][6];
    if (!dt || st === "Applied" || st === "Blocked") continue;
    if (dt.getTime() !== td.getTime() || st !== "Approved") continue;
    swapTodayShift(data[i][1], data[i][2]);
    sheet.getRange(i+1, 7).setValue("Applied");
  }
  clearCache();
}


// ============================================================
// SECTION 10B: OT HISTORY
// ============================================================

function sendOTHistory(chatId) {
  try {
    var log   = readSheet(SHEETS.LOG);
    var td    = today0();
    var f7    = new Date(td); f7.setDate(td.getDate() - 7);
    var t7    = new Date(td); t7.setDate(td.getDate() + 7);
    var list  = [];

    for (var i = 1; i < log.length; i++) {
      var dt = toDate(log[i][0]);
      if (!dt || dt < f7 || dt > t7) continue;
      list.push({
        dt:  dt,
        ds:  Utilities.formatDate(dt, "GMT+5:30", "dd-MM-yyyy"),
        n:   log[i][1] || "",
        s:   log[i][2] || "",
        a:   log[i][3] || log[i][2] || ""
      });
    }

    list.sort(function(a, b) { return b.dt - a.dt; });
    if (list.length > 15) list = list.slice(0, 15);

    var tds  = Utilities.formatDate(td, "GMT+5:30", "dd-MM-yyyy");
    var msg  = "📊 OT History (±7 Days)\n─────────────────────\n";
    var last = "";

    if (!list.length) {
      msg += "\n📭 Koi OT record nahi.";
    } else {
      list.forEach(function(l) {
        if (l.ds !== last) {
          var tag = l.ds === tds ? " ← Aaj" : l.dt < td ? " (Past)" : " (Upcoming)";
          msg += "\n📅 " + l.ds + tag + "\n";
          last = l.ds;
        }
        var e = l.a==="Accept"||l.a==="OT Assigned" ? "✅" :
                l.a==="Refuse"            ? "❌" :
                l.a==="Skip (Same Shift)" ? "⏭" :
                l.a==="Lapse (Leave)"     ? "🏥" : "📌";
        msg += e + " " + l.n + (l.s ? " — " + l.s : "") + "\n";
      });
      msg += "\n─────────────────────\n📝 " + list.length + " records";
    }

    sendKB(chatId, msg, {
      inline_keyboard: [
        [{ text: "🔙 Options",   callback_data: "menu"      }],
        [{ text: "🏠 Dashboard", callback_data: "dashboard" }]
      ]
    });
  } catch(err) { sendTelegram(chatId, "⚠️ Error: " + err); }
}


// ============================================================
// SECTION 11: EMPLOYEE HELPERS
// ============================================================

function getEmployeeByUEC(uec) {
  try {
    var emp = readSheet(SHEETS.EMP);
    uec = uec.toString().trim().toUpperCase();
    for (var i = 1; i < emp.length; i++) {
      if (!emp[i][1] || !emp[i][3]) continue;
      if (emp[i][1].toString().trim().toUpperCase() === uec) return emp[i][3].toString().trim();
    }
    return null;
  } catch(err) { return null; }
}

function getEmployeeList() {
  try {
    var emp = readSheet(SHEETS.EMP), list = [];
    for (var i = 1; i < emp.length; i++) {
      if (!emp[i][1] || !emp[i][3]) continue;
      list.push({ uec: emp[i][1].toString().trim(), name: emp[i][3].toString().trim() });
    }
    return list;
  } catch(err) { return []; }
}

function makeEmpRows(list, prefix, excludeUec) {
  var rows = [];
  for (var i = 0; i < list.length; i += 2) {
    var row = [];
    if (list[i].uec !== excludeUec) row.push({ text: list[i].name, callback_data: prefix + list[i].uec });
    if (list[i+1] && list[i+1].uec !== excludeUec) row.push({ text: list[i+1].name, callback_data: prefix + list[i+1].uec });
    if (row.length) rows.push(row);
  }
  return rows;
}


// ============================================================
// SECTION 12: SHIFT HELPERS
// ============================================================

function getOTCandidates() {
  var seq = readSheet(SHEETS.SEQ), map = {};
  readSheet(SHEETS.TODAY).forEach(function(r, i) {
    if (i > 0) map[r[0].toString().replace(/ \(OT\)| \(Mutual\)/g, "").trim()] = r[1];
  });
  var list = [];
  for (var j = 1; j < seq.length && list.length < 5; j++) {
    list.push({ name: seq[j][1], shift: map[seq[j][1]] || "—" });
  }
  return list;
}

function getNextCandidate() {
  var seq = readSheet(SHEETS.SEQ);
  return seq.length < 2 ? null : seq[1][1];
}

// Time ke hisaab se shift priority order
function getShiftPriority() {
  var h = new Date().getHours();
  if (h >= 7  && h < 15) return ["2nd", "3rd", "1st"]; // 7am-3pm
  if (h >= 15 && h < 22) return ["3rd", "1st", "2nd"]; // 3pm-10pm
  return ["1st", "2nd", "3rd"];                          // 10pm-7am
}

function getCurrentTimeSlot() {
  var h = new Date().getHours();
  if (h >= 7  && h < 15) return "🌅 Morning (7am-3pm)";
  if (h >= 15 && h < 22) return "🌆 Evening (3pm-10pm)";
  return "🌙 Night (10pm-7am)";
}

// Kya yeh shift "kal ke liye" hai time ke hisaab se
function isTomorrowShift(shift) {
  var h        = new Date().getHours();
  var priority = getShiftPriority();
  // 7am-3pm: 1st = kal | 3pm-10pm: 1st,2nd = kal | 10pm-7am: sabhi = kal
  if (h >= 7  && h < 15) return shift === "1st";
  if (h >= 15 && h < 22) return shift === "1st" || shift === "2nd";
  return true; // 10pm-7am sab kal ke
}

// Priority order ke hisaab se pehli shortage wali shift
// Today + Tomorrow dono check karta hai
function getShiftShortage() {
  var status   = readSheet(SHEETS.STATUS);
  var priority = getShiftPriority();
  var map      = {};

  for (var i = 1; i < status.length; i++) {
    var sh = status[i][0];
    map[sh] = {
      // Today
      reqT: status[i][1] || 0,
      preT: status[i][2] || 0,
      shoT: status[i][3] || 0,
      // Tomorrow
      reqM: status[i][4] || 0,
      preM: status[i][5] || 0,
      shoM: status[i][6] || 0
    };
  }

  for (var p = 0; p < priority.length; p++) {
    var sh      = priority[p];
    var isTmr   = isTomorrowShift(sh);
    var shortage = isTmr ? (map[sh] ? map[sh].shoM : 0)
                         : (map[sh] ? map[sh].shoT : 0);
    if (shortage > 0) {
      return {
        shift:    sh,
        required: shortage,
        isTmr:    isTmr
      };
    }
  }
  return null;
}

// Saari shifts priority order mein today + tomorrow ke saath
function getAllShortages() {
  var status   = readSheet(SHEETS.STATUS);
  var priority = getShiftPriority();
  var map      = {};

  for (var i = 1; i < status.length; i++) {
    map[status[i][0]] = {
      reqT: status[i][1] || 0,
      preT: status[i][2] || 0,
      shoT: status[i][3] || 0,
      reqM: status[i][4] || 0,
      preM: status[i][5] || 0,
      shoM: status[i][6] || 0
    };
  }

  var result = [];
  for (var p = 0; p < priority.length; p++) {
    var sh = priority[p];
    if (map[sh]) result.push({ shift: sh, isTmr: isTomorrowShift(sh), data: map[sh] });
  }
  return result;
}

function detectShiftShortage() {
  var today = readSheet(SHEETS.TODAY), req = readSheet(SHEETS.REQ);
  var count = { "1st":0, "2nd":0, "3rd":0 };
  for (var i = 1; i < today.length; i++) {
    var s = today[i][1];
    if (s && s !== "REST" && count[s] !== undefined) count[s]++;
  }
  var out = {};
  for (var i = 1; i < req.length; i++) {
    if (count[req[i][0]] < req[i][1]) out[req[i][0]] = req[i][1] - count[req[i][0]];
  }
  return out;
}

function getEmployeeShift(name) {
  var today = readSheet(SHEETS.TODAY);
  name = name.toString().trim().toLowerCase();
  for (var i = 1; i < today.length; i++) {
    var n = today[i][0].toString().replace(/ \(OT\)| \(Mutual\)/g, "").trim().toLowerCase();
    if (n === name) return today[i][1];
  }
  return null;
}

function isOnLeave(name) {
  var leave = readSheet(SHEETS.LEAVE), td = today0();
  for (var i = 1; i < leave.length; i++) {
    if (!leave[i][2]) continue;
    var fr = toDate(leave[i][4]), to = toDate(leave[i][5]);
    if (!fr || !to) continue;
    if (leave[i][2].toString().trim() === name.toString().trim() && td >= fr && td <= to) return true;
  }
  return false;
}


// ============================================================
// SECTION 13: OT CONFIRM + ACTIONS
// ============================================================

// Confirm screens — pehle employee + shift dikhao
function showAcceptConfirm(chatId) {
  var name = getNextCandidate();
  var info = getShiftShortage();

  // Requirement full hai — restrict karo
  if (!info) {
    sendKB(chatId, "✅ Koi shortage nahi hai — OT assign karne ki zaroorat nahi.", {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
    return;
  }
  if (!name) {
    sendKB(chatId, "⚠️ Koi eligible candidate nahi.", {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
    return;
  }

  var empShift  = getEmployeeShift(name) || "—";
  var dayLabel  = info.isTmr ? "Kal" : "Aaj";
  var otDate    = new Date();
  if (info.isTmr) otDate.setDate(otDate.getDate() + 1);
  var otDateStr = Utilities.formatDate(otDate, "GMT+5:30", "dd-MM-yyyy");

  var msg = "✅ OT Accept Confirm?\n─────────────────────\n";
  msg += "👤 Employee  : " + name       + "\n";
  msg += "🕐 Aaj Shift : " + empShift   + "\n";
  msg += "⚡ OT Shift  : " + info.shift + " (" + dayLabel + ")\n";
  msg += "📅 OT Date   : " + otDateStr  + "\n";
  msg += "─────────────────────\nConfirm karo?";

  sendKB(chatId, msg, { inline_keyboard: [[
    { text: "✅ Haan Accept", callback_data: "accept_do" },
    { text: "🔙 Cancel",      callback_data: "dashboard" }
  ]]});
}

function showRefuseConfirm(chatId) {
  var name = getNextCandidate();
  var info = getShiftShortage();
  if (!name) { sendTelegramWithButtons(chatId, buildOTMessage()); return; }

  var msg = "❌ OT Refuse Confirm?\n─────────────────────\n";
  msg += "👤 Employee : " + name + "\n";
  if (info) msg += "⚡ Shift    : " + info.shift + "\n";
  msg += "─────────────────────\n";
  msg += "Refuse karke next candidate pe move karein?";

  sendKB(chatId, msg, { inline_keyboard: [[
    { text: "❌ Haan Refuse", callback_data: "refuse_do" },
    { text: "🔙 Cancel",     callback_data: "dashboard" }
  ]]});
}

function showSkipConfirm(chatId) {
  var name = getNextCandidate();
  if (!name) { sendTelegramWithButtons(chatId, buildOTMessage()); return; }

  var empShift = getEmployeeShift(name) || "—";
  var msg = "⏭ OT Skip Confirm?\n─────────────────────\n";
  msg += "👤 Employee : " + name     + "\n";
  msg += "🕐 Shift    : " + empShift + "\n";
  msg += "─────────────────────\n";
  msg += "Queue mein neeche move karein?";

  sendKB(chatId, msg, { inline_keyboard: [[
    { text: "⏭ Haan Skip", callback_data: "skip_do"   },
    { text: "🔙 Cancel",   callback_data: "dashboard" }
  ]]});
}

// Actual actions
function handleAccept(chatId) {
  // ── LOCK — double press prevent ──
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000); // 3 sec wait, agar dusra chal raha hai toh
  } catch(e) {
    // Lock nahi mila — dusra press already process ho raha hai
    sendKB(chatId, "⏳ Pehla request process ho raha hai, dobara mat dabao.", {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
    return;
  }

  try {
    var name = getNextCandidate(), info = getShiftShortage();
    if (!info) {
      sendKB(chatId, "✅ Koi shortage nahi hai — OT assign karne ki zaroorat nahi.", {
        inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
      });
      return;
    }
    if (!name) {
      sendKB(chatId, "⚠️ Koi eligible candidate nahi.", {
        inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
      });
      return;
    }

    // ── DOUBLE ACCEPT GUARD — same employee last 30 sec mein assign hua? ──
    var logSheet = SS.getSheetByName(SHEETS.LOG);
    var logData  = logSheet.getDataRange().getValues();
    var now      = new Date().getTime();
    for (var i = logData.length - 1; i >= 1; i--) {
      var logTime = new Date(logData[i][0]).getTime();
      if (now - logTime > 30000) break; // 30 sec se purana = stop
      if (logData[i][1] === name && logData[i][3] === "OT Assigned") {
        sendKB(chatId, "⚠️ " + name + " ka OT abhi abhi assign hua hai!\nDouble press — ignore.", {
          inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
        });
        return;
      }
    }

    assignOT(name, info.shift);
    rotateQueue(name);
    addOTTag(name, info.shift);

    var otDate = new Date();
    if (info.isTmr) otDate.setDate(otDate.getDate() + 1);
    var otDateStr = Utilities.formatDate(otDate, "GMT+5:30", "dd-MM-yyyy");
    var dayLabel  = info.isTmr ? "Kal" : "Aaj";

    logSheet.appendRow([
      new Date(), name, info.shift, "OT Assigned", otDate
    ]);

    var lastLogRow = logSheet.getLastRow();
    PropertiesService.getScriptProperties().setProperty(
      "lastOT_" + chatId,
      JSON.stringify({ name: name, shift: info.shift, logRow: lastLogRow, time: new Date().getTime() })
    );

    clearCache();
    updateOTStatus();

    var msg = "✅ OT Assigned!\n─────────────────────\n";
    msg += "👤 " + name + "\n";
    msg += "⚡ Shift : " + info.shift + " (" + dayLabel + ")\n";
    msg += "📅 Date  : " + otDateStr;

    sendKB(chatId, msg, { inline_keyboard: [[
      { text: "↩️ Undo", callback_data: "undo_ot" },
      { text: "🏠 Dashboard", callback_data: "dashboard" }
    ]]});

  } finally {
    lock.releaseLock();
  }
}

function handleUndoOT(chatId) {
  var raw = PropertiesService.getScriptProperties().getProperty("lastOT_" + chatId);
  if (!raw) {
    sendKB(chatId, "⚠️ Koi recent OT nahi mila undo karne ke liye.", {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
    return;
  }

  var last = JSON.parse(raw);

  // 5 minute window check
  if (new Date().getTime() - last.time > 300000) {
    PropertiesService.getScriptProperties().deleteProperty("lastOT_" + chatId);
    sendKB(chatId, "⏰ Undo window expire ho gayi (5 min).\n\nManually OT_Log se delete karna hoga.", {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
    return;
  }

  try {
    // 1. OT_Log se row delete karo
    var logSheet = SS.getSheetByName(SHEETS.LOG);
    if (last.logRow && last.logRow <= logSheet.getLastRow()) {
      logSheet.deleteRow(last.logRow);
    }

    // 2. Today_Shift se OT row remove karo (naam + "(OT)" tag wala)
    var tSheet = SS.getSheetByName(SHEETS.TODAY);
    var tData  = tSheet.getDataRange().getValues();
    for (var i = tData.length - 1; i >= 1; i--) {
      var n = tData[i][0].toString().trim();
      var cleanN = n.replace(/ \(OT\)/g, "").trim();
      if (cleanN === last.name) {
        if (n.indexOf("(OT)") > -1) {
          // OT tag wali row — naam se (OT) hatao, shift wahi rakho
          tSheet.getRange(i+1, 1).setValue(last.name);
        } else {
          // Agar extra row thi (assignOT ne appendRow ki thi) — delete karo
          // Dhundho duplicate row
        }
        break;
      }
    }

    // assignOT ne appendRow kiya tha — last row delete karo agar wahi naam hai
    // (addOTTag ne existing row update ki thi, assignOT ne extra append ki — undo mein extra row hatao)
    tData = tSheet.getDataRange().getValues();
    for (var i = tData.length - 1; i >= 1; i--) {
      if (tData[i][0].toString().trim() === last.name &&
          tData[i][1].toString().trim() === last.shift) {
        // Check if there's another row with same name (OT tag wali)
        var count = 0;
        for (var j = 1; j < tData.length; j++) {
          if (tData[j][0].toString().replace(/ \(OT\)/g,"").trim() === last.name) count++;
        }
        if (count > 1) {
          tSheet.deleteRow(i + 1);
          break;
        }
      }
    }

    // 3. (OT) tag hatao Today_Shift se
    tData = tSheet.getDataRange().getValues();
    for (var i = 1; i < tData.length; i++) {
      var n = tData[i][0].toString();
      if (n.replace(/ \(OT\)/g,"").trim() === last.name && n.indexOf("(OT)") > -1) {
        tSheet.getRange(i+1, 1).setValue(last.name); // tag hatao
        break;
      }
    }

    // 4. Queue mein wapas pehle position pe laao
    var seqSheet = SS.getSheetByName(SHEETS.SEQ);
    var seqData  = seqSheet.getDataRange().getValues();
    // Last row mein hoga (rotate hua tha) — pehle position pe laao
    for (var i = seqData.length - 1; i >= 1; i--) {
      if (seqData[i][1] === last.name) {
        var row = seqData.splice(i, 1)[0];
        seqData.splice(1, 0, row); // index 1 = first employee
        seqSheet.getRange(2, 1, seqData.length-1, seqData[0].length).setValues(seqData.slice(1));
        markNextOT();
        break;
      }
    }

    // 5. Undo record delete karo
    PropertiesService.getScriptProperties().deleteProperty("lastOT_" + chatId);

    clearCache();
    updateOTStatus();

    var msg = "↩️ Undo Ho Gaya!\n─────────────────────\n";
    msg += "👤 " + last.name + " ka OT cancel kiya\n";
    msg += "⚡ " + last.shift + "\n";
    msg += "─────────────────────\n";
    msg += last.name + " wapas queue mein pehle position pe aa gaya.";

    sendKB(chatId, msg, { inline_keyboard: [[
      { text: "🏠 Dashboard", callback_data: "dashboard" }
    ]]});

  } catch(err) {
    Logger.log("handleUndoOT: " + err);
    sendKB(chatId, "⚠️ Undo mein error: " + err, {
      inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "dashboard" }]]
    });
  }
}

function handleRefuse(chatId) {
  var name = getNextCandidate();
  if (name) {
    SS.getSheetByName(SHEETS.LOG).appendRow([new Date(), name, "", "Refuse", ""]);
    rotateQueue(name);
    clearCache();
  }
  sendKB(chatId, "❌ Refused.\n" + (name ? "👤 " + name + " — queue mein neeche bheja." : ""), {
    inline_keyboard: [[{ text: "🔙 Dashboard", callback_data: "dashboard" }]]
  });
}

function handleSkip(chatId) {
  var name = getNextCandidate();
  if (!name) { sendTelegramWithButtons(chatId, buildOTMessage()); return; }
  moveDownByRequirement(name, 1);
  clearCache();
  sendKB(chatId, "⏭ " + name + " skip kiya gaya.", {
    inline_keyboard: [[{ text: "🔙 Dashboard", callback_data: "dashboard" }]]
  });
}

function processOT() {
  var sheet    = SS.getSheetByName(SHEETS.SEQ);
  var data     = sheet.getDataRange().getValues();
  var shortage = detectShiftShortage();
  var shift    = Object.keys(shortage)[0];
  if (!shift) return;

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === "Accept") {
      assignOT(data[i][1], shift); rotateQueue(data[i][1]);
      SS.getSheetByName(SHEETS.LOG).appendRow([new Date(), data[i][1], shift, "Accept"]);
      clearCache(); return;
    }
    if (data[i][2] === "Refuse") {
      rotateQueue(data[i][1]);
      SS.getSheetByName(SHEETS.LOG).appendRow([new Date(), data[i][1], shift, "Refuse"]);
      clearCache(); return;
    }
  }

  if (data.length < 2) return;
  var first = data[1][1];
  if (isOnLeave(first)) {
    rotateQueue(first);
    SS.getSheetByName(SHEETS.LOG).appendRow([new Date(), first, shift, "Lapse (Leave)"]);
    clearCache(); return;
  }
  if (getEmployeeShift(first) === shift) {
    moveDownByRequirement(first, 1); // sirf 1 step neeche — shortage count se nahi
    SS.getSheetByName(SHEETS.LOG).appendRow([new Date(), first, shift, "Skip (Same Shift)"]);
    clearCache();
  }
}

function assignOT(name, shift) {
  SS.getSheetByName(SHEETS.TODAY).appendRow([name, shift]);
  clearCache();
}

function addOTTag(name, shift) {
  var sheet = SS.getSheetByName(SHEETS.TODAY);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var clean = data[i][0].toString().replace(/ \(OT\)| \(Mutual\)/g, "").trim();
    if (clean === name) {
      // Batch write — 1 call
      sheet.getRange(i+1, 1, 1, 2).setValues([[name + " (OT)", shift]]);
      clearCache();
      return;
    }
  }
  // Nahi mila toh naya row
  sheet.appendRow([name + " (OT)", shift]);
  clearCache();
}

function rotateQueue(name) {
  var sheet = SS.getSheetByName(SHEETS.SEQ);
  var data  = sheet.getDataRange().getValues();
  var idx   = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == name) { idx = i; break; }
  }
  if (idx === -1) return;

  var row = data.splice(idx, 1)[0];
  row[2] = ""; row[3] = "";
  var mySeq = parseInt(row[0]) || 0;

  // Queue structure after splice:
  // [active group: seq ascending] [done group: seq mixed/descending]
  //
  // Active group end dhundo:
  var activeEnd = 0; // data index (0 = header, so 0 means no active)
  for (var i = 1; i < data.length; i++) {
    if (i === 1) { activeEnd = 1; continue; }
    var prevSeq = parseInt(data[i-1][0]) || 0;
    var currSeq = parseInt(data[i][0])   || 0;
    if (currSeq > prevSeq) {
      activeEnd = i;
    } else {
      break;
    }
  }
  // Done group: from (activeEnd+1) to end

  // Done group mein mySeq se bade (seq>) wale ke PEHLE insert karo
  // Kyunki done group mein jo pehle rotate hua woh bada seq tha
  // Aur jo baad mein rotate hua woh chota seq tha
  // (cycle: 1,2,3,4,5 → 2 pehle rotate hua, 1 baad mein)
  // Toh done group sorted descending hota hai rotate order se
  // Insert: pehle aisa done-member dhundo jiska seq < mySeq
  //         usse PEHLE insert karo
  //         Warna bilkul end mein

  var insertAt = data.length; // default end
  for (var i = activeEnd + 1; i < data.length; i++) {
    var theirSeq = parseInt(data[i][0]) || 0;
    if (theirSeq > mySeq) {
      insertAt = i; // unka seq mujhse bada → woh baad mein aayega, main pehle
      break;
    }
  }

  data.splice(insertAt, 0, row);
  sheet.getRange(2, 1, data.length-1, data[0].length).setValues(data.slice(1));
  markNextOT();
  clearCache();
}

function moveDownByRequirement(name, steps) {
  var sheet = SS.getSheetByName(SHEETS.SEQ);
  var data  = sheet.getDataRange().getValues();
  var idx   = -1;
  for (var i = 1; i < data.length; i++) if (data[i][1] == name) { idx = i; break; }
  if (idx === -1) return;
  var ni = Math.min(idx + steps, data.length - 1);
  var row = data.splice(idx, 1)[0]; data.splice(ni, 0, row);
  sheet.getRange(2, 1, data.length-1, data[0].length).setValues(data.slice(1));
  markNextOT();
  clearCache();
}

function markNextOT() {
  var sheet = SS.getSheetByName(SHEETS.SEQ);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var blanks = []; for (var i = 1; i < data.length; i++) blanks.push([""]);
  sheet.getRange(2, 4, data.length-1, 1).setValues(blanks);
  sheet.getRange(2, 4).setValue("NEXT");
}


// ============================================================
// SECTION 14: OT STATUS UPDATE
// ============================================================

function updateOTStatus() {
  var sheet  = SS.getSheetByName(SHEETS.STATUS);
  // Duty_Auto direct padho — cache mein nahi
  var duty   = SS.getSheetByName(SHEETS.DUTY).getDataRange().getValues();
  var leave  = readSheet(SHEETS.LEAVE);
  var data   = sheet.getDataRange().getValues();

  // ── TODAY Present count — Today_Shift se directly ──
  var todaySheet = SS.getSheetByName(SHEETS.TODAY);
  var todayData  = todaySheet.getDataRange().getValues();
  var preToday   = { "1st":0, "2nd":0, "3rd":0 };
  for (var i = 1; i < todayData.length; i++) {
    var s = todayData[i][1];
    if (s && s !== "REST" && preToday[s] !== undefined) preToday[s]++;
  }

  // ── TODAY Leave adjustment — Leave wale minus karo ──
  var todayMid = new Date(); todayMid.setHours(0,0,0,0);
  for (var i = 1; i < leave.length; i++) {
    var emp = leave[i][2];
    var fr  = toDate(leave[i][4]);
    var to  = toDate(leave[i][5]);
    if (!emp || !fr || !to) continue;
    if (todayMid >= fr && todayMid <= to) {
      // Today_Shift mein us employee ki shift dhundo
      for (var j = 1; j < todayData.length; j++) {
        var clean = todayData[j][0].toString().replace(/ \(OT\)| \(Mutual\)/g, "").trim();
        if (clean === emp) {
          var s = todayData[j][1];
          if (s && s !== "REST" && preToday[s] !== undefined) {
            preToday[s] = Math.max(preToday[s] - 1, 0);
          }
          break;
        }
      }
    }
  }

  // ── TOMORROW Present count from Duty_Auto ──
  var tmr    = new Date(); tmr.setDate(tmr.getDate()+1); tmr.setHours(0,0,0,0);
  var tmrStr = Utilities.formatDate(tmr, "GMT+5:30", "dd-MM-yyyy");
  var preTmr = { "1st":0, "2nd":0, "3rd":0 };

  for (var i = 1; i < duty.length; i++) {
    var dt = Utilities.formatDate(new Date(duty[i][0]), "GMT+5:30", "dd-MM-yyyy");
    if (dt !== tmrStr) continue;
    var s = duty[i][2];
    if (s && s !== "REST" && preTmr[s] !== undefined) preTmr[s]++;
  }

  // ── TOMORROW Leave adjustment ──
  for (var i = 1; i < leave.length; i++) {
    var emp = leave[i][2];
    var fr  = toDate(leave[i][4]);
    var to  = toDate(leave[i][5]);
    if (!emp || !fr || !to) continue;
    if (tmr >= fr && tmr <= to) {
      for (var d = 1; d < duty.length; d++) {
        var dt = Utilities.formatDate(new Date(duty[d][0]), "GMT+5:30", "dd-MM-yyyy");
        if (dt === tmrStr && duty[d][1] === emp) {
          var s = duty[d][2];
          if (s && s !== "REST" && preTmr[s] !== undefined) preTmr[s] = Math.max(preTmr[s]-1, 0);
          break;
        }
      }
    }
  }

  // ── TOMORROW OT add — OT_Log Column E = OT_Date ──
  var log = SS.getSheetByName(SHEETS.LOG).getDataRange().getValues();
  for (var i = 1; i < log.length; i++) {
    var otDate = toDate(log[i][4]); // Column E = OT_Date
    var sh     = log[i][2];
    var act    = log[i][3];
    if (!otDate || otDate.getTime() !== tmr.getTime()) continue;
    if (act !== "OT Assigned") continue;
    if (sh && preTmr[sh] !== undefined) preTmr[sh]++;
  }

  // ── Batch write ──
  var updToday = [], updTmr = [];
  for (var k = 1; k < data.length; k++) {
    var sh     = data[k][0];
    var reqT   = data[k][1] || 0; // B
    var reqTmr = data[k][4] || 0; // E
    var pT     = Math.max(preToday[sh] || 0, 0);
    var pTm    = Math.max(preTmr[sh]   || 0, 0);
    updToday.push([pT,  Math.max(reqT   - pT,  0)]);
    updTmr.push(  [pTm, Math.max(reqTmr - pTm, 0)]);
  }

  if (updToday.length) sheet.getRange(2, 3, updToday.length, 2).setValues(updToday);
  if (updTmr.length)   sheet.getRange(2, 6, updTmr.length,   2).setValues(updTmr);
  clearCache();
}

function updateOTWaitingList() {
  var info = getShiftShortage();
  if (!info) return;

  var sheet  = SS.getSheetByName(SHEETS.SEQ);
  var seq    = sheet.getDataRange().getValues();
  var valid  = [];
  var others = [];

  for (var i = 1; i < seq.length; i++) {
    var name     = seq[i][1];
    var empShift = getEmployeeShift(name);

    // Leave pe hai — neeche rakho
    if (isOnLeave(name)) {
      others.push(seq[i]);
      continue;
    }

    // Same shift pe hai — AUTO SKIP, neeche rakho
    if (empShift === info.shift) {
      others.push(seq[i]);
      continue;
    }

    valid.push(seq[i]);
  }

  var q = valid.concat(others);
  if (!q.length) return;
  sheet.getRange(2, 1, q.length, q[0].length).setValues(q);
  markNextOT();
  clearCache();
}

function notifyIfShortage() {
  updateOTStatus();
  var info = getShiftShortage(); if (!info) return;
  updateOTWaitingList();
  sendTelegramWithButtons(GROUP_CHAT_ID, buildOTMessage());
}


// ============================================================
// SECTION 15: SCHEDULED ALERTS
// ============================================================

function scheduledOTAlert() {
  applyFutureMutuals();

  // Today shift update
  updateTodayShift();
  clearCache();

  // Status update — today + tomorrow dono
  updateOTStatus();
  updateOTWaitingList();

  var info = getShiftShortage();
  if (!info) { Logger.log("No shortage."); return; }
  sendTelegramWithButtons(GROUP_CHAT_ID, buildOTMessage());
}

function sendMorningOT() { scheduledOTAlert(); } // 7am
function sendEveningOT()  { scheduledOTAlert(); } // 3pm
function sendNightOT()    { scheduledOTAlert(); } // 11pm

// Midnight pe Today_Shift refresh karo
function midnightUpdate() {
  updateTodayShift();
  clearCache();
  updateOTStatus();
  Logger.log("Midnight update done.");
}


// ============================================================
// SECTION 16: onChange
// ============================================================

function onChange(e) {
  try {
    var sh = e.source.getActiveSheet().getName();
    if (sh === SHEETS.LEAVE || sh === SHEETS.TODAY || sh === SHEETS.SEQ) {
      clearCache();
      updateOTStatus();
      if (sh === SHEETS.LEAVE) notifyIfShortage();
    }
  } catch(err) { Logger.log("onChange: " + err); }
}


// ============================================================
// SECTION 17: TODAY SHIFT UPDATE
// ============================================================

function updateTodayShift() {
  var duty   = SS.getSheetByName(SHEETS.DUTY);
  var tSheet = SS.getSheetByName(SHEETS.TODAY);
  var data   = duty.getDataRange().getValues();
  var tds    = Utilities.formatDate(new Date(), "GMT+5:30", "dd-MM-yyyy");
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (Utilities.formatDate(new Date(data[i][0]), "GMT+5:30", "dd-MM-yyyy") === tds) {
      result.push([data[i][1], data[i][2]]);
    }
  }

  var order = { "1st":1,"2nd":2,"3rd":3,"REST":4 };
  result.sort(function(a, b) { return order[a[1]] - order[b[1]]; });

  tSheet.clear();
  tSheet.getRange(1, 1, 1, 2).setValues([["Employee","Shift"]]);
  if (result.length) {
    tSheet.getRange(2, 1, result.length, 2).setValues(result);
    var colors = result.map(function(r) {
      return r[1]==="1st"?["#1a73e8"]:r[1]==="2nd"?["#188038"]:r[1]==="3rd"?["#e37400"]:r[1]==="REST"?["#d93025"]:["#000000"];
    });
    tSheet.getRange(2, 2, colors.length, 1).setFontColors(colors);
  }
  clearCache();
}

function autoUpdateTodayShift() {
  if (new Date().getHours() === 23) updateTodayShift();
}


// ============================================================
// SECTION 18: DUTY CHART
// ============================================================

function generateDutyChart() {
  var duty = SS.getSheetByName(SHEETS.DUTY);
  var emp  = SS.getSheetByName(SHEETS.BASE).getDataRange().getValues();
  var start = new Date("2026-03-09"), sm = {1:3,3:2,2:1};
  var cs = {}, rd = {};
  for (var i = 1; i < emp.length; i++) { cs[emp[i][0]] = emp[i][1]; rd[emp[i][0]] = emp[i][2]; }

  var res = [["Date","Employee","Shift"]];
  for (var d = 0; d < 365; d++) {
    var cur = new Date(start); cur.setDate(start.getDate() + d);
    var wd  = cur.toLocaleString("en-US", { weekday:"long" });
    for (var i = 1; i < emp.length; i++) {
      var n = emp[i][0], rday = rd[n];
      var sh = wd === rday ? "REST" : ["","1st","2nd","3rd"][cs[n]];
      var nxt = new Date(cur); nxt.setDate(cur.getDate() + 1);
      if (nxt.toLocaleString("en-US", { weekday:"long" }) === rday) cs[n] = sm[cs[n]];
      res.push([Utilities.formatDate(cur, "GMT+5:30", "dd-MM-yyyy"), n, sh]);
    }
  }
  duty.clear();
  duty.getRange(1, 1, res.length, 3).setValues(res);
}

function autoGenerateNextYearChart() {
  var duty = SS.getSheetByName(SHEETS.DUTY);
  var last = duty.getLastRow();
  if (last < 2) { generateDutyChart(); return; }
  var diff = Math.floor((new Date(duty.getRange(last,1).getValue()) - new Date()) / 86400000);
  if (diff <= 30) generateDutyChart();
}


// ============================================================
// SECTION 19: APPSHEET MUTUAL (legacy)
// ============================================================

function applyMutualDuty() {
  var mSheet = SS.getSheetByName(SHEETS.MUTUAL);
  var tSheet = SS.getSheetByName(SHEETS.TODAY);
  var mData  = mSheet.getDataRange().getValues();
  var tData  = tSheet.getDataRange().getValues();
  var tds    = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy");

  for (var i = 1; i < mData.length; i++) {
    var e1 = mData[i][1], e2 = mData[i][2];
    var dt = Utilities.formatDate(new Date(mData[i][5]), "GMT+5:30", "dd/MM/yyyy");
    var st = mData[i][6];
    if (dt !== tds || st !== "Approved") continue;

    var r1=-1, r2=-1;
    for (var j=1;j<tData.length;j++) {
      if (tData[j][0]==e1) r1=j;
      if (tData[j][0]==e2) r2=j;
    }
    if (r1===-1||r2===-1) continue;
    var s1=tData[r1][1], s2=tData[r2][1];
    if (isOnLeave(e1)||isOnLeave(e2)) { mSheet.getRange(i+1,7).setValue("Blocked (Leave)"); continue; }
    if (s1===s2) { mSheet.getRange(i+1,7).setValue("Blocked (Same Shift)"); continue; }
    if (!s1||!s2) { mSheet.getRange(i+1,7).setValue("Blocked"); continue; }
    tSheet.getRange(r1+1,2).setValue(s2);
    tSheet.getRange(r2+1,2).setValue(s1);
    mSheet.getRange(i+1,7).setValue("Applied");
  }
  clearCache();
  updateOTStatus();
  notifyIfShortage();
}


// ============================================================
// SECTION 21: WEB APP HELPER FUNCTIONS
// ============================================================

function getOTStatusJson() {
  try {
    var info = getShiftShortage();
    var msg  = buildOTMessage();
    return { ok: true, shortage: !!info, shift: info ? info.shift : null, isTmr: info ? info.isTmr : null, message: msg };
  } catch(err) { return { ok: false, message: "Error: " + err }; }
}

function getNextCandidateJson() {
  try {
    var info = getShiftShortage();
    if (!info) return { shortage: false };
    var name = getNextCandidate();
    if (!name) return { shortage: false };
    var empShift = getEmployeeShift(name) || "—";
    var otDate   = new Date();
    if (info.isTmr) otDate.setDate(otDate.getDate() + 1);
    return {
      shortage:   true,
      name:       name,
      todayShift: empShift,
      otShift:    info.shift,
      otDay:      info.isTmr ? "Kal" : "Aaj",
      otDate:     Utilities.formatDate(otDate, "GMT+5:30", "dd-MM-yyyy")
    };
  } catch(err) { return { shortage: false, error: err.toString() }; }
}

function getEmployeeListWithUec() {
  try {
    var emp  = readSheet(SHEETS.EMP);
    var list = [];
    for (var i = 1; i < emp.length; i++) {
      var name = (emp[i][3] || "").toString().trim();
      var uec  = (emp[i][1] || "").toString().trim();
      if (name) list.push({ name: name, uec: uec });
    }
    list.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return list;
  } catch(err) { return []; }
}

function getTodayShiftJson() {
  try {
    var data   = readSheet(SHEETS.TODAY);
    var shifts = { "1st": [], "2nd": [], "3rd": [], "REST": [] };
    for (var i = 1; i < data.length; i++) {
      var name  = (data[i][0] || "").toString().trim();
      var shift = (data[i][1] || "").toString().trim();
      if (!name) continue;
      var key = shifts[shift] !== undefined ? shift : "REST";
      shifts[key].push(name);
    }
    return { ok: true, shifts: shifts };
  } catch(err) { return { ok: false, shifts: {} }; }
}

function getTomorrowShiftJson() {
  try {
    return { ok: true, text: buildTomorrowDuty() };
  } catch(err) { return { ok: false, text: "Error: " + err }; }
}

function getOTHistoryJson() {
  try {
    var log  = readSheet(SHEETS.LOG);
    var td   = today0();
    var f7   = new Date(td); f7.setDate(td.getDate() - 7);
    var t7   = new Date(td); t7.setDate(td.getDate() + 7);
    var list = [];
    for (var i = 1; i < log.length; i++) {
      var dt = toDate(log[i][0]);
      if (!dt || dt < f7 || dt > t7) continue;
      list.push({
        date:   Utilities.formatDate(dt, "GMT+5:30", "dd-MM-yyyy"),
        name:   (log[i][1] || "").toString(),
        shift:  (log[i][2] || "").toString(),
        action: (log[i][3] || "").toString(),
        ts:     dt.getTime()
      });
    }
    list.sort(function(a, b) { return b.ts - a.ts; });
    if (list.length > 30) list = list.slice(0, 30);
    return { ok: true, records: list };
  } catch(err) { return { ok: false, records: [] }; }
}

function submitLeaveWebV2(nameParam, uecParam, from, to, reason) {
  try {
    if (!nameParam || !from || !to) return { ok: false, msg: "Fields missing" };
    var fr = new Date(from); fr.setHours(0,0,0,0);
    var tr = new Date(to);   tr.setHours(0,0,0,0);
    if (fr > tr) return { ok: false, msg: "From date, To date se badi hai" };
    var emp = readSheet(SHEETS.EMP);
    var empUec = uecParam || "";
    for (var i = 1; i < emp.length; i++) {
      if ((emp[i][3] || "").toString().trim() === nameParam) {
        empUec = (emp[i][1] || "").toString().trim();
        break;
      }
    }
    SS.getSheetByName(SHEETS.LEAVE).appendRow([
      Utilities.getUuid(), empUec, nameParam, nameParam, fr, tr, reason || ""
    ]);
    clearCache();
    updateOTStatus();
    notifyIfShortage();
    return { ok: true, msg: "Leave submit ho gaya!" };
  } catch(err) { return { ok: false, msg: "Error: " + err.toString() }; }
}

function submitMutualWebV2(emp1Name, emp2Name, date) {
  try {
    if (!emp1Name || !emp2Name || !date) return { ok: false, msg: "Fields missing" };
    if (emp1Name === emp2Name) return { ok: false, msg: "Same employee select nahi kar sakte" };
    var mDate = new Date(date); mDate.setHours(0,0,0,0);
    var td    = today0();
    if (mDate < td) return { ok: false, msg: "Past date nahi ho sakta" };
    var emp = readSheet(SHEETS.EMP);
    var uec1 = "", uec2 = "";
    for (var i = 1; i < emp.length; i++) {
      var n = (emp[i][3] || "").toString().trim();
      if (n === emp1Name) uec1 = (emp[i][1] || "").toString().trim();
      if (n === emp2Name) uec2 = (emp[i][1] || "").toString().trim();
    }
    SS.getSheetByName(SHEETS.MUTUAL).appendRow([
      Utilities.getUuid(), emp1Name, emp2Name, uec1, uec2, mDate, "Approved"
    ]);
    if (mDate.getTime() === td.getTime()) swapTodayShift(emp1Name, emp2Name);
    clearCache();
    updateOTStatus();
    notifyIfShortage();
    return { ok: true, msg: "Mutual submit ho gaya!" };
  } catch(err) { return { ok: false, msg: "Error: " + err.toString() }; }
}

// PLACEHOLDER — doGet neeche Section 22 mein hai

function buildOTHistoryWeb() {
  try {
    var log   = readSheet(SHEETS.LOG);
    var td    = today0();
    var f7    = new Date(td); f7.setDate(td.getDate()-7);
    var t7    = new Date(td); t7.setDate(td.getDate()+7);
    var list  = [];

    for (var i = 1; i < log.length; i++) {
      var dt = toDate(log[i][0]);
      if (!dt || dt < f7 || dt > t7) continue;
      list.push({
        dt:  dt,
        ds:  Utilities.formatDate(dt, "GMT+5:30", "dd-MM-yyyy"),
        n:   log[i][1] || "",
        s:   log[i][2] || "",
        a:   log[i][3] || ""
      });
    }

    list.sort(function(a,b){ return b.dt - a.dt; });
    if (list.length > 15) list = list.slice(0, 15);

    var tds  = Utilities.formatDate(td, "GMT+5:30", "dd-MM-yyyy");
    var msg  = "📊 OT History (±7 Days)\n─────────────────────\n";
    var last = "";

    if (!list.length) { msg += "\n📭 Koi record nahi."; return msg; }

    list.forEach(function(l) {
      if (l.ds !== last) {
        var tag = l.ds===tds ? " ← Aaj" : l.dt<td ? " (Past)" : " (Upcoming)";
        msg += "\n📅 " + l.ds + tag + "\n";
        last = l.ds;
      }
      var e = l.a==="OT Assigned"||l.a==="Accept"      ? "✅" :
              l.a==="Refuse"                            ? "❌" :
              l.a==="Skip (Same Shift)"                 ? "⏭" :
              l.a==="Lapse (Leave)"                     ? "🏥" : "📌";
      msg += e + " " + l.n + (l.s ? " — "+l.s : "") + "\n";
    });
    msg += "\n─────────────────────\n📝 " + list.length + " records";
    return msg;
  } catch(err) { return "⚠️ Error: " + err; }
}

function getEmployeeListWeb() {
  try {
    var emp  = readSheet(SHEETS.EMP);
    var list = [];
    for (var i = 1; i < emp.length; i++) {
      if (emp[i][3]) list.push(emp[i][3].toString().trim());
    }
    list.sort();
    return list;
  } catch(err) { return []; }
}

function text(msg) {
  return ContentService.createTextOutput(msg);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setWebhook() {
  UrlFetchApp.fetch(BOT_URL + "/deleteWebhook?drop_pending_updates=true");
  Utilities.sleep(500);
  var url = "YAHAN_NAYA_URL_DAALO"; // ← Deploy ke baad naya URL yahan daalo
  Logger.log(UrlFetchApp.fetch(BOT_URL + "/setWebhook?url=" + url).getContentText());
}

function deleteWebhook() {
  Logger.log(UrlFetchApp.fetch(BOT_URL + "/deleteWebhook?drop_pending_updates=true").getContentText());
}

function checkWebhook() {
  Logger.log(UrlFetchApp.fetch(BOT_URL + "/getWebhookInfo").getContentText());
}


// ============================================================
// SECTION 22: WEB APP — doGet (JSON only)
// ============================================================

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : null;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === "getOT")      { return json(getOTStatusJson()); }
  if (action === "getNext")    { return json(getNextCandidateJson()); }
  if (action === "empList")    { return json(getEmployeeListWithUec()); }
  if (action === "todayWeb")   { return json(getTodayShiftJson()); }
  if (action === "tomorrowWeb"){ return json(getTomorrowShiftJson()); }
  if (action === "historyWeb") { return json(getOTHistoryJson()); }

  if (action === "accept") {
    var info = getShiftShortage();
    if (!info) return json({ ok: false, msg: "Koi shortage nahi hai." });
    handleAccept(GROUP_CHAT_ID);
    return json({ ok: true, msg: "✅ Accept ho gaya!" });
  }
  if (action === "refuse") {
    handleRefuse(GROUP_CHAT_ID);
    return json({ ok: true, msg: "❌ Refuse ho gaya!" });
  }
  if (action === "skip") {
    handleSkip(GROUP_CHAT_ID);
    return json({ ok: true, msg: "⏭ Skip ho gaya!" });
  }

  if (action === "leaveWeb") {
    return json(submitLeaveWebV2(
      e.parameter.name   || "",
      e.parameter.uec    || "",
      e.parameter.from   || "",
      e.parameter.to     || "",
      e.parameter.reason || ""
    ));
  }

  if (action === "mutualWeb") {
    return json(submitMutualWebV2(
      e.parameter.emp1 || "",
      e.parameter.emp2 || "",
      e.parameter.date || ""
    ));
  }

  return json({ ok: false, msg: "Invalid action" });
}
