// WordWise PH — Full Backend v4.1 (quota alerts + announcements)
// Paste this ENTIRE file into your Apps Script editor

const SHEET_USERS    = "Users";
const SHEET_SCORES   = "Scores";
const SHEET_ACTIVITY = "Activity";
const SHEET_PAYMENTS = "Payments";
const SHEET_CHILDREN = "Children";
const ADMIN_EMAIL    = "darcant01@gmail.com";

// ── QUOTA MONITORING ──
const API_DAILY_SAFE = 3000;  // est. safe daily requests before GAS runtime quota — adjust as you learn your real usage
const TG_TOKEN = "8898030018:AAGbiyas66mmdE-cFI08F_WDkeZCupOIz1g";
const TG_CHAT  = "5559542103";

function countRequest() {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = "req_" + Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    var n = (Number(props.getProperty(key)) || 0) + 1;
    props.setProperty(key, String(n));
    if (n === Math.ceil(API_DAILY_SAFE * 0.75)) alertQuota(n, 75);
    if (n === Math.ceil(API_DAILY_SAFE * 0.95)) alertQuota(n, 95);
    return n;
  } catch(e) { return 0; }
}

function alertQuota(count, pct) {
  var msg = (pct >= 95 ? "🚨" : "⚠️") + " WordWise API usage at " + pct + "% — " + count + " / " + API_DAILY_SAFE +
            " requests today. " + (pct >= 95 ? "Service may slow down or fail soon!" : "Keep an eye on it.");
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TG_TOKEN + "/sendMessage", {
      method: "post",
      payload: { chat_id: TG_CHAT, text: msg }
    });
  } catch(e) {}
  try { MailApp.sendEmail(ADMIN_EMAIL, "WordWise quota alert " + pct + "%", msg); } catch(e) {}
}

function getApiUsageToday() {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = "req_" + Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    return Number(props.getProperty(key)) || 0;
  } catch(e) { return 0; }
}

function getAnnouncement() {
  try {
    var props = PropertiesService.getScriptProperties();
    var t = props.getProperty("ann_text") || "";
    if (!t) return null;
    return { id: props.getProperty("ann_id") || "0", text: t };
  } catch(e) { return null; }
}

function setAnnouncement(p) {
  var rows = getSheet(SHEET_USERS).getDataRange().getValues();
  var ok = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === String(p.username) &&
        String(rows[i][3]) === String(p.password) &&
        String(rows[i][5]) === "admin") { ok = true; break; }
  }
  if (!ok) return { success:false, error:"Not authorized" };
  var props = PropertiesService.getScriptProperties();
  var text = (p.text || "").trim();
  if (text) {
    props.setProperty("ann_text", text);
    props.setProperty("ann_id", String(Date.now()));
  } else {
    props.deleteProperty("ann_text");
    props.deleteProperty("ann_id");
  }
  return { success:true, active: !!text };
}

function cleanOldCounters() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getKeys();
    var cutoff = Utilities.formatDate(new Date(Date.now() - 7*86400000), "GMT+8", "yyyy-MM-dd");
    for (var i = 0; i < all.length; i++) {
      if (all[i].indexOf("req_") === 0 && all[i].slice(4) < cutoff) props.deleteProperty(all[i]);
    }
  } catch(e) {}
}

// ================================================================
// STEP 1: Run this first to create all sheets
// ================================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  function makeSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) { s = ss.insertSheet(name); s.appendRow(headers); s.setFrozenRows(1); }
    return s;
  }
  makeSheet(SHEET_USERS,    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription","SubStart","SubExpiry"]);
  makeSheet(SHEET_SCORES,   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"]);
  makeSheet(SHEET_ACTIVITY, ["ID","Message","Time","Date"]);
  makeSheet(SHEET_PAYMENTS, ["ID","Name","Username","Plan","Price","Reference","Date","Status"]);
  makeSheet(SHEET_CHILDREN, ["ID","ParentUsername","ChildUsername","DateLinked"]);
  addExpiryColumns();
  const users = ss.getSheetByName(SHEET_USERS);
  if (users.getLastRow() < 2) {
    users.appendRow([1,"Admin","admin","wordwise2024","adult","admin",new Date().toLocaleDateString(),0,0,"premium","",""]);
  }
  Logger.log("✅ Setup complete!");
}

// ================================================================
// STEP 2: Run this to add SubStart/SubExpiry columns
// ================================================================
function addExpiryColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) { Logger.log("No Users sheet — run setup() first"); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers.includes("SubStart"))  { sheet.getRange(1, 11).setValue("SubStart");  Logger.log("✅ Added SubStart"); }
  if (!headers.includes("SubExpiry")) { sheet.getRange(1, 12).setValue("SubExpiry"); Logger.log("✅ Added SubExpiry"); }
  Logger.log("✅ Done — SubStart=col11, SubExpiry=col12");
}

// ================================================================
// STEP 3: Run this ONCE to set up the daily auto-expire trigger
// ================================================================
function installDailyTrigger() {
  // Remove old triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkAllExpiries") {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create daily trigger at 1 AM
  ScriptApp.newTrigger("checkAllExpiries")
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  Logger.log("✅ Daily trigger installed! checkAllExpiries runs every day at 1 AM.");
  return "Done!";
}

// ================================================================
// Auto-expire function — runs daily automatically
// ================================================================
function checkAllExpiries() {
  cleanOldCounters();
  const users = getSheet(SHEET_USERS);
  const rows  = users.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  var expired = 0;
  var expiringSoon = [];

  for (var i = 1; i < rows.length; i++) {
    var sub      = rows[i][9]  || "free";
    var subExpiry= rows[i][11] || "";
    var username = rows[i][2];
    var name     = rows[i][1];

    if ((sub === "basic" || sub === "premium") && subExpiry) {
      var expDate  = new Date(subExpiry);
      expDate.setHours(0, 0, 0, 0);
      var daysLeft = Math.ceil((expDate - today) / (1000 * 60 * 60 * 24));

      if (daysLeft <= 0) {
        users.getRange(i + 1, 10).setValue("free");
        logActivity({message: "@" + username + " subscription expired (" + sub + ")"});
        sendAdminEmail(
          "❌ Expired — @" + username,
          "<p><b>" + name + "</b> (@" + username + ") — " + sub + " plan expired on " + subExpiry + ".<br>Account reset to Free.</p>"
        );
        expired++;
      } else if (daysLeft <= 7) {
        expiringSoon.push({name: name, username: username, sub: sub, subExpiry: subExpiry, daysLeft: daysLeft});
      }
    }
  }

  if (expiringSoon.length > 0) {
    var rows_html = expiringSoon.map(function(u) {
      return "<tr><td>" + u.name + "</td><td>@" + u.username + "</td><td>" + u.sub + "</td><td><b>" + u.daysLeft + " days</b></td><td>" + u.subExpiry + "</td></tr>";
    }).join("");
    sendAdminEmail(
      "⚠️ " + expiringSoon.length + " subscription(s) expiring soon",
      "<h3>Expiring within 7 days</h3><table border='1' cellpadding='6'><tr><th>Name</th><th>Username</th><th>Plan</th><th>Days Left</th><th>Expires</th></tr>" + rows_html + "</table>"
    );
  }

  Logger.log("✅ checkAllExpiries: " + expired + " expired, " + expiringSoon.length + " expiring soon");
  return {expired: expired, expiringSoon: expiringSoon.length};
}

// ================================================================
// Remove the trigger if needed
// ================================================================
function removeDailyTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkAllExpiries") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log("Removed " + removed + " trigger(s)");
}

// ================================================================
// Manually set expiry for a specific user — edit USERNAME below
// ================================================================
function manualSetExpiry() {
  var USERNAME  = "maria";   // ← change this
  var PLAN      = "basic";   // ← basic or premium
  var DAYS      = 30;        // ← how many days

  var sheet = getSheet(SHEET_USERS);
  var rows  = sheet.getDataRange().getValues();
  var start  = new Date();
  var expiry = new Date();
  expiry.setDate(expiry.getDate() + DAYS);

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === USERNAME) {
      sheet.getRange(i+1, 10).setValue(PLAN);
      sheet.getRange(i+1, 11).setValue(fmt(start));
      sheet.getRange(i+1, 12).setValue(fmt(expiry));
      Logger.log("✅ Set @" + USERNAME + " → " + PLAN + " expires " + fmt(expiry));
      return;
    }
  }
  Logger.log("❌ User not found: " + USERNAME);
}

// ================================================================
// Test columns — run to verify sheet structure
// ================================================================
function testColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) { Logger.log("No Users sheet!"); return; }
  var data = sheet.getDataRange().getValues();
  Logger.log("Headers: " + JSON.stringify(data[0]));
  for (var i = 1; i < data.length; i++) {
    Logger.log("Row " + i + ": user=" + data[i][2] + " sub=" + data[i][9] + " start=" + data[i][10] + " expiry=" + data[i][11]);
  }
}

// ================================================================
// ROUTER
// ================================================================
function doGet(e) {
  try {
    countRequest();
    var p = e.parameter;
    var action = p.action || "";
    var result;
    switch(action) {
      case "login":              result = login(p); break;
      case "register":           result = register(p); break;
      case "saveScore":          result = saveScore(p); break;
      case "getScores":          result = getScores(p); break;
      case "getProfile":         result = getProfile(p); break;
      case "logActivity":        result = logActivity(p); break;
      case "adminData":          result = getAdminData(); break;
      case "setAnnouncement":    result = setAnnouncement(p); break;
      case "clearActivity":      result = clearActivity(); break;
      case "submitPayment":      result = submitPayment(p); break;
      case "getPendingPayments": result = getPendingPayments(); break;
      case "approvePayment":     result = approvePayment(p); break;
      case "addChild":           result = addChild(p); break;
      case "getChildren":        result = getChildren(p); break;
      case "setExpiry":          result = setExpiry(p); break;
      case "ping":               result = {success:true, message:"WordWise PH v3 running!"}; break;
      case "getAllTimeScores":    result = getAllTimeScores(p); break;
      default:                   result = {success:false, error:"Unknown action: " + action};
    }
    return respond(result);
  } catch(err) {
    return respond({success:false, error:"Server error: " + err.toString()});
  }
}

function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : "{}";
    var data = JSON.parse(raw);
    if (data.action === "submitPayment") {
      countRequest();
      return respond(submitPaymentWithReceipt(data, data.receipt || null));
    }
    return doGet(e);
  } catch(err) {
    return respond({success:false, error:"Server error: " + err.toString()});
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  if (!s) {
    var headers = {
      "Users":    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription","SubStart","SubExpiry"],
      "Scores":   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"],
      "Activity": ["ID","Message","Time","Date"],
      "Payments": ["ID","Name","Username","Plan","Price","Reference","Date","Status"],
      "Children": ["ID","ParentUsername","ChildUsername","DateLinked"]
    };
    s = ss.insertSheet(name);
    if (headers[name]) { s.appendRow(headers[name]); s.setFrozenRows(1); }
  }
  return s;
}

function fmt(d) {
  return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
}
function fmtTime() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a");
}

function login(p) {
  var sheet = getSheet(SHEET_USERS);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() === String(p.username).trim() &&
        String(rows[i][3]).trim() === String(p.password).trim()) {
      var sub      = rows[i][9]  || "free";
      var subStart = rows[i][10] || "";
      var subExpiry= rows[i][11] || "";
      if ((sub === "basic" || sub === "premium") && subExpiry) {
        var expDate = new Date(subExpiry);
        var today   = new Date(); today.setHours(0,0,0,0);
        if (expDate < today) {
          sub = "expired";
          sheet.getRange(i+1, 10).setValue("expired");
          logActivity({message: rows[i][1] + " subscription expired on login"});
        }
      }
      logActivity({message: rows[i][1] + " logged in"});
      return {success:true, user:{id:rows[i][0], name:rows[i][1], username:rows[i][2], type:rows[i][4], role:rows[i][5], joined:rows[i][6], rounds:Number(rows[i][7])||0, best:Number(rows[i][8])||0, subscription:sub, subStart:subStart, subExpiry:subExpiry}};
    }
  }
  return {success:false, error:"Wrong username or password!"};
}

function register(p) {
  if (!p.name || !p.username || !p.password) return {success:false, error:"Missing fields!"};
  if (p.password.length < 4) return {success:false, error:"Password must be at least 4 characters!"};
  var sheet = getSheet(SHEET_USERS);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === String(p.username).toLowerCase())
      return {success:false, error:"Username already taken!"};
  }
  var id     = rows.length;
  var joined = fmt(new Date());
  sheet.appendRow([id, p.name, p.username, p.password, p.type||"kid", "user", joined, 0, 0, "free", "", ""]);
  logActivity({message: p.name + " registered (" + (p.type||"kid") + ")"});
  sendAdminEmail("👤 New User", "<p><b>" + p.name + "</b> (@" + p.username + ") joined as " + p.type + ".</p>");
  return {success:true, user:{id:id, name:p.name, username:p.username, type:p.type||"kid", role:"user", joined:joined, rounds:0, best:0, subscription:"free", subStart:"", subExpiry:""}};
}

function saveScore(p) {
  var scores = getSheet(SHEET_SCORES);
  var users  = getSheet(SHEET_USERS);
  scores.appendRow([scores.getLastRow(), p.name, p.username, p.type, Number(p.score)||0, p.diff, p.stars, fmt(new Date())]);
  if (p.username && p.username !== "guest") {
    var uRows = users.getDataRange().getValues();
    for (var i = 1; i < uRows.length; i++) {
      if (uRows[i][2] === p.username) {
        users.getRange(i+1,8).setValue(Number(uRows[i][7]||0)+1);
        users.getRange(i+1,9).setValue(Math.max(Number(uRows[i][8]||0), Number(p.score)));
        break;
      }
    }
  }
  logActivity({message: (p.name||"Guest") + " scored " + p.score + " on " + p.diff});
  return {success:true};
}

function getScores(p) {
  var rows   = getSheet(SHEET_SCORES).getDataRange().getValues();
  var scores = [];
  for (var i = 1; i < rows.length; i++) {
    if (!p.diff || rows[i][5] === p.diff)
      scores.push({name:rows[i][1], username:rows[i][2], type:rows[i][3], score:Number(rows[i][4])||0, diff:rows[i][5], stars:rows[i][6], date:rows[i][7], total:Number(rows[i][8])||5});
  }
  scores.sort(function(a,b){return b.score-a.score;});
  return {success:true, scores:scores.slice(0,10), ann: getAnnouncement()};
}

function getAllTimeScores(p) {
  var rows = getSheet(SHEET_SCORES).getDataRange().getValues();
  var players = {}; // username -> {name, type, total, rounds, best, games:{grammar,spell,shark,odd}}

  for (var i = 1; i < rows.length; i++) {
    var username = rows[i][2];
    if (!username || username === 'guest') continue;
    var score = Number(rows[i][4]) || 0;
    var diff  = rows[i][5] || '';

    // Determine which game this score belongs to
    var game = 'grammar';
    if (diff.startsWith('spell')) game = 'spell';
    else if (diff.startsWith('shark')) game = 'shark';
    else if (diff.startsWith('odd'))   game = 'odd';

    if (!players[username]) {
      players[username] = {
        name: rows[i][1], username: username, type: rows[i][3],
        total: 0, rounds: 0, best: 0,
        games: {grammar:false, spell:false, shark:false, odd:false}
      };
    }
    players[username].total  += score;
    players[username].rounds += 1;
    players[username].games[game] = true;
    if (score > players[username].best) players[username].best = score;
  }

  // Only include players who have played ALL 4 games
  var result = [];
  for (var u in players) {
    var p2 = players[u];
    var g  = p2.games;
    if (g.grammar && g.spell && g.shark && g.odd) {
      result.push(p2);
    }
  }

  result.sort(function(a,b){ return b.total - a.total; });
  return {success:true, scores:result.slice(0,10)};
}

function getProfile(p) {
  var sRows  = getSheet(SHEET_SCORES).getDataRange().getValues();
  var scores = [];
  for (var i = 1; i < sRows.length; i++) {
    if (sRows[i][2] === p.username)
      scores.push({score:Number(sRows[i][4])||0, diff:sRows[i][5], stars:sRows[i][6], date:sRows[i][7]});
  }
  var uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  var sub="free", subStart="", subExpiry="";
  for (var i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      sub=uRows[i][9]||"free"; subStart=uRows[i][10]||""; subExpiry=uRows[i][11]||"";
      if ((sub==="basic"||sub==="premium") && subExpiry) {
        var expDate=new Date(subExpiry); var today=new Date(); today.setHours(0,0,0,0);
        if (expDate < today) {
          sub="expired";
          getSheet(SHEET_USERS).getRange(i+1,10).setValue("expired");
        }
      }
      break;
    }
  }
  return {success:true, scores:scores.reverse(), subscription:sub, subStart:subStart, subExpiry:subExpiry};
}

function logActivity(p) {
  getSheet(SHEET_ACTIVITY).appendRow([getSheet(SHEET_ACTIVITY).getLastRow(), p.message||"", fmtTime(), fmt(new Date())]);
  return {success:true};
}

function clearActivity() {
  var sheet = getSheet(SHEET_ACTIVITY);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow()-1);
  return {success:true};
}

function submitPayment(p) {
  if (!p.username) return {success:false, error:"Not logged in!"};
  if (!p.reference) return {success:false, error:"Reference number required!"};
  if (!p.plan) return {success:false, error:"No plan selected!"};
  var sheet = getSheet(SHEET_PAYMENTS);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).trim() === String(p.reference).trim())
      return {success:false, error:"This reference number was already submitted!"};
  }
  sheet.appendRow([sheet.getLastRow(), p.name, p.username, p.plan, Number(p.price)||0, p.reference, fmt(new Date()), "pending"]);
  var uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  var uSheet = getSheet(SHEET_USERS);
  for (var i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) { uSheet.getRange(i+1,10).setValue("pending"); break; }
  }
  logActivity({message: p.name + " submitted " + p.plan + " payment (ref: " + p.reference + ")"});
  sendAdminEmail("💳 New Payment!", "<p><b>" + p.name + "</b> (@" + p.username + ") submitted " + p.plan + " payment.<br>GCash Ref: <b style='font-size:18px;color:#BA7517'>" + p.reference + "</b></p><p><a href='https://darcant01.github.io/wordwiseph/' style='background:#D4537E;color:#fff;padding:10px 20px;border-radius:99px;text-decoration:none'>Open Admin Panel</a></p>");
  return {success:true};
}

function submitPaymentWithReceipt(p, receiptBase64) {
  var result = submitPayment(p);
  if (!result.success || !receiptBase64) return result;
  try {
    var matches = receiptBase64.match(/^data:([^;]+);base64,(.+)$/);
    if (matches) {
      var blob = Utilities.newBlob(Utilities.base64Decode(matches[2]), matches[1], "receipt_" + p.username + ".jpg");
      MailApp.sendEmail({to:ADMIN_EMAIL, subject:"WordWise PH | 📸 Receipt @"+p.username, htmlBody:"<p>Receipt for @"+p.username+" — "+p.plan+" — ref: "+p.reference+"</p>", attachments:[blob]});
    }
  } catch(e) { Logger.log("Receipt email error: "+e); }
  return result;
}

function getPendingPayments() {
  var rows = getSheet(SHEET_PAYMENTS).getDataRange().getValues();
  var payments = [];
  for (var i = 1; i < rows.length; i++)
    payments.push({id:Number(rows[i][0]), name:rows[i][1], username:rows[i][2], plan:rows[i][3], price:Number(rows[i][4])||0, reference:rows[i][5], date:rows[i][6], status:rows[i][7]});
  return {success:true, payments:payments.reverse()};
}

function approvePayment(p) {
  var approve  = p.approve === "true" || p.approve === true;
  var payments = getSheet(SHEET_PAYMENTS);
  var users    = getSheet(SHEET_USERS);
  var pRows = payments.getDataRange().getValues();
  for (var i = 1; i < pRows.length; i++) {
    if (String(pRows[i][2]) === String(p.username) && String(pRows[i][5]) === String(p.reference)) {
      payments.getRange(i+1, 8).setValue(approve ? "approved" : "rejected");
      break;
    }
  }
  var uRows = users.getDataRange().getValues();
  for (var i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      if (approve) {
        var startDate = new Date();
        var expDate   = new Date();
        expDate.setDate(expDate.getDate() + 30);
        users.getRange(i+1, 10).setValue(p.plan);
        users.getRange(i+1, 11).setValue(fmt(startDate));
        users.getRange(i+1, 12).setValue(fmt(expDate));
      } else {
        users.getRange(i+1, 10).setValue("free");
        users.getRange(i+1, 11).setValue("");
        users.getRange(i+1, 12).setValue("");
      }
      break;
    }
  }
  logActivity({message:"Admin "+(approve?"approved":"rejected")+" "+p.plan+" for @"+p.username});
  return {success:true};
}

function setExpiry(p) {
  var users = getSheet(SHEET_USERS);
  var rows  = users.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === p.username) {
      if (p.plan)      users.getRange(i+1, 10).setValue(p.plan);
      if (p.subStart)  users.getRange(i+1, 11).setValue(p.subStart);
      if (p.subExpiry) users.getRange(i+1, 12).setValue(p.subExpiry);
      Logger.log("✅ Updated @" + p.username);
      return {success:true};
    }
  }
  return {success:false, error:"User not found"};
}

function addChild(p) {
  var users    = getSheet(SHEET_USERS);
  var children = getSheet(SHEET_CHILDREN);
  var uRows = users.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.childUsername) { found = true; break; }
  }
  if (!found) return {success:false, error:"Username not found!"};
  var cRows = children.getDataRange().getValues();
  for (var i = 1; i < cRows.length; i++) {
    if (cRows[i][1] === p.parentUsername && cRows[i][2] === p.childUsername)
      return {success:false, error:"Already linked!"};
  }
  children.appendRow([children.getLastRow(), p.parentUsername, p.childUsername, fmt(new Date())]);
  logActivity({message: "@"+p.parentUsername+" linked child @"+p.childUsername});
  return {success:true};
}

function getChildren(p) {
  var cRows = getSheet(SHEET_CHILDREN).getDataRange().getValues();
  var uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  var sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  var result = [];
  for (var i = 1; i < cRows.length; i++) {
    if (cRows[i][1] !== p.parentUsername) continue;
    var childUsername = cRows[i][2];
    var childUser = null;
    for (var j = 1; j < uRows.length; j++) {
      if (uRows[j][2] === childUsername) { childUser = {name:uRows[j][1], username:uRows[j][2], type:uRows[j][4]}; break; }
    }
    if (!childUser) continue;
    var childScores = [];
    for (var j = 1; j < sRows.length; j++) {
      if (sRows[j][2] === childUsername)
        childScores.push({score:Number(sRows[j][4])||0, diff:sRows[j][5], stars:sRows[j][6], date:sRows[j][7]});
    }
    result.push({name:childUser.name, username:childUser.username, type:childUser.type, scores:childScores.reverse()});
  }
  return {success:true, children:result};
}

function getAdminData() {
  var uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  var sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  var aRows = getSheet(SHEET_ACTIVITY).getDataRange().getValues();
  var users = [], scores = [], activity = [];
  for (var i = 1; i < uRows.length; i++)
    users.push({id:uRows[i][0], name:uRows[i][1], username:uRows[i][2], type:uRows[i][4], role:uRows[i][5], joined:uRows[i][6], rounds:Number(uRows[i][7])||0, best:Number(uRows[i][8])||0, subscription:uRows[i][9]||"free", subStart:uRows[i][10]||"", subExpiry:uRows[i][11]||""});
  for (var i = 1; i < sRows.length; i++)
    scores.push({name:sRows[i][1], username:sRows[i][2], type:sRows[i][3], score:Number(sRows[i][4])||0, diff:sRows[i][5], stars:sRows[i][6], date:sRows[i][7]});
  for (var i = 1; i < aRows.length; i++)
    activity.push({m:aRows[i][1], t:aRows[i][2], d:aRows[i][3]});
  return {success:true, users:users, scores:scores, activity:activity.reverse().slice(0,50), apiToday: getApiUsageToday(), apiLimit: API_DAILY_SAFE};
}

function sendAdminEmail(subject, body) {
  try {
    MailApp.sendEmail({to:ADMIN_EMAIL, subject:"WordWise PH | "+subject, htmlBody:"<div style='font-family:Arial,sans-serif;max-width:600px'><div style='background:#1D2B55;padding:20px;text-align:center;border-radius:12px 12px 0 0'><h2 style='color:#FAC775;margin:0'>WordWise PH</h2></div><div style='background:#fff;padding:24px;border:1px solid #F4C0D1;border-radius:0 0 12px 12px'>"+body+"</div></div>"});
  } catch(e) { Logger.log("Email error: "+e); }
}
