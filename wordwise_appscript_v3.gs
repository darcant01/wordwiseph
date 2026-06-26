// ============================================================
// WordWise PH — Google Apps Script Backend v3
// ============================================================

const SHEET_USERS    = "Users";
const SHEET_SCORES   = "Scores";
const SHEET_ACTIVITY = "Activity";
const SHEET_PAYMENTS = "Payments";
const SHEET_CHILDREN = "Children";
const ADMIN_EMAIL    = "darcant01@gmail.com";

// ── RUN THIS FIRST: adds SubStart & SubExpiry to Users sheet ──
function addExpiryColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);

  if (!sheet) {
    Logger.log("❌ Users sheet not found. Run setup() first.");
    return;
  }

  const lastCol    = sheet.getLastColumn();
  const headerRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log("Current headers: " + JSON.stringify(headerRow));

  // Add SubStart at column 11 if missing
  if (!headerRow.includes("SubStart")) {
    sheet.getRange(1, 11).setValue("SubStart");
    sheet.getRange(1, 11).setFontWeight("bold").setBackground("#1D2B55").setFontColor("#FAC775");
    Logger.log("✅ Added SubStart at column 11");
  } else {
    Logger.log("ℹ️ SubStart already exists at column " + (headerRow.indexOf("SubStart") + 1));
  }

  // Add SubExpiry at column 12 if missing
  if (!headerRow.includes("SubExpiry")) {
    sheet.getRange(1, 12).setValue("SubExpiry");
    sheet.getRange(1, 12).setFontWeight("bold").setBackground("#1D2B55").setFontColor("#FAC775");
    Logger.log("✅ Added SubExpiry at column 12");
  } else {
    Logger.log("ℹ️ SubExpiry already exists at column " + (headerRow.indexOf("SubExpiry") + 1));
  }

  Logger.log("✅ Done! SubStart = col 11, SubExpiry = col 12");
}

// ── AUTO-EXPIRE: run daily via trigger ──────────────────────────
// Checks ALL users and expires any subscriptions past their subExpiry date
// Set up a daily trigger by running installDailyTrigger() once

function checkAllExpiries() {
  const users  = getSheet(SHEET_USERS);
  const rows   = users.getDataRange().getValues();
  const today  = new Date();
  today.setHours(0, 0, 0, 0);
  
  let expired = 0;
  let notified = 0;
  let expiringSoon = [];

  for (let i = 1; i < rows.length; i++) {
    const sub      = rows[i][9]  || "free";
    const subExpiry= rows[i][11] || "";
    const username = rows[i][2];
    const name     = rows[i][1];

    if ((sub === "basic" || sub === "premium") && subExpiry) {
      const expDate = new Date(subExpiry);
      expDate.setHours(0, 0, 0, 0);
      const daysLeft = Math.ceil((expDate - today) / (1000 * 60 * 60 * 24));

      if (daysLeft <= 0) {
        // ── EXPIRED: reset to free ──
        users.getRange(i+1, 10).setValue("free");
        logActivity({message: "@" + username + " subscription expired (was " + sub + ")"});
        
        // Email admin about expiry
        sendAdminEmail(
          "❌ Subscription Expired — @" + username,
          `<h3 style="color:#1D2B55">Subscription Expired</h3>
           <table style="width:100%;border-collapse:collapse;font-family:Arial">
             <tr><td style="padding:8px;color:#888;font-weight:bold">User</td><td style="padding:8px">${name} (@${username})</td></tr>
             <tr style="background:#f9f9f9"><td style="padding:8px;color:#888;font-weight:bold">Plan</td><td style="padding:8px">${sub}</td></tr>
             <tr><td style="padding:8px;color:#888;font-weight:bold">Expired On</td><td style="padding:8px;color:#E24B4A;font-weight:bold">${subExpiry}</td></tr>
           </table>
           <div style="margin-top:14px;padding:12px;background:#FCEBEB;border-radius:8px;color:#791F1F;font-weight:bold">
             Their account has been automatically set to Free plan.
           </div>
           <div style="margin-top:12px;text-align:center">
             <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:10px 24px;border-radius:99px;text-decoration:none;font-weight:bold">Open Admin Panel</a>
           </div>`
        );
        expired++;

      } else if (daysLeft <= 7) {
        // ── EXPIRING SOON: collect for summary email ──
        expiringSoon.push({name, username, sub, subExpiry, daysLeft});
      }
    }
  }

  // Send expiring-soon summary to admin
  if (expiringSoon.length > 0) {
    const rows_html = expiringSoon.map(u =>
      `<tr>
        <td style="padding:8px">${u.name}</td>
        <td style="padding:8px">@${u.username}</td>
        <td style="padding:8px">${u.sub}</td>
        <td style="padding:8px;color:#BA7517;font-weight:bold">${u.daysLeft} day${u.daysLeft===1?'':'s'}</td>
        <td style="padding:8px">${u.subExpiry}</td>
      </tr>`
    ).join('');

    sendAdminEmail(
      "⚠️ " + expiringSoon.length + " Subscription(s) Expiring Soon",
      `<h3 style="color:#1D2B55">Subscriptions Expiring Within 7 Days</h3>
       <table style="width:100%;border-collapse:collapse;font-family:Arial;border:1px solid #eee">
         <thead><tr style="background:#1D2B55;color:#FAC775">
           <th style="padding:8px;text-align:left">Name</th>
           <th style="padding:8px;text-align:left">Username</th>
           <th style="padding:8px;text-align:left">Plan</th>
           <th style="padding:8px;text-align:left">Days Left</th>
           <th style="padding:8px;text-align:left">Expires</th>
         </tr></thead>
         <tbody>${rows_html}</tbody>
       </table>
       <div style="margin-top:14px;padding:12px;background:#FAEEDA;border-radius:8px;color:#412402;font-weight:bold">
         💡 Consider reaching out to these subscribers to offer renewal!
       </div>
       <div style="margin-top:12px;text-align:center">
         <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:10px 24px;border-radius:99px;text-decoration:none;font-weight:bold">Open Admin Panel</a>
       </div>`
    );
    notified = expiringSoon.length;
  }

  const summary = "checkAllExpiries: " + expired + " expired, " + notified + " expiring-soon notified";
  Logger.log("✅ " + summary);
  logActivity({message: "Daily check: " + expired + " expired, " + notified + " expiring soon"});
  return {success: true, expired, expiringSoon: notified};
}

// ── INSTALL DAILY TRIGGER ──────────────────────────────────────
// Run this ONCE to set up automatic daily expiry checks
function installDailyTrigger() {
  // Remove existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAllExpiries') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Create new daily trigger at 1:00 AM Philippine time
  ScriptApp.newTrigger('checkAllExpiries')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  
  Logger.log("✅ Daily trigger installed — checkAllExpiries will run every day at 1:00 AM");
  return "Daily trigger installed!";
}

// ── REMOVE DAILY TRIGGER ────────────────────────────────────────
function removeDailyTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAllExpiries') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log("Removed " + removed + " trigger(s)");
}

// ── SETUP: create all sheets ──────────────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function makeSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.appendRow(headers);
      s.setFrozenRows(1);
      s.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#1D2B55").setFontColor("#FAC775");
    }
    return s;
  }

  makeSheet(SHEET_USERS,    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription","SubStart","SubExpiry"]);
  makeSheet(SHEET_SCORES,   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"]);
  makeSheet(SHEET_ACTIVITY, ["ID","Message","Time","Date"]);
  makeSheet(SHEET_PAYMENTS, ["ID","Name","Username","Plan","Price","Reference","Date","Status"]);
  makeSheet(SHEET_CHILDREN, ["ID","ParentUsername","ChildUsername","DateLinked"]);

  // Always ensure SubStart/SubExpiry exist
  addExpiryColumns();

  // Default admin
  const users = ss.getSheetByName(SHEET_USERS);
  if (users.getLastRow() < 2) {
    users.appendRow([1,"Admin","admin","wordwise2024","adult","admin",new Date().toLocaleDateString(),0,0,"premium","",""]);
  }

  return "✅ Setup complete!";
}

// ── TEST: verify columns and show all user data ───────────────
function testColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) { Logger.log("❌ No Users sheet"); return; }
  const data = sheet.getDataRange().getValues();
  Logger.log("Total columns: " + sheet.getLastColumn());
  Logger.log("Headers: " + JSON.stringify(data[0]));
  for (let i = 1; i < data.length; i++) {
    Logger.log("Row " + i + ": username=" + data[i][2] + " | sub=" + data[i][9] + " | subStart=" + data[i][10] + " | subExpiry=" + data[i][11]);
  }
}

// ── MANUAL: set expiry for a specific user ────────────────────
function manualSetExpiry() {
  // EDIT THESE VALUES before running:
  const USERNAME  = "maria";      // ← change to actual username
  const PLAN      = "basic";      // ← basic or premium
  const DAYS      = 30;           // ← number of days

  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const rows   = sheet.getDataRange().getValues();
  const fmt    = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
  const start  = new Date();
  const expiry = new Date(); expiry.setDate(expiry.getDate() + DAYS);

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === USERNAME) {
      sheet.getRange(i+1, 10).setValue(PLAN);
      sheet.getRange(i+1, 11).setValue(fmt(start));
      sheet.getRange(i+1, 12).setValue(fmt(expiry));
      Logger.log("✅ Set @" + USERNAME + " → " + PLAN + " | expires: " + fmt(expiry));
      return;
    }
  }
  Logger.log("❌ User not found: " + USERNAME);
}

// ── CORS / RESPOND ────────────────────────────────────────────
function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── ROUTER ───────────────────────────────────────────────────
function doGet(e) {
  try {
    const p      = e.parameter;
    const action = p.action || "";
    let result;
    switch(action) {
      case "login":              result = login(p); break;
      case "register":           result = register(p); break;
      case "saveScore":          result = saveScore(p); break;
      case "getScores":          result = getScores(p); break;
      case "getProfile":         result = getProfile(p); break;
      case "logActivity":        result = logActivity(p); break;
      case "adminData":          result = getAdminData(); break;
      case "clearActivity":      result = clearActivity(); break;
      case "submitPayment":      result = submitPayment(p); break;
      case "getPendingPayments": result = getPendingPayments(); break;
      case "approvePayment":     result = approvePayment(p); break;
      case "addChild":           result = addChild(p); break;
      case "getChildren":        result = getChildren(p); break;
      case "setExpiry":          result = setExpiry(p); break;
      case "ping":               result = {success:true, message:"WordWise PH API v3 running!"}; break;
      default:                   result = {success:false, error:"Unknown action: " + action};
    }
    return respond(result);
  } catch(err) {
    return respond({success:false, error:"Server error: " + err.toString()});
  }
}

function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    if (data.action === "submitPayment") {
      return respond(submitPaymentWithReceipt(data, data.receipt || null));
    }
    return doGet(e);
  } catch(err) {
    return respond({success:false, error:"Server error: " + err.toString()});
  }
}

// ── HELPER ───────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    const headers = {
      [SHEET_USERS]:    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription","SubStart","SubExpiry"],
      [SHEET_SCORES]:   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"],
      [SHEET_ACTIVITY]: ["ID","Message","Time","Date"],
      [SHEET_PAYMENTS]: ["ID","Name","Username","Plan","Price","Reference","Date","Status"],
      [SHEET_CHILDREN]: ["ID","ParentUsername","ChildUsername","DateLinked"]
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

// ── LOGIN ────────────────────────────────────────────────────
function login(p) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() === String(p.username).trim() &&
        String(rows[i][3]).trim() === String(p.password).trim()) {
      logActivity({message: rows[i][1] + " logged in"});
      // Auto-expire check on login
      let sub      = rows[i][9]  || "free";
      const subStart  = rows[i][10] || "";
      const subExpiry = rows[i][11] || "";

      if ((sub === "basic" || sub === "premium") && subExpiry) {
        const expDate = new Date(subExpiry);
        const today   = new Date();
        today.setHours(0, 0, 0, 0);
        if (expDate < today) {
          sub = "expired";
          sheet.getRange(i+1, 10).setValue("expired");
          logActivity({message: rows[i][1] + " subscription expired on login"});
        }
      }

      return {
        success: true,
        user: {
          id:           rows[i][0],
          name:         rows[i][1],
          username:     rows[i][2],
          type:         rows[i][4],
          role:         rows[i][5],
          joined:       rows[i][6],
          rounds:       Number(rows[i][7]) || 0,
          best:         Number(rows[i][8]) || 0,
          subscription: sub,
          subStart,
          subExpiry
        }
      };
    }
  }
  return {success:false, error:"Wrong username or password!"};
}

// ── REGISTER ─────────────────────────────────────────────────
function register(p) {
  if (!p.name || !p.username || !p.password) return {success:false, error:"Missing required fields!"};
  if (p.password.length < 4) return {success:false, error:"Password must be at least 4 characters!"};
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === String(p.username).toLowerCase())
      return {success:false, error:"Username already taken!"};
  }
  const id     = rows.length;
  const joined = fmt(new Date());
  sheet.appendRow([id, p.name, p.username, p.password, p.type||"kid", "user", joined, 0, 0, "free", "", ""]);
  logActivity({message: p.name + " registered (" + (p.type||"kid") + ")"});
  sendAdminEmail("👤 New User", `<p><b>${p.name}</b> (@${p.username}) registered as ${p.type}.</p>`);
  return {success:true, user:{id, name:p.name, username:p.username, type:p.type||"kid", role:"user", joined, rounds:0, best:0, subscription:"free", subStart:"", subExpiry:""}};
}

// ── SAVE SCORE ───────────────────────────────────────────────
function saveScore(p) {
  const scores = getSheet(SHEET_SCORES);
  const users  = getSheet(SHEET_USERS);
  scores.appendRow([scores.getLastRow(), p.name, p.username, p.type, Number(p.score)||0, p.diff, p.stars, fmt(new Date())]);
  if (p.username && p.username !== "guest") {
    const uRows = users.getDataRange().getValues();
    for (let i = 1; i < uRows.length; i++) {
      if (uRows[i][2] === p.username) {
        users.getRange(i+1,8).setValue(Number(uRows[i][7]||0)+1);
        users.getRange(i+1,9).setValue(Math.max(Number(uRows[i][8]||0), Number(p.score)));
        break;
      }
    }
  }
  logActivity({message:(p.name||"Guest")+" scored "+p.score+"/5 on "+p.diff});
  return {success:true};
}

// ── GET SCORES ───────────────────────────────────────────────
function getScores(p) {
  const rows = getSheet(SHEET_SCORES).getDataRange().getValues();
  let scores = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = rows[i];
    if (!p.diff || diff === p.diff)
      scores.push({name,username,type,score:Number(score)||0,diff,stars,date});
  }
  scores.sort((a,b)=>b.score-a.score);
  return {success:true, scores:scores.slice(0,10).map(s=>({...s, total:s.total||5}))};
}

// ── GET PROFILE ──────────────────────────────────────────────
function getProfile(p) {
  // Scores
  const sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  let scores = [];
  for (let i = 1; i < sRows.length; i++) {
    if (sRows[i][2] === p.username)
      scores.push({score:Number(sRows[i][4])||0, diff:sRows[i][5], stars:sRows[i][6], date:sRows[i][7]});
  }

  // User subscription info
  const uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  let sub="free", subStart="", subExpiry="";
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      sub      = uRows[i][9]  || "free";
      subStart = uRows[i][10] || "";
      subExpiry= uRows[i][11] || "";
      break;
    }
  }

  // Auto-expire check
  if ((sub==="basic"||sub==="premium") && subExpiry) {
    const exp = new Date(subExpiry);
    const now = new Date(); now.setHours(0,0,0,0);
    if (exp < now) {
      sub = "expired";
      const users = getSheet(SHEET_USERS);
      const rows2 = users.getDataRange().getValues();
      for (let i = 1; i < rows2.length; i++) {
        if (rows2[i][2] === p.username) { users.getRange(i+1,10).setValue("expired"); break; }
      }
    }
  }

  return {success:true, scores:scores.reverse(), subscription:sub, subStart, subExpiry};
}

// ── LOG ACTIVITY ─────────────────────────────────────────────
function logActivity(p) {
  const sheet = getSheet(SHEET_ACTIVITY);
  sheet.appendRow([sheet.getLastRow(), p.message||"", fmtTime(), fmt(new Date())]);
  return {success:true};
}

function clearActivity() {
  const sheet = getSheet(SHEET_ACTIVITY);
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last-1);
  return {success:true};
}

// ── SUBMIT PAYMENT ───────────────────────────────────────────
function submitPayment(p) {
  if (!p.username) return {success:false, error:"Not logged in!"};
  if (!p.reference) return {success:false, error:"GCash reference number required!"};
  if (!p.plan) return {success:false, error:"No plan selected!"};
  const sheet = getSheet(SHEET_PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).trim() === String(p.reference).trim())
      return {success:false, error:"This reference number was already submitted!"};
  }
  sheet.appendRow([sheet.getLastRow(), p.name, p.username, p.plan, Number(p.price)||0, p.reference, fmt(new Date()), "pending"]);
  // Set user to pending
  const users = getSheet(SHEET_USERS);
  const uRows = users.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) { users.getRange(i+1,10).setValue("pending"); break; }
  }
  logActivity({message:p.name+" submitted "+p.plan+" payment (ref: "+p.reference+")"});
  sendAdminEmail("💳 New Payment!",
    `<h3>Payment from @${p.username}</h3>
     <p><b>Plan:</b> ${p.plan} — ₱${p.price}/mo<br>
     <b>GCash Ref #:</b> <span style="font-size:20px;color:#BA7517">${p.reference}</span><br>
     <b>Date:</b> ${fmt(new Date())}</p>
     <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:10px 24px;border-radius:99px;text-decoration:none;font-weight:bold">Open Admin Panel</a>`);
  return {success:true};
}

function submitPaymentWithReceipt(p, receiptBase64) {
  const result = submitPayment(p);
  if (!result.success || !receiptBase64) return result;
  try {
    const matches = receiptBase64.match(/^data:([^;]+);base64,(.+)$/);
    if (matches) {
      const blob = Utilities.newBlob(Utilities.base64Decode(matches[2]), matches[1], "receipt_"+p.username+".jpg");
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: "WordWise PH | 📸 Receipt @"+p.username,
        htmlBody: `<p>Receipt attached for <b>@${p.username}</b> — ${p.plan} plan — ref: <b>${p.reference}</b></p>`,
        attachments: [blob]
      });
    }
  } catch(e) { Logger.log("Receipt email error: "+e); }
  return result;
}

// ── GET PENDING PAYMENTS ─────────────────────────────────────
function getPendingPayments() {
  const rows = getSheet(SHEET_PAYMENTS).getDataRange().getValues();
  let payments = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,plan,price,reference,date,status] = rows[i];
    payments.push({id:Number(id),name,username,plan,price:Number(price)||0,reference,date,status});
  }
  return {success:true, payments:payments.reverse()};
}

// ── APPROVE PAYMENT ──────────────────────────────────────────
function approvePayment(p) {
  const approve  = p.approve === "true" || p.approve === true;
  const payments = getSheet(SHEET_PAYMENTS);
  const users    = getSheet(SHEET_USERS);

  // Update payment status
  const pRows = payments.getDataRange().getValues();
  for (let i = 1; i < pRows.length; i++) {
    if (String(pRows[i][2]) === String(p.username) && String(pRows[i][5]) === String(p.reference)) {
      payments.getRange(i+1, 8).setValue(approve ? "approved" : "rejected");
      break;
    }
  }

  // Update user subscription + set SubStart & SubExpiry
  const uRows = users.getDataRange().getValues();
  Logger.log("Looking for user: " + p.username);
  Logger.log("Total user rows: " + uRows.length);
  Logger.log("Sheet last column: " + users.getLastColumn());

  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      Logger.log("Found user at row " + (i+1));
      if (approve) {
        const startDate = new Date();
        const expDate   = new Date();
        expDate.setDate(expDate.getDate() + 30);
        Logger.log("Setting subscription: " + p.plan);
        Logger.log("Setting subStart: " + fmt(startDate));
        Logger.log("Setting subExpiry: " + fmt(expDate));
        users.getRange(i+1, 10).setValue(p.plan);
        users.getRange(i+1, 11).setValue(fmt(startDate));
        users.getRange(i+1, 12).setValue(fmt(expDate));
        Logger.log("✅ All values written to sheet");
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

// ── SET EXPIRY (manual) ──────────────────────────────────────
function setExpiry(p) {
  const users = getSheet(SHEET_USERS);
  const rows  = users.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === p.username) {
      if (p.plan)      users.getRange(i+1, 10).setValue(p.plan);
      if (p.subStart)  users.getRange(i+1, 11).setValue(p.subStart);
      if (p.subExpiry) users.getRange(i+1, 12).setValue(p.subExpiry);
      Logger.log("✅ Updated @" + p.username + ": sub=" + p.plan + " start=" + p.subStart + " expiry=" + p.subExpiry);
      return {success:true, message:"Updated @"+p.username};
    }
  }
  return {success:false, error:"User not found: "+p.username};
}

// ── ADD CHILD ────────────────────────────────────────────────
function addChild(p) {
  const users    = getSheet(SHEET_USERS);
  const children = getSheet(SHEET_CHILDREN);
  const uRows    = users.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.childUsername) { found=true; break; }
  }
  if (!found) return {success:false, error:"Username not found. Make sure your child has registered first!"};
  const cRows = children.getDataRange().getValues();
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][1]===p.parentUsername && cRows[i][2]===p.childUsername)
      return {success:false, error:"This child is already linked!"};
  }
  children.appendRow([children.getLastRow(), p.parentUsername, p.childUsername, fmt(new Date())]);
  logActivity({message:"@"+p.parentUsername+" linked child @"+p.childUsername});
  return {success:true};
}

// ── GET CHILDREN ─────────────────────────────────────────────
function getChildren(p) {
  const cRows = getSheet(SHEET_CHILDREN).getDataRange().getValues();
  const uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  const sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  let result = [];
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][1] !== p.parentUsername) continue;
    const childUser = uRows.find(r=>r[2]===cRows[i][2]);
    if (!childUser) continue;
    const childScores = sRows.filter(r=>r[2]===cRows[i][2]).map(r=>({score:Number(r[4])||0,diff:r[5],stars:r[6],date:r[7]})).reverse();
    result.push({name:childUser[1], username:childUser[2], type:childUser[4], scores:childScores});
  }
  return {success:true, children:result};
}

// ── ADMIN DATA ───────────────────────────────────────────────
function getAdminData() {
  const uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  const sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  const aRows = getSheet(SHEET_ACTIVITY).getDataRange().getValues();
  const users = [], scores = [], activity = [];
  for (let i = 1; i < uRows.length; i++) {
    users.push({id:uRows[i][0],name:uRows[i][1],username:uRows[i][2],type:uRows[i][4],role:uRows[i][5],joined:uRows[i][6],rounds:Number(uRows[i][7])||0,best:Number(uRows[i][8])||0,subscription:uRows[i][9]||"free",subStart:uRows[i][10]||"",subExpiry:uRows[i][11]||""});
  }
  for (let i = 1; i < sRows.length; i++) {
    scores.push({name:sRows[i][1],username:sRows[i][2],type:sRows[i][3],score:Number(sRows[i][4])||0,diff:sRows[i][5],stars:sRows[i][6],date:sRows[i][7]});
  }
  for (let i = 1; i < aRows.length; i++) {
    activity.push({m:aRows[i][1],t:aRows[i][2],d:aRows[i][3]});
  }
  return {success:true, users, scores, activity:activity.reverse().slice(0,50)};
}

// ── EMAIL ────────────────────────────────────────────────────
function sendAdminEmail(subject, body) {
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "WordWise PH | " + subject,
      htmlBody: `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
        <div style="background:#1D2B55;padding:20px;text-align:center;border-radius:12px 12px 0 0">
          <h2 style="color:#FAC775;margin:0">WordWise PH</h2>
        </div>
        <div style="background:#fff;padding:24px;border:1px solid #F4C0D1;border-radius:0 0 12px 12px">
          ${body}
        </div>
      </div>`
    });
  } catch(e) { Logger.log("Email error: "+e); }
}
