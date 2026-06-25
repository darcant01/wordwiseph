// ============================================================
// WordWise PH — Google Apps Script Backend v2
// Includes: Users, Scores, Activity, Payments, Children
// ============================================================

const SHEET_USERS    = "Users";
const SHEET_SCORES   = "Scores";
const SHEET_ACTIVITY = "Activity";
const SHEET_PAYMENTS = "Payments";
const SHEET_CHILDREN = "Children";

// ── SETUP ──────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // USERS: ID | Name | Username | Password | Type | Role | Joined | Rounds | BestScore | Subscription
  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) {
    users = ss.insertSheet(SHEET_USERS);
    users.appendRow(["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription"]);
    users.appendRow([1,"Admin","admin","wordwise2024","adult","admin",new Date().toLocaleDateString(),0,0,"premium"]);
    users.setFrozenRows(1);
  } else if(users.getRange(1,10).getValue() !== "Subscription"){
    // Add subscription column if missing
    users.getRange(1,10).setValue("Subscription");
  }

  // SCORES: ID | Name | Username | Type | Score | Difficulty | Stars | Date
  let scores = ss.getSheetByName(SHEET_SCORES);
  if (!scores) {
    scores = ss.insertSheet(SHEET_SCORES);
    scores.appendRow(["ID","Name","Username","Type","Score","Difficulty","Stars","Date"]);
    scores.setFrozenRows(1);
  }

  // ACTIVITY: ID | Message | Time | Date
  let activity = ss.getSheetByName(SHEET_ACTIVITY);
  if (!activity) {
    activity = ss.insertSheet(SHEET_ACTIVITY);
    activity.appendRow(["ID","Message","Time","Date"]);
    activity.setFrozenRows(1);
  }

  // PAYMENTS: ID | Name | Username | Plan | Price | Reference | Date | Status
  let payments = ss.getSheetByName(SHEET_PAYMENTS);
  if (!payments) {
    payments = ss.insertSheet(SHEET_PAYMENTS);
    payments.appendRow(["ID","Name","Username","Plan","Price","Reference","Date","Status"]);
    payments.setFrozenRows(1);
  }

  // CHILDREN: ID | ParentUsername | ChildUsername | DateLinked
  let children = ss.getSheetByName(SHEET_CHILDREN);
  if (!children) {
    children = ss.insertSheet(SHEET_CHILDREN);
    children.appendRow(["ID","ParentUsername","ChildUsername","DateLinked"]);
    children.setFrozenRows(1);
  }

  return "Setup complete! All 5 sheets created.";
}

// ── EMAIL CONFIG ───────────────────────────────────────────
const ADMIN_EMAIL = "darcant01@gmail.com"; // Your email here

function sendAdminEmail(subject, body) {
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "WordWise PH | " + subject,
      htmlBody: `
        <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
          <div style="background:#1D2B55;padding:20px;border-radius:12px 12px 0 0;text-align:center">
            <h2 style="color:#FAC775;margin:0">WordWise PH</h2>
            <p style="color:#B5D4F4;margin:4px 0 0">Admin Notification</p>
          </div>
          <div style="background:#fff;padding:24px;border:1px solid #F4C0D1;border-top:none;border-radius:0 0 12px 12px">
            ${body}
          </div>
          <p style="text-align:center;color:#888780;font-size:12px;margin-top:12px">WordWise PH · darcant01.github.io/wordwiseph</p>
        </div>`
    });
  } catch(e) {
    Logger.log("Email error: " + e.toString());
  }
}

// ── MAIN ROUTER ────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    switch(action){
      case "login":            result = login(data); break;
      case "register":         result = register(data); break;
      case "saveScore":        result = saveScore(data); break;
      case "getScores":        result = getScores(data); break;
      case "getProfile":       result = getProfile(data); break;
      case "logActivity":      result = logActivity(data); break;
      case "adminData":        result = getAdminData(); break;
      case "clearActivity":    result = clearActivity(); break;
      case "submitPayment":    result = submitPayment(data); break;
      case "getPendingPayments": result = getPendingPayments(); break;
      case "approvePayment":   result = approvePayment(data); break;
      case "addChild":         result = addChild(data); break;
      case "getChildren":      result = getChildren(data); break;
      default: result = {success:false, error:"Unknown action: "+action};
    }
    return respond(result);
  } catch(err) {
    return respond({success:false, error:err.toString()});
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if(action==="getScores") return respond(getScores({diff:e.parameter.diff}));
  if(action==="adminData")  return respond(getAdminData());
  return respond({success:true, message:"WordWise PH API v2 is running!"});
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── AUTH: LOGIN ─────────────────────────────────────────────
function login(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [id, name, username, password, type, role, joined, rounds, best, subscription] = rows[i];
    if (username === data.username && password === data.password) {
      return {
        success: true,
        user: { id, name, username, type, role, joined, rounds:Number(rounds), best:Number(best), subscription:subscription||'free' }
      };
    }
  }
  return { success:false, error:"Wrong username or password!" };
}

// ── AUTH: REGISTER ──────────────────────────────────────────
function register(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.username) return {success:false, error:"Username already taken!"};
  }
  const newId  = rows.length;
  const joined = new Date().toLocaleDateString();
  sheet.appendRow([newId, data.name, data.username, data.password, data.type, "user", joined, 0, 0, "free"]);
  logActivity({message: data.name + " created an account (" + data.type + ")"});
  // Notify admin of new registration
  sendAdminEmail(
    "👤 New User Registered",
    `<h3 style="color:#1D2B55">New Account Created</h3>
    <table style="width:100%;border-collapse:collapse">
      <tr><td style="padding:8px;font-weight:bold;color:#888">Name</td><td style="padding:8px">${data.name}</td></tr>
      <tr style="background:#FFF8F0"><td style="padding:8px;font-weight:bold;color:#888">Username</td><td style="padding:8px">@${data.username}</td></tr>
      <tr><td style="padding:8px;font-weight:bold;color:#888">Type</td><td style="padding:8px">${data.type}</td></tr>
      <tr style="background:#FFF8F0"><td style="padding:8px;font-weight:bold;color:#888">Joined</td><td style="padding:8px">${new Date().toLocaleDateString()}</td></tr>
    </table>`
  );
  return {
    success: true,
    user: {id:newId, name:data.name, username:data.username, type:data.type, role:"user", joined, rounds:0, best:0, subscription:"free"}
  };
}

// ── SCORES: SAVE ────────────────────────────────────────────
function saveScore(data) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const scores = ss.getSheetByName(SHEET_SCORES);
  const users  = ss.getSheetByName(SHEET_USERS);
  const newId  = scores.getLastRow();
  const date   = new Date().toLocaleDateString();
  scores.appendRow([newId, data.name, data.username, data.type, data.score, data.diff, data.stars, date]);
  if (data.username && data.username !== "guest") {
    const userRows = users.getDataRange().getValues();
    for (let i = 1; i < userRows.length; i++) {
      if (userRows[i][2] === data.username) {
        users.getRange(i+1,8).setValue(Number(userRows[i][7])+1);
        users.getRange(i+1,9).setValue(Math.max(Number(userRows[i][8]), data.score));
        break;
      }
    }
  }
  logActivity({message: data.name + " scored " + data.score + "/5 on " + data.diff});
  return {success:true};
}

// ── SCORES: GET LEADERBOARD ─────────────────────────────────
function getScores(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES);
  const rows  = sheet.getDataRange().getValues();
  let scores  = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = rows[i];
    if (!data.diff || diff === data.diff) scores.push({name,username,type,score:Number(score),diff,stars,date});
  }
  scores.sort((a,b)=>b.score-a.score);
  return {success:true, scores:scores.slice(0,10)};
}

// ── PROFILE ─────────────────────────────────────────────────
function getProfile(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES);
  const rows  = sheet.getDataRange().getValues();
  let scores  = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = rows[i];
    if (username === data.username) scores.push({score:Number(score),diff,stars,date});
  }
  return {success:true, scores:scores.reverse()};
}

// ── ACTIVITY ────────────────────────────────────────────────
function logActivity(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ACTIVITY);
  const id    = sheet.getLastRow();
  const time  = new Date().toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"});
  const date  = new Date().toLocaleDateString();
  sheet.appendRow([id, data.message, time, date]);
  return {success:true};
}

function clearActivity() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ACTIVITY);
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last-1);
  return {success:true};
}

// ── PAYMENTS: SUBMIT ────────────────────────────────────────
function submitPayment(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  // Check for duplicate reference number
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][5] === data.reference) return {success:false, error:"This reference number has already been submitted!"};
  }
  const id = sheet.getLastRow();
  sheet.appendRow([id, data.name, data.username, data.plan, data.price, data.reference, data.date, "pending"]);
  logActivity({message: data.name + " submitted " + data.plan + " payment (ref: " + data.reference + ")"});
  // Send email to admin
  sendAdminEmail(
    "💳 New Payment Submitted!",
    `<h3 style="color:#1D2B55">New Payment Request</h3>
    <table style="width:100%;border-collapse:collapse">
      <tr><td style="padding:8px;font-weight:bold;color:#888">Name</td><td style="padding:8px;color:#1D2B55">${data.name}</td></tr>
      <tr style="background:#FFF8F0"><td style="padding:8px;font-weight:bold;color:#888">Username</td><td style="padding:8px;color:#1D2B55">@${data.username}</td></tr>
      <tr><td style="padding:8px;font-weight:bold;color:#888">Plan</td><td style="padding:8px;color:#D4537E;font-weight:bold">${data.plan.toUpperCase()} — ₱${data.price}/month</td></tr>
      <tr style="background:#FFF8F0"><td style="padding:8px;font-weight:bold;color:#888">GCash Ref #</td><td style="padding:8px;color:#FAC775;font-weight:bold;font-size:18px">${data.reference}</td></tr>
      <tr><td style="padding:8px;font-weight:bold;color:#888">Date</td><td style="padding:8px;color:#1D2B55">${data.date}</td></tr>
    </table>
    <div style="margin-top:16px;padding:14px;background:#FAEEDA;border-radius:8px;border-left:4px solid #BA7517">
      <strong>⚡ Action Required:</strong> Log in to the Admin Panel to approve or reject this payment.
    </div>
    <div style="margin-top:14px;text-align:center">
      <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:12px 28px;border-radius:99px;text-decoration:none;font-weight:bold;display:inline-block">Open Admin Panel</a>
    </div>`
  );
  // Update user subscription to pending
  const users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const uRows = users.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === data.username) {
      users.getRange(i+1,10).setValue("pending");
      break;
    }
  }
  return {success:true};
}

// ── PAYMENTS: GET PENDING ───────────────────────────────────
function getPendingPayments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  let payments = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,plan,price,reference,date,status] = rows[i];
    payments.push({id:Number(id),name,username,plan,price:Number(price),reference,date,status});
  }
  // Sort by date desc
  payments.reverse();
  return {success:true, payments};
}

// ── PAYMENTS: APPROVE / REJECT ──────────────────────────────
function approvePayment(data) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const payments = ss.getSheetByName(SHEET_PAYMENTS);
  const users    = ss.getSheetByName(SHEET_USERS);
  const pRows    = payments.getDataRange().getValues();

  // Update payment status
  for (let i = 1; i < pRows.length; i++) {
    if (pRows[i][2] === data.username && pRows[i][5] === data.reference) {
      payments.getRange(i+1, 8).setValue(data.approve ? "approved" : "rejected");
      break;
    }
  }

  // Update user subscription
  const uRows = users.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === data.username) {
      users.getRange(i+1,10).setValue(data.approve ? data.plan : "free");
      break;
    }
  }

  const action = data.approve ? "approved" : "rejected";
  logActivity({message: "Admin " + action + " " + data.plan + " payment for @" + data.username});
  // Notify user via email (if we had their email — log for now)
  // Send confirmation to admin
  sendAdminEmail(
    data.approve ? "✅ Payment Approved" : "❌ Payment Rejected",
    `<h3 style="color:#1D2B55">Payment ${data.approve ? 'Approved' : 'Rejected'}</h3>
    <p>You have ${data.approve ? '<strong style="color:#1D9E75">approved</strong>' : '<strong style="color:#E24B4A">rejected</strong>'} the payment for:</p>
    <table style="width:100%;border-collapse:collapse;margin-top:10px">
      <tr><td style="padding:8px;font-weight:bold;color:#888">Username</td><td style="padding:8px">@${data.username}</td></tr>
      <tr style="background:#FFF8F0"><td style="padding:8px;font-weight:bold;color:#888">Plan</td><td style="padding:8px">${data.plan}</td></tr>
      <tr><td style="padding:8px;font-weight:bold;color:#888">Reference</td><td style="padding:8px">${data.reference}</td></tr>
    </table>`
  );
  return {success:true};
}

// ── CHILDREN: ADD ───────────────────────────────────────────
function addChild(data) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const children = ss.getSheetByName(SHEET_CHILDREN);
  const users    = ss.getSheetByName(SHEET_USERS);

  // Check child username exists and is a 'kid' type
  const uRows = users.getDataRange().getValues();
  let childFound = false;
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === data.childUsername) {
      childFound = true;
      break;
    }
  }
  if (!childFound) return {success:false, error:"Username not found. Make sure your child has registered first!"};

  // Check not already linked
  const cRows = children.getDataRange().getValues();
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][1] === data.parentUsername && cRows[i][2] === data.childUsername) {
      return {success:false, error:"This child account is already linked to your dashboard!"};
    }
  }

  const id = children.getLastRow();
  children.appendRow([id, data.parentUsername, data.childUsername, new Date().toLocaleDateString()]);
  logActivity({message: "@" + data.parentUsername + " linked child account @" + data.childUsername});
  return {success:true};
}

// ── CHILDREN: GET WITH SCORES ───────────────────────────────
function getChildren(data) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const children = ss.getSheetByName(SHEET_CHILDREN);
  const users    = ss.getSheetByName(SHEET_USERS);
  const scores   = ss.getSheetByName(SHEET_SCORES);

  const cRows = children.getDataRange().getValues();
  const uRows = users.getDataRange().getValues();
  const sRows = scores.getDataRange().getValues();

  let result = [];
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][1] !== data.parentUsername) continue;
    const childUsername = cRows[i][2];

    // Get child user info
    let childUser = null;
    for (let j = 1; j < uRows.length; j++) {
      if (uRows[j][2] === childUsername) {
        childUser = {name:uRows[j][1], username:uRows[j][2], type:uRows[j][4]};
        break;
      }
    }
    if (!childUser) continue;

    // Get child scores
    let childScores = [];
    for (let j = 1; j < sRows.length; j++) {
      if (sRows[j][2] === childUsername) {
        childScores.push({score:Number(sRows[j][4]), diff:sRows[j][5], stars:sRows[j][6], date:sRows[j][7]});
      }
    }
    childScores.reverse(); // most recent first

    result.push({...childUser, scores:childScores});
  }

  return {success:true, children:result};
}

// ── ADMIN: ALL DATA ──────────────────────────────────────────
function getAdminData() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const uRows   = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  const sRows   = ss.getSheetByName(SHEET_SCORES).getDataRange().getValues();
  const aRows   = ss.getSheetByName(SHEET_ACTIVITY).getDataRange().getValues();

  const users = [];
  for (let i = 1; i < uRows.length; i++) {
    const [id,name,username,password,type,role,joined,rounds,best,subscription] = uRows[i];
    users.push({id,name,username,type,role,joined,rounds:Number(rounds),best:Number(best),subscription:subscription||'free'});
  }

  const scores = [];
  for (let i = 1; i < sRows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = sRows[i];
    scores.push({name,username,type,score:Number(score),diff,stars,date});
  }

  const activity = [];
  for (let i = 1; i < aRows.length; i++) {
    const [id,message,time,date] = aRows[i];
    activity.push({m:message,t:time,d:date});
  }

  return {success:true, users, scores, activity:activity.reverse().slice(0,50)};
}
