// ============================================================
// WordWise PH — Google Apps Script Backend v3
// Uses GET requests to avoid ALL CORS issues
// ============================================================

const SHEET_USERS    = "Users";
const SHEET_SCORES   = "Scores";
const SHEET_ACTIVITY = "Activity";
const SHEET_PAYMENTS = "Payments";
const SHEET_CHILDREN = "Children";
const ADMIN_EMAIL    = "darcant01@gmail.com";

// ── SETUP ──────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function makeSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.appendRow(headers);
      s.setFrozenRows(1);
      s.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1D2B55").setFontColor("#FAC775");
    }
    return s;
  }

  makeSheet(SHEET_USERS,    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription","SubStart","SubExpiry"]);
  makeSheet(SHEET_SCORES,   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"]);
  makeSheet(SHEET_ACTIVITY, ["ID","Message","Time","Date"]);
  makeSheet(SHEET_PAYMENTS, ["ID","Name","Username","Plan","Price","Reference","Date","Status"]);
  makeSheet(SHEET_CHILDREN, ["ID","ParentUsername","ChildUsername","DateLinked"]);

  // Add default admin if Users sheet is empty
  const users = ss.getSheetByName(SHEET_USERS);
  if (users.getLastRow() < 2) {
    users.appendRow([1,"Admin","admin","wordwise2024","adult","admin",new Date().toLocaleDateString(),0,0,"premium"]);
  }

  return "✅ Setup complete! All 5 sheets ready.";
}

// ── CORS HEADERS ───────────────────────────────────────────
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

// ── RESPOND ────────────────────────────────────────────────
function respond(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── MAIN GET ROUTER ────────────────────────────────────────
// All requests come as GET with ?action=xxx&param=value
function doGet(e) {
  try {
    const p = e.parameter;
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
      case "ping":               result = {success:true, message:"WordWise PH API v3 is running!"}; break;
      default:                   result = {success:false, error:"Unknown action: "+action};
    }
    return respond(result);
  } catch(err) {
    return respond({success:false, error:"Server error: "+err.toString()});
  }
}

// doPost handles receipt image uploads (large data via POST body)
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    const action = data.action || "";

    if(action === "submitPayment") {
      // Handle receipt image if included
      const receipt = data.receipt || null;
      const p = {
        username:  data.username,
        name:      data.name,
        plan:      data.plan,
        price:     data.price,
        reference: data.reference,
        date:      data.date,
        hasReceipt: data.hasReceipt || 'no'
      };
      return respond(submitPaymentWithReceipt(p, receipt));
    }

    // Fallback to GET handler for all other actions
    return doGet(e);
  } catch(err) {
    return respond({success:false, error:"Server error: "+err.toString()});
  }
}

// ── HELPER: get sheet safely ───────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Auto-create if missing
    const headers = {
      [SHEET_USERS]:    ["ID","Name","Username","Password","Type","Role","Joined","Rounds","BestScore","Subscription"],
      [SHEET_SCORES]:   ["ID","Name","Username","Type","Score","Difficulty","Stars","Date"],
      [SHEET_ACTIVITY]: ["ID","Message","Time","Date"],
      [SHEET_PAYMENTS]: ["ID","Name","Username","Plan","Price","Reference","Date","Status"],
      [SHEET_CHILDREN]: ["ID","ParentUsername","ChildUsername","DateLinked"]
    };
    sheet = ss.insertSheet(name);
    if (headers[name]) {
      sheet.appendRow(headers[name]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── LOGIN ───────────────────────────────────────────────────
function login(p) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,password,type,role,joined,rounds,best,subscription] = rows[i];
    if (String(username).trim() === String(p.username).trim() && 
        String(password).trim() === String(p.password).trim()) {
      logActivity({message: name + " logged in"});
      const [,,,,,,,,,sub,subStart,subExpiry] = rows[i];
      return {
        success: true,
        user: {id, name, username, type, role, joined, 
               rounds:Number(rounds)||0, best:Number(best)||0, 
               subscription: subscription||"free",
               subStart: subStart||"",
               subExpiry: subExpiry||""}
      };
    }
  }
  return {success:false, error:"Wrong username or password!"};
}

// ── REGISTER ────────────────────────────────────────────────
function register(p) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  
  if (!p.username || !p.password || !p.name) {
    return {success:false, error:"Missing required fields!"};
  }
  
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim().toLowerCase() === String(p.username).trim().toLowerCase()) {
      return {success:false, error:"Username already taken! Please choose another."};
    }
  }
  
  if (p.password.length < 4) {
    return {success:false, error:"Password must be at least 4 characters!"};
  }
  
  const newId  = rows.length;
  const joined = new Date().toLocaleDateString();
  sheet.appendRow([newId, p.name, p.username, p.password, p.type||"kid", "user", joined, 0, 0, "free"]);
  
  logActivity({message: p.name + " created an account (" + (p.type||"kid") + ")"});
  sendAdminEmail("👤 New User Registered",
    `<h3>New Account</h3><p><b>Name:</b> ${p.name}<br><b>Username:</b> @${p.username}<br><b>Type:</b> ${p.type}</p>`);
  
  return {
    success: true,
    user: {id:newId, name:p.name, username:p.username, type:p.type||"kid", 
           role:"user", joined, rounds:0, best:0, subscription:"free"}
  };
}

// ── SAVE SCORE ──────────────────────────────────────────────
function saveScore(p) {
  const scores = getSheet(SHEET_SCORES);
  const users  = getSheet(SHEET_USERS);
  const date   = new Date().toLocaleDateString();
  const id     = scores.getLastRow();
  
  scores.appendRow([id, p.name, p.username, p.type, Number(p.score), p.diff, p.stars, date]);
  
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
  
  logActivity({message: (p.name||"Guest") + " scored " + p.score + "/5 on " + p.diff});
  return {success:true};
}

// ── GET SCORES (leaderboard) ────────────────────────────────
function getScores(p) {
  const sheet = getSheet(SHEET_SCORES);
  const rows  = sheet.getDataRange().getValues();
  let scores  = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = rows[i];
    if (!p.diff || diff === p.diff) {
      scores.push({name:name||"Guest", username, type:type||"guest", 
                   score:Number(score)||0, diff, stars, date});
    }
  }
  scores.sort((a,b) => b.score - a.score);
  return {success:true, scores: scores.slice(0,10)};
}

// ── GET PROFILE ─────────────────────────────────────────────
function getProfile(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get scores
  const scoreSheet = getSheet(SHEET_SCORES);
  const sRows = scoreSheet.getDataRange().getValues();
  let scores = [];
  for (let i = 1; i < sRows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = sRows[i];
    if (username === p.username) {
      scores.push({score:Number(score)||0, diff, stars, date});
    }
  }

  // Get subscription info from Users sheet
  let subscription = "free", subStart = "", subExpiry = "";
  const userSheet = getSheet(SHEET_USERS);
  const uRows = userSheet.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      subscription = uRows[i][9] || "free";
      subStart     = uRows[i][10] || "";
      subExpiry    = uRows[i][11] || "";
      break;
    }
  }

  // Check if subscription has expired
  if ((subscription === "basic" || subscription === "premium") && subExpiry) {
    const expDate = new Date(subExpiry);
    const today   = new Date();
    today.setHours(0,0,0,0);
    if (expDate < today) {
      subscription = "expired";
      // Update user sheet
      for (let i = 1; i < uRows.length; i++) {
        if (uRows[i][2] === p.username) {
          userSheet.getRange(i+1, 10).setValue("expired");
          break;
        }
      }
    }
  }

  return {success:true, scores:scores.reverse(), subscription, subStart, subExpiry};
}

// ── LOG ACTIVITY ────────────────────────────────────────────
function logActivity(p) {
  const sheet = getSheet(SHEET_ACTIVITY);
  const id    = sheet.getLastRow();
  const time  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a");
  const date  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  sheet.appendRow([id, p.message||"", time, date]);
  return {success:true};
}

// ── CLEAR ACTIVITY ──────────────────────────────────────────
function clearActivity() {
  const sheet = getSheet(SHEET_ACTIVITY);
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last-1);
  return {success:true};
}

// ── SUBMIT PAYMENT ──────────────────────────────────────────
function submitPayment(p) {
  if (!p.username) return {success:false, error:"Not logged in!"};
  if (!p.reference) return {success:false, error:"Reference number is required!"};
  if (!p.plan) return {success:false, error:"No plan selected!"};

  const sheet = getSheet(SHEET_PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  
  // Check duplicate reference
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).trim() === String(p.reference).trim()) {
      return {success:false, error:"This GCash reference number was already submitted!"};
    }
  }
  
  const id   = sheet.getLastRow();
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  sheet.appendRow([id, p.name, p.username, p.plan, Number(p.price)||0, p.reference, date, "pending"]);
  
  // Set user subscription to pending
  const users = getSheet(SHEET_USERS);
  const uRows = users.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      users.getRange(i+1, 10).setValue("pending");
      break;
    }
  }
  
  logActivity({message: p.name + " submitted " + p.plan + " payment (ref: " + p.reference + ")"});
  
  sendAdminEmail("💳 New Payment Submitted!",
    `<h3 style="color:#1D2B55">New Payment Request</h3>
     <table style="width:100%;border-collapse:collapse;font-family:Arial">
       <tr><td style="padding:8px;color:#888;font-weight:bold">Name</td><td style="padding:8px">${p.name}</td></tr>
       <tr style="background:#f9f9f9"><td style="padding:8px;color:#888;font-weight:bold">Username</td><td style="padding:8px">@${p.username}</td></tr>
       <tr><td style="padding:8px;color:#888;font-weight:bold">Plan</td><td style="padding:8px;color:#D4537E;font-weight:bold">${p.plan.toUpperCase()} — ₱${p.price}/mo</td></tr>
       <tr style="background:#f9f9f9"><td style="padding:8px;color:#888;font-weight:bold">GCash Ref #</td><td style="padding:8px;font-size:20px;font-weight:bold;color:#BA7517">${p.reference}</td></tr>
       <tr><td style="padding:8px;color:#888;font-weight:bold">Date</td><td style="padding:8px">${date}</td></tr>
     </table>
     <div style="margin-top:16px;padding:14px;background:#FAEEDA;border-radius:8px">
       <b>⚡ Action Required:</b> Login to the Admin Panel to approve this payment.
     </div>
     <div style="margin-top:14px;text-align:center">
       <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:12px 28px;border-radius:99px;text-decoration:none;font-weight:bold">Open Admin Panel</a>
     </div>`);
  
  return {success:true, message:"Payment submitted successfully!"};
}

// ── SUBMIT PAYMENT WITH RECEIPT ────────────────────────────────────
function submitPaymentWithReceipt(p, receiptBase64) {
  // Run normal submitPayment logic
  const result = submitPayment(p);
  if(!result.success) return result;

  // Send email with receipt image attached
  if(receiptBase64 && receiptBase64.startsWith('data:image')) {
    try {
      const matches = receiptBase64.match(/^data:([^;]+);base64,(.+)$/);
      if(matches) {
        const mimeType = matches[1];
        const base64Data = matches[2];
        const blob = Utilities.newBlob(
          Utilities.base64Decode(base64Data),
          mimeType,
          'gcash_receipt_' + p.username + '.jpg'
        );
        const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
        MailApp.sendEmail({
          to: ADMIN_EMAIL,
          subject: "WordWise PH | 📸 Receipt for @" + p.username + " — " + p.plan.toUpperCase(),
          htmlBody: `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
            <div style="background:#1D2B55;padding:20px;border-radius:12px 12px 0 0;text-align:center">
              <h2 style="color:#FAC775;margin:0">WordWise PH</h2>
              <p style="color:#B5D4F4;margin:4px 0 0">GCash Receipt Submitted</p>
            </div>
            <div style="background:#fff;padding:24px;border:1px solid #F4C0D1;border-top:none;border-radius:0 0 12px 12px">
              <h3 style="color:#1D2B55">Payment Receipt from @${p.username}</h3>
              <table style="width:100%;border-collapse:collapse;font-family:Arial">
                <tr><td style="padding:8px;color:#888;font-weight:bold">Name</td><td style="padding:8px">${p.name}</td></tr>
                <tr style="background:#f9f9f9"><td style="padding:8px;color:#888;font-weight:bold">Username</td><td style="padding:8px">@${p.username}</td></tr>
                <tr><td style="padding:8px;color:#888;font-weight:bold">Plan</td><td style="padding:8px;color:#D4537E;font-weight:bold">${p.plan.toUpperCase()} — ₱${p.price}/mo</td></tr>
                <tr style="background:#f9f9f9"><td style="padding:8px;color:#888;font-weight:bold">GCash Ref #</td><td style="padding:8px;font-size:20px;font-weight:bold;color:#BA7517">${p.reference}</td></tr>
                <tr><td style="padding:8px;color:#888;font-weight:bold">Date</td><td style="padding:8px">${date}</td></tr>
              </table>
              <div style="margin-top:16px;padding:14px;background:#E1F5EE;border-radius:8px;border-left:4px solid #1D9E75">
                <b>📸 Receipt screenshot is attached to this email.</b>
              </div>
              <div style="margin-top:14px;text-align:center">
                <a href="https://darcant01.github.io/wordwiseph/" style="background:#D4537E;color:#fff;padding:12px 28px;border-radius:99px;text-decoration:none;font-weight:bold">Open Admin Panel to Approve</a>
              </div>
            </div>
          </div>`,
          attachments: [blob]
        });
      }
    } catch(imgErr) {
      Logger.log("Receipt email error: " + imgErr.toString());
    }
  }

  return result;
}

// ── GET PENDING PAYMENTS ────────────────────────────────────
function getPendingPayments() {
  const sheet = getSheet(SHEET_PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  let payments = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,name,username,plan,price,reference,date,status] = rows[i];
    payments.push({id:Number(id), name, username, plan, price:Number(price)||0, reference, date, status});
  }
  return {success:true, payments: payments.reverse()};
}

// ── APPROVE PAYMENT ─────────────────────────────────────────
function approvePayment(p) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const payments = getSheet(SHEET_PAYMENTS);
  const users    = getSheet(SHEET_USERS);
  const approve  = p.approve === "true" || p.approve === true;
  
  const pRows = payments.getDataRange().getValues();
  for (let i = 1; i < pRows.length; i++) {
    if (String(pRows[i][2]) === String(p.username) && String(pRows[i][5]) === String(p.reference)) {
      payments.getRange(i+1, 8).setValue(approve ? "approved" : "rejected");
      break;
    }
  }
  
  const uRows = users.getDataRange().getValues();
  for (let i = 1; i < uRows.length; i++) {
    if (uRows[i][2] === p.username) {
      if (approve) {
        const startDate = new Date();
        const expDate   = new Date();
        expDate.setDate(expDate.getDate() + 30); // 30-day subscription
        const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
        users.getRange(i+1, 10).setValue(p.plan);       // Subscription
        users.getRange(i+1, 11).setValue(fmt(startDate)); // SubStart
        users.getRange(i+1, 12).setValue(fmt(expDate));   // SubExpiry
      } else {
        users.getRange(i+1, 10).setValue("free");
        users.getRange(i+1, 11).setValue("");
        users.getRange(i+1, 12).setValue("");
      }
      break;
    }
  }
  
  const action = approve ? "approved" : "rejected";
  logActivity({message: "Admin " + action + " " + p.plan + " plan for @" + p.username});
  sendAdminEmail(approve ? "✅ Payment Approved" : "❌ Payment Rejected",
    `<p>You have <b>${action}</b> the ${p.plan} payment for <b>@${p.username}</b>.<br>Reference: ${p.reference}</p>`);
  
  return {success:true};
}

// ── ADD CHILD ───────────────────────────────────────────────
function addChild(p) {
  const users    = getSheet(SHEET_USERS);
  const children = getSheet(SHEET_CHILDREN);
  
  const uRows = users.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < uRows.length; i++) {
    if (String(uRows[i][2]).trim() === String(p.childUsername).trim()) {
      found = true; break;
    }
  }
  if (!found) return {success:false, error:"Username not found. Make sure your child has registered first!"};
  
  const cRows = children.getDataRange().getValues();
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][1] === p.parentUsername && cRows[i][2] === p.childUsername) {
      return {success:false, error:"This child is already linked to your account!"};
    }
  }
  
  const date = new Date().toLocaleDateString();
  children.appendRow([children.getLastRow(), p.parentUsername, p.childUsername, date]);
  logActivity({message: "@" + p.parentUsername + " linked child @" + p.childUsername});
  return {success:true};
}

// ── GET CHILDREN ────────────────────────────────────────────
function getChildren(p) {
  const users    = getSheet(SHEET_USERS);
  const children = getSheet(SHEET_CHILDREN);
  const scores   = getSheet(SHEET_SCORES);
  
  const cRows = children.getDataRange().getValues();
  const uRows = users.getDataRange().getValues();
  const sRows = scores.getDataRange().getValues();
  
  let result = [];
  for (let i = 1; i < cRows.length; i++) {
    if (String(cRows[i][1]) !== String(p.parentUsername)) continue;
    const childUsername = cRows[i][2];
    
    let childUser = null;
    for (let j = 1; j < uRows.length; j++) {
      if (uRows[j][2] === childUsername) {
        childUser = {name:uRows[j][1], username:uRows[j][2], type:uRows[j][4]};
        break;
      }
    }
    if (!childUser) continue;
    
    let childScores = [];
    for (let j = 1; j < sRows.length; j++) {
      if (sRows[j][2] === childUsername) {
        childScores.push({score:Number(sRows[j][4])||0, diff:sRows[j][5], stars:sRows[j][6], date:sRows[j][7]});
      }
    }
    result.push({...childUser, scores: childScores.reverse()});
  }
  return {success:true, children:result};
}

// ── ADMIN DATA ──────────────────────────────────────────────
function getAdminData() {
  const uRows = getSheet(SHEET_USERS).getDataRange().getValues();
  const sRows = getSheet(SHEET_SCORES).getDataRange().getValues();
  const aRows = getSheet(SHEET_ACTIVITY).getDataRange().getValues();
  
  const users = [];
  for (let i = 1; i < uRows.length; i++) {
    const [id,name,username,,type,role,joined,rounds,best,subscription,subStart,subExpiry] = uRows[i];
    users.push({id, name, username, type, role, joined, 
                rounds:Number(rounds)||0, best:Number(best)||0, 
                subscription:subscription||"free",
                subStart:subStart||"", subExpiry:subExpiry||""});
  }
  
  const scores = [];
  for (let i = 1; i < sRows.length; i++) {
    const [id,name,username,type,score,diff,stars,date] = sRows[i];
    scores.push({name, username, type, score:Number(score)||0, diff, stars, date});
  }
  
  const activity = [];
  for (let i = 1; i < aRows.length; i++) {
    const [id,message,time,date] = aRows[i];
    activity.push({m:message, t:time, d:date});
  }
  
  return {success:true, users, scores, activity:activity.reverse().slice(0,50)};
}

// ── EMAIL ───────────────────────────────────────────────────
function sendAdminEmail(subject, body) {
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "WordWise PH | " + subject,
      htmlBody: `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
        <div style="background:#1D2B55;padding:20px;border-radius:12px 12px 0 0;text-align:center">
          <h2 style="color:#FAC775;margin:0">WordWise PH</h2>
          <p style="color:#B5D4F4;margin:4px 0 0">Admin Notification</p>
        </div>
        <div style="background:#fff;padding:24px;border:1px solid #F4C0D1;border-top:none;border-radius:0 0 12px 12px">
          ${body}
        </div>
        <p style="text-align:center;color:#888;font-size:12px;margin-top:12px">
          WordWise PH · <a href="https://darcant01.github.io/wordwiseph/">darcant01.github.io/wordwiseph</a>
        </p>
      </div>`
    });
  } catch(e) {
    Logger.log("Email error: " + e.toString());
  }
}
