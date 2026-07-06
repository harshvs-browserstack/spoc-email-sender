// ============================================================
// CONFIGURATION
// ============================================================
var SHEET_NAME = "Invitelist"; // Update this to your exact tab name

// ============================================================
// MENU SETUP
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📧 Mailer Menu")
    .addItem("1. Send Original Emails", "sendOriginals")
    .addSeparator()
    .addItem("2. Refresh Reply Report", "refreshReplyReport")
    .addSeparator()
    .addItem("3. Send Follow Ups", "sendFollowUps")
    .addToUi();
}

// ============================================================
// HELPER: GET OR CREATE COLUMN DYNAMICALLY
// ============================================================
function getColIndex(sheet, headers, colName) {
  var lowerName = colName.trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === lowerName) {
      return i;
    }
  }
  // If column doesn't exist, create it at the end
  var newColIndex = headers.length;
  sheet.getRange(1, newColIndex + 1).setValue(colName).setFontWeight("bold").setBackground("#E8EAED");
  headers.push(colName); // Update local headers array
  return newColIndex;
}

// ============================================================
// HELPER: CONSTRUCT RAW MIME EMAIL
// ============================================================
function buildRawMime(to, cc, aliasEmail, senderName, subject, htmlBody, threadId, msgId) {
  var raw = [];
  
  // Handle Alias & Name
  if (aliasEmail) {
    var fromStr = senderName ? '"' + senderName + '" <' + aliasEmail + '>' : aliasEmail;
    raw.push("From: " + fromStr);
  } else if (senderName) {
    var activeEmail = Session.getActiveUser().getEmail();
    raw.push("From: \"" + senderName + "\" <" + activeEmail + ">");
  }

  raw.push("To: " + to);
  if (cc) raw.push("Cc: " + cc);
  
  var encodedSubject = "=?" + "UTF-8" + "?B?" + Utilities.base64Encode(Utilities.newBlob(subject).getBytes()) + "?=";
  raw.push("Subject: " + encodedSubject);
  
  // Add threading headers if follow-up
  if (threadId && msgId) {
    raw.push("In-Reply-To: " + msgId);
    raw.push("References: " + msgId);
  }

  raw.push("MIME-Version: 1.0");
  raw.push("Content-Type: text/html; charset=UTF-8");
  raw.push("");
  raw.push(htmlBody);

  var rawMessage = raw.join("\r\n");
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(rawMessage).getBytes());
}

// ============================================================
// 1. SEND ORIGINAL EMAILS
// ============================================================
function sendOriginals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Sheet not found: " + SHEET_NAME); return; }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0];

  // Get dynamic column indices
  var colMap = {
    email: getColIndex(sheet, headers, "Email"),
    firstName: getColIndex(sheet, headers, "First Name"),
    sendFrom: getColIndex(sheet, headers, "Send From"),   // actual sending mailbox (blank = your own)
    senderName: getColIndex(sheet, headers, "Sender Name"),
    aliasEmail: getColIndex(sheet, headers, "Alias Email"),
    designation: getColIndex(sheet, headers, "Sender Designation"),
    subject: getColIndex(sheet, headers, "Subject"),
    body: getColIndex(sheet, headers, "Body"),
    cc: getColIndex(sheet, headers, "cc"),
    status: getColIndex(sheet, headers, "Original Email Status"),
    threadId: getColIndex(sheet, headers, "Gmail Thread ID"),
    msgId: getColIndex(sheet, headers, "Gmail Message ID")
  };

  var sent = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentEmail = row[colMap.email];
    var status = String(row[colMap.status]).trim().toLowerCase();

    if (!currentEmail || status === "sent") continue;

    var firstName = row[colMap.firstName];
    var body = row[colMap.body];
    var account = String(row[colMap.sendFrom] || "").trim(); // mailbox that actually sends

    // Personalize body
    var htmlBody = body.replace(/\{\[first name\]\}|\{\{first name\}\}/gi, firstName)
                       .replace(/\{\[sender designation\]\}|\{\{sender designation\}\}/gi, row[colMap.designation])
                       .replace(/\n/g, "<br>");

    try {
      var rawPayload = buildRawMime(
        currentEmail, row[colMap.cc], row[colMap.aliasEmail], row[colMap.senderName],
        row[colMap.subject], htmlBody, null, null
      );

      // Send from your own mailbox OR (via impersonation) from someone else's.
      // The From header is still the alias, set inside buildRawMime.
      var sentResult = sendFromAccount_(account, rawPayload, null);

      // Fetch RFC Message-ID (from whichever mailbox actually sent it)
      var rfcMsgId = rfcMessageIdFromAccount_(account, sentResult.id);

      // Write back to sheet
      sheet.getRange(i + 1, colMap.status + 1).setValue("Sent").setBackground("#E8F5E9");
      sheet.getRange(i + 1, colMap.threadId + 1).setValue(sentResult.threadId);
      sheet.getRange(i + 1, colMap.msgId + 1).setValue(rfcMsgId);
      SpreadsheetApp.flush();
      
      sent++;
      Utilities.sleep(Math.floor(Math.random() * 5000) + 2000); // 2-7 second lag

    } catch (e) {
      sheet.getRange(i + 1, colMap.status + 1).setValue("Failed: " + e.message).setBackground("#FFEBEE");
    }
  }
  SpreadsheetApp.getUi().alert("Original Emails Completed!\nSent: " + sent);
}

// ============================================================
// 2. SEND FOLLOW UPS
// ============================================================
// ============================================================
// 2. SEND FOLLOW UPS
// ============================================================
function sendFollowUps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var colMap = {
    email: getColIndex(sheet, headers, "Email"),
    firstName: getColIndex(sheet, headers, "First Name"),
    sendFrom: getColIndex(sheet, headers, "Send From"),   // actual sending mailbox (blank = your own)
    senderName: getColIndex(sheet, headers, "Sender Name"),
    aliasEmail: getColIndex(sheet, headers, "Alias Email"),
    designation: getColIndex(sheet, headers, "Sender Designation"),
    subject: getColIndex(sheet, headers, "Subject"),
    followUpBody: getColIndex(sheet, headers, "Follow up Body"),
    cc: getColIndex(sheet, headers, "cc"),
    threadId: getColIndex(sheet, headers, "Gmail Thread ID"),
    msgId: getColIndex(sheet, headers, "Gmail Message ID"), 
    followUpStatus: getColIndex(sheet, headers, "Follow Up Status")
  };

  var sent = 0;
  var skippedReplied = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentEmail = String(row[colMap.email]).trim();
    var status = String(row[colMap.followUpStatus]).trim().toLowerCase();
    var threadId = row[colMap.threadId];
    var msgId = row[colMap.msgId];
    var followUpContent = row[colMap.followUpBody];
    var account = String(row[colMap.sendFrom] || "").trim(); // mailbox that sent the original

    // Skip if no email, no content, or if it's already sent/replied
    if (!currentEmail || !followUpContent || status === "sent (in thread)" || status === "already replied") continue;
    
    if (!threadId) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Skipped - No Thread ID");
      continue;
    }

    // --- NEW: Bulletproof Reply Detection ---
    // Checks the thread in whichever mailbox actually sent the original.
    // More than 1 message => a reply/bounce/manual response arrived, so abort.
    var hasReplied = false;
    try {
      hasReplied = threadHasReplyFromAccount_(account, threadId);
    } catch (e) {
      Logger.log("Error checking thread for " + currentEmail + ": " + e.message);
    }

    if (hasReplied) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Already Replied").setBackground("#E3F2FD");
      skippedReplied++;
      continue;
    }

    // --- FALLBACK: Fetch missing Message ID if blank ---
    if (!msgId) {
      try {
        msgId = firstMessageIdFromAccount_(account, threadId);
        if (msgId) {
          sheet.getRange(i + 1, colMap.msgId + 1).setValue(msgId);
          SpreadsheetApp.flush();
        }
      } catch (e) {
        Logger.log("Could not recover Message ID for " + currentEmail);
      }
    }

    if (!msgId) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Skipped - Missing Message ID");
      continue;
    }

    // --- SEND THE FOLLOW UP ---
    var firstName = row[colMap.firstName];
    var htmlBody = followUpContent.replace(/\{\[first name\]\}|\{\{first name\}\}/gi, firstName)
                                  .replace(/\{\[sender designation\]\}|\{\{sender designation\}\}/gi, row[colMap.designation])
                                  .replace(/\n/g, "<br>");

    try {
      var rawPayload = buildRawMime(
        currentEmail, row[colMap.cc], row[colMap.aliasEmail], row[colMap.senderName],
        "Re: " + row[colMap.subject], htmlBody, threadId, msgId
      );

      sendFromAccount_(account, rawPayload, threadId);
      
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Sent (in thread)").setBackground("#E8F5E9");
      SpreadsheetApp.flush();
      
      sent++;
      Utilities.sleep(Math.floor(Math.random() * 5000) + 2000); 

    } catch (e) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Failed: " + e.message).setBackground("#FFEBEE");
    }
  }
  
  var alertMsg = "Follow Ups Completed!\nSent: " + sent;
  if (skippedReplied > 0) {
    alertMsg += "\nSkipped (Already Replied/Updated): " + skippedReplied;
  }
  SpreadsheetApp.getUi().alert(alertMsg);
}

// ============================================================
// 3. REFRESH REPLY REPORT
// ============================================================
function refreshReplyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SHEET_NAME);
  var data = sourceSheet.getDataRange().getValues();
  var headers = data[0];

  var colMap = {
    email: getColIndex(sourceSheet, headers, "Email"),
    firstName: getColIndex(sourceSheet, headers, "First Name"),
    threadId: getColIndex(sourceSheet, headers, "Gmail Thread ID")
  };

  var myEmail = Session.getActiveUser().getEmail().toLowerCase();
  var reportData = [];
  var repliedCount = 0;
  var notRepliedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][colMap.email]).toLowerCase().trim();
    var name = data[i][colMap.firstName];
    var threadId = data[i][colMap.threadId];

    if (!email) continue;

    var replied = false;
    var replyDate = "";
    var replyText = "";
    var replySender = "";

    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (thread) {
          var messages = thread.getMessages();
          for (var m = 1; m < messages.length; m++) {
            var from = messages[m].getFrom();
            var msgEmail = (from.match(/<(.+)>/) || [, from])[1].toLowerCase().trim();
            var senderNameRaw = (from.split("<")[0] || "").replace(/"/g, "").trim();

            if (msgEmail !== myEmail && msgEmail !== "me") {
              replied = true;
              replyDate = messages[m].getDate();
              replySender = senderNameRaw || msgEmail;
              
              var body = messages[m].getPlainBody();
              replyText = body.split(/\nOn .*? wrote:/i)[0]
                              .split(/\n> /i)[0]
                              .split(/\n--- Original message ---/i)[0]
                              .trim().substring(0, 500);
              break;
            }
          }
        }
      } catch (e) {} // Ignore deleted threads
    }

    if (replied) {
      reportData.push([name, email, "Replied", replyDate, replySender, replyText, ""]);
      repliedCount++;
    } else {
      reportData.push([name, email, "Not Replied", "", "", "", "Pending Follow Up"]);
      notRepliedCount++;
    }
  }

  // Build Report Sheet
  var reportSheet = ss.getSheetByName("Reply Report") || ss.insertSheet("Reply Report");
  reportSheet.clearContents().clearFormats();

  reportSheet.getRange(1, 1, 1, 7).setValues([["Name", "Email", "Status", "Reply Date", "Latest Reply Sender Name", "Latest Reply Text", "Notes"]])
    .setFontWeight("bold").setBackground("#4A86E8").setFontColor("#FFFFFF");

  if (reportData.length > 0) {
    reportSheet.getRange(2, 1, reportData.length, 7).setValues(reportData);
    for (var r = 0; r < reportData.length; r++) {
      var isReplied = reportData[r][2] === "Replied";
      reportSheet.getRange(r + 2, 1, 1, 7).setBackground(isReplied ? "#E1F5EE" : "#FFF8E1");
      reportSheet.getRange(r + 2, 3).setFontColor(isReplied ? "#0F6E56" : "#BA7517").setFontWeight("bold");
    }
  }

  reportSheet.autoResizeColumns(1, 5);
  reportSheet.setColumnWidth(6, 400); // Give the reply text room to breathe
  reportSheet.activate();
}
