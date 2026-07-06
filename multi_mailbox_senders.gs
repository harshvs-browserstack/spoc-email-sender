// ============================================================
//  MULTI-MAILBOX versions of Send Originals / Send Follow Ups
// ------------------------------------------------------------
//  Drop-in replacements for your sendOriginals() / sendFollowUps()
//  that send from EACH row's "Send From" mailbox instead of "me".
//  Requires multi_mailbox.gs and the Service Account + DWD setup.
//
//  New required column in your Invitelist tab:
//      "Send From"  -> the SPOC mailbox to send as (a DL member).
//                      Must be a real user in your Workspace domain.
//  Optional:
//      "Sender Name" -> display name shown to the recipient.
// ============================================================

function sendOriginalsMulti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Sheet not found: " + SHEET_NAME); return; }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0];

  var colMap = {
    email:       getColIndex(sheet, headers, "Email"),
    firstName:   getColIndex(sheet, headers, "First Name"),
    sendFrom:    getColIndex(sheet, headers, "Send From"),     // <-- NEW
    senderName:  getColIndex(sheet, headers, "Sender Name"),
    designation: getColIndex(sheet, headers, "Sender Designation"),
    subject:     getColIndex(sheet, headers, "Subject"),
    body:        getColIndex(sheet, headers, "Body"),
    cc:          getColIndex(sheet, headers, "cc"),
    status:      getColIndex(sheet, headers, "Original Email Status"),
    threadId:    getColIndex(sheet, headers, "Gmail Thread ID"),
    msgId:       getColIndex(sheet, headers, "Gmail Message ID")
  };

  var sent = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentEmail = String(row[colMap.email]).trim();
    var status = String(row[colMap.status]).trim().toLowerCase();
    if (!currentEmail || status === "sent") continue;

    var fromEmail = String(row[colMap.sendFrom]).trim();
    if (!fromEmail) {
      sheet.getRange(i + 1, colMap.status + 1).setValue("Failed: no 'Send From' mailbox").setBackground("#FFEBEE");
      continue;
    }

    var firstName = row[colMap.firstName];
    var htmlBody = String(row[colMap.body])
      .replace(/\{\[first name\]\}|\{\{first name\}\}/gi, firstName)
      .replace(/\{\[sender designation\]\}|\{\{sender designation\}\}/gi, row[colMap.designation])
      .replace(/\n/g, "<br>");

    try {
      var raw = buildRawMimeMulti(
        fromEmail, row[colMap.senderName], currentEmail, row[colMap.cc],
        String(row[colMap.subject]), htmlBody, null
      );

      var res = mmbxSendAs_(fromEmail, raw, null);
      var rfcMsgId = mmbxGetRfcMessageId_(fromEmail, res.id);

      sheet.getRange(i + 1, colMap.status   + 1).setValue("Sent").setBackground("#E8F5E9");
      sheet.getRange(i + 1, colMap.threadId + 1).setValue(res.threadId);
      sheet.getRange(i + 1, colMap.msgId    + 1).setValue(rfcMsgId);
      SpreadsheetApp.flush();

      sent++;
      Utilities.sleep(Math.floor(Math.random() * 5000) + 2000);

    } catch (e) {
      sheet.getRange(i + 1, colMap.status + 1).setValue("Failed: " + e.message).setBackground("#FFEBEE");
    }
  }
  SpreadsheetApp.getUi().alert("Original Emails (multi-mailbox) done!\nSent: " + sent);
}


function sendFollowUpsMulti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var colMap = {
    email:          getColIndex(sheet, headers, "Email"),
    firstName:      getColIndex(sheet, headers, "First Name"),
    sendFrom:       getColIndex(sheet, headers, "Send From"),     // <-- NEW
    senderName:     getColIndex(sheet, headers, "Sender Name"),
    designation:    getColIndex(sheet, headers, "Sender Designation"),
    subject:        getColIndex(sheet, headers, "Subject"),
    followUpBody:   getColIndex(sheet, headers, "Follow up Body"),
    cc:             getColIndex(sheet, headers, "cc"),
    threadId:       getColIndex(sheet, headers, "Gmail Thread ID"),
    msgId:          getColIndex(sheet, headers, "Gmail Message ID"),
    followUpStatus: getColIndex(sheet, headers, "Follow Up Status")
  };

  var sent = 0, skippedReplied = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentEmail    = String(row[colMap.email]).trim();
    var fromEmail       = String(row[colMap.sendFrom]).trim();
    var status          = String(row[colMap.followUpStatus]).trim().toLowerCase();
    var threadId        = String(row[colMap.threadId]).trim();
    var msgId           = String(row[colMap.msgId]).trim();
    var followUpContent = row[colMap.followUpBody];

    if (!currentEmail || !followUpContent || status === "sent (in thread)" || status === "already replied") continue;
    if (!fromEmail)  { sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Skipped - No Send From"); continue; }
    if (!threadId)   { sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Skipped - No Thread ID"); continue; }

    // Reply detection in the SPOC's own mailbox (via API, not GmailApp).
    var hasReplied = false;
    try { hasReplied = mmbxThreadHasReply_(fromEmail, threadId); }
    catch (e) { Logger.log("Reply check failed for " + currentEmail + ": " + e.message); }

    if (hasReplied) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Already Replied").setBackground("#E3F2FD");
      skippedReplied++;
      continue;
    }

    // Recover the RFC Message-ID from the SPOC mailbox if blank.
    if (!msgId) {
      try {
        var t = mmbxGetThread_(fromEmail, threadId);
        if (t && t.messages && t.messages.length) {
          msgId = mmbxGetRfcMessageId_(fromEmail, t.messages[0].id);
          if (msgId) { sheet.getRange(i + 1, colMap.msgId + 1).setValue(msgId); SpreadsheetApp.flush(); }
        }
      } catch (e) { Logger.log("Could not recover Message ID for " + currentEmail); }
    }
    if (!msgId) { sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Skipped - Missing Message ID"); continue; }

    var firstName = row[colMap.firstName];
    var htmlBody = String(followUpContent)
      .replace(/\{\[first name\]\}|\{\{first name\}\}/gi, firstName)
      .replace(/\{\[sender designation\]\}|\{\{sender designation\}\}/gi, row[colMap.designation])
      .replace(/\n/g, "<br>");

    try {
      var raw = buildRawMimeMulti(
        fromEmail, row[colMap.senderName], currentEmail, row[colMap.cc],
        "Re: " + String(row[colMap.subject]), htmlBody, msgId
      );

      mmbxSendAs_(fromEmail, raw, threadId);

      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Sent (in thread)").setBackground("#E8F5E9");
      SpreadsheetApp.flush();
      sent++;
      Utilities.sleep(Math.floor(Math.random() * 5000) + 2000);

    } catch (e) {
      sheet.getRange(i + 1, colMap.followUpStatus + 1).setValue("Failed: " + e.message).setBackground("#FFEBEE");
    }
  }

  var msg = "Follow Ups (multi-mailbox) done!\nSent: " + sent;
  if (skippedReplied > 0) msg += "\nSkipped (Already Replied): " + skippedReplied;
  SpreadsheetApp.getUi().alert(msg);
}
