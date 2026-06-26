// ============================================================
//  FILL IN THESE PARAMETERS ONLY
// ============================================================

var SHEET_NAME       = "Invitelist";          // ← tab with your customer list
var EMAIL_COLUMN     = 0;                     // ← 0 = Column A
var NAME_COLUMN      = 1;                     // ← 1 = Column B
var CITY_COLUMN      = 2;                     // ← 2 = Column C (Campaign City)

// ============================================================
//  DO NOT EDIT BELOW
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mail Merge Tools")
    .addItem("1. Send Original Emails",           "confirmAndSendOriginal")
    .addSeparator()
    .addItem("2. Refresh Reply Report",           "refreshReplyReport")
    .addSeparator()
    .addItem("3. Send Follow Up to Non-Repliers", "confirmAndSend")
    .addSeparator()
    .addItem("4. Schedule a One-Off Run",         "promptForOneOffSchedule")
    .addToUi();
}

// ============================================================
//  HELPER — Read all City Templates from Content Sheet
// ============================================================

function getContentMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contentSheet = ss.getSheetByName("Content");
  var lastRow = contentSheet.getLastRow();
  var map = {};
  
  if (lastRow < 2) return map;
  
  // Start from row 2 to skip headers
  for (var i = 2; i <= lastRow; i++) {
    var city = String(contentSheet.getRange(i, 1).getValue()).trim().toLowerCase();
    if (!city) continue;
    
    map[city] = {
      subject:      String(contentSheet.getRange(i, 2).getValue()).trim(),
      originalBody: cellToHtml(contentSheet.getRange(i, 3)),
      followUpBody: cellToHtml(contentSheet.getRange(i, 4))
    };
  }
  return map;
}

// ============================================================
//  HELPER — Convert Spreadsheet Rich Text to HTML
// ============================================================

function cellToHtml(range) {
  var richText = range.getRichTextValue();
  if (!richText) {
    return range.getValue().replace(/\n/g, "<br>");
  }
  
  var runs = richText.getRuns();
  var html = "";
  
  for (var i = 0; i < runs.length; i++) {
    var text = runs[i].getText();
    if (!text) continue;
    
    text = text.replace(/&/g, "&amp;")
               .replace(/</g, "&lt;")
               .replace(/>/g, "&gt;")
               .replace(/\n/g, "<br>");
    
    var style = runs[i].getTextStyle();
    if (style.isBold()) text = "<b>" + text + "</b>";
    if (style.isItalic()) text = "<i>" + text + "</i>";
    if (style.isUnderline()) text = "<u>" + text + "</u>";
    if (style.isStrikethrough()) text = "<s>" + text + "</s>";
    
    html += text;
  }
  return html;
}

// ============================================================
//  HELPER — Get or create a named column, returns 0-based index
// ============================================================

function getOrCreateColumn(sheet, columnName) {
  var lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim() === columnName) return h;
    }
  }
  var newIndex = (lastCol === 0) ? 0 : lastCol;
  sheet.getRange(1, newIndex + 1)
    .setValue(columnName)
    .setFontWeight("bold")
    .setBackground("#E8EAED");
  SpreadsheetApp.flush();
  return newIndex;
}

// ============================================================
//  HELPER — Ask user for lag seconds (0–10)
// ============================================================

function askLagSeconds(defaultSeconds) {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Email Delay",
    "Enter delay between emails in seconds (0 to 10):\n(Default is " + defaultSeconds + " seconds)",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return -1;
  var input = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(input) || input < 0 || input > 10) {
    ui.alert("Invalid input. Please enter a number between 0 and 10.");
    return -1;
  }
  return input * 1000; 
}

// ============================================================
//  STEP 1 — SEND ORIGINAL EMAILS
// ============================================================

function confirmAndSendOriginal() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + SHEET_NAME);
      return;
    }

    var statusCol = getOrCreateColumn(sheet, "Original Email Status");
    var rows      = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      SpreadsheetApp.getUi().alert("No customer data found.");
      return;
    }

    var pending = 0;
    for (var i = 1; i < rows.length; i++) {
      var email  = String(rows[i][EMAIL_COLUMN] || "").trim();
      var status = String(rows[i][statusCol]    || "").trim().toLowerCase();
      if (email && status !== "sent") pending++;
    }

    if (pending === 0) {
      SpreadsheetApp.getUi().alert("All original emails already sent!");
      return;
    }

    var lagMs = askLagSeconds(2);
    if (lagMs === -1) return;

    var ui       = SpreadsheetApp.getUi();
    var response = ui.alert(
      "Send Original Emails?",
      "You are about to send the original email to " + pending + " customer(s).\n\n"
      + "Content will be dynamically pulled based on the 'City' column.\n"
      + "Do you want to continue?",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) sendOriginalEmails(lagMs);

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }
}

function sendOriginalEmails(lagMs) {
  try {
    var ss           = SpreadsheetApp.getActiveSpreadsheet();
    var sheet        = ss.getSheetByName(SHEET_NAME);
    var contentMap   = getContentMap();

    var statusCol   = getOrCreateColumn(sheet, "Original Email Status");
    var threadIdCol = getOrCreateColumn(sheet, "Gmail Thread ID");
    var msgIdCol    = getOrCreateColumn(sheet, "Gmail Message ID");
    var rows        = sheet.getDataRange().getValues();

    var startTime = new Date().getTime();
    var sent = 0, failed = 0, skipped = 0;

    for (var i = 1; i < rows.length; i++) {
      var email  = String(rows[i][EMAIL_COLUMN] || "").trim();
      var name   = String(rows[i][NAME_COLUMN]  || "").trim();
      var city   = String(rows[i][CITY_COLUMN]  || "").trim().toLowerCase();
      var status = String(rows[i][statusCol]    || "").trim().toLowerCase();

      if (!email || status === "sent") { skipped++; continue; }

      if (!contentMap[city]) {
        sheet.getRange(i + 1, statusCol + 1)
          .setValue("Failed: City '" + city + "' not found in Content sheet")
          .setFontColor("#C62828").setBackground("#FFEBEE");
        failed++;
        continue;
      }

      if ((new Date().getTime() - startTime) / 1000 > 330) {
        ScriptApp.newTrigger("sendOriginalEmails").timeBased().after(10000).create();
        SpreadsheetApp.getUi().alert("Script limit reached — auto-resuming in 10 seconds.");
        return;
      }

      try {
        var firstName = name.split(" ")[0] || name;
        var htmlBody  = contentMap[city].originalBody.replace(/\{\{FirstName\}\}/g, firstName);
        var subject   = contentMap[city].subject;

        var encodedSubject = "=?" + "UTF-8" + "?B?" + Utilities.base64Encode(Utilities.newBlob(subject).getBytes()) + "?=";

        var rawOriginal = [
          "To: "           + email,
          "Subject: "      + encodedSubject,
          "MIME-Version: 1.0",
          "Content-Type: text/html; charset=UTF-8",
          "",
          htmlBody
        ].join("\r\n");

        var rawBytes = Utilities.newBlob(rawOriginal).getBytes();
        var sentResult = Gmail.Users.Messages.send(
          { raw: Utilities.base64EncodeWebSafe(rawBytes) },
          "me"
        );

        var rfcMsgId = "";
        try {
          var msgMeta = Gmail.Users.Messages.get("me", sentResult.id, { format: "metadata", metadataHeaders: ["Message-ID"] });
          for (var h = 0; h < msgMeta.payload.headers.length; h++) {
            if (msgMeta.payload.headers[h].name === "Message-ID") rfcMsgId = msgMeta.payload.headers[h].value;
          }
        } catch (metaErr) {}

        sheet.getRange(i + 1, statusCol   + 1).setValue("Sent").setFontColor("#0F6E56").setBackground("#E8F5E9");
        sheet.getRange(i + 1, threadIdCol + 1).setValue(sentResult.threadId);
        sheet.getRange(i + 1, msgIdCol    + 1).setValue(rfcMsgId);
        SpreadsheetApp.flush();

        sent++;
        if (lagMs > 0) Utilities.sleep(lagMs); 

      } catch (emailErr) {
        sheet.getRange(i + 1, statusCol + 1).setValue("Failed: " + emailErr.message).setFontColor("#C62828").setBackground("#FFEBEE");
        failed++;
      }
    }

    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === "sendOriginalEmails") ScriptApp.deleteTrigger(t);
    });

    SpreadsheetApp.getUi().alert("Original Emails Done!\n\nSent: " + sent + "\nFailed: " + failed + "\nSkipped: " + skipped);

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }
}

// ============================================================
//  STEP 2 — REPLY REPORT (DUAL-LAYER: SUBJECT + THREAD MATCH)
// ============================================================

function refreshReplyReport() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SHEET_NAME);
  var contentMap  = getContentMap();

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Sheet not found: " + SHEET_NAME);
    return;
  }

  var threadIdCol = getOrCreateColumn(sourceSheet, "Gmail Thread ID");
  var rows        = sourceSheet.getDataRange().getValues();
  var activeUser  = Session.getActiveUser().getEmail();
  var myEmail     = activeUser ? activeUser.toLowerCase() : "";

  var repliedMap = {};

  // -- HELPER: Extract latest text from a message --
  function processMessages(messages) {
    for (var m = 0; m < messages.length; m++) {
      var from     = messages[m].getFrom();
      var msgEmail = (from.match(/<(.+)>/) || [, from])[1].toLowerCase().trim();
      var date     = messages[m].getDate();

      // Ensure we don't count our own outgoing emails as replies
      if (msgEmail !== myEmail && msgEmail !== "me") {
        var fullBody = messages[m].getPlainBody();
        var dividers = [
          /\nOn .*? wrote:/i,           
          /\n> /i,                      
          /\n--- Original message ---/i,
          /\nFrom:/i,                   
          /\n_+\n/                      
        ];
        
        var latestText = fullBody;
        for (var d = 0; d < dividers.length; d++) {
          latestText = latestText.split(dividers[d])[0]; 
        }
        latestText = latestText.trim();
        if (latestText.length > 1000) latestText = latestText.substring(0, 1000) + "... [Truncated]";

        // Save it if it's the first time we see this email, or if it's a newer reply
        if (!repliedMap[msgEmail] || date > repliedMap[msgEmail].date) {
          repliedMap[msgEmail] = { date: date, body: latestText };
        }
      }
    }
  }

  // -- LAYER 1: FALLBACK SUBJECT SEARCH --
  var uniqueSubjects = [];
  for (var cityKey in contentMap) {
    var subj = contentMap[cityKey].subject;
    if (subj && uniqueSubjects.indexOf(subj) === -1) uniqueSubjects.push(subj);
  }

  uniqueSubjects.forEach(function(subject) {
    var threads = GmailApp.search('subject:"Re: ' + subject + '"');
    threads.forEach(function(thread) {
      processMessages(thread.getMessages());
    });
  });

  // -- LAYER 2: EXACT THREAD ID MATCH --
  for (var i = 1; i < rows.length; i++) {
    var threadId = String(rows[i][threadIdCol] || "").trim();
    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (thread) {
          processMessages(thread.getMessages());
        }
      } catch (e) {
        // Thread deleted from Gmail, ignore
      }
    }
  }

  // -- BUILD THE REPORT --
  var reportData = [], repliedCount = 0, notRepliedCount = 0;
  for (var i = 1; i < rows.length; i++) {
    var name  = String(rows[i][NAME_COLUMN]  || "").trim();
    var email = String(rows[i][EMAIL_COLUMN] || "").toLowerCase().trim();
    if (!email) continue;

    if (repliedMap[email]) {
      reportData.push([name, email, "Replied", repliedMap[email].date, repliedMap[email].body, ""]);
      repliedCount++;
    } else {
      reportData.push([name, email, "Not Replied", "", "", "Pending Follow Up"]);
      notRepliedCount++;
    }
  }

  var total     = repliedCount + notRepliedCount;
  var replyRate = total > 0 ? Math.round((repliedCount / total) * 100) : 0;

  var reportSheet = ss.getSheetByName("Reply Report") || ss.insertSheet("Reply Report");
  reportSheet.clearContents().clearFormats();

  reportSheet.getRange(1, 1, 1, 6)
    .setValues([["TOTAL SENT: " + total, "REPLIED: " + repliedCount, "NOT REPLIED: " + notRepliedCount, "REPLY RATE: " + replyRate + "%", "Last Refreshed: " + new Date().toLocaleString(), ""]])
    .setFontWeight("bold").setBackground("#E8EAED");

  reportSheet.getRange(2, 1, 1, 6)
    .setValues([["Name", "Email", "Status", "Reply Date", "Latest Reply Text", "Notes"]])
    .setFontWeight("bold").setBackground("#4A86E8").setFontColor("#FFFFFF");

  if (reportData.length > 0) {
    reportSheet.getRange(3, 1, reportData.length, 6).setValues(reportData);
    for (var r = 0; r < reportData.length; r++) {
      var rowColor    = reportData[r][2] === "Replied" ? "#E1F5EE" : "#FFF8E1";
      var statusColor = reportData[r][2] === "Replied" ? "#0F6E56" : "#BA7517";
      reportSheet.getRange(r + 3, 1, 1, 6).setBackground(rowColor);
      reportSheet.getRange(r + 3, 3).setFontColor(statusColor).setFontWeight("bold");
      reportSheet.getRange(r + 3, 5).setWrap(true);
    }
  }

  reportSheet.autoResizeColumns(1, 4);
  reportSheet.setColumnWidth(5, 400); 
  reportSheet.autoResizeColumn(6);
  reportSheet.activate();
}

// ============================================================
//  STEP 3 — CONFIRM FOLLOW-UPS (DUAL-LAYER CHECK)
// ============================================================

function confirmAndSend() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SHEET_NAME);
  var rows        = sourceSheet.getDataRange().getValues();
  var contentMap  = getContentMap();
  
  var threadIdCol   = getOrCreateColumn(sourceSheet, "Gmail Thread ID");
  var activeUser    = Session.getActiveUser().getEmail();
  var myEmail       = activeUser ? activeUser.toLowerCase() : "";
  var repliedEmails = new Set();
  
  // -- LAYER 1: FALLBACK SUBJECT SEARCH --
  var uniqueSubjects = [];
  for (var cityKey in contentMap) {
    var subj = contentMap[cityKey].subject;
    if (subj && uniqueSubjects.indexOf(subj) === -1) uniqueSubjects.push(subj);
  }

  uniqueSubjects.forEach(function(subject) {
    var threads = GmailApp.search('subject:"Re: ' + subject + '"');
    threads.forEach(function(thread) {
      var messages = thread.getMessages();
      for (var m = 0; m < messages.length; m++) {
        var from     = messages[m].getFrom();
        var msgEmail = (from.match(/<(.+)>/) || [, from])[1].toLowerCase().trim();
        if (msgEmail !== myEmail && msgEmail !== "me") {
          repliedEmails.add(msgEmail);
        }
      }
    });
  });

  // -- LAYER 2: EXACT THREAD ID MATCH --
  for (var i = 1; i < rows.length; i++) {
    var email    = String(rows[i][EMAIL_COLUMN] || "").toLowerCase().trim();
    var threadId = String(rows[i][threadIdCol]  || "").trim();

    if (email && threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (thread) {
          var messages = thread.getMessages();
          for (var m = 1; m < messages.length; m++) {
            var from     = messages[m].getFrom();
            var msgEmail = (from.match(/<(.+)>/) || [, from])[1].toLowerCase().trim();
            if (msgEmail !== myEmail && msgEmail !== "me") {
              repliedEmails.add(email); // specifically tag the spreadsheet email
              break; 
            }
          }
        }
      } catch (e) {}
    }
  }

  // Calculate pending follow-ups
  var pending = 0;
  for (var i = 1; i < rows.length; i++) {
    var email = String(rows[i][EMAIL_COLUMN] || "").toLowerCase().trim();
    if (email && !repliedEmails.has(email)) pending++;
  }

  var lagMs = askLagSeconds(10);
  if (lagMs === -1) return;

  var ui       = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Send Follow Up Emails?",
    "You are about to send follow-up emails to " + pending + " customer(s) who have not replied.\n\n"
    + "Emails will be sent INSIDE the original thread, matching their assigned City.\n\n"
    + "Do you want to continue?",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) sendFollowUps(repliedEmails, lagMs);
}

// ============================================================
//  STEP 3 (CONTINUED) — SEND FOLLOW-UPS INSIDE ORIGINAL THREAD
// ============================================================

function sendFollowUps(repliedEmails, lagMs) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet  = ss.getSheetByName(SHEET_NAME);
  var contentMap   = getContentMap();

  var followUpStatusCol = getOrCreateColumn(sourceSheet, "Follow Up Status");
  var threadIdCol       = getOrCreateColumn(sourceSheet, "Gmail Thread ID");
  var msgIdCol          = getOrCreateColumn(sourceSheet, "Gmail Message ID");
  var rows              = sourceSheet.getDataRange().getValues();

  var sent = 0, failed = 0, skipped = 0;

  for (var i = 1; i < rows.length; i++) {
    var email    = String(rows[i][EMAIL_COLUMN]      || "").toLowerCase().trim();
    var name     = String(rows[i][NAME_COLUMN]       || "").trim();
    var city     = String(rows[i][CITY_COLUMN]       || "").trim().toLowerCase();
    var status   = String(rows[i][followUpStatusCol] || "").trim().toLowerCase();
    var threadId = String(rows[i][threadIdCol]       || "").trim();
    var msgId    = String(rows[i][msgIdCol]          || "").trim();

    if (!email || status === "sent" || status === "sent (in thread)") { skipped++; continue; }

    if (repliedEmails.has(email)) {
      sourceSheet.getRange(i + 1, followUpStatusCol + 1).setValue("Already Replied").setFontColor("#185FA5").setBackground("#E3F2FD");
      skipped++; continue;
    }

    if (!contentMap[city]) {
      sourceSheet.getRange(i + 1, followUpStatusCol + 1).setValue("Skipped — City missing in Content").setFontColor("#E65100").setBackground("#FFF3E0");
      skipped++; continue;
    }

    var subject         = contentMap[city].subject;
    var rawHtmlTemplate = contentMap[city].followUpBody;

    try {
      if (!threadId || !msgId) {
        var searchString = "to:" + email + " subject:\"" + subject + "\" in:sent";
        var threads = GmailApp.search(searchString, 0, 1);
        
        if (threads.length > 0) {
          var originalThread = threads[0];
          threadId = originalThread.getId();
          var rawContent = originalThread.getMessages()[0].getRawContent();
          var match = rawContent.match(/^Message-ID:\s*(<[^>]+>)/im);
          
          if (match) {
            msgId = match[1].trim();
            sourceSheet.getRange(i + 1, threadIdCol + 1).setValue(threadId);
            sourceSheet.getRange(i + 1, msgIdCol + 1).setValue(msgId);
            SpreadsheetApp.flush();
          }
        }
      }

      if (!threadId || !msgId) {
        sourceSheet.getRange(i + 1, followUpStatusCol + 1).setValue("Skipped — thread not found").setFontColor("#E65100").setBackground("#FFF3E0");
        skipped++; continue;
      }

      var firstName = name.split(" ")[0] || name;
      var htmlBody  = rawHtmlTemplate.replace(/\{\{FirstName\}\}/g, firstName);
      var encodedSubject = "=?" + "UTF-8" + "?B?" + Utilities.base64Encode(Utilities.newBlob("Re: " + subject).getBytes()) + "?=";

      var rawFollowUp = [
        "To: "           + email,
        "Subject: "      + encodedSubject,
        "In-Reply-To: "  + msgId,
        "References: "   + msgId,
        "MIME-Version: 1.0",
        "Content-Type: text/html; charset=UTF-8",
        "",
        htmlBody
      ].join("\r\n");

      var rawBytes = Utilities.newBlob(rawFollowUp).getBytes();
      Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(rawBytes), threadId: threadId }, "me");

      sourceSheet.getRange(i + 1, followUpStatusCol + 1).setValue("Sent (in thread)").setFontColor("#0F6E56").setBackground("#E8F5E9");
      SpreadsheetApp.flush();
      sent++;

      if (lagMs > 0) Utilities.sleep(lagMs); 

    } catch (e) {
      sourceSheet.getRange(i + 1, followUpStatusCol + 1).setValue("Failed: " + e.message).setFontColor("#C62828").setBackground("#FFEBEE");
      failed++;
    }
  }

  refreshReplyReport();
  SpreadsheetApp.getUi().alert("Follow Ups Complete!\n\nSent in thread: " + sent + "\nFailed: " + failed + "\nSkipped: " + skipped + "\n\nReply Report refreshed.");
}

function triggerAuth() {
  ScriptApp.getProjectTriggers();
}

// ============================================================
//  ONE-OFF SCHEDULER
// ============================================================

function promptForOneOffSchedule() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Ask the user for the date and time
  var dateResponse = ui.prompt(
    "Schedule a One-Off Execution",
    "Enter the exact date and time to run.\nFormat: YYYY-MM-DD HH:MM (in 24-hour time)\nExample: 2026-05-20 14:30",
    ui.ButtonSet.OK_CANCEL
  );

  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputStr = dateResponse.getResponseText().trim();
  
  // Parse the input into a JavaScript Date object
  // Replacing space with 'T' and adding seconds makes it safely parseable across browsers/V8
  var dateStr = inputStr.replace(" ", "T") + ":00";
  var runDate = new Date(dateStr);

  // Validate the date
  if (isNaN(runDate.getTime())) {
    ui.alert("❌ Invalid format. Please strictly use YYYY-MM-DD HH:MM.");
    return;
  }
  if (runDate.getTime() < new Date().getTime()) {
    ui.alert("❌ The scheduled time must be in the future.");
    return;
  }

  // 2. Ask which function they want to run
  var actionResponse = ui.prompt(
    "Which action to schedule?",
    "Type '1' for Original Emails\nType '2' for Follow Ups",
    ui.ButtonSet.OK_CANCEL
  );

  if (actionResponse.getSelectedButton() !== ui.Button.OK) return;
  var actionStr = actionResponse.getResponseText().trim();
  
  var functionToRun = "";
  var actionName = "";
  
  if (actionStr === "1") {
    functionToRun = "scheduledSendOriginals";
    actionName = "Original Emails";
  } else if (actionStr === "2") {
    functionToRun = "scheduledSendFollowUps";
    actionName = "Follow Up Emails";
  } else {
    ui.alert("❌ Invalid selection. Please type 1 or 2.");
    return;
  }

  // 3. Create the one-off time-based trigger
  try {
    ScriptApp.newTrigger(functionToRun)
      .timeBased()
      .at(runDate)
      .create();
      
    ui.alert("✅ Success!\n\n" + actionName + " is scheduled to run exactly once on:\n" + runDate.toLocaleString());
  } catch (e) {
    ui.alert("❌ Error creating schedule: " + e.message);
  }
}
