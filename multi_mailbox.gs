// ============================================================
//  MULTI-MAILBOX SENDER  (API-wrapper via Service Account + DWD)
// ------------------------------------------------------------
//  This module lets you send email AS other people in your
//  Google Workspace domain (e.g. the members of a Distribution
//  List), instead of only from the account that runs the script.
//
//  How it works:
//    1. A Service Account is granted Domain-Wide Delegation (DWD)
//       by a Workspace Super Admin, for the gmail.send +
//       gmail.readonly scopes.
//    2. For each SPOC we mint a short-lived OAuth access token
//       that IMPERSONATES that SPOC (JWT "sub" = spoc email).
//    3. We call the Gmail REST API directly with UrlFetchApp,
//       using the SPOC's address as the path userId. The mail
//       lands in the SPOC's own Sent folder and appears fully
//       as them.
//
//  Setup steps are in SETUP_MULTI_MAILBOX.md.
//  You must set two Script Properties first:
//       SA_CLIENT_EMAIL   -> service account client_email
//       SA_PRIVATE_KEY    -> service account private_key (PEM)
// ============================================================

// Scopes the service account is authorized for in Admin Console.
// gmail.send  -> to send.  gmail.readonly -> to read threads for
// follow-up threading + reply detection across SPOC mailboxes.
var MMBX_SCOPES = "https://www.googleapis.com/auth/gmail.send https://www.googleapis.com/auth/gmail.readonly";

// In-memory token cache so we don't re-mint a token per row.
var MMBX_TOKEN_CACHE = {};

// ------------------------------------------------------------
//  Read the service account credentials from Script Properties
// ------------------------------------------------------------
function mmbxGetCreds_() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty("SA_CLIENT_EMAIL");
  var key   = props.getProperty("SA_PRIVATE_KEY");
  if (!email || !key) {
    throw new Error("Missing SA_CLIENT_EMAIL / SA_PRIVATE_KEY in Script Properties. See SETUP_MULTI_MAILBOX.md.");
  }
  // Script Properties often store newlines as the literal 2 chars "\n".
  return { clientEmail: email, privateKey: key.replace(/\\n/g, "\n") };
}

function mmbxBase64Url_(str) {
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(str).getBytes()).replace(/=+$/, "");
}

// ------------------------------------------------------------
//  Mint an impersonated OAuth access token for one SPOC.
//  This is the heart of the "API wrapper" approach.
// ------------------------------------------------------------
function mmbxGetTokenFor_(subjectEmail) {
  subjectEmail = String(subjectEmail || "").trim().toLowerCase();
  if (!subjectEmail) throw new Error("mmbxGetTokenFor_: empty sender email.");

  var now = Math.floor(Date.now() / 1000);
  var cached = MMBX_TOKEN_CACHE[subjectEmail];
  if (cached && cached.exp - 120 > now) return cached.token; // reuse if >2 min left

  var creds  = mmbxGetCreds_();
  var header = { alg: "RS256", typ: "JWT" };
  var claim  = {
    iss:   creds.clientEmail,
    sub:   subjectEmail,             // <-- impersonate this mailbox
    scope: MMBX_SCOPES,
    aud:   "https://oauth2.googleapis.com/token",
    iat:   now,
    exp:   now + 3600
  };

  var toSign = mmbxBase64Url_(JSON.stringify(header)) + "." + mmbxBase64Url_(JSON.stringify(claim));
  var sigBytes = Utilities.computeRsaSha256Signature(toSign, creds.privateKey);
  var jwt = toSign + "." + Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, "");

  var resp = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  var body = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    // Most common: "unauthorized_client" = DWD not approved for these
    // scopes, or "invalid_grant" = subject not a real user in the domain.
    throw new Error("Token error for " + subjectEmail + ": " + (body.error_description || body.error || resp.getContentText()));
  }

  MMBX_TOKEN_CACHE[subjectEmail] = { token: body.access_token, exp: now + (body.expires_in || 3600) };
  return body.access_token;
}

// ------------------------------------------------------------
//  Send a pre-built (base64url) raw MIME message AS `fromEmail`.
//  `raw` must already be base64url-encoded (buildRawMimeMulti does this).
//  Pass threadId to keep a follow-up inside the original thread.
//  Returns the Gmail API response: { id, threadId, labelIds }.
// ------------------------------------------------------------
function mmbxSendAs_(fromEmail, raw, threadId) {
  var token = mmbxGetTokenFor_(fromEmail);
  var payload = { raw: raw };
  if (threadId) payload.threadId = threadId;

  var url = "https://gmail.googleapis.com/gmail/v1/users/" +
            encodeURIComponent(fromEmail) + "/messages/send";

  var resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var body = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error("Send failed as " + fromEmail + ": " + (body.error && body.error.message ? body.error.message : resp.getContentText()));
  }
  return body;
}

// ------------------------------------------------------------
//  Fetch the RFC Message-ID header of a message in a SPOC's
//  mailbox (needed to thread follow-ups correctly).
// ------------------------------------------------------------
function mmbxGetRfcMessageId_(fromEmail, gmailMessageId) {
  var token = mmbxGetTokenFor_(fromEmail);
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(fromEmail) +
            "/messages/" + gmailMessageId + "?format=metadata&metadataHeaders=Message-ID";
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return "";
  var body = JSON.parse(resp.getContentText());
  var headers = (body.payload && body.payload.headers) || [];
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h].name).toLowerCase() === "message-id") return headers[h].value;
  }
  return "";
}

// ------------------------------------------------------------
//  Read a thread from a SPOC's mailbox. Used for reply detection
//  (replies land in the SPOC's inbox, not yours). Returns null if
//  the thread doesn't exist. Message count > 1 => someone replied.
// ------------------------------------------------------------
function mmbxGetThread_(fromEmail, threadId) {
  var token = mmbxGetTokenFor_(fromEmail);
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(fromEmail) +
            "/threads/" + threadId + "?format=metadata";
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return null;
  return JSON.parse(resp.getContentText()); // { id, messages:[...] }
}

// True if the thread has progressed beyond the single message we sent.
function mmbxThreadHasReply_(fromEmail, threadId) {
  var t = mmbxGetThread_(fromEmail, threadId);
  return !!(t && t.messages && t.messages.length > 1);
}

// ============================================================
//  HYBRID DISPATCHERS  (used by code.gs sendOriginals / sendFollowUps)
// ------------------------------------------------------------
//  These pick the send path per row based on the "Send From" column:
//    - blank OR your own address  -> native Gmail service ("me"),
//      exactly as before. No service account needed for these rows.
//    - any other domain mailbox   -> impersonate it via DWD.
//  In BOTH cases the From header (the alias) is set inside
//  buildRawMime, so the recipient always sees the alias.
// ============================================================

// Lowercased email of the account running the script.
function mmbxActiveUser_() {
  return String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
}

// Is this row sent from the script owner's own mailbox?
function mmbxIsSelf_(account) {
  account = String(account || "").trim().toLowerCase();
  return !account || account === mmbxActiveUser_();
}

// Unified send. `account` = the actual sending mailbox ("" => your own).
// `raw` = base64url MIME (From already set to the alias). Returns {id, threadId}.
function sendFromAccount_(account, raw, threadId) {
  if (mmbxIsSelf_(account)) {
    var opt = { raw: raw };
    if (threadId) opt.threadId = threadId;
    var r = Gmail.Users.Messages.send(opt, "me");
    return { id: r.id, threadId: r.threadId };
  }
  var r2 = mmbxSendAs_(account, raw, threadId);
  return { id: r2.id, threadId: r2.threadId };
}

// Unified RFC Message-ID fetch for a message we just sent.
function rfcMessageIdFromAccount_(account, gmailId) {
  if (mmbxIsSelf_(account)) {
    try {
      var meta = Gmail.Users.Messages.get("me", gmailId, { format: "metadata", metadataHeaders: ["Message-ID"] });
      var hs = (meta.payload && meta.payload.headers) || [];
      for (var h = 0; h < hs.length; h++) if (hs[h].name === "Message-ID") return hs[h].value;
    } catch (e) {}
    return "";
  }
  return mmbxGetRfcMessageId_(account, gmailId);
}

// Unified reply check: does the thread hold more than the 1 message we sent?
function threadHasReplyFromAccount_(account, threadId) {
  if (mmbxIsSelf_(account)) {
    var t = GmailApp.getThreadById(threadId);
    return !!(t && t.getMessageCount() > 1);
  }
  return mmbxThreadHasReply_(account, threadId);
}

// Unified: RFC Message-ID of the FIRST message in a thread (recovery path).
function firstMessageIdFromAccount_(account, threadId) {
  if (mmbxIsSelf_(account)) {
    var t = GmailApp.getThreadById(threadId);
    if (!t) return "";
    var raw = t.getMessages()[0].getRawContent();
    var m = raw.match(/^Message-ID:\s*(<[^>]+>)/im);
    return m ? m[1].trim() : "";
  }
  var th = mmbxGetThread_(account, threadId);
  if (th && th.messages && th.messages.length) return mmbxGetRfcMessageId_(account, th.messages[0].id);
  return "";
}

// ------------------------------------------------------------
//  Quick self-test. Edit the vars, then Run once from the editor.
//  Sends a mail to yourself FROM another mailbox, posing as the alias.
// ------------------------------------------------------------
function mmbxSelfTest() {
  var TEST_ACCOUNT = "z@yourdomain.com";        // the mailbox that actually sends (impersonated)
  var TEST_ALIAS   = "x@yourdomain.com";        // the alias shown in From (must be verified on TEST_ACCOUNT)
  var TEST_NAME    = "Person X";
  var TEST_TO      = Session.getActiveUser().getEmail();

  // buildRawMime lives in code.gs and sets From to the alias.
  var raw = buildRawMime(
    TEST_TO, "", TEST_ALIAS, TEST_NAME,
    "Multi-mailbox alias test",
    "Sent from <b>" + TEST_ACCOUNT + "</b>'s mailbox, posing as <b>" + TEST_ALIAS + "</b>.",
    null, null
  );
  var res = sendFromAccount_(TEST_ACCOUNT, raw, null);
  Logger.log("Sent OK. Gmail id=%s threadId=%s", res.id, res.threadId);
}
