# Sending as other mailboxes (Distribution List members)

Your current script sends every email with:

```js
Gmail.Users.Messages.send({ raw: rawPayload }, "me");
```

The `"me"` is locked to the single account that authorized the script. The
Apps Script "advanced Gmail service" only ever uses that account's OAuth
token, so it can never send as anyone else.

To send **as other people** (your SPOCs / DL members) you replace that call
with a thin **wrapper around the Gmail REST API** that authenticates as each
SPOC individually. That impersonation is done with a **Service Account** that
your Workspace admin grants **Domain-Wide Delegation (DWD)**.

The mail lands in the SPOC's own **Sent** folder and appears fully as them —
not "on behalf of".

> **Authorization matters.** DWD lets code send as any user in the domain, so
> it can only be enabled by a **Super Admin**, and you should have the SPOCs'
> (and your admin's) agreement before sending on their behalf. That admin
> approval step is the control point — you cannot do this alone, by design.

> **A Distribution List / Google Group is not a mailbox.** You can't
> "impersonate a group" — a group has no Sent folder. You impersonate the
> **individual member users** of the DL. If you want the *From* to read as the
> group address, that's a separate "send-as alias" configured on each user.

---

## Part A — Google Cloud (create the service account)

1. Go to <https://console.cloud.google.com> → create or pick a project.
2. **APIs & Services → Library →** enable **Gmail API**.
3. **APIs & Services → Credentials → Create credentials → Service account.**
   Give it a name (e.g. `spoc-mailer`) and create it. No roles needed.
4. Open the service account → **Keys → Add key → Create new key → JSON.**
   A JSON file downloads. Keep it safe. You need two fields from it:
   - `client_email`  (e.g. `spoc-mailer@yourproject.iam.gserviceaccount.com`)
   - `private_key`   (the `-----BEGIN PRIVATE KEY-----...` block)
5. On the service account **Details** tab, copy the **Unique ID / Client ID**
   (a long number) — you need it for the admin step.

## Part B — Admin Console (authorize Domain-Wide Delegation)

Done by a **Super Admin** at <https://admin.google.com>:

1. **Security → Access and data control → API controls → Domain-wide
   delegation → Manage domain-wide delegation → Add new.**
2. **Client ID:** paste the service account's numeric Client ID from A.5.
3. **OAuth scopes:** paste exactly (comma-separated):

   ```
   https://www.googleapis.com/auth/gmail.send,https://www.googleapis.com/auth/gmail.readonly
   ```

4. Authorize. (Propagation can take a few minutes to ~24h.)

## Part C — Apps Script (store the credentials, add the code)

1. In your Apps Script project: **Project Settings → Script Properties → Add
   script property** twice:
   - `SA_CLIENT_EMAIL` = the `client_email` from the JSON.
   - `SA_PRIVATE_KEY`  = the full `private_key` value from the JSON
     (paste it including the `BEGIN/END` lines; literal `\n` are handled).
2. Add the two files from this repo to your project:
   - `multi_mailbox.gs`         — the API wrapper (token minting + send/read).
   - `multi_mailbox_senders.gs` — multi-mailbox `sendOriginalsMulti()` /
     `sendFollowUpsMulti()` (uses your existing `getColIndex` helper).
3. Make sure `appsscript.json` still lists the Gmail scopes your project uses.
   The wrapper itself only needs `UrlFetchApp` + `Script Properties`, but keep
   `https://www.googleapis.com/auth/script.external_request` available.

## Part D — Sheet changes

Add one column to your **Invitelist** tab:

| Column        | Meaning                                                        |
|---------------|---------------------------------------------------------------|
| **Send From** | The SPOC mailbox to send as (a DL member, real domain user).  |

`Sender Name` (optional) still controls the display name. The address in
**Send From** must be the SPOC's own address, or a send-as alias verified on
that SPOC's account — you cannot forge an arbitrary From.

## Part E — Test, then wire the menu

1. In the editor, open `multi_mailbox.gs`, edit the two vars in
   `mmbxSelfTest()`, and **Run** it. You should get a test mail that appears
   to come from the SPOC. Authorize the `external_request` scope when prompted.
2. Point your menu at the new functions:

   ```js
   .addItem("1. Send Original Emails", "sendOriginalsMulti")
   .addItem("3. Send Follow Ups",      "sendFollowUpsMulti")
   ```

---

## Troubleshooting

| Error | Cause / fix |
|-------|-------------|
| `unauthorized_client` | DWD not approved (or wrong Client ID / scopes) in Admin Console, or still propagating. Recheck Part B. |
| `invalid_grant` | `sub` (Send From) is not a real user in the domain, or clock skew. |
| `Delegation denied for <email>` | The scope isn't authorized for DWD, or you're impersonating outside your domain. |
| `403 ... failedPrecondition` on send | The `From` isn't an address that account may send as. Use the SPOC's own address or a verified alias. |
| `Missing SA_CLIENT_EMAIL / SA_PRIVATE_KEY` | Script Properties not set — see Part C.1. |

## Notes on threading & the reply report

- Thread IDs and message IDs are **per-mailbox**. Because the original and its
  follow-up are both sent from the *same* SPOC, the stored `Gmail Thread ID`
  stays valid for that SPOC — follow-ups thread correctly.
- Replies now land in the **SPOC's** inbox, not yours. `sendFollowUpsMulti()`
  already checks for replies via the API (`mmbxThreadHasReply_`). If you also
  want a cross-mailbox **Reply Report**, rebuild it on `mmbxGetThread_(sendFrom,
  threadId)` per row instead of `GmailApp` (which only sees your mailbox).
