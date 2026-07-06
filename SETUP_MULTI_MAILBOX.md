# Multi-mailbox sending (split one alias across several mailboxes)

## What this solves

You want 5 recipients to receive mail **from person X**, but split the sending
across mailboxes to stay under Gmail's per-account daily limits — e.g. 2 sent
from **your** mailbox (posing as X) and 3 from **Z's** mailbox (posing as X).

The trick is that two things are independent:

| Concept | Controlled by | Column |
|---------|---------------|--------|
| **Who it appears from** (the identity) | the `From:` header = the alias | `Alias Email` |
| **Which mailbox actually sends** (whose quota + Sent folder) | which account's token we use | `Send From` |

- `Alias Email` = **X** for all 5 rows → everyone sees "From: X".
- `Send From` = **your address** on 2 rows, **Z's address** on 3 rows.

Gmail's sending quota is charged to the **actual sending mailbox**, *not* to the
alias. So sending as X from two different mailboxes genuinely gives you two
separate quotas — exactly the point of the split.

## How the code decides the path (already wired in `code.gs`)

For each row, `sendOriginals` / `sendFollowUps` look at `Send From`:

- **blank, or = the account running the script** → sends natively via the
  built-in Gmail service (`"me"`), exactly like before. **No service account
  needed for these rows.**
- **any other mailbox** → impersonates it via a Service Account with
  Domain-Wide Delegation and sends through the Gmail REST API.

Either way, `buildRawMime` puts the alias in the `From:` header.

---

## The one hard requirement: the alias must be verified on EACH sending mailbox

For "From: X" to actually stick, **X must be a verified "Send mail as" address
on every mailbox that sends as X.** This is the same Gmail feature you already
use on your own account (Settings → Accounts → *Send mail as*).

- Your mailbox already has X verified (that's your current alias setup). ✅
- **Z's mailbox must ALSO have X added and verified** under Z's
  Settings → Accounts → *Send mail as*.

If a mailbox tries to send as an alias it hasn't verified, Gmail rewrites the
From to the real account (or the send fails). Two ways to add it on Z:

1. **Z does it manually** in their Gmail settings and clicks the verification
   link (simplest; works for any address).
2. **Admin/API**, if X is a domain-owned address (like a Group/alias
   `x@browserstack.com`): a domain-owned send-as is auto-verified. This can be
   done with the Gmail API `users.settings.sendAs.create` (scope
   `gmail.settings.sharing`) or by an admin.

> If X is one specific person's personal address, option 1 (they verify it) is
> the reliable route. Get their consent first — you're sending as them.

---

## Setup — only needed for the "other mailbox" (Z) part

Sending from **your own** mailbox as X needs nothing new. The steps below only
enable sending from **other people's** mailboxes (Z, and any future SPOC).

### Part A — Google Cloud (create the service account)

1. <https://console.cloud.google.com> → create/pick a project.
2. **APIs & Services → Library →** enable **Gmail API**.
3. **Credentials → Create credentials → Service account** (e.g. `spoc-mailer`).
4. Open it → **Keys → Add key → Create new key → JSON.** Download it. You need:
   - `client_email` (…@…iam.gserviceaccount.com)
   - `private_key` (`-----BEGIN PRIVATE KEY-----…`)
5. On the **Details** tab, copy the numeric **Unique ID / Client ID**.

### Part B — Admin Console (authorize DWD) — needs a Super Admin

At <https://admin.google.com>:

1. **Security → Access and data control → API controls → Domain-wide
   delegation → Manage → Add new.**
2. **Client ID:** the numeric ID from A.5.
3. **OAuth scopes:**
   ```
   https://www.googleapis.com/auth/gmail.send,https://www.googleapis.com/auth/gmail.readonly
   ```
4. Authorize. (Can take minutes to ~24h to propagate.)

### Part C — Apps Script

1. **Project Settings → Script Properties**, add two:
   - `SA_CLIENT_EMAIL` = the `client_email`.
   - `SA_PRIVATE_KEY`  = the full `private_key` (paste with BEGIN/END lines).
2. Add `multi_mailbox.gs` to the project (it sits alongside `code.gs`).
3. Keep the `script.external_request` scope available in `appsscript.json`.

### Part D — Sheet

Add one column to **Invitelist**:

| Column        | Fill with                                                        |
|---------------|-----------------------------------------------------------------|
| **Send From** | The mailbox that should send this row. Blank = your own mailbox. |

`Alias Email` stays as X on every row. `Sender Name` = display name for X.

Example for your 5-recipient case:

| Email            | Alias Email | Send From          |
|------------------|-------------|--------------------|
| recipient1@…     | x@domain    | *(blank = you)*    |
| recipient2@…     | x@domain    | *(blank = you)*    |
| recipient3@…     | x@domain    | z@domain           |
| recipient4@…     | x@domain    | z@domain           |
| recipient5@…     | x@domain    | z@domain           |

### Part E — Test

Edit the vars in `mmbxSelfTest()` (in `multi_mailbox.gs`) to `TEST_ACCOUNT = z`,
`TEST_ALIAS = x`, and **Run** it. Authorize the `external_request` scope when
prompted. You should receive a mail that comes from Z's mailbox but shows
"From: X".

---

## Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| From shows Z, not X | X isn't a verified *Send mail as* alias on Z's account. Add + verify it (see above). |
| `unauthorized_client` | DWD not approved / wrong Client ID or scopes / still propagating (Part B). |
| `invalid_grant` | `Send From` isn't a real user in the domain, or clock skew. |
| `403 failedPrecondition` on send | The From alias isn't permitted on that account — same fix as row 1. |
| `Missing SA_CLIENT_EMAIL / SA_PRIVATE_KEY` | Script Properties not set (Part C.1). |

## Notes

- **Threading is per-mailbox.** A row's original and follow-up must use the same
  `Send From`, or the stored Thread/Message IDs won't match. Don't change
  `Send From` on a row after the original has gone out.
- **Reply detection** in follow-ups already checks the correct mailbox (yours
  via GmailApp, others via the API). The **Reply Report** still only reads your
  own mailbox — replies to X sent from Z land in **Z's** inbox. To include those
  you'd rebuild `refreshReplyReport` on `mmbxGetThread_(sendFrom, threadId)` per
  row. Ask if you want that.
- **Quotas:** ~2,000 external recipients/day per Workspace user. Splitting
  across N mailboxes multiplies that. The alias does not get its own quota.
