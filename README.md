# Outlook Email MCP

A local `stdio` MCP server for Claude Desktop that reads Outlook emails and their attachments — and sends local files (e.g. a daily dashboard `.pptx`) as Outlook attachments — through the Microsoft Graph API.  

> Status: Functional for personal single-user local use with Claude Desktop.

---

## What's New in v0.2.0

- **`send_outlook_email`** — compose an Outlook email with **local files attached** (e.g. a `.pptx`) and either save it as a **draft** for review (default) or send it immediately (`send_now: true`). Small attachments only (each < 3 MB, inlined via Microsoft Graph). This starts a **new** conversation.
- **`reply_outlook_email`** — reply to an existing message so the response stays in the **same Outlook conversation** (Graph `createReply` / `reply`, not `sendMail`). Draft by default; `send_now: true` sends immediately; `replyAll: true` includes all original recipients.
- **`read_email`** — read the full **text/body** of a specific Outlook email (subject, sender, recipients, date; HTML converted to plain text). The original tools read *attachments*; this reads the *message body* itself.
- **`list_recent_messages`** now also returns a `bodyPreview` snippet per message.
- **Persistent login** — the MSAL token cache (incl. the refresh token) is now persisted to disk, so auth **survives app/process restarts**: sign in once and the server silently refreshes tokens (no more re-auth every restart). This makes unattended/scheduled use possible.
- **Scopes widened** to `User.Read`, `Mail.ReadWrite`, `Mail.Send` (was `User.Read`, `Mail.Read`). Existing users must add the new delegated permissions in Entra and re-consent.
- **Manifest upgraded to DXT v0.2** (`mcp_config` nested under `server`) so it installs as a Desktop Extension (`.mcpb`) on current Claude Desktop. New optional config: `M365_SEND_MAX_BYTES`, `M365_SEND_ALLOWED_DIRS`.

---

## Why This Exists

Claude's built-in Microsoft 365 connector can list emails, read message bodies, and check calendars. But it cannot read the actual content inside email attachments.

That means when you say "What does the PDF in my latest email say?", Claude can see the attachment metadata, but not the text, tables, images, or nested documents inside it.

This project fills that gap — running entirely on your local machine over `stdio`, with no public endpoints or tunnels required.

---

## What It Does

This server runs as a local MCP process started by Claude Desktop. It:

1. Authenticates with Microsoft 365 via device code flow
2. Lists Outlook emails and their attachments through Microsoft Graph
3. Downloads and parses attachment contents locally
4. Returns structured text and image blocks directly to Claude Desktop
5. Reads the **text/body** of a specific email (HTML converted to plain text) via `read_email`, or in offset-based slices via `read_email_body_chunk` for long emails that exceed the single-call character limit
6. Composes outgoing Outlook emails with **local files attached** (e.g. PPTX), saved as a draft for review or sent immediately
7. Replies to an existing Outlook message **in the same conversation thread** via `reply_outlook_email`

### Supported Formats

| Format | What Gets Extracted |
|---|---|
| PDF | Full text content |
| Scanned PDF | OCR text, plus optional rendered page images |
| DOCX | Text and embedded images |
| DOC | Text content |
| PPTX / PPTM / PPSX / POTX | Slide text, notes, and embedded images |
| PPT | Best-effort legacy text extraction |
| XLSX / XLS / CSV | All sheets converted to CSV |
| JPG / JPEG / PNG / GIF / WEBP / BMP / TIFF | Returned as MCP image blocks for visual analysis |
| ZIP / RAR / 7Z | Archive contents recursively parsed file by file |
| MSG | Subject, sender, body, and embedded attachments |
| TXT / MD / JSON / XML / HTML | Raw text |
| Outlook `itemAttachment` | Text content |

### MCP Tools

| Tool | Description |
|---|---|
| `health_check` | Check if the server is alive |
| `begin_auth` | Start device code login flow |
| `auth_status` | Check authentication status |
| `list_recent_messages` | List recent Outlook emails. Each result includes a `messageId` to pass to `read_email` / `reply_outlook_email`. |
| `search_messages` | Full-text search. Each hit includes a `messageId` for the same chain. |
| `list_flagged_messages` | List flagged emails, each with a `messageId`. |
| `list_email_attachments` | List attachments for a specific email (`messageId`) |
| `read_email_attachment` | Download, parse, and return attachment content |
| `read_email` | Read one email by `messageId` (subject, sender, recipients, date, body) |
| `read_email_body_chunk` | Read the message body in offset-based slices when `read_email` truncated it |
| `send_outlook_email` | Compose a **new** email, attach local files (e.g. `.pptx`), and save a draft (default) or send immediately. Starts a new conversation. |
| `reply_outlook_email` | Reply to an existing message **in the same Outlook thread** using that `messageId`. Draft by default; `send_now: true` sends; `replyAll: true` includes all original recipients. |

**Typical chain (list → read → reply in-thread):**

1. `list_recent_messages` / `search_messages` / `list_flagged_messages` — copy `messageId`
2. `read_email` with that `messageId` (then `read_email_body_chunk` if the body was truncated)
3. `reply_outlook_email` with the **same** `messageId` — this stays in the Outlook conversation

> **Outbound send is opt-in per call.** `send_outlook_email` and `reply_outlook_email` default to creating a **draft** in your Drafts folder — nothing leaves your mailbox until you review and send it in Outlook. Pass `send_now: true` to send directly. Each attachment must be **under 3 MB** (see Limitations). Use `reply_outlook_email` for replies; `send_outlook_email` with an `RE:` subject still starts a **new** thread.

---

## Real-World Use Cases

### Retail / Sales Operations
> "Pull the last 5 Daily Dashboard emails, read the Excel attachments, and analyze the sales trend across all store locations over the past week."

### Finance / Accounting
> "Find the latest email from our vendor with 'Invoice' in the subject, read the PDF attachment, and extract the total amount, due date, and line items."

### Legal / Contract Review
> "Open the most recent email from legal@partner.com, read the Word or PowerPoint attachment, and summarize the key terms."

### HR / Recruiting
> "Find emails from recruiting@company.com with attachments, read each resume PDF, and create a comparison table of candidates."

### Sending a report (outbound)
> "Attach `~/Desktop/Daily Dash.pptx` and save a draft to my manager with the subject 'Daily Dash — today' — I'll review it before sending."

> "Send `~/Desktop/Daily Dash.pptx` to jane@contoso.com right now with the subject 'Daily Dash' and a one-line note."

---

## Prerequisites

- Windows 10/11, macOS, or Linux
- [Node.js 20 or later](https://nodejs.org)
- Claude Desktop
- A Microsoft 365 / Outlook account
- A Microsoft Entra app registration (see Step 1 below)

---

## Setup

### 1. Create a Microsoft Entra App Registration

Go to [Microsoft Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**.

- **Name:** anything you like, e.g. `m365-mcp-local`
- **Supported account types:** Accounts in any organizational directory and personal Microsoft accounts

Then:

1. Copy the **Application (client) ID** from the Overview page
2. Go to **Authentication** → enable **Allow public client flows** → **Save**
3. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → add `User.Read`, `Mail.ReadWrite`, and `Mail.Send` → **Grant admin consent**
4. Go to **Manifest** → find `requestedAccessTokenVersion` (it may be nested inside `api`) → set it to `2` → **Save**

> **Why `Mail.ReadWrite` + `Mail.Send`?** Reading emails/attachments needs read access (`Mail.ReadWrite` includes the old `Mail.Read`). Saving a draft needs `Mail.ReadWrite`; sending needs `Mail.Send`. The `send_outlook_email` tool uses both. If you only ever want to *read*, you can drop down to `User.Read` + `Mail.Read` and edit `SCOPES` in `server.mjs` accordingly.

> **Already set up an earlier read-only version?** Adding new scopes does **not** upgrade an existing login automatically. After adding `Mail.ReadWrite` and `Mail.Send` (and granting consent), restart Claude Desktop and run `begin_auth` again — you'll be prompted to consent to the new permissions, and a fresh token with send rights is issued.

> **Why step 4?** When your app supports personal Microsoft accounts, Microsoft Entra requires access tokens to be v2. The portal doesn't always set this automatically, and the `common` endpoint will fail with `AADSTS50059` if the token version is still `null` or `1`. If you skip this step you will get `invalid_grant` errors during `begin_auth`.

> **Using v1 API integrations?** Only set this to `2` if all your Graph/API permissions support v2 tokens (all Microsoft Graph delegated permissions do). If you're integrating custom APIs that only accept v1 tokens, use `M365_TENANT_ID=consumers` (personal accounts only) or a specific tenant ID instead of `common`, and leave `requestedAccessTokenVersion` at its default.

### 2. Clone and Install

```bash
git clone https://github.com/dtvillafana/Outlook-Email-MCP.git
cd Outlook-Email-MCP
npm install
```

### 3. Configure Environment Variables

Copy the example file:

```bash
cp .env.example .env
```

Edit `.env` and fill in your client ID:

```env
M365_CLIENT_ID=your-application-client-id-here
M365_TENANT_ID=common
M365_AUTO_OPEN_BROWSER=true
```

Variable reference:

| Variable | Required | Description |
|---|---|---|
| `M365_CLIENT_ID` | ✅ Yes | Your Entra app's Application (client) ID |
| `M365_TENANT_ID` | No | Default `common` works for most accounts |
| `M365_AUTO_OPEN_BROWSER` | No | Set `true` to auto-open Microsoft login page |
| `M365_MCP_DATA_DIR` | No | Custom path for auth cache; auto-detected if omitted |
| `M365_SEND_MAX_BYTES` | No | Total attachment-size budget for `send_outlook_email` (default `3145728` = 3 MB; must stay under Graph's ~4 MB request limit) |
| `M365_SEND_ALLOWED_DIRS` | No | OS-path-separated allowlist of directories `send_outlook_email` may read from. Empty (default) = any readable path |

### 4. Find Your Node.js Path

You'll need the full path to `node.exe` (Windows) or `node` (macOS/Linux) in the next step.

```powershell
# Windows
where.exe node

# macOS / Linux
which node
```

Example output: `C:\Program Files\nodejs\node.exe`

### 5. Open Claude Desktop's Config File

Locate and open the config file for your platform:

| Platform | Path |
|---|---|
| Windows (standard) | `%APPDATA%\Claude\claude_desktop_config.json` |
| Windows (Store) | `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json` |
| macOS | `~/Library/Application Support/Claude/claude_desktop_config.json` |

If the file does not exist yet, create it.

### 6. Add the Server to Claude Desktop

Add the following entry to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365-attachment-reader-local": {
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": [
        "C:\\path\\to\\Outlook-Email-MCP\\server.mjs"
      ],
      "env": {
        "M365_CLIENT_ID": "your-client-id",
        "M365_TENANT_ID": "common",
        "M365_AUTO_OPEN_BROWSER": "true"
      }
    }
  }
}
```

> **Tips:**
> - Use the full absolute path from Step 4 for `command`.
> - Replace `args[0]` with the actual path to `server.mjs` on your machine.
> - If you already have other MCP servers in the config, merge this entry into the existing `mcpServers` object — do not overwrite the whole file.

### 7. Restart Claude Desktop

Completely quit Claude Desktop and reopen it. Claude Desktop starts the MCP server automatically — you do not need to run `node server.mjs` manually.

### 8. Authenticate with Microsoft 365

In Claude Desktop, type:

```text
Please call begin_auth
```

A browser window will open (or you'll receive a login URL + device code). Complete the Microsoft login flow, then verify:

```text
Please call auth_status
```

You should see your Microsoft account listed as authenticated.

### 9. Verify It Works

Run a quick health check:

```text
Please call health_check
```

Then try a real request:

```text
Show me my recent Outlook emails with attachments
```

```text
Summarize the contents of the attachments from the latest email
```

---

## Suggested Claude Prompts

```text
Please call begin_auth
```

```text
Please call auth_status
```

```text
Show me my recent Outlook emails with attachments
```

```text
Summarize the contents of the attachments from the email
```

```text
Find the latest invoice email and extract the total amount, due date, and line items from the PDF attachment
```

```text
Attach ~/Desktop/Daily Dash.pptx and save a draft to my manager — I'll send it from Outlook
```

```text
Reply to that email in the same thread and save a draft for me to review
```

---

## Sending Email with Local Attachments

The `send_outlook_email` tool composes a **new** outgoing Outlook message with one or more **local files** attached and either saves it as a draft (default) or sends it. It does **not** join an existing conversation — use `reply_outlook_email` for that.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `to` | string \| string[] | — (required) | Recipient(s). A comma/semicolon-separated string or an array. |
| `cc`, `bcc` | string \| string[] | — | Optional additional recipients. |
| `subject` | string | `""` | Email subject. |
| `body` | string | `""` | Email body. |
| `bodyType` | `"Text"` \| `"HTML"` | `"Text"` | Body content type. |
| `attachments` | string[] | `[]` | Local file paths to attach (e.g. a `.pptx`). |
| `send_now` | boolean | `false` | `false` saves a draft for review; `true` sends immediately. |
| `save_to_sent` | boolean | `true` | When sending, also save to Sent Items. |
| `mailbox` | string | `"me"` | Mailbox to act as (defaults to the signed-in user). |

**Behavior**

- **Draft (default):** creates the message via `POST /me/messages` and returns the draft's `webLink` so you can open, review, and send it in Outlook.
- **Send now (`send_now: true`):** sends via `POST /me/sendMail` in a single call.
- **Size guard:** each file must be **under 3 MB** and the total stays within `M365_SEND_MAX_BYTES`. Oversize files are rejected with a clear message (large-attachment upload sessions are not enabled in this build).
- **Path access:** by default any readable path is allowed. Set `M365_SEND_ALLOWED_DIRS` to restrict reads to specific folders.

---

## Replying in the Same Outlook Thread

`send_outlook_email` always starts a **new** conversation, even if the subject is `RE: …`. Use `reply_outlook_email` instead: it calls Microsoft Graph `createReply` / `reply` (or `createReplyAll` / `replyAll`) so Outlook keeps `conversationId`, `In-Reply-To`, and the quoted original.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `messageId` | string | — (required) | ID of the message to reply to (from `list_recent_messages`, `search_messages`, or `read_email`). |
| `body` | string | `""` | Reply text. Graph prepends this to the quoted original. |
| `bodyType` | `"Text"` \| `"HTML"` | `"Text"` | Body content type. |
| `replyAll` | boolean | `false` | `true` replies to the sender and all original recipients. |
| `to`, `cc`, `bcc` | string \| string[] | — | Optional recipient overrides. Omit to use Graph's defaults. |
| `attachments` | string[] | `[]` | Local file paths to attach. Same size limits as send. |
| `send_now` | boolean | `false` | `false` saves a threaded draft; `true` sends immediately. |
| `mailbox` | string | `"me"` | Mailbox to act as (defaults to the signed-in user). |

**Behavior**

- **Draft (default):** `POST /me/messages/{id}/createReply` (or `createReplyAll`), then patches in your reply text and returns the draft's `webLink`.
- **Send now (`send_now: true`):** `POST /me/messages/{id}/reply` (or `replyAll`) in a single call. The reply is saved to Sent Items.
- Recipients default to the original sender (`reply`) or everyone on the original (`replyAll`).

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Claude cannot find MCP tools | Restart Claude Desktop completely. Check that `command` and `args` paths in the config are correct and absolute. |
| `invalid_grant` error with empty `userCode` in logs | Almost always a token-version or account-type mismatch on the Entra app. See the next two rows. |
| `AADSTS50059: No tenant-identifying information found` | Your app does not support the `common` endpoint. Open the Entra app → **Authentication** → set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**, then save. |
| `Property api.requestedAccessTokenVersion is invalid` when saving account types | Open the Entra app → **Manifest** → set `requestedAccessTokenVersion` to `2` → save. Then retry changing the account type. |
| Device code not showing | Make sure `begin_auth` was called successfully. Do not manually enter a code. |
| Want to switch Microsoft accounts | Restart Claude Desktop and call `begin_auth` again in a private browser window. |
| `send_outlook_email` or `reply_outlook_email` fails with a permissions/`ErrorAccessDenied` error | Your login predates the send scopes. Add `Mail.ReadWrite` + `Mail.Send` to the Entra app, grant consent, restart Claude Desktop, and run `begin_auth` again to re-consent. |
| Reply showed up as a new email instead of in the thread | The reply was sent with `send_outlook_email`. Use `reply_outlook_email` with the original `messageId`. |
| "Attachment … exceeds the 3.00 MB per-file limit" | This build only inlines attachments under 3 MB. Compress the file or split the deck; large-attachment upload sessions are not enabled. |
| "File is outside the allowed send directories" | The file is not under `M365_SEND_ALLOWED_DIRS`. Move the file into an allowed folder or update that env var (leave it blank to allow any path). |
| Debug log location | `<M365_MCP_DATA_DIR>\debug.log` — defaults to a subdirectory auto-created next to `server.mjs`. |

---

## Manual Development Run

For debugging outside Claude Desktop, start the server manually:

```powershell
cd Outlook-Email-MCP
node .\server.mjs
```

> **Note:** Do not type into that terminal. It is a `stdio` MCP process and expects an MCP client on standard input/output.

---

## Docker

A `Dockerfile` is included for containerized testing:

```powershell
docker build -t m365-attachment-reader-mcp-local .
docker run --rm -i `
  -e M365_CLIENT_ID=your-client-id `
  -e M365_TENANT_ID=common `
  -e M365_AUTO_OPEN_BROWSER=false `
  m365-attachment-reader-mcp-local
```

> The container still runs as a `stdio` server. For everyday Claude Desktop use, the direct `node` approach in Step 6 is simpler.

---

## Project Structure

```text
Outlook-Email-MCP/
├── server.mjs
├── package.json
├── manifest.json
├── server.json
├── glama.json
├── Dockerfile
├── .env.example
├── .gitignore
├── LICENSE
└── README.md
```

---

## Limitations

- **Single-user only** — one server instance supports one Microsoft account at a time
- **Auth persists across restarts** — the MSAL token cache (incl. refresh token) is written to `msal-cache.json` in the data dir (gitignored); after one `begin_auth` the server silently refreshes tokens. You only re-authenticate if the refresh token expires (~90 days) or is revoked
- **You must create your own Entra app** and supply your own client ID
- **Very large images** may be downscaled or skipped to stay within Claude Desktop payload limits
- **Legacy `.xls`** parsing is best-effort and less reliable than `.xlsx`
- **Outbound attachments are capped at ~3 MB each** — `send_outlook_email` inlines files in a single Graph request; large files (which need chunked upload sessions) are rejected
- **Not suitable** for public or multi-user hosting

---

## License

MIT
