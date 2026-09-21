# astrbot_plugin_m365

Connects AstrBot to **Microsoft 365** via the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview).

Features are toggled individually — turn on only what you need:

| Feature | Config toggle | Default |
|---------|--------------|---------|
| Read Teams channel messages | `enable_teams_read` | ✅ ON |
| Read inbox emails | `enable_email_read` | ✅ ON |
| Save email drafts | `enable_email_draft` | ✅ ON |
| Send / reply to emails | `enable_email_send` | ❌ OFF |
| Background new-email notifications | `enable_background_poll` | ❌ OFF |

> **Teams send:** Microsoft deprecated the Incoming Webhook connector and the
> `ChannelMessage.Send` Graph permission requires Resource-Specific Consent (RSC)
> which is not available as a standard app-only permission. The recommended
> replacement is **Power Automate Workflows** (HTTP trigger → "Post message to a
> channel"). Once you have a Workflow URL, it can be wired up as a config field
> — see the note at the bottom of this file.

---

## 1 — Azure App Registration (one-time setup)

### 1.1 Create the app

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it something like `AstrBot M365 Bridge`
3. **Supported account types:** *Accounts in this organisational directory only*
4. No redirect URI needed — this plugin uses client credentials, not user login
5. Click **Register**

### 1.2 Note your IDs

On the app's **Overview** page, copy:
- **Application (client) ID** → `client_id` in plugin config
- **Directory (tenant) ID** → `tenant_id` in plugin config

### 1.3 Upload your certificate

This plugin uses certificate authentication (client secrets are not required).

1. **Certificates & secrets** → **Certificates** tab → **Upload certificate**
2. Upload your `.cer` or `.pem` public key file
3. Note the **Thumbprint** shown after upload — you'll need it in the plugin config

To generate a self-signed certificate if you don't have one (PowerShell):
```powershell
$cert = New-SelfSignedCertificate -Subject "CN=AstrBot M365 Plugin" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable -KeySpec Signature `
    -NotAfter (Get-Date).AddYears(2)
Export-Certificate -Cert $cert -FilePath "$env:DESKTOP\astrbot_m365.cer"
Export-PfxCertificate -Cert $cert -FilePath "$env:DESKTOP\astrbot_m365.pfx" `
    -Password (ConvertTo-SecureString "your-passphrase" -Force -AsPlainText)
```

To get the thumbprint of an existing PEM file:
```bash
openssl x509 -in cert.pem -noout -fingerprint -sha1
```

### 1.4 Grant API permissions

1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
2. Add all of the following:

| Permission | Purpose |
|-----------|---------|
| `Team.ReadBasic.All` | List teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read Teams channel messages |
| `Mail.ReadWrite` | Read inbox, save drafts |
| `Mail.Send` | Send and reply to emails |

> **Note:** `ChannelMessage.Send` is not available as a standard application
> permission — it requires Resource-Specific Consent (RSC). Teams channel
> posting is handled via Power Automate Workflows instead (see below).

3. Click **Grant admin consent for [your org]** — this is required; application permissions never work without it.

---

## 2 — Plugin config

| Field | Description |
|-------|-------------|
| `tenant_id` | Azure Directory (tenant) ID |
| `client_id` | App Registration application ID |
| `cert_thumbprint` | SHA-1 thumbprint of the uploaded certificate (40 hex chars, no colons) |
| `cert_private_key` | Full PEM content of the private key (paste including `-----BEGIN...` lines) |
| `user_email` | The M365 user whose mailbox and Teams the plugin acts on (e.g. `user@yourcompany.com`) |
| `watch_team_name` | Default Teams team name (used when no team is given in a command) |
| `watch_channel_name` | Default channel name (default: `General`) |
| `email_inbox_folder` | Folder for /m365-email-read (default: `inbox`) |
| `email_max_results` | Max emails per read (default: 10) |
| `teams_max_results` | Max Teams messages per read (default: 10) |
| `poll_interval_seconds` | Background poll interval in seconds (default: 120) |

---

## 5 — Teams send via Power Automate Workflows

Microsoft deprecated Incoming Webhooks. The replacement is a **Power Automate Workflow** with an HTTP trigger:

1. In Teams, open the channel you want to post to
2. Click **...** (More options) → **Workflows**
3. Search for **"Post to a channel when a webhook request is received"**
4. Follow the wizard — it creates a Flow with an HTTP POST trigger URL
5. Copy the trigger URL

Once this plugin has a `teams_workflow_url` config field, paste the URL there.
The plugin will POST `{"text": "..."}` to that URL to send messages.

> If you don't see the Workflows option, ask your Microsoft 365 admin to enable
> Power Automate for your tenant.

---

## 3 — Slash commands (current)

```
/m365-teams-list
    List all accessible Teams and their channels.

/m365-teams-read [team] [channel] [count]
    Read the last N messages from a Teams channel.
    Defaults to watch_team_name / watch_channel_name if not given.
    Example: /m365-teams-read "Contoso IT" General 5

/m365-email-read [n]
    Show the last n emails from the inbox.
    Example: /m365-email-read 5

/m365-email-reply <id> <reply text>
    Reply to an email. <id> is the 8-char short ID from /m365-email-read.
    Requires enable_email_send = true.
    Example: /m365-email-reply a3f2b901 Thanks, will follow up tomorrow.

/m365-email-draft <to> <subject> | <body>
    Save a draft email (does not send).
    Example: /m365-email-draft alice@example.com Meeting recap | Hi Alice, here are the notes...

/m365-email-send <to> <subject> | <body>
    Send an email immediately.
    Requires enable_email_send = true.
    Example: /m365-email-send bob@example.com Quick update | Hi Bob, all done on my end.
```

---

## 4 — Natural language (LLM tools)

When your AstrBot model supports tool/function calls, you can use plain language:

> *"Show me the last 5 messages in the General channel"*
> *"Post a message to Teams saying the build is done"*
> *"What emails have I got today?"*
> *"Draft a reply to email a3f2b901 saying I'll call tomorrow"*

---

## 6 — Background email notifications

Set `enable_background_poll = true` and `poll_interval_seconds` to receive automatic
notifications of new unread emails in the active AstrBot session.

**Note:** Microsoft Graph API has rate limits. Keep `poll_interval_seconds` at 60 or
higher to avoid throttling.

---

## 7 — Security notes

- The certificate private key grants **app-level access** to your tenant — treat it like a password and never commit it to source control.
- `enable_email_send` is **off by default**. Only turn it on after verifying config.
- This plugin uses **application permissions** (no user login required), which means the
  app can read/write on behalf of `user_email` without a logged-in session. Scope it
  carefully — revoke the app registration if it is no longer needed.
