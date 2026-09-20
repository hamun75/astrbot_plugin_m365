# astrbot_plugin_m365

An [AstrBot](https://github.com/Soulter/AstrBot) plugin that connects your AstrBot instance to **Microsoft 365**, giving it the ability to read and interact with Microsoft Teams channels and Outlook email — all from within your AstrBot conversations.

Once installed, you can ask AstrBot things like:
> *"What are the latest messages in the General channel?"*
> *"Send a message to Teams saying the deployment is done."*
> *"What emails have I got in my inbox today?"*
> *"Draft a reply to that last email saying I'll follow up tomorrow."*

AstrBot handles the conversation; this plugin handles the Microsoft 365 side.

---

## What this plugin enables AstrBot to do

### Microsoft Teams
- **List** all Teams and channels the app has access to
- **Read** recent messages from any channel
- **Send** messages to any channel

### Outlook Email
- **Read** emails from your inbox (or any mail folder)
- **Save drafts** without sending
- **Reply** to an email by referencing its ID
- **Send** a new email directly

All features are **toggled independently** in the plugin config — turn on only what you need. Sending email is **off by default** as a safety measure.

---

## How it works

The plugin authenticates to the **Microsoft Graph API** using an Azure App Registration with client credentials (app-only auth). No user login or browser redirect is required — it runs entirely in the background using a tenant ID, client ID, and client secret that you configure once. All Graph API calls are made asynchronously so they don't block AstrBot's event loop.

When your AstrBot model supports tool/function calls, the plugin registers LLM tools so you can trigger everything via natural language. Slash commands are always available as a fallback.

---

## Requirements

- AstrBot v4.x or later
- A Microsoft 365 tenant (business or education — personal accounts are not supported by the Graph API permissions this plugin uses)
- An Azure App Registration with the permissions listed below (free to create)

---

## 1 — Azure App Registration (one-time setup)

### 1.1 Create the app

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it something like `AstrBot M365 Bridge`
3. **Supported account types:** *Accounts in this organisational directory only*
4. No redirect URI needed
5. Click **Register**

### 1.2 Note your IDs

On the app's **Overview** page, copy:
- **Application (client) ID** → `client_id` in plugin config
- **Directory (tenant) ID** → `tenant_id` in plugin config

### 1.3 Create a client secret

1. **Certificates & secrets** → **New client secret**
2. Choose an expiry (24 months recommended)
3. **Copy the Value immediately** — it is only shown once
4. Paste it into `client_secret` in the plugin config

### 1.4 Grant API permissions

1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
2. Add all of the following:

| Permission | Purpose |
|-----------|---------|
| `Team.ReadBasic.All` | List teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read Teams channel messages |
| `ChannelMessage.Send` | Post to Teams channels |
| `Mail.ReadWrite` | Read inbox, save drafts |
| `Mail.Send` | Send and reply to emails |

3. Click **Grant admin consent for [your org]** — required for application permissions to work.

---

## 2 — Installation

Install directly from this repo via AstrBot's plugin manager:

```
https://github.com/hamun75/astrbot_plugin_m365
```

Or clone manually into your AstrBot `plugins/` directory:

```bash
cd /path/to/astrbot/plugins
git clone https://github.com/hamun75/astrbot_plugin_m365
```

AstrBot will install the dependencies from `requirements.txt` automatically on next startup.

---

## 3 — Configuration

Open the plugin settings in AstrBot's admin UI and fill in:

| Field | Description |
|-------|-------------|
| `tenant_id` | Azure Directory (tenant) ID |
| `client_id` | App Registration application ID |
| `client_secret` | Client secret value |
| `user_email` | The M365 user this plugin acts on behalf of (e.g. `you@yourdomain.com`) |
| `watch_team_name` | Default Teams team name (used when no team is specified in a command) |
| `watch_channel_name` | Default channel (default: `General`) |
| `email_inbox_folder` | Folder for email reads (default: `inbox`) |
| `email_max_results` | Max emails per read (default: `10`) |
| `teams_max_results` | Max Teams messages per read (default: `10`) |
| `poll_interval_seconds` | Background poll interval in seconds (default: `120`) |

### Feature toggles

| Toggle | Default | What it controls |
|--------|---------|-----------------|
| `enable_teams_read` | ✅ ON | Read Teams channel messages |
| `enable_teams_send` | ✅ ON | Post messages to Teams channels |
| `enable_email_read` | ✅ ON | Read inbox emails |
| `enable_email_draft` | ✅ ON | Save email drafts |
| `enable_email_send` | ❌ OFF | Send and reply to emails (turn on when ready) |
| `enable_background_poll` | ❌ OFF | Auto-notify on new unread emails |

---

## 4 — Slash commands

```
/m365-teams-list
    List all accessible Teams and their channels.

/m365-teams-read [team] [channel] [count]
    Read the last N messages from a Teams channel.
    Example: /m365-teams-read "Melbit Internal" General 5

/m365-teams-send <channel> <message>
    Post a message to a channel in the default team.
    Example: /m365-teams-send General Deployment complete ✅

/m365-email-read [n]
    Show the last n emails from the inbox.
    Example: /m365-email-read 5

/m365-email-reply <id> <reply text>
    Reply to an email. <id> is the 8-char short ID shown by /m365-email-read.
    Requires enable_email_send = true.
    Example: /m365-email-reply a3f2b901 Thanks, will follow up tomorrow.

/m365-email-draft <to> <subject> | <body>
    Save a draft (does not send).
    Example: /m365-email-draft alice@example.com Meeting recap | Hi Alice, here are the notes...

/m365-email-send <to> <subject> | <body>
    Send an email immediately. Requires enable_email_send = true.
    Example: /m365-email-send bob@example.com Quick update | Hi Bob, all done on my end.
```

---

## 5 — Natural language (LLM tools)

When your AstrBot model supports tool/function calls, you can skip the slash commands entirely and just talk to it:

> *"Show me the last 5 messages in the General channel"*
> *"Post a message to Teams saying the build is done"*
> *"What emails have I got today?"*
> *"Draft a reply to email a3f2b901 saying I'll call tomorrow"*
> *"Send an email to alice@example.com with the subject Meeting Notes"*

---

## 6 — Background email notifications

Set `enable_background_poll = true` to receive automatic notifications of new unread emails in your active AstrBot session. The plugin checks on the interval set by `poll_interval_seconds` (minimum 60 recommended to stay within Microsoft Graph API rate limits).

---

## 7 — Security notes

- The client secret grants **app-level access** to your M365 tenant — treat it like a password and never commit it to source control.
- `enable_email_send` is **off by default**. Verify your configuration before turning it on.
- This plugin uses **application permissions** (no interactive user login). The app can read and write on behalf of `user_email` without a browser session. Revoke the app registration in Entra ID if it is no longer needed.

---

## Author

Built by [Melbit Services](https://melbits.com.au) — Melbourne-based MSP and cybersecurity consultancy.
