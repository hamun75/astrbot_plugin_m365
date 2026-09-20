"""
AstrBot <-> Microsoft 365 Bridge
=================================
Connects AstrBot to Microsoft 365 via the Microsoft Graph API, using an
Azure App Registration with client-credentials (app-only) authentication.

Capabilities (each toggled independently in config):
  - Teams: read channel messages, send channel messages
  - Email: read inbox, save drafts, reply to emails, send email

Slash commands:
  /m365-teams-list               — list accessible teams and channels
  /m365-teams-read [team] [ch]   — read last N messages from a channel
  /m365-teams-send <ch> <msg>    — post a message to a channel
  /m365-email-read [n]           — read last N emails from inbox
  /m365-email-reply <id> <text>  — reply to an email by its short ID
  /m365-email-draft <to> <subj> <body>  — save a draft
  /m365-email-send  <to> <subj> <body>  — send an email

LLM tools (natural-language triggers):
  list_teams, read_teams_messages, send_teams_message,
  read_emails, reply_to_email, draft_email, send_email

Setup — see README.md for the full Azure App Registration walkthrough.
Required Microsoft Graph application permissions (admin consent needed):
  Team.ReadBasic.All, Channel.ReadBasic.All,
  ChannelMessage.Read.All, ChannelMessage.Send,
  Mail.ReadWrite, Mail.Send
"""

import asyncio
import re
from datetime import datetime, timezone
from typing import Optional

import aiohttp
import msal

from astrbot.api import logger, llm_tool
from astrbot.api.event import AstrMessageEvent, MessageChain, filter
from astrbot.api.star import Context, Star, register

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
TOKEN_REFRESH_BUFFER = 300  # refresh token 5 min before expiry


# ---------------------------------------------------------------------------
# Graph API client
# ---------------------------------------------------------------------------

class GraphClient:
    """Thin async wrapper around the Microsoft Graph REST API.

    Handles token acquisition and renewal via MSAL client credentials.
    All HTTP calls are async (aiohttp).
    """

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self._app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        self._token: Optional[str] = None
        self._token_expiry: float = 0.0
        self._session: Optional[aiohttp.ClientSession] = None

    async def _ensure_session(self):
        if self._session is None or self._session.closed:
            self._session = aiohttp.ClientSession()

    async def close(self):
        if self._session and not self._session.closed:
            await self._session.close()

    def _token_fresh(self) -> bool:
        import time
        return self._token is not None and time.time() < self._token_expiry - TOKEN_REFRESH_BUFFER

    async def _get_token(self) -> str:
        import time
        if not self._token_fresh():
            result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPE)
            if "access_token" not in result:
                err = result.get("error_description", result.get("error", "Unknown MSAL error"))
                raise RuntimeError(f"MSAL token error: {err}")
            self._token = result["access_token"]
            self._token_expiry = time.time() + result.get("expires_in", 3600)
        return self._token  # type: ignore[return-value]

    async def get(self, path: str, params: dict = None) -> dict:
        await self._ensure_session()
        token = await self._get_token()
        url = f"{GRAPH_BASE}{path}"
        async with self._session.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params or {},
        ) as resp:
            data = await resp.json()
            if not resp.ok:
                raise RuntimeError(f"Graph GET {path} → {resp.status}: {data.get('error', {}).get('message', data)}")
            return data

    async def post(self, path: str, body: dict) -> dict:
        await self._ensure_session()
        token = await self._get_token()
        url = f"{GRAPH_BASE}{path}"
        async with self._session.post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=body,
        ) as resp:
            if resp.status == 204:
                return {}
            data = await resp.json()
            if not resp.ok:
                raise RuntimeError(f"Graph POST {path} → {resp.status}: {data.get('error', {}).get('message', data)}")
            return data


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _strip_html(text: str) -> str:
    """Very basic HTML tag stripper for Teams/email body previews."""
    return re.sub(r"<[^>]+>", "", text or "").strip()


def _fmt_dt(iso: str) -> str:
    """Format an ISO-8601 timestamp to a readable local-ish string."""
    if not iso:
        return "unknown"
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.astimezone().strftime("%d %b %Y %H:%M")
    except Exception:
        return iso


def _short_id(full_id: str) -> str:
    """Return the last 8 chars of a Graph object ID as a short handle."""
    return (full_id or "")[-8:]


# ---------------------------------------------------------------------------
# Plugin
# ---------------------------------------------------------------------------

@register(
    "astrbot_plugin_m365",
    "Melbit Services",
    "Connects AstrBot to Microsoft 365 — read/send Teams messages and email via Microsoft Graph API.",
    "1.0.0",
)
class M365Plugin(Star):
    def __init__(self, context: Context, config: dict = None):
        super().__init__(context)
        self.cfg = config or {}

        # ---- feature flags ------------------------------------------------
        self.teams_read  = bool(self.cfg.get("enable_teams_read",  True))
        self.teams_send  = bool(self.cfg.get("enable_teams_send",  True))
        self.email_read  = bool(self.cfg.get("enable_email_read",  True))
        self.email_draft = bool(self.cfg.get("enable_email_draft", True))
        self.email_send  = bool(self.cfg.get("enable_email_send",  False))
        self.bg_poll     = bool(self.cfg.get("enable_background_poll", False))

        # ---- limits / defaults -------------------------------------------
        self.email_max   = int(self.cfg.get("email_max_results",  10) or 10)
        self.teams_max   = int(self.cfg.get("teams_max_results",  10) or 10)
        self.poll_secs   = int(self.cfg.get("poll_interval_seconds", 120) or 120)
        self.user_email  = str(self.cfg.get("user_email", "") or "").strip()
        self.dflt_team   = str(self.cfg.get("watch_team_name",   "") or "").strip()
        self.dflt_chan   = str(self.cfg.get("watch_channel_name", "General") or "General").strip()
        self.inbox_folder = str(self.cfg.get("email_inbox_folder", "inbox") or "inbox").strip()

        # ---- auth ---------------------------------------------------------
        tenant = str(self.cfg.get("tenant_id",     "") or "").strip()
        c_id   = str(self.cfg.get("client_id",     "") or "").strip()
        c_sec  = str(self.cfg.get("client_secret", "") or "").strip()

        if not all([tenant, c_id, c_sec]):
            logger.warning("[m365] tenant_id / client_id / client_secret not fully set — plugin disabled.")
            self._ready = False
            self._graph: Optional[GraphClient] = None
            return

        self._graph = GraphClient(tenant, c_id, c_sec)
        self._ready = True

        # ---- team/channel ID cache (name → id) ----------------------------
        self._team_cache: dict[str, str]          = {}   # team display_name → id
        self._chan_cache: dict[str, dict[str, str]] = {}  # team_id → {chan_name → chan_id}

        # ---- background poll state ----------------------------------------
        self._last_email_id: Optional[str] = None
        self._poll_task: Optional[asyncio.Task] = None
        self._poll_ctx: Optional[AstrMessageEvent] = None

        if self.bg_poll:
            self._poll_task = asyncio.create_task(self._bg_poll_loop())
            logger.info(f"[m365] Background poll started (interval={self.poll_secs}s).")

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    async def terminate(self):
        if self._poll_task:
            self._poll_task.cancel()
        if self._graph:
            await self._graph.close()

    # ------------------------------------------------------------------
    # Internal: team/channel resolution
    # ------------------------------------------------------------------

    async def _load_teams(self) -> dict[str, str]:
        """Return {display_name: id} for all teams the app can see."""
        if self._team_cache:
            return self._team_cache
        data = await self._graph.get("/teams", params={"$select": "id,displayName", "$top": "50"})
        for t in data.get("value", []):
            self._team_cache[t["displayName"].lower()] = t["id"]
        return self._team_cache

    async def _resolve_team(self, name: str) -> tuple[str, str]:
        """Return (display_name_lower, team_id) or raise ValueError."""
        teams = await self._load_teams()
        key = name.strip().lower()
        if key not in teams:
            raise ValueError(
                f"Team '{name}' not found. Available: {', '.join(t.title() for t in teams)}"
            )
        return key, teams[key]

    async def _load_channels(self, team_id: str) -> dict[str, str]:
        """Return {channel_name_lower: channel_id} for a team."""
        if team_id in self._chan_cache:
            return self._chan_cache[team_id]
        data = await self._graph.get(
            f"/teams/{team_id}/channels",
            params={"$select": "id,displayName"},
        )
        mapping = {c["displayName"].lower(): c["id"] for c in data.get("value", [])}
        self._chan_cache[team_id] = mapping
        return mapping

    async def _resolve_channel(self, team_id: str, name: str) -> tuple[str, str]:
        """Return (channel_name_lower, channel_id) or raise ValueError."""
        channels = await self._load_channels(team_id)
        key = name.strip().lower()
        if key not in channels:
            raise ValueError(
                f"Channel '{name}' not found. Available: {', '.join(c.title() for c in channels)}"
            )
        return key, channels[key]

    # ------------------------------------------------------------------
    # Internal: Graph reads
    # ------------------------------------------------------------------

    async def _read_teams_messages(self, team_name: str, chan_name: str, count: int) -> str:
        _, team_id = await self._resolve_team(team_name)
        _, chan_id  = await self._resolve_channel(team_id, chan_name)
        data = await self._graph.get(
            f"/teams/{team_id}/channels/{chan_id}/messages",
            params={"$top": str(min(count, 50)), "$orderby": "createdDateTime desc"},
        )
        msgs = data.get("value", [])
        if not msgs:
            return f"No messages in {team_name.title()} / {chan_name.title()}."
        lines = [f"📨 **{team_name.title()} › {chan_name.title()}** — last {len(msgs)} message(s)\n"]
        for m in msgs:
            sender = (m.get("from") or {}).get("user", {}).get("displayName", "Unknown")
            body   = _strip_html(m.get("body", {}).get("content", ""))[:300]
            ts     = _fmt_dt(m.get("createdDateTime", ""))
            mid    = _short_id(m.get("id", ""))
            lines.append(f"[{ts}] **{sender}** (id: {mid})\n{body}\n")
        return "\n".join(lines)

    async def _send_teams_message(self, team_name: str, chan_name: str, text: str) -> str:
        _, team_id = await self._resolve_team(team_name)
        _, chan_id  = await self._resolve_channel(team_id, chan_name)
        await self._graph.post(
            f"/teams/{team_id}/channels/{chan_id}/messages",
            {"body": {"contentType": "text", "content": text}},
        )
        return f"✅ Message sent to {team_name.title()} › {chan_name.title()}."

    async def _list_teams_text(self) -> str:
        teams = await self._load_teams()
        if not teams:
            return "No teams found (check app permissions and admin consent)."
        lines = ["**Accessible Teams & Channels**\n"]
        for tname, tid in teams.items():
            channels = await self._load_channels(tid)
            chan_list = ", ".join(c.title() for c in channels) or "(no channels)"
            lines.append(f"• **{tname.title()}** — {chan_list}")
        return "\n".join(lines)

    async def _read_emails(self, count: int, folder: str) -> str:
        if not self.user_email:
            return "⚠️ user_email is not configured in plugin settings."
        data = await self._graph.get(
            f"/users/{self.user_email}/mailFolders/{folder}/messages",
            params={
                "$top": str(min(count, 50)),
                "$orderby": "receivedDateTime desc",
                "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead",
            },
        )
        emails = data.get("value", [])
        if not emails:
            return f"No emails in {folder}."
        lines = [f"📧 **Inbox** — last {len(emails)} email(s)\n"]
        for e in emails:
            sender  = (e.get("from") or {}).get("emailAddress", {}).get("address", "?")
            subj    = e.get("subject", "(no subject)")
            preview = (e.get("bodyPreview") or "")[:200].replace("\n", " ")
            ts      = _fmt_dt(e.get("receivedDateTime", ""))
            eid     = _short_id(e.get("id", ""))
            unread  = "🔵 " if not e.get("isRead") else ""
            lines.append(f"{unread}[{ts}] **{sender}** | {subj} (id: {eid})\n{preview}\n")
        return "\n".join(lines)

    async def _resolve_email_id(self, short_id: str) -> str:
        """Find the full Graph message ID from a short (last-8-char) id."""
        data = await self._graph.get(
            f"/users/{self.user_email}/messages",
            params={"$top": "50", "$select": "id"},
        )
        for e in data.get("value", []):
            if e["id"].endswith(short_id):
                return e["id"]
        raise ValueError(f"Email id '{short_id}' not found in recent messages.")

    async def _reply_to_email(self, short_id: str, reply_text: str) -> str:
        full_id = await self._resolve_email_id(short_id)
        await self._graph.post(
            f"/users/{self.user_email}/messages/{full_id}/reply",
            {"comment": reply_text},
        )
        return "✅ Reply sent."

    async def _draft_email(self, to: str, subject: str, body: str) -> str:
        draft = await self._graph.post(
            f"/users/{self.user_email}/messages",
            {
                "subject": subject,
                "toRecipients": [{"emailAddress": {"address": to}}],
                "body": {"contentType": "Text", "content": body},
            },
        )
        return f"✅ Draft saved (id: {_short_id(draft.get('id', ''))}) — check your Drafts folder."

    async def _send_email(self, to: str, subject: str, body: str) -> str:
        await self._graph.post(
            f"/users/{self.user_email}/sendMail",
            {
                "message": {
                    "subject": subject,
                    "toRecipients": [{"emailAddress": {"address": to}}],
                    "body": {"contentType": "Text", "content": body},
                },
                "saveToSentItems": True,
            },
        )
        return f"✅ Email sent to {to}."

    # ------------------------------------------------------------------
    # Background polling
    # ------------------------------------------------------------------

    async def _bg_poll_loop(self):
        await asyncio.sleep(10)  # brief startup delay
        while True:
            try:
                await self._poll_tick()
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.warning(f"[m365] Poll tick error: {e}")
            await asyncio.sleep(self.poll_secs)

    async def _poll_tick(self):
        if not self._ready or not self._poll_ctx:
            return
        if self.email_read and self.user_email:
            data = await self._graph.get(
                f"/users/{self.user_email}/mailFolders/inbox/messages",
                params={
                    "$top": "5",
                    "$orderby": "receivedDateTime desc",
                    "$select": "id,subject,from,receivedDateTime,isRead",
                    "$filter": "isRead eq false",
                },
            )
            emails = data.get("value", [])
            if emails:
                first_id = emails[0]["id"]
                if first_id != self._last_email_id:
                    self._last_email_id = first_id
                    lines = ["📬 **New unread emails:**"]
                    for e in emails[:3]:
                        sender = (e.get("from") or {}).get("emailAddress", {}).get("address", "?")
                        subj   = e.get("subject", "(no subject)")
                        ts     = _fmt_dt(e.get("receivedDateTime", ""))
                        lines.append(f"• [{ts}] **{sender}** — {subj}")
                    await self._poll_ctx.send(MessageChain().message("\n".join(lines)))

    # ------------------------------------------------------------------
    # Guard helpers
    # ------------------------------------------------------------------

    def _check(self, flag: bool, feature: str) -> Optional[str]:
        if not self._ready:
            return "⚠️ M365 plugin not ready — check tenant_id, client_id, client_secret in config."
        if not flag:
            return f"⚠️ '{feature}' is disabled in plugin config."
        return None

    async def _safe_run(self, event: AstrMessageEvent, flag: bool, feature: str, coro):
        err = self._check(flag, feature)
        if err:
            yield event.plain_result(err)
            return
        try:
            result = await coro
            yield event.plain_result(result)
        except Exception as e:
            logger.exception(f"[m365] {feature} error")
            yield event.plain_result(f"❌ {feature} failed: {e}")

    # ------------------------------------------------------------------
    # Slash commands
    # ------------------------------------------------------------------

    @filter.command("m365-teams-list")
    async def cmd_teams_list(self, event: AstrMessageEvent):
        """List all accessible Teams and their channels."""
        async for r in self._safe_run(event, self.teams_read, "teams_read",
                                       self._list_teams_text()):
            yield r

    @filter.command("m365-teams-read")
    async def cmd_teams_read(self, event: AstrMessageEvent):
        """
        /m365-teams-read [team name] [channel name] [count]
        Read the last N messages from a Teams channel.
        If team/channel are omitted, the plugin defaults are used.
        """
        err = self._check(self.teams_read, "teams_read")
        if err:
            yield event.plain_result(err)
            return

        args = (event.message_str or "").split(maxsplit=3)[1:]
        team  = args[0] if len(args) > 0 else self.dflt_team
        chan  = args[1] if len(args) > 1 else self.dflt_chan
        count = int(args[2]) if len(args) > 2 and args[2].isdigit() else self.teams_max

        if not team:
            yield event.plain_result(
                "Usage: /m365-teams-read <team> <channel> [count]\n"
                "Or set watch_team_name in plugin config to use a default team."
            )
            return

        async for r in self._safe_run(event, self.teams_read, "teams_read",
                                       self._read_teams_messages(team, chan, count)):
            yield r

    @filter.command("m365-teams-send")
    async def cmd_teams_send(self, event: AstrMessageEvent):
        """
        /m365-teams-send <channel> <message>
        Send a message to the default team's channel (or prefix with team name).
        Example: /m365-teams-send General Hello team!
        """
        err = self._check(self.teams_send, "teams_send")
        if err:
            yield event.plain_result(err)
            return

        raw = (event.message_str or "").split(maxsplit=2)
        if len(raw) < 3:
            yield event.plain_result(
                "Usage: /m365-teams-send <channel> <message>\n"
                "The default team from watch_team_name is used unless you specify one."
            )
            return

        chan = raw[1]
        text = raw[2]
        team = self.dflt_team
        if not team:
            yield event.plain_result(
                "⚠️ watch_team_name is not set in config — cannot determine which team to post to."
            )
            return

        async for r in self._safe_run(event, self.teams_send, "teams_send",
                                       self._send_teams_message(team, chan, text)):
            yield r

    @filter.command("m365-email-read")
    async def cmd_email_read(self, event: AstrMessageEvent):
        """
        /m365-email-read [n]
        Show the last n emails from the configured inbox folder.
        """
        args  = (event.message_str or "").split()
        count = int(args[1]) if len(args) > 1 and args[1].isdigit() else self.email_max

        async for r in self._safe_run(event, self.email_read, "email_read",
                                       self._read_emails(count, self.inbox_folder)):
            yield r

    @filter.command("m365-email-reply")
    async def cmd_email_reply(self, event: AstrMessageEvent):
        """
        /m365-email-reply <id> <reply text>
        Reply to an email. <id> is the short 8-char ID shown by /m365-email-read.
        """
        err = self._check(self.email_send, "email_send")
        if err:
            yield event.plain_result(err)
            return

        raw = (event.message_str or "").split(maxsplit=2)
        if len(raw) < 3:
            yield event.plain_result("Usage: /m365-email-reply <id> <reply text>")
            return

        async for r in self._safe_run(event, self.email_send, "email_send",
                                       self._reply_to_email(raw[1], raw[2])):
            yield r

    @filter.command("m365-email-draft")
    async def cmd_email_draft(self, event: AstrMessageEvent):
        """
        /m365-email-draft <to> <subject> | <body>
        Save an email draft. Separate subject and body with a pipe ( | ).
        Example: /m365-email-draft alice@example.com Meeting notes | Hi Alice, ...
        """
        err = self._check(self.email_draft, "email_draft")
        if err:
            yield event.plain_result(err)
            return

        raw = (event.message_str or "").split(maxsplit=2)
        if len(raw) < 3 or "|" not in raw[2]:
            yield event.plain_result(
                "Usage: /m365-email-draft <to> <subject> | <body>\n"
                "Example: /m365-email-draft alice@example.com Meeting notes | Hi Alice, here are the notes..."
            )
            return

        to = raw[1]
        subj, _, body = raw[2].partition("|")
        async for r in self._safe_run(event, self.email_draft, "email_draft",
                                       self._draft_email(to, subj.strip(), body.strip())):
            yield r

    @filter.command("m365-email-send")
    async def cmd_email_send(self, event: AstrMessageEvent):
        """
        /m365-email-send <to> <subject> | <body>
        Send an email immediately. enable_email_send must be ON in config.
        Example: /m365-email-send alice@example.com Hello | Hi Alice, just checking in.
        """
        err = self._check(self.email_send, "email_send")
        if err:
            yield event.plain_result(err)
            return

        raw = (event.message_str or "").split(maxsplit=2)
        if len(raw) < 3 or "|" not in raw[2]:
            yield event.plain_result(
                "Usage: /m365-email-send <to> <subject> | <body>\n"
                "Example: /m365-email-send alice@example.com Hello | Hi Alice, just checking in."
            )
            return

        to = raw[1]
        subj, _, body = raw[2].partition("|")
        async for r in self._safe_run(event, self.email_send, "email_send",
                                       self._send_email(to, subj.strip(), body.strip())):
            yield r

    # ------------------------------------------------------------------
    # LLM tools (natural-language triggers)
    # ------------------------------------------------------------------

    @llm_tool(
        name="list_teams",
        desc="List all Microsoft Teams and their channels accessible to this plugin.",
    )
    async def tool_list_teams(self, event: AstrMessageEvent, **_):
        """List all teams and channels."""
        async for r in self._safe_run(event, self.teams_read, "teams_read",
                                       self._list_teams_text()):
            yield r

    @llm_tool(
        name="read_teams_messages",
        desc="Read recent messages from a Microsoft Teams channel.",
    )
    async def tool_read_teams(
        self,
        event: AstrMessageEvent,
        team_name: str,
        channel_name: str,
        count: int = 10,
    ):
        """
        Read Teams channel messages.
        :param team_name: Display name of the Team (e.g. 'Melbit Internal')
        :param channel_name: Display name of the channel (e.g. 'General')
        :param count: How many messages to return (default 10)
        """
        t = team_name or self.dflt_team
        c = channel_name or self.dflt_chan
        if not t:
            yield event.plain_result("Please specify a team name.")
            return
        async for r in self._safe_run(event, self.teams_read, "teams_read",
                                       self._read_teams_messages(t, c, count)):
            yield r

    @llm_tool(
        name="send_teams_message",
        desc="Post a message to a Microsoft Teams channel.",
    )
    async def tool_send_teams(
        self,
        event: AstrMessageEvent,
        channel_name: str,
        message: str,
        team_name: str = "",
    ):
        """
        Send a message to a Teams channel.
        :param channel_name: Target channel (e.g. 'General')
        :param message: Message text to post
        :param team_name: Team name — uses default if omitted
        """
        t = team_name or self.dflt_team
        if not t:
            yield event.plain_result("Please specify a team name or set watch_team_name in config.")
            return
        async for r in self._safe_run(event, self.teams_send, "teams_send",
                                       self._send_teams_message(t, channel_name, message)):
            yield r

    @llm_tool(
        name="read_emails",
        desc="Read recent emails from the Microsoft 365 inbox.",
    )
    async def tool_read_emails(
        self,
        event: AstrMessageEvent,
        count: int = 10,
        folder: str = "",
    ):
        """
        Read emails from inbox (or another folder).
        :param count: Number of emails to return (default 10)
        :param folder: Mail folder — defaults to the configured inbox_folder
        """
        f = folder or self.inbox_folder
        async for r in self._safe_run(event, self.email_read, "email_read",
                                       self._read_emails(count, f)):
            yield r

    @llm_tool(
        name="reply_to_email",
        desc="Reply to an email by its short ID (shown in /m365-email-read output).",
    )
    async def tool_reply_email(
        self,
        event: AstrMessageEvent,
        email_id: str,
        reply_text: str,
    ):
        """
        Reply to an email.
        :param email_id: Short 8-char email ID from email listing
        :param reply_text: The reply body text
        """
        async for r in self._safe_run(event, self.email_send, "email_send",
                                       self._reply_to_email(email_id, reply_text)):
            yield r

    @llm_tool(
        name="draft_email",
        desc="Save an email as a draft in Microsoft 365 (does not send).",
    )
    async def tool_draft_email(
        self,
        event: AstrMessageEvent,
        to: str,
        subject: str,
        body: str,
    ):
        """
        Create an email draft.
        :param to: Recipient email address
        :param subject: Email subject
        :param body: Email body (plain text)
        """
        async for r in self._safe_run(event, self.email_draft, "email_draft",
                                       self._draft_email(to, subject, body)):
            yield r

    @llm_tool(
        name="send_email",
        desc="Send an email via Microsoft 365. Only works when enable_email_send is ON in config.",
    )
    async def tool_send_email(
        self,
        event: AstrMessageEvent,
        to: str,
        subject: str,
        body: str,
    ):
        """
        Send an email.
        :param to: Recipient email address
        :param subject: Email subject
        :param body: Email body (plain text)
        """
        async for r in self._safe_run(event, self.email_send, "email_send",
                                       self._send_email(to, subject, body)):
            yield r

    # ------------------------------------------------------------------
    # Hook background poll context onto first inbound message
    # ------------------------------------------------------------------

    @filter.on_decorating_result()
    async def _capture_ctx(self, event: AstrMessageEvent):
        """Store the latest event so the background poller can send to it."""
        if self.bg_poll and self._poll_ctx is None:
            self._poll_ctx = event
