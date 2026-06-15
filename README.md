<div align="center">

# **CyberHawk Threat Intel**

<a href="https://www.cyberhawkthreatintel.com">
  <img src="https://media.cyberhawkthreatintel.com/general/1771234479938-y9566.png" width="140" alt="CyberHawk Threat Intel"/>
</a>

## onenote-mcp

**Bidirectional sync between Claude Code projects and Microsoft OneNote**

*Built by Rudra Verma | Senior Cyber Security Architect & Researcher | [CyberHawk Threat Intel](https://www.cyberhawkthreatintel.com)*

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![MCP 1.0+](https://img.shields.io/badge/MCP-1.0%2B-purple)](https://modelcontextprotocol.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)
[![CyberHawk Threat Intel](https://img.shields.io/badge/by-CyberHawk%20Threat%20Intel-0066cc)](https://www.cyberhawkthreatintel.com)

*Push your `CLAUDE.md`, `MEMORY.md`, and project docs to OneNote with a single command. Pull them back. No manual copy-paste, no format headaches.*

</div>

---

## What it does

Every Claude Code project has files Claude reads and updates during sessions — `CLAUDE.md`, `docs/MEMORY.md`, `docs/ARCHITECTURE.md`, and others. This MCP server syncs those files into a dedicated OneNote section so you can view, search, and reference them from any device.

```
Your project/                  OneNote
├── CLAUDE.md          →       Claude Code Projects/
├── docs/                      └── my-project/
│   ├── MEMORY.md      →           ├── CLAUDE.md
│   ├── ARCHITECTURE.md →          ├── docs/MEMORY.md
│   └── WORKFLOWS.md   →          ├── docs/ARCHITECTURE.md
└── .claude/                       └── .claude/settings.json
    └── settings.json  →
```

Subsequent pushes do direct PATCH updates via cached page IDs — no searching, no duplicates.

---

## Tools (13 total)

| Tool | What it does |
|---|---|
| `onenote_authenticate` | One-time device flow login — shows a short code, you paste it at microsoft.com/devicelogin |
| `onenote_token_status` | Check auth status and token expiry |
| `onenote_list_notebooks` | List all accessible notebooks |
| `onenote_get_or_create_notebook` | Get notebook by name or create it |
| `onenote_list_sections` | List sections in a notebook |
| `onenote_get_or_create_section` | Get section by name or create it |
| `onenote_list_pages` | List all pages in a section |
| `onenote_get_page_content` | Read a page and return it as markdown |
| `onenote_push_file` | Push a single markdown string as a OneNote page |
| `onenote_pull_page_to_file` | Pull a page and write it to a local file |
| `onenote_sync_project_to_onenote` | **Push entire project** — all markdown files in one call |
| `onenote_pull_project_from_onenote` | **Pull entire section** back to local files |
| `onenote_project_status` | Show sync config, last push time, page count |

---

## Installation

### 1. Install the package

```bash
git clone https://github.com/rudraverma/onenote-mcp.git
cd onenote-mcp
pip install -e .
```

### 2. Add to Claude Code

Edit `~/.claude/settings.json`:

**Windows** (use the full path to the exe):
```json
{
  "mcpServers": {
    "onenote-mcp": {
      "command": "C:\\Users\\YourName\\AppData\\Roaming\\Python\\Python313\\Scripts\\onenote-mcp.exe",
      "env": {
        "ONENOTE_CLIENT_ID": "your-azure-app-client-id"
      }
    }
  }
}
```

**macOS / Linux:**
```json
{
  "mcpServers": {
    "onenote-mcp": {
      "command": "onenote-mcp",
      "env": {
        "ONENOTE_CLIENT_ID": "your-azure-app-client-id"
      }
    }
  }
}
```

> Find the exe path on Windows: `where onenote-mcp` in PowerShell after install.

---

## Azure App Registration (one-time setup, ~10 minutes, free)

You need to register a free Azure app so Microsoft can authenticate you. No subscription required.

### Step 1 — Open Azure App Registrations

Go to: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps

Sign in with the **same Microsoft account** that has your OneNote notebooks.

### Step 2 — Create a new app registration

1. Click **+ New registration**
2. **Name:** `onenote-mcp` (or any name you like)
3. **Supported account types** — choose based on your account type:
   - Personal Microsoft account (Outlook.com, Hotmail, Live.com) → **"Personal Microsoft accounts only"**
   - Work / school Microsoft 365 account → **"Accounts in this organizational directory only"**
   - Both → **"Accounts in any organizational directory and personal Microsoft accounts"**
4. **Redirect URI:** leave blank — not needed for device flow
5. Click **Register**

### Step 3 — Add API permissions

1. Click **API permissions** in the left sidebar
2. Click **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Search for and select these 4 permissions:

| Permission | Purpose |
|---|---|
| `Notes.ReadWrite` | Read and write your personal OneNote notebooks |
| `Notes.ReadWrite.All` | Required for work/school accounts and shared notebooks |
| `offline_access` | Keeps you signed in via refresh tokens |
| `User.Read` | Basic Graph API access (usually pre-added) |

4. Click **Add permissions**
5. **Work/school accounts only:** click **Grant admin consent for [your org]** → Yes

> `Notes.ReadWrite.All` is required even for personal notebooks when accessed via the
> SharePoint sites endpoint (needed for accounts with large OneDrive libraries — see below).

### Step 4 — Enable public client flows (required for device flow)

1. Click **Authentication** in the left sidebar
2. Scroll to **Advanced settings**
3. Set **Allow public client flows** to **Yes**
4. Click **Save**

> Without this you'll get: `AADSTS7000218: The request body must contain 'client_assertion' or 'client_secret'`

### Step 5 — Copy your Client ID

1. Click **Overview** in the left sidebar
2. Copy **Application (client) ID** — format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

### Step 6 — Configure env vars

**Minimum config (personal Microsoft account):**
```json
"env": {
  "ONENOTE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

**Work / school Microsoft 365 account:**
```json
"env": {
  "ONENOTE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "ONENOTE_TENANT": "your-tenant-id"
}
```

Find your Tenant ID: Azure portal → **Azure Active Directory** → **Overview** → **Tenant ID**

### Step 7 — First-time authentication

Restart Claude Code so the MCP server loads, then run:

```
onenote_authenticate()
```

You'll see:
```
Go to: https://microsoft.com/devicelogin
Enter code: ABCD1234
(Code expires in 15 minutes — waiting...)
```

Open the URL on any device, sign in, enter the code, click **Allow**.
Token is saved to `~/.claude/onenote_token.json` and auto-refreshes — you never need to auth again.

---

## Enterprise / Large OneDrive Accounts

> If your account's OneDrive library has more than **5,000 OneNote items**, Microsoft Graph
> blocks all `/me/onenote/` API calls with error 10008. This affects many enterprise M365 accounts.

This MCP handles it automatically using the `/sites/{siteId}/onenote/` endpoint instead.
You just need to provide your SharePoint site ID:

### How to find your Site ID

1. Go to your OneDrive in a browser (e.g. `https://yourorg-my.sharepoint.com/personal/yourname_org_com`)
2. Open browser DevTools → Network tab
3. Navigate around and look for Graph API calls, or use:

```powershell
# PowerShell — run after authenticating once to get your token
$token = (Get-Content "~/.claude/onenote_token.json" | ConvertFrom-Json).access_token
$r = Invoke-RestMethod "https://graph.microsoft.com/v1.0/me/drive?`$select=sharepointIds" -Headers @{Authorization="Bearer $token"}
$ids = $r.sharepointIds
"$($ids.siteUrl.split('/')[2]),$($ids.siteId),$($ids.webId)"
```

The output is your site ID in the format `hostname,siteId,webId`.

### Add it to your MCP config

```json
"env": {
  "ONENOTE_CLIENT_ID": "your-client-id",
  "ONENOTE_TENANT": "your-tenant-id",
  "ONENOTE_SITE_ID": "yourorg-my.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

When `ONENOTE_SITE_ID` is set, all API calls route through the SharePoint sites endpoint
and the 5,000-item limit is bypassed. This is the recommended config for all enterprise accounts.

---

## Usage

### Push a project to OneNote

```
onenote_sync_project_to_onenote(
  project_path="D:/my-project",
  notebook_name="Claude Code Projects"
)
```

Syncs all `*.md` files in the root and `docs/` folder, plus `.claude/settings.json`.
Creates the notebook and section on first run; reuses cached IDs on subsequent pushes.

### Pull from OneNote back to local

```
onenote_pull_project_from_onenote(project_path="D:/my-project")
```

### Check sync status

```
onenote_project_status(project_path="D:/my-project")
```

### Typical session workflow

```
# Start of session — pull latest from OneNote
onenote_pull_project_from_onenote(project_path="D:/my-project")

# ... do your work, Claude updates local files ...

# End of session — push back to OneNote
onenote_sync_project_to_onenote(project_path="D:/my-project")
```

---

## Configuration reference

| Env var | Required | Description |
|---|---|---|
| `ONENOTE_CLIENT_ID` | Yes | Azure app Client ID |
| `ONENOTE_TENANT` | For work accounts | Tenant ID or `common` for personal accounts |
| `ONENOTE_SITE_ID` | For enterprise (>5k items) | SharePoint site ID in `host,siteId,webId` format |

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `AADSTS7000218` | Public client not enabled | Azure portal → Authentication → Enable "Allow public client flows" |
| `AADSTS65001` | Missing admin consent (work accounts) | Azure portal → API permissions → Grant admin consent |
| `AADSTS70011: scope notes` | Notes permissions missing | Add `Notes.ReadWrite` and `Notes.ReadWrite.All` |
| `AADSTS700016` | Wrong Client ID | Re-copy from Azure portal → Overview → Application (client) ID |
| `AADSTS50059` | Wrong tenant | Use actual Tenant ID, not `common`, for work accounts |
| `Graph 403 / error 10008` | OneDrive has >5000 OneNote items | Set `ONENOTE_SITE_ID` env var (see Enterprise section above) |
| `Graph 403 / error 40004` | Missing `Notes.ReadWrite.All` | Add this permission in Azure portal and re-authenticate |
| MCP server not found (Windows) | Bare command name doesn't resolve | Use full `.exe` path in settings.json (see Installation) |

---

## Design

- **Local files are the source of truth.** OneNote is a sync target, not a database.
- **Device Authorization Grant** — no redirect URI, no client secret, works from any CLI or desktop app.
- **Token stored globally** at `~/.claude/onenote_token.json` — one auth covers all projects.
- **Page mapping cached** in `.claude/onenote_config.json` per project — direct PATCH on every re-push, no duplicates.
- **Markdown → HTML** via the `markdown` Python library; **HTML → Markdown** via `html2text`.
- **SharePoint sites endpoint** used by default when `ONENOTE_SITE_ID` is set — bypasses the 5,000-item OneDrive limit that blocks `/me/onenote/` on large enterprise accounts.

---

## Requirements

- Python 3.10+
- Microsoft account (personal Outlook / Hotmail / Live, or work/school M365)
- Azure app registration (free, no Azure subscription required)

---

<div align="center">

<a href="https://www.cyberhawkthreatintel.com">
  <img src="https://media.cyberhawkthreatintel.com/general/1771234479938-y9566.png" alt="CyberHawk Threat Intel" width="80"/>
</a>

**[CyberHawk Threat Intel](https://www.cyberhawkthreatintel.com)**

*Rudra Verma | Senior Cyber Security Architect & Researcher | CyberHawk Threat Intel*

[YouTube @cyberhawkconsultancy](https://youtube.com/@cyberhawkconsultancy) · [YouTube @cyberhawkk](https://youtube.com/@cyberhawkk) · [TikTok](https://tiktok.com/@cyberhawkthreatintel) · [X @cyberhawkintel](https://x.com/cyberhawkintel) · [Telegram](https://t.me/cyberhawkthreatintel)

*Authorized security research & penetration testing only. Unauthorized use is illegal.*

`#cyberhawkthreatintel` `#cyberhawkconsultancy` `#cyberhawkk` `#cybersecurity` `#ethicalhacking` `#pentesting` `#redteam` `#threatintel` `#infosec` `#claudecode` `#mcp`

</div>
