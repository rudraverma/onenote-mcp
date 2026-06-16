"""
Microsoft Graph API client for OneNote.
Handles OAuth 2.0 Device Authorization Grant — no browser, no redirect URI,
no client secret. Just a client_id from your Azure app registration.

Uses /sites/{siteId}/onenote/ instead of /me/onenote/ to bypass the Graph API
403 (error 10008) that occurs on accounts whose OneDrive library has >5000
OneNote items. Site ID is read from ONENOTE_SITE_ID env var (set in
~/.claude/settings.json) or auto-discovered on first use.
"""

import asyncio
import json
import os
import time
import urllib.parse
from pathlib import Path
from typing import Any, Optional

import html2text
import httpx
import markdown
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
AUTH_BASE  = "https://login.microsoftonline.com"
TOKEN_FILE = Path.home() / ".claude" / "onenote_token.json"
SCOPES     = "Notes.ReadWrite Notes.ReadWrite.All offline_access User.Read"


class OneNoteError(Exception):
    pass


class OneNoteClient:
    def __init__(self):
        self.client_id = os.getenv("ONENOTE_CLIENT_ID", "")
        self.tenant    = os.getenv("ONENOTE_TENANT", "common")
        # ONENOTE_SITE_ID bypasses /me/onenote/ (blocked when OneDrive has >5000 items).
        # Set in ~/.claude/settings.json MCP env block; auto-discovered if missing.
        self._site_id    = os.getenv("ONENOTE_SITE_ID", "")
        self._token_data: Optional[dict] = None
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)

    # ─── Site ID (routes all calls through /sites/ to avoid 10008 limit) ──────

    @property
    def _onenote_base(self) -> str:
        if self._site_id:
            return f"{GRAPH_BASE}/sites/{self._site_id}/onenote"
        return f"{GRAPH_BASE}/me/onenote"

    async def _discover_site_id(self) -> str:
        """Discover SharePoint site ID from the user's OneDrive drive URL."""
        token = await self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        async with httpx.AsyncClient(timeout=30) as c:
            # Try sharepointIds on the drive (works for OneDrive for Business)
            for path in ["/me/drive", "/me/drive/root"]:
                r = await c.get(f"{GRAPH_BASE}{path}?$select=sharepointIds", headers=headers)
                if r.status_code == 200:
                    ids     = r.json().get("sharepointIds", {})
                    host    = ids.get("siteUrl", "").split("/")[2]
                    site_id = ids.get("siteId", "")
                    web_id  = ids.get("webId", "")
                    if host and site_id and web_id:
                        return f"{host},{site_id},{web_id}"
            # Fallback: build site path from the drive's webUrl
            r = await c.get(f"{GRAPH_BASE}/me/drive?$select=webUrl", headers=headers)
            if r.status_code == 200:
                web_url = r.json().get("webUrl", "")
                if "sharepoint.com" in web_url:
                    parts    = web_url.split("/")
                    host     = parts[2]
                    rel_path = "/" + "/".join(parts[3:])
                    r2 = await c.get(
                        f"{GRAPH_BASE}/sites/{host}:{rel_path}?$select=id",
                        headers=headers,
                    )
                    if r2.status_code == 200:
                        return r2.json().get("id", "")
        return ""

    async def ensure_site_id(self) -> None:
        """Populate _site_id if not already set via env var."""
        if not self._site_id:
            self._site_id = await self._discover_site_id()

    # ─── Token management ─────────────────────────────────────────────────────

    @property
    def _token_url(self) -> str:
        return f"{AUTH_BASE}/{self.tenant}/oauth2/v2.0/token"

    @property
    def _device_code_url(self) -> str:
        return f"{AUTH_BASE}/{self.tenant}/oauth2/v2.0/devicecode"

    def _load_token(self) -> Optional[dict]:
        if TOKEN_FILE.exists():
            try:
                return json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
            except Exception:
                return None
        return None

    def _save_token(self, data: dict) -> None:
        TOKEN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def _is_expired(self, data: dict) -> bool:
        return time.time() >= (data.get("expires_at", 0) - 300)

    async def _do_refresh(self, data: dict) -> dict:
        async with httpx.AsyncClient(timeout=30) as c:
            resp = await c.post(self._token_url, data={
                "client_id":     self.client_id,
                "grant_type":    "refresh_token",
                "refresh_token": data["refresh_token"],
                "scope":         SCOPES,
            })
            resp.raise_for_status()
            new = resp.json()
            new["expires_at"] = time.time() + new.get("expires_in", 3600)
            new.setdefault("refresh_token", data["refresh_token"])
            self._save_token(new)
            return new

    async def get_access_token(self) -> str:
        if not self.client_id:
            raise OneNoteError("ONENOTE_CLIENT_ID not set. Run /onenote-setup.")
        token = self._token_data or self._load_token()
        if not token:
            raise OneNoteError("Not authenticated. Run onenote_authenticate() first.")
        if self._is_expired(token):
            token = await self._do_refresh(token)
        self._token_data = token
        return token["access_token"]

    def get_token_status(self) -> dict:
        token = self._load_token()
        if not token:
            return {"authenticated": False, "reason": "No token — run onenote_authenticate()"}
        if self._is_expired(token):
            return {
                "authenticated":    False,
                "reason":           "Token expired — will auto-refresh on next API call",
                "has_refresh_token": "refresh_token" in token,
            }
        return {
            "authenticated":    True,
            "expires_at":       time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(token.get("expires_at", 0))),
            "has_refresh_token": "refresh_token" in token,
        }

    # ─── Device flow ──────────────────────────────────────────────────────────

    async def start_device_flow(self) -> dict:
        if not self.client_id:
            raise OneNoteError("ONENOTE_CLIENT_ID not set.")
        async with httpx.AsyncClient(timeout=30) as c:
            resp = await c.post(self._device_code_url, data={
                "client_id": self.client_id,
                "scope":     SCOPES,
            })
            resp.raise_for_status()
            return resp.json()

    async def poll_device_flow(self, device_code: str, interval: int = 5, timeout: int = 300) -> dict:
        deadline = time.time() + timeout
        async with httpx.AsyncClient(timeout=30) as c:
            while time.time() < deadline:
                await asyncio.sleep(interval)
                resp = await c.post(self._token_url, data={
                    "client_id":  self.client_id,
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "device_code": device_code,
                })
                data  = resp.json()
                error = data.get("error", "")
                if error == "authorization_pending":
                    continue
                elif error == "slow_down":
                    interval = min(interval + 5, 30)
                    continue
                elif error:
                    raise OneNoteError(f"Auth failed: {data.get('error_description', error)}")
                data["expires_at"] = time.time() + data.get("expires_in", 3600)
                self._token_data   = data
                self._save_token(data)
                return data
        raise OneNoteError("Authentication timed out after 5 minutes.")

    # ─── HTTP helpers ─────────────────────────────────────────────────────────

    async def _request(self, method: str, path: str, **kwargs) -> Any:
        token = await self.get_access_token()
        url  = path if path.startswith("https://") else f"{GRAPH_BASE}{path}"
        hdrs = {"Authorization": f"Bearer {token}"}
        if "headers" in kwargs:
            hdrs.update(kwargs.pop("headers"))

        last_status, last_text = 0, ""
        async with httpx.AsyncClient(timeout=60) as c:
            for attempt in range(6):
                r = await getattr(c, method)(url, headers=hdrs, **kwargs)
                last_status, last_text = r.status_code, r.text

                if r.status_code in (200, 201, 204):
                    return {} if r.status_code == 204 else r.json()

                if r.status_code in (429, 503):
                    retry_after = int(r.headers.get("Retry-After", 0)) or int(2 ** attempt * 2)
                    await asyncio.sleep(retry_after)
                    continue

                if r.status_code == 409:
                    try:
                        err_code = r.json().get("error", {}).get("code", "")
                    except Exception:
                        err_code = ""
                    if err_code == "30103":
                        # OneNote write throttle — fail fast so the per-file error
                        # handler catches it. Sleeping 30s here would blow the MCP timeout.
                        raise OneNoteError("30103: OneNote write throttle — run sync again to retry")
                    raise OneNoteError(f"Graph API {r.status_code}: {r.text[:500]}")

                raise OneNoteError(f"Graph API {r.status_code}: {r.text[:500]}")

        raise OneNoteError(f"Graph API {last_status} after 6 attempts: {last_text[:500]}")

    async def get(self, path: str) -> Any:
        return await self._request("get", path)

    async def post_json(self, path: str, data: dict) -> Any:
        return await self._request("post", path, json=data)

    async def post_html(self, path: str, html_bytes: bytes) -> Any:
        return await self._request(
            "post", path,
            content=html_bytes,
            headers={"Content-Type": "text/html; charset=utf-8"},
        )

    async def patch_page(self, path: str, commands: list) -> Any:
        return await self._request(
            "patch", path,
            content=json.dumps(commands).encode(),
            headers={"Content-Type": "application/json"},
        )

    async def get_raw(self, path: str) -> bytes:
        token = await self.get_access_token()
        url   = path if path.startswith("https://") else f"{GRAPH_BASE}{path}"
        async with httpx.AsyncClient(timeout=60) as c:
            r = await c.get(url, headers={"Authorization": f"Bearer {token}"})
            if r.status_code not in (200, 201):
                raise OneNoteError(f"Graph API {r.status_code}: {r.text[:300]}")
            return r.content

    # ─── Format conversion ────────────────────────────────────────────────────

    @staticmethod
    def md_to_onenote_html(title: str, md_content: str) -> bytes:
        html_body = markdown.markdown(
            md_content,
            extensions=["tables", "fenced_code", "nl2br"],
        )
        html = (
            "<!DOCTYPE html>\n<html>\n"
            f"<head><title>{title}</title><meta charset='utf-8'/></head>\n"
            f"<body>\n{html_body}\n</body>\n</html>"
        )
        return html.encode("utf-8")

    @staticmethod
    def onenote_html_to_md(html_content: str | bytes) -> str:
        if isinstance(html_content, bytes):
            html_content = html_content.decode("utf-8", errors="replace")
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.body_width   = 0
        h.protect_links = True
        h.wrap_links   = False
        return h.handle(html_content).strip()

    # ─── OneNote CRUD helpers (all use _onenote_base) ─────────────────────────

    async def list_notebooks(self) -> list[dict]:
        await self.ensure_site_id()
        data = await self.get(f"{self._onenote_base}/notebooks?$select=id,displayName,lastModifiedDateTime")
        return data.get("value", [])

    async def get_or_create_notebook(self, name: str, cached_id: str = "") -> dict:
        """Return notebook dict. Uses cached_id fast-path; falls back to filter then create."""
        await self.ensure_site_id()
        base  = self._onenote_base
        token = await self.get_access_token()
        hdrs  = {"Authorization": f"Bearer {token}"}

        if cached_id:
            async with httpx.AsyncClient(timeout=30) as c:
                r = await c.get(f"{base}/notebooks/{cached_id}", headers=hdrs)
                if r.status_code == 200:
                    return r.json()

        async with httpx.AsyncClient(timeout=30) as c:
            fq = urllib.parse.quote(f"displayName eq '{name}'")
            r  = await c.get(f"{base}/notebooks?$filter={fq}", headers=hdrs)
            if r.status_code == 200:
                hits = r.json().get("value", [])
                if hits:
                    return hits[0]
            r = await c.post(
                f"{base}/notebooks",
                headers={**hdrs, "Content-Type": "application/json"},
                json={"displayName": name},
            )
            if r.status_code in (200, 201):
                return r.json()
            if r.status_code == 409:
                # Notebook already exists — list all and find by name.
                r2 = await c.get(f"{base}/notebooks?$select=id,displayName", headers=hdrs)
                if r2.status_code == 200:
                    for nb in r2.json().get("value", []):
                        if nb.get("displayName") == name:
                            return nb
            raise OneNoteError(f"Could not get or create notebook '{name}': {r.status_code} {r.text[:300]}")

    async def list_sections(self, notebook_id: str) -> list[dict]:
        await self.ensure_site_id()
        data = await self.get(f"{self._onenote_base}/notebooks/{notebook_id}/sections?$select=id,displayName")
        return data.get("value", [])

    async def get_or_create_section(self, notebook_id: str, name: str, cached_id: str = "") -> dict:
        """Return section dict. Uses cached_id fast-path; creates if missing; handles 409."""
        await self.ensure_site_id()
        base  = self._onenote_base
        token = await self.get_access_token()
        hdrs  = {"Authorization": f"Bearer {token}"}

        if cached_id:
            async with httpx.AsyncClient(timeout=30) as c:
                r = await c.get(f"{base}/sections/{cached_id}", headers=hdrs)
                if r.status_code == 200:
                    return r.json()

        async with httpx.AsyncClient(timeout=30) as c:
            r = await c.post(
                f"{base}/notebooks/{notebook_id}/sections",
                headers={**hdrs, "Content-Type": "application/json"},
                json={"displayName": name},
            )
            if r.status_code in (200, 201):
                return r.json()
            if r.status_code == 409:
                # Section already exists — list all sections and find by name.
                # More reliable than OData $filter which can fail on special chars.
                r2 = await c.get(
                    f"{base}/notebooks/{notebook_id}/sections?$select=id,displayName",
                    headers=hdrs,
                )
                if r2.status_code == 200:
                    for sec in r2.json().get("value", []):
                        if sec.get("displayName") == name:
                            return sec
            raise OneNoteError(f"Could not get or create section '{name}': {r.status_code} {r.text[:300]}")

    async def list_pages(self, section_id: str) -> list[dict]:
        await self.ensure_site_id()
        data = await self.get(
            f"{self._onenote_base}/sections/{section_id}/pages"
            "?$select=id,title,lastModifiedDateTime&$orderby=title"
        )
        return data.get("value", [])

    async def get_page_content_md(self, page_id: str) -> str:
        await self.ensure_site_id()
        html_bytes = await self.get_raw(f"{self._onenote_base}/pages/{page_id}/content")
        return self.onenote_html_to_md(html_bytes)

    async def push_file_to_section(self, section_id: str, title: str, md_content: str, config: dict) -> dict:
        """Create or replace a page. Updates config['file_to_page'] in place."""
        await self.ensure_site_id()
        base      = self._onenote_base
        html_bytes = self.md_to_onenote_html(title, md_content)
        file_map  = config.setdefault("file_to_page", {})
        page_id   = file_map.get(title)

        if page_id:
            html_str = html_bytes.decode("utf-8")
            await self.patch_page(
                f"{base}/pages/{page_id}/content",
                [{"target": "body", "action": "replace", "content": f"<div>{html_str}</div>"}],
            )
            return {"action": "updated", "page_id": page_id, "title": title}

        page   = await self.post_html(f"{base}/sections/{section_id}/pages", html_bytes)
        new_id = page.get("id", "")
        file_map[title] = new_id
        return {"action": "created", "page_id": new_id, "title": title}
