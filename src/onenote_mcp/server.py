"""
OneNote MCP Server — bidirectional sync between Claude Code projects and Microsoft OneNote.
13 tools covering auth, notebooks, sections, pages, and full project push/pull.
"""

import asyncio
import datetime
import json
import os
import time
from pathlib import Path

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

from .client import OneNoteClient, OneNoteError

load_dotenv()

mcp = FastMCP(
    "onenote-mcp",
    instructions=(
        "Bidirectional sync between Claude Code project files and Microsoft OneNote. "
        "Push CLAUDE.md, MEMORY.md, and docs/ to OneNote sections. Pull changes back. "
        "Requires one-time Azure app setup — run onenote_authenticate() first."
    ),
)

CONFIG_FILE = ".claude/onenote_config.json"


def _c() -> OneNoteClient:
    return OneNoteClient()


def _fmt(data) -> str:
    return json.dumps(data, indent=2, default=str)


def _load_config(project_path: str) -> dict:
    cfg_path = Path(project_path) / CONFIG_FILE
    if cfg_path.exists():
        return json.loads(cfg_path.read_text(encoding="utf-8"))
    return {}


def _save_config(project_path: str, config: dict) -> None:
    cfg_path = Path(project_path) / CONFIG_FILE
    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    cfg_path.write_text(json.dumps(config, indent=2), encoding="utf-8")


def _project_memory_dir(project_path: str) -> Path:
    """Resolve ~/.claude/projects/<slug>/memory/ for the given project path."""
    import re
    slug = re.sub(r"[^a-zA-Z0-9-]", "-", project_path)
    return Path.home() / ".claude" / "projects" / slug / "memory"


def _collect_project_files(project_path: str, since: float | None = None) -> list[tuple[str, str]]:
    """Return (page_title, file_content) for syncable project files.
    If since is an epoch timestamp, only include files with mtime > since."""
    root  = Path(project_path)
    files = []
    for f in sorted(root.glob("*.md")):
        if since is None or f.stat().st_mtime > since:
            files.append((f.name, f.read_text(encoding="utf-8")))
    docs = root / "docs"
    if docs.is_dir():
        for f in sorted(docs.glob("*.md")):
            if since is None or f.stat().st_mtime > since:
                files.append((f"docs/{f.name}", f.read_text(encoding="utf-8")))
    settings = root / ".claude" / "settings.json"
    if settings.exists() and (since is None or settings.stat().st_mtime > since):
        content = f"```json\n{settings.read_text(encoding='utf-8')}\n```"
        files.append((".claude/settings.json", content))
    # Memory files — critical for project context recovery
    memory_dir = _project_memory_dir(project_path)
    if memory_dir.is_dir():
        for f in sorted(memory_dir.glob("*.md")):
            if since is None or f.stat().st_mtime > since:
                files.append((f"memory/{f.name}", f.read_text(encoding="utf-8")))
    return files


# ═══════════════════════════════════════════════════════════════════════════════
# AUTH
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
async def onenote_authenticate() -> str:
    """
    Authenticate with Microsoft OneNote using device flow (no browser redirect needed).
    Shows a short code — paste it at microsoft.com/devicelogin on any device.
    Run once; token is saved globally and auto-refreshes.
    """
    client = _c()
    if not client.client_id:
        return (
            "ONENOTE_CLIENT_ID is not set.\n\n"
            "1. Register an Azure app at https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps\n"
            "2. Copy the Application (client) ID\n"
            "3. Add it to the MCP env in ~/.claude/settings.json:\n"
            '   "ONENOTE_CLIENT_ID": "your-id-here"\n\n'
            "Full guide: skills/onenote-sync/references/azure-setup.md"
        )
    flow        = await client.start_device_flow()
    user_code   = flow["user_code"]
    verify_url  = flow["verification_uri"]
    expires     = flow.get("expires_in", 900)
    device_code = flow["device_code"]
    interval    = flow.get("interval", 5)

    print(f"Go to: {verify_url}\nEnter code: {user_code}\n(Expires in {expires // 60} min) — waiting...")

    try:
        token = await client.poll_device_flow(device_code, interval=interval, timeout=expires)
        return (
            f"Authenticated.\n"
            f"Token saved to ~/.claude/onenote_token.json\n"
            f"Expires in {token.get('expires_in', 3600) // 3600}h (auto-refreshes)"
        )
    except OneNoteError as e:
        return f"Authentication failed: {e}"


@mcp.tool()
async def onenote_token_status() -> str:
    """Check whether a valid OneNote token exists and when it expires."""
    return _fmt(_c().get_token_status())


# ═══════════════════════════════════════════════════════════════════════════════
# NOTEBOOKS
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
async def onenote_list_notebooks() -> str:
    """List all OneNote notebooks accessible to the authenticated user."""
    try:
        client    = _c()
        notebooks = await client.list_notebooks()
        return _fmt([
            {"id": nb["id"], "name": nb["displayName"], "modified": nb.get("lastModifiedDateTime", "")}
            for nb in notebooks
        ])
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_get_or_create_notebook(name: str = "Claude Code Projects") -> str:
    """
    Get a notebook by name or create it if it doesn't exist.
    Returns the notebook ID needed for section operations.
    """
    try:
        nb = await _c().get_or_create_notebook(name)
        return _fmt({"id": nb["id"], "name": nb["displayName"]})
    except OneNoteError as e:
        return f"ERROR: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
# SECTIONS
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
async def onenote_list_sections(notebook_id: str) -> str:
    """List all sections inside a notebook."""
    try:
        client   = _c()
        sections = await client.list_sections(notebook_id)
        return _fmt([{"id": s["id"], "name": s["displayName"]} for s in sections])
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_get_or_create_section(notebook_id: str, section_name: str) -> str:
    """
    Get a section by name inside a notebook, or create it.
    Returns section ID needed for page operations.
    """
    try:
        sec = await _c().get_or_create_section(notebook_id, section_name)
        return _fmt({"id": sec["id"], "name": sec["displayName"]})
    except OneNoteError as e:
        return f"ERROR: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
# PAGES
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
async def onenote_list_pages(section_id: str) -> str:
    """List all pages in a OneNote section."""
    try:
        client = _c()
        pages  = await client.list_pages(section_id)
        return _fmt([
            {"id": p["id"], "title": p.get("title", ""), "modified": p.get("lastModifiedDateTime", "")}
            for p in pages
        ])
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_get_page_content(page_id: str) -> str:
    """
    Get a OneNote page's content and return it as markdown.
    Use this to read a page back into Claude's context.
    """
    try:
        return await _c().get_page_content_md(page_id)
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_push_file(
    section_id: str,
    title: str,
    markdown_content: str,
    project_path: str = "",
) -> str:
    """
    Push a markdown string to OneNote as a page with the given title.
    Creates a new page if it doesn't exist; replaces content if it does.
    Optionally pass project_path to update the local onenote_config.json mapping.
    """
    try:
        client = _c()
        config = _load_config(project_path) if project_path else {}
        result = await client.push_file_to_section(section_id, title, markdown_content, config)
        if project_path:
            _save_config(project_path, config)
        return _fmt(result)
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_pull_page_to_file(page_id: str, output_path: str) -> str:
    """
    Pull a OneNote page and write it as a markdown file at output_path.
    Creates parent directories if needed.
    """
    try:
        client = _c()
        md     = await client.get_page_content_md(page_id)
        out    = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(md, encoding="utf-8")
        return f"Written: {output_path} ({len(md)} chars)"
    except OneNoteError as e:
        return f"ERROR: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
# PROJECT SYNC
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
async def onenote_sync_project_to_onenote(
    project_path: str,
    notebook_name: str = "Claude Code Projects",
    section_name: str = "",
    force: bool = False,
) -> str:
    """
    Push changed project files (CLAUDE.md, docs/*.md, .claude/settings.json, memory/*.md)
    to OneNote. On repeated calls only pushes files modified since the last successful sync.
    Pass force=True to push all files regardless of modification time.
    Creates the notebook and section on first run; reuses cached IDs on subsequent runs.
    Config saved to .claude/onenote_config.json.
    """
    try:
        root = Path(project_path).resolve()
        if not root.is_dir():
            return f"ERROR: {project_path} is not a directory"

        effective_section = section_name or root.name
        client = _c()
        config = _load_config(str(root))

        nb  = await client.get_or_create_notebook(notebook_name, cached_id=config.get("notebook_id", ""))
        sec = await client.get_or_create_section(nb["id"], effective_section, cached_id=config.get("section_id", ""))

        config["notebook_id"]   = nb["id"]
        config["notebook_name"] = nb["displayName"]
        config["section_id"]    = sec["id"]
        config["section_name"]  = sec["displayName"]

        file_map = config.get("file_to_page", {})
        last_push_str = config.get("last_push", "")

        if force:
            # force=True: collect all files regardless of mtime
            files = _collect_project_files(str(root), since=None)
        else:
            # Default: push (a) files modified since last successful sync
            #          + (b) files never yet pushed (no page_id), regardless of mtime.
            # This lets the initial sync converge quickly without re-PATCHing everything.
            since: float | None = None
            if last_push_str:
                try:
                    dt = datetime.datetime.fromisoformat(last_push_str.rstrip("Z"))
                    since = dt.replace(tzinfo=datetime.timezone.utc).timestamp()
                except Exception:
                    since = None
            changed = _collect_project_files(str(root), since=since)
            all_files = _collect_project_files(str(root), since=None)
            never_pushed = [(t, c) for t, c in all_files if t not in file_map]
            # Merge: never-pushed first, then changed (deduped)
            seen: set[str] = set()
            files = []
            for t, c in (never_pushed + changed):
                if t not in seen:
                    seen.add(t)
                    files.append((t, c))

        if not files:
            return (
                f"Nothing to sync — no files changed since last push ({last_push_str})\n"
                f"Pass force=True to push all files regardless of mtime."
            )

        # Sort never-pushed files (no page_id) first so they always complete
        # before the time budget runs out, even when there are many existing pages.
        files.sort(key=lambda ft: (ft[0] in file_map, ft[0]))

        # 38s hard budget — safely fires before the ~45s MCP transport timeout.
        # Semaphore(1) fully serializes writes — OneNote returns 30103 on concurrent
        # writes to the same section.
        # POSTs (new pages) sleep 0.5s — they trigger 30103 most easily.
        # PATCHes (updates) sleep 0.1s — less throttle risk, saves ~6s on bulk syncs.
        deadline = time.monotonic() + 38
        sem = asyncio.Semaphore(1)

        async def push_one(title: str, content: str) -> dict:
            async with sem:
                if time.monotonic() > deadline:
                    return {"action": "skipped", "title": title}
                is_new = title not in file_map
                await asyncio.sleep(0.5 if is_new else 0.1)
                try:
                    return await client.push_file_to_section(sec["id"], title, content, config)
                except OneNoteError as e:
                    return {"action": "error", "title": title, "error": str(e)}

        results = list(await asyncio.gather(*[push_one(t, c) for t, c in files]))

        pushed  = sum(1 for r in results if r["action"] == "created")
        updated = sum(1 for r in results if r["action"] == "updated")
        skipped = [r for r in results if r["action"] == "skipped"]
        errors  = [r for r in results if r["action"] == "error"]

        # Only advance last_push when everything completed — ensures skipped/failed
        # files are retried on the next call (their mtime will still be > last_push).
        if not skipped and not errors:
            config["last_push"] = datetime.datetime.utcnow().isoformat() + "Z"
        _save_config(str(root), config)

        lines = [f"  {r['action']:8} {r['title']}" for r in results if r["action"] not in ("error", "skipped")]
        if errors:
            lines += [f"  ERROR    {r['title']}: {r.get('error', '')}" for r in errors]
        if skipped:
            lines += [f"  SKIPPED  {r['title']} (time budget — run sync again)" for r in skipped]

        summary = f"Sync: {pushed} created, {updated} updated"
        if errors:
            summary += f", {len(errors)} failed"
        if skipped:
            summary += f", {len(skipped)} skipped — run again to finish"
        return (
            f"{summary}\n"
            f"Notebook: {notebook_name}  Section: {effective_section}\n\n"
            + "\n".join(lines)
        )
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_pull_project_from_onenote(
    project_path: str,
    section_id: str = "",
) -> str:
    """
    Pull all pages from a OneNote section back to local markdown files.
    section_id: optional — reads from .claude/onenote_config.json if omitted.
    Page titles become relative file paths (e.g. 'docs/MEMORY.md').
    """
    try:
        root   = Path(project_path).resolve()
        config = _load_config(str(root))
        sid    = section_id or config.get("section_id", "")
        if not sid:
            return (
                "ERROR: No section_id provided and no config found.\n"
                "Run onenote_sync_project_to_onenote() first, or pass section_id explicitly."
            )

        client = _c()
        pages  = await client.list_pages(sid)
        if not pages:
            return "No pages found in section."

        results = []
        for page in pages:
            title = page.get("title", page["id"])
            out   = root / title
            try:
                md = await client.get_page_content_md(page["id"])
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_text(md, encoding="utf-8")
                results.append(f"  pulled  {title}")
            except Exception as ex:
                results.append(f"  ERROR   {title}: {ex}")

        return f"Pull complete: {len(results)} files\n\n" + "\n".join(results)
    except OneNoteError as e:
        return f"ERROR: {e}"


@mcp.tool()
async def onenote_project_status(project_path: str) -> str:
    """
    Show the current OneNote sync status for a project.
    Reads .claude/onenote_config.json and verifies the section is reachable.
    """
    try:
        config = _load_config(project_path)
        if not config:
            return "No OneNote config found. Run onenote_sync_project_to_onenote() first."

        client     = _c()
        auth       = client.get_token_status()
        section_id = config.get("section_id", "")
        page_count = 0
        if section_id and auth.get("authenticated"):
            try:
                pages      = await client.list_pages(section_id)
                page_count = len(pages)
            except Exception:
                pass

        return _fmt({
            "auth":            auth,
            "notebook":        config.get("notebook_name", ""),
            "section":         config.get("section_name", ""),
            "last_push":       config.get("last_push", "never"),
            "pages_in_onenote": page_count,
            "files_mapped":    len(config.get("file_to_page", {})),
        })
    except OneNoteError as e:
        return f"ERROR: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    mcp.run()


if __name__ == "__main__":
    main()
