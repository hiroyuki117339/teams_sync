"""Upload exported chat.md files to a Notion database."""
import json
import os
import re
import sys
import requests
from pathlib import Path
from datetime import datetime

NOTION_API_KEY = os.environ.get("NOTION_API_KEY") or open(Path(__file__).parent / ".env").read().split("=", 1)[1].strip()
DATABASE_ID = "553158dae11d41288deca0c6efb46725"
NOTION_VERSION = "2022-06-28"
HEADERS = {
    "Authorization": f"Bearer {NOTION_API_KEY}",
    "Notion-Version": NOTION_VERSION,
    "Content-Type": "application/json",
}
API_BASE = "https://api.notion.com/v1"


def md_to_notion_blocks(md_text: str) -> list:
    """Convert Markdown text to Notion block objects."""
    blocks = []
    lines = md_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]

        # H1
        if line.startswith('# '):
            blocks.append(heading_block(line[2:].strip(), level=1))
            i += 1
            continue

        # H2
        if line.startswith('## '):
            blocks.append(heading_block(line[3:].strip(), level=2))
            i += 1
            continue

        # H3
        if line.startswith('### '):
            blocks.append(heading_block(line[4:].strip(), level=3))
            i += 1
            continue

        # Horizontal rule
        if line.strip() == '---':
            blocks.append({"object": "block", "type": "divider", "divider": {}})
            i += 1
            continue

        # Blockquote (collect consecutive > lines)
        if line.startswith('> '):
            quote_lines = []
            while i < len(lines) and lines[i].startswith('> '):
                quote_lines.append(lines[i][2:])
                i += 1
            blocks.append({
                "object": "block",
                "type": "quote",
                "quote": {"rich_text": parse_inline_md('\n'.join(quote_lines))}
            })
            continue

        # Empty line
        if line.strip() == '':
            i += 1
            continue

        # Regular paragraph (collect until empty line or special line)
        para_lines = []
        while i < len(lines) and lines[i].strip() != '' and not lines[i].startswith('#') and lines[i].strip() != '---' and not lines[i].startswith('> '):
            para_lines.append(lines[i])
            i += 1
        if para_lines:
            text = '\n'.join(para_lines)
            blocks.append(paragraph_block(text))

    return blocks


def parse_inline_md(text: str) -> list:
    """Parse inline markdown (bold, links, etc.) into Notion rich_text array."""
    rich_text = []
    # Split by bold markers and links
    # Pattern: **bold**, [text](url)
    pattern = r'(\*\*[^*]+\*\*|\[[^\]]*\]\([^)]*\))'
    parts = re.split(pattern, text)
    for part in parts:
        if not part:
            continue
        # Bold
        if part.startswith('**') and part.endswith('**'):
            content = part[2:-2]
            for chunk_start in range(0, len(content), 2000):
                chunk = content[chunk_start:chunk_start + 1900]
                rich_text.append({
                    "type": "text",
                    "text": {"content": chunk},
                    "annotations": {"bold": True}
                })
        # Link
        elif re.match(r'\[([^\]]*)\]\(([^)]*)\)', part):
            m = re.match(r'\[([^\]]*)\]\(([^)]*)\)', part)
            link_text = m.group(1)[:2000]
            url = m.group(2)
            # Validate URL: Notion rejects non-http URLs and overly long ones
            if url.startswith(('http://', 'https://')) and len(url) <= 2000:
                rich_text.append({
                    "type": "text",
                    "text": {"content": link_text, "link": {"url": url}},
                })
            else:
                # Fallback: just show as text
                rich_text.append({
                    "type": "text",
                    "text": {"content": f"{link_text} ({url[:100]}...)"}
                })
        # Plain text
        else:
            # Notion has 2000 char limit per rich_text segment
            for chunk_start in range(0, len(part), 1900):
                chunk = part[chunk_start:chunk_start + 1900]
                rich_text.append({
                    "type": "text",
                    "text": {"content": chunk}
                })
    return rich_text if rich_text else [{"type": "text", "text": {"content": " "}}]


def heading_block(text: str, level: int) -> dict:
    key = f"heading_{level}"
    return {
        "object": "block",
        "type": key,
        key: {"rich_text": parse_inline_md(text)}
    }


def paragraph_block(text: str) -> dict:
    return {
        "object": "block",
        "type": "paragraph",
        "paragraph": {"rich_text": parse_inline_md(text)}
    }


def create_page(title: str, blocks: list, export_date: str) -> dict:
    """Create a Notion page in the database."""
    # Notion API limits: max 100 blocks per request
    first_batch = blocks[:100]

    page_data = {
        "parent": {"database_id": DATABASE_ID},
        "properties": {
            "件名": {"title": [{"text": {"content": title}}]},
            "カテゴリ": {"select": {"name": "チャット"}},
            "ステータス": {"select": {"name": "完了"}},
            "登録日": {"date": {"start": export_date}},
        },
        "children": first_batch
    }

    resp = requests.post(f"{API_BASE}/pages", headers=HEADERS, json=page_data)
    if resp.status_code != 200:
        print(f"  ERROR creating page: {resp.status_code} {resp.text[:300]}", flush=True)
        return None

    page = resp.json()
    page_id = page["id"]

    # Append remaining blocks in batches of 100
    remaining = blocks[100:]
    batch_num = 1
    while remaining:
        batch = remaining[:100]
        remaining = remaining[100:]
        batch_num += 1
        resp = requests.patch(
            f"{API_BASE}/blocks/{page_id}/children",
            headers=HEADERS,
            json={"children": batch}
        )
        if resp.status_code != 200:
            print(f"  ERROR appending batch {batch_num}: {resp.status_code} {resp.text[:300]}", flush=True)
            break
        print(f"  Appended block batch {batch_num} ({len(batch)} blocks)", flush=True)

    return page


def main():
    saved_chats_dir = Path(__file__).parent / "teams-chat-exporter" / "saved_chats"

    if not saved_chats_dir.exists():
        print(f"No saved_chats directory found at {saved_chats_dir}")
        sys.exit(1)

    md_files = sorted(saved_chats_dir.glob("*/chat.md"))
    if not md_files:
        print("No chat.md files found.")
        sys.exit(1)

    print(f"Found {len(md_files)} exported chats to upload.\n", flush=True)

    for md_path in md_files:
        folder_name = md_path.parent.name
        # Extract title from folder name (remove timestamp suffix)
        # e.g. "53_被害把握ツール連携_251127_2026-02-09_211409" → "53_被害把握ツール連携_251127"
        parts = folder_name.rsplit('_', 2)
        if len(parts) >= 3 and re.match(r'\d{4}-\d{2}-\d{2}', parts[-2]):
            title = '_'.join(parts[:-2])
            export_date = parts[-2]
        else:
            title = folder_name
            export_date = datetime.now().strftime("%Y-%m-%d")

        print(f"Uploading: {title}", flush=True)

        md_text = md_path.read_text(encoding="utf-8")
        blocks = md_to_notion_blocks(md_text)

        print(f"  {len(blocks)} blocks to upload", flush=True)

        page = create_page(title, blocks, export_date)
        if page:
            print(f"  OK: {page.get('url', 'created')}", flush=True)
        else:
            print(f"  FAILED", flush=True)

        print(flush=True)

    print("Done!", flush=True)


if __name__ == "__main__":
    main()
