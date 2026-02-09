"""Inspect channel message DOM structure in detail."""
import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

PROJECT_ROOT = Path(__file__).parent.parent
USER_DATA_DIR = PROJECT_ROOT / "playwright_user_data"

async def main():
    async with async_playwright() as p:
        browser_context = await p.chromium.launch_persistent_context(USER_DATA_DIR, headless=False)
        page = browser_context.pages[0] if browser_context.pages else await browser_context.new_page()

        if "teams.microsoft.com" not in page.url:
            await page.goto("https://teams.microsoft.com/")

        print("\n=== Teams チャネルのページに移動してください（30秒待ちます） ===")
        await asyncio.sleep(30)

        # Inspect channel-pane-message structure
        detail = await page.evaluate("""() => {
            const results = [];

            // Look at channel-pane-message (top-level threads)
            const threads = document.querySelectorAll('[data-tid="channel-pane-message"]');
            let threadIdx = 0;
            for (const thread of threads) {
                threadIdx++;
                const threadInfo = {
                    type: 'thread',
                    index: threadIdx,
                    tag: thread.tagName,
                    dataMid: thread.querySelector('[data-mid]')?.getAttribute('data-mid') || 'none',
                    outerHTML_snippet: thread.outerHTML.substring(0, 300),
                    children_summary: []
                };

                // Look at immediate structure
                const walkChildren = (el, depth) => {
                    if (depth > 4) return;
                    for (const child of el.children) {
                        const tid = child.getAttribute('data-tid') || '';
                        const testid = child.getAttribute('data-testid') || '';
                        const mid = child.getAttribute('data-mid') || '';
                        const role = child.getAttribute('role') || '';
                        if (tid || testid || mid || role) {
                            threadInfo.children_summary.push({
                                depth,
                                tag: child.tagName.toLowerCase(),
                                dataTid: tid,
                                dataTestid: testid,
                                dataMid: mid,
                                role: role,
                                text: child.textContent?.substring(0, 60) || ''
                            });
                        }
                        walkChildren(child, depth + 1);
                    }
                };
                walkChildren(thread, 0);
                results.push(threadInfo);
                if (threadIdx >= 3) break; // Only inspect first 3 threads
            }

            // Also check data-mid elements specifically
            const midElements = document.querySelectorAll('[data-mid]');
            const midInfo = [];
            for (const el of midElements) {
                midInfo.push({
                    tag: el.tagName.toLowerCase(),
                    dataMid: el.getAttribute('data-mid'),
                    dataTid: el.getAttribute('data-tid') || '',
                    parentTid: el.parentElement?.getAttribute('data-tid') || '',
                    parentTestid: el.parentElement?.getAttribute('data-testid') || '',
                    text: el.textContent?.substring(0, 50) || ''
                });
            }

            // Check message-body elements and their parents
            const bodies = document.querySelectorAll('[data-tid="message-body"]');
            const bodyInfo = [];
            for (const body of bodies) {
                const parent = body.parentElement;
                const grandparent = parent?.parentElement;
                bodyInfo.push({
                    tag: body.tagName.toLowerCase(),
                    parentTag: parent?.tagName.toLowerCase(),
                    parentTid: parent?.getAttribute('data-tid') || '',
                    parentMid: parent?.getAttribute('data-mid') || '',
                    grandparentTid: grandparent?.getAttribute('data-tid') || '',
                    grandparentMid: grandparent?.getAttribute('data-mid') || '',
                    bodyId: body.id || '',
                    contentId: body.querySelector('[id^="content-"]')?.id || '',
                    text: body.textContent?.substring(0, 50) || ''
                });
            }

            // Check post-message-subheader structure
            const subheaders = document.querySelectorAll('[data-tid="post-message-subheader"]');
            const subheaderInfo = [];
            for (const sh of subheaders) {
                const authorEl = sh.querySelector('[data-tid="message-author-name"]')
                    || sh.querySelector('.fui-StyledText');
                subheaderInfo.push({
                    html: sh.innerHTML.substring(0, 300),
                    authorText: authorEl?.textContent || 'not found',
                    hasAuthorName: !!sh.querySelector('[data-tid="message-author-name"]')
                });
            }

            // Check reply-message-header structure
            const replyHeaders = document.querySelectorAll('[data-tid="reply-message-header"]');
            const replyInfo = [];
            for (const rh of replyHeaders) {
                const authorEl = rh.querySelector('[data-tid="message-author-name"]')
                    || rh.querySelector('.fui-StyledText');
                replyInfo.push({
                    html: rh.innerHTML.substring(0, 300),
                    authorText: authorEl?.textContent || 'not found',
                    hasAuthorName: !!rh.querySelector('[data-tid="message-author-name"]')
                });
                if (replyInfo.length >= 3) break;
            }

            return { threads: results, midElements: midInfo, bodies: bodyInfo, subheaders: subheaderInfo, replyHeaders: replyInfo };
        }""")

        import json
        print("\n=== data-mid elements ===")
        for m in detail['midElements']:
            print(f"  <{m['tag']}> data-mid={m['dataMid']} data-tid={m['dataTid']} parent-tid={m['parentTid']} | {m['text'][:40]}")

        print(f"\n=== message-body elements ({len(detail['bodies'])}) ===")
        for b in detail['bodies']:
            print(f"  id={b['bodyId']} content-id={b['contentId']} parent-mid={b['parentMid']} parent-tid={b['parentTid']} gp-tid={b['grandparentTid']} | {b['text'][:40]}")

        print(f"\n=== post-message-subheader ({len(detail['subheaders'])}) ===")
        for s in detail['subheaders']:
            print(f"  hasAuthorName={s['hasAuthorName']} author={s['authorText'][:40]}")

        print(f"\n=== reply-message-header (first 3 of {len(detail['replyHeaders'])}) ===")
        for r in detail['replyHeaders']:
            print(f"  hasAuthorName={r['hasAuthorName']} author={r['authorText'][:40]}")

        print(f"\n=== Thread structure (first 3) ===")
        for t in detail['threads']:
            print(f"\n  Thread #{t['index']} data-mid={t['dataMid']}")
            for c in t['children_summary']:
                indent = "    " + "  " * c['depth']
                print(f"{indent}<{c['tag']}> tid={c['dataTid']} testid={c['dataTestid'][:40]} mid={c['dataMid']} role={c['role']} | {c['text'][:40]}")

        print("\n=== 調査完了 ===")

asyncio.run(main())
