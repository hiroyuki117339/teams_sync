"""Inspect the exact relationship between data-mid, sender, timestamp, and avatar elements."""
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

        print("\n=== Teams チャネルに移動してください（30秒待ちます） ===")
        await asyncio.sleep(30)

        detail = await page.evaluate("""() => {
            const results = [];

            // For each data-mid element, trace up to find the nearest ancestor
            // that contains sender, timestamp, avatar
            const midElements = document.querySelectorAll('[data-mid]');

            for (const midEl of midElements) {
                const mid = midEl.getAttribute('data-mid');

                // Check: does this element contain sender/timestamp?
                const hasSenderInside = !!midEl.querySelector('[data-tid="post-message-subheader"]')
                    || !!midEl.querySelector('[data-tid="reply-message-header"]');
                const hasTimestampInside = !!midEl.querySelector('[data-tid="timestamp"]');
                const hasBodyInside = !!midEl.querySelector('[data-tid="message-body"]');

                // Walk up to find the container that has sender + timestamp + body
                let container = midEl;
                let depth = 0;
                let containerInfo = null;
                while (container && depth < 10) {
                    const hasSender = !!container.querySelector('[data-tid="post-message-subheader"]')
                        || !!container.querySelector('[data-tid="reply-message-header"]');
                    const hasTime = !!container.querySelector('[data-tid="timestamp"]');
                    const hasBody = !!container.querySelector('[data-tid="message-body"]');

                    if (hasSender && hasTime && hasBody) {
                        containerInfo = {
                            depth: depth,
                            tag: container.tagName.toLowerCase(),
                            dataTid: container.getAttribute('data-tid') || '',
                            dataTestid: container.getAttribute('data-testid') || '',
                            className: container.className?.toString().substring(0, 60) || '',
                            role: container.getAttribute('role') || ''
                        };
                        break;
                    }
                    container = container.parentElement;
                    depth++;
                }

                // Get the midEl's own outerHTML (first 200 chars)
                const midOuterSnippet = midEl.outerHTML.substring(0, 200);

                // Check parent structure
                const parent = midEl.parentElement;
                const grandparent = parent?.parentElement;
                const greatGP = grandparent?.parentElement;

                results.push({
                    mid: mid,
                    hasSenderInside,
                    hasTimestampInside,
                    hasBodyInside,
                    midTag: midEl.tagName.toLowerCase(),
                    midClass: midEl.className?.toString().substring(0, 60) || '',
                    parentTag: parent?.tagName.toLowerCase(),
                    parentTid: parent?.getAttribute('data-tid') || '',
                    parentClass: parent?.className?.toString().substring(0, 60) || '',
                    gpTag: grandparent?.tagName.toLowerCase(),
                    gpTid: grandparent?.getAttribute('data-tid') || '',
                    ggpTag: greatGP?.tagName.toLowerCase(),
                    ggpTid: greatGP?.getAttribute('data-tid') || '',
                    containerInfo: containerInfo,
                    snippet: midOuterSnippet
                });

                if (results.length >= 5) break;
            }

            // Also check: what is the parent of post-message-subheader?
            const subheaders = document.querySelectorAll('[data-tid="post-message-subheader"]');
            const shParents = [];
            for (const sh of subheaders) {
                let p = sh.parentElement;
                const chain = [];
                for (let i = 0; i < 5 && p; i++) {
                    chain.push({
                        tag: p.tagName.toLowerCase(),
                        tid: p.getAttribute('data-tid') || '',
                        mid: p.getAttribute('data-mid') || '',
                        testid: (p.getAttribute('data-testid') || '').substring(0, 40),
                        class: p.className?.toString().substring(0, 60) || ''
                    });
                    p = p.parentElement;
                }
                // Get sender text
                const senderText = sh.querySelector('.fui-StyledText')?.textContent || '';
                shParents.push({ senderText: senderText.substring(0, 40), chain });
                if (shParents.length >= 3) break;
            }

            // Same for reply-message-header
            const replyHeaders = document.querySelectorAll('[data-tid="reply-message-header"]');
            const rhParents = [];
            for (const rh of replyHeaders) {
                let p = rh.parentElement;
                const chain = [];
                for (let i = 0; i < 5 && p; i++) {
                    chain.push({
                        tag: p.tagName.toLowerCase(),
                        tid: p.getAttribute('data-tid') || '',
                        mid: p.getAttribute('data-mid') || '',
                        class: p.className?.toString().substring(0, 60) || ''
                    });
                    p = p.parentElement;
                }
                const senderText = rh.querySelector('.fui-StyledText')?.textContent || '';
                rhParents.push({ senderText: senderText.substring(0, 40), chain });
                if (rhParents.length >= 3) break;
            }

            return { messages: results, subheaderParents: shParents, replyHeaderParents: rhParents };
        }""")

        print("\n=== data-mid elements (first 5) ===")
        for m in detail['messages']:
            print(f"\n  mid={m['mid']}")
            print(f"    hasSenderInside={m['hasSenderInside']} hasTimeInside={m['hasTimestampInside']} hasBodyInside={m['hasBodyInside']}")
            print(f"    midEl: <{m['midTag']}> class={m['midClass'][:40]}")
            print(f"    parent: <{m['parentTag']}> tid={m['parentTid']}")
            print(f"    grandparent: <{m['gpTag']}> tid={m['gpTid']}")
            print(f"    great-gp: <{m['ggpTag']}> tid={m['ggpTid']}")
            if m['containerInfo']:
                c = m['containerInfo']
                print(f"    --> Container found at depth {c['depth']}: <{c['tag']}> tid={c['dataTid']} testid={c['dataTestid']} role={c['role']}")
            else:
                print(f"    --> No container found with sender+time+body!")

        print("\n=== post-message-subheader parent chains ===")
        for sh in detail['subheaderParents']:
            print(f"\n  sender: {sh['senderText']}")
            for i, p in enumerate(sh['chain']):
                print(f"    {'  ' * i}parent[{i}]: <{p['tag']}> tid={p['tid']} mid={p['mid']} class={p['class'][:40]}")

        print("\n=== reply-message-header parent chains ===")
        for rh in detail['replyHeaderParents']:
            print(f"\n  sender: {rh['senderText']}")
            for i, p in enumerate(rh['chain']):
                print(f"    {'  ' * i}parent[{i}]: <{p['tag']}> tid={p['tid']} mid={p['mid']} class={p['class'][:40]}")

        print("\n=== 完了 ===")

asyncio.run(main())
