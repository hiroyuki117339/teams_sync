"""Inspect Teams channel DOM to discover correct selectors."""
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

        # Search in both main page and iframes
        scopes = [("main page", page)] + [(f"iframe[{i}]", f) for i, f in enumerate(page.frames) if f.parent_frame]

        for scope_name, scope in scopes:
            print(f"\n--- Searching in: {scope_name} (URL: {scope.url[:80]}) ---")

            # Find all elements with data-tid attribute
            tids = await scope.evaluate("""() => {
                const elements = document.querySelectorAll('[data-tid]');
                const result = {};
                elements.forEach(el => {
                    const tid = el.getAttribute('data-tid');
                    const tag = el.tagName.toLowerCase();
                    const classes = el.className ? el.className.toString().substring(0, 60) : '';
                    const key = `${tid}`;
                    if (!result[key]) {
                        result[key] = { tag, classes, count: 0, sample_text: el.textContent?.substring(0, 50) || '' };
                    }
                    result[key].count++;
                });
                return result;
            }""")

            if tids:
                print(f"\n  Found {len(tids)} unique data-tid values:")
                for tid, info in sorted(tids.items()):
                    print(f"    data-tid=\"{tid}\" | <{info['tag']}> x{info['count']} | {info.get('sample_text', '')[:40]}")

            # Find elements with data-testid
            testids = await scope.evaluate("""() => {
                const elements = document.querySelectorAll('[data-testid]');
                const result = {};
                elements.forEach(el => {
                    const testid = el.getAttribute('data-testid');
                    const tag = el.tagName.toLowerCase();
                    const key = `${testid}`;
                    if (!result[key]) {
                        result[key] = { tag, count: 0 };
                    }
                    result[key].count++;
                });
                return result;
            }""")

            if testids:
                print(f"\n  Found {len(testids)} unique data-testid values:")
                for testid, info in sorted(testids.items()):
                    print(f"    data-testid=\"{testid}\" | <{info['tag']}> x{info['count']}")

            # Specifically look for scroll containers and message-like structures
            scroll_info = await scope.evaluate("""() => {
                const scrollables = [];
                const all = document.querySelectorAll('*');
                for (const el of all) {
                    const style = getComputedStyle(el);
                    if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > 500) {
                        scrollables.push({
                            tag: el.tagName.toLowerCase(),
                            id: el.id || '',
                            dataTid: el.getAttribute('data-tid') || '',
                            dataTestid: el.getAttribute('data-testid') || '',
                            role: el.getAttribute('role') || '',
                            className: el.className?.toString().substring(0, 80) || '',
                            scrollHeight: el.scrollHeight,
                            childCount: el.children.length
                        });
                    }
                }
                return scrollables;
            }""")

            if scroll_info:
                print(f"\n  Scrollable containers (scrollHeight > 500):")
                for s in scroll_info:
                    print(f"    <{s['tag']}> id=\"{s['id']}\" data-tid=\"{s['dataTid']}\" data-testid=\"{s['dataTestid']}\" role=\"{s['role']}\" scrollH={s['scrollHeight']} children={s['childCount']}")
                    print(f"      class: {s['className']}")

        print("\n=== 調査完了 ===")

asyncio.run(main())
