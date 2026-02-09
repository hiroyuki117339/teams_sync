"""Inspect the DOM when inside a thread view to understand how to navigate back."""
import asyncio, sys
from pathlib import Path
from playwright.async_api import async_playwright

PROJECT_ROOT = Path(__file__).parent.parent
USER_DATA_DIR = PROJECT_ROOT / "playwright_user_data"

def p(msg):
    print(msg, flush=True)

async def main():
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(USER_DATA_DIR, headless=False)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        if "teams.microsoft.com" not in page.url:
            await page.goto("https://teams.microsoft.com/")

        p("Navigate to a channel thread view. Inspecting in 30 seconds...")
        await asyncio.sleep(30)
        p("=== Inspecting thread view DOM ===")

        info = await page.evaluate("""() => {
            const results = {};
            const selectors = [
                '[data-tid="channel-pane-viewport"]',
                '[role="complementary"]',
                '[role="main"]',
                '[data-tid="thread-pane"]',
                '[data-tid="thread-pane-viewport"]',
                '[data-tid="message-pane-list-viewport"]',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                results[sel] = el ? 'FOUND' : 'not found';
            }
            const allTids = [];
            document.querySelectorAll('[data-tid]').forEach(el => {
                const tid = el.getAttribute('data-tid');
                if (tid && (tid.includes('thread') || tid.includes('back') || tid.includes('pane') || tid.includes('nav') || tid.includes('reply') || tid.includes('channel')))
                    allTids.push(tid);
            });
            results['relevant_tids'] = [...new Set(allTids)];
            const backButtons = [];
            document.querySelectorAll('button').forEach(btn => {
                const label = btn.getAttribute('aria-label') || '';
                const tid = btn.getAttribute('data-tid') || '';
                const text = btn.textContent?.trim()?.substring(0, 50) || '';
                if (label.includes('Back') || label.includes('戻') || label.includes('閉') || label.includes('Close')
                    || tid.includes('back') || tid.includes('close') || tid.includes('nav')
                    || label.includes('チャネル') || label.includes('channel')) {
                    backButtons.push({label, tid, text: text.substring(0, 50)});
                }
            });
            results['back_buttons'] = backButtons;
            results['data_mid_count'] = document.querySelectorAll('[data-mid]').length;
            results['url'] = window.location.href;
            return results;
        }""")

        for key, val in info.items():
            p(f"{key}: {val}")

        p("\n=== Done ===")
        await asyncio.sleep(300)

asyncio.run(main())
