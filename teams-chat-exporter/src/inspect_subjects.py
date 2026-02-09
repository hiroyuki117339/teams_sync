"""Inspect thread subjects and collapsed reply buttons."""
import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

PROJECT_ROOT = Path(__file__).parent.parent
USER_DATA_DIR = PROJECT_ROOT / "playwright_user_data"

async def main():
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(USER_DATA_DIR, headless=False)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        if "teams.microsoft.com" not in page.url:
            await page.goto("https://teams.microsoft.com/")

        print("\n=== チャネルに移動してください（30秒） ===")
        await asyncio.sleep(30)

        data = await page.evaluate("""() => {
            const results = [];

            // For each subject-line, find which data-mid it belongs to
            const subjects = document.querySelectorAll('[data-tid="subject-line"]');
            for (const sub of subjects) {
                const text = sub.textContent?.trim() || '';
                // Walk up to find the nearest [data-mid] ancestor or sibling
                let el = sub;
                let midValue = '';
                for (let i = 0; i < 10; i++) {
                    el = el.parentElement;
                    if (!el) break;
                    const mid = el.getAttribute('data-mid');
                    if (mid) { midValue = mid; break; }
                    // Check if any child has data-mid
                    const midChild = el.querySelector('[data-mid]');
                    if (midChild) { midValue = midChild.getAttribute('data-mid'); break; }
                }
                results.push({ subject: text, nearestMid: midValue });
            }

            // Check for collapsed reply buttons
            const replyButtons = document.querySelectorAll('[data-tid="response-summary-button"]');
            const buttons = [];
            for (const btn of replyButtons) {
                const text = btn.textContent?.trim() || '';
                // Find nearest data-mid ancestor
                let el = btn;
                let midValue = '';
                for (let i = 0; i < 10; i++) {
                    el = el.parentElement;
                    if (!el) break;
                    const mid = el.getAttribute('data-mid');
                    if (mid) { midValue = mid; break; }
                }
                buttons.push({ text: text.substring(0, 80), nearestMid: midValue, ariaExpanded: btn.getAttribute('aria-expanded') });
            }

            // Check "see more" / expand elements
            const seeMoreBtns = document.querySelectorAll('[data-testid^="see-more-content-"] button, [data-tid="quick-reply-placeholder-reply"]');
            const seeMore = [];
            for (const btn of seeMoreBtns) {
                seeMore.push({
                    text: btn.textContent?.trim().substring(0, 50),
                    tid: btn.getAttribute('data-tid') || '',
                    testid: btn.getAttribute('data-testid') || '',
                    tag: btn.tagName.toLowerCase()
                });
            }

            // Count total visible data-mid vs total replies per thread
            const threadMessages = document.querySelectorAll('[data-tid="channel-pane-message"]');
            const threads = [];
            for (const thread of threadMessages) {
                const threadMid = thread.getAttribute('data-mid') || '';
                const allMids = thread.querySelectorAll('[data-mid]');
                const replySurface = thread.querySelector('[data-tid="response-surface"]');
                const replyMids = replySurface ? replySurface.querySelectorAll('[data-mid]') : [];
                const summaryBtn = thread.querySelector('[data-tid="response-summary-button"]');
                threads.push({
                    threadMid,
                    totalMidsInThread: allMids.length,
                    replyMidsCount: replyMids.length,
                    hasSummaryBtn: !!summaryBtn,
                    summaryText: summaryBtn?.textContent?.trim().substring(0, 80) || ''
                });
            }

            return { subjects: results, replyButtons: buttons, seeMore, threads };
        }""")

        print("\n=== Subject lines ===")
        for s in data['subjects']:
            print(f"  mid={s['nearestMid']} | {s['subject']}")

        print(f"\n=== Reply expansion buttons ({len(data['replyButtons'])}) ===")
        for b in data['replyButtons']:
            print(f"  mid={b['nearestMid']} expanded={b['ariaExpanded']} | {b['text']}")

        print(f"\n=== See more / expand ({len(data['seeMore'])}) ===")
        for s in data['seeMore']:
            print(f"  tid={s['tid']} testid={s['testid']} | {s['text']}")

        print(f"\n=== Thread structure ({len(data['threads'])} threads) ===")
        for t in data['threads']:
            print(f"  thread mid={t['threadMid']} | total_mids={t['totalMidsInThread']} replies={t['replyMidsCount']} hasSummary={t['hasSummaryBtn']} | {t['summaryText']}")

        print("\n=== 完了 ===")

asyncio.run(main())
