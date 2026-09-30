import asyncio
from playwright.async_api import async_playwright


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        )

        page = await browser.new_page()
        await page.goto("https://web.whatsapp.com")

        await page.wait_for_timeout(10000)

        await browser.close()


asyncio.run(main())
