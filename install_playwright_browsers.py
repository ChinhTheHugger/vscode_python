import asyncio
from playwright.async_api import async_playwright

async def install_browsers():
    async with async_playwright() as p:
        # This will trigger the download of the required binaries for all supported browsers
        await p.chromium.launch()
        await p.firefox.launch()
        await p.webkit.launch()

asyncio.run(install_browsers())
