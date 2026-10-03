import { chromium } from "playwright";
import { mkdir } from "node:fs/promises";
import { resolve } from "node:path";

const fixture = process.argv[2];
if (!fixture) throw new Error("用法：node scripts/smoke-local.mjs <二维码照片路径>");

const browser = await chromium.launch({
  executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  headless: true,
});
const page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
const errors = [];
page.on("pageerror", (error) => errors.push(error.message));
page.on("console", (message) => {
  if (message.type() === "error") errors.push(message.text());
});
try {
  await page.goto("http://localhost:5173/", { waitUntil: "networkidle" });
  await page.locator("#photo-input").setInputFiles(resolve(fixture));
  await page.getByText(/识别完成：1 个不同的二维码/).waitFor({ timeout: 30_000 });
  if (await page.locator("#paper-grid .paper-cell").count() !== 9) {
    throw new Error("排版预览不是九个二维码");
  }
  const output = resolve("test-artifacts");
  await mkdir(output, { recursive: true });
  await page.screenshot({ path: resolve(output, "screenshot.png"), fullPage: true });
  const downloadEvent = page.waitForEvent("download", { timeout: 30_000 });
  await page.locator("#download-button").click();
  const download = await downloadEvent;
  await download.saveAs(resolve(output, "sheet.docx"));
  if (errors.length) throw new Error(errors.join("\n"));
  console.log(`OK: ${download.suggestedFilename()}`);
} finally {
  await browser.close();
}
