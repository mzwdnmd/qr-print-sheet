import { chromium } from "playwright";
import { mkdir } from "node:fs/promises";
import { resolve } from "node:path";

const fixture = process.argv[2];
if (!fixture) throw new Error("用法：node scripts/smoke-local.mjs <二维码照片路径>");
const baseSize = Number(process.argv[3] || 1.60);

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
  await page.goto(process.env.TEST_URL || "http://localhost:5173/", { waitUntil: "networkidle" });
  await page.locator("#photo-input").setInputFiles(resolve(fixture));
  await page.getByText(/识别完成：1 个不同的二维码/).waitFor({ timeout: 30_000 });
  if (baseSize !== 1.60) await page.locator("#base-size").fill(baseSize.toFixed(2));
  await page.locator("#paper-grid .paper-module").first().waitFor({ timeout: 30_000 });
  const modules = await page.locator("#paper-grid .paper-module").count();
  const expectedModules = baseSize === 1.60 ? 15 : baseSize === 2.80 ? 6 : null;
  if (modules !== expectedModules || await page.locator("#paper-grid .paper-module img").count() !== modules * 9) {
    throw new Error(`排版预览组数不符：预计 ${expectedModules}，实际 ${modules}`);
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
