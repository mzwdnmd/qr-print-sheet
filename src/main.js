import "./style.css";
import {
  prepareZXingModule,
  readBarcodes,
} from "zxing-wasm/reader";
import QRCode from "qrcode";
import jsQR from "jsqr";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  VerticalAlignTable,
  WidthType,
} from "docx";

const BASE_URL = new URL(import.meta.env.BASE_URL, window.location.href);
prepareZXingModule({
  overrides: {
    locateFile: (path, prefix) =>
      path.endsWith(".wasm") ? new URL(path, BASE_URL).href : prefix + path,
  },
});

const elements = {
  input: document.querySelector("#photo-input"),
  dropzone: document.querySelector("#dropzone"),
  message: document.querySelector("#message"),
  fileList: document.querySelector("#file-list"),
  baseSize: document.querySelector("#base-size"),
  chips: document.querySelector("#size-chips"),
  download: document.querySelector("#download-button"),
  grid: document.querySelector("#paper-grid"),
  placeholder: document.querySelector("#paper-placeholder"),
  previewCount: document.querySelector("#preview-count"),
};

const state = { items: [], selected: null, busy: false };
// The source sticker has an 11-module square cutout around a roughly
// 5.2-module center mark. Keep the cutout aligned with the module grid.
const LOGO_MODULES = 5.2;
const WHITE_PATCH_MODULES = 11;
const QUIET_MODULES = 1;
let logoBitmapPromise;

function setMessage(message, tone = "") {
  elements.message.textContent = message;
  elements.message.className = `message ${tone}`.trim();
}

function getSizes() {
  const base = Number(elements.baseSize.value);
  if (!Number.isFinite(base) || base < 0.70 || base > 5.00) {
    return null;
  }
  return Array.from({ length: 9 }, (_, index) =>
    Math.round((base + (index - 4) * 0.05) * 100) / 100,
  );
}

function renderSizes() {
  const sizes = getSizes();
  elements.chips.replaceChildren();
  if (!sizes) {
    setMessage("基准尺寸请输入 0.70–5.00 cm。", "error");
    elements.download.disabled = true;
    return;
  }
  sizes.forEach((size, index) => {
    const chip = document.createElement("span");
    chip.className = `size-chip${index === 4 ? " base" : ""}`;
    chip.textContent = `${size.toFixed(2)} cm`;
    elements.chips.append(chip);
  });
  elements.download.disabled = state.busy || !state.selected;
  renderPreview();
}

function renderFileList() {
  elements.fileList.replaceChildren();
  state.items.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `file-item${state.selected === item ? " active" : ""}`;
    button.setAttribute("aria-pressed", state.selected === item ? "true" : "false");
    const icon = document.createElement("span");
    icon.className = "file-thumb";
    icon.textContent = "▦";
    const name = document.createElement("span");
    name.className = "file-name";
    name.title = item.names.join("、");
    name.textContent = item.names[0];
    const status = document.createElement("span");
    status.className = "file-state";
    status.textContent = item.names.length > 1 ? `${item.names.length} 张合并` : "已识别";
    button.append(icon, name, status);
    button.addEventListener("click", () => {
      state.selected = item;
      renderFileList();
      renderPreview();
      elements.download.disabled = state.busy || !getSizes();
    });
    elements.fileList.append(button);
  });
}

function renderPreview() {
  const sizes = getSizes();
  elements.grid.replaceChildren();
  if (!state.selected || !sizes) {
    elements.placeholder.hidden = false;
    elements.previewCount.textContent = "0 / 9";
    return;
  }
  elements.placeholder.hidden = true;
  const cellCm = Math.max(...sizes) + 0.6;
  elements.grid.style.width = `${((cellCm * 3) / 21) * 100}%`;
  sizes.forEach((size) => {
    const cell = document.createElement("div");
    cell.className = "paper-cell";
    const image = document.createElement("img");
    image.src = state.selected.previewUrl;
    image.alt = `${size.toFixed(2)} 厘米二维码`;
    image.style.width = `${(size / cellCm) * 100}%`;
    cell.append(image);
    elements.grid.append(cell);
  });
  elements.previewCount.textContent = "9 / 9";
}

async function getLogoBitmap() {
  if (!logoBitmapPromise) {
    logoBitmapPromise = fetch(new URL("logo.webp", BASE_URL))
      .then((response) => {
        if (!response.ok) throw new Error("无法读取校徽素材");
        return response.blob();
      })
      .then((blob) => createImageBitmap(blob));
  }
  return logoBitmapPromise;
}

async function getImageData(file, maxEdge = 2400) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, maxEdge / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  return context.getImageData(0, 0, canvas.width, canvas.height);
}

async function decodePhoto(file) {
  const options = { formats: ["QRCode"], tryHarder: true, maxNumberOfSymbols: 1 };
  let imageData = await getImageData(file);
  let result = (await readBarcodes(imageData, options)).find(
    (barcode) => barcode.isValid && barcode.format === "QRCode",
  );
  if (!result && Math.max(imageData.width, imageData.height) === 2400) {
    imageData = await getImageData(file, 6000);
    result = (await readBarcodes(imageData, options)).find(
      (barcode) => barcode.isValid && barcode.format === "QRCode",
    );
  }
  if (!result) throw new Error("未识别到清晰的二维码");
  let metadata = {};
  try {
    metadata = JSON.parse(result.extra || "{}");
  } catch {
    // Older reader builds may omit the optional metadata.
  }
  const maskPattern = Number(metadata.DataMask);
  const version = Number(metadata.Version);
  const errorCorrectionLevel = String(metadata.ECLevel || "M");
  return {
    bytes: result.bytes,
    maskPattern: Number.isInteger(maskPattern) && maskPattern >= 0 && maskPattern <= 7 ? maskPattern : undefined,
    version: Number.isInteger(version) && version >= 1 && version <= 40 ? version : undefined,
    errorCorrectionLevel: /^[LMQH]$/.test(errorCorrectionLevel) ? errorCorrectionLevel : "M",
  };
}

async function makeQrImage(source) {
  const symbol = QRCode.create([{ data: new Uint8ClampedArray(source.bytes), mode: "byte" }], {
    errorCorrectionLevel: source.errorCorrectionLevel,
    version: source.version,
    maskPattern: source.maskPattern,
  });
  const modulePixels = 24;
  const quietModules = QUIET_MODULES;
  const moduleCount = symbol.modules.size;
  const side = (moduleCount + quietModules * 2) * modulePixels;
  const logoBitmap = await getLogoBitmap();
  const canvas = document.createElement("canvas");
  canvas.width = side;
  canvas.height = side;
  const context = canvas.getContext("2d");
  context.fillStyle = "#fff";
  context.fillRect(0, 0, side, side);
  context.fillStyle = "#000";
  for (let row = 0; row < moduleCount; row += 1) {
    for (let column = 0; column < moduleCount; column += 1) {
      if (symbol.modules.get(row, column)) {
        context.fillRect(
          (column + quietModules) * modulePixels,
          (row + quietModules) * modulePixels,
          modulePixels,
          modulePixels,
        );
      }
    }
  }
  context.fillStyle = "#fff";
  const patchPixels = WHITE_PATCH_MODULES * modulePixels;
  context.fillRect((side - patchPixels) / 2, (side - patchPixels) / 2, patchPixels, patchPixels);
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  const logoSide = modulePixels * LOGO_MODULES;
  context.drawImage(logoBitmap, (side - logoSide) / 2, (side - logoSide) / 2, logoSide, logoSide);
  const blob = await new Promise((resolve, reject) =>
    canvas.toBlob((value) => value ? resolve(value) : reject(new Error("无法生成二维码图片")), "image/png"),
  );
  return { blob, previewUrl: URL.createObjectURL(blob), version: symbol.version, maskPattern: symbol.maskPattern };
}

function payloadKey(payload) {
  return Array.from(payload, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

async function ingest(files) {
  if (state.busy || !files.length) return;
  if (files.length > 20) {
    setMessage("一次最多选择 20 张照片。", "error");
    return;
  }
  state.busy = true;
  elements.input.disabled = true;
  elements.download.disabled = true;
  const errors = [];
  let added = 0;
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    setMessage(`正在识别第 ${index + 1} / ${files.length} 张：${file.name}`);
    try {
      if (!file.type.startsWith("image/")) throw new Error("请选择图片文件");
      if (file.size > 25 * 1024 * 1024) throw new Error("文件超过 25 MB");
      const source = await decodePhoto(file);
      const bytes = source.bytes;
      const key = `${payloadKey(bytes)}:${source.maskPattern ?? "auto"}:${source.version ?? "auto"}:${source.errorCorrectionLevel}`;
      const duplicate = state.items.find((item) => item.key === key);
      if (duplicate) {
        if (!duplicate.names.includes(file.name)) duplicate.names.push(file.name);
      } else {
        const image = await makeQrImage(source);
        const item = { key, names: [file.name], payload: bytes, ...image };
        state.items.push(item);
        state.selected = item;
        added += 1;
      }
      renderFileList();
      renderPreview();
    } catch (error) {
      errors.push(`${file.name}：${error.message || String(error)}`);
    }
  }
  state.busy = false;
  elements.input.disabled = false;
  elements.input.value = "";
  renderSizes();
  if (errors.length) {
    setMessage(`新增 ${added} 个二维码。${errors.join("；")}`, "error");
  } else {
    setMessage(`识别完成：${state.items.length} 个不同的二维码，可下载单页 Word。`, "success");
  }
}

const toTwip = (centimeters) => Math.round((centimeters / 2.54) * 1440);
const toPixels = (centimeters) => (centimeters / 2.54) * 96;

async function makeWord(item, sizes) {
  const imageBytes = new Uint8Array(await item.blob.arrayBuffer());
  const cellTwip = toTwip(Math.max(...sizes) + 0.6);
  const marginTwip = toTwip(1.5);
  const rows = Array.from({ length: 3 }, (_, rowIndex) =>
    new TableRow({
      cantSplit: true,
      height: { value: cellTwip, rule: HeightRule.EXACT },
      children: Array.from({ length: 3 }, (_, columnIndex) => {
        const size = sizes[rowIndex * 3 + columnIndex];
        return new TableCell({
          width: { size: cellTwip, type: WidthType.DXA },
          verticalAlign: VerticalAlignTable.CENTER,
          margins: { top: 0, right: 0, bottom: 0, left: 0 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new ImageRun({
              type: "png",
              data: imageBytes,
              transformation: { width: toPixels(size), height: toPixels(size) },
            })],
          })],
        });
      }),
    }),
  );
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const table = new Table({
    rows,
    columnWidths: [cellTwip, cellTwip, cellTwip],
    width: { size: cellTwip * 3, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.LEFT,
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    borders: { top: none, right: none, bottom: none, left: none, insideHorizontal: none, insideVertical: none },
  });
  const document = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: toTwip(21), height: toTwip(29.7) },
          margin: { top: marginTwip, right: marginTwip, bottom: marginTwip, left: marginTwip },
        },
      },
      children: [table],
    }],
  });
  return Packer.toBlob(document);
}

async function validatePrintSizes(item, sizes) {
  const expectedPayload = payloadKey(item.payload);
  const bitmap = await createImageBitmap(item.blob);
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { willReadFrequently: true });
  try {
    for (const size of sizes) {
      const pixels = Math.round((size / 2.54) * 300);
      canvas.width = pixels;
      canvas.height = pixels;
      context.imageSmoothingEnabled = true;
      context.imageSmoothingQuality = "high";
      context.drawImage(bitmap, 0, 0, pixels, pixels);
      const imageData = context.getImageData(0, 0, pixels, pixels);
      const independentResult = jsQR(imageData.data, pixels, pixels, { inversionAttempts: "dontInvert" });
      if (!independentResult || payloadKey(independentResult.binaryData) !== expectedPayload) {
        throw new Error(`${size.toFixed(2)} cm 在 300 dpi 下无法由独立扫码器识别，请增大基准尺寸`);
      }
      const wasmResults = await readBarcodes(imageData, {
        formats: ["QRCode"], tryHarder: true, maxNumberOfSymbols: 1,
      });
      if (!wasmResults.some((result) => result.isValid && payloadKey(result.bytes) === expectedPayload)) {
        throw new Error(`${size.toFixed(2)} cm 在 300 dpi 下无法识别，请增大基准尺寸`);
      }
    }
  } finally {
    bitmap.close();
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

elements.input.addEventListener("change", (event) => ingest(Array.from(event.target.files || [])));
elements.dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropzone.classList.add("dragging");
});
elements.dropzone.addEventListener("dragleave", () => elements.dropzone.classList.remove("dragging"));
elements.dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropzone.classList.remove("dragging");
  ingest(Array.from(event.dataTransfer.files || []));
});
elements.baseSize.addEventListener("input", renderSizes);
elements.download.addEventListener("click", async () => {
  const sizes = getSizes();
  if (!state.selected || !sizes || state.busy) return;
  elements.download.disabled = true;
  setMessage("正在检查九档二维码的可读性…");
  try {
    await validatePrintSizes(state.selected, sizes);
    setMessage("检查通过，正在生成 Word 文件…");
    const blob = await makeWord(state.selected, sizes);
    const stem = state.selected.names[0]
      .replace(/\.[^.]+$/, "")
      .replace(/[<>:"/\\|?*\x00-\x1f]/g, "_")
      .slice(0, 50) || "二维码";
    downloadBlob(blob, `${stem}_九档同页.docx`);
    setMessage("Word 文件已下载。打印时选择实际大小 / 100%。", "success");
  } catch (error) {
    setMessage(`生成失败：${error.message || String(error)}`, "error");
  } finally {
    elements.download.disabled = false;
  }
});

renderSizes();
