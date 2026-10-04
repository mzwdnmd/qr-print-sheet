import "./style.css";
import {
  prepareZXingModule,
  readBarcodes,
} from "zxing-wasm/reader";
import QRCode from "qrcode";
import jsQR from "jsqr";
import {
  AlignmentType,
  Document,
  ImageRun,
  Packer,
  Paragraph,
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
// Ruler-calibrated source: the emblem is 41 / 153 of the sticker width,
// and its rounded white backing is 45 / 153 (39 modules including border).
const LOGO_MODULES = 10.5;
const WHITE_PATCH_MODULES = 11.5;
const PATCH_CORNER_MODULES = 1;
const QUIET_MODULES = 1;
const PAGE = { width: 21, height: 29.7, sheetWidth: 19.5, sheetHeight: 27.75 };
const MIN_CUT_GAP_CM = 0.15;
const PRINT_DPI = 300;
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
    elements.previewCount.textContent = "0 组 / 0 枚";
    return;
  }
  elements.placeholder.hidden = true;
  const layout = getSheetLayout(sizes);
  elements.grid.style.gridTemplateColumns = `repeat(${layout.columns}, 1fr)`;
  elements.grid.style.gridTemplateRows = `repeat(${layout.rows}, 1fr)`;
  for (let block = 0; block < layout.columns * layout.rows; block += 1) {
    const module = document.createElement("div");
    module.className = "paper-module";
    if (block % layout.columns === layout.columns - 1) module.classList.add("last-column");
    if (Math.floor(block / layout.columns) === layout.rows - 1) module.classList.add("last-row");
    sizes.forEach((size, index) => {
      const image = document.createElement("img");
      image.src = state.selected.previewUrl;
      image.alt = `${size.toFixed(2)} 厘米二维码`;
      image.style.width = `${(size / layout.blockWidth) * 100}%`;
      image.style.height = `${(size / layout.blockHeight) * 100}%`;
      image.style.left = `${((layout.centersX[index % 3] - size / 2) / layout.blockWidth) * 100}%`;
      image.style.top = `${((layout.centersY[Math.floor(index / 3)] - size / 2) / layout.blockHeight) * 100}%`;
      module.append(image);
    });
    elements.grid.append(module);
  }
  const count = layout.columns * layout.rows;
  elements.previewCount.textContent = `${count} 组 / ${count * 9} 枚`;
}

function getSheetLayout(sizes) {
  const columnWidths = [sizes[6], sizes[7], sizes[8]];
  const rowHeights = [sizes[2], sizes[5], sizes[8]];
  const columnSum = columnWidths.reduce((sum, value) => sum + value, 0);
  const rowSum = rowHeights.reduce((sum, value) => sum + value, 0);
  const minimumWidth = columnSum + 4 * MIN_CUT_GAP_CM;
  const minimumHeight = rowSum + 4 * MIN_CUT_GAP_CM;
  const columns = Math.max(1, Math.floor((PAGE.sheetWidth + 1e-6) / minimumWidth));
  const rows = Math.max(1, Math.floor((PAGE.sheetHeight + 1e-6) / minimumHeight));
  if (minimumWidth > PAGE.sheetWidth || minimumHeight > PAGE.sheetHeight) {
    throw new Error("九档二维码无法放进 A4 页面，请减小基准尺寸");
  }
  const blockWidth = PAGE.sheetWidth / columns;
  const blockHeight = PAGE.sheetHeight / rows;
  const columnGap = (blockWidth - columnSum) / 4;
  const rowGap = (blockHeight - rowSum) / 4;
  const centersX = columnWidths.map((width, index) =>
    columnGap * (index + 1) + columnWidths.slice(0, index).reduce((sum, value) => sum + value, 0) + width / 2,
  );
  const centersY = rowHeights.map((height, index) =>
    rowGap * (index + 1) + rowHeights.slice(0, index).reduce((sum, value) => sum + value, 0) + height / 2,
  );
  return { columns, rows, blockWidth, blockHeight, centersX, centersY };
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
  context.beginPath();
  context.roundRect(
    (side - patchPixels) / 2,
    (side - patchPixels) / 2,
    patchPixels,
    patchPixels,
    PATCH_CORNER_MODULES * modulePixels,
  );
  context.fill();
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
const toPrintPixels = (centimeters) => Math.round((centimeters / 2.54) * PRINT_DPI);

async function makeSheetImage(item, sizes, layout) {
  const bitmap = await createImageBitmap(item.blob);
  const canvas = document.createElement("canvas");
  canvas.width = toPrintPixels(PAGE.sheetWidth);
  canvas.height = toPrintPixels(PAGE.sheetHeight);
  const context = canvas.getContext("2d");
  context.fillStyle = "#fff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  try {
    for (let blockRow = 0; blockRow < layout.rows; blockRow += 1) {
      for (let blockColumn = 0; blockColumn < layout.columns; blockColumn += 1) {
        sizes.forEach((size, index) => {
          const x = blockColumn * layout.blockWidth + layout.centersX[index % 3] - size / 2;
          const y = blockRow * layout.blockHeight + layout.centersY[Math.floor(index / 3)] - size / 2;
          const left = toPrintPixels(x);
          const top = toPrintPixels(y);
          context.drawImage(bitmap, left, top, toPrintPixels(x + size) - left, toPrintPixels(y + size) - top);
        });
      }
    }
  } finally {
    bitmap.close();
  }
  context.strokeStyle = "#777";
  context.lineWidth = 2;
  for (let column = 1; column < layout.columns; column += 1) {
    const x = toPrintPixels(column * layout.blockWidth);
    context.beginPath();
    context.moveTo(x, 0);
    context.lineTo(x, canvas.height);
    context.stroke();
  }
  for (let row = 1; row < layout.rows; row += 1) {
    const y = toPrintPixels(row * layout.blockHeight);
    context.beginPath();
    context.moveTo(0, y);
    context.lineTo(canvas.width, y);
    context.stroke();
  }
  context.strokeRect(3, 3, canvas.width - 6, canvas.height - 6);
  return new Promise((resolve, reject) =>
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("无法生成打印页")), "image/png"),
  );
}

async function makeWord(item, sizes) {
  const layout = getSheetLayout(sizes);
  const sheetBlob = await makeSheetImage(item, sizes, layout);
  const imageBytes = new Uint8Array(await sheetBlob.arrayBuffer());
  const document = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: toTwip(PAGE.width), height: toTwip(PAGE.height) },
          margin: { top: toTwip(0.9), right: toTwip(0.75), bottom: toTwip(0.9), left: toTwip(0.75) },
        },
      },
      children: [new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({
          type: "png",
          data: imageBytes,
          transformation: { width: toPixels(PAGE.sheetWidth), height: toPixels(PAGE.sheetHeight) },
        })],
      })],
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
    downloadBlob(blob, `${stem}_九档重复排版.docx`);
    const layout = getSheetLayout(sizes);
    setMessage(`Word 文件已下载：${layout.columns * layout.rows} 组，组间有裁剪线。打印时选择实际大小 / 100%。`, "success");
  } catch (error) {
    setMessage(`生成失败：${error.message || String(error)}`, "error");
  } finally {
    elements.download.disabled = false;
  }
});

renderSizes();
