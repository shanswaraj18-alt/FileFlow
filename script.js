"use strict";

const $ = (selector) => document.querySelector(selector);

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const pdfToDocxTab = $("#pdfToDocxTab");
const fileInput = $("#fileInput");
const dropZone = $("#dropZone");
const chooseButton = $("#chooseButton");
const convertButton = $("#convertButton");
const selectedFileBox = $("#selectedFile");
const removeButton = $("#removeButton");
const fileNameElement = $("#fileName");
const fileSizeElement = $("#fileSize");
const fileIcon = $("#fileIcon");
const uploadTitle = $("#uploadTitle");
const uploadDescription = $("#uploadDescription");
const fileLimit = $("#fileLimit");
const message = $("#message");
const progressWrap = $("#progressWrap");
const progressBar = $("#progressBar");
const progressPercent = $("#progressPercent");
const statusText = $("#statusText");
const result = $("#result");
const resultName = $("#resultName");
const downloadButton = $("#downloadButton");

const MAX_FILE_SIZE = 20 * 1024 * 1024;
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";
let busy = false;

function setMode(mode) {
  // FileFlow supports PDF to DOCX only.
}

chooseButton.addEventListener("click", () => { if (!busy) fileInput.click(); });

dropZone.addEventListener("click", (event) => {
  if (!busy && !event.target.closest("button")) fileInput.click();
});
dropZone.addEventListener("keydown", (event) => {
  if (!busy && (event.key === "Enter" || event.key === " ")) {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", () => {
  if (fileInput.files && fileInput.files[0]) acceptFile(fileInput.files[0]);
});

["dragenter", "dragover"].forEach((name) => {
  dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    if (!busy) dropZone.classList.add("dragging");
  });
});
["dragleave", "drop"].forEach((name) => {
  dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});
dropZone.addEventListener("drop", (event) => {
  if (!busy && event.dataTransfer.files[0]) acceptFile(event.dataTransfer.files[0]);
});
removeButton.addEventListener("click", clearSelectedFile);

async function acceptFile(file) {
  if (busy) return;
  const ext = ".pdf";

  if (!file.name.toLowerCase().endsWith(ext)) {
    showMessage(`Please choose a ${ext} file.`, "error");
    return;
  }
  if (!file.size) {
    showMessage("The selected file is empty.", "error");
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    showMessage("The maximum file size is 20 MB.", "error");
    return;
  }

  try {
    const bytes = new Uint8Array(await file.slice(0, 8).arrayBuffer());
    const signature = new TextDecoder().decode(bytes.slice(0, 5));
    if (signature !== "%PDF-") {
      showMessage("This file does not appear to be a valid PDF.", "error");
      return;
    }

    selectedFile = file;
    convertedBlob = null;
    convertedFileName = "";
    fileNameElement.textContent = file.name;
    fileSizeElement.textContent = formatFileSize(file.size);
    fileIcon.textContent = "PDF";
    fileIcon.className = "file-icon";
    selectedFileBox.hidden = false;
    result.hidden = true;
    progressWrap.hidden = true;
    hideMessage();
    convertButton.disabled = false;
  } catch (error) {
    showMessage("Could not read this file. Please try another one.", "error");
  }
}

function clearSelectedFile() {
  if (busy) return;
  selectedFile = null;
  convertedBlob = null;
  convertedFileName = "";
  fileInput.value = "";
  selectedFileBox.hidden = true;
  progressWrap.hidden = true;
  result.hidden = true;
  convertButton.disabled = true;
  progressBar.style.width = "0%";
  progressPercent.textContent = "0%";
  hideMessage();
}

convertButton.addEventListener("click", async () => {
  if (!selectedFile || busy) return;
  busy = true;
  setControlsDisabled(true);
  result.hidden = true;
  progressWrap.hidden = false;
  hideMessage();
  setProgress(2, "Starting conversion…");

  try {
    await convertPdfToDocx(selectedFile);
    if (!(convertedBlob instanceof Blob) || convertedBlob.size === 0) {
      throw new Error("The converter returned an empty file. Please try another document.");
    }

    convertedFileName = makeOutputName(selectedFile.name);
    setProgress(100, "Conversion complete");
    resultName.textContent = convertedFileName;
    result.hidden = false;
    showMessage("Your converted file is ready to download.", "info");
  } catch (error) {
    console.error("FileFlow conversion error:", error);
    convertedBlob = null;
    convertedFileName = "";
    showMessage(error?.message || "Conversion failed. Please try another file.", "error");
  } finally {
    busy = false;
    setControlsDisabled(false);
    convertButton.disabled = !selectedFile;
  }
});

function setControlsDisabled(disabled) {
  chooseButton.disabled = disabled;
  removeButton.disabled = disabled;
  pdfToDocxTab.disabled = disabled;
}

async function convertPdfToDocx(file) {
  if (!window.pdfjsLib || !window.docx) {
    throw new Error("A required library failed to load. Refresh the page and try again.");
  }

  // Visual-preservation mode: render every PDF page and place that full page
  // into the DOCX. Unlike text extraction/OCR, this retains the page's visible
  // text, graphics, tables, and positioning as one image.
  const pdf = await pdfjsLib.getDocument({
    data: new Uint8Array(await file.arrayBuffer())
  }).promise;

  const sections = [];
  const MAX_CANVAS_PIXELS = 18_000_000;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    setProgress(
      5 + Math.round(((pageNumber - 1) / pdf.numPages) * 82),
      `Preserving page ${pageNumber} of ${pdf.numPages}…`
    );

    const page = await pdf.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });

    // Keep the raster size reasonable for browser memory while retaining
    // enough detail for normal document viewing and printing.
    let scale = 2;
    const estimatedPixels = baseViewport.width * baseViewport.height * scale * scale;
    if (estimatedPixels > MAX_CANVAS_PIXELS) {
      scale = Math.sqrt(MAX_CANVAS_PIXELS / (baseViewport.width * baseViewport.height));
    }
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);

    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error(`Could not render PDF page ${pageNumber}.`);

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: context, viewport }).promise;

    const dataUrl = canvas.toDataURL("image/png");
    const imageBytes = dataUrlToUint8Array(dataUrl);

    // PDF.js viewport dimensions at scale 1 are CSS pixels. 15 twips per
    // CSS pixel maps the page to the same physical dimensions in Word.
    const pageWidthTwips = Math.max(1, Math.round(baseViewport.width * 15));
    const pageHeightTwips = Math.max(1, Math.round(baseViewport.height * 15));
    const imageWidthPx = baseViewport.width;
    const imageHeightPx = baseViewport.height;

    const imageRun = new docx.ImageRun({
      data: imageBytes,
      transformation: {
        width: imageWidthPx,
        height: imageHeightPx
      }
    });

    sections.push({
      properties: {
        type: pageNumber === 1 ? docx.SectionType.CONTINUOUS : docx.SectionType.NEXT_PAGE,
        page: {
          size: {
            width: pageWidthTwips,
            height: pageHeightTwips
          },
          margin: {
            top: 0,
            right: 0,
            bottom: 0,
            left: 0,
            header: 0,
            footer: 0,
            gutter: 0
          }
        }
      },
      children: [
        new docx.Paragraph({
          children: [imageRun],
          spacing: { before: 0, after: 0, line: 240 },
          indent: { left: 0, right: 0 }
        })
      ]
    });

    canvas.width = 0;
    canvas.height = 0;
  }

  setProgress(93, "Building layout-preserving Word document…");
  const documentFile = new docx.Document({ sections });
  convertedBlob = await docx.Packer.toBlob(documentFile);
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function convertDocxToPdf(file) {
  if (!window.mammoth || !window.html2pdf) {
    throw new Error("A required conversion library did not load. Refresh the page and try again.");
  }

  setProgress(10, "Reading Word document…");
  const arrayBuffer = await file.arrayBuffer();
  const mammothResult = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.imgElement((image) =>
        image.read("base64").then((base64) => ({
          src: `data:${image.contentType};base64,${base64}`
        }))
      )
    }
  );

  const html = mammothResult.value || "";
  const textOnly = html.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, " ").replace(/&nbsp;|&#160;/gi, " ").trim();
  const hasImages = /<img\b/i.test(html);
  if (!textOnly && !hasImages) {
    throw new Error("No readable text or images were found in this DOCX. Try opening it in Word and saving it again as .docx.");
  }

  setProgress(30, "Preparing Word content…");
  const container = document.createElement("div");
  container.className = "fileflow-pdf-render";
  Object.assign(container.style, {
    position: "absolute",
    left: "0",
    top: "0",
    width: "794px",
    minHeight: "1123px",
    boxSizing: "border-box",
    padding: "38px",
    margin: "0",
    background: "#ffffff",
    color: "#111111",
    fontFamily: "Arial, sans-serif",
    fontSize: "12pt",
    lineHeight: "1.45",
    overflowWrap: "break-word",
    zIndex: "2147483647",
    opacity: "1",
    visibility: "visible",
    pointerEvents: "none"
  });
  container.innerHTML = html;

  const style = document.createElement("style");
  style.textContent = `
    .fileflow-pdf-render, .fileflow-pdf-render * { box-sizing: border-box; }
    .fileflow-pdf-render { background: #fff !important; color: #111 !important; }
    .fileflow-pdf-render img { max-width: 100%; height: auto; }
    .fileflow-pdf-render table { max-width: 100%; border-collapse: collapse; }
    .fileflow-pdf-render td, .fileflow-pdf-render th { padding: 4px; }
    .fileflow-pdf-render p { orphans: 3; widows: 3; }
    .fileflow-pdf-render pre { white-space: pre-wrap; overflow-wrap: anywhere; }
  `;
  container.prepend(style);
  document.body.appendChild(container);

  try {
    const images = Array.from(container.querySelectorAll("img"));
    await Promise.all(images.map(async (img) => {
      if (!img.complete) {
        await new Promise((resolve) => {
          img.onload = resolve;
          img.onerror = resolve;
        });
      }
      if (typeof img.decode === "function") {
        try { await img.decode(); } catch (_) { /* checked below */ }
      }
    }));

    const failedImages = images.filter((img) => !img.naturalWidth || !img.naturalHeight);
    if (failedImages.length) {
      throw new Error(`Could not load ${failedImages.length} image(s) from this DOCX. Please try another DOCX file.`);
    }
    if (document.fonts?.ready) await document.fonts.ready;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    if (container.scrollHeight < 50 || container.scrollWidth < 50) {
      throw new Error("The DOCX content could not be rendered. Try saving the document again as .docx.");
    }

    setProgress(55, "Rendering document pages…");
    const options = {
      margin: 0,
      filename: makeOutputName(file.name),
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: {
        scale: 1.8,
        useCORS: true,
        allowTaint: false,
        backgroundColor: "#ffffff",
        logging: false,
        scrollX: 0,
        scrollY: 0,
        windowWidth: Math.max(document.documentElement.clientWidth, 900),
        windowHeight: Math.max(window.innerHeight, container.scrollHeight),
        onclone: (clonedDoc) => {
          const clone = clonedDoc.querySelector(".fileflow-pdf-render");
          if (clone) {
            clone.style.position = "static";
            clone.style.left = "auto";
            clone.style.top = "auto";
            clone.style.zIndex = "auto";
            clone.style.width = "794px";
            clone.style.opacity = "1";
            clone.style.visibility = "visible";
          }
        }
      },
      jsPDF: { unit: "px", format: "a4", orientation: "portrait", hotfixes: ["px_scaling"] },
      pagebreak: { mode: ["css", "legacy"] }
    };

    // Use the documented Worker -> PDF -> jsPDF Blob path. This avoids
    // accidentally treating the html2pdf Worker itself as the generated Blob.
    const worker = html2pdf().set(options).from(container).toPdf();
    const pdf = await worker.get("pdf");
    convertedBlob = pdf.output("blob");

    if (!(convertedBlob instanceof Blob) || convertedBlob.size < 500) {
      throw new Error("The generated PDF is empty. Please try another DOCX file.");
    }
    const signature = new TextDecoder().decode(
      new Uint8Array(await convertedBlob.slice(0, 5).arrayBuffer())
    );
    if (signature !== "%PDF-") {
      throw new Error("The converter did not create a valid PDF. Please try again.");
    }
  } finally {
    container.remove();
  }

  setProgress(92, "Finishing PDF…");
}

function makeOutputName(name) {
  const base = name.replace(/\.[^.]+$/, "");
  return `${base}.docx`;
}
function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}
function setProgress(percent, text) {
  progressBar.style.width = `${percent}%`;
  progressPercent.textContent = `${percent}%`;
  statusText.textContent = text;
}
function showMessage(text, type) {
  message.textContent = text;
  message.className = `message ${type}`;
  message.hidden = false;
}
function hideMessage() {
  message.hidden = true;
  message.textContent = "";
  message.className = "message";
}

downloadButton.addEventListener("click", () => {
  if (!convertedBlob || !convertedFileName) return;
  const url = URL.createObjectURL(convertedBlob);
  const link = document.createElement("a");
  link.href = url;
  link.download = convertedFileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
});
