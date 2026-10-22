"use strict";

const $ = (selector) => document.querySelector(selector);

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const pdfToDocxTab = $("#pdfToDocxTab");
const docxToPdfTab = $("#docxToPdfTab");
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
let conversionMode = "pdf-to-docx";
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";
let busy = false;

function setMode(mode) {
  if (busy) return;
  conversionMode = mode;
  clearSelectedFile();

  const isPdfToDocx = mode === "pdf-to-docx";
  pdfToDocxTab.classList.toggle("active", isPdfToDocx);
  docxToPdfTab.classList.toggle("active", !isPdfToDocx);
  fileInput.accept = isPdfToDocx ? ".pdf,application/pdf" :
    ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  uploadTitle.textContent = isPdfToDocx ? "Drop your PDF here" : "Drop your DOCX here";
  uploadDescription.textContent = isPdfToDocx
    ? "or choose a PDF file from your computer"
    : "or choose a Word document from your computer";
  fileLimit.textContent = `${isPdfToDocx ? "PDF" : "DOCX"} files only · Maximum 20 MB`;
  convertButton.innerHTML = isPdfToDocx
    ? 'Convert to DOCX <span aria-hidden="true">→</span>'
    : 'Convert to PDF <span aria-hidden="true">→</span>';
}

pdfToDocxTab.addEventListener("click", () => setMode("pdf-to-docx"));
docxToPdfTab.addEventListener("click", () => setMode("docx-to-pdf"));
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
  const ext = conversionMode === "pdf-to-docx" ? ".pdf" : ".docx";

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
    if (conversionMode === "pdf-to-docx") {
      const signature = new TextDecoder().decode(bytes.slice(0, 5));
      if (signature !== "%PDF-") {
        showMessage("This file does not appear to be a valid PDF.", "error");
        return;
      }
    } else {
      const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
      if (!isZip) {
        showMessage("This file does not appear to be a valid DOCX file.", "error");
        return;
      }
    }

    selectedFile = file;
    convertedBlob = null;
    convertedFileName = "";
    fileNameElement.textContent = file.name;
    fileSizeElement.textContent = formatFileSize(file.size);
    fileIcon.textContent = conversionMode === "pdf-to-docx" ? "PDF" : "DOCX";
    fileIcon.className = conversionMode === "pdf-to-docx" ? "file-icon" : "file-icon word";
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
    if (conversionMode === "pdf-to-docx") {
      await convertPdfToDocx(selectedFile);
    } else {
      await convertDocxToPdf(selectedFile);
    }
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
  docxToPdfTab.disabled = disabled;
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
    throw new Error("A required library failed to load. Refresh the page and try again.");
  }

  setProgress(10, "Reading Word document…");
  const arrayBuffer = await file.arrayBuffer();

  // Explicitly embed DOCX images as data URLs so html2canvas can render them.
  const resultHtml = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.imgElement((image) =>
        image.read("base64").then((imageBuffer) => ({
          src: `data:${image.contentType};base64,${imageBuffer}`
        }))
      )
    }
  );

  if (!resultHtml.value || !resultHtml.value.trim() ||
      !resultHtml.value.replace(/<[^>]*>/g, "").trim() &&
      !resultHtml.value.includes("<img")) {
    throw new Error("No readable text or images were found in this DOCX file.");
  }

  setProgress(35, "Preparing document for PDF…");
  const container = document.createElement("div");
  Object.assign(container.style, {
    // Keep the render element inside the viewport. Positioning it at -10000px
    // can cause html2canvas to capture a blank canvas in some browsers.
    position: "fixed",
    left: "0",
    top: "0",
    width: "760px",
    boxSizing: "border-box",
    padding: "24px",
    backgroundColor: "#ffffff",
    color: "#111111",
    fontFamily: "Arial, sans-serif",
    fontSize: "12px",
    lineHeight: "1.5",
    overflowWrap: "anywhere",
    zIndex: "-1",
    opacity: "1",
    visibility: "visible",
    pointerEvents: "none"
  });
  container.innerHTML = resultHtml.value;

  const style = document.createElement("style");
  style.textContent = `
    .fileflow-pdf-render, .fileflow-pdf-render * { box-sizing: border-box; }
    .fileflow-pdf-render img { max-width: 100%; height: auto; }
    .fileflow-pdf-render table { width: 100%; border-collapse: collapse; table-layout: auto; }
    .fileflow-pdf-render td, .fileflow-pdf-render th { border: 1px solid #ccc; padding: 5px; }
    .fileflow-pdf-render p { margin: 0 0 10px; }
    .fileflow-pdf-render h1 { font-size: 24px; margin: 0 0 16px; }
    .fileflow-pdf-render h2 { font-size: 19px; margin: 14px 0 10px; }
    .fileflow-pdf-render pre { white-space: pre-wrap; overflow-wrap: anywhere; }
  `;
  container.className = "fileflow-pdf-render";
  container.prepend(style);

  // Ensure all embedded images have decoded before html2canvas starts.
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
        try { await img.decode(); } catch (_) { /* handled by render check below */ }
      }
    }));

    if (document.fonts && document.fonts.ready) {
      await document.fonts.ready;
    }
    // Allow the browser one frame to lay out the newly inserted content.
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    if (container.scrollHeight <= 30 || container.scrollWidth <= 30) {
      throw new Error("The DOCX content could not be laid out. Try saving the file again as .docx.");
    }

    setProgress(60, "Rendering PDF pages…");
    const worker = html2pdf().set({
      margin: [10, 10, 10, 10],
      filename: makeOutputName(file.name),
      image: { type: "jpeg", quality: 0.96 },
      html2canvas: {
        scale: 1.5,
        useCORS: false,
        allowTaint: false,
        backgroundColor: "#ffffff",
        logging: false,
        scrollX: 0,
        scrollY: 0,
        windowWidth: Math.max(800, container.scrollWidth),
        windowHeight: Math.max(window.innerHeight, container.scrollHeight)
      },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] }
    }).from(container).toPdf();

    const pdf = await worker.get("pdf");
    if (!pdf || pdf.getNumberOfPages() < 1) {
      throw new Error("The PDF renderer did not create any pages. Please try another DOCX file.");
    }
    convertedBlob = pdf.output("blob");
    if (!(convertedBlob instanceof Blob) || convertedBlob.size < 100) {
      throw new Error("The generated PDF is empty. Please try another DOCX file.");
    }
  } finally {
    container.remove();
  }
  setProgress(92, "Finishing PDF…");
}

function makeOutputName(name) {
  const base = name.replace(/\.[^.]+$/, "");
  return conversionMode === "pdf-to-docx" ? `${base}.docx` : `${base}.pdf`;
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
