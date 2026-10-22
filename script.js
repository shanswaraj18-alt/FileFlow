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

  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const paragraphs = [];
  let extractedCharacters = 0;
  let ocrWorker = null;

  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      setProgress(5 + Math.round(((pageNumber - 1) / pdf.numPages) * 80),
        `Reading page ${pageNumber} of ${pdf.numPages}…`);

      const content = await page.getTextContent();
      let pageText = content.items
        .filter(item => typeof item.str === "string")
        .map(item => item.str)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();

      // Scanned/image-only pages need OCR.
      if (!pageText) {
        if (!window.Tesseract) {
          throw new Error("OCR library did not load. Check your internet connection and refresh.");
        }

        if (!ocrWorker) {
          setProgress(10, "Loading English OCR engine (first use may take a little while)…");
          ocrWorker = await Tesseract.createWorker("eng", 1, {
            workerPath: "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js",
            langPath: "https://tessdata.projectnaptha.com/4.0.0",
            corePath: "https://cdn.jsdelivr.net/npm/tesseract.js-core@v5.0.0",
            logger: (info) => {
              if (info.status === "recognizing text") {
                const overall = 10 + Math.round(((pageNumber - 1 + info.progress) / pdf.numPages) * 80);
                setProgress(Math.min(overall, 90), `Recognizing text on page ${pageNumber}…`);
              }
            }
          });
        }

        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        const context = canvas.getContext("2d", { willReadFrequently: true });

        if (!context) throw new Error("Could not prepare the PDF page for OCR.");

        await page.render({ canvasContext: context, viewport }).promise;
        const ocrResult = await ocrWorker.recognize(canvas);
        pageText = (ocrResult.data.text || "").trim();

        canvas.width = 0;
        canvas.height = 0;
      }

      if (pageText) {
        // Preserve line breaks where possible for searchable PDFs and OCR output.
        const lines = pageText.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
        for (const line of lines) {
          paragraphs.push(new docx.Paragraph({
            children: [new docx.TextRun(line)],
            spacing: { after: 120 }
          }));
          extractedCharacters += line.length;
        }
      }

      if (pageNumber < pdf.numPages) {
        paragraphs.push(new docx.Paragraph({
          children: [new docx.TextRun("")],
          pageBreakBefore: true
        }));
      }
    }
  } finally {
    if (ocrWorker) await ocrWorker.terminate();
  }

  if (!extractedCharacters) {
    throw new Error("No readable text was found. This PDF may be blank, damaged, or too blurry for OCR.");
  }

  setProgress(93, "Building Word document…");
  const documentFile = new docx.Document({
    sections: [{ properties: {}, children: paragraphs }]
  });
  convertedBlob = await docx.Packer.toBlob(documentFile);
}

async function convertDocxToPdf(file) {
  if (!window.mammoth || !window.html2pdf) {
    throw new Error("A required library failed to load. Refresh the page and try again.");
  }

  setProgress(10, "Reading Word document…");
  const arrayBuffer = await file.arrayBuffer();
  const resultHtml = await mammoth.convertToHtml({ arrayBuffer });

  if (!resultHtml.value || !resultHtml.value.trim()) {
    throw new Error("No readable content was found in this DOCX file.");
  }

  setProgress(35, "Preparing PDF layout…");
  const container = document.createElement("div");
  Object.assign(container.style, {
    position: "fixed", left: "-10000px", top: "0",
    width: "720px", padding: "30px", background: "#fff",
    color: "#111", fontFamily: "Arial, sans-serif",
    fontSize: "12px", lineHeight: "1.5", overflowWrap: "anywhere"
  });
  container.innerHTML = resultHtml.value;

  container.querySelectorAll("h1").forEach(el => {
    el.style.fontSize = "24px";
    el.style.margin = "0 0 16px";
  });
  container.querySelectorAll("h2").forEach(el => {
    el.style.fontSize = "19px";
    el.style.margin = "14px 0 10px";
  });
  container.querySelectorAll("p").forEach(el => {
    el.style.margin = "0 0 10px";
  });
  container.querySelectorAll("table").forEach(table => {
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.style.marginBottom = "12px";
  });
  container.querySelectorAll("td, th").forEach(cell => {
    cell.style.border = "1px solid #ccc";
    cell.style.padding = "5px";
  });
  container.querySelectorAll("img").forEach(img => {
    if (!img.src.startsWith("data:image/")) img.remove();
    else {
      img.style.maxWidth = "100%";
      img.style.height = "auto";
    }
  });

  document.body.appendChild(container);
  try {
    setProgress(55, "Rendering PDF…");
    convertedBlob = await html2pdf().set({
      margin: [10, 10, 10, 10],
      image: { type: "jpeg", quality: 0.95 },
      html2canvas: { scale: 1.5, useCORS: false, backgroundColor: "#fff", logging: false },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] }
    }).from(container).outputPdf("blob");
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
