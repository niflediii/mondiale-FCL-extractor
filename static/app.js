const fileInput = document.getElementById("file-input");
const dropzone = document.getElementById("dropzone");
const processButton = document.getElementById("process-button");
const fileName = document.getElementById("file-name");
const statusText = document.getElementById("status-text");
const uploadForm = document.getElementById("upload-form");
const loadingOverlay = document.getElementById("loading-overlay");
const loadingText = document.getElementById("loading-text");

let selectedFile = null;

function setSelectedFile(file) {
  selectedFile = file ?? null;
  processButton.disabled = !selectedFile;
  processButton.classList.toggle("is-ready", Boolean(selectedFile));
  fileName.textContent = selectedFile ? selectedFile.name : "Select `.pdf` file";
  statusText.textContent = selectedFile
    ? "File is ready. Process it to download the FCL Rates workbook."
    : "100% it automatically download the result.";
}

function validateFile(file) {
  return file && file.name.toLowerCase().endsWith(".pdf");
}

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (!validateFile(file)) {
    setSelectedFile(null);
    statusText.textContent = "Please choose the Mondiale FCL quotation `.pdf`.";
    return;
  }

  setSelectedFile(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (!validateFile(file)) {
    setSelectedFile(null);
    statusText.textContent = "Only `.pdf` files are supported right now.";
    return;
  }

  fileInput.files = event.dataTransfer.files;
  setSelectedFile(file);
});

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!selectedFile) {
    statusText.textContent = "Please select a quotation PDF first.";
    return;
  }

  processButton.disabled = true;
  processButton.classList.remove("is-ready");
  processButton.textContent = "Processing...";
  statusText.textContent = "Reading the quotation and mapping every rate row...";
  loadingOverlay.hidden = false;
  loadingText.textContent = "Preparing the workbook and your download...";

  try {
    const formData = new FormData();
    formData.append("file", selectedFile);

    const response = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const errorPayload = await response.json().catch(() => ({ error: "Processing failed." }));
      throw new Error(errorPayload.error || "Processing failed.");
    }

    const blob = await response.blob();
    loadingText.textContent = "Finalizing the file and starting your download...";
    const contentDisposition = response.headers.get("Content-Disposition") || "";
    // Flask only quotes the filename when it has to, so accept both forms.
    const downloadNameMatch = contentDisposition.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
    const downloadName = downloadNameMatch ? downloadNameMatch[1].trim() : "Mondiale_FCL.xlsx";

    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(downloadUrl);

    statusText.textContent = "Done. Your FCL Rates workbook has been downloaded.";
  } catch (error) {
    statusText.textContent = error.message;
  } finally {
    loadingOverlay.hidden = true;
    processButton.textContent = "Process File";
    processButton.disabled = !selectedFile;
    processButton.classList.toggle("is-ready", Boolean(selectedFile));
  }
});
