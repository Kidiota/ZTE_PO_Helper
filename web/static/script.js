const i18n = {
  zh: {
    title: "PO 解析与导出工具",
    subtitle: "上传多个 PO PDF，等待处理完成后下载结果",
    uploadPDF: "点击或拖拽上传 PDF 文件（可多选）",
    start: "开始处理",
    wait: "等待开始",
    download: "下载结果",
    noFiles: "请先选择 PDF 文件"
  },
  en: {
    title: "PO Extraction & Export Tool",
    subtitle: "Upload multiple PO PDFs, wait and download the Excel result",
    uploadPDF: "Click or drag to upload PDF files (multiple allowed)",
    start: "Start Processing",
    wait: "Waiting to start",
    download: "Download Result",
    noFiles: "Please select PDF files first"
  }
};

let lang = "zh";
const pdfArea = document.getElementById("pdfUploadArea");
const pdfInput = document.getElementById("pdfInput");
const fileListDiv = document.getElementById("fileList");
const startBtn = document.getElementById("startBtn");
const progressBar = document.getElementById("progressBar");
const statusText = document.getElementById("statusText");
const resultArea = document.getElementById("resultArea");
const downloadLink = document.getElementById("downloadLink");
const langToggle = document.getElementById("langToggle");

let selectedPDFs = [];

function applyLang() {
  document.querySelector("[data-i18n='title']").textContent = i18n[lang].title;
  document.querySelector("[data-i18n='subtitle']").textContent = i18n[lang].subtitle;
  pdfArea.querySelector("[data-i18n='uploadPDF']").textContent = i18n[lang].uploadPDF;
  startBtn.textContent = i18n[lang].start;
  statusText.textContent = i18n[lang].wait;
  downloadLink.textContent = i18n[lang].download;
  langToggle.textContent = lang === "zh" ? "EN" : "中文";
}
applyLang();

langToggle.addEventListener("click", () => {
  lang = (lang === "zh") ? "en" : "zh";
  applyLang();
});

function updateFileList() {
  fileListDiv.innerHTML = "";
  if (selectedPDFs.length > 0) {
    selectedPDFs.forEach(f => { fileListDiv.innerHTML += `<p><i class="fas fa-file-pdf"></i> ${f.name}</p>`; });
    fileListDiv.style.display = "block";
    startBtn.disabled = false;
  } else {
    fileListDiv.style.display = "none";
    startBtn.disabled = true;
  }
}

pdfArea.addEventListener("click", () => pdfInput.click());
pdfArea.addEventListener("dragover", e => { e.preventDefault(); pdfArea.classList.add("dragover"); });
pdfArea.addEventListener("dragleave", () => pdfArea.classList.remove("dragover"));
pdfArea.addEventListener("drop", e => {
  e.preventDefault();
  pdfArea.classList.remove("dragover");
  const files = Array.from(e.dataTransfer.files).filter(f => f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf"));
  selectedPDFs = selectedPDFs.concat(files);
  updateFileList();
});
pdfInput.addEventListener("change", e => {
  selectedPDFs = Array.from(e.target.files);
  updateFileList();
});

startBtn.addEventListener("click", () => {
  if (!selectedPDFs.length) { alert(i18n[lang].noFiles); return; }

  const formData = new FormData();
  selectedPDFs.forEach(f => formData.append("pdfs", f));

  progressBar.style.width = "0%";
  statusText.textContent = i18n[lang].wait;
  resultArea.style.display = "none";
  downloadLink.href = "#";

  fetch("/start_process", { method: "POST", body: formData })
    .then(res => res.json())
    .then(data => {
      if (data.error) { alert(data.error); return; }
      pollProgress(data.task_id);
    })
    .catch(err => { alert("上传失败：" + err); });
});

function pollProgress(task_id) {
  const interval = setInterval(() => {
    fetch(`/progress?task_id=${encodeURIComponent(task_id)}`)
      .then(r => r.json())
      .then(data => {
        if (data.error) {
          clearInterval(interval);
          statusText.textContent = data.error;
          return;
        }
        const p = data.percent || 0;
        progressBar.style.width = p + "%";
        statusText.textContent = data.message || "";

        if (data.finished) {
          clearInterval(interval);
          if (data.files && data.files.length > 0) {
            const fname = data.files[0];
            downloadLink.href = `/download_file/${encodeURIComponent(fname)}`;
            resultArea.style.display = "block";
            statusText.textContent = i18n[lang].download;
          }
        }
      })
      .catch(err => {
        clearInterval(interval);
        statusText.textContent = "获取进度失败";
      });
  }, 1000);
}
