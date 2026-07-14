const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v314.0.2/full/";
const ASSET_VERSION = "20260714-2";
const MAX_FILE_SIZE = 25 * 1024 * 1024;

const elements = {
  input: document.querySelector("#file-input"),
  dropZone: document.querySelector("#drop-zone"),
  fileCard: document.querySelector("#file-card"),
  fileName: document.querySelector("#file-name"),
  fileSize: document.querySelector("#file-size"),
  removeFile: document.querySelector("#remove-file"),
  generate: document.querySelector("#generate-button"),
  status: document.querySelector("#status-text"),
  engine: document.querySelector("#engine-state"),
  progressTrack: document.querySelector("#progress-track"),
  progressBar: document.querySelector("#progress-bar"),
  downloadCard: document.querySelector("#download-card"),
  downloadLink: document.querySelector("#download-link"),
  outputName: document.querySelector("#output-name"),
  emptyReview: document.querySelector("#empty-review"),
  reviewContent: document.querySelector("#review-content"),
  reviewStatus: document.querySelector("#review-status"),
  qcList: document.querySelector("#qc-list"),
  qcSummary: document.querySelector("#qc-summary"),
  previewList: document.querySelector("#preview-list"),
  previewCount: document.querySelector("#preview-count"),
};

let pyodideRuntime = null;
let selectedFile = null;
let downloadUrl = null;

function setEngineState(type, text) {
  elements.engine.className = `engine-state ${type}`;
  elements.engine.querySelector("b").textContent = text;
}

function setStatus(text, error = false) {
  elements.status.textContent = text;
  elements.status.classList.toggle("error", error);
}

function setProgress(value) {
  elements.progressTrack.classList.remove("hidden");
  elements.progressBar.style.width = `${value}%`;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function outputFilename(name) {
  const stem = name.replace(/\.docx$/i, "").replace(/[^\p{L}\p{N}._-]+/gu, "_");
  return `${stem}_phieu_cham_rut_gon_MaiAnhTuan_TUMP.docx`;
}

function validateFile(file) {
  if (!file || !file.name.toLowerCase().endsWith(".docx")) return "Vui lòng chọn đúng file DOCX.";
  if (file.size > MAX_FILE_SIZE) return "File lớn hơn giới hạn 25 MB.";
  if (!file.size) return "File đang rỗng.";
  return null;
}

function selectFile(file) {
  const error = validateFile(file);
  if (error) {
    setStatus(error, true);
    return;
  }
  selectedFile = file;
  elements.fileName.textContent = file.name;
  elements.fileSize.textContent = formatBytes(file.size);
  elements.fileCard.classList.remove("hidden");
  elements.downloadCard.classList.add("hidden");
  elements.generate.disabled = !pyodideRuntime;
  setStatus(pyodideRuntime ? "File hợp lệ. Sẵn sàng tạo phiếu chấm." : "File hợp lệ. Đang chờ bộ xử lý sẵn sàng…");
}

function resetFile() {
  selectedFile = null;
  elements.input.value = "";
  elements.fileCard.classList.add("hidden");
  elements.downloadCard.classList.add("hidden");
  elements.generate.disabled = true;
  setStatus("Chọn một file DOCX để bắt đầu.");
}

async function prepareRuntime() {
  try {
    setEngineState("loading", "Đang chuẩn bị bộ xử lý…");
    pyodideRuntime = await loadPyodide({ indexURL: PYODIDE_INDEX });
    await pyodideRuntime.loadPackage(["micropip", "lxml"]);
    const micropip = pyodideRuntime.pyimport("micropip");
    await micropip.install("python-docx==1.2.0");
    micropip.destroy();

    const processorResponse = await fetch(`./processor.py?v=${ASSET_VERSION}`, { cache: "no-store" });
    if (!processorResponse.ok) throw new Error("Không tải được bộ quy tắc xử lý.");
    const processorSource = await processorResponse.text();
    const librarySource = processorSource.split('\nif __name__ == "__main__":')[0];
    pyodideRuntime.globals.set("__file__", "/processor.py");
    pyodideRuntime.runPython(librarySource);
    pyodideRuntime.runPython(`
def process_for_web(input_path, output_path, template_path):
    report = process_docx(input_path, output_path, template_path)
    previews = []
    if not report["errors"]:
        rendered = Document(output_path)
        for table in rendered.tables:
            if _is_question_table(table) and len(table.rows) >= 3:
                question = normalize_space(table.cell(1, 0).text)
                items = [normalize_space(row.cells[1].text) for row in table.rows[1:-1] if normalize_space(row.cells[1].text)]
                previews.append({"question": question, "items": items})
    return json.dumps({"report": report, "previews": previews}, ensure_ascii=False)
`);
    setEngineState("ready", "Bộ xử lý sẵn sàng");
    elements.generate.disabled = !selectedFile;
    if (selectedFile) setStatus("File hợp lệ. Sẵn sàng tạo phiếu chấm.");
  } catch (error) {
    console.error(error);
    pyodideRuntime = null;
    setEngineState("error", "Không thể khởi tạo");
    setStatus("Không tải được bộ xử lý. Kiểm tra kết nối mạng rồi tải lại trang.", true);
  }
}

function addQcItem(text, type = "success") {
  const item = document.createElement("li");
  item.textContent = text;
  if (type !== "success") item.classList.add(type);
  elements.qcList.appendChild(item);
}

function renderReview(data) {
  const { report, previews } = data;
  elements.emptyReview.classList.add("hidden");
  elements.reviewContent.classList.remove("hidden");
  document.querySelector("#metric-questions").textContent = report.questions_found;
  document.querySelector("#metric-rows").textContent = report.score_rows_written;
  document.querySelector("#metric-shortened").textContent = report.cells_shortened;
  document.querySelector("#metric-ellipsis").textContent = report.natural_ellipsis_processed;

  elements.qcList.replaceChildren();
  addQcItem(`${report.score_rows_found} dòng 0,25 được nhận diện từ file nguồn.`);
  addQcItem(`${report.score_rows_written} dòng được ghi vào phiếu chấm.`);
  addQcItem(`${report.cells_shortened} ô nội dung được rút gọn theo quy tắc 5 + 3.`);
  if (report.natural_ellipsis_processed) {
    addQcItem(`${report.natural_ellipsis_processed} ô có dấu ba chấm tự nhiên vẫn được xử lý.`);
  }
  report.warnings.forEach((warning) => addQcItem(warning, "warning"));
  report.errors.forEach((error) => addQcItem(error, "error"));

  const hasErrors = report.errors.length > 0;
  const hasWarnings = report.warnings.length > 0;
  elements.reviewStatus.className = `review-status ${hasErrors ? "error" : hasWarnings ? "warning" : "success"}`;
  elements.reviewStatus.textContent = hasErrors ? "Có lỗi" : hasWarnings ? "Cần kiểm tra" : "Đạt QC";
  elements.qcSummary.textContent = hasErrors ? `${report.errors.length} lỗi` : hasWarnings ? `${report.warnings.length} cảnh báo` : "Không có cảnh báo";

  elements.previewList.replaceChildren();
  previews.forEach((group) => {
    group.items.forEach((text, index) => {
      const article = document.createElement("article");
      article.className = "preview-item";
      const heading = document.createElement("strong");
      heading.textContent = `${group.question} · Ý ${index + 1}`;
      const paragraph = document.createElement("p");
      paragraph.textContent = text;
      article.append(heading, paragraph);
      elements.previewList.appendChild(article);
    });
  });
  const previewTotal = previews.reduce((total, group) => total + group.items.length, 0);
  elements.previewCount.textContent = `${previewTotal} ý`;
}

async function generateDocument() {
  if (!selectedFile || !pyodideRuntime) return;
  const outputName = outputFilename(selectedFile.name);
  elements.generate.disabled = true;
  elements.generate.classList.add("processing");
  elements.downloadCard.classList.add("hidden");
  setStatus("Đang đọc file đáp án…");
  setProgress(14);

  try {
    const [inputBuffer, templateResponse] = await Promise.all([
      selectedFile.arrayBuffer(),
      fetch("./assets/grading-sheet-template.docx"),
    ]);
    if (!templateResponse.ok) throw new Error("Không tải được template phiếu chấm.");
    const templateBuffer = await templateResponse.arrayBuffer();
    setProgress(38);
    setStatus("Đang rút gọn nội dung và dựng phiếu chấm…");

    ["/input.docx", "/template.docx", "/output.docx", "/output.qc.json"].forEach((path) => {
      try { pyodideRuntime.FS.unlink(path); } catch (_) { /* file chưa tồn tại */ }
    });
    pyodideRuntime.FS.writeFile("/input.docx", new Uint8Array(inputBuffer));
    pyodideRuntime.FS.writeFile("/template.docx", new Uint8Array(templateBuffer));
    setProgress(58);

    const resultJson = pyodideRuntime.runPython('process_for_web("/input.docx", "/output.docx", "/template.docx")');
    const result = JSON.parse(resultJson);
    renderReview(result);
    if (result.report.errors.length) throw new Error(result.report.errors.join(" · "));

    setProgress(88);
    const outputBytes = pyodideRuntime.FS.readFile("/output.docx");
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(new Blob([outputBytes], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }));
    elements.downloadLink.href = downloadUrl;
    elements.downloadLink.download = outputName;
    elements.outputName.textContent = outputName;
    elements.downloadCard.classList.remove("hidden");
    setProgress(100);
    setStatus("Hoàn tất. Kiểm tra review bên phải trước khi tải file.");
    window.setTimeout(() => elements.progressTrack.classList.add("hidden"), 900);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Không thể xử lý file DOCX.", true);
    elements.reviewStatus.className = "review-status error";
    elements.reviewStatus.textContent = "Có lỗi";
  } finally {
    elements.generate.classList.remove("processing");
    elements.generate.disabled = !selectedFile || !pyodideRuntime;
  }
}

elements.input.addEventListener("change", (event) => selectFile(event.target.files[0]));
elements.removeFile.addEventListener("click", resetFile);
elements.generate.addEventListener("click", generateDocument);
elements.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropZone.classList.add("dragover");
});
elements.dropZone.addEventListener("dragleave", () => elements.dropZone.classList.remove("dragover"));
elements.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropZone.classList.remove("dragover");
  selectFile(event.dataTransfer.files[0]);
});

prepareRuntime();
