from __future__ import annotations

import base64
import hashlib
import json
import re
import tempfile
from pathlib import Path

import streamlit as st

from processor import process_docx


ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "assets" / "grading-sheet-template.docx"
MAX_FILE_SIZE = 25 * 1024 * 1024
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def output_filename(name: str) -> str:
    stem = Path(name).stem
    safe_stem = re.sub(r"[^\w.-]+", "_", stem, flags=re.UNICODE).strip("._") or "phieu_cham"
    return f"{safe_stem}_phieu_cham_rut_gon_MaiAnhTuan_TUMP.docx"


def process_upload(filename: str, payload: bytes) -> tuple[dict, bytes | None]:
    with tempfile.TemporaryDirectory(prefix="phieu-cham-") as temp_dir:
        workdir = Path(temp_dir)
        input_path = workdir / Path(filename).name
        output_path = workdir / output_filename(filename)
        input_path.write_bytes(payload)
        report = process_docx(str(input_path), str(output_path), str(TEMPLATE_PATH))
        output_bytes = output_path.read_bytes() if output_path.is_file() and not report["errors"] else None
    return report, output_bytes


def render_docx_preview(document_bytes: bytes) -> None:
    encoded = base64.b64encode(document_bytes).decode("ascii")
    encoded_json = json.dumps(encoded)
    preview_html = f"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/docx-preview@0.3.6/dist/docx-preview.min.js"></script>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; color: #12373b; background: #e8edeb; font-family: Inter, system-ui, sans-serif; }}
    .toolbar {{ position: sticky; top: 0; z-index: 5; display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; border-bottom: 1px solid #d4ddda; background: #f8faf8; font-size: 12px; font-weight: 800; }}
    .controls {{ display: flex; align-items: center; gap: 8px; }}
    button {{ width: 30px; height: 30px; border: 1px solid #cbd8d4; border-radius: 8px; color: #075d65; background: white; cursor: pointer; font-size: 16px; font-weight: 900; }}
    button:disabled {{ cursor: not-allowed; opacity: .4; }}
    #zoom-label {{ min-width: 42px; text-align: center; }}
    #viewport {{ height: 700px; overflow: auto; padding: 18px; }}
    #pages {{ width: max-content; min-width: 100%; transform-origin: top left; }}
    #pages .docx-wrapper {{ padding: 0 !important; background: transparent !important; }}
    #pages section.docx {{ margin: 0 auto 18px !important; box-shadow: 0 8px 24px rgba(28, 59, 58, .18) !important; }}
    #error {{ display: none; margin: 80px auto; max-width: 420px; padding: 16px; border-radius: 12px; color: #a23c3c; background: #fff0f0; text-align: center; line-height: 1.5; }}
  </style>
</head>
<body>
  <div class="toolbar">
    <span id="page-count">Đang dựng trang…</span>
    <div class="controls">
      <button id="zoom-out" type="button" aria-label="Thu nhỏ">−</button>
      <span id="zoom-label">65%</span>
      <button id="zoom-in" type="button" aria-label="Phóng to">+</button>
    </div>
  </div>
  <div id="viewport">
    <div id="pages"></div>
    <p id="error"></p>
  </div>
  <script>
    const encoded = {encoded_json};
    const bytes = Uint8Array.from(atob(encoded), character => character.charCodeAt(0));
    const pages = document.getElementById("pages");
    const pageCount = document.getElementById("page-count");
    const zoomLabel = document.getElementById("zoom-label");
    const zoomOut = document.getElementById("zoom-out");
    const zoomIn = document.getElementById("zoom-in");
    let zoom = 0.65;

    function applyZoom() {{
      pages.style.zoom = String(zoom);
      zoomLabel.textContent = `${{Math.round(zoom * 100)}}%`;
      zoomOut.disabled = zoom <= 0.45;
      zoomIn.disabled = zoom >= 1.25;
    }}

    function changeZoom(delta) {{
      zoom = Math.min(1.25, Math.max(0.45, Math.round((zoom + delta) * 100) / 100));
      applyZoom();
    }}

    zoomOut.addEventListener("click", () => changeZoom(-0.1));
    zoomIn.addEventListener("click", () => changeZoom(0.1));
    applyZoom();

    docx.renderAsync(bytes, pages, pages, {{
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      useBase64URL: true,
      experimental: false,
      debug: false,
    }}).then(() => {{
      const count = pages.querySelectorAll("section.docx").length || 1;
      pageCount.textContent = `${{count}} trang · DOCX`;
    }}).catch(error => {{
      pages.replaceChildren();
      const errorBox = document.getElementById("error");
      errorBox.style.display = "block";
      errorBox.textContent = error.message || "Không thể dựng bản xem trước DOCX.";
      pageCount.textContent = "Lỗi preview";
    }});
  </script>
</body>
</html>
"""
    st.iframe(preview_html, width="stretch", height=760)


def render_report(report: dict) -> None:
    metric_columns = st.columns(4)
    metric_columns[0].metric("Câu hỏi", report["questions_found"])
    metric_columns[1].metric("Dòng 0,25", report["score_rows_written"])
    metric_columns[2].metric("Đã rút gọn", report["cells_shortened"])
    metric_columns[3].metric("Dấu … tự nhiên", report["natural_ellipsis_processed"])

    st.markdown("#### Kiểm tra chất lượng")
    st.success(f"{report['score_rows_found']} dòng 0,25 được nhận diện từ file nguồn.")
    st.success(f"{report['score_rows_written']} dòng được ghi vào phiếu chấm.")
    st.success(f"{report['cells_shortened']} ô được rút gọn theo quy tắc 5 + 3.")
    if report["natural_ellipsis_processed"]:
        st.success(f"{report['natural_ellipsis_processed']} ô có dấu ba chấm tự nhiên vẫn được xử lý.")
    for warning in report["warnings"]:
        st.warning(warning)
    for error in report["errors"]:
        st.error(error)


st.set_page_config(
    page_title="shorten-docx-score-sheets-maianhtuan-tump",
    page_icon="📝",
    layout="wide",
)

st.markdown(
    """
<style>
  .stApp { background: radial-gradient(circle at 8% 8%, rgba(185,217,76,.16), transparent 28rem), #f5f3ea; }
  .block-container { max-width: 1480px; padding-top: 2.2rem; }
  .tool-hero h1 { margin: 0; color: #12373b; font-size: clamp(2.1rem, 5vw, 4.6rem); line-height: .98; letter-spacing: -.05em; overflow-wrap: anywhere; }
  .tool-hero h1 span { color: #087f83; }
  .tool-hero p { margin: 1rem 0 .35rem; color: #456467; font-size: 1.05rem; }
  .tool-hero small { color: #667d7f; }
  div[data-testid="stFileUploader"] { padding: 1rem; border: 1px solid #dce5df; border-radius: 1rem; background: rgba(255,254,249,.9); }
  div[data-testid="stMetric"] { padding: .75rem; border: 1px solid #dce5df; border-radius: .8rem; background: white; }
  .privacy-note { margin-top: .6rem; padding: .75rem .9rem; border-radius: .75rem; color: #76561a; background: #fff4d9; font-size: .82rem; }
</style>
<div class="tool-hero">
  <h1>shorten-docx-score-sheets-<span>maianhtuan-tump</span></h1>
  <p>Chuyển file đáp án thành phiếu chấm rút gọn, kiểm tra QC và xem trước theo trang.</p>
  <small><strong>Mai Anh Tuấn TUMP</strong> · Social Medicine Public Health</small>
</div>
""",
    unsafe_allow_html=True,
)

st.divider()
input_column, review_column = st.columns([0.82, 1.18], gap="large")

with input_column:
    st.subheader("01 · Chọn file đáp án")
    uploaded_file = st.file_uploader("File DOCX", type=["docx"], help="Tối đa 25 MB")
    st.markdown(
        '<div class="privacy-note">Bản Streamlit gửi file tới máy chủ Streamlit để xử lý tạm thời. '
        "Không dùng bản này cho tài liệu nhạy cảm nếu chưa chấp nhận việc truyền file lên cloud.</div>",
        unsafe_allow_html=True,
    )

    if uploaded_file is not None:
        input_bytes = uploaded_file.getvalue()
        fingerprint = hashlib.sha256(input_bytes).hexdigest()
        current_fingerprint = st.session_state.get("input_fingerprint")
        if fingerprint != current_fingerprint:
            st.session_state.input_fingerprint = fingerprint
            st.session_state.pop("output_bytes", None)
            st.session_state.pop("report", None)
            st.session_state.pop("output_name", None)

        if len(input_bytes) > MAX_FILE_SIZE:
            st.error("File lớn hơn giới hạn 25 MB.")
        elif not input_bytes:
            st.error("File đang rỗng.")
        elif st.button("Tạo phiếu chấm →", type="primary", use_container_width=True):
            with st.spinner("Đang dựng phiếu chấm và kiểm tra QC…"):
                report, output_bytes = process_upload(uploaded_file.name, input_bytes)
                st.session_state.report = report
                st.session_state.output_bytes = output_bytes
                st.session_state.output_name = output_filename(uploaded_file.name)

    output_bytes = st.session_state.get("output_bytes")
    report = st.session_state.get("report")
    output_name = st.session_state.get("output_name")

    if output_bytes and output_name:
        st.download_button(
            "Tải file DOCX",
            data=output_bytes,
            file_name=output_name,
            mime=DOCX_MIME,
            type="primary",
            use_container_width=True,
        )

with review_column:
    st.subheader("02 · Review kết quả")
    if report:
        render_report(report)
        if output_bytes:
            st.markdown("#### Xem trước phiếu chấm")
            render_docx_preview(output_bytes)
    else:
        st.info("Chọn file DOCX và nhấn “Tạo phiếu chấm” để xem QC và bản xem trước.")

st.divider()
st.caption("Mai Anh Tuấn TUMP · Social Medicine Public Health")
