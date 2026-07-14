from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]


def test_required_static_files_exist():
    for relative in (
        "index.html",
        "styles.css",
        "app.js",
        "processor.py",
        "assets/grading-sheet-template.docx",
    ):
        assert (ROOT / relative).is_file(), relative


def test_interface_contains_requested_content():
    html = (ROOT / "index.html").read_text(encoding="utf-8")
    assert "shorten-docx-score-sheets-maianhtuan-tump" in html
    assert "Mai Anh Tuấn TUMP" in html
    assert "Social Medicine Public Health" in html
    assert 'id="file-input"' in html
    assert 'id="generate-button"' in html
    assert 'id="review-content"' in html
    assert 'id="download-link"' in html
    assert "DOCX GRADING SHEET STUDIO" not in html
    assert 'class="brand-row"' not in html
    assert "jszip@3.10.1" in html
    assert "docx-preview@0.3.6" in html
    assert 'src="./app.js?v=20260714-3"' in html
    assert 'id="docx-preview-pages"' in html
    assert 'id="zoom-out"' in html
    assert 'id="zoom-in"' in html


def test_browser_runtime_sets_pyodide_file_and_cache_busts_processor():
    javascript = (ROOT / "app.js").read_text(encoding="utf-8")
    assert 'const ASSET_VERSION = "20260714-3"' in javascript
    assert 'fetch(`./processor.py?v=${ASSET_VERSION}`, { cache: "no-store" })' in javascript
    assert 'pyodideRuntime.globals.set("__file__", "/processor.py")' in javascript
    assert "globalThis.docx.renderAsync" in javascript
    assert "outputBytes.slice()" in javascript


def test_template_is_a_readable_docx():
    document = Document(ROOT / "assets" / "grading-sheet-template.docx")
    assert len(document.tables) == 7


def test_processor_compiles():
    source = (ROOT / "processor.py").read_text(encoding="utf-8")
    compile(source, "processor.py", "exec")


def test_processor_loads_without_dunder_file_like_pyodide():
    source = (ROOT / "processor.py").read_text(encoding="utf-8")
    library_source = source.split('\nif __name__ == "__main__":')[0]
    namespace = {"__name__": "__main__"}

    exec(compile(library_source, "processor.py", "exec"), namespace)

    assert namespace["_SCRIPT_DIR"] == Path.cwd()
    assert callable(namespace["process_docx"])
