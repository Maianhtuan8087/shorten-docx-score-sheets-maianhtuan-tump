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


def test_template_is_a_readable_docx():
    document = Document(ROOT / "assets" / "grading-sheet-template.docx")
    assert len(document.tables) == 7


def test_processor_compiles():
    source = (ROOT / "processor.py").read_text(encoding="utf-8")
    compile(source, "processor.py", "exec")
