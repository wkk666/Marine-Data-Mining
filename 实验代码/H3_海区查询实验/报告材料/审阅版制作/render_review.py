"""复用打包DOCX渲染器的PDF→PNG流程；PDF由原生Word提前导出。

本脚本不启动Word、不调用soffice、不修改DOCX或PDF。
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import tempfile

HERE = Path(__file__).resolve().parent
REPORT_ROOT = HERE.parent
DEFAULT_RENDERER = Path(r"C:\Users\wkc\.codex\plugins\cache\openai-primary-runtime\documents\26.1110.11613\skills\documents\render_docx.py")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", type=Path, default=REPORT_ROOT / "实验报告_审阅版.docx")
    parser.add_argument("--pdf", type=Path, default=REPORT_ROOT / "实验报告_审阅版.pdf")
    parser.add_argument("--output-dir", type=Path, default=REPORT_ROOT / "审阅版预览")
    parser.add_argument("--renderer", type=Path, default=DEFAULT_RENDERER)
    parser.add_argument("--dpi", type=int, default=140)
    args = parser.parse_args()
    docx, pdf, out, renderer = [p.resolve() for p in (args.docx, args.pdf, args.output_dir, args.renderer)]
    for source in (docx, pdf, renderer):
        if not source.is_file() or source.stat().st_size == 0:
            parser.error(f"Missing or empty input: {source}")
    if args.dpi < 96 or args.dpi > 300:
        parser.error("Use 96–300 dpi for the review PNGs.")
    try:
        out.relative_to(REPORT_ROOT.resolve())
    except ValueError:
        parser.error("Preview directory must stay inside the report-material directory.")
    if out == REPORT_ROOT.resolve():
        parser.error("Use a dedicated preview subdirectory, not the report root.")
    out.mkdir(parents=True, exist_ok=True)
    temp_root = HERE / ".tmp"
    temp_root.mkdir(exist_ok=True)
    os.environ["TEMP"] = str(temp_root)
    os.environ["TMP"] = str(temp_root)
    tempfile.tempdir = str(temp_root)

    before = {"docx_sha256": digest(docx), "pdf_sha256": digest(pdf)}
    word_info_path = out / "Word分页.json"
    if word_info_path.is_file():
        word_info = json.loads(word_info_path.read_text(encoding="utf-8-sig"))
        if word_info.get("docx_sha256") != before["docx_sha256"] or word_info.get("pdf_sha256") != before["pdf_sha256"]:
            parser.error("DOCX/PDF no longer match Word分页.json; export Word again before rendering.")
    else:
        word_info = None
    # Only remove old generated page-N PNGs inside this checked preview directory.
    # The complete source report and experimental records are never touched.
    for old in out.glob("page-*.png"):
        if re.fullmatch(r"page-\d+\.png", old.name) and old.resolve().parent == out:
            old.unlink()

    spec = importlib.util.spec_from_file_location("codex_packaged_docx_renderer", renderer)
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)

    def existing_word_pdf(doc_path, user_profile, convert_tmp_dir, stem, verbose):
        # Exact packaged convert_to_pdf signature, bypassing unavailable soffice.
        return str(pdf), "Native Microsoft Word COM already exported the PDF."

    helper.convert_to_pdf = existing_word_pdf
    render_input, repair_temp = helper.make_renderable_docx_copy(str(docx), verbose=False)
    try:
        pages = helper.rasterize(render_input, str(out), args.dpi, verbose=False, emit_pdf=False)
    finally:
        if repair_temp is not None:
            repair_temp.cleanup()
    assert digest(docx) == before["docx_sha256"] and digest(pdf) == before["pdf_sha256"], "Rendering changed an input."
    if word_info is not None:
        assert len(pages) == word_info["physical_pages"], "PNG count disagrees with native Word page count."
    info = {"status": "rendered_existing_word_pdf", "dpi": args.dpi,
            "renderer": str(renderer), "renderer_sha256": digest(renderer),
            "conversion": "Native Word PDF supplied to packaged renderer.rasterize; no soffice",
            "docx": str(docx), "pdf": str(pdf), **before,
            "pages": [{"physical_page": i, "png": str(p), "sha256": digest(Path(p))} for i, p in enumerate(pages, 1)],
            "docx_pdf_unchanged": True}
    (out / "页面渲染.json").write_text(json.dumps(info, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": info["status"], "pages": len(pages), "dpi": args.dpi, "output": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
