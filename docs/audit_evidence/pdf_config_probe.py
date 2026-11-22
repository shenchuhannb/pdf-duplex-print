# -*- coding: utf-8 -*-
"""CodeArts 审计探针：PDF/图片/配置，仅使用临时目录，不调用打印或 Office。"""
import json
import pathlib
import sys
import tempfile

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

# 英文 Windows（如 GitHub Actions runner）控制台是 cp1252，ensure_ascii=False 的
# 中文结果会触发 UnicodeEncodeError；统一强制 UTF-8 输出，任何代码页下都可运行。
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

import fitz
from PIL import Image
from image_convert import image_to_pdf
from core_utils import PdfError
from pdf_layout import LayoutOptions, compose_layout_pdf, crop_rect_for_page
from settings_store import load_settings

MM = 72.0 / 25.4
results = {}


def rect(value):
    return [round(value.x0, 2), round(value.y0, 2), round(value.x1, 2), round(value.y1, 2)]


with tempfile.TemporaryDirectory(prefix="pdf-duplex-audit-") as temp:
    root = pathlib.Path(temp)
    normal_src = root / "normal.pdf"
    doc = fitz.open()
    for number in range(1, 4):
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), f"UNIQUE_PAGE_{number}")
    doc.save(normal_src)
    doc.close()
    normal_out = root / "normal-out.pdf"
    composed = compose_layout_pdf(
        str(normal_src), [1, 2, 3], str(normal_out), LayoutOptions(manual_duplex=True))
    doc = fitz.open(normal_out)
    output_text = "\n".join(page.get_text() for page in doc)
    results["normal_three_page_duplex"] = {
        "output_pages": doc.page_count,
        "selected_count": composed.selected_count,
        "all_source_markers_preserved": all(
            f"UNIQUE_PAGE_{number}" in output_text for number in range(1, 4)),
    }
    doc.close()

    annot_src = root / "annot.pdf"
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 72), "ANNOT_CONTENT")
    page.add_text_annot((100, 100), "AUDIT_NOTE")
    doc.save(annot_src)
    source_annots = sum(1 for _ in doc[0].annots())
    doc.close()
    annot_out = root / "annot-out.pdf"
    compose_layout_pdf(str(annot_src), [1], str(annot_out),
                       LayoutOptions(manual_duplex=False, annots=True))
    plain_out = root / "plain-out.pdf"
    compose_layout_pdf(str(annot_src), [1], str(plain_out),
                       LayoutOptions(manual_duplex=False, annots=False))
    doc = fitz.open(annot_out)
    plain = fitz.open(plain_out)
    results["annotations_enabled"] = {
        "source_annots": source_annots,
        "output_interactive_annots": sum(1 for page in doc for _ in (page.annots() or [])),
        "content_preserved": "ANNOT_CONTENT" in doc[0].get_text(),
        "visible_annotation_flattened": doc[0].get_pixmap().samples != plain[0].get_pixmap().samples,
    }
    doc.close()
    plain.close()

    crop_doc = fitz.open()
    crop_page = crop_doc.new_page(width=595, height=842)
    crop_page.set_mediabox(fitz.Rect(100, 100, 695, 942))
    actual_crop = crop_rect_for_page(crop_page, LayoutOptions(
        crop_left_mm=10, crop_top_mm=10, crop_right_mm=10, crop_bottom_mm=10))
    box = crop_page.cropbox
    results["offset_mediabox_crop"] = {
        "actual": rect(actual_crop),
        "expected": [round(box.x0 + 10 * MM, 2), round(box.y0 + 10 * MM, 2),
                     round(box.x1 - 10 * MM, 2), round(box.y1 - 10 * MM, 2)],
        "matches_cropbox_coordinates": actual_crop == fitz.Rect(
            box.x0 + 10 * MM, box.y0 + 10 * MM, box.x1 - 10 * MM, box.y1 - 10 * MM),
    }
    crop_doc.close()

    invalid_nup_src = root / "invalid-nup.pdf"
    doc = fitz.open()
    for number in range(1, 4):
        page = doc.new_page()
        page.insert_text((72, 72), f"NUP_PAGE_{number}")
    doc.save(invalid_nup_src)
    doc.close()
    invalid_nup_out = root / "invalid-nup-out.pdf"
    try:
        compose_layout_pdf(str(invalid_nup_src), [1, 2, 3], str(invalid_nup_out),
                           LayoutOptions(manual_duplex=False, nup=3))
        rejected = False
    except PdfError as exc:
        rejected = str(exc)
    results["invalid_nup_three"] = {
        "rejected_instead_of_dropping_pages": bool(rejected),
        "message": rejected,
        "ui_guard": "ui_settings.py:1062-1064 falls back to nup=1",
    }

    nan_path = root / "nan-settings.json"
    nan_path.write_text('{"dpi": NaN}', encoding="utf-8")
    try:
        sanitized = load_settings(str(nan_path))
        results["nan_config"] = {"accepted_without_error": True,
                                 "fallback_dpi": sanitized["dpi"]}
    except Exception as exc:
        results["nan_config"] = {
            "accepted_without_error": False,
            "exception": type(exc).__name__, "message": str(exc)}

    bool_path = root / "bool-settings.json"
    bool_path.write_text('{"manual_duplex": "false", "orientation": "diagonal"}',
                         encoding="utf-8")
    loaded = load_settings(str(bool_path))
    results["string_bool_and_enum"] = {
        "manual_duplex": loaded["manual_duplex"], "orientation": loaded["orientation"]}

    rgba_path = root / "透明 图片.png"
    Image.new("RGBA", (16, 8), (37, 99, 235, 128)).save(rgba_path, dpi=(96, 96))
    rgba_pdf = root / "rgba.pdf"
    rgba_result = image_to_pdf(str(rgba_path), str(rgba_pdf))
    doc = fitz.open(rgba_pdf)
    results["rgba_image"] = {
        "result": rgba_result, "pdf_pages": doc.page_count,
        "leftover_frame_files": [p.name for p in root.glob(".frame_*")]}
    doc.close()

    tiff_path = root / "multi.tiff"
    frames = [Image.new("RGB", (10, 10), color) for color in ("red", "blue")]
    frames[0].save(tiff_path, save_all=True, append_images=frames[1:])
    results["multipage_tiff"] = image_to_pdf(str(tiff_path), str(root / "multi.pdf"))

print(json.dumps(results, ensure_ascii=False, indent=2))
