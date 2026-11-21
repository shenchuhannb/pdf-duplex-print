# -*- coding: utf-8 -*-
"""图片 → 临时 PDF（jpg/png/bmp/gif/tif/webp …）。

规则（与界面说明一致）：
  * 每张图片 = 1 个 PDF 页；多页 TIFF 逐帧拆成多页；
  * GIF/动态图只取第一帧；
  * 透明背景合成到白底；按 EXIF 方向自动转正；
  * 页面尺寸按图片像素与内嵌 DPI 换算成点（96 DPI 时 1 像素≈0.75 点），
    后续仍由现有 PDF 排版/补页/双面逻辑统一处理。
"""

from __future__ import annotations

import os
from typing import Dict, List

import fitz
from PIL import Image, ImageOps

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp")
FIRST_FRAME_ONLY = (".gif", ".webp")


class ImageConvertError(Exception):
    """图片转换失败；message 为可直接展示的中文提示。"""


def is_image_path(path: str) -> bool:
    return str(path).strip().lower().endswith(IMAGE_EXTENSIONS)


def supported_extensions_text() -> str:
    return "JPG/JPEG、PNG、BMP、GIF、TIF/TIFF、WEBP"


def _flatten(image: Image.Image) -> Image.Image:
    """统一成 RGB：透明背景合成白底，避免打印出黑底。"""
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        background = Image.new("RGB", rgba.size, (255, 255, 255))
        background.paste(rgba, mask=rgba.split()[-1])
        return background
    return image.convert("RGB")


def _page_points(image: Image.Image) -> tuple:
    """按图片像素与内嵌 DPI 计算页面尺寸（点）。"""
    dpi = image.info.get("dpi") or (96.0, 96.0)
    try:
        dpi_x, dpi_y = float(dpi[0]), float(dpi[1])
    except (TypeError, ValueError, IndexError):
        dpi_x = dpi_y = 96.0
    dpi_x = dpi_x if dpi_x > 1 else 96.0
    dpi_y = dpi_y if dpi_y > 1 else 96.0
    width = max(1.0, image.width * 72.0 / dpi_x)
    height = max(1.0, image.height * 72.0 / dpi_y)
    # 防止异常元数据生成超大页面：最长边限制在 2000 点（约 70cm）
    longest = max(width, height)
    if longest > 2000.0:
        factor = 2000.0 / longest
        width, height = width * factor, height * factor
    return width, height


def image_to_pdf(src: str, dst: str) -> Dict[str, object]:
    """把图片文件转换成 PDF，返回 {"pages", "converter", "first_frame_only"}。"""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise ImageConvertError(f"找不到图片文件：{src}")
    try:
        image = Image.open(src)
    except Exception as exc:
        raise ImageConvertError(f"无法读取图片：{os.path.basename(src)}（{exc}）") from exc

    first_only = src.lower().endswith(FIRST_FRAME_ONLY)
    frames = 1 if first_only else max(1, int(getattr(image, "n_frames", 1)))
    frame_files: List[str] = []
    document = fitz.open()
    try:
        for index in range(frames):
            try:
                image.seek(index)
            except EOFError:
                break
            frame = ImageOps.exif_transpose(image)
            rgb = _flatten(frame)
            width_pt, height_pt = _page_points(frame)
            page = document.new_page(width=width_pt, height=height_pt)
            frame_path = os.path.join(
                os.path.dirname(dst) or ".",
                f".frame_{os.getpid()}_{index}_{os.path.basename(src)}.png")
            rgb.save(frame_path, "PNG")
            frame_files.append(frame_path)
            page.insert_image(page.rect, filename=frame_path, keep_proportion=False)
        if document.page_count <= 0:
            raise ImageConvertError("图片没有可转换的帧。")
        document.save(dst, deflate=True)
        return {"pages": document.page_count,
                "converter": "图片",
                "first_frame_only": first_only}
    except ImageConvertError:
        raise
    except Exception as exc:
        raise ImageConvertError(f"图片转换失败：{exc}") from exc
    finally:
        document.close()
        for path in frame_files:
            try:
                os.remove(path)
            except OSError:
                pass
        try:
            image.close()
        except Exception:
            pass
