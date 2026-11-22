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

import math
import os
import tempfile
import time
from typing import Dict

import fitz
from PIL import Image, ImageOps

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp")
FIRST_FRAME_ONLY = (".gif", ".webp")
MAX_IMAGE_PIXELS = 50_000_000
MAX_IMAGE_TOTAL_PIXELS = 100_000_000
MAX_IMAGE_FRAMES = 200
IMAGE_CONVERT_TIMEOUT_SECONDS = 60.0


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
    dpi_x = dpi_x if math.isfinite(dpi_x) and dpi_x > 1 else 96.0
    dpi_y = dpi_y if math.isfinite(dpi_y) and dpi_y > 1 else 96.0
    width = max(1.0, image.width * 72.0 / dpi_x)
    height = max(1.0, image.height * 72.0 / dpi_y)
    # 防止异常元数据生成超大页面：最长边限制在 2000 点（约 70cm）
    longest = max(width, height)
    if longest > 2000.0:
        factor = 2000.0 / longest
        width, height = width * factor, height * factor
    return width, height


def image_to_pdf(src: str, dst: str) -> Dict[str, object]:
    """把图片转换成 PDF；成功后才原子替换目标，避免半成品和并发冲突。"""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise ImageConvertError(f"找不到图片文件：{src}")
    if os.path.normcase(src) == os.path.normcase(dst):
        raise ImageConvertError("源图片与输出 PDF 不能是同一路径。")
    try:
        image = Image.open(src)
    except Exception as exc:
        raise ImageConvertError(f"无法读取图片：{os.path.basename(src)}（{exc}）") from exc

    first_only = src.lower().endswith(FIRST_FRAME_ONLY)
    frames = 1 if first_only else max(1, int(getattr(image, "n_frames", 1)))
    if frames > MAX_IMAGE_FRAMES:
        image.close()
        raise ImageConvertError(f"图片帧数超过安全上限（{MAX_IMAGE_FRAMES} 帧）。")
    document = fitz.open()
    started = time.monotonic()
    total_pixels = 0
    folder = os.path.dirname(dst) or "."
    os.makedirs(folder, exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix=".image_frames_", dir=folder) as workdir:
            for index in range(frames):
                if time.monotonic() - started > IMAGE_CONVERT_TIMEOUT_SECONDS:
                    raise ImageConvertError("图片转换超时，请拆分文件或降低图片规模后重试。")
                try:
                    image.seek(index)
                except EOFError:
                    break
                pixels = image.width * image.height
                total_pixels += pixels
                if pixels > MAX_IMAGE_PIXELS or total_pixels > MAX_IMAGE_TOTAL_PIXELS:
                    raise ImageConvertError("图片像素总量超过安全上限，请缩小图片后重试。")
                frame = ImageOps.exif_transpose(image)
                rgb = _flatten(frame)
                width_pt, height_pt = _page_points(frame)
                page = document.new_page(width=width_pt, height=height_pt)
                frame_path = os.path.join(workdir, f"frame_{index}.png")
                rgb.save(frame_path, "PNG")
                page.insert_image(page.rect, filename=frame_path, keep_proportion=False)
            if document.page_count <= 0:
                raise ImageConvertError("图片没有可转换的帧。")
            temp_pdf = os.path.join(workdir, "converted.pdf")
            document.save(temp_pdf, deflate=True)
            page_count = document.page_count
            os.replace(temp_pdf, dst)
        return {"pages": page_count,
                "converter": "图片",
                "first_frame_only": first_only}
    except ImageConvertError:
        raise
    except Exception as exc:
        raise ImageConvertError(f"图片转换失败：{exc}") from exc
    finally:
        if document is not None:
            document.close()
        try:
            image.close()
        except Exception:
            pass
