# -*- coding: utf-8 -*-
"""生成最终图标（风格 A：单张纸 + 折角）与多尺寸 app_icon.ico / png。

运行：D:\\Anaconda3\\python.exe tools\\make_icons.py
输出：
  docs/icon_a.png、docs/icon_a_512.png       256/512px 预览
  docs/icon_a_16.png                        16px 小尺寸可读性预览（放大 8 倍展示）
  docs/icon_a_dark.png                      深底版本
  app_icon.ico                              16/24/32/48/64/256 多尺寸
  docs/app_icon_256.png、docs/app_icon_512.png  备用 PNG

设计原则：极简几何、单一色相、无渐变、16px 使用专属简化稿（不是 256px 直接缩小）。
"""

import io
import os
import struct

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")

ACCENT = (37, 99, 235, 255)        # #2563eb 蓝色强调
ACCENT_LIGHT = (76, 141, 255, 255)  # 深色背景用的亮蓝
PAPER = (255, 255, 255, 255)
PAPER_EDGE = (223, 228, 234, 255)
FOLD = (203, 213, 225, 255)

SS = 4  # 超采样倍数，保证边缘抗锯齿


def _rounded_tile(size: int, color) -> Image.Image:
    """圆角方形底板：任务栏/桌面缩放时轮廓清晰。"""
    work = size * SS
    image = Image.new("RGBA", (work, work), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    radius = int(work * 0.22)
    draw.rounded_rectangle((0, 0, work - 1, work - 1), radius=radius, fill=color)
    return image


def _finish(image: Image.Image, size: int) -> Image.Image:
    return image.resize((size, size), Image.LANCZOS)


def _paper(draw: ImageDraw.ImageDraw, box, fold: int, fill=PAPER, edge=PAPER_EDGE,
           fold_fill=FOLD) -> None:
    """画一张带折角的纸：折角在右上角，是“文档”的最简识别符号。"""
    x0, y0, x1, y1 = box
    cut = fold
    draw.polygon(
        [(x0, y0), (x1 - cut, y0), (x1, y0 + cut), (x1, y1), (x0, y1)],
        fill=fill, outline=edge)
    draw.polygon(
        [(x1 - cut, y0), (x1 - cut, y0 + cut), (x1, y0 + cut)],
        fill=fold_fill, outline=fold_fill)


def draw_icon_a(size: int, dark: bool = False) -> Image.Image:
    """风格 A：单张纸 + 折角（最稳、小尺寸可读性最好）。"""
    tile_color = ACCENT_LIGHT if dark else ACCENT
    image = _rounded_tile(size, tile_color)
    work = size * SS
    draw = ImageDraw.Draw(image)
    simple = size <= 24  # 16/24 使用专属简化稿：去掉文字横线，加大纸面
    if simple:
        box = (int(work * 0.22), int(work * 0.14), int(work * 0.78), int(work * 0.86))
        _paper(draw, box, fold=int(work * 0.20))
    else:
        box = (int(work * 0.25), int(work * 0.16), int(work * 0.75), int(work * 0.84))
        _paper(draw, box, fold=int(work * 0.16))
        # 两条文字线，暗示“文档内容”；16px 下省略
        line_x0 = int(work * 0.33)
        line_x1 = int(work * 0.67)
        for ratio in (0.48, 0.60, 0.72):
            y = int(work * ratio)
            draw.rounded_rectangle((line_x0, y, line_x1, y + max(2, work // 90)),
                                   radius=work // 90, fill=(37, 99, 235, 255))
    return _finish(image, size)


def _bmp_entry(image: Image.Image) -> bytes:
    """把 RGBA 图像编码为 ICO 内部的未压缩 BMP(DIB) 条目。

    Windows 资源管理器只对 256x256 允许 PNG 压缩；16/24/32/48/64 等小尺寸
    必须是未压缩 DIB，否则 Explorer 会静默忽略该图标并回退到目标程序的图标
    （表现为“快捷方式改了 IconLocation 但桌面还是旧图标”）。

    结构：BITMAPINFOHEADER(40B, biHeight=2*h) + BGRA 自底向上 XOR 位图
          + 1bpp AND 掩码（每行 4 字节对齐，alpha 有效时掩码全 0）。
    """
    width, height = image.size
    rgba = image.convert("RGBA")
    pixels = rgba.load()
    xor_rows = []
    for y in range(height - 1, -1, -1):  # DIB 行自底向上存储
        row = bytearray()
        for x in range(width):
            r, g, b, a = pixels[x, y]
            row += bytes((b, g, r, a))
        xor_rows.append(bytes(row))
    xor_bitmap = b"".join(xor_rows)
    mask_row_bytes = ((width + 31) // 32) * 4
    and_mask = b"\x00" * (mask_row_bytes * height)
    header = struct.pack(
        "<IiiHHIIiiII",
        40,                    # biSize
        width,                 # biWidth
        height * 2,            # biHeight = XOR 高 + AND 高
        1,                     # biPlanes
        32,                    # biBitCount（BGRA）
        0,                     # biCompression = BI_RGB（未压缩）
        len(xor_bitmap) + len(and_mask),  # biSizeImage
        0, 0,                  # 每米像素（不使用）
        0, 0,                  # biClrUsed / biClrImportant
    )
    return header + xor_bitmap + and_mask


def build_ico(path: str, drawer, sizes=(16, 24, 32, 48, 64, 256)) -> None:
    """手写 ICO 容器：小尺寸写 BMP(DIB) 条目，256 写 PNG 压缩条目。

    不能只把 256px 缩小：Windows 会按近似尺寸缩放，细碎元素在 16px 会糊成一片。
    也不能把小尺寸写成 PNG：Explorer 不支持，会静默回退到目标程序图标。
    """
    images = [drawer(size) for size in sizes]
    payloads = []
    for size, image in zip(sizes, images):
        if size >= 256:
            buffer = io.BytesIO()
            image.save(buffer, format="PNG", optimize=True)
            payloads.append(buffer.getvalue())
        else:
            payloads.append(_bmp_entry(image))

    header = struct.pack("<HHH", 0, 1, len(images))
    directory = b""
    offset = 6 + 16 * len(images)
    for size, data in zip(sizes, payloads):
        width = 0 if size >= 256 else size
        height = 0 if size >= 256 else size
        directory += struct.pack("<BBBBHHII", width, height, 0, 0, 1, 32,
                                 len(data), offset)
        offset += len(data)
    with open(path, "wb") as handle:
        handle.write(header + directory + b"".join(payloads))


def verify_ico(path: str) -> list:
    """逐条解析 ICO：确认 16/24/32/48/64 是未压缩 BMP、256 是 PNG 或 BMP。

    不依赖 PIL（PIL 对“能打开”很宽容，无法反映 Explorer 的真实取图规则），
    直接按字节校验目录项与 BITMAPINFOHEADER。
    """
    with open(path, "rb") as handle:
        data = handle.read()
    reserved, kind, count = struct.unpack("<HHH", data[:6])
    if reserved != 0 or kind != 1:
        raise ValueError(f"ICO 头非法: reserved={reserved} type={kind}")
    report = []
    for index in range(count):
        entry = data[6 + 16 * index: 6 + 16 * (index + 1)]
        width_b, height_b, colors, resv, planes, bitcount, size, offset = struct.unpack(
            "<BBBBHHII", entry)
        width = width_b or 256
        height = height_b or 256
        blob = data[offset: offset + size]
        if blob[:8] == b"\x89PNG\r\n\x1a\n":
            if width < 256:
                raise ValueError(f"{width}x{height} 使用了 PNG 压缩（Explorer 不接受）")
            fmt = "PNG"
        else:
            bi_size, bi_w, bi_h, bi_planes, bi_bits, bi_comp = struct.unpack(
                "<IiiHHI", blob[:20])
            if bi_size != 40 or bi_comp != 0:
                raise ValueError(
                    f"{width}x{height} 不是未压缩 BITMAPINFOHEADER"
                    f"（biSize={bi_size} biCompression={bi_comp}）")
            if bi_w != width or bi_h != height * 2:
                raise ValueError(
                    f"{width}x{height} DIB 尺寸不符（biWidth={bi_w} biHeight={bi_h}）")
            if bi_bits != 32:
                raise ValueError(f"{width}x{height} DIB 位深不是 32（{bi_bits}）")
            mask_row = ((width + 31) // 32) * 4
            expected = 40 + width * height * 4 + mask_row * height
            if size < expected:
                raise ValueError(
                    f"{width}x{height} DIB 长度不足（{size} < {expected}）")
            fmt = "BMP"
        report.append((width, height, fmt, bitcount, size))
    return report


def main() -> None:
    os.makedirs(DOCS, exist_ok=True)
    draw_icon_a(256).save(os.path.join(DOCS, "icon_a.png"))
    draw_icon_a(512).save(os.path.join(DOCS, "icon_a_512.png"))
    # 16px 专属简化稿预览（放大 8 倍便于查看，但源图仍是 16px）
    draw_icon_a(16).resize((128, 128), Image.NEAREST).save(
        os.path.join(DOCS, "icon_a_16.png"))
    draw_icon_a(256, dark=True).save(os.path.join(DOCS, "icon_a_dark.png"))

    build_ico(os.path.join(ROOT, "app_icon.ico"), draw_icon_a)
    draw_icon_a(256).save(os.path.join(DOCS, "app_icon_256.png"))
    draw_icon_a(512).save(os.path.join(DOCS, "app_icon_512.png"))
    print("已生成图标：", os.path.join(ROOT, "app_icon.ico"))


if __name__ == "__main__":
    main()
