# -*- coding: utf-8 -*-
"""从离屏空状态截图定位蓝色徽标，测量圆与白色加号的像素对称性。"""
import json
import pathlib

from PIL import Image, ImageChops

path = pathlib.Path(__file__).with_name("current_badge_input.png")
if not path.exists():
    path = pathlib.Path(r"C:\Users\Lenovo\AppData\Local\Temp\codearts-current-empty-badge.png")
image = Image.open(path).convert("RGB")
pixels = image.load()
width, height = image.size
mask = {(x, y) for y in range(height) for x in range(width)
        if (lambda r, g, b: b > 170 and b > r + 70 and b > g + 60)(*pixels[x, y])}
components = []
while mask:
    seed = mask.pop()
    stack = [seed]
    points = [seed]
    while stack:
        x, y = stack.pop()
        for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
            if (nx, ny) in mask:
                mask.remove((nx, ny))
                stack.append((nx, ny))
                points.append((nx, ny))
    components.append(points)
candidates = []
for points in components:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    box = (min(xs), min(ys), max(xs) + 1, max(ys) + 1)
    w, h = box[2] - box[0], box[3] - box[1]
    if 15 <= w <= 40 and 15 <= h <= 40 and abs(w - h) <= 3:
        candidates.append((box, points))
result = {"image": str(path), "candidate_count": len(candidates), "badges": []}
for box, points in candidates:
    crop = image.crop(box)
    white = Image.new("RGB", crop.size, (255, 255, 255))
    diff = ImageChops.difference(crop, white)
    horizontal = ImageChops.difference(crop, crop.transpose(Image.Transpose.FLIP_LEFT_RIGHT))
    vertical = ImageChops.difference(crop, crop.transpose(Image.Transpose.FLIP_TOP_BOTTOM))
    result["badges"].append({
        "box": box,
        "size": crop.size,
        "not_white_bbox": diff.getbbox(),
        "horizontal_bbox": horizontal.getbbox(),
        "vertical_bbox": vertical.getbbox(),
        "center_pixel": crop.getpixel((crop.width // 2, crop.height // 2)),
    })
    crop.resize((crop.width * 8, crop.height * 8), Image.NEAREST).save(
        path.with_name(f"current_badge_crop_{box[0]}_{box[1]}.png"))
print(json.dumps(result, indent=2))