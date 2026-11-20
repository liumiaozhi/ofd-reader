# -*- coding: utf-8 -*-
"""make_icon.py —— 生成程序图标（assets/logo.ico 与 logo.png）

用 Pillow 直接绘制，不依赖 ImageMagick / rsvg-convert。
图形含义：蓝色圆角底 + 白色文档页 + 红色印章 + OFD 字样，
小尺寸下靠"蓝底白纸"的轮廓辨识，大尺寸下能看清字与印章。

用法：
    python tools/make_icon.py
"""

import math
import os

from PIL import Image, ImageDraw, ImageFont

SIZE = 1024
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "assets")

BLUE_TOP = (23, 82, 150)
BLUE_BOTTOM = (46, 139, 216)
PAGE_WHITE = (255, 255, 255)
PAGE_SHADE = (226, 233, 242)
LINE_COLOR = (176, 199, 226)
INK = (26, 84, 150)
SEAL_RED = (192, 57, 43)


def rounded_mask(size, radius):
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return mask


def gradient(size, top, bottom):
    img = Image.new("RGB", (size, size), top)
    draw = ImageDraw.Draw(img)
    for y in range(size):
        ratio = y / float(size - 1)
        color = (int(top[0] + (bottom[0] - top[0]) * ratio),
                 int(top[1] + (bottom[1] - top[1]) * ratio),
                 int(top[2] + (bottom[2] - top[2]) * ratio))
        draw.line([(0, y), (size, y)], fill=color)
    return img


def load_font(px, bold=True):
    windir = os.environ.get("WINDIR") or r"C:\Windows"
    names = ["arialbd.ttf", "ariblk.ttf", "seguisb.ttf", "calibrib.ttf"]
    for name in names:
        path = os.path.join(windir, "Fonts", name)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, px)
            except Exception:
                continue
    return ImageFont.load_default()


def star_points(cx, cy, outer, inner, points=5, rotate=-90.0):
    result = []
    for i in range(points * 2):
        radius = outer if i % 2 == 0 else inner
        angle = math.radians(rotate + i * 180.0 / points)
        result.append((cx + radius * math.cos(angle),
                       cy + radius * math.sin(angle)))
    return result


def build():
    bg = gradient(SIZE, BLUE_TOP, BLUE_BOTTOM).convert("RGBA")
    bg.putalpha(rounded_mask(SIZE, 210))

    canvas = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    canvas.alpha_composite(bg)

    # 顶部高光，让纯色块看起来不那么平
    gloss = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(gloss)
    gdraw.ellipse([-260, -520, SIZE + 260, 430], fill=(255, 255, 255, 26))
    canvas.alpha_composite(Image.composite(
        gloss, Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0)),
        rounded_mask(SIZE, 210)))

    # 页面阴影
    shadow = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    sdraw.rounded_rectangle([262, 196, 800, 866], radius=30, fill=(10, 40, 80, 90))
    from PIL import ImageFilter
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    canvas.alpha_composite(shadow)

    # 页面
    page = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    pdraw = ImageDraw.Draw(page)
    left, top, right, bottom, fold = 250, 176, 786, 846, 118
    pdraw.rounded_rectangle([left, top, right, bottom], radius=26, fill=PAGE_WHITE)
    # 折角：用背景色盖出一个三角
    pdraw.polygon([(right - fold, top - 2), (right + 2, top - 2),
                   (right + 2, top + fold)], fill=PAGE_SHADE)
    canvas.alpha_composite(page)

    draw = ImageDraw.Draw(canvas)

    # 顶部三条文字线
    for index, width in enumerate((372, 300, 336)):
        y = 292 + index * 66
        draw.rounded_rectangle([312, y, 312 + width, y + 26], radius=13,
                               fill=LINE_COLOR)

    # OFD 字样（在页面内水平居中，纵向留在印章上方）
    font = load_font(152)
    text = "OFD"
    page_center = (left + right) / 2.0
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        draw.text((page_center - tw / 2.0 - bbox[0], 528 - bbox[1]), text,
                  font=font, fill=INK)
    except Exception:
        draw.text((page_center - 160, 528), text, font=font, fill=INK)

    # 红色印章（压在页面右下角，与文字纵向错开）
    cx, cy, radius = 748, 786, 92
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius],
                 outline=SEAL_RED, width=16)
    draw.ellipse([cx - radius + 27, cy - radius + 27,
                  cx + radius - 27, cy + radius - 27],
                 outline=SEAL_RED, width=6)
    draw.polygon(star_points(cx, cy, radius - 46, (radius - 46) * 0.42),
                 fill=SEAL_RED)

    return canvas


def main():
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)
    logo = build()
    png_path = os.path.join(OUT_DIR, "logo.png")
    logo.save(png_path)
    ico_path = os.path.join(OUT_DIR, "logo.ico")
    logo.save(ico_path, format="ICO",
              sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
                     (64, 64), (128, 128), (256, 256)])
    print("已生成：%s" % png_path)
    print("已生成：%s（16/24/32/48/64/128/256）" % ico_path)

    # 拼一张多尺寸预览，方便肉眼检查小图标是否还认得出
    sizes = [16, 24, 32, 48, 64]
    pad = 12
    total_w = sum(sizes) + pad * (len(sizes) + 1)
    total_h = max(sizes) + pad * 2
    sheet = Image.new("RGB", (total_w, total_h), (245, 245, 245))
    x = pad
    for size in sizes:
        small = logo.resize((size, size), Image.LANCZOS)
        sheet.paste(small, (x, pad + (max(sizes) - size) // 2), small)
        x += size + pad
    sheet_path = os.path.join(OUT_DIR, "logo_sizes.png")
    sheet.save(sheet_path)
    print("已生成：%s（多尺寸预览）" % sheet_path)


if __name__ == "__main__":
    main()
