# -*- coding: utf-8 -*-
"""ofd_view.py —— OFD 页面渲染与 PDF 导出

不做像素级复刻，只把阅读器需要的东西画准：
- 矢量图形（填充多边形 + 描边折线）
- 图片（按 Boundary 缩放放置，保留透明通道）
- 文字（逐字形定位，用基线锚点，字体优先取系统中文字体，缺字体时回退到 OFD 内嵌字体）

对外接口：
    FontResolver(res)                字体解析器（可复用，内部有缓存）
    render_page(page, resolver, dpi) 把 PageModel 画成 PIL.Image
    export_pdf(images, path, dpi)    多页图像合成 PDF
"""

import io
import os

from PIL import Image, ImageChops, ImageDraw, ImageFont

import ofd2docx as ox


# 中文字体名 → 系统字体文件名（精确匹配，不跨字族）
FONT_FILES = {
    "宋体": ["simsun.ttc", "simsun.ttf"],
    "新宋体": ["simsun.ttc"],
    "黑体": ["simhei.ttf"],
    "楷体": ["simkai.ttf"],
    "楷体_GB2312": ["simkai.ttf"],
    "仿宋": ["simfang.ttf"],
    "仿宋_GB2312": ["simfang.ttf"],
    "微软雅黑": ["msyh.ttc", "msyh.ttf"],
    "方正小标宋简体": ["FZXBSJW.TTF", "FZXBSK.TTF", "FZXBSJW.ttf"],
    "方正仿宋简体": ["FZFSJW.TTF", "FZFSK.TTF"],
    "方正黑体简体": ["FZHTJW.TTF", "FZHTK.TTF"],
    "方正书宋简体": ["FZSSJW.TTF", "FZSSK.TTF"],
    "Times New Roman": ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"],
    "Arial": ["arial.ttf", "arialbd.ttf"],
    "Calibri": ["calibri.ttf", "calibrib.ttf"],
    "Cambria": ["cambria.ttc", "cambria.ttf"],
    "Courier New": ["cour.ttf"],
    "Helvetica": ["arial.ttf"],
}

# 系统里确实找不到时才用的兜底字体
GENERIC_CJK = ["simsun.ttc", "msyh.ttc", "simhei.ttf", "simfang.ttf"]
GENERIC_LATIN = ["times.ttf", "arial.ttf", "calibri.ttf"]

_SEAL_CACHE = {}


class FontResolver(object):
    """把字体名解析成可用的 PIL 字体对象，带多级回退。"""

    def __init__(self, res=None):
        # type: (object) -> None
        self.res = res
        self._files = {}          # type: dict
        self._cache = {}          # type: dict
        self._bytes_cache = {}    # type: dict
        self._scan_system_fonts()

    def _scan_system_fonts(self):
        # type: () -> None
        windir = os.environ.get("WINDIR") or r"C:\Windows"
        local = os.environ.get("LOCALAPPDATA") or ""
        dirs = [os.path.join(windir, "Fonts")]
        if local:
            dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
        for folder in dirs:
            try:
                names = os.listdir(folder)
            except OSError:
                continue
            for name in names:
                low = name.lower()
                if not low.endswith((".ttf", ".ttc", ".otf")):
                    continue
                self._files.setdefault(low, os.path.join(folder, name))

    # -- 字体来源 ---------------------------------------------------------
    def _exact_names(self, family):
        # type: (str) -> list
        """只返回该字体名本身（含映射后的中文名），不带任何兜底字体。"""
        names = []
        if family:
            names.append(family)
            mapped, _known = ox.normalize_font_name(family)
            if mapped and mapped not in names:
                names.append(mapped)
        return names

    def _candidate_names(self, family, is_cjk):
        # type: (str, bool) -> list
        names = self._exact_names(family)
        chain = GENERIC_CJK if is_cjk else GENERIC_LATIN
        for extra in chain:
            if extra not in names:
                names.append(extra)
        return names

    def _exact_file(self, family, is_cjk):
        # type: (str, bool) -> object
        for name in self._exact_names(family):
            for fname in FONT_FILES.get(name, []):
                path = self._files.get(fname.lower())
                if path:
                    return path
        if family:
            for ext in (".ttf", ".ttc", ".otf"):
                path = self._files.get((family + ext).lower())
                if path:
                    return path
        return None

    def _generic_file(self, is_cjk):
        # type: (bool) -> object
        chain = GENERIC_CJK if is_cjk else GENERIC_LATIN
        for fname in chain:
            path = self._files.get(fname.lower())
            if path:
                return path
        for name, path in self._files.items():
            if name.endswith((".ttf", ".ttc")):
                return path
        return None

    def _bytes_source(self, family, is_cjk):
        # type: (str, bool) -> object
        if self.res is None or not hasattr(self.res, "font_bytes_for"):
            return None
        for name in self._candidate_names(family, is_cjk):
            data = self.res.font_bytes_for(name)
            if data:
                return ("bytes", name, data)
        return None

    def _make_font(self, family, is_cjk, px, bold):
        # type: (str, bool, int, bool) -> ImageFont
        """取字体：系统同名字体 → 通用兜底 → OFD 内嵌字体 → PIL 内置。

        注意内嵌字体排在很后面：OFD 内嵌的往往是"按字形索引"的子集，
        文件里根本没有 cmap 表，Pillow 按 Unicode 取字只会得到空白/方框。
        只有系统里连一个中文字体都找不到时才值得一试。
        """
        px = max(6, int(px))

        # 1) 系统里装了同名字体（最理想：字形正确、度量一致）
        path = self._exact_file(family, is_cjk)
        if path:
            cache_key = ("file", path, px)
            cached = self._cache.get(cache_key)
            if cached is not None:
                return cached
            try:
                font = ImageFont.truetype(path, px, index=0)
                self._cache[cache_key] = font
                return font
            except Exception:
                pass

        # 2) 通用中/西文兜底（宋体 / 微软雅黑 / Times）
        path = self._generic_file(is_cjk)
        if path:
            cache_key = ("gen", path, px)
            cached = self._cache.get(cache_key)
            if cached is not None:
                return cached
            try:
                font = ImageFont.truetype(path, px, index=0)
                self._cache[cache_key] = font
                return font
            except Exception:
                pass

        # 3) 最后才试 OFD 内嵌字体
        source = self._bytes_source(family, is_cjk)
        if source is not None:
            name, data = source[1], source[2]
            cache_key = ("bytes", name, px)
            cached = self._cache.get(cache_key)
            if cached is not None:
                return cached
            try:
                font = ImageFont.truetype(io.BytesIO(data), px, index=0)
                self._cache[cache_key] = font
                return font
            except Exception:
                pass

        # 4) PIL 内置点阵字体（中文会显示为方块，但不会崩）
        cache_key = ("default", px)
        cached = self._cache.get(cache_key)
        if cached is not None:
            return cached
        try:
            font = ImageFont.load_default()
        except Exception:
            font = None
        self._cache[cache_key] = font
        return font

    def get(self, ch, size_mm, family_latin, family_cjk, bold, px_per_mm):
        # type: (str, float, str, str, bool, float) -> ImageFont
        px = max(6, int(round(size_mm * px_per_mm)))
        is_cjk = ox.is_wide(ch)
        family = family_cjk if is_cjk else family_latin
        key = (family, is_cjk, bool(bold), px)
        hit = self._cache.get(key)
        if hit is not None:
            return hit
        font = self._make_font(family, is_cjk, px, bool(bold))
        self._cache[key] = font
        return font


def _poly_area(pts):
    # type: (list) -> float
    """带符号面积。正负代表环绕方向，用来区分字形外轮廓与内部的"洞"。"""
    n = len(pts)
    if n < 3:
        return 0.0
    total = 0.0
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        total += x1 * y2 - x2 * y1
    return total * 0.5


def _group_vectors(items):
    # type: (list) -> list
    """把同一 PathObject 拆出的子路径收回一组。

    Even-Odd 规则只在"同一个 PathObject 的多个子路径"之间生效：外轮廓与
    中间的洞必须一起参与运算，单独填每一条只会得到实心块或碎片。
    """
    groups = []       # type: list
    cur_key = object()
    for item in items:
        key = item.get("group")
        if key is None:
            groups.append([item])
            cur_key = object()
            continue
        if not groups or key != cur_key:
            groups.append([item])
            cur_key = key
        else:
            groups[-1].append(item)
    return groups


def _mask_union(polys, size):
    # type: (list, tuple) -> object
    mask = Image.new("1", size, 0)
    draw = ImageDraw.Draw(mask)
    for pts in polys:
        if len(pts) < 3:
            continue
        try:
            draw.polygon(pts, fill=1)
        except Exception:
            continue
    return mask


def _vector_group_rgba(items, ppm, page_w, page_h):
    # type: (list, float, float, float) -> object
    """把一组矢量路径画成一张局部 RGBA 图，返回 (x, y, image) 或 None。

    要点：
    - 按 PathObject 的 Rule 合成子路径（奇偶 XOR / 非零按环绕方向挖洞）；
    - 2 倍超采样后缩小，边缘不会出现锯齿（公文红头字大，锯齿很明显）。
    """
    xs = []           # type: list
    ys = []           # type: list
    stroke_max = 0.0
    has_fill = False
    has_stroke = False
    for item in items:
        for (x, y) in item["points"]:
            xs.append(x)
            ys.append(y)
        if item.get("fill"):
            has_fill = True
        if item.get("stroke"):
            has_stroke = True
            stroke_max = max(stroke_max, float(item.get("width") or 0.0))
    if not xs or (not has_fill and not has_stroke):
        return None

    pad = stroke_max * 0.5 + 0.6
    x0 = max(0.0, min(xs) - pad)
    y0 = max(0.0, min(ys) - pad)
    x1 = min(page_w, max(xs) + pad)
    y1 = min(page_h, max(ys) + pad)
    if x1 - x0 < 0.05 or y1 - y0 < 0.05:
        return None

    ss = 2
    w = int(round((x1 - x0) * ppm * ss))
    h = int(round((y1 - y0) * ppm * ss))
    if w < 1 or h < 1:
        return None
    if w * h > 16000000:          # 整页底纹这类巨图不超采样，免得住满内存
        ss = 1
        w = max(1, int(round((x1 - x0) * ppm)))
        h = max(1, int(round((y1 - y0) * ppm)))

    def local(pts):
        return [((x - x0) * ppm * ss, (y - y0) * ppm * ss) for (x, y) in pts]

    fw = max(1, w // ss)
    fh = max(1, h // ss)
    size = (w, h)                 # 超采样空间，用来画
    out_size = (fw, fh)           # 最终交付尺寸
    layer = Image.new("RGBA", out_size, (0, 0, 0, 0))

    fill_items = [it for it in items if it.get("fill") and len(it["points"]) >= 3]
    if fill_items:
        rule = str(fill_items[0].get("rule") or "Non-Zero").lower()
        polys = [local(it["points"]) for it in fill_items]
        if rule.startswith("even"):
            mask = None
            for pts in polys:
                one = _mask_union([pts], size)
                mask = one if mask is None else ImageChops.logical_xor(mask, one)
        else:
            pos = [p for p in polys if _poly_area(p) >= 0]
            neg = [p for p in polys if _poly_area(p) < 0]
            s_pos = sum(abs(_poly_area(p)) for p in pos)
            s_neg = sum(abs(_poly_area(p)) for p in neg)
            if pos and neg and min(s_pos, s_neg) > 0:
                outer, inner = (pos, neg) if s_pos >= s_neg else (neg, pos)
                mask = ImageChops.subtract(_mask_union(outer, size),
                                           _mask_union(inner, size))
            else:
                mask = _mask_union(polys, size)
        if mask is not None:
            mask = mask.convert("L").resize(out_size, Image.LANCZOS)
            color = tuple(fill_items[0]["fill_color"]) + (255,)
            layer.paste(Image.new("RGBA", out_size, color), (0, 0), mask)

    stroke_items = [it for it in items if it.get("stroke") and len(it["points"]) >= 2]
    if stroke_items:
        smask = Image.new("L", size, 0)
        sdraw = ImageDraw.Draw(smask)
        for it in stroke_items:
            pts = local(it["points"])
            width = max(1, int(round(float(it.get("width") or 0.35) * ppm * ss)))
            color = tuple(it["stroke_color"]) + (255,)
            try:
                if len(pts) >= 3:
                    sdraw.line(pts, fill=255, width=width, joint="curve")
                else:
                    sdraw.line(pts, fill=255, width=width)
            except Exception:
                continue
        s_color = tuple(stroke_items[0]["stroke_color"]) + (255,)
        smask = smask.resize(out_size, Image.LANCZOS)
        layer.paste(Image.new("RGBA", out_size, s_color), (0, 0), smask)

    return (int(round(x0 * ppm)), int(round(y0 * ppm)), layer)


def _draw_vectors(base, page, ppm):
    # type: (Image.Image, object, float) -> None
    for group in _group_vectors(page.vectors):
        try:
            drawn = _vector_group_rgba(group, ppm, page.box.w, page.box.h)
        except Exception:
            drawn = None
        if drawn is None:
            continue
        pos, layer = (drawn[0], drawn[1]), drawn[2]
        base.paste(layer, pos, layer)



def _draw_images(base, page, ppm):
    # type: (Image.Image, object, float) -> int
    drawn = 0
    for item in page.images:
        try:
            im = Image.open(io.BytesIO(item.data))
            im.load()
            tw = max(1, int(round(item.box.w * ppm)))
            th = max(1, int(round(item.box.h * ppm)))
            im = im.resize((tw, th), Image.LANCZOS)
            pos = (int(round(item.box.x * ppm)), int(round(item.box.y * ppm)))
            if im.mode in ("RGBA", "LA", "P"):
                im = im.convert("RGBA")
                base.paste(im, pos, im)
            else:
                base.paste(im.convert("RGB"), pos)
            drawn += 1
        except Exception:
            continue
    return drawn


def _draw_text(draw, page, resolver, ppm):
    # type: (ImageDraw.ImageDraw, object, FontResolver, float) -> int
    count = 0
    for line in page.lines:
        for g in line.glyphs:
            ch = g.ch
            if not ch or ch.isspace():
                continue
            font = resolver.get(ch, g.size_mm, g.font_latin, g.font_cjk,
                                g.bold, ppm)
            if font is None:
                continue
            try:
                draw.text((round(g.x * ppm, 2), round(g.y * ppm, 2)), ch,
                          font=font, fill=tuple(g.color), anchor="ls")
                count += 1
            except Exception:
                continue
    return count


def seal_rgba(data):
    # type: (bytes) -> object
    """把印章图转成"白底透明"的 RGBA。

    签章里的印章多数是白底不透明的 GIF，直接贴会把下面的文字整块盖住；
    把接近纯白的像素设成透明后，红色笔画仍然清晰，文字也读得出来。
    """
    import hashlib
    key = hashlib.sha1(data).hexdigest()
    hit = _SEAL_CACHE.get(key)
    if hit is not None:
        return hit
    im = Image.open(io.BytesIO(data))
    im.load()
    im = im.convert("RGBA")
    gray = im.convert("L")
    alpha = gray.point(lambda v: 0 if v > 238 else 255)
    im.putalpha(alpha)
    _SEAL_CACHE[key] = im
    return im


def _draw_stamps(base, page, ppm):
    # type: (Image.Image, object, float) -> int
    """画电子印章。印章盖在文字之上，符合实际观感。"""
    drawn = 0
    for stamp in getattr(page, "stamps", []) or []:
        try:
            im = seal_rgba(stamp["data"])
            box = stamp["box"]
            tw = max(1, int(round(box.w * ppm)))
            th = max(1, int(round(box.h * ppm)))
            im = im.resize((tw, th), Image.LANCZOS)
            pos = (int(round(box.x * ppm)), int(round(box.y * ppm)))
            base.paste(im, pos, im)
            drawn += 1
        except Exception:
            continue
    return drawn


def render_page(page, resolver=None, dpi=110.0, scale=1.0, draw_blank_hint=False):
    # type: (object, FontResolver, float, float, bool) -> Image.Image
    """把一页 PageModel 画成 PIL 图像。dpi 控制清晰度，scale 用于缩放预览。"""
    if resolver is None:
        resolver = FontResolver()
    ppm = max(0.5, float(dpi) / 25.4 * float(scale))
    width = max(1, int(round(page.box.w * ppm)))
    height = max(1, int(round(page.box.h * ppm)))
    image = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(image)

    _draw_vectors(image, page, ppm)
    _draw_images(image, page, ppm)
    _draw_text(draw, page, resolver, ppm)
    _draw_stamps(image, page, ppm)

    if draw_blank_hint and not page.lines and not page.images and not page.vectors:
        try:
            draw.rectangle([2, 2, width - 3, height - 3],
                           outline=(200, 200, 200), width=1)
        except Exception:
            pass
    return image


def _rect_matches(box, rects, tol=0.8):
    # type: (tuple, list, float) -> bool
    if not rects:
        return False
    x0, y0, x1, y1 = box
    for (sx0, sy0, sx1, sy1) in rects:
        if (abs(x0 - sx0) <= tol and abs(y0 - sy0) <= tol and
                abs(x1 - sx1) <= tol and abs(y1 - sy1) <= tol):
            return True
    return False


def overlay_png(page, skip_rects=None, dpi=150.0):
    # type: (object, object, float) -> object
    """把一页的矢量图形与印章画成一张整页透明 PNG，供 Word 浮动图层使用。

    文字仍由 Word 自己排（可编辑可搜索），这里只补"不是文字"的东西：
    红头横线、版记横线、框线、底纹、电子印章等。
    """
    vectors = getattr(page, "vectors", None) or []
    stamps = getattr(page, "stamps", None) or []
    if not vectors and not stamps:
        return None
    ppm = max(0.5, float(dpi) / 25.4)
    width = max(1, int(round(page.box.w * ppm)))
    height = max(1, int(round(page.box.h * ppm)))
    image = Image.new("RGBA", (width, height), (255, 255, 255, 0))

    for group in _group_vectors(vectors):
        raw_group = []
        for item in group:
            if skip_rects:
                raw_pts = item["points"]
                xs = [p[0] for p in raw_pts]
                ys = [p[1] for p in raw_pts]
                if _rect_matches((min(xs), min(ys), max(xs), max(ys)), skip_rects):
                    continue        # 已改用段落边框表达
            raw_group.append(item)
        if not raw_group:
            continue
        try:
            drawn = _vector_group_rgba(raw_group, ppm, page.box.w, page.box.h)
        except Exception:
            drawn = None
        if drawn is None:
            continue
        layer = drawn[2]
        image.paste(layer, (drawn[0], drawn[1]), layer)

    for stamp in stamps:
        try:
            im = seal_rgba(stamp["data"])
            box = stamp["box"]
            tw = max(1, int(round(box.w * ppm)))
            th = max(1, int(round(box.h * ppm)))
            im = im.resize((tw, th), Image.LANCZOS)
            pos = (int(round(box.x * ppm)), int(round(box.y * ppm)))
            image.paste(im, pos, im)
        except Exception:
            continue

    buf = io.BytesIO()
    image.save(buf, "PNG", optimize=True)
    return buf.getvalue()


def export_pdf(images, path, dpi=150.0):
    # type: (list, str, float) -> None
    """把若干页图像写成一个多页 PDF。

    说明：这里的 PDF 是"每页一张位图"，适合查看与打印；文字不可选中、不可搜索。
    若要可编辑/可检索的文档，请用"导出 Word"。
    """
    if not images:
        raise ox.UsageError("没有可导出的页面")
    converted = []
    for im in images:
        if im.mode != "RGB":
            im = im.convert("RGB")
        converted.append(im)
    first = converted[0]
    rest = converted[1:]
    tmp = "%s.tmp-%d" % (path, os.getpid())
    try:
        if rest:
            first.save(tmp, "PDF", save_all=True, append_images=rest,
                       resolution=float(dpi))
        else:
            first.save(tmp, "PDF", resolution=float(dpi))
        os.replace(tmp, path)
    except PermissionError:
        raise ox.OfdError("无法写入 PDF（文件可能正被打开）", part=path,
                          hint="请先关闭同名的 PDF 文件，或换个文件名")
    except Exception as exc:
        raise ox.OfdError("生成 PDF 失败：%s" % exc, part=path)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
