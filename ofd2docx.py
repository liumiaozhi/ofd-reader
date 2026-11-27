# -*- coding: utf-8 -*-
"""
ofd2docx.py —— OFD 版式文档转 Word（.docx）流式转换器

适用范围
--------
把符合 GB/T 33190-2016 的 .ofd 文件转换为可编辑的 .docx。转换采用"流式还原"：
一个 OFD 文字行输出为一个 Word 段落（用左缩进 + 固定行距还原位置），
由线段构成的规则表格输出为可编辑的 Word 真表格，图片按原始尺寸内联，
每个 OFD 页面独立成一个 Word 节并沿用原始页面尺寸。

依赖
----
    python-docx == 1.1.2   （生成 docx）
    lxml        >= 4.9,<6  （解析 OFD 内的 XML，同时用于 XXE 防护）
    Pillow      >= 8.0     （可选：仅在遇到 JPEG2000 等非白名单图像时转码）

    pip install "python-docx==1.1.2" "lxml==5.2.2" "Pillow==10.4.0"

命令行
------
    python ofd2docx.py 输入.ofd                       # 输出同名 .docx
    python ofd2docx.py 输入.ofd -o 输出.docx           # 指定输出
    python ofd2docx.py 输入.ofd --report r.json -v     # 输出转换报告
    python ofd2docx.py --self-test                     # 运行内建自测

退出码
------
    0  成功
    1  参数错误或文件读写失败
    2  OFD 格式错误（不是 ZIP、缺 OFD.xml、XML 语法错、引用部件不存在）
    3  文件已加密或包含不支持的特性
    4  转换完成但有内容损失（有告警），详见报告

已知转换限制
------------
1. 本工具做的是"流式还原"，不是像素级版式复刻。OFD 的绝对定位被转换成
   "段落左缩进 + 固定行距"，因此换行位置、字间距可能与原件有细微差异。
2. 一个 OFD 文字行对应一个 Word 段落。同一视觉行内的多栏文本会被合并为一行。
3. 表格仅支持由水平/垂直线段构成的规则表格。斜线表头、圆角单元格、跨页表格、
   旋转表格、嵌套表格均不支持；非矩形合并区会被拆分为多个单元格。
4. 电子签章、数字签名、印章、水印、手写批注、表单域、超链接、书签、大纲
   均不转换，仅在转换报告中统计数量。
5. 下划线、删除线、文字描边、背景底纹、艺术字、竖排/旋转文字不做还原；
   竖排文字按字符顺序降级为横排输出。
6. 字体不嵌入。文档中记录的是字体名称，打开文档的机器若缺少该字体，
   Word 会自动替换（中文缺字体时通常回退为宋体）。
7. 图片方面：JBIG2 编码无法解码，会被跳过并告警；JPEG2000 会自动转为 PNG；
   CMYK 色彩空间的 JPEG 在屏幕上可能偏色；透明 PNG 的透明区域显示为白色。
8. 加密或带口令保护的 OFD 不支持，会直接报错退出。
9. 仅以字形索引（Glyphs/CGTransform）描述、不含 Unicode 文本的文字对象无法还原。
10. OFD 没有页眉页脚概念，原件中的页眉页脚内容会作为正文输出，页码会成为
    页面底部的普通段落。
11. 极窄的排版间隙（例如页码 "— 1 —" 中破折号与数字之间）可能识别不到，
    导致少一个空格。
"""

import argparse
import io
import json
import math
import os
import posixpath
import re
import statistics
import struct
import sys
import tempfile
import time
import unicodedata
import zipfile

from typing import Any, Dict, List, Optional, Sequence, Tuple

__version__ = "1.0.0"

# ---------------------------------------------------------------------------
# B1  单位换算与几何
# ---------------------------------------------------------------------------

MM_PER_INCH = 25.4
TWIPS_PER_MM = 1440.0 / 25.4      # 1 mm = 56.6929 twips
EMU_PER_MM = 36000.0              # 914400 / 25.4，精确整数
PT_PER_MM = 72.0 / 25.4           # 1 mm = 2.8346 pt

# 各类容差（单位 mm）
TOL_ON_PAGE = 0.5          # 判定"落在页内"的外扩容差
TOL_ANGLE = 0.6            # 判定水平/垂直线的角度容差（度）
TOL_MIN_SEG = 0.35         # 线段最短长度
TOL_SNAP_BASE = 0.8        # 坐标吸附基础容差


def mm_to_twips(mm):
    # type: (float) -> int
    return int(round(mm * TWIPS_PER_MM))


def mm_to_emu(mm):
    # type: (float) -> int
    return int(round(mm * EMU_PER_MM))


def mm_to_pt(mm):
    # type: (float) -> float
    return mm * PT_PER_MM


def mm_to_half_pt(mm):
    # type: (float) -> int
    """Word 的 w:sz 用半磅，字号换算。"""
    return max(2, int(round(mm_to_pt(mm) * 2.0)))


def mm_to_eighth_pt(mm):
    # type: (float) -> int
    """Word 的边框宽度 w:sz 用 1/8 磅。"""
    return max(2, min(96, int(round(mm_to_pt(mm) * 8.0))))


class Box(object):
    """轴对齐矩形，单位 mm，原点在页面左上角。"""

    __slots__ = ("x", "y", "w", "h")

    def __init__(self, x=0.0, y=0.0, w=0.0, h=0.0):
        # type: (float, float, float, float) -> None
        self.x = float(x)
        self.y = float(y)
        self.w = float(w)
        self.h = float(h)

    @property
    def x2(self):
        # type: () -> float
        return self.x + self.w

    @property
    def y2(self):
        # type: () -> float
        return self.y + self.h

    @property
    def area(self):
        # type: () -> float
        return max(0.0, self.w) * max(0.0, self.h)

    def contains_point(self, x, y, pad=0.0):
        # type: (float, float, float) -> bool
        return (self.x - pad) <= x <= (self.x2 + pad) and \
               (self.y - pad) <= y <= (self.y2 + pad)

    def __repr__(self):
        # type: () -> str
        return "Box(%.2f, %.2f, %.2f, %.2f)" % (self.x, self.y, self.w, self.h)


def parse_box(text):
    # type: (Optional[str]) -> Optional[Box]
    """解析 ST_Box。规范中后两位是宽高，但存在写成 x1 y1 x2 y2 的文件。"""
    vals = parse_array(text)
    if len(vals) < 4:
        return None
    x, y, w, h = vals[0], vals[1], vals[2], vals[3]
    if w <= 0 or h <= 0:
        w, h = w - x, h - y
    if w <= 0 or h <= 0:
        return None
    return Box(x, y, w, h)


def bbox_of(points):
    # type: (Sequence[Tuple[float, float]]) -> Optional[Box]
    if not points:
        return None
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return Box(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))


def box_iou(a, b):
    # type: (Optional[Box], Optional[Box]) -> float
    if a is None or b is None or a.area <= 0 or b.area <= 0:
        return 0.0
    ix = max(0.0, min(a.x2, b.x2) - max(a.x, b.x))
    iy = max(0.0, min(a.y2, b.y2) - max(a.y, b.y))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    union = a.area + b.area - inter
    return inter / union if union > 0 else 0.0


def apply_matrix(m, x, y):
    # type: (Sequence[float], float, float) -> Tuple[float, float]
    a, b, c, d, e, f = m[0], m[1], m[2], m[3], m[4], m[5]
    return (a * x + c * y + e, b * x + d * y + f)


def matrix_is_identity(m):
    # type: (Optional[Sequence[float]]) -> bool
    if not m or len(m) < 6:
        return True
    return (abs(m[0] - 1.0) < 1e-6 and abs(m[1]) < 1e-6 and
            abs(m[2]) < 1e-6 and abs(m[3] - 1.0) < 1e-6 and
            abs(m[4]) < 1e-6 and abs(m[5]) < 1e-6)


# ---------------------------------------------------------------------------
# B2  异常层次
# ---------------------------------------------------------------------------

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_FORMAT = 2
EXIT_UNSUPPORTED = 3
EXIT_PARTIAL = 4


class OfdError(Exception):
    """所有转换错误的基类，携带定位信息与用户可读的处置建议。"""

    kind = "错误"
    exit_code = EXIT_FORMAT

    def __init__(self, message, part=None, page=None, obj_id=None, hint=None):
        # type: (str, Optional[str], Optional[int], Optional[str], Optional[str]) -> None
        super(OfdError, self).__init__(message)
        self.message = message
        self.part = part
        self.page = page
        self.obj_id = obj_id
        self.hint = hint

    def __str__(self):
        # type: () -> str
        return self.user_message()

    def user_message(self):
        # type: () -> str
        lines = ["[%s] %s" % (self.kind, self.message)]
        if self.part:
            lines.append("    位置：部件 %s" % self.part)
        if self.page is not None:
            extra = ""
            if self.obj_id is not None:
                extra = "，对象 ID=%s" % self.obj_id
            lines.append("    位置：第 %s 页%s" % (self.page, extra))
        elif self.obj_id is not None:
            lines.append("    位置：对象 ID=%s" % self.obj_id)
        if self.hint:
            lines.append("    建议：%s" % self.hint)
        return "\n".join(lines)


class OfdFormatError(OfdError):
    """文件不是合法 OFD、XML 损坏、引用的部件不存在。"""
    kind = "格式错误"
    exit_code = EXIT_FORMAT


class OfdEncryptedError(OfdError):
    """文件受口令/加密保护。"""
    kind = "不支持"
    exit_code = EXIT_UNSUPPORTED


class OfdUnsupportedError(OfdError):
    """使用了本工具不支持的 OFD 特性。"""
    kind = "不支持"
    exit_code = EXIT_UNSUPPORTED


class OfdGeometryError(OfdError):
    """页面尺寸或坐标系无法确定。"""
    kind = "几何错误"
    exit_code = EXIT_UNSUPPORTED


class UsageError(OfdError):
    """命令行参数本身有问题（与文件内容无关）。"""
    kind = "参数错误"
    exit_code = EXIT_USAGE


# ---------------------------------------------------------------------------
# B3  数据模型
# ---------------------------------------------------------------------------

class Warning_ (object):
    """一条转换告警。用独立类避免与内置 Warning 冲突。"""

    __slots__ = ("level", "code", "message", "page", "obj_id")

    def __init__(self, level, code, message, page=None, obj_id=None):
        self.level = level
        self.code = code
        self.message = message
        self.page = page
        self.obj_id = obj_id

    def as_dict(self):
        # type: () -> Dict[str, Any]
        d = {"level": self.level, "code": self.code, "message": self.message}
        if self.page is not None:
            d["page"] = self.page
        if self.obj_id is not None:
            d["obj"] = str(self.obj_id)
        return d

    def __str__(self):
        # type: () -> str
        loc = ""
        if self.page is not None:
            loc = "第 %s 页 " % self.page
            if self.obj_id is not None:
                loc += "ID=%s " % self.obj_id
        return "%s%s: %s" % (loc, self.level.upper(), self.message)


class Report(object):
    """转换报告：逐页统计 + 告警列表。"""

    def __init__(self):
        # type: () -> None
        self.warnings = []          # type: List[Warning_]
        self.pages = []             # type: List[Dict[str, Any]]
        self.meta = {
            "ok": False,
            "ofd_version": None,
            "doc_bodies": 0,
            "outlines_skipped": 0,
            "signatures_skipped": 0,
            "annotations_skipped": 0,
        }
        self.elapsed_ms = 0

    def warn(self, code, message, page=None, obj_id=None):
        # type: (str, str, Optional[int], Optional[str]) -> None
        self.warnings.append(Warning_("warn", code, message, page, obj_id))

    def info(self, code, message, page=None, obj_id=None):
        # type: (str, str, Optional[int], Optional[str]) -> None
        self.warnings.append(Warning_("info", code, message, page, obj_id))

    def has_warning(self):
        # type: () -> bool
        for w in self.warnings:
            if w.level == "warn":
                return True
        return False

    def totals(self):
        # type: () -> Dict[str, Any]
        keys = ("paragraphs", "tables", "images", "images_skipped", "lines",
                "empty_paragraphs")
        out = {"pages": len(self.pages)}
        for k in keys:
            out[k] = sum(p.get(k, 0) for p in self.pages)
        out["annotations_skipped"] = self.meta.get("annotations_skipped", 0)
        out["signatures_skipped"] = self.meta.get("signatures_skipped", 0)
        out["outlines_skipped"] = self.meta.get("outlines_skipped", 0)
        out["warnings"] = sum(1 for w in self.warnings if w.level == "warn")
        out["infos"] = sum(1 for w in self.warnings if w.level == "info")
        return out

    def as_dict(self):
        # type: () -> Dict[str, Any]
        d = dict(self.meta)
        d["totals"] = self.totals()
        d["pages"] = self.pages
        d["warnings"] = [w.as_dict() for w in self.warnings]
        d["elapsed_ms"] = self.elapsed_ms
        return d


class Piece(object):
    """同一段落内样式一致的一段文字。"""

    __slots__ = ("text", "size_mm", "color", "font_latin", "font_cjk",
                 "bold", "italic")

    def __init__(self, text, size_mm, color, font_latin, font_cjk, bold, italic):
        self.text = text
        self.size_mm = size_mm
        self.color = color
        self.font_latin = font_latin
        self.font_cjk = font_cjk
        self.bold = bold
        self.italic = italic

    def style_key(self):
        # type: () -> Tuple[Any, ...]
        return (round(self.size_mm, 2), self.color, self.font_latin,
                self.font_cjk, self.bold, self.italic)


class Glyph(object):
    """一个字形：字符 + 页面坐标 + 样式。"""

    __slots__ = ("ch", "x", "y", "adv", "size_mm", "color", "font_latin",
                 "font_cjk", "bold", "italic", "obj_id", "seq")

    def __init__(self, ch, x, y, adv, size_mm, color, font_latin, font_cjk,
                 bold, italic, obj_id, seq):
        self.ch = ch
        self.x = x
        self.y = y
        self.adv = adv
        self.size_mm = size_mm
        self.color = color
        self.font_latin = font_latin
        self.font_cjk = font_cjk
        self.bold = bold
        self.italic = italic
        self.obj_id = obj_id
        self.seq = seq


class TextLine(object):
    """一个视觉行，由若干字形拼成，最终拆成若干样式段。"""

    __slots__ = ("glyphs", "pieces", "y", "size_mm", "left", "right",
                 "gap_to_next", "page")

    def __init__(self, glyphs, page=0):
        # type: (List[Glyph], int) -> None
        self.glyphs = glyphs
        self.pieces = []        # type: List[Piece]
        self.y = 0.0
        self.size_mm = 0.0
        self.left = 0.0
        self.right = 0.0
        self.gap_to_next = None  # type: Optional[float]
        self.page = page

    @property
    def text(self):
        # type: () -> str
        return "".join(p.text for p in self.pieces)

    def is_blank(self):
        # type: () -> bool
        return not self.text.strip()

    @property
    def top(self):
        # type: () -> float
        return self.y - self.size_mm


class Seg(object):
    """一条水平或垂直线段（页面坐标 mm）。"""

    __slots__ = ("orient", "c", "a", "b", "width", "color", "from_fill")

    def __init__(self, orient, c, a, b, width, color=None, from_fill=False):
        self.orient = orient   # 'h' | 'v'
        self.c = c             # h: y 坐标；v: x 坐标
        self.a = a             # 起点（h: x1；v: y1）
        self.b = b             # 终点
        self.width = width
        self.color = color     # 填充色（仅 from_fill 时有意义）
        self.from_fill = from_fill  # 是否来自"细长填充矩形"

    def __repr__(self):
        # type: () -> str
        return "Seg(%s c=%.2f [%.2f,%.2f] w=%.2f)" % (
            self.orient, self.c, self.a, self.b, self.width)


class Cell(object):
    __slots__ = ("r", "c", "rs", "cs", "lines", "borders")

    def __init__(self, r, c, rs, cs):
        self.r = r
        self.c = c
        self.rs = rs
        self.cs = cs
        self.lines = []        # type: List[TextLine]
        self.borders = {"top": None, "left": None, "bottom": None, "right": None}


class Table(object):
    __slots__ = ("col_xs", "row_ys", "cells", "box", "score")

    def __init__(self, col_xs, row_ys, cells, score=1.0):
        # type: (List[float], List[float], List[Cell], float) -> None
        self.col_xs = col_xs
        self.row_ys = row_ys
        self.cells = cells
        self.score = score
        self.box = Box(col_xs[0], row_ys[0],
                       col_xs[-1] - col_xs[0], row_ys[-1] - row_ys[0])

    @property
    def nrows(self):
        # type: () -> int
        return len(self.row_ys) - 1

    @property
    def ncols(self):
        # type: () -> int
        return len(self.col_xs) - 1


class ImageItem(object):
    __slots__ = ("data", "ext", "box", "page", "resource_id", "sha1", "alpha")

    def __init__(self, data, ext, box, page, resource_id, sha1, alpha=255.0):
        self.data = data
        self.ext = ext
        self.box = box
        self.page = page
        self.resource_id = resource_id
        self.sha1 = sha1
        self.alpha = alpha


class PageModel(object):
    __slots__ = ("index", "box", "lines", "tables", "images", "coord_mode",
                 "text_objects", "segments", "margin", "stats_images_skipped",
                 "vectors", "stamps")

    def __init__(self, index, box):
        # type: (int, Box) -> None
        self.index = index
        self.box = box
        self.lines = []          # type: List[TextLine]
        self.tables = []         # type: List[Table]
        self.images = []         # type: List[ImageItem]
        self.coord_mode = None
        self.text_objects = 0
        self.segments = []       # type: List[Seg]
        self.margin = (0.0, 0.0, 0.0, 0.0)
        self.stats_images_skipped = 0
        self.vectors = []        # type: List[Dict[str, Any]]
        self.stamps = []         # type: List[Dict[str, Any]]

    def flow_lines(self):
        # type: () -> List[TextLine]
        """排除已被表格吸收的行，返回正文行（按阅读顺序）。"""
        if not self.tables:
            return [ln for ln in self.lines if not ln.is_blank()]
        inside = []             # type: List[Tuple[float, float, float, float]]
        for t in self.tables:
            inside.append((t.box.x, t.box.y, t.box.x2, t.box.y2))
        out = []
        for ln in self.lines:
            if ln.is_blank():
                continue
            if any(b[0] <= ln.left + 0.5 and ln.right - 0.5 <= b[2] and
                   b[1] <= ln.y <= b[3] for b in inside):
                continue
            out.append(ln)
        return out


# ---------------------------------------------------------------------------
# B4  XML 工具
# ---------------------------------------------------------------------------

_ETREE = None


def _get_etree():
    # type: () -> Any
    global _ETREE
    if _ETREE is None:
        try:
            from lxml import etree
        except ImportError:
            raise OfdUnsupportedError(
                "缺少 lxml 依赖",
                hint="请执行 pip install \"lxml>=4.9,<6\"，"
                     "或 python-docx 安装时自动带上的 lxml 被卸载了")
        _ETREE = etree
    return _ETREE


def make_parser():
    # type: () -> Any
    etree = _get_etree()
    return etree.XMLParser(
        resolve_entities=False,   # 不展开实体：XXE 与实体膨胀炸弹同时失效
        no_network=True,
        load_dtd=False,
        recover=False,
        huge_tree=False,
        remove_comments=True,
        remove_pis=True,
    )


_PARSER = None


def parse_xml(data, part="(未知部件)"):
    # type: (bytes, str) -> Any
    """带安全防护的 XML 解析。"""
    global _PARSER
    head = data[:4096]
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise OfdFormatError(
            "部件包含 DTD 或实体声明，出于安全考虑拒绝解析",
            part=part,
            hint="OFD 规范不允许 DTD，该文件可能被篡改；"
                 "请从原始出处重新获取文件")
    etree = _get_etree()
    if _PARSER is None:
        _PARSER = make_parser()
    try:
        return etree.fromstring(data, parser=_PARSER)
    except etree.XMLSyntaxError as exc:
        lineno = getattr(exc, "lineno", None)
        msg = getattr(exc, "msg", str(exc))
        where = "第 %s 行" % lineno if lineno else "未知行"
        raise OfdFormatError(
            "XML 语法错误（%s）：%s" % (where, msg),
            part=part,
            hint="文件可能被截断，或编码不是 UTF-8/GB18030；"
                 "可尝试用 OFD 阅读器另存后重试")
    except Exception as exc:                      # pragma: no cover
        raise OfdFormatError("XML 解析失败：%s" % exc, part=part)


def lname(el):
    # type: (Any) -> str
    """取元素的本地名，忽略命名空间。"""
    tag = el.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def children(el, name):
    # type: (Any, str) -> List[Any]
    if el is None:
        return []
    return [c for c in el if lname(c) == name]


def child(el, name):
    # type: (Any, str) -> Any
    if el is None:
        return None
    for c in el:
        if lname(c) == name:
            return c
    return None


def iter_local(el, name):
    # type: (Any, str) -> List[Any]
    if el is None:
        return []
    return [e for e in el.iter() if lname(e) == name]


def text_of(el, default=""):
    # type: (Any, str) -> str
    if el is None or el.text is None:
        return default
    return el.text.strip()


def attr(el, name, default=None):
    # type: (Any, str, Optional[str]) -> Optional[str]
    if el is None:
        return default
    v = el.get(name)
    return default if v is None else v


def attr_float(el, name, default=0.0):
    # type: (Any, str, float) -> float
    v = attr(el, name)
    if v is None:
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def attr_true(el, name):
    # type: (Any, str) -> bool
    v = attr(el, name)
    return v is not None and str(v).strip().lower() in ("true", "1")


_ARRAY_SPLIT = re.compile(r"[,\s]+")


def parse_array(text):
    # type: (Optional[str]) -> List[float]
    """解析 ST_Array，支持 "g N X" 压缩写法（N 个 X）与逗号分隔。"""
    if not text:
        return []
    toks = [t for t in _ARRAY_SPLIT.split(str(text).strip()) if t]
    out = []            # type: List[float]
    i = 0
    n = len(toks)
    while i < n:
        tok = toks[i]
        if tok in ("g", "G") and i + 2 < n:
            try:
                cnt = int(float(toks[i + 1]))
                val = float(toks[i + 2])
            except (TypeError, ValueError):
                i += 1
                continue
            if cnt > 0:
                out.extend([val] * cnt)
            i += 3
            continue
        try:
            out.append(float(tok))
        except (TypeError, ValueError):
            pass
        i += 1
    return out


def is_wide(ch):
    # type: (str) -> bool
    """全角判定。A（Ambiguous）在中文文档里按全角处理，
    否则 "“" 会被当成半角，导致引号后凭空多出空格。"""
    if not ch:
        return False
    return unicodedata.east_asian_width(ch) in ("W", "F", "A")


def est_advance(ch, size_mm):
    # type: (str, float) -> float
    """在缺少 DeltaX 时估算单字步进（mm）。"""
    if not ch:
        return 0.0
    if ch == " ":
        return 0.28 * size_mm
    if ch == "\u3000":
        return 1.0 * size_mm
    if is_wide(ch):
        return 1.0 * size_mm
    if ch.isdigit():
        # Times 系字体的数字是 0.5em；按 0.56 估会把页码 "— 1 —" 的间隙吃掉
        return 0.50 * size_mm
    if "A" <= ch <= "Z":
        return 0.62 * size_mm
    if "a" <= ch <= "z":
        return 0.50 * size_mm
    if ch in ".,:;'!|":
        return 0.28 * size_mm
    if ch in "-/":
        return 0.33 * size_mm
    return 0.60 * size_mm


def parse_color_value(value, kind="RGB"):
    # type: (Optional[str], str) -> Tuple[int, int, int]
    """把颜色分量解析为 RGB 三元组。

    支持 "#RRGGBB"、0-255 整数分量、0-1 归一化分量、CMYK 四分量。
    """
    if not value:
        return (0, 0, 0)
    text = str(value).strip()
    if text.startswith("#"):
        hexed = text[1:]
        if len(hexed) == 6:
            try:
                return (int(hexed[0:2], 16), int(hexed[2:4], 16), int(hexed[4:6], 16))
            except ValueError:
                return (0, 0, 0)
        if len(hexed) == 3:
            try:
                return tuple(int(c * 2, 16) for c in hexed)  # type: ignore
            except ValueError:
                return (0, 0, 0)
    comps = parse_array(text)
    if not comps:
        return (0, 0, 0)
    if kind.upper() == "CMYK" or len(comps) >= 4:
        vals = list(comps[:4]) + [0.0] * (4 - min(4, len(comps)))
        c, m, y, k = [v / 100.0 if v > 1.0 else v for v in vals]
        return (int(255 * (1 - c) * (1 - k) + 0.5),
                int(255 * (1 - m) * (1 - k) + 0.5),
                int(255 * (1 - y) * (1 - k) + 0.5))
    if len(comps) == 1 or kind.upper() == "GRAY":
        g = comps[0]
        g = int(g * 255 + 0.5) if g <= 1.0 else int(g + 0.5)
        g = max(0, min(255, g))
        return (g, g, g)
    rgb = list(comps[:3])
    if all(abs(v) <= 1.0 for v in rgb) and any(abs(v) > 0.0 for v in rgb):
        rgb = [v * 255.0 for v in rgb]
    return tuple(max(0, min(255, int(v + 0.5))) for v in rgb)  # type: ignore


# ---------------------------------------------------------------------------
# B5  包访问
# ---------------------------------------------------------------------------

ENCRYPT_MARKERS = ("Encryptions.xml", "decryptseed.dat", "entriesmap.dat",
                   "entriesmap.dat".upper())

MAX_ENTRIES = 20000
MAX_PART_BYTES = 256 * 1024 * 1024
MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024


class OfdPackage(object):
    """OFD 包（ZIP 容器）的只读访问封装。"""

    def __init__(self, path):
        # type: (str) -> None
        try:
            self._zf = zipfile.ZipFile(path, "r")
        except zipfile.BadZipFile:
            raise OfdFormatError(
                "该文件不是有效的 OFD（ZIP 容器损坏）",
                part=os.path.basename(path),
                hint="OFD 本质是 ZIP 包。请确认文件没有被截断、"
                     "没有被当成 HTML/PDF 下载（可先用 7-Zip 试着打开）")
        except OSError as exc:
            raise OfdFormatError("无法打开文件：%s" % exc, part=path)
        self.path = path
        self._infos = {}
        self._lower = {}
        total = 0
        for info in self._zf.infolist():
            name = info.filename
            if name.endswith("/"):
                continue
            if name.startswith("__MACOSX/") or name.endswith("Thumbs.db"):
                continue
            self._infos[name] = info
            self._lower[name.lower()] = name
            total += max(0, info.file_size)
        if len(self._infos) > MAX_ENTRIES:
            raise OfdUnsupportedError(
                "包内条目过多（%d 个），疑似异常文件" % len(self._infos))
        if total > MAX_TOTAL_BYTES:
            raise OfdUnsupportedError(
                "解压后体积超过 2GB，已中止处理")
        self._check_encrypted()

    # -- 加密检测 ---------------------------------------------------------
    def _check_encrypted(self):
        # type: () -> None
        for marker in ENCRYPT_MARKERS:
            for name in self._infos:
                if posixpath.basename(name).lower() == marker.lower():
                    raise OfdEncryptedError(
                        "该 OFD 已加密（检测到 %s）" % posixpath.basename(name),
                        hint="本工具不处理加密文档。请用 OFD 阅读器打开并"
                             "输入口令后“另存为”未加密的 OFD，或先导出 PDF 再转。")

    # -- 基本访问 ---------------------------------------------------------
    def exists(self, name):
        # type: (str) -> bool
        return name in self._infos

    def names(self):
        # type: () -> List[str]
        return list(self._infos.keys())

    def has_root_ofd(self):
        # type: () -> bool
        return "OFD.xml" in self._infos or "ofd.xml" in self._lower

    def read(self, name):
        # type: (bytes or str) -> bytes
        real = self._real_name(name)
        info = self._infos[real]
        if info.file_size > MAX_PART_BYTES:
            raise OfdUnsupportedError(
                "单个部件过大（%.1f MB），已中止" % (info.file_size / 1048576.0),
                part=real)
        try:
            return self._zf.read(real)
        except (KeyError, zipfile.BadZipFile) as exc:
            raise OfdFormatError("读取部件失败：%s" % exc, part=real)

    def _real_name(self, name):
        # type: (str) -> str
        if name in self._infos:
            return name
        alt = self._lower.get(str(name).lower())
        if alt is not None:
            return alt
        raise OfdFormatError("包内不存在该部件", part=str(name))

    def resolve(self, loc, base_part):
        # type: (Optional[str], str) -> str
        """把 ST_Loc 解析成包内绝对路径。相对路径的基准是"引用它的部件"所在目录。"""
        if not loc:
            raise OfdFormatError("空的部件引用（BaseLoc 为空）", part=base_part)
        text = str(loc).strip().replace("\\", "/")
        if text.startswith("/"):
            cand = text.lstrip("/")
        else:
            base_dir = posixpath.dirname(base_part.replace("\\", "/"))
            cand = posixpath.normpath(posixpath.join(base_dir, text))
        if cand.startswith("../") or cand.startswith("/"):
            raise OfdFormatError(
                "部件引用越出包根：%s" % loc, part=base_part,
                hint="该 OFD 的路径写法不符合规范，可能被人工改动过")
        try:
            return self._real_name(cand)
        except OfdFormatError:
            raise OfdFormatError(
                "引用的部件不存在：%s（解析为 %s）" % (loc, cand),
                part=base_part,
                hint="该 OFD 可能未完整打包，或使用了非标准路径分隔符；"
                     "请用 OFD 阅读器另存后重试")

    def resolve_in(self, base_dir, loc):
        # type: (str, str) -> str
        """在指定目录下解析相对路径（用于 ResourceTable 里记录的 BaseLoc）。"""
        if not loc:
            raise OfdFormatError("空的资源文件引用", part=base_dir)
        text = str(loc).strip().replace("\\", "/")
        if text.startswith("/"):
            cand = text.lstrip("/")
        else:
            cand = posixpath.normpath(posixpath.join(base_dir or "", text))
        if cand.startswith("../"):
            raise OfdFormatError("资源引用越出包根：%s" % loc, part=base_dir)
        try:
            return self._real_name(cand)
        except OfdFormatError:
            raise OfdFormatError(
                "资源文件不存在：%s（解析为 %s）" % (loc, cand), part=base_dir,
                hint="该 OFD 可能未完整打包")

    def close(self):
        # type: () -> None
        try:
            self._zf.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# B6  资源表
# ---------------------------------------------------------------------------

# OFD 里常见的内嵌字体名（多为全小写合成名）到系统中文字体名的映射
FONT_NAME_MAP = {
    "simsun": "宋体",
    "nsimsun": "新宋体",
    "simhei": "黑体",
    "simkai": "楷体",
    "kaiti": "楷体",
    "kaiti_gb2312": "楷体_GB2312",
    "fangsong": "仿宋",
    "fangsong_gb2312": "仿宋_GB2312",
    "fzssk--gbk1-0": "方正书宋简体",
    "fzxbsk--gbk1-0": "方正小标宋简体",
    "fzfsk--gbk1-0": "方正仿宋简体",
    "fzhtk--gbk1-0": "方正黑体简体",
    "fzktk--gbk1-0": "方正楷体简体",
    "timesnewroman": "Times New Roman",
    "timesnewromanpsmt": "Times New Roman",
    "timesnewromanps-boldmt": "Times New Roman",
    "arial": "Arial",
    "arialmt": "Arial",
    "calibri": "Calibri",
    "cambria": "Cambria",
    "couriernew": "Courier New",
    "helvetica": "Helvetica",
}

_FONT_SUFFIX_RE = re.compile(
    r"[-_ ]?(psmt|mt|regular|bold|italic|boldmt|italicmt|light|medium)$")

CJK_FALLBACK = "宋体"
LATIN_FALLBACK = "Times New Roman"


def normalize_font_name(raw):
    # type: (Optional[str]) -> Tuple[str, bool]
    """把 OFD 里的字体名归一为可用的字体名。

    返回 (字体名, 是否已识别)。中文名同时可用于 w:eastAsia。
    """
    if not raw:
        return (CJK_FALLBACK, False)
    key = str(raw).strip()
    low = key.lower()
    if low in FONT_NAME_MAP:
        return (FONT_NAME_MAP[low], True)
    stripped = _FONT_SUFFIX_RE.sub("", low)
    if stripped in FONT_NAME_MAP:
        return (FONT_NAME_MAP[stripped], True)
    # 形如 "fzssk--gbk1-0" 的方正系字体，保留原名（用户机器上装了才有意义）
    for candidate, chinese in FONT_NAME_MAP.items():
        if stripped.replace("-", "") == candidate.replace("-", ""):
            return (chinese, True)
    return (key, False)


class FontInfo(object):
    __slots__ = ("name", "latin", "cjk", "bold", "italic", "known")

    def __init__(self, name, latin, cjk, bold, italic, known):
        self.name = name
        self.latin = latin
        self.cjk = cjk
        self.bold = bold
        self.italic = italic
        self.known = known


class ResourceTable(object):
    """按 公共资源 → 文档资源 → 页面资源 三层合并的资源表。"""

    def __init__(self):
        # type: () -> None
        self.fonts = {}        # type: Dict[str, Any]
        self.colors = {}       # type: Dict[str, str]
        self.params = {}       # type: Dict[str, Any]
        self.media = {}        # type: Dict[str, Any]
        self.font_files = {}   # type: Dict[str, Tuple[str, str]]
        self._font_cache = {}  # type: Dict[Tuple[str, str], FontInfo]
        self._unmapped = set()
        self._pkg = None

    def merge_res(self, pkg, res_part, report):
        # type: (OfdPackage, str, Report) -> None
        data = pkg.read(res_part)
        root = parse_xml(data, res_part)
        base_loc = attr(root, "BaseLoc") or "."
        base_dir = posixpath.join(posixpath.dirname(res_part),
                                  base_loc.strip().replace("\\", "/"))
        base_dir = posixpath.normpath(base_dir)
        self.merge_res_root(root, base_dir, report, pkg)

    def merge_res_root(self, root, base_dir, report=None, pkg=None):
        # type: (Any, str, Optional[Report], Optional[OfdPackage]) -> None
        """合并一个 Res 元素下的资源（也用于页面级内联 PageRes）。"""
        if pkg is not None:
            self._pkg = pkg
        for el in children(root, "Fonts"):
            for font in children(el, "Font"):
                fid = attr(font, "ID")
                if fid is not None:
                    self.fonts[fid] = font
                self._register_font_file(font, base_dir)
                for cfont in children(font, "CompositeFont"):
                    cid = attr(cfont, "ID")
                    if cid is not None:
                        self.fonts[cid] = cfont
                for sfont in children(font, "SubFonts"):
                    for sub in children(sfont, "SubFont"):
                        sid = attr(sub, "ID")
                        if sid is not None:
                            self.fonts.setdefault(sid, font)

        for el in children(root, "ColorSpaces"):
            for cs in children(el, "ColorSpace"):
                cid = attr(cs, "ID")
                if cid is None:
                    continue
                palette = []
                pal = child(cs, "Palette")
                if pal is not None:
                    for entry in children(pal, "PaletteEntry"):
                        palette.append(
                            parse_color_value(attr(entry, "Value"), "RGB"))
                self.colors[cid] = {"type": attr(cs, "Type") or "RGB",
                                    "palette": palette,
                                    "bits": attr(cs, "BitsPerComponent")}

        for el in children(root, "DrawParams"):
            for dp in children(el, "DrawParam"):
                did = attr(dp, "ID")
                if did is not None:
                    self.params[did] = dp

        for el in children(root, "MultiMedias"):
            for mm in children(el, "MultiMedia"):
                mid = attr(mm, "ID")
                if mid is None:
                    continue
                mf = child(mm, "MediaFile")
                loc = text_of(mf) if mf is not None else ""
                self.media[mid] = {
                    "id": mid,
                    "type": attr(mm, "Type") or "Image",
                    "format": (attr(mm, "Format") or "").upper(),
                    "base_dir": base_dir,
                    "loc": loc,
                }
        self._font_cache.clear()

    # -- 字色 -------------------------------------------------------------
    def color_space_type(self, ref):
        # type: (Optional[str]) -> str
        if ref is None:
            return "RGB"
        info = self.colors.get(str(ref))
        if isinstance(info, dict):
            return str(info.get("type") or "RGB")
        return str(info) if info else "RGB"

    def color_from(self, value, ref=None):
        # type: (Optional[str], Optional[str]) -> Tuple[int, int, int]
        """按颜色空间解析颜色。

        Index / Separation 的 Value 是调色板下标，不是 RGB 分量：
        以前直接当 RGB 读，红色会解析成 (1, 0, 0) 这种近乎黑的颜色，
        这也是同一份文件在别的阅读器里是红头、在本程序里发黑的原因。
        """
        kind = self.color_space_type(ref)
        if kind in ("Index", "Separation"):
            info = self.colors.get(str(ref)) if ref is not None else None
            palette = info.get("palette") if isinstance(info, dict) else None
            try:
                index = int(parse_array(value)[0])
            except Exception:
                return (0, 0, 0)
            if palette and 0 <= index < len(palette):
                return palette[index]
            return (0, 0, 0)
        return parse_color_value(value, kind)

    def color_of(self, fill_el, dp_el, depth=0):
        # type: (Any, Any, int) -> Tuple[int, int, int]
        """按 对象自身 → DrawParam 继承链 的顺序取填充色。"""
        if fill_el is not None:
            return self.color_from(attr(fill_el, "Value"),
                                   attr(fill_el, "ColorSpace"))
        if dp_el is None or depth > 8:
            return (0, 0, 0)
        fc = child(dp_el, "FillColor")
        if fc is not None:
            return self.color_from(attr(fc, "Value"),
                                   attr(fc, "ColorSpace"))
        rel = attr(dp_el, "Relative")
        if rel is not None:
            parent = self.params.get(str(rel))
            return self.color_of(None, parent, depth + 1)
        return (0, 0, 0)

    def object_color(self, obj, name, dp_el):
        # type: (Any, str, Any) -> Tuple[int, int, int]
        """取对象某个颜色属性（FillColor / StrokeColor），支持 DrawParam 继承链。"""
        el = child(obj, name)
        if el is not None:
            return self.color_from(attr(el, "Value"),
                                   attr(el, "ColorSpace"))
        node = dp_el
        depth = 0
        while node is not None and depth < 8:
            e = child(node, name)
            if e is not None:
                return self.color_from(attr(e, "Value"),
                                       attr(e, "ColorSpace"))
            rel = attr(node, "Relative")
            node = self.params.get(str(rel)) if rel is not None else None
            depth += 1
        return (0, 0, 0)

    def draw_param(self, ref):
        # type: (Optional[str]) -> Any
        if ref is None:
            return None
        return self.params.get(str(ref))

    def line_width_of(self, obj, dp_el):
        # type: (Any, Any) -> float
        lw = attr_float(obj, "LineWidth", -1.0)
        if lw >= 0:
            return lw
        depth = 0
        node = dp_el
        while node is not None and depth < 8:
            v = attr_float(node, "LineWidth", -1.0)
            if v >= 0:
                return v
            rel = attr(node, "Relative")
            node = self.params.get(str(rel)) if rel is not None else None
            depth += 1
        return 0.35

    # -- 字体 -------------------------------------------------------------
    def font(self, ref):
        # type: (Optional[str]) -> FontInfo
        cache_key = (str(ref), "")
        hit = self._font_cache.get(cache_key)
        if hit is not None:
            return hit
        el = self.fonts.get(str(ref)) if ref is not None else None
        raw_name = None
        family = None
        bold = False
        italic = False
        if el is not None:
            raw_name = attr(el, "FontName") or attr(el, "FamilyName")
            family = attr(el, "FamilyName")
            bold = attr_true(el, "Bold") or attr_true(el, "Weight")
            italic = attr_true(el, "Italic")
        name, known = normalize_font_name(raw_name)
        if family and not known:
            name2, known2 = normalize_font_name(family)
            if known2:
                name, known = name2, True
        latin = name
        cjk = name if _looks_cjk_font(name) else CJK_FALLBACK
        if not _looks_cjk_font(name) and not known:
            cjk = name
        info = FontInfo(name, latin, cjk, bold, italic, known)
        if not known and raw_name and raw_name not in self._unmapped:
            self._unmapped.add(raw_name)
        self._font_cache[cache_key] = info
        return info

    def _register_font_file(self, font_el, base_dir):
        # type: (Any, str) -> None
        """记录字体文件位置，供阅读器取内嵌字体渲染。"""
        ff = child(font_el, "FontFile")
        loc = text_of(ff) if ff is not None else ""
        if not loc:
            for sub in font_el.iter():
                if lname(sub) == "FontFile" and text_of(sub):
                    loc = text_of(sub)
                    break
        if not loc:
            return
        raw = attr(font_el, "FontName") or attr(font_el, "FamilyName") or ""
        mapped, _known = normalize_font_name(raw)
        keys = []          # type: List[str]
        if raw:
            keys.append(raw.strip())
            keys.append(raw.strip().lower())
        if mapped:
            keys.append(mapped)
        for key in keys:
            if key:
                self.font_files[key] = (base_dir, loc)

    def font_bytes_for(self, name):
        # type: (Optional[str]) -> Optional[bytes]
        """按字体名取内嵌字体字节；取不到返回 None。"""
        if not name or self._pkg is None:
            return None
        key = str(name)
        src = self.font_files.get(key)
        if src is None:
            src = self.font_files.get(key.lower())
        if src is None:
            return None
        base_dir, loc = src
        try:
            return self._pkg.read(self._pkg.resolve_in(base_dir, loc))
        except OfdError:
            return None

    def unmapped_fonts(self):
        # type: () -> List[str]
        return sorted(self._unmapped)

    # -- 媒体 -------------------------------------------------------------
    def media_info(self, ref):
        # type: (Optional[str]) -> Optional[Dict[str, Any]]
        if ref is None:
            return None
        return self.media.get(str(ref))


def _looks_cjk_font(name):
    # type: (str) -> bool
    if not name:
        return False
    if any(ord(c) > 0x2E80 for c in name):
        return True
    low = name.lower()
    for token in ("song", "hei", "kai", "fang", "sun", "ming", "yahei",
                  "gb2312", "gbk", "cjk", "sc", "fz"):
        if token in low:
            return True
    return False


# ---------------------------------------------------------------------------
# B7  文档装配
# ---------------------------------------------------------------------------

class DocumentMeta(object):
    __slots__ = ("version", "doc_bodies", "page_box", "pages", "title",
                 "stamps_by_id")

    def __init__(self):
        self.version = None
        self.doc_bodies = 0
        self.page_box = Box(0.0, 0.0, 210.0, 297.0)
        self.pages = []      # type: List[Tuple[str, Optional[str], Any]]
        self.title = ""
        self.stamps_by_id = {}   # type: Dict[str, List[Dict[str, Any]]]


def open_package(path):
    # type: (str) -> OfdPackage
    pkg = OfdPackage(path)
    if not pkg.has_root_ofd():
        names = pkg.names()
        raise OfdFormatError(
            "包内缺少 OFD.xml 入口文件",
            hint="这通常不是 OFD 文件，或其内部结构被破坏。"
                 "当前包内前几个条目：%s" % "、".join(names[:5] or ["(空包)"]))
    return pkg


def read_entry(pkg):
    # type: (OfdPackage) -> Tuple[Any, str]
    part = "OFD.xml" if pkg.exists("OFD.xml") else "ofd.xml"
    root = parse_xml(pkg.read(part), part)
    if lname(root) != "OFD":
        raise OfdFormatError(
            "OFD.xml 的根元素是 <%s>，应为 <OFD>" % lname(root), part=part)
    return root, part


def load_document(pkg, report):
    # type: (OfdPackage, Report) -> Tuple[DocumentMeta, Any, str, ResourceTable]
    """解析 OFD.xml 与 Document.xml，返回 (文档元信息, Document 根, 部件名, 资源表)。"""
    meta = DocumentMeta()
    ofd_root, ofd_part = read_entry(pkg)
    meta.version = attr(ofd_root, "Version")

    bodies = children(ofd_root, "DocBody")
    meta.doc_bodies = len(bodies)
    if not bodies:
        raise OfdFormatError("OFD.xml 中没有任何 DocBody", part=ofd_part,
                             hint="文件可能被截断")
    if len(bodies) > 1:
        report.info("MULTI_DOCBODY",
                    "该文件包含 %d 个文档版本（DocBody），已转换第 1 个，"
                    "其余 %d 个未处理" % (len(bodies), len(bodies) - 1))
    body = bodies[0]

    root_el = child(body, "DocRoot")
    if root_el is None or not text_of(root_el):
        versions = children(body, "Versions")
        if versions:
            last = versions[-1]
            for v in reversed(versions):
                cand = child(v, "DocRoot")
                if cand is not None and text_of(cand):
                    last = v
                    break
            root_el = child(last, "DocRoot")
        if root_el is None or not text_of(root_el):
            raise OfdFormatError("DocBody 中没有指向文档根的 DocRoot", part=ofd_part)

    doc_part = pkg.resolve(text_of(root_el), ofd_part)
    doc_root = parse_xml(pkg.read(doc_part), doc_part)
    if lname(doc_root) != "Document":
        raise OfdFormatError(
            "文档根元素是 <%s>，应为 <Document>" % lname(doc_root), part=doc_part)

    info = child(body, "DocInfo")
    if info is not None:
        meta.title = text_of(child(info, "Title"))

    # 页面尺寸（文档级默认）
    common = child(doc_root, "CommonData")
    page_area = child(common, "PageArea") if common is not None else None
    if page_area is not None:
        pb = parse_box(text_of(child(page_area, "PhysicalBox")))
        if pb is not None:
            meta.page_box = pb

    # 资源表：公共 → 文档
    res = ResourceTable()
    if common is not None:
        for tag in ("PublicRes", "DocumentRes"):
            el = child(common, tag)
            if el is not None and text_of(el):
                part = pkg.resolve(text_of(el), doc_part)
                try:
                    res.merge_res(pkg, part, report)
                except OfdFormatError as exc:
                    report.warn("BAD_RES", "资源文件解析失败，已忽略：%s"
                                % exc.message, obj_id=str(exc.part))

    # 模板页（正文之前绘制，用于页眉、水印、边框）
    templates = {}       # type: Dict[str, str]
    if common is not None:
        for tp in children(common, "TemplatePage"):
            tid = attr(tp, "ID")
            loc = attr(tp, "BaseLoc")
            if tid is not None and loc:
                try:
                    templates[str(tid)] = pkg.resolve(loc, doc_part)
                except OfdError as exc:
                    report.warn("TEMPLATE_MISSING",
                                "模板页引用无效，已忽略：%s" % exc.message)

    # 页树（规范要求按前序遍历取叶子节点）
    pages_el = child(doc_root, "Pages")
    page_locs = []       # type: List[Tuple[str, Optional[str]]]
    if pages_el is not None:
        _walk_pages(pkg, doc_part, pages_el, page_locs, report, depth=0)
    resolved = []        # type: List[Tuple[str, Optional[str], Any]]
    for (part, tid, page_id) in page_locs:
        tpl = None
        if tid is not None:
            tpl = templates.get(str(tid))
            if tpl is None and templates:
                report.warn("TEMPLATE_UNKNOWN",
                            "页面引用了不存在的模板页（TemplateID=%s），已忽略" % tid)
        resolved.append((part, tpl, page_id))
    meta.pages = resolved

    # 大纲与签章只统计
    outlines = child(doc_root, "Outlines")
    if outlines is not None:
        n = len([e for e in outlines.iter() if lname(e) == "OutlineElem"])
        report.meta["outlines_skipped"] = n
        if n:
            report.info("OUTLINES_SKIPPED",
                        "已跳过 %d 条大纲/书签（OFD 大纲不转换）" % n)

    return meta, ofd_root, ofd_part, doc_root, doc_part, res


def _walk_pages(pkg, base_part, node, out, report, depth):
    # type: (OfdPackage, str, Any, List[Tuple[str, Optional[str]]], Report, int) -> None
    if depth > 32:
        report.warn("PAGE_TREE_DEEP", "页树嵌套过深，已停止遍历")
        return
    for el in node:
        name = lname(el)
        if name == "Page":
            loc = attr(el, "BaseLoc")
            if loc:
                out.append((pkg.resolve(loc, base_part),
                            attr(el, "TemplateID"),
                            attr(el, "ID")))
        elif name == "Pages":
            _walk_pages(pkg, base_part, el, out, report, depth + 1)


SEAL_MAGICS = ((b"\x89PNG\r\n\x1a\n", "png"), (b"GIF89a", "gif"),
               (b"GIF87a", "gif"), (b"\xff\xd8\xff", "jpg"))


def extract_embedded_image(blob):
    # type: (bytes) -> Tuple[Optional[bytes], str]
    """从签章数据里挖出印章图片。

    SES 签章把印章图片以 OCTET STRING 放进签名值里，不同厂商外层结构差异很大，
    直接按魔数在二进制里定位比完整解析 ASN.1 更耐操。
    """
    if not blob:
        return (None, "")
    for magic, ext in SEAL_MAGICS:
        pos = blob.find(magic)
        if pos >= 0 and len(blob) - pos > 128:
            return (blob[pos:], ext)
    return (None, "")


def collect_stamps(pkg, ofd_root, ofd_part, report):
    # type: (OfdPackage, Any, str, Report) -> Dict[str, List[Dict[str, Any]]]
    """解析签章里的印章批注，返回 页 ID -> [{box, data, ext}]。"""
    stamps = {}      # type: Dict[str, List[Dict[str, Any]]]
    sig_total = 0
    for body in children(ofd_root, "DocBody"):
        el = child(body, "Signatures")
        if el is None or not text_of(el):
            continue
        try:
            part = pkg.resolve(text_of(el), ofd_part)
            root = parse_xml(pkg.read(part), part)
        except OfdError:
            continue
        for sig in iter_local(root, "Signature"):
            sig_total += 1
            loc = attr(sig, "BaseLoc")
            if not loc:
                continue
            try:
                sig_part = pkg.resolve(loc, part)
                sroot = parse_xml(pkg.read(sig_part), sig_part)
            except OfdError as exc:
                report.warn("SIGN_UNREADABLE",
                            "签名部件读取失败，已跳过：%s" % exc.message)
                continue
            annots = iter_local(sroot, "StampAnnot")
            if not annots:
                report.info("SIGN_NO_STAMP",
                            "该签名不含可见印章批注（StampAnnot）")
                continue
            blob = b""
            sv = child(sroot, "SignedValue")
            if sv is not None and text_of(sv):
                try:
                    blob = pkg.read(pkg.resolve(text_of(sv), sig_part))
                except OfdError as exc:
                    report.warn("SEAL_UNREADABLE",
                                "签章数据读取失败：%s" % exc.message)
            data, ext = extract_embedded_image(blob)
            if data is None:
                report.warn("SEAL_NO_IMAGE",
                            "签章里没有可用的印章图片（已尝试 GIF/PNG/JPEG），"
                            "该印章无法还原")
                continue
            for annot in annots:
                box = parse_box(attr(annot, "Boundary"))
                page_ref = attr(annot, "PageRef")
                if box is None or page_ref is None:
                    continue
                stamps.setdefault(str(page_ref), []).append(
                    {"box": box, "data": data, "ext": ext})
    if sig_total:
        report.meta["signatures_skipped"] = sig_total
        visible = sum(len(v) for v in stamps.values())
        if visible:
            report.info("SEALS_FOUND",
                        "识别到 %d 个电子印章（共 %d 个签名），将一并绘制"
                        % (visible, sig_total))
        else:
            report.info("SIGNATURES_SKIPPED",
                        "已跳过 %d 个数字签名（无可见印章图形）" % sig_total)
    return stamps


class CoordResolver(object):
    """判定 TextCode 的 X/Y 属于哪一种坐标解释。

    候选：
        A  P = CTM · p
        B  P = Boundary.origin + p          （国产生成器最常见）
        C  P = p                            （绝对页面坐标）
        D  P = Boundary.origin + Linear(CTM) · p
    先过"落在页内"的门，再用与 Boundary 的 IoU 决胜——因为错误的绝对坐标
    解释同样落在页面左上角区域，只看在页率区分不开。
    """

    MODES = ("B", "A", "D", "C")

    def __init__(self, page_box, report, page_no):
        # type: (Box, Report, int) -> None
        self.page_box = page_box
        self.report = report
        self.page_no = page_no
        self.votes = {}      # type: Dict[str, int]

    def dominant_mode(self):
        # type: () -> Optional[str]
        """本页已经形成多数派时返回它，否则 None。

        同一页里的文字对象应当采用同一种坐标解释；某些大字号标题对象在两种
        解释下 IoU 很接近，逐对象独立判定会各自选边，结果同一行的字落在不同
        基线上（看起来就是"红头错乱"）。这里用多数派做一致性约束。
        """
        total = sum(self.votes.values())
        if total < 3:
            return None
        mode, count = max(self.votes.items(), key=lambda kv: kv[1])
        if count / float(total) >= 0.6:
            return mode
        return None

    def transform(self, mode, x, y, box, ctm):
        # type: (str, float, float, Optional[Box], Optional[Sequence[float]]) -> Tuple[float, float]
        if mode == "C":
            return (x, y)
        if mode == "B":
            if box is None:
                return (x, y)
            return (box.x + x, box.y + y)
        if mode == "A":
            if ctm is None:
                return (x, y)
            return apply_matrix(ctm, x, y)
        if mode == "D":
            if box is None or ctm is None:
                return (x, y)
            lin = (ctm[0], ctm[1], ctm[2], ctm[3], 0.0, 0.0)
            px, py = apply_matrix(lin, x, y)
            return (box.x + px, box.y + py)
        return (x, y)

    def probe(self, samples, box, ctm, obj_id, size_mm, quiet=False):
        # type: (List[Tuple[float, float, float]], Optional[Box], Optional[Sequence[float]], Any, float, bool) -> str
        """samples: [(x, y, advance)]，返回选定的模式。

        quiet=True 时不产生告警，用于 Boundary 明显带内边距的短文本与纯空格对象——
        这些对象与外框的 IoU 天然偏低，报出来只是噪声。
        """
        if not samples:
            return "B"
        size = max(size_mm, 0.5)
        up = 0.88 * size      # 基线以上大致是字高
        down = 0.22 * size    # 基线以下留出降部
        if matrix_is_identity(ctm):
            allowed = ("B", "C") if box is not None else ("C",)
        else:
            allowed = tuple(m for m in self.MODES
                            if (m in ("A",) and ctm) or
                            (m == "B" and box is not None) or
                            (m == "C") or
                            (m == "D" and box is not None and ctm))
        pb = self.page_box
        best_mode = None
        best_score = -1.0
        scores = {}       # type: Dict[str, float]
        for mode in allowed:
            xs = []
            ys = []
            for (x, y, adv) in samples:
                px, py = self.transform(mode, x, y, box, ctm)
                xs.append(px)
                xs.append(px + max(adv, 0.0))
                ys.append(py - up)
                ys.append(py + down)
            inside = 0
            total = 0
            for i in range(0, len(xs), 2):
                for v, lo, hi in ((xs[i], 0.0, pb.w), (ys[i], 0.0, pb.h)):
                    total += 1
                    if (lo - TOL_ON_PAGE) <= v <= (hi + TOL_ON_PAGE):
                        inside += 1
            ratio = inside / float(total) if total else 0.0
            if ratio < 0.85:
                continue
            gb = Box(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
            score = box_iou(gb, box) if box is not None else 0.5
            scores[mode] = score
            if score > best_score or (abs(score - best_score) < 1e-9 and best_mode is None):
                best_score = score
                best_mode = mode

        # 页内一致性：如果本对象没有明显优势，就跟随本页的多数派，
        # 避免同一行文字被拆到不同基线上。
        dominant = self.dominant_mode()
        if (dominant is not None and best_mode is not None
                and dominant != best_mode and dominant in scores):
            if scores[dominant] >= max(0.60, best_score * 0.85):
                best_mode = dominant
                best_score = scores[dominant]
        if best_mode is None:
            # 兜底：全页多数票
            best_mode = max(self.votes.items(), key=lambda kv: kv[1])[0] \
                if self.votes else ("B" if box is not None else "C")
            if not quiet:
                self.report.warn(
                    "COORD_FALLBACK",
                    "对象 ID=%s 的文字坐标无法可靠定位，已按 %s 解释输出，位置可能偏移"
                    % (obj_id, best_mode), page=self.page_no, obj_id=obj_id)
        elif best_score < 0.35 and not quiet:
            self.report.warn(
                "COORD_UNCERTAIN",
                "对象 ID=%s 的坐标解释置信度偏低（IoU=%.2f），已按 %s 解释输出"
                % (obj_id, best_score, best_mode),
                page=self.page_no, obj_id=obj_id)
        self.votes[best_mode] = self.votes.get(best_mode, 0) + 1
        return best_mode


def _expand_deltas(dx, n, text, size_mm):
    # type: (List[float], int, str, float) -> List[float]
    """把 DeltaX 对齐到字符数。

    实测：有的生成器只给出 n-1 个（最后一字用中位值），
    有的只给 n-k 个（规范行为是末值补齐）。这里两种都兼容。
    """
    if n <= 0:
        return []
    if not dx:
        return [est_advance(c, size_mm) for c in text]
    if len(dx) == n + 1:
        dx = dx[:n]
    if len(dx) == n - 1 and len(dx) > 0:
        vals = list(dx) + [statistics.median(dx)]
    elif len(dx) < n:
        vals = list(dx) + [dx[-1]] * (n - len(dx))
    else:
        vals = list(dx[:n])
    return vals


def _text_codes(t_obj):
    # type: (Any) -> List[Any]
    out = []
    for c in t_obj:
        if lname(c) == "TextCode":
            out.append(c)
    return out


def build_glyphs(t_obj, res, page_box, page_no, report, resolver, seq_start):
    # type: (Any, ResourceTable, Box, int, Report, CoordResolver, int) -> List[Glyph]
    obj_id = attr(t_obj, "ID") or "?"
    if attr(t_obj, "Visible") is not None and not attr_true(t_obj, "Visible"):
        return []
    if attr(t_obj, "Alpha") is not None:
        try:
            if float(attr(t_obj, "Alpha")) <= 0.01:
                return []
        except (TypeError, ValueError):
            pass
    size_mm = attr_float(t_obj, "Size", 0.0)
    if size_mm <= 0:
        report.warn("TEXT_NO_SIZE",
                    "文字对象缺少字号（Size），已跳过", page=page_no, obj_id=obj_id)
        return []

    hscale = attr_float(t_obj, "HScale", 1.0) or 1.0
    font = res.font(attr(t_obj, "Font"))
    dp = res.draw_param(attr(t_obj, "DrawParam"))
    fill_el = child(t_obj, "FillColor")
    color = res.color_of(fill_el, dp)
    bold = font.bold or (attr_float(t_obj, "Weight", 400.0) >= 600.0)
    italic = font.italic or attr_true(t_obj, "Italic")

    # 竖排/旋转：不做还原，只告警
    direction = attr(t_obj, "CharDirection") or attr(t_obj, "ReadDirection")
    if direction and str(direction).strip() not in ("0", ""):
        report.warn("TEXT_DIRECTION",
                    "文字对象为竖排/旋转排版（CharDirection=%s），已按横排输出"
                    % direction, page=page_no, obj_id=obj_id)
    rotate = attr_float(t_obj, "Rotate", 0.0)
    if abs(rotate) > 0.01 and abs(abs(rotate) - 360.0) > 0.01:
        report.warn("TEXT_ROTATE",
                    "文字对象带旋转（Rotate=%s°），已按未旋转输出" % rotate,
                    page=page_no, obj_id=obj_id)

    boundary = parse_box(attr(t_obj, "Boundary"))
    ctm = parse_array(attr(t_obj, "CTM"))
    if len(ctm) < 6:
        ctm = None

    codes = _text_codes(t_obj)
    if not codes:
        if child(t_obj, "Glyphs") is not None:
            report.warn("TEXT_GLYPHS_ONLY",
                        "文字对象只以字形索引（Glyphs）描述、没有 Unicode 文本，"
                        "无法还原，已跳过", page=page_no, obj_id=obj_id)
        return []

    # 第一遍：把每个 TextCode 展开成 (x, y, text, deltas)
    prepared = []
    samples = []
    prev_x = 0.0
    prev_y = 0.0
    for tc in codes:
        text = tc.text if tc.text is not None else ""
        if not text:
            continue
        x = attr_float(tc, "X", prev_x)
        y = attr_float(tc, "Y", prev_y)
        prev_x, prev_y = x, y
        deltas = _expand_deltas(parse_array(attr(tc, "DeltaX")), len(text),
                                text, size_mm * hscale)
        prepared.append((x, y, text, deltas))
        offset = 0.0
        for idx, ch in enumerate(text):
            samples.append((x + offset, y, deltas[idx]))
            offset += deltas[idx]

    if not prepared:
        return []

    est_w = 0.0
    if prepared:
        est_w = (prepared[-1][0] - prepared[0][0]) + sum(prepared[-1][3])
    has_text = any("".join(t).strip() for (_x, _y, t, _d) in prepared)
    padded = boundary is not None and boundary.w >= 1.8 * max(est_w, 0.1)
    mode = resolver.probe(samples, boundary, ctm, obj_id, size_mm,
                          quiet=(not has_text or padded))
    if ctm is not None and mode in ("A", "D"):
        scale = math.hypot(ctm[0], ctm[1])
        if abs(scale - 1.0) > 0.1:
            report.info("CTM_SCALE",
                        "文字对象的 CTM 含 %.3f 倍缩放，字号按原值输出，"
                        "实际观感可能偏大或偏小" % scale,
                        page=page_no, obj_id=obj_id)

    glyphs = []
    seq = seq_start
    for (x, y, text, deltas) in prepared:
        offset = 0.0
        for idx, ch in enumerate(text):
            px, py = resolver.transform(mode, x + offset, y, boundary, ctm)
            glyphs.append(Glyph(ch, px, py, deltas[idx], size_mm, color,
                                font.latin, font.cjk, bold, italic,
                                obj_id, seq))
            seq += 1
            offset += deltas[idx]
    return glyphs


def cluster_lines(glyphs, page_no):
    # type: (List[Glyph], int) -> List[List[Glyph]]
    """按基线把字形聚成视觉行。

    只按基线聚类，不限制横向间隙：公文版记这类"两端对齐"的行，中间空白很大，
    但视觉上仍是一行，拆开会让它在 Word 里变成两段。表格里跨单元格的同一行
    会在表格识别阶段按单元格边界再切开。
    """
    if not glyphs:
        return []
    ordered = sorted(glyphs, key=lambda g: (g.y, g.x, g.seq))
    clusters = []        # type: List[Dict[str, Any]]
    for g in ordered:
        target = None
        for cl in reversed(clusters[-6:]):
            tol = max(0.8, 0.42 * max(g.size_mm, cl["size"]))
            if abs(g.y - cl["y"]) <= tol:
                target = cl
                break
        if target is None:
            clusters.append({"y": g.y, "size": g.size_mm, "items": [g],
                             "right": g.x + g.adv, "left": g.x})
        else:
            target["items"].append(g)
            items = target["items"]
            target["y"] = sum(i.y for i in items) / len(items)
            target["size"] = max(target["size"], g.size_mm)
            target["right"] = max(target["right"], g.x + g.adv)
            target["left"] = min(target["left"], g.x)

    # 二次合并：基线接近
    merged = []          # type: List[Dict[str, Any]]
    for cl in sorted(clusters, key=lambda c: c["y"]):
        placed = False
        for m in merged:
            tol = max(0.8, 0.30 * max(cl["size"], m["size"]))
            if abs(cl["y"] - m["y"]) > tol:
                continue
            m["items"].extend(cl["items"])
            items = m["items"]
            m["y"] = sum(i.y for i in items) / len(items)
            m["size"] = max(m["size"], cl["size"])
            m["right"] = max(m["right"], cl["right"])
            m["left"] = min(m["left"], cl["left"])
            placed = True
            break
        if not placed:
            merged.append(cl)
    return [sorted(m["items"], key=lambda g: (g.x, g.seq)) for m in merged]


def finish_line(glyphs, page_no):
    # type: (List[Glyph], int) -> Optional[TextLine]
    """把一行字形拼成文字，处理跨对象空格与样式分段。"""
    if not glyphs:
        return None
    line = TextLine(glyphs, page_no)
    line.y = sum(g.y for g in glyphs) / len(glyphs)
    line.size_mm = max(g.size_mm for g in glyphs)

    pieces = []          # type: List[Piece]

    def cur_piece():
        if pieces:
            return pieces[-1]
        p = Piece("", glyphs[0].size_mm, glyphs[0].color,
                  glyphs[0].font_latin, glyphs[0].font_cjk,
                  glyphs[0].bold, glyphs[0].italic)
        pieces.append(p)
        return p

    prev = None
    for g in glyphs:
        if prev is not None:
            if abs(g.x - prev.x) < 0.15 * max(g.size_mm, 1e-6) and g.ch == prev.ch:
                continue        # 描边+填充重复绘制，去重
            if g.obj_id != prev.obj_id:
                gap = g.x - (prev.x + prev.adv)
                if is_wide(g.ch) or is_wide(prev.ch):
                    # 中文之间不存在词间空格，阈值放宽到 0.45em；
                    # 取较小字号，避免大字号把页码这类混排行的间隙吃掉
                    base = min(g.size_mm, prev.size_mm)
                    thr = 0.45 * base
                else:
                    base = max(g.size_mm, prev.size_mm)
                    thr = 0.25 * base
                if gap > 0.95 * max(g.size_mm, prev.size_mm):
                    cur_piece().text += "  "
                elif gap > thr:
                    cur_piece().text += " "
        p = cur_piece()
        if (round(g.size_mm, 2), g.color, g.font_latin, g.font_cjk,
                g.bold, g.italic) != p.style_key():
            p = Piece("", g.size_mm, g.color, g.font_latin, g.font_cjk,
                      g.bold, g.italic)
            pieces.append(p)
        p.text += g.ch
        prev = g

    # 去掉首尾空白
    while pieces:
        pieces[0].text = pieces[0].text.lstrip()
        if pieces[0].text:
            break
        pieces.pop(0)
    while pieces:
        pieces[-1].text = pieces[-1].text.rstrip()
        if pieces[-1].text:
            break
        pieces.pop()

    if not pieces:
        return None
    line.pieces = pieces
    line.left = glyphs[0].x
    line.right = glyphs[-1].x + glyphs[-1].adv
    # 行左右边界以首个/末个非空白字形为准
    for g in glyphs:
        if g.ch.strip():
            line.left = g.x
            break
    for g in reversed(glyphs):
        if g.ch.strip():
            line.right = g.x + g.adv
            break
    return line


def build_text_lines(text_objects, res, page_box, page_no, report):
    # type: (List[Any], ResourceTable, Box, int, Report) -> Tuple[List[TextLine], str]
    resolver = CoordResolver(page_box, report, page_no)
    glyphs = []          # type: List[Glyph]
    seq = 0
    for t_obj in text_objects:
        got = build_glyphs(t_obj, res, page_box, page_no, report, resolver, seq)
        if got:
            seq += len(got)
            glyphs.extend(got)
    clusters = cluster_lines(glyphs, page_no)
    lines = []
    for cl in clusters:
        ln = finish_line(cl, page_no)
        if ln is not None and not ln.is_blank():
            lines.append(ln)
    lines.sort(key=lambda l: (l.y, l.left))
    _assign_gaps(lines)
    mode = None
    if resolver.votes:
        mode = max(resolver.votes.items(), key=lambda kv: kv[1])[0]
    return lines, mode


def _assign_gaps(lines):
    # type: (List[TextLine]) -> None
    """用相邻行的基线差推算固定行距；离群值回退到 1.32 倍字号。"""
    for i, ln in enumerate(lines):
        if i + 1 < len(lines):
            gap = lines[i + 1].y - ln.y
            if gap <= 0 or gap > 3.0 * ln.size_mm:
                ln.gap_to_next = 1.32 * ln.size_mm
            else:
                ln.gap_to_next = gap
        else:
            ln.gap_to_next = 1.32 * ln.size_mm


# ---------------------------------------------------------------------------
# B9  表格识别
# ---------------------------------------------------------------------------

# OFD 路径命令的操作数个数。注意与 PDF 不同：
#   B = 三次贝塞尔（6 个参数），C = 闭合路径（0 个参数）
ARITY = {"S": 2, "M": 2, "L": 2, "Q": 4, "B": 6, "A": 7, "C": 0, "Z": 0}
_CMD_RE = re.compile(r"^([SMLQBACZ])(.*)$", re.IGNORECASE)


def parse_path_data(text):
    # type: (Optional[str]) -> List[Tuple[str, List[float]]]
    """解析 AbbreviatedData，返回 [(命令, 参数列表)]。"""
    if not text or not text.strip():
        return []
    toks = [t for t in re.split(r"[,\s]+", text.strip()) if t]
    out = []           # type: List[Tuple[str, List[float]]]
    i = 0
    n = len(toks)
    while i < n:
        m = _CMD_RE.match(toks[i])
        if not m:
            i += 1
            continue
        cmd = m.group(1).upper()
        rest = m.group(2)
        nums = []      # type: List[float]
        i += 1
        if rest:
            try:
                nums.append(float(rest))
            except ValueError:
                pass
        need = ARITY.get(cmd, 0)
        while len(nums) < need and i < n and not _CMD_RE.match(toks[i]):
            try:
                nums.append(float(toks[i]))
            except ValueError:
                break
            i += 1
        out.append((cmd, nums[:need] if need else []))
    return out


def _quad_point(p0, c1, p1, t):
    mt = 1.0 - t
    return (mt * mt * p0[0] + 2 * mt * t * c1[0] + t * t * p1[0],
            mt * mt * p0[1] + 2 * mt * t * c1[1] + t * t * p1[1])


def _cubic_point(p0, c1, c2, p1, t):
    mt = 1.0 - t
    a = mt * mt * mt
    b = 3 * mt * mt * t
    c = 3 * mt * t * t
    d = t * t * t
    return (a * p0[0] + b * c1[0] + c * c2[0] + d * p1[0],
            a * p0[1] + b * c1[1] + c * c2[1] + d * p1[1])


def path_subpaths(cmds):
    # type: (List[Tuple[str, List[float]]]) -> List[Tuple[List[Tuple[float, float]], bool]]
    """把命令序列折线化。返回 [(点列表, 是否含曲线)]。"""
    subs = []          # type: List[Tuple[List[Tuple[float, float]], bool]]
    cur = None         # type: Optional[List[Tuple[float, float]]]
    curved = False
    x = y = 0.0
    start = (0.0, 0.0)
    for cmd, nums in cmds:
        if cmd in ("S", "M"):
            if len(nums) < 2:
                continue
            if cur is not None and len(cur) >= 2:
                subs.append((cur, curved))
            x, y = nums[0], nums[1]
            start = (x, y)
            cur = [(x, y)]
            curved = False
        elif cmd == "L":
            if len(nums) < 2:
                continue
            x, y = nums[0], nums[1]
            if cur is None:
                # "S" 之后可能跟着多个 "C"（闭合）分段：闭合后的第一个 L 才是
                # 新子路径的起点。若继续沿用最初 S 的坐标，每段快结束时都会被
                # 强拉回同一个点，字形上就出现一根根斜线（红头字体尤其明显）。
                cur = [(x, y)]
                start = (x, y)
            else:
                cur.append((x, y))
        elif cmd in ("C", "Z"):
            if cur is not None:
                if cur[-1] != start:
                    cur.append(start)
                subs.append((cur, curved))
                cur = None
                curved = False
        elif cmd == "Q":
            if len(nums) < 4:
                continue
            p0 = (x, y)
            c1 = (nums[0], nums[1])
            p1 = (nums[2], nums[3])
            if cur is None:
                cur = [p0]
                start = p0
            for t in (0.25, 0.5, 0.75, 1.0):
                cur.append(_quad_point(p0, c1, p1, t))
            x, y = p1
            curved = True
        elif cmd == "B":
            if len(nums) < 6:
                continue
            p0 = (x, y)
            c1 = (nums[0], nums[1])
            c2 = (nums[2], nums[3])
            p1 = (nums[4], nums[5])
            if cur is None:
                cur = [p0]
                start = p0
            for t in (0.2, 0.4, 0.6, 0.8, 1.0):
                cur.append(_cubic_point(p0, c1, c2, p1, t))
            x, y = p1
            curved = True
        elif cmd == "A":
            if len(nums) < 7:
                continue
            p1 = (nums[5], nums[6])
            if cur is None:
                cur = [(x, y)]
                start = (x, y)
            cur.append(p1)
            x, y = p1
            curved = True
    if cur is not None and len(cur) >= 2:
        subs.append((cur, curved))
    return subs


def _bbox_match_score(gb, b):
    # type: (Optional[Box], Optional[Box]) -> float
    """比较两个矩形的四条边，用于在线段这类"零高度"包围盒上判定坐标解释。"""
    if gb is None or b is None:
        return 0.0
    dist = (abs(gb.x - b.x) + abs(gb.y - b.y) +
            abs(gb.x2 - b.x2) + abs(gb.y2 - b.y2))
    scale = max(b.w + b.h, 1.0)
    return max(0.0, 1.0 - dist / (2.0 * scale))


def _points_mode(pts, boundary, ctm, page_box):
    # type: (List[Tuple[float, float]], Optional[Box], Optional[Sequence[float]], Optional[Box]) -> List[Tuple[float, float]]
    """为路径点选择坐标解释。

    与文字同一思路：候选解释里，取"包围盒与 Boundary 最吻合且落在页内"的那一种。
    由于直线段的包围盒高度为 0，不能用 IoU，改为比较四条边的距离。
    """
    if not pts:
        return []
    cands = [("C", list(pts))]      # 绝对页面坐标
    if boundary is not None:
        cands.append(("B", [(boundary.x + p[0], boundary.y + p[1]) for p in pts]))
    if ctm is not None and not matrix_is_identity(ctm):
        cands.append(("A", [apply_matrix(ctm, p[0], p[1]) for p in pts]))
        if boundary is not None:
            lin = (ctm[0], ctm[1], ctm[2], ctm[3], 0.0, 0.0)
            shifted = []
            for p in pts:
                px, py = apply_matrix(lin, p[0], p[1])
                shifted.append((boundary.x + px, boundary.y + py))
            cands.append(("D", shifted))
    if boundary is None or page_box is None:
        return cands[0][1]
    best = None
    best_score = -1.0
    for _mode, cp in cands:
        inside = 0
        for (x, y) in cp:
            if page_box.contains_point(x, y, TOL_ON_PAGE):
                inside += 1
        if inside < len(cp) * 0.85:
            continue
        score = _bbox_match_score(bbox_of(cp), boundary)
        if score > best_score:
            best_score = score
            best = cp
    return best if best is not None else cands[0][1]


def path_data_of(el):
    # type: (Any) -> str
    """取图形轮廓数据。

    规范把 AbbreviatedData 定义为 PathObject 的属性，但实测有生成器（本样本就是）
    写成子元素 <ofd:AbbreviatedData>…</ofd:AbbreviatedData>。两种都要认，
    否则整条路径会被当成空占位符丢掉——红头横线、版记横线就是这么丢的。
    """
    raw = attr(el, "AbbreviatedData")
    if raw and raw.strip():
        return raw
    node = child(el, "AbbreviatedData")
    if node is not None:
        return text_of(node)
    return ""


def _thin_rect_rule(points, fill_width_hint=0.0):
    # type: (List[Tuple[float, float]], float) -> Optional[Seg]
    """把"细长的填充矩形"识别成一条表格线。

    不少生成器画边框/横线时不用描边，而是填一个 0.5~1.5mm 厚的窄矩形。
    不这样处理就完全看不到表格线。
    """
    if len(points) < 4:
        return None
    xs = sorted(set(round(p[0], 3) for p in points))
    ys = sorted(set(round(p[1], 3) for p in points))
    if len(xs) != 2 or len(ys) != 2:
        return None
    w = xs[1] - xs[0]
    h = ys[1] - ys[0]
    if h <= 1.8 and w >= 3.0:
        return Seg("h", (ys[0] + ys[1]) / 2.0, xs[0], xs[1], max(h, 0.1))
    if w <= 1.8 and h >= 3.0:
        return Seg("v", (xs[0] + xs[1]) / 2.0, ys[0], ys[1], max(w, 0.1))
    return None


def extract_segments(page_root, res, page_box, page_no, report):
    # type: (Any, ResourceTable, Box, int, Report) -> List[Seg]
    segs = []          # type: List[Seg]
    for po in iter_local(page_root, "PathObject"):
        data = path_data_of(po)
        if not data or not data.strip():
            continue
        if parse_array(attr(po, "DashPattern")):
            continue
        dp = res.draw_param(attr(po, "DrawParam"))
        width = res.line_width_of(po, dp)
        stroke_attr = attr(po, "Stroke")
        do_stroke = True if stroke_attr is None else attr_true(po, "Stroke")
        do_fill = attr_true(po, "Fill")
        if not do_stroke and not do_fill:
            continue
        if width <= 0.02:
            width = 0.35
        boundary = parse_box(attr(po, "Boundary"))
        ctm = parse_array(attr(po, "CTM"))
        if len(ctm) < 6:
            ctm = None
        for pts, curved in path_subpaths(parse_path_data(data)):
            if curved:
                continue          # 含曲线的子路径整体排除（圆角框、印章轮廓）
            mapped = _points_mode(pts, boundary, ctm, page_box)
            if not do_stroke:
                # 纯填充：只可能是"厚线"，按细长矩形取中线
                rule = _thin_rect_rule(mapped)
                if rule is not None:
                    rule.color = res.object_color(po, "FillColor", dp)
                    rule.from_fill = True
                    segs.append(rule)
                continue
            if width > 3.0:
                # 粗线通常是装饰色块边框，不当表格线
                continue
            for i in range(len(mapped) - 1):
                x1, y1 = mapped[i]
                x2, y2 = mapped[i + 1]
                dx = x2 - x1
                dy = y2 - y1
                length = math.hypot(dx, dy)
                if length < TOL_MIN_SEG:
                    continue
                ang = math.degrees(math.atan2(dy, dx)) % 180.0
                if ang <= TOL_ANGLE or ang >= 180.0 - TOL_ANGLE:
                    segs.append(Seg("h", (y1 + y2) / 2.0,
                                    min(x1, x2), max(x1, x2), width))
                elif abs(ang - 90.0) <= TOL_ANGLE:
                    segs.append(Seg("v", (x1 + x2) / 2.0,
                                    min(y1, y2), max(y1, y2), width))
    return segs


def extract_vectors(page_root, res, page_box, page_no, report):
    # type: (Any, ResourceTable, Box, int, Report) -> List[Dict[str, Any]]
    """抽取页面上的矢量图形，供阅读器预览使用。

    与表格识别用的 extract_segments 不同：这里不过滤曲线、粗线，也不判方向，
    只把路径折线化后原样交给渲染层，尽量还原图表、边框、底纹等图形。
    仍然跳过 AbbreviatedData 为空的占位路径。
    """
    out = []           # type: List[Dict[str, Any]]
    for po in iter_local(page_root, "PathObject"):
        raw = path_data_of(po)
        if not raw or not raw.strip():
            continue
        boundary = parse_box(attr(po, "Boundary"))
        ctm = parse_array(attr(po, "CTM"))
        if len(ctm) < 6:
            ctm = None
        dp = res.draw_param(attr(po, "DrawParam"))
        stroke_attr = attr(po, "Stroke")
        do_stroke = True if stroke_attr is None else attr_true(po, "Stroke")
        do_fill = attr_true(po, "Fill")
        if not do_stroke and not do_fill:
            continue
        width = res.line_width_of(po, dp)
        if width <= 0.02:
            width = 0.35
        stroke_color = res.object_color(po, "StrokeColor", dp)
        fill_color = res.object_color(po, "FillColor", dp)
        dashed = bool(parse_array(attr(po, "DashPattern")))
        # Rule 决定同一 PathObject 内多个子路径如何合成：Even-Odd 奇偶填充、
        # 其余按 Non-Zero。默认 Non-Zero（GB/T 33190）。
        rule = attr(po, "Rule") or "Non-Zero"
        group = "po:%s:%d" % (attr(po, "ID") or "?", len(out))
        sub_no = 0
        for pts, curved in path_subpaths(parse_path_data(raw)):
            if len(pts) < 2:
                continue
            mapped = _points_mode(pts, boundary, ctm, page_box)
            closed = bool(curved) or (
                abs(mapped[0][0] - mapped[-1][0]) < 1e-6 and
                abs(mapped[0][1] - mapped[-1][1]) < 1e-6)
            out.append({
                "points": mapped,
                "closed": closed,
                "curved": bool(curved),
                # 填充不要求显式闭合：按 PDF/OFD 语义，填充时子路径隐式闭合。
                # 之前用 closed 卡填充，会把"没写 C 命令"的字形轮廓整条丢掉，
                # 红头这类"文字转矢量"的图形因此只剩几条边，看起来就是字被切碎。
                "fill": do_fill,
                "stroke": do_stroke,
                "rule": str(rule),
                "group": group,
                "sub": sub_no,
                "stroke_color": stroke_color,
                "fill_color": fill_color,
                "width": max(0.05, min(width, 5.0)),
                "dash": dashed,
            })
            sub_no += 1
    return out


def _snap(values):
    # type: (List[Tuple[float, float]]) -> List[float]
    """一维加权吸附聚类，返回升序的代表坐标。"""
    if not values:
        return []
    vals = sorted(values)
    groups = [[vals[0]]]
    for item in vals[1:]:
        last = groups[-1][-1]
        tol = max(TOL_SNAP_BASE, 2.2 * max(item[1], last[1]))
        if item[0] - last[0] <= tol:
            groups[-1].append(item)
        else:
            groups.append([item])
    out = []
    for grp in groups:
        tot = sum(max(w, 0.05) for _, w in grp)
        out.append(sum(v * max(w, 0.05) for v, w in grp) / tot)
    return out


def _coverage_ratio(a0, a1, b0, b1):
    # type: (float, float, float, float) -> float
    if b1 <= b0:
        return 0.0
    ov = min(a1, b1) - max(a0, b0)
    if ov <= 0:
        return 0.0
    return ov / (b1 - b0)


class _UnionFind(object):
    def __init__(self, n):
        self.parent = list(range(n))

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def detect_tables(segs, lines, page_box, page_no, report, enabled=True):
    # type: (List[Seg], List[TextLine], Box, int, Report, bool) -> List[Table]
    if not enabled or not segs:
        return []
    hs = [s for s in segs if s.orient == "h"]
    vs = [s for s in segs if s.orient == "v"]
    if len(hs) < 3 or len(vs) < 3:
        return []

    ys = _snap([(s.c, s.width) for s in hs])
    xs = _snap([(s.c, s.width) for s in vs])
    if len(ys) < 3 or len(xs) < 3:
        return []
    m = len(ys) - 1
    n = len(xs) - 1

    # 每条线覆盖了哪些格子边
    h_at = [[False] * n for _ in range(m + 1)]
    for s in hs:
        i = min(range(len(ys)), key=lambda k: abs(ys[k] - s.c))
        for j in range(n):
            if _coverage_ratio(s.a, s.b, xs[j], xs[j + 1]) >= 0.6:
                h_at[i][j] = True
    v_at = [[False] * m for _ in range(n + 1)]
    for s in vs:
        j = min(range(len(xs)), key=lambda k: abs(xs[k] - s.c))
        for i in range(m):
            if _coverage_ratio(s.a, s.b, ys[i], ys[i + 1]) >= 0.6:
                v_at[j][i] = True

    # 只保留真正被线条穿过的行列，去掉零散噪声坐标
    row_used = [False] * (m + 1)
    col_used = [False] * (n + 1)
    for i in range(m + 1):
        if any(h_at[i]):
            row_used[i] = True
    for j in range(n + 1):
        if any(v_at[j]):
            col_used[j] = True
    if not all(row_used) or not all(col_used):
        return []

    # 并查集：没有分隔线的相邻格合并
    uf = _UnionFind(m * n)

    def idx(i, j):
        return i * n + j

    for i in range(m):
        for j in range(n - 1):
            if not v_at[j + 1][i]:
                uf.union(idx(i, j), idx(i, j + 1))
    for i in range(m - 1):
        for j in range(n):
            if not h_at[i + 1][j]:
                uf.union(idx(i, j), idx(i + 1, j))

    groups = {}        # type: Dict[int, List[Tuple[int, int]]]
    for i in range(m):
        for j in range(n):
            groups.setdefault(uf.find(idx(i, j)), []).append((i, j))

    rects = []         # type: List[Tuple[int, int, int, int]]
    for cells in groups.values():
        i0 = min(c[0] for c in cells)
        i1 = max(c[0] for c in cells)
        j0 = min(c[1] for c in cells)
        j1 = max(c[1] for c in cells)
        if (i1 - i0 + 1) * (j1 - j0 + 1) != len(cells):
            # 非矩形（斜线表头/缺线），按行切成若干直矩形
            rows = {}    # type: Dict[int, List[int]]
            for (i, j) in cells:
                rows.setdefault(i, []).append(j)
            for i in sorted(rows):
                cols = sorted(rows[i])
                start = cols[0]
                prev = cols[0]
                for c in cols[1:] + [None]:
                    if c is not None and c == prev + 1:
                        prev = c
                        continue
                    rects.append((i, i, start, prev))
                    if c is not None:
                        start = prev = c
            continue
        rects.append((i0, i1, j0, j1))

    tbl = Table(xs, ys, [])
    box = tbl.box
    if box.area > 0.98 * page_box.area:
        return []

    # 置信度
    total_edges = (m + 1) * n + (n + 1) * m
    present = sum(1 for row in h_at for v in row if v) + \
        sum(1 for row in v_at for v in row if v)
    f2 = present / float(total_edges) if total_edges else 0.0

    inside = [ln for ln in lines
              if box.x <= ln.left <= box.x2 and box.y <= ln.y <= box.y2]
    f3 = 1.0 if not inside else min(1.0, len(inside) / float(len(inside)))
    if inside:
        in_cell = 0
        for ln in inside:
            for (i0, i1, j0, j1) in rects:
                if xs[j0] - 0.6 <= ln.left <= xs[j1 + 1] + 0.6 and \
                        ys[i0] - 0.6 <= ln.y <= ys[i1 + 1] + 0.6:
                    in_cell += 1
                    break
        f3 = in_cell / float(len(inside))
    else:
        f3 = 0.0

    heights = [ys[i + 1] - ys[i] for i in range(m)]
    fonts = [ln.size_mm for ln in inside] or [4.0]
    row_per_font = statistics.median(heights) / max(statistics.median(fonts), 0.5)
    f4 = min(row_per_font / 2.0, 1.0) if row_per_font >= 1.15 else 0.0

    score = 0.5 * f2 + 0.35 * f3 + 0.15 * f4
    if score < 0.6:
        report.info("TABLE_REJECTED",
                    "检测到候选网格但置信度仅 %.2f（完整度 %.2f、落格率 %.2f），"
                    "已按普通文本输出" % (score, f2, f3),
                    page=page_no, obj_id="grid@%.1f,%.1f" % (box.x, box.y))
        return []

    # 生成单元格
    cells = []         # type: List[Cell]
    for (i0, i1, j0, j1) in rects:
        cell = Cell(i0, j0, i1 - i0 + 1, j1 - j0 + 1)
        if all(h_at[i0][j] for j in range(j0, j1 + 1)):
            cell.borders["top"] = _seg_width(hs, ys[i0])
        if all(h_at[i1 + 1][j] for j in range(j0, j1 + 1)):
            cell.borders["bottom"] = _seg_width(hs, ys[i1 + 1])
        if all(v_at[j0][i] for i in range(i0, i1 + 1)):
            cell.borders["left"] = _seg_width(vs, xs[j0])
        if all(v_at[j1 + 1][i] for i in range(i0, i1 + 1)):
            cell.borders["right"] = _seg_width(vs, xs[j1 + 1])
        cells.append(cell)

    # 文字归位：一行可能横跨多个单元格，按单元格边界切开
    for ln in inside:
        hits = []          # type: List[Cell]
        for cell in cells:
            x0 = xs[cell.c]
            x1 = xs[cell.c + cell.cs]
            y0 = ys[cell.r]
            y1 = ys[cell.r + cell.rs]
            if not (y0 - 0.8 <= ln.y <= y1 + 0.8):
                continue
            if ln.left <= x1 + 0.8 and ln.right >= x0 - 0.8:
                hits.append(cell)
        if not hits:
            continue
        if len(hits) == 1:
            hits[0].lines.append(ln)
            continue
        for cell in hits:
            x0 = xs[cell.c]
            x1 = xs[cell.c + cell.cs]
            sub = [g for g in ln.glyphs if x0 - 0.9 <= g.x <= x1 + 0.9]
            if not sub:
                continue
            sub_line = finish_line(sorted(sub, key=lambda g: (g.x, g.seq)), page_no)
            if sub_line is not None and not sub_line.is_blank():
                cell.lines.append(sub_line)
    for cell in cells:
        cell.lines.sort(key=lambda l: (l.y, l.left))

    tbl.cells = cells
    tbl.score = score
    report.info("TABLE_FOUND",
                "识别到 %d 行 × %d 列表格（置信度 %.2f，%d 个单元格）"
                % (m, n, score, len(cells)), page=page_no)
    return [tbl]


def _seg_width(segs, coord):
    # type: (List[Seg], float) -> float
    best = 0.35
    for s in segs:
        if abs(s.c - coord) <= max(TOL_SNAP_BASE, 1.5 * s.width):
            return s.width
    return best


# ---------------------------------------------------------------------------
# B10  图片
# ---------------------------------------------------------------------------

IMAGE_EXTS = {
    "PNG": "png", "JPEG": "jpg", "JPG": "jpg", "GIF": "gif",
    "BMP": "bmp", "TIFF": "tiff", "TIF": "tiff",
}
CONVERTIBLE = {"JPX", "JPEG2000", "J2K", "JBIG2", "JBG", "CCITT"}


def sniff_image(data, hint=None):
    # type: (bytes, Optional[str]) -> str
    if hint:
        hint = hint.upper()
        if hint in IMAGE_EXTS:
            return IMAGE_EXTS[hint]
        if hint in CONVERTIBLE:
            return hint.lower()
    head = data[:12]
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if head.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if head.startswith(b"GIF8"):
        return "gif"
    if head.startswith(b"BM"):
        return "bmp"
    if head.startswith(b"II*\x00") or head.startswith(b"MM\x00*"):
        return "tiff"
    if head.startswith(b"\x00\x00\x00\x0c") and data[4:8] == b"jP  ":
        return "jpx"
    return ""


def image_to_png(data):
    # type: (bytes) -> Optional[bytes]
    try:
        from PIL import Image
    except ImportError:
        return None
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
        if im.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            im = im.convert("RGB")
        out = io.BytesIO()
        im.save(out, "PNG")
        return out.getvalue()
    except Exception:
        return None


def load_images(page_root, pkg, res, page_box, page_no, report):
    # type: (Any, OfdPackage, ResourceTable, Box, int, Report) -> Tuple[List[ImageItem], int]
    items = []         # type: List[ImageItem]
    skipped = 0
    for obj in iter_local(page_root, "ImageObject"):
        rid = attr(obj, "ResourceID")
        media = res.media_info(rid)
        if media is None:
            report.warn("IMAGE_NO_MEDIA",
                        "图片对象引用的资源不存在（ResourceID=%s），已跳过" % rid,
                        page=page_no, obj_id=attr(obj, "ID"))
            skipped += 1
            continue
        loc = media.get("loc")
        if not loc:
            report.warn("IMAGE_NO_FILE",
                        "图片资源没有 MediaFile 指向实体文件，已跳过",
                        page=page_no, obj_id=attr(obj, "ID"))
            skipped += 1
            continue
        try:
            part = pkg.resolve_in(media.get("base_dir") or "", loc)
            data = pkg.read(part)
        except OfdError as exc:
            report.warn("IMAGE_READ_FAIL",
                        "读取图片失败（%s），已跳过" % exc.message,
                        page=page_no, obj_id=attr(obj, "ID"))
            skipped += 1
            continue
        ext = sniff_image(data, media.get("format"))
        if ext in CONVERTIBLE or ext not in ("png", "jpg", "gif", "bmp", "tiff"):
            converted = image_to_png(data)
            if converted is None:
                report.warn("IMAGE_FORMAT",
                            "图片格式为 %s，无法解码（缺少 Pillow 或不支持的编码），"
                            "已跳过" % (ext or "未知"),
                            page=page_no, obj_id=attr(obj, "ID"))
                skipped += 1
                continue
            data = converted
            ext = "png"
        box = parse_box(attr(obj, "Boundary"))
        if box is None:
            report.warn("IMAGE_NO_BOX", "图片对象没有有效边界，已跳过",
                        page=page_no, obj_id=attr(obj, "ID"))
            skipped += 1
            continue
        import hashlib
        digest = hashlib.sha1(data).hexdigest()
        # Alpha 是 0–255 的整体不透明度，电子印章常见 160 左右；
        # 以前没读，印章会以全不透明盖上，把日期压得看不清。
        alpha = attr_float(obj, "Alpha", 255.0)
        alpha = max(0.0, min(255.0, alpha))
        items.append(ImageItem(data, ext, box, page_no, rid, digest, alpha))
    return items, skipped


# ---------------------------------------------------------------------------
# B11  DOCX 组装
# ---------------------------------------------------------------------------

_DOCX = {}


def _dx(name):
    # type: (str) -> Any
    if not _DOCX:
        try:
            from docx import Document
            from docx.enum.section import WD_SECTION
            from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
            from docx.oxml import parse_xml as oxml_parse_xml
            from docx.oxml.ns import nsdecls, qn
            from docx.shared import Emu, Pt, RGBColor, Twips
        except ImportError:
            raise OfdUnsupportedError(
                "缺少 python-docx，无法生成 Word 文档",
                hint="请执行：pip install \"python-docx==1.1.2\"")
        _DOCX.update({
            "Document": Document, "WD_SECTION": WD_SECTION,
            "WD_ROW_HEIGHT_RULE": WD_ROW_HEIGHT_RULE,
            "WD_TABLE_ALIGNMENT": WD_TABLE_ALIGNMENT,
            "oxml_parse_xml": oxml_parse_xml, "nsdecls": nsdecls, "qn": qn,
            "Emu": Emu, "Pt": Pt, "RGBColor": RGBColor, "Twips": Twips,
        })
    return _DOCX[name]


class DocxBuilder(object):
    def __init__(self, report):
        # type: (Report) -> None
        self.report = report
        self.doc = _dx("Document")()
        self._first_section = True
        self._last_block = "paragraph"
        self._setup_styles()

    def _setup_styles(self):
        # type: () -> None
        Pt = _dx("Pt")
        qn = _dx("qn")
        style = self.doc.styles["Normal"]
        try:
            style.font.name = LATIN_FALLBACK
            style.font.size = Pt(10.5)
        except Exception:
            pass
        try:
            rpr = style.element.get_or_add_rPr()
            rpr.get_or_add_rFonts().set(qn("w:eastAsia"), CJK_FALLBACK)
        except Exception:
            pass

    # -- 分节 -------------------------------------------------------------
    def add_section(self, box, margins):
        # type: (Box, Tuple[float, float, float, float]) -> Any
        WD_SECTION = _dx("WD_SECTION")
        if self._first_section:
            sec = self.doc.sections[0]
            self._first_section = False
        else:
            sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        Emu = _dx("Emu")
        qn = _dx("qn")
        sec.page_width = Emu(mm_to_emu(box.w))
        sec.page_height = Emu(mm_to_emu(box.h))
        left, top, right, bottom = margins
        sec.left_margin = Emu(mm_to_emu(left))
        sec.right_margin = Emu(mm_to_emu(right))
        sec.top_margin = Emu(mm_to_emu(top))
        sec.bottom_margin = Emu(mm_to_emu(bottom))
        try:
            sec.header_distance = Emu(mm_to_emu(5.0))
            sec.footer_distance = Emu(mm_to_emu(5.0))
        except Exception:
            pass
        try:
            sectPr = sec._sectPr
            pgSz = sectPr.find(qn("w:pgSz"))
            if pgSz is not None:
                pgSz.set(qn("w:orient"),
                         "landscape" if box.w > box.h else "portrait")
        except Exception:
            pass
        self._last_block = "section"
        return sec

    # -- 段落 -------------------------------------------------------------
    def _apply_paragraph_format(self, p, line, left_margin, avail_w):
        # type: (Any, TextLine, float, float) -> bool
        Twips = _dx("Twips")
        pf = p.paragraph_format
        pf.space_before = Twips(0)
        pf.space_after = Twips(0)
        try:
            pf.widow_control = False
        except Exception:
            pass
        size_pt = mm_to_pt(line.size_mm)
        gap = line.gap_to_next if line.gap_to_next else 1.32 * line.size_mm
        line_tw = max(mm_to_twips(gap), int(round(size_pt * 20.0 * 1.12)))
        pf.line_spacing = Twips(line_tw)
        indent_mm = max(0.0, line.left - left_margin)
        est_w = max(0.0, line.right - line.left)
        overflow = False
        if avail_w > 0 and indent_mm + est_w > avail_w:
            if avail_w >= est_w:
                indent_mm = max(0.0, avail_w - est_w)
            else:
                indent_mm = 0.0
                overflow = True
        pf.left_indent = Twips(mm_to_twips(indent_mm))
        pf.first_line_indent = Twips(0)
        return overflow

    def _add_runs(self, p, line):
        # type: (Any, TextLine) -> None
        Pt = _dx("Pt")
        RGBColor = _dx("RGBColor")
        qn = _dx("qn")
        for piece in line.pieces:
            if not piece.text:
                continue
            r = p.add_run(piece.text)
            r.font.size = Pt(round(mm_to_pt(piece.size_mm), 1))
            r.font.bold = piece.bold
            r.font.italic = piece.italic
            if piece.color != (0, 0, 0):
                r.font.color.rgb = RGBColor(*piece.color)
            r.font.name = piece.font_latin
            rpr = r._element.get_or_add_rPr()
            # 中文必须设置 eastAsia，否则中文字体会整体回退
            rpr.get_or_add_rFonts().set(qn("w:eastAsia"), piece.font_cjk)

    def add_line(self, line, left_margin, avail_w, page_no):
        # type: (TextLine, float, float, int) -> Any
        p = self.doc.add_paragraph()
        overflow = self._apply_paragraph_format(p, line, left_margin, avail_w)
        if overflow:
            self.report.warn(
                "LINE_TOO_WIDE",
                "有一行文字宽度超过正文区域，已紧贴左边距输出，Word 可能重新折行",
                page=page_no)
        self._add_runs(p, line)
        self._last_block = "paragraph"
        return p

    def add_empty_paragraph(self):
        # type: () -> Any
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = _dx("Twips")(0)
        p.paragraph_format.space_after = _dx("Twips")(0)
        self._last_block = "paragraph"
        return p

    def add_spacer(self, height_mm):
        # type: (float) -> Any
        """插入一个高度精确的空白段落，用来还原纵向空隙。"""
        Twips = _dx("Twips")
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Twips(0)
        pf.space_after = Twips(0)
        pf.line_spacing = Twips(max(20, mm_to_twips(max(0.0, height_mm))))
        self._last_block = "paragraph"
        return p

    # -- 表格 -------------------------------------------------------------
    def add_table(self, table, left_margin, avail_w, page_no):
        # type: (Table, float, float, int) -> Any
        Twips = _dx("Twips")
        Emu = _dx("Emu")
        WD_TABLE_ALIGNMENT = _dx("WD_TABLE_ALIGNMENT")
        WD_ROW_HEIGHT_RULE = _dx("WD_ROW_HEIGHT_RULE")
        qn = _dx("qn")

        nrows, ncols = table.nrows, table.ncols
        t = self.doc.add_table(rows=nrows, cols=ncols)
        t.autofit = False
        try:
            t.alignment = WD_TABLE_ALIGNMENT.LEFT
        except Exception:
            pass

        widths = [table.col_xs[j + 1] - table.col_xs[j] for j in range(ncols)]
        total = sum(widths) or 1.0
        scale = 1.0
        if avail_w > 0 and total > avail_w:
            scale = avail_w / total
        tw = [max(120, mm_to_twips(w * scale)) for w in widths]
        for j, w in enumerate(tw):
            t.columns[j].width = Twips(w)
        for row in t.rows:
            for j, cell in enumerate(row.cells):
                cell.width = Twips(tw[j])

        # 表级：无边框，逐格再画识别到的线
        self._set_table_borders(t)
        # 缩进：用表格整体左缩进对齐 OFD 里的 x
        try:
            indent = max(0.0, table.box.x - left_margin)
            tblPr = t._tbl.tblPr
            ind = tblPr.find(qn("w:tblInd"))
            if ind is None:
                ind = _dx("oxml_parse_xml")(
                    '<w:tblInd %s w:w="%d" w:type="dxa"/>'
                    % (_dx("nsdecls")("w"), mm_to_twips(indent)))
                tblPr.append(ind)
            else:
                ind.set(qn("w:w"), str(mm_to_twips(indent)))
        except Exception:
            pass

        # 合并单元格（先合并，再填内容与边框）
        merged = {}
        for cell in table.cells:
            if cell.rs == 1 and cell.cs == 1:
                continue
            try:
                origin = t.cell(cell.r, cell.c)
                other = t.cell(cell.r + cell.rs - 1, cell.c + cell.cs - 1)
                merged[(cell.r, cell.c)] = origin.merge(other)
            except Exception as exc:
                self.report.warn("TABLE_MERGE_FAIL",
                                 "合并单元格失败，已按未合并输出：%s" % exc,
                                 page=page_no)

        for cell in table.cells:
            tc = merged.get((cell.r, cell.c))
            if tc is None:
                try:
                    tc = t.cell(cell.r, cell.c)
                except Exception:
                    continue
            self._set_cell_borders(tc, cell.borders)
            first = True
            for ln in cell.lines:
                if first:
                    p = tc.paragraphs[0]
                    first = False
                else:
                    p = tc.add_paragraph()
                self._apply_paragraph_format(p, ln, table.box.x, 0.0)
                try:
                    p.paragraph_format.left_indent = Twips(0)
                    p.paragraph_format.first_line_indent = Twips(0)
                except Exception:
                    pass
                self._add_runs(p, ln)

        for i in range(nrows):
            h_mm = table.row_ys[i + 1] - table.row_ys[i]
            try:
                t.rows[i].height = Emu(mm_to_emu(h_mm))
                t.rows[i].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            except Exception:
                pass

        self._last_block = "table"
        return t

    def _set_table_borders(self, t):
        # type: (Any) -> None
        qn = _dx("qn")
        try:
            tblPr = t._tbl.tblPr
            if tblPr.find(qn("w:tblBorders")) is not None:
                return
            xml = ['<w:tblBorders %s>' % _dx("nsdecls")("w")]
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                xml.append('<w:%s w:val="nil"/>' % side)
            xml.append("</w:tblBorders>")
            tblPr.append(_dx("oxml_parse_xml")("".join(xml)))
        except Exception:
            pass

    def _set_cell_borders(self, cell, borders):
        # type: (Any, Dict[str, Optional[float]]) -> None
        qn = _dx("qn")
        try:
            tcPr = cell._tc.get_or_add_tcPr()
            old = tcPr.find(qn("w:tcBorders"))
            if old is not None:
                tcPr.remove(old)
            xml = ['<w:tcBorders %s>' % _dx("nsdecls")("w")]
            for side in ("top", "left", "bottom", "right"):
                width = borders.get(side)
                if width:
                    xml.append('<w:%s w:val="single" w:sz="%d" '
                               'w:space="0" w:color="auto"/>'
                               % (side, mm_to_eighth_pt(width)))
                else:
                    xml.append('<w:%s w:val="nil"/>' % side)
            xml.append("</w:tcBorders>")
            tcPr.append(_dx("oxml_parse_xml")("".join(xml)))
        except Exception:
            pass

    # -- 图片 -------------------------------------------------------------
    def add_image(self, item, left_margin, avail_w, page_no):
        # type: (ImageItem, float, float, int) -> Any
        Emu = _dx("Emu")
        Twips = _dx("Twips")
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Twips(0)
        pf.space_after = Twips(0)
        pf.left_indent = Twips(mm_to_twips(max(0.0, item.box.x - left_margin)))
        width_mm = item.box.w
        if avail_w > 0 and width_mm > avail_w:
            width_mm = avail_w
        try:
            run = p.add_run()
            run.add_picture(io.BytesIO(item.data), width=Emu(mm_to_emu(width_mm)))
        except Exception as exc:
            self.report.warn("IMAGE_INSERT_FAIL",
                             "图片插入 Word 失败，已跳过：%s" % exc,
                             page=page_no)
        self._last_block = "paragraph"
        return p

    # -- 段落边框（贴近文字的横线）-----------------------------------------
    def add_paragraph_bottom_border(self, paragraph, color, width_mm,
                                    right_indent_mm):
        # type: (Any, Tuple[int, int, int], float, float) -> None
        """给段落加一条底边框，用来还原贴文字下方的横线。

        线随段落走，Word 怎么排都不会压到字；右侧缩进按原件的线长收一下，
        避免边框一直画到页边。
        """
        Twips = _dx("Twips")
        qn = _dx("qn")
        try:
            pPr = paragraph._p.get_or_add_pPr()
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is None:
                # python-docx 1.1 不认 pBdr，自己按 schema 顺序插进去
                pBdr = _dx("oxml_parse_xml")(
                    "<w:pBdr %s/>" % _dx("nsdecls")("w"))
                successors = (
                    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                    "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                    "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
                    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                    "w:contextualSpacing", "w:mirrorIndents",
                    "w:suppressOverlap", "w:jc", "w:textDirection",
                    "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                    "w:pPrChange")
                inserter = None
                for name in successors:
                    found = pPr.find(qn(name))
                    if found is not None:
                        inserter = found
                        break
                if inserter is not None:
                    inserter.addprevious(pBdr)
                else:
                    pPr.append(pBdr)
            r, g, b = color or (0, 0, 0)
            xml = ('<w:bottom %s w:val="single" w:sz="%d" w:space="1" '
                   'w:color="%02X%02X%02X"/>'
                   % (_dx("nsdecls")("w"), mm_to_eighth_pt(width_mm),
                      r, g, b))
            pBdr.append(_dx("oxml_parse_xml")(xml))
            if right_indent_mm and right_indent_mm > 0:
                paragraph.paragraph_format.right_indent = \
                    Twips(mm_to_twips(right_indent_mm))
        except Exception as exc:                   # pragma: no cover
            self.report.warn("BORDER_FAIL", "段落横线设置失败：%s" % exc)

    # -- 浮动图形（矢量横线 / 印章）---------------------------------------
    def add_floating_picture(self, paragraph, data, x_mm, y_mm, w_mm, h_mm,
                             behind=False):
        # type: (Any, bytes, float, float, float, float, bool) -> bool
        """往指定段落挂一张"绝对定位"的图片。

        python-docx 只能插入随文图片，这里把它生成的 wp:inline 改写成 wp:anchor，
        坐标以页面左上角为基准。锚定段落只提供挂载点，不占排版高度。
        """
        Emu = _dx("Emu")
        Twips = _dx("Twips")
        qn = _dx("qn")
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        WP_NS = ("http://schemas.openxmlformats.org/drawingml/2006/"
                 "wordprocessingDrawing")
        try:
            run = paragraph.add_run()
            run.font.size = _dx("Pt")(1)
            run.add_picture(io.BytesIO(data),
                            width=Emu(mm_to_emu(w_mm)),
                            height=Emu(mm_to_emu(h_mm)))
            drawing = run._element.find(qn("w:drawing"))
            if drawing is None:
                return False
            inline = drawing[0]
            graphic = inline.find("{%s}graphic" % A_NS)
            doc_pr = inline.find("{%s}docPr" % WP_NS)
            extent = inline.find("{%s}extent" % WP_NS)
            if extent is None or graphic is None:
                return False
            cx = extent.get("cx")
            cy = extent.get("cy")

            xml = (
                '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" '
                'simplePos="0" relativeHeight="251658240" behindDoc="%d" '
                'locked="0" layoutInCell="1" allowOverlap="1">'
                '<wp:simplePos x="0" y="0"/>'
                '<wp:positionH relativeFrom="page"><wp:posOffset>%d'
                '</wp:posOffset></wp:positionH>'
                '<wp:positionV relativeFrom="page"><wp:posOffset>%d'
                '</wp:posOffset></wp:positionV>'
                '<wp:extent cx="%s" cy="%s"/>'
                '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
                '<wp:wrapNone/>'
                '</wp:anchor>'
            ) % (_dx("nsdecls")("wp", "a", "pic", "r"), 1 if behind else 0,
                 mm_to_emu(x_mm), mm_to_emu(y_mm), cx, cy)
            anchor_el = _dx("oxml_parse_xml")(xml)
            if doc_pr is not None:
                anchor_el.append(doc_pr)
            anchor_el.append(graphic)
            drawing.remove(inline)
            drawing.append(anchor_el)
            return True
        except Exception as exc:                   # pragma: no cover
            self.report.warn("FLOAT_IMAGE_FAIL",
                             "浮动图形插入失败，已跳过：%s" % exc)
            return False

    # -- 保存 -------------------------------------------------------------
    def save(self, path):
        # type: (str) -> None
        tmp = "%s.tmp-%d" % (path, os.getpid())
        try:
            if self._last_block == "table":
                self.add_empty_paragraph()
            self.doc.save(tmp)
            os.replace(tmp, path)
        except PermissionError:
            raise OfdError(
                "无法写入输出文件（可能正被 Word 打开）",
                part=path,
                hint="请先关闭已经打开的 “%s”，或改用 -o 指定另一个文件名"
                     % os.path.basename(path))
        except OSError as exc:
            raise OfdError("写入输出文件失败：%s" % exc, part=path)
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass


# ---------------------------------------------------------------------------
# B12  转换主流程
# ---------------------------------------------------------------------------

class Options(object):
    __slots__ = ("pages", "no_table", "strict", "verbose", "quiet",
                 "overlay_func")

    def __init__(self):
        self.pages = ""        # type: str
        self.no_table = False
        self.strict = False
        self.verbose = 0
        self.quiet = False
        # 可选：callable(page_model) -> PNG 字节。把矢量图形与印章画成一张
        # 覆盖整页的透明图，插入 Word 后浮动在文字之上/之下，用来还原
        # 红头横线、版记横线、表格外的框线、电子印章等非文字元素。
        self.overlay_func = None


def parse_page_spec(spec, total):
    # type: (str, int) -> List[int]
    if not spec:
        return list(range(1, total + 1))
    wanted = []            # type: List[int]
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            try:
                start, end = int(a), int(b)
            except ValueError:
                raise UsageError("页范围写法不正确：%s" % chunk,
                               hint="正确写法示例：--pages 1-3,5")
            if start > end:
                start, end = end, start
            wanted.extend(range(start, end + 1))
        else:
            try:
                wanted.append(int(chunk))
            except ValueError:
                raise UsageError("页码不是数字：%s" % chunk,
                               hint="正确写法示例：--pages 1-3,5")
    seen = set()
    out = []               # type: List[int]
    for i in wanted:
        if 1 <= i <= total and i not in seen:
            seen.add(i)
            out.append(i)
    if not out:
        raise UsageError("指定的页码超出文档范围（共 %d 页）" % total)
    return out


def count_annotations(doc_root, report):
    # type: (Any, Report) -> None
    total = 0
    for el in iter_local(doc_root, "Annot"):
        total += 1
    if total:
        report.meta["annotations_skipped"] = total
        report.info("ANNOTATIONS_SKIPPED",
                    "已跳过 %d 个注释（印章、水印、手写批注等）" % total)


def load_page(pkg, part, index, meta, res, report, options, template_part,
              stamps=None):
    # type: (OfdPackage, str, int, DocumentMeta, ResourceTable, Report, Options, Optional[str], Any) -> PageModel
    data = pkg.read(part)
    root = parse_xml(data, part)
    if lname(root) != "Page":
        raise OfdFormatError(
            "页面部件根元素是 <%s>，应为 <Page>" % lname(root),
            part=part, page=index)

    box = meta.page_box
    area = child(root, "Area")
    if area is not None:
        pb = parse_box(text_of(child(area, "PhysicalBox")))
        if pb is not None and pb.w > 1.0 and pb.h > 1.0:
            box = pb
    if box.w <= 1.0 or box.h <= 1.0:
        raise OfdGeometryError(
            "页面尺寸无效（%.2f × %.2f mm）" % (box.w, box.h), page=index)

    page_res = child(root, "PageRes")
    if page_res is not None:
        res.merge_res_root(page_res, posixpath.dirname(part), report)

    page = PageModel(index, box)
    roots = []             # type: List[Any]
    if template_part:
        try:
            tpl_root = parse_xml(pkg.read(template_part), template_part)
            roots.append(tpl_root)
        except OfdError as exc:
            report.warn("TEMPLATE_FAIL",
                        "模板页解析失败，已忽略：%s" % exc.message, page=index)
    roots.append(root)

    text_objects = []      # type: List[Any]
    all_segments = []      # type: List[Seg]
    all_images = []        # type: List[ImageItem]
    all_vectors = []       # type: List[Dict[str, Any]]
    skipped_images = 0
    for node in roots:
        for el in node.iter():
            name = lname(el)
            if name == "TextObject":
                text_objects.append(el)
        all_segments.extend(extract_segments(node, res, box, index, report))
        all_vectors.extend(extract_vectors(node, res, box, index, report))
        imgs, skipped = load_images(node, pkg, res, box, index, report)
        all_images.extend(imgs)
        skipped_images += skipped

    page.text_objects = len(text_objects)
    lines, mode = build_text_lines(text_objects, res, box, index, report)
    page.lines = lines
    page.coord_mode = mode
    page.segments = all_segments
    page.images = all_images
    page.vectors = all_vectors
    page.stamps = list(stamps or [])

    if not options.no_table:
        page.tables = detect_tables(all_segments, lines, box, index,
                                    report, enabled=True)
        if page.tables:
            absorbed = set()
            for t in page.tables:
                for c in t.cells:
                    for ln in c.lines:
                        absorbed.add(id(ln))
            if absorbed:
                report.info("TABLE_ABSORB",
                            "%d 行文字已归入表格" % len(absorbed), page=index)

    page.margin = compute_margins(page)
    page.stats_images_skipped = skipped_images
    return page


def compute_margins(page):
    # type: (PageModel) -> Tuple[float, float, float, float]
    xs = []                # type: List[float]
    ys = []                # type: List[float]
    for ln in page.lines:
        xs.extend([ln.left, ln.right])
        ys.append(ln.y)
    for t in page.tables:
        xs.extend([t.box.x, t.box.x2])
        ys.extend([t.box.y, t.box.y2])
    for im in page.images:
        xs.extend([im.box.x, im.box.x2])
        ys.extend([im.box.y, im.box.y2])
    if not xs:
        return (0.0, 0.0, 0.0, 0.0)
    min_x, max_x = min(xs), max(xs)
    w = page.box.w
    left = min(3.0, max(0.0, min_x - 1.0))
    right = min(3.0, max(0.0, w - max_x - 1.0))
    if w - left - right < w * 0.5:
        left, right = 1.0, 1.0
    return (left, 0.0, right, 0.0)


def _item_top(page, obj, kind):
    # type: (PageModel, Any, str) -> float
    if kind == "line":
        return max(0.0, obj.y - obj.size_mm)
    return obj.box.y


def _item_height(obj, kind):
    # type: (Any, str) -> float
    if kind == "line":
        gap = obj.gap_to_next if obj.gap_to_next else 1.32 * obj.size_mm
        return max(gap, obj.size_mm * 1.2)
    return obj.box.h


def classify_border_rules(page):
    # type: (PageModel) -> Tuple[Dict[int, List[Seg]], List[Tuple[float, float, float, float]]]
    """把"紧贴文字下方的横线"挑出来，改用段落边框表达。

    Word 是流式排版，文字的实际落点与 OFD 基线有几毫米浮动；如果这类线还画在
    整页图层里，就会从文字中间穿过去（版记区最常见）。做成段落底边框后，
    线永远跟着文字走，不会再压字。

    返回 (行对象 id -> [Seg], 需要从图层里排除的矩形)。
    """
    rules = {}      # type: Dict[int, List[Seg]]
    skip = []       # type: List[Tuple[float, float, float, float]]
    if not page.lines:
        return (rules, skip)
    for seg in page.segments:
        if seg.orient != "h" or not getattr(seg, "from_fill", False):
            continue
        best = None
        best_gap = 99.0
        for line in page.lines:
            gap = seg.c - line.y          # 正数：线在基线下方
            if 0.8 <= gap <= 5.0 and gap < best_gap:
                best = line
                best_gap = gap
        if best is None:
            continue
        rules.setdefault(id(best), []).append(seg)
        skip.append((seg.a, seg.c - seg.width / 2.0,
                     seg.b, seg.c + seg.width / 2.0))
    return (rules, skip)


def render_page(builder, page, report, options=None):
    # type: (DocxBuilder, PageModel, Report, Optional[Options]) -> None
    items = []             # type: List[Tuple[float, int, str, Any]]
    for ln in page.flow_lines():
        items.append((ln.y, 0, "line", ln))
    for t in page.tables:
        items.append((t.box.y, 1, "table", t))
    for im in page.images:
        items.append((im.box.y, 2, "image", im))
    items.sort(key=lambda it: (it[0], it[1]))

    # 顶边距取第一块内容的上沿，这样纵向空隙由"段前空白"精确补齐，
    # 而不是靠堆空段落或绝对定位。
    top_margin = 0.0
    if items:
        first_top = _item_top(page, items[0][3], items[0][2])
        top_margin = min(max(0.0, first_top - 0.5), page.box.h * 0.45)
    left, _t, right, _b = page.margin
    builder.add_section(page.box, (left, top_margin, right, 0.0))
    avail_w = max(1.0, page.box.w - left - right)

    if not items:
        builder.add_empty_paragraph()
        report.warn("EMPTY_PAGE", "该页没有任何可转换内容，已输出空白页",
                    page=page.index)
        report.pages.append({
            "index": page.index, "size_mm": [round(page.box.w, 2), round(page.box.h, 2)],
            "paragraphs": 1, "empty_paragraphs": 1, "tables": 0,
            "images": 0, "lines": 0,
            "images_skipped": 0, "coord_mode": page.coord_mode,
        })
        return

    border_rules, skip_rects = classify_border_rules(page)
    cursor = top_margin
    paragraphs = 0
    tables = 0
    images = 0
    empties = 0
    prev_kind = None
    first_para = None
    for (_sort_y, _order, kind, obj) in items:
        top = _item_top(page, obj, kind)
        gap = top - cursor
        if gap > 1.0:
            first_para = builder.add_spacer(gap) or first_para
            paragraphs += 1
            empties += 1
        elif prev_kind == "table" and kind == "table":
            first_para = builder.add_spacer(0.0) or first_para
            paragraphs += 1
            empties += 1

        if kind == "line":
            para = builder.add_line(obj, left, avail_w, page.index)
            first_para = first_para or para
            paragraphs += 1
            for seg in border_rules.get(id(obj), []):
                right_indent = page.box.w - right - seg.b
                builder.add_paragraph_bottom_border(
                    para, seg.color, seg.width, max(0.0, right_indent))
        elif kind == "table":
            builder.add_table(obj, left, avail_w, page.index)
            tables += 1
        else:
            para = builder.add_image(obj, left, avail_w, page.index)
            first_para = first_para or para
            images += 1
        cursor = top + _item_height(obj, kind)
        prev_kind = kind

    # 矢量图形与印章：画一张覆盖整页的透明图，浮动到页面上。
    # 这样红头横线、版记横线、印章都能还原，且不影响流式排版。
    # 紧贴文字下方的横线已改成段落边框，这里要从图层里排除，避免画两遍。
    if options is not None and options.overlay_func is not None:
        try:
            blob = options.overlay_func(page, skip_rects)
        except TypeError:
            blob = options.overlay_func(page)
        except Exception as exc:               # pragma: no cover
            blob = None
            report.warn("OVERLAY_FAIL", "图形图层生成失败，已跳过：%s" % exc,
                        page=page.index)
        if blob:
            target = first_para or builder.add_spacer(0.0)
            builder.add_floating_picture(target, blob, 0.0, 0.0,
                                         page.box.w, page.box.h)

    report.pages.append({
        "index": page.index,
        "size_mm": [round(page.box.w, 2), round(page.box.h, 2)],
        "paragraphs": paragraphs,
        "empty_paragraphs": empties,
        "tables": tables,
        "images": images,
        "lines": len(page.lines),
        "images_skipped": getattr(page, "stats_images_skipped", 0),
        "coord_mode": page.coord_mode,
    })


def convert_ofd(input_path, output_path, options, report):
    # type: (str, str, Options, Report) -> None
    pkg = open_package(input_path)
    try:
        meta, ofd_root, ofd_part, doc_root, doc_part, res = load_document(pkg, report)
        report.meta["ofd_version"] = meta.version
        report.meta["doc_bodies"] = meta.doc_bodies
        meta.stamps_by_id = collect_stamps(pkg, ofd_root, ofd_part, report)
        count_annotations(doc_root, report)

        page_entries = meta.pages
        total = len(page_entries)
        if total == 0:
            report.warn("NO_PAGES", "该文档没有任何页面，将输出一份空文档")
            chosen = []
        else:
            wanted = parse_page_spec(options.pages, total)
            chosen = [page_entries[i - 1] for i in wanted]

        builder = DocxBuilder(report)
        for idx, (part, template_part, page_id) in enumerate(chosen, 1):
            page = load_page(pkg, part, idx, meta, res, report, options,
                             template_part,
                             meta.stamps_by_id.get(str(page_id)))
            render_page(builder, page, report, options)
        builder.save(output_path)
        report.meta["ok"] = True
        unmapped = res.unmapped_fonts()
        if unmapped:
            report.info("FONT_UNMAPPED",
                        "以下字体名未能映射到系统中文字体，将直接写入原名：%s"
                        % "、".join(unmapped[:8]))
    finally:
        pkg.close()


class OfdDocument(object):
    """阅读器用的文档门面：持有已打开的包与解析结果，按页懒加载。"""

    def __init__(self, path, report=None):
        # type: (str, Optional[Report]) -> None
        self.path = path
        self.report = report if report is not None else Report()
        self.pkg = open_package(path)
        try:
            (self.meta, self.ofd_root, self.ofd_part,
             self.doc_root, self.doc_part, self.res) = \
                load_document(self.pkg, self.report)
            self.meta.stamps_by_id = collect_stamps(
                self.pkg, self.ofd_root, self.ofd_part, self.report)
            count_annotations(self.doc_root, self.report)
        except Exception:
            self.pkg.close()
            raise
        self.options = Options()

    @property
    def page_count(self):
        # type: () -> int
        return len(self.meta.pages)

    @property
    def title(self):
        # type: () -> str
        return self.meta.title or os.path.splitext(os.path.basename(self.path))[0]

    def get_page(self, index, options=None):
        # type: (int, Optional[Options]) -> PageModel
        if index < 1 or index > self.page_count:
            raise UsageError("页码超出范围：%s（共 %s 页）"
                             % (index, self.page_count))
        part, template, page_id = self.meta.pages[index - 1]
        return load_page(self.pkg, part, index, self.meta, self.res,
                         self.report, options or self.options, template,
                         self.meta.stamps_by_id.get(str(page_id)))

    def close(self):
        # type: () -> None
        try:
            self.pkg.close()
        except Exception:
            pass

    def __enter__(self):
        # type: () -> OfdDocument
        return self

    def __exit__(self, exc_type, exc, tb):
        # type: (Any, Any, Any) -> None
        self.close()


# ---------------------------------------------------------------------------
# B13  命令行
# ---------------------------------------------------------------------------

def build_argparser():
    # type: () -> argparse.ArgumentParser
    parser = argparse.ArgumentParser(
        prog="ofd2docx",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="把 OFD 版式文档转换为可编辑的 Word（.docx）文档",
        epilog=("示例：\n"
                "  python ofd2docx.py 通知.ofd\n"
                "  python ofd2docx.py 通知.ofd -o 通知.docx --report r.json -v\n"
                "  python ofd2docx.py --self-test\n"))
    parser.add_argument("input", nargs="?", help="输入 .ofd 文件路径")
    parser.add_argument("-o", "--out", default="",
                        help="输出 .docx 路径（默认与输入同名同目录）")
    parser.add_argument("--force", action="store_true",
                        help="输出文件已存在时覆盖（默认拒绝）")
    parser.add_argument("--pages", default="",
                        help="只转换指定页，如 1-3,5（默认全部）")
    parser.add_argument("--no-table", action="store_true", dest="no_table",
                        help="关闭表格识别，全部按正文输出")
    parser.add_argument("--redhead", action="store_true",
                        help="把灰度红头（机关标志 + 分隔线）按公文红绘制到 Word 图层")
    parser.add_argument("--seal-color", action="store_true", dest="seal_color",
                        help="把灰度印章按朱红绘制到 Word 图层")
    parser.add_argument("--strict", action="store_true",
                        help="有任何告警时返回退出码 4")
    parser.add_argument("--report", default="",
                        help="把转换报告写入指定 JSON 文件")
    parser.add_argument("-q", "--quiet", action="store_true", help="只输出错误")
    parser.add_argument("-v", "--verbose", action="count", default=0,
                        help="输出更多过程信息（可重复）")
    parser.add_argument("--selftest-dir", default="",
                        help="指定自测产物目录（默认临时目录）")
    parser.add_argument("--version", action="version",
                        version="ofd2docx %s" % __version__)
    parser.add_argument("--self-test", action="store_true", dest="self_test",
                        help="运行内建自测（生成样本并校验转换结果）")
    return parser


def _fail(message, code, quiet=False):
    # type: (str, int, bool) -> int
    if not quiet:
        sys.stderr.write(message.rstrip() + "\n")
    return code


def ensure_console_encoding():
    # type: () -> None
    """控制台编码写不了中文时改用 UTF-8。

    英文版 Windows 的代码页是 1252，直接 print 中文会抛 UnicodeEncodeError，
    整个自测会因此中断。这里先探测：能编码就不动（中文系统保持 cp936），
    不能编码才切 UTF-8 并允许替换字符，宁可显示成问号也不要崩。
    """
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            continue
        encoding = getattr(stream, "encoding", None) or "ascii"
        try:
            u"中文".encode(encoding)
            continue
        except Exception:
            pass
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main(argv=None):
    # type: (Optional[List[str]]) -> int
    ensure_console_encoding()
    args = build_argparser().parse_args(argv)

    if args.self_test:
        return run_self_test(args.selftest_dir, args.verbose)

    if not args.input:
        return _fail("错误：请指定输入的 .ofd 文件；"
                     "若只想验证工具本身，可运行 --self-test", EXIT_USAGE,
                     args.quiet)

    src = os.path.abspath(args.input)
    if not os.path.exists(src):
        return _fail("错误：输入文件不存在\n    路径：%s" % src, EXIT_USAGE,
                     args.quiet)
    if os.path.isdir(src):
        return _fail("错误：输入路径是一个目录，不是文件\n    路径：%s" % src,
                     EXIT_USAGE, args.quiet)
    if not os.access(src, os.R_OK):
        return _fail("错误：没有读取权限\n    路径：%s" % src, EXIT_USAGE,
                     args.quiet)
    if os.path.getsize(src) == 0:
        return _fail("错误：输入文件大小为 0\n    路径：%s" % src, EXIT_FORMAT,
                     args.quiet)

    if args.out:
        dst = os.path.abspath(args.out)
    else:
        stem, _ext = os.path.splitext(src)
        dst = stem + ".docx"

    if os.path.abspath(dst) == os.path.abspath(src):
        return _fail("错误：输出路径不能与输入文件相同", EXIT_USAGE, args.quiet)
    if os.path.isdir(dst):
        return _fail("错误：输出路径是一个目录\n    路径：%s" % dst, EXIT_USAGE,
                     args.quiet)
    if os.path.exists(dst) and not args.force:
        return _fail("错误：输出文件已存在，未覆盖\n    路径：%s\n"
                     "    若确认覆盖，请加 --force" % dst, EXIT_USAGE, args.quiet)
    parent = os.path.dirname(dst)
    if parent and not os.path.isdir(parent):
        try:
            os.makedirs(parent)
        except OSError as exc:
            return _fail("错误：无法创建输出目录：%s" % exc, EXIT_USAGE,
                         args.quiet)

    if not src.lower().endswith(".ofd") and not args.quiet:
        sys.stderr.write("提示：输入文件扩展名不是 .ofd，仍按 OFD 格式尝试解析\n")

    options = Options()
    options.pages = args.pages
    options.no_table = args.no_table
    options.strict = args.strict
    options.verbose = args.verbose
    options.quiet = args.quiet
    # 矢量图形与印章靠渲染层画成透明图层，这里延迟导入避免核心模块强依赖 Pillow
    try:
        import ofd_view as _view
        red = bool(getattr(args, "redhead", False))
        seal = bool(getattr(args, "seal_color", False))
        if red or seal:
            options.overlay_func = (
                lambda page, skip=None: _view.overlay_png(
                    page, skip, redhead=red, seal_red=seal))
        else:
            options.overlay_func = _view.overlay_png
    except Exception:
        options.overlay_func = None

    report = Report()
    started = time.time()
    try:
        convert_ofd(src, dst, options, report)
    except OfdError as exc:
        report.elapsed_ms = int((time.time() - started) * 1000)
        if args.report:
            _write_report(args.report, src, dst, report, exc)
        return _fail(exc.user_message(), exc.exit_code, args.quiet)
    except KeyboardInterrupt:
        return _fail("已取消", EXIT_USAGE, args.quiet)
    except Exception as exc:                       # pragma: no cover
        if args.verbose:
            import traceback
            traceback.print_exc()
        report.elapsed_ms = int((time.time() - started) * 1000)
        return _fail("[意外错误] %s: %s\n"
                     "    建议：加 -v 重新运行以查看完整堆栈，"
                     "并把该文件反馈给维护者" % (type(exc).__name__, exc),
                     EXIT_FORMAT, args.quiet)

    report.elapsed_ms = int((time.time() - started) * 1000)
    if args.report:
        _write_report(args.report, src, dst, report, None)

    if not args.quiet:
        tot = report.totals()
        print("转换完成：%s" % dst)
        print("  页数 %d｜段落 %d｜表格 %d｜图片 %d｜耗时 %d ms"
              % (tot["pages"], tot["paragraphs"], tot["tables"],
                 tot["images"], report.elapsed_ms))
        if tot["images_skipped"]:
            print("  跳过图片 %d 张" % tot["images_skipped"])
        if tot["signatures_skipped"] or tot["annotations_skipped"]:
            print("  跳过签章 %d 个、注释 %d 个"
                  % (tot["signatures_skipped"], tot["annotations_skipped"]))
        warns = [w for w in report.warnings if w.level == "warn"]
        if warns:
            print("  告警 %d 条（加 -v 查看明细，或 --report 导出完整报告）" % len(warns))
            if args.verbose:
                for w in warns[:40]:
                    print("    - %s" % w)
            else:
                for w in warns[:5]:
                    print("    - %s" % w)
                if len(warns) > 5:
                    print("    ...（其余 %d 条略）" % (len(warns) - 5))

    if args.strict and report.has_warning():
        return EXIT_PARTIAL
    return EXIT_OK


def _write_report(path, src, dst, report, exc):
    # type: (str, str, str, Report, Optional[OfdError]) -> None
    payload = report.as_dict()
    payload["ok"] = exc is None and report.meta.get("ok", False)
    payload["input"] = src
    payload["output"] = dst if exc is None else None
    if exc is not None:
        payload["error"] = {
            "kind": exc.kind, "message": exc.message, "part": exc.part,
            "page": exc.page, "obj": exc.obj_id, "hint": exc.hint,
        }
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False, indent=2))
    except OSError as write_exc:
        sys.stderr.write("警告：转换报告写入失败（%s）\n" % write_exc)


# ---------------------------------------------------------------------------
# B14  内建自测：现场合成 OFD 样本并校验转换结果
# ---------------------------------------------------------------------------

OFD_NS = "http://www.ofdspec.org/2016"


def _xml_decl():
    # type: () -> str
    return '<?xml version="1.0" encoding="UTF-8"?>'


def _esc(text):
    # type: (str) -> str
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _ofd_xml(doc_part):
    # type: (str) -> str
    return (_xml_decl() +
            '<ofd:OFD xmlns:ofd="%s" Version="1.0" DocType="OFD">'
            '<ofd:DocBody><ofd:DocInfo><ofd:DocID>TEST0001</ofd:DocID>'
            '</ofd:DocInfo><ofd:DocRoot>%s</ofd:DocRoot>'
            '</ofd:DocBody></ofd:OFD>' % (OFD_NS, doc_part))


def _document_xml(page_parts, box):
    # type: (List[str], Tuple[float, float]) -> str
    pages = "".join('<ofd:Page ID="%d" BaseLoc="%s"/>' % (i + 1, p)
                    for i, p in enumerate(page_parts))
    return (_xml_decl() +
            '<ofd:Document xmlns:ofd="%s"><ofd:CommonData>'
            '<ofd:MaxUnitID>999</ofd:MaxUnitID>'
            '<ofd:PageArea><ofd:PhysicalBox>0 0 %.2f %.2f</ofd:PhysicalBox>'
            '</ofd:PageArea><ofd:PublicRes>PublicRes.xml</ofd:PublicRes>'
            '<ofd:DocumentRes>DocumentRes.xml</ofd:DocumentRes>'
            '</ofd:CommonData><ofd:Pages>%s</ofd:Pages></ofd:Document>'
            % (OFD_NS, box[0], box[1], pages))


def _public_res(fonts):
    # type: (List[Tuple[str, str]]) -> str
    items = "".join('<ofd:Font ID="%s" FontName="%s"/>' % (fid, _esc(name))
                    for fid, name in fonts)
    return (_xml_decl() +
            '<ofd:Res xmlns:ofd="%s" BaseLoc="Res"><ofd:Fonts>%s</ofd:Fonts>'
            '</ofd:Res>' % (OFD_NS, items))


def _document_res(media):
    # type: (List[Tuple[str, str, bytes]]) -> str
    items = []
    for (mid, fname, _data) in media:
        items.append('<ofd:MultiMedia ID="%s" Type="Image" Format="PNG">'
                     '<ofd:MediaFile>%s</ofd:MediaFile></ofd:MultiMedia>'
                     % (mid, fname))
    return (_xml_decl() +
            '<ofd:Res xmlns:ofd="%s" BaseLoc="Res"><ofd:DrawParams>'
            '<ofd:DrawParam ID="2"><ofd:FillColor Value="0 0 0"/>'
            '<ofd:StrokeColor Value="0 0 0"/></ofd:DrawParam>'
            '</ofd:DrawParams><ofd:MultiMedias>%s</ofd:MultiMedias></ofd:Res>'
            % (OFD_NS, "".join(items)))


def _page_content(box, body):
    # type: (Tuple[float, float], str) -> str
    return (_xml_decl() +
            '<ofd:Page xmlns:ofd="%s"><ofd:Area><ofd:PhysicalBox>0 0 %.2f %.2f'
            '</ofd:PhysicalBox></ofd:Area><ofd:Content><ofd:Layer ID="1">%s'
            '</ofd:Layer></ofd:Content></ofd:Page>'
            % (OFD_NS, box[0], box[1], body))


def _text_obj(oid, x, y, text, size=4.0, font="3", mode="B"):
    # type: (int, float, float, str, float, str, str) -> str
    width = sum(est_advance(c, size) for c in text) or size
    height = size * 1.25
    if mode == "B":
        tx, ty = 0.0, size
    else:
        tx, ty = x, y + size
    return ('<ofd:TextObject ID="%d" Boundary="%.3f %.3f %.3f %.3f" Font="%s" '
            'Size="%.4f" DrawParam="2"><ofd:TextCode X="%.4f" Y="%.4f">%s'
            '</ofd:TextCode></ofd:TextObject>'
            % (oid, x, y, width, height, font, size, tx, ty, _esc(text)))


def _path_obj(oid, x1, y1, x2, y2, width=0.35):
    # type: (int, float, float, float, float, float) -> str
    bw = abs(x2 - x1) or 0.1
    bh = abs(y2 - y1) or 0.1
    return ('<ofd:PathObject ID="%d" Boundary="%.3f %.3f %.3f %.3f" '
            'LineWidth="%.3f" Stroke="true" Fill="false" '
            'AbbreviatedData="M %.3f %.3f L %.3f %.3f"/>'
            % (oid, min(x1, x2), min(y1, y2), bw, bh, width, x1, y1, x2, y2))


def _image_obj(oid, rid, x, y, w, h):
    # type: (int, str, float, float, float, float) -> str
    return ('<ofd:ImageObject ID="%d" Boundary="%.3f %.3f %.3f %.3f" '
            'ResourceID="%s"/>' % (oid, x, y, w, h, rid))


def _png_bytes(w, h, rgb):
    # type: (int, int, Tuple[int, int, int]) -> bytes
    import zlib
    raw = bytearray()
    row = bytes(bytearray(rgb)) * w
    for _ in range(h):
        raw.append(0)
        raw.extend(row)

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data +
                struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) +
            chunk(b"IDAT", zlib.compress(bytes(raw))) + chunk(b"IEND", b""))


def _write_package(path, pages, box=(210.0, 297.0), fonts=None, media=None):
    # type: (str, List[Dict[str, Any]], Tuple[float, float], Optional[List[Tuple[str, str]]], Optional[List[Tuple[str, str, bytes]]]) -> None
    fonts = fonts or [("3", "simsun"), ("6", "simhei")]
    media = media or []
    parts = ["Pages/Page_%d/Content.xml" % i for i in range(len(pages))]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("OFD.xml", _ofd_xml("Doc_0/Document.xml"))
        zf.writestr("Doc_0/Document.xml", _document_xml(parts, box))
        zf.writestr("Doc_0/PublicRes.xml", _public_res(fonts))
        zf.writestr("Doc_0/DocumentRes.xml", _document_res(media))
        for i, pg in enumerate(pages):
            zf.writestr("Doc_0/" + parts[i],
                        _page_content(pg.get("box", box), pg["body"]))
        for (_mid, fname, data) in media:
            zf.writestr("Doc_0/Res/" + fname, data)


def _docx_open(path):
    # type: (str) -> Any
    Document = _dx("Document")
    return Document(path)


def _docx_all_text(doc):
    # type: (Any) -> str
    chunks = [p.text for p in doc.paragraphs]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                chunks.append(cell.text)
    return "\n".join(chunks)


def _cross_check(path):
    # type: (str) -> Tuple[bool, str]
    """用另一个 Python 版本上的 python-docx 独立读取一次，确认 OOXML 结构没问题。"""
    import subprocess
    candidate = r"C:\Python314\python.exe"
    if not os.path.exists(candidate):
        return (True, "跳过（未找到第二解释器）")
    code = (
        "import sys\n"
        "from docx import Document\n"
        "d = Document(sys.argv[1])\n"
        "print('%d|%d|%d|%d' % (len(d.paragraphs), len(d.tables),\n"
        "      len(d.sections), len(d.inline_shapes)))\n")
    try:
        out = subprocess.check_output([candidate, "-c", code, path],
                                      stderr=subprocess.STDOUT, timeout=120)
    except Exception as exc:
        return (False, "第二读取器执行失败：%s" % exc)
    text = out.decode("utf-8", "replace").strip()
    try:
        doc = _docx_open(path)
        mine = "%d|%d|%d|%d" % (len(doc.paragraphs), len(doc.tables),
                                len(doc.sections), len(doc.inline_shapes))
    except Exception as exc:
        return (False, "主读取器执行失败：%s" % exc)
    if text == mine:
        return (True, mine)
    return (False, "主读取器 %s ≠ 第二读取器 %s" % (mine, text))


def _check_ooxml(path):
    # type: (str) -> Tuple[bool, str]
    """结构校验：XML 合法、无相邻表格、每个单元格有段落。"""
    etree = _get_etree()
    qn = _dx("qn")
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                etree.fromstring(zf.read(name))
            except Exception as exc:
                return (False, "部件 %s 不是合法 XML：%s" % (name, exc))
        body = etree.fromstring(zf.read("word/document.xml"))
    w = "{%s}" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = body.find(w + "body")
    if root is None:
        return (False, "document.xml 缺少 body")
    prev = None
    for el in root:
        tag = el.tag
        if tag == w + "tbl" and prev == w + "tbl":
            return (False, "存在两个相邻的 w:tbl，Word 会把它们合并")
        prev = tag
    for tc in root.iter(w + "tc"):
        if tc.find(w + "p") is None:
            return (False, "存在没有段落的 w:tc")
    return (True, "")


def _run_convert(src, dst, pages="", no_table=False):
    # type: (str, str, str, bool) -> Report
    report = Report()
    opts = Options()
    opts.pages = pages
    opts.no_table = no_table
    convert_ofd(src, dst, opts, report)
    return report


def _cli(argv):
    # type: (List[str]) -> int
    return main(list(argv) + ["-q"])


def run_self_test(out_dir="", verbose=0):
    # type: (str, int) -> int
    base = out_dir or tempfile.mkdtemp(prefix="ofd2docx_selftest_")
    try:
        os.makedirs(base)
    except OSError:
        pass
    results = []           # type: List[Tuple[str, bool, str]]

    def check(name, cond, detail=""):
        # type: (str, bool, str) -> None
        ok = bool(cond)
        results.append((name, ok, detail))
        print("  [%s] %s%s" % ("OK" if ok else "FAIL", name,
                               "" if ok else ("  ← " + str(detail))))

    print("ofd2docx %s 自测，产物目录：%s\n" % (__version__, base))

    # ---- 用例 1：中文正文 + 4×3 表格 + 图片，双页 A4 ----------------
    print("[1] 中文正文 + 4x3 表格 + 图片（双页 A4）")
    body1 = []           # type: List[str]
    body1.append(_text_obj(10, 55.0, 22.0, "浙江农商银行测试文档", size=7.0))
    body1.append(_text_obj(11, 20.0, 48.0, "一、测试段落标题", size=5.6))
    body1.append(_text_obj(12, 20.0, 58.0,
                           "这是一段用于验证转换正确性的中文正文，包含数字2026与英文ABC。",
                           size=5.6))
    xs = [20.0, 70.0, 120.0, 170.0]
    ys = [78.0, 88.0, 98.0, 108.0, 118.0]
    oid = 100
    for y in ys:
        body1.append(_path_obj(oid, xs[0], y, xs[-1], y))
        oid += 1
    for x in xs:
        body1.append(_path_obj(oid, x, ys[0], x, ys[-1]))
        oid += 1
    cells = [["项目", "金额", "备注"], ["存款", "100", "测试"],
             ["贷款", "200", "验证"], ["合计", "300", "完毕"]]
    for r in range(4):
        for c in range(3):
            body1.append(_text_obj(oid, xs[c] + 4.0, ys[r] + 2.5,
                                   cells[r][c], size=4.0))
            oid += 1
    png = _png_bytes(32, 24, (200, 60, 60))
    body1.append(_image_obj(oid, "90", 20.0, 128.0, 40.0, 30.0))

    body2 = []
    body2.append(_text_obj(200, 20.0, 40.0, "中文", size=5.0))
    body2.append(_text_obj(201, 30.0, 40.0, "ABC", size=5.0))
    body2.append(_text_obj(202, 38.4, 40.0, "测试", size=5.0))
    body2.append(_text_obj(203, 20.0, 60.0, "第二页结束标记", size=5.6))

    src1 = os.path.join(base, "case1_basic.ofd")
    dst1 = os.path.join(base, "case1_basic.docx")
    _write_package(src1, [{"body": "".join(body1)}, {"body": "".join(body2)}],
                   media=[("90", "img0.png", png)])
    try:
        rep1 = _run_convert(src1, dst1)
        check("转换未抛异常", True)
        doc1 = _docx_open(dst1)
        text1 = _docx_all_text(doc1)
        check("生成 docx 可被 python-docx 打开", True)
        check("页数为 2（sections）", len(doc1.sections) == 2,
              "实际 %d" % len(doc1.sections))
        check("识别到 1 个表格", len(doc1.tables) == 1,
              "实际 %d" % len(doc1.tables))
        if doc1.tables:
            t1 = doc1.tables[0]
            check("表格 4 行", len(t1.rows) == 4, "实际 %d" % len(t1.rows))
            check("表格 3 列", len(t1.columns) == 3, "实际 %d" % len(t1.columns))
            check("单元格文字命中", "存款" in t1.rows[1].cells[0].text,
                  repr(t1.rows[1].cells[0].text))
        check("标题文字逐字命中", "浙江农商银行测试文档" in text1)
        check("正文文字逐字命中",
              "这是一段用于验证转换正确性的中文正文，包含数字2026与英文ABC。" in text1)
        check("跨对象同行未被拆开", "中文ABC测试" in text1,
              "未找到「中文ABC测试」")
        check("图片已插入", len(doc1.inline_shapes) == 1,
              "实际 %d" % len(doc1.inline_shapes))
        w1 = doc1.sections[0].page_width.twips
        check("页面宽度按 210mm 还原", abs(w1 - mm_to_twips(210.0)) <= 2,
              "twips=%d" % w1)
        ok_x, detail_x = _cross_check(dst1)
        check("第二读取器交叉校验", ok_x, detail_x)
        ok_s, detail_s = _check_ooxml(dst1)
        check("OOXML 结构校验", ok_s, detail_s)
        check("转换报告记录了页数", rep1.totals()["pages"] == 2)
    except Exception as exc:
        check("用例 1 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 2：坐标写法 B / C 一致性 ------------------------------
    print("[2] 坐标写法一致性（Boundary 相对 vs 页面绝对）")
    try:
        texts = ["坐标一致性验证甲段落", "坐标一致性验证乙段落", "坐标一致性验证丙段落"]
        for mode, name in (("B", "coords_b"), ("C", "coords_c")):
            body = []
            for i, txt in enumerate(texts):
                body.append(_text_obj(300 + i, 30.0, 40.0 + i * 12.0, txt,
                                      size=5.0, mode=mode))
            _write_package(os.path.join(base, name + ".ofd"), [{"body": "".join(body)}])
        _run_convert(os.path.join(base, "coords_b.ofd"),
                     os.path.join(base, "coords_b.docx"))
        _run_convert(os.path.join(base, "coords_c.ofd"),
                     os.path.join(base, "coords_c.docx"))
        db = _docx_open(os.path.join(base, "coords_b.docx"))
        dc = _docx_open(os.path.join(base, "coords_c.docx"))
        tb = "\n".join(p.text for p in db.paragraphs)
        tc = "\n".join(p.text for p in dc.paragraphs)
        check("两种坐标写法输出文本一致", tb == tc,
              "%r vs %r" % (tb[:40], tc[:40]))
        same_para = len(db.paragraphs) == len(dc.paragraphs)
        check("两种坐标写法段落数一致", same_para,
              "%d vs %d" % (len(db.paragraphs), len(dc.paragraphs)))
        if same_para:
            worst = 0
            for pb, pc in zip(db.paragraphs, dc.paragraphs):
                ib = pb.paragraph_format.left_indent
                ic = pc.paragraph_format.left_indent
                vb = ib.twips if ib is not None else 0
                vc = ic.twips if ic is not None else 0
                worst = max(worst, abs(vb - vc))
            check("左缩进偏差 < 1mm", worst <= 57, "最大偏差 %d twips" % worst)
        check("三行文字全部命中", all(t in tb for t in texts))
    except Exception as exc:
        check("用例 2 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 3：整页扫描件 ----------------------------------------
    print("[3] 整页扫描件（单张图铺满 A4）")
    try:
        scan_png = _png_bytes(80, 110, (240, 240, 240))
        body = [_image_obj(400, "91", 0.0, 0.0, 210.0, 297.0)]
        _write_package(os.path.join(base, "case3_scan.ofd"), [{"body": "".join(body)}],
                       media=[("91", "scan.png", scan_png)])
        _run_convert(os.path.join(base, "case3_scan.ofd"),
                     os.path.join(base, "case3_scan.docx"))
        doc3 = _docx_open(os.path.join(base, "case3_scan.docx"))
        check("扫描件不产生表格", len(doc3.tables) == 0)
        check("扫描件图片已插入", len(doc3.inline_shapes) == 1)
        if doc3.inline_shapes:
            wemu = doc3.inline_shapes[0].width
            check("图片宽度接近页宽", abs(wemu - mm_to_emu(210.0)) <= mm_to_emu(5.0),
                  "%.1f mm" % (wemu / EMU_PER_MM))
    except Exception as exc:
        check("用例 3 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 4：横向页与非标尺寸 ----------------------------------
    print("[4] 横向页 297x210")
    try:
        body = [_text_obj(500, 30.0, 40.0, "横向页面测试", size=6.0)]
        _write_package(os.path.join(base, "case4_land.ofd"),
                       [{"body": "".join(body), "box": (297.0, 210.0)}],
                       box=(297.0, 210.0))
        _run_convert(os.path.join(base, "case4_land.ofd"),
                     os.path.join(base, "case4_land.docx"))
        doc4 = _docx_open(os.path.join(base, "case4_land.docx"))
        sec = doc4.sections[0]
        check("页宽按 297mm 还原",
              abs(sec.page_width.twips - mm_to_twips(297.0)) <= 2,
              "twips=%d" % sec.page_width.twips)
        check("页高按 210mm 还原",
              abs(sec.page_height.twips - mm_to_twips(210.0)) <= 2,
              "twips=%d" % sec.page_height.twips)
        check("方向标记为横向", int(sec.orientation) == 1,
              "orientation=%s" % sec.orientation)
    except Exception as exc:
        check("用例 4 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 5：损坏文件 ------------------------------------------
    print("[5] 损坏文件与退出码")
    try:
        bad1 = os.path.join(base, "bad_notzip.ofd")
        with io.open(bad1, "wb") as fh:
            fh.write(b"this is not a zip file")
        code = _cli([bad1, "-o", os.path.join(base, "bad1.docx")])
        check("非 ZIP 文件返回退出码 2", code == EXIT_FORMAT, "实际 %d" % code)

        bad2 = os.path.join(base, "bad_noofd.ofd")
        with zipfile.ZipFile(bad2, "w") as zf:
            zf.writestr("readme.txt", "no ofd here")
        code = _cli([bad2, "-o", os.path.join(base, "bad2.docx")])
        check("缺少 OFD.xml 返回退出码 2", code == EXIT_FORMAT, "实际 %d" % code)

        bad3 = os.path.join(base, "bad_xml.ofd")
        with zipfile.ZipFile(bad3, "w") as zf:
            zf.writestr("OFD.xml", _ofd_xml("Doc_0/Document.xml"))
            zf.writestr("Doc_0/Document.xml",
                        _xml_decl() + "<ofd:Document xmlns:ofd=\"%s\">" % OFD_NS)
        code = _cli([bad3, "-o", os.path.join(base, "bad3.docx")])
        check("XML 语法错误返回退出码 2", code == EXIT_FORMAT, "实际 %d" % code)

        bad4 = os.path.join(base, "bad_dangling.ofd")
        with zipfile.ZipFile(bad4, "w") as zf:
            zf.writestr("OFD.xml", _ofd_xml("Doc_0/Document.xml"))
            zf.writestr("Doc_0/Document.xml",
                        _document_xml(["Pages/Page_0/Content.xml"], (210.0, 297.0)))
        code = _cli([bad4, "-o", os.path.join(base, "bad4.docx")])
        check("BaseLoc 悬空返回退出码 2", code == EXIT_FORMAT, "实际 %d" % code)

        bad5 = os.path.join(base, "bad_encrypted.ofd")
        with zipfile.ZipFile(bad5, "w") as zf:
            zf.writestr("OFD.xml", _ofd_xml("Doc_0/Document.xml"))
            zf.writestr("decryptseed.dat", b"\x00\x01")
        code = _cli([bad5, "-o", os.path.join(base, "bad5.docx")])
        check("加密文件返回退出码 3", code == EXIT_UNSUPPORTED, "实际 %d" % code)

        code = _cli([os.path.join(base, "nope.ofd")])
        check("文件不存在返回退出码 1", code == EXIT_USAGE, "实际 %d" % code)
    except Exception as exc:
        check("用例 5 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 6：边缘情况 ------------------------------------------
    print("[6] 边缘情况（无页面 / 仅下划线 / 仅图片）")
    try:
        _write_package(os.path.join(base, "empty.ofd"), [])
        rep = _run_convert(os.path.join(base, "empty.ofd"),
                           os.path.join(base, "empty.docx"))
        check("零页面仍产出 docx", os.path.exists(os.path.join(base, "empty.docx")))
        check("零页面被记录为告警",
              any(w.code == "NO_PAGES" for w in rep.warnings))

        body = [_text_obj(600, 20.0, 40.0, "下划线反例，不应识别为表格", size=5.0)]
        for i, y in enumerate((60.0, 70.0, 80.0)):
            body.append(_path_obj(610 + i, 20.0, y, 170.0, y))
        _write_package(os.path.join(base, "underline.ofd"), [{"body": "".join(body)}])
        _run_convert(os.path.join(base, "underline.ofd"),
                     os.path.join(base, "underline.docx"))
        doc6 = _docx_open(os.path.join(base, "underline.docx"))
        check("只有横线时不误判为表格", len(doc6.tables) == 0,
              "实际 %d 个表格" % len(doc6.tables))

        only_png = _png_bytes(20, 20, (10, 120, 200))
        body = [_image_obj(700, "92", 40.0, 40.0, 60.0, 40.0)]
        _write_package(os.path.join(base, "imageonly.ofd"), [{"body": "".join(body)}],
                       media=[("92", "only.png", only_png)])
        _run_convert(os.path.join(base, "imageonly.ofd"),
                     os.path.join(base, "imageonly.docx"))
        doc7 = _docx_open(os.path.join(base, "imageonly.docx"))
        check("纯图片页不产生表格", len(doc7.tables) == 0)
        check("纯图片页图片已插入", len(doc7.inline_shapes) == 1)
    except Exception as exc:
        check("用例 6 执行", False, "%s: %s" % (type(exc).__name__, exc))

    try:
        # 用例 7：调色板颜色（ColorSpace Type="Index"）不能当成 RGB 解析
        palette_res = (
            _xml_decl() +
            '<ofd:Res xmlns:ofd="%s" BaseLoc="Res"><ofd:ColorSpaces>'
            '<ofd:ColorSpace ID="3" Type="Index" BitsPerComponent="8">'
            '<ofd:Palette>'
            '<ofd:PaletteEntry Value="0 0 0"/>'
            '<ofd:PaletteEntry Value="255 0 0"/>'
            '</ofd:Palette></ofd:ColorSpace></ofd:ColorSpaces>'
            '<ofd:DrawParams><ofd:DrawParam ID="5">'
            '<ofd:FillColor Value="1" ColorSpace="3"/>'
            '</ofd:DrawParam></ofd:DrawParams></ofd:Res>' % OFD_NS)
        palette_body = (
            '<ofd:PathObject ID="800" Boundary="20 40 100 2" Fill="true" '
            'Stroke="false" AbbreviatedData="M 0 0 L 100 0 L 100 2 L 0 2 C">'
            '<ofd:FillColor Value="1" ColorSpace="3"/></ofd:PathObject>'
            + _text_obj(801, 20.0, 60.0, "调色板红色", size=6.0))
        palette_path = os.path.join(base, "palette.ofd")
        with zipfile.ZipFile(palette_path, "w") as zf:
            zf.writestr("OFD.xml", _ofd_xml("Doc_0/Document.xml"))
            zf.writestr("Doc_0/Document.xml",
                        _document_xml(["Pages/Page_1/Content.xml"],
                                      (210.0, 297.0)))
            zf.writestr("Doc_0/PublicRes.xml", palette_res)
            zf.writestr("Doc_0/DocumentRes.xml",
                        _xml_decl() + ('<ofd:Res xmlns:ofd="%s"/>' % OFD_NS))
            zf.writestr("Doc_0/Pages/Page_1/Content.xml",
                        _page_content((210.0, 297.0), palette_body))
        pkg_p = open_package(palette_path)
        try:
            meta_p, _o1, _o2, _d1, _d2, res_p = load_document(pkg_p, Report())
            page_p = load_page(pkg_p, meta_p.pages[0][0], 1, meta_p, res_p,
                               Report(), Options(), meta_p.pages[0][1])
            fills = [v["fill_color"] for v in page_p.vectors]
            check("调色板索引解析为红色", bool(fills) and fills[0] == (255, 0, 0),
                  "实际 %r" % (fills[:1],))
            check("调色板颜色空间被记录",
                  res_p.colors.get("3", {}).get("type") == "Index",
                  repr(res_p.colors.get("3")))
        finally:
            pkg_p.close()
    except Exception as exc:
        check("用例 7 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 8：矢量轮廓的多段闭合与奇偶填充 ----------------------
    print("[8] 矢量轮廓：多段 C 闭合 / 填充不要求显式闭合 / Even-Odd")
    try:
        # 一个 S 后面跟两个 C 段：这类写法在 WPS 导出的公文里很常见
        raw = ("S 10 10 L 40 10 L 40 40 L 10 40 C L 20 20 L 30 20 L 30 30 L 20 30 C")
        subs = path_subpaths(parse_path_data(raw))
        check("识别出 2 段子路径", len(subs) == 2, "实际 %d" % len(subs))
        if len(subs) == 2:
            check("第二段以自身首点收尾（不被拉回原点）",
                  subs[1][0][0] == (20.0, 20.0) and subs[1][0][-1] == (20.0, 20.0),
                  "起点 %r 收尾 %r" % (subs[1][0][0], subs[1][0][-1]))

        po = ('<ofd:PathObject ID="700" Boundary="0 0 210 297" Rule="Even-Odd" '
              'Stroke="false" Fill="true"><ofd:AbbreviatedData>%s'
              '</ofd:AbbreviatedData></ofd:PathObject>' % raw)
        root = parse_xml((_xml_decl() + '<ofd:Page xmlns:ofd="%s"><ofd:Content>'
                          '<ofd:Layer ID="1">%s</ofd:Layer></ofd:Content>'
                          '</ofd:Page>' % (OFD_NS, po)).encode("utf-8"), "selftest")
        vectors = extract_vectors(root, ResourceTable(), Box(0.0, 0.0, 210.0, 297.0),
                                  1, Report())
        check("两条子路径都参与填充（Fill=true）",
              len(vectors) == 2 and all(v["fill"] for v in vectors),
              repr([v["fill"] for v in vectors]))
        check("同一 PathObject 归为同一组",
              len({v["group"] for v in vectors}) == 1)
        check("Rule 透传到渲染层",
              all(v["rule"] == "Even-Odd" for v in vectors))

        import ofd_view as _ov
        drawn = _ov._vector_group_rgba(vectors, 10.0, 210.0, 297.0)
        if drawn is None:
            check("整组可栅格化", False, "返回 None")
        else:
            alpha = drawn[2].split()[3]
            ink = sum(1 for v in alpha.getdata() if v > 0) / 100.0
            # 30x30 外框减 10x10 内孔 = 800 mm²
            check("奇偶填充挖掉内孔（面积≈800mm²）", abs(ink - 800.0) <= 60.0,
                  "实测 %.0f mm²" % ink)
    except Exception as exc:
        check("用例 8 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 用例 9：红头 / 印章着色的识别边界 --------------------------
    print("[9] 红头与印章着色：识别命中与不误伤")
    try:
        import ofd_view as _ov2
        from PIL import Image as _Img, ImageDraw as _ImgDraw

        class _FakePage(object):
            def __init__(self, index, box, vectors):
                self.index = index
                self.box = box
                self.vectors = vectors

        def _glyph(gid, x, y, w=8.0, h=12.0, color=(77, 77, 77)):
            return {"points": [(x, y), (x + w, y), (x + w, y + h), (x, y + h)],
                    "fill": True, "fill_color": color, "stroke": False,
                    "stroke_color": color, "group": gid}

        def _line(gid, color=(77, 77, 77)):
            return {"points": [(28.0, 62.0), (190.0, 62.0),
                               (190.0, 63.0), (28.0, 63.0)],
                    "fill": True, "fill_color": color, "stroke": False,
                    "stroke_color": color, "group": gid}

        a4 = Box(0.0, 0.0, 210.0, 297.0)
        glyphs = [_glyph("g%d" % i, 30.0 + i * 9.0, 38.0) for i in range(6)]
        page_ok = _FakePage(1, a4, glyphs + [_line("ln")])
        keys = _ov2.redhead_groups(page_ok)
        check("灰度红头被识别（6 字 + 1 线）", len(keys) == 7, repr(sorted(keys)))

        colored = [_glyph("c%d" % i, 30.0 + i * 9.0, 38.0, color=(200, 20, 20))
                   for i in range(6)]
        check("已有彩色红头不改", not _ov2.redhead_groups(
            _FakePage(1, a4, colored + [_line("ln2", (200, 20, 20))])))

        low = [_glyph("d%d" % i, 30.0 + i * 9.0, 240.0) for i in range(6)]
        check("页面底部的大字不算红头",
              not _ov2.redhead_groups(_FakePage(1, a4, low + [_line("ln3")])))

        check("第 2 页不做红头识别", not _ov2.redhead_groups(
            _FakePage(2, a4, glyphs + [_line("ln4")])))

        changed = _ov2.apply_redhead(page_ok)
        reds = [it["fill_color"] for it in changed if it["group"] == "g0"]
        grays = [it["fill_color"] for it in page_ok.vectors if it["group"] == "g0"]
        check("着色只改副本，不动页面模型",
              reds == [_ov2.REDHEAD_RED] and grays == [(77, 77, 77)],
              "%r / %r" % (reds, grays))

        seal = _Img.new("L", (300, 300), 255)
        _ImgDraw.Draw(seal).ellipse((24, 24, 276, 276), outline=0, width=16)
        check("灰度圆章被认成印章", _ov2._looks_like_seal(seal, Box(40, 40, 40, 40)))
        check("灰度照片不被认成印章",
              not _ov2._looks_like_seal(seal, Box(40, 40, 130, 90)))
        colored_seal = _Img.new("RGB", (300, 300), (255, 255, 255))
        _ImgDraw.Draw(colored_seal).ellipse((24, 24, 276, 276),
                                           outline=(200, 20, 20), width=16)
        check("彩色印章不再染色",
              not _ov2._looks_like_seal(colored_seal, Box(40, 40, 40, 40)))
        red_seal = _ov2.gray_to_seal(seal, 160.0)
        px = red_seal.getpixel((150, 20))
        check("灰度印章染成朱红且带半透明",
              px[0] > px[1] and px[3] < 255 and px[0] > 0, repr(px))
    except Exception as exc:
        check("用例 9 执行", False, "%s: %s" % (type(exc).__name__, exc))

    # ---- 汇总 -------------------------------------------------------
    failed = [r for r in results if not r[1]]
    print("\n合计 %d 项，通过 %d 项，失败 %d 项"
          % (len(results), len(results) - len(failed), len(failed)))
    if failed:
        print("产物目录（可人工用 Word 打开核对）：%s" % base)
        for name, _ok, detail in failed:
            print("  失败：%s  %s" % (name, detail))
        return 1
    return 0


if __name__ == "__main__":
    # ofd_view 会 import ofd2docx；直接运行本文件时模块名是 __main__，
    # 不注册别名的话会加载出第二份模块，异常类型都对不上。
    sys.modules.setdefault("ofd2docx", sys.modules["__main__"])
    sys.exit(main())
