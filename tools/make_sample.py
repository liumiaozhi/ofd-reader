# -*- coding: utf-8 -*-
"""生成一份用于演示与 CI 的合成 OFD 样本（内容虚构，不含任何真实公文）。

用法：
    python tools/make_sample.py out/demo.ofd [--pages 3]

做两件事：
  1. 给界面截图提供素材（README 里的两张图就是用它拍的）；
  2. 给 GitHub Actions 提供 `--selftest` / `--smoke-gui` 需要的输入文件，
     这样仓库里不必存放任何 .ofd 文件。

复用 ofd2docx 自测里的私有构造器（同一套 XML 写法，保证样本与自测一致）。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ofd2docx as ox          # noqa: E402


def _width(text, size):
    # type: (str, float) -> float
    return sum(ox.est_advance(ch, size) for ch in text)


def _text(oid, x, y, text, size=4.0, font="3"):
    # type: (int, float, float, str, float, str) -> str
    return ox._text_obj(oid, x, y, text, size, font)          # noqa: SLF001


def _center(text, size, y):
    # type: (str, float, float) -> str
    return _text(0, (210.0 - _width(text, size)) / 2.0, y, text, size)


def _page_one(oid):
    # type: (int) -> tuple
    """第一页：标题 + 两节正文 + 页码。返回 (body, 下一个 oid)。"""
    out = []
    out.append(_text(oid, 28.0, 16.0,
                     "示例文档 · 内容为虚构，仅用于界面演示", 3.2))
    oid += 1
    out.append(_center("OFD 阅读转换器", 8.2, 34.0))
    oid += 1
    out.append(_center("多标签阅读 · 连续滚动 · 导出 Word / PDF", 4.6, 48.0))
    oid += 1
    out.append(ox._path_obj(oid, 28.0, 58.0, 182.0, 58.0, 0.8))   # noqa: SLF001
    oid += 1

    out.append(_text(oid, 29.0, 68.0, "一、多标签与连续滚动", 5.0, "6"))
    oid += 1
    body1 = [
        "打开多个文件时，每个文件占一个标签，标签从左到右排列，",
        "与浏览器的习惯一致。每个标签各自记住页码、缩放比例和单双",
        "页设置，切换时互不影响。",
        "所有页面纵向连续排布，滚轮和右侧滑块可以一路滚到底，状态",
        "栏的页码会跟随视口自动更新。双击标签可以把该文档拆成独立",
        "窗口，适合并排对照两份文件。",
    ]
    y = 76.0
    for line in body1:
        out.append(_text(oid, 29.0, y, line, 4.0))
        oid += 1
        y += 6.2

    out.append(_text(oid, 29.0, y + 4.0, "二、导出与还原", 5.0, "6"))
    oid += 1
    y += 12.0
    body2 = [
        "导出 Word 时按阅读顺序重排为流式段落，表格会还原成可编辑",
        "的 Word 真表格；红头横线、版记横线与电子印章通过图层或段",
        "落边框保留。导出 PDF 则是每页一张位图，适合查看与打印。",
        "阅读界面支持鼠标框选文字并复制，也可以按 Ctrl+滚轮缩放，",
        "每次 10% 一档。",
    ]
    for line in body2:
        out.append(_text(oid, 29.0, y, line, 4.0))
        oid += 1
        y += 6.2

    out.append(_text(oid, 29.0, y + 4.0, "三、阅读与交互", 5.0, "6"))
    oid += 1
    y += 12.0
    body3 = [
        "页面支持单页与双页两种视图，工具栏的两个小图标即可切换；",
        "双击页面空白处也能切换。窗口变宽变窄时，页面始终水平居中。",
        "在页面上按住左键拖动可以框选文字，选中的内容以蓝色高亮显",
        "示，按 Ctrl+C 复制；右击空白处可取消选择。",
        "左侧标签较多时，标签栏支持横向滚动，激活的标签会自动滚入",
        "可见区域。",
    ]
    for line in body3:
        out.append(_text(oid, 29.0, y, line, 4.0))
        oid += 1
        y += 6.2

    out.append(_center("— 1 —", 3.6, 282.0))
    oid += 1
    return ("".join(out), oid)


def _page_two(oid):
    # type: (int) -> tuple
    """第二页：三列四行的表格 + 一张示意图 + 页码。"""
    out = []
    out.append(_text(oid, 29.0, 24.0, "四、表格与图片", 5.0, "6"))
    oid += 1

    left, top = 30.0, 36.0
    cols = [left, left + 52.0, left + 104.0, left + 152.0]
    rows = [top + i * 9.0 for i in range(5)]
    for y in rows:
        out.append(ox._path_obj(oid, cols[0], y, cols[-1], y, 0.4))   # noqa: SLF001
        oid += 1
    for x in cols:
        out.append(ox._path_obj(oid, x, rows[0], x, rows[-1], 0.4))   # noqa: SLF001
        oid += 1

    header = ["项目", "说明", "取值"]
    for col, text in enumerate(header):
        out.append(_text(oid, cols[col] + 3.0, rows[0] + 2.2, text, 3.6))
        oid += 1
    rows_data = [
        ("行数", "表格识别出的行数", "4"),
        ("列数", "表格识别出的列数", "3"),
        ("边框", "仅绘制识别到的线", "是"),
    ]
    for index, (a, b, c) in enumerate(rows_data, 1):
        base = rows[index] + 2.2
        out.append(_text(oid, cols[0] + 3.0, base, a, 3.6))
        oid += 1
        out.append(_text(oid, cols[1] + 3.0, base, b, 3.6))
        oid += 1
        out.append(_text(oid, cols[2] + 3.0, base, c, 3.6))
        oid += 1

    out.append(_text(oid, 29.0, 96.0,
                     "下图是一张占位图片，用来检查图片的尺寸与位置：", 4.0))
    oid += 1
    out.append(ox._image_obj(oid, "10", 55.0, 104.0, 100.0, 56.0))    # noqa: SLF001
    oid += 1
    out.append(_text(oid, 29.0, 172.0,
                     "图片按原始尺寸内联，超过正文宽度时等比缩放。", 4.0))
    oid += 1

    out.append(_center("— 2 —", 3.6, 282.0))
    oid += 1
    return ("".join(out), oid)


def _page_three(oid):
    # type: (int) -> tuple
    """第三页：快捷键说明 + 版记样式区块。"""
    out = []
    out.append(_text(oid, 29.0, 24.0, "五、常用快捷键", 5.0, "6"))
    oid += 1
    lines = [
        ("Ctrl+O", "打开文件（可多选）"),
        ("Ctrl+W", "关闭当前标签"),
        ("Ctrl+Tab", "切换到下一个标签"),
        ("Ctrl+0 / Ctrl+1", "适应宽度 / 实际大小"),
        ("Ctrl+S / Ctrl+P", "导出 Word / 导出 PDF"),
    ]
    y = 34.0
    for key, desc in lines:
        out.append(_text(oid, 29.0, y, key, 4.0, "6"))
        oid += 1
        out.append(_text(oid, 82.0, y, desc, 4.0))
        oid += 1
        y += 7.0

    out.append(ox._path_obj(oid, 28.0, 240.0, 182.0, 240.0, 0.6))     # noqa: SLF001
    oid += 1
    out.append(_text(oid, 28.0, 247.0,
                     "内部发送：示例部门、示例岗位。   （共印 0 份）", 3.6))
    oid += 1
    out.append(ox._path_obj(oid, 28.0, 254.0, 182.0, 254.0, 0.6))     # noqa: SLF001
    oid += 1
    out.append(_text(oid, 28.0, 261.0,
                     "示例单位办公室   2026 年 1 月 1 日印发", 3.6))
    oid += 1
    out.append(ox._path_obj(oid, 28.0, 268.0, 182.0, 268.0, 0.6))     # noqa: SLF001
    oid += 1

    out.append(_center("— 3 —", 3.6, 282.0))
    oid += 1
    return ("".join(out), oid)


def build(path, pages=3):
    # type: (str, int) -> str
    oid = 100
    builders = [_page_one, _page_two, _page_three][:max(1, min(pages, 3))]
    bodies = []
    for builder in builders:
        body, oid = builder(oid)
        bodies.append({"body": body})

    image = ox._png_bytes(160, 90, (214, 226, 240))               # noqa: SLF001
    media = [("10", "demo.png", image)]
    ox._write_package(path, bodies, box=(210.0, 297.0),            # noqa: SLF001
                      fonts=[("3", "simsun"), ("6", "simhei")],
                      media=media)
    return path


def main():
    # type: () -> int
    parser = argparse.ArgumentParser(description="生成演示用合成 OFD")
    parser.add_argument("output", help="输出的 .ofd 路径")
    parser.add_argument("--pages", type=int, default=3, choices=(1, 2, 3),
                        help="页数（默认 3）")
    args = parser.parse_args()

    out = os.path.abspath(args.output)
    folder = os.path.dirname(out)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    build(out, args.pages)
    print("已生成 %s（%d 页，%.1f KB）"
          % (out, args.pages, os.path.getsize(out) / 1024.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
