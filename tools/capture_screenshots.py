# -*- coding: utf-8 -*-
"""生成 README 用的界面截图（全部基于合成样本，不含任何真实文档）。

用法：
    python tools/capture_screenshots.py --out docs/screenshots

会现场合成三个演示用 OFD（内容虚构），然后：
  1. 开三个标签、框选一段文字，截主窗口 -> tabs.png
  2. 把当前标签拆成独立窗口、翻到第 2 页（表格 + 图片），截该窗口 -> detached.png

需要能创建窗口的桌面环境（Windows）。CI 里不调用本脚本。
"""

import argparse
import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk                       # noqa: E402

import make_sample                         # noqa: E402
import ofd_reader as R                     # noqa: E402

DEMO_NAMES = ["演示-通知.ofd", "演示-报告.ofd", "演示-纪要.ofd"]


def pump(root, seconds):
    # type: (object, float) -> None
    end = time.time() + seconds
    while time.time() < end:
        root.update()
        time.sleep(0.02)


def grab(window, path):
    # type: (object, str) -> None
    from PIL import ImageGrab
    window.lift()
    try:
        window.attributes("-topmost", True)
    except Exception:
        pass
    window.update()
    time.sleep(0.5)
    window.update()
    x, y = window.winfo_rootx(), window.winfo_rooty()
    box = (x, y, x + window.winfo_width(), y + window.winfo_height())
    ImageGrab.grab(bbox=box).save(path)
    try:
        window.attributes("-topmost", False)
    except Exception:
        pass
    print("  已保存 %s（%dx%d）"
          % (path, window.winfo_width(), window.winfo_height()))


def main():
    # type: () -> int
    parser = argparse.ArgumentParser(description="生成 README 界面截图")
    parser.add_argument("--out", default="docs/screenshots",
                        help="截图输出目录（默认 docs/screenshots）")
    args = parser.parse_args()

    out_dir = os.path.abspath(args.out)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)

    try:
        import ctypes
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

    work = tempfile.mkdtemp(prefix="ofd_shots_")
    paths = []
    for name in DEMO_NAMES:
        path = os.path.join(work, name)
        make_sample.build(path, pages=3)
        paths.append(path)
    print("已合成 %d 个演示文档" % len(paths))

    root = tk.Tk()
    root.geometry("+60+50")
    app = R.ReaderApp(root, paths=paths)
    pump(root, 2.6)

    # 最后一个标签是当前标签，回到第一个标签再框选，标题顺序更自然
    app.activate_view(app._views[0])
    pump(root, 0.8)
    app.fit_width()          # 去掉横向滚动条，画面更干净
    pump(root, 0.8)

    view = app.view
    slot = view._slots.get(1)
    model = view._models.get(1)
    if slot is not None and model is not None and len(model.lines) > 4:
        line = model.lines[4]
        ppm = view._ppm()
        x0 = slot["x"] + min(g.x for g in line.glyphs) * ppm - 2
        x1 = slot["x"] + max(g.x + g.adv for g in line.glyphs) * ppm + 2
        y0 = slot["y"] + line.y * ppm - line.size_mm * ppm
        y1 = slot["y"] + line.y * ppm + line.size_mm * ppm * 0.35
        view._paint_selection((x0, y0, x1, y1))
        root.update()
        pump(root, 0.5)

    grab(root, os.path.join(out_dir, "tabs.png"))

    # 拆成独立窗口，翻到第 2 页（有表格和图片，画面更直观）
    second = app.detach_view(app.view)
    pump(root, 1.4)
    if second is not None:
        second.root.geometry("1040x900+120+40")
        pump(root, 1.0)
        second.view.fit_width()
        second.view._goto_page(2 if second.view.page_count > 1 else 1)
        pump(root, 1.2)
        grab(second.root, os.path.join(out_dir, "detached.png"))

    app.on_close()
    shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
