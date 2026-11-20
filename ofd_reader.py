# -*- coding: utf-8 -*-
"""ofd_reader.py —— OFD 阅读转换器（Windows 桌面程序）

功能：
  · 多文档：一次打开多个 .ofd，浏览器式标签从左到右排列，可拆成独立窗口
  · 连续滚动阅读 .ofd（像 Word 一样铺开所有页，滚轮/右侧滑块浏览）
  · 单页 / 双页两种视图，可随时切换，各标签互不影响
  · 鼠标框选页面文字并复制（Ctrl+C）
  · 缩放步进 10%，另有适应宽度、实际大小
  · 导出为 Word（可编辑流式排版）、PDF（每页一张位图）
  · 一键注册到资源管理器右键「打开方式」

命令行：
  ofd阅读转换器.exe a.ofd b.ofd    一次打开多个文件，各占一个标签
  ofd阅读转换器.exe --register     注册右键「打开方式」
  ofd阅读转换器.exe --unregister   移除注册
  ofd阅读转换器.exe --selftest <文件.ofd> [--report out.txt]
                                   无界面自检（打包后验证依赖是否完整）
  ofd阅读转换器.exe --smoke-gui <文件.ofd> [--report out.txt]
                                   建窗口跑一遍多标签/多窗口流程后自动退出

依赖：python-docx、lxml、Pillow（见 requirements.txt）
"""

import argparse
import ctypes
import os
import sys
import threading
import traceback
from collections import OrderedDict

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ofd2docx as ox
import ofd_view as vw

try:
    from PIL import Image, ImageDraw, ImageTk
except Exception:                                  # pragma: no cover
    Image = None
    ImageDraw = None
    ImageTk = None

APP_NAME = "OFD 阅读转换器"
APP_VERSION = "1.2.2"
PROGID = "OFDReader.Document"
APP_ID = "Local.OFDReader.1"

BASE_DPI = 96.0          # 缩放 100% 时的屏幕像素密度
ZOOM_STEP = 0.10         # 每次放大/缩小 10%
ZOOM_MIN = 0.20
ZOOM_MAX = 4.00
PAGE_GAP = 14            # 页与页的纵向间距（像素）
PAGE_GAP_H = 12          # 双页模式下左右间距
CANVAS_MARGIN = 18       # 画布四周留白

MODE_SINGLE = "single"
MODE_SPREAD = "spread"

# 标签栏配色（浅色，和 vista 主题一致）
TAB_BAR_BG = "#e9edf0"
TAB_BG = "#dbe1e6"
TAB_BG_HOVER = "#eef3f7"
TAB_ACTIVE_BG = "#ffffff"
TAB_FG = "#3d4b57"
TAB_ACTIVE_FG = "#101a22"
TAB_BAR_H = 31


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

def resource_path(rel):
    # type: (str) -> str
    """打包后从 _MEIPASS 取资源，源码运行时从脚本目录取。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_exe_path():
    # type: () -> str
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return os.path.abspath(sys.executable)


def attach_parent_console():
    # type: () -> None
    """窗口程序在命令行里运行时，把 stdout 接回父控制台便于看输出。"""
    try:
        if ctypes.windll.kernel32.AttachConsole(-1):
            # 按控制台当前代码页输出，否则中文在 cmd 里会是乱码
            code_page = ctypes.windll.kernel32.GetConsoleOutputCP()
            encoding = "utf-8"
            if code_page and code_page != 65001:
                candidate = "cp%d" % code_page
                try:
                    u"\u4e2d".encode(candidate)
                    encoding = candidate
                except Exception:
                    encoding = "gbk"
            sys.stdout = open("CONOUT$", "w", buffering=1,            # noqa: SIM115
                              encoding=encoding, errors="replace")
            sys.stderr = sys.stdout
    except Exception:
        pass


def set_app_user_model_id():
    # type: () -> None
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 文件关联（HKCU，写当前用户，不需要管理员）
# ---------------------------------------------------------------------------

def _winreg():
    try:
        import winreg
    except ImportError:                            # pragma: no cover
        return None
    return winreg


def _set_value(root, key_path, name, value):
    # type: (object, str, str, str) -> None
    winreg = _winreg()
    with winreg.CreateKeyEx(root, key_path, 0, winreg.KEY_WRITE) as handle:
        winreg.SetValueEx(handle, name, 0, winreg.REG_SZ, value)


def _delete_value(root, key_path, name):
    # type: (object, str, str) -> bool
    winreg = _winreg()
    try:
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_SET_VALUE) as handle:
            winreg.DeleteValue(handle, name)
            return True
    except OSError:
        return False


def _delete_tree(root, key_path):
    # type: (object, str) -> bool
    winreg = _winreg()
    try:
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_ALL_ACCESS) as handle:
            while True:
                try:
                    sub = winreg.EnumKey(handle, 0)
                except OSError:
                    break
                _delete_tree(root, key_path + "\\" + sub)
    except OSError:
        return False
    try:
        winreg.DeleteKey(root, key_path)
        return True
    except OSError:
        return False


def _notify_shell():
    # type: () -> None
    try:
        SHCNE_ASSOCCHANGED = 0x08000000
        SHCNF_IDLIST = 0x0000
        ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST,
                                             None, None)
    except Exception:
        pass


def register_association(exe_path, make_default=False):
    # type: (str, bool) -> None
    """把本程序注册到 .ofd 的「打开方式」。全部写 HKCU，不需要管理员权限。"""
    winreg = _winreg()
    if winreg is None:
        raise ox.OfdUnsupportedError("当前系统缺少 winreg 模块，无法注册文件关联")
    if not os.path.exists(exe_path):
        raise ox.OfdError("程序文件不存在，无法注册", part=exe_path)
    HKCU = winreg.HKEY_CURRENT_USER
    classes = r"Software\Classes"
    exe_name = os.path.basename(exe_path)
    command = '"%s" "%%1"' % exe_path

    app_key = classes + "\\Applications\\" + exe_name
    _set_value(HKCU, app_key, "FriendlyAppName", APP_NAME)
    _set_value(HKCU, app_key + r"\DefaultIcon", "", '"%s",0' % exe_path)
    _set_value(HKCU, app_key + r"\shell\open", "FriendlyAppName", APP_NAME)
    _set_value(HKCU, app_key + r"\shell\open\command", "", command)
    _set_value(HKCU, app_key + r"\SupportedTypes", ".ofd", "")

    prog_key = classes + "\\" + PROGID
    _set_value(HKCU, prog_key, "", "OFD 文档")
    _set_value(HKCU, prog_key + r"\DefaultIcon", "", '"%s",0' % exe_path)
    _set_value(HKCU, prog_key + r"\shell\open", "FriendlyAppName", APP_NAME)
    _set_value(HKCU, prog_key + r"\shell\open\command", "", command)

    _set_value(HKCU, classes + r"\.ofd\OpenWithProgids", PROGID, "")
    _set_value(HKCU,
               r"Software\Microsoft\Windows\CurrentVersion\Explorer"
               r"\FileExts\.ofd\OpenWithList", "a", exe_name)

    if make_default:
        _set_value(HKCU, classes + r"\.ofd", "", PROGID)

    _notify_shell()


def unregister_association(exe_path=None):
    # type: (str) -> int
    """移除本程序写入的关联，返回移除的项数。"""
    winreg = _winreg()
    if winreg is None:
        return 0
    HKCU = winreg.HKEY_CURRENT_USER
    classes = r"Software\Classes"
    exe_name = os.path.basename(exe_path or app_exe_path())
    removed = 0

    if _delete_tree(HKCU, classes + "\\Applications\\" + exe_name):
        removed += 1
    if _delete_value(HKCU, classes + r"\.ofd\OpenWithProgids", PROGID):
        removed += 1
    if _delete_tree(HKCU, classes + "\\" + PROGID):
        removed += 1

    try:
        with winreg.OpenKey(HKCU, classes + r"\.ofd", 0, winreg.KEY_READ) as key:
            value, _kind = winreg.QueryValueEx(key, "")
        if str(value) == PROGID:
            _set_value(HKCU, classes + r"\.ofd", "", "")
            removed += 1
    except OSError:
        pass

    list_key = (r"Software\Microsoft\Windows\CurrentVersion\Explorer"
                r"FileExts\.ofd\OpenWithList")
    for slot in ("a", "b", "c", "d"):
        try:
            with winreg.OpenKey(HKCU, list_key, 0, winreg.KEY_READ) as key:
                value, _kind = winreg.QueryValueEx(key, slot)
        except OSError:
            continue
        if str(value).lower() == exe_name.lower():
            _delete_value(HKCU, list_key, slot)
            removed += 1

    _notify_shell()
    return removed


def association_installed(exe_path=None):
    # type: (str) -> bool
    winreg = _winreg()
    if winreg is None:
        return False
    exe_name = os.path.basename(exe_path or app_exe_path())
    key = (r"Software\Classes\Applications" + "\\" + exe_name +
           r"\shell\open\command")
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0,
                            winreg.KEY_READ) as handle:
            value, _kind = winreg.QueryValueEx(handle, "")
        return bool(value)
    except OSError:
        return False


# ---------------------------------------------------------------------------
# 视图切换按钮的小图标（Pillow 画，避免依赖字体是否含对应符号）
# ---------------------------------------------------------------------------

def make_view_icons():
    # type: () -> tuple
    """返回 (单页图标, 双页图标)，失败时返回 (None, None) 由调用方退回文字。"""
    if Image is None or ImageDraw is None or ImageTk is None:
        return (None, None)
    size = 22
    result = []        # type: list
    try:
        for pages in (1, 2):
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            page_w, page_h = 8, 14
            gap = 2
            total = pages * page_w + (pages - 1) * gap
            left = (size - total) // 2
            top = (size - page_h) // 2
            for idx in range(pages):
                x = left + idx * (page_w + gap)
                draw.rectangle([x, top, x + page_w, top + page_h],
                               fill=(255, 255, 255), outline=(70, 70, 70))
                for line in range(3):
                    y = top + 3 + line * 3
                    draw.line([x + 2, y, x + page_w - 2, y],
                              fill=(150, 150, 150))
            result.append(ImageTk.PhotoImage(img))
    except Exception:
        return (None, None)
    return (result[0], result[1])


# ---------------------------------------------------------------------------
# 滚轮分发
#
# Tk 的 bind_all 是"全应用级"的，而且不带 add="+" 时是**替换**而不是追加：
# 每开一个新窗口就调一次 bind_all，会把上一个窗口的回调顶掉，结果是先开的
# 那个窗口滚轮失灵。所以这里只安装一次分发器，再按鼠标位置找到所属窗口。
# ---------------------------------------------------------------------------

_WINDOWS = []            # type: list  当前活着的窗口，按创建顺序
_WHEEL_INSTALLED = False


def _install_wheel_dispatch(root):
    # type: (object) -> None
    global _WHEEL_INSTALLED
    if _WHEEL_INSTALLED:
        return
    _WHEEL_INSTALLED = True
    root.bind_all("<MouseWheel>", _dispatch_wheel)
    root.bind_all("<Control-MouseWheel>", _dispatch_ctrl_wheel)


def _dispatch_wheel(event):
    # type: (object) -> object
    for window in list(_WINDOWS):
        if window.scroll_wheel(event):
            return "break"
    return None


def _dispatch_ctrl_wheel(event):
    # type: (object) -> object
    for window in list(_WINDOWS):
        if window.zoom_wheel(event):
            return "break"
    return None


# ---------------------------------------------------------------------------
# 单文档视图：一个标签页里的画布与全部阅读逻辑
# ---------------------------------------------------------------------------

class DocView(object):
    """一份文档的阅读视图：画布、分页布局、渲染、文字选中、导出。

    每个 DocView 对应一个标签。页码、缩放、视图模式、滚动位置、选中内容
    全部存在实例里，因此多个标签可以各看各的，互不干扰。
    """

    def __init__(self, window, path, page_index=1, zoom=None,
                 mode=MODE_SINGLE, fit_mode=True):
        # type: (ReaderApp, str, int, float, str, bool) -> None
        self.win = window
        self.root = window.root
        self.frame = ttk.Frame(window.content)

        self.doc = None
        self.resolver = None
        self.path = ""
        self.page_count = 0
        self.page_index = max(1, int(page_index or 1))
        self.mode = mode or MODE_SINGLE
        self.zoom = max(ZOOM_MIN, min(ZOOM_MAX, float(zoom or 1.0)))
        self.fit_mode = bool(fit_mode) and zoom is None
        self.info_text = ""
        self.status_text = ""

        self._assumed_box = ox.Box(0.0, 0.0, 210.0, 297.0)
        self._boxes = {}             # type: dict  每页真实尺寸（mm）
        self._slots = {}             # type: dict  page -> {x,y,w,h}
        self._rows = []              # type: list  每行包含的页码
        self._slot_items = {}        # type: dict  page -> 画布图片项 id
        self._cache = OrderedDict()  # type: OrderedDict  page -> PhotoImage
        self._models = OrderedDict()  # type: OrderedDict page -> PageModel
        self._content_w = 1.0
        self._content_h = 1.0
        self._visible_after = None
        self._relayout_after = None
        self._need_relayout = False

        self._sel_start = None
        self._dragging = False
        self._sel_hits = []
        self._last_report = None

        self._build_ui()
        if path:
            self.open(path, page_index=page_index, zoom=zoom)

    # -- 界面 -------------------------------------------------------------
    def _build_ui(self):
        # type: () -> None
        body = self.frame
        self.canvas = tk.Canvas(body, bg="#767676", highlightthickness=0,
                                yscrollincrement=1, xscrollincrement=1)
        self.vbar = ttk.Scrollbar(body, orient="vertical",
                                  command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(body, orient="horizontal",
                                  command=self.canvas.xview)
        self.canvas.configure(
            yscrollcommand=lambda a, b: self._on_yscroll(a, b, self.vbar),
            xscrollcommand=self.hbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)

        self._selection_menu = tk.Menu(self.root, tearoff=0)
        self._selection_menu.add_command(label="复制\tCtrl+C",
                                         command=self.copy_selection)
        self._selection_menu.add_separator()
        self._selection_menu.add_command(
            label="在新窗口中打开本文档",
            command=lambda: self.win.detach_view(self))
        self._selection_menu.add_command(label="取消选择",
                                         command=self.clear_selection)

    # -- 打开 / 关闭 -------------------------------------------------------
    def open(self, path, page_index=1, zoom=None):
        # type: (str, int, float) -> bool
        if self.win.busy:
            return False
        path = os.path.abspath(path)
        self.win.set_status("正在解析 %s …" % os.path.basename(path))
        self.root.update_idletasks()
        try:
            doc = ox.OfdDocument(path)
        except ox.OfdError as exc:
            messagebox.showerror(APP_NAME, exc.user_message(), parent=self.root)
            self.win.set_status("打开失败：%s" % exc.message)
            return False
        except Exception as exc:                   # pragma: no cover
            messagebox.showerror(
                APP_NAME, "打开失败：%s: %s" % (type(exc).__name__, exc),
                parent=self.root)
            self.win.set_status("打开失败：%s" % exc)
            return False

        self._release_doc()
        self.doc = doc
        self.path = path
        self.resolver = vw.FontResolver(doc.res)
        self.page_count = doc.page_count
        self._assumed_box = doc.meta.page_box
        self._boxes = {}
        self._cache.clear()
        self._models.clear()
        self._sel_hits = []
        self._last_report = None
        self.info_text = ""

        if zoom is not None:
            self.zoom = max(ZOOM_MIN, min(ZOOM_MAX, float(zoom)))
            self.fit_mode = False
        self.page_index = max(1, min(int(page_index or 1),
                                     max(1, self.page_count)))

        if self.page_count == 0:
            messagebox.showwarning(APP_NAME, "该文档没有任何页面。",
                                   parent=self.root)
            self.info_text = ""
            return True

        self._relayout(keep_position=False)
        self._goto_page(self.page_index)
        self.win.set_status("已打开 %s（%d 页）"
                            % (os.path.basename(path), self.page_count))
        return True

    def dispose(self):
        # type: () -> None
        """关闭文档并销毁界面，退出前调用。"""
        self._release_doc()
        try:
            self.frame.destroy()
        except Exception:
            pass

    def _release_doc(self):
        # type: () -> None
        if self.doc is not None:
            try:
                self.doc.close()
            except Exception:
                pass
        self.doc = None
        self.resolver = None
        self._cache.clear()
        self._models.clear()
        self._slots = {}
        self._slot_items = {}
        self._rows = []

    def snapshot(self):
        # type: () -> dict
        """拆窗/合并时用来把阅读位置带过去。"""
        return {"page_index": self.page_index, "zoom": self.zoom,
                "mode": self.mode, "fit_mode": self.fit_mode}

    def on_show(self):
        # type: () -> None
        """标签被激活时调用：按真实画布宽度重新定位。"""
        if self.doc is None:
            return
        self.frame.update_idletasks()
        if self.fit_mode:
            zoom = self._fit_zoom()
            if abs(zoom - self.zoom) > 0.002:
                self.zoom = zoom
                self._update_zoom_label()
                self._cache.clear()
        self._relayout(keep_position=False)
        self._goto_page(self.page_index)
        self._schedule_visible()

    def on_hide(self):
        # type: () -> None
        """切到别的标签时释放位图，多开几个文档也不会把内存吃光。"""
        for item in list(self._slot_items.values()):
            try:
                self.canvas.delete(item)
            except Exception:
                pass
        self._slot_items.clear()
        self._cache.clear()

    # -- 布局 -------------------------------------------------------------
    def _ppm(self):
        # type: () -> float
        return BASE_DPI / 25.4 * self.zoom

    def _box_of(self, page):
        # type: (int) -> object
        return self._boxes.get(page, self._assumed_box)

    def _page_px(self, page):
        # type: (int) -> tuple
        ppm = self._ppm()
        box = self._box_of(page)
        return (box.w * ppm, box.h * ppm)

    def _build_rows(self):
        # type: () -> list
        pages = list(range(1, self.page_count + 1))
        if self.mode == MODE_SPREAD and len(pages) > 1:
            return [pages[i:i + 2] for i in range(0, len(pages), 2)]
        return [[p] for p in pages]

    def _relayout(self, keep_position=True):
        # type: (bool) -> None
        if self.doc is None or self.page_count <= 0:
            return
        anchor = self._capture_anchor() if keep_position else None

        self._relayout_after = None
        self.canvas.delete("all")
        self._slot_items = {}
        self._rows = self._build_rows()

        canvas_w = max(1, self.canvas.winfo_width())
        canvas_h = max(1, self.canvas.winfo_height())

        row_widths = []            # type: list
        for row in self._rows:
            width = sum(self._page_px(p)[0] for p in row)
            width += PAGE_GAP_H * (len(row) - 1)
            row_widths.append(width)
        max_row_w = max(row_widths) if row_widths else 0.0

        # 内容宽度取"画布宽度"与"最宽行"的较大者，窄内容也会水平居中
        content_w = max(canvas_w, max_row_w + 2 * CANVAS_MARGIN)
        self._content_w = content_w

        slots = {}                 # type: dict
        y = CANVAS_MARGIN
        for row, row_w in zip(self._rows, row_widths):
            row_h = max(self._page_px(p)[1] for p in row)
            x = (content_w - row_w) / 2.0
            for page in row:
                w, h = self._page_px(page)
                slots[page] = {"page": page, "x": x, "y": y, "w": w, "h": h}
                x += w + PAGE_GAP_H
            y += row_h + PAGE_GAP
        content_h = max(canvas_h, y - PAGE_GAP + CANVAS_MARGIN)
        self._content_h = content_h

        self._slots = slots
        for page in sorted(slots):
            self._create_placeholder(slots[page])
        self.canvas.config(scrollregion=(0, 0, content_w, content_h))
        # 不需要横向滚动时把滚动条收起来，别留一条空槽
        try:
            if content_w <= canvas_w + 0.5:
                self.hbar.grid_remove()
            else:
                self.hbar.grid()
        except Exception:
            pass

        if anchor is not None:
            self._restore_anchor(anchor)
        self._update_zoom_label()
        self._update_current_page(self.canvas.canvasy(0))
        self._schedule_visible()

    def _create_placeholder(self, slot):
        # type: (dict) -> None
        x, y = slot["x"], slot["y"]
        w, h = slot["w"], slot["h"]
        self.canvas.create_rectangle(x, y, x + w, y + h, fill="#ffffff",
                                     outline="#c9c9c9", tags=("ph",))
        self.canvas.create_text(x + w / 2.0, y + h / 2.0,
                                text="第 %d 页" % slot["page"], fill="#b0b0b0",
                                font=("Microsoft YaHei UI", 11), tags=("ph",))

    def _capture_anchor(self):
        # type: () -> object
        if not self._slots:
            return None
        top = self.canvas.canvasy(0)
        best = None
        for page, slot in self._slots.items():
            if slot["y"] + slot["h"] > top:
                if best is None or slot["y"] < self._slots[best]["y"]:
                    best = page
        if best is None:
            return None
        slot = self._slots[best]
        ratio = (top - slot["y"]) / max(1.0, slot["h"])
        return (best, max(0.0, min(1.0, ratio)))

    def _restore_anchor(self, anchor):
        # type: (object) -> None
        page, ratio = anchor
        slot = self._slots.get(page)
        if slot is None:
            return
        target = slot["y"] + ratio * slot["h"]
        self.canvas.yview_moveto(max(0.0, target / max(1.0, self._content_h)))

    def _schedule_relayout(self):
        # type: () -> None
        if self._relayout_after is not None:
            try:
                self.root.after_cancel(self._relayout_after)
            except Exception:
                pass
        self._relayout_after = self.root.after(120, self._do_relayout)

    def _do_relayout(self):
        # type: () -> None
        self._relayout_after = None
        self._relayout(keep_position=True)

    # -- 渲染 -------------------------------------------------------------
    def _schedule_visible(self):
        # type: () -> None
        if self._visible_after is not None:
            try:
                self.root.after_cancel(self._visible_after)
            except Exception:
                pass
        self._visible_after = self.root.after(25, self._render_visible)

    def _render_visible(self):
        # type: () -> None
        self._visible_after = None
        if self.doc is None or not self._slots:
            return
        if self.win.view is not self:
            return                  # 隐藏的标签不渲染，省内存也省 CPU
        top = self.canvas.canvasy(0)
        height = max(1, self.canvas.winfo_height())
        lo = top - height * 0.6
        hi = top + height * 1.6
        todo = []                  # type: list
        for page, slot in self._slots.items():
            if slot["y"] + slot["h"] < lo or slot["y"] > hi:
                continue
            if self._slot_items.get(page) is None:
                todo.append(page)
        for page in sorted(todo):
            self._render_page_into(page)
        # 滚动后页码要跟着走，否则状态栏一直停在上一次跳转的页
        self._update_current_page(self.canvas.canvasy(0))
        if self._need_relayout:
            self._need_relayout = False
            self._schedule_relayout()

    def _get_model(self, page):
        # type: (int) -> object
        model = self._models.get(page)
        if model is not None:
            self._models.move_to_end(page)
            return model
        model = self.doc.get_page(page)
        self._models[page] = model
        while len(self._models) > 10:
            self._models.popitem(last=False)
        return model

    def _cache_limit(self):
        # type: () -> int
        if self.zoom >= 2.5:
            return 3
        if self.zoom >= 1.5:
            return 6
        return 12

    def _render_page_into(self, page):
        # type: (int) -> None
        slot = self._slots.get(page)
        if slot is None:
            return
        try:
            model = self._get_model(page)
            image = vw.render_page(model, self.resolver, dpi=BASE_DPI,
                                   scale=self.zoom)
        except Exception as exc:                   # pragma: no cover
            self._set_status("第 %d 页渲染失败：%s" % (page, exc))
            return
        if ImageTk is None:
            return
        try:
            photo = ImageTk.PhotoImage(image)
        except Exception as exc:                   # pragma: no cover
            self._set_status("第 %d 页显示失败：%s" % (page, exc))
            return

        # 实际页面尺寸与假定尺寸不一致时，记录并重排一次
        assumed = self._assumed_box
        if (abs(model.box.w - assumed.w) > 0.5 or
                abs(model.box.h - assumed.h) > 0.5):
            if page not in self._boxes:
                self._boxes[page] = model.box
                self._need_relayout = True

        old = self._slot_items.pop(page, None)
        if old is not None:
            self.canvas.delete(old)
        item = self.canvas.create_image(slot["x"], slot["y"], anchor="nw",
                                        image=photo, tags=("page",))
        self._slot_items[page] = item
        self.canvas.tag_raise("hl")

        self._cache[page] = photo
        self._cache.move_to_end(page)
        limit = self._cache_limit()
        while len(self._cache) > limit:
            victim, kept = self._cache.popitem(last=False)
            if victim == page:
                self._cache[victim] = kept
                break
            stale = self._slot_items.pop(victim, None)
            if stale is not None:
                self.canvas.delete(stale)
        self._update_info()

    def _update_info(self):
        # type: () -> None
        if self.doc is None or not self._slots:
            self.info_text = ""
            self.win.refresh_info()
            return
        box = self._box_of(self.page_index)
        warns = [w for w in self.doc.report.warnings if w.level == "warn"]
        suffix = (" ｜ 告警 %d 条" % len(warns)) if warns else ""
        self.info_text = ("%.1f × %.1f mm ｜ 缩放 %d%% ｜ 第 %d/%d 页%s"
                          % (box.w, box.h, int(round(self.zoom * 100)),
                             self.page_index, self.page_count, suffix))
        self.win.refresh_info()

    # -- 滚动 -------------------------------------------------------------
    def _on_yscroll(self, first, last, vbar):
        # type: (str, str, object) -> None
        vbar.set(first, last)
        self._schedule_visible()

    def _update_current_page(self, top):
        # type: (float) -> None
        if not self._slots:
            return
        bottom = top + max(1, self.canvas.winfo_height())
        # 取"与视口重叠面积最大"的页作为当前页。用面积而不是"顶部第一页"，
        # 是因为跳到最后一页时滚动会被底部边界夹住，顶部可能还压着上一页。
        current = None
        best = -1.0
        for page, slot in self._slots.items():
            overlap = min(bottom, slot["y"] + slot["h"]) - max(top, slot["y"])
            if overlap > best:
                best = overlap
                current = page
        if current is None:
            return
        self.page_index = current
        self.win.sync_page_var(self)
        self._update_info()

    # -- 翻页 -------------------------------------------------------------
    def _row_index_of(self, page):
        # type: (int) -> int
        for index, row in enumerate(self._rows):
            if page in row:
                return index
        return 0

    def _goto_page(self, page):
        # type: (int) -> None
        slot = self._slots.get(page)
        if slot is None:
            return
        target = max(0.0, slot["y"] - CANVAS_MARGIN * 0.5)
        self.canvas.yview_moveto(target / max(1.0, self._content_h))
        self.page_index = page
        self.win.sync_page_var(self)
        self._update_info()
        self._schedule_visible()

    def prev_page(self):
        # type: () -> None
        if self.doc is None or self.win.busy:
            return
        row = self._row_index_of(self.page_index)
        if row > 0:
            self._goto_page(self._rows[row - 1][0])

    def next_page(self):
        # type: () -> None
        if self.doc is None or self.win.busy:
            return
        row = self._row_index_of(self.page_index)
        if row + 1 < len(self._rows):
            self._goto_page(self._rows[row + 1][0])

    def go_to_page_text(self, text):
        # type: (str) -> None
        if self.doc is None:
            return
        value = text.strip().split("/")[0].strip()
        try:
            target = int(value)
        except ValueError:
            self.win.sync_page_var(self, force=True)
            return
        target = max(1, min(self.page_count, target))
        self._goto_page(target)

    def _on_double_click(self, event):
        # type: (object) -> None
        """双击页面切换单页/双页。"""
        if self.doc is None:
            return
        self.set_mode(MODE_SPREAD if self.mode == MODE_SINGLE else MODE_SINGLE)

    # -- 视图模式 ---------------------------------------------------------
    def set_mode(self, mode):
        # type: (str) -> None
        self.mode = mode
        self.win.sync_mode_var(self)
        if self.doc is None:
            return
        anchor = self.page_index
        self._clear_selection_internal()
        if self.fit_mode:
            self.zoom = self._fit_zoom()
        self._cache.clear()
        self._relayout(keep_position=False)
        self._goto_page(anchor)
        self._set_status("已切换为%s"
                         % ("双页" if mode == MODE_SPREAD else "单页"))

    # -- 缩放 -------------------------------------------------------------
    def _update_zoom_label(self):
        # type: () -> None
        self.win.sync_zoom_var(self)

    def _fit_zoom(self):
        # type: () -> float
        canvas_w = self.canvas.winfo_width()
        if canvas_w < 60:                  # 还没真正布局，按窗口宽度估算
            canvas_w = max(600, self.root.winfo_width() - 40)
        usable = canvas_w - 2 * CANVAS_MARGIN - 8
        box = self._assumed_box
        natural = max(1.0, box.w / 25.4 * BASE_DPI)
        if self.mode == MODE_SPREAD and self.page_count > 1:
            natural = natural * 2 + PAGE_GAP_H
        return max(ZOOM_MIN, min(ZOOM_MAX, usable / natural))

    def _apply_zoom(self, zoom):
        # type: (float) -> None
        zoom = round(max(ZOOM_MIN, min(ZOOM_MAX, zoom)), 3)
        if abs(zoom - self.zoom) < 1e-4:
            return
        self.zoom = zoom
        self._update_zoom_label()
        self._clear_selection_internal()
        self._cache.clear()
        self._relayout(keep_position=True)

    def zoom_in(self):
        # type: () -> None
        if self.doc is None:
            return
        self.fit_mode = False
        self._apply_zoom(self.zoom + ZOOM_STEP)

    def zoom_out(self):
        # type: () -> None
        if self.doc is None:
            return
        self.fit_mode = False
        self._apply_zoom(self.zoom - ZOOM_STEP)

    def fit_width(self):
        # type: () -> None
        if self.doc is None:
            return
        self.fit_mode = True
        self._apply_zoom(self._fit_zoom())

    def actual_size(self):
        # type: () -> None
        if self.doc is None:
            return
        self.fit_mode = False
        self._apply_zoom(1.0)

    def set_zoom_percent(self, percent):
        # type: (float) -> None
        if self.doc is None:
            return
        self.fit_mode = False
        self._apply_zoom(percent / 100.0)

    def _on_canvas_resize(self, _event=None):
        # type: (object) -> None
        if self.doc is None or self.win.view is not self:
            return
        if self.fit_mode:
            zoom = self._fit_zoom()
            if abs(zoom - self.zoom) > 0.002:
                self.zoom = zoom
                self._update_zoom_label()
                self._cache.clear()
                self._relayout(keep_position=True)
                return
        # 宽度变化时重排，保证页面始终水平居中
        self._schedule_relayout()

    # -- 文字选中 ---------------------------------------------------------
    def _on_press(self, event):
        # type: (object) -> None
        if self.doc is None:
            return
        self.canvas.focus_set()
        self._sel_start = (self.canvas.canvasx(event.x),
                           self.canvas.canvasy(event.y))
        self._dragging = False

    def _on_drag(self, event):
        # type: (object) -> None
        if self.doc is None or self._sel_start is None:
            return
        x0, y0 = self._sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        if abs(x1 - x0) < 3 and abs(y1 - y0) < 3:
            return
        self._dragging = True
        rect = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
        self._paint_selection(rect)

    def _on_release(self, event):
        # type: (object) -> None
        was_dragging = self._dragging
        self._dragging = False
        self._sel_start = None
        if not was_dragging:
            self.clear_selection()

    def _on_right_click(self, event):
        # type: (object) -> None
        if self.doc is None:
            return
        has = bool(self._sel_hits)
        self._selection_menu.entryconfig(
            0, state="normal" if has else "disabled")
        try:
            self._selection_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._selection_menu.grab_release()

    def clear_selection(self, _event=None):
        # type: (object) -> None
        self._clear_selection_internal()
        if self.doc is not None:
            self._set_status("就绪")

    def _clear_selection_internal(self):
        # type: () -> None
        self._sel_hits = []
        try:
            self.canvas.delete("hl")
        except Exception:
            pass

    def _paint_selection(self, rect):
        # type: (tuple) -> None
        x0, y0, x1, y1 = rect
        hits = []                  # type: list
        ppm = self._ppm()
        for page, slot in self._slots.items():
            if slot["y"] > y1 or slot["y"] + slot["h"] < y0:
                continue
            model = self._models.get(page)
            if model is None:
                continue
            for line_index, line in enumerate(model.lines):
                for gi, g in enumerate(line.glyphs):
                    gx0 = slot["x"] + g.x * ppm
                    gy_base = slot["y"] + g.y * ppm
                    gx1 = gx0 + max(g.adv, 0.1) * ppm
                    gy0 = gy_base - 0.88 * g.size_mm * ppm
                    gy1 = gy_base + 0.22 * g.size_mm * ppm
                    if gx0 < x1 and gx1 > x0 and gy0 < y1 and gy1 > y0:
                        hits.append({
                            "page": page, "line": line_index, "idx": gi,
                            "x0": gx0, "y0": gy0, "x1": gx1, "y1": gy1,
                            "ch": g.ch, "x_mm": g.x, "adv_mm": g.adv,
                            "size_mm": g.size_mm,
                        })
        self._sel_hits = hits
        self.canvas.delete("hl")
        if not hits:
            self._set_status("未选中文字")
            return
        # 同一行内连续的字形合并成一个矩形，避免几百个小方块拖慢重绘
        groups = OrderedDict()
        for hit in hits:
            groups.setdefault((hit["page"], hit["line"]), []).append(hit)
        for _key, items in groups.items():
            run_x0 = items[0]["x0"]
            run_x1 = items[0]["x1"]
            run_y0 = min(h["y0"] for h in items)
            run_y1 = max(h["y1"] for h in items)
            prev_idx = items[0]["idx"]
            for hit in items[1:]:
                if hit["idx"] == prev_idx + 1:
                    run_x1 = max(run_x1, hit["x1"])
                else:
                    self._draw_hl(run_x0, run_y0, run_x1, run_y1)
                    run_x0, run_x1 = hit["x0"], hit["x1"]
                    run_y0 = min(run_y0, hit["y0"])
                    run_y1 = max(run_y1, hit["y1"])
                prev_idx = hit["idx"]
            self._draw_hl(run_x0, run_y0, run_x1, run_y1)
        self.canvas.tag_raise("hl")
        self._set_status("已选中 %d 个字符，Ctrl+C 复制" % len(hits))

    def _draw_hl(self, x0, y0, x1, y1):
        # type: (float, float, float, float) -> None
        self.canvas.create_rectangle(x0 - 1, y0 - 1, x1 + 1, y1 + 1,
                                     fill="#3D8BFD", outline="",
                                     stipple="gray25", tags=("hl",))

    def _selection_text(self):
        # type: () -> str
        if not self._sel_hits:
            return ""
        groups = OrderedDict()
        for hit in self._sel_hits:
            groups.setdefault((hit["page"], hit["line"]), []).append(hit)
        chunks = []                # type: list
        for key in sorted(groups, key=lambda k: (k[0], k[1])):
            items = sorted(groups[key], key=lambda h: h["idx"])
            text = ""
            prev = None
            for hit in items:
                if prev is not None:
                    gap = hit["x_mm"] - (prev["x_mm"] + prev["adv_mm"])
                    if gap > 0.3 * max(hit["size_mm"], prev["size_mm"]):
                        text += " "
                text += hit["ch"]
                prev = hit
            if text.strip():
                chunks.append(text)
        return "\n".join(chunks)

    def copy_selection(self, _event=None):
        # type: (object) -> None
        text = self._selection_text()
        if not text:
            if self.doc is not None:
                self._set_status("没有选中文字：按住鼠标左键拖动选择")
            return "break"
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception as exc:                   # pragma: no cover
            self._set_status("复制失败：%s" % exc)
            return "break"
        self._set_status("已复制 %d 个字符" % len(text))
        return "break"

    # -- 状态 -------------------------------------------------------------
    def _set_status(self, text, info=None):
        # type: (str, str) -> None
        self.status_text = text
        if info is not None:
            self.info_text = info
        if self.win.view is self:
            self.win.set_status(text, self.info_text)

    # -- 导出 -------------------------------------------------------------
    def _ask_save(self, ext, filetypes):
        # type: (str, list) -> str
        base = os.path.splitext(os.path.basename(self.path or "输出"))[0]
        initial = (os.path.dirname(self.path) or os.getcwd()) + os.sep
        return filedialog.asksaveasfilename(
            title="导出为 %s" % ext.upper(), initialdir=initial,
            initialfile=base + ext, defaultextension=ext,
            filetypes=filetypes)

    def export_word(self):
        # type: () -> None
        if self.doc is None or self.win.busy:
            return
        target = self._ask_save(".docx", [("Word 文档", "*.docx")])
        if not target:
            return
        src = self.path
        options = ox.Options()
        # 矢量横线与印章通过透明图层还原，缺了它红头/版记横线和印章会丢
        options.overlay_func = vw.overlay_png

        def job():
            report = ox.Report()
            ox.convert_ofd(src, target, options, report)
            return report

        def done(report):
            totals = report.totals()
            warns = len([w for w in report.warnings if w.level == "warn"])
            msg = ("已导出 Word：\n%s\n\n页数 %d ｜ 段落 %d ｜ 表格 %d ｜ 图片 %d"
                   % (target, totals["pages"], totals["paragraphs"],
                      totals["tables"], totals["images"]))
            if warns:
                msg += ("\n\n有 %d 条告警，可在「工具 → 查看转换报告」里查看。"
                        % warns)
                self._last_report = report
            messagebox.showinfo(APP_NAME, msg, parent=self.root)
            self._set_status("已导出 %s" % os.path.basename(target))

        self.win.start_job("正在导出 Word…", job, done)

    def export_pdf(self):
        # type: () -> None
        if self.doc is None or self.win.busy:
            return
        target = self._ask_save(".pdf", [("PDF 文件", "*.pdf")])
        if not target:
            return
        src = self.path

        def job():
            doc = ox.OfdDocument(src)
            try:
                resolver = vw.FontResolver(doc.res)
                images = []
                for index in range(1, doc.page_count + 1):
                    page = doc.get_page(index)
                    images.append(vw.render_page(page, resolver, dpi=150))
            finally:
                doc.close()
            vw.export_pdf(images, target, dpi=150)
            return len(images)

        def done(count):
            size_kb = os.path.getsize(target) / 1024.0
            messagebox.showinfo(
                APP_NAME,
                "已导出 PDF（%d 页，%.0f KB）：\n%s\n\n"
                "说明：PDF 每页是一张位图，适合查看与打印；"
                "需要可编辑文档请用「导出 Word」。"
                % (count, size_kb, target), parent=self.root)
            self._set_status("已导出 %s" % os.path.basename(target))

        self.win.start_job("正在导出 PDF…", job, done)


# ---------------------------------------------------------------------------
# 窗口：工具栏 + 标签栏 + 多份 DocView + 状态栏
# ---------------------------------------------------------------------------

class ReaderApp(object):
    """一个窗口。主窗口和「独立窗口」是同一个类，靠 standalone 区分：

    standalone=True 时，最后一个标签关掉就自动关窗（适合单独看一份文件）。
    """

    def __init__(self, root, paths=None, standalone=False):
        # type: (object, list, bool) -> None
        self.root = root
        self.standalone = standalone
        self._views = []            # type: list[DocView]
        self._active = -1
        self._busy = False
        self._tab_items = {}        # type: dict  view -> {frame,label,close}
        self._frame_map = {}        # type: dict  widget -> view
        self._tab_menu = None
        self._empty = None
        self._closed = False

        if isinstance(self.root, tk.Tk):
            self.root.geometry("1120x840")
            self.root.minsize(820, 560)
        else:
            self.root.geometry("1080x800")
            self.root.minsize(760, 520)

        _WINDOWS.append(self)
        _install_wheel_dispatch(self.root)

        self._build_ui()
        self._bind_keys()

        for path in (paths or []):
            if path:
                self.open_in_new_tab(path, activate=True)
        if not self._views:
            self._show_empty(True)
        self._sync_ui()
        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        except Exception:
            pass

    # -- 属性 -------------------------------------------------------------
    @property
    def view(self):
        # type: () -> object
        if 0 <= self._active < len(self._views):
            return self._views[self._active]
        return None

    @property
    def busy(self):
        # type: () -> bool
        return self._busy

    def view_count(self):
        # type: () -> int
        return len(self._views)

    # -- 界面 -------------------------------------------------------------
    def _build_ui(self):
        # type: () -> None
        self.root.title(APP_NAME)
        self._set_icon()

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        try:
            style.configure(".", font=("Microsoft YaHei UI", 9))
        except Exception:
            pass

        # 菜单与工具栏都要绑定它，必须在两者之前创建，且只创建一次
        self.mode_var = tk.StringVar(value=MODE_SINGLE)
        self.page_var = tk.StringVar(value="-/0")
        self.zoom_var = tk.StringVar(value="100%")
        self._build_menu()

        self._build_toolbar()
        self._build_tabbar()

        self.content = ttk.Frame(self.root)
        self.content.pack(side="top", fill="both", expand=True)
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

        self._build_empty_pane()
        self._build_status()

    def _build_toolbar(self):
        # type: () -> None
        toolbar = ttk.Frame(self.root, padding=(6, 5))
        toolbar.pack(side="top", fill="x")
        ttk.Button(toolbar, text="打开", width=7,
                   command=self.choose_file).pack(side="left")
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y",
                                                       padx=6)
        self.btn_prev = ttk.Button(toolbar, text="◀ 上一页", width=9,
                                   command=self.prev_page, state="disabled")
        self.btn_prev.pack(side="left")
        self.page_entry = ttk.Entry(toolbar, textvariable=self.page_var, width=7,
                                    justify="center", state="disabled")
        self.page_entry.pack(side="left", padx=4)
        self.page_entry.bind("<Return>", self._on_page_entry)
        self.btn_next = ttk.Button(toolbar, text="下一页 ▶", width=9,
                                   command=self.next_page, state="disabled")
        self.btn_next.pack(side="left")

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y",
                                                       padx=6)
        self._view_icons = make_view_icons()
        if self._view_icons[0] is not None:
            self.rb_single = ttk.Radiobutton(
                toolbar, style="Toolbutton", image=self._view_icons[0],
                variable=self.mode_var, value=MODE_SINGLE,
                command=lambda: self.set_mode(MODE_SINGLE))
            self.rb_double = ttk.Radiobutton(
                toolbar, style="Toolbutton", image=self._view_icons[1],
                variable=self.mode_var, value=MODE_SPREAD,
                command=lambda: self.set_mode(MODE_SPREAD))
        else:
            self.rb_single = ttk.Radiobutton(
                toolbar, style="Toolbutton", text="单页",
                variable=self.mode_var, value=MODE_SINGLE,
                command=lambda: self.set_mode(MODE_SINGLE))
            self.rb_double = ttk.Radiobutton(
                toolbar, style="Toolbutton", text="双页",
                variable=self.mode_var, value=MODE_SPREAD,
                command=lambda: self.set_mode(MODE_SPREAD))
        self.rb_single.pack(side="left")
        self.rb_double.pack(side="left", padx=(2, 0))

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y",
                                                       padx=6)
        ttk.Button(toolbar, text="－", width=3,
                   command=self.zoom_out).pack(side="left")
        self.zoom_box = ttk.Combobox(toolbar, textvariable=self.zoom_var, width=7,
                                     state="readonly",
                                     values=["25%", "50%", "75%", "100%",
                                             "125%", "150%", "200%", "300%",
                                             "400%"])
        self.zoom_box.pack(side="left", padx=3)
        self.zoom_box.bind("<<ComboboxSelected>>", self._on_zoom_selected)
        ttk.Button(toolbar, text="＋", width=3,
                   command=self.zoom_in).pack(side="left")
        ttk.Button(toolbar, text="适应宽度", width=9,
                   command=self.fit_width).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="实际大小", width=9,
                   command=self.actual_size).pack(side="left", padx=(4, 0))

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y",
                                                       padx=6)
        self.btn_word = ttk.Button(toolbar, text="导出 Word", width=10,
                                   command=self.export_word, state="disabled")
        self.btn_word.pack(side="left")
        self.btn_pdf = ttk.Button(toolbar, text="导出 PDF", width=10,
                                  command=self.export_pdf, state="disabled")
        self.btn_pdf.pack(side="left", padx=(4, 0))

    def _build_tabbar(self):
        # type: () -> None
        bar = tk.Frame(self.root, bg=TAB_BAR_BG, height=TAB_BAR_H)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)

        self.tab_canvas = tk.Canvas(bar, bg=TAB_BAR_BG, height=TAB_BAR_H,
                                    highlightthickness=0, bd=0,
                                    xscrollincrement=20)
        self.tab_canvas.pack(side="left", fill="both", expand=True)
        self.tab_holder = tk.Frame(self.tab_canvas, bg=TAB_BAR_BG)
        self._tab_window = self.tab_canvas.create_window(
            (0, 0), window=self.tab_holder, anchor="nw")
        self.tab_holder.bind("<Configure>", self._on_tabs_configure)
        self.tab_canvas.bind("<Configure>", self._on_tabs_configure)

        # 右侧的「＋」，像浏览器一样新建标签
        self.btn_add_tab = tk.Label(bar, text="＋", bg=TAB_BAR_BG, fg="#4a5a66",
                                    font=("Microsoft YaHei UI", 11),
                                    padx=10, cursor="hand2")
        self.btn_add_tab.pack(side="right", fill="y")
        self.btn_add_tab.bind("<Button-1>", lambda e: self.choose_file())
        self.btn_add_tab.bind("<Enter>",
                              lambda e: self.btn_add_tab.config(fg="#0b6cd4"))
        self.btn_add_tab.bind("<Leave>",
                              lambda e: self.btn_add_tab.config(fg="#4a5a66"))

    def _build_empty_pane(self):
        # type: () -> None
        self._empty = ttk.Frame(self.content)
        inner = ttk.Frame(self._empty)
        inner.pack(expand=True)
        ttk.Label(inner, text="还没有打开文档",
                  font=("Microsoft YaHei UI", 15)).pack(pady=(0, 8))
        ttk.Label(inner,
                  text="点击下方按钮选择一个或多个 .ofd 文件；"
                       "也可以把文件拖到程序图标上，或用右键「打开方式」。",
                  foreground="#666666").pack()
        ttk.Button(inner, text="打开文件…", width=16,
                   command=self.choose_file).pack(pady=14)

    def _build_status(self):
        # type: () -> None
        status = ttk.Frame(self.root, padding=(8, 4))
        status.pack(side="bottom", fill="x")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status, textvariable=self.status_var,
                  anchor="w").pack(side="left", fill="x", expand=True)
        self.info_var = tk.StringVar(value="")
        ttk.Label(status, textvariable=self.info_var, anchor="e",
                  foreground="#555555").pack(side="right")

    def _build_menu(self):
        # type: () -> None
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="打开…", accelerator="Ctrl+O",
                              command=self.choose_file)
        file_menu.add_command(label="新建窗口", accelerator="Ctrl+N",
                              command=self.new_window)
        file_menu.add_separator()
        file_menu.add_command(label="关闭当前标签", accelerator="Ctrl+W",
                              command=self.close_current)
        file_menu.add_command(label="关闭其他标签",
                              command=lambda: self.close_others(self.view))
        file_menu.add_command(label="全部关闭", command=self.close_all)
        file_menu.add_separator()
        file_menu.add_command(label="在新窗口中打开当前文档",
                              command=lambda: self.detach_view(self.view))
        file_menu.add_separator()
        file_menu.add_command(label="导出为 Word…", accelerator="Ctrl+S",
                              command=self.export_word)
        file_menu.add_command(label="导出为 PDF…", accelerator="Ctrl+P",
                              command=self.export_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_close)
        menubar.add_cascade(label="文件", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="复制", accelerator="Ctrl+C",
                              command=self.copy_selection)
        edit_menu.add_command(label="取消选择", accelerator="Esc",
                              command=self.clear_selection)
        menubar.add_cascade(label="编辑", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_radiobutton(label="单页", variable=self.mode_var,
                                  value=MODE_SINGLE,
                                  command=lambda: self.set_mode(MODE_SINGLE))
        view_menu.add_radiobutton(label="双页", variable=self.mode_var,
                                  value=MODE_SPREAD,
                                  command=lambda: self.set_mode(MODE_SPREAD))
        view_menu.add_separator()
        view_menu.add_command(label="上一页", accelerator="←",
                              command=self.prev_page)
        view_menu.add_command(label="下一页", accelerator="→",
                              command=self.next_page)
        view_menu.add_separator()
        view_menu.add_command(label="放大 10%", accelerator="Ctrl+=",
                              command=self.zoom_in)
        view_menu.add_command(label="缩小 10%", accelerator="Ctrl+-",
                              command=self.zoom_out)
        view_menu.add_command(label="适应宽度", accelerator="Ctrl+0",
                              command=self.fit_width)
        view_menu.add_command(label="实际大小", accelerator="Ctrl+1",
                              command=self.actual_size)
        view_menu.add_separator()
        view_menu.add_command(label="下一个标签", accelerator="Ctrl+Tab",
                              command=lambda: self.cycle_tab(1))
        view_menu.add_command(label="上一个标签",
                              accelerator="Ctrl+Shift+Tab",
                              command=lambda: self.cycle_tab(-1))
        menubar.add_cascade(label="视图", menu=view_menu)

        tool_menu = tk.Menu(menubar, tearoff=0)
        tool_menu.add_command(label="添加到右键「打开方式」…",
                              command=self.install_association)
        tool_menu.add_command(label="从右键菜单移除…",
                              command=self.remove_association)
        tool_menu.add_separator()
        tool_menu.add_command(label="查看转换报告…",
                              command=self.show_report)
        menubar.add_cascade(label="工具", menu=tool_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="已知限制", command=self.show_limits)
        help_menu.add_separator()
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    def _set_icon(self):
        # type: () -> None
        icon = resource_path(os.path.join("assets", "logo.ico"))
        try:
            if os.path.exists(icon):
                self.root.iconbitmap(default=icon)
        except Exception:
            pass

    def _bind_keys(self):
        # type: () -> None
        self.root.bind("<Control-o>", lambda e: self.choose_file())
        self.root.bind("<Control-t>", lambda e: self.choose_file())
        self.root.bind("<Control-w>", lambda e: self.close_current())
        self.root.bind("<Control-n>", lambda e: self.new_window())
        self.root.bind("<Control-s>", lambda e: self.export_word())
        self.root.bind("<Control-p>", lambda e: self.export_pdf())
        self.root.bind("<Control-c>", self.copy_selection)
        self.root.bind("<Escape>", lambda e: self.clear_selection())
        self.root.bind("<Control-equal>", lambda e: self.zoom_in())
        self.root.bind("<Control-plus>", lambda e: self.zoom_in())
        self.root.bind("<Control-minus>", lambda e: self.zoom_out())
        self.root.bind("<Control-Key-0>", lambda e: self.fit_width())
        self.root.bind("<Control-Key-1>", lambda e: self.actual_size())

        # 方向键翻页：在页码输入框里按左右键要能移动光标，不能翻页
        for seq, action in (("<Left>", self.prev_page),
                            ("<Right>", self.next_page),
                            ("<Prior>", self.prev_page),
                            ("<Next>", self.next_page)):
            self.root.bind(seq, self._guard_key(action))

        for seq in ("<Control-Tab>", "<Control-Next>"):
            try:
                self.root.bind(seq, self._guard_key(lambda: self.cycle_tab(1)))
            except Exception:
                pass
        for seq in ("<Control-Shift-Tab>", "<Control-Prior>"):
            try:
                self.root.bind(seq, self._guard_key(lambda: self.cycle_tab(-1)))
            except Exception:
                pass

        # 滚轮的全局绑定由 _install_wheel_dispatch 统一安装，
        # 这里不能再 bind_all，否则会把别的窗口顶掉。

    def _guard_key(self, action):
        # type: (object) -> object
        def handler(event=None):
            if self._entry_focused():
                return None
            action()
            return "break"
        return handler

    def _entry_focused(self):
        # type: () -> bool
        try:
            focused = self.root.focus_get()
        except Exception:
            focused = None
        return focused is self.page_entry

    # -- 标签栏 -----------------------------------------------------------
    def _on_tabs_configure(self, _event=None):
        # type: (object) -> None
        self._refresh_tab_strip()

    def _refresh_tab_strip(self):
        # type: () -> None
        """标签增删后重算滚动范围。

        不能只靠 <Configure>：holder 被拉宽到画布宽度后尺寸不再变化，
        不会再触发事件，滚动范围就会一直停在上一次的值。
        """
        try:
            self.tab_holder.update_idletasks()
            total = max(1, self.tab_holder.winfo_reqwidth())
            canvas_w = max(1, self.tab_canvas.winfo_width())
            width = max(canvas_w, total)
            self.tab_canvas.itemconfigure(self._tab_window, width=width)
            self.tab_canvas.configure(scrollregion=(0, 0, width, TAB_BAR_H))
            if total <= canvas_w:
                self.tab_canvas.xview_moveto(0)
        except Exception:
            pass

    def _unique_title(self, path):
        # type: (str) -> str
        """同一个文件被打开多次时，标题后面加序号区分。"""
        name = os.path.basename(path) or "(未命名)"
        try:
            key = os.path.normcase(os.path.abspath(path))
            same = sum(1 for v in self._views
                       if v.path
                       and os.path.normcase(os.path.abspath(v.path)) == key)
        except Exception:
            same = 1
        # 调用时该视图已加入 _views，所以第一次 same == 1、第二次 same == 2
        suffix = (" ·%d" % same) if same > 1 else ""
        limit = 20 - len(suffix)
        if len(name) > limit:
            name = name[:max(8, limit - 1)] + "…"
        return name + suffix

    def _create_tab(self, view):
        # type: (DocView) -> None
        holder = tk.Frame(self.tab_holder, bg=TAB_BG, cursor="hand2")
        holder.pack(side="left", fill="y", padx=(0, 1), pady=(4, 0))
        label = tk.Label(holder, text=self._unique_title(view.path), bg=TAB_BG,
                         fg=TAB_FG, font=("Microsoft YaHei UI", 9),
                         padx=10, pady=3, anchor="w")
        label.pack(side="left")
        close = tk.Label(holder, text="×", bg=TAB_BG, fg="#7c8b97",
                         font=("Microsoft YaHei UI", 10), padx=6,
                         cursor="hand2")
        close.pack(side="left", fill="y")

        self._tab_items[view] = {"frame": holder, "label": label, "close": close}
        self._refresh_tab_strip()

        for widget in (holder, label):
            widget.bind("<Button-1>",
                        lambda e, v=view: self.activate_view(v))
            widget.bind("<Double-Button-1>",
                        lambda e, v=view: self.detach_view(v))
            widget.bind("<Button-2>",
                        lambda e, v=view: self.close_view(v))
            widget.bind("<Button-3>",
                        lambda e, v=view: self._popup_tab_menu(e, v))
            widget.bind("<Enter>", lambda e, v=view: self._hover_tab(v, True))
            widget.bind("<Leave>", lambda e, v=view: self._hover_tab(v, False))
        close.bind("<Button-1>", lambda e, v=view: self._on_tab_close(e, v))
        close.bind("<Enter>", lambda e, v=view: close.config(fg="#c62828"))
        close.bind("<Leave>", lambda e, v=view: self._hover_tab(v, False))

    def _tab_tip(self, view):
        # type: (DocView) -> str
        return ("%s\n%d 页 ｜ 双击拆成独立窗口，中键或右键可关闭"
                % (os.path.basename(view.path), view.page_count))

    def _hover_tab(self, view, entering):
        # type: (DocView, bool) -> None
        tab = self._tab_items.get(view)
        if tab is None:
            return
        if view is self.view:
            bg = TAB_ACTIVE_BG
        else:
            bg = TAB_BG_HOVER if entering else TAB_BG
        tab["frame"].config(bg=bg)
        tab["label"].config(bg=bg)
        tab["close"].config(
            bg=bg, fg=("#8a97a3" if view is self.view else "#93a1ad"))

    def _style_tabs(self):
        # type: () -> None
        for view, tab in self._tab_items.items():
            active = (view is self.view)
            bg = TAB_ACTIVE_BG if active else TAB_BG
            fg = TAB_ACTIVE_FG if active else TAB_FG
            tab["frame"].config(bg=bg)
            tab["label"].config(
                bg=bg, fg=fg,
                font=("Microsoft YaHei UI", 9, "bold" if active else "normal"))
            tab["close"].config(bg=bg, fg=("#8a97a3" if active else "#93a1ad"))

    def _scroll_tab_into_view(self, view):
        # type: (DocView) -> None
        tab = self._tab_items.get(view)
        if tab is None:
            return
        try:
            self.tab_holder.update_idletasks()
            total = float(max(1, self.tab_holder.winfo_reqwidth()))
            canvas_w = float(max(1, self.tab_canvas.winfo_width()))
            if total <= canvas_w:
                self.tab_canvas.xview_moveto(0)
                return
            left = tab["frame"].winfo_x()
            width = tab["frame"].winfo_width()
            offset = self.tab_canvas.canvasx(0)
            if left < offset:
                self.tab_canvas.xview_moveto(left / total)
            elif left + width > offset + canvas_w:
                self.tab_canvas.xview_moveto(
                    max(0.0, (left + width - canvas_w) / total))
        except Exception:
            pass

    def _popup_tab_menu(self, event, view):
        # type: (object, DocView) -> None
        if view is None:
            return
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="在新窗口中打开",
                         command=lambda: self.detach_view(view))
        menu.add_separator()
        menu.add_command(label="关闭", command=lambda: self.close_view(view))
        menu.add_command(label="关闭其他标签",
                         command=lambda: self.close_others(view))
        menu.add_command(label="全部关闭", command=self.close_all)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _on_tab_close(self, event, view):
        # type: (object, DocView) -> str
        self.close_view(view)
        return "break"

    # -- 标签管理 ---------------------------------------------------------
    def open_in_new_tab(self, path, activate=True, state=None):
        # type: (str, bool, dict) -> object
        state = state or {}
        zoom = state.get("zoom")
        if state.get("fit_mode"):
            zoom = None                 # 交给新窗口按自己的宽度重新适配
        view = DocView(self, path,
                       page_index=state.get("page_index", 1),
                       zoom=zoom,
                       mode=state.get("mode", MODE_SINGLE),
                       fit_mode=state.get("fit_mode", True))
        if view.doc is None:
            view.dispose()          # 打开失败，不留下空标签
            return None
        self._views.append(view)
        self._frame_map[view.frame] = view
        self._create_tab(view)
        self._style_tabs()
        if activate:
            self.activate_view(view)
        else:
            self._sync_ui()
        return view

    def activate_view(self, view):
        # type: (DocView) -> None
        if view is None or view not in self._views:
            return
        if view is self.view and view.frame.winfo_ismapped():
            self._sync_ui()
            return
        old = self.view
        if old is not None and old is not view:
            try:
                old.frame.grid_remove()
            except Exception:
                pass
            old.on_hide()
        index = self._views.index(view)
        self._active = index
        view.frame.grid(row=0, column=0, sticky="nsew")
        view.frame.tkraise()
        self._show_empty(False)
        self._style_tabs()
        self._refresh_tab_strip()
        self._scroll_tab_into_view(view)
        self._update_title()
        view.on_show()
        self._sync_ui()
        try:
            view.canvas.focus_set()
        except Exception:
            pass

    def close_view(self, view):
        # type: (DocView) -> None
        if view is None or view not in self._views:
            return
        index = self._views.index(view)
        was_active = (self.view is view)
        self._views.remove(view)
        self._frame_map.pop(view.frame, None)
        tab = self._tab_items.pop(view, None)
        if tab is not None:
            try:
                tab["frame"].destroy()
            except Exception:
                pass
        self._refresh_tab_strip()
        view.dispose()

        if not self._views:
            self._active = -1
            self._show_empty(True)
            self._update_title()
            self._sync_ui()
            self.set_status("就绪", "")
            if self.standalone:
                self.close_window()
            return

        if was_active:
            self._active = -1
            self.activate_view(self._views[min(index, len(self._views) - 1)])
        else:
            self._style_tabs()
            self._update_title()
            self._sync_ui()

    def close_current(self):
        # type: () -> None
        self.close_view(self.view)

    def close_others(self, view):
        # type: (DocView) -> None
        if view is None:
            return
        for other in list(self._views):
            if other is not view:
                self.close_view(other)
        self.activate_view(view)

    def close_all(self):
        # type: () -> None
        for view in list(self._views):
            self.close_view(view)

    def cycle_tab(self, delta):
        # type: (int) -> None
        if len(self._views) < 2:
            return
        index = (self._active + delta) % len(self._views)
        self.activate_view(self._views[index])

    # -- 多窗口 -----------------------------------------------------------
    def new_window(self, paths=None):
        # type: (list) -> object
        top = tk.Toplevel(self.root)
        win = ReaderApp(top, paths=paths or [], standalone=True)
        try:
            top.deiconify()
            top.lift()
        except Exception:
            pass
        return win

    def detach_view(self, view=None):
        # type: (DocView) -> object
        """把某个标签拆成独立窗口（阅读位置、缩放、单双页一起带过去）。"""
        view = view or self.view
        if view is None or view.doc is None:
            return None
        state = view.snapshot()
        path = view.path
        win = self.new_window([])
        new_view = win.open_in_new_tab(path, activate=True, state=state)
        if new_view is None:
            win.close_window()
            return None
        self.close_view(view)
        return win

    def close_window(self):
        # type: () -> None
        self.on_close()

    def on_close(self):
        # type: () -> None
        if self._closed:
            return
        self._closed = True
        if self in _WINDOWS:
            _WINDOWS.remove(self)
        for view in list(self._views):
            view.dispose()
        self._views = []
        self._tab_items = {}
        self._frame_map = {}
        try:
            self.root.destroy()
        except Exception:
            pass

    # -- 空状态 -----------------------------------------------------------
    def _show_empty(self, show):
        # type: (bool) -> None
        if self._empty is None:
            return
        try:
            if show:
                self._empty.place(relx=0.5, rely=0.45, anchor="center")
            else:
                self._empty.place_forget()
        except Exception:
            pass

    # -- 状态同步 ---------------------------------------------------------
    def set_status(self, text, info=None):
        # type: (str, str) -> None
        try:
            self.status_var.set(text)
            if info is not None:
                self.info_var.set(info)
        except Exception:
            pass

    def refresh_info(self):
        # type: () -> None
        view = self.view
        try:
            self.info_var.set(view.info_text if view is not None else "")
        except Exception:
            pass

    def sync_page_var(self, view, force=False):
        # type: (DocView, bool) -> None
        if view is not self.view:
            return
        if force or not self._entry_focused():
            try:
                self.page_var.set("%d/%d" % (view.page_index, view.page_count))
            except Exception:
                pass

    def sync_zoom_var(self, view):
        # type: (DocView) -> None
        if view is not self.view:
            return
        try:
            self.zoom_var.set("%d%%" % int(round(view.zoom * 100)))
        except Exception:
            pass

    def sync_mode_var(self, view):
        # type: (DocView) -> None
        if view is not self.view:
            return
        try:
            self.mode_var.set(view.mode)
        except Exception:
            pass

    def _update_title(self):
        # type: () -> None
        view = self.view
        count = len(self._views)
        try:
            if view is None:
                self.root.title(APP_NAME)
            elif count <= 1:
                self.root.title("%s - %s" % (os.path.basename(view.path),
                                             APP_NAME))
            else:
                self.root.title("%s 等 %d 个文档 - %s"
                                % (os.path.basename(view.path), count, APP_NAME))
        except Exception:
            pass

    def _sync_ui(self):
        # type: () -> None
        view = self.view
        has = view is not None and view.doc is not None
        enabled = "normal" if (has and not self._busy) else "disabled"
        for btn in (self.btn_prev, self.btn_next, self.btn_word, self.btn_pdf):
            try:
                btn.config(state=enabled)
            except Exception:
                pass
        try:
            self.page_entry.config(state="normal" if has else "disabled")
        except Exception:
            pass
        if has:
            self.sync_page_var(view, force=True)
            self.zoom_var.set("%d%%" % int(round(view.zoom * 100)))
            self.mode_var.set(view.mode)
        else:
            self.page_var.set("-/0")
            self.zoom_var.set("100%")
        self.refresh_info()
        self._update_title()

    def set_busy(self, busy, text=""):
        # type: (bool, str) -> None
        self._busy = busy
        try:
            self.root.config(cursor="watch" if busy else "")
        except Exception:
            pass
        self._sync_ui()
        if text:
            self.set_status(text)

    # -- 打开 -------------------------------------------------------------
    def choose_file(self):
        # type: () -> None
        current = self.view.path if self.view is not None else ""
        initial = os.path.dirname(current) if current else ""
        paths = filedialog.askopenfilenames(
            title="选择 OFD 文件（可按住 Ctrl 多选）",
            initialdir=initial or None,
            filetypes=[("OFD 文档", "*.ofd"), ("所有文件", "*.*")])
        if not paths:
            return
        if isinstance(paths, str):                 # 老版本 Tk 返回 Tcl 字符串
            try:
                paths = self.root.tk.splitlist(paths)
            except Exception:
                paths = [paths]
        opened = 0
        for path in paths:
            if self.open_in_new_tab(path, activate=True) is not None:
                opened += 1
        if len(paths) > 1:
            self.set_status("已打开 %d 个文件，共 %d 个标签"
                            % (opened, len(self._views)))

    # -- 事件 -------------------------------------------------------------
    def _view_of_widget(self, widget):
        # type: (object) -> object
        while widget is not None:
            view = self._frame_map.get(widget)
            if view is not None:
                return view
            widget = getattr(widget, "master", None)
        return None

    def _in_tabbar(self, widget):
        # type: (object) -> bool
        while widget is not None:
            if widget in (self.tab_canvas, self.tab_holder, self.btn_add_tab):
                return True
            if widget in [t["frame"] for t in self._tab_items.values()]:
                return True
            widget = getattr(widget, "master", None)
        return False

    def scroll_wheel(self, event):
        # type: (object) -> bool
        """处理滚轮。返回 True 表示这个事件归本窗口，调用方会 break。"""
        if self._closed:
            return False
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            widget = None
        if widget is None:
            return False
        if self._in_tabbar(widget):
            step = -int(event.delta / 120) * 60
            if step:
                self.tab_canvas.xview_scroll(step, "units")
            return True
        if self._entry_focused():
            return False
        if event.state & 0x0004:               # Ctrl 交给 zoom_wheel
            return False
        view = self._view_of_widget(widget)
        if view is None:
            return False
        step = -int(event.delta / 120) * 60
        if step:
            view.canvas.yview_scroll(step, "units")
        return True

    def zoom_wheel(self, event):
        # type: (object) -> bool
        """处理 Ctrl+滚轮缩放。"""
        if self._closed:
            return False
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            widget = None
        view = self._view_of_widget(widget) if widget is not None else None
        if view is None:
            return False
        if event.delta > 0:
            view.zoom_in()
        else:
            view.zoom_out()
        return True

    def _on_page_entry(self, _event=None):
        # type: (object) -> None
        view = self.view
        if view is None:
            return
        view.go_to_page_text(self.page_var.get())

    def _on_zoom_selected(self, _event=None):
        # type: (object) -> None
        view = self.view
        if view is None:
            return
        text = self.zoom_var.get().replace("%", "").strip()
        try:
            value = float(text)
        except ValueError:
            return
        view.set_zoom_percent(value)

    # -- 动作转发 ---------------------------------------------------------
    def prev_page(self):
        # type: () -> None
        if self.view is not None:
            self.view.prev_page()

    def next_page(self):
        # type: () -> None
        if self.view is not None:
            self.view.next_page()

    def set_mode(self, mode):
        # type: (str) -> None
        if self.view is not None:
            self.view.set_mode(mode)
        else:
            self.mode_var.set(mode)

    def zoom_in(self):
        # type: () -> None
        if self.view is not None:
            self.view.zoom_in()

    def zoom_out(self):
        # type: () -> None
        if self.view is not None:
            self.view.zoom_out()

    def fit_width(self):
        # type: () -> None
        if self.view is not None:
            self.view.fit_width()

    def actual_size(self):
        # type: () -> None
        if self.view is not None:
            self.view.actual_size()

    def copy_selection(self, _event=None):
        # type: (object) -> object
        if self.view is not None:
            return self.view.copy_selection()
        return "break"

    def clear_selection(self, _event=None):
        # type: (object) -> None
        if self.view is not None:
            self.view.clear_selection()

    def export_word(self):
        # type: () -> None
        if self.view is not None:
            self.view.export_word()

    def export_pdf(self):
        # type: () -> None
        if self.view is not None:
            self.view.export_pdf()

    # -- 后台任务 ---------------------------------------------------------
    def start_job(self, label, func, on_ok=None):
        # type: (str, object, object) -> None
        if self._busy:
            return
        self.set_busy(True, label)

        def worker():
            try:
                result = func()
            except Exception as exc:               # noqa: BLE001
                self.root.after(0, lambda e=exc: self._job_failed(e))
                return
            self.root.after(0, lambda r=result: self._job_done(r, on_ok))

        threading.Thread(target=worker, daemon=True).start()

    def _job_done(self, result, on_ok):
        # type: (object, object) -> None
        self.set_busy(False, "就绪")
        if on_ok is not None:
            try:
                on_ok(result)
            except Exception:                      # pragma: no cover
                pass

    def _job_failed(self, exc):
        # type: (object) -> None
        self.set_busy(False, "导出失败")
        if isinstance(exc, ox.OfdError):
            messagebox.showerror(APP_NAME, exc.user_message(), parent=self.root)
        else:
            messagebox.showerror(
                APP_NAME, "导出失败：%s: %s" % (type(exc).__name__, exc),
                parent=self.root)

    # -- 工具菜单 ---------------------------------------------------------
    def install_association(self):
        # type: () -> None
        exe = app_exe_path()
        if not getattr(sys, "frozen", False):
            if not messagebox.askyesno(
                    APP_NAME,
                    "当前是源码运行模式，注册后右击 .ofd 会用\n%s\n打开，"
                    "而不是本程序。\n\n仍要继续吗？" % exe, parent=self.root):
                return
        make_default = messagebox.askyesno(
            APP_NAME,
            "是否同时把本程序设为 .ofd 的默认打开方式？\n\n"
            "选「否」只加入右键「打开方式」列表，不改动当前默认程序。",
            parent=self.root)
        try:
            register_association(exe, make_default=make_default)
        except Exception as exc:
            messagebox.showerror(APP_NAME, "注册失败：%s" % exc, parent=self.root)
            return
        messagebox.showinfo(
            APP_NAME,
            "已注册。现在可以：\n"
            "· 右击 .ofd 文件 →「打开方式」→ 选择「%s」\n"
            "%s\n\n若列表没立刻出现，注销或重启资源管理器后再看。"
            % (APP_NAME, "· 双击 .ofd 将直接用本程序打开" if make_default else ""),
            parent=self.root)
        self.set_status("已添加到右键「打开方式」")

    def remove_association(self):
        # type: () -> None
        if not messagebox.askyesno(APP_NAME,
                                   "确定要从右键「打开方式」中移除本程序吗？",
                                   parent=self.root):
            return
        count = unregister_association(app_exe_path())
        messagebox.showinfo(APP_NAME, "已移除 %d 项注册信息。" % count,
                            parent=self.root)
        self.set_status("已移除右键注册")

    def show_report(self):
        # type: () -> None
        view = self.view
        report = None
        if view is not None:
            if view._last_report is not None:
                report = view._last_report
            elif view.doc is not None:
                report = view.doc.report
        if report is None:
            messagebox.showinfo(APP_NAME, "还没有可显示的转换报告。",
                                parent=self.root)
            return
        win = tk.Toplevel(self.root)
        win.title("转换报告")
        win.geometry("760x460")
        text = tk.Text(win, wrap="word", font=("Consolas", 9))
        bar = ttk.Scrollbar(win, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=bar.set)
        text.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        totals = report.totals()
        lines = ["汇总：页数 %d，段落 %d，表格 %d，图片 %d，跳过签章 %d，跳过注释 %d"
                 % (totals["pages"], totals["paragraphs"], totals["tables"],
                    totals["images"], totals["signatures_skipped"],
                    totals["annotations_skipped"]), ""]
        if not report.warnings:
            lines.append("没有告警。")
        for item in report.warnings:
            lines.append("[%s] %s" % (item.level.upper(), str(item)))
        text.insert("1.0", "\n".join(lines))
        text.config(state="disabled")

    def show_help(self):
        # type: () -> None
        messagebox.showinfo(
            APP_NAME,
            "同时看多个文件\n"
            "· 「打开」可以按住 Ctrl 一次选多个 .ofd，每个文件一个标签。\n"
            "· 标签从左到右排列，点标签即切换；每个标签自己记页码、缩放和选中状态。\n"
            "· Ctrl+W 关闭当前标签，Ctrl+Tab / Ctrl+Shift+Tab 切换标签。\n"
            "· 双击标签（或右键「在新窗口中打开」）把该文档拆成独立窗口；\n"
            "  独立窗口里也能继续打开别的文件。Ctrl+N 直接开一个新窗口。\n\n"
            "浏览\n"
            "· 滚轮上下翻，右侧滑块可拖到任意位置；所有页面是连续排布的。\n"
            "· ← → 或 PageUp/PageDown 跳到上一页/下一页。\n"
            "· 双击页面可在单页/双页之间切换，也可用工具栏的两个小图标。\n\n"
            "缩放\n"
            "· Ctrl+滚轮、Ctrl+加号/减号，每次 10%。\n"
            "· Ctrl+0 适应宽度，Ctrl+1 实际大小。\n\n"
            "选中文字\n"
            "· 在页面上按住左键拖动框选，蓝色高亮即为选中。\n"
            "· Ctrl+C 复制；右击可弹出「复制」菜单。\n\n"
            "导出\n"
            "· 「导出 Word」得到可编辑的 .docx（表格为真表格）。\n"
            "· 「导出 PDF」得到每页一张位图的 PDF，适合查看和打印。",
            parent=self.root)

    def show_limits(self):
        # type: () -> None
        messagebox.showinfo(
            APP_NAME,
            "已知限制：\n"
            "· 阅读视图按页面渲染，阴影、渐变、裁剪路径等装饰可能显示不全。\n"
            "· 印章会画出来；纯数字签名的签章没有可见图形。\n"
            "· 选中文字按字形位置还原，跨栏或跨页时顺序为从上到下、从左到右。\n"
            "· 导出 Word 为流式排版，换行与字距可能与原件有细微差异。\n"
            "· 导出 PDF 为位图，文字不可选中、不可搜索。\n"
            "· 缺字体时会用系统里的相近字体替代（如方正小标宋简体）。\n"
            "· 同时打开很多大文档会占用较多内存，切走的标签会自动释放已渲染的页面。",
            parent=self.root)

    def show_about(self):
        # type: () -> None
        messagebox.showinfo(
            APP_NAME,
            "%s %s\n\n"
            "OFD 版式文档阅读与转换工具\n"
            "· 阅读：多标签、连续滚动、单页/双页、框选复制\n"
            "· 窗口：标签式浏览，可拆成多个独立窗口\n"
            "· 导出：Word（可编辑）、PDF（位图）\n"
            "· 关联：可注册到右键「打开方式」\n\n"
            "转换内核：ofd2docx %s" % (APP_NAME, APP_VERSION, ox.__version__),
            parent=self.root)


# ---------------------------------------------------------------------------
# 无界面自检（打包后验证依赖是否完整）
# ---------------------------------------------------------------------------

def run_selftest(source, report_path=""):
    # type: (str, str) -> int
    import tempfile
    lines = []          # type: list

    def log(text):
        # type: (str) -> None
        lines.append(text)
        try:
            print(text)
        except Exception:
            pass

    ok = True
    workdir = tempfile.mkdtemp(prefix="ofd_reader_check_")
    log("OFD 阅读转换器自检")
    log("  冻结打包：%s" % bool(getattr(sys, "frozen", False)))
    log("  Python：%s" % sys.version.split()[0])
    log("  工作目录：%s" % workdir)
    try:
        import docx
        import lxml.etree
        import PIL
        log("  依赖：python-docx %s / lxml %s / Pillow %s"
            % (docx.__version__, lxml.etree.__version__, PIL.__version__))
    except Exception as exc:
        ok = False
        log("  [FAIL] 依赖导入失败：%s" % exc)

    if not source or not os.path.exists(source):
        ok = False
        log("  [FAIL] 样本文件不存在：%s" % source)
    else:
        try:
            doc = ox.OfdDocument(source)
            log("  [OK] 打开成功：%d 页，%d 个字体资源"
                % (doc.page_count, len(doc.res.fonts)))
            resolver = vw.FontResolver(doc.res)
            image = vw.render_page(doc.get_page(1), resolver, dpi=96)
            png = os.path.join(workdir, "page1.png")
            image.save(png)
            log("  [OK] 渲染第 1 页：%dx%d -> %s"
                % (image.width, image.height, png))
            vec_total = 0
            seal_total = 0
            for index in range(1, doc.page_count + 1):
                page = doc.get_page(index)
                vec_total += len(page.vectors)
                seal_total += len(page.stamps)
            log("  [OK] 图形元素：矢量 %d 条，印章 %d 个"
                % (vec_total, seal_total))

            docx_path = os.path.join(workdir, "out.docx")
            report = ox.Report()
            options = ox.Options()
            options.overlay_func = vw.overlay_png
            ox.convert_ofd(source, docx_path, options, report)
            size = os.path.getsize(docx_path)
            from docx import Document
            check = Document(docx_path)
            log("  [OK] 导出 Word：%d 页 %d 段 %.0f KB"
                % (len(check.sections), len(check.paragraphs), size / 1024.0))

            images = [vw.render_page(doc.get_page(i), resolver, dpi=110)
                      for i in range(1, doc.page_count + 1)]
            pdf_path = os.path.join(workdir, "out.pdf")
            vw.export_pdf(images, pdf_path, dpi=110)
            log("  [OK] 导出 PDF：%d 页 %.0f KB"
                % (len(images), os.path.getsize(pdf_path) / 1024.0))
            doc.close()
        except Exception as exc:
            ok = False
            log("  [FAIL] %s: %s" % (type(exc).__name__, exc))
            if os.environ.get("OFD_READER_DEBUG"):
                log(traceback.format_exc())

    log("")
    log("结果：%s" % ("全部通过" if ok else "存在失败项"))
    if not report_path:
        report_path = os.path.join(tempfile.gettempdir(),
                                   "ofd_reader_selftest.txt")
    if report_path:
        try:
            with open(report_path, "w", encoding="utf-8") as handle:
                handle.write("\n".join(lines) + "\n")
        except OSError:
            pass
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# 界面冒烟测试
# ---------------------------------------------------------------------------

def run_gui_smoke(source, report_path=""):
    # type: (str, str) -> int
    """真正建窗口，跑一遍多标签/多窗口流程后自动退出。

    覆盖：首页渲染、第二个标签、标签间状态独立、翻页、双页、缩放步进、
    框选复制、拆成独立窗口、关闭标签。
    """
    import time

    lines = []          # type: list

    def log(text):
        # type: (str) -> None
        lines.append(text)
        try:
            print(text)
        except Exception:
            pass

    ok = True
    root = None
    app = None
    log("OFD 阅读转换器 GUI 冒烟测试")
    log("  冻结打包：%s" % bool(getattr(sys, "frozen", False)))

    def pump(seconds):
        # type: (float) -> None
        deadline = time.time() + seconds
        while time.time() < deadline:
            root.update()
            time.sleep(0.03)

    def check(condition, text):
        # type: (bool, str) -> bool
        log("  [%s] %s" % ("OK" if condition else "FAIL", text))
        return bool(condition)

    def grab(window, path):
        # type: (object, str) -> None
        try:
            from PIL import ImageGrab
            window.lift()
            try:
                window.attributes("-topmost", True)
            except Exception:
                pass
            window.update()
            pump(0.4)
            x = window.winfo_rootx()
            y = window.winfo_rooty()
            box = (x, y, x + window.winfo_width(), y + window.winfo_height())
            ImageGrab.grab(bbox=box).save(path)
            try:
                window.attributes("-topmost", False)
            except Exception:
                pass
            log("  [OK] 界面截图：%s" % path)
        except Exception as exc:
            log("  [WARN] 截图失败：%s" % exc)

    try:
        root = tk.Tk()
        app = ReaderApp(root, paths=[source])
        root.update()
        pump(2.2)

        view1 = app.view
        rendered = view1 is not None and len(view1._cache) > 0
        ok = check(rendered, "首页渲染：%s"
                   % ("已生成页面位图" if rendered else "未生成")) and ok
        if view1 is not None:
            log("      窗口 %dx%d，共 %d 页，缩放 %d%%，标签 %d 个"
                % (root.winfo_width(), root.winfo_height(), view1.page_count,
                   int(view1.zoom * 100), app.view_count()))

        # 第二个标签：同一文件再开一份，验证多标签机制
        view2 = app.open_in_new_tab(source, activate=True)
        pump(1.2)
        two = (view2 is not None and app.view_count() == 2
               and app.view is view2)
        ok = check(two, "第二个标签：当前 %d 个标签，激活的是第 %d 个"
                   % (app.view_count(), app._active + 1)) and ok
        if view2 is not None:
            log("      标签标题：%s"
                % " | ".join(t["label"].cget("text")
                             for t in app._tab_items.values()))

        # 可选：额外打开几个文件，用来验证多标签排布（截图也更好看）
        extra = os.environ.get("OFD_READER_SMOKE_EXTRA", "")
        extra_paths = [item for item in extra.split(";") if item.strip()]
        for path in extra_paths:
            if os.path.exists(path):
                app.open_in_new_tab(path.strip(), activate=True)
        if extra_paths:
            pump(1.2)
            ok = check(app.view_count() == 2 + len(extra_paths),
                       "追加 %d 个标签后共 %d 个"
                       % (len(extra_paths), app.view_count())) and ok
            app.activate_view(view1)
            pump(0.6)

        # 两个标签状态独立：改第二个的缩放，切回第一个应保持原样
        zoom1 = view1.zoom if view1 is not None else 1.0
        if view2 is not None:
            view2.actual_size()
            pump(0.3)
        app.activate_view(view1)
        pump(0.6)
        independent = (view1 is not None
                       and abs(view1.zoom - zoom1) < 0.002
                       and abs(app.view.zoom - zoom1) < 0.002)
        ok = check(independent, "标签状态独立：切回后缩放 %d%%（另一标签为 %d%%）"
                   % (int(round((view1.zoom if view1 else 0) * 100)),
                      int(round((view2.zoom if view2 else 0) * 100)))) and ok

        # 滚动：滚到第 2 页，页码要跟随且该页被渲染
        if view1 is not None and view1.page_count > 1:
            slot2 = view1._slots.get(2)
            if slot2 is not None:
                view1.canvas.yview_moveto(
                    max(0.0, (slot2["y"] + 4) / max(1.0, view1._content_h)))
                pump(0.8)
                scrolled = (view1.page_index == 2
                            and view1._slot_items.get(2) is not None)
                ok = check(scrolled, "滚动到第 2 页：页码 %d，该页已渲染 %s"
                           % (view1.page_index,
                              view1._slot_items.get(2) is not None)) and ok
                view1._goto_page(1)
                pump(0.4)

            view1.next_page()
            pump(0.5)
            ok = check(view1.page_index >= 2,
                       "下一页按钮跳到第 %d 页" % view1.page_index) and ok

            view1.set_mode(MODE_SPREAD)
            pump(0.7)
            spread = (view1.mode == MODE_SPREAD and view1._rows
                      and len(view1._rows[0]) == 2)
            ok = check(bool(spread), "切到双页：首行 %d 页"
                       % (len(view1._rows[0]) if view1._rows else 0)) and ok
            view1.set_mode(MODE_SINGLE)
            pump(0.4)

        before = view1.zoom
        view1.zoom_in()
        pump(0.4)
        step = round(view1.zoom - before, 3)
        ok = check(abs(step - ZOOM_STEP) < 1e-6,
                   "放大一档：%d%% -> %d%%（步进 %.0f%%）"
                   % (int(round(before * 100)), int(round(view1.zoom * 100)),
                      step * 100)) and ok

        # 框选一行文字
        view1.actual_size()
        view1._goto_page(1)
        pump(0.9)
        slot = view1._slots.get(1)
        model = view1._models.get(1)
        copied = ""
        if slot is not None and model is not None and model.lines:
            line = model.lines[min(3, len(model.lines) - 1)]
            if line.glyphs:
                ppm = view1._ppm()
                x0 = slot["x"] + line.glyphs[0].x * ppm - 2
                y0 = slot["y"] + line.y * ppm - line.size_mm * ppm
                view1._paint_selection(
                    (x0, y0, x0 + 300, y0 + line.size_mm * ppm * 2))
                root.update()
                copied = view1._selection_text()
        ok = check(bool(copied), "框选文字：%d 字，样例 %r"
                   % (len(copied), copied[:16])) and ok

        shot = os.environ.get("OFD_READER_SCREENSHOT")
        if shot:
            grab(root, shot)

        # 拆成独立窗口（标签数用相对值，冒烟时可能已经开了好几个标签）
        before_count = app.view_count()
        detached = app.detach_view(app.view)
        pump(1.0)
        moved = (detached is not None
                 and app.view_count() == before_count - 1
                 and detached.view_count() == 1)
        ok = check(moved, "拆成独立窗口：原窗口 %d -> %d 个标签，新窗口 %d 个"
                   % (before_count, app.view_count(),
                      detached.view_count() if detached is not None else -1)) and ok

        if detached is not None:
            # 两个窗口都要能用滚轮翻页：曾经因为 bind_all 不带 add="+"，
            # 后开的窗口把先开窗口的滚轮回调顶掉，导致一个窗口滚不动。
            def wheel_scrolls(win):
                # type: (object) -> bool
                win.root.lift()
                win.root.update()
                view = win.view
                if view is None:
                    return False
                canvas = view.canvas
                canvas.yview_moveto(0.0)
                win.root.update()
                before = canvas.yview()[0]
                cx = canvas.winfo_rootx() + canvas.winfo_width() // 2
                cy = canvas.winfo_rooty() + canvas.winfo_height() // 2
                canvas.event_generate("<MouseWheel>", delta=-120,
                                      x=60, y=60, rootx=cx, rooty=cy)
                win.root.update()
                return canvas.yview()[0] > before

            ok = check(wheel_scrolls(detached), "独立窗口滚轮翻页") and ok
            ok = check(wheel_scrolls(app), "主窗口滚轮翻页") and ok

            same_page = (detached.view is not None
                         and detached.view.page_index == view1.page_index)
            ok = check(same_page, "拆窗保留阅读位置：第 %d 页"
                       % (detached.view.page_index
                          if detached.view is not None else -1)) and ok
            second = os.environ.get("OFD_READER_SCREENSHOT2")
            if second:
                grab(detached.root, second)
            detached.close_window()
            pump(0.5)
            ok = check(True, "关闭独立窗口") and ok

        # 关掉主窗口最后一个标签，应回到空状态且不报错
        app.close_all()
        pump(0.5)
        empty_ok = (app.view_count() == 0 and app.view is None)
        ok = check(empty_ok, "关闭全部标签后回到空状态") and ok
    except Exception as exc:
        ok = False
        log("  [FAIL] %s: %s" % (type(exc).__name__, exc))
        if os.environ.get("OFD_READER_DEBUG"):
            log(traceback.format_exc())
    finally:
        if app is not None:
            try:
                app.on_close()
            except Exception:
                pass
        elif root is not None:
            try:
                root.destroy()
            except Exception:
                pass
    log("")
    log("结果：%s" % ("通过" if ok else "失败"))
    if report_path:
        try:
            with open(report_path, "w", encoding="utf-8") as handle:
                handle.write("\n".join(lines) + "\n")
        except OSError:
            pass
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def build_argparser():
    # type: () -> argparse.ArgumentParser
    parser = argparse.ArgumentParser(prog="ofd阅读转换器", add_help=True)
    parser.add_argument("files", nargs="*",
                        help="要打开的 .ofd 文件（可给多个，各占一个标签）")
    parser.add_argument("--register", action="store_true",
                        help="注册到右键「打开方式」")
    parser.add_argument("--unregister", action="store_true",
                        help="移除右键「打开方式」注册")
    parser.add_argument("--default", action="store_true",
                        help="与 --register 合用，同时设为默认打开程序")
    parser.add_argument("--selftest", action="store_true",
                        help="无界面自检（需配合样本文件）")
    parser.add_argument("--smoke-gui", action="store_true", dest="smoke_gui",
                        help="建窗口跑一遍多标签流程后自动退出")
    parser.add_argument("--report", default="", help="自检结果写入指定文件")
    parser.add_argument("--version", action="version",
                        version="%s %s" % (APP_NAME, APP_VERSION))
    return parser


def main(argv=None):
    # type: (object) -> int
    args = build_argparser().parse_args(argv)
    attach_parent_console()
    ox.ensure_console_encoding()
    set_app_user_model_id()

    if args.register:
        try:
            register_association(app_exe_path(), make_default=args.default)
        except Exception as exc:
            print("注册失败：%s" % exc)
            return 1
        print("已注册到右键「打开方式」：%s" % app_exe_path())
        return 0

    if args.unregister:
        count = unregister_association(app_exe_path())
        print("已移除 %d 项注册信息" % count)
        return 0

    first = args.files[0] if args.files else ""

    if args.selftest:
        return run_selftest(first, args.report)

    if args.smoke_gui:
        return run_gui_smoke(first, args.report)

    try:
        root = tk.Tk()
    except Exception as exc:                       # pragma: no cover
        try:
            ctypes.windll.user32.MessageBoxW(0, "无法创建窗口：%s" % exc,
                                             APP_NAME, 0x10)
        except Exception:
            pass
        return 1
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
    app = ReaderApp(root, paths=[p for p in args.files if p])
    try:
        root.mainloop()
    finally:
        del app
    return 0


if __name__ == "__main__":
    sys.exit(main())
