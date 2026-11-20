# -*- coding: utf-8 -*-
"""纵向精度校验：把 Word 里每段的实际位置与 OFD 原始基线对比。

Word 的 wdVerticalPositionRelativeToPage 返回的是段落起始处文字的顶部（磅），
OFD 里记的是基线（mm），所以比较时把基线减去约 0.88 倍字高，得到近似的顶部。

用法：
    python tests/position_check.py <ofd> <docx>
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ofd2docx as ox          # noqa: E402


def collect_ofd_lines(path):
    report = ox.Report()
    options = ox.Options()
    pkg = ox.open_package(path)
    try:
        meta, _ofd_root, _ofd_part, _doc_root, _doc_part, res = \
            ox.load_document(pkg, report)
        meta.stamps_by_id = ox.collect_stamps(pkg, _ofd_root, _ofd_part, report)
        out = []
        for idx, (part, tpl, page_id) in enumerate(meta.pages, 1):
            page = ox.load_page(pkg, part, idx, meta, res, report, options, tpl,
                                meta.stamps_by_id.get(str(page_id)))
            for ln in page.flow_lines():
                out.append((idx, ln.y, ln.size_mm, ln.text))
        return out
    finally:
        pkg.close()


def collect_word_lines(path):
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=False,
                                  AddToRecentFiles=False)
        out = []
        for i in range(1, doc.Paragraphs.Count + 1):
            rng = doc.Paragraphs(i).Range
            text = rng.Text.replace("\r", "").replace("\x07", "").strip()
            if not text:
                continue
            page = rng.Information(3)          # wdActiveEndPageNumber
            vpt = rng.Information(6)           # wdVerticalPositionRelativeToPage
            out.append((page, vpt * 25.4 / 72.0, text))
        doc.Close(SaveChanges=0)
        doc = None
        return out
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def main():
    if len(sys.argv) < 3:
        print("用法：python tests/position_check.py <ofd> <docx>")
        return 2
    src, dst = sys.argv[1], sys.argv[2]
    want = collect_ofd_lines(src)
    have = collect_word_lines(dst)
    print("OFD 行数 %d，Word 段落数 %d" % (len(want), len(have)))
    if len(want) != len(have):
        print("行数不一致，仅比较前 %d 行" % min(len(want), len(have)))

    n = min(len(want), len(have))
    worst = 0.0
    worst_i = -1
    per_page = {}
    for i in range(n):
        wpage, wy, wsize, wtext = want[i]
        hpage, hy, htext = have[i]
        expect_top = wy - 0.88 * wsize
        residual = hy - expect_top
        per_page.setdefault(wpage, []).append(residual)
        if abs(residual) > abs(worst):
            worst = residual
            worst_i = i

    for page in sorted(per_page):
        vals = per_page[page]
        print("  第 %d 页：%2d 行，纵向偏差 平均 %+5.2f mm，最大 %+5.2f mm"
              % (page, len(vals), sum(vals) / len(vals),
                 max(vals, key=abs)))

    if worst_i >= 0:
        print("偏差最大的行：第 %d 行 %r（偏差 %+.2f mm）"
              % (worst_i + 1, want[worst_i][3][:20], worst))

    # 文本一致性
    mismatch = 0
    for i in range(n):
        a = want[i][3].strip()
        b = have[i][2].strip()
        if a != b:
            mismatch += 1
            if mismatch <= 5:
                print("  文本不一致：%r vs %r" % (a, b))
    print("文本不一致 %d 行" % mismatch)

    ok = mismatch == 0 and abs(worst) <= 4.0
    print("\n结果：%s" % ("通过（纵向偏差 ≤ 4mm）" if ok else "未通过"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
