# -*- coding: utf-8 -*-
"""用本机 Microsoft Word 打开生成的 docx，验证可用性与分页。

Word 若对 OOXML 有意见，会在打开时弹“内容有问题”并修复；这里通过
Documents.Open 的返回值与实际页数判断，并把结果导出为 PDF 备查。

用法：
    python tests/word_check.py <docx> [导出.pdf]
"""

import os
import sys


def main():
    if len(sys.argv) < 2:
        print("用法：python tests/word_check.py <docx> [pdf]")
        return 2
    docx_path = os.path.abspath(sys.argv[1])
    pdf_path = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else ""

    import win32com.client as win32          # pywin32
    from win32com.client import constants as c

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=False,
                                  AddToRecentFiles=False)
        pages = doc.ComputeStatistics(c.wdStatisticPages)
        words = doc.ComputeStatistics(c.wdStatisticWords)
        chars = doc.ComputeStatistics(c.wdStatisticCharacters)
        print("打开成功：页数 %d，字数 %d，字符数 %d" % (pages, words, chars))
        print("段落数 %d，表格数 %d，内嵌图形 %d"
              % (doc.Paragraphs.Count, doc.Tables.Count,
                 doc.InlineShapes.Count))
        if pdf_path:
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(pdf_path, 17)
            print("已导出 PDF：%s" % pdf_path)
        doc.Close(SaveChanges=0)
        doc = None
        return 0
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


if __name__ == "__main__":
    sys.exit(main())
