# -*- coding: utf-8 -*-
"""文本一致性校验：把导出的 docx 与一份期望文本逐行比对。

用法：
    python tests/verify_docx_text.py 结果.docx 期望文本.txt

期望文本一行一段（空行会被忽略），用来做回归：只要生成结果与它逐行一致，
就能确认转换没有多字、少字或错序。期望文本属于具体文档，请放在仓库之外
（例如 samples/ 目录，已在 .gitignore 里排除）。
"""

import io
import sys

from docx import Document


def load_expected(path):
    # type: (str) -> list
    with io.open(path, encoding="utf-8") as handle:
        return [line.rstrip("\n") for line in handle if line.strip()]


def load_actual(path):
    # type: (str) -> list
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def main():
    # type: () -> int
    if len(sys.argv) < 3:
        print("用法：python tests/verify_docx_text.py <docx> <期望文本.txt>")
        return 2

    docx_path, expect_path = sys.argv[1], sys.argv[2]
    expected = load_expected(expect_path)
    actual = load_actual(docx_path)

    doc = Document(docx_path)
    print("文档：%d 节，%d 段，%d 个表格，%d 张图片"
          % (len(doc.sections), len(doc.paragraphs), len(doc.tables),
             len(doc.inline_shapes)))
    print("期望 %d 行，实际 %d 行" % (len(expected), len(actual)))

    bad = 0
    for index in range(max(len(expected), len(actual))):
        want = expected[index] if index < len(expected) else "<缺失>"
        got = actual[index] if index < len(actual) else "<缺失>"
        if want != got:
            bad += 1
            if bad <= 20:
                print("第 %d 行不一致\n  期望：%r\n  实际：%r"
                      % (index + 1, want, got))

    if bad == 0:
        print("\n结果：逐行一致")
        return 0
    print("\n结果：%d 行不一致" % bad)
    return 1


if __name__ == "__main__":
    sys.exit(main())
