# -*- coding: utf-8 -*-
"""OFD 红头体检：把某页顶部区域的对象逐个列出来，定位"红头显示不对"的原因。

用法：
    python tools/inspect_ofd.py 文件.ofd [--page 1] [--top 90] [--all]

输出包括：
  · 文件概况、资源表（字体 / 颜色空间 / DrawParam / 图片）
  · 顶部区域内每个 TextObject / PathObject / ImageObject 的关键属性：
    坐标、字号（同时给出 mm 与 pt）、字体、CTM、文字内容、解析出的颜色
  · 可疑点汇总（调色板索引色、非单位 CTM、Clips、模板页、字号异常等）
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ofd2docx as ox          # noqa: E402


def _font_name(res, fid):
    # type: (object, str) -> str
    if not fid:
        return "-"
    el = res.fonts.get(str(fid))
    if el is None:
        return "%s(未登记)" % fid
    name = ox.attr(el, "FontName") or ox.attr(el, "FamilyName") or "?"
    if ox.child(el, "CompositeFont") is not None:
        name = "*" + name
    return name


def _color_text(res, obj, dp, name):
    # type: (object, object, object, str) -> str
    el = ox.child(obj, name)
    if el is not None:
        kind = res.color_space_type(ox.attr(el, "ColorSpace"))
        value = ox.attr(el, "Value")
        rgb = res.color_from(value, ox.attr(el, "ColorSpace"))
        return "%s=%r (ColorSpace=%s %s)" % (
            name, value, ox.attr(el, "ColorSpace"), kind)
    if dp is None:
        return ""
    node = dp
    depth = 0
    while node is not None and depth < 8:
        e = ox.child(node, name)
        if e is not None:
            cs = ox.attr(e, "ColorSpace")
            return "%s=%r (DrawParam 继承, ColorSpace=%s %s)" % (
                name, ox.attr(e, "Value"), cs, res.color_space_type(cs))
        rel = ox.attr(node, "Relative")
        node = res.params.get(str(rel)) if rel is not None else None
        depth += 1
    return ""


def main():
    # type: () -> int
    parser = argparse.ArgumentParser(description="OFD 页面对象体检")
    parser.add_argument("file")
    parser.add_argument("--page", type=int, default=1)
    parser.add_argument("--top", type=float, default=90.0,
                        help="只看 y 小于该值（mm）的对象，默认 90")
    parser.add_argument("--all", action="store_true", help="输出整页对象")
    args = parser.parse_args()

    path = os.path.abspath(args.file)
    pkg = ox.open_package(path)
    report = ox.Report()
    meta, ofd_root, ofd_part, doc_root, doc_part, res = \
        ox.load_document(pkg, report)

    print("文件：%s" % path)
    print("版本：%s ｜ 页数：%d ｜ 文档页尺寸：%.1f x %.1f mm"
          % (meta.version or "?", len(meta.pages),
             meta.page_box.w, meta.page_box.h))
    print("资源：字体 %d ｜ 颜色空间 %d ｜ DrawParam %d ｜ 多媒体 %d"
          % (len(res.fonts), len(res.colors), len(res.params), len(res.media)))

    if res.colors:
        print("\n颜色空间：")
        for cid, info in sorted(res.colors.items(),
                                key=lambda kv: str(kv[0])):
            kind = info.get("type") if isinstance(info, dict) else info
            palette = info.get("palette") if isinstance(info, dict) else None
            extra = ""
            if palette:
                extra = "，调色板 %d 项，前几项 %s" % (
                    len(palette), palette[:4])
            print("  ID=%-4s Type=%-10s%s" % (cid, kind, extra))

    index = max(1, min(args.page, len(meta.pages)))
    part, template, page_id = meta.pages[index - 1]
    page = ox.load_page(pkg, part, index, meta, res, report, ox.Options(),
                        template,
                        meta.stamps_by_id.get(str(page_id)))
    root = ox.parse_xml(pkg.read(part), part)

    print("\n第 %d 页：%.2f x %.2f mm ｜ 行 %d ｜ 表格 %d ｜ 图片 %d ｜ 矢量 %d"
          % (index, page.box.w, page.box.h, len(page.lines), len(page.tables),
             len(page.images), len(page.vectors)))
    if template:
        print("模板页：%s（模板内容会先于页面内容绘制）" % template)

    text_objs = list(ox.iter_local(root, "TextObject"))
    path_objs = list(ox.iter_local(root, "PathObject"))
    image_objs = list(ox.iter_local(root, "ImageObject"))
    print("本页对象：TextObject %d ｜ PathObject %d ｜ ImageObject %d"
          % (len(text_objs), len(path_objs), len(image_objs)))

    clips = list(ox.iter_local(root, "Clip"))
    limit = 9999.0 if args.all else args.top
    ctm_count = 0
    odd_size = []
    rows = []          # type: list

    for obj in text_objs:
        box = ox.parse_box(ox.attr(obj, "Boundary"))
        if box is None or box.y > limit:
            continue
        ctm = ox.parse_array(ox.attr(obj, "CTM"))
        size = float(ox.attr(obj, "Size") or 0)
        texts = [t.text or "" for t in obj if ox.lname(t) == "TextCode"]
        body = "".join(texts)
        info = {
            "kind": "Text", "y": box.y, "id": ox.attr(obj, "ID"),
            "box": box, "size": size, "ctm": ctm, "text": body,
            "font": _font_name(res, ox.attr(obj, "Font")),
            "dp": ox.attr(obj, "DrawParam"), "codes": len(texts),
        }
        rows.append(info)
        if ctm and not ox.matrix_is_identity(ctm):
            ctm_count += 1
        if size > 30.0:
            odd_size.append((ox.attr(obj, "ID"), size, body[:16]))

    for obj in path_objs:
        box = ox.parse_box(ox.attr(obj, "Boundary"))
        if box is None or box.y > limit:
            continue
        data = ox.path_data_of(obj)
        rows.append({
            "kind": "Path", "y": box.y, "id": ox.attr(obj, "ID"),
            "box": box, "data_len": len(data or ""),
            "fill": ox.attr_true(obj, "Fill"),
            "stroke": (ox.attr_true(obj, "Stroke")
                       if ox.attr(obj, "Stroke") is not None else True),
            "width": ox.attr(obj, "LineWidth"),
            "ctm": ox.parse_array(ox.attr(obj, "CTM")),
            "dp": ox.attr(obj, "DrawParam"),
        })

    for obj in image_objs:
        box = ox.parse_box(ox.attr(obj, "Boundary"))
        if box is None or box.y > limit:
            continue
        rows.append({
            "kind": "Image", "y": box.y, "id": ox.attr(obj, "ID"),
            "box": box, "res": ox.attr(obj, "ResourceID"),
            "ctm": ox.parse_array(ox.attr(obj, "CTM")),
            "dp": ox.attr(obj, "DrawParam"),
        })

    rows.sort(key=lambda r: (r["y"], r["kind"]))
    print("\n顶部 %.0f mm 内对象（按 y 排序）：" % limit)
    for row in rows:
        box = row["box"]
        head = ("  [%s] ID=%-6s y=%7.2f x=%7.2f w=%7.2f h=%7.2f"
                % (row["kind"], row["id"], box.y, box.x, box.w, box.h))
        if row["kind"] == "Text":
            head += ("  字号=%.3fmm(%.1fpt) 字体=%s"
                     % (row["size"], row["size"] * 72.0 / 25.4, row["font"]))
        print(head)
        if row["kind"] == "Text":
            print("        TextCode %d 个，内容 %r"
                  % (row["codes"], row["text"][:40]))
            ctm = row["ctm"]
            if ctm and not ox.matrix_is_identity(ctm):
                print("        CTM=%r（非单位矩阵）" % (ctm,))
            dp = res.draw_param(row["dp"])
            color = _color_text(res, ox.child(root, "x") or None, dp,
                                "FillColor")
            # 直接在 TextObject 上找颜色子元素更准确
            src = next((o for o in text_objs if ox.attr(o, "ID") == row["id"]),
                       None)
            color = _color_text(res, src, dp, "FillColor")
            if color:
                print("        %s" % color)
        elif row["kind"] == "Path":
            dp = res.draw_param(row["dp"])
            src = next((o for o in path_objs if ox.attr(o, "ID") == row["id"]),
                       None)
            print("        路径数据 %d 字符 ｜ Fill=%s Stroke=%s 线宽=%s"
                  % (row["data_len"], row["fill"], row["stroke"], row["width"]))
            for name in ("FillColor", "StrokeColor"):
                text = _color_text(res, src, dp, name)
                if text:
                    print("        %s" % text)
        else:
            print("        ResourceID=%s" % row["res"])
        ctm = row.get("ctm")
        if ctm and not ox.matrix_is_identity(ctm):
            print("        CTM=%r（非单位矩阵）" % (ctm,))

    print("\n可疑点：")
    flagged = False
    index_spaces = [(cid, info) for cid, info in res.colors.items()
                    if (info.get("type") if isinstance(info, dict) else info)
                    == "Index"]
    if index_spaces:
        flagged = True
        print("  · 存在 %d 个 Index（调色板）颜色空间：%s"
              % (len(index_spaces), [c for c, _ in index_spaces]))
        print("    若红头变黑/变灰，多半是调色板索引被当成 RGB 解析。")
    if ctm_count:
        flagged = True
        print("  · 顶部有 %d 个 TextObject 带非单位 CTM，需确认缩放/平移是否应用。"
              % ctm_count)
    if clips:
        flagged = True
        print("  · 本页有 %d 个 Clip，当前渲染未实现裁剪。" % len(clips))
    if template:
        flagged = True
        print("  · 引用了模板页，红头若在模板页里，注意模板与页面是否重复绘制。")
    if odd_size:
        flagged = True
        print("  · 字号异常大的文本对象：%r（可能是 Size 单位被当作 pt 处理）"
              % odd_size[:3])
    if not flagged:
        print("  · 未发现明显异常。")

    pkg.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
