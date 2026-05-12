#!/usr/bin/env python3
"""
md2gongwen.py — Markdown → 标准公文格式 DOCX（零依赖版本）

用法:
    python3 md2gongwen.py 输入.md [输出.docx] [选项]

选项:
    --scan-fonts      扫描系统字体并打印推荐配置
    --save-config     将当前字体配置保存到 ~/.gongwen_fonts.json
    --config FILE     使用指定的字体配置文件
    --verify          生成后验证 DOCX 结构（纯 stdlib）

公文格式标准（GB/T 9704—2012）:
    标题        小标宋二号 (22pt) 居中
    主送单位     仿宋三号 (16pt) 顶格（> 标记）
    正文        仿宋三号 (16pt) 首行缩进 2 字符
    一级标题     黑体三号 (16pt)  — ##
    二级标题     楷体三号 (16pt)  — ###
    三级标题     仿宋加粗三号    — ####
    落款        右对齐（--- 分隔，尾行按日期识别）
    附件标记     仿宋三号 顶格（## 附件 或 附件：开头）
    页面        A4, 上37/下35/左28/右26mm, 行距固定28磅, 页脚页码

Inline 格式：
    **加粗**  → 仿宋加粗
    *斜体*   → 仿宋斜体（仅正文内生效）

Markdown 写作约定：
    # 公文标题
    > 主送单位
    正文段落，支持 **加粗** 和 *斜体*
    ## 一、一级标题
    ### （一）二级标题
    #### 1. 三级标题
    ---
    发文机关署名
    2026年1月1日
    ## 附件
    1. 附件一名称
    2. 附件二名称

零依赖：仅需 Python 3.8+ 标准库 (zipfile + xml.etree.ElementTree)。
可直接拷贝到内网环境运行，无需 pip install 任何包。
"""

import os
import sys
import json
import re
import platform
import subprocess
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree.ElementTree import Element, SubElement, tostring

# ═══════════════════════════════════════════════════════════════════
# 字体配置
# ═══════════════════════════════════════════════════════════════════

_DARWIN_DEFAULTS = {
    "标题":     "Songti SC",
    "主送单位": "STFangsong",
    "正文":     "STFangsong",
    "一级标题": "Heiti SC",
    "二级标题": "STKaiti",
    "三级标题": "STFangsong",
    "附件":     "STFangsong",
    "落款":     "STFangsong",
}

_WINDOWS_DEFAULTS = {
    "标题":     "方正小标宋简体",
    "主送单位": "仿宋",
    "正文":     "仿宋",
    "一级标题": "黑体",
    "二级标题": "楷体",
    "三级标题": "仿宋",
    "附件":     "仿宋",
    "落款":     "仿宋",
}

_LINUX_DEFAULTS = {
    "标题":     "Noto Serif CJK SC",
    "主送单位": "Noto Serif CJK SC",
    "正文":     "Noto Serif CJK SC",
    "一级标题": "Noto Sans CJK SC",
    "二级标题": "Noto Serif CJK SC",
    "三级标题": "Noto Serif CJK SC",
    "附件":     "Noto Serif CJK SC",
    "落款":     "Noto Serif CJK SC",
}

CONFIG_PATH = Path.home() / ".gongwen_fonts.json"

# 字号/间距常量（单位见注释）
TITLE_SIZE_PT = 22          # 二号
BODY_SIZE_PT  = 16          # 三号
LINE_SPACING_PT = 28        # 固定 28 磅
FIRST_LINE_INDENT_PT = 32   # 首行缩进 ≈ 2 个三号字宽
SIGNATURE_LINE_SPACING_PT = 28  # 落款行距

# A4 页边距 (mm)
MARGIN_TOP    = 37
MARGIN_BOTTOM = 35
MARGIN_LEFT   = 28
MARGIN_RIGHT  = 26

# ── XML 名称空间 ──
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

# ElementTree 注册名称空间（写入时用）
for _pfx, _uri in [("w", NS_W), ("r", NS_R)]:
    try:
        from xml.etree.ElementTree import register_namespace
        register_namespace(_pfx, _uri)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════
# 字体扫描
# ═══════════════════════════════════════════════════════════════════

def get_platform_defaults():
    system = platform.system()
    if system == "Darwin":
        return dict(_DARWIN_DEFAULTS)
    elif system == "Windows":
        return dict(_WINDOWS_DEFAULTS)
    else:
        return dict(_LINUX_DEFAULTS)


def load_font_map(config_path=None):
    """优先级：--config > ~/.gongwen_fonts.json > 平台默认"""
    font_map = get_platform_defaults()
    config_file = Path(config_path) if config_path else CONFIG_PATH
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            font_map.update(saved)
        except Exception:
            pass
    return font_map


def scan_fonts_macos():
    fonts = set()
    try:
        result = subprocess.run(
            ["fc-list", ":lang=zh", "family"],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stdout.splitlines():
            for name in line.split(","):
                name = name.strip()
                if name:
                    fonts.add(name)
    except Exception:
        pass
    if not fonts:
        try:
            for font_dir in ["/System/Library/Fonts",
                             "/System/Library/Fonts/Supplemental",
                             "/Library/Fonts",
                             os.path.expanduser("~/Library/Fonts")]:
                if os.path.isdir(font_dir):
                    for f in os.listdir(font_dir):
                        if f.endswith((".ttf", ".ttc", ".otf")):
                            fonts.add(f.rsplit(".", 1)[0])
        except Exception:
            pass
    return sorted(fonts)


def scan_fonts_windows():
    fonts = set()
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing');"
             "(New-Object System.Drawing.Text.InstalledFontCollection).Families.Name"],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stdout.splitlines():
            line = line.strip()
            if line:
                fonts.add(line)
    except Exception:
        pass
    return sorted(fonts)


def scan_fonts_linux():
    fonts = set()
    try:
        result = subprocess.run(
            ["fc-list", ":lang=zh", "family"],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stdout.splitlines():
            for name in line.split(","):
                name = name.strip()
                if name:
                    fonts.add(name)
    except Exception:
        pass
    return sorted(fonts)


def scan_fonts():
    system = platform.system()
    if system == "Darwin":
        return scan_fonts_macos()
    elif system == "Windows":
        return scan_fonts_windows()
    else:
        return scan_fonts_linux()


def fuzzy_match_fonts(all_fonts):
    keywords = {
        "标题":     ["小标宋", "标宋", "Songti", "宋体", "宋體", "SimSun", "Song",
                     "STSong", "华文宋体"],
        "主送单位": ["仿宋", "Fangsong", "FangSong", "Fang", "STFangsong"],
        "正文":     ["仿宋", "Fangsong", "FangSong", "Fang", "STFangsong"],
        "一级标题": ["黑体", "黑體", "Heiti", "SimHei", "Hei", "STHeiti"],
        "二级标题": ["楷体", "楷體", "Kaiti", "KaiTi", "SimKai", "Kai", "STKaiti"],
        "三级标题": ["仿宋", "Fangsong", "FangSong", "Fang", "STFangsong"],
        "附件":     ["仿宋", "Fangsong", "FangSong", "Fang", "STFangsong"],
        "落款":     ["仿宋", "Fangsong", "FangSong", "Fang", "STFangsong"],
    }
    result = {}
    for role, kws in keywords.items():
        best = None
        for font in all_fonts:
            for kw in kws:
                if kw.lower() in font.lower():
                    best = font
                    break
            if best:
                break
        result[role] = best or "请手动指定"
    return result


# ═══════════════════════════════════════════════════════════════════
# Inline 格式解析
# ═══════════════════════════════════════════════════════════════════

_INLINE_RE = re.compile(
    r"(\*\*(.+?)\*\*|\*(.+?)\*)"
)


def _parse_inline(text):
    """解析行内格式，返回 [(text, bold, italic), ...]"""
    runs = []
    pos = 0
    for m in _INLINE_RE.finditer(text):
        # 匹配前的普通文本
        if m.start() > pos:
            runs.append((text[pos:m.start()], False, False))
        full = m.group(1)
        if full.startswith("**") and full.endswith("**"):
            runs.append((m.group(2), True, False))
        elif full.startswith("*") and full.endswith("*"):
            runs.append((m.group(3), False, True))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], False, False))
    return runs if runs else [(text, False, False)]


# ═══════════════════════════════════════════════════════════════════
# 日期识别
# ═══════════════════════════════════════════════════════════════════

_DATE_RE = re.compile(r"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$")


def _is_date_line(text):
    """判断一行是否为日期格式"""
    return bool(_DATE_RE.match(text.strip()))


# ═══════════════════════════════════════════════════════════════════
# DOCX 生成（纯标准库：zipfile + xml.etree.ElementTree）
# ═══════════════════════════════════════════════════════════════════

def _pt_to_half_pt(pt):
    """字号：pt → half-points (w:sz 单位)"""
    return str(int(pt * 2))


def _pt_to_twips(pt):
    """通用：pt → twips (1/20 pt)"""
    return str(int(pt * 20))


def _mm_to_twips(mm):
    """毫米 → twips (1mm ≈ 56.69 twips)"""
    return str(int(mm * 1440 / 25.4))


def _make_rpr(font_name, size_pt, bold=False, italic=False):
    """构建 <w:rPr> 元素"""
    rPr = Element(f"{{{NS_W}}}rPr")

    rFonts = SubElement(rPr, f"{{{NS_W}}}rFonts")
    rFonts.set(f"{{{NS_W}}}ascii", font_name)
    rFonts.set(f"{{{NS_W}}}hAnsi", font_name)
    rFonts.set(f"{{{NS_W}}}eastAsia", font_name)

    sz = SubElement(rPr, f"{{{NS_W}}}sz")
    sz.set(f"{{{NS_W}}}val", _pt_to_half_pt(size_pt))

    szCs = SubElement(rPr, f"{{{NS_W}}}szCs")
    szCs.set(f"{{{NS_W}}}val", _pt_to_half_pt(size_pt))

    if bold:
        SubElement(rPr, f"{{{NS_W}}}b")
    if italic:
        SubElement(rPr, f"{{{NS_W}}}i")

    return rPr


def _make_ppr(line_spacing_pt=None, first_line_indent_pt=None,
              alignment=None):
    """构建 <w:pPr> 元素"""
    pPr = Element(f"{{{NS_W}}}pPr")

    if line_spacing_pt is not None:
        spacing = SubElement(pPr, f"{{{NS_W}}}spacing")
        spacing.set(f"{{{NS_W}}}line", _pt_to_twips(line_spacing_pt))
        spacing.set(f"{{{NS_W}}}lineRule", "exact")

    if first_line_indent_pt is not None:
        ind = SubElement(pPr, f"{{{NS_W}}}ind")
        ind.set(f"{{{NS_W}}}firstLine", _pt_to_twips(first_line_indent_pt))

    if alignment:
        jc = SubElement(pPr, f"{{{NS_W}}}jc")
        jc.set(f"{{{NS_W}}}val", alignment)

    return pPr


def _make_single_paragraph(text, font_name, size_pt, bold=False, italic=False,
                           alignment=None, first_line_indent_pt=None,
                           line_spacing_pt=LINE_SPACING_PT):
    """构建单 run 段落（向后兼容）"""
    p = Element(f"{{{NS_W}}}p")

    pPr = _make_ppr(
        line_spacing_pt=line_spacing_pt,
        first_line_indent_pt=first_line_indent_pt,
        alignment=alignment,
    )
    p.append(pPr)

    r = SubElement(p, f"{{{NS_W}}}r")
    rPr = _make_rpr(font_name, size_pt, bold, italic)
    r.append(rPr)

    t = SubElement(r, f"{{{NS_W}}}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text

    return p


def _make_formatted_paragraph(runs, font_name, size_pt,
                              alignment=None, first_line_indent_pt=None,
                              line_spacing_pt=LINE_SPACING_PT):
    """构建多 run 段落（支持 inline 格式）"""
    p = Element(f"{{{NS_W}}}p")

    pPr = _make_ppr(
        line_spacing_pt=line_spacing_pt,
        first_line_indent_pt=first_line_indent_pt,
        alignment=alignment,
    )
    p.append(pPr)

    for text, bold, italic in runs:
        r = SubElement(p, f"{{{NS_W}}}r")
        rPr = _make_rpr(font_name, size_pt, bold, italic)
        r.append(rPr)

        t = SubElement(r, f"{{{NS_W}}}t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text

    return p


def _make_empty_paragraph(line_spacing_pt=LINE_SPACING_PT):
    """构建空段落"""
    p = Element(f"{{{NS_W}}}p")
    pPr = _make_ppr(line_spacing_pt=line_spacing_pt)
    p.append(pPr)
    return p


def _make_sect_pr():
    """构建 <w:sectPr> 页面设置（含页码引用）"""
    sectPr = Element(f"{{{NS_W}}}sectPr")

    # 页脚引用
    footerRef = SubElement(sectPr, f"{{{NS_W}}}footerReference")
    footerRef.set(f"{{{NS_W}}}type", "default")
    footerRef.set(f"{{{NS_R}}}id", "rId3")

    # A4 页面大小
    pgSz = SubElement(sectPr, f"{{{NS_W}}}pgSz")
    pgSz.set(f"{{{NS_W}}}w", _mm_to_twips(210))
    pgSz.set(f"{{{NS_W}}}h", _mm_to_twips(297))

    # 页边距
    pgMar = SubElement(sectPr, f"{{{NS_W}}}pgMar")
    pgMar.set(f"{{{NS_W}}}top", _mm_to_twips(MARGIN_TOP))
    pgMar.set(f"{{{NS_W}}}bottom", _mm_to_twips(MARGIN_BOTTOM))
    pgMar.set(f"{{{NS_W}}}left", _mm_to_twips(MARGIN_LEFT))
    pgMar.set(f"{{{NS_W}}}right", _mm_to_twips(MARGIN_RIGHT))

    return sectPr


def _xml_bytes(element):
    """Element → bytes with XML declaration"""
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
           tostring(element, encoding="utf-8", xml_declaration=False)


# ── DOCX ZIP 组装 ──

def build_content_types():
    """[Content_Types].xml"""
    root = Element(f"{{{NS_CT}}}Types")

    SubElement(root, f"{{{NS_CT}}}Default",
               Extension="rels",
               ContentType="application/vnd.openxmlformats-package.relationships+xml")
    SubElement(root, f"{{{NS_CT}}}Default",
               Extension="xml",
               ContentType="application/xml")
    SubElement(root, f"{{{NS_CT}}}Override",
               PartName="/word/document.xml",
               ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    SubElement(root, f"{{{NS_CT}}}Override",
               PartName="/word/styles.xml",
               ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
    SubElement(root, f"{{{NS_CT}}}Override",
               PartName="/word/settings.xml",
               ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
    SubElement(root, f"{{{NS_CT}}}Override",
               PartName="/word/footer1.xml",
               ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

    return _xml_bytes(root)


def build_root_rels():
    """_rels/.rels"""
    root = Element(f"{{{NS_RELS}}}Relationships")
    SubElement(root, f"{{{NS_RELS}}}Relationship",
               Id="rId1",
               Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
               Target="word/document.xml")
    return _xml_bytes(root)


def build_doc_rels():
    """word/_rels/document.xml.rels"""
    root = Element(f"{{{NS_RELS}}}Relationships")
    SubElement(root, f"{{{NS_RELS}}}Relationship",
               Id="rId1",
               Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
               Target="styles.xml")
    SubElement(root, f"{{{NS_RELS}}}Relationship",
               Id="rId2",
               Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
               Target="settings.xml")
    SubElement(root, f"{{{NS_RELS}}}Relationship",
               Id="rId3",
               Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
               Target="footer1.xml")
    return _xml_bytes(root)


def build_styles(font_map):
    """word/styles.xml — 定义 Normal 样式"""
    root = Element(f"{{{NS_W}}}styles")

    style = SubElement(root, f"{{{NS_W}}}style",
                       {f"{{{NS_W}}}type": "paragraph",
                        f"{{{NS_W}}}styleId": "Normal",
                        f"{{{NS_W}}}default": "1"})
    SubElement(style, f"{{{NS_W}}}name", {f"{{{NS_W}}}val": "Normal"})

    rPr = SubElement(style, f"{{{NS_W}}}rPr")
    body_font = font_map.get("正文", "SimSun")
    rFonts = SubElement(rPr, f"{{{NS_W}}}rFonts")
    rFonts.set(f"{{{NS_W}}}ascii", body_font)
    rFonts.set(f"{{{NS_W}}}hAnsi", body_font)
    rFonts.set(f"{{{NS_W}}}eastAsia", body_font)
    SubElement(rPr, f"{{{NS_W}}}sz", {f"{{{NS_W}}}val": _pt_to_half_pt(BODY_SIZE_PT)})

    return _xml_bytes(root)


def build_settings():
    """word/settings.xml"""
    root = Element(f"{{{NS_W}}}settings")
    SubElement(root, f"{{{NS_W}}}defaultTabStop", {f"{{{NS_W}}}val": "420"})
    SubElement(root, f"{{{NS_W}}}characterSpacingControl",
               {f"{{{NS_W}}}val": "doNotCompress"})
    return _xml_bytes(root)


def build_footer(font_map):
    """word/footer1.xml — 页脚页码（居中，仿宋五号 10.5pt）"""
    footer = Element(f"{{{NS_W}}}ftr")

    p = SubElement(footer, f"{{{NS_W}}}p")
    pPr = SubElement(p, f"{{{NS_W}}}pPr")
    jc = SubElement(pPr, f"{{{NS_W}}}jc")
    jc.set(f"{{{NS_W}}}val", "center")

    # 页码字段
    r = SubElement(p, f"{{{NS_W}}}r")
    rPr = SubElement(r, f"{{{NS_W}}}rPr")
    font_name = font_map.get("正文", "STFangsong")
    rFonts = SubElement(rPr, f"{{{NS_W}}}rFonts")
    rFonts.set(f"{{{NS_W}}}ascii", font_name)
    rFonts.set(f"{{{NS_W}}}hAnsi", font_name)
    rFonts.set(f"{{{NS_W}}}eastAsia", font_name)
    sz = SubElement(rPr, f"{{{NS_W}}}sz")
    sz.set(f"{{{NS_W}}}val", _pt_to_half_pt(10.5))  # 五号

    # PAGE 字段
    fldChar1 = SubElement(r, f"{{{NS_W}}}fldChar")
    fldChar1.set(f"{{{NS_W}}}fldCharType", "begin")
    SubElement(r, f"{{{NS_W}}}instrText", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " PAGE "
    fldChar2 = SubElement(r, f"{{{NS_W}}}fldChar")
    fldChar2.set(f"{{{NS_W}}}fldCharType", "end")

    return _xml_bytes(footer)


def build_document(parsed_elements, font_map):
    """构建 word/document.xml 主体"""
    body = Element(f"{{{NS_W}}}body")

    i = 0
    while i < len(parsed_elements):
        typ = parsed_elements[i][0]

        if typ == "blank":
            i += 1
            continue

        elif typ == "h1":
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["标题"], TITLE_SIZE_PT,
                bold=False, alignment="center",
                first_line_indent_pt=None,
                line_spacing_pt=36,
            )
            body.append(p)
            i += 1
            # 跳过后面的空行，保留一个空行
            while i < len(parsed_elements) and parsed_elements[i][0] == "blank":
                i += 1
            body.append(_make_empty_paragraph())

        elif typ == "recipient":
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["主送单位"], BODY_SIZE_PT,
                bold=False, first_line_indent_pt=None,
            )
            body.append(p)
            i += 1

        elif typ == "h2":
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["一级标题"], BODY_SIZE_PT,
                bold=False, first_line_indent_pt=FIRST_LINE_INDENT_PT,
            )
            body.append(p)
            i += 1

        elif typ == "h3":
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["二级标题"], BODY_SIZE_PT,
                bold=False, first_line_indent_pt=FIRST_LINE_INDENT_PT,
            )
            body.append(p)
            i += 1

        elif typ == "h4":
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["三级标题"], BODY_SIZE_PT,
                bold=True, first_line_indent_pt=FIRST_LINE_INDENT_PT,
            )
            body.append(p)
            i += 1

        elif typ == "text":
            text = parsed_elements[i][1]
            runs = _parse_inline(text)
            # 使用 runs 中的格式检测
            has_format = any(bold or italic for _, bold, italic in runs)
            if has_format:
                p = _make_formatted_paragraph(
                    runs, font_map["正文"], BODY_SIZE_PT,
                    first_line_indent_pt=FIRST_LINE_INDENT_PT,
                )
            else:
                p = _make_single_paragraph(
                    text, font_map["正文"], BODY_SIZE_PT,
                    first_line_indent_pt=FIRST_LINE_INDENT_PT,
                )
            body.append(p)
            i += 1

        elif typ == "sig_sep":
            # --- 分隔线：空行
            body.append(_make_empty_paragraph())
            i += 1

        elif typ == "sig_unit":
            # 发文机关署名：右对齐，标准行距
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["落款"], BODY_SIZE_PT,
                alignment="right", first_line_indent_pt=None,
                line_spacing_pt=SIGNATURE_LINE_SPACING_PT,
            )
            body.append(p)
            i += 1

        elif typ == "sig_date":
            # 日期：右对齐
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["落款"], BODY_SIZE_PT,
                alignment="right", first_line_indent_pt=None,
                line_spacing_pt=SIGNATURE_LINE_SPACING_PT,
            )
            body.append(p)
            i += 1

        elif typ == "attachment_header":
            # 附件标题
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["附件"], BODY_SIZE_PT,
                bold=False, first_line_indent_pt=None,
            )
            body.append(p)
            i += 1

        elif typ == "attachment_item":
            # 附件列表项
            text = parsed_elements[i][1]
            p = _make_single_paragraph(
                text, font_map["附件"], BODY_SIZE_PT,
                first_line_indent_pt=None,
            )
            body.append(p)
            i += 1

        else:
            i += 1

    # 页面设置（必须放在 body 末尾）
    body.append(_make_sect_pr())

    document = Element(f"{{{NS_W}}}document")
    document.append(body)

    return _xml_bytes(document)


def write_docx(parsed_elements, font_map, output_path):
    """组装完整 DOCX（ZIP 包）"""
    with ZipFile(output_path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types())
        zf.writestr("_rels/.rels", build_root_rels())
        zf.writestr("word/_rels/document.xml.rels", build_doc_rels())
        zf.writestr("word/styles.xml", build_styles(font_map))
        zf.writestr("word/settings.xml", build_settings())
        zf.writestr("word/footer1.xml", build_footer(font_map))
        zf.writestr("word/document.xml", build_document(parsed_elements, font_map))


# ═══════════════════════════════════════════════════════════════════
# Markdown 解析
# ═══════════════════════════════════════════════════════════════════

def parse_markdown(md_text):
    """解析 Markdown 文本，返回 [(type, text, level), ...]

    type: 'blank' | 'h1'-'h4' | 'recipient' | 'text' |
          'sig_sep' | 'sig_unit' | 'sig_date' |
          'attachment_header' | 'attachment_item'
    """
    lines = md_text.splitlines()
    parsed = []
    in_signature = False
    in_attachment = False

    for line in lines:
        stripped = line.strip()

        if not stripped:
            parsed.append(("blank", "", 0))
            continue

        # --- 分隔线 → 进入落款区域
        if stripped == "---":
            in_signature = True
            in_attachment = False
            parsed.append(("sig_sep", "---", 0))
            continue

        # 落款区域内的行
        if in_signature:
            # 检查是否退出落款（遇到标题行）
            h_match_sig = re.match(r"^(#{1,4})\s+(.+)$", stripped)
            if h_match_sig:
                # 落款结束，回退到正常解析
                in_signature = False
                # 不 continue — 让后续逻辑处理这行
            else:
                if _is_date_line(stripped):
                    parsed.append(("sig_date", stripped, 0))
                else:
                    parsed.append(("sig_unit", stripped, 0))
                continue

        # 附件区域检测（## 附件 或 附件：开头）
        h_match = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if h_match:
            level = len(h_match.group(1))
            text = h_match.group(2).strip()
            if re.match(r"^附件", text):
                if level == 2:
                    # ## 附件 → 附件标题
                    in_attachment = True
                    in_signature = False
                    parsed.append(("attachment_header", text, level))
                    continue
            in_attachment = False  # 非附件标题则退出附件模式
            parsed.append((f"h{level}", text, level))
            continue

        # 单独的附件行（非标题）
        if re.match(r"^附件[：:]", stripped):
            in_attachment = True
            in_signature = False
            parsed.append(("attachment_header", stripped, 0))
            continue

        # 附件列表项（数字. 或 （数字） 开头）
        if in_attachment and re.match(r"^\d+[\.\、\．\）\)]", stripped):
            parsed.append(("attachment_item", stripped, 0))
            continue

        # 引用块 → 主送单位
        if stripped.startswith(">"):
            parsed.append(("recipient", stripped[1:].strip(), 0))
            continue

        # 普通正文
        in_attachment = False  # 普通正文退出附件模式
        parsed.append(("text", stripped, 0))

    return parsed


# ═══════════════════════════════════════════════════════════════════
# DOCX 验证（纯 stdlib）
# ═══════════════════════════════════════════════════════════════════

def verify_docx(docx_path):
    """使用纯 stdlib 验证 DOCX 结构，返回字典"""
    result = {
        "valid_zip": False,
        "parts": [],
        "paragraphs": 0,
        "fonts_used": set(),
        "has_footer": False,
        "page_setup": {},
        "errors": [],
    }

    # 检查 ZIP 结构
    try:
        zf = ZipFile(docx_path, "r")
        names = zf.namelist()
        result["valid_zip"] = True
        result["parts"] = names

        required = ["word/document.xml", "word/styles.xml",
                    "[Content_Types].xml", "_rels/.rels"]
        for r in required:
            if r not in names:
                result["errors"].append(f"缺少必需部件: {r}")

        if "word/footer1.xml" in names:
            result["has_footer"] = True

        # 解析 document.xml
        from xml.etree import ElementTree as ET
        doc_xml = ET.fromstring(zf.read("word/document.xml"))
        ns = {"w": NS_W, "r": NS_R}

        body = doc_xml.find("w:body", ns)
        if body is None:
            result["errors"].append("document.xml 中缺少 body")

        # 段落统计 + 字体收集
        for p in body.findall("w:p", ns) if body is not None else []:
            result["paragraphs"] += 1
            for rf in p.iter(f"{{{NS_W}}}rFonts"):
                ea = rf.get(f"{{{NS_W}}}eastAsia")
                if ea:
                    result["fonts_used"].add(ea)

        # 页面设置
        sectPr = body.find("w:sectPr", ns) if body is not None else None
        if sectPr is not None:
            pgSz = sectPr.find("w:pgSz", ns)
            if pgSz is not None:
                result["page_setup"]["width"] = pgSz.get(f"{{{NS_W}}}w")
                result["page_setup"]["height"] = pgSz.get(f"{{{NS_W}}}h")
            pgMar = sectPr.find("w:pgMar", ns)
            if pgMar is not None:
                result["page_setup"]["margin_top"] = pgMar.get(f"{{{NS_W}}}top")
                result["page_setup"]["margin_bottom"] = pgMar.get(f"{{{NS_W}}}bottom")
                result["page_setup"]["margin_left"] = pgMar.get(f"{{{NS_W}}}left")
                result["page_setup"]["margin_right"] = pgMar.get(f"{{{NS_W}}}right")

        zf.close()
    except Exception as e:
        result["errors"].append(f"文件错误: {e}")

    result["fonts_used"] = sorted(result["fonts_used"])
    return result


# ═══════════════════════════════════════════════════════════════════
# 库函数接口
# ═══════════════════════════════════════════════════════════════════

def convert_md_to_docx(md_text, output_path, font_map=None):
    """库函数：直接转换 Markdown 文本为 DOCX

    Args:
        md_text: Markdown 文本字符串
        output_path: 输出 .docx 文件路径
        font_map: 字体映射 dict，为 None 时使用默认配置

    Returns:
        parsed_elements list
    """
    if font_map is None:
        font_map = load_font_map()
    parsed = parse_markdown(md_text)
    write_docx(parsed, font_map, output_path)
    return parsed


# ═══════════════════════════════════════════════════════════════════
# 命令行
# ═══════════════════════════════════════════════════════════════════

def print_font_scan():
    print("🔍 正在扫描系统字体...\n")
    all_fonts = scan_fonts()
    if not all_fonts:
        print("❌ 未扫描到任何字体")
        return

    recommendations = fuzzy_match_fonts(all_fonts)
    defaults = get_platform_defaults()

    print("字体配置（扫描推荐 + 当前备选）：\n")
    print(f"  {'角色':　<6} {'扫描匹配':　<20} {'备选（当前默认）':　<20}")
    print(f"  {'─'*6:　<6} {'─'*20:　<20} {'─'*20:　<20}")
    roles = ["标题", "主送单位", "正文", "一级标题", "二级标题", "三级标题", "附件", "落款"]
    for role in roles:
        matched = recommendations.get(role, "未匹配")
        default = defaults.get(role, "")
        marker = "✅" if matched != "请手动指定" else "  "
        print(f"  {marker} {role:　<4} {matched:　<20} {default:　<20}")

    print(f"\n系统共发现 {len(all_fonts)} 个中文字体")
    print(f"\n运行 --save-config 可将推荐配置保存到 {CONFIG_PATH}")
    print("（未匹配的角色将使用备选默认值）")


def save_config(font_map):
    config = {k: v for k, v in font_map.items()}
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"✅ 配置已保存到 {CONFIG_PATH}")


def print_verify(docx_path):
    print(f"🔍 验证 DOCX: {docx_path}\n")
    result = verify_docx(docx_path)

    if result["valid_zip"]:
        print("✅ 有效 ZIP/OOXML 包")
    else:
        print("❌ 无效 ZIP")

    print(f"\n部件清单 ({len(result['parts'])} 个):")
    for p in result["parts"]:
        print(f"  📄 {p}")

    print(f"\n段落数: {result['paragraphs']}")
    print(f"页脚页码: {'✅ 已生成' if result['has_footer'] else '❌ 缺失'}")
    print(f"使用字体: {', '.join(result['fonts_used']) if result['fonts_used'] else '(无)'}")

    if result["page_setup"]:
        ps = result["page_setup"]
        print(f"页面尺寸: {ps.get('width', '?')}×{ps.get('height', '?')} twips")
        print(f"页边距: 上={ps.get('margin_top')} 下={ps.get('margin_bottom')} "
              f"左={ps.get('margin_left')} 右={ps.get('margin_right')} twips")

    if result["errors"]:
        print(f"\n⚠️  问题 ({len(result['errors'])} 个):")
        for e in result["errors"]:
            print(f"  ❌ {e}")
    else:
        print("\n✅ 无结构性问题")


def main():
    args = sys.argv[1:]

    # 特殊选项
    if "--scan-fonts" in args:
        print_font_scan()
        return

    if "--save-config" in args:
        font_map = load_font_map(None)
        save_config(font_map)
        return

    # 解析参数
    config_file = None
    positional = []
    do_verify = False

    i = 0
    while i < len(args):
        if args[i] == "--config" and i + 1 < len(args):
            config_file = args[i + 1]
            i += 2
        elif args[i] == "--verify":
            do_verify = True
            i += 1
        elif args[i].startswith("--"):
            print(f"未知选项: {args[i]}")
            print("用法: python3 md2gongwen.py 输入.md [输出.docx] [选项]")
            print("选项: --scan-fonts, --save-config, --config FILE, --verify")
            sys.exit(1)
        else:
            positional.append(args[i])
            i += 1

    if not positional:
        print("用法: python3 md2gongwen.py 输入.md [输出.docx] [选项]")
        print()
        print("选项:")
        print("  --scan-fonts       扫描系统字体并打印推荐配置")
        print("  --save-config      保存字体配置到 ~/.gongwen_fonts.json")
        print("  --config FILE      使用指定字体配置文件")
        print("  --verify           生成后验证 DOCX 结构")
        print()
        print("示例:")
        print("  python3 md2gongwen.py 通知.md")
        print("  python3 md2gongwen.py 通知.md 输出.docx --verify")
        print("  python3 md2gongwen.py --scan-fonts")
        print("  python3 md2gongwen.py --save-config")
        sys.exit(1)

    input_path = positional[0]
    output_path = positional[1] if len(positional) > 1 else input_path.rsplit(".", 1)[0] + ".docx"

    if not os.path.exists(input_path):
        print(f"❌ 文件不存在: {input_path}")
        sys.exit(1)

    # 加载字体配置
    font_map = load_font_map(config_file)

    # 读取 Markdown
    with open(input_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    # 解析 & 生成
    print(f"📄 输入: {input_path}")
    print(f"📝 输出: {output_path}")
    font_roles = ["标题", "正文", "一级标题", "二级标题", "三级标题", "附件", "落款"]
    print(f"🔤 字体: {' | '.join(f'{r}={font_map.get(r)}' for r in font_roles)}")
    print()

    parsed = parse_markdown(md_text)
    write_docx(parsed, font_map, output_path)

    print(f"✅ 转换完成 → {output_path}")
    print(f"   (零依赖 — 仅用 Python 标准库)")

    # 统计
    counts = {}
    for typ, _, _ in parsed:
        counts[typ] = counts.get(typ, 0) + 1
    has_sig = any(t in ("sig_unit", "sig_date") for t, _, _ in parsed)
    has_attachment = any(t == "attachment_header" for t, _, _ in parsed)
    has_footer = True  # 总是生成

    _h1 = counts.get('h1', 0)
    _h2 = counts.get('h2', 0)
    _h3 = counts.get('h3', 0)
    _h4 = counts.get('h4', 0)
    print(f"   段落: {len([p for p in parsed if p[0] != 'blank'])} "
          f"(标题{_h1} 一级{_h2} 二级{_h3} 三级{_h4} "
          f"落款{'✅' if has_sig else '—'} "
          f"附件{'✅' if has_attachment else '—'} 页码✅)")

    if platform.system() == "Darwin":
        if "songti" in font_map.get("标题", "").lower():
            print(f"   ⚠️  小标宋字体在 macOS 上不可用，已用宋体({font_map['标题']})替代。")

    if do_verify:
        print()
        print_verify(output_path)


if __name__ == "__main__":
    main()
