#!/usr/bin/env python3
"""
GB/T 9704-2012 Markdown to DOCX Converter v2

Converts Markdown documents to DOCX following Chinese official document standards.
Usage: python md2docx.py input.md output.docx

字体跨平台适配：
- Windows: 方正小标宋简体, 仿宋_GB2312, 黑体, 楷体_GB2312
- Mac: 方正小标宋简体, 仿宋, STHeiti, 楷体
- Linux: 使用可用字体（如无则使用回退字体）
"""

import sys
import re
import platform
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
except ImportError:
    print("Error: python-docx is required. Install with: pip install python-docx")
    sys.exit(1)


# ==================== 字体配置 ====================
def get_available_fonts():
    """检测系统可用字体"""
    fonts = {
        'title': ['方正小标宋简体', '方正小标宋', 'SimHei', '黑体'],
        'body': ['仿宋_GB2312', '仿宋', 'FangSong', 'FangSong_GB2312'],
        'heading1': ['黑体', 'SimHei', 'Heiti'],
        'heading2': ['楷体_GB2312', '楷体', 'Kaiti', 'STKaiti'],
    }
    
    # 检查系统类型并调整
    system = platform.system()
    
    # 如果是 Linux，尝试使用 fc-list
    if system == 'Linux':
        try:
            import subprocess
            result = subprocess.run(['fc-list', ':lang=zh', '-f', '%{family}\n'],
                                    capture_output=True, text=True, timeout=5)
            available = set()
            for line in result.stdout.split('\n'):
                for families in line.split(','):
                    available.add(families.strip())
            
            for category in fonts:
                fonts[category] = [f for f in fonts[category] if f in available] or fonts[category]
        except Exception:
            pass
    
    return fonts

# 全局字体配置
FONTS = get_available_fonts()

def get_font( category):
    """获取指定类别的首选字体"""
    return FONTS.get(category, FONTS['body'])[0]


# ==================== GB/T 9704-2012 标准配置 ====================
PAGE_MARGIN = {
    "top": 3.7,    # cm
    "bottom": 3.5,  # cm
    "left": 2.7,    # cm (27mm)
    "right": 2.7,   # cm (27mm)
}

# 字号对照：二号=22pt，三号=16pt
LINE_SPACING = 28  # 固定值28磅
CHAR_INDENT = 32    # 2字符缩进 ≈ 9pt (GB/T 9704-2012: 2字符 ≈ 3.17mm ≈ 9pt)

# 字体配置
FONT_SIZES = {
    "title": 22,       # 2号 = 22pt
    "heading1": 16,    # 3号 = 16pt
    "heading2": 16,    # 3号 = 16pt
    "heading3": 16,    # 3号 = 16pt
    "heading4": 16,    # 3号 = 16pt
    "body": 16,         # 3号 = 16pt
}


def set_page_setup(doc):
    """设置A4页面及边距"""
    section = doc.sections[0]
    section.page_height = Cm(29.7)  # A4高度
    section.page_width = Cm(21.0)   # A4宽度
    section.top_margin = Cm(PAGE_MARGIN["top"])
    section.bottom_margin = Cm(PAGE_MARGIN["bottom"])
    section.left_margin = Cm(PAGE_MARGIN["left"])
    section.right_margin = Cm(PAGE_MARGIN["right"])


def apply_paragraph_style(para, font_name, font_size, bold=False,
                          align='left', indent=0, line_spacing=LINE_SPACING,
                          space_before=0, space_after=0):
    """应用统一的段落样式"""
    # 设置段落对齐
    align_map = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    para.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    
    # 设置段落格式
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line_spacing)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    
    # 首行缩进（pt 转为 twips: 1pt = 20twips）
    if indent > 0:
        pf.first_line_indent = Pt(indent)
    
    # 设置字体
    for run in para.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        # 中文字体
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    return para


def add_empty_paragraph(doc):
    """添加空行（用于标题与正文之间的空行）"""
    para = doc.add_paragraph()
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    para.paragraph_format.line_spacing = Pt(LINE_SPACING)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    return para


def parse_markdown(content):
    """
    解析Markdown内容，识别标题层级。
    返回: [(level, text, ptype), ...]
    level: 0=标题, 1-4=正文层次, None=正文
    ptype: title, heading1, heading2, heading3, heading4, body
    """
    lines = content.split('\n')
    result = []
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # 主标题 (# 标题)
        if stripped.startswith('# ') and len(stripped) > 2:
            result.append((0, stripped[2:].strip(), 'title'))
        # 二级标题 (##)
        elif stripped.startswith('## ') and not stripped.startswith('### '):
            text = stripped[3:].strip()
            # 检查标题前缀
            if re.match(r'^[一二三四五六七八九十]+、', text):
                result.append((1, text, 'heading1'))
            elif re.match(r'^（[一二三四五六七八九十]+）', text) or re.match(r'^\([一二三四五六七八九十]+\)', text):
                result.append((2, text, 'heading2'))
            elif re.match(r'^\d+\.', text):
                result.append((3, text, 'heading3'))
            elif re.match(r'^（\d+）', text) or re.match(r'^\(\d+\)', text):
                result.append((4, text, 'heading4'))
            else:
                result.append((1, text, 'heading1'))
        # 三级标题 (###)
        elif stripped.startswith('### '):
            text = stripped[4:].strip()
            if re.match(r'^（[一二三四五六七八九十]+）', text):
                result.append((2, text, 'heading2'))
            elif re.match(r'^\d+\.', text):
                result.append((3, text, 'heading3'))
            else:
                result.append((None, text, 'body'))
        # 四级标题 (####)
        elif stripped.startswith('#### '):
            text = stripped[5:].strip()
            if re.match(r'^（\d+）', text):
                result.append((4, text, 'heading4'))
            else:
                result.append((None, text, 'body'))
        else:
            # 普通文本，检查是否为一、二、三、四级标题
            if re.match(r'^[一二三四五六七八九十]+、', stripped):
                result.append((1, stripped, 'heading1'))
            elif re.match(r'^（[一二三四五六七八九十]+）', stripped) or re.match(r'^\([一二三四五六七八九十]+\)', stripped):
                result.append((2, stripped, 'heading2'))
            elif re.match(r'^\d+\.', stripped):
                result.append((3, stripped, 'heading3'))
            elif re.match(r'^（\d+）', stripped) or re.match(r'^\(\d+\)', stripped):
                result.append((4, stripped, 'heading4'))
            else:
                result.append((None, stripped, 'body'))
    
    return result


def create_styled_paragraph(doc, text, ptype):
    """根据段落类型创建并格式化段落"""
    para = doc.add_paragraph()
    run = para.add_run(text)
    
    if ptype == 'title':
        # 主标题：居中，方正小标宋简体，2号（22pt），不加粗
        font_name = get_font('title')
        apply_paragraph_style(para, font_name, FONT_SIZES["title"],
                             bold=False, align='center',
                             indent=0, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    elif ptype == 'heading1':
        # 一级标题：一、黑体，3号（16pt），加粗，首行缩进2字符
        font_name = get_font('heading1')
        apply_paragraph_style(para, font_name, FONT_SIZES["heading1"],
                             bold=True, align='left',
                             indent=CHAR_INDENT, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    elif ptype == 'heading2':
        # 二级标题：（一）楷体，3号（16pt），加粗，首行缩进2字符
        font_name = get_font('heading2')
        apply_paragraph_style(para, font_name, FONT_SIZES["heading2"],
                             bold=True, align='left',
                             indent=CHAR_INDENT, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    elif ptype == 'heading3':
        # 三级标题：1. 仿宋，3号（16pt），加粗，首行缩进2字符
        font_name = get_font('body')
        apply_paragraph_style(para, font_name, FONT_SIZES["heading3"],
                             bold=True, align='left',
                             indent=CHAR_INDENT, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    elif ptype == 'heading4':
        # 四级标题：（1）仿宋，3号（16pt），不加粗，首行缩进2字符
        font_name = get_font('body')
        apply_paragraph_style(para, font_name, FONT_SIZES["heading4"],
                             bold=False, align='left',
                             indent=CHAR_INDENT, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    else:  # body
        # 正文：仿宋，3号（16pt），两端对齐，首行缩进2字符
        font_name = get_font('body')
        apply_paragraph_style(para, font_name, FONT_SIZES["body"],
                             bold=False, align='justify',
                             indent=CHAR_INDENT, line_spacing=LINE_SPACING,
                             space_before=0, space_after=0)
    
    return para


def add_page_numbers(doc):
    """
    添加页码（GB/T 9704-2012 标准）
    - 4号半角宋体（Times New Roman）14pt
    - 格式：— 页码 —（带边框线）
    - 单页居右，双页居左
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    section = doc.sections[0]
    
    # 启用奇偶页不同的页眉页脚
    section.different_first_page_header_footer = True
    
    def make_page_num_para(footer, align):
        """创建页码段落"""
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        para.clear()
        para.alignment = align
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        return para
    
    # 偶数页（左对齐）
    footer_even = section.even_page_footer
    footer_even.is_linked_to_previous = False
    p_even = make_page_num_para(footer_even, WD_ALIGN_PARAGRAPH.LEFT)
    run_even = p_even.add_run(' \u2014 1 \u2014 ')
    run_even.font.name = 'Times New Roman'
    run_even.font.size = Pt(14)
    run_even._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    
    # 奇数页（右对齐）
    footer_primary = section.footer
    footer_primary.is_linked_to_previous = False
    p_primary = make_page_num_para(footer_primary, WD_ALIGN_PARAGRAPH.RIGHT)
    run_primary = p_primary.add_run(' \u2014 1 \u2014 ')
    run_primary.font.name = 'Times New Roman'
    run_primary.font.size = Pt(14)
    run_primary._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')


def convert_markdown_to_docx(input_path, output_path):
    """主转换函数"""
    # 读取输入文件
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 创建文档
    doc = Document()
    
    # 设置页面
    set_page_setup(doc)
    
    # 解析Markdown
    parsed = parse_markdown(content)
    
    # 处理段落
    prev_was_title = False
    prev_was_heading = False
    
    for level, text, ptype in parsed:
        if not text.strip():
            continue
        
        # 主标题之后添加空行
        if prev_was_title and ptype != 'title':
            add_empty_paragraph(doc)
        
        # 创建段落
        para = create_styled_paragraph(doc, text, ptype)
        
        # 更新状态
        prev_was_title = (ptype == 'title')
        prev_was_heading = ptype in ('heading1', 'heading2', 'heading3', 'heading4')
    
    # 添加页码
    add_page_numbers(doc)
    
    # 保存文档
    doc.save(output_path)
    print(f"✓ 已转换: {input_path}")
    print(f"  → {output_path}")
    print(f"  字体配置: {get_font('title')} / {get_font('body')}")


def main():
    if len(sys.argv) < 3:
        print("用法: python md2docx.py <输入.md> <输出.docx>")
        print("示例: python md2docx.py input.md output.docx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not Path(input_file).exists():
        print(f"错误: 找不到输入文件: {input_file}")
        sys.exit(1)
    
    convert_markdown_to_docx(input_file, output_file)


if __name__ == '__main__':
    main()
