#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx —— 把 Markdown 转成 Word(.docx)，零第三方依赖，只用 Python 标准库。

为什么要它：手机上写完的笔记、AI 导出的 Markdown、GitHub 上的 README，
想发给领导/客户时对方只认 Word。在线转换要登录、要限页数、还可能把格式弄乱；
Word 另存为又得装 Office。这个脚本直接手写 OOXML，一个文件搞定。

支持：
  标题 # ~ ######（带大纲级别，Word 导航窗格能跳转）
  正文段落、空行分段
  有序列表 1. / 1)   无序列表 - * +（支持缩进层级）
  表格 | a | b |（首行为表头，自动加底色、等宽分列）
  代码块 ```（等宽字体 + 浅灰底，保留缩进）
  引用 >（左缩进 + 灰色斜体）
  行内 **粗体**、*斜体*、`等宽`
  链接 [文字](地址)（保留文字，把地址附在后面）
  分隔线 ---（转成分页符）

用法：
  python3 md2docx.py 笔记.md                  # 生成 笔记.docx
  python3 md2docx.py 笔记.md -o 汇报.docx
  python3 md2docx.py 笔记.md --font 微软雅黑 --size 11
  python3 md2docx.py 笔记.md --landscape      # 横向（宽表格用）
  python3 md2docx.py 笔记.md --title "周报" --author "Marco"
  python3 md2docx.py --check 笔记.md          # 只体检不写盘：行数/字数/表格数

纯标准库，Windows/macOS/Linux 都能跑，Python 3.8+。
"""
import argparse
import os
import re
import sys
import zipfile
from datetime import datetime, timezone

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def _disp_len(s):
    # 中文/全角算2个宽度，英文数字算1个，用来分配表格列宽
    return sum(2 if ord(c) > 0x2E7F else 1 for c in s)


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))


def styles_xml(font, size_half_pt, mono):
    """默认字体 + 6级标题样式（带 outlineLvl，导航窗格可跳）。"""
    heads = [(1, 40, True), (2, 32, True), (3, 28, True), (4, 26, True), (5, 24, True), (6, 22, True)]
    head_styles = []
    for lvl, sz, bold in heads:
        head_styles.append(
            '<w:style w:type="paragraph" w:styleId="Heading%d">'
            '<w:name w:val="heading %d"/><w:basedOn w:val="Normal"/>'
            '<w:next w:val="Normal"/><w:qFormat/>'
            '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>'
            '<w:outlineLvl w:val="%d"/></w:pPr>'
            '<w:rPr>%s<w:sz w:val="%d"/><w:color w:val="1F2A37"/></w:rPr>'
            '</w:style>' % (lvl, lvl, lvl - 1, '<w:b/>' if bold else '', sz))
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>
<w:sz w:val="%d"/><w:szCs w:val="%d"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="300" w:lineRule="auto" w:after="60"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
%s
<w:style w:type="character" w:styleId="VerbatimChar"><w:name w:val="Verbatim Char"/>
<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/></w:rPr></w:style>
</w:styles>''' % (NS_W, font, font, font, font, size_half_pt, size_half_pt,
                  '\n'.join(head_styles), mono, mono, mono, size_half_pt - 2)


CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>%s</dc:title><dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>'''

APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>md2docx</Application></Properties>'''

INLINE = re.compile(r'(\*\*.+?\*\*|__.+?__|\*[^*\n]+?\*|`[^`\n]+?`|\[[^\]\n]*\]\([^)\n]*\))')


def runs(text, base_bold=False, mono_style=False, color=None, italic=False):
    """把行内标记拆成 w:r 片段。返回 XML 字符串（多个 run）。"""
    out = []
    pos = 0
    for m in INLINE.finditer(text):
        if m.start() > pos:
            out.append(_run(text[pos:m.start()], base_bold, mono_style, color, italic))
        tok = m.group(0)
        if tok.startswith('**') or tok.startswith('__'):
            out.append(_run(tok[2:-2], True, mono_style, color, italic))
        elif tok.startswith('`'):
            out.append(_run(tok[1:-1], base_bold, True, color, italic))
        elif tok.startswith('['):
            mm = re.match(r'\[([^\]]*)\]\(([^)]*)\)', tok)
            label, url = mm.group(1), mm.group(2)
            out.append(_run(label, base_bold, mono_style, color, italic))
            if url.startswith('http'):
                out.append(_run('（%s）' % url, False, False, '5B6573', italic))
        else:
            out.append(_run(tok[1:-1], base_bold, mono_style, color, True))
        pos = m.end()
    if pos < len(text):
        out.append(_run(text[pos:], base_bold, mono_style, color, italic))
    return ''.join(out) or _run('', base_bold, mono_style, color, italic)


def _run(text, bold=False, mono=False, color=None, italic=False):
    rpr = ['<w:rPr>']
    if mono:
        rpr.append('<w:rStyle w:val="VerbatimChar"/><w:rFonts w:ascii="Consolas" '
                   'w:hAnsi="Consolas" w:eastAsia="Consolas"/>')
    if bold:
        rpr.append('<w:b/>')
    if italic:
        rpr.append('<w:i/>')
    if color:
        rpr.append('<w:color w:val="%s"/>' % color)
    rpr.append('</w:rPr>')
    body = esc(text).replace('\t', '    ')
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (''.join(rpr), body)


def para(inner, style=None, indent=None, shading=None, align=None, space_after=None,
         break_before=False):
    ppr = ['<w:pPr>']
    if style:
        ppr.append('<w:pStyle w:val="%s"/>' % style)
    if break_before:
        ppr.append('<w:pageBreakBefore/>')
    if shading:
        ppr.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shading)
    if indent:
        left = indent[0] if isinstance(indent, tuple) else indent
        hang = indent[1] if isinstance(indent, tuple) else 0
        ppr.append('<w:ind w:left="%d" w:hanging="%d"/>' % (left, hang))
    if align:
        ppr.append('<w:jc w:val="%s"/>' % align)
    if space_after is not None:
        ppr.append('<w:spacing w:after="%d"/>' % space_after)
    ppr.append('</w:pPr>')
    return '<w:p>%s%s</w:p>' % (''.join(ppr), inner)


def table_xml(rows, total_width):
    n = max(len(r) for r in rows)
    # 列宽按每列最长内容分配，最短给 0.8 份，避免窄列被挤成一条线
    lens = []
    for j in range(n):
        mx = max(_disp_len(r[j]) if j < len(r) else 0 for r in rows)
        lens.append(max(mx, 3))
    unit = int(total_width / max(sum(lens), 1))
    col_width = [max(int(l * unit), 720) for l in lens]
    borders = ''.join(
        '<w:%s w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>' % side
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    grid = ''.join('<w:gridCol w:w="%d"/>' % col_width[j] for j in range(n))
    trs = []
    for i, row in enumerate(rows):
        cells = []
        for j in range(n):
            txt = row[j] if j < len(row) else ''
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="F3F5F7"/>' if i == 0 else ''
            cells.append(
                '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>'
                '<w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
                '</w:tcPr>%s</w:tc>'
                % (col_width[j], shd, para(runs(txt, base_bold=(i == 0)))))
        trs.append('<w:tr>%s</w:tr>' % ''.join(cells))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>%s</w:tblBorders>'
            '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
            % (borders, grid, ''.join(trs)))


def convert(md, opt):
    lines = md.splitlines()
    body = []
    i = 0
    n = len(lines)
    text_width = 11906 - 2 * opt.margin if not opt.landscape else 16838 - 2 * opt.margin
    total_width = text_width
    stats = {'table': 0, 'code': 0, 'head': 0, 'list': 0}

    while i < n:
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if stripped.startswith('```'):
            stats['code'] += 1
            fence = stripped[:3]
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith(fence):
                buf.append(lines[i])
                i += 1
            i += 1
            if not buf:
                body.append(para(_run(' ', False, True, None, False), shading='F6F8FA'))
            for cl in buf:
                body.append(para(_run(cl if cl else ' ', False, True, '333B45', False),
                                 shading='F6F8FA', indent=240, space_after=0))
            body.append(para('', space_after=120))
            continue

        # 表格：| a | b | 且下一行是 |---|
        if stripped.startswith('|') and i + 1 < n and re.match(r'^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$', lines[i + 1]):
            stats['table'] += 1
            rows = []
            header = [c.strip() for c in stripped.strip('|').split('|')]
            rows.append(header)
            i += 2
            while i < n and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                rows.append(cells)
                i += 1
            body.append(table_xml(rows, total_width))
            body.append(para('', space_after=120))
            continue

        if re.match(r'^(---|___|\*\*\*)\s*$', stripped):
            body.append(para('', break_before=True))
            i += 1
            continue

        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            stats['head'] += 1
            lvl = len(m.group(1))
            body.append(para(runs(m.group(2).strip()), style='Heading%d' % lvl))
            i += 1
            continue

        m = re.match(r'^(>)\s?(.*)$', stripped)
        if m:
            body.append(para(runs(m.group(2), color='5B6573', italic=True), indent=360,
                             space_after=40))
            i += 1
            continue

        m = re.match(r'^(\s*)([-*+])\s+(.*)$', line)
        if m:
            stats['list'] += 1
            depth = min(len(m.group(1)) // 2, 4)
            body.append(para(_run('•  ') + runs(m.group(3)),
                             indent=240 + depth * 240))
            i += 1
            continue

        m = re.match(r'^(\s*)(\d+)[.)]\s+(.*)$', line)
        if m:
            stats['list'] += 1
            depth = min(len(m.group(1)) // 2, 4)
            body.append(para(_run(m.group(2) + '.  ') + runs(m.group(3)),
                             indent=240 + depth * 240))
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        body.append(para(runs(stripped)))
        i += 1

    if not body:
        body.append(para(_run('（空文档）')))

    sect = ('<w:sectPr><w:pgSz w:w="%d" w:h="%d" w:orient="%s"/>'
            '<w:pgMar w:top="1440" w:right="%d" w:bottom="1440" w:left="%d" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
            % (11906 if not opt.landscape else 16838,
               16838 if not opt.landscape else 11906,
               'landscape' if opt.landscape else 'portrait',
               opt.margin, opt.margin))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="%s" xmlns:r="%s"><w:body>%s%s</w:body></w:document>'
           % (NS_W, NS_R, ''.join(body), sect))
    return doc, stats


def write_docx(path, doc_xml, opt, title, author):
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc_xml)
        z.writestr('word/styles.xml', styles_xml(opt.font, opt.size * 2, opt.mono))
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('docProps/core.xml', CORE % (esc(title), esc(author), esc(author), now, now))
        z.writestr('docProps/app.xml', APP)


def main():
    ap = argparse.ArgumentParser(description='Markdown 转 Word(.docx)，零依赖')
    ap.add_argument('input', help='Markdown 文件路径（.md/.markdown/.txt）')
    ap.add_argument('-o', '--output', help='输出 .docx 路径，默认同名 .docx')
    ap.add_argument('--font', default='等线', help='正文字体，默认 等线')
    ap.add_argument('--mono', default='Consolas', help='代码字体，默认 Consolas')
    ap.add_argument('--size', type=float, default=10.5, help='字号(pt)，默认 10.5')
    ap.add_argument('--margin', type=int, default=1134, help='页边距(twips)，默认 1134≈2cm')
    ap.add_argument('--landscape', action='store_true', help='横向页面')
    ap.add_argument('--title', default='', help='文档标题属性，默认取一级标题或文件名')
    ap.add_argument('--author', default='md2docx', help='作者属性')
    ap.add_argument('--check', action='store_true', help='只体检不写盘')
    opt = ap.parse_args()

    if not os.path.isfile(opt.input):
        print('找不到文件：%s' % opt.input)
        return 2
    raw = open(opt.input, 'rb').read()
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'gb18030', 'big5'):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if text is None:
        print('这个文件无法识别编码（试过 utf-8/gb18030/big5）')
        return 2

    doc_xml, st = convert(text, opt)
    title = opt.title
    if not title:
        m = re.search(r'^#\s+(.+)$', text, re.M)
        title = m.group(1).strip() if m else os.path.splitext(os.path.basename(opt.input))[0]

    chars = len(re.sub(r'\s', '', text))
    lines_n = len(text.splitlines())
    print('%s：%d行，%d字（不含空白），标题%d，表格%d，代码块%d，列表项%d'
          % (os.path.basename(opt.input), lines_n, chars, st['head'], st['table'],
             st['code'], st['list']))

    if opt.check:
        print('体检模式，未写盘。加 -o 或直接运行即可生成。')
        return 0

    out = opt.output or (os.path.splitext(opt.input)[0] + '.docx')
    write_docx(out, doc_xml, opt, title, opt.author)
    size_kb = os.path.getsize(out) / 1024.0
    print('已生成：%s（%.1f KB）' % (out, size_kb))
    print('用 Word/WPS 打开检查后再发人；首行缩进、字号、字体都在 styles.xml 里，改 --font/--size 即可。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
