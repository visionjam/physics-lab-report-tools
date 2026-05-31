"""
将 Word 文档中 $...$ 和 $$...$$ 的 LaTeX 源码转换为 OMML 公式。
用法: python convert_latex_in_docx.py <输入.docx> [输出.docx]
"""
import sys, re
from lxml import etree
from docx import Document
from latex2mathml import converter as latex2mathml
from mathml2omml import convert as mathml2omml

NSW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSM = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML = '{http://www.w3.org/XML/1998/namespace}'
W = '{' + NSW + '}'
M = '{' + NSM + '}'

etree.register_namespace('m', NSM)

def tex2omml(tex, display=False):
    """LaTeX → OMML element"""
    mathml = latex2mathml.convert(tex, display="block" if display else "inline")
    raw = mathml2omml(mathml)
    raw = raw.replace('<m:oMath>', '<m:oMath xmlns:m="' + NSM + '">', 1)
    return etree.fromstring(raw.encode())

def split_fragments(text):
    """[(type, content)]  type in ('text','inline','display')"""
    parts, i = [], 0
    while i < len(text):
        if text[i:i+2] == '$$':
            j = text.find('$$', i+2)
            if j != -1:
                parts.append(('display', text[i+2:j])); i = j+2; continue
        if text[i] == '$':
            j = text.find('$', i+1)
            if j != -1 and text[i:i+2] != '$$':
                c = text[i+1:j]
                parts.append(('inline', c) if c.strip() else ('text', '$'))
                i = j+1; continue
        j = text.find('$', i)
        if j == -1: parts.append(('text', text[i:])); break
        else: parts.append(('text', text[i:j])); i = j
    return parts

def make_run(text):
    r = etree.SubElement(etree.Element('dummy'), W + 'r')
    t = etree.SubElement(r, W + 't')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set(XML + 'space', 'preserve')
    return r

def process_paragraph(para):
    """返回 True 表示有修改"""
    full = ''.join(t.text or '' for t in para.iter(W + 't'))
    if '$' not in full: return False

    parts = split_fragments(full)
    if not any(t != 'text' for t, _ in parts): return False

    ppr = para.find(W + 'pPr')
    for child in list(para):
        if child.tag != W + 'pPr': para.remove(child)

    for typ, content in parts:
        if typ == 'text':
            if content: para.append(make_run(content))
        else:
            try:
                elem = tex2omml(content, typ == 'display')
                if typ == 'display':
                    wp = etree.SubElement(para, M + 'oMathPara')
                    wp.append(elem)
                else:
                    para.append(elem)
            except Exception:
                d = '$$' if typ == 'display' else '$'
                para.append(make_run(d + content + d))
    return True

def main():
    if len(sys.argv) < 2:
        print('用法: python convert_latex_in_docx.py <输入.docx> [输出.docx]', file=sys.stderr)
        sys.exit(1)

    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else inp.replace('.docx', '_公式版.docx')
    if out == inp: out = inp.replace('.docx', '_公式版.docx')

    doc = Document(inp)
    total = mod = 0
    for para in doc.element.iter(W + 'p'):
        total += 1
        if process_paragraph(para): mod += 1

    # 处理页眉页脚
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            for para in hf.paragraphs:
                if '$' in para.text and process_paragraph(para._element): mod += 1

    doc.save(out)
    print(f'完成: 处理段落 {total}，修改 {mod}，输出 → {out}')

if __name__ == '__main__':
    main()
