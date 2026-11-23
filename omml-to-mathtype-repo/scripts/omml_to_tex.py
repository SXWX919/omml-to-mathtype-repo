# -*- coding: utf-8 -*-
"""Replace every OMML equation in a .docx with LaTeX text ready for MathType Toggle TeX.

Display equations -> \\[...\\]   Inline equations -> $...$
Bare all-caps payloads (MMF, DSS, ...) are wrapped in \\mathrm{} because MathType
refuses them in raw form.

Usage: python omml_to_tex.py <src.docx> <dst.docx>
"""
import sys, io, os, re, zipfile, shutil
import lxml.etree as ET
import pypandoc

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def extract_segments(docx):
    tex = pypandoc.convert_file(docx, 'latex', format='docx')
    raw = re.findall(r'\\\[.*?\\\]|\\\(.*?\\\)', tex, re.S)
    out = []
    for s in raw:
        if s.startswith('\\(') and s.endswith('\\)'):
            out.append('$' + s[2:-2] + '$')
        else:
            out.append(s)
    return out


def fix_bare_caps(seg):
    if seg.startswith('$') and seg.endswith('$'):
        inner, open_c, close_c = seg[1:-1], '$', '$'
    elif seg.startswith('\\[') and seg.endswith('\\]'):
        inner, open_c, close_c = seg[2:-2], '\\[', '\\]'
    else:
        return seg
    core = inner.strip()
    if re.fullmatch(r'[A-Z]{2,}', core):
        return open_c + '\\mathrm{' + core + '}' + close_c
    return seg


def main():
    src, dst = sys.argv[1], sys.argv[2]
    shutil.copyfile(src, dst)

    segs = [fix_bare_caps(s) for s in extract_segments(src)]
    print("pandoc segments:", len(segs),
          "| display:", sum(1 for s in segs if s.startswith('\\[')))

    with zipfile.ZipFile(src) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    omaths = list(root.iter('{%s}oMath' % M))
    print("omaths:", len(omaths))
    if len(omaths) != len(segs):
        raise SystemExit("equation count mismatch: %d oMath vs %d pandoc segments"
                         % (len(omaths), len(segs)))

    for om, tex in zip(omaths, segs):
        parent = om.getparent()
        target = parent if parent.tag == '{%s}oMathPara' % M else om
        r = ET.Element('{%s}r' % W)
        t = ET.SubElement(r, '{%s}t' % W)
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = tex
        grand = target.getparent()
        idx = grand.index(target)
        grand.remove(target)
        grand.insert(idx, r)

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        data = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == 'word/document.xml'
                          else zin.read(item.filename))
    print("written:", dst, os.path.getsize(dst), "bytes")

    with zipfile.ZipFile(dst) as z:
        x = z.read('word/document.xml').decode('utf-8')
    print("oMath left:", len(re.findall(r'<m:oMath[ >]', x)),
          "| \\[ :", x.count('\\['), "| $ :", x.count('$'))


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    main()
