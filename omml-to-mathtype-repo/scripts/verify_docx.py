# -*- coding: utf-8 -*-
"""Verify an OMML->MathType conversion and optionally export a PDF for spot checks.

Usage:
  python verify_docx.py <docx> [-o out.pdf]
Prints:
  oMath          must be 0
  OLEObject      should equal (previous MathType objects + converted count)
  leftover \\[ or $ markers  must be 0
"""
import sys, io, os, re, zipfile


def counts(path):
    with zipfile.ZipFile(path) as z:
        x = z.read('word/document.xml').decode('utf-8')
    return {
        'oMath': len(re.findall(r'<m:oMath[ >]', x)),
        'OLEObject': len(re.findall(r'<o:OLEObject', x)),
        'leftover_bracket': x.count(chr(92) + '['),
        'leftover_dollar': x.count('$'),
    }


def main():
    args = sys.argv[1:]
    path = args[0]
    pdf = None
    if '-o' in args:
        pdf = args[args.index('-o') + 1]
    c = counts(path)
    print("file:", path)
    for k, v in c.items():
        print("  %-17s %d" % (k, v))
    ok = c['oMath'] == 0 and c['leftover_bracket'] == 0 and c['leftover_dollar'] == 0
    print("PASS" if ok else "NOT CLEAN")
    if pdf:
        import win32com.client as wc
        app = wc.Dispatch('Word.Application')
        app.DisplayAlerts = 0
        d = app.Documents.Open(os.path.abspath(path), AddToRecentFiles=False)
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        pages = d.ComputeStatistics(2)
        d.Close(0)
        print("pdf:", pdf, "pages:", pages)


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    main()
