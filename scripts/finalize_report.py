"""Fill the report's table of contents with page numbers and export a PDF copy.

Usage: python3 finalize_report.py <report.docx> <report.pdf>

docx-js writes the contents list as an empty TOC field, which only Word can fill in. This script
renders the document with LibreOffice, finds the page on which each Heading 1/2 appears and writes
those entries into the field's cached result, so the contents list shows in every viewer. Word
still refreshes the field when the document is opened (updateFields is set in settings.xml).
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

from lxml import etree

DOCX, PDF = sys.argv[1:3]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TOC_RE = re.compile(r'(<w:sdt><w:sdtPr><w:alias w:val="Contents"/></w:sdtPr><w:sdtContent>)(.*?)(</w:sdtContent></w:sdt>)', re.S)
TAB_POS = 9402  # right edge of the text block (A4 minus margins), in twentieths of a point


def read_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf8")


def write_docx(src, dst, xml):
    tmp = dst + ".tmp"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = xml.encode("utf8") if item.filename == "word/document.xml" else zi.read(item.filename)
            zo.writestr(item, data)
    os.replace(tmp, dst)


def headings(xml):
    body = etree.fromstring(xml.encode("utf8")).find(W + "body")
    out = []
    for p in body.iter(W + "p"):
        st = p.find(f"{W}pPr/{W}pStyle")
        if st is not None and st.get(W + "val") in ("Heading1", "Heading2"):
            text = "".join(t.text or "" for t in p.iter(W + "t"))
            out.append((1 if st.get(W + "val") == "Heading1" else 2, text))
    return out


def toc_xml(entries):
    paras = []
    for i, (lvl, text, page) in enumerate(entries):
        r = ""
        if i == 0:
            r += ('<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
                  '<w:r><w:instrText xml:space="preserve">TOC \\h \\o &quot;1-2&quot;</w:instrText></w:r>'
                  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        r += f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>{page}</w:t></w:r>'
        if i == len(entries) - 1:
            r += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        paras.append(f'<w:p><w:pPr><w:pStyle w:val="TOC{lvl}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB_POS}"/></w:tabs></w:pPr>{r}</w:p>')
    return "".join(paras)


def with_toc(xml, entries):
    new, n = TOC_RE.subn(lambda m: m.group(1) + toc_xml(entries) + m.group(3), xml, count=1)
    if n != 1:
        sys.exit("TOC field not found")
    return new


def render(docx, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, docx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    txt = subprocess.run(["pdftotext", "-layout", pdf, "-"], check=True, capture_output=True, text=True).stdout
    return pdf, txt.split("\f")


def norm(s):
    return re.sub(r"\s+", " ", s.replace(" ", " ")).strip()


def find_pages(heads, pages):
    lines = [[norm(ln) for ln in pg.splitlines() if ln.strip()] for pg in pages]
    start = next(i for i, ls in enumerate(lines) if "Abbreviations" in ls)  # first page after the contents
    out, cur = [], start
    for lvl, text in heads:
        h = norm(text)
        for i in range(cur, len(lines)):
            if any(ln == h or (len(ln) >= 20 and h.startswith(ln)) for ln in lines[i]):
                out.append((lvl, text, i + 1))
                cur = i
                break
        else:
            sys.exit(f"heading not found in PDF: {text}")
    return out


tmpdir = tempfile.mkdtemp(prefix="toc_")
try:
    base = read_xml(DOCX)
    heads = headings(base)
    work = os.path.join(tmpdir, "report.docx")
    entries = [(lvl, text, "00") for lvl, text in heads]  # placeholders take the same lines as the final entries
    for attempt in range(3):
        write_docx(DOCX, work, with_toc(base, entries))
        pdf, pages = render(work, tmpdir)
        found = find_pages(heads, pages)
        if found == entries:
            break
        entries = found
    else:
        sys.exit("page numbers did not settle")
    write_docx(DOCX, DOCX, with_toc(base, entries))
    shutil.copyfile(pdf, PDF)
    print(f"{len(entries)} contents entries; {len(pages) - 1} pages; wrote {DOCX} and {PDF}")
finally:
    shutil.rmtree(tmpdir, ignore_errors=True)
