#!/usr/bin/env python3
"""A foreign arbiter over the produced .docx.

Every other axis in this repository judges the conversion with the same hand that
writes it: `tests/corpus.py` reads the package with the same code family as the
reader, and `officecli validate` is the host's own checker. A construct the
reader drops is then dropped on both sides of the comparison, and the axis stays
green while the output is short.

This one puts implementations that are not ours on both ends:

  source  .odt  → odfpy 1.4.1        (a third-party ODF implementation)
  output  .docx → python-docx 1.2.0  (a third-party OOXML implementation)

and a second, unrelated opinion on the package where the host has one: Apple's
Office.mdimporter (`mdimport`), macOS only. Where it is absent the check says so
and does not hold anything against the package.

What is compared is the block sequence — paragraph, heading with its level, list
item with its depth, table — in order. A block may go or change shape only when
the reader's own report declared it, by the same two rules the block axis uses in
`corpus.py`:

  * declared dropped at block level → no output block is expected for it;
  * declared "not converted; its text is kept" → it comes out flattened, so its
    items are body paragraphs rather than list items.

Pictures are counted and printed but not gated. What a frame carries is not a
text block, and the block axis in `corpus.py` already matches pictures in order
with the reader's declared drops excused.

Requirements: `pip install odfpy python-docx`. The reader itself needs neither —
they are needed only to read both ends of this check.

Usage:

    pip install odfpy python-docx
    python3 tests/corpus/fetch.py
    OFFICECLI_BIN=/path/to/officecli python3 tests/foreign_axis.py [name.odt ...]

Exit status is 1 when any produced document's block sequence disagrees with the
source and the reader did not declare the difference, and 0 otherwise.
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PLUGIN = os.path.join(ROOT, "plugin")
FILES = os.path.join(HERE, "corpus", "files")

OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

# Containers the source walk descends into. Not frames or text boxes: the reader
# declares their payload dropped, so their contents are not blocks of the output.
CONTAINERS = ("list", "list-item", "list-header", "section", "index-body",
              "table-of-content")
SKIP_TEXT = ("frame", "text-box")

# A refusal the corpus records as clean: the file is not well-formed XML and the
# reader says so. Named here so an unexpected second refusal still shows up.
REFUSAL = ("not well-formed", "not a readable", "not an OpenDocument", "not an ODF",
           "no office:text body", "has no content.xml")


def binary():
    b = os.environ.get("OFFICECLI_BIN")
    if not b:
        sys.exit("set OFFICECLI_BIN to the officecli binary to test against")
    return b


ENV = dict(os.environ)
ENV.update({
    "OFFICECLI_NO_AUTO_INSTALL": "1",
    "OFFICECLI_NO_AUTO_RESIDENT": "1",
    "OFFICECLI_SKIP_UPDATE": "1",
    "OFFICECLI_PLUGIN_DUMP_READER_ODT": PLUGIN,
})


def need(module, package):
    try:
        return __import__(module)
    except ImportError:
        sys.exit("%s is needed for the foreign arbiter: pip install %s" % (package, package))


# ---------------------------------------------------------------- source side

def declarations(odt):
    """What the reader itself says it did not convert.

    Read out of the report before the source is walked, so that an absence the
    reader announced is not counted against it. Two shapes matter: a construct
    dropped at block level expects no output block at all, and a construct whose
    text is kept comes out flattened rather than as a list.
    """
    r = subprocess.run([PLUGIN, "dump", odt], capture_output=True, text=True)
    dropped, flat = set(), set()
    for line in r.stderr.splitlines():
        m = re.match(r"\[warning\] odt: skipped \d+ x (\S+) (.*)", line)
        if not m:
            continue
        name = m.group(1).split(":")[-1]
        rest = m.group(2)
        if "at block level (dropped)" in rest:
            dropped.add(name)
        elif "not converted; its text is kept" in rest:
            flat.add(name)
    return dropped, flat


def source_blocks(odt, dropped, flat):
    odf = need("odf", "odfpy")
    from odf.opendocument import load
    from odf import draw as DR

    def text_of(el):
        out = []
        for n in el.childNodes:
            q = getattr(n, "qname", None)
            if q and q[1] in SKIP_TEXT:
                continue          # the reader declares a frame's payload dropped
            d = getattr(n, "data", None)
            if isinstance(d, str):
                out.append(d)
            elif hasattr(n, "childNodes"):
                out.append(text_of(n))
        return "".join(out)

    def depth_of(el):
        if "list" in flat:
            # Declared not converted, its text kept: its items are body
            # paragraphs in the output, so they are counted at that depth here.
            return 0
        d, n = 0, getattr(el, "parentNode", None)
        while n is not None:
            if getattr(n, "qname", None) and n.qname[1] == "list":
                d += 1
            n = getattr(n, "parentNode", None)
        return d

    def heading_level(el):
        lvl = el.attributes.get((TEXT, "outline-level"))
        if lvl:
            return str(lvl)
        style = el.attributes.get((TEXT, "style-name")) or ""
        if "Heading_20_" in style:
            return style.split("Heading_20_")[-1] or "1"
        return "1"

    def walk(el, seq):
        for ch in el.childNodes:
            q = getattr(ch, "qname", None)
            if q is None:
                continue
            ln = q[1]
            if ln in dropped:
                continue
            if ln == "h":
                seq.append("h" + heading_level(ch))
            elif ln == "p":
                depth = depth_of(ch)
                if depth:
                    seq.append("li%d" % (depth - 1))
                elif text_of(ch).strip():
                    seq.append("p")
            elif ln == "table":
                seq.append("TBL")
            elif ln in CONTAINERS:
                walk(ch, seq)

    def from_xml():
        # Some packages load through odfpy without a body element — the two
        # templates in the corpus are among them — so the same block model is
        # applied to content.xml with the standard library's parser instead.
        with zipfile.ZipFile(odt) as z:
            root = ET.fromstring(z.read("content.xml"))
        text = root.find("{%s}body/{%s}text" % (OFFICE, OFFICE))
        if text is None:
            raise ValueError("no office:text body")
        seq = []

        def wt(el, depth=0):
            for ch in el:
                tag = ch.tag.split("}")[-1]
                if tag in dropped:
                    continue
                if tag == "h":
                    seq.append("h" + (ch.get("{%s}outline-level" % TEXT) or "1"))
                elif tag == "p":
                    if depth:
                        seq.append("li%d" % (depth - 1))
                    elif "".join(ch.itertext()).strip():
                        seq.append("p")
                elif tag == "table":
                    seq.append("TBL")
                elif tag in CONTAINERS:
                    deeper = 1 if tag == "list" and "list" not in flat else 0
                    wt(ch, depth + deeper)

        wt(text)
        return seq

    try:
        doc = load(odt)
        txt = doc.text
        if not hasattr(txt, "childNodes"):
            raise ValueError("no body element through odfpy")
        seq = []
        walk(txt, seq)
        return {"seq": seq, "pics": len(txt.getElementsByType(DR.Image)),
                "reader": "odfpy"}
    except Exception:
        return {"seq": from_xml(), "pics": 0, "reader": "elementtree"}


# ---------------------------------------------------------------- output side

def docx_blocks(docx):
    need("docx", "python-docx")
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    d = Document(docx)
    seq = []

    def kind(p, el):
        style = (p.style.name or "") if p.style is not None else ""
        low = style.lower().replace(" ", "")
        if low.startswith("heading") and low[7:].isdigit():
            return "h" + low[7:]
        if el.findall("%spPr/%snumPr" % (W, W)):
            lvl = el.findall("%spPr/%snumPr/%silvl" % (W, W, W))
            v = lvl[0].get(W + "val") if lvl else None
            return "li%s" % (v if v is not None else "0")
        return "p"

    for child in d.element.body.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, d)
            k = kind(p, child)
            # A block is a text-bearing paragraph. The source side drops a
            # paragraph carrying only a frame for the same reason — what it
            # carries is a picture or an object, not text — and keeping it on
            # one side only would offset every later block.
            if k == "p" and not (p.text or "").strip():
                continue
            seq.append(k)
        elif child.tag == qn("w:tbl"):
            seq.append("TBL")
    return {"seq": seq, "pics": len(d.inline_shapes)}


def apple_says(docx):
    """A second implementation's verdict, where the host ships one."""
    if not shutil.which("mdimport"):
        return {"checked": False, "imported": None}
    r = subprocess.run(["mdimport", "-t", "-d2", docx], capture_output=True, text=True)
    return {"checked": True, "imported": "Imported '" in (r.stdout + r.stderr)}


# ---------------------------------------------------------------- the measure

def measure(odt):
    name = os.path.basename(odt)
    work = tempfile.mkdtemp(prefix="foreign-")
    try:
        src = os.path.join(work, name)
        shutil.copy2(odt, src)
        docx = os.path.splitext(src)[0] + ".docx"
        r = subprocess.run([binary(), "get", src, "/body"], capture_output=True,
                           text=True, env=ENV, cwd=work)
        rec = {"name": name, "produced": os.path.exists(docx)}
        if not rec["produced"]:
            text = r.stdout + r.stderr
            rec["verdict"] = ("REFUSED" if any(k in text for k in REFUSAL)
                              else "NOT-PRODUCED")
            rec["why"] = [l for l in (r.stderr or r.stdout).splitlines() if l.strip()][:2]
            return rec
        dropped, flat = declarations(src)
        want = source_blocks(src, dropped, flat)
        got = docx_blocks(docx)
        rec.update(want=want, got=got, apple=apple_says(docx),
                   declared=sorted(dropped | flat))
        first = None
        for i in range(max(len(want["seq"]), len(got["seq"]))):
            a = want["seq"][i] if i < len(want["seq"]) else None
            b = got["seq"][i] if i < len(got["seq"]) else None
            if a != b:
                first = (i, a, b)
                break
        rec["diff"] = first
        rec["verdict"] = "SAME" if first is None else "DIFF"
        return rec
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main(argv):
    files = sorted(glob.glob(os.path.join(FILES, "*.odt")))
    if argv:
        files = [f for f in files if os.path.basename(f) in argv]
    if not files:
        sys.exit("no .odt documents in %s — run tests/corpus/fetch.py" % FILES)
    rows = [measure(f) for f in files]
    counts, blocks = {}, 0
    for r in rows:
        counts[r["verdict"]] = counts.get(r["verdict"], 0) + 1
        detail = ""
        if r["verdict"] == "DIFF":
            i, a, b = r["diff"]
            detail = "  block %d: source %r, output %r  (%d/%d blocks)" % (
                i, a, b, len(r["want"]["seq"]), len(r["got"]["seq"]))
        elif r["verdict"] in ("REFUSED", "NOT-PRODUCED"):
            detail = "  " + str(r.get("why"))
        else:
            blocks += len(r["want"]["seq"])
        apple = r.get("apple") or {}
        print("%-32s %-13s pics %s/%s apple=%-5s%s"
              % (r["name"], r["verdict"], r["got"]["pics"] if "got" in r else None,
                 r["want"]["pics"] if "want" in r else None,
                 apple.get("imported"), detail))
    print("\n%d document(s): %d agree over %d text blocks, %s"
          % (len(rows), counts.get("SAME", 0), blocks,
             ", ".join("%d %s" % (v, k) for k, v in sorted(counts.items())
                       if k != "SAME") or "0 disagree"))
    unexpected = [r for r in rows if r["verdict"] not in ("SAME", "REFUSED")]
    return 1 if unexpected else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
