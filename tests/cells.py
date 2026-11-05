#!/usr/bin/env python3
"""Two-sided cells for the .odt dump-reader.

Each cell is one feature measured twice: the feature is present in one fixture
and absent from the other, and the assertions run against both. A cell only
counts as passed when **all three** hold:

  1. the feature appears in the converted .docx (the positive half),
  2. `officecli validate` on that .docx exits 0 (the package is sound),
  3. the feature-absent control does *not* gain the feature and still validates.

Without 3 a cell would pass on a reader that invented the feature everywhere;
without 2 a cell would pass on a reader that produced a document Word refuses
to open. Both halves convert through the same discovery path, so the only
difference between them is the construct under test.

    OFFICECLI_BIN=/path/to/officecli python3 tests/cells.py [-v]
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PLUGIN = os.path.join(ROOT, "plugin")
BIN = os.environ.get("OFFICECLI_BIN")
if not BIN:
    sys.exit("set OFFICECLI_BIN to the officecli binary to test against")

ENV = dict(os.environ)
ENV.update({
    "OFFICECLI_NO_AUTO_INSTALL": "1",
    "OFFICECLI_NO_AUTO_RESIDENT": "1",
    "OFFICECLI_SKIP_UPDATE": "1",
    "OFFICECLI_PLUGIN_DUMP_READER_ODT": PLUGIN,
})

VERBOSE = False
results = []          # (cell, half, name, ok, detail)


def check(cell, half, name, ok, detail=""):
    results.append((cell, half, name, bool(ok), detail))
    if VERBOSE or not ok:
        print("%s  [%s/%s] %s%s" % ("PASS" if ok else "FAIL", cell, half, name,
                                    ("  — " + detail) if detail else ""))


def run(args):
    p = subprocess.run([BIN] + args, capture_output=True, text=True, env=ENV)
    return p.returncode, p.stdout, p.stderr


def convert(feature, on, work):
    """Build a fixture and convert it; returns (docx_path, rc, stderr)."""
    stem = "%s-%s" % (feature, "on" if on else "off")
    src = os.path.join(work, stem + ".odt")
    args = [sys.executable, os.path.join(HERE, "make_cell_fixture.py"),
            "--feature", feature, "--out", src]
    if not on:
        args.append("--without")
    subprocess.run(args, check=True)
    rc, out, err = run(["get", src, "/body"])
    return os.path.join(work, stem + ".docx"), rc, err


def xml(docx, part):
    with zipfile.ZipFile(docx) as z:
        try:
            return z.read(part).decode("utf-8")
        except KeyError:
            return ""


def rels(docx):
    return xml(docx, "word/_rels/document.xml.rels")


def read_body(docx):
    rc, out, err = run(["get", docx, "/body", "--depth", "3"])
    return rc, out


def read_root(docx):
    """A header/footer hangs off the document root, not off /body."""
    rc, out, err = run(["get", docx, "/", "--depth", "1"])
    return rc, out


def parts(docx):
    with zipfile.ZipFile(docx) as z:
        return set(z.namelist())


def num_formats(numbering):
    """ilvl → numFmt, read out of the numbering part's own level definitions."""
    out = {}
    for m in re.finditer(r'<w:lvl w:ilvl="(\d+)"[^>]*>(.*?)</w:lvl>', numbering, re.S):
        fmt = re.search(r'<w:numFmt w:val="([^"]+)"', m.group(2))
        out[m.group(1)] = fmt.group(1) if fmt else None
    return out


# ---------------------------------------------------------------------------
# cells
# ---------------------------------------------------------------------------

def cell_hyperlink(work):
    """text:a with an absolute href → w:hyperlink + an external relationship."""
    cell = "hyperlink"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    body = xml(on, "word/document.xml")
    check(cell, "on", "a w:hyperlink element is present", "<w:hyperlink" in body)
    r = rels(on)
    check(cell, "on", "the relationship targets the source URL",
          'Target="https://example.com/odt"' in r, r[:200])
    check(cell, "on", "the relationship is external",
          re.search(r'Target="https://example\.com/odt"[^>]*TargetMode="External"|'
                    r'TargetMode="External"[^>]*Target="https://example\.com/odt"', r)
          is not None)
    _, out = read_body(on)
    check(cell, "on", "the link text survives", "the site" in out)
    check(cell, "on", "read-back reports it as a hyperlink",
          "isHyperlink=true" in out and "url=https://example.com/odt" in out, out[:200])

    # control: same sentence, no text:a
    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    check(cell, "off", "control gains NO hyperlink element",
          "<w:hyperlink" not in xml(off, "word/document.xml"))
    check(cell, "off", "control gains NO hyperlink relationship",
          "relationships/hyperlink" not in rels(off))
    _, out = read_body(off)
    check(cell, "off", "control text is intact", "See the site now." in out, out[:160])
    return cell


def cell_mailto(work):
    """A mailto: href is an external target too, and must survive as one."""
    cell = "mailto"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    body = xml(on, "word/document.xml")
    check(cell, "on", "a w:hyperlink element is present",
          re.search(r"<w:hyperlink\b", body) is not None)
    r = rels(on)
    check(cell, "on", "the relationship keeps the mailto target",
          'Target="mailto:team@example.com"' in r, r[:200])
    check(cell, "on", "the relationship is external",
          re.search(r'Target="mailto:team@example.com"[^>]*TargetMode="External"|'
                    r'TargetMode="External"[^>]*Target="mailto:team@example.com"', r)
          is not None)
    _, out = read_body(on)
    check(cell, "on", "the link text survives", "the team" in out)
    check(cell, "on", "read-back reports it as a hyperlink",
          "isHyperlink=true" in out and "mailto:team@example.com" in out, out[:200])

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    check(cell, "off", "control gains NO hyperlink element",
          re.search(r"<w:hyperlink\b", xml(off, "word/document.xml")) is None)
    check(cell, "off", "control gains NO hyperlink relationship",
          "relationships/hyperlink" not in rels(off))
    _, out = read_body(off)
    check(cell, "off", "control text is intact", "Write to the team." in out, out[:160])
    return cell


def cell_merged_cells(work):
    """number-columns-spanned / number-rows-spanned → w:gridSpan / w:vMerge."""
    cell = "merged-cells"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    body = xml(on, "word/document.xml")
    check(cell, "on", "w:gridSpan is written",
          re.search(r'<w:gridSpan w:val="2"\s*/>', body) is not None, body[:200])
    check(cell, "on", "w:vMerge restart is written",
          re.search(r'<w:vMerge w:val="restart"', body) is not None)
    check(cell, "on", "w:vMerge continue is written",
          re.search(r"<w:vMerge\s*/>", body) is not None)
    check(cell, "on", "the merged cell keeps its text", "Corner" in body)
    check(cell, "on", "the neighbour cell keeps its text", "Head" in body)
    check(cell, "on", "the second row keeps its text", "Body" in body)

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    off_body = xml(off, "word/document.xml")
    check(cell, "off", "control gains NO gridSpan", "<w:gridSpan" not in off_body)
    check(cell, "off", "control gains NO vMerge",
          re.search(r"<w:vMerge\b", off_body) is None)
    check(cell, "off", "control grid is 2 rows x 3 cells",
          len(re.findall(r"<w:tr[ >]", off_body)) == 2
          and len(re.findall(r"<w:tc>", off_body)) == 6,
          "rows=%d tc=%d" % (len(re.findall(r"<w:tr[ >]", off_body)),
                             len(re.findall(r"<w:tc>", off_body))))
    return cell


def cell_vertical_text(work):
    """style:text-position / text-line-through → w:vertAlign / w:strike."""
    cell = "vertical-text"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    body = xml(on, "word/document.xml")
    check(cell, "on", "w:strike is written",
          re.search(r"<w:strike\s*/>", body) is not None)
    check(cell, "on", "superscript is written",
          re.search(r'<w:vertAlign w:val="superscript"\s*/>', body) is not None, body[:200])
    check(cell, "on", "subscript is written",
          re.search(r'<w:vertAlign w:val="subscript"\s*/>', body) is not None, body[:200])
    _, out = read_body(on)
    check(cell, "on", "the characters survive", "Water is H" in out and "O and E=mc" in out,
          out[:200])
    check(cell, "on", "read-back reports superscript", "superscript=true" in out, out[:220])
    check(cell, "on", "read-back reports subscript", "subscript=true" in out, out[:220])
    check(cell, "on", "read-back reports strike", "strike=true" in out, out[:220])

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    off_body = xml(off, "word/document.xml")
    check(cell, "off", "control gains NO strike",
          re.search(r"<w:strike\b", off_body) is None)
    check(cell, "off", "control gains NO vertAlign", "<w:vertAlign" not in off_body)
    _, out = read_body(off)
    check(cell, "off", "control text is intact", "but this is wrong." in out, out[:160])
    return cell


def note_cell(work, cell, kind):
    """text:note → w:<kind>Reference at the anchor + word/<kind>s.xml holding the body."""
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)
    part = "word/%ss.xml" % kind                      # footnotes.xml / endnotes.xml
    other = "word/%ss.xml" % ("endnote" if kind == "footnote" else "footnote")
    ref = "<w:%sReference" % kind

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    body = xml(on, "word/document.xml")
    check(cell, "on", "a %s element is written" % ref, ref in body)
    check(cell, "on", "the reference carries an id",
          re.search(r'<w:%sReference[^>]*w:id="\d+"' % kind, body) is not None)
    notes = xml(on, part)
    check(cell, "on", "%s is written" % part, bool(notes))
    check(cell, "on", "the note body lands in it", "The note body." in notes)
    found = len(re.findall(r"<w:%s\b" % kind, notes))
    check(cell, "on", "one real note beside the two separators", found == 3,
          "%d <w:%s> in %s" % (found, kind, part))
    # The marker's position is the whole point of an inline note: it has to sit
    # between the text on either side of it, not at the end of the sentence.
    ib, im, ia = (body.find("Body text"), body.find("%sReference" % kind),
                  body.find("more text."))
    check(cell, "on", "the marker sits between the two text runs",
          -1 < ib < im < ia, "head=%d marker=%d tail=%d" % (ib, im, ia))
    # A note body is not paragraph text: a reader that flattened it would read
    # back as if the note had been typed into the sentence.
    check(cell, "on", "the note body is NOT flattened into the paragraph",
          "The note body." not in body)
    check(cell, "on", "the other note part stays empty", xml(on, other) == "")
    _, out = read_body(on)
    check(cell, "on", "the paragraph text is unchanged", "Body text more text." in out,
          out[:200])

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    off_body = xml(off, "word/document.xml")
    check(cell, "off", "control gains NO %s" % ref,
          "<w:%sReference" % kind not in off_body)
    check(cell, "off", "control gains NO %s" % part, xml(off, part) == "")
    _, out = read_body(off)
    check(cell, "off", "control text is intact", "Body text more text." in out, out[:200])
    return cell


def cell_footnote(work):
    """text:note with note-class=footnote → w:footnoteReference + word/footnotes.xml."""
    return note_cell(work, "footnote", "footnote")


def cell_endnote(work):
    """text:note with note-class=endnote → w:endnoteReference + word/endnotes.xml."""
    return note_cell(work, "endnote", "endnote")


def cell_header_footer(work):
    """office:master-styles header/footer → word/header*.xml + footer*.xml + references."""
    cell = "header-footer"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    names = parts(on)
    hdr = sorted(n for n in names if re.match(r"word/header\d+\.xml$", n))
    ftr = sorted(n for n in names if re.match(r"word/footer\d+\.xml$", n))
    check(cell, "on", "a header part is written", len(hdr) == 1, ", ".join(sorted(names)))
    check(cell, "on", "a footer part is written", len(ftr) == 1, ", ".join(sorted(names)))
    hbody = xml(on, hdr[0]) if hdr else ""
    fbody = xml(on, ftr[0]) if ftr else ""
    check(cell, "on", "the header carries its text", "Quarterly report" in hbody, hbody[:160])
    check(cell, "on", "the footer carries its text", "Internal use only" in fbody, fbody[:160])
    check(cell, "on", "the header keeps its alignment",
          re.search(r'<w:jc w:val="right"\s*/>', hbody) is not None, hbody[:200])
    check(cell, "on", "the footer keeps its alignment",
          re.search(r'<w:jc w:val="center"\s*/>', fbody) is not None, fbody[:200])
    doc = xml(on, "word/document.xml")
    check(cell, "on", "the section references the header",
          re.search(r'<w:headerReference w:type="default"', doc) is not None, doc[:200])
    check(cell, "on", "the section references the footer",
          re.search(r'<w:footerReference w:type="default"', doc) is not None)
    # A decoration is its own part and never body text. Without this, a reader
    # that dumped the header into the paragraph would pass on the part existing.
    check(cell, "on", "the header text is NOT in the document body",
          "Quarterly report" not in doc)
    check(cell, "on", "the footer text is NOT in the document body",
          "Internal use only" not in doc)
    check(cell, "on", "the body paragraph survives", "Page body text." in doc)
    _, out = read_root(on)
    check(cell, "on", "read-back reports both references",
          "headerRef.default=/header[1]" in out
          and "footerRef.default=/footer[1]" in out, out[:200])

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    onames = parts(off)
    check(cell, "off", "control gains NO header part",
          not [n for n in onames if re.match(r"word/header\d+\.xml$", n)],
          ", ".join(sorted(onames)))
    check(cell, "off", "control gains NO footer part",
          not [n for n in onames if re.match(r"word/footer\d+\.xml$", n)])
    odoc = xml(off, "word/document.xml")
    check(cell, "off", "control section has no header/footer reference",
          "<w:headerReference" not in odoc and "<w:footerReference" not in odoc)
    check(cell, "off", "control body text is intact", "Page body text." in odoc)
    return cell


def cell_nested_list(work):
    """A list inside a list item → the next numLevel, with its own number format."""
    cell = "nested-list"
    on, rc, err = convert(cell, True, work)
    off, rc2, err2 = convert(cell, False, work)

    check(cell, "on", "conversion exits 0", rc == 0, "rc=%d %s" % (rc, err.strip()[:160]))
    check(cell, "on", "generated .docx validates", run(["validate", on])[0] == 0)
    doc = xml(on, "word/document.xml")
    levels = re.findall(r'<w:ilvl w:val="(\d+)"', doc)
    check(cell, "on", "four list paragraphs", len(levels) == 4, str(levels))
    # Position, not just counts: the nested pair must follow its parent item.
    check(cell, "on", "the nesting sits where the source put it",
          levels == ["0", "0", "1", "1"], str(levels))
    nums = re.findall(r'<w:numId w:val="(\d+)"', doc)
    check(cell, "on", "every item shares one numbering instance",
          len(nums) == 4 and len(set(nums)) == 1, str(nums))
    fmts = num_formats(xml(on, "word/numbering.xml"))
    check(cell, "on", "level 0 numbers in decimal", fmts.get("0") == "decimal", str(fmts))
    check(cell, "on", "level 1 numbers in its own format",
          fmts.get("1") == "lowerLetter", str(fmts))
    _, out = read_body(on)
    check(cell, "on", "read-back reports the nested level",
          "numLevel=1" in out, out[:200])
    check(cell, "on", "read-back gives the nested level a distinct format",
          re.search(r'"Nested one".*numLevel=1 numFmt=lowerLetter', out) is not None,
          out[:300])
    check(cell, "on", "read-back keeps the top level decimal",
          re.search(r'"Top level one".*numLevel=0 numFmt=decimal', out) is not None,
          out[:300])
    check(cell, "on", "all four texts survive",
          all(t in doc for t in ("Top level one", "Top level two",
                                 "Nested one", "Nested two")))

    check(cell, "off", "control conversion exits 0", rc2 == 0, err2.strip()[:160])
    check(cell, "off", "control .docx validates", run(["validate", off])[0] == 0)
    odoc = xml(off, "word/document.xml")
    offlevels = re.findall(r'<w:ilvl w:val="(\d+)"', odoc)
    check(cell, "off", "control keeps every item at level 0",
          offlevels == ["0", "0", "0", "0"], str(offlevels))
    check(cell, "off", "control text is intact",
          all(t in odoc for t in ("Top level one", "Top level two",
                                  "Nested one", "Nested two")))
    return cell


CELLS = (cell_hyperlink, cell_mailto, cell_merged_cells, cell_vertical_text,
         cell_footnote, cell_endnote, cell_header_footer, cell_nested_list)


def main():
    global VERBOSE
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("--keep", action="store_true", help="keep the work directory")
    a = ap.parse_args()
    VERBOSE = a.verbose

    work = tempfile.mkdtemp(prefix="odt-cells-")
    cells = []
    try:
        for fn in CELLS:
            cells.append(fn(work))
    finally:
        if a.keep:
            print("artifacts kept in", work)
        else:
            shutil.rmtree(work, ignore_errors=True)

    print()
    total_ok = 0
    for cell in cells:
        rows = [r for r in results if r[0] == cell]
        on_ok = all(r[3] for r in rows if r[1] == "on")
        off_ok = all(r[3] for r in rows if r[1] == "off")
        passed = len([r for r in rows if r[3]])
        # A cell counts when its positive half and its control both hold.
        good = on_ok and off_ok
        total_ok += 1 if good else 0
        print("%s  cell %-15s %2d/%2d assertions  (positive half %s, control %s)"
              % ("PASS" if good else "FAIL", cell, passed, len(rows),
                 "ok" if on_ok else "FAILED", "ok" if off_ok else "FAILED"))

    failed = [r for r in results if not r[3]]
    print("\n%d cells, %d passed with their feature-absent control; "
          "%d assertions, %d failed"
          % (len(cells), total_ok, len(results), len(failed)))
    return 0 if not failed and total_ok == len(cells) else 1


if __name__ == "__main__":
    sys.exit(main())
