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


CELLS = (cell_hyperlink, cell_mailto, cell_merged_cells, cell_vertical_text)


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
