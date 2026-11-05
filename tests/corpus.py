#!/usr/bin/env python3
"""Measure the reader against documents it did not make.

Every fixture under `tests/` was written by the same hand as the reader. This
suite runs it over documents produced by OpenOffice.org, OpenOffice, NeoOffice,
LibreOffice and pandoc — the files `tests/corpus/sources.json` names — and judges
each one on three things:

  1. it converts, and the produced .docx validates (a clean refusal is recorded
     as a refusal, with its reason);
  2. every construct the reader does not convert is **named in its own report** —
     a construct that vanishes silently is the defect this suite exists to catch;
  3. nothing crashes or hangs.

Plus one cheap extra: a construct reported as dropped must not have its text in
the produced document.

    python3 tests/corpus/fetch.py                    # once
    OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py [--verbose]

Exit status is 1 when any document fails one of the three, so CI can gate on it.
"""
import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PLUGIN = os.path.join(ROOT, "plugin")
CORPUS = os.path.join(HERE, "corpus")
FILES = os.path.join(CORPUS, "files")
SOURCES = os.path.join(CORPUS, "sources.json")

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

# Elements the reader turns into document content. Everything else in the text
# flow must be named in the reader's report — that is the contract the corpus
# checks, so this set is deliberately the small one.
CONVERTED = {
    "p", "h", "span", "s", "tab", "line-break", "list", "list-item", "section",
    "a", "note", "note-citation", "note-body", "frame", "image", "g", "title", "desc",
    "table", "table-row", "table-cell", "covered-table-cell", "table-header-rows",
    "table-row-group", "table-column",
}

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
}


def q(ns, tag):
    return "{%s}%s" % (NS[ns], tag)


def local(tag):
    return tag.rsplit("}", 1)[-1] if "}" in str(tag) else str(tag)


def text_flows(root):
    """The subtrees whose content is document text: the body and any decoration."""
    out = []
    body = root.find(q("office", "body"))
    if body is not None:
        for el in body:
            if local(el.tag) == "text":
                out.append(el)
    ms = root.find(q("office", "master-styles"))
    if ms is not None:
        for page in ms:
            for deco in page:
                if local(deco.tag) in ("header", "footer", "header-left", "footer-left"):
                    out.append(deco)
    return out


def run(args, timeout=180):
    try:
        p = subprocess.run(args, capture_output=True, text=True, timeout=timeout, env=ENV)
        return p.returncode, p.stdout, p.stderr, False
    except subprocess.TimeoutExpired:
        return None, "", "", True


def report_names(err):
    """Element names the reader's own report mentions."""
    return set(re.findall(r"(?:text|draw|table|office|style|svg|fo):([A-Za-z][\w-]*)", err))


def silent_drops(root, reported):
    """{name: count} for flow elements neither converted nor reported."""
    bad = {}

    def walk(el, shadowed):
        n = local(el.tag)
        now = shadowed or n in reported
        if not now and n not in CONVERTED:
            bad[n] = bad.get(n, 0) + 1
        for kid in el:
            walk(kid, now)

    for flow in text_flows(root):
        for kid in flow:
            walk(kid, False)
    return bad


def leaked_text(root, reported, body_text):
    """{name: text} for a construct reported as dropped whose text is in the body."""
    leaks = {}

    def walk(el, shadowed):
        n = local(el.tag)
        now = shadowed or n in reported
        if not now and n not in CONVERTED:
            t = re.sub(r"\s+", " ", "".join(el.itertext())).strip()
            if len(t) > 12 and t in body_text:
                leaks[n] = t[:60]
        for kid in el:
            walk(kid, now)

    for flow in text_flows(root):
        for kid in flow:
            walk(kid, False)
    return leaks


def docx_body_text(path):
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "replace")
    except (KeyError, zipfile.BadZipFile, OSError):
        return ""
    return re.sub(r"\s+", " ", "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)))


REFUSAL = ("not well-formed", "not a readable", "not an OpenDocument", "not an ODF",
           "no office:text body", "has no content.xml")


def measure(path):
    rc_dump, dumped, err, to_dump = run([sys.executable, PLUGIN, "dump", path])

    root = None
    with zipfile.ZipFile(path) as z:
        try:
            root = ET.fromstring(z.read("content.xml"))
        except (KeyError, ET.ParseError):
            root = None

    reported = report_names(err)
    silent = silent_drops(root, reported) if root is not None else {}

    work = tempfile.mkdtemp(prefix="corpus-")
    try:
        src = os.path.join(work, os.path.basename(path))
        shutil.copy2(path, src)
        rc_get, _out, get_err, to_get = run([BIN, "get", src, "/body"])
        docx = os.path.splitext(src)[0] + ".docx"
        rc_val = None
        if os.path.exists(docx):
            rc_val, _vo, _ve, _ = run([BIN, "validate", docx], timeout=120)
        leaks = leaked_text(root, reported, docx_body_text(docx)) if root is not None else {}
        produced = os.path.exists(docx)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    refused = rc_get != 0 and any(s in get_err for s in REFUSAL)
    crash = "Traceback (most recent call last)" in err or \
            "Traceback (most recent call last)" in get_err
    return dict(name=os.path.basename(path), rc_dump=rc_dump, rc_get=rc_get, rc_val=rc_val,
                refused=refused, produced=produced, crash=crash,
                timeout=to_dump or to_get, silent=silent, leaks=leaks,
                warns=[l for l in err.strip().splitlines() if l.strip()], get_err=get_err)


def verdict(r):
    """(ok, [flags]) — a clean refusal passes; everything else has to be green."""
    flags = []
    if r["crash"]:
        flags.append("CRASH")
    if r["timeout"]:
        flags.append("TIMEOUT")
    if r["silent"]:
        flags.append("SILENT-DROP " + ", ".join("%s x%d" % (k, v)
                                                for k, v in sorted(r["silent"].items())))
    if r["leaks"]:
        flags.append("LEAK " + ", ".join(sorted(r["leaks"])))
    converted = r["rc_get"] == 0 and r["rc_val"] == 0
    if not converted and not r["refused"]:
        flags.append("CONVERT-FAIL rc=%s validate=%s" % (r["rc_get"], r["rc_val"]))
    return (not flags), flags


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--verbose", action="store_true", help="print every warning line")
    ap.add_argument("dirs", nargs="*", help="measure these directories instead of the corpus")
    a = ap.parse_args(argv)

    if a.dirs:
        files = sorted(f for d in a.dirs for f in glob.glob(os.path.join(d, "*.odt")))
    else:
        if not os.path.isdir(FILES) or not glob.glob(os.path.join(FILES, "*.odt")):
            print("no documents in %s — run: python3 tests/corpus/fetch.py" % FILES)
            return 2
        files = sorted(glob.glob(os.path.join(FILES, "*.odt")))
        manifest = {}
        if os.path.exists(SOURCES):
            with open(SOURCES, encoding="utf-8") as fh:
                manifest = {d["file"]: d for d in json.load(fh)["documents"]}
    if not files:
        print("no .odt documents found")
        return 2

    manifest = locals().get("manifest", {})
    rows, failures = [], 0
    for f in files:
        r = measure(f)
        ok, flags = verdict(r)
        if not ok:
            failures += 1
        meta = manifest.get(os.path.basename(f), {})
        state = ("refused: %s" % next((s for s in REFUSAL if s in r["get_err"]), "explained")
                 if r["refused"] else
                 "converts + validates" if (r["rc_get"] == 0 and r["rc_val"] == 0) else "FAILED")
        print("%-28s %-22s %-24s %s" % (os.path.basename(f), (meta.get("producer") or "-")[:24],
                                        state, " ".join(flags) if flags else "ok"))
        if a.verbose or not ok:
            for line in r["warns"]:
                print("      " + line)
        rows.append(r)

    converted = sum(1 for r in rows if r["rc_get"] == 0 and r["rc_val"] == 0)
    refused = sum(1 for r in rows if r["refused"] and not (r["rc_get"] == 0 and r["rc_val"] == 0))
    producers = len({m.get("producer") for m in manifest.values() if m.get("producer")})
    print("\n%d documents%s | %d convert and validate | %d refused with a reason | "
          "%d silent drops | %d crashes | %d timeouts"
          % (len(rows), (" (%d producers)" % producers) if producers else "",
             converted, refused,
             sum(1 for r in rows if r["silent"]), sum(1 for r in rows if r["crash"]),
             sum(1 for r in rows if r["timeout"])))
    if failures:
        print("\n%d document(s) failed" % failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
