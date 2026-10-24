#!/usr/bin/env python3
"""End-to-end verification of the .odt dump-reader plugin against a real
officecli binary. Every assertion prints PASS/FAIL; the process exits 1 if any
failed. Run as:

    OFFICECLI_BIN=/path/to/officecli python3 tests/verify.py

The binary under test must be a build of the tree the plugin targets — the
README records which one this was measured against.
"""

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
})

results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print(("PASS  " if ok else "FAIL  ") + name + (("  — " + detail) if detail else ""))


def run(args, env=None):
    p = subprocess.run([BIN] + args, capture_output=True, text=True, env=env or ENV)
    return p.returncode, p.stdout, p.stderr


def body_lines(docx):
    rc, out, err = run(["get", docx, "/body", "--depth", "1"])
    return rc, [l for l in out.splitlines() if l.strip()], err


def main():
    work = tempfile.mkdtemp(prefix="odt-verify-")
    src = os.path.join(work, "report.odt")
    subprocess.run([sys.executable, os.path.join(HERE, "make_fixture.py"), src], check=True)

    env = dict(ENV)
    env["OFFICECLI_PLUGIN_DUMP_READER_ODT"] = PLUGIN

    # -- arm 1: conversion through discovery path 1 (env var) -----------------
    rc, out, err = run(["get", src, "/body", "--depth", "1"], env=env)
    sibling = os.path.join(work, "report.docx")
    check("conversion exits 0", rc == 0, "rc=%d stderr=%s" % (rc, err.strip()[:200]))
    check("sibling .docx created", os.path.isfile(sibling),
          "expected %s" % sibling)

    if not os.path.isfile(sibling):
        print("\nverification stopped: no output file to inspect")
        return 1

    # -- arm 2: the produced docx validates ----------------------------------
    rc, out, err = run(["validate", sibling])
    check("generated .docx validates", rc == 0, (out + err).strip()[:200])

    # -- arm 3: content read back --------------------------------------------
    rc, lines, err = body_lines(sibling)
    joined = "\n".join(lines)

    def has(pattern):
        return re.search(pattern, joined) is not None

    check("heading 1 present", has(r'\(paragraph\) "Quarterly Report" .*style=heading 1'))
    check("heading 2 present", has(r'\(paragraph\) "Summary" .*style=heading 2'))
    check("no duplicated paragraph text",
          not has(r'"Quarterly ReportQuarterly Report"'),
          "each paragraph must appear once")
    check("inline paragraph text intact",
          has(r'\(paragraph\) "Plain paragraph with bold, italic and underline runs\."'))
    check("centered paragraph", has(r'"Centered line" .*align=center'))
    check("bullet list items", len(re.findall(r'listStyle=bullet', joined)) == 2,
          "found %d" % len(re.findall(r'listStyle=bullet', joined)))
    check("ordered list items", len(re.findall(r'listStyle=ordered', joined)) == 2,
          "found %d" % len(re.findall(r'listStyle=ordered', joined)))
    check("table present as 2x2", has(r'\(table\).*cols=2 rows=2'))
    check("trailing paragraph after table", has(r'"After the table"'))

    # run-level formatting (deeper read of the formatted paragraph)
    rc, out, err = run(["get", sibling, "/body/p[3]", "--depth", "1"])
    check("bold run", "bold=true" in out, out.splitlines()[0][:120] if out else "")
    check("italic run", "italic=true" in out)
    check("underline run", "underline=single" in out)

    # table cell content
    rc, out, err = run(["get", sibling, "/body/tbl[1]/tr[2]/tc[1]"])
    check("table cell text", '"EMEA"' in out, out.strip()[:120])

    # -- arm 4: images -------------------------------------------------------
    # The source declares five translatable images (three PNGs, one SVG, and a
    # second reference to one of the PNGs) plus two the reader must refuse: an
    # EMF it cannot sniff, and one inside a table cell.
    with zipfile.ZipFile(src) as z:
        source_pics = {n: z.read(n) for n in z.namelist() if n.startswith("Pictures/")}
    with zipfile.ZipFile(sibling) as z:
        media = {n: z.read(n) for n in z.namelist()
                 if re.match(r"^(word/)?media/", n)}

    # 5 pictures, embedded as 6 parts: the SVG also carries the 1x1 PNG
    # fallback officecli generates when `fallback` is not given.
    check("image parts landed in the .docx", len(media) == 6,
          "%d part(s): %s" % (len(media), ", ".join(sorted(media))))
    check("the svg is embedded as svg",
          any(n.endswith(".svg") for n in media), ", ".join(sorted(media)))
    lost = [n for n, blob in source_pics.items()
            if n.endswith(".png") and blob not in media.values()]
    check("every source PNG appears byte-for-byte in the .docx", not lost,
          "missing: %s" % ", ".join(lost))

    rc, out, err = run(["get", sibling, "/body", "--depth", "1"])
    n_paras = len([l for l in out.splitlines() if re.match(r"^/body/p\[", l)])
    pics = []
    for i in range(1, n_paras + 1):
        rc, pout, perr = run(["get", sibling, "/body/p[%d]" % i, "--depth", "1"])
        for line in pout.splitlines():
            if "(picture)" in line:
                pics.append((i, line))
    check("five pictures read back", len(pics) == 5,
          "%d found in paragraph(s) %s" % (len(pics), [i for i, _ in pics]))

    # the paragraph that mixes text and an image: order must survive
    rc, out, err = run(["get", sibling, "/body/p[5]", "--depth", "1"])
    lines = out.splitlines()
    ordered = (len(lines) >= 4 and "(run)" in lines[1] and "(picture)" in lines[2]
               and "(run)" in lines[3])
    check("text, image, text keep their order", ordered, out.strip()[:200])
    check("surrounding text survives around the image",
          "Text before" in lines[0] and "text after" in lines[0],
          lines[0][:160] if lines else "")
    check("frame geometry is honoured",
          "width=2.0cm" in lines[2] and "height=1.5cm" in lines[2],
          lines[2][:200] if len(lines) > 2 else "")
    check("svg:title becomes the alt text", "alt=Red square" in lines[2],
          lines[2][:200] if len(lines) > 2 else "")

    # no frame geometry at all → pixel dimensions at 96dpi (96x48 → 2.54x1.27cm;
    # officecli prints lengths rounded to 0.1cm)
    blue = [l for _, l in pics if "blue.png" in l]
    check("size falls back to the pixel dimensions when the frame has none",
          any("width=2.5cm" in l and "height=1.3cm" in l for l in blue),
          " | ".join(blue)[:220])
    rels = [re.search(r"relId=(\S+)", l).group(1) for l in blue
            if re.search(r"relId=(\S+)", l)]
    check("two references to one source image get their own part",
          len(rels) == 2 and len(set(rels)) == 2, ", ".join(rels))

    # -- arm 5: the plugin's own output, checked directly --------------------
    # Main collects a dump-reader's stderr but only surfaces it when the plugin
    # *fails* (DumpReaderInvoker passes no OnStderrLine), so the skip report is
    # only observable by running the plugin itself. Same for the JSONL contract.
    p = subprocess.run([sys.executable, PLUGIN, "dump", src],
                       capture_output=True, text=True, env=ENV)
    check("direct dump exits 0", p.returncode == 0, p.stderr.strip()[:200])
    bad = []
    for n, line in enumerate(p.stdout.splitlines(), 1):
        if not line.strip():
            continue
        try:
            obj = json.loads(line)
        except ValueError:
            bad.append(n)
            continue
        if "command" not in obj:
            bad.append(n)
    check("every emitted line is a JSON command object", not bad,
          "line(s): %s" % bad[:5])
    check("a refused image format is reported, not dropped in silence",
          "vector.emf" in p.stderr and "not png/jpeg/gif/bmp/tiff/svg" in p.stderr,
          (p.stderr.strip().splitlines() or [""])[0][:200])
    check("an image inside a table cell is reported",
          "inside a table cell" in p.stderr, (p.stderr.strip().splitlines() or [""])[0][:200])
    check("the end-of-run summary counts images",
          "and 5 image(s)" in p.stderr, (p.stderr.strip().splitlines() or [""])[-1][:200])
    rc, out, err = run(["get", sibling, "/body/tbl[1]/tr[2]/tc[2]"])
    check("the cell keeps its text next to a dropped image", '"1200"' in out,
          out.strip()[:120])

    # -- arm 6: sibling reuse (no reconversion) ------------------------------
    rc, out, err = run(["get", src, "/body", "--depth", "1"], env=env)
    check("sibling reused on second open", "generated" not in err,
          "stderr: " + err.strip()[:160])

    # -- precondition for the two "no plugin" arms ---------------------------
    # These arms can only be faked if nothing else resolves a .odt dump-reader.
    # A real user may legitimately have the plugin installed at discovery path 2
    # or 3, in which case `main` falls through to it and the arms would report a
    # false failure. Detect that and say so instead of inventing a defect.
    stray = []
    rc, out, err = run(["plugins", "list", "--json"])
    try:
        for item in json.loads(out).get("data", []):
            if "dump-reader" in item.get("kinds", []) and ".odt" in item.get("extensions", []):
                stray.append(item.get("path", "?"))
    except (ValueError, AttributeError):
        pass

    lonely = os.path.join(work, "lonely")
    os.makedirs(lonely, exist_ok=True)
    shutil.copy(src, os.path.join(lonely, "report.odt"))

    if stray:
        print("SKIP  negative arms: a .odt dump-reader is already installed at %s\n"
              "      remove it (e.g. <officecli-dir>/plugins/dump-reader/odt or "
              "~/.officecli/plugins/dump-reader/odt) and re-run to exercise them"
              % ", ".join(stray))
    else:
        # -- arm 7 (negative): no plugin installed ---------------------------
        bare = dict(ENV)
        bare.pop("OFFICECLI_PLUGIN_DUMP_READER_ODT", None)
        rc, out, err = run(["get", os.path.join(lonely, "report.odt"), "/body"], env=bare)
        msg = (out + err)
        check("missing plugin: non-zero exit", rc != 0, "rc=%d" % rc)
        check("missing plugin: clear error, not a crash",
              ("plugin" in msg.lower()) and ("Unhandled exception" not in msg),
              " | ".join(msg.strip().splitlines()[:2]))

        # -- arm 8 (negative): manifest protocol != 1 ------------------------
        bad = os.path.join(work, "badproto")
        os.makedirs(bad, exist_ok=True)
        bad_plugin = os.path.join(bad, "plugin")
        with open(bad_plugin, "w") as fh:
            fh.write('#!/usr/bin/env python3\nimport json,sys\n'
                     'print(json.dumps({"name":"officecli-odt-badproto","version":"0.0.1",'
                     '"protocol":2,"kinds":["dump-reader"],"extensions":[".odt"],'
                     '"target":"docx","runtime":"python",'
                     '"idle_timeout_seconds":{"default":30}}))\n')
        os.chmod(bad_plugin, 0o755)
        badenv = dict(ENV)
        badenv["OFFICECLI_PLUGIN_DUMP_READER_ODT"] = bad_plugin
        rc, out, err = run(["get", os.path.join(lonely, "report.odt"), "/body"], env=badenv)
        msg = out + err
        check("protocol mismatch: refused (non-zero exit)", rc != 0, "rc=%d" % rc)
        check("protocol mismatch: rejection is explained",
              "protocol" in msg.lower(), " | ".join(msg.strip().splitlines()[:2]))
        # plugin-protocol.md §4.1 specifies exit code 5 for this case; record
        # what the binary actually returns so the README states the measured
        # value rather than the documented one.
        print("NOTE  protocol-mismatch exit code measured: %d (protocol §4.1 says 5)" % rc)

    # -- arm 9 (negative): corrupt input -------------------------------------
    junk = os.path.join(work, "junk.odt")
    with open(junk, "w") as fh:
        fh.write("this is not a zip file")
    jen = dict(ENV)
    jen["OFFICECLI_PLUGIN_DUMP_READER_ODT"] = PLUGIN
    rc, out, err = run(["get", junk, "/body"], env=jen)
    msg = out + err
    check("corrupt input: reported, not crashed",
          rc != 0 and "Unhandled exception" not in msg, "rc=%d" % rc)

    failed = [r for r in results if not r[1]]
    print("\n%d checks, %d failed" % (len(results), len(failed)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
