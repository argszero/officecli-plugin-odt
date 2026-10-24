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

    # -- arm 4: sibling reuse (no reconversion) ------------------------------
    rc, out, err = run(["get", src, "/body", "--depth", "1"], env=env)
    check("sibling reused on second open", "generated" not in err,
          "stderr: " + err.strip()[:160])

    # -- arm 5 (negative): no plugin installed -------------------------------
    bare = dict(ENV)
    bare.pop("OFFICECLI_PLUGIN_DUMP_READER_ODT", None)
    lonely = os.path.join(work, "lonely")
    os.makedirs(lonely, exist_ok=True)
    shutil.copy(src, os.path.join(lonely, "report.odt"))
    os.remove(os.path.join(lonely, "report.docx")) if os.path.exists(
        os.path.join(lonely, "report.docx")) else None
    rc, out, err = run(["get", os.path.join(lonely, "report.odt"), "/body"], env=bare)
    msg = (out + err)
    check("missing plugin: non-zero exit", rc != 0, "rc=%d" % rc)
    check("missing plugin: clear error, not a crash",
          ("plugin" in msg.lower()) and ("Unhandled exception" not in msg),
          " | ".join(msg.strip().splitlines()[:2]))

    # -- arm 6 (negative): manifest protocol != 1 ----------------------------
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
    # plugin-protocol.md §4.1 specifies exit code 5 for this case; record what
    # the binary actually returns so the README can state the measured value.
    print("NOTE  protocol-mismatch exit code measured: %d (protocol §4.1 says 5)" % rc)

    # -- arm 7 (negative): corrupt input -------------------------------------
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
