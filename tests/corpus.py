#!/usr/bin/env python3
"""Measure the reader against documents it did not make.

Every fixture under `tests/` was written by the same hand as the reader. This
suite runs it over documents produced by OpenOffice.org, OpenOffice, NeoOffice,
LibreOffice and pandoc — the files `tests/corpus/sources.json` names — and judges
each one on four things:

  1. it converts, and the produced .docx validates (a clean refusal is recorded
     as a refusal, with its reason);
  2. every construct the reader does not convert is **named in its own report** —
     a construct that vanishes silently is one defect this suite exists to catch;
  3. **no word is lost in silence** — the second, and the reason this suite grew
     a fourth axis: a document can convert, validate, report every construct it
     skipped, and still swallow a sentence. Words are compared between an
     independent extraction of the source and the produced document; a missing
     word passes only when the reader named the construct that held it as
     dropped. A word the reader *promised* to keep and then lost fails too.
  4. nothing crashes or hangs.

Plus one cheap extra: a construct reported as dropped must not have its text in
the produced document.

    python3 tests/corpus/fetch.py                    # once
    OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py [--verbose]

Exit status is 1 when any document fails one of the four, so CI can gate on it.
"""
import argparse
import collections
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PLUGIN = os.path.join(ROOT, "plugin")
CORPUS = os.path.join(HERE, "corpus")
FILES = os.path.join(CORPUS, "files")
SOURCES = os.path.join(CORPUS, "sources.json")

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

# Elements the reader turns into document content. Everything else in the text
# flow must be named in the reader's report — that is the contract the corpus
# checks, so this set is deliberately the small one.
CONVERTED = {
    "p", "h", "span", "s", "tab", "line-break", "list", "list-item", "section",
    "a", "note", "note-citation", "note-body", "frame", "image", "g", "title", "desc",
    "table", "table-row", "table-cell", "covered-table-cell", "table-header-rows",
    "table-row-group", "table-column",
}

# A note's body is a different part in the docx vocabulary (word/footnotes.xml),
# so both sides of the word comparison set notes aside: their text is judged by
# verify.py and cells.py, not here.
NOTES_ASIDE = {"note"}

# ODF spells its whitespace as elements, so a walker that ignores them runs words
# together and then blames the reader for the join. Reading them as what they are
# is reading the file correctly, not normalising a failure away.
WHITESPACE = {"s": " ", "tab": "\t", "line-break": "\n"}

# Paragraph containers in the ODF text flow.
PARA = {"p", "h"}

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "dc": "http://purl.org/dc/elements/1.1/",
    "xlink": "http://www.w3.org/1999/xlink",
}
PREFIX = {v: k for k, v in NS.items()}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# A word: a run of non-whitespace. Nothing else is normalised — no case folding,
# no punctuation stripping, no stemming — because an axis that tidies away its
# own failures measures nothing. Unicode NFC is the one equivalence applied, and
# it is applied to both sides: the same character can be encoded two ways, and
# that difference is encoding, not loss.
WORD = re.compile(r"\S+")


def q(ns, tag):
    return "{%s}%s" % (NS[ns], tag)


def local(tag):
    return tag.rsplit("}", 1)[-1] if "}" in str(tag) else str(tag)


def prefixed(tag):
    """The element as ODF spells it, e.g. 'draw:text-box' — the form the reader's
    report uses, so an ancestry chain and a report line can be compared."""
    t = str(tag)
    if not t.startswith("{"):
        return t
    uri, name = t[1:].split("}", 1)
    prefix = PREFIX.get(uri)
    if prefix is None:
        # Never an empty prefix: cutting a namespace URI that ends in '/' yields
        # nothing, and ':creator' names no vocabulary and no element at all.
        prefix = uri.rstrip("/:#").rsplit("/", 1)[-1].rsplit(":", 1)[-1] or "unknown"
    return "%s:%s" % (prefix, name)


def norm(word):
    return unicodedata.normalize("NFC", word)


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


# ---------------------------------------------------------------------------
# the reader's own report
# ---------------------------------------------------------------------------
REPORTED = re.compile(r"\b((?:text|draw|table|office|style|svg|dc|meta|fo|number|xlink)"
                     r":[A-Za-z][\w-]*)")
# A line that promises the text survives. Held to that promise by axis 4.
KEPT_HINT = ("text kept", "value kept", "text is kept", "its text is kept",
             "into its cached text", "to its cached text")
# A line that neither drops the construct nor promises its text: a tab
# flattened to a space, a list assumed bullet, a citation anchor dropped while
# the citation stays. Such a line is no licence for a word to vanish, so it
# excuses nothing — it is the case a careless reader of the report would let pass.
NEUTRAL_HINT = ("assumed bullet", "flattened", "joined with a line break",
                "anchor dropped", "assumed")


def report_names(err):
    """Local names the reader's own report mentions (axis 2's reading)."""
    return set(re.findall(r"(?:text|draw|table|office|style|svg|dc|meta|fo|number|xlink)"
                          r":([A-Za-z][\w-]*)", err))


def report_claims(err):
    """(dropped, kept): prefixed names the report says were dropped / were kept."""
    dropped, kept = set(), set()
    for line in err.splitlines():
        if not line.strip().startswith("[warning]"):
            continue
        names = set(REPORTED.findall(line))
        if not names:
            continue
        low = line.lower()
        if any(h in low for h in KEPT_HINT):
            kept |= names
        elif any(h in low for h in NEUTRAL_HINT):
            continue
        elif "drop" in low:
            dropped |= names
    return dropped, kept


# ---------------------------------------------------------------------------
# axis 4: text fidelity
# ---------------------------------------------------------------------------
def source_words(roots):
    """(Counter of words, [(word, ancestry)]) from the source, independently read.

    Deliberately dumb: paragraphs and their text, joined in document order, each
    token carrying the chain of prefixed element names that held it. No style is
    consulted, no construct is interpreted — the only way to be wrong here is to
    disagree with what the file plainly says.
    """
    words, origin = collections.Counter(), []
    paragraphs = []

    def chunks(el, anc, acc):
        name = prefixed(el.tag)
        anc2 = anc + (name,)
        n = local(el.tag)
        if n in NOTES_ASIDE:
            return
        if n in WHITESPACE:
            times = 1
            if n == "s":
                try:
                    times = int(el.get(q("text", "c")) or 1)
                except (TypeError, ValueError):
                    times = 1
            acc.append((WHITESPACE[n] * max(1, min(times, 64)), anc2))
            return
        if el.text:
            acc.append((el.text, anc2))
        for kid in el:
            chunks(kid, anc2, acc)
            if kid.tail:
                acc.append((kid.tail, anc2))

    def walk(el, anc):
        name = prefixed(el.tag)
        anc2 = anc + (name,)
        if local(el.tag) in NOTES_ASIDE:
            return
        if local(el.tag) in PARA:
            acc = []
            if el.text:
                acc.append((el.text, anc2))
            for kid in el:
                chunks(kid, anc2, acc)
                if kid.tail:
                    acc.append((kid.tail, anc2))
            paragraphs.append(acc)
            return
        for kid in el:
            walk(kid, anc2)

    for root in roots:
        if root is None:
            continue
        for flow in text_flows(root):
            walk(flow, ())

    for acc in paragraphs:
        text, spans, pos = [], [], 0
        for chunk, anc in acc:
            spans.append((pos, pos + len(chunk), anc))
            text.append(chunk)
            pos += len(chunk)
        joined = "".join(text)
        for m in WORD.finditer(joined):
            at = spans[-1][2] if spans else ()
            for start, end, anc in spans:
                if start <= m.start() < end:
                    at = anc
                    break
            w = norm(m.group())
            words[w] += 1
            origin.append((w, at))
    return words, origin, len(paragraphs)


def _top_paragraphs(el, out):
    """Every w:p that is not inside another w:p (text boxes nest)."""
    if el.tag == W + "p":
        out.append(el)
        return
    for kid in el:
        _top_paragraphs(kid, out)


def docx_words(path):
    """(Counter of words, paragraph count) from the body and every header/footer.

    A header or footer is a separate part, so the document's words are its body
    plus all of them — the same scope the source side reads.
    """
    words, paras = collections.Counter(), 0
    if not os.path.exists(path):
        return words, 0
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return words, 0
    with z:
        names = [n for n in z.namelist()
                 if re.fullmatch(r"word/(document|header\d*|footer\d*)\.xml", n)]
        for n in names:
            try:
                root = ET.fromstring(z.read(n))
            except (ET.ParseError, KeyError):
                continue
            found = []
            _top_paragraphs(root, found)
            for p in found:
                text = []
                for el in p.iter():
                    if el.tag == W + "t":
                        text.append(el.text or "")
                    elif el.tag in (W + "tab", W + "br"):
                        text.append(" ")
                joined = "".join(text)
                for m in WORD.finditer(joined):
                    words[norm(m.group())] += 1
                paras += 1
    return words, paras


def word_losses(src, origin, out, dropped, kept):
    """(silent, false_claims, compared) — the words the produced document lost.

    A missing word is explained only when every lost occurrence sat inside a
    construct the report named as dropped. If it sat inside a construct the
    report promised to keep, the report and the document disagree, which is its
    own failure: a false claim is worse than a silence, because it is a silence
    with a receipt.
    """
    by_word = collections.defaultdict(list)
    for w, anc in origin:
        by_word[w].append(anc)

    silent, false_claims = {}, {}
    compared = sum(src.values())
    for w, n in src.items():
        lost = n - out.get(w, 0)
        if lost <= 0:
            continue
        explained, promised = 0, set()
        for anc in by_word.get(w, ()):
            names = set(anc) | {local(a) for a in anc}
            if names & dropped:
                explained += 1
            elif names & kept:
                promised |= (names & kept)
        if lost <= explained:
            continue
        if promised:
            false_claims[w] = (lost - explained, sorted(promised))
        else:
            silent[w] = lost - explained
    return silent, false_claims, compared


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

    root = styles = None
    with zipfile.ZipFile(path) as z:
        try:
            root = ET.fromstring(z.read("content.xml"))
        except (KeyError, ET.ParseError):
            root = None
        try:
            styles = ET.fromstring(z.read("styles.xml"))
        except (KeyError, ET.ParseError):
            styles = None

    reported = report_names(err)
    dropped, kept = report_claims(err)
    silent = silent_drops(root, reported) if root is not None else {}

    work = tempfile.mkdtemp(prefix="corpus-")
    try:
        src = os.path.join(work, os.path.basename(path))
        shutil.copy2(path, src)
        rc_get, _out, get_err, to_get = run([binary(), "get", src, "/body"])
        docx = os.path.splitext(src)[0] + ".docx"
        rc_val = None
        if os.path.exists(docx):
            rc_val, _vo, _ve, _ = run([binary(), "validate", docx], timeout=120)
        leaks = leaked_text(root, reported, docx_body_text(docx)) if root is not None else {}
        produced = os.path.exists(docx)
        words_src = words_out = compared = 0
        word_silent, word_false = {}, {}
        if produced and (root is not None or styles is not None):
            s_words, s_origin, _sp = source_words([root, styles])
            o_words, _op = docx_words(docx)
            words_src, words_out = sum(s_words.values()), sum(o_words.values())
            word_silent, word_false, compared = word_losses(
                s_words, s_origin, o_words, dropped, kept)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    refused = rc_get != 0 and any(s in get_err for s in REFUSAL)
    crash = "Traceback (most recent call last)" in err or \
            "Traceback (most recent call last)" in get_err
    return dict(name=os.path.basename(path), rc_dump=rc_dump, rc_get=rc_get, rc_val=rc_val,
                refused=refused, produced=produced, crash=crash,
                timeout=to_dump or to_get, silent=silent, leaks=leaks,
                words_src=words_src, words_out=words_out, compared=compared,
                word_silent=word_silent, word_false=word_false,
                warns=[l for l in err.strip().splitlines() if l.strip()], get_err=get_err)


def _sample(d, limit=6):
    items = sorted(d.items())
    shown = ", ".join("%r x%d%s" % (w, v[0] if isinstance(v, tuple) else v,
                                    ("" if not isinstance(v, tuple) else "/" + ",".join(v[1])))
                      for w, v in items[:limit])
    return shown + ("" if len(items) <= limit else ", +%d more" % (len(items) - limit))


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
    if r["word_silent"]:
        flags.append("LOST-WORDS %d (%s)" % (sum(r["word_silent"].values()),
                                             _sample(r["word_silent"])))
    if r["word_false"]:
        flags.append("FALSE-CLAIM %d (%s)" % (sum(v[0] for v in r["word_false"].values()),
                                              _sample(r["word_false"])))
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

    manifest = {}
    if a.dirs:
        files = sorted(f for d in a.dirs for f in glob.glob(os.path.join(d, "*.odt")))
    else:
        if not os.path.isdir(FILES) or not glob.glob(os.path.join(FILES, "*.odt")):
            print("no documents in %s — run: python3 tests/corpus/fetch.py" % FILES)
            return 2
        files = sorted(glob.glob(os.path.join(FILES, "*.odt")))
        if os.path.exists(SOURCES):
            with open(SOURCES, encoding="utf-8") as fh:
                manifest = {d["file"]: d for d in json.load(fh)["documents"]}
    if not files:
        print("no .odt documents found")
        return 2

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
        words = ("%d/%d" % (r["compared"] - sum(r["word_silent"].values())
                            - sum(v[0] for v in r["word_false"].values()), r["compared"])
                 if r["compared"] else "-")
        print("%-28s %-22s %-24s %-11s %s"
              % (os.path.basename(f), (meta.get("producer") or "-")[:24], state, words,
                 " ".join(flags) if flags else "ok"))
        if a.verbose or not ok:
            for line in r["warns"]:
                print("      " + line)
        rows.append(r)

    converted = sum(1 for r in rows if r["rc_get"] == 0 and r["rc_val"] == 0)
    refused = sum(1 for r in rows if r["refused"] and not (r["rc_get"] == 0 and r["rc_val"] == 0))
    producers = len({m.get("producer") for m in manifest.values() if m.get("producer")})
    compared = sum(r["compared"] for r in rows)
    lost = sum(sum(r["word_silent"].values()) for r in rows)
    false = sum(sum(v[0] for v in r["word_false"].values()) for r in rows)
    print("\n%d documents%s | %d convert and validate | %d refused with a reason | "
          "%d silent drops | %d crashes | %d timeouts"
          % (len(rows), (" (%d producers)" % producers) if producers else "",
             converted, refused,
             sum(1 for r in rows if r["silent"]), sum(1 for r in rows if r["crash"]),
             sum(1 for r in rows if r["timeout"])))
    print("%d words compared | %d lost without a word (%d document(s)) | "
          "%d lost after being promised (%d document(s))"
          % (compared, lost, sum(1 for r in rows if r["word_silent"]),
             false, sum(1 for r in rows if r["word_false"])))
    if failures:
        print("\n%d document(s) failed" % failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
