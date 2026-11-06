#!/usr/bin/env python3
"""Measure the reader against documents it did not make.

Every fixture under `tests/` was written by the same hand as the reader. This
suite runs it over documents produced by OpenOffice.org, OpenOffice, NeoOffice,
LibreOffice and pandoc — the files `tests/corpus/sources.json` names — and judges
each one on five things:

  1. it converts, and the produced .docx validates (a clean refusal is recorded
     as a refusal, with its reason);
  2. every construct the reader does not convert is **named in its own report** —
     a construct that vanishes silently is one defect this suite exists to catch;
  3. **no word is lost in silence** — the reason this suite grew a word axis:
     a document can convert, validate, report every construct it skipped, and
     still swallow a sentence. Words are compared between an independent
     extraction of the source and the produced document; a missing word passes
     only when the reader named the construct that held it as dropped. A word
     the reader *promised* to keep and then lost fails too.
  4. **no block is lost in silence** — the axis word fidelity cannot supply: a table flattened into loose paragraphs, a nesting level
     collapsed and two paragraphs merged all score 100% on words. The source's
     block sequence (paragraph / heading with its level / list item with its
     depth / table with its row-by-column grid / image and the paragraph holding
     it) is read independently and compared **block by block, in order**,
     against the reader's own command stream. A structural difference passes only
     when the report names the construct that caused it. Totals are never
     compared: two documents with the same number of blocks in a different order
     are not the same document.
  5. nothing crashes or hangs.

Plus one cheap extra: a construct reported as dropped must not have its text in
the produced document.

    python3 tests/corpus/fetch.py                    # once
    OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py [--verbose]

Exit status is 1 when any document fails one of them, so CI can gate on it.
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


# ---------------------------------------------------------------------------
# axis 5: block structure
# ---------------------------------------------------------------------------
# Three pools, because a report says one of three different things about a
# construct: it is gone with everything inside it, only its *text* is gone, or it
# survives in a shallower shape. Collapsing them excuses a picture on a line that
# only ever promised to drop text — on this corpus that one collapse turned four
# correct conversions into eleven reported differences, so the pools are kept
# apart rather than tidied into one.
REMOVED_HINT = ("dropped", "not a package part", "missing from the package")
TEXT_ONLY_HINT = ("its text is dropped",)
RECONVERTED_HINT = ("not converted",)

PARA_BLOCKS = ("p", "h", "li")


DROPPED_COUNT = re.compile(r"skipped (\d+) x ")


def report_structure(err):
    """(whole, text_gone, kept, drops) — names, by what the report says became of them.

    whole: the construct and everything in it reached nothing.
    text_gone: its text is gone, so the block it formed reached nothing either —
      but a picture it merely *held* is not excused by this line, which is why
      this pool is kept apart from the first.
    kept: not converted, its text is kept — the block reaches the output in
      another shape, so it consumes an output block without a kind to compare.

    drops carries the count the report itself declared per dropped construct. A
    name cannot say *which* of a hundred pictures went, so a census that excused
    on the name would excuse every picture in the document the moment one of them
    was named — the same always-fires mistake, reached from the other side.
    """
    whole, text_gone, kept, drops = set(), set(), set(), {}
    for line in err.splitlines():
        if not line.strip().startswith("[warning]"):
            continue
        names = set(REPORTED.findall(line))
        if not names:
            continue
        low = line.lower()
        if any(h in low for h in RECONVERTED_HINT):
            kept |= names
        elif any(h in low for h in TEXT_ONLY_HINT):
            text_gone |= names
        elif any(h in low for h in REMOVED_HINT):
            whole |= names
            m = DROPPED_COUNT.search(line)
            n = int(m.group(1)) if m else 1
            for name in names:
                drops[name] = drops.get(name, 0) + n
    return whole, text_gone, kept, drops


def _runs_text(el, skip_draw):
    """The element's text, with text:s/tab/line-break spelled as they render."""
    acc = []
    if el.text:
        acc.append(el.text)

    def tail(kid):
        # Whatever followed a skipped element is still its parent's own text: a
        # caption sits as often after the frame that illustrates it as before.
        if kid.tail:
            acc.append(kid.tail)

    def walk(node):
        for kid in node:
            n = local(kid.tag)
            if skip_draw and (prefixed(kid.tag).startswith("draw:")
                              or n in ("table", "note")):
                tail(kid)
                continue
            if n in WHITESPACE:
                times = 1
                if n == "s":
                    try:
                        times = int(kid.get(q("text", "c")) or 1)
                    except (TypeError, ValueError):
                        times = 1
                acc.append(WHITESPACE[n] * max(1, min(times, 64)))
                tail(kid)
                continue
            if kid.text:
                acc.append(kid.text)
            walk(kid)
            if kid.tail:
                acc.append(kid.tail)

    walk(el)
    return " ".join("".join(acc).split())


def plain_text(el):
    """Everything the element says, as a reader sees it."""
    return _runs_text(el, False)


def own_text(el):
    """A paragraph's own text: its runs, not what the frames it carries hold.

    A text box inside a paragraph is a construct of its own, which the report
    names; the text inside it is not the paragraph's text, and counting it here
    would make a paragraph look kept while the reader had dropped its box.
    """
    return _runs_text(el, True)


def _pictures(el, anc):
    """[(own, ancestry)] for every picture in a subtree, notes aside."""
    out = []

    def walk(node, a):
        n = local(node.tag)
        if n in NOTES_ASIDE:
            return
        a2 = a + (prefixed(node.tag),)
        if n == "image":
            out.append((a2[-1], a))
            return
        for kid in node:
            walk(kid, a2)

    walk(el, anc)
    return out


def _table_grid(el):
    """(rows, columns) for the table's own grid, row groups transparent."""
    rows = []

    def rows_of(node):
        for kid in node:
            n = local(kid.tag)
            if n == "table-row":
                rows.append(kid)
            elif n in ("table-header-rows", "table-row-group"):
                rows_of(kid)

    rows_of(el)
    width = 0
    for r in rows:
        w = 0
        for c in r:
            if local(c.tag) in ("table-cell", "covered-table-cell"):
                try:
                    span = int(c.get(q("table", "number-columns-spanned")) or 1)
                except (TypeError, ValueError):
                    span = 1
                w += max(1, span)
        width = max(width, w)
    return len(rows), width


def source_blocks(root, styles):
    """(body, pictures, decorations) — the source's own account of its structure.

    body is the block sequence: paragraph, heading with its level, list item with
    its depth, table with its grid. A paragraph holding nothing but the pictures
    anchored to it is not a block on either side — the picture is, and counting
    the carrier would report a block the file never had.

    pictures is a census rather than a sequence, because extraction and anchoring
    legitimately disagree about *where* a picture lands: the reader places a
    paragraph-anchored frame inline, so its position in the output is not the
    position of its element in the source. Whether it arrived is still a fact.

    decorations are the header and footer parts that carry text, which live in
    the page layouts of styles.xml and nowhere near a text flow.
    """
    body, decos = [], []

    def block(kind, detail, anc, own):
        return dict(kind=kind, detail=detail, anc=anc, own=own)

    def walk(el, anc, depth):
        n = local(el.tag)
        a2 = anc + (prefixed(el.tag),)
        if n in NOTES_ASIDE:
            return
        if n in ("h", "p"):
            if n == "h":
                try:
                    lvl = int(el.get(q("text", "outline-level")) or 1)
                except (TypeError, ValueError):
                    lvl = 1
                body.append(block("h", max(1, min(6, lvl)), anc, prefixed(el.tag)))
                return
            if _pictures(el, anc) and not own_text(el):
                return                      # a carrier: the picture is the block
            body.append(block("li" if depth else "p",
                              depth - 1 if depth else None, anc, prefixed(el.tag)))
            return
        if n in ("frame", "image", "g"):
            for kid in el:
                walk(kid, a2, depth)
            return
        if n == "list":
            for kid in el:
                k = local(kid.tag)
                if k == "list-header":
                    for k2 in kid:
                        walk(k2, a2, depth)
                elif k == "list-item":
                    for k2 in kid:
                        if local(k2.tag) == "list":
                            walk(k2, a2, depth + 1)
                        else:
                            walk(k2, a2, depth + 1 if depth else 1)
                else:
                    walk(kid, a2, depth)
            return
        if n == "table":
            rows, cols = _table_grid(el)
            body.append(block("table", "%dx%d" % (rows, cols), anc, prefixed(el.tag)))
            return
        for kid in el:
            walk(kid, a2, depth)

    pictures = []
    if root is not None:
        for flow in text_flows(root)[:1]:
            for kid in flow:
                walk(kid, (), 0)
            pictures = _pictures(flow, ())
    return body, pictures, decos


def _para_kind(pr):
    """(kind, detail) for a paragraph command: heading level, list depth, plain."""
    if pr.get("liststyle"):
        try:
            return "li", int(pr.get("level") or 0)
        except (TypeError, ValueError):
            return "li", 0
    style = str(pr.get("style", ""))
    if style.startswith("Heading"):
        try:
            return "h", int(style[len("Heading"):] or 1)
        except ValueError:
            return "h", 1
    return "p", None


def stream_blocks(dumped):
    """(body, pictures, decorations) read from the reader's own command stream.

    The stream is the artifact's product, so this side is what the reader asked
    officecli to build — not what a re-reading of the package claims it built.
    """
    seq, decos = [], []
    pictures = 0
    last = None
    for line in dumped.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            o = json.loads(line)
        except ValueError:
            continue
        if o.get("command") != "add":
            continue
        t = o.get("type")
        pr = o.get("props") or {}
        if t == "paragraph":
            if (o.get("parent") or "") not in ("/body", "/"):
                last = None
                continue                    # a table cell's or a part's paragraph
            rec = [_para_kind(pr)[0], _para_kind(pr)[1], 0, 0]
            seq.append(("para", rec))
            last = rec
        elif t == "run":
            if last is not None and (pr.get("text") or "").strip():
                last[2] += 1
        elif t == "picture":
            pictures += 1
            if last is not None:
                last[3] += 1
        elif t == "table":
            seq.append(("table", "%sx%s" % (pr.get("rows"), pr.get("cols"))))
            last = None

    body = []
    for tag, val in seq:
        if tag == "table":
            body.append(("table", val))
            continue
        kind, detail, runs, pics = val
        if pics and not runs:
            continue                        # a carrier: the picture is the block
        body.append((kind, detail))
    return body, pictures, decos


def _hit(b, pool):
    """Whether one of the block's *containers* is one of these report names.

    A block's own element never counts. text:p is an ancestor of every paragraph,
    draw:image is the element of every picture, and table:table of every table: a
    single report line naming one of them would excuse its whole class, and an
    excuse that always fires proves nothing. What the report says about a whole
    class is a count, and the census reads it as one.
    """
    return set(b["anc"]) & pool


def block_losses(src, out, whole, text_gone, kept):
    """(silent, excused, compared) — the body sequence, block by block and in order.

    A construct the report dropped contributes no output block, so it leaves the
    expected sequence rather than shifting it; one the report says was *not
    converted* but kept its text does reach the output in another shape, so it
    stays and consumes one output block with no kind to compare. Totals are never
    compared: two documents with the same number of blocks in a different order
    are not the same document.
    """
    expected, excused = [], 0
    for b in src:
        if _hit(b, whole) or _hit(b, text_gone):
            excused += 1
            continue
        expected.append((b["kind"], b["detail"], bool(_hit(b, kept)), b))

    silent = []
    i = j = 0
    while i < len(expected) and j < len(out):
        kind, detail, loose, b = expected[i]
        if (kind, detail) == out[j] or loose:
            i += 1
            j += 1
            continue
        if _hit(b, kept):
            excused += 1
        else:
            silent.append((i, (kind, detail), out[j], b))
        i += 1
        j += 1
    while i < len(expected):
        kind, detail, loose, b = expected[i]
        if loose:
            excused += 1
        else:
            silent.append((i, (kind, detail), None, b))
        i += 1
    while j < len(out):
        silent.append((j, None, out[j], None))
        j += 1
    return silent, excused, max(len(expected), len(out))


def picture_losses(src, out, whole, drops):
    """(silent, excused, compared) — did each picture arrive, not where it landed.

    A picture inside a construct the report dropped needs no allowance. The rest
    are covered by the number the report declared for draw:image: the name cannot
    say which of them went, and the count can.
    """
    kept = [p for p in src if not _hit(dict(kind="img", anc=p[1], own=p[0]), whole)]
    gone_with_container = len(src) - len(kept)
    unread = len(kept) - out
    allowance = drops.get("draw:image", 0)
    if unread <= 0:
        if unread < 0:
            return [(0, ("img", len(kept)), ("img", out), None)], gone_with_container, len(src)
        return [], gone_with_container, len(src)
    if unread <= allowance:
        return [], gone_with_container + unread, len(src)
    return ([(0, ("img", out + allowance), ("img", out), None)],
            gone_with_container + allowance, len(src))


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
        struct_silent, struct_excused, blocks_cmp, pics_cmp = [], 0, 0, 0
        # Not `and dumped`: a document whose every block the report dropped
        # emits an empty command stream, and guarding on it would skip the
        # structure check for exactly the conversion that kept nothing — the
        # failure this axis exists to catch.
        if produced:
            whole, text_gone, kept, drops = report_structure(err)
            s_body, s_pics, s_deco = source_blocks(root, styles)
            o_body, o_pics, o_deco = stream_blocks(dumped)
            lost, excused, cmp_body = block_losses(
                s_body, o_body, whole, text_gone, kept)
            struct_silent += lost
            struct_excused += excused
            lost, excused, cmp_pics = picture_losses(s_pics, o_pics, whole, drops)
            struct_silent += lost
            struct_excused += excused
            blocks_cmp = cmp_body + cmp_pics
            pics_cmp = cmp_pics
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
                struct_silent=struct_silent, struct_excused=struct_excused,
                blocks=blocks_cmp, pics=pics_cmp,
                warns=[l for l in err.strip().splitlines() if l.strip()], get_err=get_err)


def _sample(d, limit=6):
    items = sorted(d.items())
    shown = ", ".join("%r x%d%s" % (w, v[0] if isinstance(v, tuple) else v,
                                    ("" if not isinstance(v, tuple) else "/" + ",".join(v[1])))
                      for w, v in items[:limit])
    return shown + ("" if len(items) <= limit else ", +%d more" % (len(items) - limit))


def _struct_sample(items, limit=4):
    shown = []
    for i, want, got, b in items[:limit]:
        shown.append("#%d %s -> %s" % (i, want, got))
    return ", ".join(shown) + ("" if len(items) <= limit else ", +%d more" % (len(items) - limit))


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
    if r["struct_silent"]:
        flags.append("STRUCT %d (%s)" % (len(r["struct_silent"]),
                                         _struct_sample(r["struct_silent"])))
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
    blocks = sum(r["blocks"] for r in rows)
    pics = sum(r["pics"] for r in rows)
    silent_blocks = sum(len(r["struct_silent"]) for r in rows)
    excused_blocks = sum(r["struct_excused"] for r in rows)
    print("%d words compared | %d lost without a word (%d document(s)) | "
          "%d lost after being promised (%d document(s))"
          % (compared, lost, sum(1 for r in rows if r["word_silent"]),
             false, sum(1 for r in rows if r["word_false"])))
    print("%d body block(s) compared | %d lost without a name (%d document(s)) | "
          "%d explained by the report | %d picture(s) accounted for"
          % (blocks, silent_blocks, sum(1 for r in rows if r["struct_silent"]),
             excused_blocks,
             pics))
    if failures:
        print("\n%d document(s) failed" % failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
