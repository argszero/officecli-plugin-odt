# officecli-plugin-odt

A **dump-reader plugin** for [officecli](https://github.com/iOfficeAI/OfficeCLI) that
converts OpenDocument Text (`.odt`, LibreOffice Writer) into `.docx` — text,
runs, headings, alignment, lists (including **nested** lists), tables (including
**merged cells**), **images**, **hyperlinks**, **footnotes / endnotes**, and
**headers / footers**. The [coverage table](#coverage-table) is the one-page
answer to "will my document survive the conversion".

`officecli get report.odt /body` converts the file once, writes the sibling
`report.docx`, and every later command works on that `.docx`. The plugin speaks
[`plugins/plugin-protocol.md` v1](https://github.com/iOfficeAI/OfficeCLI/blob/main/plugins/plugin-protocol.md)
§2.1/§5.1: a one-shot process that streams officecli batch commands as JSONL on
stdout, which main replays into a fresh native file.

**Version 0.9.0.** Verified against **officecli 1.0.160** (upstream `main` at
[`c56512c8`](https://github.com/iOfficeAI/OfficeCLI/commit/c56512c8)) and against
**1.0.158** ([`5e3e2627`](https://github.com/iOfficeAI/OfficeCLI/commit/5e3e2627)),
with **51 checks**, **8 cells / 118 assertions** and a **39-document foreign
corpus** green on both, on macOS arm64 with Python 3.13. That corpus is now judged
by a **foreign arbiter** as well: **odfpy** reads the source `.odt`, **python-docx**
reads the produced `.docx`, and **38 documents agree over 4,386 blocks** with
neither end of the comparison in this repository's code — the fourth suite, and
gating in CI on both arms. On that corpus the
conversion is measured **word by word** — **40,536 source words** compared
against the produced `.docx`, **0 lost without a word, 0 lost after being
promised** — and **block by block**: **4,536 body blocks** (paragraph, heading
with its level, list item with its depth, table with its grid) matched in order,
plus **20 pictures** as a census, **0 lost without a name, 47 explained by the
report**. The same three
suites also run in CI on Linux against the released binaries — see
[Tests in CI](#tests-in-ci). The suites need a binary; no other build tooling is
involved.

## Requirements

- **Python 3.8+** on `PATH` (`#!/usr/bin/env python3`). Standard library only —
  no third-party packages, nothing to install.
- An officecli build that implements plugin protocol v1.

## Install

The plugin is a single executable file named `plugin`. Put it at any of the
discovery paths in protocol §3 — the first match wins:

| # | Path | Measured on this machine |
|---|---|---|
| 1 | `$OFFICECLI_PLUGIN_DUMP_READER_ODT` (absolute path to the file) | ✅ works, **not listed** by `officecli plugins list` |
| 2 | `~/.officecli/plugins/dump-reader/odt/plugin` | ✅ works and **is listed**; the release-asset install path (measured with `HOME` redirected into a scratch dir) |
| 3 | `<dir of the officecli binary>/plugins/dump-reader/odt/plugin` | ✅ works, listed by `officecli plugins list` |
| 4 | an executable `officecli-dump-reader-odt` or `officecli-odt` on `PATH` | not tested |

**Measured (0.9.0):** path 2 works end to end — a clean
`HOME=<dir>` with `plugin` at `~/.officecli/plugins/dump-reader/odt/plugin` is
listed by `officecli plugins list` (name `officecli-odt`, version `0.9.0`) and a
`.odt` converts through it with no environment variable set, producing a
sibling `.docx` that validates. That is the path the release notes use.

So the simplest machine-local install is:

```bash
mkdir -p ~/.officecli/plugins/dump-reader/odt
cp plugin ~/.officecli/plugins/dump-reader/odt/plugin
chmod +x ~/.officecli/plugins/dump-reader/odt/plugin
officecli plugins list          # the plugin should now appear
```

Then:

```bash
officecli get report.odt /body          # converts, then reads the .docx
officecli set report.odt /body/p[1] --prop bold=true   # edits report.docx
```

Conversion is one-shot and cached: `report.docx` is regenerated only when
`report.odt` is newer, or when you delete the sibling. Edits go to the `.docx`,
never back to the `.odt`.

## Script plugins: what works where

The plugin is a Python script with a `#!` line, not a compiled binary. That is
allowed by the protocol (every path is a plain file path, and `runtime` is
declarative), and it is measured to work on Unix:

> `PluginRegistry.cs:140` — `PluginExeNames()` yields `plugin.exe` and `plugin` on
> Windows, and only `plugin` elsewhere, and the file is started with
> `Process.Start(FileName: <path>)`. On macOS/Linux the kernel honours the
> shebang, so an extensionless script named `plugin` runs. On Windows there is no
> `.bat`/`.py` in that name list, so a script plugin is **not** discoverable
> there — this is exactly the gap reported in
> [OfficeCLI issue #202](https://github.com/iOfficeAI/OfficeCLI/issues/202).

On Windows either compile this reader into `plugin.exe`, or use a launcher that
is itself `plugin.exe`. That limitation is upstream's, not this plugin's.

## What it converts

Declared in the manifest `supports` field, and each item is asserted by
`tests/verify.py`:

- paragraphs
- runs, with **bold / italic / underline / strikethrough**, and
  **superscript / subscript** (`style:text-line-through-style`,
  `style:text-position` → `strike`, `vertAlign`)
- heading levels **1–6** (`text:h @text:outline-level` → `style=Heading<N>`)
- alignment (`fo:text-align` → `align=left|center|right|justify`)
- **bullet** and **ordered** lists (`text:list-style` → `liststyle=bullet|ordered`),
  including **nested** ones: a `text:list` inside a `text:list-item` is one docx
  `level` deeper, so an ordered list reads back as `numLevel=0` decimal followed
  by `numLevel=1` lowerLetter instead of one flat run of items
- simple tables (`table:table` → `add --type table`, text into each cell), and
  **merged cells** — `table:number-columns-spanned` → `colspan` (`w:gridSpan`)
  and `table:number-rows-spanned` → `vmerge=restart` / `vmerge=continue`
  (`w:vMerge`). A `table:covered-table-cell` is the slot a merge swallowed, so
  it is consumed by its neighbour's span instead of becoming a cell of its own
- **hyperlinks** — `text:a` with an absolute `xlink:href` → `add --type
  hyperlink`, which gives a `w:hyperlink` element plus an external relationship
  (`TargetMode="External"`) pointing at the URL. `mailto:` targets work the same
  way
- **footnotes and endnotes** — `text:note` → `add --type footnote|endnote` on the
  paragraph that contains it, so the marker lands at the anchor instead of at the
  end of the sentence, and the body goes to `word/footnotes.xml` or
  `word/endnotes.xml`. `text:note-class` picks which: `endnote` is routed to the
  endnote part, everything else to the footnote part. A note holds one paragraph
  in the docx vocabulary, so a multi-paragraph body is joined with a line break
  (`<w:br/>`) and the join is reported. The source's citation number is not
  carried — Word numbers the note itself
- **headers and footers** — the `style:header` / `style:footer` of the master
  page, which ODF keeps in **`styles.xml`** and not in `content.xml` →
  `add --type header|footer` at the document root, which writes
  `word/header1.xml` / `word/footer1.xml` and a `w:headerReference` /
  `w:footerReference` into the section properties. `style:header-left` and
  `style:footer-left` (the mirrored layout's even pages) are carried as the
  `even` scope, and a one-paragraph decoration keeps its paragraph alignment.
  ODF has no "different first page" flag: that is a separate master page, so the
  page the first block names (else `Standard`) is the one converted and the rest
  are reported. Fields keep only their cached text — see below
- **images** — `draw:frame`/`draw:image` for **PNG, JPEG, GIF, BMP, TIFF and
  SVG**, placed inline where they sit in the text flow (see below)

## Coverage table

One row per construct, with what asserts it. `verify.py` is the end-to-end suite
over one big fixture; a **cell** is the two-half test in `cells.py` — the feature
present, and the same document with it removed — with its assertion count in
brackets. Every ✅ is a measured assertion, not an intention: a row says ✅ only
if a suite goes red without the code that earns it (see
[how the ✔s are earned](#verify-it-yourself)).

All of it was measured on **macOS arm64, officecli 1.0.160 and 1.0.158**; the
same suites also run in CI on Linux against the same released binaries
([Tests in CI](#tests-in-ci)).

| construct | 0.9.0 | asserted by | notes |
|---|---|---|---|
| **block structure** — every source block is still a block | ✅ **new in 0.8.0** | `tests/corpus.py` axis 5, 4,536 blocks + 20 pictures | paragraph, heading + level, list item + depth, table + grid and picture matched **in order**; a block may go or change shape only when the report names the construct that held it |
| **foreign arbiter** — an implementation that is not ours agrees on the block sequence | ✅ **new in 0.9.0** | `tests/foreign_axis.py`, 38 documents / 4,386 blocks | source read with **odfpy**, output read with **python-docx**: with neither end in this repository's code, a block the reader drops cannot disappear from both sides at once. 1 clean refusal (`nasty.odt`, not well-formed XML). Gating in CI on both arms — see [the section below](#a-foreign-arbiter-on-the-output) |
| **word fidelity** — every source word accounted for | ✅ **new in 0.7.0** | `tests/corpus.py` axis 4, 40,536 words | a word may be absent only when the report names the construct that held it; a report that promises the text is held to that promise |
| paragraphs and text | ✅ | `verify.py`, every cell | each paragraph's text appears exactly once |
| bold / italic / underline | ✅ | `verify.py` | `fo:font-weight`, `fo:font-style`, `style:text-underline-style` |
| strikethrough, superscript, subscript | ✅ | cell `vertical-text` (14) | `style:text-line-through-style`, `style:text-position` → `strike`, `vertAlign` |
| headings 1–6 | ✅ | `verify.py` | `text:outline-level` → `style=Heading<N>` |
| paragraph alignment | ✅ | `verify.py`, cell `header-footer` (20) | `fo:text-align`, in the body and in a decoration |
| bullet and ordered lists | ✅ | `verify.py` | `text:list-style` → `liststyle=bullet\|ordered` |
| **nested lists** | ✅ **new in 0.5.0** | cell `nested-list` (15) | one `level` deeper per `text:list` inside an item; ordered levels get their own number format |
| simple tables | ✅ | `verify.py` | `table:table` → `add --type table` |
| merged cells | ✅ | cell `merged-cells` (13), `verify.py` | `number-columns-spanned` → `w:gridSpan`, `number-rows-spanned` → `w:vMerge` |
| absolute hyperlinks | ✅ | cell `hyperlink` (12), `verify.py` | `w:hyperlink` plus an external relationship |
| `mailto:` links | ✅ | cell `mailto` (12) | same shape, `mailto:` target |
| footnotes | ✅ | cell `footnote` (16), `verify.py` | `w:footnoteReference` at the anchor, body in `word/footnotes.xml` |
| endnotes | ✅ | cell `endnote` (16), `verify.py` | `word/endnotes.xml`, and *not* the footnote part |
| **headers** | ✅ **new in 0.5.0** | cell `header-footer` (20), `verify.py` | master page in `styles.xml` → `word/header1.xml` + `w:headerReference` |
| **footers** | ✅ **new in 0.5.0** | cell `header-footer` (20), `verify.py` | `word/footer1.xml` + `w:footerReference`, alignment carried |
| even-page (mirrored) header/footer | ✅ | measured by hand, not a cell | `style:header-left` / `style:footer-left` → the `even` scope: 3 parts (`header1`, `header2`, `footer1`), `w:headerReference` types `even`/`default`, `validate` clean |
| images: PNG, JPEG, GIF, BMP, TIFF, SVG | ✅ | `verify.py` | inline, bytes byte-identical, sized from the frame |
| SVG 1×1 PNG fallback | ✅ | `verify.py` | 5 pictures → 6 media parts |
| shapes, charts, embedded objects, text boxes | ❌ | `verify.py`; corpus `ClimateChange2.odt`, `PageBackground.odt`, `fdo68839.odt`, `fdo90130-1.odt` | a text box's text is dropped; reported. **No corpus document carries a chart** |
| floating image position | ❌ | corpus `PageBackground.odt` (page-anchored), `ClimateChange2.odt` (paragraph-anchored) | size kept, anchor lost, reported |
| images inside a table cell | ❌ | corpus `fdo47267-3.odt`; `verify.py` | reported |
| EMF / WMF and other image formats | ❌ | `verify.py` only — **no corpus document carries one** | refused with a warning, never embedded under a guessed MIME |
| fields: page number, page count, date, … | ❌ | corpus `api-for-odfpy.odt`, `headerfooter.odt`, `annotation-formatting.odt` | only the cached text survives, reported |
| bookmarks, internal anchors | ❌ | corpus `api-for-odfpy.odt` (540 start / 540 end), `empty-svg-family-name.odt` | text kept, link dropped, reported |
| **cross-references** (`bookmark-ref`, `reference-ref`, `note-ref`) | ❌ **named in 0.6.0** | corpus `api-for-odfpy.odt` (4,013 of them) | cached text kept, reference dropped, each construct reported by its ODF name |
| **user-field values** (`text:user-field-get`) | ❌ **named in 0.6.0** | corpus `userfields.odt`, `userfields_ooo3.odt` | cached value kept, the field dropped, reported |
| **bibliography marks, form controls, embedded objects** | ❌ **named in 0.6.0** | corpus `BibliographyEntryField.odt`, `citation.odt`, `dateFormFormats.odt`, `fdo68839.odt` | `text:bibliography-mark`, `draw:control`, `draw:object` — each reported by name |
| relative hrefs | ❌ | corpus `api-for-odfpy.odt` | text kept, link dropped, reported |
| tracked changes | ❌ | corpus `fdo47267-3.odt` | flattened to its text, reported |
| nested tables | ❌ | corpus `fdo37606.odt`, `fdo47267-3.odt`, `simpletable.odt` | reported |
| page styles, page size, margins, column layout | ❌ | corpus — a `style:page-layout` in **38 of 39** documents, `style:columns` in 5 (`BibliographyEntryField.odt`, `Electra.odt`, `PageBackground.odt`, `api-for-odfpy.odt`, `fdo53210.odt`) | not converted, not reported |
| table cell formatting beyond text | ❌ | corpus `fdo37606.odt`, `fdo47267-3.odt`, `simpletable.odt` | bold/italic inside a cell is flattened |
| `.fodt`, encrypted packages | ❌ | not a corpus case — refused before anything is read | refused, exit 3 |

## The foreign corpus

Every other fixture in `tests/` is written by the same hand as the reader, so a
construct read from the wrong package part — or skipped without a word — stays
invisible in it. `tests/corpus.py` measures the reader against documents it did
not make: **39 documents from 22 producers** — OpenOffice.org 2.0–3.1, OpenOffice
4.x, NeoOffice 2.x–3.0, LibreOffice 3.5–5.2 and its dev alphas, odfpy, pandoc.

The documents are **not committed**. They come from odfpy (GPL-2.0), LibreOffice
(MPL-2.0) and pandoc (GPL-2.0-or-later), whose licences do not allow their bytes
in an MIT repository, so what is committed is the manifest:
`tests/corpus/sources.json` pins each file by repository, **commit**, path, size
and sha256, and `tests/corpus/fetch.py` reconstitutes the exact bytes and refuses
any that does not hash to the recorded value. A commit rather than a branch is
what is pinned, so a run next month measures the same document. Provenance,
per-source licences and the fetch protocol are in
[`tests/corpus/README.md`](tests/corpus/README.md).

```bash
python3 tests/corpus/fetch.py                             # once, and after a manifest change
OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py
```

Each document is judged on five things, and the last three are what found the
defects this release fixes:

1. **it converts** — `get <doc> /body` exits 0 and the produced `.docx` validates.
   A refusal that explains itself is recorded as a refusal, not as a failure.
2. **every construct it cannot convert is named** — the reader's own report must
   mention it. A construct that leaves the document without a word is the defect
   worth hunting, because it looks exactly like a document that read fine.
3. **nothing crashes or hangs** — no traceback, no timeout.
4. **every source word is accounted for** — the source words are extracted a
   second, deliberately dumb way (paragraphs and their text, with `text:s`,
   `text:tab` and `text:line-break` read as the whitespace they are) and compared
   against the words in the produced `.docx`. A word may be absent only when the
   report names the construct that held it. Two failures are told apart, and the
   second is the one a reader of the report would wave through: **lost without a
   word** (nobody said anything — a silent loss) and **lost after being promised**
   (a report line says the text was kept, and the word is not there anyway).
5. **every block keeps its shape** — the source block sequence is read from
   `content.xml` and compared **in order** against the blocks the reader asked
   officecli to build, so a paragraph, a heading's level, a list item's depth, a
   table's grid and a picture each have to arrive. Absence or reshaping is
   allowed only when the report names the construct that held it; a construct
   reported as *not converted, its text is kept* consumes one output block
   instead of vanishing, which is what a list flattened into plain paragraphs
   looks like from the outside.

Plus one cheap extra: a construct reported as *dropped* must not have its text in
the produced document. A "dropped" that quietly lands is worse than silence.

**Measured on both arms — macOS arm64, 1.0.160 and 1.0.158:** 38 documents convert
and validate, 1 is refused with its reason (`nasty.odt`, an external DTD entity
the well-formedness check rejects rather than parses), **0 silent drops, 0
crashes, 0 timeouts**, **40,536 words compared, 0 lost without a word, 0 lost
after being promised**, and **4,556 blocks compared, 0 lost without a name** —
4,536 body blocks in order, and the corpus's **20 pictures** as a census — with
47 differences explained by the report. Every one of those 20 pictures is
accounted for: 17 arrive as pictures, and 3 are dropped under the name of their
construct and the count the report declares (`fdo68839.odt` and `formula.odt`
each carry an `Object 1` **VCL metafile**; `textWithImages.odt`'s picture has an
external href, so it is not a package part at all). Under 0.5.0 the same corpus had **8 documents with silent
drops** across five constructs: `text:bookmark-ref` (4,013 of them in one
document), `text:user-field-get`, `text:bibliography-mark`, `draw:control` and
`draw:object` inside a frame. 0.6.0 reports all five, reports each by its **ODF
name** rather than its bare local name (`text:note` and `draw:frame` are
different vocabularies, and the prefix is what says which), and closes the class
with a catch-all: an element the reader has no rule for is named, so a construct
cannot leave a document in silence again.

### What the word axis found in 0.7.0

Three defects, none of which any earlier suite could see — all three were
reported by their document, and all three are fixed here:

1. **`text:list-header` was skipped without a word.** In
   `annotation-formatting.odt` *every* paragraph sits in a list header, so the
   whole document converted to an empty `.docx`: `get` exited 0, `validate`
   exited 0, the summary said `emitted 0 paragraph(s)`, and 6 words were gone. A
   list header is the list's lead-in — paragraphs that carry no marker — and
   Word has no such slot, so they become plain paragraphs. The document now
   converts to 3 paragraphs, 6/6 words.
2. **A decoration report printed the wrong vocabulary.** What it dropped from a
   header or footer was always named `text:<name>`, so a `table:table` inside a
   footer (`fdo47267-3.odt`) came out as `text:table` — an element that exists in
   no ODF vocabulary, and therefore a name that could never match the construct
   it named. The real prefix is used now.
3. **The same shape one level down.** `dc:creator` in an annotation report came
   out as `:creator`, because the prefix was derived by cutting a namespace URI
   that ends in `/` — cutting there yields nothing. It prints `dc:creator` now,
   and an unknown namespace can no longer produce an empty prefix.

**The axis is proved able to fail.** Silencing the reader's drop report — a
two-line change in a scratch copy — turns the corpus **red**: exit 1, **286 words
lost without a word across 8 documents, 38 silent drops**, each naming the word
and the document. Reverting only the `text:list-header` fix moves
`annotation-formatting.odt` from `6/6 ok` to `6 lost without a word`, exit 1, while
the rest of the corpus stays green. Both breaks were then reverted to a plugin
byte-identical to the committed one (`shasum -a 256`), because an axis that cannot
go red proves nothing about the documents that pass it.

The same numbers come from a **different platform and a different transport**: the
CI runner has no authenticated `gh`, so `fetch.py` falls through to plain HTTPS,
and there it reports `0 kept, 39 fetched, 0 failed` — every byte matching the
manifest's sha256 — before the suite prints the same 39-document verdict on
Linux with 1.0.160. The development machine cannot reach `raw.githubusercontent.com`
at all, so that half of the fetcher is the runner's measurement, not a local one.

### A foreign arbiter on the output

Every axis above judges the conversion with the same hand that writes it:
`corpus.py` reads the package with code from this repository's own family, and
`officecli validate` is the host's checker. A block the reader drops is then
dropped on both sides of the comparison, and the axis stays green while the output
is short.

`tests/foreign_axis.py` puts implementations that are **not ours on either end** —
**odfpy 1.4.1** reading the source `.odt`, **python-docx 1.2.0** reading the
produced `.docx` — and compares the block sequence in order: paragraph, heading
with its level, list item with its depth, table. Where the host ships Apple's
Office.mdimporter, `mdimport` adds a second, unrelated opinion on whether the
package imports at all; where it does not, the check says so and holds nothing
against the package.

Measured on the 39-document corpus, macOS arm64, officecli 1.0.160:
**38 documents agree over 4,386 text blocks, and 1 is the corpus's own clean
refusal** (`nasty.odt` — `content.xml` is not well-formed XML, which the reader
reports in those words).

Microsoft Word 16.114.4 is installed on the development host and would be the
arbiter to want. It cannot be driven from this process: every document-level Apple
Event answers `-10004` while a version query succeeds, for Word, TextEdit and
Finder alike. The two readers above stand in for it and are named as themselves.

Two rules keep the comparison fair, and both are read out of the reader's own
report rather than out of a list of exceptions — the same rule the block axis
applies, in a different currency:

- a construct the report declares **dropped at block level** expects no output
  block, so its subtree is not counted (`text:table-of-content` in
  `api-for-odfpy.odt`, 37 entries);
- a construct the report declares **"not converted; its text is kept"** comes out
  flattened, so its items are counted as body paragraphs rather than list items
  (`text:list` in `annotation-formatting.odt`).

Pictures are printed but not gated: what a frame carries is not a text block, and
the block axis above already matches pictures in order with the declared drops
excused.

**It was proved able to fail before it was trusted.** Suppressing one line of the
reader — `props["liststyle"] = list_kind` — turns `simplelist.odt`,
`definitionlists.odt` and `twolevellist.odt` red (`source 'li0', output 'p'`, exit
1); restoring the file byte-identically (`f300b52c…`, the 0.8.0 asset) turns them
green again. An axis that cannot go red proves nothing about the documents that
pass it.

```bash
python3 -m pip install odfpy python-docx
python3 tests/corpus/fetch.py
OFFICECLI_BIN=/path/to/officecli python3 tests/foreign_axis.py
```

The reader itself still needs nothing but Python: these two packages are needed
only to read both ends of this check.

**The runner has now carried it, so it gates.** On commit `9bfe02c` the step ran
on both arms and both concluded `success` — the 1.0.158 arm included, whose job is
`continue-on-error` for reasons that concern officecli's Linux asset and not this
check. Its own reading on Linux is the same verdict as the one on macOS:
**39 documents, 38 agree over 4,386 text blocks, 1 refused**. A different platform
and a different transport (`fetch.py` falls through to plain HTTPS on the runner,
where this host cannot reach `raw.githubusercontent.com` at all) reaching the same
4,386 is what makes the number a property of the output rather than of one
machine. The step is gating now, and the two packages it needs are installed by
the workflow; the reader itself still needs only Python.

### What the block axis found in 0.8.0

Word fidelity cannot see structure. Two documents can hold exactly the same words
in the same order and be different documents: one table, or forty paragraphs. This
release adds the axis that reads structure on both sides — the source's block
sequence out of `content.xml`, the reader's out of the commands it emitted — and
compares them **in order**. On the corpus: **4,536 body blocks and 20 pictures
compared, 0 lost without a name, 47 explained by the report.**

**Two defects, one in the reader and one in the axis itself:**

1. **An image part was reported missing while it was in the package.**
   `fdo68839.odt` and `formula.odt` write their `xlink:href` as
   `./ObjectReplacements/Object 1`. The reader used the href as a zip name without
   dropping the `./` prefix, so `zf.read` raised `KeyError` and the report said
   *"image part missing from the package"* — the construct named, the reason
   false. The prefix is stripped now, and what the reader refuses it refuses for
   its real reason: the part is a **`VCLMTF`** metafile, an unsupported format. The
   two image report lines also name their construct (`draw:image …`) instead of a
   bare file name, which is what lets the census excuse them by the count the
   report declares rather than by a name no reader could tie to an element.
2. **The axis skipped exactly the conversion it exists to catch.** `formula.odt`'s
   body is a single framed object, so the reader emits an empty command stream,
   and the structural comparison was guarded on that stream being non-empty — the
   one document that kept nothing was the one the axis never looked at. It now
   runs on every document that produced a file, so an empty stream is compared as
   an empty sequence, and `formula.odt`'s single picture is counted and excused by
   the count the report declares.

**The boundary: decorations are not this axis.** A source header or footer is
declared once per `style:page-layout`, and one layout can serve several master
pages, while docx has one part per section — comparing decoration text per
page-layout against one output part is a **model mismatch**, not a defect. Header
and footer text is the **word axis's** territory, where `styles.xml` is read
independently and every header word is accounted for.

**The axis is proved able to fail, three times.** Each break was confirmed red
against the code in this release, then reverted to a plugin byte-identical to the
committed one (`shasum -a 256`):

| break | words | blocks |
|---|---|---|
| every second paragraph merged into the one before it | **3,787 lost without a word** in 13 documents, 404 after being promised | **2,731 lost without a name** in 15 documents, exit 1 |
| a list item's `level` never emitted, so nesting collapses | **40,536 / 0 / 0 — unchanged** | **20 lost without a name** in 3 documents, exit 1 |
| a table's cells emitted as ordinary paragraphs, no `add --type table` | **40,536 / 0 / 0 — unchanged** | **8 lost without a name** in 2 documents, exit 1 |

The last two rows are this release's whole argument. The document those two breaks
produce is structurally wrong and the word axis cannot tell: every word is there,
in order, accounted for. Only the block axis says the nesting and the grid are
gone.

**An excuse is scoped to the block's own construct, never to a universal
ancestor.** `text:p` is an ancestor of every block in the corpus, so a report line
naming a dropped `text:p` would excuse an entire document if ancestry alone were
enough: the axis refuses a block's own element as its excuse and reads only its
containers. It also excuses a class by the **count the report declares**, not by
the name, because one line naming `draw:image` must not excuse every picture in
the file.

## Images

An image in the source becomes `add --type picture` on the paragraph that
contains it, so `Text before <image> text after.` keeps its three pieces in
order — measured, not assumed: `r[1]` run, `r[2]` picture, `r[3]` run.

**How the bytes travel.** The plugin reads them straight out of the ODT package
(`Pictures/*`, resolved through `xlink:href`) and carries them in the emitted
command as a `data:` URI (`props.src`). That is deliberate, and measured:

- main invokes a dump-reader as `plugin dump <file>` and nothing else — it never
  passes `--media-dir`, although protocol §5.1 allows one (`DumpReaderInvoker.cs`
  builds `Arguments = new[] { "dump", sourceFullPath }`);
- main **replays the buffered commands after the plugin process has exited**, so
  a scratch file written by the plugin would either be deleted too early or leak
  on every conversion;
- a data URI leaves nothing behind and works when the `.odt` sits on a
  read-only volume.

`props.path` (a plain file path) was measured to work too and keeps the original
file name as the picture's `name`; the data URI names it from the package part
unless `<svg:title>`/`<svg:desc>` supplies alt text.

**Size.** `svg:width`/`svg:height` on the frame, in any ODF length unit
(cm/mm/in/pt/pc/px), converted to points — officecli wants the unit, and a bare
number would be read as raw EMU. When the frame carries no geometry the pixel
dimensions are read from the image header (PNG/JPEG/GIF/BMP) at the CSS 96dpi
ratio: a 96×48px image lands as 2.5cm × 1.3cm.

**Formats.** Only the six above: they are the ones officecli's own
`ImageSource.TrySniffContentType` recognises, so the MIME in the data URI always
matches the real bytes. An unknown format is refused with a warning rather than
embedded under a guessed content type — a data URI is trusted by main without a
magic check, so a wrong MIME would produce a schema-valid document with a broken
image in it. An SVG picture additionally carries the 1×1 PNG fallback officecli
generates when `fallback` is not given (5 pictures → 6 media parts).

## What it does **not** convert

The plugin writes a `[warning] odt: …` line to stderr for every construct it
skips, with a count, and `[warning] odt: emitted N paragraph(s), M table(s) and
K image(s)` at the end. **Note where those lines go**: main collects a
dump-reader's stderr but only shows it when the plugin *fails*, so a successful
`officecli get` prints none of them (see the third observation below). To read
the report, run the plugin itself:

```bash
python3 plugin dump report.odt > /dev/null   # stdout is the JSONL command stream
```

- drawings that are not pictures — shapes (`draw:custom-shape`, `rect`, …),
  charts, embedded objects, and text boxes (`draw:text-box`): the text inside a
  text box is dropped. *Corpus:* text boxes in `ClimateChange2.odt`,
  `PageBackground.odt`, `fdo68839.odt`, `fdo90130-1.odt`; embedded objects in
  `fdo68839.odt` and `formula.odt`. No corpus document carries a chart.
- **floating images**: everything is placed inline. A frame anchored to the page
  or the paragraph keeps its size but loses its position, and the plugin says so.
  *Corpus:* `PageBackground.odt` (page-anchored), `ClimateChange2.odt`
  (paragraph-anchored).
- images inside a **table cell** (a cell carries text only; the surrounding text
  of that cell survives). *Corpus:* `fdo47267-3.odt`.
- image formats outside the six above — EMF and WMF in particular, which have no
  magic bytes this reader can trust. *No corpus document carries one*; the
  fixture in `tests/make_fixture.py` does.
- fields, bookmarks and tracked changes — flattened to text, each reported.
  *Corpus:* bookmarks in `api-for-odfpy.odt` (540 start / 540 end) and
  `empty-svg-family-name.odt`; tracked changes in `fdo47267-3.odt`.
- **cross-references** — `text:bookmark-ref`, `text:reference-ref` and
  `text:note-ref` keep their cached text and lose the reference.
  *Corpus:* `api-for-odfpy.odt` carries 4,013 of them, and each one is counted in
  the report.
- **user fields** (`text:user-field-get`), **bibliography marks**
  (`text:bibliography-mark`), **form controls** (`draw:control`) and **embedded
  objects** (`draw:object`, `draw:object-ole`, `draw:plugin`): dropped or
  flattened to their text, reported by ODF name. *Corpus:* `userfields.odt`,
  `userfields_ooo3.odt`, `BibliographyEntryField.odt`, `citation.odt`,
  `dateFormFormats.odt`, `fdo68839.odt`.
- any element the reader has no rule for: reported by name from a catch-all, so
  nothing in the text flow leaves without a word
- nested tables. *Corpus:* `fdo37606.odt`, `fdo47267-3.odt`, `simpletable.odt`.
- **internal anchors** — `text:a xlink:href="#name"` points at a
  `text:bookmark`, and bookmarks are not converted, so the text is kept and the
  link is dropped (with a warning). A `w:anchor` pointing at a bookmark that
  does not exist would be a link that lands nowhere, so it is not emitted
- `text:a` with a **relative** href, which would need a relationship to a part
  this conversion does not carry. *Corpus:* `api-for-odfpy.odt`.
- page styles, page size and margins, and sections' column layout. *Corpus:* a
  `style:page-layout` in 38 of the 39 documents, `style:columns` in
  `BibliographyEntryField.odt`, `Electra.odt`, `PageBackground.odt`,
  `api-for-odfpy.odt` and `fdo53210.odt`.
- **fields** — `text:page-number`, `text:page-count`, `text:date` and the rest.
  A field is computed by the writer, so the package holds only a cached value: a
  footer reading "Page 1 of 3" converts to exactly that text, on every page. The
  plugin reports the flattening rather than implying the field survived.
  *Corpus:* `api-for-odfpy.odt`, `headerfooter.odt`, `annotation-formatting.odt`.
- a header or footer on **any other master page** — a different first page is a
  separate master page in ODF, so the one the first block names is converted and
  the others are reported. *Corpus:* `PageBackground.odt` (1 more),
  `fdo53210.odt` (3 more).
- table cell formatting beyond text (bold/italic inside a cell is flattened).
  *Corpus:* `fdo37606.odt`, `fdo47267-3.odt`, `simpletable.odt`.
- `.fodt` (flat XML) and password-protected packages — refused before anything is
  read, so no corpus document exercises them.

Every ❌ above either names the corpus document that exhibits it or says outright
that the corpus has none — a row whose evidence is a fixture says so.

`.odt` files that are really a spreadsheet or presentation package are refused
with exit 3 rather than converted into a wrong-shaped document.

## Verify it yourself

Three suites, all standard-library Python, all driven by `OFFICECLI_BIN`:

```bash
OFFICECLI_BIN=/path/to/officecli python3 tests/verify.py   # 51 checks
OFFICECLI_BIN=/path/to/officecli python3 tests/cells.py    # 8 cells, 118 assertions
python3 tests/corpus/fetch.py                              # 39 foreign documents, by sha256
OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py   # 39 documents, 22 producers
```

`tests/verify.py` is the end-to-end pass over one big fixture
(`tests/make_fixture.py` — real PNG/SVG bytes, a second reference to one image,
an EMF the reader must refuse). It converts through discovery path 1, then
asserts on the produced `.docx`: it validates, every paragraph's text is present
exactly once, headings/styles/alignment/lists/table cells read back correctly,
the images land byte-for-byte with their geometry, order and alt text, the
sibling is reused without reconversion — plus three negative arms (no plugin
installed, `protocol` ≠ 1, corrupt input), each of which must fail *clearly*
rather than crash. **51 checks, 0 failed.**

`tests/cells.py` is the suite that measures **this version's new coverage**, and
its unit is a *cell*: one feature, measured **twice**. Every feature cell builds
two documents from `tests/make_cell_fixture.py` — the feature present, and the
same document with the construct removed — and a cell counts as passed only when
**all three** hold:

1. the feature appears in the converted `.docx` (the positive half);
2. `officecli validate` on that `.docx` exits 0;
3. the **feature-absent control** does not gain the feature, and still validates.

The third is why a cell is not just a smoke test. Without it, a reader that
invented a `w:hyperlink` everywhere would pass; without the second, one that
produced a package Word refuses to open would pass. Both halves go through the
same discovery path, so the construct under test is the only difference between
them.

| cell | feature on | control (feature absent) |
|---|---|---|
| `hyperlink` | `text:a` + absolute `href` → `w:hyperlink` + external rel | same sentence, no `text:a` → no `w:hyperlink`, no rel |
| `mailto` | `text:a href="mailto:…"` → external rel with the `mailto:` target | same sentence, no `text:a` |
| `merged-cells` | 2×3 grid with a 2-column, 2-row merge → `w:gridSpan`, `w:vMerge` restart + continue | same strings in a plain 2×3 grid → no `w:gridSpan`, no `w:vMerge`, 6 cells |
| `vertical-text` | `text-position` / `text-line-through-style` → `w:vertAlign`, `w:strike` | identical text, no character styles → neither element |
| `footnote` | `text:note class="footnote"` mid-sentence → `w:footnoteReference` at the anchor + `word/footnotes.xml` holding the body | same sentence, note markup removed → no reference, no footnotes part, text unchanged |
| `endnote` | `text:note class="endnote"` → `w:endnoteReference` + `word/endnotes.xml`, and **not** the footnote part | same sentence, note markup removed → no reference, no endnotes part |
| `header-footer` | a master page in **`styles.xml`** with `style:header` + `style:footer` → `word/header1.xml`, `word/footer1.xml`, `w:headerReference`/`w:footerReference`, alignment kept | the same master page and page layout with the two elements removed → no header/footer part, no reference, body unchanged |
| `nested-list` | numbered list with a `text:list` inside an item → `numLevel=0` decimal, then `numLevel=1` with its own format (`lowerLetter`) | the same four items all at the top level → every item `numLevel=0`, one format |

**8 cells, 118 assertions, 0 failed.** Run the same suite against the
**previous** reader and it goes red exactly where the new work is: against
**0.4.0**, **6 of 8 cells pass and 12 assertions fail**, every failure in the
positive half of `header-footer` or `nested-list` and **both control halves still
green** (`0.4.0` read only `content.xml`, so it never saw the master page; and it
passed a nested list's `level` through unchanged, so all four items read back
`numLevel=0`). The same suite at the size it had then puts `0.3.0` at 4 of 6
cells / 16 failed assertions — it read a note's body *into the sentence*
(`"Body text1The note body. more text."`) and wrote no note part at all — and
`0.2.0` at 0 of 6 / 33 failed. That is what makes a green run mean something:
the failures are the features, not the harness.

### `officecli plugins lint`

The gate the upstream tool provides, run against every fixture:

```bash
officecli plugins lint ./plugin --fixture tests/…      # → Result: OK
```

It replays `plugin dump` and checks every emitted `--prop` key against the
target format's own schema tree. All eight cell fixtures (4–8 batch items each)
and the big fixture (55 batch items) report `OK — every emitted prop is declared
in the docx schema` on the version this README documents.

The two "no plugin" arms in `verify.py` need nothing else to resolve a `.odt`
dump-reader, so if you already have this plugin installed at discovery path 2 or
3 the harness reports them as `SKIP` with the path that shadowed them, rather
than failing — a false failure here would accuse the binary of a defect it does
not have.

## Tests in CI

`.github/workflows/test.yml` runs all three suites on every push, in a matrix over
the two released officecli binaries (`1.0.160`, `1.0.158`), downloaded from
upstream releases rather than built — the plugin is measured against the binary a
user installs. The corpus step fetches its documents from the pinned commits and
fails the build on a single silent drop. The workflow also lints every fixture
and, on the newer binary, installs the **published release asset** over HTTPS
into discovery path 2 and converts a note fixture through it with no environment
variable set.

That last job exists because the machine this plugin was developed on cannot
reach `github.com` with `curl` (the connection times out), so the release-asset
line was the one thing the README could not measure locally. The runner measures
it, and the job fails if the asset is not byte-identical to the tree it was built
from (`cmp plugin …`) or if the converted output does not validate. On the 0.8.0
tag run it printed `release tag: v0.8.0`, `officecli plugins list` reported
`officecli-odt 0.8.0`, and the note fixture converted through the installed asset
with no environment variable set.

The workflow installs the **.NET 10 runtime** explicitly, because the released
officecli assets are framework-dependent single files (`Microsoft.NETCore.App`
10.0.13 for 1.0.158, 10.0.14 for 1.0.160) and a runner that only carries an older
runtime still starts the app and answers `--version` — then dies at the first
document with `Could not load file or assembly 'System.IO.Compression.ZipFile,
Version=10.0.0.0'`. That was measured here: two runs of the same commit differed,
one green on both arms and one red on 1.0.158 alone. A `--version` answer is not
evidence that a binary can open a document.

**One arm is non-gating.** The `1.0.158` job stays in the matrix but does not fail
the build: across six runs of unchanged commits its *Linux* asset has passed
twice and failed four times, each failure an assembly load raised on a document —
`Could not load file or assembly 'System.IO.Compression.ZipFile,
Version=10.0.0.0'` on three of them and `'System.Private.Xml, Version=10.0.0.0'` on
the 0.6.0 run — before this plugin has done anything with the output, so that
failure measures the environment rather than the reader. The 1.0.158 numbers in
this README come from macOS arm64, where the same suites are green; the Linux
asset's behaviour is upstream's to fix, not this plugin's.

The 0.8.0 release measured that flakiness in a single cycle rather than
attributing it: the same commit `a7c0c42` was run twice, once from the `v0.8.0`
tag and once from a push to `main`, and the 1.0.158 arm **passed on the tag run
and failed on the `main` run**, with `System.IO.Compression.ZipFile` raised on the
first conversion. Both runs are the same tree; only the runner differed. The tag
run's other arm and every macOS run were green, including the foreign corpus at
40,536 words and 4,536 body blocks with nothing lost.

The 0.9.0 release reproduced that pair a second time, on commit `637f9b7`, and
added a **third** assembly to the class: the `v0.9.0` tag run was green on both
arms — the arbiter reading `39 document(s): 38 agree over 4386 text blocks, 1
REFUSED` on each — while the push-to-`main` run of the same commit failed the
1.0.158 arm at its first conversion with `Could not load file or assembly
'System.Private.Xml.Linq, Version=10.0.0.0'`, before the plugin had produced
anything. Three different assemblies on three different runs is what says the
variable is the Linux asset's runtime resolution and not the document under test:
a single measured name would have been equally consistent with one broken code
path.

## Three upstream observations made while building this

All are measurements on officecli 1.0.158 (`5e3e2627`), reported here because
they affect anyone writing a plugin:

1. **`protocol` mismatch exits 1, not 5.** Protocol §4.1 says "Main rejects
   mismatches with exit code 5", and §6.5 reserves 5 for exactly this. In
   practice `PluginRegistry.TryReadManifest` prints
   `[warning] plugin at … declares protocol=2 but main supports protocol=1;
   plugin will not load.` and returns false, so discovery falls through to the
   generic "Unsupported file type" error — **exit 1**. The warning is clear; only
   the documented exit code differs. Reported upstream as
   [#548](https://github.com/iOfficeAI/OfficeCLI/issues/548).
2. **`officecli plugins list` does not enumerate discovery path 1.** It walks
   `~/.officecli/plugins` and `<binary>/plugins` only, so a plugin installed via
   `$OFFICECLI_PLUGIN_<KIND>_<EXT>` works but is invisible to `plugins list`
   (measured: `[]` when the env var was the only install, and listed once the
   same file was placed under `<binary>/plugins/dump-reader/odt/`).
3. **A dump-reader's stderr is discarded on success.** Protocol §5.1/§5.4 make
   stderr the diagnostics channel, and a reader is expected to say what it
   skipped there — but `DumpReaderInvoker` passes no `OnStderrLine`, so
   `PluginProcess` collects stderr into `RunResult.Stderr` and only `Truncate`s
   it into the message of a *failed* run. Measured: a conversion that skips two
   images prints only `[note] generated report.docx …`; `--verbose` adds nothing;
   `--log-file` (protocol §5.2) is not an officecli option
   (`Unrecognized command or argument '--log-file'`). So the skip report is only
   visible when the plugin is run by hand. If you are writing a reader, do not
   assume your warnings reach the user.

## License

MIT — see [LICENSE](LICENSE).
