# officecli-plugin-odt

A **dump-reader plugin** for [officecli](https://github.com/iOfficeAI/OfficeCLI) that
converts OpenDocument Text (`.odt`, LibreOffice Writer) into `.docx` — text,
runs, headings, alignment, lists, tables (including **merged cells**),
**images** and **hyperlinks**.

`officecli get report.odt /body` converts the file once, writes the sibling
`report.docx`, and every later command works on that `.docx`. The plugin speaks
[`plugins/plugin-protocol.md` v1](https://github.com/iOfficeAI/OfficeCLI/blob/main/plugins/plugin-protocol.md)
§2.1/§5.1: a one-shot process that streams officecli batch commands as JSONL on
stdout, which main replays into a fresh native file.

**Version 0.3.0.** Verified against **officecli 1.0.160** (upstream `main` at
[`c56512c8`](https://github.com/iOfficeAI/OfficeCLI/commit/c56512c8)) and against
**1.0.158** ([`5e3e2627`](https://github.com/iOfficeAI/OfficeCLI/commit/5e3e2627))
for the 0.2.0 behaviour, on macOS arm64 with Python 3.13. The suites need a
binary; no other build tooling is involved.

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

**Measured (0.3.0):** path 2 works end to end — a clean
`HOME=<dir>` with `plugin` at `~/.officecli/plugins/dump-reader/odt/plugin` is
listed by `officecli plugins list` (name `officecli-odt`, version `0.3.0`) and a
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
- **bullet** and **ordered** lists (`text:list-style` → `liststyle=bullet|ordered`)
- simple tables (`table:table` → `add --type table`, text into each cell), and
  **merged cells** — `table:number-columns-spanned` → `colspan` (`w:gridSpan`)
  and `table:number-rows-spanned` → `vmerge=restart` / `vmerge=continue`
  (`w:vMerge`). A `table:covered-table-cell` is the slot a merge swallowed, so
  it is consumed by its neighbour's span instead of becoming a cell of its own
- **hyperlinks** — `text:a` with an absolute `xlink:href` → `add --type
  hyperlink`, which gives a `w:hyperlink` element plus an external relationship
  (`TargetMode="External"`) pointing at the URL. `mailto:` targets work the same
  way
- **images** — `draw:frame`/`draw:image` for **PNG, JPEG, GIF, BMP, TIFF and
  SVG**, placed inline where they sit in the text flow (see below)

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
  text box is dropped
- **floating images**: everything is placed inline. A frame anchored to the page
  or the paragraph keeps its size but loses its position, and the plugin says so
- images inside a **table cell** (a cell carries text only; the surrounding text
  of that cell survives)
- image formats outside the six above — EMF and WMF in particular, which have no
  magic bytes this reader can trust
- footnotes and endnotes (`text:note`)
- fields, cross-references, bookmarks, tracked changes (flattened to text)
- nested tables
- **internal anchors** — `text:a xlink:href="#name"` points at a
  `text:bookmark`, and bookmarks are not converted, so the text is kept and the
  link is dropped (with a warning). A `w:anchor` pointing at a bookmark that
  does not exist would be a link that lands nowhere, so it is not emitted
- `text:a` with a **relative** href, which would need a relationship to a part
  this conversion does not carry
- true sublists — a nested list item keeps its `level` tag, so it renders at the
  right depth, but it starts a new numbering instance rather than joining the
  parent list
- page styles, headers, footers, sections' column layout
- table cell formatting beyond text (bold/italic inside a cell is flattened)
- `.fodt` (flat XML) and password-protected packages

`.odt` files that are really a spreadsheet or presentation package are refused
with exit 3 rather than converted into a wrong-shaped document.

## Verify it yourself

Two suites, both standard-library Python, both driven by `OFFICECLI_BIN`:

```bash
OFFICECLI_BIN=/path/to/officecli python3 tests/verify.py   # 38 checks
OFFICECLI_BIN=/path/to/officecli python3 tests/cells.py    # 4 cells, 51 assertions
```

`tests/verify.py` is the end-to-end pass over one big fixture
(`tests/make_fixture.py` — real PNG/SVG bytes, a second reference to one image,
an EMF the reader must refuse). It converts through discovery path 1, then
asserts on the produced `.docx`: it validates, every paragraph's text is present
exactly once, headings/styles/alignment/lists/table cells read back correctly,
the images land byte-for-byte with their geometry, order and alt text, the
sibling is reused without reconversion — plus three negative arms (no plugin
installed, `protocol` ≠ 1, corrupt input), each of which must fail *clearly*
rather than crash. **38 checks, 0 failed.**

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

**4 cells, 51 assertions, 0 failed** — and the same suite run against the
previous `0.2.0` reader gives **0 of 4 cells passed, 17 failed assertions**,
with every control half still green on that arm. That is the measurement that
makes a green run mean something: the failures are the features, not the
harness.

### `officecli plugins lint`

The gate the upstream tool provides, run against every fixture:

```bash
officecli plugins lint ./plugin --fixture tests/…      # → Result: OK
```

It replays `plugin dump` and checks every emitted `--prop` key against the
target format's own schema tree. All four cell fixtures and the big fixture
(43 batch items) report `OK — every emitted prop is declared in the docx schema`
on the version this README documents.

The two "no plugin" arms in `verify.py` need nothing else to resolve a `.odt`
dump-reader, so if you already have this plugin installed at discovery path 2 or
3 the harness reports them as `SKIP` with the path that shadowed them, rather
than failing — a false failure here would accuse the binary of a defect it does
not have.

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
