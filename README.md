# officecli-plugin-odt

A **dump-reader plugin** for [officecli](https://github.com/iOfficeAI/OfficeCLI) that
converts OpenDocument Text (`.odt`, LibreOffice Writer) into `.docx`.

`officecli get report.odt /body` converts the file once, writes the sibling
`report.docx`, and every later command works on that `.docx`. The plugin speaks
[`plugins/plugin-protocol.md` v1](https://github.com/iOfficeAI/OfficeCLI/blob/main/plugins/plugin-protocol.md)
§2.1/§5.1: a one-shot process that streams officecli batch commands as JSONL on
stdout, which main replays into a fresh native file.

Verified against **officecli 1.0.158**, built from upstream `main` at
[`5e3e2627`](https://github.com/iOfficeAI/OfficeCLI/commit/5e3e2627)
(`dotnet build -c Release src/officecli/officecli.csproj`), on macOS arm64 with
Python 3.13.

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
| 2 | `~/.officecli/plugins/dump-reader/odt/plugin` | not tested (the test sandbox could not write there); same code path as #3 |
| 3 | `<dir of the officecli binary>/plugins/dump-reader/odt/plugin` | ✅ works, listed by `officecli plugins list` |
| 4 | an executable `officecli-dump-reader-odt` or `officecli-odt` on `PATH` | not tested |

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
- runs, with **bold / italic / underline**
- heading levels **1–6** (`text:h @text:outline-level` → `style=Heading<N>`)
- alignment (`fo:text-align` → `align=left|center|right|justify`)
- **bullet** and **ordered** lists (`text:list-style` → `liststyle=bullet|ordered`)
- simple tables (`table:table` → `add --type table`, text into each cell)

## What it does **not** convert

Nothing here is dropped silently: the plugin writes a `[warning] odt: …` line to
stderr for every construct it skips, with a count, plus
`[warning] odt: emitted N paragraph(s) and M table(s)` at the end.

- images and drawings (`draw:frame`, `draw:image`, shapes, charts)
- hyperlinks (`text:a` — the text is kept, the link target is lost)
- footnotes and endnotes (`text:note`)
- fields, cross-references, bookmarks, tracked changes (flattened to text)
- nested tables; tables with merged cells lose the merge
- true sublists — a nested list item keeps its `level` tag, so it renders at the
  right depth, but it starts a new numbering instance rather than joining the
  parent list
- page styles, headers, footers, sections' column layout
- table cell formatting beyond text (bold/italic inside a cell is flattened)
- `.fodt` (flat XML) and password-protected packages

`.odt` files that are really a spreadsheet or presentation package are refused
with exit 3 rather than converted into a wrong-shaped document.

## Verify it yourself

```bash
OFFICECLI_BIN=/path/to/officecli python3 tests/verify.py
```

`tests/verify.py` builds a fixture `.odt` covering every supported construct
(`tests/make_fixture.py`), converts it through discovery path 1, then asserts on
the produced `.docx`: it validates, every paragraph's text is present exactly
once, headings/styles/alignment/lists/table cells read back correctly, the
sibling is reused without reconversion — plus three negative arms (no plugin
installed, `protocol` ≠ 1, corrupt input), each of which must fail *clearly*
rather than crash.

Result on the build named above: **22 checks, 0 failed.**

## Two upstream observations made while building this

Both are measurements on officecli 1.0.158 (`5e3e2627`), reported here because
they affect anyone writing a plugin:

1. **`protocol` mismatch exits 1, not 5.** Protocol §4.1 says "Main rejects
   mismatches with exit code 5", and §6.5 reserves 5 for exactly this. In
   practice `PluginRegistry.TryReadManifest` prints
   `[warning] plugin at … declares protocol=2 but main supports protocol=1;
   plugin will not load.` and returns false, so discovery falls through to the
   generic "Unsupported file type" error — **exit 1**. The warning is clear; only
   the documented exit code differs.
2. **`officecli plugins list` does not enumerate discovery path 1.** It walks
   `~/.officecli/plugins` and `<binary>/plugins` only, so a plugin installed via
   `$OFFICECLI_PLUGIN_<KIND>_<EXT>` works but is invisible to `plugins list`
   (measured: `[]` when the env var was the only install, and listed once the
   same file was placed under `<binary>/plugins/dump-reader/odt/`).

## License

MIT — see [LICENSE](LICENSE).
