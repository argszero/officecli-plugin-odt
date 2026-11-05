# The corpus

Documents this reader did not write, used to measure it. Every other fixture in
`tests/` was produced by the same hand as the reader, so a shared mistake — a
construct read from the wrong package part, a report never emitted — stays
invisible in them. This directory is the answer to that: `tests/corpus.py` runs
the reader over the documents named here and judges three things per file.

## What is not here, and why

**The documents themselves are not committed.** They come from projects whose
licences do not allow their bytes to be copied into an MIT repository:

| source | repository | licence | documents |
|---|---|---|---|
| odfpy | [`eea/odfpy`](https://github.com/eea/odfpy) | GPL-2.0 | 19 |
| LibreOffice QA | [`LibreOffice/core`](https://github.com/LibreOffice/core) `sw/qa/extras/odfimport/data` | MPL-2.0 | 12 |
| pandoc | [`jgm/pandoc`](https://github.com/jgm/pandoc) `test/odt` | GPL-2.0-or-later | 8 |

So what is committed is the **manifest**: `sources.json` names each document by
repository, commit, path, size and sha256, and `fetch.py` reconstitutes the exact
bytes into `files/` (git-ignored). A commit, not a branch, is what is pinned, so
a run measures the same document next month. Nothing is redistributed; the
bytes are fetched by whoever runs the suite.

`files/` may also hold `.odt` files written here by hand. Those are ours and are
committed.

## Running it

```bash
python3 tests/corpus/fetch.py                  # once, and after a manifest change
OFFICECLI_BIN=/path/to/officecli python3 tests/corpus.py
```

`fetch.py` uses `gh api` where the GitHub CLI is authenticated and a plain HTTPS
GET of `raw.githubusercontent.com` otherwise — the CI runner has the second and
not always the first. Every transport is checked against the manifest's sha256;
a mismatch is a failure, not a warning.

## What a document is judged on

1. **It converts.** `officecli get <doc> /body` exits 0 and `validate` on the
   produced `.docx` exits 0. A document the reader *refuses* is not a failure
   when the refusal explains itself (a package that is not well-formed XML, or
   that is not an ODF text document) — it is recorded as a refusal, with the
   reason.
2. **Every construct it cannot convert is reported.** For every element in the
   document's text flow (the body, and any header/footer on the master page) that
   the reader does not turn into content, the reader's own report must name it. A
   construct that disappears without a word is the defect this file exists to
   catch: the same class as a footnote body folded into its sentence, which is
   invisible in a document that reads fine.
3. **Nothing crashes or hangs.** No `Traceback`, no timeout, on either the dump
   or the conversion.

There is a fourth, cheap check: a construct reported as *dropped* must not have
its text in the produced document. A warning that says "dropped" while the text
lands in the body is worse than no warning.

## The denominator

The suite prints it every run. As of 0.6.0, on macOS arm64 with officecli
1.0.160: **39 documents, 22 producers** (OpenOffice.org 2.0–3.1, OpenOffice 4.x,
NeoOffice 2.x–3.0, LibreOffice 3.5–5.2, LibreOfficeDev alphas, pandoc, odfpy) —
38 convert and validate, 1 is refused with its reason, **0 silent drops**,
0 crashes, 0 timeouts.

The silent drops are the interesting number, because they were 8 files / 5
constructs in 0.5.0. Which constructs, and the files that showed them, are in the
README's coverage table.
