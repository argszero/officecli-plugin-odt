#!/usr/bin/env python3
"""Build a minimal .odt that either carries one feature or does not.

The suite's unit is a *cell*: one feature, measured twice. `--feature X` puts
the construct in; `--without` builds the control, which is the same document
with the construct replaced by its plainest equivalent. A cell that cannot
tell those two apart is not a test, so both halves are built here rather than
written by hand twice.

Every feature is expressed only through ODF, so the fixture depends on nothing
but the standard library.

    python3 tests/make_cell_fixture.py --feature hyperlink --out a.odt
    python3 tests/make_cell_fixture.py --feature hyperlink --without --out b.odt
"""
import argparse
import zipfile

FEATURES = ("hyperlink", "mailto", "merged-cells", "vertical-text",
            "footnote", "endnote")

HEAD = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 office:version="1.2">
 <office:automatic-styles>
  <style:style style:name="TStrike" style:family="text">
   <style:text-properties style:text-line-through-style="solid"/>
  </style:style>
  <style:style style:name="TSuper" style:family="text">
   <style:text-properties style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="TSub" style:family="text">
   <style:text-properties style:text-position="sub 58%"/>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:text>
'''

TAIL = '''  </office:text>
 </office:body>
</office:document-content>
'''


def body(feature, on):
    """The document body for one feature, in its feature-on or control form."""
    if feature == "hyperlink":
        if on:
            return ('   <text:p>See <text:a xlink:type="simple" '
                    'xlink:href="https://example.com/odt">the site</text:a> now.</text:p>\n')
        # Control: the same sentence, the link markup removed. The text is
        # identical, so a reader that gains a hyperlink here invented one.
        return '   <text:p>See the site now.</text:p>\n'

    if feature == "mailto":
        if on:
            return ('   <text:p>Write to <text:a xlink:type="simple" '
                    'xlink:href="mailto:team@example.com">the team</text:a>.</text:p>\n')
        return '   <text:p>Write to the team.</text:p>\n'

    if feature == "merged-cells":
        if on:
            # 2x3 grid. Row 1: one cell spanning two columns, then one cell.
            # Column 1 of row 1 spans both rows.
            return (
                '   <table:table table:name="Merged">\n'
                '    <table:table-row>\n'
                '     <table:table-cell table:number-columns-spanned="2" '
                'table:number-rows-spanned="2"><text:p>Corner</text:p></table:table-cell>\n'
                '     <table:table-cell><text:p>Head</text:p></table:table-cell>\n'
                '     <table:covered-table-cell/>\n'
                '    </table:table-row>\n'
                '    <table:table-row>\n'
                '     <table:covered-table-cell/>\n'
                '     <table:table-cell><text:p>Body</text:p></table:table-cell>\n'
                '    </table:table-row>\n'
                '   </table:table>\n')
        # Control: the same three strings in a plain 2x3 grid, merges removed.
        return (
            '   <table:table table:name="Plain">\n'
            '    <table:table-row>\n'
            '     <table:table-cell><text:p>Corner</text:p></table:table-cell>\n'
            '     <table:table-cell><text:p>Head</text:p></table:table-cell>\n'
            '     <table:table-cell/>\n'
            '    </table:table-row>\n'
            '    <table:table-row>\n'
            '     <table:table-cell/>\n'
            '     <table:table-cell><text:p>Body</text:p></table:table-cell>\n'
            '     <table:table-cell/>\n'
            '    </table:table-row>\n'
            '   </table:table>\n')

    if feature == "vertical-text":
        if on:
            return ('   <text:p>Water is H<text:span text:style-name="TSub">2</text:span>O '
                    'and E=mc<text:span text:style-name="TSuper">2</text:span>, '
                    'but this is <text:span text:style-name="TStrike">wrong</text:span>.</text:p>\n')
        # Control: identical text, no character styles applied.
        return ('   <text:p>Water is H2O and E=mc2, '
                'but this is wrong.</text:p>\n')

    if feature in ("footnote", "endnote"):
        # The note sits mid-sentence, so the marker's position is measurable:
        # "Body text" <marker> " more text." The control keeps the visible
        # sentence and drops the note markup entirely — including its body, so
        # a reader that leaks the body into the paragraph is caught by the
        # positive half's "body is not flattened" assertion.
        if on:
            return ('   <text:p>Body text<text:note text:note-class="%s" text:id="nt1">'
                    '<text:note-citation>1</text:note-citation>'
                    '<text:note-body><text:p>The note body.</text:p></text:note-body>'
                    '</text:note> more text.</text:p>\n' % feature)
        return '   <text:p>Body text more text.</text:p>\n'

    raise SystemExit("unknown feature: %s (known: %s)" % (feature, ", ".join(FEATURES)))


MANIFEST = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
'''


def build(feature, on, out):
    content = HEAD + body(feature, on) + TAIL
    with zipfile.ZipFile(out, "w") as z:
        # mimetype must be first and stored, or the package is not an ODF one.
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/manifest.xml", MANIFEST, zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", content, zipfile.ZIP_DEFLATED)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--feature", required=True, choices=FEATURES)
    ap.add_argument("--without", action="store_true",
                    help="build the feature-absent control instead")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    build(a.feature, not a.without, a.out)


if __name__ == "__main__":
    main()
