#!/usr/bin/env python3
"""Build a small .odt exercising every construct the plugin declares it supports.

Images are generated here as real bytes (hand-built PNGs, a tiny SVG, and a
fake EMF that must be refused) so the fixture depends on nothing but the
standard library.
"""
import struct
import sys
import zipfile
import zlib
import os

CONTENT = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 office:version="1.2">
 <office:automatic-styles>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  <style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
  <style:style style:name="T3" style:family="text"><style:text-properties style:text-underline-style="solid"/></style:style>
  <style:style style:name="PCenter" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>
  <text:list-style style:name="LBullet"><text:list-level-style-bullet text:level="1" text:bullet-char="\u2022"/></text:list-style>
  <text:list-style style:name="LNum"><text:list-level-style-number text:level="1" style:num-format="1"/></text:list-style>
 </office:automatic-styles>
 <office:body>
  <office:text>
   <text:h text:outline-level="1">Quarterly Report</text:h>
   <text:h text:outline-level="2">Summary</text:h>
   <text:p>Plain paragraph with <text:span text:style-name="T1">bold</text:span>, <text:span text:style-name="T2">italic</text:span> and <text:span text:style-name="T3">underline</text:span> runs.</text:p>
   <text:p text:style-name="PCenter">Centered line</text:p>
   <text:p>Text before <draw:frame draw:name="Figure 1" text:anchor-type="as-char" svg:width="2cm" svg:height="1.5cm"><svg:title>Red square</svg:title><draw:image xlink:href="Pictures/red.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame> text after.</text:p>
   <text:p><draw:frame draw:name="Unframed blue"><draw:image xlink:href="Pictures/blue.png"/></draw:frame></text:p>
   <text:p><draw:frame draw:name="Logo" svg:width="1.5cm" svg:height="1cm"><draw:image xlink:href="Pictures/logo.svg"/></draw:frame></text:p>
   <text:p><draw:frame draw:name="Vector" svg:width="1cm" svg:height="1cm"><draw:image xlink:href="Pictures/vector.emf"/></draw:frame></text:p>
   <text:list text:style-name="LBullet">
    <text:list-item><text:p>First bullet</text:p></text:list-item>
    <text:list-item><text:p>Second bullet</text:p></text:list-item>
   </text:list>
   <text:list text:style-name="LNum">
    <text:list-item><text:p>Step one</text:p></text:list-item>
    <text:list-item><text:p>Step two</text:p></text:list-item>
   </text:list>
   <table:table table:name="Metrics">
    <table:table-row><table:table-cell><text:p>Region</text:p></table:table-cell><table:table-cell><text:p>Revenue</text:p></table:table-cell></table:table-row>
    <table:table-row><table:table-cell><text:p>EMEA</text:p></table:table-cell><table:table-cell><text:p>1200<draw:frame svg:width="0.4cm" svg:height="0.4cm"><draw:image xlink:href="Pictures/red.png"/></draw:frame></text:p></table:table-cell></table:table-row>
   </table:table>
   <draw:frame draw:name="Block level" svg:width="1cm" svg:height="1cm"><draw:image xlink:href="Pictures/green.png"/></draw:frame>
   <text:p>After the table</text:p>
   <text:p>Text before <draw:frame draw:name="Unframed blue again"><draw:image xlink:href="Pictures/blue.png"/></draw:frame> text after.</text:p>
  </office:text>
 </office:body>
</office:document-content>
'''

MANIFEST = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Pictures/red.png" manifest:media-type="image/png"/>
 <manifest:file-entry manifest:full-path="Pictures/blue.png" manifest:media-type="image/png"/>
 <manifest:file-entry manifest:full-path="Pictures/green.png" manifest:media-type="image/png"/>
 <manifest:file-entry manifest:full-path="Pictures/logo.svg" manifest:media-type="image/svg+xml"/>
 <manifest:file-entry manifest:full-path="Pictures/vector.emf" manifest:media-type="image/x-emf"/>
</manifest:manifest>
'''

LOGO_SVG = ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="8">'
            '<rect width="12" height="8" fill="#2a7"/></svg>\n')


def png(width, height, rgb):
    """A minimal valid PNG of one solid colour."""
    def chunk(tag, data):
        body = tag + data
        return (struct.pack(">I", len(data)) + body
                + struct.pack(">I", zlib.crc32(body) & 0xffffffff))
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))


def main(out):
    with zipfile.ZipFile(out, "w") as z:
        # mimetype must be the first entry and stored uncompressed
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/manifest.xml", MANIFEST, zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", CONTENT, zipfile.ZIP_DEFLATED)
        z.writestr("Pictures/red.png", png(3, 3, (255, 0, 0)))
        z.writestr("Pictures/blue.png", png(96, 48, (0, 0, 255)))
        z.writestr("Pictures/green.png", png(2, 2, (0, 200, 0)))
        z.writestr("Pictures/logo.svg", LOGO_SVG)
        # EMF magic (record type 1, size 108): a format the reader refuses by
        # design, because it cannot be sniffed and a wrong MIME would corrupt
        # the document silently.
        z.writestr("Pictures/vector.emf",
                   b"\x01\x00\x00\x00\x6c\x00\x00\x00" + b"\x00" * 100)
    print("wrote", out, os.path.getsize(out), "bytes")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "tests/fixture.odt")
