#!/usr/bin/env python3
"""Build a small .odt exercising every construct the plugin declares it supports."""
import sys, zipfile, os

CONTENT = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
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
    <table:table-row><table:table-cell><text:p>EMEA</text:p></table:table-cell><table:table-cell><text:p>1200</text:p></table:table-cell></table:table-row>
   </table:table>
   <text:p>After the table</text:p>
  </office:text>
 </office:body>
</office:document-content>
'''

MANIFEST = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
'''

def main(out):
    with zipfile.ZipFile(out, "w") as z:
        # mimetype must be the first entry and stored uncompressed
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/manifest.xml", MANIFEST, zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", CONTENT, zipfile.ZIP_DEFLATED)
    print("wrote", out, os.path.getsize(out), "bytes")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "tests/fixture.odt")
