"""Remove the shield watermark from the Steel King PowerPoint template.

The watermark is a background image on the ten white content layouts.
This replaces each of those backgrounds with a solid white fill and
leaves the dark photo layouts (COVER, SECTION, KEY_MESSAGE, CLOSING,
COVER_MASCOT) untouched.

Run:  python3 strip_watermark.py <source.pptx> <output.pptx>
"""
import sys
from copy import deepcopy
from lxml import etree
from pptx import Presentation

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

DARK = {"COVER", "SECTION", "KEY_MESSAGE", "CLOSING", "COVER_MASCOT", "DEFAULT"}

WHITE_BG = (
    '<p:bg xmlns:p="%s" xmlns:a="%s"><p:bgPr>'
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/>'
    '</p:bgPr></p:bg>' % (NS["p"], NS["a"])
)


def strip(src, out):
    prs = Presentation(src)
    changed = []
    for layout in prs.slide_layouts:
        if layout.name in DARK:
            continue
        cSld = layout._element.find("p:cSld", NS)
        bg = cSld.find("p:bg", NS)
        if bg is None:
            continue
        blip = bg.find(".//a:blip", NS)
        if blip is None:
            continue
        rId = blip.get("{%s}embed" % NS["r"])
        cSld.replace(bg, etree.fromstring(WHITE_BG))
        try:
            layout.part.drop_rel(rId)
        except Exception:
            pass
        changed.append(layout.name)
    prs.save(out)
    return changed


if __name__ == "__main__":
    src, out = sys.argv[1], sys.argv[2]
    print("white background set on:", ", ".join(strip(src, out)))
