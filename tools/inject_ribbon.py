#!/usr/bin/env python3
"""Inject the Slide Aid custom ribbon (incl. icons) into a .pptm/.ppam.

Usage:  python3 tools/inject_ribbon.py "Slide Aid.ppam"
        python3 tools/inject_ribbon.py --make-ppam "Slide Aid.pptm"

Office files are zip archives. This script:
  1. adds ribbon/customUI14.xml as customUI/customUI14.xml
  2. embeds every PNG from ribbon/images/ as customUI/images/<name>.png
  3. writes customUI/_rels/customUI14.xml.rels linking image names
  4. registers the customUI part in _rels/.rels
  5. adds a png content type to [Content_Types].xml if missing

--make-ppam: additionally converts a .pptm into a .ppam add-in
(same package; only the main content type and extension differ).
Used by the BuildSlideAid macro, because Mac PowerPoint VBA cannot
SaveAs directly to the add-in format.

Run it AFTER saving the file from PowerPoint (re-run after every
re-save, since PowerPoint strips parts it didn't write).
"""
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

CT_PPTM = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
CT_PPAM = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOMUI_RELTYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
IMAGE_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

def main() -> None:
    args = sys.argv[1:]
    make_ppam = "--make-ppam" in args
    args = [a for a in args if a != "--make-ppam"]
    if len(args) != 1:
        sys.exit(__doc__)
    target = Path(args[0])
    if not target.exists() or target.suffix.lower() not in (".pptm", ".ppam", ".pptx"):
        sys.exit(f"Not a PowerPoint file: {target}")
    if make_ppam and target.suffix.lower() != ".pptm":
        sys.exit("--make-ppam expects a .pptm input")

    root_dir = Path(__file__).resolve().parent.parent
    ribbon_xml = root_dir / "ribbon" / "customUI14.xml"
    images_dir = root_dir / "ribbon" / "images"
    if not ribbon_xml.exists():
        sys.exit(f"Ribbon XML not found: {ribbon_xml}")
    ET.fromstring(ribbon_xml.read_bytes())  # validate before touching the file

    images = sorted(images_dir.glob("*.png")) if images_dir.exists() else []

    # Every image="..." referenced in the XML must exist as a PNG
    import re
    refs = set(re.findall(r'\bimage="([^"]+)"', ribbon_xml.read_text()))
    have = {p.stem for p in images}
    if refs - have:
        sys.exit(f"Missing icon files for: {', '.join(sorted(refs - have))} "
                 f"- run tools/make_icons.py first")

    backup = target.with_suffix(target.suffix + ".bak")
    shutil.copy2(target, backup)

    tmp = target.with_suffix(target.suffix + ".tmp")
    try:
        _rewrite(target, tmp, ribbon_xml, images, make_ppam)
    except BaseException:
        tmp.unlink(missing_ok=True)   # don't leave a half-written .tmp behind
        raise

    if make_ppam:
        out = target.with_suffix(".ppam")
        tmp.replace(out)
        target.unlink()          # remove the intermediate .pptm
        print(f"Ribbon + {len(images)} icons injected; converted to {out.name}")
    else:
        tmp.replace(target)
        print(f"Ribbon + {len(images)} icons injected into {target} (backup: {backup.name})")

def _rewrite(target: Path, tmp: Path, ribbon_xml: Path, images: list, make_ppam: bool) -> None:
    # customUI relationships part (image name -> images/<name>.png)
    ui_rels_root = ET.Element(f"{{{RELS_NS}}}Relationships")
    for p in images:
        rel = ET.SubElement(ui_rels_root, f"{{{RELS_NS}}}Relationship")
        rel.set("Id", p.stem)
        rel.set("Type", IMAGE_RELTYPE)
        rel.set("Target", f"images/{p.name}")
    ET.register_namespace("", RELS_NS)
    ui_rels = ET.tostring(ui_rels_root, encoding="unicode", xml_declaration=True)

    with zipfile.ZipFile(target, "r") as zin:
        names = zin.namelist()
        rels = zin.read("_rels/.rels").decode("utf-8")
        ctypes = zin.read("[Content_Types].xml").decode("utf-8")

        # Register the customUI part (skip if already present)
        if "customUI" not in rels:
            root = ET.fromstring(rels)
            rel = ET.SubElement(root, f"{{{RELS_NS}}}Relationship")
            rel.set("Id", "slideAidCustomUI")
            rel.set("Type", CUSTOMUI_RELTYPE)
            rel.set("Target", "customUI/customUI14.xml")
            rels = ET.tostring(root, encoding="unicode", xml_declaration=True)

        # Convert pptm -> ppam content type if requested
        if make_ppam:
            ctypes = ctypes.replace(CT_PPTM, CT_PPAM)

        # Ensure png default content type
        if 'Extension="png"' not in ctypes:
            ET.register_namespace("", CT_NS)
            croot = ET.fromstring(ctypes)
            d = ET.SubElement(croot, f"{{{CT_NS}}}Default")
            d.set("Extension", "png")
            d.set("ContentType", "image/png")
            ctypes = ET.tostring(croot, encoding="unicode", xml_declaration=True)

        skip = {"_rels/.rels", "[Content_Types].xml", "customUI/customUI14.xml",
                "customUI/_rels/customUI14.xml.rels"}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in skip or name.startswith("customUI/images/"):
                    continue
                zout.writestr(name, zin.read(name))
            zout.writestr("_rels/.rels", rels)
            zout.writestr("[Content_Types].xml", ctypes)
            zout.writestr("customUI/customUI14.xml", ribbon_xml.read_bytes())
            if images:
                zout.writestr("customUI/_rels/customUI14.xml.rels", ui_rels)
                for p in images:
                    zout.writestr(f"customUI/images/{p.name}", p.read_bytes())

if __name__ == "__main__":
    main()
