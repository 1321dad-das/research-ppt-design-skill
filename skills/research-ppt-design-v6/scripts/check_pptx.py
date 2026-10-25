#!/usr/bin/env python3
"""Read-only PPTX structure/size inspection. No size cap unless explicitly supplied."""

import argparse
import json
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def inspect(path, max_bytes=None):
    size = path.stat().st_size
    issues = []
    warnings = []
    if max_bytes is not None and size > max_bytes:
        issues.append(f"File exceeds budget by {size - max_bytes:,} bytes.")
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            issues.append(f"ZIP CRC failure: {bad}")
        infos = archive.infolist()
        names = set(archive.namelist())
        if len(names) != len(infos):
            issues.append("ZIP contains duplicate entry names.")
        for required in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"):
            if required not in names:
                issues.append(f"Required part missing: {required}")

        # Resolve relationship targets without extracting or modifying any file.
        for name in sorted(names):
            if not name.endswith(".rels"):
                continue
            root = ET.fromstring(archive.read(name))
            base = posixpath.dirname(posixpath.dirname(name))
            for rel in root.findall(f"{{{REL_NS}}}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "").split("#", 1)[0]
                if not target:
                    continue
                resolved = (
                    posixpath.normpath(target.lstrip("/"))
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join(base, target))
                )
                if resolved not in names:
                    issues.append(f"Missing relationship target: {name} -> {target}")

        slide_names = sorted(
            (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"slide(\d+)\.xml$", n).group(1)),
        )
        if not slide_names:
            issues.append("No slide XML parts found.")
        slides = []
        for name in slide_names:
            root = ET.fromstring(archive.read(name))
            text_nodes = root.findall(f".//{{{A_NS}}}t")
            item = {
                "part": name,
                "text_runs": len(text_nodes),
                "text_characters": sum(len(n.text or "") for n in text_nodes),
                "shapes": len(root.findall(f".//{{{P_NS}}}sp")),
                "pictures": len(root.findall(f".//{{{P_NS}}}pic")),
                "connectors": len(root.findall(f".//{{{P_NS}}}cxnSp")),
                "graphic_frames": len(root.findall(f".//{{{P_NS}}}graphicFrame")),
                "hidden": root.get("show") in ("0", "false"),
            }
            if not text_nodes:
                warnings.append(f"{name}: no native text runs; review editability if required.")
            slides.append(item)
        media = sorted(
            (
                {"part": i.filename, "bytes": i.file_size, "compressed_bytes": i.compress_size}
                for i in infos
                if i.filename.startswith("ppt/media/") and not i.is_dir()
            ),
            key=lambda i: i["compressed_bytes"],
            reverse=True,
        )
        embeds = [n for n in sorted(names) if n.startswith(("ppt/embeddings/", "ppt/fonts/"))]
        if embeds:
            warnings.append("Embedded files/fonts present; review necessity, license and size.")
        return {
            "filename": path.name,
            "file_bytes": size,
            "limit_bytes": max_bytes,
            "remaining_bytes": max_bytes - size if max_bytes is not None else None,
            "slide_part_count": len(slides),
            "media_compressed_bytes": sum(m["compressed_bytes"] for m in media),
            "media": media,
            "embedded_parts": embeds,
            "slides": slides,
            "issues": issues,
            "warnings": warnings,
            "structural_check_passed": not issues,
            "not_checked": [
                "Scientific correctness and sources",
                "Actual rendering, missing glyphs, clipping and overlap",
                "Readability and visual quality",
                "Target-application editability and compatibility",
                "Relationship completeness beyond declared targets",
            ],
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--max-bytes", type=int, default=None,
                        help="Optional user-specified byte limit; no limit by default.")
    parser.add_argument("--json", action="store_true", help="Print a machine-readable report.")
    args = parser.parse_args()
    if args.max_bytes is not None and args.max_bytes <= 0:
        parser.error("--max-bytes must be positive")
    try:
        report = inspect(args.pptx, args.max_bytes)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, RuntimeError) as exc:
        print(f"Cannot inspect file: {exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        if report["limit_bytes"] is None:
            print(f"{report['filename']}: {report['file_bytes']:,} bytes; no size limit set")
            print(f"Slide parts: {report['slide_part_count']}")
        else:
            print(f"{report['filename']}: {report['file_bytes']:,} / {report['limit_bytes']:,} bytes")
            print(f"Slide parts: {report['slide_part_count']}; budget remaining: {report['remaining_bytes']:,} bytes")
        print("Largest media entries (compressed bytes):")
        for item in report["media"][:10]:
            print(f"  {item['compressed_bytes']:>10,}  {item['part']}")
        for issue in report["issues"]:
            print(f"FAIL: {issue}")
        for warning in report["warnings"]:
            print(f"REVIEW: {warning}")
        print("Appearance, science, readability and application compatibility require separate review.")
        print("Structural/budget result: " + ("PASS" if report["structural_check_passed"] else "FAIL"))
    return 0 if report["structural_check_passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
