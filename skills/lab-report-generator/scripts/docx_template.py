#!/usr/bin/env python3
"""Inspect and patch verified plain-text DOCX slots without rebuilding the package."""

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
from zipfile import ZipFile, BadZipFile

from lxml import etree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def tag(name):
    return f"{{{W}}}{name}"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def load_package(path):
    if path.suffix.lower() != ".docx":
        raise ValueError("Expected .docx; convert legacy .doc first.")
    with ZipFile(path) as archive:
        infos = archive.infolist()
        if len({i.filename for i in infos}) != len(infos):
            raise ValueError("Duplicate package members are unsupported.")
        data = {info.filename: archive.read(info) for info in infos}
        comment = archive.comment
    if "word/document.xml" not in data:
        raise ValueError("Missing word/document.xml.")
    if any(n.startswith("_xmlsignatures/") for n in data):
        raise ValueError("Signed template: ask the user before invalidating signatures.")
    return infos, data, comment


def parse(data):
    root = ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError("DTD-bearing XML is unsupported.")
    if root.tag.startswith("{http://purl.oclc.org/ooxml/"):
        raise ValueError("Strict OOXML is unsupported; convert a copy using Word first.")
    return root


def text_parts(data):
    return sorted(n for n in data if n.startswith("word/") and n.endswith(".xml")
                  and (n == "word/document.xml" or Path(n).name.startswith(
                      ("header", "footer", "footnotes", "endnotes"))))


def nearest_paragraph(element):
    return next((a for a in element.iterancestors() if a.tag == tag("p")), None)


def visible_text(paragraph):
    chunks = []
    for element in paragraph.iterdescendants():
        if nearest_paragraph(element) is not paragraph:
            continue
        if element.tag == tag("t"):
            chunks.append(element.text or "")
        elif element.tag in {tag("br"), tag("cr")}:
            chunks.append("\n")
        elif element.tag == tag("tab"):
            chunks.append("\t")
    return "".join(chunks)


def node_path(element):
    indices = []
    while element.getparent() is not None:
        parent = element.getparent()
        indices.append(parent.index(element) + 1)
        element = parent
    return "/*[1]" + "".join(f"/*[{i}]" for i in reversed(indices))


def edit_blocker(paragraph):
    if any(a.tag in {tag(n) for n in ("sdt", "ins", "del", "moveFrom", "moveTo")}
           for a in paragraph.iterancestors()):
        return "content control or tracked revision ancestor"
    for child in paragraph:
        if child.tag == tag("pPr"):
            if child.find("w:rPr/w:del", NS) is not None:
                return "tracked paragraph deletion"
            continue
        if child.tag != tag("r"):
            return "non-text paragraph structure"
        if any(c.tag not in {tag(n) for n in ("rPr", "t", "tab", "br", "cr")}
               for c in child):
            return "field, drawing, or another non-text run"
        if any(c.tag == tag("br") and c.get(tag("type"), "textWrapping") != "textWrapping"
               for c in child):
            return "explicit page or column break"
        if child.find("w:rPr/w:rPrChange", NS) is not None:
            return "tracked run-format change"
    if paragraph.find("w:pPr/w:pPrChange", NS) is not None:
        return "tracked paragraph-format change"
    return None


def inspect(path):
    _, data, _ = load_package(path)
    result = {"template_sha256": digest(path.read_bytes()), "paragraphs": [],
              "parts": {n: {"bytes": len(b), "sha256": digest(b)} for n, b in data.items()},
              "sections": [], "tables": [], "fields": [], "controls": [],
              "media": [n for n in data if n.startswith("word/media/")], "revisions": 0}
    for name in text_parts(data):
        root = parse(data[name])
        for paragraph in root.iter(tag("p")):
            blocker = edit_blocker(paragraph)
            style = paragraph.find("w:pPr/w:pStyle", NS)
            result["paragraphs"].append({"part": name, "path": node_path(paragraph),
                "text": visible_text(paragraph), "style": style.get(tag("val")) if style is not None else None,
                "editable": blocker is None, "blocker": blocker})
        for section in root.iter(tag("sectPr")):
            result["sections"].append({"part": name, "path": node_path(section),
                "xml": ET.tostring(section, encoding="unicode")})
        for table in root.iter(tag("tbl")):
            result["tables"].append({"part": name, "path": node_path(table),
                "rows": len(table.findall("w:tr", NS)),
                "cells": len(table.findall(".//w:tc", NS))})
        for field in root.iter(tag("instrText")):
            result["fields"].append({"part": name, "instruction": field.text or ""})
        for field in root.iter(tag("fldSimple")):
            result["fields"].append({"part": name, "instruction": field.get(tag("instr"), "")})
        for control in root.iter(tag("sdt")):
            result["controls"].append({"part": name, "path": node_path(control)})
        result["revisions"] += sum(1 for e in root.iter()
            if e.tag in {tag("ins"), tag("del"), tag("moveFrom"), tag("moveTo")})
    return result


def write_run(run, text):
    for child in list(run):
        if child.tag != tag("rPr"):
            run.remove(child)
    buffer = []

    def flush():
        if buffer:
            node = ET.SubElement(run, tag("t"))
            node.set(XML_SPACE, "preserve")
            node.text = "".join(buffer)
            buffer.clear()

    for character in text:
        if character in "\n\t":
            flush()
            ET.SubElement(run, tag("br" if character == "\n" else "tab"))
        else:
            buffer.append(character)
    flush()


def run_text(run):
    result = []
    for child in run:
        if child.tag == tag("t"):
            result.append(child.text or "")
        elif child.tag in {tag("br"), tag("cr")}:
            result.append("\n")
        elif child.tag == tag("tab"):
            result.append("\t")
    return "".join(result)


def set_text(paragraph, value):
    runs = paragraph.findall("w:r", NS)
    if not runs:
        runs = [ET.SubElement(paragraph, tag("r"))]
    write_run(runs[0], value)
    for run in runs[1:]:
        write_run(run, "")


def replace_text(paragraph, old, new):
    original = visible_text(paragraph)
    if not old or original.count(old) != 1:
        raise ValueError("Replacement old text must occur exactly once in the paragraph.")
    start = original.index(old)
    end = start + len(old)
    offset = 0
    inserted = False
    for run in paragraph.findall("w:r", NS):
        current = run_text(run)
        next_offset = offset + len(current)
        if offset < end and next_offset > start:
            prefix = current[:max(0, start - offset)]
            suffix = current[max(0, end - offset):]
            replacement = prefix + (new if not inserted else "") + suffix
            write_run(run, replacement)
            inserted = True
        offset = next_offset
    if not inserted:
        raise ValueError("Could not locate replacement inside editable runs.")


def fill(source, plan_path, output):
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError("Output must be a new path different from the source.")
    if output.suffix.lower() != ".docx":
        raise ValueError("Output must end in .docx.")
    plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    if plan.get("template_sha256") != digest(source.read_bytes()):
        raise ValueError("Template SHA-256 mismatch; inspect the current template again.")
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise ValueError("Plan requires a non-empty operations list.")
    infos, data, comment = load_package(source)
    if "word/settings.xml" in data:
        settings = parse(data["word/settings.xml"])
        protection = settings.find("w:documentProtection", NS)
        if protection is not None and protection.get(tag("enforcement"), "0") in {"1", "true", "on"}:
            raise ValueError("Protected document: ask the user before editing.")
    roots, seen = {}, set()
    for operation in operations:
        name, path = operation["part"], operation["path"]
        if name not in text_parts(data):
            raise ValueError(f"Unsupported text part: {name}")
        if (name, path) in seen:
            raise ValueError("Only one operation per paragraph is allowed.")
        seen.add((name, path))
        if name not in roots:
            roots[name] = parse(data[name])
        matches = roots[name].getroottree().xpath(path)
        if len(matches) != 1 or not isinstance(matches[0], ET._Element) or matches[0].tag != tag("p"):
            raise ValueError(f"Path must locate exactly one paragraph: {path}")
        paragraph = matches[0]
        blocker = edit_blocker(paragraph)
        if blocker:
            raise ValueError(f"Unsafe slot {name}:{path}: {blocker}")
        if visible_text(paragraph) != operation["expected"]:
            raise ValueError(f"Original text mismatch at {name}:{path}")
        if ("text" in operation) == ("replace" in operation):
            raise ValueError("Each operation requires exactly one of text or replace.")
        if "text" in operation:
            if not isinstance(operation["text"], str):
                raise ValueError("text must be a string.")
            set_text(paragraph, operation["text"])
        else:
            replacement = operation["replace"]
            if not all(isinstance(replacement[k], str) for k in ("old", "new")):
                raise ValueError("Replacement values must be strings.")
            replace_text(paragraph, replacement["old"], replacement["new"])
    for name, root in roots.items():
        data[name] = ET.tostring(root, xml_declaration=True, encoding="UTF-8",
                                 standalone=root.getroottree().docinfo.standalone)
    output.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix=".docx", dir=output.parent)
    os.close(handle)
    try:
        with ZipFile(temporary, "w") as archive:
            archive.comment = comment
            for info in infos:
                archive.writestr(info, data[info.filename])
        created = False
        try:
            with output.open("xb") as final:
                created = True
                final.write(Path(temporary).read_bytes())
        except BaseException:
            if created:
                output.unlink(missing_ok=True)
            raise
    finally:
        Path(temporary).unlink(missing_ok=True)
    return {"output": str(output), "changed_parts": sorted(roots), "operations": len(operations)}


def compare(before, after):
    _, original, _ = load_package(before)
    _, final, _ = load_package(after)
    return {"added": sorted(final.keys() - original.keys()),
            "removed": sorted(original.keys() - final.keys()),
            "changed": sorted(n for n in original.keys() & final.keys() if original[n] != final[n]),
            "unchanged_count": sum(original[n] == final[n] for n in original.keys() & final.keys())}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    check = commands.add_parser("inspect")
    check.add_argument("source", type=Path)
    check.add_argument("--output", type=Path)
    patch = commands.add_parser("fill")
    patch.add_argument("source", type=Path)
    patch.add_argument("--plan", type=Path, required=True)
    patch.add_argument("--output", type=Path, required=True)
    diff = commands.add_parser("compare")
    diff.add_argument("before", type=Path)
    diff.add_argument("after", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "inspect":
            result = inspect(args.source)
        elif args.command == "fill":
            result = fill(args.source, args.plan, args.output)
        else:
            result = compare(args.before, args.after)
        content = json.dumps(result, ensure_ascii=False, indent=2)
        if args.command == "inspect" and args.output:
            if args.output.resolve() == args.source.resolve():
                raise ValueError("Inspection output cannot replace the source.")
            args.output.parent.mkdir(parents=True, exist_ok=True)
            with args.output.open("x", encoding="utf-8") as stream:
                stream.write(content + "\n")
        else:
            print(content)
        return 0
    except (ValueError, OSError, KeyError, TypeError, BadZipFile, ET.Error) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
