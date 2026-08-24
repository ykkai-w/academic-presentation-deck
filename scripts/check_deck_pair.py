#!/usr/bin/env python3
"""Read-only structural preflight for an academic submission/preparation PPTX pair."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import unicodedata
import zipfile
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
RID = f"{{{NS['r']}}}id"
REMBED = f"{{{NS['r']}}}embed"
RLINK = f"{{{NS['r']}}}link"
LONG_NUMBER_RE = re.compile(r"(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d{5,})(?!\d)")
SOURCES_RE = re.compile(r"\[\s*Sources\s*\]", re.IGNORECASE)
MAX_PACKAGE_PARTS = 20_000
MAX_TOTAL_UNCOMPRESSED_BYTES = 2_000_000_000
MAX_SINGLE_PART_BYTES = 500_000_000
MAX_COMPRESSION_RATIO = 1_000
MANUAL_CHECKS_REQUIRED = [
    "render every final slide in both variants and compare at full size",
    "verify claims and displayed numbers against the authoritative paper",
    "inspect font resolution, clipping, overlap, and editability in a PowerPoint-compatible app",
    "time a complete spoken rehearsal at the presenter's realistic pace",
    "review every external relationship and embedded object before high-stakes transfer",
]


def normalized_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", text).strip()


def compact_char_count(text: str) -> int:
    return len(re.sub(r"\s+", "", text))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_slide_spec(value: str) -> set[int]:
    slides: set[int] = set()
    if not value.strip():
        return slides
    for token in value.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            start, end = int(start_text), int(end_text)
            if start < 1 or end < start:
                raise ValueError(f"invalid slide range: {token}")
            slides.update(range(start, end + 1))
        else:
            slide = int(token)
            if slide < 1:
                raise ValueError(f"invalid slide number: {token}")
            slides.add(slide)
    return slides


def relationship_part(owner: str) -> str:
    if not owner:
        return "_rels/.rels"
    directory, name = posixpath.split(owner)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def resolve_target(owner: str, target: str) -> str:
    target = unquote(target.split("#", 1)[0].split("?", 1)[0])
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base = posixpath.dirname(owner)
    return posixpath.normpath(posixpath.join(base, target))


def parse_xml(zf: zipfile.ZipFile, part: str, errors: list[str]) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(part))
    except KeyError:
        errors.append(f"missing required part: {part}")
    except ET.ParseError as exc:
        errors.append(f"invalid XML in {part}: {exc}")
    return None


def relationships_for(
    zf: zipfile.ZipFile,
    owner: str,
    parts: set[str],
    errors: list[str],
) -> dict[str, dict[str, str | bool]]:
    rel_part = relationship_part(owner)
    if rel_part not in parts:
        return {}
    root = parse_xml(zf, rel_part, errors)
    if root is None:
        return {}

    relationships: dict[str, dict[str, str | bool]] = {}
    for rel in root.findall("pr:Relationship", NS):
        rel_id = rel.get("Id", "")
        target = rel.get("Target", "")
        rel_type = rel.get("Type", "")
        external = rel.get("TargetMode", "").lower() == "external"
        if not external and urlsplit(target).scheme:
            external = True
        resolved = target if external else resolve_target(owner, target)
        relationships[rel_id] = {
            "target": resolved,
            "type": rel_type,
            "external": external,
        }
    return relationships


def extract_all_text(root: ET.Element) -> str:
    return "\n".join(
        node.text or "" for node in root.findall(".//a:t", NS) if (node.text or "").strip()
    )


def extract_note_text(root: ET.Element) -> str:
    body_chunks: list[str] = []
    fallback_chunks: list[str] = []
    excluded_types = {"sldImg", "hdr", "ftr", "dt", "sldNum"}

    for shape in root.findall(".//p:sp", NS):
        placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
        placeholder_type = placeholder.get("type", "") if placeholder is not None else ""
        chunk = "\n".join(
            node.text or ""
            for node in shape.findall(".//a:t", NS)
            if (node.text or "").strip()
        ).strip()
        if not chunk:
            continue
        if placeholder_type == "body":
            body_chunks.append(chunk)
        elif placeholder_type not in excluded_types:
            fallback_chunks.append(chunk)

    for frame in root.findall(".//p:graphicFrame", NS):
        chunk = "\n".join(
            node.text or ""
            for node in frame.findall(".//a:t", NS)
            if (node.text or "").strip()
        ).strip()
        if chunk:
            fallback_chunks.append(chunk)

    return "\n".join(body_chunks + fallback_chunks).strip()


def split_sources(note_text: str) -> tuple[str, str, bool]:
    match = SOURCES_RE.search(note_text)
    if not match:
        return note_text.strip(), "", False
    return note_text[: match.start()].strip(), note_text[match.end() :].strip(), True


def transform_signature(element: ET.Element) -> dict:
    candidates = [
        element.find("./p:spPr/a:xfrm", NS),
        element.find("./p:grpSpPr/a:xfrm", NS),
        element.find("./p:xfrm", NS),
    ]
    transform = next((item for item in candidates if item is not None), None)
    if transform is None:
        return {}
    signature: dict = {"attrs": dict(sorted(transform.attrib.items()))}
    for child_name in ("off", "ext", "chOff", "chExt"):
        child = transform.find(f"a:{child_name}", NS)
        if child is not None:
            signature[child_name] = dict(sorted(child.attrib.items()))
    return signature


def style_signature(element: ET.Element) -> list[dict]:
    style_roots = {
        "spPr",
        "grpSpPr",
        "style",
        "bodyPr",
        "lstStyle",
        "pPr",
        "rPr",
        "defRPr",
        "endParaRPr",
        "tblPr",
        "tcPr",
    }
    signatures: list[dict] = []
    for node in element.iter():
        if node.tag.rsplit("}", 1)[-1] in style_roots:
            signatures.append(semantic_element(node))
    return signatures


def semantic_element(element: ET.Element) -> dict:
    attributes: list[tuple[str, str]] = []
    for key, value in element.attrib.items():
        if key in {RID, REMBED, RLINK}:
            value = "__RELATIONSHIP__"
        attributes.append((key, value))
    return {
        "tag": element.tag,
        "attrs": sorted(attributes),
        "text": normalized_text(element.text or ""),
        "children": [semantic_element(child) for child in list(element)],
    }


def semantic_part_hash(
    part: str,
    zf: zipfile.ZipFile,
    cache: dict[str, str],
) -> str:
    if part in cache:
        return cache[part]
    payload = zf.read(part)
    if part.lower().endswith(".xml"):
        try:
            root = ET.fromstring(payload)
            payload = json.dumps(
                semantic_element(root),
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        except ET.ParseError:
            pass
    digest = hashlib.sha256(payload).hexdigest()
    cache[part] = digest
    return digest


def recursive_part_hash(
    part: str,
    zf: zipfile.ZipFile,
    parts: set[str],
    errors: list[str],
    content_cache: dict[str, str],
    graph_cache: dict[str, str],
    stack: set[str] | None = None,
) -> str:
    if part in graph_cache:
        return graph_cache[part]
    stack = set() if stack is None else set(stack)
    content_hash = semantic_part_hash(part, zf, content_cache)
    if part in stack:
        return hashlib.sha256(f"cycle:{content_hash}".encode()).hexdigest()
    stack.add(part)

    edges: list[dict] = []
    for rel in relationships_for(zf, part, parts, errors).values():
        if rel["external"]:
            edges.append(
                {
                    "type": rel["type"],
                    "external": str(rel["target"]),
                }
            )
            continue
        target = str(rel["target"])
        if target in parts:
            edges.append(
                {
                    "type": rel["type"],
                    "child": recursive_part_hash(
                        target,
                        zf,
                        parts,
                        errors,
                        content_cache,
                        graph_cache,
                        stack,
                    ),
                }
            )
    edges.sort(key=lambda item: json.dumps(item, sort_keys=True))
    encoded = json.dumps(
        {"content": content_hash, "edges": edges},
        sort_keys=True,
        separators=(",", ":"),
    ).encode()
    digest = hashlib.sha256(encoded).hexdigest()
    graph_cache[part] = digest
    return digest


def referenced_payload_signatures(
    element: ET.Element,
    relationships: dict[str, dict[str, str | bool]],
    zf: zipfile.ZipFile,
    parts: set[str],
    errors: list[str],
    content_cache: dict[str, str],
    graph_cache: dict[str, str],
) -> tuple[list[str], list[str]]:
    hashes: set[str] = set()
    external: set[str] = set()
    for node in element.iter():
        for attribute in (RID, REMBED, RLINK):
            rel_id = node.get(attribute)
            if not rel_id:
                continue
            rel = relationships.get(rel_id)
            if rel is None:
                continue
            if rel["external"]:
                external.add(f"{rel['type']}|{rel['target']}")
                continue
            target = str(rel["target"])
            if target in parts:
                hashes.add(
                    recursive_part_hash(
                        target,
                        zf,
                        parts,
                        errors,
                        content_cache,
                        graph_cache,
                    )
                )
    return sorted(hashes), sorted(external)


def visual_object_fingerprint(
    slide_root: ET.Element,
    slide_relationships: dict[str, dict[str, str | bool]],
    zf: zipfile.ZipFile,
    parts: set[str],
    errors: list[str],
    content_cache: dict[str, str],
    graph_cache: dict[str, str],
) -> tuple[str, list[str], list[str]]:
    records: list[dict] = []
    all_payload_hashes: set[str] = set()
    all_external_targets: set[str] = set()

    background = slide_root.find("./p:cSld/p:bg", NS)
    if background is not None:
        records.append(
            {
                "kind": "slideBackground",
                "style": semantic_element(background),
            }
        )

    def visit(container: ET.Element) -> None:
        for element in list(container):
            local_name = element.tag.rsplit("}", 1)[-1]
            if local_name not in {"sp", "pic", "graphicFrame", "cxnSp", "grpSp"}:
                continue
            payload_hashes, external_targets = referenced_payload_signatures(
                element,
                slide_relationships,
                zf,
                parts,
                errors,
                content_cache,
                graph_cache,
            )
            all_payload_hashes.update(payload_hashes)
            all_external_targets.update(external_targets)
            crop = element.find(".//a:srcRect", NS)
            hidden_values = sorted(
                {
                    node.get("hidden", "")
                    for node in element.findall(".//p:cNvPr", NS)
                    if node.get("hidden") is not None
                }
            )
            record = {
                "kind": local_name,
                "transform": transform_signature(element),
                "crop": dict(sorted(crop.attrib.items())) if crop is not None else {},
                "style": style_signature(element),
                "hidden": hidden_values,
                "payload_hashes": payload_hashes,
                "external_targets": external_targets,
            }
            records.append(record)
            if local_name == "grpSp":
                visit(element)

    shape_tree = slide_root.find("./p:cSld/p:spTree", NS)
    if shape_tree is not None:
        visit(shape_tree)
    encoded = json.dumps(
        records, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return (
        hashlib.sha256(encoded).hexdigest(),
        sorted(all_payload_hashes),
        sorted(all_external_targets),
    )


def inspect_package(path: Path) -> dict:
    report: dict = {
        "path": str(path.resolve()),
        "sha256": None,
        "errors": [],
        "warnings": [],
        "slide_size": None,
        "theme_parts": {},
        "all_theme_parts": {},
        "orphan_theme_parts": [],
        "external_relationships": [],
        "risky_parts": [],
        "dangerous_parts": [],
        "embedded_parts": [],
        "slides": [],
        "total_script_chars": 0,
        "notes_relationship_slides": [],
        "notes_nonempty_slides": [],
        "sources_marker_slides": [],
        "sources_nonempty_slides": [],
        "orphan_notes_parts": [],
    }
    errors: list[str] = report["errors"]

    if not path.is_file():
        errors.append("file does not exist")
        return report

    try:
        report["sha256"] = sha256_file(path)
        with zipfile.ZipFile(path) as zf:
            infos = [info for info in zf.infolist() if not info.is_dir()]
            names = [info.filename for info in infos]
            total_uncompressed = sum(info.file_size for info in infos)
            oversized_parts = [
                info.filename for info in infos if info.file_size > MAX_SINGLE_PART_BYTES
            ]
            extreme_ratios = [
                info.filename
                for info in infos
                if info.compress_size > 0
                and info.file_size / info.compress_size > MAX_COMPRESSION_RATIO
            ]
            if len(infos) > MAX_PACKAGE_PARTS:
                errors.append(
                    f"package has too many parts: {len(infos)} > {MAX_PACKAGE_PARTS}"
                )
            if total_uncompressed > MAX_TOTAL_UNCOMPRESSED_BYTES:
                errors.append(
                    "package uncompressed size exceeds the safety limit: "
                    f"{total_uncompressed} > {MAX_TOTAL_UNCOMPRESSED_BYTES}"
                )
            if oversized_parts:
                errors.append(
                    "package parts exceed the single-part safety limit: "
                    + ", ".join(oversized_parts)
                )
            if extreme_ratios:
                errors.append(
                    "package parts exceed the compression-ratio safety limit: "
                    + ", ".join(extreme_ratios)
                )
            if any(
                message.startswith("package ")
                for message in errors
                if "safety limit" in message or "too many parts" in message
            ):
                return report

            parts = set(names)
            content_hash_cache: dict[str, str] = {}
            graph_hash_cache: dict[str, str] = {}
            duplicates = sorted(
                name for name, count in Counter(names).items() if count > 1
            )
            if duplicates:
                errors.append(f"duplicate package parts: {', '.join(duplicates)}")

            dangerous_parts = sorted(
                part
                for part in parts
                if part.endswith("vbaProject.bin")
                or "/activeX/" in part
            )
            embedded_parts = sorted(
                part for part in parts if "/embeddings/" in part
            )
            report["dangerous_parts"] = dangerous_parts
            report["embedded_parts"] = embedded_parts
            report["risky_parts"] = sorted(set(dangerous_parts + embedded_parts))
            if dangerous_parts:
                errors.append(
                    "macro or ActiveX package parts are not allowed in this preflight: "
                    + ", ".join(dangerous_parts)
                )
            if embedded_parts:
                report["warnings"].append(
                    "embedded package parts require review: "
                    + ", ".join(embedded_parts)
                )

            bad_crc = zf.testzip()
            if bad_crc:
                errors.append(f"CRC failure in package part: {bad_crc}")

            report["all_theme_parts"] = {
                part: hashlib.sha256(zf.read(part)).hexdigest()
                for part in sorted(parts)
                if re.fullmatch(r"ppt/(?:[^/]+/)*theme/[^/]+\.xml", part)
            }
            report["theme_parts"] = {
                part: digest
                for part, digest in report["all_theme_parts"].items()
                if re.fullmatch(r"ppt/theme/[^/]+\.xml", part)
            }

            content_types = parse_xml(zf, "[Content_Types].xml", errors)
            if content_types is not None:
                defaults = {
                    node.get("Extension", "").lower(): node.get("ContentType", "")
                    for node in content_types.findall("ct:Default", NS)
                }
                overrides = {
                    unquote(node.get("PartName", "")).lstrip("/")
                    for node in content_types.findall("ct:Override", NS)
                }
                undeclared: list[str] = []
                for part in sorted(parts - {"[Content_Types].xml"}):
                    extension = part.rsplit(".", 1)[1].lower() if "." in part else ""
                    if part not in overrides and extension not in defaults:
                        undeclared.append(part)
                if undeclared:
                    errors.append(
                        "package parts without content-type declarations: "
                        + ", ".join(undeclared)
                    )

            referenced_theme_parts: set[str] = set()
            for rel_part in sorted(name for name in parts if name.endswith(".rels")):
                if rel_part == "_rels/.rels":
                    owner = ""
                else:
                    directory, rel_name = posixpath.split(rel_part)
                    if posixpath.basename(directory) != "_rels" or not rel_name.endswith(".rels"):
                        continue
                    owner = posixpath.join(
                        posixpath.dirname(directory), rel_name[: -len(".rels")]
                    )
                for rel_id, rel in relationships_for(zf, owner, parts, errors).items():
                    if rel["external"]:
                        report["external_relationships"].append(
                            {
                                "owner": owner or "/",
                                "id": rel_id,
                                "type": rel["type"],
                                "target": rel["target"],
                            }
                        )
                        continue
                    target = str(rel["target"])
                    if str(rel["type"]).endswith("/theme"):
                        referenced_theme_parts.add(target)
                    if target.startswith("../") or target not in parts:
                        errors.append(
                            f"broken internal relationship {rel_part}#{rel_id} -> {target}"
                        )

            orphan_themes = sorted(
                set(report["all_theme_parts"]) - referenced_theme_parts
            )
            report["orphan_theme_parts"] = orphan_themes
            if orphan_themes:
                report["warnings"].append(
                    "unreferenced theme parts: " + ", ".join(orphan_themes)
                )
            if report["external_relationships"]:
                report["warnings"].append(
                    f"package contains {len(report['external_relationships'])} external relationships; review linked content"
                )
                non_hyperlink_external = [
                    rel
                    for rel in report["external_relationships"]
                    if not str(rel["type"]).endswith("/hyperlink")
                ]
                if non_hyperlink_external:
                    errors.append(
                        "package contains external non-hyperlink relationships that may break offline"
                    )

            presentation = parse_xml(zf, "ppt/presentation.xml", errors)
            if presentation is None:
                return report

            size = presentation.find("p:sldSz", NS)
            if size is not None:
                report["slide_size"] = [size.get("cx"), size.get("cy")]

            presentation_rels = relationships_for(
                zf, "ppt/presentation.xml", parts, errors
            )
            slide_parts: list[str] = []
            referenced_notes_parts: set[str] = set()
            for slide_id in presentation.findall("./p:sldIdLst/p:sldId", NS):
                rel_id = slide_id.get(RID, "")
                rel = presentation_rels.get(rel_id)
                if rel is None or rel["external"]:
                    errors.append(f"missing slide relationship for {rel_id or '<empty>'}")
                    continue
                slide_parts.append(str(rel["target"]))

            for number, slide_part in enumerate(slide_parts, start=1):
                slide_root = parse_xml(zf, slide_part, errors)
                if slide_root is None:
                    continue
                visible_text = normalized_text(extract_all_text(slide_root))
                shape_counts = {
                    "text_shapes": sum(
                        1
                        for shape in slide_root.findall(".//p:sp", NS)
                        if normalized_text(extract_all_text(shape))
                    ),
                    "pictures": len(slide_root.findall(".//p:pic", NS)),
                    "graphic_frames": len(slide_root.findall(".//p:graphicFrame", NS)),
                    "groups": len(slide_root.findall(".//p:grpSp", NS)),
                }

                slide_rels = relationships_for(zf, slide_part, parts, errors)
                (
                    visual_fingerprint,
                    payload_hashes,
                    external_targets,
                ) = visual_object_fingerprint(
                    slide_root,
                    slide_rels,
                    zf,
                    parts,
                    errors,
                    content_hash_cache,
                    graph_hash_cache,
                )
                layout_targets = [
                    str(rel["target"])
                    for rel in slide_rels.values()
                    if not rel["external"]
                    and str(rel["type"]).endswith("/slideLayout")
                ]
                effective_theme_part = ""
                if len(layout_targets) != 1:
                    errors.append(
                        f"slide {number} must resolve exactly one slide layout, found {len(layout_targets)}"
                    )
                else:
                    layout_rels = relationships_for(
                        zf, layout_targets[0], parts, errors
                    )
                    master_targets = [
                        str(rel["target"])
                        for rel in layout_rels.values()
                        if not rel["external"]
                        and str(rel["type"]).endswith("/slideMaster")
                    ]
                    if len(master_targets) != 1:
                        errors.append(
                            f"slide {number} layout must resolve exactly one slide master, found {len(master_targets)}"
                        )
                    else:
                        master_rels = relationships_for(
                            zf, master_targets[0], parts, errors
                        )
                        theme_targets = [
                            str(rel["target"])
                            for rel in master_rels.values()
                            if not rel["external"]
                            and str(rel["type"]).endswith("/theme")
                        ]
                        if len(theme_targets) != 1:
                            errors.append(
                                f"slide {number} master must resolve exactly one theme, found {len(theme_targets)}"
                            )
                        else:
                            effective_theme_part = theme_targets[0]
                notes_targets = [
                    str(rel["target"])
                    for rel in slide_rels.values()
                    if not rel["external"]
                    and str(rel["type"]).endswith("/notesSlide")
                ]
                if len(notes_targets) > 1:
                    errors.append(f"slide {number} has multiple notes-slide relationships")
                notes_target = notes_targets[0] if notes_targets else ""
                note_text = ""
                if notes_target:
                    report["notes_relationship_slides"].append(number)
                    referenced_notes_parts.add(notes_target)
                    note_root = parse_xml(zf, notes_target, errors)
                    if note_root is not None:
                        note_text = extract_note_text(note_root)
                script, sources, has_sources = split_sources(note_text)
                script_chars = compact_char_count(script)
                long_numbers = sorted(set(LONG_NUMBER_RE.findall(script)))

                if script_chars:
                    report["notes_nonempty_slides"].append(number)
                if has_sources:
                    report["sources_marker_slides"].append(number)
                    if compact_char_count(sources):
                        report["sources_nonempty_slides"].append(number)
                report["total_script_chars"] += script_chars
                report["slides"].append(
                    {
                        "number": number,
                        "part": slide_part,
                        "visible_text": visible_text,
                        "shape_counts": shape_counts,
                        "visual_object_fingerprint": visual_fingerprint,
                        "referenced_payload_hashes": payload_hashes,
                        "external_targets": external_targets,
                        "effective_theme_part": effective_theme_part or None,
                        "effective_theme_sha256": report["all_theme_parts"].get(
                            effective_theme_part
                        ),
                        "notes_part": notes_target or None,
                        "script_chars": script_chars,
                        "has_sources_marker": has_sources,
                        "source_chars": compact_char_count(sources),
                        "long_numbers_in_script": long_numbers,
                    }
                )

            all_notes_parts = {
                part
                for part in parts
                if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", part)
            }
            orphan_notes = sorted(all_notes_parts - referenced_notes_parts)
            report["orphan_notes_parts"] = orphan_notes
            if orphan_notes:
                errors.append("orphan notes-slide parts: " + ", ".join(orphan_notes))

    except zipfile.BadZipFile:
        errors.append("not a valid ZIP/PPTX package")
    except OSError as exc:
        errors.append(f"cannot read file: {exc}")

    return report


def compare_pair(
    submission: dict,
    preparation: dict,
    template: dict | None,
    expected_slides: int | None,
    allow_empty: set[int],
    require_sources: set[int],
    min_script_chars: int | None,
    max_script_chars: int | None,
) -> dict:
    errors: list[str] = []
    warnings: list[str] = []
    sub_slides = submission["slides"]
    prep_slides = preparation["slides"]

    if submission["slide_size"] != preparation["slide_size"]:
        errors.append("slide dimensions differ between variants")
    if len(sub_slides) != len(prep_slides):
        errors.append(
            f"slide counts differ: submission={len(sub_slides)}, preparation={len(prep_slides)}"
        )
    if expected_slides is not None:
        if len(sub_slides) != expected_slides:
            errors.append(
                f"submission slide count differs from expected: {len(sub_slides)} != {expected_slides}"
            )
        if len(prep_slides) != expected_slides:
            errors.append(
                f"preparation slide count differs from expected: {len(prep_slides)} != {expected_slides}"
            )

    for number, (sub_slide, prep_slide) in enumerate(
        zip(sub_slides, prep_slides), start=1
    ):
        if sub_slide["visible_text"] != prep_slide["visible_text"]:
            errors.append(f"visible text differs on slide {number}")
        if sub_slide["shape_counts"] != prep_slide["shape_counts"]:
            errors.append(f"native object counts differ on slide {number}")
        if sub_slide["referenced_payload_hashes"] != prep_slide["referenced_payload_hashes"]:
            errors.append(f"referenced media or chart payloads differ on slide {number}")
        if sub_slide["external_targets"] != prep_slide["external_targets"]:
            errors.append(f"external relationship targets differ on slide {number}")
        if sub_slide["visual_object_fingerprint"] != prep_slide["visual_object_fingerprint"]:
            errors.append(
                f"object geometry, style, z-order, visibility, or payload placement differs on slide {number}"
            )
        if sub_slide["effective_theme_sha256"] != prep_slide["effective_theme_sha256"]:
            errors.append(f"effective visible theme differs on slide {number}")

    if submission["notes_relationship_slides"]:
        errors.append(
            "submission deck contains notes-slide relationships on slides "
            + ", ".join(map(str, submission["notes_relationship_slides"]))
        )

    required_note_slides = set(range(1, len(prep_slides) + 1)) - allow_empty
    missing_note_relationships = sorted(
        required_note_slides - set(preparation["notes_relationship_slides"])
    )
    if missing_note_relationships:
        errors.append(
            "preparation deck lacks notes-slide relationships on slides "
            + ", ".join(map(str, missing_note_relationships))
        )
    missing_notes = sorted(
        required_note_slides - set(preparation["notes_nonempty_slides"])
    )
    if missing_notes:
        errors.append(
            "preparation deck lacks spoken notes on slides "
            + ", ".join(map(str, missing_notes))
        )

    out_of_range = sorted(
        (allow_empty | require_sources) - set(range(1, len(prep_slides) + 1))
    )
    if out_of_range:
        errors.append(
            "slide options reference nonexistent slides "
            + ", ".join(map(str, out_of_range))
        )

    missing_sources = sorted(
        require_sources - set(preparation["sources_nonempty_slides"])
    )
    if missing_sources:
        errors.append(
            "required non-empty [Sources] blocks are absent on preparation slides "
            + ", ".join(map(str, missing_sources))
        )

    empty_source_blocks = sorted(
        set(preparation["sources_marker_slides"])
        - set(preparation["sources_nonempty_slides"])
    )
    if empty_source_blocks:
        warnings.append(
            "empty [Sources] blocks on preparation slides "
            + ", ".join(map(str, empty_source_blocks))
        )

    total_chars = preparation["total_script_chars"]
    if min_script_chars is not None and total_chars < min_script_chars:
        errors.append(
            f"spoken script is below the requested coarse minimum: {total_chars} < {min_script_chars}"
        )
    if max_script_chars is not None and total_chars > max_script_chars:
        errors.append(
            f"spoken script exceeds the requested coarse maximum: {total_chars} > {max_script_chars}"
        )

    if template is not None:
        template_themes = template["theme_parts"]
        allowed_visible_theme_hashes = set(template_themes.values()) | {
            slide["effective_theme_sha256"]
            for slide in template["slides"]
            if slide["effective_theme_sha256"]
        }
        for label, deck in (("submission", submission), ("preparation", preparation)):
            if template["slide_size"] and deck["slide_size"] != template["slide_size"]:
                errors.append(f"{label} deck slide dimensions differ from the template")
            deck_themes = deck["theme_parts"]
            for part, expected_hash in template_themes.items():
                actual_hash = deck_themes.get(part)
                if actual_hash is None:
                    errors.append(f"{label} deck is missing template theme part {part}")
                elif actual_hash != expected_hash:
                    errors.append(f"{label} deck changed template theme part {part}")
            extra_themes = sorted(set(deck_themes) - set(template_themes))
            if extra_themes:
                warnings.append(
                    f"{label} deck contains theme parts absent from the template: "
                    + ", ".join(extra_themes)
                )
            for slide in deck["slides"]:
                effective_hash = slide["effective_theme_sha256"]
                if effective_hash and effective_hash not in allowed_visible_theme_hashes:
                    errors.append(
                        f"{label} slide {slide['number']} uses a visible theme not present in the template"
                    )

    for slide in prep_slides:
        if slide["long_numbers_in_script"]:
            warnings.append(
                f"slide {slide['number']} spoken notes contain long numeric tokens: "
                + ", ".join(slide["long_numbers_in_script"])
            )
        counts = slide["shape_counts"]
        if not slide["visible_text"] and counts["pictures"] and not counts["graphic_frames"]:
            warnings.append(
                f"slide {slide['number']} appears image-led with no native visible text; confirm intended editability"
            )

    return {
        "ok": not (
            errors
            or submission["errors"]
            or preparation["errors"]
            or (template is not None and template["errors"])
        ),
        "errors": errors,
        "warnings": warnings,
        "visible_pair_checks": min(len(sub_slides), len(prep_slides)),
        "preparation_script_chars_excluding_sources": total_chars,
    }


def print_human(result: dict) -> None:
    submission = result["submission"]
    preparation = result["preparation"]
    template = result["template"]
    pair = result["pair"]
    status = "PASS" if result["ok"] else "FAIL"
    print(f"Academic deck pair preflight: {status}")
    print(
        f"Slides: submission={len(submission['slides'])}, preparation={len(preparation['slides'])}"
    )
    print(
        "Preparation spoken characters (sources excluded): "
        f"{preparation['total_script_chars']}"
    )
    print(f"Submission SHA-256: {submission['sha256']}")
    print(f"Preparation SHA-256: {preparation['sha256']}")
    if template is not None:
        print(f"Template SHA-256: {template['sha256']}")

    all_errors = [
        *(f"submission: {item}" for item in submission["errors"]),
        *(f"preparation: {item}" for item in preparation["errors"]),
        *(
            (f"template: {item}" for item in template["errors"])
            if template is not None
            else ()
        ),
        *(f"pair: {item}" for item in pair["errors"]),
    ]
    all_warnings = [
        *(f"submission: {item}" for item in submission["warnings"]),
        *(f"preparation: {item}" for item in preparation["warnings"]),
        *(
            (f"template: {item}" for item in template["warnings"])
            if template is not None
            else ()
        ),
        *(f"pair: {item}" for item in pair["warnings"]),
    ]
    if all_errors:
        print("Errors:")
        for item in all_errors:
            print(f"- {item}")
    if all_warnings:
        print("Warnings:")
        for item in all_warnings:
            print(f"- {item}")
    print("Rendering, semantic source checks, font fidelity, and timed rehearsal remain required.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--submission", required=True, type=Path)
    parser.add_argument("--preparation", required=True, type=Path)
    parser.add_argument("--expected-slides", required=True, type=int)
    parser.add_argument(
        "--allow-empty-prep-slides",
        default="",
        help="Comma-separated slide numbers or ranges allowed to lack spoken notes.",
    )
    parser.add_argument(
        "--require-sources-slides",
        default="",
        help="Comma-separated preparation slides that must contain a [Sources] marker.",
    )
    parser.add_argument("--min-script-chars", type=int)
    parser.add_argument("--max-script-chars", type=int)
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.expected_slides is not None and args.expected_slides < 1:
        print("--expected-slides must be positive", file=sys.stderr)
        return 2
    try:
        allow_empty = parse_slide_spec(args.allow_empty_prep_slides)
        require_sources = parse_slide_spec(args.require_sources_slides)
    except ValueError as exc:
        print(f"invalid slide specification: {exc}", file=sys.stderr)
        return 2

    template = inspect_package(args.template) if args.template else None
    submission = inspect_package(args.submission)
    preparation = inspect_package(args.preparation)
    pair = compare_pair(
        submission,
        preparation,
        template,
        args.expected_slides,
        allow_empty,
        require_sources,
        args.min_script_chars,
        args.max_script_chars,
    )
    result = {
        "ok": pair["ok"],
        "scope": "structural_preflight",
        "manual_checks_required": MANUAL_CHECKS_REQUIRED,
        "template": template,
        "submission": submission,
        "preparation": preparation,
        "pair": pair,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human(result)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
