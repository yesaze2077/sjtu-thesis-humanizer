#!/usr/bin/env python3
"""Apply mapped DOCX direct formatting while preserving semantic styles.

Python standard library only. Existing XML namespace declarations and all
untouched package parts are preserved. No prose, fields or tables are rebuilt.
"""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from xml.dom import minidom
from zipfile import BadZipFile, ZipFile
from xml.parsers.expat import ExpatError

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/2000/xmlns/"
NOT_CHECKED = ["unmapped_paragraphs_and_sections", "headers_and_footers", "settings_mirror_margins_and_gutter", "TOC_and_cross_reference_refresh", "caption_position_and_numbering", "table_borders_and_continuations", "fonts_installed", "pagination_and_visual_render", "reference_content_and_standard", "institutional_acceptance"]
ROLE_KEYS = {"font_cn", "font_latin", "font_pt", "align", "first_line_chars", "line_pt", "line_multiple", "line_rule", "space_before_pt", "space_after_pt", "space_before_lines", "space_after_lines", "bold", "outline_level", "new_page"}
ORDER = {
    "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange".split(),
    "rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange".split(),
    "sectPr": "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange".split(),
}
PAGE_KEYS = {"width_mm", "height_mm", "top_mm", "bottom_mm", "left_mm", "right_mm", "header_mm", "footer_mm", "gutter_mm"}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def children(node, local):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE and n.namespaceURI == W and n.localName == local]


def first(node, local):
    items = children(node, local)
    return items[0] if items else None


def element(doc, local, **attrs):
    node = doc.createElementNS(W, "w:" + local)
    for key, val in attrs.items():
        node.setAttributeNS(W, "w:" + key, str(val))
    return node


def ensure(node, local, first_child=False):
    found = first(node, local)
    if found is None:
        found = element(node.ownerDocument, local)
        if first_child and node.firstChild:
            node.insertBefore(found, node.firstChild)
        else:
            insert_ordered(node, found)
    return found


def set_attrs(node, attrs):
    for key, val in attrs.items():
        node.setAttributeNS(W, "w:" + key, str(val))


def remove_attrs(node, names):
    for key in names:
        if node.hasAttributeNS(W, key):
            node.removeAttributeNS(W, key)


def insert_ordered(parent, new):
    order = ORDER.get(parent.localName, [])
    rank = order.index(new.localName) if new.localName in order else len(order)
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.namespaceURI == W and child.localName in order and order.index(child.localName) > rank:
            parent.insertBefore(new, child)
            return
    parent.appendChild(new)


def replace_child(parent, new):
    old = children(parent, new.localName)
    if old:
        parent.replaceChild(new, old[0])
        for extra in old[1:]:
            parent.removeChild(extra)
    else:
        insert_ordered(parent, new)


def load_package(path):
    with ZipFile(path) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError("duplicate ZIP member names")
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    if "word/document.xml" not in parts or "word/styles.xml" not in parts:
        raise ValueError("DOCX must contain document.xml and styles.xml")
    roots = [minidom.parseString(parts[name]) for name in ("word/document.xml", "word/styles.xml")]
    for doc in roots:
        # Retain unknown prefixes, including those mentioned only in mc:Ignorable.
        root = doc.documentElement
        if not root.hasAttribute("xmlns:w"):
            root.setAttributeNS(XMLNS, "xmlns:w", W)
    body = roots[0].getElementsByTagNameNS(W, "body")
    if len(body) != 1:
        raise ValueError("expected one document body")
    return infos, parts, roots, list(body[0].getElementsByTagNameNS(W, "p")), list(body[0].getElementsByTagNameNS(W, "sectPr"))


def text_content(p):
    return "".join(n.firstChild.data if n.firstChild else "" for n in p.getElementsByTagNameNS(W, "t"))


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def validate(profile, mapping, ps, sections, source_hash):
    if not isinstance(profile.get("roles"), dict) or not isinstance(profile.get("page", {}), dict):
        raise ValueError("profile requires roles object and optional page object")
    if set(profile.get("page", {})) - PAGE_KEYS:
        raise ValueError("unsupported page property")
    if set(mapping) - {"paragraphs", "sections", "source_sha256"}:
        raise ValueError("unsupported map property")
    if mapping.get("source_sha256", source_hash) != source_hash:
        raise ValueError("map source_sha256 does not match input")
    if not isinstance(mapping.get("paragraphs", {}), dict):
        raise ValueError("map paragraphs must be an object")
    seen = set()
    for role, indexes in mapping.get("paragraphs", {}).items():
        if not re.fullmatch(r"[a-z][a-z0-9_-]{0,39}", role) or role not in profile["roles"]:
            raise ValueError("unknown or invalid role: " + str(role))
        spec = profile["roles"][role]
        if not isinstance(spec, dict) or set(spec) - ROLE_KEYS:
            raise ValueError("unsupported role property: " + role)
        if not isinstance(indexes, list):
            raise ValueError("paragraph indexes must be a list")
        for idx in indexes:
            if type(idx) is not int or not 1 <= idx <= len(ps) or idx in seen:
                raise ValueError("duplicate, conflicting or out-of-range paragraph index")
            seen.add(idx)
        if spec.get("align", "left") not in {"left", "right", "center", "justify"}:
            raise ValueError("invalid alignment")
        if "line_pt" in spec and "line_multiple" in spec:
            raise ValueError("choose line_pt or line_multiple")
        if spec.get("line_rule") not in {None, "atLeast", "exact", "auto"}:
            raise ValueError("invalid line_rule")
        if "line_multiple" in spec and spec.get("line_rule", "auto") != "auto":
            raise ValueError("line_multiple requires auto line_rule")
        if spec.get("line_rule") == "auto" and "line_multiple" not in spec:
            raise ValueError("auto line_rule requires line_multiple")
        if "outline_level" in spec and (type(spec["outline_level"]) is not int or not 0 <= spec["outline_level"] <= 9):
            raise ValueError("invalid outline_level")
        if any(a in spec and b in spec for a, b in (("space_before_pt", "space_before_lines"), ("space_after_pt", "space_after_lines"))):
            raise ValueError("choose points or lines for paragraph spacing")
        for key in ("bold", "new_page"):
            if key in spec and type(spec[key]) is not bool:
                raise ValueError(key + " must be boolean")
        for key in {"font_pt", "line_pt", "line_multiple", "first_line_chars", "space_before_pt", "space_after_pt", "space_before_lines", "space_after_lines"} & set(spec):
            value = spec[key]
            if type(value) not in (int, float) or not 0 <= value < 1000 or (key in {"font_pt", "line_pt", "line_multiple"} and value == 0):
                raise ValueError("invalid numeric property: " + key)
    si = mapping.get("sections", [])
    if not isinstance(si, list) or any(type(i) is not int or not 1 <= i <= len(sections) for i in si) or len(si) != len(set(si)):
        raise ValueError("duplicate or out-of-range section index")
    if si and not profile.get("page"):
        raise ValueError("selected sections require page properties")
    for key, value in profile.get("page", {}).items():
        if type(value) not in (int, float) or not 0 <= value < 2000 or (key in {"width_mm", "height_mm"} and value == 0):
            raise ValueError("invalid page property: " + key)
    if not seen and not si:
        raise ValueError("map must select at least one paragraph or section")
    return seen


def paragraph_properties(doc, spec):
    ppr = element(doc, "pPr")
    if "align" in spec:
        replace_child(ppr, element(doc, "jc", val={"justify": "both"}.get(spec["align"], spec["align"])))
    if "first_line_chars" in spec:
        replace_child(ppr, element(doc, "ind", firstLineChars=round(spec["first_line_chars"] * 100)))
    spacing = {}
    for key, xmlkey in (("space_before_pt", "before"), ("space_after_pt", "after")):
        if key in spec:
            spacing[xmlkey] = round(spec[key] * 20)
    for key, xmlkey in (("space_before_lines", "beforeLines"), ("space_after_lines", "afterLines")):
        if key in spec:
            spacing[xmlkey] = round(spec[key] * 100)
    if "line_multiple" in spec:
        spacing.update(line=round(spec["line_multiple"] * 240), lineRule="auto")
    elif "line_pt" in spec:
        spacing.update(line=round(spec["line_pt"] * 20), lineRule=spec.get("line_rule", "atLeast"))
    if spacing:
        replace_child(ppr, element(doc, "spacing", **spacing))
    if "outline_level" in spec:
        replace_child(ppr, element(doc, "outlineLvl", val=spec["outline_level"]))
    if "new_page" in spec:
        replace_child(ppr, element(doc, "pageBreakBefore", val=int(spec["new_page"])))
    return ppr


def format_run(rpr, spec):
    if "font_cn" in spec or "font_latin" in spec:
        fonts = ensure(rpr, "rFonts")
        if "font_cn" in spec:
            remove_attrs(fonts, ["eastAsiaTheme"])
            set_attrs(fonts, {"eastAsia": spec["font_cn"]})
        if "font_latin" in spec:
            remove_attrs(fonts, ["asciiTheme", "hAnsiTheme"])
            set_attrs(fonts, {"ascii": spec["font_latin"], "hAnsi": spec["font_latin"]})
    if "font_pt" in spec:
        replace_child(rpr, element(rpr.ownerDocument, "sz", val=round(spec["font_pt"] * 2)))
        replace_child(rpr, element(rpr.ownerDocument, "szCs", val=round(spec["font_pt"] * 2)))


def has_revision(node):
    cur = node
    while cur and cur.nodeType == cur.ELEMENT_NODE:
        if cur.namespaceURI == W and cur.localName in {"ins", "del", "moveFrom", "moveTo"}:
            return True
        cur = cur.parentNode
    return any(n.localName in {"ins", "del", "moveFrom", "moveTo"} or n.localName.endswith("Change") for n in node.getElementsByTagNameNS(W, "*"))


def apply_paragraph(p, role, spec):
    ppr = ensure(p, "pPr", first_child=True)
    # Retain the original semantic style and its inherited numbering/keep rules.
    # Existing TOC and STYLEREF fields continue to refer to the same style.
    expected = paragraph_properties(p.ownerDocument, spec)
    for item in expected.childNodes:
        # Copy only requested attributes. Other paragraph properties stay intact.
        target = ensure(ppr, item.localName)
        if item.localName == "ind":
            remove_attrs(target, ["firstLine", "hanging", "hangingChars"])
        if item.localName == "spacing":
            remove_attrs(target, ["beforeAutospacing", "afterAutospacing"])
            remove_attrs(target, ["beforeLines"] if "space_before_pt" in spec else ["before"] if "space_before_lines" in spec else [])
            remove_attrs(target, ["afterLines"] if "space_after_pt" in spec else ["after"] if "space_after_lines" in spec else [])
        for attr in item.attributes.values():
            target.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
    # Direct text runs only: do not change math, symbols, pictures or field code.
    for run in p.getElementsByTagNameNS(W, "r"):
        ancestor = run.parentNode
        while ancestor is not None and not (ancestor.namespaceURI == W and ancestor.localName == "p"):
            ancestor = ancestor.parentNode
        if ancestor is not p:
            continue
        if children(run, "t") and not any(children(run, k) for k in ("drawing", "pict", "object")):
            rpr = ensure(run, "rPr", first_child=True)
            format_run(rpr, spec)
            # A requested bold heading must override explicit non-bold text.
            # Otherwise retain direct emphasis, italics and super/subscripts.
            if spec.get("bold") is True:
                replace_child(rpr, element(p.ownerDocument, "b", val=1))
                replace_child(rpr, element(p.ownerDocument, "bCs", val=1))
    mark_rpr = first(ppr, "rPr")
    if mark_rpr is not None:
        format_run(mark_rpr, spec)


def apply_section(section, page):
    marg = {k[:-3]: round(v * 1440 / 25.4) for k, v in page.items() if k not in {"width_mm", "height_mm"}}
    if marg:
        set_attrs(ensure(section, "pgMar"), marg)
    size = {"w" if k == "width_mm" else "h": round(page[k] * 1440 / 25.4) for k in ("width_mm", "height_mm") if k in page}
    if size:
        if "width_mm" in page and "height_mm" in page:
            size["orient"] = "landscape" if page["width_mm"] > page["height_mm"] else "portrait"
        set_attrs(ensure(section, "pgSz"), size)


def inspect(path, include_text=False):
    _, parts, _, ps, sections = load_package(path)
    result = {"source_sha256": sha(Path(path).read_bytes()), "paragraph_index": "1-based body descendant w:p including table/textbox paragraphs", "paragraphs": [], "sections": [], "not_checked": NOT_CHECKED}
    for i, p in enumerate(ps, 1):
        st = first(first(p, "pPr"), "pStyle") if first(p, "pPr") is not None else None
        item = {"index": i, "style_id": st.getAttributeNS(W, "val") if st else None, "has_revision": has_revision(p)}
        if include_text:
            item["text"] = text_content(p)
        result["paragraphs"].append(item)
    for i, sec in enumerate(sections, 1):
        item = {"index": i}
        for tag in ("pgSz", "pgMar", "pgNumType"):
            n = first(sec, tag)
            if n is not None:
                item[tag] = {a.localName: a.value for a in n.attributes.values()}
        result["sections"].append(item)
    return result


def process(path, profile, mapping, out=None):
    source = Path(path).read_bytes()
    infos, parts, docs, ps, sections = load_package(path)
    seen = validate(profile, mapping, ps, sections, sha(source))
    doc = docs[0]
    original_doc = doc.toxml()
    for role, indexes in mapping.get("paragraphs", {}).items():
        spec = profile["roles"][role]
        for idx in indexes:
            if has_revision(ps[idx - 1]):
                raise ValueError("selected paragraph contains tracked changes; resolve explicitly first")
            apply_paragraph(ps[idx - 1], role, spec)
    for idx in mapping.get("sections", []):
        if has_revision(sections[idx - 1]):
            raise ValueError("selected section contains tracked changes")
        apply_section(sections[idx - 1], profile["page"])
    changed = []
    if doc.toxml() != original_doc:
        changed.append("word/document.xml")
        parts["word/document.xml"] = doc.toxml(encoding="utf-8")
    result = {"status": "partial", "profile_id": profile.get("id"), "source_sha256": sha(source), "mapped_paragraph_count": len(seen), "mapped_section_count": len(mapping.get("sections", [])), "changed_parts": changed, "not_checked": NOT_CHECKED}
    if out is None:
        result["mapped_checks_passed"] = not changed
        return result
    dest = Path(out)
    if dest.resolve() == Path(path).resolve() or dest.exists():
        raise ValueError("output must be a new file, different from input")
    # Exclusive open prevents accidental replacement, including a racing creator.
    with dest.open("xb") as handle:
        with ZipFile(handle, "w") as z:
            for info in infos:
                z.writestr(info, parts[info.filename])
    with ZipFile(path) as before, ZipFile(dest) as after:
        result["untouched_parts_preserved"] = all(before.read(i.filename) == after.read(i.filename) for i in infos if i.filename not in changed)
    result["output_sha256"] = sha(dest.read_bytes())
    result["source_unchanged"] = sha(Path(path).read_bytes()) == sha(source)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    inventory = sub.add_parser("inventory")
    inventory.add_argument("input")
    inventory.add_argument("--include-text", action="store_true", help="include private thesis text in local output")
    for command in ("apply", "audit"):
        p = sub.add_parser(command)
        p.add_argument("input")
        p.add_argument("--profile", required=True)
        p.add_argument("--map", required=True)
        if command == "apply":
            p.add_argument("--out", required=True)
    args = parser.parse_args()
    try:
        if args.command == "inventory":
            result = inspect(args.input, args.include_text)
        else:
            result = process(args.input, read_json(args.profile), read_json(args.map), getattr(args, "out", None))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1 if args.command == "audit" and not result["mapped_checks_passed"] else 0
    except (ValueError, OSError, KeyError, TypeError, BadZipFile, ExpatError) as exc:
        print(json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
