"""Behavior tests on synthetic OOXML; no copyrighted thesis fixtures."""
import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from xml.dom import minidom
from zipfile import ZipFile

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "docx_format.py"
MODULE = importlib.util.spec_from_file_location("docx_format", SCRIPT)
fmt = importlib.util.module_from_spec(MODULE)
MODULE.loader.exec_module(fmt)
W = fmt.W

DOCUMENT = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" mc:Ignorable="w14"><w:body>
<w:p w14:paraId="ABC"><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>封面保留</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:beforeLines="100"/><w:ind w:firstLine="999" w:hanging="100"/></w:pPr><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:b/><w:i/><w:vertAlign w:val="superscript"/><w:sz w:val="44"/></w:rPr><w:t>数据 5.2% [1]</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF target </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><m:oMath><m:r><m:t>x=2</m:t></m:r></m:oMath><w:bookmarkEnd w:id="1"/></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表内原值 9</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:ins w:id="4"><w:r><w:t>修订文字</w:t></w:r></w:ins></w:p>
<w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="60"/></w:rPr><w:t>嵌套文本保留</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r><w:r><w:t>外层正文</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="10000" w:h="16000"/><w:pgMar w:top="1000" w:left="1200"/></w:sectPr>
</w:body></w:document>'''.encode()
STYLES = f'<w:styles xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>'.encode()
PROFILE = {"id": "synthetic-v1", "roles": {"body": {"font_cn": "宋体", "font_latin": "Times New Roman", "font_pt": 12, "align": "justify", "first_line_chars": 2, "line_pt": 20, "line_rule": "atLeast", "space_before_pt": 0, "space_after_pt": 0}, "heading1": {"font_cn": "黑体", "font_pt": 16, "bold": True, "align": "center", "outline_level": 0, "new_page": True}}, "page": {"width_mm": 210, "height_mm": 297, "top_mm": 35, "bottom_mm": 40, "left_mm": 25, "right_mm": 25, "header_mm": 25, "footer_mm": 30, "gutter_mm": 5}}
MAPPING = {"paragraphs": {"body": [2]}, "sections": [1]}


class FormatTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.src = Path(self.tmp.name) / "source.docx"
        self.out = Path(self.tmp.name) / "new.docx"
        with ZipFile(self.src, "w") as z:
            z.writestr("word/document.xml", DOCUMENT)
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/comments.xml", b"<private-test-comment/>")
            z.writestr("word/_rels/document.xml.rels", b"<Relationships/>")
            z.writestr("word/media/image1.bin", b"\x00\xffopaque")
        self.profile = copy.deepcopy(PROFILE)
        self.mapping = copy.deepcopy(MAPPING)

    def test_apply_preserves_source_content_and_opaque_parts(self):
        before = self.src.read_bytes()
        result = fmt.process(self.src, self.profile, self.mapping, self.out)
        self.assertEqual(before, self.src.read_bytes())
        self.assertTrue(result["source_unchanged"])
        self.assertTrue(result["untouched_parts_preserved"])
        self.assertEqual(result["changed_parts"], ["word/document.xml"])
        self.assertEqual(result["status"], "partial")
        _, _, _, old, _ = fmt.load_package(self.src)
        _, _, _, new, secs = fmt.load_package(self.out)
        self.assertEqual([fmt.text_content(p) for p in old], [fmt.text_content(p) for p in new])
        for i in (0, 2, 3, 4, 5):
            self.assertEqual(old[i].toxml(), new[i].toxml())
        for tag in ("bookmarkStart", "bookmarkEnd", "instrText", "fldChar"):
            self.assertEqual([n.toxml() for n in old[1].getElementsByTagNameNS(W, tag)], [n.toxml() for n in new[1].getElementsByTagNameNS(W, tag)])
        self.assertEqual(new[1].getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/math", "oMath")[0].toxml(), old[1].getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/math", "oMath")[0].toxml())
        for tag in ("b", "i", "vertAlign"):
            self.assertEqual(old[1].getElementsByTagNameNS(W, tag)[0].toxml(), new[1].getElementsByTagNameNS(W, tag)[0].toxml())
        self.assertEqual(fmt.first(secs[0], "pgMar").getAttributeNS(W, "top"), str(round(35 * 1440 / 25.4)))
        self.assertEqual(fmt.first(secs[0], "pgMar").getAttributeNS(W, "gutter"), str(round(5 * 1440 / 25.4)))

    def test_apply_then_audit_is_idempotent_and_checks_direct_formats(self):
        fmt.process(self.src, self.profile, self.mapping, self.out)
        self.assertTrue(fmt.process(self.out, self.profile, self.mapping)["mapped_checks_passed"])
        with ZipFile(self.out) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        parts["word/document.xml"] = parts["word/document.xml"].replace(b'w:sz w:val="24"', b'w:sz w:val="30"', 1)
        tampered = Path(self.tmp.name) / "tampered.docx"
        with ZipFile(tampered, "w") as z:
            for n, data in parts.items():
                z.writestr(n, data)
        self.assertFalse(fmt.process(tampered, self.profile, self.mapping)["mapped_checks_passed"])

    def test_heading_bold_overrides_direct_false_but_body_keeps_emphasis(self):
        self.mapping = {"paragraphs": {"heading1": [2]}}
        # Set an explicit non-bold run, which would override a bold style.
        with ZipFile(self.src) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        parts["word/document.xml"] = parts["word/document.xml"].replace(b"<w:b/>", b'<w:b w:val="0"/>')
        with ZipFile(self.src, "w") as z:
            for n, data in parts.items():
                z.writestr(n, data)
        fmt.process(self.src, self.profile, self.mapping, self.out)
        p = fmt.load_package(self.out)[3][1]
        self.assertEqual(p.getElementsByTagNameNS(W, "b")[0].getAttributeNS(W, "val"), "1")
        self.assertTrue(fmt.process(self.out, self.profile, self.mapping)["mapped_checks_passed"])

    def test_namespace_prefixes_named_only_in_ignorable_survive(self):
        fmt.process(self.src, self.profile, self.mapping, self.out)
        with ZipFile(self.out) as z:
            for name in ("word/document.xml", "word/styles.xml"):
                root = minidom.parseString(z.read(name)).documentElement
                self.assertTrue(root.hasAttribute("xmlns:w14"))
                self.assertEqual(root.getAttribute("mc:Ignorable"), "w14")

    def test_inventory_is_private_by_default_and_counts_table_textbox(self):
        data = fmt.inspect(self.src)
        self.assertEqual(len(data["paragraphs"]), 6)
        self.assertFalse(any("text" in p for p in data["paragraphs"]))
        self.assertEqual(fmt.inspect(self.src, True)["paragraphs"][2]["text"], "表内原值 9")

    def test_outer_paragraph_mapping_does_not_reformat_nested_textbox(self):
        self.mapping = {"paragraphs": {"body": [5]}}
        fmt.process(self.src, self.profile, self.mapping, self.out)
        old = fmt.load_package(self.src)[3][5]
        new = fmt.load_package(self.out)[3][5]
        self.assertEqual(old.toxml(), new.toxml())

    def test_refuses_overlapping_or_out_of_range_indexes(self):
        for mapping in ({"paragraphs": {"body": [2], "heading1": [2]}}, {"paragraphs": {"body": [0]}}, {"paragraphs": {"body": [7]}}, {"paragraphs": {"body": [True]}}, {"sections": [2]}, {"sections": [1, 1]}, {}):
            with self.subTest(mapping=mapping), self.assertRaises(ValueError):
                fmt.process(self.src, self.profile, mapping, self.out)
        self.assertFalse(self.out.exists())

    def test_refuses_overwrite_bad_hash_and_tracked_changes(self):
        with self.assertRaises(ValueError):
            fmt.process(self.src, self.profile, self.mapping, self.src)
        self.out.write_bytes(b"existing")
        with self.assertRaises(ValueError):
            fmt.process(self.src, self.profile, self.mapping, self.out)
        self.assertEqual(self.out.read_bytes(), b"existing")
        with self.assertRaises(ValueError):
            fmt.process(self.src, self.profile, {"paragraphs": {"body": [4]}})
        with self.assertRaises(ValueError):
            fmt.process(self.src, self.profile, {**self.mapping, "source_sha256": "wrong"})

    def test_multiple_and_spacing_lines_use_separate_units(self):
        spec = self.profile["roles"]["body"]
        spec.pop("line_pt")
        spec.pop("space_before_pt")
        spec.update(line_multiple=1.25, line_rule="auto", space_before_lines=1.5)
        fmt.process(self.src, self.profile, self.mapping, self.out)
        p = fmt.load_package(self.out)[3][1]
        spacing = fmt.first(fmt.first(p, "pPr"), "spacing")
        self.assertEqual(spacing.getAttributeNS(W, "line"), "300")
        self.assertEqual(spacing.getAttributeNS(W, "lineRule"), "auto")
        self.assertEqual(spacing.getAttributeNS(W, "beforeLines"), "150")
        self.assertFalse(spacing.hasAttributeNS(W, "before"))
        self.assertTrue(fmt.process(self.out, self.profile, self.mapping)["mapped_checks_passed"])

    def test_profile_changes_leave_previously_mapped_paragraph_unchanged(self):
        fmt.process(self.src, self.profile, self.mapping, self.out)
        self.profile["roles"]["body"]["font_pt"] = 13
        next_out = Path(self.tmp.name) / "next.docx"
        fmt.process(self.out, self.profile, {"paragraphs": {"body": [3]}}, next_out)
        old = fmt.load_package(self.out)[3][1]
        new = fmt.load_package(next_out)[3][1]
        self.assertEqual(old.toxml(), new.toxml())

    def test_original_heading_style_numbering_and_style_fields_are_preserved(self):
        with ZipFile(self.src) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        heading_style = b'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
        parts["word/styles.xml"] = parts["word/styles.xml"].replace(b"</w:styles>", heading_style + b"</w:styles>")
        parts["word/document.xml"] = parts["word/document.xml"].replace(b'<w:p><w:pPr><w:spacing', b'<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:spacing')
        style_fields = rb'<w:r><w:instrText xml:space="preserve"> STYLEREF Heading2 </w:instrText></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r>'
        parts["word/document.xml"] = parts["word/document.xml"].replace(b'<w:bookmarkEnd w:id="1"/>', style_fields + b'<w:bookmarkEnd w:id="1"/>')
        # A real STYLEREF may also live in an opaque header part.
        parts["word/header1.xml"] = ('<w:hdr xmlns:w="' + W + '"><w:p><w:r><w:instrText> STYLEREF Heading2 </w:instrText></w:r></w:p></w:hdr>').encode()
        with ZipFile(self.src, "w") as z:
            for name, data in parts.items():
                z.writestr(name, data)
        old_p = fmt.load_package(self.src)[3][1]
        result = fmt.process(self.src, self.profile, self.mapping, self.out)
        new_p = fmt.load_package(self.out)[3][1]
        self.assertEqual(fmt.first(fmt.first(new_p, "pPr"), "pStyle").getAttributeNS(W, "val"), "Heading2")
        with ZipFile(self.out) as z:
            self.assertEqual(z.read("word/styles.xml"), parts["word/styles.xml"])
            self.assertEqual(z.read("word/header1.xml"), parts["word/header1.xml"])
            styles = minidom.parseString(z.read("word/styles.xml"))
            heading = [n for n in styles.getElementsByTagNameNS(W, "style") if n.getAttributeNS(W, "styleId") == "Heading2"][0]
            self.assertEqual(heading.getElementsByTagNameNS(W, "numId")[0].getAttributeNS(W, "val"), "7")
            self.assertTrue(heading.getElementsByTagNameNS(W, "keepNext"))
        self.assertEqual([n.toxml() for n in old_p.getElementsByTagNameNS(W, "instrText")], [n.toxml() for n in new_p.getElementsByTagNameNS(W, "instrText")])
        self.assertEqual(result["changed_parts"], ["word/document.xml"])
        self.assertTrue(fmt.process(self.out, self.profile, self.mapping)["mapped_checks_passed"])

    def test_cli_exit_codes_and_no_file_edit_on_audit(self):
        profile = Path(self.tmp.name) / "profile.json"
        mapping = Path(self.tmp.name) / "map.json"
        profile.write_text(json.dumps(self.profile))
        mapping.write_text(json.dumps(self.mapping))
        args = [sys.executable, str(SCRIPT), "audit", str(self.src), "--profile", str(profile), "--map", str(mapping)]
        result = subprocess.run(args, capture_output=True, text=True)
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertFalse(json.loads(result.stdout)["mapped_checks_passed"])
        self.assertFalse(self.out.exists())
        args[3] = str(Path(self.tmp.name) / "absent.docx")
        self.assertEqual(subprocess.run(args, capture_output=True).returncode, 2)


if __name__ == "__main__":
    unittest.main()
