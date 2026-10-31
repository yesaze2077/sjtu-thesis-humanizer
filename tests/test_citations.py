"""Synthetic citation integrity cases; no real thesis or fabricated bibliography."""
import importlib.util
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

SPEC = importlib.util.spec_from_file_location('citations', Path(__file__).resolve().parents[1] / 'scripts/check_citations.py')
MOD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MOD)


class CitationTests(unittest.TestCase):
    def check(self, text):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'test.md'
            path.write_text(text, encoding='utf-8')
            before = path.read_bytes()
            result = MOD.audit(path)
            self.assertEqual(path.read_bytes(), before)
            return result

    def test_clean_ranges_repetition_and_code(self):
        r = self.check('# 正文\n甲[1]。乙[2-3]。再见[1]。`code[99]`\n```\n[88]\n```\n# 参考文献\n[1] 合成条目甲\n[2] 合成条目乙\n[3] 合成条目丙\n# 致谢\n[77] 非文献\n')
        self.assertEqual(r['issues'], [])
        self.assertEqual(r['first_appearance'], [1, 2, 3])
        self.assertEqual(r['status'], 'partial')

    def test_missing_unused_duplicates_and_order(self):
        r = self.check('乙[2]，甲[1]，丙[5]。\n# 参考文献\n[1] 甲\n[2] 乙\n[2] 重复乙\n[3] 未引用丙\n')
        kinds = {x['kind'] for x in r['issues']}
        self.assertTrue({'duplicate_reference_numbers', 'citation_without_reference_candidate', 'uncited_reference_candidate', 'first_appearance_order_candidate', 'reference_list_sequence'} <= kinds)

    def test_fullwidth_and_heading_candidate(self):
        r = self.check('# 方法［１，２］\n# 参考文献\n［１］ 条目甲\n［２］ 条目乙\n')
        self.assertEqual(r['first_appearance'], [1, 2])
        self.assertEqual(r['issues'][0]['kind'], 'citation_in_heading_candidate')

    def test_empty_list_and_missing_heading_refused(self):
        for text in ['无边界[1]', '# 参考文献\n自动编号未展开']:
            with self.assertRaises(ValueError):
                self.check(text)

    def test_duplicate_doi_and_bounded_ranges(self):
        r = self.check('甲[1]乙[2]范围[2-5000]\n# 参考文献\n[1] 甲 10.1234/same\n[2] 乙 https://doi.org/10.1234/SAME.\n')
        self.assertTrue({'duplicate_DOI_candidate', 'ambiguous_numeric_bracket'} <= {i['kind'] for i in r['issues']})

    def test_wrapped_title_does_not_silently_end_reference_list(self):
        r = self.check('正文引用[1]。\n参考文献\n[1] 合成作者.\n学术成果评价方法[M].合成信息。\n[2] 合成作者乙.合成条目乙。\n')
        self.assertEqual(r['reference_entries'], 2)
        self.assertEqual(r['references_end_boundary'], None)
        self.assertIn({'kind': 'uncited_reference_candidate', 'numbers': [2]}, r['issues'])

    def test_plain_end_heading_requires_boundary_confirmation(self):
        r = self.check('正文引用[1]。\n参考文献\n[1] 合成条目。\n致谢\n感谢。\n')
        self.assertEqual(r['references_end_boundary']['method'], 'plain_heading_candidate')
        self.assertIn('references_end_boundary_candidate', {i['kind'] for i in r['issues']})

    def test_docx_inline_run_split_is_read_only(self):
        ns = MOD.W[1:-1]
        xml = f'<w:document xmlns:w="{ns}"><w:body><w:p><w:r><w:t>主张[</w:t></w:r><w:r><w:t>1]</w:t></w:r></w:p><w:p><w:r><w:t>参考文献</w:t></w:r></w:p><w:p><w:r><w:t>[1] 合成条目</w:t></w:r></w:p></w:body></w:document>'
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / 'test.docx'
            with ZipFile(p, 'w') as z:
                z.writestr('word/document.xml', xml)
            before = p.read_bytes()
            r = MOD.audit(p)
            self.assertEqual(r['issues'], [])
            self.assertEqual(before, p.read_bytes())

    def test_docx_inherited_heading_and_nested_text_not_duplicated(self):
        ns = MOD.W[1:-1]
        xml = f'<w:document xmlns:w="{ns}"><w:body><w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>标题[1]</w:t></w:r><w:txbxContent><w:p><w:r><w:t>框内[2]</w:t></w:r></w:p></w:txbxContent></w:p><w:p><w:r><w:t>参考文献</w:t></w:r></w:p><w:p><w:r><w:t>[1] 合成甲</w:t></w:r></w:p><w:p><w:r><w:t>[2] 合成乙</w:t></w:r></w:p></w:body></w:document>'
        styles = f'<w:styles xmlns:w="{ns}"><w:style w:styleId="Base"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:styleId="Derived"><w:basedOn w:val="Base"/></w:style></w:styles>'
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / 'test.docx'
            with ZipFile(p, 'w') as z:
                z.writestr('word/document.xml', xml)
                z.writestr('word/styles.xml', styles)
            r = MOD.audit(p)
            self.assertEqual(r['citation_marker_candidates'], 2)
            self.assertEqual([i['kind'] for i in r['issues']], ['citation_in_heading_candidate'])


if __name__ == '__main__':
    unittest.main()
