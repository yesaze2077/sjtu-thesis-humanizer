#!/usr/bin/env python3
"""Read-only numeric citation cross-check; never verifies source authenticity."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MARKER = re.compile(r'[\[［]([0-9０-９]+(?:\s*[,，、\-–—－]\s*[0-9０-９]+)*)[\]］]')
ENTRY = re.compile(r'^\s*[\[［]([0-9０-９]+)[\]］]\s*(.*)$')
REF_HEADING = re.compile(r'^\s*(?:#{1,6}\s*)?参考文献\s*$')
END_HEADING = re.compile(r'^\s*(?:#{1,6}\s*)?(?:附录|致\s*谢|攻读|学术成果)')
PLAIN_END = re.compile(r'^(?:致\s*谢|附录(?:\s*[A-ZＡ-Ｚ0-9０-９一二三四五六七八九十])?|学术成果(?:目录)?|攻读(?:硕士|博士)?学位期间(?:发表的?)?(?:学术论文|研究成果)(?:目录)?)$')
UNVERIFIED = ['source_existence', 'fulltext_reading', 'claim_support',
              'quotation_fidelity', 'GB_T_7714_full_conformance',
              'Word_field_refresh', 'superscript_rendering', 'institutional_acceptance']


def plain_number(text):
    return text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))


def expand(text):
    """Bound ranges so a malformed numeric bracket cannot explode memory."""
    numbers = []
    for part in re.split(r'[,，、]', plain_number(text)):
        pair = re.split(r'[-–—－]', part.strip())
        if len(pair) == 1:
            numbers.append(int(pair[0]))
        elif len(pair) == 2:
            lo, hi = map(int, pair)
            if hi < lo or hi - lo > 1000:
                raise ValueError('invalid or excessively large citation range')
            numbers.extend(range(lo, hi + 1))
        else:
            raise ValueError('invalid citation range')
    if not numbers or min(numbers) < 1:
        raise ValueError('citation numbers must be positive')
    return numbers


def read_blocks(path):
    if path.suffix.lower() == '.docx':
        with ZipFile(path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
            body = root.find(W + 'body')
            if body is None:
                raise ValueError('document body missing')
            styles = {}
            if 'word/styles.xml' in z.namelist():
                st = ET.fromstring(z.read('word/styles.xml'))
                styles = {s.get(W + 'styleId'): s for s in st.findall(W + 'style')}
            def outline(ppr, visited=None):
                if ppr is None:
                    return None
                direct = ppr.find(W + 'outlineLvl')
                if direct is not None:
                    return direct.get(W + 'val')
                ref = ppr.find(W + 'pStyle')
                sid = ref.get(W + 'val') if ref is not None else None
                visited = set() if visited is None else visited
                while sid and sid not in visited and sid in styles:
                    visited.add(sid)
                    style = styles[sid]
                    level = style.find(W + 'pPr/' + W + 'outlineLvl')
                    if level is not None:
                        return level.get(W + 'val')
                    base = style.find(W + 'basedOn')
                    sid = base.get(W + 'val') if base is not None else None
                return None
            parents = {child: parent for parent in body.iter() for child in parent}
            def own_text(p):
                values = []
                for node in p.iter(W + 't'):
                    parent = parents.get(node)
                    while parent is not None and parent.tag != W + 'p':
                        parent = parents.get(parent)
                    if parent is p:
                        values.append(node.text or '')
                return ''.join(values)
            blocks = []
            for index, p in enumerate(body.iter(W + 'p'), 1):
                text = own_text(p)
                ppr = p.find(W + 'pPr')
                level = outline(ppr)
                heading = level is not None and level != '9'
                blocks.append({'location': f'body_p:{index}', 'text': text, 'heading': heading})
            return blocks, next(root.iter(W + 'ins'), None) is not None or next(root.iter(W + 'del'), None) is not None
    if path.suffix.lower() not in {'.md', '.txt'}:
        raise ValueError('supported inputs: .md, .txt, .docx')
    blocks, fence = [], False
    for index, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
        if re.match(r'^\s*(```|~~~)', line):
            fence = not fence
            continue
        if fence:
            continue
        line = re.sub(r'`[^`]*`', '', line)
        blocks.append({'location': f'line:{index}', 'text': line,
                       'heading': bool(re.match(r'^\s*#{1,6}\s', line))})
    return blocks, False


def audit(path):
    blocks, revisions = read_blocks(path)
    headings = [i for i, b in enumerate(blocks) if REF_HEADING.match(b['text'])]
    if not headings:
        raise ValueError('exact 参考文献 heading not found; review document structure first')
    # Last exact heading avoids most cached TOC occurrences; the selected boundary is reported.
    boundary = headings[-1]
    body, tail = blocks[:boundary], blocks[boundary + 1:]
    entries, current, boundary_issues, end_boundary = [], None, [], None
    for b in tail:
        structured_end = b['heading'] and END_HEADING.match(b['text'])
        plain_end = not b['heading'] and PLAIN_END.fullmatch(b['text'].strip())
        if structured_end or plain_end:
            end_boundary = {'location': b['location'], 'method': 'semantic_heading' if structured_end else 'plain_heading_candidate'}
            if plain_end:
                boundary_issues.append({'kind': 'references_end_boundary_candidate', 'location': b['location']})
            break
        match = ENTRY.match(b['text'])
        if match:
            current = {'number': int(plain_number(match[1])), 'location': b['location'], 'text': match[2]}
            entries.append(current)
        elif current and b['text'].strip():
            current['text'] += ' ' + b['text'].strip()
    if not entries:
        raise ValueError('no [n] reference entries found; automatic list labels may require Word display/export review')
    seen, citations, issues = [], [], list(boundary_issues)
    for b in body:
        for m in MARKER.finditer(b['text']):
            try:
                numbers = expand(m[1])
            except ValueError as exc:
                issues.append({'kind': 'ambiguous_numeric_bracket', 'location': b['location'], 'detail': str(exc)})
                continue
            citations.append({'location': b['location'], 'numbers': numbers})
            if b['heading']:
                issues.append({'kind': 'citation_in_heading_candidate', 'location': b['location'], 'numbers': numbers})
            for n in numbers:
                if n not in seen:
                    seen.append(n)
    ids = [e['number'] for e in entries]
    repeated = sorted(n for n, count in Counter(ids).items() if count > 1)
    missing, unused = sorted(set(seen) - set(ids)), sorted(set(ids) - set(seen))
    if repeated:
        issues.append({'kind': 'duplicate_reference_numbers', 'numbers': repeated})
    if missing:
        issues.append({'kind': 'citation_without_reference_candidate', 'numbers': missing})
    if unused:
        issues.append({'kind': 'uncited_reference_candidate', 'numbers': unused})
    if seen != list(range(1, len(seen) + 1)):
        issues.append({'kind': 'first_appearance_order_candidate', 'observed': seen})
    if ids != list(range(1, len(ids) + 1)):
        issues.append({'kind': 'reference_list_sequence', 'observed': ids})
    dois = {}
    for entry in entries:
        match = re.search(r'10\.\d{4,9}/[^\s，。；]+', entry['text'], re.I)
        if match:
            key = match[0].rstrip('.,;)]').casefold()
            dois.setdefault(key, []).append(entry['number'])
    for doi, ns in dois.items():
        if len(ns) > 1:
            issues.append({'kind': 'duplicate_DOI_candidate', 'numbers': ns, 'doi': doi})
    return {'status': 'partial', 'read_only': True,
            'references_boundary': blocks[boundary]['location'],
            'references_end_boundary': end_boundary,
            'reference_entries': len(entries), 'citation_marker_candidates': len(citations),
            'first_appearance': seen, 'issues': issues, 'citations': citations,
            'tracked_changes_present': revisions,
            'caveats': ['Numeric brackets may be mathematical intervals or other labels; confirm candidates in context.',
                        'DOCX automatic numbering, revisions, footnotes/endnotes, textboxes and cached fields need separate review.',
                        'A clean result does not establish complete citation coverage or bibliographic conformance.'],
            'not_verified': UNVERIFIED}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    args = parser.parse_args()
    try:
        result = audit(args.input)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1 if result['issues'] else 0
    except (OSError, ValueError, KeyError, ET.ParseError, BadZipFile) as exc:
        print(json.dumps({'status': 'error', 'message': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
