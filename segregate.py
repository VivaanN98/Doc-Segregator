"""
DOCX Section Segregation Script
================================
Splits a monolithic math textbook DOCX into individual files
organized by Unit / Chapter / Section (A-G, skipping E).

Output: output/Unit {roman} Ch {n} Sec {letter}.docx

Preserves ALL formatting by cloning the source document's internal
parts (styles, numbering, fonts, themes) and then replacing the body
with only the relevant paragraphs.
"""

import os
import re
import copy
import io
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

INPUT_FILE = r'input\Core Math XI HS 3.0 NEW.docx'
OUTPUT_DIR = 'output'
SECTIONS_TO_SKIP = {'E'}  # Skip Section E (Question Bank)

ROMAN_MAP = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}


def parse_structure(doc):
    """
    Parse the document to identify Unit, Chapter, and Section boundaries.
    Returns a list of dicts with unit/chapter/section info and paragraph ranges.
    """
    paragraphs = doc.paragraphs
    total = len(paragraphs)

    markers = []

    for i, p in enumerate(paragraphs):
        txt = p.text.strip()
        if not txt:
            continue

        # Unit marker
        m = re.match(r'^(Unit\s+[IVXLC]+:\s+.+)$', txt)
        if m:
            markers.append((i, 'unit', m.group(1)))
            continue

        # Chapter marker (may contain Section A via newline)
        m = re.match(r'^(Chapter\s*[\d:].+?)(?:\n(.+))?$', txt, re.DOTALL)
        if m:
            markers.append((i, 'chapter', m.group(1).strip()))
            if m.group(2):
                sec_text = m.group(2).strip()
                sec_m = re.match(r'^([A-G])\.\s', sec_text)
                if sec_m:
                    markers.append((i, 'section', sec_m.group(1), True))
            continue

        # Section marker
        m = re.match(r'^([A-G])\.\s', txt)
        if m:
            markers.append((i, 'section', m.group(1), False))
            continue

    # Build hierarchical structure
    entries = []
    current_unit = None
    current_unit_num = 0
    current_chapter = None
    current_chapter_num = 0
    unit_chapter_counter = {}

    roman_to_int = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6}

    for idx, marker in enumerate(markers):
        para_idx = marker[0]
        mtype = marker[1]

        if mtype == 'unit':
            unit_text = marker[2]
            um = re.match(r'Unit\s+([IVXLC]+):', unit_text)
            if um:
                current_unit = unit_text
                roman = um.group(1)
                current_unit_num = roman_to_int.get(roman, 0)
                if current_unit_num not in unit_chapter_counter:
                    unit_chapter_counter[current_unit_num] = 0

        elif mtype == 'chapter':
            chapter_text = marker[2]
            if current_unit_num > 0:
                unit_chapter_counter[current_unit_num] += 1
                current_chapter_num = unit_chapter_counter[current_unit_num]
            current_chapter = chapter_text

        elif mtype == 'section':
            section_letter = marker[2]
            is_inline = marker[3] if len(marker) > 3 else False

            if current_unit is None or current_chapter is None:
                continue

            # End = next marker's paragraph - 1, or end of doc
            end_para = total - 1
            for next_idx in range(idx + 1, len(markers)):
                end_para = markers[next_idx][0] - 1
                break

            entries.append({
                'unit': current_unit,
                'unit_num': current_unit_num,
                'chapter': current_chapter,
                'chapter_num': current_chapter_num,
                'section': section_letter,
                'start_para': para_idx,
                'end_para': end_para,
                'is_inline': is_inline,
            })

    return entries


def get_actual_section_title(doc, entry):
    """Extract the actual section title text from the document."""
    para = doc.paragraphs[entry['start_para']]
    txt = para.text.strip()
    if entry['is_inline']:
        parts = txt.split('\n', 1)
        if len(parts) > 1:
            return parts[1].strip()
    else:
        return txt
    return None


def create_section_file(doc, raw_bytes, entry, output_path):
    """
    Create a new DOCX for a single section by cloning the source document
    (preserving all styles, numbering, fonts, themes) and replacing the body.
    """
    # Clone the source document from raw bytes so all internal parts are preserved
    new_doc = Document(io.BytesIO(raw_bytes))

    # Clear the body of the cloned document
    body = new_doc.element.body
    for child in list(body):
        # Keep sectPr (page layout settings) but remove all paragraphs and tables
        if child.tag != qn('w:sectPr'):
            body.remove(child)

    source_paras = doc.paragraphs

    # --- 1. Build Unit header paragraph from the source Unit paragraph ---
    # Find the source Unit header paragraph to clone its formatting
    unit_para_idx = find_unit_para(doc, entry['unit'])
    if unit_para_idx is not None:
        # Clone the Unit paragraph from source
        unit_elem = copy.deepcopy(source_paras[unit_para_idx]._element)
        # Insert before sectPr
        _insert_before_sectpr(body, unit_elem)
    else:
        # Fallback: create a simple bold paragraph
        _insert_before_sectpr(body, _make_bold_para(entry['unit']))

    # --- 2. Build Chapter + Section header paragraph ---
    # Clone the original chapter paragraph and modify it to include the section title
    chapter_para_idx = find_chapter_para(doc, entry)
    sec_title = get_actual_section_title(doc, entry)

    if chapter_para_idx is not None:
        orig_chapter_elem = source_paras[chapter_para_idx]._element
        # We need to build a new paragraph that has:
        #   - The chapter title runs (from the original chapter paragraph)
        #   - A line break
        #   - The section title (bold)
        # But we want to preserve the original paragraph's formatting properties

        ch_elem = copy.deepcopy(orig_chapter_elem)

        # If this is an inline section (Chapter + Section A in same para),
        # the paragraph already has both. We keep it as-is.
        if entry['is_inline']:
            _insert_before_sectpr(body, ch_elem)
        else:
            # Remove all runs and content from the cloned chapter paragraph,
            # keeping only pPr (paragraph properties)
            for child in list(ch_elem):
                if child.tag != qn('w:pPr'):
                    ch_elem.remove(child)

            # Get the run formatting from the original chapter paragraph's first run
            orig_rPr = _get_first_rPr(orig_chapter_elem)

            # Add chapter title run
            ch_run = _make_run(entry['chapter'], orig_rPr)
            ch_elem.append(ch_run)

            # Add line break
            br_run = _make_br_run(orig_rPr)
            ch_elem.append(br_run)

            # Add section title run
            sec_run = _make_run(sec_title or f"{entry['section']}.", orig_rPr)
            ch_elem.append(sec_run)

            _insert_before_sectpr(body, ch_elem)
    else:
        # Fallback
        _insert_before_sectpr(body, _make_bold_para(f"{entry['chapter']}\n{sec_title}"))

    # --- 3. Copy content paragraphs ---
    start = entry['start_para']
    end = entry['end_para']
    content_start = start + 1  # Skip the section marker paragraph

    for i in range(content_start, end + 1):
        if i < len(source_paras):
            elem = copy.deepcopy(source_paras[i]._element)
            _insert_before_sectpr(body, elem)

    new_doc.save(output_path)


def find_unit_para(doc, unit_text):
    """Find the paragraph index of a Unit header matching the given text."""
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == unit_text:
            return i
    return None


def find_chapter_para(doc, entry):
    """Find the paragraph index of the Chapter header for this entry."""
    start = entry['start_para']
    # For inline sections, the chapter para IS the start_para
    if entry['is_inline']:
        return start
    # Otherwise, search backwards from the section start to find the chapter para
    for i in range(start - 1, max(start - 10, -1), -1):
        txt = doc.paragraphs[i].text.strip()
        if txt.startswith('Chapter'):
            return i
    return None


def _get_first_rPr(para_elem):
    """Extract the rPr (run properties) from the first run of a paragraph element."""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for run in para_elem.findall(f'{{{ns}}}r'):
        rPr = run.find(f'{{{ns}}}rPr')
        if rPr is not None:
            return copy.deepcopy(rPr)
    return None


def _make_run(text, rPr=None):
    """Create a w:r element with given text and optional rPr."""
    run = OxmlElement('w:r')
    if rPr is not None:
        run.append(copy.deepcopy(rPr))
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)
    return run


def _make_br_run(rPr=None):
    """Create a w:r element containing a line break."""
    run = OxmlElement('w:r')
    if rPr is not None:
        run.append(copy.deepcopy(rPr))
    br = OxmlElement('w:br')
    run.append(br)
    return run


def _make_bold_para(text):
    """Create a simple bold paragraph element (fallback)."""
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    b = OxmlElement('w:b')
    rPr.append(b)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _insert_before_sectpr(body, element):
    """Insert an element before the sectPr in the body, or at the end."""
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(element)
    else:
        body.append(element)


def get_unit_roman(unit_num):
    return ROMAN_MAP.get(unit_num, str(unit_num))


def main():
    print(f"Loading input document: {INPUT_FILE}")

    # Read raw bytes once for efficient cloning
    with open(INPUT_FILE, 'rb') as f:
        raw_bytes = f.read()

    doc = Document(INPUT_FILE)
    print(f"Total paragraphs: {len(doc.paragraphs)}")

    print("\nParsing document structure...")
    entries = parse_structure(doc)
    print(f"Found {len(entries)} sections total")

    entries = [e for e in entries if e['section'] not in SECTIONS_TO_SKIP]
    print(f"After skipping Section E: {len(entries)} sections to extract")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    created_files = []
    for i, entry in enumerate(entries):
        unit_roman = get_unit_roman(entry['unit_num'])
        filename = f"Unit {unit_roman} Ch {entry['chapter_num']} Sec {entry['section']}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)

        print(f"  [{i+1}/{len(entries)}] Creating: {filename}")
        create_section_file(doc, raw_bytes, entry, output_path)
        created_files.append(filename)

    print(f"\n✓ Created {len(created_files)} files in '{OUTPUT_DIR}/' directory")

    # Summary
    units = {}
    for entry in entries:
        u = get_unit_roman(entry['unit_num'])
        if u not in units:
            units[u] = set()
        units[u].add(entry['chapter_num'])

    print("\nSummary:")
    for u in sorted(units.keys()):
        chapters = sorted(units[u])
        print(f"  Unit {u}: {len(chapters)} chapters (Ch {', '.join(map(str, chapters))})")


if __name__ == '__main__':
    main()
