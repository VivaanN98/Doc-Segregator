"""
Universal DOCX Section Segregation Script
==========================================
Processes all .docx files in input/ subdirectories and splits them by
section markers (A-G, skipping E). Each section is output as a separate
file, cloned from the source to preserve ALL formatting (fonts, bold,
italic, styles, numbering, images).

Output: output/{SubjectFolder}/Unit {N} Sec {letter}.docx
"""

import os
import re
import copy
import io
import glob
from docx import Document
from docx.oxml.ns import qn

INPUT_DIR = 'input'
OUTPUT_DIR = 'output'
SECTIONS_TO_SKIP = {'E'}

# Regex to detect section markers like "A. Summary", "B. In-Depth..."
SECTION_RE = re.compile(r'^([A-G])\.\s')

# Known section title fragments — used to validate that a match is a real
# section marker and not a random paragraph starting with "C. Think of..."
SECTION_TITLE_PATTERNS = [
    'summary', 'concept map', 'unit number', 'unit information',
    'in-depth', 'explanation', 'analysis',
    'flashcard', 'memory',
    'key terms', 'keyword',
    'question bank',
    'exam strategy', 'strategy note',
    'fun fact',
    # Geography uses inline unit/chapter info like "A. Unit I, Chapter 1:"
    'unit i', 'unit ii', 'unit iii', 'unit iv', 'unit v', 'unit vi',
]

# Body element tags
W_P = qn('w:p')      # paragraph
W_TBL = qn('w:tbl')  # table
W_SECT_PR = qn('w:sectPr')  # section properties (page layout)


def get_paragraph_text(element):
    """
    Extract plain text from a w:p element, inserting \\n for w:br elements.
    This allows detection of inline section markers after line breaks.
    """
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    parts = []
    for run in element.findall(f'{{{ns}}}r'):
        for child in run:
            if child.tag == f'{{{ns}}}t' and child.text:
                parts.append(child.text)
            elif child.tag == f'{{{ns}}}br':
                parts.append('\n')
    return ''.join(parts).strip()


def is_real_section_marker(full_text, letter):
    """
    Verify that a detected section marker is a real structural section,
    not a content paragraph that happens to start with a letter and period.
    """
    # Get the text after the section letter (e.g. "A. Summary & Concept Map" -> "Summary & Concept Map")
    after_letter = full_text.split('.', 1)[1].strip().lower() if '.' in full_text else ''

    # Check if any known section title fragment appears
    for pattern in SECTION_TITLE_PATTERNS:
        if pattern in after_letter:
            return True

    return False


def find_section_boundaries(body):
    """
    Scan all body children to find section boundaries.
    Returns a list of (element_index, section_letter) for each section marker.

    Also detects inline section markers embedded via newline characters
    (e.g. "Chapter 1: Sets\\nA. Summary & Concept Map").
    """
    boundaries = []
    children = list(body)

    for idx, child in enumerate(children):
        if child.tag == W_P:
            txt = get_paragraph_text(child)
            if not txt:
                continue

            # Check if the paragraph itself starts with a section marker
            m = SECTION_RE.match(txt)
            if m:
                letter = m.group(1)
                if is_real_section_marker(txt, letter):
                    boundaries.append((idx, letter))
                continue

            # Check for inline section marker after a newline
            # (handles cases like "Chapter 1: Sets\nA. Summary...")
            if '\n' in txt:
                lines = txt.split('\n')
                for line in lines[1:]:  # skip first line
                    line = line.strip()
                    m = SECTION_RE.match(line)
                    if m:
                        letter = m.group(1)
                        if is_real_section_marker(line, letter):
                            boundaries.append((idx, letter))
                            break  # only take the first inline match

    return boundaries


def group_into_units(boundaries):
    """
    Group section boundaries into unit-chunks.
    Each chunk starts with Section A and contains all subsequent sections
    until the next A (or end of document).

    Returns list of lists: [[boundary1, boundary2, ...], ...]
    Each boundary is (element_index, section_letter).
    """
    units = []
    current_unit = []

    for boundary in boundaries:
        letter = boundary[1]
        if letter == 'A' and current_unit:
            units.append(current_unit)
            current_unit = []
        current_unit.append(boundary)

    if current_unit:
        units.append(current_unit)

    return units


def create_section_file(raw_bytes, body, section_start_idx, section_end_idx, output_path):
    """
    Create a new DOCX for a single section by:
    1. Cloning the source document (preserves styles, numbering, fonts, themes, images)
    2. Replacing the body with only the elements for this section
    """
    # Clone from raw bytes
    new_doc = Document(io.BytesIO(raw_bytes))
    new_body = new_doc.element.body

    # Collect all source body children to copy
    source_children = list(body)

    # Identify elements to keep: from section_start_idx to section_end_idx (inclusive)
    elements_to_keep = []
    for i in range(section_start_idx, section_end_idx + 1):
        if i < len(source_children):
            elements_to_keep.append(copy.deepcopy(source_children[i]))

    # Clear the cloned document's body (keep sectPr for page layout)
    sectPr = None
    for child in list(new_body):
        if child.tag == W_SECT_PR:
            sectPr = child
        else:
            new_body.remove(child)

    # Insert our section elements
    for elem in elements_to_keep:
        if sectPr is not None:
            sectPr.addprevious(elem)
        else:
            new_body.append(elem)

    new_doc.save(output_path)


def process_file(input_path, output_folder):
    """
    Process a single DOCX file: find sections, group into units,
    and create individual output files.
    """
    print(f"\n{'='*60}")
    print(f"Processing: {input_path}")
    print(f"{'='*60}")

    # Read raw bytes for efficient cloning
    with open(input_path, 'rb') as f:
        raw_bytes = f.read()

    doc = Document(input_path)
    body = doc.element.body
    body_children = list(body)
    total_elements = len(body_children)

    print(f"  Total body elements: {total_elements}")

    # Find all section boundaries
    boundaries = find_section_boundaries(body)
    print(f"  Found {len(boundaries)} section markers")

    if not boundaries:
        print(f"  WARNING: No section markers found, skipping this file!")
        return 0

    # Group into unit-chunks
    units = group_into_units(boundaries)
    print(f"  Found {len(units)} unit-chunks")

    # Create output directory
    os.makedirs(output_folder, exist_ok=True)

    created = 0
    for unit_idx, unit_boundaries in enumerate(units):
        unit_num = unit_idx + 1

        for sec_idx, (elem_start, letter) in enumerate(unit_boundaries):
            if letter in SECTIONS_TO_SKIP:
                continue

            # Determine end of this section:
            # - If there's a next section in this unit, end = next_start - 1
            # - If this is the last section in this unit, check if there's a next unit
            # - Otherwise, end = last body element (before sectPr)
            if sec_idx + 1 < len(unit_boundaries):
                elem_end = unit_boundaries[sec_idx + 1][0] - 1
            elif unit_idx + 1 < len(units):
                # Next unit's first section A
                elem_end = units[unit_idx + 1][0][0] - 1
            else:
                # Last section of last unit - go to end of body
                elem_end = total_elements - 1
                # Skip sectPr at the end
                while elem_end >= 0 and body_children[elem_end].tag == W_SECT_PR:
                    elem_end -= 1

            filename = f"Unit {unit_num} Sec {letter}.docx"
            output_path = os.path.join(output_folder, filename)

            create_section_file(raw_bytes, body, elem_start, elem_end, output_path)
            created += 1

    print(f"  [OK] Created {created} files in '{output_folder}/'")
    return created


def main():
    print("Universal DOCX Section Segregation")
    print("===================================\n")

    # Find all input files
    input_files = []
    for subject_dir in sorted(os.listdir(INPUT_DIR)):
        subject_path = os.path.join(INPUT_DIR, subject_dir)
        if os.path.isdir(subject_path):
            for fname in os.listdir(subject_path):
                if fname.lower().endswith('.docx') and not fname.startswith('~'):
                    input_files.append((
                        os.path.join(subject_path, fname),
                        subject_dir
                    ))

    print(f"Found {len(input_files)} input files:")
    for fpath, subj in input_files:
        print(f"  [{subj}] {os.path.basename(fpath)}")

    total_created = 0
    for input_path, subject_folder in input_files:
        output_folder = os.path.join(OUTPUT_DIR, subject_folder)
        count = process_file(input_path, output_folder)
        total_created += count

    print(f"\n{'='*60}")
    print(f"TOTAL: Created {total_created} files across {len(input_files)} subjects")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
