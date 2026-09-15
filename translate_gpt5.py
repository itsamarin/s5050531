#!/usr/bin/env python3
"""
English to Thai Translation using OpenAI GPT-5
Supports .pptx, .docx, and .xlsx with structure and formatting preservation
"""

import os
import sys
import json
import shutil
import re
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Config ──────────────────────────────────────────────────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MODEL = "gpt-5"          # update if OpenAI releases a versioned alias
BATCH_SIZE = 40          # strings per API call (keep prompts manageable)
SUPPORTED_EXTS = {".pptx", ".docx", ".xlsx"}

# ── OpenAI client ────────────────────────────────────────────────────────────
try:
    from openai import OpenAI
    _client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
except ImportError:
    _client = None


# ── Glossary ─────────────────────────────────────────────────────────────────
def load_glossary() -> dict:
    p = Path("glossary.json")
    if p.exists():
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


# ── Core translation ─────────────────────────────────────────────────────────
def _check_client():
    if not OPENAI_API_KEY:
        print("Error: OPENAI_API_KEY not set in .env")
        sys.exit(1)
    if _client is None:
        print("Error: openai package not installed — run: pip install openai")
        sys.exit(1)


def _is_translatable(text: str) -> bool:
    """Return False for blank strings, numbers-only, or single symbols."""
    stripped = text.strip()
    if not stripped:
        return False
    if re.fullmatch(r"[\d\s.,/%$€£¥+\-=<>*#@!?()[\]{}&|^~`\"'\\/:;]+", stripped):
        return False
    return True


def translate_batch(texts: list[str], glossary: dict | None = None) -> list[str]:
    """
    Translate a list of English strings to Thai via GPT-5.
    Returns a list of the same length.
    """
    if not texts:
        return []

    glossary_note = ""
    if glossary:
        pairs = "; ".join(f'"{k}" → "{v}"' for k, v in list(glossary.items())[:20])
        glossary_note = f"\n\nGlossary (prefer these translations): {pairs}"

    prompt = (
        "You are a professional English-to-Thai translator.\n"
        "Translate each item in the JSON array below from English to Thai.\n"
        "Rules:\n"
        "- Return ONLY a JSON array with the same number of items, in the same order.\n"
        "- Preserve numbers, punctuation, line breaks, and special characters exactly.\n"
        "- Do not translate brand names, proper nouns, acronyms, or file paths.\n"
        "- If an item is already in Thai or is untranslatable, return it unchanged."
        f"{glossary_note}\n\n"
        f"Input: {json.dumps(texts, ensure_ascii=False)}"
    )

    response = _client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Return only valid JSON — no markdown, no explanation."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
    )

    raw = response.choices[0].message.content.strip()
    # Strip markdown code fences if GPT wraps the response
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)

    parsed = json.loads(raw)
    # GPT sometimes wraps the array in {"translations": [...]}
    if isinstance(parsed, dict):
        parsed = next(iter(parsed.values()))

    if len(parsed) != len(texts):
        raise ValueError(f"GPT returned {len(parsed)} items but expected {len(texts)}")

    return [str(t) for t in parsed]


def translate_texts(texts: list[str], glossary: dict | None = None) -> list[str]:
    """Translate a large list by splitting into BATCH_SIZE chunks."""
    results: list[str] = []
    for i in range(0, len(texts), BATCH_SIZE):
        chunk = texts[i: i + BATCH_SIZE]
        print(f"   Batch {i // BATCH_SIZE + 1}/{(len(texts) - 1) // BATCH_SIZE + 1} "
              f"({len(chunk)} strings)...")
        results.extend(translate_batch(chunk, glossary))
    return results


# ── PPTX ─────────────────────────────────────────────────────────────────────
def _pptx_collect(prs) -> tuple[list[str], list[tuple]]:
    """
    Walk all slides (content + notes) and collect translatable paragraph texts.
    Returns (texts, locations) where each location is (paragraph_obj, original_text).
    """
    from pptx.util import Pt

    texts: list[str] = []
    locations: list[tuple] = []

    for slide in prs.slides:
        # Slide shapes
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                full_text = para.text
                if _is_translatable(full_text):
                    texts.append(full_text)
                    locations.append(("para", para, full_text))

        # Speaker notes
        if slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            for para in tf.paragraphs:
                full_text = para.text
                if _is_translatable(full_text):
                    texts.append(full_text)
                    locations.append(("para", para, full_text))

    return texts, locations


def _pptx_apply(locations: list[tuple], translations: list[str]):
    """Write translated text back into paragraph runs."""
    for (kind, para, original), translated in zip(locations, translations):
        if not para.runs:
            continue
        # Put the full translation into the first run, clear the rest
        para.runs[0].text = translated
        for run in para.runs[1:]:
            run.text = ""


def translate_pptx(input_path: Path, output_path: Path, glossary: dict) -> bool:
    try:
        from pptx import Presentation
    except ImportError:
        print("   Error: python-pptx not installed — run: pip install python-pptx")
        return False

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)

        prs = Presentation(output_path)
        texts, locations = _pptx_collect(prs)

        if not texts:
            print("   No translatable text found.")
            return True

        print(f"   Found {len(texts)} text elements to translate.")
        translations = translate_texts(texts, glossary)
        _pptx_apply(locations, translations)
        prs.save(output_path)

        print(f"   Saved to: {output_path}")
        return True

    except Exception as e:
        print(f"   Error: {e}")
        return False


# ── DOCX ─────────────────────────────────────────────────────────────────────
def _docx_collect(doc) -> tuple[list[str], list]:
    texts: list[str] = []
    locations: list = []

    def collect_paragraph(para):
        full_text = para.text
        if _is_translatable(full_text):
            texts.append(full_text)
            locations.append(para)

    # Body paragraphs
    for para in doc.paragraphs:
        collect_paragraph(para)

    # Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    collect_paragraph(para)

    # Headers and footers
    for section in doc.sections:
        for hf in (section.header, section.footer):
            if hf is not None:
                for para in hf.paragraphs:
                    collect_paragraph(para)

    return texts, locations


def _docx_apply(locations, translations: list[str]):
    for para, translated in zip(locations, translations):
        if not para.runs:
            continue
        para.runs[0].text = translated
        for run in para.runs[1:]:
            run.text = ""


def translate_docx(input_path: Path, output_path: Path, glossary: dict) -> bool:
    try:
        from docx import Document
    except ImportError:
        print("   Error: python-docx not installed — run: pip install python-docx")
        return False

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)

        doc = Document(output_path)
        texts, locations = _docx_collect(doc)

        if not texts:
            print("   No translatable text found.")
            return True

        print(f"   Found {len(texts)} paragraphs to translate.")
        translations = translate_texts(texts, glossary)
        _docx_apply(locations, translations)
        doc.save(output_path)

        print(f"   Saved to: {output_path}")
        return True

    except Exception as e:
        print(f"   Error: {e}")
        return False


# ── XLSX ─────────────────────────────────────────────────────────────────────
def _xlsx_collect(wb) -> tuple[list[str], list[tuple]]:
    """Collect string cell values (skip formulas and numbers)."""
    texts: list[str] = []
    locations: list[tuple] = []  # (worksheet, row, col)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "s" and cell.value and _is_translatable(str(cell.value)):
                    texts.append(str(cell.value))
                    locations.append((ws, cell.row, cell.column))

    return texts, locations


def _xlsx_apply(wb, locations: list[tuple], translations: list[str]):
    for (ws, row, col), translated in zip(locations, translations):
        ws.cell(row=row, column=col).value = translated


def translate_xlsx(input_path: Path, output_path: Path, glossary: dict) -> bool:
    try:
        import openpyxl
    except ImportError:
        print("   Error: openpyxl not installed — run: pip install openpyxl")
        return False

    try:
        import openpyxl
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)

        wb = openpyxl.load_workbook(output_path)
        texts, locations = _xlsx_collect(wb)

        if not texts:
            print("   No translatable text found.")
            return True

        print(f"   Found {len(texts)} cells to translate.")
        translations = translate_texts(texts, glossary)
        _xlsx_apply(wb, locations, translations)
        wb.save(output_path)

        print(f"   Saved to: {output_path}")
        return True

    except Exception as e:
        print(f"   Error: {e}")
        return False


# ── Dispatcher ───────────────────────────────────────────────────────────────
def translate_file(input_path: Path, output_path: Path | None = None) -> bool:
    ext = input_path.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        print(f"   Unsupported format: {ext}  (supported: {', '.join(sorted(SUPPORTED_EXTS))})")
        return False

    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}_thai{input_path.suffix}"

    glossary = load_glossary()

    print(f"\nTranslating: {input_path.name}")
    print(f"   Format : {ext.upper()}")
    print(f"   Engine : GPT-5")
    print(f"   Target : Thai")

    if ext == ".pptx":
        return translate_pptx(input_path, output_path, glossary)
    elif ext == ".docx":
        return translate_docx(input_path, output_path, glossary)
    elif ext == ".xlsx":
        return translate_xlsx(input_path, output_path, glossary)


def translate_folder(folder: Path, out_folder: Path):
    files = [
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS and not f.name.startswith(".")
    ]
    if not files:
        print(f"No supported files found in {folder}")
        return

    print(f"\n{'='*55}")
    print(f"Translating {len(files)} file(s) from {folder}")
    print(f"Output folder: {out_folder}")
    print(f"{'='*55}")

    ok = 0
    for f in sorted(files):
        if translate_file(f, out_folder / f.name):
            ok += 1

    print(f"\n{'='*55}")
    print(f"Done: {ok}/{len(files)} files translated successfully")
    print(f"{'='*55}")


# ── CLI ───────────────────────────────────────────────────────────────────────
def print_usage():
    print("""
English to Thai Translation  —  GPT-5
======================================

Usage:
  python translate_gpt5.py <command|file>

Commands:
  presentations    Translate all files in presentations/english/
  scripts          Translate all .docx/.xlsx in scripts/english/
  all              Translate presentations + scripts
  <file.pptx>      Translate a single file (any supported format)

Supported formats:
  .pptx   PowerPoint  (slides + speaker notes)
  .docx   Word document
  .xlsx   Excel spreadsheet

Setup:
  1. pip install -r requirements.txt
  2. Add OPENAI_API_KEY=sk-... to your .env file

Output:
  Single file  →  same folder,  <name>_thai.<ext>
  Folder mode  →  presentations/thai/  or  scripts/thai/
""")


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help", "help"):
        print_usage()
        sys.exit(0)

    _check_client()
    command = sys.argv[1]

    if command == "presentations":
        translate_folder(Path("presentations/english"), Path("presentations/thai"))

    elif command == "scripts":
        translate_folder(Path("scripts/english"), Path("scripts/thai"))

    elif command == "all":
        translate_folder(Path("presentations/english"), Path("presentations/thai"))
        translate_folder(Path("scripts/english"), Path("scripts/thai"))

    else:
        p = Path(command)
        if not p.exists():
            print(f"Error: file not found — {command}")
            sys.exit(1)
        translate_file(p)


if __name__ == "__main__":
    main()
