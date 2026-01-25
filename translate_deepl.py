#!/usr/bin/env python3
"""
English to Thai Translation using DeepL API
Optimized for PowerPoint presentations with format preservation
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import deepl

# Load environment variables
load_dotenv()

# Initialize DeepL client
DEEPL_API_KEY = os.getenv("DEEPL_API_KEY")

if DEEPL_API_KEY:
    translator = deepl.Translator(DEEPL_API_KEY)
else:
    translator = None


def check_api_key():
    """Verify DeepL API key is configured."""
    if not DEEPL_API_KEY:
        print("❌ Error: DEEPL_API_KEY not set")
        print("Add your DeepL API key to .env file")
        sys.exit(1)


def get_usage_info():
    """Display current API usage."""
    try:
        usage = translator.get_usage()
        if usage.character.limit:
            used = usage.character.count
            limit = usage.character.limit
            percent = (used / limit) * 100
            print(f"📊 API Usage: {used:,} / {limit:,} characters ({percent:.1f}%)")
        return usage
    except Exception as e:
        print(f"Could not fetch usage: {e}")
        return None


def translate_text(
    text: str,
    target_lang: str = "TH",
    formality: str = "default",
    context: str = None
) -> str:
    """
    Translate text using DeepL API.

    Args:
        text: English text to translate
        target_lang: Target language code (TH for Thai)
        formality: "less", "more", "default", "prefer_less", "prefer_more"
        context: Additional context to improve translation
    """
    check_api_key()

    result = translator.translate_text(
        text,
        source_lang="EN",
        target_lang=target_lang,
        formality=formality if formality != "default" else None,
        context=context
    )

    return result.text


def translate_document(
    input_path: Path,
    output_path: Path,
    target_lang: str = "TH",
    formality: str = "default"
) -> bool:
    """
    Translate a document (PPTX, DOCX, PDF, etc.) using DeepL API.
    Preserves original formatting.

    Args:
        input_path: Path to source document
        output_path: Path for translated document
        target_lang: Target language code
        formality: Formality level
    """
    check_api_key()

    try:
        print(f"\n📄 Translating: {input_path.name}")
        print(f"   Format: {input_path.suffix.upper()}")
        print(f"   Target: Thai (TH)")

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # DeepL document translation (preserves formatting)
        print(f"   🔄 Uploading and translating...")

        with open(input_path, "rb") as in_file, open(output_path, "wb") as out_file:
            translator.translate_document(
                in_file,
                out_file,
                source_lang="EN",
                target_lang=target_lang,
                formality=formality if formality != "default" else None
            )

        print(f"   ✅ Saved to: {output_path}")
        return True

    except deepl.DocumentTranslationException as e:
        print(f"   ❌ Translation failed: {e}")
        print(f"      Document ID: {e.document_handle.id}")
        return False
    except deepl.DeepLException as e:
        print(f"   ❌ DeepL Error: {e}")
        return False
    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def translate_text_file(
    input_path: Path,
    output_path: Path,
    formality: str = "default"
) -> bool:
    """Translate a text file (.txt, .md)."""
    try:
        print(f"\n📄 Translating: {input_path.name}")

        with open(input_path, 'r', encoding='utf-8') as f:
            english_content = f.read()

        if not english_content.strip():
            print(f"   ⚠️  Skipping empty file")
            return False

        word_count = len(english_content.split())
        print(f"   Words: ~{word_count}")
        print(f"   🔄 Translating...")

        thai_content = translate_text(english_content, formality=formality)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(thai_content)

        print(f"   ✅ Saved to: {output_path}")
        return True

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def translate_presentations(formality: str = "default"):
    """Translate all presentations in the presentations folder."""
    english_dir = Path("presentations/english")
    thai_dir = Path("presentations/thai")

    if not english_dir.exists():
        print("No presentations/english directory found")
        return

    # Supported presentation formats
    patterns = ["*.pptx", "*.ppt", "*.txt", "*.md"]
    files = []
    for pattern in patterns:
        files.extend(english_dir.glob(pattern))

    files = [f for f in files if not f.name.startswith('.')]

    if not files:
        print("No presentation files found in presentations/english/")
        return

    print(f"\n{'='*50}")
    print(f"Translating {len(files)} presentation(s) with DeepL")
    print(f"From: {english_dir}")
    print(f"To:   {thai_dir}")
    print(f"{'='*50}")

    success_count = 0
    for file_path in sorted(files):
        output_path = thai_dir / file_path.name

        if file_path.suffix.lower() in ['.pptx', '.ppt']:
            # Use document translation API (preserves formatting)
            if translate_document(file_path, output_path, formality=formality):
                success_count += 1
        else:
            # Use text translation for .txt/.md files
            if translate_text_file(file_path, output_path, formality=formality):
                success_count += 1

    print(f"\n{'='*50}")
    print(f"Completed: {success_count}/{len(files)} presentations translated")
    print(f"{'='*50}")


def translate_scripts(formality: str = "default"):
    """Translate all scripts in the scripts folder."""
    english_dir = Path("scripts/english")
    thai_dir = Path("scripts/thai")

    if not english_dir.exists():
        print("No scripts/english directory found")
        return

    files = list(english_dir.glob("*.txt")) + list(english_dir.glob("*.md"))
    files = [f for f in files if not f.name.startswith('.')]

    if not files:
        print("No script files found in scripts/english/")
        return

    print(f"\n{'='*50}")
    print(f"Translating {len(files)} script(s) with DeepL")
    print(f"From: {english_dir}")
    print(f"To:   {thai_dir}")
    print(f"{'='*50}")

    success_count = 0
    for file_path in sorted(files):
        output_path = thai_dir / file_path.name
        if translate_text_file(file_path, output_path, formality=formality):
            success_count += 1

    print(f"\n{'='*50}")
    print(f"Completed: {success_count}/{len(files)} scripts translated")
    print(f"{'='*50}")


def interactive_translate():
    """Interactive translation mode."""
    check_api_key()

    print("\n🌐 Interactive English → Thai Translation (DeepL)")
    print("Type 'quit' to exit\n")

    while True:
        print("-" * 40)
        text = input("English: ").strip()

        if text.lower() == 'quit':
            break

        if not text:
            continue

        print("\nTranslating...")
        thai = translate_text(text)
        print(f"\nThai: {thai}\n")


def print_usage():
    """Print usage instructions."""
    print("""
🌐 English to Thai Translation Tool (DeepL API)
================================================

Usage:
  python translate_deepl.py <command> [options]

Commands:
  presentations    Translate all presentations (.pptx, .txt, .md)
  scripts          Translate all scripts (.txt, .md)
  all              Translate everything
  interactive      Interactive translation mode
  usage            Show API usage statistics
  <file>           Translate a single file

Options:
  --formal         Use formal Thai language
  --informal       Use informal Thai language

Supported Formats:
  📊 Presentations: .pptx (preserves formatting), .txt, .md
  📝 Scripts: .txt, .md
  📄 Documents: .docx, .pdf, .xlsx (via single file mode)

Examples:
  python translate_deepl.py presentations
  python translate_deepl.py scripts --formal
  python translate_deepl.py all
  python translate_deepl.py my_slides.pptx
  python translate_deepl.py interactive

Setup:
  1. pip install -r requirements.txt
  2. Add DEEPL_API_KEY to .env file
""")


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print_usage()
        sys.exit(1)

    args = sys.argv[1:]
    command = args[0]

    # Parse formality option
    formality = "default"
    if "--formal" in args:
        formality = "more"
    elif "--informal" in args:
        formality = "less"

    if command == "presentations":
        check_api_key()
        get_usage_info()
        translate_presentations(formality)
        get_usage_info()

    elif command == "scripts":
        check_api_key()
        get_usage_info()
        translate_scripts(formality)
        get_usage_info()

    elif command == "all":
        check_api_key()
        get_usage_info()
        translate_scripts(formality)
        translate_presentations(formality)
        get_usage_info()

    elif command == "interactive":
        interactive_translate()

    elif command == "usage":
        check_api_key()
        get_usage_info()

    elif command in ("--help", "-h", "help"):
        print_usage()

    else:
        # Single file translation
        check_api_key()
        input_path = Path(command)

        if not input_path.exists():
            print(f"❌ File not found: {command}")
            sys.exit(1)

        output_path = input_path.parent / f"{input_path.stem}_thai{input_path.suffix}"

        if input_path.suffix.lower() in ['.pptx', '.ppt', '.docx', '.pdf', '.xlsx']:
            translate_document(input_path, output_path, formality=formality)
        else:
            translate_text_file(input_path, output_path, formality=formality)


if __name__ == "__main__":
    main()
