#!/usr/bin/env python3
"""
PowerPoint Script Manager
- Attach scripts to slides as speaker notes
- Extract scripts from slides
- Translate slides + scripts together
"""

import os
import sys
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt


def attach_script_to_pptx(pptx_path: Path, script_path: Path, output_path: Path = None) -> bool:
    """
    Attach a script file to PowerPoint slides as speaker notes.

    Script format (script.txt):
    ---
    [Slide 1]
    This is what I say for slide 1...

    [Slide 2]
    This is what I say for slide 2...
    ---

    Args:
        pptx_path: Path to PowerPoint file
        script_path: Path to script file
        output_path: Output path (default: adds _with_notes suffix)
    """
    if not pptx_path.exists():
        print(f"❌ PowerPoint not found: {pptx_path}")
        return False

    if not script_path.exists():
        print(f"❌ Script not found: {script_path}")
        return False

    # Parse script file
    scripts = parse_script_file(script_path)

    if not scripts:
        print("❌ No scripts found in file")
        return False

    print(f"📄 Loading: {pptx_path.name}")
    prs = Presentation(str(pptx_path))

    slide_count = len(prs.slides)
    print(f"   Slides: {slide_count}")
    print(f"   Scripts: {len(scripts)}")

    # Attach scripts to slides
    attached = 0
    for slide_num, script_text in scripts.items():
        if slide_num <= slide_count:
            slide = prs.slides[slide_num - 1]  # 0-indexed

            # Get or create notes slide
            notes_slide = slide.notes_slide
            notes_frame = notes_slide.notes_text_frame

            # Set the script as speaker notes
            notes_frame.text = script_text.strip()

            print(f"   ✓ Slide {slide_num}: Added speaker notes")
            attached += 1
        else:
            print(f"   ⚠️  Slide {slide_num}: Skipped (slide doesn't exist)")

    # Save output
    if output_path is None:
        output_path = pptx_path.parent / f"{pptx_path.stem}_with_notes.pptx"

    prs.save(str(output_path))
    print(f"\n✅ Saved: {output_path}")
    print(f"   Attached {attached} speaker notes")

    return True


def parse_script_file(script_path: Path) -> dict:
    """
    Parse a script file into slide-indexed dictionary.

    Supported formats:

    Format 1 - Bracketed:
    [Slide 1]
    Script for slide 1...

    [Slide 2]
    Script for slide 2...

    Format 2 - Numbered:
    1. Script for slide 1...

    2. Script for slide 2...

    Format 3 - Separator:
    Script for slide 1...
    ---
    Script for slide 2...
    ---
    """
    with open(script_path, 'r', encoding='utf-8') as f:
        content = f.read()

    scripts = {}

    # Try Format 1: [Slide N] or [N]
    import re
    pattern = r'\[(?:Slide\s*)?(\d+)\]\s*\n(.*?)(?=\[(?:Slide\s*)?\d+\]|\Z)'
    matches = re.findall(pattern, content, re.DOTALL | re.IGNORECASE)

    if matches:
        for slide_num, text in matches:
            scripts[int(slide_num)] = text.strip()
        return scripts

    # Try Format 2: Numbered list (1. 2. 3.)
    pattern = r'^(\d+)\.\s*(.*?)(?=^\d+\.|\Z)'
    matches = re.findall(pattern, content, re.DOTALL | re.MULTILINE)

    if matches:
        for slide_num, text in matches:
            scripts[int(slide_num)] = text.strip()
        return scripts

    # Try Format 3: Separator (---)
    parts = re.split(r'\n---+\n', content)
    if len(parts) > 1:
        for i, text in enumerate(parts, 1):
            if text.strip():
                scripts[i] = text.strip()
        return scripts

    # Fallback: Single script for slide 1
    if content.strip():
        scripts[1] = content.strip()

    return scripts


def extract_scripts_from_pptx(pptx_path: Path, output_path: Path = None) -> bool:
    """
    Extract speaker notes from PowerPoint into a script file.

    Args:
        pptx_path: Path to PowerPoint file
        output_path: Output path (default: same name with .txt)
    """
    if not pptx_path.exists():
        print(f"❌ PowerPoint not found: {pptx_path}")
        return False

    print(f"📄 Loading: {pptx_path.name}")
    prs = Presentation(str(pptx_path))

    scripts = []
    for i, slide in enumerate(prs.slides, 1):
        try:
            notes_slide = slide.notes_slide
            notes_text = notes_slide.notes_text_frame.text.strip()

            if notes_text:
                scripts.append(f"[Slide {i}]\n{notes_text}")
                print(f"   ✓ Slide {i}: Extracted notes ({len(notes_text)} chars)")
            else:
                print(f"   - Slide {i}: No notes")
        except Exception:
            print(f"   - Slide {i}: No notes")

    if not scripts:
        print("\n⚠️  No speaker notes found in presentation")
        return False

    # Save output
    if output_path is None:
        output_path = pptx_path.parent / f"{pptx_path.stem}_script.txt"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n\n".join(scripts))

    print(f"\n✅ Saved: {output_path}")
    print(f"   Extracted {len(scripts)} slide scripts")

    return True


def list_slide_content(pptx_path: Path):
    """
    List all slides with their content and notes.
    """
    if not pptx_path.exists():
        print(f"❌ PowerPoint not found: {pptx_path}")
        return

    print(f"📄 {pptx_path.name}")
    print("=" * 60)

    prs = Presentation(str(pptx_path))

    for i, slide in enumerate(prs.slides, 1):
        print(f"\n[Slide {i}]")

        # Get slide text
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text.strip())

        if slide_text:
            print(f"  Content: {slide_text[0][:50]}..." if len(slide_text[0]) > 50 else f"  Content: {slide_text[0]}")
        else:
            print("  Content: (no text)")

        # Get notes
        try:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                preview = notes[:80].replace('\n', ' ')
                print(f"  Notes: {preview}..." if len(notes) > 80 else f"  Notes: {preview}")
            else:
                print("  Notes: (none)")
        except Exception:
            print("  Notes: (none)")

    print("\n" + "=" * 60)
    print(f"Total: {len(prs.slides)} slides")


def create_script_template(pptx_path: Path, output_path: Path = None):
    """
    Create a script template file based on slide count.
    """
    if not pptx_path.exists():
        print(f"❌ PowerPoint not found: {pptx_path}")
        return

    prs = Presentation(str(pptx_path))
    slide_count = len(prs.slides)

    if output_path is None:
        output_path = pptx_path.parent / f"{pptx_path.stem}_script_template.txt"

    lines = []
    for i, slide in enumerate(prs.slides, 1):
        # Get slide title if available
        title = ""
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                title = shape.text.strip()[:50]
                break

        lines.append(f"[Slide {i}]")
        if title:
            lines.append(f"# {title}")
        lines.append("Write your script for this slide here...\n")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))

    print(f"✅ Created template: {output_path}")
    print(f"   {slide_count} slides")
    print("\nEdit the template, then run:")
    print(f"   python pptx_scripts.py attach {pptx_path} {output_path}")


def print_usage():
    """Print usage instructions."""
    print("""
📝 PowerPoint Script Manager
=============================

Attach scripts to PowerPoint slides as speaker notes for Presenter View.

Commands:

  attach <pptx> <script>    Attach script file to slides as speaker notes
  extract <pptx>            Extract speaker notes into a script file
  list <pptx>               List all slides with content and notes
  template <pptx>           Create a script template from slides

Examples:

  python pptx_scripts.py attach presentation.pptx script.txt
  python pptx_scripts.py extract presentation.pptx
  python pptx_scripts.py list presentation.pptx
  python pptx_scripts.py template presentation.pptx

Script File Format:

  [Slide 1]
  This is the script for slide 1.
  You can have multiple lines.

  [Slide 2]
  This is the script for slide 2.
  It will appear in Presenter View.

  [Slide 3]
  And so on...

Presenter View:
  When presenting, press F5 or use Slideshow > Presenter View
  to see your slides with speaker notes.
""")


def main():
    if len(sys.argv) < 2:
        print_usage()
        sys.exit(1)

    command = sys.argv[1]

    if command == "attach":
        if len(sys.argv) < 4:
            print("Usage: python pptx_scripts.py attach <pptx> <script>")
            sys.exit(1)
        pptx_path = Path(sys.argv[2])
        script_path = Path(sys.argv[3])
        output_path = Path(sys.argv[4]) if len(sys.argv) > 4 else None
        attach_script_to_pptx(pptx_path, script_path, output_path)

    elif command == "extract":
        if len(sys.argv) < 3:
            print("Usage: python pptx_scripts.py extract <pptx>")
            sys.exit(1)
        pptx_path = Path(sys.argv[2])
        output_path = Path(sys.argv[3]) if len(sys.argv) > 3 else None
        extract_scripts_from_pptx(pptx_path, output_path)

    elif command == "list":
        if len(sys.argv) < 3:
            print("Usage: python pptx_scripts.py list <pptx>")
            sys.exit(1)
        list_slide_content(Path(sys.argv[2]))

    elif command == "template":
        if len(sys.argv) < 3:
            print("Usage: python pptx_scripts.py template <pptx>")
            sys.exit(1)
        pptx_path = Path(sys.argv[2])
        output_path = Path(sys.argv[3]) if len(sys.argv) > 3 else None
        create_script_template(pptx_path, output_path)

    elif command in ("--help", "-h", "help"):
        print_usage()

    else:
        print(f"Unknown command: {command}")
        print_usage()
        sys.exit(1)


if __name__ == "__main__":
    main()
