#!/usr/bin/env python3
"""
English to Thai Translation Script using Claude API
Enhanced with advanced LLM techniques for best translation quality
"""

import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv
import anthropic

# Load environment variables
load_dotenv()

# Initialize Anthropic client
client = anthropic.Anthropic()

# Model to use for translation
MODEL = "claude-sonnet-4-20250514"

# Load glossary if exists
GLOSSARY_PATH = Path("glossary.json")


def load_glossary() -> dict:
    """Load translation glossary for consistent terminology."""
    if GLOSSARY_PATH.exists():
        with open(GLOSSARY_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


GLOSSARY = load_glossary()


def get_glossary_text() -> str:
    """Format glossary for prompt inclusion."""
    if not GLOSSARY:
        return ""

    lines = ["## Glossary - Use these specific translations:"]
    for eng, thai in GLOSSARY.items():
        lines.append(f"- {eng} → {thai}")
    return "\n".join(lines)


# System prompt for expert translation
SYSTEM_PROMPT = """You are an expert English to Thai translator with deep knowledge of both languages and cultures. You have years of experience translating professional content including scripts, presentations, and technical documents.

Your translation philosophy:
1. **Accuracy First**: Capture the exact meaning and intent of the original
2. **Natural Flow**: Thai output should read naturally, not like a translation
3. **Cultural Adaptation**: Adapt idioms and expressions appropriately for Thai audience
4. **Consistency**: Maintain consistent terminology throughout
5. **Preserve Structure**: Keep formatting, bullet points, numbering intact

Your expertise includes:
- Understanding context and nuance
- Choosing appropriate register (formal/informal) based on content
- Handling technical terms appropriately
- Maintaining the tone and style of the original"""


def translate_with_context(
    text: str,
    content_type: str = "general",
    additional_context: str = ""
) -> str:
    """
    Translate text using advanced prompting for best quality.

    Args:
        text: English text to translate
        content_type: Type of content (script, presentation, general)
        additional_context: Any additional context about the content
    """

    glossary_text = get_glossary_text()

    # Content-specific instructions
    content_instructions = {
        "script": """This is a SCRIPT for video/audio content.
- Maintain natural speaking rhythm in Thai
- Keep sentences concise for easy reading/speaking
- Preserve timing cues and speaker labels
- Adapt expressions to sound natural when spoken aloud""",

        "presentation": """This is a PRESENTATION (slides/deck).
- Keep text concise and impactful
- Maintain bullet point structure
- Preserve slide titles and hierarchy
- Ensure text fits typical slide constraints
- Keep key terms prominent""",

        "general": """This is general content.
- Maintain the original structure and formatting
- Preserve all important information"""
    }

    instruction = content_instructions.get(content_type, content_instructions["general"])

    user_prompt = f"""## Task
Translate the following English text to Thai with professional quality.

## Content Type Instructions
{instruction}

{glossary_text}

{f"## Additional Context: {additional_context}" if additional_context else ""}

## English Text to Translate:
---
{text}
---

## Instructions:
1. First, analyze the text to understand its purpose, tone, and key messages
2. Translate to Thai maintaining meaning, tone, and structure
3. Ensure the Thai reads naturally and fluently
4. Preserve all formatting (headers, bullets, numbers, etc.)

Provide ONLY the Thai translation, no explanations or notes."""

    message = client.messages.create(
        model=MODEL,
        max_tokens=8192,
        system=SYSTEM_PROMPT,
        messages=[
            {"role": "user", "content": user_prompt}
        ]
    )

    return message.content[0].text


def review_translation(english: str, thai: str, content_type: str = "general") -> str:
    """
    Review and improve a translation using a second LLM pass.
    """

    review_prompt = f"""## Task
You are a senior Thai language editor reviewing a translation. Your job is to improve the Thai translation for quality, accuracy, and naturalness.

## Original English:
---
{english}
---

## Current Thai Translation:
---
{thai}
---

## Review Checklist:
1. **Accuracy**: Does the Thai accurately convey the English meaning?
2. **Naturalness**: Does it read like native Thai, not translated text?
3. **Terminology**: Are technical terms handled appropriately?
4. **Tone**: Does it match the original's tone and register?
5. **Completeness**: Is anything missing or added incorrectly?
6. **Grammar**: Is Thai grammar correct?
7. **Flow**: Does the text flow well and connect naturally?

## Instructions:
- If the translation is already excellent, return it unchanged
- If improvements are needed, provide the improved version
- Make only necessary changes, don't over-edit
- Return ONLY the final Thai text, no explanations

## Final Thai Translation:"""

    message = client.messages.create(
        model=MODEL,
        max_tokens=8192,
        system="You are a senior Thai language editor with expertise in translation quality assurance. You refine translations to be accurate, natural, and professional.",
        messages=[
            {"role": "user", "content": review_prompt}
        ]
    )

    return message.content[0].text


def translate_text(
    english_text: str,
    content_type: str = "general",
    use_review: bool = True,
    context: str = ""
) -> str:
    """
    Full translation pipeline with optional review pass.

    Args:
        english_text: Text to translate
        content_type: script, presentation, or general
        use_review: Whether to run a review/improvement pass
        context: Additional context about the content
    """

    # First pass: Initial translation
    thai_translation = translate_with_context(english_text, content_type, context)

    # Second pass: Review and improve (if enabled)
    if use_review:
        thai_translation = review_translation(english_text, thai_translation, content_type)

    return thai_translation


def translate_file(
    input_path: Path,
    output_path: Path,
    content_type: str = "general",
    use_review: bool = True
) -> bool:
    """Translate a single file from English to Thai."""
    try:
        print(f"\n📄 Translating: {input_path.name}")
        print(f"   Type: {content_type}")
        print(f"   Review pass: {'enabled' if use_review else 'disabled'}")

        # Read English content
        with open(input_path, 'r', encoding='utf-8') as f:
            english_content = f.read()

        if not english_content.strip():
            print(f"   ⚠️  Skipping empty file")
            return False

        # Check file size and warn if large
        word_count = len(english_content.split())
        print(f"   Words: ~{word_count}")

        # Translate to Thai
        print(f"   🔄 Translating...")
        thai_content = translate_text(
            english_content,
            content_type=content_type,
            use_review=use_review
        )

        # Save Thai translation
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(thai_content)

        print(f"   ✅ Saved to: {output_path}")
        return True

    except anthropic.APIError as e:
        print(f"   ❌ API Error: {e}")
        return False
    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def translate_directory(
    english_dir: Path,
    thai_dir: Path,
    content_type: str,
    use_review: bool = True
):
    """Translate all files in a directory."""

    if not english_dir.exists():
        print(f"Directory not found: {english_dir}")
        return

    # Find all translatable files
    files = list(english_dir.glob("*.txt")) + list(english_dir.glob("*.md"))
    files = [f for f in files if not f.name.startswith('.')]

    if not files:
        print(f"No files found in {english_dir}")
        return

    print(f"\n{'='*50}")
    print(f"Translating {len(files)} {content_type}(s)")
    print(f"From: {english_dir}")
    print(f"To:   {thai_dir}")
    print(f"{'='*50}")

    success_count = 0
    for file_path in sorted(files):
        output_path = thai_dir / file_path.name
        if translate_file(file_path, output_path, content_type, use_review):
            success_count += 1

    print(f"\n{'='*50}")
    print(f"Completed: {success_count}/{len(files)} files translated")
    print(f"{'='*50}")


def translate_scripts(use_review: bool = True):
    """Translate all scripts."""
    translate_directory(
        Path("scripts/english"),
        Path("scripts/thai"),
        "script",
        use_review
    )


def translate_presentations(use_review: bool = True):
    """Translate all presentations."""
    translate_directory(
        Path("presentations/english"),
        Path("presentations/thai"),
        "presentation",
        use_review
    )


def interactive_translate():
    """Interactive translation mode for testing."""
    print("\n🌐 Interactive English → Thai Translation")
    print("Type 'quit' to exit\n")

    while True:
        print("-" * 40)
        text = input("English: ").strip()

        if text.lower() == 'quit':
            break

        if not text:
            continue

        print("\nTranslating...")
        thai = translate_text(text, use_review=True)
        print(f"\nThai: {thai}\n")


def print_usage():
    """Print usage instructions."""
    print("""
🌐 English to Thai Translation Tool (Claude API)
================================================

Usage:
  python translate.py <command> [options]

Commands:
  scripts          Translate all scripts in scripts/english/
  presentations    Translate all presentations in presentations/english/
  all              Translate everything
  interactive      Interactive translation mode (for testing)
  <file>           Translate a single file

Options:
  --no-review      Skip the review/improvement pass (faster but lower quality)
  --type=TYPE      Content type: script, presentation, general (default: general)

Examples:
  python translate.py scripts
  python translate.py presentations --no-review
  python translate.py all
  python translate.py myfile.txt --type=script
  python translate.py interactive

Setup:
  1. pip install -r requirements.txt
  2. cp .env.example .env
  3. Add your ANTHROPIC_API_KEY to .env
  4. (Optional) Create glossary.json for consistent terminology
""")


def main():
    """Main entry point."""
    if not os.getenv("ANTHROPIC_API_KEY"):
        print("❌ Error: ANTHROPIC_API_KEY not set")
        print("Please set your API key in a .env file or environment variable")
        sys.exit(1)

    if len(sys.argv) < 2:
        print_usage()
        sys.exit(1)

    # Parse arguments
    args = sys.argv[1:]
    command = args[0]

    use_review = "--no-review" not in args
    content_type = "general"

    for arg in args:
        if arg.startswith("--type="):
            content_type = arg.split("=")[1]

    # Execute command
    if command == "scripts":
        translate_scripts(use_review)
    elif command == "presentations":
        translate_presentations(use_review)
    elif command == "all":
        translate_scripts(use_review)
        translate_presentations(use_review)
    elif command == "interactive":
        interactive_translate()
    elif command in ("--help", "-h", "help"):
        print_usage()
    else:
        # Assume it's a file path
        input_path = Path(command)
        if not input_path.exists():
            print(f"❌ File not found: {command}")
            sys.exit(1)

        output_path = input_path.parent / f"{input_path.stem}_thai{input_path.suffix}"
        translate_file(input_path, output_path, content_type, use_review)


if __name__ == "__main__":
    main()
