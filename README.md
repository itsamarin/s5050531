# English to Thai Translation Project

Professional translation from English to Thai using **Claude API** or **DeepL API**.

## Quick Start

```
INPUT                              OUTPUT
-----                              ------
English Presentation       ->      Thai Presentation
English Script             ->      Thai Script (as speaker notes)
```

---

## Two Translation Engines

| Feature | Claude API | DeepL API |
|---------|------------|-----------|
| **Best for** | Scripts, nuanced content | Presentations (.pptx) |
| **Format preservation** | Text only | Full document formatting |
| **PPTX support** | No (text only) | Yes (native) |
| **Review pass** | Yes (two-pass) | No |
| **Glossary** | Custom JSON | DeepL glossary |
| **Script** | `translate.py` | `translate_deepl.py` |

### Recommendation

- **Presentations (.pptx)**: Use **DeepL** - preserves slides, formatting, images
- **Scripts (.txt, .md)**: Use **Claude** - better context understanding, review pass

---

## Project Structure

```
.
├── scripts/
│   ├── english/          # English scripts
│   └── thai/             # Thai translations
├── presentations/
│   ├── english/          # English presentations (.pptx, .txt)
│   └── thai/             # Thai translations
├── translate.py          # Claude API translator
├── translate_deepl.py    # DeepL API translator
├── pptx_scripts.py       # Attach scripts to slides
├── quality_check.py      # Translation quality checker
├── glossary.json         # Terminology (for Claude)
└── requirements.txt
```

---

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure API Keys

```bash
cp .env.example .env
```

Edit `.env`:
```
ANTHROPIC_API_KEY=sk-ant-xxxxx    # For Claude
DEEPL_API_KEY=xxxxx               # For DeepL
```

Get your DeepL API key at: https://www.deepl.com/pro-api

---

## Step-by-Step Guide: Translate PowerPoint with Scripts

### Step 1: Prepare Your Files

**Option A: You have separate PowerPoint and Script files**

```
presentations/english/
├── my_presentation.pptx    <- Your English PowerPoint
└── my_script.txt           <- Your English script
```

**Option B: You only have PowerPoint (no script yet)**

```
presentations/english/
└── my_presentation.pptx    <- Your English PowerPoint
```

### Step 2: Create or Format Your Script

If you don't have a script yet, generate a template:

```bash
python pptx_scripts.py template presentations/english/my_presentation.pptx
```

This creates: `presentations/english/my_presentation_script_template.txt`

**Script Format:**

```
[Slide 1]
Welcome everyone to today's presentation.
I'm excited to share our quarterly results with you.

[Slide 2]
Let's start with the key highlights.
Our revenue grew by 25% compared to last year.

[Slide 3]
Here you can see the detailed breakdown.
The blue bars represent Q1, and the orange bars represent Q2.

[Slide 4]
In conclusion, we had an excellent quarter.
Thank you for your attention. Any questions?
```

### Step 3: Attach Script to PowerPoint

```bash
python pptx_scripts.py attach \
    presentations/english/my_presentation.pptx \
    presentations/english/my_script.txt
```

**Output:** `presentations/english/my_presentation_with_notes.pptx`

Verify the attachment:

```bash
python pptx_scripts.py list presentations/english/my_presentation_with_notes.pptx
```

### Step 4: Translate with DeepL

```bash
python translate_deepl.py presentations/english/my_presentation_with_notes.pptx
```

**Output:** `presentations/english/my_presentation_with_notes_thai.pptx`

DeepL translates:
- All slide content (text, titles, bullets)
- All speaker notes (your scripts)
- Preserves formatting, images, layouts

Move to Thai folder:

```bash
mv presentations/english/my_presentation_with_notes_thai.pptx \
   presentations/thai/my_presentation.pptx
```

### Step 5: Verify Translation

```bash
python pptx_scripts.py list presentations/thai/my_presentation.pptx
```

### Step 6: Present with Scripts

1. Open the Thai PowerPoint
2. Start slideshow: **F5** (or Slideshow -> From Beginning)
3. Enter Presenter View:
   - **Windows:** Alt + F5
   - **Mac:** Option + Return

**Presenter View:**

```
+------------------------------------------------+
|   +------------------+  +------------------+   |
|   |  Current Slide   |  |   Next Slide     |   |
|   |                  |  |   (preview)      |   |
|   +------------------+  +------------------+   |
|                                                |
|   +----------------------------------------+   |
|   |  Speaker Notes (Your Thai Script):     |   |
|   |  ...                                   |   |
|   +----------------------------------------+   |
|   [Timer: 00:05:32]        [Slide 1 of 10]     |
+------------------------------------------------+
        ^ Only YOU see this
        v Audience sees only the slide
```

---

## Complete Example

```bash
# 1. Setup (first time only)
pip install -r requirements.txt
cp .env.example .env
# Add your DEEPL_API_KEY to .env

# 2. Add your PowerPoint
cp ~/Downloads/quarterly_report.pptx presentations/english/

# 3. Create script template
python pptx_scripts.py template presentations/english/quarterly_report.pptx

# 4. Edit the script template (add your speaking notes)
# Open: presentations/english/quarterly_report_script_template.txt

# 5. Attach script to PowerPoint
python pptx_scripts.py attach \
    presentations/english/quarterly_report.pptx \
    presentations/english/quarterly_report_script_template.txt

# 6. Translate to Thai
python translate_deepl.py presentations/english/quarterly_report_with_notes.pptx

# 7. Move to Thai folder
mv presentations/english/quarterly_report_with_notes_thai.pptx \
   presentations/thai/quarterly_report.pptx

# 8. Verify
python pptx_scripts.py list presentations/thai/quarterly_report.pptx

# Done! Open presentations/thai/quarterly_report.pptx and present!
```

---

## Command Reference

### PowerPoint Script Manager (`pptx_scripts.py`)

| Command | Description |
|---------|-------------|
| `python pptx_scripts.py template <pptx>` | Create script template from slides |
| `python pptx_scripts.py attach <pptx> <script>` | Attach script as speaker notes |
| `python pptx_scripts.py list <pptx>` | View slides + notes |
| `python pptx_scripts.py extract <pptx>` | Extract notes to file |

### DeepL Translator (`translate_deepl.py`)

| Command | Description |
|---------|-------------|
| `python translate_deepl.py presentations` | Translate all presentations |
| `python translate_deepl.py scripts` | Translate all scripts |
| `python translate_deepl.py all` | Translate everything |
| `python translate_deepl.py <file>` | Translate single file |
| `python translate_deepl.py usage` | Check API usage |
| `--formal` | Use formal Thai |
| `--informal` | Use informal Thai |

### Claude Translator (`translate.py`)

| Command | Description |
|---------|-------------|
| `python translate.py scripts` | Translate all scripts (with review) |
| `python translate.py presentations` | Translate presentations (text only) |
| `python translate.py all` | Translate everything |
| `python translate.py interactive` | Test mode |
| `--no-review` | Skip review pass (faster) |

### Quality Check (`quality_check.py`)

| Command | Description |
|---------|-------------|
| `python quality_check.py <translated.pptx>` | Check translation quality |
| `python quality_check.py <translated.pptx> <original.pptx>` | Compare with original |
| `--html` | Save HTML report only |
| `--text` | Save text report only |

---

## Quality Check Process

After translation, run the quality check to identify potential issues.
The translated slides remain **untouched** - issues are flagged in a separate report.

```bash
python quality_check.py my_presentation_thai.pptx my_presentation.pptx
```

### Output Files

| File | Description |
|------|-------------|
| `*_thai.pptx` | Clean translated slides (unchanged) |
| `*_quality_report.html` | HTML report - open in browser |
| `*_quality_report.txt` | Text report - for terminal/logs |

### What It Checks

| Severity | Issue Type | Description |
|----------|------------|-------------|
| Critical | Character Encoding | Garbled or corrupted text |
| Warning | Untranslated Text | English words that should be Thai |
| Warning | Missing Numbers | Numbers lost in translation |
| Warning | Length Difference | Text significantly longer/shorter |
| Info | Empty Content | Empty slides or notes |
| Info | Formatting | Unmatched brackets/parentheses |

### Sample Output

```
======================================================================
TRANSLATION QUALITY REPORT
======================================================================

File: presentation_thai.pptx
Date: 2026-01-25

Quality Score: 85/100 (Good)

----------------------------------------------------------------------
SUMMARY
----------------------------------------------------------------------
Total Slides: 10
Total Issues: 5
  Warnings: 3
  Info: 2

----------------------------------------------------------------------
ISSUES BY SLIDE
----------------------------------------------------------------------

+-- SLIDE 3 (1 issue)
|
|  [Warning] Untranslated Text
|     Location: Slide Content
|     Found: Customer, Revenue, Growth
|     -> Review if these English words should be translated to Thai
+--------------------------------------------------

----------------------------------------------------------------------
SLIDES TO REVIEW
----------------------------------------------------------------------

Review these slides: 3, 7

======================================================================
NOTE: The translated presentation file is NOT modified.
This report is for review purposes only.
======================================================================
```

### Complete Workflow with Quality Check

```bash
# 1. Translate
python translate_deepl.py my_presentation.pptx

# 2. Quality Check (generates separate report)
python quality_check.py my_presentation_thai.pptx my_presentation.pptx

# 3. Open report and presentation side by side to review

# 4. Present!
```

---

## DeepL Supported Formats

| Format | Description |
|--------|-------------|
| `.pptx` | PowerPoint (formatting preserved) |
| `.docx` | Word documents |
| `.pdf` | PDF files |
| `.xlsx` | Excel files |
| `.txt` | Plain text |

---

## Tips

### For Better Translation Quality

1. **Use formal Thai**: Add `--formal` flag
   ```bash
   python translate_deepl.py my_slides.pptx --formal
   ```

2. **Keep sentences short** in your script for natural Thai flow

3. **Technical terms**: DeepL handles most technical terms well, but review specialized vocabulary

### Troubleshooting

| Issue | Solution |
|-------|----------|
| Script not appearing | Run `pptx_scripts.py list` to verify attachment |
| Wrong slide numbers | Check script format uses `[Slide 1]`, `[Slide 2]`, etc. |
| Translation missing notes | Ensure you translated the `_with_notes.pptx` file |
| Presenter View not showing | Press Alt+F5 (Windows) or check Display Settings |

---

## API Pricing Notes

### DeepL
- **Free tier**: 500,000 characters/month
- **Pro**: Pay per character
- **Document minimum**: 50,000 characters per .pptx/.docx/.pdf

### Claude
- Pay per token (input + output)
- Review pass doubles API calls but improves quality

---

## Sources

- [DeepL API Documentation](https://developers.deepl.com/docs)
- [DeepL PowerPoint Translation](https://www.deepl.com/en/features/document-translation/ppt)
- [DeepL Python Library](https://github.com/DeepLcom/deepl-python)
- [Anthropic Claude API](https://docs.anthropic.com/)
- [python-pptx Notes Documentation](https://python-pptx.readthedocs.io/en/latest/user/notes.html)
