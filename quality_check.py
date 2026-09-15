#!/usr/bin/env python3
"""
Translation Quality Check Tool — powered by GPT-5
Analyzes translated files (.pptx, .docx, .xlsx) and creates a SEPARATE quality report.
The translated file is never modified — issues are flagged in the report only.
"""

import os
import sys
import re
import json
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MODEL = "gpt-5"
BATCH_SIZE = 30          # translation pairs per GPT-5 call
SUPPORTED_EXTS = {".pptx", ".docx", ".xlsx"}

try:
    from openai import OpenAI
    _client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
except ImportError:
    _client = None

# Fallback regex patterns (used when no API key)
_THAI_RE = re.compile(r"[฀-๿]")
_ENGLISH_WORD_RE = re.compile(r"\b[A-Za-z]{4,}\b")
_NUMBER_RE = re.compile(r"\b\d+(?:[.,]\d+)*%?\b")

_ALWAYS_ENGLISH = {
    "ai", "api", "iot", "ceo", "cto", "cfo", "coo", "dna", "gps", "usb",
    "pdf", "url", "html", "css", "javascript", "python", "java", "react",
    "angular", "vue", "blockchain", "cloud", "aws", "azure", "google",
    "microsoft", "apple", "meta", "facebook", "linkedin", "twitter",
    "iphone", "android", "ios", "windows", "linux", "excel", "powerpoint",
    "word", "outlook", "teams", "wifi", "bluetooth", "sql", "nosql",
    "docker", "kubernetes", "devops", "agile", "scrum", "roi", "kpi",
    "saas", "paas", "iaas", "crm", "erp", "usd", "eur", "thb",
}


# ── Data model ────────────────────────────────────────────────────────────────
class QualityIssue:
    def __init__(self, item_num: int, item_label: str, issue_type: str,
                 location: str, original: str, translated: str,
                 suggestion: str, severity: str):
        self.item_num = item_num        # slide / paragraph / row number
        self.item_label = item_label    # "Slide" / "Paragraph" / "Row"
        self.issue_type = issue_type
        self.location = location
        self.original = original
        self.translated = translated
        self.suggestion = suggestion
        self.severity = severity        # "critical" | "warning" | "info"

    # Backwards compat alias used by report generators
    @property
    def slide_num(self):
        return self.item_num


# ── GPT-5 quality analysis ────────────────────────────────────────────────────
def _gpt5_available() -> bool:
    return bool(OPENAI_API_KEY and _client)


def _analyze_batch_gpt5(pairs: list[dict]) -> list[dict]:
    """
    Send a batch of {index, english, thai} pairs to GPT-5 for quality review.
    Returns a list of issue dicts (only for pairs that have problems).
    """
    prompt = (
        "You are a professional English-to-Thai translation quality assessor.\n\n"
        "Evaluate each translation pair below. For pairs with good quality, return "
        "{\"index\": N, \"ok\": true}. For pairs with issues return:\n"
        "{\n"
        "  \"index\": N,\n"
        "  \"ok\": false,\n"
        "  \"severity\": \"critical\" | \"warning\" | \"info\",\n"
        "  \"issue_type\": \"Semantic Error\" | \"Untranslated Text\" | \"Awkward Phrasing\" "
        "| \"Missing Content\" | \"Number Mismatch\" | \"Formatting Issue\" | \"Other\",\n"
        "  \"issue\": \"<one sentence describing the specific problem>\",\n"
        "  \"suggestion\": \"<one sentence on how to fix it>\"\n"
        "}\n\n"
        "Rules:\n"
        "- Brand names, acronyms, and proper nouns kept in English are acceptable.\n"
        "- Numbers, dates, percentages must match the original exactly.\n"
        "- Flag if Thai text appears non-fluent, overly literal, or unnatural.\n"
        "- If 'english' is empty or already Thai, return ok: true.\n\n"
        "Return ONLY a JSON array — no markdown, no explanation.\n\n"
        f"Pairs:\n{json.dumps(pairs, ensure_ascii=False)}"
    )

    response = _client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Return only valid JSON arrays."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
    )

    raw = response.choices[0].message.content.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)

    parsed = json.loads(raw)
    if isinstance(parsed, dict):
        parsed = next(iter(parsed.values()))
    return parsed


def analyze_pairs_gpt5(pairs: list[dict], item_label: str) -> list[QualityIssue]:
    """Run GPT-5 quality check across all pairs, in batches."""
    issues: list[QualityIssue] = []

    for i in range(0, len(pairs), BATCH_SIZE):
        chunk = pairs[i: i + BATCH_SIZE]
        print(f"   GPT-5 batch {i // BATCH_SIZE + 1}/{(len(pairs) - 1) // BATCH_SIZE + 1} "
              f"({len(chunk)} pairs)...")

        try:
            results = _analyze_batch_gpt5(chunk)
        except Exception as e:
            print(f"   Warning: GPT-5 batch failed — {e}")
            continue

        for result in results:
            if result.get("ok"):
                continue
            idx = result.get("index", 0)
            source_pair = chunk[idx] if idx < len(chunk) else {}

            issues.append(QualityIssue(
                item_num=source_pair.get("item_num", idx + 1),
                item_label=item_label,
                issue_type=result.get("issue_type", "Other"),
                location=source_pair.get("location", "Content"),
                original=result.get("issue", ""),
                translated="",
                suggestion=result.get("suggestion", ""),
                severity=result.get("severity", "warning"),
            ))

    return issues


# ── Fallback rule-based analysis ──────────────────────────────────────────────
def _rule_based_issues(item_num: int, item_label: str, location: str,
                        english: str, thai: str) -> list[QualityIssue]:
    issues = []

    if not thai.strip():
        issues.append(QualityIssue(
            item_num, item_label, "Missing Translation", location,
            english[:80], "", "This text appears untranslated", "warning"
        ))
        return issues

    # Untranslated English words (only when Thai chars are present)
    if _THAI_RE.search(thai):
        words = _ENGLISH_WORD_RE.findall(thai)
        flagged = [w for w in words if w.lower() not in _ALWAYS_ENGLISH]
        if flagged:
            issues.append(QualityIssue(
                item_num, item_label, "Untranslated Text", location,
                ", ".join(set(flagged[:5])), "",
                "Review if these English words should be translated to Thai",
                "warning"
            ))

    # Missing numbers
    if english:
        orig_nums = set(_NUMBER_RE.findall(english))
        thai_nums = set(_NUMBER_RE.findall(thai))
        missing = orig_nums - thai_nums
        if missing:
            issues.append(QualityIssue(
                item_num, item_label, "Number Mismatch", location,
                ", ".join(missing), "",
                "Verify these numbers appear correctly in the translation",
                "warning"
            ))

    # Length ratio (only when comparing against English)
    if english and len(english) > 20:
        ratio = len(thai) / len(english)
        if ratio > 2.0:
            issues.append(QualityIssue(
                item_num, item_label, "Unusual Length", location,
                f"{len(english)} chars (EN)", f"{len(thai)} chars (TH, {int((ratio-1)*100)}% longer)",
                "Text may be significantly longer than expected",
                "info"
            ))

    return issues


# ── PPTX extraction ───────────────────────────────────────────────────────────
def _pptx_pairs(translated_prs, original_prs) -> list[dict]:
    pairs = []
    idx = 0

    for slide_i, slide in enumerate(translated_prs.slides, 1):
        orig_slide = None
        if original_prs and slide_i <= len(original_prs.slides):
            orig_slide = original_prs.slides[slide_i - 1]

        def _orig_shape_text(shape_idx):
            if orig_slide is None:
                return ""
            orig_shapes = [s for s in orig_slide.shapes if s.has_text_frame]
            if shape_idx < len(orig_shapes):
                return orig_shapes[shape_idx].text_frame.text
            return ""

        content_shapes = [s for s in slide.shapes if s.has_text_frame]
        for shape_idx, shape in enumerate(content_shapes):
            thai_text = shape.text_frame.text.strip()
            if not thai_text:
                continue
            pairs.append({
                "index": idx,
                "item_num": slide_i,
                "location": "Slide Content",
                "english": _orig_shape_text(shape_idx).strip(),
                "thai": thai_text,
            })
            idx += 1

        # Speaker notes
        thai_notes = ""
        orig_notes = ""
        try:
            thai_notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            pass
        if orig_slide:
            try:
                orig_notes = orig_slide.notes_slide.notes_text_frame.text.strip()
            except Exception:
                pass
        if thai_notes:
            pairs.append({
                "index": idx,
                "item_num": slide_i,
                "location": "Speaker Notes",
                "english": orig_notes,
                "thai": thai_notes,
            })
            idx += 1

    return pairs


def analyze_pptx(translated_path: Path, original_path: Path | None) -> dict:
    from pptx import Presentation

    translated_prs = Presentation(str(translated_path))
    original_prs = Presentation(str(original_path)) if (original_path and original_path.exists()) else None

    pairs = _pptx_pairs(translated_prs, original_prs)
    print(f"   Found {len(pairs)} text elements to check.")

    if _gpt5_available():
        print("   Engine: GPT-5")
        issues = analyze_pairs_gpt5(pairs, "Slide")
    else:
        print("   Engine: Rule-based (set OPENAI_API_KEY for GPT-5 analysis)")
        issues = []
        for p in pairs:
            issues.extend(_rule_based_issues(
                p["item_num"], "Slide", p["location"], p["english"], p["thai"]
            ))

    return {
        "filename": translated_path.name,
        "original": original_path.name if original_path else None,
        "timestamp": datetime.now().isoformat(),
        "file_type": "pptx",
        "item_label": "Slide",
        "item_count": len(translated_prs.slides),
        "total_issues": len(issues),
        "issues": issues,
    }


# ── DOCX extraction ───────────────────────────────────────────────────────────
def _docx_paragraphs(doc) -> list[str]:
    texts = []
    for para in doc.paragraphs:
        if para.text.strip():
            texts.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        texts.append(para.text.strip())
    return texts


def analyze_docx(translated_path: Path, original_path: Path | None) -> dict:
    from docx import Document

    translated_doc = Document(str(translated_path))
    original_doc = Document(str(original_path)) if (original_path and original_path.exists()) else None

    thai_texts = _docx_paragraphs(translated_doc)
    orig_texts = _docx_paragraphs(original_doc) if original_doc else []

    pairs = []
    for i, thai in enumerate(thai_texts):
        pairs.append({
            "index": i,
            "item_num": i + 1,
            "location": "Paragraph",
            "english": orig_texts[i] if i < len(orig_texts) else "",
            "thai": thai,
        })

    print(f"   Found {len(pairs)} paragraphs to check.")

    if _gpt5_available():
        print("   Engine: GPT-5")
        issues = analyze_pairs_gpt5(pairs, "Paragraph")
    else:
        print("   Engine: Rule-based (set OPENAI_API_KEY for GPT-5 analysis)")
        issues = []
        for p in pairs:
            issues.extend(_rule_based_issues(
                p["item_num"], "Paragraph", p["location"], p["english"], p["thai"]
            ))

    return {
        "filename": translated_path.name,
        "original": original_path.name if original_path else None,
        "timestamp": datetime.now().isoformat(),
        "file_type": "docx",
        "item_label": "Paragraph",
        "item_count": len(thai_texts),
        "total_issues": len(issues),
        "issues": issues,
    }


# ── XLSX extraction ───────────────────────────────────────────────────────────
def analyze_xlsx(translated_path: Path, original_path: Path | None) -> dict:
    import openpyxl

    translated_wb = openpyxl.load_workbook(str(translated_path))
    original_wb = openpyxl.load_workbook(str(original_path)) if (original_path and original_path.exists()) else None

    pairs = []
    idx = 0

    for ws in translated_wb.worksheets:
        orig_ws = original_wb[ws.title] if (original_wb and ws.title in original_wb.sheetnames) else None

        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "s" or not cell.value:
                    continue
                thai = str(cell.value).strip()
                if not thai:
                    continue

                english = ""
                if orig_ws:
                    orig_cell = orig_ws.cell(row=cell.row, column=cell.column)
                    if orig_cell.value:
                        english = str(orig_cell.value).strip()

                pairs.append({
                    "index": idx,
                    "item_num": cell.row,
                    "location": f"Sheet '{ws.title}' cell {cell.coordinate}",
                    "english": english,
                    "thai": thai,
                })
                idx += 1

    print(f"   Found {len(pairs)} cells to check.")

    if _gpt5_available():
        print("   Engine: GPT-5")
        issues = analyze_pairs_gpt5(pairs, "Row")
    else:
        print("   Engine: Rule-based (set OPENAI_API_KEY for GPT-5 analysis)")
        issues = []
        for p in pairs:
            issues.extend(_rule_based_issues(
                p["item_num"], "Row", p["location"], p["english"], p["thai"]
            ))

    return {
        "filename": translated_path.name,
        "original": original_path.name if original_path else None,
        "timestamp": datetime.now().isoformat(),
        "file_type": "xlsx",
        "item_label": "Row",
        "item_count": len(pairs),
        "total_issues": len(issues),
        "issues": issues,
    }


# ── Report generation ─────────────────────────────────────────────────────────
def _score(issues: list[QualityIssue]) -> tuple[int, str]:
    critical = sum(1 for i in issues if i.severity == "critical")
    warnings = sum(1 for i in issues if i.severity == "warning")
    info = sum(1 for i in issues if i.severity == "info")
    s = max(0, 100 - critical * 20 - warnings * 5 - info * 1)
    label = "Excellent" if s >= 90 else "Good" if s >= 70 else "Fair" if s >= 50 else "Needs Review"
    return s, label


def generate_report_text(analysis: dict) -> str:
    issues = analysis["issues"]
    critical = [i for i in issues if i.severity == "critical"]
    warnings = [i for i in issues if i.severity == "warning"]
    info = [i for i in issues if i.severity == "info"]
    score, score_label = _score(issues)
    item_label = analysis.get("item_label", "Slide")
    item_count = analysis.get("item_count", 0)
    engine = "GPT-5" if _gpt5_available() else "Rule-based"

    lines = []
    lines.append("=" * 70)
    lines.append("TRANSLATION QUALITY REPORT")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"File:     {analysis['filename']}")
    if analysis["original"]:
        lines.append(f"Original: {analysis['original']}")
    lines.append(f"Date:     {analysis['timestamp'][:10]}")
    lines.append(f"Engine:   {engine}")
    lines.append("")
    lines.append(f"Quality Score: {score}/100 ({score_label})")
    lines.append("")
    lines.append("-" * 70)
    lines.append("SUMMARY")
    lines.append("-" * 70)
    lines.append(f"Total {item_label}s: {item_count}")
    lines.append(f"Total Issues: {len(issues)}")
    if critical:
        lines.append(f"  Critical: {len(critical)}")
    if warnings:
        lines.append(f"  Warnings: {len(warnings)}")
    if info:
        lines.append(f"  Info:     {len(info)}")
    lines.append("")

    by_item = defaultdict(list)
    for issue in issues:
        by_item[issue.item_num].append(issue)

    if issues:
        lines.append("-" * 70)
        lines.append(f"ISSUES BY {item_label.upper()}")
        lines.append("-" * 70)
        lines.append("")

        for item_num in sorted(by_item.keys()):
            item_issues = by_item[item_num]
            lines.append(f"+-- {item_label} {item_num} ({len(item_issues)} issue{'s' if len(item_issues) > 1 else ''})")
            for issue in item_issues:
                sev_tag = {"critical": "[CRITICAL]", "warning": "[Warning]", "info": "[Info]"}[issue.severity]
                lines.append(f"|")
                lines.append(f"|  {sev_tag} {issue.issue_type}")
                lines.append(f"|     Location: {issue.location}")
                if issue.original:
                    lines.append(f"|     Detail:   {issue.original}")
                if issue.translated:
                    lines.append(f"|     Value:    {issue.translated}")
                lines.append(f"|     Fix:      {issue.suggestion}")
            lines.append(f"+{'─' * 50}")
            lines.append("")

        problem_items = sorted(by_item.keys())
        lines.append("-" * 70)
        lines.append(f"{item_label}S TO REVIEW")
        lines.append("-" * 70)
        lines.append("")
        lines.append(f"Review these {item_label.lower()}s: {', '.join(map(str, problem_items))}")
        lines.append("")

    lines.append("=" * 70)
    lines.append("NOTE: The translated file is NOT modified.")
    lines.append("This report is for review purposes only.")
    lines.append("=" * 70)

    return "\n".join(lines)


def generate_report_html(analysis: dict) -> str:
    issues = analysis["issues"]
    critical = [i for i in issues if i.severity == "critical"]
    warnings = [i for i in issues if i.severity == "warning"]
    info_items = [i for i in issues if i.severity == "info"]
    score, score_label = _score(issues)
    item_label = analysis.get("item_label", "Slide")
    item_count = analysis.get("item_count", 0)
    engine = "GPT-5" if _gpt5_available() else "Rule-based"

    score_color = "#22c55e" if score >= 90 else "#eab308" if score >= 70 else "#f97316" if score >= 50 else "#ef4444"

    by_item = defaultdict(list)
    for issue in issues:
        by_item[issue.item_num].append(issue)

    issues_html = ""
    if not issues:
        issues_html = """
            <div class="no-issues">
                <div style="font-size:48px;margin-bottom:16px">✅</div>
                <div style="font-size:18px;font-weight:600">No issues found!</div>
                <div style="color:#666;margin-top:8px">Your translation looks great.</div>
            </div>"""
    else:
        for item_num in sorted(by_item.keys()):
            item_issues = by_item[item_num]
            count = len(item_issues)
            issues_html += f"""
            <div class="slide-group">
                <div class="slide-header">{item_label} {item_num} ({count} issue{'s' if count > 1 else ''})</div>"""
            for issue in item_issues:
                found_html = f"<div class='issue-details'><strong>Detail:</strong> {issue.original}</div>" if issue.original else ""
                trans_html = f"<div class='issue-details'><strong>Value:</strong> {issue.translated}</div>" if issue.translated else ""
                issues_html += f"""
                <div class="issue">
                    <div class="issue-type">
                        <span class="severity {issue.severity}"></span>
                        <strong>{issue.issue_type}</strong>
                        <span style="color:#999;font-size:12px">({issue.location})</span>
                    </div>
                    {found_html}{trans_html}
                    <div class="suggestion">Fix: {issue.suggestion}</div>
                </div>"""
            issues_html += "\n            </div>"

    original_line = f"<strong>Original:</strong> {analysis['original']}<br>" if analysis["original"] else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Quality Report — {analysis['filename']}</title>
    <style>
        *{{margin:0;padding:0;box-sizing:border-box}}
        body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f5f5f5;padding:20px}}
        .container{{max-width:900px;margin:0 auto}}
        .card{{background:white;border-radius:12px;padding:24px;margin-bottom:20px;box-shadow:0 2px 4px rgba(0,0,0,.1)}}
        h1{{font-size:24px;margin-bottom:8px}}
        .meta{{color:#666;font-size:14px}}
        .score-box{{display:inline-flex;align-items:center;gap:12px;background:{score_color}15;border:2px solid {score_color};border-radius:8px;padding:12px 20px;margin-top:16px}}
        .score-number{{font-size:32px;font-weight:bold;color:{score_color}}}
        .score-label{{font-size:16px;color:{score_color}}}
        h2{{font-size:18px;margin-bottom:16px}}
        .stats{{display:flex;gap:20px;flex-wrap:wrap}}
        .stat{{background:#f8f9fa;padding:12px 20px;border-radius:8px}}
        .stat-value{{font-size:24px;font-weight:bold}}
        .stat-label{{font-size:12px;color:#666;text-transform:uppercase}}
        .slide-group{{border:1px solid #e5e7eb;border-radius:8px;margin-bottom:16px;overflow:hidden}}
        .slide-header{{background:#f8f9fa;padding:12px 16px;font-weight:600;border-bottom:1px solid #e5e7eb}}
        .issue{{padding:12px 16px;border-bottom:1px solid #f0f0f0}}
        .issue:last-child{{border-bottom:none}}
        .issue-type{{display:flex;align-items:center;gap:8px;margin-bottom:8px}}
        .severity{{width:10px;height:10px;border-radius:50%}}
        .severity.critical{{background:#ef4444}}
        .severity.warning{{background:#eab308}}
        .severity.info{{background:#3b82f6}}
        .issue-details{{font-size:14px;color:#666;margin-left:18px;margin-bottom:4px}}
        .suggestion{{background:#fef3c7;padding:8px 12px;border-radius:4px;margin-top:8px;font-size:13px;margin-left:18px}}
        .no-issues{{text-align:center;padding:40px;color:#22c55e}}
        .note{{background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:16px;margin-top:20px;font-size:14px}}
        .engine-badge{{display:inline-block;background:#6366f1;color:white;font-size:11px;padding:2px 8px;border-radius:999px;margin-left:8px;vertical-align:middle}}
    </style>
</head>
<body>
<div class="container">
    <div class="card">
        <h1>Translation Quality Report <span class="engine-badge">{engine}</span></h1>
        <div class="meta">
            <strong>File:</strong> {analysis['filename']}<br>
            {original_line}
            <strong>Date:</strong> {analysis['timestamp'][:10]}
        </div>
        <div class="score-box">
            <span class="score-number">{score}</span>
            <span class="score-label">/100<br>{score_label}</span>
        </div>
    </div>
    <div class="card">
        <h2>Summary</h2>
        <div class="stats">
            <div class="stat"><div class="stat-value">{item_count}</div><div class="stat-label">{item_label}s</div></div>
            <div class="stat"><div class="stat-value">{len(issues)}</div><div class="stat-label">Total Issues</div></div>
            <div class="stat"><div class="stat-value" style="color:#ef4444">{len(critical)}</div><div class="stat-label">Critical</div></div>
            <div class="stat"><div class="stat-value" style="color:#eab308">{len(warnings)}</div><div class="stat-label">Warnings</div></div>
            <div class="stat"><div class="stat-value" style="color:#3b82f6">{len(info_items)}</div><div class="stat-label">Info</div></div>
        </div>
    </div>
    <div class="card">
        <h2>Issues by {item_label}</h2>
        {issues_html}
        <div class="note">
            <strong>Note:</strong> The translated file is <strong>not modified</strong>.
            This report is for review purposes only.
        </div>
    </div>
</div>
</body>
</html>"""


# ── Runner ────────────────────────────────────────────────────────────────────
def run_quality_check(translated_path: Path, original_path: Path | None = None,
                      save_report: bool = True, output_format: str = "both"):
    ext = translated_path.suffix.lower()

    print(f"\nQuality Check: {translated_path.name}")
    print("=" * 60)

    if ext == ".pptx":
        try:
            analysis = analyze_pptx(translated_path, original_path)
        except ImportError:
            print("Error: python-pptx not installed — run: pip install python-pptx")
            sys.exit(1)
    elif ext == ".docx":
        try:
            analysis = analyze_docx(translated_path, original_path)
        except ImportError:
            print("Error: python-docx not installed — run: pip install python-docx")
            sys.exit(1)
    elif ext == ".xlsx":
        try:
            analysis = analyze_xlsx(translated_path, original_path)
        except ImportError:
            print("Error: openpyxl not installed — run: pip install openpyxl")
            sys.exit(1)
    else:
        print(f"Unsupported format: {ext}  (supported: .pptx, .docx, .xlsx)")
        sys.exit(1)

    text_report = generate_report_text(analysis)
    print("\n" + text_report)

    if save_report:
        base = translated_path.stem
        out_dir = translated_path.parent

        if output_format in ("text", "both"):
            p = out_dir / f"{base}_quality_report.txt"
            p.write_text(text_report, encoding="utf-8")
            print(f"\nText report: {p}")

        if output_format in ("html", "both"):
            p = out_dir / f"{base}_quality_report.html"
            p.write_text(generate_report_html(analysis), encoding="utf-8")
            print(f"HTML report: {p}")

    return analysis


# ── CLI ───────────────────────────────────────────────────────────────────────
def print_usage():
    print("""
Translation Quality Check  —  GPT-5
======================================

Analyzes a translated file and creates a SEPARATE quality report.
The translated file is NEVER modified.

Usage:
  python quality_check.py <translated_file> [original_file] [options]

Supported formats:
  .pptx   PowerPoint
  .docx   Word document
  .xlsx   Excel spreadsheet

Options:
  --no-save    Print to console only, don't save report files
  --html       Save HTML report only
  --text       Save text report only

Examples:
  python quality_check.py presentation_thai.pptx
  python quality_check.py report_thai.docx report_english.docx
  python quality_check.py data_thai.xlsx data_english.xlsx --html

Output:
  <name>_quality_report.txt   — plain text
  <name>_quality_report.html  — open in browser (recommended)

Setup:
  Add OPENAI_API_KEY to .env for GPT-5 powered analysis.
  Without it, falls back to rule-based checks.
""")


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("--help", "-h", "help"):
        print_usage()
        sys.exit(0)

    translated_path = Path(sys.argv[1])
    if not translated_path.exists():
        print(f"Error: file not found — {translated_path}")
        sys.exit(1)

    save_report = "--no-save" not in sys.argv
    output_format = "html" if "--html" in sys.argv else "text" if "--text" in sys.argv else "both"

    original_path = None
    for arg in sys.argv[2:]:
        if not arg.startswith("--"):
            p = Path(arg)
            if p.exists():
                original_path = p
            else:
                print(f"Warning: original file not found — {arg}")
            break

    run_quality_check(translated_path, original_path, save_report, output_format)


if __name__ == "__main__":
    main()
