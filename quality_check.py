#!/usr/bin/env python3
"""
Translation Quality Check Tool
Analyzes translated PowerPoint files and creates a SEPARATE quality report
The translated slides remain untouched - issues are flagged in the report only
"""

import os
import sys
import re
import json
from pathlib import Path
from datetime import datetime
from pptx import Presentation
from collections import defaultdict

# Thai character range
THAI_PATTERN = re.compile(r'[\u0E00-\u0E7F]')
# English word pattern
ENGLISH_WORD_PATTERN = re.compile(r'\b[A-Za-z]{3,}\b')
# Number pattern
NUMBER_PATTERN = re.compile(r'\b\d+(?:[.,]\d+)*%?\b')


# Words commonly kept in English (acceptable, won't be flagged)
ACCEPTABLE_ENGLISH = {
    # Tech terms
    'AI', 'API', 'IT', 'IoT', 'CEO', 'CTO', 'CFO', 'COO',
    'DNA', 'RNA', 'GPS', 'USB', 'PDF', 'URL', 'HTML', 'CSS',
    'JavaScript', 'Python', 'Java', 'React', 'Angular', 'Vue',
    'Machine', 'Learning', 'Deep', 'Blockchain', 'Cloud',
    'AWS', 'Azure', 'Google', 'Microsoft', 'Apple', 'Meta',
    'Facebook', 'LinkedIn', 'Twitter', 'Instagram', 'TikTok',
    'iPhone', 'Android', 'iOS', 'Windows', 'Linux', 'Mac',
    'Excel', 'PowerPoint', 'Word', 'Outlook', 'Teams',
    'WiFi', 'Bluetooth', 'LTE', 'SQL', 'NoSQL',
    'Docker', 'Kubernetes', 'DevOps', 'Agile', 'Scrum',
    # Business terms
    'ROI', 'KPI', 'B2B', 'B2C', 'SaaS', 'PaaS', 'IaaS',
    'CRM', 'ERP', 'HR', 'PR', 'Marketing', 'Branding',
    # Common
    'vs', 'etc', 'USD', 'EUR', 'THB', 'Q1', 'Q2', 'Q3', 'Q4',
    'km', 'kg', 'mg', 'ml', 'GB', 'TB', 'MB', 'KB',
    'Email', 'Website', 'Online', 'Offline', 'Digital',
}

ACCEPTABLE_ENGLISH_LOWER = {word.lower() for word in ACCEPTABLE_ENGLISH}


class QualityIssue:
    """Represents a quality issue found in translation."""

    def __init__(self, slide_num: int, issue_type: str, location: str,
                 original: str, translated: str, suggestion: str, severity: str):
        self.slide_num = slide_num
        self.issue_type = issue_type
        self.location = location  # "slide" or "notes"
        self.original = original
        self.translated = translated
        self.suggestion = suggestion
        self.severity = severity  # "warning", "info", "critical"


def extract_slide_content(slide):
    """Extract all text content from a slide."""
    texts = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            texts.append(shape.text.strip())
    return "\n".join(texts)


def extract_notes(slide):
    """Extract speaker notes from a slide."""
    try:
        return slide.notes_slide.notes_text_frame.text.strip()
    except:
        return ""


def find_untranslated_words(text: str) -> list:
    """Find English words that might need translation."""
    if not THAI_PATTERN.search(text):
        return []  # No Thai text, skip check

    words = ENGLISH_WORD_PATTERN.findall(text)
    untranslated = []

    for word in words:
        if word.lower() not in ACCEPTABLE_ENGLISH_LOWER and len(word) > 3:
            untranslated.append(word)

    return list(set(untranslated))


def check_numbers_match(original: str, translated: str) -> list:
    """Check if numbers from original appear in translation."""
    orig_nums = set(NUMBER_PATTERN.findall(original))
    trans_nums = set(NUMBER_PATTERN.findall(translated))
    return list(orig_nums - trans_nums)


def analyze_presentation(translated_path: Path, original_path: Path = None) -> dict:
    """
    Analyze translated presentation and generate quality report.
    Does NOT modify the translated file.
    """

    translated_prs = Presentation(str(translated_path))
    original_prs = None

    if original_path and original_path.exists():
        original_prs = Presentation(str(original_path))

    issues = []
    slides_data = []

    for i, slide in enumerate(translated_prs.slides, 1):
        slide_content = extract_slide_content(slide)
        notes_content = extract_notes(slide)

        orig_content = ""
        orig_notes = ""

        if original_prs and i <= len(original_prs.slides):
            orig_slide = original_prs.slides[i - 1]
            orig_content = extract_slide_content(orig_slide)
            orig_notes = extract_notes(orig_slide)

        slide_issues = []

        # Check slide content
        if slide_content:
            # Untranslated words
            untranslated = find_untranslated_words(slide_content)
            if untranslated:
                slide_issues.append(QualityIssue(
                    slide_num=i,
                    issue_type="Untranslated Text",
                    location="Slide Content",
                    original=", ".join(untranslated[:5]),
                    translated="",
                    suggestion="Review if these English words should be translated to Thai",
                    severity="warning"
                ))

            # Missing numbers
            if orig_content:
                missing_nums = check_numbers_match(orig_content, slide_content)
                if missing_nums:
                    slide_issues.append(QualityIssue(
                        slide_num=i,
                        issue_type="Missing Numbers",
                        location="Slide Content",
                        original=", ".join(missing_nums),
                        translated="",
                        suggestion="Verify these numbers appear correctly in translation",
                        severity="warning"
                    ))

                # Length check
                if len(orig_content) > 20:
                    ratio = len(slide_content) / len(orig_content)
                    if ratio > 1.5:
                        slide_issues.append(QualityIssue(
                            slide_num=i,
                            issue_type="Text Length",
                            location="Slide Content",
                            original=f"{len(orig_content)} chars",
                            translated=f"{len(slide_content)} chars ({int((ratio-1)*100)}% longer)",
                            suggestion="Text may not fit in slide - consider shortening",
                            severity="info"
                        ))

        # Check speaker notes
        if notes_content:
            untranslated = find_untranslated_words(notes_content)
            if untranslated:
                slide_issues.append(QualityIssue(
                    slide_num=i,
                    issue_type="Untranslated Text",
                    location="Speaker Notes",
                    original=", ".join(untranslated[:5]),
                    translated="",
                    suggestion="Review if these English words should be translated",
                    severity="warning"
                ))

        # Store slide data
        slides_data.append({
            "slide_num": i,
            "has_content": bool(slide_content),
            "has_notes": bool(notes_content),
            "content_preview": slide_content[:100] + "..." if len(slide_content) > 100 else slide_content,
            "notes_preview": notes_content[:100] + "..." if len(notes_content) > 100 else notes_content,
            "issue_count": len(slide_issues)
        })

        issues.extend(slide_issues)

    return {
        "filename": translated_path.name,
        "original": original_path.name if original_path else None,
        "timestamp": datetime.now().isoformat(),
        "slide_count": len(translated_prs.slides),
        "total_issues": len(issues),
        "issues": issues,
        "slides": slides_data
    }


def generate_report_text(analysis: dict) -> str:
    """Generate a text report for console and file output."""

    issues = analysis["issues"]

    # Group by severity
    critical = [i for i in issues if i.severity == "critical"]
    warnings = [i for i in issues if i.severity == "warning"]
    info = [i for i in issues if i.severity == "info"]

    # Calculate score
    score = max(0, 100 - (len(critical) * 20) - (len(warnings) * 5) - (len(info) * 1))

    if score >= 90:
        score_icon, score_label = "🟢", "Excellent"
    elif score >= 70:
        score_icon, score_label = "🟡", "Good"
    elif score >= 50:
        score_icon, score_label = "🟠", "Fair"
    else:
        score_icon, score_label = "🔴", "Needs Review"

    lines = []
    lines.append("=" * 70)
    lines.append("TRANSLATION QUALITY REPORT")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"File: {analysis['filename']}")
    if analysis['original']:
        lines.append(f"Original: {analysis['original']}")
    lines.append(f"Date: {analysis['timestamp'][:10]}")
    lines.append("")
    lines.append(f"{score_icon} Quality Score: {score}/100 ({score_label})")
    lines.append("")
    lines.append("-" * 70)
    lines.append("SUMMARY")
    lines.append("-" * 70)
    lines.append(f"Total Slides: {analysis['slide_count']}")
    lines.append(f"Total Issues: {analysis['total_issues']}")
    if critical:
        lines.append(f"  🔴 Critical: {len(critical)}")
    if warnings:
        lines.append(f"  🟡 Warnings: {len(warnings)}")
    if info:
        lines.append(f"  🔵 Info: {len(info)}")
    lines.append("")

    # Issues by slide
    if issues:
        lines.append("-" * 70)
        lines.append("ISSUES BY SLIDE")
        lines.append("-" * 70)
        lines.append("")

        # Group issues by slide
        by_slide = defaultdict(list)
        for issue in issues:
            by_slide[issue.slide_num].append(issue)

        for slide_num in sorted(by_slide.keys()):
            slide_issues = by_slide[slide_num]
            lines.append(f"┌─ SLIDE {slide_num} ({len(slide_issues)} issue{'s' if len(slide_issues) > 1 else ''})")

            for issue in slide_issues:
                severity_icon = {"critical": "🔴", "warning": "🟡", "info": "🔵"}[issue.severity]
                lines.append(f"│")
                lines.append(f"│  {severity_icon} {issue.issue_type}")
                lines.append(f"│     Location: {issue.location}")
                if issue.original:
                    lines.append(f"│     Found: {issue.original}")
                if issue.translated:
                    lines.append(f"│     Translation: {issue.translated}")
                lines.append(f"│     → {issue.suggestion}")

            lines.append(f"└{'─' * 50}")
            lines.append("")

    # Slides needing review
    problem_slides = list(by_slide.keys()) if issues else []
    if problem_slides:
        lines.append("-" * 70)
        lines.append("SLIDES TO REVIEW")
        lines.append("-" * 70)
        lines.append("")
        lines.append(f"Review these slides: {', '.join(map(str, sorted(problem_slides)))}")
        lines.append("")
        lines.append("Open the translated presentation alongside this report")
        lines.append("to review and fix any issues before presenting.")
        lines.append("")

    lines.append("=" * 70)
    lines.append("NOTE: The translated presentation file is NOT modified.")
    lines.append("This report is for review purposes only.")
    lines.append("=" * 70)

    return "\n".join(lines)


def generate_report_html(analysis: dict) -> str:
    """Generate an HTML report for better viewing."""

    issues = analysis["issues"]
    critical = [i for i in issues if i.severity == "critical"]
    warnings = [i for i in issues if i.severity == "warning"]
    info = [i for i in issues if i.severity == "info"]
    score = max(0, 100 - (len(critical) * 20) - (len(warnings) * 5) - (len(info) * 1))

    if score >= 90:
        score_color, score_label = "#22c55e", "Excellent"
    elif score >= 70:
        score_color, score_label = "#eab308", "Good"
    elif score >= 50:
        score_color, score_label = "#f97316", "Fair"
    else:
        score_color, score_label = "#ef4444", "Needs Review"

    # Group issues by slide
    by_slide = defaultdict(list)
    for issue in issues:
        by_slide[issue.slide_num].append(issue)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Quality Report - {analysis['filename']}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; padding: 20px; }}
        .container {{ max-width: 900px; margin: 0 auto; }}
        .header {{ background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
        .header .meta {{ color: #666; font-size: 14px; }}
        .score-box {{ display: inline-flex; align-items: center; gap: 12px; background: {score_color}15; border: 2px solid {score_color}; border-radius: 8px; padding: 12px 20px; margin-top: 16px; }}
        .score-number {{ font-size: 32px; font-weight: bold; color: {score_color}; }}
        .score-label {{ font-size: 16px; color: {score_color}; }}
        .summary {{ background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .summary h2 {{ font-size: 18px; margin-bottom: 16px; }}
        .stats {{ display: flex; gap: 20px; flex-wrap: wrap; }}
        .stat {{ background: #f8f9fa; padding: 12px 20px; border-radius: 8px; }}
        .stat-value {{ font-size: 24px; font-weight: bold; }}
        .stat-label {{ font-size: 12px; color: #666; text-transform: uppercase; }}
        .issues {{ background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .issues h2 {{ font-size: 18px; margin-bottom: 16px; }}
        .slide-group {{ border: 1px solid #e5e7eb; border-radius: 8px; margin-bottom: 16px; overflow: hidden; }}
        .slide-header {{ background: #f8f9fa; padding: 12px 16px; font-weight: 600; border-bottom: 1px solid #e5e7eb; }}
        .issue {{ padding: 12px 16px; border-bottom: 1px solid #f0f0f0; }}
        .issue:last-child {{ border-bottom: none; }}
        .issue-type {{ display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }}
        .severity {{ width: 10px; height: 10px; border-radius: 50%; }}
        .severity.critical {{ background: #ef4444; }}
        .severity.warning {{ background: #eab308; }}
        .severity.info {{ background: #3b82f6; }}
        .issue-details {{ font-size: 14px; color: #666; margin-left: 18px; }}
        .suggestion {{ background: #fef3c7; padding: 8px 12px; border-radius: 4px; margin-top: 8px; font-size: 13px; margin-left: 18px; }}
        .no-issues {{ text-align: center; padding: 40px; color: #22c55e; }}
        .note {{ background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 16px; margin-top: 20px; font-size: 14px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Translation Quality Report</h1>
            <div class="meta">
                <strong>File:</strong> {analysis['filename']}<br>
                {"<strong>Original:</strong> " + analysis['original'] + "<br>" if analysis['original'] else ""}
                <strong>Date:</strong> {analysis['timestamp'][:10]}
            </div>
            <div class="score-box">
                <span class="score-number">{score}</span>
                <span class="score-label">/100<br>{score_label}</span>
            </div>
        </div>

        <div class="summary">
            <h2>Summary</h2>
            <div class="stats">
                <div class="stat">
                    <div class="stat-value">{analysis['slide_count']}</div>
                    <div class="stat-label">Slides</div>
                </div>
                <div class="stat">
                    <div class="stat-value">{analysis['total_issues']}</div>
                    <div class="stat-label">Total Issues</div>
                </div>
                <div class="stat">
                    <div class="stat-value" style="color: #ef4444">{len(critical)}</div>
                    <div class="stat-label">Critical</div>
                </div>
                <div class="stat">
                    <div class="stat-value" style="color: #eab308">{len(warnings)}</div>
                    <div class="stat-label">Warnings</div>
                </div>
                <div class="stat">
                    <div class="stat-value" style="color: #3b82f6">{len(info)}</div>
                    <div class="stat-label">Info</div>
                </div>
            </div>
        </div>

        <div class="issues">
            <h2>Issues by Slide</h2>
"""

    if not issues:
        html += """
            <div class="no-issues">
                <div style="font-size: 48px; margin-bottom: 16px;">✅</div>
                <div style="font-size: 18px; font-weight: 600;">No issues found!</div>
                <div style="color: #666; margin-top: 8px;">Your translation looks great.</div>
            </div>
"""
    else:
        for slide_num in sorted(by_slide.keys()):
            slide_issues = by_slide[slide_num]
            html += f"""
            <div class="slide-group">
                <div class="slide-header">Slide {slide_num} ({len(slide_issues)} issue{'s' if len(slide_issues) > 1 else ''})</div>
"""
            for issue in slide_issues:
                html += f"""
                <div class="issue">
                    <div class="issue-type">
                        <span class="severity {issue.severity}"></span>
                        <strong>{issue.issue_type}</strong>
                        <span style="color: #999; font-size: 12px;">({issue.location})</span>
                    </div>
                    {"<div class='issue-details'><strong>Found:</strong> " + issue.original + "</div>" if issue.original else ""}
                    {"<div class='issue-details'><strong>Translation:</strong> " + issue.translated + "</div>" if issue.translated else ""}
                    <div class="suggestion">💡 {issue.suggestion}</div>
                </div>
"""
            html += """
            </div>
"""

    html += """
            <div class="note">
                <strong>Note:</strong> The translated presentation file is NOT modified.
                This report is for review purposes only. Open the translated presentation
                alongside this report to review and fix any issues before presenting.
            </div>
        </div>
    </div>
</body>
</html>
"""

    return html


def run_quality_check(translated_path: Path, original_path: Path = None,
                      save_report: bool = True, output_format: str = "both"):
    """Run quality check and generate reports."""

    print(f"\n🔍 Quality Check: {translated_path.name}")
    print("=" * 60)

    # Analyze
    analysis = analyze_presentation(translated_path, original_path)

    # Generate text report
    text_report = generate_report_text(analysis)
    print(text_report)

    # Save reports
    if save_report:
        base_name = translated_path.stem
        report_dir = translated_path.parent

        if output_format in ("text", "both"):
            text_path = report_dir / f"{base_name}_quality_report.txt"
            with open(text_path, 'w', encoding='utf-8') as f:
                f.write(text_report)
            print(f"\n📄 Text report saved: {text_path}")

        if output_format in ("html", "both"):
            html_report = generate_report_html(analysis)
            html_path = report_dir / f"{base_name}_quality_report.html"
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_report)
            print(f"🌐 HTML report saved: {html_path}")

    return analysis


def print_usage():
    """Print usage instructions."""
    print("""
🔍 Translation Quality Check Tool
==================================

Analyzes translated PowerPoint and creates a SEPARATE quality report.
The translated slides remain UNTOUCHED - issues are flagged in the report only.

Usage:
  python quality_check.py <translated.pptx> [original.pptx] [options]

Arguments:
  translated.pptx    The translated PowerPoint file to check
  original.pptx      (Optional) Original file for comparison

Options:
  --no-save          Don't save report files (console only)
  --html             Save HTML report only
  --text             Save text report only

Output:
  📄 <filename>_quality_report.txt   - Text report
  🌐 <filename>_quality_report.html  - HTML report (open in browser)

Examples:
  python quality_check.py presentation_thai.pptx
  python quality_check.py translated.pptx original.pptx
  python quality_check.py translated.pptx --html

The HTML report is recommended - open it in a browser alongside
your PowerPoint to review issues slide by slide.
""")


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("--help", "-h", "help"):
        print_usage()
        sys.exit(0)

    translated_path = Path(sys.argv[1])

    if not translated_path.exists():
        print(f"❌ File not found: {translated_path}")
        sys.exit(1)

    # Parse arguments
    original_path = None
    save_report = "--no-save" not in sys.argv

    output_format = "both"
    if "--html" in sys.argv:
        output_format = "html"
    elif "--text" in sys.argv:
        output_format = "text"

    # Find original file if provided
    for arg in sys.argv[2:]:
        if not arg.startswith("--") and arg.endswith(".pptx"):
            original_path = Path(arg)
            if not original_path.exists():
                print(f"⚠️  Original file not found: {original_path}")
                original_path = None
            break

    # Run check
    run_quality_check(translated_path, original_path, save_report, output_format)


if __name__ == "__main__":
    main()
