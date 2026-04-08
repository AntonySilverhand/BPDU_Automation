#!/usr/bin/env python3
"""
BP Debate Union — Activity Review Document Validator

Validates a Word document (.docx) for compliance with school activity review standards.

Checks:
1. Photo count: 2-3 embedded images
2. Word count: 50-100 words (WPS-style — Chinese chars = 1 word each)
3. Language ratio: >50%% Chinese characters
4. Time reference: uses last-week format (上周三/上周五 etc.) and is plausible

Usage:
    python validate_review.py "path/to/activity_review.docx"

Exit codes:
    0 = all checks pass
    1 = one or more checks fail
"""

import argparse
import os
import re
import sys

from docx import Document


VALID_ACTIVITY_TYPES = {
    "常规活动", "比赛", "培训", "讲座",
    "苏格拉底式研讨会", "辩论训练", "workshop", "seminar",
}
VALID_WEEKDAY_NAMES = {"周一", "周二", "周三", "周四", "周五", "周六", "周日"}


def count_wps_words(text):
    """
    WPS word count style:
    - Each Chinese character = 1 word
    - Each English word (whitespace-separated token) = 1 word
    """
    chinese_chars = sum(1 for c in text if "\u4e00" <= c <= "\u9fff")
    english_words = len([w for w in text.split() if w.isalpha()])
    return chinese_chars + english_words


def count_chinese_ratio(text):
    """Return fraction of Chinese characters in the text."""
    if not text.strip():
        return 0.0
    chinese_chars = sum(1 for c in text if "\u4e00" <= c <= "\u9fff")
    total_chars = len(text)
    return chinese_chars / total_chars if total_chars > 0 else 0.0


def extract_last_week_date(text):
    """
    Extract last-week date references like '上周三', '上周五'.
    Returns the weekday name if found, else None.
    """
    match = re.search(r"上周([一二三四五六日])", text)
    if match:
        weekday = {"一": "周一", "二": "周二", "三": "周三", "四": "周四",
                   "五": "周五", "六": "周六", "日": "周日"}.get(match.group(1))
        return weekday
    return None


def validate(docx_path):
    """Run all validation checks. Returns (passed: bool, results: list[str])."""
    results = []

    if not os.path.exists(docx_path):
        return False, [f"File not found: {docx_path}"]

    # Reject legacy .doc format — python-docx only supports .docx
    lower = docx_path.lower()
    if lower.endswith(".doc") and not lower.endswith(".docx"):
        return False, ["Unsupported format: .doc (legacy). python-docx only supports .docx. Please convert to .docx first."]

    try:
        doc = Document(docx_path)
    except Exception as e:
        return False, [f"Cannot read .docx file: {e}"]

    # ---- Extract text ----
    full_text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())

    # ---- 1. Photo count ----
    image_parts = [rel for rel in doc.part.rels.values()
                   if "image" in rel.target_ref]
    photo_count = len(image_parts)
    if 2 <= photo_count <= 3:
        results.append(f"[PASS] Photo count: {photo_count} (2-3 required)")
    else:
        results.append(f"[FAIL] Photo count: {photo_count} (expected 2-3)")

    # ---- 2. Word count ----
    word_count = count_wps_words(full_text)
    if 50 <= word_count <= 100:
        results.append(f"[PASS] Word count: {word_count} (50-100 required)")
    else:
        results.append(f"[FAIL] Word count: {word_count} (expected 50-100)")

    # ---- 3. Chinese ratio ----
    ratio = count_chinese_ratio(full_text)
    if ratio > 0.5:
        results.append(f"[PASS] Chinese ratio: {ratio:.1%} (>50% required)")
    else:
        results.append(f"[FAIL] Chinese ratio: {ratio:.1%} (expected >50%)")

    # ---- 4. Time reference ----
    last_week_ref = extract_last_week_date(full_text)
    if last_week_ref:
        results.append(f"[PASS] Time reference: found '{last_week_ref}' (last-week format)")
    else:
        results.append(f"[FAIL] Time reference: no '上周X' pattern found")

    # ---- 5. Third-person perspective ----
    first_person = any(word in full_text for word in ["我们", "我社", "笔者", "本人"])
    if first_person:
        results.append(f"[FAIL] Perspective: first-person words detected (must be third-person)")
    else:
        results.append(f"[PASS] Perspective: third-person confirmed")

    # ---- Summary ----
    passed = all("[PASS]" in r for r in results)

    return passed, results


def main():
    parser = argparse.ArgumentParser(description="Validate BP DU Activity Review document.")
    parser.add_argument("docx_path", help="Path to the activity review .docx file")
    args = parser.parse_args()

    passed, results = validate(args.docx_path)

    print(f"\n=== Validation Results ===")
    for r in results:
        print(r)
    print()
    print("OVERALL:", "PASS" if passed else "FAIL")

    sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()
