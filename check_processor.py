# -*- coding: utf-8 -*-
"""
==================================================================
  עיבוד שיק בודד -> שורת תוצאה
==================================================================
לוגיקה משותפת ל-main.py (עם דפדפן) ול-process_only.py (בלי דפדפן):
מקבל שיק (צילום + נתוני שורה), קורא אותו עם Claude, מתאים ללקוח
לפי ח.פ/שם/סכום, ומחזיר ResultRow מוכן לכתיבה לאקסל.

כאן גם נמצאת ההגנה החשובה: שיק שחזר (סכום שלילי) -> לא מוציאים קבלה.
==================================================================
"""

from typing import Optional

from check_reader import CheckReader
from customer_matcher import CustomerMatcher
from excel_writer import (
    ResultRow, STATUS_READY, STATUS_REVIEW, STATUS_BOUNCED,
)


def process_one(
    reader: CheckReader,
    matcher: CustomerMatcher,
    image_path: Optional[str],
    row_reference: Optional[str] = None,
    row_amount: Optional[str] = None,
    is_bounced: bool = False,
) -> ResultRow:
    """
    מעבד שיק בודד ומחזיר שורת תוצאה. מטפל בכל המקרים:
    שיק שחזר, צילום חסר, צילום לא קריא, והתאמה רגילה.
    """
    # --- מקרה 1: שיק שחזר (סכום שלילי) -> לא מוציאים קבלה! ---
    if is_bounced:
        print("      ⛔ שיק שחזר (סכום שלילי) - לא להוציא קבלה.")
        return ResultRow(
            detected_name=None, matched_name=None, confidence=0,
            check_number=None, bank_name=None, branch_number=None,
            account_number=None, due_date=None, amount=row_amount,
            status=STATUS_BOUNCED, maven_reference=row_reference,
            image_path=image_path, company_id=None,
            evidence="זוהה כשיק שחזר לפי סכום שלילי בתנועה",
        )

    # --- מקרה 2: אין צילום -> בדיקה ידנית ---
    if not image_path:
        return ResultRow(
            detected_name=None, matched_name=None, confidence=0,
            check_number=None, bank_name=None, branch_number=None,
            account_number=None, due_date=None, amount=row_amount,
            status=STATUS_REVIEW, maven_reference=row_reference,
            image_path=None, evidence="לא הורד צילום של השיק",
        )

    # --- קריאת השיק עם Claude ---
    details = reader.read_check(image_path)

    # --- מקרה 3: צילום לא קריא -> בדיקה ידנית ---
    if not details.is_readable:
        print("      🔶 הצילום לא קריא מספיק - דורש בדיקה ידנית.")
        return ResultRow(
            detected_name=details.drawer_name, matched_name=None, confidence=0,
            check_number=details.check_number, bank_name=details.bank_name,
            branch_number=details.branch_number, account_number=details.account_number,
            due_date=details.due_date, amount=details.amount or row_amount,
            status=STATUS_REVIEW, maven_reference=row_reference,
            image_path=image_path, company_id=details.company_id,
            evidence="הצילום לא קריא",
        )

    # --- מקרה 4: התאמה רגילה (ח.פ + שם + סכום) ---
    # הסכום להתאמה: מעדיפים את מה שנקרא מהצילום; אם חסר - מהשורה.
    amount_for_match = details.amount or row_amount
    match = matcher.match(
        detected_name=details.drawer_name,
        company_id=details.company_id,
        amount=amount_for_match,
    )
    status = STATUS_READY if not match.needs_review else STATUS_REVIEW
    evidence_text = "; ".join(match.evidence) if match.evidence else match.reason

    icon = "✅" if status == STATUS_READY else "🔶"
    print(f"      {icon} {details.drawer_name} -> "
          f"{match.matched_name or '?'} ({match.reason})")

    return ResultRow(
        detected_name=details.drawer_name,
        matched_name=match.matched_name,
        confidence=match.confidence,
        # מספר השיק נלקח מהצילום בלבד (האסמכתא מהשורה היא לא מספר השיק).
        check_number=details.check_number,
        bank_name=details.bank_name,
        branch_number=details.branch_number,
        account_number=details.account_number,
        due_date=details.due_date,
        amount=details.amount or row_amount,
        status=status,
        maven_reference=row_reference,
        image_path=image_path,
        company_id=details.company_id,
        evidence=evidence_text,
    )
