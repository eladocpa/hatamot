# -*- coding: utf-8 -*-
"""
==================================================================
  עיבוד שיקים -> שורות תוצאה (כולל זיהוי שיקים שחזרו)
==================================================================
לוגיקה משותפת ל-main.py (עם דפדפן) ול-process_only.py (בלי דפדפן).

>>> זיהוי שיק שחזר - הנקודה החשובה <<<
שיק שחזר מופיע ב-Maven בשתי תנועות:
  1. ההפקדה המקורית - שורת "הפקדת שיק". *זו* שעליה אסור להוציא קבלה.
  2. פעולת ההחזרה - שורת "החזרת שיק" (גם עם צילום השיק).

לכן אנחנו עובדים בשני מעברים:
  מעבר 1: קוראים את *כל* השיקים (הפקדות + החזרות) ומוציאים מספר שיק.
  מעבר 2: מספרי השיק שהופיעו בשורת "החזרת שיק" = "שיקים שחזרו".
          כל הפקדה עם מספר שיק כזה מסומנת "לא להוציא קבלה".
==================================================================
"""

import re
from dataclasses import dataclass
from typing import List, Optional

from check_reader import CheckReader, CheckDetails
from customer_matcher import CustomerMatcher
from excel_writer import (
    ResultRow, STATUS_READY, STATUS_REVIEW, STATUS_BOUNCED,
)


@dataclass
class CheckItem:
    """פריט לעיבוד - שיק בודד (חיובי או שלילי)."""
    image_path: Optional[str]
    row_reference: Optional[str] = None
    row_amount: Optional[str] = None
    is_bounced: bool = False


def _digits_only(value: Optional[str]) -> Optional[str]:
    """משאיר רק ספרות (להשוואת מספרי שיק). מחזיר None אם אין ספרות."""
    if not value:
        return None
    digits = re.sub(r"\D", "", str(value))
    return digits or None


def process_all(
    reader: CheckReader,
    matcher: CustomerMatcher,
    items: List[CheckItem],
) -> List[ResultRow]:
    """
    מעבד את כל השיקים בשני מעברים, ומחזיר רשימת שורות תוצאה.
    מזהה שיקים שחזרו ומחריג את ההפקדות החיוביות התואמות.
    """
    # ---- מעבר 1: קריאת כל הצילומים (חיוביים ושליליים) ----
    # שומרים את התוצאה לכל פריט כדי לא לקרוא פעמיים.
    print("\n🤖 מעבר 1/2: קורא את כל הצילומים עם Claude...")
    read_cache: List[Optional[CheckDetails]] = []
    for idx, item in enumerate(items, start=1):
        kind = "החזרה (שלילי)" if item.is_bounced else "הפקדה"
        print(f"  [{idx}/{len(items)}] קורא {kind}...")
        if item.image_path:
            read_cache.append(reader.read_check(item.image_path))
        else:
            read_cache.append(None)

    # ---- בניית רשימת "שיקים שחזרו" מתוך התנועות השליליות ----
    # לפי מספר שיק (מדויק), ולפי סכום (גיבוי).
    bounced_check_numbers = set()
    bounced_amounts = set()
    for item, details in zip(items, read_cache):
        if not item.is_bounced:
            continue
        if details is not None:
            num = _digits_only(details.check_number)
            if num:
                bounced_check_numbers.add(num)
        amt = _digits_only(item.row_amount)
        if amt:
            bounced_amounts.add(amt)

    if bounced_check_numbers or bounced_amounts:
        print(f"\n  ⛔ זוהו שיקים שחזרו: "
              f"{len(bounced_check_numbers)} לפי מספר שיק, "
              f"{len(bounced_amounts)} לפי סכום (גיבוי).")

    # ---- מעבר 2: בניית שורות התוצאה + החרגת הפקדות שחזרו ----
    print("\n🔗 מעבר 2/2: מתאים ללקוחות ומחריג שיקים שחזרו...")
    results: List[ResultRow] = []
    for idx, (item, details) in enumerate(zip(items, read_cache), start=1):
        print(f"\n  [{idx}/{len(items)}] מעבד...")

        # שורת "החזרת שיק" עצמה = פעולת ההחזרה. מציגים לתיעוד,
        # אבל ברור שלא מוציאים עליה קבלה.
        if item.is_bounced:
            num = _digits_only(details.check_number) if details else None
            results.append(_bounced_action_row(item, details))
            print(f"      ⛔ פעולת החזרת שיק (מספר שיק: {num or '?'}).")
            continue

        # התאמת לקוח להפקדה החיובית.
        result = _match_positive(reader, matcher, item, details)

        # האם ההפקדה הזו היא של שיק שחזר? בודקים לפי מספר שיק, ואז סכום.
        num = _digits_only(details.check_number) if details else None
        amt = _digits_only(item.row_amount)
        is_returned = (num is not None and num in bounced_check_numbers) or \
                      (num is None and amt is not None and amt in bounced_amounts)

        if is_returned:
            # מחריגים: לא להוציא קבלה, גם אם ההתאמה ללקוח ודאית.
            result.status = STATUS_BOUNCED
            note = f"שיק זה חזר (מספר שיק {num or amt}) - אין להוציא קבלה"
            result.evidence = note
            print(f"      ⛔ ההפקדה הזו היא של שיק שחזר - לא להוציא קבלה!")

        results.append(result)

    return results


def _match_positive(reader, matcher, item: CheckItem, details) -> ResultRow:
    """בונה שורת תוצאה להפקדה חיובית: קריאה + התאמת לקוח."""
    # אין צילום -> בדיקה ידנית.
    if not item.image_path or details is None:
        return ResultRow(
            detected_name=None, matched_name=None, confidence=0,
            check_number=None, bank_name=None, branch_number=None,
            account_number=None, due_date=None, amount=item.row_amount,
            status=STATUS_REVIEW, maven_reference=item.row_reference,
            image_path=item.image_path, evidence="לא הורד צילום של השיק",
        )

    # צילום לא קריא -> בדיקה ידנית.
    if not details.is_readable:
        print("      🔶 הצילום לא קריא מספיק - דורש בדיקה ידנית.")
        return ResultRow(
            detected_name=details.drawer_name, matched_name=None, confidence=0,
            check_number=details.check_number, bank_name=details.bank_name,
            branch_number=details.branch_number, account_number=details.account_number,
            due_date=details.due_date, amount=details.amount or item.row_amount,
            status=STATUS_REVIEW, maven_reference=item.row_reference,
            image_path=item.image_path, company_id=details.company_id,
            evidence="הצילום לא קריא",
        )

    # התאמת לקוח (ח.פ + שם + סכום).
    amount_for_match = details.amount or item.row_amount
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
        check_number=details.check_number,
        bank_name=details.bank_name,
        branch_number=details.branch_number,
        account_number=details.account_number,
        due_date=details.due_date,
        amount=details.amount or item.row_amount,
        status=status,
        maven_reference=item.row_reference,
        image_path=item.image_path,
        company_id=details.company_id,
        evidence=evidence_text,
        matched_customer_id=match.matched_id,
        matched_company_id=match.matched_company_id,
    )


def _bounced_action_row(item: CheckItem, details) -> ResultRow:
    """שורת תיעוד לפעולת ההחזרה עצמה (התנועה השלילית)."""
    check_number = details.check_number if details else None
    return ResultRow(
        detected_name=details.drawer_name if details else None,
        matched_name=None, confidence=0,
        check_number=check_number, bank_name=details.bank_name if details else None,
        branch_number=details.branch_number if details else None,
        account_number=details.account_number if details else None,
        due_date=details.due_date if details else None,
        amount=item.row_amount,
        status=STATUS_BOUNCED, maven_reference=item.row_reference,
        image_path=item.image_path,
        company_id=details.company_id if details else None,
        evidence="פעולת החזרת שיק (תנועה שלילית) - לתיעוד בלבד",
    )
