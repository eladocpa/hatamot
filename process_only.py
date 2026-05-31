# -*- coding: utf-8 -*-
"""
==================================================================
  עיבוד מחדש של צילומים שכבר ירדו - בלי דפדפן ובלי כניסה ל-Maven
==================================================================
שימושי כשהצילומים כבר ירדו (בתיקייה checks_images) ואתה רוצה
רק לקרוא אותם שוב עם Claude ולבנות את טבלת האקסל -
בלי לעבור שוב את כל תהליך הכניסה והסריקה.

הרצה:   python process_only.py
==================================================================
"""

import glob
import json
import os

from dotenv import load_dotenv

import config
from check_reader import CheckReader
from customer_matcher import CustomerMatcher, load_customers
from excel_writer import ResultRow, STATUS_READY, STATUS_REVIEW, write_results
# משתמשים בבדיקת המפתח החכמה שכבר כתבנו ב-main.
from main import _load_api_key


def _load_index():
    """
    טוען את רשימת הצילומים לעיבוד.
    מעדיף את קובץ האינדקס (שנשמר בסריקה האחרונה); אם אין -
    פשוט עובר על כל קבצי ה-PNG בתיקיית הצילומים.
    """
    index_path = os.path.join(config.IMAGES_DIR, "checks_index.json")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return json.load(f)

    # גיבוי: אין אינדקס - עוברים על כל הצילומים לפי שם הקובץ.
    images = sorted(glob.glob(os.path.join(config.IMAGES_DIR, "*.png")))
    return [{"image_path": p, "row_reference": None, "row_text": ""} for p in images]


def main():
    if not _load_api_key():
        return

    print("=" * 60)
    print("  🔁  עיבוד מחדש של צילומים קיימים (בלי דפדפן)")
    print("=" * 60)

    # טעינת רשימת הלקוחות
    try:
        customers = load_customers()
    except FileNotFoundError:
        print(f"\n❌ לא נמצא קובץ הלקוחות '{config.CUSTOMERS_FILE}'.")
        return
    matcher = CustomerMatcher(customers)

    entries = _load_index()
    if not entries:
        print(f"\nℹ️  לא נמצאו צילומים בתיקייה '{config.IMAGES_DIR}'.")
        print("   הרץ קודם את main.py כדי להוריד צילומים מ-Maven.")
        return

    print(f"\n🤖 מעבד {len(entries)} צילומים עם Claude...")
    reader = CheckReader()
    results = []

    for idx, entry in enumerate(entries, start=1):
        image_path = entry.get("image_path")
        reference = entry.get("row_reference")
        print(f"\n  [{idx}/{len(entries)}] {image_path}")

        if not image_path or not os.path.exists(image_path):
            print("      ⚠️  הצילום לא נמצא - מדלג.")
            continue

        details = reader.read_check(image_path)

        if not details.is_readable:
            results.append(ResultRow(
                detected_name=details.drawer_name, matched_name=None, confidence=0,
                check_number=details.check_number, bank_name=details.bank_name,
                branch_number=details.branch_number, account_number=details.account_number,
                due_date=details.due_date, amount=details.amount,
                status=STATUS_REVIEW, maven_reference=reference, image_path=image_path,
            ))
            print("      🔶 הצילום לא קריא מספיק - דורש בדיקה ידנית.")
            continue

        match = matcher.match(details.drawer_name)
        status = STATUS_READY if not match.needs_review else STATUS_REVIEW
        results.append(ResultRow(
            detected_name=details.drawer_name, matched_name=match.matched_name,
            confidence=match.confidence, check_number=details.check_number,
            bank_name=details.bank_name, branch_number=details.branch_number,
            account_number=details.account_number, due_date=details.due_date,
            amount=details.amount, status=status,
            maven_reference=reference, image_path=image_path,
        ))
        icon = "✅" if status == STATUS_READY else "🔶"
        print(f"      {icon} {details.drawer_name} -> "
              f"{match.matched_name or '?'} ({match.reason})")

    if results:
        write_results(results)
        ready = sum(1 for r in results if r.status == STATUS_READY)
        print("\n" + "=" * 60)
        print(f"  סיכום: {len(results)} שיקים. ✅ מוכנים: {ready}  "
              f"🔶 לבדיקה: {len(results) - ready}")
        print("=" * 60)
        print(f"\n👉 פתח את הקובץ '{config.OUTPUT_FILE}' ובדוק את הטבלה.")


if __name__ == "__main__":
    main()
