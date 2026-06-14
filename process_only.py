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
from check_processor import CheckItem, process_all
from customer_matcher import CustomerMatcher, load_customers, load_income_index
from excel_writer import STATUS_READY, STATUS_BOUNCED, write_results
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
    return [{"image_path": p} for p in images]


def main():
    if not _load_api_key():
        return

    print("=" * 60)
    print("  🔁  עיבוד מחדש של צילומים קיימים (בלי דפדפן)")
    print("=" * 60)

    # טעינת רשימת הלקוחות + תנועות הכנסה (אם קיימות)
    try:
        customers = load_customers()
    except FileNotFoundError:
        print(f"\n❌ לא נמצא קובץ הלקוחות '{config.CUSTOMERS_FILE}'.")
        return
    income_index = load_income_index()
    if income_index:
        print(f"   נטענו תנועות הכנסה ({len(income_index)} סכומים) לחיזוק ההתאמה.")
    matcher = CustomerMatcher(customers, income_index)

    entries = _load_index()
    if not entries:
        print(f"\nℹ️  לא נמצאו צילומים בתיקייה '{config.IMAGES_DIR}'.")
        print("   הרץ קודם את main.py כדי להוריד צילומים מ-Maven.")
        return

    reader = CheckReader()
    items = [
        CheckItem(
            image_path=entry.get("image_path"),
            row_reference=entry.get("row_reference"),
            row_amount=entry.get("row_amount"),
            row_date=entry.get("row_date"),
            is_bounced=entry.get("is_bounced", False),
        )
        for entry in entries
    ]
    results = process_all(reader, matcher, items)

    if results:
        write_results(results, customers=customers)
        ready = sum(1 for r in results if r.status == STATUS_READY)
        bounced = sum(1 for r in results if r.status == STATUS_BOUNCED)
        review = len(results) - ready - bounced
        print("\n" + "=" * 60)
        print(f"  סיכום: {len(results)} שיקים. ✅ מוכנים: {ready}  "
              f"🔶 לבדיקה: {review}  ⛔ חזרו: {bounced}")
        print("=" * 60)
        print(f"\n👉 פתח את הקובץ '{config.OUTPUT_FILE}' ובדוק את הטבלה.")


if __name__ == "__main__":
    main()
