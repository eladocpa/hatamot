# -*- coding: utf-8 -*-
"""
==================================================================
  הכלי הראשי - מריצים את הקובץ הזה כדי להפעיל הכל
==================================================================
התהליך, צעד אחר צעד:
  1. טוען את רשימת הלקוחות שלך מקובץ האקסל.
  2. פותח דפדפן ונותן לך להתחבר ידנית ל-Maven.
  3. עובר למסך התאמות הבנק.
  4. מאתר את כל תנועות 'הפקדת שיק' ומוריד את הצילומים.
  5. שולח כל צילום ל-Claude ומחלץ את פרטי השיק.
  6. מתאים כל שיק ללקוח ברשימה (או מסמן 'דורש בדיקה ידנית').
  7. בונה קובץ אקסל מסודר לבדיקה שלך.

בשלב הזה הכלי לא מוציא שום קבלה - רק בונה את הטבלה.
==================================================================
"""

import os

from dotenv import load_dotenv

import config
from bank_match_scraper import BankMatchScraper
from check_reader import CheckReader
from check_processor import CheckItem, process_all
from customer_matcher import CustomerMatcher, load_customers, load_income_index
from excel_writer import (
    ResultRow, STATUS_READY, STATUS_REVIEW, STATUS_BOUNCED, write_results,
)


def _load_api_key() -> bool:
    """
    טוען את מפתח ה-API מקובץ .env (או .env.txt), מנקה אותו אם צריך,
    ובודק שהוא תקין. מחזיר True אם הכל תקין, או מדפיס אבחון ומחזיר False.
    """
    # טוענים משני השמות האפשריים (Notepad לפעמים שומר כ-.env.txt).
    for fname in (".env", ".env.txt"):
        if os.path.exists(fname):
            load_dotenv(fname, override=True)

    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip().strip('"').strip("'")

    # ניקוי אוטומטי: אם נשאר טקסט הדוגמה בעברית דבוק לסוף המפתח - חותכים אותו.
    if "הדבק" in api_key:
        api_key = api_key.split("הדבק")[0].rstrip("- ").strip()

    # אם נראה תקין - מעדכנים לערך הנקי וממשיכים.
    if api_key.startswith("sk-ant-") and len(api_key) > 30:
        os.environ["ANTHROPIC_API_KEY"] = api_key
        return True

    # נכשל - מדפיסים אבחון מדויק שיעזור להבין מה הבעיה.
    print("\n❌ מפתח ה-API של Claude לא נמצא או לא תקין.")
    found = [f for f in (".env", ".env.txt") if os.path.exists(f)]
    if found:
        print(f"   📄 נמצאו הקבצים: {', '.join(found)}")
    else:
        print("   ⚠️  לא נמצא קובץ .env בתיקייה הזו בכלל!")

    if api_key:
        preview = (api_key[:14] + "..." + api_key[-4:]) if len(api_key) > 20 else api_key
        print(f"   🔎 המפתח שנקרא: {preview}  (אורך {len(api_key)} תווים)")
        if not api_key.startswith("sk-ant-"):
            print("   👉 המפתח לא מתחיל ב-sk-ant- . כנראה הועתק לא נכון.")
    else:
        print("   🔎 לא נקרא שום ערך עבור ANTHROPIC_API_KEY.")

    print("\n   💡 הדרך הקלה והבטוחה ביותר לתקן: הרץ")
    print("        python set_key.py")
    print("      והדבק את המפתח. זה ייצור קובץ .env תקין אוטומטית.")
    return False


def main():
    # טוען את מפתח ה-API. אם משהו לא תקין - עוצרים כאן,
    # לפני שפותחים דפדפן (כדי לא להתחבר ל-Maven סתם).
    if not _load_api_key():
        return

    print("=" * 60)
    print("  🧾  כלי התאמות בנק - הפקדות שיקים  🧾")
    print("=" * 60)

    # ---- שלב 1: טעינת רשימת הלקוחות ----
    print(f"\n📂 טוען את רשימת הלקוחות מהקובץ '{config.CUSTOMERS_FILE}'...")
    try:
        customers = load_customers()
    except FileNotFoundError:
        print(f"\n❌ לא נמצא קובץ הלקוחות '{config.CUSTOMERS_FILE}'.")
        print("   צור אותו (אפשר להריץ: python create_customer_template.py)")
        return
    except ValueError as e:
        print(f"\n❌ {e}")
        return

    print(f"   נטענו {len(customers)} לקוחות.")

    # טוענים את קובץ תנועות ההכנסה (אם קיים) - לחיזוק התאמה לפי סכום.
    income_index = load_income_index()
    if income_index:
        print(f"   נטענו תנועות הכנסה מ-'{config.INCOME_FILE}' "
              f"({len(income_index)} סכומים) - ישמשו לחיזוק ההתאמה.")

    matcher = CustomerMatcher(customers, income_index)

    # ---- שלב 2-4: דפדפן, כניסה, וסריקת השיקים ----
    scraper = BankMatchScraper()
    reader = CheckReader()
    results = []

    try:
        scraper.start_and_login()
        scraper.focus_bank_match_page()
        check_rows = scraper.collect_check_rows()

        if not check_rows:
            print("\nℹ️  לא נמצאו תנועות 'הפקדת שיק' לא־מותאמות. אין מה לעבד.")
            return

        # ---- שלב 5-6: קריאת כל השיקים (שני מעברים) והתאמה ללקוח ----
        # מעבירים את כל השורות (כולל שליליות) כדי לזהות שיקים שחזרו.
        items = [
            CheckItem(
                image_path=cr.image_path,
                row_reference=cr.row_reference,
                row_amount=cr.row_amount,
                is_bounced=cr.is_bounced,
            )
            for cr in check_rows
        ]
        results = process_all(reader, matcher, items)

    finally:
        # תמיד סוגרים את הדפדפן, גם אם הייתה שגיאה.
        scraper.close()

    # ---- שלב 7: כתיבת קובץ האקסל ----
    if results:
        write_results(results)
        ready = sum(1 for r in results if r.status == STATUS_READY)
        bounced = sum(1 for r in results if r.status == STATUS_BOUNCED)
        review = len(results) - ready - bounced
        print("\n" + "=" * 60)
        print(f"  סיכום: {len(results)} שיקים עובדו.")
        print(f"    ✅ מוכנים לקבלה:    {ready}")
        print(f"    🔶 לבדיקה ידנית:   {review}")
        print(f"    ⛔ שיקים שחזרו:    {bounced}")
        print("=" * 60)
        print(f"\n👉 פתח את הקובץ '{config.OUTPUT_FILE}' ובדוק את הטבלה.")
        print("   אחרי שתאשר - נוסיף את שלב הוצאת הקבלות.")


if __name__ == "__main__":
    main()
