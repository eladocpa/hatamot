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

from dotenv import load_dotenv

import config
from bank_match_scraper import BankMatchScraper
from check_reader import CheckReader
from customer_matcher import CustomerMatcher, load_customers
from excel_writer import ResultRow, STATUS_READY, STATUS_REVIEW, write_results


def main():
    # טוען את מפתח ה-API מקובץ .env אל תוך משתני הסביבה.
    # ככה Claude מקבל את המפתח בלי שהוא כתוב בקוד.
    load_dotenv()

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
    matcher = CustomerMatcher(customers)

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

        # ---- שלב 5-6: קריאת כל שיק והתאמה ללקוח ----
        print("\n🤖 שולח את הצילומים ל-Claude וקורא את הפרטים...")
        for idx, check_row in enumerate(check_rows, start=1):
            print(f"\n  [{idx}/{len(check_rows)}] מעבד שיק...")

            # אם לא הצלחנו להוריד צילום - מסמנים לבדיקה ידנית.
            if not check_row.image_path:
                results.append(_blank_review_row(
                    reason="לא הורד צילום של השיק"
                ))
                continue

            # קוראים את השיק עם Claude.
            details = reader.read_check(check_row.image_path)

            # אם הצילום לא היה קריא - לבדיקה ידנית.
            if not details.is_readable:
                results.append(ResultRow(
                    detected_name=details.drawer_name,
                    matched_name=None, confidence=0,
                    check_number=details.check_number or check_row.row_reference,
                    bank_name=details.bank_name,
                    branch_number=details.branch_number,
                    account_number=details.account_number,
                    due_date=details.due_date,
                    amount=details.amount,
                    status=STATUS_REVIEW,
                    image_path=check_row.image_path,
                ))
                print("      🔶 הצילום לא קריא מספיק - דורש בדיקה ידנית.")
                continue

            # מתאימים את השם שזוהה ללקוח ברשימה.
            match = matcher.match(details.drawer_name)
            status = STATUS_READY if not match.needs_review else STATUS_REVIEW

            # מספר השיק: מעדיפים את מה שנקרא מהצילום; אם חסר -
            # משתמשים במספר האסמכתא שמופיע בשורת הטבלה.
            check_number = details.check_number or check_row.row_reference

            results.append(ResultRow(
                detected_name=details.drawer_name,
                matched_name=match.matched_name,
                confidence=match.confidence,
                check_number=check_number,
                bank_name=details.bank_name,
                branch_number=details.branch_number,
                account_number=details.account_number,
                due_date=details.due_date,
                amount=details.amount,
                status=status,
                image_path=check_row.image_path,
            ))

            icon = "✅" if status == STATUS_READY else "🔶"
            print(f"      {icon} {details.drawer_name} -> "
                  f"{match.matched_name or '?'} ({match.reason})")

    finally:
        # תמיד סוגרים את הדפדפן, גם אם הייתה שגיאה.
        scraper.close()

    # ---- שלב 7: כתיבת קובץ האקסל ----
    if results:
        write_results(results)
        ready = sum(1 for r in results if r.status == STATUS_READY)
        review = len(results) - ready
        print("\n" + "=" * 60)
        print(f"  סיכום: {len(results)} שיקים עובדו.")
        print(f"    ✅ מוכנים לקבלה:    {ready}")
        print(f"    🔶 לבדיקה ידנית:   {review}")
        print("=" * 60)
        print(f"\n👉 פתח את הקובץ '{config.OUTPUT_FILE}' ובדוק את הטבלה.")
        print("   אחרי שתאשר - נוסיף את שלב הוצאת הקבלות.")


def _blank_review_row(reason: str) -> ResultRow:
    """שורה ריקה שמסומנת לבדיקה ידנית (כשלא הצלחנו להוריד צילום)."""
    return ResultRow(
        detected_name=None, matched_name=None, confidence=0,
        check_number=None, bank_name=None, branch_number=None,
        account_number=None, due_date=None, amount=None,
        status=STATUS_REVIEW, image_path=None,
    )


if __name__ == "__main__":
    main()
