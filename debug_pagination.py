# -*- coding: utf-8 -*-
"""
==================================================================
  אבחון אזור הדפדוף - בלי קריאות API
==================================================================
מתחבר, נועל על מסך התאמות הבנק, ומדפיס את ה-HTML של אזור הדפדוף
(כפתורי העמודים). זה עוזר לדייק את הסלקטור של "העמוד הבא".
לא שולח שום צילום ל-Claude, אז לא עולה כסף.

הרצה:   python debug_pagination.py
==================================================================
"""

from bank_match_scraper import BankMatchScraper


def main():
    scraper = BankMatchScraper()
    try:
        scraper.start_and_login()
        scraper.focus_bank_match_page()
        scraper.print_pagination_html()
        print("\n👆 העתק את כל ה-HTML שמופיע למעלה ושלח לי.")
    finally:
        input("\nלחץ Enter לסגירת הדפדפן... ")
        scraper.close()


if __name__ == "__main__":
    main()
