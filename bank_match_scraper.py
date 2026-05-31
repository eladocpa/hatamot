# -*- coding: utf-8 -*-
"""
==================================================================
  סריקת מסך התאמות הבנק (ה"ידיים" של הכלי)
==================================================================
המודול הזה פותח דפדפן, נכנס למערכת Maven, עובר על טבלת
התאמות הבנק, מאתר את שורות "הפקדת שיק", לוחץ על אייקון
המסמך בכל שורה ומוריד את צילום השיק.

>>> חשוב לדעת <<<
המבנה הפנימי של אתר Maven אינו ידוע לי מראש (זה אתר דינמי).
לכן כל הסלקטורים (ה"כתובות" של הכפתורים בתוך הדף) מרוכזים
למעלה, ומסומנים בבירור. בהרצה הראשונה ייתכן שנצטרך לכוונן
אותם יחד מול האתר האמיתי - זה נורמלי לגמרי.
==================================================================
"""

import json
import os
import re
import time
from dataclasses import dataclass
from typing import List, Optional

from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

import config


# ==================================================================
#  >>> אזור הכיוונון <<<
#  אם הכלי לא מוצא את השורות או את אייקון המסמך,
#  כאן המקום לעדכן את ה"כתובות" של הרכיבים בדף.
# ==================================================================

# סלקטור לשורות בטבלה. ברירת מחדל: כל שורת טבלה (<tr>) בדף.
ROW_SELECTOR = "table tr"

# אלמנטים שייחשבו כ"אייקון השטר" בתוך תא הפעולה, לפי סדר עדיפות.
# מנסים קודם אלמנטים שנראים כמו אייקון (תמונה/אייקון/כפתור), ואז קישורים.
# הכלי לוחץ על כל מועמד ובודק אם נפתח צילום; אם כן - עוצר. אם לא - ממשיך לבא.
ICON_CANDIDATE_SELECTOR = "img, i, svg, button, a, [onclick]"

# כשנפתח צילום השיק - איפה התמונה נמצאת?
# הכלי ינסה למצוא את התמונה הגדולה ביותר בחלון/חלונית שנפתחה.
POPUP_IMAGE_SELECTOR = "img"

# מעבר ל"עמוד הבא" במנוע הדפדוף (RichFaces dataScroller) = תנועות מוקדמות
# יותר. הכלי לוחץ על כפתור "הבא"; כשמגיעים לעמוד האחרון הכפתור מקבל
# class="disabled" ואין בו <a>, אז הסלקטורים האלה פשוט לא ימצאו כלום -> סוף.
NEXT_PAGE_SELECTORS = [
    "li.paginate_button.next:not(.disabled) a",   # המבנה המדויק באתר
    "li.page-item.next:not(.disabled) a",
    ".paginate_button.next a",
    "li.next a",
    # גיבויים למסגרות אחרות (ליתר ביטחון)
    ".ui-paginator-next:not(.ui-state-disabled)",
    ".p-paginator-next:not(.p-disabled)",
    "a[aria-label*='Next']",
]

# מיכל אזור הדפדוף - בתוכו נחפש את מספר העמוד הפעיל ואת כפתורי המספרים.
PAGINATOR_CONTAINER_SELECTORS = [
    "ul.pagination",
    ".pagination",
    "[class*='paginat']",
    ".ui-paginator",
    ".p-paginator",
    "tfoot",
]

# סלקטורים למספר העמוד *הפעיל* (כדי לזהות שבאמת עברנו עמוד).
ACTIVE_PAGE_SELECTORS = [
    "li.paginate_button.active",
    "li.page-item.active",
    "li.active",
    ".active",
]

# תבנית לחילוץ מספר האסמכתא מתוך טקסט השורה, למשל "הפקדת שיק - (88635)".
ROW_REFERENCE_PATTERN = re.compile(r"\((\d+)\)")

# ==================================================================


@dataclass
class CheckRow:
    """שורת "הפקדת שיק" שמצאנו, יחד עם הנתיב לצילום שהורדנו."""
    row_index: int                   # מספר השורה בטבלה (לתיעוד)
    row_text: str                    # הטקסט המלא של השורה (לתיעוד)
    image_path: Optional[str]        # נתיב לקובץ הצילום שהורדנו (או None)
    row_reference: Optional[str] = None  # המספר בסוגריים מהשורה, למשל "88635"


class BankMatchScraper:
    """מנהל את כל האינטראקציה עם אתר Maven."""

    def __init__(self):
        os.makedirs(config.IMAGES_DIR, exist_ok=True)
        self._playwright = None
        self._browser = None
        self._page: Optional[Page] = None

    # --------------------------------------------------------------
    #  פתיחת דפדפן + כניסה ידנית למערכת
    # --------------------------------------------------------------
    def start_and_login(self):
        """
        פותח דפדפן על מסך הכניסה, ועוצר כדי שתעשה ידנית את כל מה שצריך:
        כניסה עם שם משתמש וסיסמה, בחירת הלקוח המיוצג, ומעבר למסך
        התאמות הבנק שלו. אחרי שתגיע למסך - תלחץ Enter בטרמינל.

        הגישה הזו (ניווט ידני) מאפשרת לך להריץ את הכלי על כל לקוח שתרצה,
        פשוט על ידי בחירתו לפני שאתה לוחץ Enter.
        """
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(
            headless=not config.SHOW_BROWSER  # גלוי או נסתר, לפי ההגדרה
        )
        self._page = self._browser.new_page()
        self._page.set_default_timeout(config.PAGE_TIMEOUT_SECONDS * 1000)

        print("\n🌐 פותח את מסך הכניסה של Maven...")
        self._page.goto(config.LOGIN_URL)

        # כאן עוצרים ומחכים לך. הסיסמה נשארת רק אצלך - לא נשמרת בשום מקום.
        print("\n" + "=" * 60)
        print("  ✋ עצור! עכשיו תורך (הכל ידני בדפדפן):")
        print("  1. הקלד שם משתמש וסיסמה והתחבר ל-Maven.")
        print("  2. בחר מהרשימה את הלקוח המיוצג שאתה רוצה לעבד.")
        print("  3. נווט אל מסך 'התאמות בנק' של אותו לקוח.")
        print("  4. ודא שאתה רואה את הטבלה עם התנועות הלא־מותאמות.")
        print("  5. חזור לכאן ולחץ Enter כדי שהכלי ישתלט על המסך.")
        print("=" * 60)
        input("  >>> לחץ Enter כשאתה על מסך התאמות הבנק... ")

    # --------------------------------------------------------------
    #  מעבר למסך התאמות הבנק
    # --------------------------------------------------------------
    def focus_bank_match_page(self):
        """
        'נועל' את הכלי על המסך שאליו ניווטת ידנית.
        לא מנווט לשום מקום - פשוט מוודא שהוא קורא את הלשונית הפעילה
        (במקרה ש-Maven פתח את מסך ההתאמות בלשונית חדשה) ומחכה שתסיים להיטען.
        """
        print("\n🧭 נועל על מסך התאמות הבנק שפתחת...")

        # אם פתחת לשונית חדשה תוך כדי הניווט - עוברים אליה (האחרונה שנפתחה).
        pages = self._page.context.pages
        if pages:
            self._page = pages[-1]
            self._page.set_default_timeout(config.PAGE_TIMEOUT_SECONDS * 1000)

        # נותנים לטבלה רגע להיטען (אתרי JSF טוענים חלק מהתוכן בנפרד).
        try:
            self._page.wait_for_load_state("networkidle")
        except PWTimeout:
            pass
        time.sleep(2)

    # --------------------------------------------------------------
    #  איתור שורות "הפקדת שיק" והורדת הצילומים
    # --------------------------------------------------------------
    def collect_check_rows(self) -> List[CheckRow]:
        """
        עובר על כל העמודים (עם דפדוף אחורה), ובכל עמוד מאתר את שורות
        "הפקדת שיק", מוריד את הצילום, וממשיך עד שאין יותר עמודים אחורה.
        מונע כפילויות לפי מספר האסמכתא.
        """
        results: List[CheckRow] = []
        seen_references = set()  # אסמכתאות שכבר עיבדנו (למניעת כפילויות)
        page_num = 1

        while True:
            print(f"\n📄 ===== עמוד {page_num} =====")
            new_count = self._scan_current_page(results, seen_references)
            print(f"   נמצאו {new_count} שיקים חדשים בעמוד זה "
                  f"(סה\"כ עד כה: {len(results)}).")

            # הגנה: לא חורגים ממספר העמודים המקסימלי.
            if page_num >= config.MAX_PAGES:
                print(f"   ⚠️  הגעתי למגבלת {config.MAX_PAGES} עמודים - עוצר.")
                break

            # מנסים לעבור לעמוד הבא (תנועות מוקדמות יותר). אם אין - סיימנו.
            if not self._go_to_next_page(page_num + 1):
                print("   ✅ אין עוד עמודים - הגענו למסך האחרון.")
                break

            page_num += 1

        print(f"\n✅ סיימתי לסרוק {page_num} עמודים. "
              f"סה\"כ {len(results)} תנועות 'הפקדת שיק'.")

        # שומרים "אינדקס" קטן של מה שמצאנו, ליד הצילומים.
        # זה מאפשר לעבד מחדש את הצילומים בלי להיכנס שוב ל-Maven
        # (באמצעות הסקריפט process_only.py).
        self._save_index(results)
        return results

    def _scan_current_page(self, results: List[CheckRow], seen_references: set) -> int:
        """
        סורק את העמוד הנוכחי: מאתר שורות 'הפקדת שיק', מוריד צילום לכל אחת,
        ומוסיף ל-results. מדלג על אסמכתאות שכבר ראינו. מחזיר כמה חדשות נמצאו.
        """
        rows = self._page.locator(ROW_SELECTOR)
        total_rows = rows.count()
        new_count = 0
        page_refs = []   # כל האסמכתאות של 'הפקדת שיק' בעמוד (לאבחון)
        dup_count = 0    # כמה כפילויות דילגנו

        for i in range(total_rows):
            row = rows.nth(i)
            try:
                row_text = row.inner_text(timeout=2000)
            except PWTimeout:
                continue

            # מתעלמים מכל שורה שאינה הפקדת שיק.
            if config.CHECK_DEPOSIT_LABEL not in row_text:
                continue

            # מחלצים את מספר האסמכתא (המספר בסוגריים) ומדלגים על כפילויות.
            ref_match = ROW_REFERENCE_PATTERN.search(row_text)
            row_reference = ref_match.group(1) if ref_match else None
            page_refs.append(row_reference or "?")
            if row_reference and row_reference in seen_references:
                dup_count += 1
                continue
            if row_reference:
                seen_references.add(row_reference)

            check_number = len(results) + 1
            short_text = " ".join(row_text.split())[:70]
            print(f"\n  💳 שיק #{check_number}: {short_text}...")

            image_path = self._capture_check_image(row, check_number)
            results.append(
                CheckRow(
                    row_index=i,
                    row_text=row_text,
                    image_path=image_path,
                    row_reference=row_reference,
                )
            )
            new_count += 1

        # אבחון: מה נמצא בעמוד הזה (עוזר להבין אם יש שיקים בעמודים נוספים).
        if page_refs:
            shown = ", ".join(page_refs[:15])
            print(f"   🔎 שורות 'הפקדת שיק' בעמוד: {len(page_refs)} "
                  f"(אסמכתאות: {shown})")
            if dup_count:
                print(f"      ({dup_count} כפילויות דולגו - כבר עובדו בעמוד קודם)")
        else:
            print("   🔎 אין שורות 'הפקדת שיק' בעמוד הזה.")

        return new_count

    def _data_signature(self) -> str:
        """
        מחזיר 'טביעת אצבע' של *שורות הנתונים* בעמוד (שורות שמכילות תאריך).
        אלה השורות שמתחלפות בין עמודים, אז זו הדרך האמינה לדעת אם
        תוכן העמוד באמת התעדכן (ולא רק מספר העמוד באזור הדפדוף).
        """
        try:
            rows = self._page.locator(ROW_SELECTOR)
            total = rows.count()
            parts = []
            for i in range(total):
                if len(parts) >= 6:
                    break
                try:
                    t = rows.nth(i).inner_text(timeout=800)
                except Exception:
                    continue
                if "/20" in t:  # שורה שמכילה תאריך = שורת תנועה
                    parts.append(" ".join(t.split()))
            return " | ".join(parts)[:600]
        except Exception:
            return ""

    def _find_paginator(self):
        """מאתר את מיכל אזור הדפדוף. מחזיר locator או None."""
        for selector in PAGINATOR_CONTAINER_SELECTORS:
            cand = self._page.locator(selector).first
            try:
                if cand.count() > 0 and cand.is_visible():
                    return cand
            except Exception:
                continue
        return None

    def _go_to_next_page(self, target_page: int) -> bool:
        """
        עובר לעמוד הבא (תנועות מוקדמות יותר). שתי אסטרטגיות:
        (1) ללחוץ ישירות על מספר העמוד הבא (target_page),
        (2) ואם לא נמצא - ללחוץ על חץ "הבא".
        מחזיר True אם עברנו לעמוד חדש, או False אם הגענו לסוף.
        """
        before = self._data_signature()
        clicked = False

        # אסטרטגיה 1: לחיצה על כפתור "הבא" (החץ) - מדויק לפי מבנה האתר.
        for selector in NEXT_PAGE_SELECTORS:
            cand = self._page.locator(selector).first
            try:
                if cand.count() > 0 and cand.is_visible():
                    cand.click(timeout=3000)
                    clicked = True
                    break
            except Exception:
                continue

        # אסטרטגיה 2: גיבוי - לחיצה ישירה על מספר העמוד הבא.
        if not clicked:
            paginator = self._find_paginator()
            scope = paginator if paginator is not None else self._page
            try:
                number_btn = scope.get_by_text(str(target_page), exact=True).first
                if number_btn.count() > 0 and number_btn.is_visible():
                    number_btn.click(timeout=3000)
                    clicked = True
            except Exception:
                clicked = False

        if not clicked:
            if config.DEBUG_PRINT_HTML:
                self._dump_pagination_html()
            return False

        # ממתינים עד ש*שורות הנתונים עצמן* יתחלפו (ולא רק מספר העמוד).
        # זה מונע סריקה של תוכן ישן לפני שה-AJAX סיים לטעון את העמוד החדש.
        deadline = time.time() + 10
        while time.time() < deadline:
            now = self._data_signature()
            if now and now != before:
                time.sleep(0.5)  # רגע קטן נוסף שהכל יתייצב
                return True
            time.sleep(0.5)

        # תוכן השורות לא השתנה תוך 10 שניות - כנראה הגענו לסוף (או תקלה).
        if config.DEBUG_PRINT_HTML:
            self._dump_pagination_html()
        return False

    def _dump_pagination_html(self):
        """
        מדפיס את ה-HTML של אזור הדפדוף, כדי שנוכל לדייק את הסלקטורים.
        (תוכל להעתיק ולשלוח לי אם הדפדוף לא עובד.)
        """
        print("      🐞 לא הצלחתי לדפדף. העתק את ה-HTML הבא ושלח לי:")
        self.print_pagination_html()

    def print_pagination_html(self):
        """מדפיס את ה-HTML של אזור הדפדוף (לאבחון, בלי קריאות API)."""
        print("=" * 60)
        print("  HTML של אזור הדפדוף:")
        print("=" * 60)
        found_any = False
        for selector in PAGINATOR_CONTAINER_SELECTORS:
            try:
                area = self._page.locator(selector).first
                if area.count() > 0 and area.is_visible():
                    print(f"\n--- נמצא ב: {selector} ---")
                    print(area.inner_html(timeout=3000)[:3000])
                    found_any = True
                    break
            except Exception:
                continue
        if not found_any:
            print("  ⚠️  לא נמצא אזור דפדוף לפי הסלקטורים המוכרים.")
        print("=" * 60)

    def _save_index(self, rows: List[CheckRow]):
        """שומר קובץ אינדקס (checks_index.json) שמקשר צילום -> אסמכתא + טקסט שורה."""
        index = [
            {
                "image_path": r.image_path,
                "row_reference": r.row_reference,
                "row_text": r.row_text,
            }
            for r in rows
        ]
        index_path = os.path.join(config.IMAGES_DIR, "checks_index.json")
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)

    # --------------------------------------------------------------
    #  לחיצה על אייקון המסמך ושמירת התמונה
    # --------------------------------------------------------------
    def _capture_check_image(self, row, check_number: int) -> Optional[str]:
        """
        בתוך שורה נתונה - מאתר את תא ה'פעולה' (זה שמכיל 'הפקדת שיק'),
        מנסה ללחוץ על אייקון השטר שבו, ושומר את צילום השיק לקובץ.
        מחזיר את נתיב הקובץ, או None אם נכשל.
        """
        image_path = os.path.join(
            config.IMAGES_DIR, f"check_{check_number:03d}.png"
        )

        # שלב 1: מאתרים את התא בשורה שמכיל את הטקסט 'הפקדת שיק' (תא הפעולה).
        # שם נמצא אייקון השטר. אם לא נמצא - מחפשים בכל השורה.
        action_cell = row.locator(
            "td", has_text=config.CHECK_DEPOSIT_LABEL
        ).first
        if action_cell.count() == 0:
            action_cell = row

        # שלב 2: אוספים את כל ה"מועמדים" להיות אייקון השטר בתוך התא.
        candidates = action_cell.locator(ICON_CANDIDATE_SELECTOR)
        count = candidates.count()

        # שלב 3: לוחצים על כל מועמד בתורו ובודקים אם נפתח צילום.
        # מדלגים על האלמנט שמכיל את הטקסט 'הפקדת שיק' עצמו (לא האייקון).
        for j in range(count):
            cand = candidates.nth(j)
            try:
                if not cand.is_visible():
                    continue
                text = cand.inner_text(timeout=500)
            except Exception:
                text = ""
            if config.CHECK_DEPOSIT_LABEL in text:
                continue  # זה קישור הטקסט, לא אייקון הצילום

            if self._click_and_capture(cand, image_path):
                return image_path

        # שלב 4: לא הצלחנו. מדפיסים את ה-HTML של התא לצורך כיוונון.
        # (תוכל להעתיק את הפלט הזה ולשלוח לי כדי שאדייק את הזיהוי.)
        if config.DEBUG_PRINT_HTML:
            try:
                html = action_cell.inner_html(timeout=2000)
                print("      🐞 לא מצאתי את אייקון הצילום. "
                      "העתק את ה-HTML הבא ושלח לי:")
                print("      " + "-" * 50)
                print(html[:1500])
                print("      " + "-" * 50)
            except Exception:
                pass

        print("      ⚠️  לא נמצא צילום ברור לשיק הזה - מסומן לבדיקה ידנית.")
        return None

    def _click_and_capture(self, element, image_path: str) -> bool:
        """
        לוחץ על אלמנט נתון. אם נפתחה חלונית קופצת (מודאל) עם צילום השיק -
        שומר אותו וסוגר את החלונית. מחזיר True אם נשמר צילום.
        """
        try:
            element.click(timeout=2000)
        except Exception:
            return False

        # החלונית נפתחת באותו מסך - מחכים שתופיע בה תמונה גדולה (צילום השיק).
        if self._save_image_from(self._page, image_path):
            self._close_modal()
            return True

        # שום צילום לא נפתח -> כנראה לחצנו על אלמנט שגוי. סוגרים ליתר ביטחון.
        self._close_modal()
        return False

    def _close_modal(self):
        """סוגר חלונית קופצת (בדרך כלל Esc סוגר; מנסה גם כפתור סגירה)."""
        try:
            self._page.keyboard.press("Escape")
            time.sleep(0.3)
        except Exception:
            pass

    def _save_image_from(self, page: Page, image_path: str, poll_seconds: int = 5) -> bool:
        """
        ממתין עד שתופיע בדף תמונה גדולה (צילום השיק) ושומר אותה לקובץ.
        מחזיר True אם הצליח, או False אם לא הופיעה תמונה תוך poll_seconds שניות.
        """
        deadline = time.time() + poll_seconds
        while time.time() < deadline:
            best = self._find_largest_image(page)
            if best is not None:
                try:
                    best.screenshot(path=image_path)
                    print(f"      📸 צילום השיק נשמר: {image_path}")
                    return True
                except Exception:
                    pass
            time.sleep(0.5)
        return False

    def _find_largest_image(self, page: Page):
        """
        מחזיר את אלמנט התמונה הגדול ביותר בדף - בדרך כלל זה צילום השיק
        שנפתח בחלונית (ולא אייקון קטן). מחזיר None אם אין תמונה גדולה.
        """
        images = page.locator(POPUP_IMAGE_SELECTOR)
        count = images.count()
        best = None
        # שטח מינימלי (בפיקסלים) כדי להתעלם מאייקונים ולוגו - דורשים תמונה גדולה.
        best_area = 15000
        for idx in range(count):
            img = images.nth(idx)
            try:
                if not img.is_visible():
                    continue
                box = img.bounding_box()
            except Exception:
                box = None
            if box is None:
                continue
            area = box["width"] * box["height"]
            if area > best_area:
                best_area = area
                best = img
        return best

    # --------------------------------------------------------------
    #  סגירה מסודרת
    # --------------------------------------------------------------
    def close(self):
        """סוגר את הדפדפן בצורה מסודרת."""
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()
