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
        עובר על כל השורות בטבלה, מאתר רק את שורות "הפקדת שיק",
        ולכל אחת לוחץ על אייקון המסמך ומוריד את צילום השיק.
        """
        results: List[CheckRow] = []

        rows = self._page.locator(ROW_SELECTOR)
        total_rows = rows.count()
        print(f"\n🔎 נמצאו {total_rows} שורות בטבלה. סורק אחרי 'הפקדת שיק'...")

        check_counter = 0
        for i in range(total_rows):
            row = rows.nth(i)
            try:
                row_text = row.inner_text(timeout=2000)
            except PWTimeout:
                continue

            # מתעלמים מכל שורה שאינה הפקדת שיק.
            if config.CHECK_DEPOSIT_LABEL not in row_text:
                continue

            check_counter += 1
            short_text = " ".join(row_text.split())[:70]
            print(f"\n  💳 שיק #{check_counter} (שורה {i}): {short_text}...")

            # מחלצים את מספר האסמכתא מתוך טקסט השורה (המספר בסוגריים).
            ref_match = ROW_REFERENCE_PATTERN.search(row_text)
            row_reference = ref_match.group(1) if ref_match else None

            # מנסים לפתוח ולהוריד את צילום השיק של השורה הזו.
            image_path = self._capture_check_image(row, check_counter)
            results.append(
                CheckRow(
                    row_index=i,
                    row_text=row_text,
                    image_path=image_path,
                    row_reference=row_reference,
                )
            )

        print(f"\n✅ סיימתי לסרוק. נמצאו {len(results)} תנועות 'הפקדת שיק'.")
        return results

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
        לוחץ על אלמנט נתון ומנסה לתפוס את צילום השיק שנפתח -
        בין אם בחלון/לשונית חדשה ובין אם בחלונית קופצת (מודאל).
        מחזיר True אם נשמר צילום.
        """
        try:
            # אפשרות א': נפתח חלון/לשונית חדשה בעקבות הלחיצה.
            try:
                with self._page.context.expect_page(timeout=3000) as popup_info:
                    element.click(timeout=2000)
                popup = popup_info.value
                popup.wait_for_load_state("domcontentloaded")
                saved = self._save_image_from(popup, image_path)
                popup.close()
                if saved:
                    return True
            except PWTimeout:
                pass  # לא נפתח חלון חדש -> בודקים מודאל באותו דף

            # אפשרות ב': נפתחה חלונית/תמונה גדולה באותו דף.
            time.sleep(1)
            if self._save_image_from(self._page, image_path, in_modal=True):
                self._page.keyboard.press("Escape")  # סוגרים את החלונית
                time.sleep(0.4)
                return True

        except Exception:
            pass
        return False

    def _save_image_from(self, page: Page, image_path: str, in_modal: bool = False) -> bool:
        """
        מאתר את תמונת השיק בדף/חלון נתון ושומר אותה לקובץ.
        מחזיר True אם הצליח.
        """
        images = page.locator(POPUP_IMAGE_SELECTOR)
        count = images.count()
        if count == 0:
            return False

        # בוחרים את התמונה הגדולה ביותר - בדרך כלל זו תמונת השיק עצמה
        # (ולא אייקון קטן). מודדים את הגודל של כל תמונה ובוחרים את הגדולה.
        best_index = -1
        best_area = 0
        for idx in range(count):
            img = images.nth(idx)
            try:
                box = img.bounding_box()
            except Exception:
                box = None
            if box is None:
                continue
            area = box["width"] * box["height"]
            # מתעלמים מתמונות זעירות (אייקונים) - דורשים שטח מינימלי.
            if area > best_area and area > 5000:
                best_area = area
                best_index = idx

        if best_index == -1:
            return False

        # שומרים צילום מסך של אלמנט התמונה לקובץ.
        try:
            images.nth(best_index).screenshot(path=image_path)
            print(f"      📸 צילום השיק נשמר: {image_path}")
            return True
        except Exception:
            return False

    # --------------------------------------------------------------
    #  סגירה מסודרת
    # --------------------------------------------------------------
    def close(self):
        """סוגר את הדפדפן בצורה מסודרת."""
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()
