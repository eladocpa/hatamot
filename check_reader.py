# -*- coding: utf-8 -*-
"""
==================================================================
  קריאת שיק באמצעות Claude (ה"עיניים" של הכלי)
==================================================================
המודול הזה מקבל תמונה של שיק, שולח אותה ל-Claude,
ומחזיר בחזרה את הפרטים בצורה מסודרת (כמו טופס מלא).

אנחנו משתמשים ב"פלט מובנה" (structured outputs) - זאת טכניקה
שמכריחה את Claude להחזיר בדיוק את השדות שאנחנו צריכים,
בלי טקסט מיותר, כך שאי אפשר לטעות בקריאה.
==================================================================
"""

import base64
import re
from typing import Optional

import anthropic
from pydantic import BaseModel, Field

import config


# ------------------------------------------------------------------
# הגדרת ה"טופס" שאנחנו רוצים ש-Claude ימלא עבור כל שיק.
# כל שדה כאן הוא פרט אחד שמחלצים מהצילום.
# ------------------------------------------------------------------
class CheckDetails(BaseModel):
    """הפרטים שמחלצים מצילום של שיק אחד."""

    # לכל השדות יש ברירת מחדל None, כדי שנוכל לבנות אובייקט "ריק"
    # במקרה של כשל בקריאה בלי שהקוד יקרוס.
    drawer_name: Optional[str] = Field(
        default=None,
        description="שם הלקוח / המושך - מי שכתב את השיק. "
                    "בדרך כלל מודפס בחלק העליון של השיק. "
                    "אם לא ברור או לא קריא - החזר null."
    )
    company_id: Optional[str] = Field(
        default=None,
        description="מספר ח.פ / עוסק מורשה / ע.מ של המושך - אם מודפס על השיק "
                    "(בדרך כלל ליד השם, מספר בן 9 ספרות). אם לא מופיע - החזר null."
    )
    check_number: Optional[str] = Field(
        default=None,
        description="מספר השיק - מופיע בדרך כלל בפינה השמאלית התחתונה של השיק "
                    "(לעיתים גם בפינה הימנית העליונה). זהו מספר השיק עצמו, "
                    "ולא מספר אסמכתא או מספר חשבון."
    )
    bank_name: Optional[str] = Field(
        default=None,
        description="שם הבנק (למשל: לאומי, פועלים, דיסקונט, מזרחי טפחות)."
    )
    branch_number: Optional[str] = Field(
        default=None,
        description="מספר הסניף - מספר בן 3 ספרות בדרך כלל."
    )
    account_number: Optional[str] = Field(
        default=None,
        description="מספר חשבון הבנק שמופיע על השיק."
    )
    due_date: Optional[str] = Field(
        default=None,
        description="תאריך הפירעון של השיק, בפורמט DD/MM/YYYY. "
                    "אם זה שיק דחוי - זה התאריך העתידי הכתוב עליו. "
                    "שים לב: בישראל התאריך נכתב תמיד משמאל לימין - "
                    "המספר השמאלי ביותר הוא היום, האמצעי הוא החודש, "
                    "והימני ביותר הוא השנה. אין לכך חריגים. "
                    "אם נכתב בכתב יד תאריך מקוצר (למשל 9/6/26) - "
                    "פענח אותו כ-09/06/2026: רפד יום וחודש לשתי ספרות, "
                    "והרחב שנה דו-ספרתית (26) לשנה מלאה בת ארבע ספרות (2026)."
    )
    amount: Optional[str] = Field(
        default=None,
        description="סכום השיק במספרים בלבד (למשל '1500.00'), בלי סימן שקל ובלי פסיקים."
    )
    is_readable: bool = Field(
        default=False,
        description="האם הצילום ברור מספיק כדי לקרוא ממנו את הפרטים בביטחון? "
                    "אם הצילום מטושטש/חתוך/לא שיק - החזר false."
    )


# ההנחיה הקבועה ל-Claude. היא זהה עבור כל השיקים, ולכן אפשר "לשמור אותה במטמון"
# (prompt caching) כדי לחסוך בעלות כשמעבדים הרבה שיקים ברצף.
SYSTEM_PROMPT = (
    "אתה מומחה לקריאת שיקים בנקאיים ישראליים מתוך צילומים. "
    "המשימה שלך: לחלץ מהצילום את הפרטים המבוקשים בדייקנות מרבית. "
    "קרא בעיון גם טקסט מודפס וגם כתב יד. "
    "אל תנחש - אם פרט מסוים אינו קריא או אינו מופיע, החזר עבורו null. "
    "החזר את הסכום במספרים בלבד, ותאריכים בפורמט DD/MM/YYYY. "
    "שם המושך הוא שם בעל החשבון שמודפס על השיק (לא שם המוטב/הנפרע). "
    "מספר השיק נמצא בדרך כלל בפינה השמאלית התחתונה של השיק - "
    "זהו מספר השיק עצמו, לא מספר חשבון ולא אסמכתא. "
    "אם מודפס על השיק מספר ח.פ / עוסק מורשה (מספר בן 9 ספרות, "
    "בדרך כלל ליד שם החברה) - חלץ גם אותו. "
    "חשוב מאוד לגבי תאריך הפירעון: בישראל מבנה התאריך הוא תמיד "
    "משמאל לימין - המספר השמאלי ביותר הוא היום, האמצעי הוא החודש, "
    "והימני ביותר הוא השנה. אין לכך חריגים. "
    "תאריכים שנכתבו בכתב יד מופיעים לעיתים מקוצרים (למשל 9/6/26); "
    "השלם אותם לפורמט DD/MM/YYYY - רפד יום וחודש לשתי ספרות, "
    "והרחב שנה דו-ספרתית לשנה מלאה בת ארבע ספרות (26 -> 2026)."
)


def normalize_due_date(raw: Optional[str]) -> Optional[str]:
    """
    רשת ביטחון לתאריך הפירעון: מבטיחה פורמט DD/MM/YYYY עקבי.

    בישראל מבנה התאריך הוא תמיד משמאל לימין - יום / חודש / שנה,
    ואין לכך חריגים. הפונקציה הזו:
      * מרפדת יום וחודש חד-ספרתיים לשתי ספרות (9/6 -> 09/06).
      * מרחיבה שנה דו-ספרתית לארבע ספרות (26 -> 2026).
      * תומכת במפרידים /, -, . (לפעמים מעורבים).

    אם אי אפשר לפענח את המבנה (לא בדיוק שלושה חלקים מספריים) -
    מחזירים את הערך המקורי כמות שהוא, בלי לנחש.
    """
    if not raw:
        return raw
    text = str(raw).strip()
    if not text:
        return raw
    # אם הודבקה שעה (למשל "09/06/2026 00:00:00") - לוקחים רק את החלק הראשון.
    text = text.split(" ")[0]
    # מפצלים לפי כל מפריד נפוץ. דורשים בדיוק שלושה חלקים, כולם ספרות.
    parts = re.split(r"[/\-.]", text)
    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        return raw

    day, month, year = parts  # תמיד משמאל לימין: יום, חודש, שנה.

    # שנה דו-ספרתית -> מאה ה-21 (26 -> 2026). שנה חד-ספרתית נדירה -> גם כן.
    if len(year) <= 2:
        year = f"20{int(year):02d}"

    # רפד יום וחודש לשתי ספרות. אימות בסיסי לטווח חוקי.
    try:
        d, m = int(day), int(month)
    except ValueError:
        return raw
    if not (1 <= d <= 31 and 1 <= m <= 12):
        return raw

    return f"{d:02d}/{m:02d}/{year}"


class CheckReader:
    """עוטף את החיבור ל-Claude וקורא שיקים."""

    def __init__(self):
        # הלקוח קורא את המפתח אוטומטית ממשתנה הסביבה ANTHROPIC_API_KEY
        # (שנטען מקובץ ה-.env ב-main). אנחנו לא שותלים את המפתח בקוד.
        self.client = anthropic.Anthropic()

    def read_check(self, image_path: str) -> CheckDetails:
        """
        מקבל נתיב לקובץ תמונה של שיק, ומחזיר אובייקט CheckDetails מלא.
        אם הקריאה נכשלת - מחזיר אובייקט עם is_readable=False.
        """
        # שלב 1: קוראים את קובץ התמונה וממירים אותו לפורמט שאפשר לשלוח ב-API.
        with open(image_path, "rb") as f:
            image_bytes = f.read()
        image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

        # מזהים את סוג התמונה לפי הסיומת (png / jpeg).
        media_type = "image/png"
        lower = image_path.lower()
        if lower.endswith(".jpg") or lower.endswith(".jpeg"):
            media_type = "image/jpeg"

        # שלב 2: שולחים ל-Claude את התמונה + בקשה למלא את ה"טופס".
        try:
            response = self.client.messages.parse(
                model=config.CLAUDE_MODEL,
                max_tokens=1024,
                system=[
                    {
                        "type": "text",
                        "text": SYSTEM_PROMPT,
                        # שומר את ההנחיה במטמון - חוסך בעלות על שיקים הבאים.
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": image_b64,
                                },
                            },
                            {
                                "type": "text",
                                "text": "חלץ את כל פרטי השיק מהצילום הזה.",
                            },
                        ],
                    }
                ],
                # ה"טופס" שהגדרנו למעלה - מכריח את Claude להחזיר בדיוק את השדות האלה.
                output_format=CheckDetails,
            )

            # response.parsed_output הוא אובייקט CheckDetails מוכן לשימוש.
            if response.parsed_output is not None:
                parsed = response.parsed_output
                # רשת ביטחון: מנרמלים את תאריך הפירעון ל-DD/MM/YYYY עקבי,
                # כולל הרחבת שנה דו-ספרתית ורפוד יום/חודש (9/6/26 -> 09/06/2026).
                parsed.due_date = normalize_due_date(parsed.due_date)
                return parsed

            # מקרה נדיר: Claude לא הצליח להחזיר טופס תקין.
            return CheckDetails(is_readable=False)

        except Exception as e:
            # אם הייתה שגיאת רשת / API - לא מפילים את כל הכלי,
            # רק מסמנים שהשיק הזה לא נקרא, וממשיכים הלאה.
            print(f"      ⚠️  שגיאה בקריאת השיק {image_path}: {e}")
            return CheckDetails(is_readable=False)
