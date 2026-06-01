# -*- coding: utf-8 -*-
"""
==================================================================
  כתיבת קובץ האקסל (ה"מזכיר" של הכלי)
==================================================================
מקבל את כל השורות שעובדו ובונה מהן קובץ אקסל מסודר
עם כל העמודות שביקשת, כולל צביעה לפי סטטוס:
ירוק = מוכן לקבלה, צהוב = דורש בדיקה ידנית.
==================================================================
"""

from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import config


@dataclass
class ResultRow:
    """שורה אחת בטבלת הפלט - שיק אחד מעובד."""
    detected_name: Optional[str]   # שם שזוהה על השיק
    matched_name: Optional[str]    # לקוח מותאם ברשימה
    confidence: int                # רמת ודאות (0-100)
    check_number: Optional[str]
    bank_name: Optional[str]
    branch_number: Optional[str]
    account_number: Optional[str]
    due_date: Optional[str]
    amount: Optional[str]
    status: str                    # מוכן / בדיקה ידנית / שיק שחזר
    maven_reference: Optional[str] # מספר האסמכתא מהשורה ב-Maven (88635 וכו')
    image_path: Optional[str]      # נתיב לצילום (לתיעוד)
    company_id: Optional[str] = None       # ח.פ שזוהה על השיק
    evidence: Optional[str] = None         # על מה התבססה ההתאמה (שם/ח.פ/סכום)
    matched_customer_id: Optional[str] = None   # מספר הלקוח במערכת (מהרשימה)
    matched_company_id: Optional[str] = None    # ח.פ הלקוח מהרשימה (להשוואה)


# הכותרות של העמודות, בסדר שביקשת.
HEADERS = [
    "לאשר?",            # <<< אתה ממלא: כתוב "כן" בשורות שברצונך להוציא להן קבלה
    "שם לקוח שזוהה",
    "לקוח מותאם ברשימה",
    "מספר לקוח במערכת",  # המזהה של הלקוח כפי שמופיע ברשימה/Maven
    "רמת ודאות",
    "ח.פ שזוהה בשיק",    # ה-ח.פ שזוהה בצילום השיק
    "ח.פ במערכת",        # ה-ח.פ של הלקוח המותאם, מהרשימה - להשוואה
    "מספר שיק",
    "בנק",
    "סניף",
    "חשבון",
    "תאריך פירעון",
    "סכום",
    "סטטוס",
    "בסיס ההתאמה",
    "אסמכתא Maven",
    "קובץ צילום",
]

# צבעים לסטטוס
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

STATUS_READY = "מוכן לקבלה"
STATUS_REVIEW = "דורש בדיקה ידנית"
STATUS_BOUNCED = "שיק שחזר - לא להוציא קבלה"


def write_results(rows: List[ResultRow], output_file: str = config.OUTPUT_FILE):
    """בונה את קובץ האקסל מתוך רשימת השורות."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "שיקים לבדיקה"
    # הצגה מימין לשמאל (עברית)
    sheet.sheet_view.rightToLeft = True

    # שלב 1: כותבים את שורת הכותרות ומעצבים אותה.
    for col_index, header in enumerate(HEADERS, start=1):
        cell = sheet.cell(row=1, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # שלב 2: כותבים שורה לכל שיק, וצובעים לפי הסטטוס.
    for row_index, r in enumerate(rows, start=2):
        # עמודת "לאשר?": מציעים "כן" כברירת מחדל לשורות מוכנות לקבלה,
        # וריק לשורות שדורשות בדיקה / שיקים שחזרו. אתה יכול לשנות ידנית.
        approve_default = "כן" if r.status == STATUS_READY else ""
        values = [
            approve_default,
            r.detected_name or "",
            r.matched_name or "",
            r.matched_customer_id or "",
            f"{r.confidence}%" if r.confidence else "",
            r.company_id or "",
            r.matched_company_id or "",
            r.check_number or "",
            r.bank_name or "",
            r.branch_number or "",
            r.account_number or "",
            r.due_date or "",
            r.amount or "",
            r.status,
            r.evidence or "",
            r.maven_reference or "",
            r.image_path or "",
        ]
        # צבע לפי הסטטוס: ירוק=מוכן, אדום=שיק שחזר, צהוב=בדיקה ידנית.
        if r.status == STATUS_READY:
            fill = GREEN_FILL
        elif r.status == STATUS_BOUNCED:
            fill = RED_FILL
        else:
            fill = YELLOW_FILL
        for col_index, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_index, column=col_index, value=value)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # שלב 3: מרחיבים את העמודות שיהיה נוח לקרוא.
    # סדר: לאשר, שם שזוהה, לקוח מותאם, מספר לקוח, ודאות, ח.פ בשיק, ח.פ במערכת,
    #      מספר שיק, בנק, סניף, חשבון, תאריך, סכום, סטטוס, בסיס, אסמכתא, צילום
    widths = [8, 22, 22, 14, 9, 13, 13, 12, 12, 7, 13, 13, 11, 20, 24, 13, 22]
    for col_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            openpyxl.utils.get_column_letter(col_index)
        ].width = width

    # שמירה חסינה: אם הקובץ פתוח באקסל (PermissionError) - לא קורסים
    # ולא מאבדים את העבודה. שומרים לקובץ חלופי עם חותמת זמן ומודיעים.
    try:
        workbook.save(output_file)
        print(f"\n📊 קובץ התוצאות נשמר: {output_file}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt_file = output_file.replace(".xlsx", f"_{timestamp}.xlsx")
        workbook.save(alt_file)
        print(f"\n⚠️  הקובץ '{output_file}' היה פתוח (אולי באקסל), אז לא ניתן היה לדרוס אותו.")
        print(f"📊 שמרתי במקום זאת לקובץ חדש: {alt_file}")
        print("   (טיפ: סגור את הקובץ באקסל לפני הרצה כדי שיישמר בשם הרגיל.)")
