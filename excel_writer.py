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
    status: str                    # "מוכן לקבלה" / "דורש בדיקה ידנית"
    maven_reference: Optional[str] # מספר האסמכתא מהשורה ב-Maven (88635 וכו')
    image_path: Optional[str]      # נתיב לצילום (לתיעוד)


# הכותרות של העמודות, בסדר שביקשת.
HEADERS = [
    "שם לקוח שזוהה",
    "לקוח מותאם ברשימה",
    "רמת ודאות",
    "מספר שיק",
    "בנק",
    "סניף",
    "חשבון",
    "תאריך פירעון",
    "סכום",
    "סטטוס",
    "אסמכתא Maven",
    "קובץ צילום",
]

# צבעים לסטטוס
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

STATUS_READY = "מוכן לקבלה"
STATUS_REVIEW = "דורש בדיקה ידנית"


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
        values = [
            r.detected_name or "",
            r.matched_name or "",
            f"{r.confidence}%" if r.confidence else "",
            r.check_number or "",
            r.bank_name or "",
            r.branch_number or "",
            r.account_number or "",
            r.due_date or "",
            r.amount or "",
            r.status,
            r.maven_reference or "",
            r.image_path or "",
        ]
        fill = GREEN_FILL if r.status == STATUS_READY else YELLOW_FILL
        for col_index, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_index, column=col_index, value=value)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # שלב 3: מרחיבים את העמודות שיהיה נוח לקרוא.
    widths = [22, 22, 10, 12, 14, 8, 14, 14, 12, 18, 14, 24]
    for col_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            openpyxl.utils.get_column_letter(col_index)
        ].width = width

    workbook.save(output_file)
    print(f"\n📊 קובץ התוצאות נשמר: {output_file}")
