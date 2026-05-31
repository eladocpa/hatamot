# -*- coding: utf-8 -*-
"""
==================================================================
  התאמת לקוחות (ה"שדכן" של הכלי)
==================================================================
המודול הזה טוען את רשימת הלקוחות שלך מקובץ האקסל,
ולכל שם שזוהה על שיק - מנסה למצוא את הלקוח המתאים ביותר.

הכלל החשוב ביותר: לא לנחש!
אם אין התאמה ברורה, או שיש כמה לקוחות דומים -
מחזירים "דורש בדיקה ידנית" ונותנים לך להחליט.
==================================================================
"""

from dataclasses import dataclass
from typing import List, Optional

import openpyxl
from thefuzz import fuzz, process

import config


@dataclass
class Customer:
    """לקוח בודד מהרשימה שלך."""
    name: str
    customer_id: str


@dataclass
class MatchResult:
    """תוצאת ההתאמה של שם שיק ללקוח."""
    matched_name: Optional[str]      # שם הלקוח שנמצא ברשימה (או None)
    matched_id: Optional[str]        # מזהה הלקוח שנמצא (או None)
    confidence: int                  # רמת ודאות 0-100
    needs_review: bool               # האם השורה דורשת בדיקה ידנית?
    reason: str                      # הסבר קצר (למה ודאי / למה לבדיקה)


def load_customers(file_path: str = config.CUSTOMERS_FILE) -> List[Customer]:
    """
    טוען את רשימת הלקוחות מקובץ האקסל.
    מצפה לעמודות בשמות שהוגדרו ב-config (ברירת מחדל: 'שם' ו'מזהה').
    """
    workbook = openpyxl.load_workbook(file_path, read_only=True)
    sheet = workbook.active

    # שלב 1: מוצאים באיזו עמודה נמצא השם ובאיזו המזהה,
    # לפי הכותרות בשורה הראשונה.
    header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    name_col_index = None
    id_col_index = None
    for index, cell_value in enumerate(header_row):
        if cell_value == config.CUSTOMERS_NAME_COLUMN:
            name_col_index = index
        elif cell_value == config.CUSTOMERS_ID_COLUMN:
            id_col_index = index

    if name_col_index is None or id_col_index is None:
        raise ValueError(
            f"לא נמצאו העמודות '{config.CUSTOMERS_NAME_COLUMN}' ו-"
            f"'{config.CUSTOMERS_ID_COLUMN}' בקובץ הלקוחות. "
            f"בדוק שהכותרות בשורה הראשונה כתובות בדיוק כך."
        )

    # שלב 2: קוראים את כל השורות (החל מהשנייה) ובונים רשימת לקוחות.
    customers: List[Customer] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[name_col_index]
        customer_id = row[id_col_index]
        if name is None or str(name).strip() == "":
            continue  # מדלגים על שורות ריקות
        customers.append(
            Customer(name=str(name).strip(), customer_id=str(customer_id).strip())
        )

    workbook.close()
    return customers


def _normalize(text: str) -> str:
    """
    מנקה שם לפני השוואה: מוריד רווחים מיותרים ותווים שלא משנים,
    כדי שהשוואה תהיה הוגנת (למשל 'בע״מ' מול 'בעמ').
    """
    text = text.strip()
    for ch in ['"', "'", "״", "׳", ".", ",", "-"]:
        text = text.replace(ch, " ")
    # מאחדים רווחים כפולים לרווח אחד
    return " ".join(text.split())


class CustomerMatcher:
    """מתאים שמות שזוהו על שיקים ללקוחות ברשימה."""

    def __init__(self, customers: List[Customer]):
        self.customers = customers
        # מילון מ"שם מנורמל" -> לקוח, לשליפה מהירה אחרי ההתאמה.
        self._normalized_map = {_normalize(c.name): c for c in customers}
        self._normalized_names = list(self._normalized_map.keys())

    def match(self, detected_name: Optional[str]) -> MatchResult:
        """
        מקבל שם שזוהה על השיק, ומחזיר את ההתאמה הטובה ביותר.
        מחיל את הכלל: בספק - לבדיקה ידנית.
        """
        # מקרה 1: לא זוהה שם בכלל על השיק.
        if not detected_name or detected_name.strip() == "":
            return MatchResult(
                matched_name=None, matched_id=None, confidence=0,
                needs_review=True, reason="לא זוהה שם על השיק",
            )

        normalized_detected = _normalize(detected_name)

        # שלב 1: מבקשים מ-thefuzz את שני הלקוחות הכי דומים, עם ציון לכל אחד.
        # extract מחזיר רשימה של (שם_מנורמל, ציון).
        top_matches = process.extract(
            normalized_detected,
            self._normalized_names,
            scorer=fuzz.token_sort_ratio,  # סדר המילים לא משנה ("אבי כהן" = "כהן אבי")
            limit=2,
        )

        if not top_matches:
            return MatchResult(
                matched_name=None, matched_id=None, confidence=0,
                needs_review=True, reason="אין לקוחות ברשימה",
            )

        best_name, best_score = top_matches[0]
        best_customer = self._normalized_map[best_name]

        # שלב 2: בודקים אם יש לקוח שני קרוב מדי (עמימות).
        second_score = top_matches[1][1] if len(top_matches) > 1 else 0
        gap = best_score - second_score

        # מקרה 2: ההתאמה חלשה מדי -> בדיקה ידנית.
        if best_score < config.MATCH_CONFIDENCE_THRESHOLD:
            return MatchResult(
                matched_name=best_customer.name,
                matched_id=best_customer.customer_id,
                confidence=best_score,
                needs_review=True,
                reason=f"התאמה חלשה ({best_score}%)",
            )

        # מקרה 3: שני לקוחות דומים מדי -> לא מנחשים, בדיקה ידנית.
        if gap < config.MATCH_AMBIGUITY_GAP and second_score >= config.MATCH_CONFIDENCE_THRESHOLD:
            return MatchResult(
                matched_name=best_customer.name,
                matched_id=best_customer.customer_id,
                confidence=best_score,
                needs_review=True,
                reason=f"כמה לקוחות דומים (פער {gap}% בלבד)",
            )

        # מקרה 4: התאמה ודאית!
        return MatchResult(
            matched_name=best_customer.name,
            matched_id=best_customer.customer_id,
            confidence=best_score,
            needs_review=False,
            reason="התאמה ודאית",
        )
