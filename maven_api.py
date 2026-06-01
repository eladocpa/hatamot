# -*- coding: utf-8 -*-
"""
==================================================================
  מתאם (Adapter) ל-API של Maven להוצאת קבלות
==================================================================
>>> זה הקובץ היחיד שצריך להשלים לפי תיעוד ה-API של Maven. <<<

כל שאר הכלי כבר בנוי ומוכן. כאן נמצא החיבור בפועל ל-Maven:
שליחת בקשה ליצירת קבלה. ברגע שתשלח את תיעוד ה-API, אמלא כאן
את הפרטים המדויקים (כתובת, אימות, מבנה הבקשה).

עד אז - הקובץ זורק שגיאה ברורה שמסבירה מה חסר, כך שאי אפשר
להוציא קבלה בטעות לפני שהחיבור הוגדר נכון.
==================================================================
"""

import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class ReceiptRequest:
    """כל הפרטים הדרושים להוצאת קבלה אחת."""
    customer_name: str
    customer_id: Optional[str]
    amount: str
    check_number: Optional[str]
    bank_name: Optional[str]
    branch_number: Optional[str]
    account_number: Optional[str]
    due_date: Optional[str]
    maven_reference: Optional[str]


class MavenReceiptClient:
    """
    לקוח ל-API של Maven להוצאת קבלות.

    >>> להשלמה לפי התיעוד <<<
    קרא את פרטי ההתחברות ממשתני סביבה (כמו מפתח ה-Claude),
    לעולם לא מהקוד. הוסף ל-.env את מה שצריך, למשל:
        MAVEN_API_TOKEN=...
        MAVEN_API_BASE_URL=https://app.invoice-maven.co.il/api/...
    """

    def __init__(self):
        # נטען את פרטי ההתחברות ל-API מהסביבה (להשלמה לפי התיעוד).
        self.api_token = os.environ.get("MAVEN_API_TOKEN", "").strip()
        self.base_url = os.environ.get("MAVEN_API_BASE_URL", "").strip()

    def create_receipt(self, request: ReceiptRequest) -> str:
        """
        מוציא קבלה אחת ב-Maven ומחזיר את מזהה הקבלה שנוצרה.

        >>> חלק זה ימולא לפי תיעוד ה-API של Maven. <<<
        המבנה הכללי יהיה בערך כך (דוגמה - לא סופי):

            import requests
            response = requests.post(
                f"{self.base_url}/receipts",
                headers={"Authorization": f"Bearer {self.api_token}"},
                json={
                    "customer_id": request.customer_id,
                    "amount": request.amount,
                    "payment": {
                        "type": "check",
                        "check_number": request.check_number,
                        "bank": request.bank_name,
                        "branch": request.branch_number,
                        "account": request.account_number,
                        "due_date": request.due_date,
                    },
                },
                timeout=30,
            )
            response.raise_for_status()
            return str(response.json()["receipt_id"])
        """
        raise NotImplementedError(
            "החיבור ל-API של Maven עדיין לא הוגדר.\n"
            "       שלח את תיעוד ה-API (כתובת, אימות, מבנה בקשת קבלה),\n"
            "       וזה יושלם כאן. עד אז אפשר להריץ רק במצב 'הרצה יבשה'."
        )
