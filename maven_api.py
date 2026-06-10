# -*- coding: utf-8 -*-
"""
==================================================================
  מתאם (Adapter) ל-API של Maven להוצאת קבלות
==================================================================
מבוסס על התיעוד הרשמי:
https://www.invoice-maven.co.il/support/api/add-document/

  Endpoint:  POST https://app.invoice-maven.co.il/api/documents/addDocument
  אימות:     שדה api_key בגוף הבקשה (JSON)
  קבלה:      doc_type = 400
  תשלום בשיק: payment_type = 3
  מצב בדיקה: test = 1 (בדיקה, לא נוצרת קבלה אמיתית) / 0 (אמיתי)

>>> מפתח API לכל לקוח מיוצג <<<
לכל חברה מיוצגת יש מפתח API משלה. שומרים אותם בקובץ .env, ומעבירים
לכלי איזה מפתח להשתמש בהרצה הנוכחית (לפי הלקוח המיוצג שאתה מעבד).
==================================================================
"""

import os
from dataclasses import dataclass
from typing import Optional

import requests

import config


# כתובת ה-API להוספת מסמך (קבלה).
ADD_DOCUMENT_URL = "https://app.invoice-maven.co.il/api/documents/addDocument"

# קודים מהתיעוד.
DOC_TYPE_RECEIPT = 400      # קבלה
PAYMENT_TYPE_CHECK = 3      # תשלום בהמחאה / שיק


@dataclass
class ReceiptRequest:
    """כל הפרטים הדרושים להוצאת קבלה אחת על תשלום בשיק."""
    customer_name: str
    customer_id: Optional[str]        # מזהה הלקוח אצלך (לא נשלח ל-API ישירות)
    amount: str
    check_number: Optional[str]
    bank_name: Optional[str]
    branch_number: Optional[str]
    account_number: Optional[str]
    due_date: Optional[str]           # תאריך פירעון השיק (לתיעוד)
    maven_reference: Optional[str]    # אסמכתא - תשמש כ-doc_id למניעת כפילויות
    company_id: Optional[str] = None  # ח.פ הלקוח (לשדה identification)
    payment_date: Optional[str] = None  # תאריך הקבלה (תאריך ההפקדה מדף הבנק)


def _clean_amount(amount: str) -> float:
    """ממיר סכום מחרוזת (אולי עם ₪/פסיקים) למספר."""
    text = str(amount).replace("₪", "").replace(",", "").strip()
    return float(text)


class MavenReceiptClient:
    """
    לקוח ל-API של Maven להוצאת קבלות.

    api_key: מפתח ה-API של הלקוח המיוצג הספציפי שמעבדים.
             אם לא מועבר במפורש, נטען ממשתנה הסביבה MAVEN_API_KEY.
    test_mode: True = הבקשה נשלחת עם test=1 (Maven בודק אך לא יוצר קבלה).
    """

    def __init__(self, api_key: Optional[str] = None, test_mode: bool = True):
        self.api_key = (api_key or os.environ.get("MAVEN_API_KEY", "")).strip()
        self.test_mode = test_mode
        # פרטי קשר אופציונליים לבקשה (אפשר להגדיר ב-.env, לא חובה).
        self.contact_email = os.environ.get("MAVEN_CONTACT_EMAIL", "").strip()
        self.contact_phone = os.environ.get("MAVEN_CONTACT_PHONE", "").strip()

    def is_configured(self) -> bool:
        """האם יש מפתח API להשתמש בו?"""
        return bool(self.api_key)

    def create_receipt(self, request: ReceiptRequest) -> str:
        """
        מוציא קבלה אחת ב-Maven על תשלום בשיק.
        מחזיר את מספר המסמך (doc_no) שנוצר.
        זורק חריגה עם הסבר אם השרת החזיר שגיאה.
        """
        if not self.api_key:
            raise RuntimeError(
                "אין מפתח API. הגדר MAVEN_API_KEY בקובץ .env "
                "(המפתח של הלקוח המיוצג שאתה מעבד)."
            )

        # תאריך הקבלה = תאריך ההפקדה מדף הבנק. אם משום מה חסר -
        # נופלים חזרה לתאריך הפירעון, כדי לא לאבד הוצאת קבלה.
        payment_date = request.payment_date or request.due_date

        # payment_date הוא שדה חובה. אם חסר לגמרי - עוצרים עם הסבר ברור,
        # במקום לשלוח ערך ריק שייכשל בשרת.
        if not payment_date:
            raise RuntimeError(
                "חסר תאריך תשלום (payment_date). לא נמצא תאריך הפקדה "
                "ולא תאריך פירעון. מלא תאריך הפקדה בקובץ התוצאות והרץ שוב."
            )

        # בונים את גוף הבקשה לפי התיעוד. כל השדות ב-lowercase.
        payload = {
            "api_key": self.api_key,
            "test": 1 if self.test_mode else 0,
            "doc_type": DOC_TYPE_RECEIPT,
            "customer": {
                "name": request.customer_name,
                "save_customer": 0,  # לא יוצרים לקוח חדש - הוא כבר קיים
            },
            "payments": [
                {
                    "payment_date": payment_date or "",
                    "payment_type": PAYMENT_TYPE_CHECK,
                    "amount": _clean_amount(request.amount),
                    "bank": request.bank_name or "",
                    "bank_branch": request.branch_number or "",
                    "bank_account": request.account_number or "",
                    "cheque_number": request.check_number or "",
                }
            ],
        }

        # ח.פ הלקוח (אם זוהה) -> שדה identification.
        if request.company_id:
            payload["customer"]["identification"] = request.company_id

        # אסמכתא ייחודית -> doc_id, כדי שגם השרת ימנע כפילויות.
        if request.maven_reference:
            payload["doc_id"] = str(request.maven_reference)

        # פרטי קשר אופציונליים, אם הוגדרו.
        if self.contact_email:
            payload["contact_email"] = self.contact_email
        if self.contact_phone:
            payload["contact_phone"] = self.contact_phone

        # שולחים את הבקשה.
        response = requests.post(
            ADD_DOCUMENT_URL,
            json=payload,
            headers={"Content-Type": "application/json; charset=utf-8"},
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()

        # לפי התיעוד: status_code=0 פירושו הצלחה.
        # ה-API עשוי להחזיר את הקוד כמספר (0) או כמחרוזת ("0") - מטפלים בשניהם.
        status_code = data.get("status_code")
        is_success = str(status_code).strip() == "0"
        if not is_success:
            description = data.get("status_description", "שגיאה לא ידועה")
            raise RuntimeError(
                f"Maven החזיר שגיאה (status_code={status_code}): {description}"
            )

        # מחזירים את מספר המסמך שנוצר.
        return str(data.get("doc_no", "?"))
