# -*- coding: utf-8 -*-
"""
==================================================================
  הוצאת קבלות דרך ה-API של Maven
==================================================================
>>> שלב רגיש! הוצאת קבלה משנה נתונים אמיתיים. <<<

זרימת העבודה:
  1. main.py / process_only.py יוצרים את 'תוצאות_שיקים.xlsx'.
  2. אתה פותח את הקובץ, בודק, ומסמן "כן" בעמודת "לאשר?"
     בשורות שברצונך להוציא להן קבלה (כברירת מחדל כבר מסומן "כן"
     לכל שורה "מוכן לקבלה" - אתה יכול למחוק/לשנות).
  3. מריצים:  python issue_receipts.py
     הכלי קורא רק את השורות המסומנות "כן", ומוציא להן קבלה דרך ה-API.

הגנות מובנות:
  - מוציאים קבלה רק לשורות עם "כן" מפורש בעמודת "לאשר?".
  - לעולם לא מוציאים קבלה לשיק שחזר (גם אם סומן "כן" בטעות).
  - מצב "הרצה יבשה" (DRY RUN): מראה מה *היה* קורה בלי לשלוח כלום.
  - יומן (receipts_log.csv) מתעד כל קבלה שהוצאה, למניעת כפילויות.
==================================================================
"""

import csv
import os
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import openpyxl

import config
from customer_matcher import load_customers
from maven_api import MavenReceiptClient, ReceiptRequest


# שם עמודת האישור בקובץ האקסל, והערך שמסמן אישור.
APPROVE_COLUMN = "לאשר?"
APPROVE_VALUE = "כן"
STATUS_BOUNCED = "שיק שחזר - לא להוציא קבלה"

# קובץ יומן: מתעד אילו קבלות כבר הוצאו (כדי לא להוציא פעמיים).
RECEIPTS_LOG = "receipts_log.csv"


@dataclass
class ApprovedRow:
    """שורה שאושרה להוצאת קבלה, עם הפרטים הדרושים."""
    row_number: int                  # מספר השורה באקסל (לתיעוד)
    matched_id: Optional[str]        # מזהה הלקוח ב-Maven
    matched_name: Optional[str]
    check_number: Optional[str]
    bank_name: Optional[str]
    branch_number: Optional[str]
    account_number: Optional[str]
    due_date: Optional[str]
    amount: Optional[str]
    maven_reference: Optional[str]
    status: str


def _read_approved_rows(file_path: str) -> List[ApprovedRow]:
    """קורא מהאקסל את כל השורות שסומנו 'כן' בעמודת 'לאשר?'."""
    workbook = openpyxl.load_workbook(file_path, read_only=True)
    sheet = workbook.active

    # ממפים שם עמודה -> אינדקס, לפי שורת הכותרות.
    header = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    col = {name: idx for idx, name in enumerate(header)}

    def get(row, name):
        idx = col.get(name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    approved: List[ApprovedRow] = []
    for row_num, row in enumerate(
        sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        approve = get(row, APPROVE_COLUMN)
        if approve is None or str(approve).strip() != APPROVE_VALUE:
            continue  # רק שורות שסומנו "כן"

        approved.append(ApprovedRow(
            row_number=row_num,
            matched_id=None,  # מזהה הלקוח ב-Maven נשלף בהמשך לפי השם
            matched_name=_str(get(row, "לקוח מותאם ברשימה")),
            check_number=_str(get(row, "מספר שיק")),
            bank_name=_str(get(row, "בנק")),
            branch_number=_str(get(row, "סניף")),
            account_number=_str(get(row, "חשבון")),
            due_date=_str(get(row, "תאריך פירעון")),
            amount=_str(get(row, "סכום")),
            maven_reference=_str(get(row, "אסמכתא Maven")),
            status=_str(get(row, "סטטוס")) or "",
        ))

    workbook.close()
    return approved


def _str(value) -> Optional[str]:
    """המרה בטוחה למחרוזת (None נשאר None)."""
    if value is None:
        return None
    return str(value).strip()


def _load_issued_references() -> set:
    """טוען מהיומן את האסמכתאות שכבר הוצאה להן קבלה (מניעת כפילויות)."""
    issued = set()
    if not os.path.exists(RECEIPTS_LOG):
        return issued
    with open(RECEIPTS_LOG, "r", encoding="utf-8", newline="") as f:
        for record in csv.DictReader(f):
            ref = record.get("maven_reference")
            if ref:
                issued.add(ref)
    return issued


def _append_to_log(row: ApprovedRow, receipt_id: str, dry_run: bool):
    """מוסיף שורה ליומן הקבלות."""
    is_new = not os.path.exists(RECEIPTS_LOG)
    with open(RECEIPTS_LOG, "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow([
                "timestamp", "maven_reference", "check_number",
                "customer", "amount", "receipt_id", "dry_run",
            ])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            row.maven_reference or "", row.check_number or "",
            row.matched_name or "", row.amount or "",
            receipt_id, "כן" if dry_run else "לא",
        ])


def issue_receipts(file_path: str = config.OUTPUT_FILE, dry_run: bool = True):
    """
    מוציא קבלות לכל השורות שסומנו 'כן' בקובץ האקסל.
    dry_run=True (ברירת מחדל) = הרצה יבשה: מראה מה היה קורה, בלי לשלוח.
    dry_run=False = הוצאה אמיתית של הקבלות.
    """
    mode = "🧪 הרצה יבשה (DRY RUN - לא נשלח כלום)" if dry_run \
        else "🔴 הוצאה אמיתית של קבלות"
    print("=" * 60)
    print(f"  הוצאת קבלות — {mode}")
    print("=" * 60)

    if not os.path.exists(file_path):
        print(f"\n❌ לא נמצא קובץ התוצאות '{file_path}'.")
        print("   הרץ קודם את main.py או process_only.py.")
        return

    approved = _read_approved_rows(file_path)
    if not approved:
        print(f"\nℹ️  לא נמצאו שורות מסומנות '{APPROVE_VALUE}' בעמודת '{APPROVE_COLUMN}'.")
        print("   פתח את הקובץ, סמן 'כן' בשורות הרצויות, ושמור.")
        return

    print(f"\nנמצאו {len(approved)} שורות מאושרות.")

    # בונים מילון שם-לקוח -> מזהה Maven, מתוך קובץ הלקוחות,
    # כדי שנשלח ל-API את מזהה הלקוח הנכון (ולא רק את השם).
    name_to_id = {}
    try:
        for c in load_customers():
            name_to_id[c.name.strip()] = c.customer_id
    except Exception as e:
        print(f"⚠️  לא הצלחתי לטעון את קובץ הלקוחות למזהים: {e}")

    issued_refs = _load_issued_references()
    client = MavenReceiptClient()
    success, skipped, failed = 0, 0, 0

    for row in approved:
        label = f"שורה {row.row_number}: {row.matched_name or '?'} " \
                f"(שיק {row.check_number or '?'}, ₪{row.amount or '?'})"

        # הגנה 1: לעולם לא מוציאים קבלה לשיק שחזר.
        if row.status == STATUS_BOUNCED:
            print(f"  ⛔ דילוג - שיק שחזר: {label}")
            skipped += 1
            continue

        # הגנה 2: לא מוציאים פעמיים לאותה אסמכתא.
        if row.maven_reference and row.maven_reference in issued_refs:
            print(f"  ⏭️  דילוג - כבר הוצאה קבלה: {label}")
            skipped += 1
            continue

        # הגנה 3: חייב להיות לקוח מזוהה וסכום.
        if not row.matched_name or not row.amount:
            print(f"  ⚠️  דילוג - חסר לקוח/סכום: {label}")
            skipped += 1
            continue

        # שולפים את מזהה הלקוח ב-Maven לפי השם (מקובץ הלקוחות).
        customer_id = name_to_id.get((row.matched_name or "").strip())
        if not customer_id:
            print(f"  ⚠️  דילוג - לא נמצא מזהה Maven ללקוח: {label}")
            skipped += 1
            continue

        request = ReceiptRequest(
            customer_name=row.matched_name,
            customer_id=customer_id,
            amount=row.amount,
            check_number=row.check_number,
            bank_name=row.bank_name,
            branch_number=row.branch_number,
            account_number=row.account_number,
            due_date=row.due_date,
            maven_reference=row.maven_reference,
        )

        try:
            if dry_run:
                print(f"  🧪 [יבש] היה מוציא קבלה: {label}")
                receipt_id = "DRY_RUN"
            else:
                receipt_id = client.create_receipt(request)
                print(f"  ✅ הוצאה קבלה ({receipt_id}): {label}")
            _append_to_log(row, receipt_id, dry_run)
            if row.maven_reference:
                issued_refs.add(row.maven_reference)
            success += 1
        except Exception as e:
            print(f"  ❌ נכשל: {label}\n       {e}")
            failed += 1

    print("\n" + "=" * 60)
    print(f"  סיכום: {success} {'(יבש)' if dry_run else 'הוצאו'}, "
          f"{skipped} דולגו, {failed} נכשלו.")
    print("=" * 60)
    if dry_run:
        print("\n💡 זו הייתה הרצה יבשה. כדי להוציא קבלות באמת, הרץ:")
        print("      python issue_receipts.py --real")


if __name__ == "__main__":
    import sys
    # ברירת מחדל: הרצה יבשה. רק עם הדגל --real מוציאים קבלות אמיתיות.
    real = "--real" in sys.argv
    issue_receipts(dry_run=not real)
