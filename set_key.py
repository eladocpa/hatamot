# -*- coding: utf-8 -*-
"""
==================================================================
  הגדרת מפתח ה-API בקלות ובבטחה
==================================================================
מריצים את הקובץ הזה, מדביקים את המפתח, וזהו - הוא כותב קובץ .env
תקין בשבילך. בלי Notepad, בלי בעיית .env.txt, בלי טקסט דוגמה תקוע.

הרצה:   python set_key.py
==================================================================
"""


def main():
    print("=" * 60)
    print("  🔑  הגדרת מפתח ה-API של Claude")
    print("=" * 60)
    print("\nהדבק כאן את המפתח (מתחיל ב-sk-ant-...) ולחץ Enter:")
    print("(משיגים אותו ב: https://console.anthropic.com -> API Keys)\n")

    key = input("מפתח: ").strip()

    # ניקוי: מורידים מירכאות, רווחים, וטקסט דוגמה אם נדבק בטעות.
    key = key.strip('"').strip("'").strip()
    if "הדבק" in key:
        key = key.split("הדבק")[0].rstrip("- ").strip()

    # בדיקה בסיסית שהמפתח נראה תקין.
    if not key.startswith("sk-ant-"):
        print("\n❌ זה לא נראה כמו מפתח תקין (אמור להתחיל ב-sk-ant-).")
        print("   נסה שוב, והקפד להעתיק את כל המפתח מ-console.anthropic.com.")
        return

    # כותבים קובץ .env נקי (תמיד בשם הנכון, בלי .txt).
    with open(".env", "w", encoding="utf-8") as f:
        f.write(f"ANTHROPIC_API_KEY={key}\n")

    preview = key[:14] + "..." + key[-4:]
    print(f"\n✅ נשמר קובץ .env תקין עם המפתח: {preview}")
    print("   עכשיו אפשר להריץ:  python process_only.py")


if __name__ == "__main__":
    main()
