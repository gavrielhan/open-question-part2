# מחולל נושאים MECE אוטומטי | Automatic MECE Topic Generator

יישום אינטרנטי ליצירת נושאים MECE אוטומטית מתשובות פתוחות באמצעות GPT 5.1 דרך Azure.

## תכונות

- **העלאת קבצים**: תמיכה בקבצי Excel (.xlsx, .xls) ו-CSV
- **בחירת עמודה**: בחרו את העמודה המכילה את התשובות הפתוחות
- **סליידר לכמות נושאים**: שליטה מלאה על מספר הנושאים (2-15)
- **יצירת נושאים MECE**: המערכת מנתחת את התשובות ויוצרת נושאים בעברית באופן אוטומטי
- **סיווג אוטומטי**: כל התשובות מסווגות לפי הנושאים שנוצרו

## דרישות מקדימות

- Python 3.9+
- גישה ל-Azure OpenAI עם GPT 5.1 (או מודל תואם)

## התקנה

### Windows Users
**📘 For detailed Windows setup instructions, see [WINDOWS_SETUP_GUIDE.md](WINDOWS_SETUP_GUIDE.md)**

Quick start on Windows:
1. Install Python 3.9+ (make sure to check "Add Python to PATH")
2. Download/clone this repository
3. Double-click `launch_app.bat` (or follow the detailed guide)

### Linux/macOS Users

1. **יצירת סביבה וירטואלית:**
```bash
python3 -m venv venv
source venv/bin/activate  # Linux/macOS
# או
venv\Scripts\activate  # Windows
```

3. **התקנת תלויות:**
```bash
pip install -r requirements.txt
```

4. **הגדרת משתני סביבה:**
```bash
cp .env.example .env
# ערכו את קובץ .env עם הפרטים שלכם
```

## הגדרת Azure OpenAI

ערכו את קובץ `.env`:

```env
OPENAI_API_KEY=your_azure_api_key
OPENAI_API_BASE_URL=https://your-resource.openai.azure.com
MODEL=gpt-5.1
AZURE_API_VERSION=2025-04-01-preview
```

## הפעלה

```bash
python topic_generator_app.py
```

הדפדפן יפתח אוטומטית בכתובת http://127.0.0.1:5001

## שימוש

1. **העלאת קובץ**: גררו קובץ Excel/CSV או לחצו לבחירה
2. **בחירת גיליון**: אם יש מספר גיליונות, בחרו את הרצוי
3. **הגדרות**:
   - הזינו את אות העמודה המכילה את התשובות (לדוגמה: H)
   - הגדירו את מספר הנושאים המקסימלי בסליידר (2-15)
4. **יצירת נושאים**: לחצו על "יצירת נושאים אוטומטית"
5. **סיווג**: לאחר שהנושאים נוצרו, לחצו על "סיווג כל התשובות"

## פלט

הקובץ המסווג יישמר על שולחן העבודה בפורמט:
`[שם_הקובץ_המקורי]_topics_classified_[תאריך_שעה].csv`

הקובץ יכיל:
- את כל העמודות המקוריות
- עמודות חדשות לכל נושא שנוצר (ערכים: 0 או 1)

## מה זה MECE?

MECE = **M**utually **E**xclusive, **C**ollectively **E**xhaustive

- **בלעדיות הדדית**: אין חפיפה בין הנושאים
- **מיצוי**: כל תשובה מכוסה על ידי לפחות נושא אחד

## פורט

האפליקציה רצה על פורט 5001 (כדי לא להתנגש עם האפליקציה המקורית על פורט 5000).

