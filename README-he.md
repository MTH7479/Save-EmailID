# Copy Email ID – תוסף ל-Outlook

תוסף (Add-in) שמוסיף כפתור ברצועת הכלים של Outlook. בלחיצה נפתח חלון צד עם ה-ID של המייל הנבחר וכפתור העתקה ללוח – אפשר להדביק אותו בכל מקום (Power Automate, Dataverse, Excel וכו').

## הקבצים
| קובץ | תפקיד |
|---|---|
| `manifest.xml` | הגדרת התוסף (כפתור, הרשאות, כתובות) |
| `taskpane.html` | ממשק חלון הצד |
| `taskpane.js` | הלוגיקה: קריאת ה-ID והעתקה ללוח |

## אילו מזהים מוצגים
1. **Graph / REST Id** – זה המזהה שמתאים ל-Power Automate / Microsoft Graph (`messageId`). מומלץ.
2. **EWS ItemId** – המזהה הגולמי של Exchange (`Office.context.mailbox.item.itemId`).
3. **Internet Message-Id** – מזהה הדואר האוניברסלי (`<...@...>`), שימושי לחיפוש בין תיבות.

## שלבי הפעלה
1. **אחסון**: התוסף חייב להיות מאוחסן ב-HTTPS. אפשרויות: GitHub Pages, Azure Static Web Apps, או SharePoint/כל שרת HTTPS.
   העלו את `taskpane.html` ו-`taskpane.js` (ואת קבצי האייקונים) לאותה כתובת.
2. **עדכון ה-manifest**: החליפו בכל הקובץ את `YOUR-HOST` בכתובת האמיתית, למשל `contoso.github.io/copyemailid`.
3. **אייקונים**: הוסיפו `assets/icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png` (PNG פשוט מספיק).
4. **התקנה (Sideload)**:
   - Outlook באינטרנט / New Outlook: הגדרות → Add-ins → My add-ins → Add a custom add-in → Add from file → בחרו את `manifest.xml`.
   - Outlook לשולחן העבודה (Windows): Home → Get Add-ins → My add-ins → Add a custom add-in → Add from file.
   - לפריסה לכל הארגון: Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps.
5. פתחו מייל → בלשונית Home יופיע הכפתור **Copy Email ID** → לחיצה פותחת את חלון הצד → כפתור העתקה.

## הערות
- ההרשאה המבוקשת היא `ReadItem` בלבד – התוסף לא קורא ולא שולח שום תוכן החוצה, הכל רץ מקומית בדפדפן/בקליינט.
- אם ההעתקה האוטומטית נחסמת על ידי הדפדפן, הטקסט מוצג בתיבה וניתן לסמן ולהקיש Ctrl+C.
- ה-Graph Id תלוי בתיבת הדואר: מזהה שנלקח מתיבה אחת לא יעבוד מול תיבה אחרת.
- לבדיקה מקומית אפשר להריץ `npx http-server -S` או להשתמש ב-`office-addin-debugging`.
