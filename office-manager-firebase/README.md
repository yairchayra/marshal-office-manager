# Office Manager Backend — Firebase Edition

Express.js API server using **Firebase Firestore** instead of PostgreSQL.

## Firestore Collections

| Collection | Description |
|---|---|
| `pcs` | מחשבים + נתוני חומרה + AutoCAD |
| `desks` | שולחנות על מפת הקומה |
| `app_accounts` | חשבונות כניסה לאפליקציה |
| `app_settings` | הגדרות (table_groups, floor_plan) |

---

## 1. הגדרת Firebase

1. לך ל-[Firebase Console](https://console.firebase.google.com/)
2. צור פרויקט חדש (או בחר קיים)
3. הפעל **Firestore Database** → Start in production mode
4. לך ל-**Project Settings → Service Accounts**
5. לחץ **Generate new private key** → הורד את הקובץ
6. שמור אותו בתיקיית הפרויקט בשם `serviceAccountKey.json`

---

## 2. התקנה מקומית

```bash
# שכפל את הריפו
git clone https://github.com/YOUR_USERNAME/marshal-office-manager.git
cd marshal-office-manager

# התקן תלויות
npm install

# הגדר סביבה
cp .env.example .env
# ערוך את .env לפי הצורך
```

---

## 3. הרצת שרת Firebase

```bash
npm start
# או בפיתוח:
npm run dev
```

---

## 4. מיגרציה מ-PostgreSQL ל-Firebase

> מריץ את הסקריפט **פעם אחת בלבד** — הוא שואב את כל הנתונים מ-Postgres ומעלה ל-Firestore.

**דרישות:**
- PostgreSQL פועל עם הנתונים הישנים
- `.env` מוגדר עם `DB_*` (Postgres) + `FIREBASE_SERVICE_ACCOUNT_PATH`

```bash
node migrate.js
```

הסקריפט ימגרר:
- ✅ כל המחשבים (`pcs` + `users`)
- ✅ כל השולחנות (`desks`)
- ✅ חשבונות אפליקציה (`app_accounts`)
- ✅ הגדרות (`app_settings`)
- ✅ מפת הקומה (`floor_plan`)

אחרי המיגרציה אפשר להסיר את `DB_*` מה-`.env`.

---

## 5. העלאה ל-GitHub

```bash
# בתיקיית הפרויקט
git init
git add .
git commit -m "Initial commit — Firebase backend"

# צור ריפו חדש ב-GitHub (ללא README):
# https://github.com/new

# קשר וידחף
git remote add origin https://github.com/YOUR_USERNAME/marshal-office-manager.git
git branch -M main
git push -u origin main
```

> ⚠️ `serviceAccountKey.json` ו-`.env` נמצאים ב-`.gitignore` ולא יועלו ל-GitHub.

---

## API Endpoints

כל הנתיבים זהים לגרסה הישנה (PostgreSQL) — אין צורך לשנות את ה-Frontend.

| Method | Path | Auth |
|---|---|---|
| POST | `/api/auth/login` | פתוח |
| GET | `/api/pcs` | משתמש |
| POST | `/api/pcs/bulk` | אדמין |
| GET/POST/PUT/DELETE | `/api/desks` | משתמש/אדמין |
| GET/POST | `/api/tablegroups` | משתמש/אדמין |
| GET/POST | `/api/floorplan` | משתמש/אדמין |
| GET/POST/PUT/DELETE | `/api/accounts` | אדמין |
