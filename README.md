# Marshal Office Manager

אפליקציית React המתחברת ישירות ל-Firebase Firestore — ללא שרת, ללא backend.

**🌐 Live:** `https://YOUR_USERNAME.github.io/marshal-office-manager/`

---

## הגדרה ראשונית (פעם אחת)

### 1. Firebase — הכנסת פרטי הפרויקט

לך ל-[Firebase Console](https://console.firebase.google.com) → הפרויקט שלך →
**Project Settings → General → Your apps → SDK setup → Config**

העתק את הערכים לתוך `src/firebase.js`:

```js
const firebaseConfig = {
  apiKey:            "...",
  authDomain:        "...",
  projectId:         "...",
  storageBucket:     "...",
  messagingSenderId: "...",
  appId:             "...",
};
```

### 2. Firebase — Firestore Rules

לך ל-Firestore → **Rules** והדבק:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true;
    }
  }
}
```

> ⚠️ זה מאפשר גישה לכולם. האפליקציה מנהלת הרשאות בעצמה דרך login.
> אחרי שהכל עובד, שקול להגביל לפי auth.

### 3. GitHub — הגדרת Pages

1. לך ל-Settings → Pages
2. **Source:** GitHub Actions
3. לחץ Save

### 4. העלאה ל-GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/marshal-office-manager.git
git branch -M main
git push -u origin main
```

אחרי כמה דקות האתר יהיה זמין בכתובת:
`https://YOUR_USERNAME.github.io/marshal-office-manager/`

---

## פיתוח מקומי

```bash
npm install
npm run dev
```

---

## מיגרציה מ-PostgreSQL

```bash
# בתיקיית office-manager-firebase (מהשלב הקודם)
node migrate.js
```
