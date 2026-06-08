# Marshal Office Manager

אפליקציית React המתחברת ישירות ל-Firebase — ללא שרת, ללא backend.

**🌐 Live:** `https://yairchayra.github.io/marshal-office-manager/`

---

## הגדרה ראשונית (פעם אחת)

### 1. Firebase config — `src/firebase.js`
Firebase Console → הפרויקט → **Project Settings → General → Your apps → Config**

### 2. Firestore Rules
```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if request.auth != null;
    }
  }
}
```

### 3. Storage Rules
```
rules_version = '2';
service firebase.storage {
  match /b/{bucket}/o {
    match /{allPaths=**} {
      allow read, write: if request.auth != null;
    }
  }
}
```

### 4. הוספת משתמשים
Firebase Console → **Authentication → Users → Add user**
אחרי כן — לך ל-Firestore ובנה document ב-`user_profiles/{uid}`:
```json
{
  "email": "user@example.com",
  "displayName": "שם המשתמש",
  "isAdmin": true
}
```

### 5. GitHub Pages
Settings → Pages → Source: **GitHub Actions**

### 6. העלאה ל-GitHub
```bash
git add .
git commit -m "Switch to Firebase Auth"
git push origin main
```

---

## מיגרציה מ-PostgreSQL (פעם אחת)
```bash
cd office-manager-firebase
npm install
node migrate.js        # PCs, Desks, Settings
node migrate-floor.js  # תמונת מפת הקומה → Storage
```
