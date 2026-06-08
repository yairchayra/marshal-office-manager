// setup-admin.js
// מגדיר משתמש כאדמין ב-Firestore + מוחק את הטבלאות הישנות
require("dotenv").config();
const admin = require("firebase-admin");
const readline = require("readline");

const serviceAccount = require(process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "./serviceAccountKey.json");
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });

const db   = admin.firestore();
const auth = admin.auth();

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = (q) => new Promise(res => rl.question(q, res));

async function deleteCollection(colName) {
  const snap = await db.collection(colName).get();
  if (snap.empty) { console.log(`  ℹ️  "${colName}" — ריקה או לא קיימת`); return; }
  const batch = db.batch();
  snap.docs.forEach(d => batch.delete(d.ref));
  await batch.commit();
  console.log(`  🗑️  "${colName}" — נמחקה (${snap.size} documents)`);
}

async function main() {
  console.log("\n🔧 Office Manager — Admin Setup\n");

  // ─── שלב 1: מחיקת טבלאות ישנות ──────────────────────────────────────────
  console.log("1️⃣  מוחק טבלאות התחברות ישנות...");
  await deleteCollection("app_accounts");
  console.log("   ✅ סיום מחיקה\n");

  // ─── שלב 2: הגדרת אדמין ──────────────────────────────────────────────────
  console.log("2️⃣  הגדרת משתמש אדמין");
  const email = await ask("   הכנס את האימייל שלך (כפי שמופיע ב-Firebase Auth): ");

  let user;
  try {
    user = await auth.getUserByEmail(email.trim());
    console.log(`   ✓ נמצא: ${user.displayName || user.email} (UID: ${user.uid})`);
  } catch (e) {
    console.error(`   ✗ משתמש עם אימייל "${email}" לא נמצא ב-Firebase Auth`);
    console.error("   וודא שהמשתמש קיים ב-Firebase Console → Authentication → Users");
    rl.close(); process.exit(1);
  }

  const displayName = await ask(`   שם תצוגה [${user.displayName || user.email}]: `) || user.displayName || user.email;

  await db.collection("user_profiles").doc(user.uid).set({
    email:       user.email,
    displayName: displayName.trim(),
    isAdmin:     true,
    createdAt:   admin.firestore.FieldValue.serverTimestamp(),
  }, { merge: true });

  console.log(`\n   ✅ ${displayName} הוגדר כאדמין!\n`);
  console.log("🎉 סיום! כעת תוכל להתחבר לאתר עם הרשאות אדמין.");
  console.log("   אם כבר מחובר — התנתק והתחבר מחדש.\n");

  rl.close();
  process.exit(0);
}

main().catch(err => { console.error("✗ שגיאה:", err.message); rl.close(); process.exit(1); });
