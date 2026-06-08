// cleanup-firestore.js
// ─────────────────────────────────────────────────────────────────────────────
// סקריפט מיגרציה: מנקה את הנתונים הישנים ב-Firestore ומעדכן למבנה החדש
//
// מה הסקריפט עושה:
//   1. מחשבים (pcs) — שומר רק: userName, operatingSystem, cpuModel, gpuModel,
//                     ramGb, ssdGb, teken, comments, userMail, softwareIds
//                     מוחק: autocadVersion, licenseType, serialNumber, productKey,
//                            cpu.*, gpu.*, ssd.*, ram.*, user.passwordHash
//   2. מוחק את הקולקציה app_accounts (הוחלפה ב-Firebase Auth)
//   3. מדפיס סיכום מה שנמחק / עודכן
//
// שימוש:
//   cd office-manager-firebase
//   node cleanup-firestore.js
// ─────────────────────────────────────────────────────────────────────────────

require("dotenv").config();
const admin = require("firebase-admin");

const serviceAccount = require(
  process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "./serviceAccountKey.json"
);
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

// שדות ישנים שצריך למחוק מכל PC
const FIELDS_TO_DELETE = [
  "autocadVersion",
  "licenseType",
  "serialNumber",
  "productKey",
  "cpu",           // כל האובייקט הישן — יוחלף ב-cpuModel
  "gpu",           // כל האובייקט הישן — יוחלף ב-gpuModel
  "ssd",           // כל האובייקט הישן — יוחלף ב-ssdGb
  "ram",           // כל האובייקט הישן — יוחלף ב-ramGb
  "user",          // כל האובייקט הישן (כולל passwordHash) — mail יועבר ל-userMail
];

function extractNewFields(data) {
  // שלוף את המידע הנחוץ מהמבנה הישן
  const newFields = {};

  // cpuModel: נקח cpu.model אם קיים
  if (data.cpu?.model && !data.cpuModel) {
    newFields.cpuModel = data.cpu.model;
  }

  // gpuModel: נקח gpu.model אם קיים
  if (data.gpu?.model && !data.gpuModel) {
    newFields.gpuModel = data.gpu.model;
  }

  // ramGb: נקח ram.sizeGb אם קיים
  if (data.ram?.sizeGb && !data.ramGb) {
    newFields.ramGb = parseInt(data.ram.sizeGb) || null;
  }

  // ssdGb: נקח ssd.capacityGb אם קיים
  if (data.ssd?.capacityGb && !data.ssdGb) {
    newFields.ssdGb = parseInt(data.ssd.capacityGb) || null;
  }

  // userMail: נקח user.mail אם קיים
  if (data.user?.mail && !data.userMail) {
    newFields.userMail = data.user.mail;
  }

  // softwareIds: וודא שקיים
  if (!data.softwareIds) {
    newFields.softwareIds = [];
  }

  return newFields;
}

async function migratePCs() {
  console.log("\n📦 מעדכן מחשבים (PCs)...");
  const snap = await db.collection("pcs").get();

  if (snap.empty) {
    console.log("   ℹ️  אין מחשבים ב-Firestore");
    return;
  }

  let updated = 0, skipped = 0;
  const CHUNK = 400;
  const docs = snap.docs;

  for (let i = 0; i < docs.length; i += CHUNK) {
    const batch = db.batch();
    const chunk = docs.slice(i, i + CHUNK);

    for (const docSnap of chunk) {
      const data = docSnap.data();
      const ref  = docSnap.ref;

      // בנה את העדכון
      const update = {};

      // העבר שדות מהמבנה הישן
      const extracted = extractNewFields(data);
      Object.assign(update, extracted);

      // מחק שדות ישנים
      const hasOldFields = FIELDS_TO_DELETE.some(f => data[f] !== undefined);
      if (!hasOldFields && Object.keys(extracted).length === 0) {
        skipped++;
        continue; // כבר נקי
      }

      for (const field of FIELDS_TO_DELETE) {
        if (data[field] !== undefined) {
          update[field] = admin.firestore.FieldValue.delete();
        }
      }

      update.updatedAt = admin.firestore.FieldValue.serverTimestamp();

      batch.update(ref, update);
      updated++;

      // לוג מה שינינו
      const changes = [];
      if (extracted.cpuModel)  changes.push(`CPU: "${extracted.cpuModel}"`);
      if (extracted.gpuModel)  changes.push(`GPU: "${extracted.gpuModel}"`);
      if (extracted.ramGb)     changes.push(`RAM: ${extracted.ramGb}GB`);
      if (extracted.ssdGb)     changes.push(`SSD: ${extracted.ssdGb}GB`);
      if (extracted.userMail)  changes.push(`mail: ${extracted.userMail}`);
      const deleted = FIELDS_TO_DELETE.filter(f => data[f] !== undefined);

      console.log(`   ✓ ${docSnap.id} (${data.userName || "ללא שם"})`);
      if (changes.length)  console.log(`     שמר: ${changes.join(", ")}`);
      if (deleted.length)  console.log(`     מחק: ${deleted.join(", ")}`);
    }

    await batch.commit();
  }

  console.log(`\n   ✅ עודכנו: ${updated} | דולגו (כבר נקיים): ${skipped}`);
}

async function deleteCollection(name) {
  console.log(`\n🗑️  מוחק קולקציה: "${name}"...`);
  const snap = await db.collection(name).get();

  if (snap.empty) {
    console.log(`   ℹ️  "${name}" — ריקה או לא קיימת`);
    return 0;
  }

  const CHUNK = 400;
  let deleted = 0;

  for (let i = 0; i < snap.docs.length; i += CHUNK) {
    const batch = db.batch();
    snap.docs.slice(i, i + CHUNK).forEach(d => batch.delete(d.ref));
    await batch.commit();
    deleted += snap.docs.slice(i, i + CHUNK).length;
  }

  console.log(`   ✅ נמחקו ${deleted} documents`);
  return deleted;
}

async function printSummary() {
  console.log("\n📋 מצב נוכחי של ה-Firestore:");
  const collections = ["pcs", "softwares", "desks", "app_settings", "user_profiles", "app_accounts"];
  for (const col of collections) {
    try {
      const snap = await db.collection(col).get();
      console.log(`   ${col}: ${snap.size} documents`);
    } catch {
      console.log(`   ${col}: לא נגיש`);
    }
  }
}

async function main() {
  console.log("🚀 מתחיל מיגרציה / ניקוי Firestore...");
  console.log(`   Firebase: ${serviceAccount.project_id}`);

  // הדפס מצב לפני
  await printSummary();

  console.log("\n" + "─".repeat(50));

  // 1. עדכן מחשבים
  await migratePCs();

  // 2. מחק app_accounts (הוחלף ב-Firebase Auth + user_profiles)
  await deleteCollection("app_accounts");

  console.log("\n" + "─".repeat(50));
  console.log("\n🎉 סיום! מצב לאחר מיגרציה:");
  await printSummary();

  console.log("\n✅ הכל עודכן בהצלחה!\n");
  process.exit(0);
}

main().catch(err => {
  console.error("\n✗ שגיאה:", err.message);
  process.exit(1);
});
