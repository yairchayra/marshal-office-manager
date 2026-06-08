// migrate-floor.js — מעלה את תמונת המפה ל-Firebase Storage
require("dotenv").config();
const { Pool } = require("pg");
const admin = require("firebase-admin");

const serviceAccount = require(process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "./serviceAccountKey.json");

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  storageBucket: `${serviceAccount.project_id}.appspot.com`,
});

const db     = admin.firestore();
const bucket = admin.storage().bucket();

const pool = new Pool({
  host: process.env.DB_HOST, port: process.env.DB_PORT,
  database: process.env.DB_NAME, user: process.env.DB_USER, password: process.env.DB_PASSWORD,
});

async function main() {
  console.log("🗺️  Migrating floor plan image to Firebase Storage...");

  const { rows } = await pool.query("SELECT image_data FROM floor_plan ORDER BY id DESC LIMIT 1");
  if (!rows.length || !rows[0].image_data) {
    console.log("ℹ️  No floor plan found in PostgreSQL — skipping.");
    process.exit(0);
  }

  const imageData = rows[0].image_data;
  const matches   = imageData.match(/^data:(.+);base64,(.+)$/s);
  if (!matches) { console.error("✗ Invalid image format"); process.exit(1); }

  const mimeType = matches[1];
  const buffer   = Buffer.from(matches[2], "base64");
  const ext      = mimeType.includes("png") ? "png" : "jpg";
  const fileName = `floor-plan/floor.${ext}`;

  console.log(`   Size: ${(buffer.length / 1024).toFixed(0)} KB`);
  console.log(`   Uploading as: ${fileName}`);

  const file = bucket.file(fileName);
  await file.save(buffer, { metadata: { contentType: mimeType } });
  await file.makePublic();

  const publicUrl = `https://storage.googleapis.com/${bucket.name}/${fileName}`;
  console.log(`   ✓ Public URL: ${publicUrl}`);

  // שמור את ה-URL ב-Firestore
  await db.collection("app_settings").doc("floor_plan").set({
    imageUrl:  publicUrl,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
  });

  console.log("✅ Done! Floor plan URL saved to Firestore.");
  await pool.end();
  process.exit(0);
}

main().catch(err => {
  console.error("✗ Failed:", err.message);
  process.exit(1);
});
