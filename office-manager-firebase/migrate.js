/**
 * migrate.js — PostgreSQL → Firebase Firestore migration
 *
 * Usage:
 *   node migrate.js
 *
 * Requires:
 *   - .env with DB_* (Postgres) and FIREBASE_SERVICE_ACCOUNT_PATH
 *   - serviceAccountKey.json (download from Firebase Console)
 *
 * Collections created in Firestore:
 *   pcs, desks, app_settings, app_accounts
 */

require("dotenv").config();
const { Pool } = require("pg");
const admin = require("firebase-admin");

// ─── Firebase init ──────────────────────────────────────────────────────────────
const serviceAccount = require(process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "./serviceAccountKey.json");
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

// ─── PostgreSQL init ────────────────────────────────────────────────────────────
const pool = new Pool({
  host: process.env.DB_HOST,
  port: process.env.DB_PORT,
  database: process.env.DB_NAME,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
});

// ─── Helpers ────────────────────────────────────────────────────────────────────
async function batchWrite(collection, docs) {
  const CHUNK = 400; // Firestore batch limit is 500
  let total = 0;
  for (let i = 0; i < docs.length; i += CHUNK) {
    const batch = db.batch();
    for (const { id, data } of docs.slice(i, i + CHUNK)) {
      batch.set(db.collection(collection).doc(String(id)), data);
    }
    await batch.commit();
    total += docs.slice(i, i + CHUNK).length;
    console.log(`  ✓ Written ${total}/${docs.length} to "${collection}"`);
  }
}

// ─── Migrate PCs ────────────────────────────────────────────────────────────────
async function migratePCs() {
  console.log("\n📦 Migrating PCs...");
  const { rows } = await pool.query(`
    SELECT p.*, u.mail, u.password_hash
    FROM pcs p
    LEFT JOIN users u ON p.pc_id = u.pc_id
    ORDER BY p.user_name
  `);

  const docs = rows.map(r => ({
    id: r.pc_id,
    data: {
      userName: r.user_name || null,
      autocadVersion: r.autocad_version || null,
      licenseType: r.license_type || null,
      serialNumber: r.serial_number || null,
      productKey: r.product_key || null,
      operatingSystem: r.operating_system || null,
      cpu: {
        manufacturer: r.cpu_manufacturer || null,
        model: r.cpu_model || null,
        cores: r.cpu_cores ? parseInt(r.cpu_cores) : null,
        ghz: r.cpu_ghz ? parseFloat(r.cpu_ghz) : null,
      },
      gpu: {
        manufacturer: r.gpu_manufacturer || null,
        model: r.gpu_model || null,
        vramGb: r.gpu_vram_gb ? parseInt(r.gpu_vram_gb) : null,
      },
      ssd: {
        brand: r.ssd_brand || null,
        capacityGb: r.ssd_capacity_gb ? parseInt(r.ssd_capacity_gb) : null,
        type: r.ssd_type || null,
      },
      ram: {
        sizeGb: r.ram_size_gb ? parseInt(r.ram_size_gb) : null,
        speedMhz: r.ram_speed_mhz ? parseInt(r.ram_speed_mhz) : null,
        sticks: r.ram_sticks ? parseInt(r.ram_sticks) : null,
      },
      teken: r.teken ? parseInt(r.teken) : 0,
      comments: r.comments || null,
      user: {
        mail: r.mail || "",
        passwordHash: r.password_hash || null,
      },
      createdAt: r.created_at ? admin.firestore.Timestamp.fromDate(new Date(r.created_at)) : null,
      updatedAt: r.updated_at ? admin.firestore.Timestamp.fromDate(new Date(r.updated_at)) : null,
    },
  }));

  await batchWrite("pcs", docs);
  console.log(`✅ PCs migrated: ${docs.length}`);
}

// ─── Migrate Desks ──────────────────────────────────────────────────────────────
async function migrateDesks() {
  console.log("\n🪑 Migrating Desks...");
  const { rows } = await pool.query("SELECT * FROM desks ORDER BY label");

  const docs = rows.map(r => ({
    id: r.id,
    data: {
      label: r.label || null,
      tableGroup: r.table_group || null,
      pcId: r.pc_id || null,
      x: r.x_pct ? parseFloat(r.x_pct) : null,
      y: r.y_pct ? parseFloat(r.y_pct) : null,
      color: r.color || "#6aa3fc",
    },
  }));

  await batchWrite("desks", docs);
  console.log(`✅ Desks migrated: ${docs.length}`);
}

// ─── Migrate App Accounts ───────────────────────────────────────────────────────
async function migrateAccounts() {
  console.log("\n👤 Migrating App Accounts...");
  const { rows } = await pool.query("SELECT * FROM app_accounts ORDER BY id");

  const docs = rows.map(r => ({
    id: String(r.id),
    data: {
      email: r.email,
      displayName: r.display_name || "",
      passwordHash: r.password_hash || null,
      isAdmin: !!r.is_admin,
      createdAt: r.created_at ? admin.firestore.Timestamp.fromDate(new Date(r.created_at)) : null,
    },
  }));

  await batchWrite("app_accounts", docs);
  console.log(`✅ App Accounts migrated: ${docs.length}`);
}

// ─── Migrate App Settings ───────────────────────────────────────────────────────
async function migrateSettings() {
  console.log("\n⚙️  Migrating App Settings...");
  const { rows } = await pool.query("SELECT * FROM app_settings");

  for (const row of rows) {
    let value;
    try { value = JSON.parse(row.value); } catch { value = row.value; }
    await db.collection("app_settings").doc(row.key).set({ value });
    console.log(`  ✓ Setting: "${row.key}"`);
  }
  console.log(`✅ Settings migrated: ${rows.length}`);
}

// ─── Migrate Floor Plan ─────────────────────────────────────────────────────────
async function migrateFloorPlan() {
  console.log("\n🗺️  Migrating Floor Plan...");
  const { rows } = await pool.query("SELECT image_data FROM floor_plan ORDER BY id DESC LIMIT 1");

  if (rows.length && rows[0].image_data) {
    await db.collection("app_settings").doc("floor_plan").set({
      imageData: rows[0].image_data,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    console.log("✅ Floor plan migrated (image stored in Firestore)");
    console.log("⚠️  NOTE: If the image is large, consider moving it to Firebase Storage instead.");
  } else {
    console.log("ℹ️  No floor plan found — skipping.");
  }
}

// ─── Main ───────────────────────────────────────────────────────────────────────
async function main() {
  console.log("🚀 Starting PostgreSQL → Firebase migration...");
  console.log(`   Postgres: ${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_NAME}`);
  console.log(`   Firebase: ${serviceAccount.project_id}`);

  try {
    await pool.query("SELECT 1"); // test connection
    console.log("✓ PostgreSQL connected\n");
  } catch (err) {
    console.error("✗ Cannot connect to PostgreSQL:", err.message);
    process.exit(1);
  }

  await migratePCs();
  await migrateDesks();
  await migrateAccounts();
  await migrateSettings();
  await migrateFloorPlan();

  console.log("\n🎉 Migration complete!");
  await pool.end();
  process.exit(0);
}

main().catch(err => {
  console.error("✗ Migration failed:", err);
  process.exit(1);
});
