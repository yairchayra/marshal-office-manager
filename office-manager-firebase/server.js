require("dotenv").config();
const express = require("express");
const cors = require("cors");
const crypto = require("crypto");
const admin = require("firebase-admin");
const { setup, getDb } = require("./db");

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(cors({ origin: "*", methods: ["GET", "POST", "PUT", "DELETE", "PATCH"] }));

app.use((req, res, next) => {
  const start = Date.now();
  res.on("finish", () => {
    const ms = Date.now() - start;
    const color = res.statusCode >= 500 ? "\x1b[31m" : res.statusCode >= 400 ? "\x1b[33m" : "\x1b[32m";
    console.log(`${color}[${res.statusCode}]\x1b[0m ${req.method} ${req.path} (${ms}ms)`);
  });
  next();
});

// ─── Helpers ────────────────────────────────────────────────────────────────────
const toNum = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
const toInt = (v) => { const n = parseInt(v);   return isNaN(n) ? null : n; };
const clean = (v) => (v === undefined || v === null || v === "" || String(v).toLowerCase() === "nan") ? null : String(v).trim();

function sha256Hash(salt, password) {
  return crypto.createHash("sha256").update(salt + password).digest("hex");
}
function makeHash(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  return salt + ":" + sha256Hash(salt, password);
}
function checkHash(password, stored) {
  if (!stored) return false;
  const [salt, hash] = stored.split(":");
  return sha256Hash(salt, password) === hash;
}

// Simple in-memory token store (fine for office LAN)
const sessions = new Map();
function makeToken() { return crypto.randomBytes(32).toString("hex"); }
function authMiddleware(req, res, next) {
  const token = req.headers["x-auth-token"];
  const session = sessions.get(token);
  if (!session) return res.status(401).json({ error: "Not logged in" });
  req.user = session;
  next();
}
function adminOnly(req, res, next) {
  if (!req.user?.isAdmin) return res.status(403).json({ error: "Admin only" });
  next();
}

// ─── PC helper: convert Firestore doc to API shape ──────────────────────────
function docToPC(id, data) {
  return {
    pcId: id,
    userName: data.userName,
    autocadVersion: data.autocadVersion,
    licenseType: data.licenseType,
    serialNumber: data.serialNumber,
    productKey: data.productKey,
    operatingSystem: data.operatingSystem,
    cpu: data.cpu || {},
    gpu: data.gpu || {},
    ssd: data.ssd || {},
    ram: data.ram || {},
    teken: data.teken || 0,
    comments: data.comments,
    user: data.user || { mail: "", passwordHash: null },
  };
}

function pcToDoc(pc) {
  return {
    userName: clean(pc.userName),
    autocadVersion: clean(pc.autocadVersion),
    licenseType: clean(pc.licenseType),
    serialNumber: clean(pc.serialNumber),
    productKey: clean(pc.productKey),
    operatingSystem: clean(pc.operatingSystem),
    cpu: {
      manufacturer: clean(pc.cpu?.manufacturer),
      model: clean(pc.cpu?.model),
      cores: toInt(pc.cpu?.cores),
      ghz: toNum(pc.cpu?.ghz),
    },
    gpu: {
      manufacturer: clean(pc.gpu?.manufacturer),
      model: clean(pc.gpu?.model),
      vramGb: toInt(pc.gpu?.vramGb),
    },
    ssd: {
      brand: clean(pc.ssd?.brand),
      capacityGb: toInt(pc.ssd?.capacityGb),
      type: clean(pc.ssd?.type),
    },
    ram: {
      sizeGb: toInt(pc.ram?.sizeGb),
      speedMhz: toInt(pc.ram?.speedMhz),
      sticks: toInt(pc.ram?.sticks),
    },
    teken: pc.teken ? 1 : 0,
    comments: clean(pc.comments),
    user: {
      mail: clean(pc.user?.mail) || "",
      passwordHash: pc.user?.passwordHash || null,
    },
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
  };
}

// ─── Health ─────────────────────────────────────────────────────────────────────
app.get("/api/health", (req, res) => res.json({ status: "ok", db: "firestore" }));

// ─── Auth ───────────────────────────────────────────────────────────────────────
app.post("/api/auth/login", async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: "Email and password required" });

    const db = getDb();
    const snap = await db.collection("app_accounts")
      .where("email", "==", email.trim().toLowerCase()).limit(1).get();

    if (snap.empty) return res.status(401).json({ error: "Invalid email or password" });

    const doc = snap.docs[0];
    const account = doc.data();

    // First-time ADMIN_SEED — set the password
    if (account.passwordHash === "ADMIN_SEED") {
      const hash = makeHash(password);
      await doc.ref.update({ passwordHash: hash });
      const token = makeToken();
      sessions.set(token, { email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
      return res.json({ token, email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
    }

    if (!checkHash(password, account.passwordHash)) {
      return res.status(401).json({ error: "Invalid email or password" });
    }

    const token = makeToken();
    sessions.set(token, { email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
    console.log(`[AUTH] Login: ${account.email} (admin: ${account.isAdmin})`);
    res.json({ token, email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
  } catch (err) {
    console.error("[AUTH] Login error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/auth/logout", (req, res) => {
  sessions.delete(req.headers["x-auth-token"]);
  res.json({ success: true });
});

app.get("/api/auth/me", authMiddleware, (req, res) => res.json(req.user));

// ─── App Accounts ───────────────────────────────────────────────────────────────
app.get("/api/accounts", authMiddleware, adminOnly, async (req, res) => {
  try {
    const snap = await getDb().collection("app_accounts").orderBy("displayName").get();
    const accounts = snap.docs.map(d => {
      const data = d.data();
      return { id: d.id, email: data.email, display_name: data.displayName, is_admin: data.isAdmin, created_at: data.createdAt };
    });
    res.json(accounts);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/accounts", authMiddleware, adminOnly, async (req, res) => {
  try {
    const { email, displayName, password, isAdmin } = req.body;
    if (!email || !password) return res.status(400).json({ error: "Email and password required" });

    const db = getDb();
    const existing = await db.collection("app_accounts")
      .where("email", "==", email.trim().toLowerCase()).limit(1).get();
    if (!existing.empty) return res.status(409).json({ error: "Email already exists" });

    const hash = makeHash(password);
    await db.collection("app_accounts").add({
      email: email.trim().toLowerCase(),
      displayName: displayName || "",
      passwordHash: hash,
      isAdmin: !!isAdmin,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put("/api/accounts/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    const { displayName, isAdmin, password } = req.body;
    const update = { displayName, isAdmin: !!isAdmin };
    if (password) update.passwordHash = makeHash(password);
    await getDb().collection("app_accounts").doc(req.params.id).update(update);
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete("/api/accounts/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    await getDb().collection("app_accounts").doc(req.params.id).delete();
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── GET all PCs ────────────────────────────────────────────────────────────────
app.get("/api/pcs", authMiddleware, async (req, res) => {
  try {
    const snap = await getDb().collection("pcs").orderBy("userName").get();
    const pcs = {};
    snap.docs.forEach(d => { pcs[d.id] = docToPC(d.id, d.data()); });
    res.json(pcs);
  } catch (err) { console.error("[GET /pcs]", err.message); res.status(500).json({ error: err.message }); }
});

// ─── BULK import ────────────────────────────────────────────────────────────────
app.post("/api/pcs/bulk", authMiddleware, adminOnly, async (req, res) => {
  try {
    const db = getDb();
    const pcs = req.body;
    if (!Array.isArray(pcs)) return res.status(400).json({ error: `Expected array, got ${typeof pcs}` });

    console.log(`\n[BULK] ── Importing ${pcs.length} PCs ──`);
    let count = 0;
    const errors = [];

    // Firestore batch writes (max 500 per batch)
    const BATCH_SIZE = 400;
    for (let i = 0; i < pcs.length; i += BATCH_SIZE) {
      const batch = db.batch();
      const chunk = pcs.slice(i, i + BATCH_SIZE);
      for (const pc of chunk) {
        try {
          if (!pc.pcId) { console.warn(`[BULK] Skip — no pcId:`, pc.userName); continue; }
          const ref = db.collection("pcs").doc(clean(pc.pcId));
          batch.set(ref, pcToDoc(pc), { merge: true });
          count++;
        } catch (rowErr) {
          errors.push({ pcId: pc.pcId, userName: pc.userName, error: rowErr.message });
        }
      }
      await batch.commit();
    }

    console.log(`[BULK] ── Done: ${count} OK, ${errors.length} failed ──\n`);
    res.json({ success: true, count, errors });
  } catch (err) { console.error("[BULK] Fatal:", err); res.status(500).json({ error: err.message }); }
});

// ─── PC swap ────────────────────────────────────────────────────────────────────
app.patch("/api/pcs/:id/swap", authMiddleware, adminOnly, async (req, res) => {
  try {
    const { newUserName } = req.body;
    await getDb().collection("pcs").doc(req.params.id).update({
      userName: newUserName,
      "user.userName": newUserName,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    console.log(`[SWAP] ${req.params.id} → ${newUserName}`);
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── Single PC CRUD ─────────────────────────────────────────────────────────────
app.get("/api/pcs/:id", authMiddleware, async (req, res) => {
  try {
    const doc = await getDb().collection("pcs").doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ error: "Not found" });
    res.json(docToPC(doc.id, doc.data()));
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/pcs", authMiddleware, adminOnly, async (req, res) => {
  try {
    const pc = req.body;
    if (!pc.pcId) return res.status(400).json({ error: "pcId required" });
    await getDb().collection("pcs").doc(clean(pc.pcId)).set(pcToDoc(pc), { merge: true });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put("/api/pcs/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    await getDb().collection("pcs").doc(req.params.id).set(pcToDoc({ ...req.body, pcId: req.params.id }), { merge: true });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete("/api/pcs/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    await getDb().collection("pcs").doc(req.params.id).delete();
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── Desks ──────────────────────────────────────────────────────────────────────
app.get("/api/desks", authMiddleware, async (req, res) => {
  try {
    const snap = await getDb().collection("desks").orderBy("label").get();
    res.json(snap.docs.map(d => {
      const r = d.data();
      return { id: d.id, label: r.label, tableGroup: r.tableGroup, pcId: r.pcId, x: r.x, y: r.y, color: r.color || "#6aa3fc" };
    }));
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/desks", authMiddleware, adminOnly, async (req, res) => {
  try {
    const { id, label, tableGroup, pcId, x, y, color } = req.body;
    await getDb().collection("desks").doc(id).set({ label, tableGroup, pcId: pcId || null, x, y, color: color || "#6aa3fc" }, { merge: true });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put("/api/desks/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    const { label, tableGroup, pcId, x, y, color } = req.body;
    await getDb().collection("desks").doc(req.params.id).set({ label, tableGroup, pcId: pcId || null, x, y, color: color || "#6aa3fc" }, { merge: true });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete("/api/desks/:id", authMiddleware, adminOnly, async (req, res) => {
  try {
    await getDb().collection("desks").doc(req.params.id).delete();
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── Table Groups ────────────────────────────────────────────────────────────────
app.get("/api/tablegroups", authMiddleware, async (req, res) => {
  try {
    const doc = await getDb().collection("app_settings").doc("table_groups").get();
    res.json(doc.exists ? doc.data().value : []);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/tablegroups", authMiddleware, adminOnly, async (req, res) => {
  try {
    const groups = req.body;
    if (!Array.isArray(groups)) return res.status(400).json({ error: "Expected array" });
    await getDb().collection("app_settings").doc("table_groups").set({ value: groups });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── Floor Plan ──────────────────────────────────────────────────────────────────
app.get("/api/floorplan", authMiddleware, async (req, res) => {
  try {
    const doc = await getDb().collection("app_settings").doc("floor_plan").get();
    res.json({ imageData: doc.exists ? doc.data().imageData : null });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/floorplan", authMiddleware, adminOnly, async (req, res) => {
  try {
    await getDb().collection("app_settings").doc("floor_plan").set({
      imageData: req.body.imageData,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── 404 ────────────────────────────────────────────────────────────────────────
app.use((req, res) => {
  console.warn(`[404] ${req.method} ${req.path}`);
  res.status(404).json({ error: `Route not found: ${req.method} ${req.path}` });
});

const PORT = process.env.PORT || 3001;
setup();
app.listen(PORT, "0.0.0.0", () => console.log(`\n✓ Office Manager API (Firebase) — port ${PORT}\n`));
