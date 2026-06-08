// src/db.js — כל גישה ל-Firestore במקום API calls
import {
  db,
  collection, doc,
  getDocs, getDoc,
  setDoc, deleteDoc,
  writeBatch,
  orderBy, query, where,
  serverTimestamp,
} from "./firebase.js";

// ─── Helpers ────────────────────────────────────────────────────────────────
const toNum = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
const toInt = (v) => { const n = parseInt(v);   return isNaN(n) ? null : n; };
const clean = (v) =>
  v === undefined || v === null || v === "" || String(v).toLowerCase() === "nan"
    ? null : String(v).trim();

export function pcToDoc(pc) {
  return {
    userName:        clean(pc.userName),
    autocadVersion:  clean(pc.autocadVersion),
    licenseType:     clean(pc.licenseType),
    serialNumber:    clean(pc.serialNumber),
    productKey:      clean(pc.productKey),
    operatingSystem: clean(pc.operatingSystem),
    cpu: {
      manufacturer: clean(pc.cpu?.manufacturer),
      model:        clean(pc.cpu?.model),
      cores:        toInt(pc.cpu?.cores),
      ghz:          toNum(pc.cpu?.ghz),
    },
    gpu: {
      manufacturer: clean(pc.gpu?.manufacturer),
      model:        clean(pc.gpu?.model),
      vramGb:       toInt(pc.gpu?.vramGb),
    },
    ssd: {
      brand:      clean(pc.ssd?.brand),
      capacityGb: toInt(pc.ssd?.capacityGb),
      type:       clean(pc.ssd?.type),
    },
    ram: {
      sizeGb:   toInt(pc.ram?.sizeGb),
      speedMhz: toInt(pc.ram?.speedMhz),
      sticks:   toInt(pc.ram?.sticks),
    },
    teken:    pc.teken ? 1 : 0,
    comments: clean(pc.comments),
    user: {
      mail:         clean(pc.user?.mail) || "",
      passwordHash: pc.user?.passwordHash || null,
    },
    updatedAt: serverTimestamp(),
  };
}

export function docToPC(id, data) {
  return {
    pcId:            id,
    userName:        data.userName,
    autocadVersion:  data.autocadVersion,
    licenseType:     data.licenseType,
    serialNumber:    data.serialNumber,
    productKey:      data.productKey,
    operatingSystem: data.operatingSystem,
    cpu:             data.cpu  || {},
    gpu:             data.gpu  || {},
    ssd:             data.ssd  || {},
    ram:             data.ram  || {},
    teken:           data.teken || 0,
    comments:        data.comments,
    user:            data.user || { mail: "", passwordHash: null },
  };
}

// ─── PCs ────────────────────────────────────────────────────────────────────
export async function getPCs() {
  const snap = await getDocs(query(collection(db, "pcs"), orderBy("userName")));
  const pcs = {};
  snap.docs.forEach(d => { pcs[d.id] = docToPC(d.id, d.data()); });
  return pcs;
}

export async function getPC(pcId) {
  const d = await getDoc(doc(db, "pcs", pcId));
  return d.exists() ? docToPC(d.id, d.data()) : null;
}

export async function savePC(pc) {
  await setDoc(doc(db, "pcs", clean(pc.pcId)), pcToDoc(pc), { merge: true });
}

export async function deletePC(pcId) {
  await deleteDoc(doc(db, "pcs", pcId));
}

export async function swapPC(pcId, newUserName) {
  await setDoc(doc(db, "pcs", pcId), {
    userName: newUserName,
    "user.userName": newUserName,
    updatedAt: serverTimestamp(),
  }, { merge: true });
}

export async function bulkSavePCs(pcArray) {
  const CHUNK = 400;
  let count = 0;
  const errors = [];
  for (let i = 0; i < pcArray.length; i += CHUNK) {
    const batch = writeBatch(db);
    for (const pc of pcArray.slice(i, i + CHUNK)) {
      try {
        if (!pc.pcId) continue;
        batch.set(doc(db, "pcs", clean(pc.pcId)), pcToDoc(pc), { merge: true });
        count++;
      } catch (e) { errors.push({ pcId: pc.pcId, error: e.message }); }
    }
    await batch.commit();
  }
  return { count, errors };
}

// ─── Desks ──────────────────────────────────────────────────────────────────
export async function getDesks() {
  const snap = await getDocs(query(collection(db, "desks"), orderBy("label")));
  return snap.docs.map(d => {
    const r = d.data();
    return { id: d.id, label: r.label, tableGroup: r.tableGroup, pcId: r.pcId, x: r.x, y: r.y, color: r.color || "#6aa3fc" };
  });
}

export async function saveDesk(desk) {
  await setDoc(doc(db, "desks", desk.id), {
    label: desk.label, tableGroup: desk.tableGroup,
    pcId: desk.pcId || null, x: desk.x, y: desk.y, color: desk.color || "#6aa3fc",
  }, { merge: true });
}

export async function deleteDesk(deskId) {
  await deleteDoc(doc(db, "desks", deskId));
}

// ─── Table Groups ────────────────────────────────────────────────────────────
export async function getTableGroups() {
  const d = await getDoc(doc(db, "app_settings", "table_groups"));
  return d.exists() ? d.data().value : [];
}

export async function saveTableGroups(groups) {
  await setDoc(doc(db, "app_settings", "table_groups"), { value: groups });
}

// ─── Floor Plan ──────────────────────────────────────────────────────────────
export async function getFloorPlan() {
  const d = await getDoc(doc(db, "app_settings", "floor_plan"));
  return d.exists() ? d.data().imageData : null;
}

export async function saveFloorPlan(imageData) {
  await setDoc(doc(db, "app_settings", "floor_plan"), {
    imageData, updatedAt: serverTimestamp(),
  });
}

// ─── App Accounts ────────────────────────────────────────────────────────────
export async function getAccounts() {
  const snap = await getDocs(query(collection(db, "app_accounts"), orderBy("displayName")));
  return snap.docs.map(d => {
    const data = d.data();
    return { id: d.id, email: data.email, display_name: data.displayName, is_admin: data.isAdmin };
  });
}

export async function getAccountByEmail(email) {
  const snap = await getDocs(
    query(collection(db, "app_accounts"), where("email", "==", email.trim().toLowerCase()))
  );
  if (snap.empty) return null;
  return { id: snap.docs[0].id, ...snap.docs[0].data() };
}

export async function createAccount({ email, displayName, passwordHash, isAdmin }) {
  // Check duplicate
  const existing = await getAccountByEmail(email);
  if (existing) throw new Error("Email already exists");
  const ref = doc(collection(db, "app_accounts"));
  await setDoc(ref, {
    email: email.trim().toLowerCase(),
    displayName: displayName || "",
    passwordHash,
    isAdmin: !!isAdmin,
    createdAt: serverTimestamp(),
  });
}

export async function updateAccount(id, { displayName, isAdmin, passwordHash }) {
  const update = { displayName, isAdmin: !!isAdmin };
  if (passwordHash) update.passwordHash = passwordHash;
  await setDoc(doc(db, "app_accounts", id), update, { merge: true });
}

export async function deleteAccount(id) {
  await deleteDoc(doc(db, "app_accounts", id));
}

export async function updateAccountPassword(id, passwordHash) {
  await setDoc(doc(db, "app_accounts", id), { passwordHash }, { merge: true });
}
