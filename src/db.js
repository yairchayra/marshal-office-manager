// src/db.js — כל גישה ל-Firestore, Storage ו-Auth
import {
  db, storage, auth,
  collection, doc,
  getDocs, getDoc,
  setDoc, deleteDoc,
  writeBatch,
  orderBy, query, where,
  serverTimestamp,
  ref, uploadString, getDownloadURL,
  signInWithEmailAndPassword,
  signOut,
  createUserWithEmailAndPassword,
  updatePassword,
} from "./firebase.js";

const clean  = (v) => (v === undefined || v === null || v === "" || String(v).toLowerCase() === "nan") ? null : String(v).trim();
const toInt  = (v) => { const n = parseInt(v);   return isNaN(n) ? null : n; };
const toNum  = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };

// ─── PC ───────────────────────────────────────────────────────────────────────
// מבנה מעודכן: ללא autocad/license/serial/productKey, CPU/GPU רק מודל, RAM/SSD רק GB, user ללא passwordHash
export function pcToDoc(pc) {
  return {
    userName:        clean(pc.userName),
    operatingSystem: clean(pc.operatingSystem),
    cpuModel:        clean(pc.cpuModel),
    gpuModel:        clean(pc.gpuModel),
    ramGb:           toInt(pc.ramGb),
    ssdGb:           toInt(pc.ssdGb),
    comments:        clean(pc.comments),
    userMail:        clean(pc.userMail) || "",
    softwareIds:     Array.isArray(pc.softwareIds) ? pc.softwareIds : [],
    updatedAt:       serverTimestamp(),
  };
}

export function docToPC(id, data) {
  return {
    pcId:            id,
    userName:        data.userName        || "",
    operatingSystem: data.operatingSystem || "",
    cpuModel:        data.cpuModel        || "",
    gpuModel:        data.gpuModel        || "",
    ramGb:           data.ramGb           || "",
    ssdGb:           data.ssdGb           || "",
    comments:        data.comments        || "",
    userMail:        data.userMail        || "",
    softwareIds:     data.softwareIds     || [],
  };
}

export async function getPCs() {
  const snap = await getDocs(query(collection(db, "pcs"), orderBy("userName")));
  const pcs = {};
  snap.docs.forEach(d => { pcs[d.id] = docToPC(d.id, d.data()); });
  return pcs;
}

export async function savePC(pc) {
  await setDoc(doc(db, "pcs", clean(pc.pcId)), pcToDoc(pc), { merge: true });
}

export async function deletePC(pcId) {
  await deleteDoc(doc(db, "pcs", pcId));
}

export async function swapPC(pcId, newUserName) {
  await setDoc(doc(db, "pcs", pcId), { userName: newUserName, updatedAt: serverTimestamp() }, { merge: true });
}

export async function bulkSavePCs(pcArray) {
  const CHUNK = 400;
  let count = 0, errors = [];
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

// ─── Software ─────────────────────────────────────────────────────────────────
// מבנה: { id, companyName, softwareName, licenseType: "annual"|"permanent",
//          expiryDate?: "YYYY-MM-DD", price: number, invoice?: string, productKey?: string }

export function softwareToDoc(sw) {
  const type = ["annual","monthly","permanent"].includes(sw.licenseType) ? sw.licenseType : "permanent";
  return {
    companyName:  clean(sw.companyName),
    softwareName: clean(sw.softwareName),
    licenseType:  type,
    expiryDate:   type !== "permanent" ? (clean(sw.expiryDate) || null) : null,
    months:       type === "monthly" ? (parseInt(sw.months) || 12) : null,
    price:        toNum(sw.price) || 0,
    invoice:      clean(sw.invoice)    || null,
    productKey:   clean(sw.productKey) || null,
    updatedAt:    serverTimestamp(),
  };
}

export function docToSoftware(id, data) {
  return {
    id,
    companyName:  data.companyName  || "",
    softwareName: data.softwareName || "",
    licenseType:  data.licenseType  || "permanent",
    expiryDate:   data.expiryDate   || null,
    months:       data.months       || 12,
    price:        data.price        || 0,
    invoice:      data.invoice      || null,
    productKey:   data.productKey   || null,
  };
}

export async function getSoftwares() {
  const snap = await getDocs(query(collection(db, "softwares"), orderBy("softwareName")));
  return snap.docs.map(d => docToSoftware(d.id, d.data()));
}

export async function saveSoftware(sw) {
  const ref_ = sw.id ? doc(db, "softwares", sw.id) : doc(collection(db, "softwares"));
  await setDoc(ref_, softwareToDoc(sw), { merge: true });
  return ref_.id;
}

export async function deleteSoftware(id) {
  // הסר את התוכנה מכל ה-PCs
  const snap = await getDocs(query(collection(db, "pcs"), where("softwareIds", "array-contains", id)));
  const batch = writeBatch(db);
  snap.docs.forEach(d => {
    const ids = (d.data().softwareIds || []).filter(x => x !== id);
    batch.update(d.ref, { softwareIds: ids });
  });
  batch.delete(doc(db, "softwares", id));
  await batch.commit();
}

// ─── Desks ────────────────────────────────────────────────────────────────────
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

export async function deleteDesk(id) { await deleteDoc(doc(db, "desks", id)); }

// ─── Table Groups ─────────────────────────────────────────────────────────────
export async function getTableGroups() {
  const d = await getDoc(doc(db, "app_settings", "table_groups"));
  return d.exists() ? d.data().value : [];
}
export async function saveTableGroups(groups) {
  await setDoc(doc(db, "app_settings", "table_groups"), { value: groups });
}

// ─── Floor Plan ───────────────────────────────────────────────────────────────
export async function getFloorPlan() {
  const d = await getDoc(doc(db, "app_settings", "floor_plan"));
  if (!d.exists()) return null;
  return d.data().imageUrl || d.data().imageData || null;
}

export async function saveFloorPlan(imageData) {
  try {
    const storageRef = ref(storage, "floor-plan/floor.jpg");
    await uploadString(storageRef, imageData, "data_url");
    const imageUrl = await getDownloadURL(storageRef);
    await setDoc(doc(db, "app_settings", "floor_plan"), { imageUrl, updatedAt: serverTimestamp() });
    return imageUrl;
  } catch (e) {
    await setDoc(doc(db, "app_settings", "floor_plan"), { imageData, updatedAt: serverTimestamp() });
    return imageData;
  }
}

// ─── Auth ─────────────────────────────────────────────────────────────────────
export async function loginUser(email, password) {
  const cred = await signInWithEmailAndPassword(auth, email, password);
  const profile = await getUserProfile(cred.user.uid);
  return { uid: cred.user.uid, email: cred.user.email, displayName: profile?.displayName || cred.user.email, isAdmin: profile?.isAdmin || false };
}

export async function logoutUser() { await signOut(auth); }

export async function getUserProfile(uid) {
  const d = await getDoc(doc(db, "user_profiles", uid));
  return d.exists() ? d.data() : null;
}

export async function getAccounts() {
  const snap = await getDocs(query(collection(db, "user_profiles"), orderBy("displayName")));
  return snap.docs.map(d => ({ id: d.id, email: d.data().email, display_name: d.data().displayName, is_admin: d.data().isAdmin }));
}

export async function updateUserProfile(uid, { displayName, isAdmin }) {
  await setDoc(doc(db, "user_profiles", uid), { displayName, isAdmin: !!isAdmin }, { merge: true });
}

export async function registerNewUser(email, password, displayName, isAdmin) {
  if (!auth.currentUser) throw new Error("Not logged in");
  const cred = await createUserWithEmailAndPassword(auth, email.trim().toLowerCase(), password);
  await setDoc(doc(db, "user_profiles", cred.user.uid), {
    email: email.trim().toLowerCase(), displayName: displayName || email, isAdmin: !!isAdmin, createdAt: serverTimestamp(),
  });
  await signOut(auth);
  return { uid: cred.user.uid, email: cred.user.email };
}

export async function changeCurrentUserPassword(newPassword) {
  if (!auth.currentUser) throw new Error("Not logged in");
  await updatePassword(auth.currentUser, newPassword);
}

// ─── Reports ──────────────────────────────────────────────────────────────────
function calcAnnualCost(sw) {
  if (sw.licenseType === "annual")   return sw.price || 0;
  if (sw.licenseType === "monthly")  return (sw.price || 0) * (sw.months || 12);
  return 0;
}

export async function getCostReport(pcs, softwares) {
  const swMap = {};
  softwares.forEach(s => { swMap[s.id] = s; });

  return Object.values(pcs).map(pc => {
    const pcSoftwares   = (pc.softwareIds || []).map(id => swMap[id]).filter(Boolean);
    const annualCost    = pcSoftwares.filter(s => s.licenseType !== "permanent").reduce((sum, s) => sum + calcAnnualCost(s), 0);
    const permanentCost = pcSoftwares.filter(s => s.licenseType === "permanent").reduce((sum, s) => sum + (s.price || 0), 0);
    return {
      pcId: pc.pcId, userName: pc.userName, userMail: pc.userMail,
      softwares: pcSoftwares, annualCost, permanentCost,
      totalCost: annualCost + permanentCost,
    };
  }).sort((a, b) => b.totalCost - a.totalCost);
}
