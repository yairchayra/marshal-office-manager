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
  deleteUser,
} from "./firebase.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────
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

// ─── Auth — Firebase Authentication ──────────────────────────────────────────

// התחברות
export async function loginUser(email, password) {
  const cred = await signInWithEmailAndPassword(auth, email, password);
  // שלוף פרופיל מ-Firestore (displayName + isAdmin)
  const profile = await getUserProfile(cred.user.uid);
  return {
    uid:         cred.user.uid,
    email:       cred.user.email,
    displayName: profile?.displayName || cred.user.email,
    isAdmin:     profile?.isAdmin || false,
  };
}

// התנתקות
export async function logoutUser() {
  await signOut(auth);
}

// שלוף פרופיל משתמש מ-Firestore
export async function getUserProfile(uid) {
  const d = await getDoc(doc(db, "user_profiles", uid));
  return d.exists() ? d.data() : null;
}

// ─── Admin: ניהול משתמשים ────────────────────────────────────────────────────
// הערה: יצירת/מחיקת משתמשים ב-Firebase Auth דורשת Admin SDK (server-side).
// לכן — יצירה נעשית ע"י המשתמש עצמו (self-registration) או דרך Firebase Console.
// אנחנו שומרים רק את הפרופיל (displayName + isAdmin) ב-Firestore.

export async function getAccounts() {
  const snap = await getDocs(query(collection(db, "user_profiles"), orderBy("displayName")));
  return snap.docs.map(d => ({
    id:           d.id,
    email:        d.data().email,
    display_name: d.data().displayName,
    is_admin:     d.data().isAdmin,
  }));
}

export async function updateUserProfile(uid, { displayName, isAdmin }) {
  await setDoc(doc(db, "user_profiles", uid), { displayName, isAdmin: !!isAdmin }, { merge: true });
}

// רישום משתמש חדש (נשתמש ב-createUserWithEmailAndPassword)
// הערה: Firebase מחבר אוטומטית את המשתמש החדש — לכן אנחנו מתנתקים ומחברים חזרה את האדמין
export async function registerNewUser(email, password, displayName, isAdmin) {
  // שמור את credentials של האדמין הנוכחי
  const adminUser = auth.currentUser;
  if (!adminUser) throw new Error("Not logged in");

  // צור את המשתמש החדש — Firebase מחבר אותו אוטומטית
  const cred = await createUserWithEmailAndPassword(auth, email.trim().toLowerCase(), password);
  const newUid = cred.user.uid;

  // שמור פרופיל ב-Firestore
  await setDoc(doc(db, "user_profiles", newUid), {
    email:       email.trim().toLowerCase(),
    displayName: displayName || email,
    isAdmin:     !!isAdmin,
    createdAt:   serverTimestamp(),
  });

  // התנתק מהמשתמש החדש וחזור להתחבר כאדמין
  // Firebase מתנתק אוטומטית כשיוצרים משתמש חדש — לכן שולחים sign-in link
  // הפתרון: האדמין מתחבר מחדש לאחר יצירת המשתמש
  await signOut(auth);
  
  return { uid: newUid, email: cred.user.email };
}

// עדכון סיסמה של המשתמש המחובר כרגע
export async function changeCurrentUserPassword(newPassword) {
  if (!auth.currentUser) throw new Error("Not logged in");
  await updatePassword(auth.currentUser, newPassword);
}

// ─── PCs ──────────────────────────────────────────────────────────────────────
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
  await setDoc(doc(db, "pcs", pcId), {
    userName: newUserName,
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

export async function deleteDesk(deskId) {
  await deleteDoc(doc(db, "desks", deskId));
}

// ─── Table Groups ─────────────────────────────────────────────────────────────
export async function getTableGroups() {
  const d = await getDoc(doc(db, "app_settings", "table_groups"));
  return d.exists() ? d.data().value : [];
}

export async function saveTableGroups(groups) {
  await setDoc(doc(db, "app_settings", "table_groups"), { value: groups });
}

// ─── Floor Plan — Storage + Firestore ─────────────────────────────────────────
export async function getFloorPlan() {
  const d = await getDoc(doc(db, "app_settings", "floor_plan"));
  if (!d.exists()) return null;
  const data = d.data();
  return data.imageUrl || data.imageData || null;
}

export async function saveFloorPlan(imageData) {
  try {
    const storageRef = ref(storage, "floor-plan/floor.jpg");
    await uploadString(storageRef, imageData, "data_url");
    const imageUrl = await getDownloadURL(storageRef);
    await setDoc(doc(db, "app_settings", "floor_plan"), {
      imageUrl,
      updatedAt: serverTimestamp(),
    });
    return imageUrl;
  } catch (e) {
    console.warn("Storage upload failed, falling back:", e.message);
    await setDoc(doc(db, "app_settings", "floor_plan"), {
      imageData,
      updatedAt: serverTimestamp(),
    });
    return imageData;
  }
}
