// src/firebase.js
// הערכים נטענים מ-.env.local (לא עולה ל-GitHub)
// ראה .env.local.example להוראות הגדרה
import { initializeApp } from "firebase/app";
import {
  getFirestore,
  collection, doc,
  getDocs, getDoc,
  setDoc, deleteDoc,
  writeBatch,
  orderBy, query, where,
  serverTimestamp,
} from "firebase/firestore";
import {
  getStorage,
  ref,
  uploadString,
  uploadBytes,
  getDownloadURL,
  listAll,
  deleteObject,
} from "firebase/storage";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  createUserWithEmailAndPassword,
  updatePassword,
  deleteUser,
  onAuthStateChanged,
} from "firebase/auth";

const firebaseConfig = {
  apiKey:            import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain:        import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId:         import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket:     import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId:             import.meta.env.VITE_FIREBASE_APP_ID,
};

// בדיקה שהערכים הוגדרו
if (!firebaseConfig.apiKey || firebaseConfig.apiKey === "undefined") {
  console.error("⚠️ Firebase config חסר! וודא שקובץ .env.local קיים עם הערכים הנכונים.");
}

const app     = initializeApp(firebaseConfig);
const db      = getFirestore(app);
const storage = getStorage(app);
const auth    = getAuth(app);

export {
  db, storage, auth,
  collection, doc,
  getDocs, getDoc,
  setDoc, deleteDoc,
  writeBatch,
  orderBy, query, where,
  serverTimestamp,
  ref, uploadString, uploadBytes, getDownloadURL, listAll, deleteObject,
  signInWithEmailAndPassword,
  signOut,
  createUserWithEmailAndPassword,
  updatePassword,
  deleteUser,
  onAuthStateChanged,
};
