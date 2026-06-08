// src/firebase.js
// ─────────────────────────────────────────────────────────────────────────────
// הכנס כאן את פרטי הפרויקט שלך מ-Firebase Console:
// Project Settings → General → Your apps → Firebase SDK snippet → Config
// ─────────────────────────────────────────────────────────────────────────────
import { initializeApp } from "firebase/app";
import {
  getFirestore,
  collection,
  doc,
  getDocs,
  getDoc,
  setDoc,
  deleteDoc,
  writeBatch,
  orderBy,
  query,
  where,
  serverTimestamp,
} from "firebase/firestore";

const firebaseConfig = {
  apiKey:            "REPLACE_WITH_YOUR_API_KEY",
  authDomain:        "REPLACE_WITH_YOUR_AUTH_DOMAIN",
  projectId:         "REPLACE_WITH_YOUR_PROJECT_ID",
  storageBucket:     "REPLACE_WITH_YOUR_STORAGE_BUCKET",
  messagingSenderId: "REPLACE_WITH_YOUR_MESSAGING_SENDER_ID",
  appId:             "REPLACE_WITH_YOUR_APP_ID",
};

const app = initializeApp(firebaseConfig);
const db  = getFirestore(app);

export {
  db,
  collection, doc,
  getDocs, getDoc,
  setDoc, deleteDoc,
  writeBatch,
  orderBy, query, where,
  serverTimestamp,
};
