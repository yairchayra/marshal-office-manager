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
  apiKey: "AIzaSyBpHFm7wKA7tR25lemACFBEVVgdXyavGn4",
  authDomain: "marshal-office-manager.firebaseapp.com",
  projectId: "marshal-office-manager",
  storageBucket: "marshal-office-manager.firebasestorage.app",
  messagingSenderId: "747589186849",
  appId: "1:747589186849:web:6737dfcd6eacf5c76dfa60"
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
