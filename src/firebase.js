// src/firebase.js
// הכנס כאן את פרטי הפרויקט שלך מ-Firebase Console:
// Project Settings → General → Your apps → Firebase SDK snippet → Config
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
  getDownloadURL,
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
  apiKey: "AIzaSyBpHFm7wKA7tR25lemACFBEVVgdXyavGn4",
  authDomain: "marshal-office-manager.firebaseapp.com",
  projectId: "marshal-office-manager",
  storageBucket: "marshal-office-manager.firebasestorage.app",
  messagingSenderId: "747589186849",
  appId: "1:747589186849:web:6737dfcd6eacf5c76dfa60"
};

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
  ref, uploadString, getDownloadURL,
  signInWithEmailAndPassword,
  signOut,
  createUserWithEmailAndPassword,
  updatePassword,
  deleteUser,
  onAuthStateChanged,
};
