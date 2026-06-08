const admin = require("firebase-admin");
require("dotenv").config();

let db;

function setup() {
  if (!admin.apps.length) {
    const serviceAccount = require(process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "./serviceAccountKey.json");
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount),
    });
  }
  db = admin.firestore();
  console.log("✓ Firebase Firestore connected");
  return db;
}

function getDb() {
  if (!db) throw new Error("Firestore not initialized — call setup() first");
  return db;
}

module.exports = { setup, getDb };
