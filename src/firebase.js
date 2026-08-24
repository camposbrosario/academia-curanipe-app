import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD05-RWc_ce4DHVUHokwPE79pr73URPq5M",
  authDomain: "academia-curanipe.firebaseapp.com",
  projectId: "academia-curanipe",
  storageBucket: "academia-curanipe.firebasestorage.app",
  messagingSenderId: "39739598232",
  appId: "1:39739598232:web:4a0ba4a68f3c42923b77ed",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
