import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyB-xxxxxxxxxxxxxxx",
  authDomain: "stock-portfolio-xxxx.firebaseapp.com",
  projectId: "stock-portfolio-xxxx",
  storageBucket: "stock-portfolio-xxxx.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef123456"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
