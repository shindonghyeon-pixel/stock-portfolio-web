import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBC7tErv9Qe14JpiOMmKFwis9h4Z789H1Y",
  authDomain: "stock-portfolio-2673c.firebaseapp.com",
  projectId: "stock-portfolio-2673c",
  storageBucket: "stock-portfolio-2673c.firebasestorage.app",
  messagingSenderId: "476479685670",
  appId: "1:476479685670:web:35f48aefdb901e67dd92e7"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
