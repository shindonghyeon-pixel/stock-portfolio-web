import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

// 아래 firebaseConfig 부분을 아까 복사한 본인의 코드로 완전히 교체하세요!
const firebaseConfig = {
  apiKey: "AIzaSyB-xxxxxxxxxxxxxxx",
  authDomain: "stock-portfolio-xxxx.firebaseapp.com",
  projectId: "stock-portfolio-xxxx",
  storageBucket: "stock-portfolio-xxxx.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef123456"
};

// Firebase 초기화 및 데이터베이스 연결
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
