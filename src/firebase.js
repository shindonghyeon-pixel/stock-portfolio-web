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

// 데이터베이스와 인증 기능을 모두 내보내어 빌드 에러를 원천 차단합니다.
export const db = getFirestore(app);
export const auth = getAuth(app);
