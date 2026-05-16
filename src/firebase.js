import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyCJS4ugnshuEcjmlm7vjz4y--UfjdzNp50",
  authDomain: "f-gestion.firebaseapp.com",
  projectId: "f-gestion",
  storageBucket: "f-gestion.firebasestorage.app",
  messagingSenderId: "1066677487813",
  appId: "1:1066677487813:web:7d5e3113905798cf1708ac",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
