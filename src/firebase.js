import firebase from 'firebase/compat/app';
import 'firebase/compat/database';
import 'firebase/compat/auth';

const firebaseConfig = {
    authDomain: "project-management-control.firebaseapp.com",
    databaseURL: "https://project-management-control-default-rtdb.asia-southeast1.firebasedatabase.app/",
    projectId: "project-management-control",
    storageBucket: "project-management-control.firebasestorage.app",
    messagingSenderId: "1065130749964",
    appId: "1:1065130749964:web:404873304275127c85fe29",
    measurementId: "G-LLNQ6P9DQ9",
    apiKey: "AIzaSyDerRfccB2e-k3SSLoo5gxh10sRy9FRRwY" // WAJIB DIISI UNTUK LOGIN
};

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}

export const db = firebase.database();
export const auth = firebase.auth();
export default firebase;
