// ════════════════════════════════════════════════════════════════
//  firebase.js — เชื่อมต่อ Firebase Firestore
//  โรงเรียนด่านช้างวิทยา
// ════════════════════════════════════════════════════════════════
//
//  ✏️  วิธีใส่ค่า:
//  1. ไปที่ https://console.firebase.google.com
//  2. Project Settings (⚙️) → Your apps → Web
//  3. ก็อป firebaseConfig มาแทนที่ด้านล่าง
//
// ════════════════════════════════════════════════════════════════

import { initializeApp } from 'firebase/app'
import {
  getFirestore, collection, doc, onSnapshot,
  setDoc, updateDoc, deleteDoc, writeBatch,
  query, orderBy, getDocs,
} from 'firebase/firestore'

// ✏️ แก้ค่าตรงนี้ให้ตรงกับ Firebase project ของคุณ
const firebaseConfig = {
  apiKey: "AIzaSyATGCcaRCrtve9MZqr-OGvaBxfyzsDQiX0",
  authDomain: "danchang-behavior.firebaseapp.com",
  projectId: "danchang-behavior",
  storageBucket: "danchang-behavior.firebasestorage.app",
  messagingSenderId: "823201706381",
  appId: "1:823201706381:web:e13bc0c6f42067038e8dc4",
  measurementId: "G-N1X0NLTDBW"
};

const app = initializeApp(firebaseConfig)
export const db = getFirestore(app)

// ── Helpers ──────────────────────────────────────────────────────
export const semKey      = (y, t)     => `${y}_${t}`
export const studentsCol = (y, t)     => collection(db, 'semesters', semKey(y, t), 'students')
export const studentDoc  = (y, t, id) => doc(db, 'semesters', semKey(y, t), 'students', String(id))
export const rulesCol    = ()         => collection(db, 'rules')
export const ruleDoc     = (id)       => doc(db, 'rules', String(id))
export const configDoc   = ()         => doc(db, 'config', 'semester')

// ── Listen: semester config ───────────────────────────────────────
export function listenConfig(cb) {
  return onSnapshot(configDoc(), snap => {
    cb(snap.exists() ? snap.data() : { year: new Date().getFullYear() + 543, term: 1 })
  })
}
export async function setActiveSemester(year, term) {
  await setDoc(configDoc(), { year: Number(year), term: Number(term) })
}

// ── Listen: students ──────────────────────────────────────────────
export function listenStudents(year, term, cb) {
  return onSnapshot(
    query(studentsCol(year, term), orderBy('class'), orderBy('studentId')),
    snap => cb(snap.docs.map(d => ({ ...d.data(), _docId: d.id })))
  )
}

// ── Listen: rules ─────────────────────────────────────────────────
export function listenRules(cb) {
  return onSnapshot(rulesCol(), snap => cb(snap.docs.map(d => ({ ...d.data() }))))
}

// ── Student writes ────────────────────────────────────────────────
export async function updateStudentScore(year, term, id, score, transactions) {
  await updateDoc(studentDoc(year, term, id), { score, transactions })
}

export async function bulkImportStudents(year, term, students, mode, existing) {
  const batch = writeBatch(db)
  if (mode === 'replace') {
    existing.forEach(s  => batch.delete(studentDoc(year, term, s.id)))
    students.forEach(s  => batch.set(studentDoc(year, term, s.id), s))
  } else {
    const ids = new Set(existing.map(s => String(s.studentId)))
    students.filter(s => !ids.has(String(s.studentId)))
            .forEach(s => batch.set(studentDoc(year, term, s.id), s))
  }
  await batch.commit()
}

// ── Rule writes ───────────────────────────────────────────────────
export async function saveRule(rule) { await setDoc(ruleDoc(rule.id), rule) }
export async function deleteRule(id) { await deleteDoc(ruleDoc(id)) }

export async function seedRulesIfEmpty(rules) {
  const snap = await getDocs(rulesCol())
  if (snap.empty) {
    const batch = writeBatch(db)
    rules.forEach(r => batch.set(ruleDoc(r.id), r))
    await batch.commit()
  }
}

// ── Semester management ───────────────────────────────────────────
export async function copySemester(fy, ft, ty, tt) {
  const from = await getDocs(studentsCol(fy, ft))
  if (from.empty) throw new Error('ไม่มีข้อมูลนักเรียนในเทอมต้นทาง')
  const to = await getDocs(studentsCol(ty, tt))
  if (!to.empty) throw new Error('เทอมปลายทางมีข้อมูลอยู่แล้ว กรุณาเคลียก่อน')
  const batch = writeBatch(db)
  from.docs.forEach(d => {
    const s = d.data()
    batch.set(studentDoc(ty, tt, s.id), { ...s, score: 100, transactions: [] })
  })
  await batch.commit()
}

export async function clearSemester(year, term) {
  const snap = await getDocs(studentsCol(year, term))
  const batch = writeBatch(db)
  snap.docs.forEach(d => batch.delete(d.ref))
  await batch.commit()
}

export async function resetScores(year, term) {
  const snap = await getDocs(studentsCol(year, term))
  const batch = writeBatch(db)
  snap.docs.forEach(d => batch.update(d.ref, { score: 100, transactions: [] }))
  await batch.commit()
}
