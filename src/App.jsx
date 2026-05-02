// ════════════════════════════════════════════════════════════════
//  App.js — ระบบคะแนนพฤติกรรมนักเรียน
//  โรงเรียนด่านช้างวิทยา
//  พัฒนาโดย นายณัฐพล กลุ่มกลัด
// ════════════════════════════════════════════════════════════════

import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  listenConfig, listenStudents, listenRules,
  setActiveSemester, updateStudentScore,
  bulkImportStudents, saveRule, deleteRule as fbDeleteRule,
  seedRulesIfEmpty, copySemester, clearSemester, resetScores,
} from "./firebase";

// ════════════════════════════════════════════════════════════════
//  ✏️  ตั้งค่าระบบ — แก้ได้ตรงนี้
// ════════════════════════════════════════════════════════════════
const CREDENTIALS = {
  admin:   { password: "admin2569",   role: "admin",   name: "ผู้ดูแลระบบ" },
  teacher: { password: "teacher2569", role: "teacher", name: "ครู" },
};

const SCHOOL_NAME = "โรงเรียนด่านช้างวิทยา";
const DEV_NAME    = "นายณัฐพล กลุ่มกลัด";

// ─── ระเบียบเริ่มต้น ───────────────────────────────────────────
const INITIAL_RULES = [
  { id: 1,  name: "มาสาย",                     type: "deduct", points: 5,  category: "วินัย"    },
  { id: 2,  name: "ขาดเรียนโดยไม่มีใบลา",     type: "deduct", points: 10, category: "วินัย"    },
  { id: 3,  name: "แต่งกายไม่ถูกระเบียบ",     type: "deduct", points: 5,  category: "วินัย"    },
  { id: 4,  name: "ใช้โทรศัพท์ในห้องเรียน",   type: "deduct", points: 10, category: "วินัย"    },
  { id: 5,  name: "ทะเลาะวิวาท",              type: "deduct", points: 20, category: "พฤติกรรม" },
  { id: 6,  name: "ทำลายทรัพย์สินโรงเรียน",   type: "deduct", points: 15, category: "พฤติกรรม" },
  { id: 7,  name: "สูบบุหรี่/ดื่มแอลกอฮอล์", type: "deduct", points: 30, category: "พฤติกรรม" },
  { id: 8,  name: "ทำดีเด่น/ช่วยเหลือสังคม", type: "add",    points: 10, category: "ความดี"   },
  { id: 9,  name: "ได้รับรางวัลระดับโรงเรียน", type: "add",   points: 15, category: "ความดี"   },
  { id: 10, name: "ได้รับรางวัลระดับจังหวัด",  type: "add",   points: 20, category: "ความดี"   },
  { id: 11, name: "เข้าร่วมกิจกรรมจิตอาสา",   type: "add",   points: 5,  category: "ความดี"   },
  { id: 12, name: "ส่งการบ้านครบทุกวิชา",      type: "add",   points: 5,  category: "ความดี"   },
];

const CLASSES = [
  "ม.1/1","ม.1/2","ม.1/3","ม.2/1","ม.2/2","ม.2/3",
  "ม.3/1","ม.3/2","ม.3/3","ม.4/1","ม.4/2","ม.4/3","ม.5/1","ม.5/2","ม.5/3","ม.6/1","ม.6/2","ม.6/3",
];

const CATEGORY_COLORS = {
  "วินัย":    "#3b82f6",
  "พฤติกรรม": "#ef4444",
  "ความดี":   "#22c55e",
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const getRiskLevel = (s) => s < 0 ? "danger" : s === 0 ? "warning" : s <= 30 ? "low" : "normal";
const getRiskLabel = (s) => s < 0 ? "เด็กมีปัญหา" : s === 0 ? "เด็กเสี่ยง" : s <= 30 ? "ระวัง" : "ปกติ";
const today  = () => new Date().toLocaleDateString("th-TH",
  { year:"numeric", month:"2-digit", day:"2-digit" }).replace(/\//g,"-");
const semLabel = (y, t) => `เทอม ${t} ปีการศึกษา ${y}`;

// ─── Excel ────────────────────────────────────────────────────────────────────
function exportToExcel(data, filename, sheet = "Sheet1") {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheet);
  ws["!cols"] = Object.keys(data[0]||{}).map(() => ({ wch: 22 }));
  XLSX.writeFile(wb, filename);
}

const exportStudentReport = (students, year, term) => exportToExcel(
  students.map(s => ({
    "รหัสนักเรียน": s.studentId, "ชื่อ-นามสกุล": s.name, "ห้องเรียน": s.class,
    "คะแนนปัจจุบัน": s.score, "สถานะ": getRiskLabel(s.score),
    "ครั้งที่บันทึก": (s.transactions||[]).length,
  })),
  `คะแนนพฤติกรรม_เทอม${term}_${year}_${today()}.xlsx`, "คะแนน"
);

const exportRiskReport = (students, year, term) => {
  const rows = students.filter(s => s.score <= 30).sort((a,b) => a.score - b.score)
    .map(s => ({
      "รหัสนักเรียน": s.studentId, "ชื่อ-นามสกุล": s.name,
      "ห้องเรียน": s.class, "คะแนน": s.score, "สถานะ": getRiskLabel(s.score),
    }));
  if (!rows.length) { alert("ไม่มีนักเรียนเสี่ยงในขณะนี้"); return; }
  exportToExcel(rows, `นักเรียนเสี่ยง_เทอม${term}_${year}_${today()}.xlsx`, "เสี่ยง");
};

const exportTxReport = (students, year, term) => {
  const rows = [];
  students.forEach(s => (s.transactions||[]).forEach(tx => rows.push({
    "รหัสนักเรียน": s.studentId, "ชื่อ-นามสกุล": s.name, "ห้องเรียน": s.class,
    "ระเบียบ": tx.ruleName, "ประเภท": tx.type==="deduct"?"หักคะแนน":"เพิ่มคะแนน",
    "คะแนน": tx.points, "หมายเหตุ": tx.note||"-",
    "ครูผู้บันทึก": tx.teacherName, "วันที่": tx.date, "เวลา": tx.time,
  })));
  if (!rows.length) { alert("ยังไม่มีประวัติการบันทึก"); return; }
  exportToExcel(rows, `ประวัติ_เทอม${term}_${year}_${today()}.xlsx`, "ประวัติ");
};

const downloadTemplate = () => exportToExcel([
  { "รหัสนักเรียน":"66001","ชื่อ-นามสกุล":"สมชาย ใจดี",   "ห้องเรียน":"ม.1/1" },
  { "รหัสนักเรียน":"66002","ชื่อ-นามสกุล":"สมหญิง รักดี", "ห้องเรียน":"ม.1/1" },
  { "รหัสนักเรียน":"66003","ชื่อ-นามสกุล":"วิชัย มีสุข",  "ห้องเรียน":"ม.1/2" },
], "Template_นำเข้านักเรียน.xlsx", "นักเรียน");

function parseExcel(file, cb) {
  const r = new FileReader();
  r.onload = (e) => {
    try {
      const wb  = XLSX.read(e.target.result, { type:"binary" });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval:"" });
      const students = raw.map((row, i) => {
        const sid  = String(row["รหัสนักเรียน"]||row["รหัส"]||"").trim();
        const name = String(row["ชื่อ-นามสกุล"]||row["ชื่อนามสกุล"]||"").trim();
        const cls  = String(row["ห้องเรียน"]||row["ห้อง"]||"").trim();
        if (!name) return null;
        return {
          id: `${Date.now()}_${i}`,
          studentId: sid || `STU${String(i+1).padStart(4,"0")}`,
          name, class: cls||"ไม่ระบุ", score: 100, transactions: [],
        };
      }).filter(Boolean);
      cb({ ok: true, students, count: students.length });
    } catch { cb({ ok: false, error: "ไม่สามารถอ่านไฟล์ได้ กรุณาตรวจสอบรูปแบบ" }); }
  };
  r.readAsBinaryString(file);
}

// ════════════════════════════════════════════════════════════════
//  ROOT APP
// ════════════════════════════════════════════════════════════════
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [students,    setStudents]    = useState([]);
  const [rules,       setRules]       = useState(INITIAL_RULES);
  const [semConfig,   setSemConfig]   = useState(null);
  const [dbReady,     setDbReady]     = useState(false);
  const [dbError,     setDbError]     = useState(false);

  // ── โหลด config + rules ทันทีตอนเปิดแอป (ไม่รอ login)
  useEffect(() => {
    // timeout 10 วินาที ถ้า Firebase ไม่ตอบ → แสดง error
    const timeout = setTimeout(() => setDbError(true), 10000);

    seedRulesIfEmpty(INITIAL_RULES).catch(console.error);

    const unsubCfg = listenConfig((cfg) => {
      clearTimeout(timeout);
      setSemConfig(cfg);
    });
    const unsubRules = listenRules((r) => {
      setRules(r.length ? r : INITIAL_RULES);
    });
    return () => { unsubCfg(); unsubRules(); clearTimeout(timeout); };
  }, []);

  // ── โหลด students เมื่อ login แล้ว + semConfig พร้อม
  useEffect(() => {
    if (!currentUser || !semConfig) return;
    setDbReady(false);
    const timeout = setTimeout(() => setDbReady(true), 8000); // fallback
    const unsub = listenStudents(semConfig.year, semConfig.term, (data) => {
      clearTimeout(timeout);
      setStudents(data);
      setDbReady(true);
    });
    return () => { unsub(); clearTimeout(timeout); };
  }, [currentUser, semConfig?.year, semConfig?.term]);

  // แสดง Login ก่อนเสมอ — ไม่รอ Firebase
  if (!currentUser) return <LoginScreen onLogin={setCurrentUser} firebaseReady={!!semConfig} />;

  // หลัง login แล้ว ถ้า Firebase ยังไม่พร้อม → แสดง loading (แต่มี timeout)
  if (!semConfig || !dbReady) return <LoadingScreen error={dbError} />;

  return currentUser.role === "admin"
    ? <AdminDashboard user={currentUser} students={students} rules={rules}
        semConfig={semConfig} onLogout={() => { setCurrentUser(null); setDbReady(false); }} />
    : <TeacherDashboard user={currentUser} students={students} rules={rules}
        semConfig={semConfig} onLogout={() => { setCurrentUser(null); setDbReady(false); }} />;
}

function LoadingScreen({ error }) {
  const [dots, setDots] = useState(".");
  useEffect(() => {
    const t = setInterval(() => setDots(d => d.length >= 3 ? "." : d + "."), 500);
    return () => clearInterval(t);
  }, []);

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center",
      justifyContent:"center", background:"linear-gradient(135deg,#0f1f3d,#1a3a6b)",
      fontFamily:"'Sarabun',sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700;800&display=swap" rel="stylesheet"/>
      <div style={{ textAlign:"center", color:"#fff", padding:32 }}>
        {error ? (
          <>
            <div style={{ fontSize:52, marginBottom:16 }}>⚠️</div>
            <p style={{ fontSize:20, fontWeight:700, marginBottom:12 }}>เชื่อมต่อ Firebase ไม่ได้</p>
            <div style={{ background:"rgba(255,255,255,0.08)", borderRadius:14,
              padding:"16px 24px", marginBottom:20, textAlign:"left", maxWidth:360 }}>
              <p style={{ fontSize:13, lineHeight:2, margin:0, opacity:0.85 }}>
                🔧 วิธีแก้:<br/>
                1. ตรวจสอบ firebaseConfig ใน <code style={{background:"rgba(255,255,255,0.15)",padding:"1px 6px",borderRadius:4}}>src/firebase.js</code><br/>
                2. ตรวจสอบ Firestore เปิดใช้งานแล้ว<br/>
                3. ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต
              </p>
            </div>
            <button onClick={() => window.location.reload()}
              style={{ padding:"10px 28px", background:"#f59e0b", border:"none",
                borderRadius:10, color:"#fff", fontSize:15, fontWeight:700,
                cursor:"pointer", fontFamily:"'Sarabun',sans-serif" }}>
              🔄 ลองใหม่
            </button>
          </>
        ) : (
          <>
            <div style={{ width:60, height:60, margin:"0 auto 20px",
              border:"4px solid rgba(255,255,255,0.15)",
              borderTop:"4px solid #f59e0b", borderRadius:"50%",
              animation:"spin 1s linear infinite" }}/>
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
            <p style={{ fontSize:18, fontWeight:600 }}>กำลังโหลดข้อมูล{dots}</p>
            <p style={{ fontSize:13, opacity:0.5, marginTop:8 }}>กำลังเชื่อมต่อ Firebase</p>
          </>
        )}
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════
//  LOGIN
// ════════════════════════════════════════════════════════════════
function LoginScreen({ onLogin, firebaseReady }) {
  const [role,     setRole]     = useState("teacher");
  const [password, setPassword] = useState("");
  const [error,    setError]    = useState("");
  const [loading,  setLoading]  = useState(false);

  const handleLogin = () => {
    if (!password) { setError("กรุณากรอกรหัสผ่าน"); return; }
    setLoading(true);
    setTimeout(() => {
      const cred = CREDENTIALS[role];
      if (cred && cred.password === password) {
        onLogin({ role: cred.role, name: cred.name });
      } else {
        setError("รหัสผ่านไม่ถูกต้อง");
        setLoading(false);
      }
    }, 400);
  };

  return (
    <div style={{ minHeight:"100vh",
      background:"linear-gradient(135deg,#0f1f3d 0%,#1a3a6b 50%,#0d2847 100%)",
      display:"flex", alignItems:"center", justifyContent:"center",
      fontFamily:"'Sarabun',sans-serif", padding:20 }}>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700;800&display=swap" rel="stylesheet"/>

      {/* Glow effects */}
      <div style={{ position:"fixed", width:500, height:500, borderRadius:"50%", top:-150, right:-150,
        background:"radial-gradient(circle, rgba(245,158,11,0.06) 0%, transparent 70%)", pointerEvents:"none" }}/>
      <div style={{ position:"fixed", width:400, height:400, borderRadius:"50%", bottom:-100, left:-100,
        background:"radial-gradient(circle, rgba(59,130,246,0.08) 0%, transparent 70%)", pointerEvents:"none" }}/>

      <div style={{ background:"rgba(255,255,255,0.05)", backdropFilter:"blur(20px)",
        border:"1px solid rgba(255,255,255,0.12)", borderRadius:24,
        padding:"48px 40px", width:"100%", maxWidth:420,
        boxShadow:"0 25px 50px rgba(0,0,0,0.5)", position:"relative", zIndex:1 }}>

        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ width:76, height:76, borderRadius:22, margin:"0 auto 16px",
            background:"linear-gradient(135deg,#f59e0b,#f97316)",
            display:"flex", alignItems:"center", justifyContent:"center",
            fontSize:34, boxShadow:"0 8px 28px rgba(245,158,11,0.45)" }}>⭐</div>
          <h1 style={{ color:"#fff", fontSize:22, fontWeight:800, margin:0 }}>ระบบคะแนนพฤติกรรม</h1>
          <p style={{ color:"#fbbf24", fontSize:15, marginTop:8, fontWeight:700 }}>{SCHOOL_NAME}</p>
          <p style={{ color:"rgba(255,255,255,0.35)", fontSize:12, marginTop:3 }}>
            สำนักงานเขตพื้นที่การศึกษามัธยมศึกษา
          </p>
        </div>

        {/* Role toggle */}
        <div style={{ display:"flex", background:"rgba(255,255,255,0.08)",
          borderRadius:14, padding:4, marginBottom:24, gap:4 }}>
          {[["teacher","👩‍🏫 เข้าสู่ระบบ (ครู)"],["admin","🔑 ผู้ดูแลระบบ"]].map(([r,l])=>(
            <button key={r} onClick={()=>{ setRole(r); setError(""); setPassword(""); }}
              style={{ flex:1, padding:"11px 0", border:"none", borderRadius:10,
                background:role===r?"rgba(255,255,255,0.18)":"transparent",
                color:role===r?"#fff":"rgba(255,255,255,0.45)",
                fontWeight:role===r?700:500, fontSize:13, cursor:"pointer",
                fontFamily:"'Sarabun',sans-serif", transition:"all 0.2s" }}>
              {l}
            </button>
          ))}
        </div>

        {/* Info box */}
        {role === "teacher" && (
          <div style={{ padding:"12px 16px", background:"rgba(59,130,246,0.15)",
            borderRadius:12, marginBottom:18, border:"1px solid rgba(59,130,246,0.3)" }}>
            <p style={{ margin:0, color:"rgba(255,255,255,0.75)", fontSize:13, lineHeight:1.6 }}>
              👥 ครูทุกคนใช้ <strong style={{color:"#93c5fd"}}>รหัสผ่านเดียวกัน</strong><br/>
              <span style={{ opacity:0.6, fontSize:12 }}>ติดต่อแอดมินหากต้องการรีเซ็ตรหัสผ่าน</span>
            </p>
          </div>
        )}

        {/* Password input */}
        <div style={{ marginBottom:10 }}>
          <label style={{ color:"rgba(255,255,255,0.7)", fontSize:13, fontWeight:600,
            display:"block", marginBottom:8 }}>รหัสผ่าน</label>
          <input type="password" value={password}
            onChange={e=>{ setPassword(e.target.value); setError(""); }}
            onKeyDown={e=>e.key==="Enter"&&handleLogin()}
            placeholder="กรอกรหัสผ่าน"
            style={{ width:"100%", padding:"13px 16px", boxSizing:"border-box",
              background:"rgba(255,255,255,0.09)", border:"1px solid rgba(255,255,255,0.18)",
              borderRadius:12, color:"#fff", fontSize:15, outline:"none",
              fontFamily:"'Sarabun',sans-serif" }}/>
        </div>

        {error && (
          <div style={{ padding:"9px 14px", background:"rgba(239,68,68,0.2)",
            borderRadius:10, marginBottom:10, border:"1px solid rgba(239,68,68,0.3)" }}>
            <p style={{ color:"#fca5a5", fontSize:13, margin:0 }}>⚠️ {error}</p>
          </div>
        )}

        <button onClick={handleLogin} disabled={loading} style={{
          width:"100%", padding:"14px", marginTop:8,
          background:loading?"rgba(245,158,11,0.4)":"linear-gradient(135deg,#f59e0b,#f97316)",
          border:"none", borderRadius:12, color:"#fff", fontSize:16,
          fontWeight:700, cursor:loading?"not-allowed":"pointer",
          fontFamily:"'Sarabun',sans-serif",
          boxShadow:loading?"none":"0 4px 20px rgba(245,158,11,0.35)" }}>
          {loading ? "กำลังเข้าสู่ระบบ..." : "เข้าสู่ระบบ"}
        </button>

        <div style={{ marginTop:24, textAlign:"center",
          borderTop:"1px solid rgba(255,255,255,0.08)", paddingTop:18 }}>
          <p style={{ margin:0, fontSize:11, color:"rgba(255,255,255,0.25)" }}>พัฒนาโดย</p>
          <p style={{ margin:"5px 0 8px", fontSize:13, color:"rgba(255,255,255,0.5)", fontWeight:600 }}>
            {DEV_NAME}
          </p>
          <div style={{ display:"inline-flex", alignItems:"center", gap:6,
            padding:"4px 12px", borderRadius:20,
            background: firebaseReady ? "rgba(34,197,94,0.15)" : "rgba(251,191,36,0.15)",
            border: `1px solid ${firebaseReady ? "rgba(34,197,94,0.3)" : "rgba(251,191,36,0.3)"}` }}>
            <div style={{ width:7, height:7, borderRadius:"50%",
              background: firebaseReady ? "#22c55e" : "#fbbf24",
              boxShadow: firebaseReady ? "0 0 6px #22c55e" : "0 0 6px #fbbf24" }}/>
            <span style={{ fontSize:11, color: firebaseReady ? "#86efac" : "#fde68a" }}>
              {firebaseReady ? "Firebase พร้อม" : "กำลังเชื่อมต่อ..."}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════
//  ADMIN DASHBOARD
// ════════════════════════════════════════════════════════════════
function AdminDashboard({ user, students, rules, semConfig, onLogout }) {
  const [tab,         setTab]         = useState("overview");
  const [ruleModal,   setRuleModal]   = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [filterRisk,  setFilterRisk]  = useState("all");
  const [filterClass, setFilterClass] = useState("all");
  const [search,      setSearch]      = useState("");
  const [rForm,       setRForm]       = useState({ name:"", type:"deduct", points:5, category:"วินัย" });
  const [saving,      setSaving]      = useState(false);

  const stats = useMemo(() => ({
    total:  students.length,
    normal: students.filter(s=>s.score>30).length,
    low:    students.filter(s=>s.score>0&&s.score<=30).length,
    risk:   students.filter(s=>s.score===0).length,
    danger: students.filter(s=>s.score<0).length,
    avg:    students.length ? Math.round(students.reduce((a,s)=>a+s.score,0)/students.length) : 100,
  }), [students]);

  const filtered = useMemo(() =>
    students.filter(s => {
      if (filterRisk!=="all" && getRiskLevel(s.score)!==filterRisk) return false;
      if (filterClass!=="all" && s.class!==filterClass) return false;
      if (search && !s.name.includes(search) && !s.studentId.includes(search)) return false;
      return true;
    }).sort((a,b) => a.score - b.score)
  , [students, filterRisk, filterClass, search]);

  const saveRuleHandler = async () => {
    if (!rForm.name || !rForm.points) return;
    setSaving(true);
    await saveRule({ ...rForm, points:Number(rForm.points),
      id: ruleModal==="new" ? Date.now() : ruleModal.id });
    setSaving(false);
    setRuleModal(null);
  };

  const TABS = [["overview","📋 รายชื่อ"],["rules","📜 ระเบียบ"],["semester","📅 จัดการเทอม"]];

  return (
    <AppShell user={user} onLogout={onLogout} role="admin" semConfig={semConfig}>
      {/* Stats */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))", gap:12, marginBottom:22 }}>
        {[["👥","ทั้งหมด",stats.total,"#3b82f6"],["✅","ปกติ (>30)",stats.normal,"#22c55e"],
          ["⚠️","ระวัง (1-30)",stats.low,"#f59e0b"],["🔶","เสี่ยง (0)",stats.risk,"#f97316"],
          ["🔴","มีปัญหา (-)",stats.danger,"#ef4444"],["📊","คะแนนเฉลี่ย",stats.avg,"#8b5cf6"]
        ].map(([icon,label,val,color])=>(
          <div key={label} style={{ background:"#fff", borderRadius:14, padding:"14px 12px",
            border:`2px solid ${color}22`, boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize:22 }}>{icon}</div>
            <div style={{ fontSize:28, fontWeight:800, color, lineHeight:1.2, marginTop:2 }}>{val}</div>
            <div style={{ fontSize:11, color:"#6b7280", marginTop:3 }}>{label}</div>
          </div>
        ))}
      </div>

      {/* Tabs */}
      <div style={{ display:"flex", gap:8, marginBottom:18, flexWrap:"wrap" }}>
        {TABS.map(([k,l])=>(
          <button key={k} onClick={()=>setTab(k)} style={{
            padding:"9px 20px", borderRadius:10, fontWeight:600, fontSize:14,
            background:tab===k?"#1a3a6b":"#fff", color:tab===k?"#fff":"#374151",
            border:tab===k?"none":"1px solid #e5e7eb", cursor:"pointer",
            fontFamily:"'Sarabun',sans-serif", transition:"all 0.2s" }}>{l}</button>
        ))}
      </div>

      {tab==="overview" && (
        <>
          <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap" }}>
            <Btn onClick={()=>setImportModal(true)}                                      color="#059669" icon="📥" label="นำเข้า Excel"/>
            <Btn onClick={()=>exportStudentReport(students,semConfig.year,semConfig.term)} color="#2563eb" icon="📊" label="Export คะแนน"/>
            <Btn onClick={()=>exportRiskReport(students,semConfig.year,semConfig.term)}    color="#dc2626" icon="🔴" label="Export เสี่ยง"/>
            <Btn onClick={()=>exportTxReport(students,semConfig.year,semConfig.term)}      color="#7c3aed" icon="📋" label="Export ประวัติ"/>
          </div>
          <div style={{ display:"flex", gap:10, marginBottom:14, flexWrap:"wrap" }}>
            <input placeholder="🔍 ค้นหาชื่อ / รหัส" value={search}
              onChange={e=>setSearch(e.target.value)} style={fs}/>
            <select value={filterClass} onChange={e=>setFilterClass(e.target.value)} style={fs}>
              <option value="all">ทุกห้องเรียน</option>
              {CLASSES.map(c=><option key={c}>{c}</option>)}
            </select>
            <select value={filterRisk} onChange={e=>setFilterRisk(e.target.value)} style={fs}>
              <option value="all">ทุกสถานะ</option>
              <option value="danger">เด็กมีปัญหา</option>
              <option value="warning">เด็กเสี่ยง</option>
              <option value="low">ระวัง</option>
              <option value="normal">ปกติ</option>
            </select>
            <span style={{ padding:"8px 14px", background:"#f3f4f6", borderRadius:10,
              fontSize:13, color:"#6b7280", alignSelf:"center", whiteSpace:"nowrap" }}>
              {filtered.length} คน
            </span>
          </div>
          <StudentTable students={filtered} showClass/>
        </>
      )}

      {tab==="rules" && (
        <>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
            <h3 style={{ margin:0, fontSize:16, fontWeight:700, color:"#1a3a6b" }}>
              ระเบียบหัก/เพิ่มคะแนน ({rules.length} รายการ)
            </h3>
            <button onClick={()=>{ setRForm({name:"",type:"deduct",points:5,category:"วินัย"}); setRuleModal("new"); }}
              style={btnP}>+ เพิ่มระเบียบ</button>
          </div>
          <RuleTable rules={rules}
            onEdit={r=>{ setRForm({...r}); setRuleModal(r); }}
            onDelete={async id=>{ if(window.confirm("ลบระเบียบนี้?")) await fbDeleteRule(id); }}/>
        </>
      )}

      {tab==="semester" && (
        <SemesterManager semConfig={semConfig} students={students}/>
      )}

      {/* Rule Modal */}
      {ruleModal && (
        <Modal title={ruleModal==="new"?"เพิ่มระเบียบ":"แก้ไขระเบียบ"} onClose={()=>setRuleModal(null)}>
          <FF label="ชื่อระเบียบ">
            <input value={rForm.name} onChange={e=>setRForm({...rForm,name:e.target.value})}
              placeholder="เช่น มาสาย, ทำดีเด่น" style={is}/>
          </FF>
          <FF label="ประเภท">
            <select value={rForm.type} onChange={e=>setRForm({...rForm,type:e.target.value})} style={is}>
              <option value="deduct">หักคะแนน</option>
              <option value="add">เพิ่มคะแนน</option>
            </select>
          </FF>
          <FF label="จำนวนคะแนน">
            <input type="number" min={1} value={rForm.points}
              onChange={e=>setRForm({...rForm,points:e.target.value})} style={is}/>
          </FF>
          <FF label="หมวดหมู่">
            <select value={rForm.category} onChange={e=>setRForm({...rForm,category:e.target.value})} style={is}>
              <option>วินัย</option><option>พฤติกรรม</option><option>ความดี</option>
            </select>
          </FF>
          <BtnRow>
            <button onClick={saveRuleHandler} disabled={saving}
              style={saving ? btnD : btnP}>{saving?"บันทึก...":"บันทึก"}</button>
            <button onClick={()=>setRuleModal(null)} style={btnS}>ยกเลิก</button>
          </BtnRow>
        </Modal>
      )}

      {importModal && (
        <ImportModal onClose={()=>setImportModal(false)}
          onImport={async(s,mode)=>{
            await bulkImportStudents(semConfig.year,semConfig.term,s,mode,students);
            setImportModal(false);
          }}/>
      )}
    </AppShell>
  );
}

// ─── Semester Manager ─────────────────────────────────────────────────────────
function SemesterManager({ semConfig, students }) {
  const curYear = new Date().getFullYear() + 543;
  const years   = [curYear-1, curYear, curYear+1];
  const [tYear, setTYear] = useState(semConfig.year);
  const [tTerm, setTTerm] = useState(semConfig.term);
  const [busy,  setBusy]  = useState(false);
  const [msg,   setMsg]   = useState(null);

  const run = async (label, fn) => {
    if (!window.confirm(`ยืนยัน: ${label}?`)) return;
    setBusy(true); setMsg(null);
    try { await fn(); setMsg({ ok:true,  text:`✅ ${label} สำเร็จ` }); }
    catch(e) {      setMsg({ ok:false, text:`❌ ${e.message}` }); }
    setBusy(false);
  };

  const nextY = semConfig.term===1 ? semConfig.year     : semConfig.year+1;
  const nextT = semConfig.term===1 ? 2                  : 1;

  return (
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>

      {/* Active semester banner */}
      <div style={{ gridColumn:"1/-1", background:"#fff", borderRadius:16,
        padding:"20px 24px", boxShadow:"0 2px 10px rgba(0,0,0,0.06)",
        display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:16 }}>
        <div>
          <div style={{ fontSize:12, color:"#6b7280", marginBottom:4 }}>เทอมที่เปิดใช้งานอยู่</div>
          <div style={{ fontSize:24, fontWeight:800, color:"#f59e0b" }}>
            {semLabel(semConfig.year, semConfig.term)}
          </div>
          <div style={{ fontSize:13, color:"#6b7280", marginTop:4 }}>
            มีนักเรียน {students.length} คน
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
          <span style={{ fontSize:13, fontWeight:600, color:"#374151" }}>สลับไปเทอม:</span>
          <select value={tYear} onChange={e=>setTYear(Number(e.target.value))} style={fs}>
            {years.map(y=><option key={y} value={y}>ปีการศึกษา {y}</option>)}
          </select>
          <select value={tTerm} onChange={e=>setTTerm(Number(e.target.value))} style={fs}>
            <option value={1}>เทอม 1</option><option value={2}>เทอม 2</option>
          </select>
          <button disabled={busy}
            onClick={async()=>{
              setBusy(true);
              await setActiveSemester(tYear,tTerm);
              setMsg({ ok:true, text:`✅ สลับไป ${semLabel(tYear,tTerm)} แล้ว` });
              setBusy(false);
            }}
            style={busy?btnD:{...btnP, flex:"initial", padding:"10px 18px"}}>
            สลับเทอม
          </button>
        </div>
      </div>

      {/* Cards */}
      <SemCard icon="🔄" title="เลื่อนขึ้นเทอมใหม่"
        desc={`คัดลอกรายชื่อนักเรียนทั้งหมดไป ${semLabel(nextY,nextT)} พร้อมรีเซ็ตคะแนนเป็น 100 ล้างประวัติทั้งหมด`}>
        <button disabled={busy}
          style={busy?btnD:{...btnP,background:"linear-gradient(135deg,#059669,#10b981)"}}
          onClick={()=>run(`เลื่อนจาก${semLabel(semConfig.year,semConfig.term)} → ${semLabel(nextY,nextT)}`,
            ()=>copySemester(semConfig.year,semConfig.term,nextY,nextT))}>
          เลื่อนขึ้นเทอมใหม่
        </button>
      </SemCard>

      <SemCard icon="♻️" title="รีเซ็ตคะแนน (คงรายชื่อ)"
        desc="คืนคะแนนทุกคนเป็น 100 และล้างประวัติการบันทึก โดยยังคงรายชื่อนักเรียนไว้">
        <button disabled={busy}
          style={busy?btnD:{...btnP,background:"linear-gradient(135deg,#f59e0b,#f97316)"}}
          onClick={()=>run(`รีเซ็ตคะแนน${semLabel(semConfig.year,semConfig.term)}`,
            ()=>resetScores(semConfig.year,semConfig.term))}>
          รีเซ็ตคะแนน
        </button>
      </SemCard>

      <SemCard icon="🗑️" title="เคลียข้อมูลสิ้นปีการศึกษา"
        desc="ลบรายชื่อและข้อมูลนักเรียนทั้งหมดของเทอมที่เลือก ใช้หลังสิ้นปีการศึกษา">
        <div style={{ display:"flex", gap:8, marginBottom:12, flexWrap:"wrap" }}>
          <select value={tYear} onChange={e=>setTYear(Number(e.target.value))} style={fs}>
            {years.map(y=><option key={y} value={y}>ปีการศึกษา {y}</option>)}
          </select>
          <select value={tTerm} onChange={e=>setTTerm(Number(e.target.value))} style={fs}>
            <option value={1}>เทอม 1</option><option value={2}>เทอม 2</option>
          </select>
        </div>
        <button disabled={busy}
          style={busy?btnD:{...btnP,background:"linear-gradient(135deg,#dc2626,#ef4444)"}}
          onClick={()=>run(`เคลียข้อมูล ${semLabel(tYear,tTerm)}`,()=>clearSemester(tYear,tTerm))}>
          เคลียข้อมูลเทอมที่เลือก
        </button>
      </SemCard>

      <SemCard icon="📖" title="คำแนะนำการใช้งาน" desc="">
        <div style={{ fontSize:13, color:"#374151", lineHeight:2 }}>
          <div>🟢 <b>ต้นเทอม 1</b> → นำเข้านักเรียนใหม่</div>
          <div>🔄 <b>ต้นเทอม 2</b> → กด "เลื่อนขึ้นเทอมใหม่"</div>
          <div>📊 <b>สิ้นเทอม</b> → Export รายงานก่อนเคลีย</div>
          <div>🗑️ <b>สิ้นปี</b> → เคลียข้อมูลทั้ง 2 เทอม</div>
        </div>
      </SemCard>

      {msg && (
        <div style={{ gridColumn:"1/-1", padding:"12px 16px", borderRadius:12, fontWeight:600,
          fontSize:14, background:msg.ok?"#f0fdf4":"#fef2f2",
          color:msg.ok?"#15803d":"#dc2626",
          border:`1px solid ${msg.ok?"#bbf7d0":"#fecaca"}` }}>
          {msg.text}
        </div>
      )}
    </div>
  );
}

function SemCard({ icon, title, desc, children }) {
  return (
    <div style={{ background:"#fff", borderRadius:16, padding:"20px 22px",
      boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
      <div style={{ fontSize:28, marginBottom:10 }}>{icon}</div>
      <h4 style={{ margin:"0 0 6px", fontSize:15, fontWeight:700, color:"#1a3a6b" }}>{title}</h4>
      {desc && <p style={{ margin:"0 0 16px", fontSize:13, color:"#6b7280", lineHeight:1.7 }}>{desc}</p>}
      {children}
    </div>
  );
}

// ════════════════════════════════════════════════════════════════
//  TEACHER DASHBOARD
// ════════════════════════════════════════════════════════════════
function TeacherDashboard({ user, students, rules, semConfig, onLogout }) {
  const [search,      setSearch]      = useState("");
  const [filterClass, setFilterClass] = useState("all");
  const [selectedId,  setSelectedId]  = useState(null);
  const [scoreModal,  setScoreModal]  = useState(false);
  const [sForm,       setSForm]       = useState({ ruleId:"", note:"" });
  const [toast,       setToast]       = useState(null);
  const [saving,      setSaving]      = useState(false);

  const filtered = useMemo(() => {
    if (!search && filterClass==="all") return [];
    return students.filter(s => {
      if (filterClass!=="all" && s.class!==filterClass) return false;
      if (search && !s.name.includes(search) && !s.studentId.includes(search) && !s.class.includes(search)) return false;
      return true;
    }).slice(0,60);
  }, [students, search, filterClass]);

  const selected = useMemo(() =>
    selectedId ? students.find(s=>s.id===selectedId)||null : null,
    [students, selectedId]);

  const showToast = (msg) => { setToast(msg); setTimeout(()=>setToast(null),3000); };

  const submitScore = async () => {
    const rule = rules.find(r=>String(r.id)===String(sForm.ruleId));
    if (!rule || !selected) return;
    setSaving(true);
    const pts = rule.type==="deduct" ? -rule.points : rule.points;
    const now = new Date();
    const tx  = {
      id: Date.now(), ruleId: rule.id, ruleName: rule.name, type: rule.type,
      points: pts, note: sForm.note, teacherName: user.name,
      date: now.toLocaleDateString("th-TH"),
      time: now.toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"}),
    };
    await updateStudentScore(semConfig.year, semConfig.term, selected.id,
      selected.score+pts, [tx,...(selected.transactions||[])]);
    setSaving(false);
    setScoreModal(false);
    showToast(`${rule.type==="deduct"?"หัก":"เพิ่ม"} ${rule.points} คะแนน — ${rule.name}`);
  };

  return (
    <AppShell user={user} onLogout={onLogout} role="teacher" semConfig={semConfig}>
      {toast && (
        <div style={{ position:"fixed", top:16, right:16, zIndex:9999,
          background:"#22c55e", color:"#fff", padding:"12px 20px", borderRadius:12,
          fontWeight:600, fontSize:14, boxShadow:"0 4px 20px rgba(0,0,0,0.2)" }}>
          ✅ {toast}
        </div>
      )}

      <div style={{ display:"grid", gridTemplateColumns:selected?"1fr 1.4fr":"1fr", gap:20 }}>
        {/* Search panel */}
        <div>
          <div style={{ background:"#fff", borderRadius:16, padding:20,
            boxShadow:"0 2px 10px rgba(0,0,0,0.06)", marginBottom:14 }}>
            <h3 style={{ margin:"0 0 14px", fontSize:15, fontWeight:700, color:"#1a3a6b" }}>
              🔍 ค้นหานักเรียน
            </h3>
            <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
              <input placeholder="พิมพ์ชื่อ หรือรหัสนักเรียน" value={search}
                onChange={e=>setSearch(e.target.value)} style={fs}/>
              <select value={filterClass} onChange={e=>setFilterClass(e.target.value)} style={fs}>
                <option value="all">— เลือกห้องเรียน —</option>
                {CLASSES.map(c=><option key={c}>{c}</option>)}
              </select>
            </div>
          </div>

          <div style={{ background:"#fff", borderRadius:16, overflow:"hidden",
            boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
            {filtered.length===0
              ? <div style={{ padding:"40px 20px", textAlign:"center", color:"#9ca3af", fontSize:14, lineHeight:2 }}>
                  {search||filterClass!=="all"
                    ? "🔍 ไม่พบนักเรียน"
                    : "พิมพ์ชื่อ หรือเลือกห้องเรียน\nเพื่อค้นหานักเรียน"}
                </div>
              : filtered.map(s=>(
                <div key={s.id} onClick={()=>setSelectedId(s.id)}
                  style={{ display:"flex", alignItems:"center", padding:"12px 16px",
                    cursor:"pointer", borderBottom:"1px solid #f3f4f6",
                    background:selected?.id===s.id?"#eff6ff":"#fff",
                    transition:"background 0.15s" }}>
                  <ScoreBadge score={s.score} size="sm"/>
                  <div style={{ flex:1, marginLeft:12 }}>
                    <div style={{ fontWeight:600, fontSize:14, color:"#111827" }}>{s.name}</div>
                    <div style={{ fontSize:12, color:"#6b7280" }}>{s.studentId} · {s.class}</div>
                  </div>
                  <RiskChip score={s.score}/>
                </div>
              ))
            }
          </div>
        </div>

        {/* Detail panel */}
        {selected && (
          <StudentDetail student={selected}
            onAddScore={()=>{ setSForm({ruleId:"",note:""}); setScoreModal(true); }}/>
        )}
      </div>

      {/* Score Modal */}
      {scoreModal && selected && (
        <Modal title={`📝 บันทึกคะแนน — ${selected.name}`} onClose={()=>setScoreModal(false)}>
          <FF label="เลือกระเบียบ">
            <select value={sForm.ruleId}
              onChange={e=>setSForm({...sForm,ruleId:e.target.value})} style={is}>
              <option value="">— กรุณาเลือกระเบียบ —</option>
              <optgroup label="🔴 หักคะแนน">
                {rules.filter(r=>r.type==="deduct").map(r=>(
                  <option key={r.id} value={r.id}>[−{r.points}] {r.name} ({r.category})</option>
                ))}
              </optgroup>
              <optgroup label="🟢 เพิ่มคะแนน">
                {rules.filter(r=>r.type==="add").map(r=>(
                  <option key={r.id} value={r.id}>[+{r.points}] {r.name} ({r.category})</option>
                ))}
              </optgroup>
            </select>
          </FF>
          {sForm.ruleId && (()=>{
            const r = rules.find(x=>String(x.id)===String(sForm.ruleId));
            if (!r) return null;
            const after = selected.score + (r.type==="deduct"?-r.points:r.points);
            return (
              <div style={{ padding:"11px 14px", borderRadius:10, marginBottom:4, fontSize:13,
                fontWeight:600, background:r.type==="deduct"?"#fef2f2":"#f0fdf4",
                color:r.type==="deduct"?"#dc2626":"#16a34a",
                border:`1px solid ${r.type==="deduct"?"#fecaca":"#bbf7d0"}` }}>
                {r.type==="deduct"?"หัก":"เพิ่ม"} {r.points} คะแนน → คะแนนใหม่: <strong>{after}</strong>
                {after<=0&&<span style={{marginLeft:8,fontSize:12}}>⚠️ {after<0?"เด็กมีปัญหา":"เด็กเสี่ยง"}</span>}
              </div>
            );
          })()}
          <FF label="หมายเหตุ (ไม่บังคับ)">
            <input value={sForm.note}
              onChange={e=>setSForm({...sForm,note:e.target.value})}
              placeholder="รายละเอียดเพิ่มเติม เช่น สาย 15 นาที" style={is}/>
          </FF>
          <BtnRow>
            <button onClick={submitScore} disabled={!sForm.ruleId||saving}
              style={(!sForm.ruleId||saving)?btnD:btnP}>
              {saving?"กำลังบันทึก...":"บันทึก"}
            </button>
            <button onClick={()=>setScoreModal(false)} style={btnS}>ยกเลิก</button>
          </BtnRow>
        </Modal>
      )}
    </AppShell>
  );
}

// ════════════════════════════════════════════════════════════════
//  IMPORT MODAL
// ════════════════════════════════════════════════════════════════
function ImportModal({ onClose, onImport }) {
  const ref = useRef();
  const [file,    setFile]    = useState(null);
  const [preview, setPreview] = useState(null);
  const [mode,    setMode]    = useState("merge");
  const [error,   setError]   = useState("");
  const [loading, setLoading] = useState(false);
  const [saving,  setSaving]  = useState(false);

  const handle = (f) => {
    if (!f) return;
    setFile(f); setError(""); setPreview(null); setLoading(true);
    parseExcel(f, (r) => {
      setLoading(false);
      if (!r.ok) { setError(r.error); return; }
      if (!r.count) { setError("ไม่พบข้อมูลนักเรียนในไฟล์"); return; }
      setPreview(r);
    });
  };

  return (
    <Modal title="📥 นำเข้ารายชื่อนักเรียนจาก Excel" onClose={onClose}>
      {/* Template download */}
      <div style={{ padding:"11px 14px", background:"#eff6ff", borderRadius:10,
        marginBottom:14, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div>
          <div style={{ fontSize:13, fontWeight:700, color:"#1d4ed8" }}>📄 ดาวน์โหลด Template</div>
          <div style={{ fontSize:11, color:"#6b7280", marginTop:2 }}>
            คอลัมน์: รหัสนักเรียน | ชื่อ-นามสกุล | ห้องเรียน
          </div>
        </div>
        <button onClick={downloadTemplate}
          style={{ padding:"7px 14px", background:"#2563eb", color:"#fff", border:"none",
            borderRadius:8, fontSize:12, fontWeight:600, cursor:"pointer",
            fontFamily:"'Sarabun',sans-serif" }}>
          ดาวน์โหลด
        </button>
      </div>

      {/* Drop zone */}
      <div onClick={()=>ref.current.click()}
        onDragOver={e=>e.preventDefault()}
        onDrop={e=>{ e.preventDefault(); handle(e.dataTransfer.files[0]); }}
        style={{ border:"2px dashed #d1d5db", borderRadius:12, padding:"28px 16px",
          textAlign:"center", cursor:"pointer", background:"#fafafa", marginBottom:14,
          transition:"border-color 0.2s" }}>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv"
          style={{ display:"none" }} onChange={e=>handle(e.target.files[0])}/>
        <div style={{ fontSize:32, marginBottom:8 }}>📂</div>
        <div style={{ fontSize:14, fontWeight:600, color:"#374151" }}>
          {file ? `📄 ${file.name}` : "คลิก หรือลากไฟล์มาวางที่นี่"}
        </div>
        <div style={{ fontSize:12, color:"#9ca3af", marginTop:4 }}>รองรับ .xlsx .xls .csv</div>
      </div>

      {loading && (
        <div style={{ textAlign:"center", color:"#6b7280", fontSize:14, marginBottom:12 }}>
          ⏳ กำลังอ่านไฟล์...
        </div>
      )}
      {error && (
        <div style={{ color:"#dc2626", fontSize:13, padding:"10px 14px",
          background:"#fef2f2", borderRadius:10, marginBottom:12 }}>⚠️ {error}</div>
      )}

      {preview && (
        <>
          <div style={{ padding:"12px 14px", background:"#f0fdf4", borderRadius:10,
            border:"1px solid #bbf7d0", marginBottom:14 }}>
            <div style={{ fontSize:14, fontWeight:700, color:"#15803d", marginBottom:6 }}>
              ✅ พบนักเรียน {preview.count} คน
            </div>
            <div style={{ fontSize:12, color:"#6b7280", marginBottom:4 }}>ตัวอย่าง 3 รายการแรก:</div>
            {preview.students.slice(0,3).map((s,i)=>(
              <div key={i} style={{ fontSize:12, color:"#374151", marginTop:3 }}>
                • {s.studentId} — {s.name} ({s.class})
              </div>
            ))}
            {preview.count>3 && (
              <div style={{ fontSize:12, color:"#9ca3af", marginTop:4 }}>
                ...และอีก {preview.count-3} คน
              </div>
            )}
          </div>

          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:13, fontWeight:700, color:"#374151", marginBottom:10 }}>
              เลือกโหมดการนำเข้า:
            </div>
            {[
              ["merge","➕ เพิ่มเข้าของเดิม","ข้ามนักเรียนที่มีรหัสซ้ำ"],
              ["replace","🔄 แทนที่ทั้งหมด","ลบรายชื่อเดิมทั้งหมดแล้วนำเข้าใหม่"],
            ].map(([v,l,sub])=>(
              <label key={v} style={{ display:"flex", alignItems:"flex-start", gap:10,
                padding:"11px 14px", borderRadius:10, marginBottom:8, cursor:"pointer",
                background:mode===v?"#eff6ff":"#f9fafb",
                border:`1.5px solid ${mode===v?"#93c5fd":"#e5e7eb"}` }}>
                <input type="radio" value={v} checked={mode===v} onChange={()=>setMode(v)}
                  style={{ accentColor:"#2563eb", marginTop:2 }}/>
                <div>
                  <div style={{ fontSize:13, fontWeight:mode===v?700:500, color:"#111827" }}>{l}</div>
                  <div style={{ fontSize:11, color:"#6b7280", marginTop:2 }}>{sub}</div>
                </div>
              </label>
            ))}
          </div>

          <BtnRow>
            <button onClick={async()=>{ setSaving(true); await onImport(preview.students,mode); setSaving(false); }}
              disabled={saving}
              style={saving?btnD:{...btnP,background:"linear-gradient(135deg,#059669,#10b981)"}}>
              {saving?"กำลังบันทึก...":"นำเข้าเลย"}
            </button>
            <button onClick={onClose} style={btnS}>ยกเลิก</button>
          </BtnRow>
        </>
      )}
    </Modal>
  );
}

// ════════════════════════════════════════════════════════════════
//  SHARED COMPONENTS
// ════════════════════════════════════════════════════════════════
function AppShell({ user, onLogout, role, semConfig, children }) {
  return (
    <div style={{ minHeight:"100vh", background:"#f1f5f9", fontFamily:"'Sarabun',sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700;800&display=swap" rel="stylesheet"/>
      <nav style={{ background:"linear-gradient(135deg,#0f1f3d,#1a3a6b)", padding:"0 20px",
        height:60, display:"flex", alignItems:"center", justifyContent:"space-between",
        boxShadow:"0 2px 12px rgba(0,0,0,0.3)", position:"sticky", top:0, zIndex:100 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
          <span style={{ fontSize:20 }}>⭐</span>
          <span style={{ color:"#fff", fontWeight:700, fontSize:14 }}>ระบบคะแนนพฤติกรรม</span>
          <span style={{ padding:"2px 10px", borderRadius:20, fontSize:11, fontWeight:700,
            background:role==="admin"?"rgba(245,158,11,0.3)":"rgba(59,130,246,0.3)",
            color:role==="admin"?"#fbbf24":"#93c5fd" }}>
            {role==="admin"?"แอดมิน":"ครู"}
          </span>
          {semConfig && (
            <span style={{ padding:"2px 10px", borderRadius:20, fontSize:11, fontWeight:600,
              background:"rgba(255,255,255,0.12)", color:"rgba(255,255,255,0.7)" }}>
              📅 {semLabel(semConfig.year,semConfig.term)}
            </span>
          )}
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span style={{ color:"rgba(255,255,255,0.6)", fontSize:13 }}>{user.name}</span>
          <button onClick={onLogout}
            style={{ padding:"6px 14px", borderRadius:8, background:"rgba(255,255,255,0.1)",
              color:"#fff", border:"1px solid rgba(255,255,255,0.2)", fontSize:12,
              cursor:"pointer", fontFamily:"'Sarabun',sans-serif" }}>
            ออกจากระบบ
          </button>
        </div>
      </nav>
      <div style={{ padding:"20px", maxWidth:1200, margin:"0 auto" }}>{children}</div>
    </div>
  );
}

function StudentDetail({ student, onAddScore }) {
  const risk = getRiskLevel(student.score);
  const rc = {
    normal:  { bg:"#f0fdf4", border:"#bbf7d0", text:"#15803d" },
    low:     { bg:"#fffbeb", border:"#fde68a", text:"#d97706" },
    warning: { bg:"#fff7ed", border:"#fdba74", text:"#ea580c" },
    danger:  { bg:"#fef2f2", border:"#fecaca", text:"#dc2626" },
  }[risk];
  return (
    <div style={{ background:"#fff", borderRadius:16, overflow:"hidden",
      boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
      <div style={{ background:"linear-gradient(135deg,#1a3a6b,#2563eb)",
        padding:"18px 22px", color:"#fff" }}>
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ width:50, height:50, borderRadius:"50%",
            background:"rgba(255,255,255,0.2)", display:"flex",
            alignItems:"center", justifyContent:"center", fontSize:22 }}>👤</div>
          <div>
            <div style={{ fontWeight:700, fontSize:18 }}>{student.name}</div>
            <div style={{ fontSize:13, opacity:0.7 }}>{student.studentId} · {student.class}</div>
          </div>
        </div>
      </div>
      <div style={{ padding:"18px 22px", borderBottom:"1px solid #f3f4f6" }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <div>
            <div style={{ fontSize:12, color:"#6b7280", marginBottom:4 }}>คะแนนพฤติกรรม</div>
            <div style={{ fontSize:50, fontWeight:800, lineHeight:1,
              color:risk==="danger"?"#dc2626":risk==="warning"?"#ea580c":"#1a3a6b" }}>
              {student.score}
            </div>
            <div style={{ fontSize:12, color:"#9ca3af", marginTop:4 }}>จาก 100 คะแนน</div>
          </div>
          <div style={{ padding:"10px 18px", borderRadius:20, fontWeight:700, fontSize:14,
            background:rc.bg, border:`1px solid ${rc.border}`, color:rc.text }}>
            {risk==="danger"?"🔴":risk==="warning"?"🔶":risk==="low"?"⚠️":"✅"} {getRiskLabel(student.score)}
          </div>
        </div>
        <div style={{ marginTop:14, height:10, background:"#f3f4f6", borderRadius:10, overflow:"hidden" }}>
          <div style={{ height:"100%", borderRadius:10, transition:"width 0.5s",
            width:`${Math.max(0,Math.min(100,student.score))}%`,
            background:risk==="danger"?"#ef4444":risk==="warning"?"#f97316":risk==="low"?"#f59e0b":"#22c55e" }}/>
        </div>
        <button onClick={onAddScore}
          style={{ ...btnP, marginTop:14, width:"100%", display:"block" }}>
          + บันทึกคะแนน
        </button>
      </div>
      <div style={{ padding:"14px 22px" }}>
        <div style={{ fontSize:14, fontWeight:700, color:"#374151", marginBottom:12 }}>
          ประวัติการบันทึก ({(student.transactions||[]).length} รายการ)
        </div>
        <div style={{ maxHeight:280, overflowY:"auto" }}>
          {!(student.transactions||[]).length
            ? <div style={{ textAlign:"center", padding:"24px 0", color:"#9ca3af", fontSize:13 }}>
                ยังไม่มีประวัติการบันทึก
              </div>
            : (student.transactions||[]).map(tx=>(
              <div key={tx.id} style={{ display:"flex", alignItems:"flex-start", gap:12,
                padding:"10px 0", borderBottom:"1px solid #f9fafb" }}>
                <div style={{ width:30, height:30, borderRadius:"50%", flexShrink:0,
                  background:tx.type==="deduct"?"#fef2f2":"#f0fdf4",
                  color:tx.type==="deduct"?"#dc2626":"#16a34a",
                  display:"flex", alignItems:"center", justifyContent:"center",
                  fontSize:15, fontWeight:800 }}>
                  {tx.type==="deduct"?"−":"+"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:13, fontWeight:600, color:"#111827" }}>{tx.ruleName}</div>
                  {tx.note&&<div style={{ fontSize:12, color:"#6b7280", marginTop:2 }}>📝 {tx.note}</div>}
                  <div style={{ fontSize:11, color:"#9ca3af", marginTop:2 }}>
                    {tx.teacherName} · {tx.date} {tx.time}
                  </div>
                </div>
                <div style={{ fontWeight:800, fontSize:15,
                  color:tx.points<0?"#dc2626":"#16a34a", flexShrink:0 }}>
                  {tx.points>0?"+":""}{tx.points}
                </div>
              </div>
            ))
          }
        </div>
      </div>
    </div>
  );
}

function StudentTable({ students, showClass }) {
  const [page, setPage] = useState(1);
  const PER = 25;
  const pages = Math.ceil(students.length/PER);
  const slice = students.slice((page-1)*PER, page*PER);

  return (
    <div>
      <div style={{ background:"#fff", borderRadius:16, overflow:"hidden",
        boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
            <thead>
              <tr style={{ background:"#f8fafc", borderBottom:"2px solid #e5e7eb" }}>
                {["รหัส","ชื่อ-นามสกุล",showClass&&"ห้อง","คะแนน","สถานะ"].filter(Boolean).map(h=>(
                  <th key={h} style={{ padding:"12px 14px", textAlign:"left",
                    fontWeight:700, color:"#374151", whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {slice.map((s,i)=>(
                <tr key={s.id} style={{ borderBottom:"1px solid #f3f4f6",
                  background:i%2===0?"#fff":"#fafafa" }}>
                  <td style={{ padding:"10px 14px", color:"#6b7280" }}>{s.studentId}</td>
                  <td style={{ padding:"10px 14px", fontWeight:600, color:"#111827" }}>{s.name}</td>
                  {showClass&&<td style={{ padding:"10px 14px", color:"#6b7280" }}>{s.class}</td>}
                  <td style={{ padding:"10px 14px" }}><ScoreBadge score={s.score} size="sm"/></td>
                  <td style={{ padding:"10px 14px" }}><RiskChip score={s.score}/></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      {pages>1 && (
        <div style={{ display:"flex", gap:5, justifyContent:"center", marginTop:14 }}>
          <PB onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}>‹</PB>
          {Array.from({length:Math.min(7,pages)},(_,i)=>{
            const p = pages<=7?i+1 : page<=4?i+1 : page>=pages-3?pages-6+i : page-3+i;
            return <PB key={p} onClick={()=>setPage(p)} active={p===page}>{p}</PB>;
          })}
          <PB onClick={()=>setPage(p=>Math.min(pages,p+1))} disabled={page===pages}>›</PB>
        </div>
      )}
    </div>
  );
}

function RuleTable({ rules, onEdit, onDelete }) {
  return (
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
      {[["deduct","🔴 ระเบียบหักคะแนน"],["add","🟢 ระเบียบเพิ่มคะแนน"]].map(([type,title])=>(
        <div key={type} style={{ background:"#fff", borderRadius:14, overflow:"hidden",
          boxShadow:"0 2px 10px rgba(0,0,0,0.06)" }}>
          <div style={{ padding:"12px 16px", fontWeight:700, fontSize:13, color:"#fff",
            background:type==="deduct"?"linear-gradient(135deg,#dc2626,#ef4444)":"linear-gradient(135deg,#16a34a,#22c55e)" }}>
            {title} ({rules.filter(r=>r.type===type).length})
          </div>
          {rules.filter(r=>r.type===type).map(r=>(
            <div key={r.id} style={{ display:"flex", alignItems:"center",
              padding:"11px 14px", borderBottom:"1px solid #f3f4f6" }}>
              <div style={{ width:34, height:34, borderRadius:9, flexShrink:0,
                background:type==="deduct"?"#fef2f2":"#f0fdf4",
                color:type==="deduct"?"#dc2626":"#16a34a",
                display:"flex", alignItems:"center", justifyContent:"center",
                fontWeight:800, fontSize:13 }}>
                {type==="deduct"?"−":"+"}{r.points}
              </div>
              <div style={{ flex:1, marginLeft:10 }}>
                <div style={{ fontSize:13, fontWeight:600 }}>{r.name}</div>
                <span style={{ fontSize:11, padding:"2px 8px", borderRadius:20, fontWeight:600,
                  background:(CATEGORY_COLORS[r.category]||"#6b7280")+"22",
                  color:CATEGORY_COLORS[r.category]||"#6b7280" }}>{r.category}</span>
              </div>
              <div style={{ display:"flex", gap:5 }}>
                <button onClick={()=>onEdit(r)}
                  style={{ padding:"4px 10px", borderRadius:7, fontSize:12, fontWeight:600,
                    background:"#eff6ff", color:"#2563eb", border:"none", cursor:"pointer",
                    fontFamily:"'Sarabun',sans-serif" }}>แก้ไข</button>
                <button onClick={()=>onDelete(r.id)}
                  style={{ padding:"4px 10px", borderRadius:7, fontSize:12, fontWeight:600,
                    background:"#fef2f2", color:"#dc2626", border:"none", cursor:"pointer",
                    fontFamily:"'Sarabun',sans-serif" }}>ลบ</button>
              </div>
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

function ScoreBadge({ score, size="md" }) {
  const risk = getRiskLevel(score);
  const c = { normal:"#22c55e", low:"#f59e0b", warning:"#f97316", danger:"#ef4444" }[risk];
  const sz = size==="sm" ? 36 : 48;
  return (
    <div style={{ width:sz, height:sz, borderRadius:"50%", display:"inline-flex",
      alignItems:"center", justifyContent:"center", flexShrink:0,
      background:c+"22", color:c, fontWeight:800,
      fontSize:size==="sm"?13:16, border:`2px solid ${c}44` }}>{score}</div>
  );
}

function RiskChip({ score }) {
  const c = {
    normal:  { bg:"#f0fdf4", color:"#16a34a", label:"ปกติ"    },
    low:     { bg:"#fffbeb", color:"#d97706", label:"ระวัง"   },
    warning: { bg:"#fff7ed", color:"#ea580c", label:"เสี่ยง"  },
    danger:  { bg:"#fef2f2", color:"#dc2626", label:"มีปัญหา" },
  }[getRiskLevel(score)];
  return (
    <span style={{ padding:"3px 10px", borderRadius:20, fontSize:12, fontWeight:700,
      background:c.bg, color:c.color, whiteSpace:"nowrap" }}>{c.label}</span>
  );
}

function Modal({ title, children, onClose }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)",
      display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}
      onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{ background:"#fff", borderRadius:20, padding:"26px 26px",
        width:"100%", maxWidth:480, boxShadow:"0 20px 60px rgba(0,0,0,0.35)",
        maxHeight:"92vh", overflowY:"auto" }}>
        <div style={{ display:"flex", justifyContent:"space-between",
          alignItems:"center", marginBottom:18 }}>
          <h3 style={{ margin:0, fontSize:17, fontWeight:700, color:"#1a3a6b" }}>{title}</h3>
          <button onClick={onClose}
            style={{ width:30, height:30, borderRadius:"50%", border:"none",
              background:"#f3f4f6", cursor:"pointer", fontSize:16, lineHeight:1 }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

const FF = ({ label, children }) => (
  <div style={{ marginBottom:14 }}>
    <label style={{ display:"block", fontSize:13, fontWeight:600, color:"#374151", marginBottom:6 }}>{label}</label>
    {children}
  </div>
);
const BtnRow  = ({ children }) => <div style={{ display:"flex", gap:10, marginTop:20 }}>{children}</div>;
const Btn = ({ onClick, color, icon, label }) => (
  <button onClick={onClick} style={{ padding:"8px 16px", borderRadius:10, fontSize:13, fontWeight:600,
    background:`linear-gradient(135deg,${color},${color}cc)`, color:"#fff",
    border:"none", cursor:"pointer", fontFamily:"'Sarabun',sans-serif" }}>
    {icon} {label}
  </button>
);
function PB({ children, onClick, disabled, active }) {
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ padding:"6px 12px", borderRadius:8, fontSize:13,
        cursor:disabled?"default":"pointer",
        border:active?"none":"1px solid #e5e7eb",
        background:active?"#1a3a6b":disabled?"#f9fafb":"#fff",
        color:active?"#fff":disabled?"#d1d5db":"#374151",
        fontFamily:"'Sarabun',sans-serif" }}>{children}</button>
  );
}

// ── Shared styles ────────────────────────────────────────────────
const btnP = { flex:1, padding:"11px", fontFamily:"'Sarabun',sans-serif",
  background:"linear-gradient(135deg,#1a3a6b,#2563eb)",
  color:"#fff", border:"none", borderRadius:10, fontSize:15, fontWeight:700, cursor:"pointer" };
const btnS = { flex:1, padding:"11px", fontFamily:"'Sarabun',sans-serif",
  background:"#f3f4f6", color:"#374151", border:"none",
  borderRadius:10, fontSize:15, fontWeight:600, cursor:"pointer" };
const btnD = { flex:1, padding:"11px", fontFamily:"'Sarabun',sans-serif",
  background:"#9ca3af", color:"#fff", border:"none",
  borderRadius:10, fontSize:15, fontWeight:700, cursor:"not-allowed" };
const is = { width:"100%", padding:"10px 14px", boxSizing:"border-box",
  border:"1px solid #e5e7eb", borderRadius:10, fontSize:14, outline:"none",
  fontFamily:"'Sarabun',sans-serif", background:"#fafafa" };
const fs = { padding:"8px 13px", borderRadius:10, border:"1px solid #e5e7eb",
  fontSize:14, fontFamily:"'Sarabun',sans-serif", outline:"none",
  background:"#fff", flex:1, minWidth:150 };
