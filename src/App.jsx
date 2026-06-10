import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  getPCs, savePC, deletePC, swapPC,
  getSoftwares, saveSoftware, deleteSoftware,
  getDesks, saveDesk, deleteDesk,
  getTableGroups, saveTableGroups,
  getFloorPlan, saveFloorPlan,
  getAccounts, updateUserProfile, registerNewUser,
  loginUser, logoutUser, getUserProfile,
  getCostReport,
  getInvoices, uploadInvoice, deleteInvoice,
} from "./db.js";
import { auth, onAuthStateChanged } from "./firebase.js";

// ─── Themes ───────────────────────────────────────────────────────────────────
const THEMES = {
  dark:  { bg:"#181b27", panel:"#20243a", card:"#272c45", border:"#363d60", accent:"#6aa3fc", accent2:"#fbb360", green:"#5de89a", red:"#fc6b6b", yellow:"#fde96a", text:"#e8ecfa", muted:"#8892b8", dim:"#4a527a", inputBg:"#20243a", shadow:"0 4px 24px #0007" },
  light: { bg:"#f0f3fa", panel:"#ffffff", card:"#f6f8ff", border:"#d0d8f0", accent:"#2563eb", accent2:"#d97706", green:"#16a34a", red:"#dc2626", yellow:"#ca8a04", text:"#1a2040", muted:"#4a5280", dim:"#9aa0c0", inputBg:"#ffffff", shadow:"0 2px 12px #0002" },
};

// ─── License helpers ──────────────────────────────────────────────────────────
// חישוב עלות שנתית אפקטיבית לפי סוג רישיון
function calcAnnualCost(sw) {
  if (sw.licenseType === "annual")   return sw.price || 0;
  if (sw.licenseType === "monthly")  return (sw.price || 0) * (sw.months || 12);
  return 0; // permanent
}

// ─── Tiny UI ──────────────────────────────────────────────────────────────────
const Tag = ({ color, children }) => (
  <span style={{ background:color+"22", color, border:`1px solid ${color}55`, borderRadius:4, padding:"2px 8px", fontSize:11, fontWeight:700 }}>{children}</span>
);

const InputField = ({ label, value, onChange, type="text", placeholder, C, error, min }) => (
  <div style={{ marginBottom:12 }}>
    {label && <label style={{ display:"block", color:C.muted, fontSize:11, textTransform:"uppercase", letterSpacing:.8, marginBottom:4 }}>{label}</label>}
    <input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} min={min}
      style={{ width:"100%", background:C.inputBg, border:`1px solid ${error?C.red:C.border}`, borderRadius:7, padding:"9px 13px", color:C.text, fontSize:13, fontFamily:"inherit", outline:"none", boxSizing:"border-box",
        colorScheme: "dark" }}
      onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor=error?C.red:C.border} />
    {error && <div style={{ color:C.red, fontSize:11, marginTop:3 }}>⚠ {error}</div>}
  </div>
);

const DatePicker = ({ label, value, onChange, C }) => (
  <div style={{ marginBottom:12 }}>
    {label && <label style={{ display:"block", color:C.muted, fontSize:11, textTransform:"uppercase", letterSpacing:.8, marginBottom:4 }}>{label}</label>}
    <input type="date" value={value||""} onChange={e=>onChange(e.target.value)}
      style={{ width:"100%", background:C.inputBg, border:`1px solid ${C.border}`, borderRadius:7, padding:"9px 13px", color:C.text, fontSize:13, fontFamily:"inherit", outline:"none", boxSizing:"border-box", cursor:"pointer", colorScheme:"dark" }}
      onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor=C.border} />
  </div>
);

const Btn = ({ children, onClick, variant="primary", style:s={}, disabled }) => {
  const colors = { primary:"#6aa3fc", danger:"#fc6b6b", success:"#5de89a", neutral:"#8892b8", warning:"#fde96a" };
  const c = colors[variant]||colors.primary;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ background:c+"22", color:c, border:`1px solid ${c}55`, borderRadius:7, padding:"8px 16px", fontSize:13, fontWeight:700, cursor:disabled?"not-allowed":"pointer", fontFamily:"inherit", opacity:disabled?.5:1, transition:"all .15s", ...s }}
      onMouseEnter={e=>!disabled&&(e.currentTarget.style.background=c+"44")}
      onMouseLeave={e=>!disabled&&(e.currentTarget.style.background=c+"22")}>
      {children}
    </button>
  );
};

// Modal — סגירה רק בלחיצה על הרקע, לא על הטופס עצמו
const Modal = ({ children, onClose, C }) => (
  <div style={{ position:"fixed", inset:0, background:"#0009", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100, padding:"clamp(16px,3vw,40px)" }}
    onMouseDown={e=>{ if(e.target===e.currentTarget) onClose(); }}>
    <div style={{ maxWidth:700, width:"100%", maxHeight:"90vh", overflowY:"auto", borderRadius:14 }}
      onMouseDown={e=>e.stopPropagation()}>
      {children}
    </div>
  </div>
);

// ─── Confirm Dialog ──────────────────────────────────────────────────────────
function ConfirmDialog({ title, message, onConfirm, onCancel, C }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"#0009", display:"flex", alignItems:"center", justifyContent:"center", zIndex:200, padding:20 }}>
      <div style={{ background:C.panel, border:`1px solid ${C.red}55`, borderRadius:14, padding:"clamp(24px,3vw,36px)", width:"min(400px,100%)", boxShadow:C.shadow }}>
        <div style={{ fontSize:36, textAlign:"center", marginBottom:12 }}>⚠️</div>
        <div style={{ color:C.text, fontWeight:800, fontSize:16, textAlign:"center", marginBottom:8 }}>{title}</div>
        {message && <div style={{ color:C.muted, fontSize:13, textAlign:"center", marginBottom:24 }}>{message}</div>}
        <div style={{ display:"flex", gap:10, justifyContent:"center" }}>
          <Btn onClick={onConfirm} variant="danger" style={{ minWidth:100 }}>🗑️ מחק</Btn>
          <Btn onClick={onCancel} variant="neutral" style={{ minWidth:100 }}>ביטול</Btn>
        </div>
      </div>
    </div>
  );
}

// ─── License type badge ───────────────────────────────────────────────────────
function LicenseTag({ sw, C }) {
  if (sw.licenseType === "annual")   return <Tag color={C.accent2}>שנתי 🔄</Tag>;
  if (sw.licenseType === "monthly")  return <Tag color={C.yellow}>חודשי · {sw.months} חודשים</Tag>;
  return <Tag color={C.green}>קבוע 🔒</Tag>;
}

// ─── Login ────────────────────────────────────────────────────────────────────
function LoginScreen({ onLogin, C }) {
  const [email,setEmail]=useState(""); const [password,setPassword]=useState("");
  const [error,setError]=useState(""); const [loading,setLoading]=useState(false);
  const handle = async () => {
    if(!email||!password){setError("נדרשים אימייל וסיסמה");return;}
    setLoading(true);setError("");
    try { const u=await loginUser(email,password); onLogin(u); }
    catch(e){ setError(e.code==="auth/invalid-credential"||e.code==="auth/wrong-password"?"אימייל או סיסמה שגויים":"שגיאה: "+e.message); }
    setLoading(false);
  };
  return (
    <div style={{ background:C.bg, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'DM Mono','Courier New',monospace", padding:20, direction:"rtl" }}>
      <div style={{ background:C.panel, border:`1px solid ${C.border}`, borderRadius:18, padding:"clamp(28px,5vw,56px)", width:"min(420px,100%)", boxShadow:C.shadow }}>
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ fontSize:44, marginBottom:10 }}>⚙</div>
          <div style={{ color:C.accent, fontWeight:800, fontSize:"clamp(16px,2.5vw,22px)", letterSpacing:2 }}>OFFICE MANAGER</div>
        </div>
        <InputField label="אימייל" value={email} onChange={setEmail} type="email" placeholder="your@email.com" C={C} />
        <InputField label="סיסמה" value={password} onChange={setPassword} type="password" placeholder="••••••••" C={C} />
        {error && <div style={{ color:C.red, fontSize:12, marginBottom:14, padding:"10px 14px", background:C.red+"18", borderRadius:8, border:`1px solid ${C.red}44` }}>✗ {error}</div>}
        <button onClick={handle} disabled={loading}
          style={{ width:"100%", background:loading?C.accent+"88":C.accent, color:"#fff", border:"none", borderRadius:9, padding:"clamp(10px,2vh,14px)", fontSize:15, fontWeight:800, cursor:loading?"not-allowed":"pointer", fontFamily:"inherit", marginTop:4 }}>
          {loading?"מתחבר...":"כניסה →"}
        </button>
      </div>
    </div>
  );
}

// ─── PC Form ──────────────────────────────────────────────────────────────────
function PCForm({ initial, softwares, onSave, onCancel, C }) {
  const blank = { pcId:"", userName:"", operatingSystem:"", cpuModel:"", gpuModel:"", ramGb:"", ssdGb:"", comments:"", userMail:"", softwareIds:[] };
  const [form,setForm]=useState(initial?{...blank,...initial}:blank);
  const [saving,setSaving]=useState(false); const [error,setError]=useState(null);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const toggleSw=(id)=>setForm(f=>({ ...f, softwareIds: f.softwareIds.includes(id)?f.softwareIds.filter(x=>x!==id):[...f.softwareIds,id] }));

  const handle = async () => {
    if(!form.pcId.trim()){setError("PC ID נדרש");return;}
    setSaving(true);setError(null);
    try { await onSave(form); } catch(e){ setError(e.message); }
    setSaving(false);
  };

  const sec=(label,color)=>(
    <div style={{ color, fontSize:11, fontWeight:700, textTransform:"uppercase", letterSpacing:1, margin:"14px 0 8px", display:"flex", alignItems:"center", gap:8 }}>
      <div style={{ flex:1, height:1, background:color+"44" }}/>{label}<div style={{ flex:1, height:1, background:color+"44" }}/>
    </div>
  );

  return (
    <div style={{ background:C.card, borderRadius:14, padding:"clamp(18px,3vw,28px)" }}>
      <h3 style={{ color:C.text, margin:"0 0 18px", fontSize:17 }}>{initial?`✏️ ${initial.pcId}`:"➕ מחשב חדש"}</h3>
      {sec("זיהוי",C.accent)}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))", gap:10 }}>
        <InputField label="PC ID *" value={form.pcId} onChange={v=>set("pcId",v)} C={C} error={!form.pcId&&error?error:null} />
        <InputField label="שם משתמש" value={form.userName} onChange={v=>set("userName",v)} C={C} />
        <InputField label="מייל" value={form.userMail} onChange={v=>set("userMail",v)} type="email" C={C} />
        <InputField label="מערכת הפעלה" value={form.operatingSystem} onChange={v=>set("operatingSystem",v)} C={C} />
      </div>
      {sec("חומרה",C.accent2)}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))", gap:10 }}>
        <InputField label="מודל CPU" value={form.cpuModel} onChange={v=>set("cpuModel",v)} placeholder="Intel i7-13700K" C={C} />
        <InputField label="מודל GPU" value={form.gpuModel} onChange={v=>set("gpuModel",v)} placeholder="RTX 4060" C={C} />
        <InputField label="RAM (GB)" value={form.ramGb} onChange={v=>set("ramGb",v)} type="number" C={C} />
        <InputField label="SSD (GB)" value={form.ssdGb} onChange={v=>set("ssdGb",v)} type="number" C={C} />
      </div>
      {sec("תוכנות",C.green)}
      <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:14 }}>
        {softwares.length===0 && <div style={{ color:C.dim, fontSize:12 }}>אין תוכנות במאגר — הוסף תחילה בטאב תוכנות</div>}
        {softwares.map(sw=>{
          const sel=form.softwareIds.includes(sw.id);
          const annual = calcAnnualCost(sw);
          return (
            <button key={sw.id} onClick={()=>toggleSw(sw.id)}
              style={{ background:sel?C.green+"33":C.bg, border:`1.5px solid ${sel?C.green:C.border}`, color:sel?C.green:C.muted, borderRadius:7, padding:"6px 12px", cursor:"pointer", fontFamily:"inherit", fontSize:12, fontWeight:sel?700:400, transition:"all .15s", textAlign:"right" }}>
              {sel?"✓ ":""}{sw.companyName} — {sw.softwareName}
              <span style={{ color:sel?C.green:C.dim, fontSize:10, marginRight:6 }}>
                {sw.licenseType==="permanent"?"קבוע":sw.licenseType==="annual"?"שנתי":"חודשי"} · ₪{sw.price}
                {sw.licenseType==="monthly"&&` × ${sw.months}`}
              </span>
            </button>
          );
        })}
      </div>
      <div style={{ marginBottom:12 }}>
        <label style={{ color:C.muted, fontSize:11, textTransform:"uppercase", marginBottom:4, display:"block" }}>הערות</label>
        <textarea value={form.comments} onChange={e=>set("comments",e.target.value)}
          style={{ width:"100%", background:C.inputBg, border:`1px solid ${C.border}`, borderRadius:7, padding:"9px 13px", color:C.text, fontSize:13, fontFamily:"inherit", resize:"vertical", minHeight:56, boxSizing:"border-box" }} />
      </div>
      {error && <div style={{ color:C.red, fontSize:13, padding:"10px 14px", background:C.red+"18", border:`1px solid ${C.red}55`, borderRadius:8, marginBottom:12 }}>✗ {error}</div>}
      <div style={{ display:"flex", gap:8 }}>
        <Btn onClick={handle} variant="success" disabled={saving}>{saving?"שומר...":"💾 שמור"}</Btn>
        <Btn onClick={onCancel} variant="neutral">ביטול</Btn>
      </div>
    </div>
  );
}

// ─── PC Detail ────────────────────────────────────────────────────────────────
function PCDetail({ pc, softwares, onEdit, onDelete, onSwap, isAdmin, C }) {
  const pcSoftwares = softwares.filter(s=>pc.softwareIds?.includes(s.id));
  const hw=(label,val)=>(
    <div style={{ background:C.bg, borderRadius:8, padding:"10px 12px", border:`1px solid ${C.border}` }}>
      <div style={{ color:C.muted, fontSize:10, marginBottom:3, textTransform:"uppercase" }}>{label}</div>
      <div style={{ color:C.text, fontSize:13, fontWeight:700 }}>{val||"—"}</div>
    </div>
  );
  return (
    <div style={{ background:C.card, borderRadius:12, padding:18 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ color:C.text, fontSize:18, fontWeight:800 }}>{pc.userName||"לא מוקצה"}</div>
          <div style={{ color:C.muted, fontSize:12, fontFamily:"monospace" }}>{pc.pcId}</div>
          {pc.userMail && <div style={{ color:C.dim, fontSize:11, marginTop:2 }}>✉ {pc.userMail}</div>}
        </div>
        {isAdmin && (
          <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
            <Btn onClick={onSwap} variant="warning" style={{ padding:"5px 9px", fontSize:11 }}>🔄</Btn>
            <Btn onClick={onEdit} variant="primary" style={{ padding:"5px 9px", fontSize:11 }}>✏️</Btn>
            <Btn onClick={onDelete} variant="danger" style={{ padding:"5px 9px", fontSize:11 }}>🗑️</Btn>
          </div>
        )}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:14 }}>
        {hw("מערכת הפעלה", pc.operatingSystem)}
        {hw("CPU", pc.cpuModel)}
        {hw("GPU", pc.gpuModel)}
        {hw("RAM", pc.ramGb?`${pc.ramGb} GB`:null)}
        {hw("SSD", pc.ssdGb?`${pc.ssdGb} GB`:null)}
      </div>
      {pcSoftwares.length>0 && (
        <div>
          <div style={{ color:C.green, fontSize:10, fontWeight:700, textTransform:"uppercase", marginBottom:8 }}>תוכנות מותקנות</div>
          <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
            {pcSoftwares.map(sw=>(
              <div key={sw.id} style={{ background:C.bg, borderRadius:7, padding:"7px 10px", border:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div>
                  <span style={{ color:C.text, fontSize:12, fontWeight:700 }}>{sw.softwareName}</span>
                  <span style={{ color:C.muted, fontSize:11, marginRight:6 }}>{sw.companyName}</span>
                </div>
                <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                  <LicenseTag sw={sw} C={C}/>
                  <span style={{ color:C.muted, fontSize:11 }}>₪{sw.price}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      {pc.comments && <div style={{ marginTop:12, color:C.muted, fontSize:12, fontStyle:"italic", background:C.bg, padding:"8px 10px", borderRadius:6 }}>💬 {pc.comments}</div>}
    </div>
  );
}

// ─── Invoice Manager ─────────────────────────────────────────────────────────
function InvoiceManager({ softwareId, softwareName, onClose, isAdmin, C }) {
  const [invoices,setInvoices]=useState([]);
  const [loading,setLoading]=useState(true);
  const [uploading,setUploading]=useState(false);
  const [confirmDel,setConfirmDel]=useState(null);

  const load=async()=>{ setLoading(true); setInvoices(await getInvoices(softwareId)); setLoading(false); };
  useEffect(()=>{ load(); },[softwareId]);

  const handleUpload=async(e)=>{
    const files=[...e.target.files]; if(!files.length)return;
    setUploading(true);
    for(const file of files){ await uploadInvoice(softwareId,file); }
    await load();
    setUploading(false);
    e.target.value="";
  };

  const handleDelete=async(filePath,fileName)=>{
    setConfirmDel({filePath,fileName});
  };

  const ext=(name)=>name.split(".").pop().toLowerCase();
  const isImg=(name)=>["jpg","jpeg","png","gif","webp"].includes(ext(name));
  const isPdf=(name)=>ext(name)==="pdf";

  return (
    <div style={{background:C.card,borderRadius:14,padding:"clamp(18px,3vw,28px)",minWidth:300}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <div>
          <div style={{color:C.text,fontWeight:800,fontSize:16}}>📄 חשבוניות</div>
          <div style={{color:C.muted,fontSize:12}}>{softwareName}</div>
        </div>
        <button onClick={onClose} style={{background:C.dim+"33",border:"none",color:C.muted,cursor:"pointer",fontSize:18,width:32,height:32,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
      </div>

      {isAdmin && (
        <label style={{display:"flex",alignItems:"center",gap:10,background:C.accent+"18",border:`2px dashed ${C.accent}55`,borderRadius:10,padding:"clamp(12px,2vw,18px)",cursor:"pointer",marginBottom:18,justifyContent:"center"}}>
          <span style={{fontSize:24}}>📎</span>
          <div>
            <div style={{color:C.accent,fontWeight:700,fontSize:13}}>{uploading?"מעלה...":"העלה חשבוניות"}</div>
            <div style={{color:C.muted,fontSize:11}}>PDF, תמונות — ניתן לבחור מספר קבצים</div>
          </div>
          <input type="file" multiple accept=".pdf,.jpg,.jpeg,.png,.gif,.webp" onChange={handleUpload} disabled={uploading} style={{display:"none"}}/>
        </label>
      )}

      {loading && <div style={{color:C.muted,fontSize:13,textAlign:"center",padding:20}}>טוען...</div>}

      {!loading && invoices.length===0 && (
        <div style={{color:C.dim,fontSize:13,textAlign:"center",padding:"20px 0"}}>אין חשבוניות מצורפות</div>
      )}

      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {invoices.map(inv=>(
          <div key={inv.path} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:20}}>{isImg(inv.name)?"🖼️":isPdf(inv.name)?"📄":"📎"}</span>
            <div style={{flex:1,minWidth:0}}>
              <div style={{color:C.text,fontSize:12,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{inv.name}</div>
            </div>
            <div style={{display:"flex",gap:6,flexShrink:0}}>
              <a href={inv.url} target="_blank" rel="noreferrer"
                style={{background:C.accent+"22",color:C.accent,border:`1px solid ${C.accent}55`,borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,textDecoration:"none"}}>
                👁️ פתח
              </a>
              {isAdmin && (
                <Btn onClick={()=>handleDelete(inv.path,inv.name)} variant="danger" style={{padding:"4px 10px",fontSize:11}}>🗑️</Btn>
              )}
            </div>
          </div>
        ))}
      </div>

      {confirmDel && (
        <ConfirmDialog
          title={`למחוק "${confirmDel.fileName}"?`}
          message="הקובץ יימחק לצמיתות מהשרת."
          onConfirm={async()=>{ await deleteInvoice(confirmDel.filePath); setConfirmDel(null); load(); }}
          onCancel={()=>setConfirmDel(null)}
          C={C}
        />
      )}
    </div>
  );
}

// ─── Software Form ────────────────────────────────────────────────────────────
function SoftwareForm({ initial, onSave, onCancel, C }) {
  const blank = { companyName:"", softwareName:"", licenseType:"permanent", expiryDate:"", months:12, price:"", productKey:"" };
  const [form,setForm]=useState(initial?{...blank,...initial,months:initial.months||12}:blank);
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));

  // חשב מספר חודשים מתאריך תפוגה אוטומטית
  const handleExpiryDate=(date)=>{
    set("expiryDate",date);
    if(date && form.licenseType==="monthly"){
      const months = Math.max(1, Math.round((new Date(date)-new Date())/(1000*60*60*24*30)));
      set("months",months);
    }
  };

  const annualEquiv = form.licenseType==="monthly" ? (parseFloat(form.price)||0)*(parseInt(form.months)||12)
                    : form.licenseType==="annual"   ? (parseFloat(form.price)||0)
                    : 0;

  const handle=async()=>{
    if(!form.companyName||!form.softwareName){return;}
    setSaving(true);
    await onSave(form);
    setSaving(false);
  };

  const LicenseBtn=({val,label})=>(
    <button onClick={()=>set("licenseType",val)}
      style={{ flex:1, padding:"10px 8px", borderRadius:7, cursor:"pointer", fontFamily:"inherit", fontSize:12, fontWeight:700,
        background:form.licenseType===val?C.accent+"33":C.bg,
        color:form.licenseType===val?C.accent:C.muted,
        border:`1px solid ${form.licenseType===val?C.accent:C.border}` }}>
      {label}
    </button>
  );

  return (
    <div style={{ background:C.card, borderRadius:14, padding:"clamp(16px,2vw,24px)" }}>
      <h4 style={{ color:C.accent, margin:"0 0 16px" }}>{initial?.id?"✏️ עריכת תוכנה":"➕ תוכנה חדשה"}</h4>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:10 }}>
        <InputField label="שם חברה *" value={form.companyName} onChange={v=>set("companyName",v)} C={C} />
        <InputField label="שם תוכנה *" value={form.softwareName} onChange={v=>set("softwareName",v)} C={C} />

        <InputField label="מפתח מוצר (אופציונלי)" value={form.productKey} onChange={v=>set("productKey",v)} C={C} />
      </div>

      {/* סוג רישיון */}
      <div style={{ marginBottom:14 }}>
        <label style={{ color:C.muted, fontSize:11, textTransform:"uppercase", marginBottom:8, display:"block" }}>סוג רישיון</label>
        <div style={{ display:"flex", gap:8 }}>
          <LicenseBtn val="permanent" label="🔒 קבוע"/>
          <LicenseBtn val="annual"    label="🔄 שנתי"/>
          <LicenseBtn val="monthly"   label="📅 חודשי"/>
        </div>
      </div>

      {/* מחיר לפי סוג */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:10 }}>
        {form.licenseType==="permanent" && (
          <InputField label="מחיר רכישה (₪)" value={form.price} onChange={v=>set("price",v)} type="number" C={C} />
        )}
        {form.licenseType==="annual" && (
          <InputField label="מחיר לשנה (₪)" value={form.price} onChange={v=>set("price",v)} type="number" C={C} />
        )}
        {form.licenseType==="monthly" && (<>
          <InputField label="מחיר לחודש (₪)" value={form.price} onChange={v=>set("price",v)} type="number" C={C} />
          <InputField label="מספר חודשים" value={form.months} onChange={v=>set("months",parseInt(v)||1)} type="number" min="1" C={C} />
        </>)}
      </div>

      {/* תאריך תפוגה — שנתי וחודשי */}
      {(form.licenseType==="annual"||form.licenseType==="monthly") && (
        <DatePicker label="תאריך תפוגה (אופציונלי)" value={form.expiryDate} onChange={handleExpiryDate} C={C}/>
      )}

      {/* סיכום עלות שנתית לחודשי */}
      {form.licenseType==="monthly" && form.price && (
        <div style={{ background:C.yellow+"18", border:`1px solid ${C.yellow}44`, borderRadius:8, padding:"8px 14px", marginBottom:14, color:C.yellow, fontSize:12 }}>
          💡 עלות שנתית מקבילה: ₪{annualEquiv.toLocaleString()} ({form.price} × {form.months} חודשים)
        </div>
      )}

      <div style={{ display:"flex", gap:8 }}>
        <Btn onClick={handle} variant="success" disabled={saving||!form.companyName||!form.softwareName}>{saving?"שומר...":"💾 שמור"}</Btn>
        <Btn onClick={onCancel} variant="neutral">ביטול</Btn>
      </div>
    </div>
  );
}

// ─── Software Manager ─────────────────────────────────────────────────────────
function SoftwareManager({ softwares, pcs, onSave, onDelete, isAdmin, C }) {
  const [editSw,setEditSw]=useState(null);
  const [showForm,setShowForm]=useState(false);
  const [expandedSw,setExpandedSw]=useState(null);
  const [confirmDelete,setConfirmDelete]=useState(null);
  const [invoiceView,setInvoiceView]=useState(null);
  const [invoiceCounts,setInvoiceCounts]=useState({});

  // טען ספירת חשבוניות לכל תוכנה
  useEffect(()=>{
    if(!softwares.length)return;
    Promise.all(softwares.map(async sw=>{
      const inv=await getInvoices(sw.id);
      return [sw.id,inv.length];
    })).then(pairs=>setInvoiceCounts(Object.fromEntries(pairs)));
  },[softwares]);

  const openNew=()=>{ setEditSw(null); setShowForm(true); };
  const openEdit=(sw)=>{ setEditSw(sw); setShowForm(true); };
  const cancel=()=>{ setShowForm(false); setEditSw(null); };
  const handleSave=async(form)=>{ await onSave({...form,id:editSw?.id||null}); cancel(); };

  // מחשבים שמשתמשים בתוכנה
  const getPCsUsing=(swId)=>Object.values(pcs).filter(pc=>(pc.softwareIds||[]).includes(swId));

  const byCompany = softwares.reduce((acc,s)=>{ (acc[s.companyName]=acc[s.companyName]||[]).push(s); return acc; },{});
  const [collapsedCompanies,setCollapsedCompanies]=useState({});
  // collapse all companies by default when softwares load
  useEffect(()=>{
    if(!softwares.length)return;
    setCollapsedCompanies(prev=>{
      const next={...prev};
      Object.keys(byCompany).forEach(c=>{ if(!(c in next)) next[c]=true; });
      return next;
    });
  },[softwares.length]);
  const toggleCompany=(name)=>setCollapsedCompanies(c=>({...c,[name]:!c[name]}));

  return (
    <div style={{ padding:"clamp(14px,2vw,28px)" }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
        <h3 style={{ color:C.text, margin:0, fontSize:"clamp(14px,1.8vw,18px)" }}>💿 מאגר תוכנות</h3>
        {isAdmin && <Btn onClick={openNew} variant="success">➕ תוכנה חדשה</Btn>}
      </div>

      {showForm && isAdmin && (
        <div style={{ marginBottom:24 }}>
          <SoftwareForm initial={editSw} onSave={handleSave} onCancel={cancel} C={C}/>
        </div>
      )}

      {Object.keys(byCompany).length===0 && !showForm && (
        <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:12, padding:48, textAlign:"center" }}>
          <div style={{ fontSize:44, marginBottom:12 }}>💿</div>
          <div style={{ color:C.muted }}>אין תוכנות במאגר</div>
        </div>
      )}

      {Object.entries(byCompany).map(([company,sws])=>(
        <div key={company} style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:12, marginBottom:14, overflow:"hidden" }}>
          <div onClick={()=>toggleCompany(company)} style={{ padding:"12px 18px", background:C.border+"33", display:"flex", justifyContent:"space-between", alignItems:"center", cursor:"pointer" }}>
            <span style={{color:C.accent,fontWeight:700,fontSize:14}}>🏢 {company}</span>
            <span style={{color:C.muted,fontSize:13}}>{collapsedCompanies[company]?`${sws.length} תוכנות ▼`:"▲"}</span>
          </div>
          {!collapsedCompanies[company]&&sws.map(sw=>{
            const usingPCs = getPCsUsing(sw.id);
            const isExpanded = expandedSw === sw.id;
            return (
              <div key={sw.id} style={{ borderTop:`1px solid ${C.border}` }}>
                <div style={{ padding:"12px 18px", display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
                  <div style={{ flex:1 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                      <span style={{ color:C.text, fontWeight:700, fontSize:14 }}>{sw.softwareName}</span>
                      <LicenseTag sw={sw} C={C}/>
                      <span style={{ color:C.accent, fontWeight:700, fontSize:13 }}>₪{sw.price?.toLocaleString()}{sw.licenseType==="monthly"?"/חודש":sw.licenseType==="annual"?"/שנה":""}</span>
                      {sw.licenseType==="monthly" && <span style={{ color:C.yellow, fontSize:12 }}>= ₪{calcAnnualCost(sw).toLocaleString()}/שנה</span>}
                      {sw.expiryDate && <Tag color={C.yellow}>פג: {sw.expiryDate}</Tag>}
                    </div>
                    <div style={{ display:"flex", gap:12, marginTop:4, flexWrap:"wrap", alignItems:"center" }}>
                      {sw.productKey && <span style={{ color:C.muted, fontSize:11, fontFamily:"monospace" }}>🔑 {sw.productKey}</span>}
                      {sw.invoice    && <span style={{ color:C.muted, fontSize:11 }}>📄 {sw.invoice}</span>}
                      {/* כפתור מחשבים */}
                      <button onClick={()=>setExpandedSw(isExpanded?null:sw.id)}
                        style={{ background:usingPCs.length>0?C.green+"22":C.dim+"22", border:`1px solid ${usingPCs.length>0?C.green+"55":C.border}`, borderRadius:6, padding:"3px 10px", color:usingPCs.length>0?C.green:C.dim, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
                        🖥️ {usingPCs.length} מחשבים {isExpanded?"▲":"▼"}
                      </button>
                    </div>
                  </div>
                  <div style={{ display:"flex", gap:6 }}>
                      <Btn onClick={()=>setInvoiceView(sw)} variant="neutral" style={{ padding:"5px 10px", fontSize:11 }}>📄 {invoiceCounts[sw.id]>0?invoiceCounts[sw.id]:""}</Btn>
                      {isAdmin && <>
                        <Btn onClick={()=>openEdit(sw)} variant="primary" style={{ padding:"5px 10px", fontSize:11 }}>✏️</Btn>
                        <Btn onClick={()=>setConfirmDelete({id:sw.id,name:sw.softwareName,count:usingPCs.length})} variant="danger" style={{ padding:"5px 10px", fontSize:11 }}>🗑️</Btn>
                      </>}
                    </div>
                </div>
                {/* רשימת מחשבים שמשתמשים בתוכנה */}
                {isExpanded && (
                  <div style={{ padding:"0 18px 14px", borderTop:`1px solid ${C.border+"66"}` }}>
                    {usingPCs.length===0 ? (
                      <div style={{ color:C.dim, fontSize:12, padding:"10px 0" }}>אין מחשבים שמשתמשים בתוכנה זו</div>
                    ) : (
                      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:8, paddingTop:10 }}>
                        {usingPCs.map(pc=>(
                          <div key={pc.pcId} style={{ background:C.bg, border:`1px solid ${C.green}44`, borderRadius:8, padding:"8px 12px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                            <div>
                              <div style={{ color:C.text, fontSize:12, fontWeight:700 }}>{pc.userName||"לא מוקצה"}</div>
                              <div style={{ color:C.muted, fontSize:11, fontFamily:"monospace" }}>{pc.pcId}</div>
                            </div>
                            {pc.userMail && <span style={{ color:C.dim, fontSize:10 }}>✉</span>}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ))}

      {/* Confirm Delete Dialog */}
      {confirmDelete && (
        <ConfirmDialog
          title={`למחוק "${confirmDelete.name}"?`}
          message={confirmDelete.count>0 ? `התוכנה מותקנת על ${confirmDelete.count} מחשבים — היא תוסר מכולם.` : "פעולה זו אינה ניתנת לביטול."}
          onConfirm={()=>{ onDelete(confirmDelete.id); setConfirmDelete(null); }}
          onCancel={()=>setConfirmDelete(null)}
          C={C}
        />
      )}
      {/* Invoice Modal */}
      {invoiceView && (
        <Modal onClose={()=>{ setInvoiceView(null); getInvoices(invoiceView.id).then(inv=>setInvoiceCounts(c=>({...c,[invoiceView.id]:inv.length}))); }} C={C}>
          <InvoiceManager softwareId={invoiceView.id} softwareName={`${invoiceView.companyName} — ${invoiceView.softwareName}`} onClose={()=>{ setInvoiceView(null); getInvoices(invoiceView.id).then(inv=>setInvoiceCounts(c=>({...c,[invoiceView.id]:inv.length}))); }} isAdmin={isAdmin} C={C}/>
        </Modal>
      )}
    </div>
  );
}

// ─── Cost Report ──────────────────────────────────────────────────────────────
function CostReport({ pcs, softwares, isMobile, C }) {
  const [report,setReport]=useState([]);
  const [filter,setFilter]=useState("");
  const [sortBy,setSortBy]=useState("totalCost");
  const [expandedRow,setExpandedRow]=useState(null);

  useEffect(()=>{ getCostReport(pcs,softwares).then(setReport); },[pcs,softwares]);

  const totalAnnual    = report.reduce((s,r)=>s+r.annualCost,0);
  const totalPermanent = report.reduce((s,r)=>s+r.permanentCost,0);

  const filtered = report
    .filter(r=>!filter||r.userName?.toLowerCase().includes(filter.toLowerCase())||r.userMail?.toLowerCase().includes(filter.toLowerCase()))
    .sort((a,b)=>sortBy==="name"?(a.userName||"").localeCompare(b.userName||""):b[sortBy]-a[sortBy]);

  const exportReport=()=>{
    const rows=filtered.map(r=>({
      "שם משתמש":r.userName,"מייל":r.userMail,"PC ID":r.pcId,
      "עלות שנתית (₪)":r.annualCost,"עלות קבועה (₪)":r.permanentCost,"סה״כ (₪)":r.totalCost,
      "תוכנות":r.softwares.map(s=>s.softwareName).join(", "),
    }));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"דוח עלויות");
    XLSX.writeFile(wb,`cost-report-${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const swColor=(s)=>s.licenseType==="permanent"?C.green:s.licenseType==="annual"?C.accent2:C.yellow;

  return (
    <div style={{ padding:`clamp(14px,2vw,28px) clamp(14px,2vw,28px) ${isMobile?"80px":"28px"}` }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16, flexWrap:"wrap", gap:10 }}>
        <h3 style={{ color:C.text, margin:0, fontSize:"clamp(14px,1.8vw,18px)" }}>📊 דוח עלויות</h3>
        <Btn onClick={exportReport} variant="warning" style={{ fontSize:12 }}>📤 Excel</Btn>
      </div>

      {/* סיכום */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(auto-fill,minmax(180px,1fr))", gap:10, marginBottom:20 }}>
        {[
          {label:"שנתי",value:`₪${totalAnnual.toLocaleString()}`,color:C.accent2},
          {label:"קבוע",value:`₪${totalPermanent.toLocaleString()}`,color:C.green},
          {label:"סה״כ",value:`₪${(totalAnnual+totalPermanent).toLocaleString()}`,color:C.accent},
          {label:"עובדים",value:report.filter(r=>r.totalCost>0).length,color:C.muted},
        ].map(card=>(
          <div key={card.label} style={{ background:C.card, border:`1px solid ${card.color}44`, borderRadius:10, padding:"12px 14px" }}>
            <div style={{ color:C.muted, fontSize:10, textTransform:"uppercase", marginBottom:4 }}>{card.label}</div>
            <div style={{ color:card.color, fontSize:isMobile?18:22, fontWeight:800 }}>{card.value}</div>
          </div>
        ))}
      </div>

      {/* פילטר */}
      <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap" }}>
        <input value={filter} onChange={e=>setFilter(e.target.value)} placeholder="🔍 חיפוש.."
          style={{ flex:1, minWidth:140, background:C.inputBg, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", color:C.text, fontSize:13, fontFamily:"inherit", outline:"none" }}/>
        {!isMobile&&<select value={sortBy} onChange={e=>setSortBy(e.target.value)}
          style={{ background:C.inputBg, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", color:C.text, fontSize:13, fontFamily:"inherit", outline:"none" }}>
          <option value="totalCost">מיון: סה״כ</option>
          <option value="annualCost">מיון: שנתי</option>
          <option value="permanentCost">מיון: קבוע</option>
          <option value="name">מיון: שם</option>
        </select>}
      </div>

      {/* נייד: כרטיסים מתקפלים */}
      {isMobile ? (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {filtered.map(r=>{
            const open=expandedRow===r.pcId;
            return (
              <div key={r.pcId} style={{ background:C.card, border:`1px solid ${open?C.accent:C.border}`, borderRadius:10, overflow:"hidden" }}>
                <div onClick={()=>setExpandedRow(open?null:r.pcId)}
                  style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 14px", cursor:"pointer" }}>
                  <div>
                    <div style={{ color:C.text, fontWeight:800, fontSize:14 }}>{r.userName||"—"}</div>
                    <div style={{ color:C.dim, fontSize:11 }}>{r.pcId}</div>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                    <div style={{ textAlign:"left" }}>
                      <div style={{ color:r.totalCost>0?C.accent:C.dim, fontWeight:800, fontSize:15 }}>{r.totalCost>0?`₪${r.totalCost.toLocaleString()}`:"—"}</div>
                      <div style={{ color:C.muted, fontSize:10 }}>לשנה</div>
                    </div>
                    <span style={{ color:C.muted, fontSize:14 }}>{open?"▲":"▼"}</span>
                  </div>
                </div>
                {open && (
                  <div style={{ padding:"0 14px 14px", borderTop:`1px solid ${C.border}` }}>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8, margin:"12px 0" }}>
                      <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", border:`1px solid ${C.accent2}44` }}>
                        <div style={{ color:C.muted, fontSize:10, marginBottom:2 }}>שנתי</div>
                        <div style={{ color:C.accent2, fontWeight:700 }}>{r.annualCost>0?`₪${r.annualCost.toLocaleString()}`:"—"}</div>
                      </div>
                      <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", border:`1px solid ${C.green}44` }}>
                        <div style={{ color:C.muted, fontSize:10, marginBottom:2 }}>קבוע</div>
                        <div style={{ color:C.green, fontWeight:700 }}>{r.permanentCost>0?`₪${r.permanentCost.toLocaleString()}`:"—"}</div>
                      </div>
                      <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", border:`1px solid ${C.accent}44` }}>
                        <div style={{ color:C.muted, fontSize:10, marginBottom:2 }}>כולל</div>
                        <div style={{ color:C.accent, fontWeight:800 }}>{r.totalCost>0?`₪${r.totalCost.toLocaleString()}`:"—"}</div>
                      </div>
                    </div>
                    {r.softwares.length>0 && (
                      <div>
                        <div style={{ color:C.muted, fontSize:10, textTransform:"uppercase", marginBottom:6 }}>תוכנות</div>
                        <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                          {r.softwares.map(s=>(
                            <span key={s.id} style={{ background:swColor(s)+"22", color:swColor(s), borderRadius:4, padding:"3px 8px", fontSize:11, fontWeight:700 }}>
                              {s.softwareName} · ₪{s.price}
                            </span>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          {filtered.length===0&&<div style={{padding:40,textAlign:"center",color:C.muted}}>אין נתונים</div>}
        </div>
      ) : (
        /* דסקטופ: טבלה */
        <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:12, overflow:"auto" }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 110px 110px 110px", padding:"10px 16px", background:C.border+"44", color:C.muted, fontSize:11, textTransform:"uppercase", minWidth:520 }}>
            <span>משתמש</span><span>תוכנות</span><span>שנתי</span><span>קבוע</span><span>סה״כ</span>
          </div>
          {filtered.map(r=>(
            <div key={r.pcId} style={{ borderTop:`1px solid ${C.border}`, padding:"12px 16px", display:"grid", gridTemplateColumns:"1fr 1fr 110px 110px 110px", alignItems:"start" }}>
              <div>
                <div style={{ color:C.text, fontWeight:700, fontSize:13 }}>{r.userName||"—"}</div>
                <div style={{ color:C.dim, fontSize:11 }}>{r.pcId}</div>
                {r.userMail&&<div style={{ color:C.dim, fontSize:11 }}>✉ {r.userMail}</div>}
              </div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                {r.softwares.map(s=>(
                  <span key={s.id} style={{ background:swColor(s)+"22", color:swColor(s), borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:700 }}>{s.softwareName}</span>
                ))}
                {r.softwares.length===0&&<span style={{color:C.dim,fontSize:11}}>אין</span>}
              </div>
              <div style={{ color:r.annualCost>0?C.accent2:C.dim, fontWeight:700, fontSize:13 }}>{r.annualCost>0?`₪${r.annualCost.toLocaleString()}`:"—"}</div>
              <div style={{ color:r.permanentCost>0?C.green:C.dim, fontWeight:700, fontSize:13 }}>{r.permanentCost>0?`₪${r.permanentCost.toLocaleString()}`:"—"}</div>
              <div style={{ color:r.totalCost>0?C.accent:C.dim, fontWeight:800, fontSize:14 }}>{r.totalCost>0?`₪${r.totalCost.toLocaleString()}`:"—"}</div>
            </div>
          ))}
          {filtered.length===0&&<div style={{padding:40,textAlign:"center",color:C.muted}}>אין נתונים</div>}
        </div>
      )}
    </div>
  );
}

// ─── Admin Panel ──────────────────────────────────────────────────────────────
function AdminPanel({ session, onLogout, C }) {
  const [accounts,setAccounts]=useState([]);
  const [editId,setEditId]=useState(null);
  const [editForm,setEditForm]=useState({displayName:"",isAdmin:false});
  const [showNew,setShowNew]=useState(false);
  const [newForm,setNewForm]=useState({email:"",displayName:"",password:"",isAdmin:false});
  const [msg,setMsg]=useState(null);
  const [registering,setRegistering]=useState(false);

  const load=async()=>{ setAccounts(await getAccounts()); };
  useEffect(()=>{ load(); },[]);
  const showMsg=(text,type="success")=>{ setMsg({text,type}); setTimeout(()=>setMsg(null),4000); };
  const saveEdit=async()=>{ try{await updateUserProfile(editId,{displayName:editForm.displayName,isAdmin:editForm.isAdmin});showMsg("✓ עודכן!");setEditId(null);load();}catch(e){showMsg(e.message,"error");} };
  const handleRegister=async()=>{
    if(!newForm.email||!newForm.password){showMsg("אימייל וסיסמה נדרשים","error");return;}
    if(newForm.password.length<6){showMsg("סיסמה מינימום 6 תווים","error");return;}
    setRegistering(true);
    try{
      await registerNewUser(newForm.email,newForm.password,newForm.displayName,newForm.isAdmin);
      showMsg(`✓ ${newForm.email} נוצר! מתנתק...`);
      setNewForm({email:"",displayName:"",password:"",isAdmin:false});setShowNew(false);load();
      setTimeout(()=>{alert(`נוצר משתמש!\nFirebase ניתק אותך — התחבר מחדש עם:\n${session.email}`);onLogout();},1500);
    }catch(e){showMsg(e.message,"error");}
    setRegistering(false);
  };
  const RoleBtns=({value,onChange})=>(
    <div style={{display:"flex",gap:6}}>
      {[["false","👁 Viewer"],["true","👑 Admin"]].map(([v,label])=>(
        <button key={v} onClick={()=>onChange(v==="true")}
          style={{flex:1,padding:"8px",borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700,
            background:String(value)===v?C.accent+"33":C.bg,color:String(value)===v?C.accent:C.muted,
            border:`1px solid ${String(value)===v?C.accent:C.border}`}}>{label}</button>
      ))}
    </div>
  );
  return (
    <div style={{padding:"clamp(14px,2vw,28px)",maxWidth:800}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <h3 style={{color:C.text,margin:0,fontSize:"clamp(14px,1.8vw,18px)"}}>👥 ניהול משתמשים</h3>
        <Btn onClick={()=>{setShowNew(p=>!p);setEditId(null);}} variant={showNew?"neutral":"success"}>{showNew?"✕ ביטול":"➕ משתמש חדש"}</Btn>
      </div>
      {msg&&<div style={{color:msg.type==="error"?C.red:C.green,fontSize:13,marginBottom:14,padding:"10px 14px",background:(msg.type==="error"?C.red:C.green)+"18",borderRadius:8,border:`1px solid ${msg.type==="error"?C.red:C.green}44`}}>{msg.text}</div>}
      {showNew&&(
        <div style={{background:C.card,border:`1px solid ${C.green}55`,borderRadius:12,padding:"clamp(16px,2vw,24px)",marginBottom:20}}>
          <div style={{color:C.green,fontWeight:700,fontSize:14,marginBottom:14}}>➕ משתמש חדש</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10,marginBottom:12}}>
            <InputField label="אימייל *" value={newForm.email} onChange={v=>setNewForm(f=>({...f,email:v}))} type="email" C={C}/>
            <InputField label="שם תצוגה" value={newForm.displayName} onChange={v=>setNewForm(f=>({...f,displayName:v}))} C={C}/>
            <InputField label="סיסמה * (מינ׳ 6)" value={newForm.password} onChange={v=>setNewForm(f=>({...f,password:v}))} type="password" C={C}/>
          </div>
          <div style={{marginBottom:14}}><label style={{display:"block",color:C.muted,fontSize:11,textTransform:"uppercase",marginBottom:8}}>הרשאה</label><RoleBtns value={newForm.isAdmin} onChange={v=>setNewForm(f=>({...f,isAdmin:v}))}/></div>
          <div style={{background:C.yellow+"18",border:`1px solid ${C.yellow}44`,borderRadius:8,padding:"8px 12px",marginBottom:12,color:C.yellow,fontSize:12}}>⚠️ לאחר יצירה תתנתק אוטומטית</div>
          <Btn onClick={handleRegister} variant="success" disabled={registering}>{registering?"⏳ יוצר...":"✓ צור משתמש"}</Btn>
        </div>
      )}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto auto",gap:12,padding:"10px 16px",background:C.border+"44",color:C.muted,fontSize:11,textTransform:"uppercase"}}>
          <span>אימייל</span><span>שם</span><span>הרשאה</span><span></span>
        </div>
        {accounts.map(acc=>(
          <div key={acc.id}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto auto",gap:12,alignItems:"center",padding:"12px 16px",borderTop:`1px solid ${C.border}`}}>
              <span style={{color:C.text,fontSize:13,wordBreak:"break-all"}}>{acc.email}</span>
              <span style={{color:C.muted,fontSize:13}}>{acc.display_name||"—"}</span>
              <Tag color={acc.is_admin?C.accent:C.muted}>{acc.is_admin?"👑 Admin":"👁 Viewer"}</Tag>
              <Btn onClick={()=>{setEditId(editId===acc.id?null:acc.id);setShowNew(false);setEditForm({displayName:acc.display_name||"",isAdmin:acc.is_admin});}} variant="primary" style={{padding:"4px 10px",fontSize:11}}>✏️</Btn>
            </div>
            {editId===acc.id&&(
              <div style={{padding:"14px 16px",background:C.bg,borderTop:`1px solid ${C.border}`}}>
                <InputField label="שם תצוגה" value={editForm.displayName} onChange={v=>setEditForm(f=>({...f,displayName:v}))} C={C}/>
                <div style={{marginBottom:12}}><label style={{display:"block",color:C.muted,fontSize:11,textTransform:"uppercase",marginBottom:8}}>הרשאה</label><RoleBtns value={editForm.isAdmin} onChange={v=>setEditForm(f=>({...f,isAdmin:v}))}/></div>
                <div style={{display:"flex",gap:8}}><Btn onClick={saveEdit} variant="success">💾 שמור</Btn><Btn onClick={()=>setEditId(null)} variant="neutral">ביטול</Btn></div>
              </div>
            )}
          </div>
        ))}
        {accounts.length===0&&<div style={{padding:32,textAlign:"center",color:C.muted,fontSize:13}}>אין משתמשים</div>}
      </div>
    </div>
  );
}

// ─── Swap Modal ───────────────────────────────────────────────────────────────
function SwapModal({ pc, pcs, onSave, onCancel, C }) {
  const [newName,setNewName]=useState(pc.userName||""); const [saving,setSaving]=useState(false);
  const handle=async()=>{ setSaving(true); await swapPC(pc.pcId,newName); onSave(newName); setSaving(false); };
  return (
    <div style={{background:C.card,borderRadius:14,padding:28,minWidth:320}}>
      <h3 style={{color:C.text,margin:"0 0 6px",fontSize:17}}>🔄 העברת מחשב</h3>
      <div style={{color:C.muted,fontSize:12,marginBottom:16}}>PC <strong style={{color:C.accent}}>{pc.pcId}</strong></div>
      <div style={{background:C.bg,borderRadius:8,padding:"10px 14px",marginBottom:14,border:`1px solid ${C.border}`}}>
        <div style={{color:C.muted,fontSize:10,marginBottom:3,textTransform:"uppercase"}}>משתמש נוכחי</div>
        <div style={{color:C.text,fontWeight:700}}>{pc.userName||"—"}</div>
      </div>
      <InputField label="משתמש חדש" value={newName} onChange={setNewName} C={C}/>
      <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:16}}>
        {Object.values(pcs).filter(p=>p.pcId!==pc.pcId&&p.userName).map(p=>(
          <button key={p.pcId} onClick={()=>setNewName(p.userName)}
            style={{background:newName===p.userName?C.accent+"33":C.bg,border:`1px solid ${newName===p.userName?C.accent:C.border}`,borderRadius:6,padding:"4px 10px",color:newName===p.userName?C.accent:C.muted,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
            {p.userName}
          </button>
        ))}
      </div>
      <div style={{display:"flex",gap:8}}><Btn onClick={handle} variant="success" disabled={saving}>{saving?"...":"✓ אשר"}</Btn><Btn onClick={onCancel} variant="neutral">ביטול</Btn></div>
    </div>
  );
}

// ─── Table Groups ─────────────────────────────────────────────────────────────
const TABLE_COLORS=["#6aa3fc","#5de89a","#fc6b6b","#fde96a","#fbb360","#c084fc","#38bdf8","#f472b6","#a3e635","#fb923c"];

function TableGroupView({ desks, pcs, tableGroups, onSelectDesk, onGroupsChange, onAskConfirm, isAdmin, C }) {
  const [editingGroup,setEditingGroup]=useState(null);const [showAdd,setShowAdd]=useState(false);
  const [newName,setNewName]=useState("");const [newColor,setNewColor]=useState(TABLE_COLORS[0]);
  const [collapsed,setCollapsed]=useState(()=>Object.fromEntries((tableGroups||[]).map(g=>[g.name,true])));

  // sync collapsed state when tableGroups change (new groups start collapsed)
  useEffect(()=>{
    setCollapsed(prev=>{
      const next={...prev};
      tableGroups.forEach(g=>{ if(!(g.name in next)) next[g.name]=true; });
      return next;
    });
  },[tableGroups]);

  const toggleCollapse=(name)=>setCollapsed(c=>({...c,[name]:!c[name]}));
  const saveNew=async()=>{ if(!newName.trim())return; await onGroupsChange([...tableGroups,{name:newName.trim(),color:newColor}]); setNewName("");setNewColor(TABLE_COLORS[0]);setShowAdd(false); };
  const saveEdit=async()=>{ await onGroupsChange(tableGroups.map(g=>g.name===editingGroup.original?{name:editingGroup.name,color:editingGroup.color}:g)); setEditingGroup(null); };
  const delGroup=(name)=>{ onAskConfirm(`למחוק קבוצה "${name}"?`,`השולחנות בקבוצה יאבדו את הצבע.`,async()=>{ await onGroupsChange(tableGroups.filter(g=>g.name!==name)); }); };
  const CP=({selected,onSelect})=>(
    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>
      {TABLE_COLORS.map(c=><button key={c} onClick={()=>onSelect(c)} style={{width:24,height:24,borderRadius:"50%",background:c,border:selected===c?"3px solid white":"2px solid transparent",cursor:"pointer",transform:selected===c?"scale(1.2)":"scale(1)",transition:"all .15s"}}/>)}
      <input type="color" value={selected} onChange={e=>onSelect(e.target.value)} style={{width:28,height:28,border:"none",background:"none",cursor:"pointer"}}/>
    </div>
  );
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:10}}>
        <h3 style={{color:C.text,margin:0}}>🗂️ קבוצות שולחנות</h3>
        {isAdmin&&<Btn onClick={()=>setShowAdd(p=>!p)} variant="success" style={{padding:"7px 14px",fontSize:12}}>{showAdd?"✕":"+ קבוצה חדשה"}</Btn>}
      </div>
      {showAdd&&isAdmin&&(
        <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:18,marginBottom:20}}>
          <InputField label="שם קבוצה" value={newName} onChange={setNewName} C={C}/>
          <div style={{color:C.muted,fontSize:11,textTransform:"uppercase",marginBottom:4}}>צבע</div><CP selected={newColor} onSelect={setNewColor}/>
          <div style={{display:"flex",gap:8,marginTop:14}}><Btn onClick={saveNew} variant="success" disabled={!newName.trim()}>💾 צור</Btn><Btn onClick={()=>setShowAdd(false)} variant="neutral">ביטול</Btn></div>
        </div>
      )}
      {tableGroups.map(group=>{
        const gd=desks.filter(d=>d.tableGroup===group.name);const isEditing=editingGroup?.original===group.name;
        return (
          <div key={group.name} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`4px solid ${group.color}`,borderRadius:12,marginBottom:14,overflow:"hidden"}}>
            {isEditing?(
              <div style={{padding:"14px 18px",background:C.border+"22"}}>
                <input value={editingGroup.name} onChange={e=>setEditingGroup(g=>({...g,name:e.target.value}))} style={{width:"100%",background:C.inputBg,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 12px",color:C.text,fontSize:14,fontFamily:"inherit",outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
                <CP selected={editingGroup.color} onSelect={c=>setEditingGroup(g=>({...g,color:c}))}/>
                <div style={{display:"flex",gap:8,marginTop:12}}>
                  <Btn onClick={saveEdit} variant="success" style={{padding:"6px 12px",fontSize:11}}>✓</Btn>
                  <Btn onClick={()=>setEditingGroup(null)} variant="neutral" style={{padding:"6px 12px",fontSize:11}}>ביטול</Btn>
                  <Btn onClick={()=>delGroup(group.name)} variant="danger" style={{padding:"6px 12px",fontSize:11}}>🗑️</Btn>
                </div>
              </div>
            ):(
              <div onClick={()=>toggleCollapse(group.name)} style={{padding:"14px 18px",display:"flex",alignItems:"center",gap:10,cursor:"pointer"}}>
                <div style={{width:14,height:14,borderRadius:"50%",background:group.color,flexShrink:0}}/>
                <span style={{color:C.text,fontWeight:700,fontSize:15,flex:1}}>{group.name}</span>
                <span style={{color:C.muted,fontSize:12,marginLeft:4}}>{gd.length} מקומות</span>
                <span style={{color:C.muted,fontSize:13,marginLeft:4}}>{collapsed[group.name]?"▼":"▲"}</span>
                {isAdmin&&<Btn onClick={e=>{e.stopPropagation();setEditingGroup({original:group.name,name:group.name,color:group.color});}} variant="primary" style={{padding:"4px 10px",fontSize:11}}>✏️</Btn>}
              </div>
            )}
            {!collapsed[group.name]&&gd.length>0&&(
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:8,padding:"0 16px 16px"}}>
                {gd.map(desk=>{const pc=desk.pcId?pcs[desk.pcId]:null;return(
                  <div key={desk.id} onClick={()=>onSelectDesk(desk)} style={{background:C.bg,border:`1.5px solid ${pc?group.color+"88":C.border}`,borderRadius:8,padding:"10px 12px",cursor:"pointer"}}>
                    <div style={{color:group.color,fontSize:10,fontWeight:700,marginBottom:4}}>שולחן {desk.label}</div>
                    {pc?<><div style={{color:C.text,fontSize:13,fontWeight:700}}>{pc.userName}</div><div style={{color:C.muted,fontSize:11}}>{pc.pcId}</div></>:<div style={{color:C.dim,fontSize:12}}>ריק</div>}
                  </div>
                );})}
              </div>
            )}
            {!collapsed[group.name]&&gd.length===0&&(
              <div style={{padding:"10px 18px 14px",color:C.dim,fontSize:12}}>אין שולחנות בקבוצה זו</div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Floor Map ────────────────────────────────────────────────────────────────
function FloorMapFull({ desks, pcs, onDeskClick, onAddDesk, editMode, floorImage, C }) {
  const containerRef=useRef(null);const [dragging,setDragging]=useState(null);const [localDesks,setLocalDesks]=useState(desks);
  useEffect(()=>setLocalDesks(desks),[desks]);
  const handleMapClick=useCallback((e)=>{ if(!editMode||dragging)return; const rect=containerRef.current.getBoundingClientRect(); onAddDesk({x:((e.clientX-rect.left)/rect.width)*100,y:((e.clientY-rect.top)/rect.height)*100}); },[editMode,dragging,onAddDesk]);
  const handleMouseMove=useCallback((e)=>{ if(!dragging||!containerRef.current)return; const rect=containerRef.current.getBoundingClientRect(); setLocalDesks(ds=>ds.map(d=>d.id===dragging?{...d,x:Math.max(1,Math.min(99,((e.clientX-rect.left)/rect.width)*100)),y:Math.max(1,Math.min(99,((e.clientY-rect.top)/rect.height)*100))}:d)); },[dragging]);
  const handleMouseUp=useCallback(()=>{ if(dragging){const moved=localDesks.find(d=>d.id===dragging);if(moved)onDeskClick(moved,true);setDragging(null);} },[dragging,localDesks,onDeskClick]);
  return (
    <div ref={containerRef} onClick={handleMapClick} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp}
      style={{position:"relative",width:"100%",height:"100%",background:C.panel,borderRadius:10,overflow:"hidden",cursor:editMode?"crosshair":"default",userSelect:"none",border:`1px solid ${C.border}`}}>
      {floorImage?<img src={floorImage} alt="floor" style={{position:"absolute",inset:0,width:"100%",height:"100%",objectFit:"contain",opacity:.92}}/>:
        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",pointerEvents:"none"}}>
          <div style={{color:C.dim,textAlign:"center"}}><div style={{fontSize:52,marginBottom:12}}>🏢</div><div style={{fontSize:15,color:C.muted}}>העלה תמונת קומה</div></div>
        </div>}
      {localDesks.map(desk=>{const pc=desk.pcId?pcs[desk.pcId]:null;const color=desk.color||(pc?C.accent:C.dim);return(
        <div key={desk.id} onMouseDown={e=>{if(!editMode)return;e.stopPropagation();setDragging(desk.id);}} onClick={e=>{e.stopPropagation();if(!dragging)onDeskClick(desk,false);}}
          style={{position:"absolute",left:`${desk.x}%`,top:`${desk.y}%`,transform:"translate(-50%,-50%)",width:36,height:36,borderRadius:"50%",background:color+"33",border:`2.5px solid ${color}`,display:"flex",alignItems:"center",justifyContent:"center",cursor:editMode?"grab":"pointer",zIndex:dragging===desk.id?10:2,boxShadow:`0 0 10px ${color}66`}}>
          <span style={{color,fontSize:9,fontWeight:800,textAlign:"center"}}>{desk.label}</span>
        </div>
      );})}
      {editMode&&<div style={{position:"absolute",top:12,right:12,background:C.accent+"22",border:`1px solid ${C.accent}66`,borderRadius:8,padding:"6px 14px",color:C.accent,fontSize:12,fontWeight:700,pointerEvents:"none"}}>✏️ לחץ להוספה · גרור להזזה</div>}
    </div>
  );
}

function DeskForm({ desk, tableGroups, pcs, onSave, onCancel, onDelete, C }) {
  const [label,setLabel]=useState(desk?.label||"");const [groupName,setGroupName]=useState(desk?.tableGroup||(tableGroups[0]?.name||""));const [pcId,setPcId]=useState(desk?.pcId||"");
  const group=tableGroups.find(g=>g.name===groupName);const groupColor=group?.color||"#6aa3fc";
  return (
    <div style={{background:C.card,borderRadius:14,padding:24,minWidth:300}}>
      <h3 style={{color:C.text,margin:"0 0 16px",fontSize:15}}>{desk?"עריכת שולחן":"שולחן חדש"}</h3>
      <InputField label="תווית שולחן" value={label} onChange={setLabel} placeholder="A1, B3..." C={C}/>
      <div style={{marginBottom:14}}>
        <label style={{display:"block",color:C.muted,fontSize:11,textTransform:"uppercase",marginBottom:6}}>קבוצת שולחנות</label>
        {tableGroups.length===0?<div style={{color:C.red,fontSize:12,padding:"8px 12px",background:C.red+"18",borderRadius:7}}>⚠ אין קבוצות — צור קבוצה בטאב 🗂️ תחילה</div>:
          <select value={groupName} onChange={e=>setGroupName(e.target.value)} style={{width:"100%",background:C.inputBg,border:`1px solid ${C.border}`,borderRadius:7,padding:"9px 12px",color:C.text,fontSize:13,fontFamily:"inherit",outline:"none"}}>
            {tableGroups.map(g=><option key={g.name} value={g.name}>{g.name}</option>)}
          </select>}
      </div>
      <div style={{marginBottom:16}}>
        <label style={{display:"block",color:C.muted,fontSize:11,textTransform:"uppercase",marginBottom:4}}>שייך מחשב</label>
        <select value={pcId} onChange={e=>setPcId(e.target.value)} style={{width:"100%",background:C.inputBg,border:`1px solid ${C.border}`,borderRadius:7,padding:"9px 12px",color:C.text,fontSize:13,fontFamily:"inherit",outline:"none"}}>
          <option value="">— ריק —</option>
          {Object.values(pcs).map(pc=><option key={pc.pcId} value={pc.pcId}>{pc.userName} ({pc.pcId})</option>)}
        </select>
      </div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <Btn onClick={()=>onSave({label,tableGroup:groupName,pcId:pcId||null,color:groupColor})} variant="success" disabled={tableGroups.length===0}>💾 שמור</Btn>
        {desk&&<Btn onClick={onDelete} variant="danger">🗑️ הסר</Btn>}
        <Btn onClick={onCancel} variant="neutral">ביטול</Btn>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [isMobile,setIsMobile]=useState(()=>window.innerWidth<700);
  useEffect(()=>{ const h=()=>setIsMobile(window.innerWidth<700); window.addEventListener("resize",h); return()=>window.removeEventListener("resize",h); },[]);
  const [session,setSession]=useState(null);
  const [themeKey,setThemeKey]=useState("dark");
  const C=THEMES[themeKey]; const isAdmin=session?.isAdmin;

  const [pcs,setPcs]=useState({}); const [softwares,setSoftwares]=useState([]);
  const [desks,setDesks]=useState([]); const [tableGroups,setTableGroups]=useState([]);
  const [floorImage,setFloorImage]=useState(null); const [selectedDesk,setSelectedDesk]=useState(null);
  const [editMode,setEditMode]=useState(false); const [activeTab,setActiveTab]=useState("map");
  const [modal,setModal]=useState(null); const [toast,setToast]=useState(null); const [loading,setLoading]=useState(false); const [confirmDlg,setConfirmDlg]=useState(null);
  const [pcSearch,setPcSearch]=useState(""); const [pcSort,setPcSort]=useState("az");
  const askConfirm=(title,message,onConfirm)=>setConfirmDlg({title,message,onConfirm});

  const showToast=(msg,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),4000); };

  useEffect(()=>{
    const unsub=onAuthStateChanged(auth,async(user)=>{
      if(user){const profile=await getUserProfile(user.uid);setSession({uid:user.uid,email:user.email,displayName:profile?.displayName||user.email,isAdmin:profile?.isAdmin||false});}
      else setSession(null);
    }); return unsub;
  },[]);

  const loadAll=async()=>{
    setLoading(true);
    try{const [p,sw,d,f,tg]=await Promise.all([getPCs(),getSoftwares(),getDesks(),getFloorPlan(),getTableGroups()]);setPcs(p);setSoftwares(sw);setDesks(d);if(f)setFloorImage(f);setTableGroups(tg);}
    catch(e){showToast("שגיאה בטעינה: "+e.message,"error");}
    setLoading(false);
  };

  const handleLogin=(data)=>setSession(data);
  useEffect(()=>{ if(session)loadAll(); },[session]);
  const handleLogout=async()=>{ await logoutUser(); setSession(null); setPcs({}); setSoftwares([]); setDesks([]); setFloorImage(null); setSelectedDesk(null); };
  const handleFloorUpload=(e)=>{ const file=e.target.files[0]; if(!file)return; const reader=new FileReader(); reader.onload=async(ev)=>{ const img=ev.target.result; setFloorImage(img); await saveFloorPlan(img); showToast("✓ מפת קומה נשמרה"); }; reader.readAsDataURL(file); };
  const handleGroupsChange=async(updated)=>{ setTableGroups(updated); await saveTableGroups(updated); setDesks(ds=>ds.map(d=>{const g=updated.find(g=>g.name===d.tableGroup);return g?{...d,color:g.color}:d;})); };
  const handleDeskClick=(desk,isPos)=>{ if(isPos){setDesks(ds=>ds.map(d=>d.id===desk.id?{...d,x:desk.x,y:desk.y}:d));saveDesk(desk);return;} if(editMode)setModal({type:"editDesk",data:desk});else setSelectedDesk(desk); };
  const handleAddDesk=({x,y})=>{ const fg=tableGroups[0]; setModal({type:"editDesk",data:{id:`desk-${Date.now()}`,x,y,pcId:null,label:`D${desks.length+1}`,tableGroup:fg?.name||"",color:fg?.color||"#6aa3fc"},isNew:true}); };
  const saveDeskModal=async(updated)=>{ const g=tableGroups.find(g=>g.name===updated.tableGroup); const wc={...updated,color:g?.color||"#6aa3fc"}; const desk=modal.isNew?{...modal.data,...wc}:{...desks.find(d=>d.id===modal.data.id),...wc}; if(modal.isNew){setDesks(ds=>[...ds,desk]);await saveDesk(desk);}else{setDesks(ds=>ds.map(d=>d.id===desk.id?desk:d));await saveDesk(desk);} setModal(null); };
  const handleDeleteDesk=()=>{ askConfirm(`למחוק שולחן "${modal.data.label}"?`,"השולחן יוסר מהמפה.",async()=>{ setDesks(ds=>ds.filter(d=>d.id!==modal.data.id)); await deleteDesk(modal.data.id); setSelectedDesk(null); setModal(null); setConfirmDlg(null); }); };
  const handleSavePC=async(pc)=>{ await savePC(pc); setPcs(p=>({...p,[pc.pcId]:pc})); setModal(null); showToast(`✓ ${pc.pcId} נשמר`); };
  const handleDeletePC=(pcId,pcName)=>{ askConfirm(`למחוק את המחשב "${pcName||pcId}"?`,"כל הנתונים של המחשב יימחקו לצמיתות.",async()=>{ setPcs(p=>{const n={...p};delete n[pcId];return n;}); setDesks(ds=>ds.map(d=>d.pcId===pcId?{...d,pcId:null}:d)); await deletePC(pcId); setSelectedDesk(null); setConfirmDlg(null); }); };
  const handleSwap=async(pc,newUserName)=>{ setPcs(p=>({...p,[pc.pcId]:{...p[pc.pcId],userName:newUserName}})); setModal(null); showToast(`✓ הועבר ל-${newUserName}`); };
  const handleSaveSoftware=async(sw)=>{ await saveSoftware(sw); setSoftwares(await getSoftwares()); showToast("✓ תוכנה נשמרה"); };
  const handleDeleteSoftware=async(id)=>{ await deleteSoftware(id); setSoftwares(await getSoftwares()); setPcs(await getPCs()); showToast("✓ תוכנה נמחקה"); };

  const selectedPC=selectedDesk?.pcId?pcs[selectedDesk.pcId]:null;
  const TABS=[
    ...(!isMobile?[{id:"map",label:"🗺️ מפה"}]:[]),
    {id:"tables",label:"🗂️ שולחנות"},{id:"list",label:"📋 מחשבים"},
    {id:"software",label:"💿 תוכנות"},{id:"report",label:"📊 דוח"},
    ...(isAdmin?[{id:"admin",label:"👥 משתמשים"}]:[]),
  ];
  // אם על נייד ונמצאים בטאב מפה — עבור לשולחנות
  useEffect(()=>{ if(isMobile&&activeTab==="map") setActiveTab("tables"); },[isMobile]);

  if(!session) return <LoginScreen onLogin={handleLogin} C={C}/>;

  return (
    <div dir="rtl" style={{ background:C.bg, height:"100vh", width:"100vw", fontFamily:"'DM Mono','Courier New',monospace", color:C.text, display:"flex", flexDirection:"column", overflow:"hidden" }}>
      {toast&&<div style={{position:"fixed",top:16,left:16,zIndex:999,background:toast.type==="error"?C.red+"22":C.green+"22",border:`1px solid ${toast.type==="error"?C.red:C.green}88`,borderRadius:10,padding:"12px 20px",color:toast.type==="error"?C.red:C.green,fontSize:13,fontWeight:700,boxShadow:C.shadow,maxWidth:360}}>{toast.msg}</div>}

      {/* Header */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:isMobile?"0 10px":"0 clamp(12px,2vw,24px)",display:"flex",alignItems:"center",justifyContent:"space-between",height:isMobile?"52px":"clamp(48px,6vh,60px)",flexShrink:0,boxShadow:C.shadow,gap:6}}>
        {/* Logo + stats */}
        <div style={{display:"flex",alignItems:"center",gap:isMobile?6:"clamp(8px,1.5vw,18px)",minWidth:0,overflow:"hidden"}}>
          <div style={{background:C.accent+"22",border:`1px solid ${C.accent}55`,borderRadius:7,padding:isMobile?"4px 8px":"5px clamp(8px,1.5vw,14px)",color:C.accent,fontWeight:800,fontSize:isMobile?12:"clamp(11px,1.4vw,14px)",letterSpacing:1,whiteSpace:"nowrap",flexShrink:0}}>⚙{!isMobile&&" OFFICE MGR"}</div>
          {!isMobile&&<><span style={{color:C.text,fontSize:"clamp(11px,1.2vw,13px)"}}><b style={{color:C.accent}}>{Object.keys(pcs).length}</b> מחשבים</span><span style={{color:C.dim}}>·</span><span style={{color:C.text,fontSize:"clamp(11px,1.2vw,13px)"}}><b style={{color:C.green}}>{softwares.length}</b> תוכנות</span></>}
          {loading&&<span style={{color:C.muted,fontSize:12}}>⏳</span>}
        </div>
        {/* Actions */}
        <div style={{display:"flex",gap:isMobile?4:"clamp(4px,.8vw,10px)",alignItems:"center",flexShrink:0}}>
          <button onClick={()=>setThemeKey(k=>k==="dark"?"light":"dark")} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:7,padding:isMobile?"6px 8px":"6px 10px",color:C.muted,fontSize:isMobile?14:15,cursor:"pointer",flexShrink:0}}>{themeKey==="dark"?"☀️":"🌙"}</button>
          {isAdmin&&<>
            {!isMobile&&<label style={{background:C.dim+"22",border:`1px solid ${C.border}`,borderRadius:7,padding:"6px clamp(8px,1vw,14px)",color:C.muted,fontSize:"clamp(11px,1.1vw,13px)",cursor:"pointer",fontWeight:700,whiteSpace:"nowrap"}}>
              🖼️ מפה<input type="file" accept="image/*" onChange={handleFloorUpload} style={{display:"none"}}/>
            </label>}
            <button onClick={()=>setModal({type:"addPC"})} style={{background:C.green+"22",border:`1px solid ${C.green}55`,borderRadius:7,padding:isMobile?"6px 10px":"6px clamp(8px,1vw,14px)",color:C.green,fontSize:isMobile?12:"clamp(11px,1.1vw,13px)",cursor:"pointer",fontFamily:"inherit",fontWeight:800,whiteSpace:"nowrap"}}>+ {!isMobile&&"מחשב"}{isMobile&&"💻"}</button>
          </>}
          <div style={{display:"flex",alignItems:"center",gap:isMobile?4:8,borderRight:`1px solid ${C.border}`,paddingRight:isMobile?6:"clamp(6px,1vw,12px)"}}>
            {!isMobile&&<div><div style={{color:C.text,fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>{session.displayName||session.email}</div><div style={{color:isAdmin?C.accent:C.muted,fontSize:10}}>{isAdmin?"Admin":"Viewer"}</div></div>}
            <button onClick={handleLogout} style={{background:C.red+"18",border:`1px solid ${C.red}44`,borderRadius:6,padding:isMobile?"6px 8px":"5px 10px",color:C.red,fontSize:isMobile?12:"clamp(10px,1vw,12px)",cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>{isMobile?"🚪":"יציאה"}</button>
          </div>
        </div>
      </div>

      {/* Tabs — desktop top bar, mobile bottom nav */}
      {!isMobile && (
        <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 clamp(10px,1.5vw,24px)",display:"flex",gap:2,flexShrink:0,overflowX:"auto"}}>
          {TABS.map(tab=>(
            <button key={tab.id} onClick={()=>setActiveTab(tab.id)}
              style={{background:"none",border:"none",borderBottom:activeTab===tab.id?`2.5px solid ${C.accent}`:"2.5px solid transparent",color:activeTab===tab.id?C.accent:C.muted,padding:"clamp(7px,1.2vh,11px) clamp(10px,1.5vw,20px)",fontSize:"clamp(11px,1.2vw,13px)",fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",marginBottom:-1}}>
              {tab.label}
            </button>
          ))}
        </div>
      )}

      {/* Body */}
      <div style={{flex:1,overflow:"hidden",position:"relative"}}>
        <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column"}}>
          {activeTab==="map"&&(
            <div style={{flex:1,padding:"clamp(10px,1.5vh,16px) clamp(12px,1.5vw,22px)",display:"flex",flexDirection:"column",overflow:"hidden"}}>
              <div style={{display:"flex",gap:10,marginBottom:10,alignItems:"center",flexShrink:0,flexWrap:"wrap"}}>
                <div style={{display:"flex",gap:"clamp(8px,1.5vw,16px)",alignItems:"center",flexWrap:"wrap"}}>
                  {tableGroups.map(g=><div key={g.name} style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:11,height:11,borderRadius:"50%",background:g.color,flexShrink:0}}/><span style={{color:C.muted,fontSize:12,whiteSpace:"nowrap"}}>{g.name}</span></div>)}
                </div>
                <div style={{flex:1}}/>
                {isAdmin&&<button onClick={()=>setEditMode(m=>!m)} style={{background:editMode?C.red+"22":C.dim+"22",border:`1px solid ${editMode?C.red+"88":C.border}`,borderRadius:7,padding:"6px 14px",color:editMode?C.red:C.muted,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{editMode?"✓ סיום עריכה":"✏️ עריכת מפה"}</button>}
              </div>
              <div style={{flex:1,minHeight:0}}><FloorMapFull desks={desks} pcs={pcs} onDeskClick={handleDeskClick} onAddDesk={handleAddDesk} editMode={editMode&&isAdmin} floorImage={floorImage} C={C}/></div>
            </div>
          )}
          {activeTab==="tables"&&<div style={{flex:1,overflowY:"auto",padding:"clamp(14px,2vh,22px) clamp(14px,2vw,28px)"}}><TableGroupView desks={desks} pcs={pcs} tableGroups={tableGroups} onSelectDesk={d=>{setSelectedDesk(d);setActiveTab("map");}} onGroupsChange={handleGroupsChange} onAskConfirm={askConfirm} isAdmin={isAdmin} C={C}/></div>}
          {activeTab==="list"&&(
            <div style={{flex:1,overflowY:"auto",padding:`clamp(14px,2vh,22px) clamp(14px,2vw,28px) ${isMobile?"80px":"22px"}`}}>
              {/* Search + Sort */}
              <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
                <input value={pcSearch} onChange={e=>setPcSearch(e.target.value)} placeholder="🔍 חיפוש לפי שם או מספר מחשב..."
                  style={{flex:1,minWidth:160,background:C.inputBg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 14px",color:C.text,fontSize:13,fontFamily:"inherit",outline:"none"}}/>
                <button onClick={()=>setPcSort(s=>s==="az"?"za":"az")}
                  style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 14px",color:C.muted,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
                  {pcSort==="az"?"א→ת ↑":"ת→א ↓"}
                </button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(auto-fill,minmax(280px,1fr))",gap:isMobile?10:14}}>
                {Object.values(pcs)
                  .filter(pc=>!pcSearch||pc.userName?.toLowerCase().includes(pcSearch.toLowerCase())||pc.pcId?.toLowerCase().includes(pcSearch.toLowerCase()))
                  .sort((a,b)=>{ const na=a.userName||a.pcId||""; const nb=b.userName||b.pcId||""; return pcSort==="az"?na.localeCompare(nb,"he"):nb.localeCompare(na,"he"); })
                  .map(pc=>(
                  <div key={pc.pcId} onClick={()=>setSelectedDesk({pcId:pc.pcId,label:"—",tableGroup:"—"})}
                    style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:10,padding:"clamp(12px,1.5vw,18px)",cursor:"pointer",transition:"all .15s"}}
                    onMouseEnter={e=>e.currentTarget.style.borderColor=C.accent+"88"} onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                      <div><div style={{color:C.text,fontWeight:800,fontSize:15}}>{pc.userName||<span style={{color:C.dim}}>לא מוקצה</span>}</div><div style={{color:C.muted,fontSize:12,fontFamily:"monospace"}}>{pc.pcId}</div></div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
                      <div style={{color:C.muted,fontSize:12}}>🖥 {pc.cpuModel||"—"}</div>
                      <div style={{color:C.muted,fontSize:12}}>🎮 {pc.gpuModel||"—"}</div>
                      <div style={{color:C.muted,fontSize:12}}>💾 {pc.ramGb?`${pc.ramGb}GB`:"—"}</div>
                      <div style={{color:C.muted,fontSize:12}}>💿 {pc.ssdGb?`${pc.ssdGb}GB`:"—"}</div>
                    </div>
                    {pc.softwareIds?.length>0&&<div style={{color:C.dim,fontSize:11,marginTop:6}}>{pc.softwareIds.length} תוכנות</div>}
                  </div>
                ))}
              </div>
            </div>
          )}
          {activeTab==="software"&&<div style={{flex:1,overflowY:"auto"}}><SoftwareManager softwares={softwares} pcs={pcs} onSave={handleSaveSoftware} onDelete={handleDeleteSoftware} isAdmin={isAdmin} C={C}/></div>}
          {activeTab==="report"&&<div style={{flex:1,overflowY:"auto"}}><CostReport pcs={pcs} softwares={softwares} isMobile={isMobile} C={C}/></div>}
          {activeTab==="admin"&&isAdmin&&<div style={{flex:1,overflowY:"auto"}}><AdminPanel session={session} onLogout={handleLogout} C={C}/></div>}
        </div>

        {/* Sidebar — desktop: panel lateral, mobile: full overlay */}
        <div style={{
          position:"absolute", top:0, left:0, bottom:0, zIndex:20,
          ...(isMobile
            ? { right:0, width:selectedDesk?"100%":"0%", background:C.panel, overflowY:"auto", overflowX:"hidden", transition:"width .22s ease" }
            : { width:selectedDesk?"clamp(300px,28vw,420px)":0, borderRight:selectedDesk?`1px solid ${C.border}`:"none", background:C.panel, overflowY:"auto", overflowX:"hidden", transition:"width .22s cubic-bezier(.4,0,.2,1)", boxShadow:selectedDesk?"4px 0 20px #0004":"none" }
          )
        }}>
          {selectedDesk&&(
            <div style={{width:isMobile?"100%":"clamp(300px,28vw,420px)",padding:"clamp(12px,1.5vw,20px)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div><div style={{color:C.muted,fontSize:10,textTransform:"uppercase",letterSpacing:1}}>שולחן {selectedDesk.label}</div><div style={{color:C.dim,fontSize:11}}>{selectedDesk.tableGroup}</div></div>
                <button onClick={()=>setSelectedDesk(null)} style={{background:C.dim+"33",border:"none",color:C.muted,cursor:"pointer",fontSize:18,width:30,height:30,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
              </div>
              {selectedPC?(
                <PCDetail pc={selectedPC} softwares={softwares} isAdmin={isAdmin} C={C}
                  onEdit={()=>setModal({type:"editPC",data:selectedPC})}
                  onDelete={()=>handleDeletePC(selectedPC.pcId,selectedPC.userName)}
                  onSwap={()=>setModal({type:"swap",data:selectedPC})}/>
              ):(
                <div style={{background:C.card,borderRadius:12,padding:"clamp(20px,3vh,36px)",textAlign:"center",border:`1px solid ${C.border}`}}>
                  <div style={{fontSize:44,marginBottom:12}}>🖥️</div>
                  <div style={{color:C.muted,fontSize:14,marginBottom:20}}>שולחן ריק</div>
                  {isAdmin&&<button onClick={()=>setModal({type:"editDesk",data:selectedDesk})} style={{background:C.accent+"22",border:`1px solid ${C.accent}55`,borderRadius:8,padding:"10px 24px",color:C.accent,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>שייך מחשב</button>}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Mobile bottom nav */}
      {isMobile && !selectedDesk && (
        <div style={{position:"fixed",bottom:0,left:0,right:0,background:C.panel,borderTop:`1px solid ${C.border}`,display:"flex",zIndex:50,boxShadow:"0 -2px 12px #0005"}}>
          {TABS.map(tab=>(
            <button key={tab.id} onClick={()=>setActiveTab(tab.id)}
              style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"10px 4px 8px",background:"none",border:"none",color:activeTab===tab.id?C.accent:C.muted,cursor:"pointer",fontFamily:"inherit",fontSize:9,fontWeight:700,borderTop:activeTab===tab.id?`2px solid ${C.accent}`:"2px solid transparent"}}>
              <span style={{fontSize:18}}>{tab.label.split(" ")[0]}</span>
              <span style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:60}}>{tab.label.split(" ").slice(1).join(" ")}</span>
            </button>
          ))}
        </div>
      )}
      {/* Modals */}
      {modal&&(
        <Modal onClose={()=>setModal(null)} C={C}>
          {(modal.type==="addPC"||modal.type==="editPC")&&<PCForm initial={modal.data} softwares={softwares} onSave={handleSavePC} onCancel={()=>setModal(null)} C={C}/>}
          {modal.type==="editDesk"&&<DeskForm desk={modal.isNew?null:modal.data} tableGroups={tableGroups} pcs={pcs} onSave={saveDeskModal} onCancel={()=>setModal(null)} onDelete={handleDeleteDesk} C={C}/>}
          {modal.type==="swap"&&<SwapModal pc={modal.data} pcs={pcs} onSave={name=>handleSwap(modal.data,name)} onCancel={()=>setModal(null)} C={C}/>}
        </Modal>
      )}
      {/* Confirm Delete Dialog */}
      {confirmDlg&&<ConfirmDialog title={confirmDlg.title} message={confirmDlg.message} onConfirm={confirmDlg.onConfirm} onCancel={()=>setConfirmDlg(null)} C={C}/>}
    </div>
  );
}
