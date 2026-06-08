import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  getPCs, savePC, deletePC, swapPC, bulkSavePCs,
  getDesks, saveDesk, deleteDesk,
  getTableGroups, saveTableGroups,
  getFloorPlan, saveFloorPlan,
  getAccounts, getAccountByEmail, createAccount, updateAccount, deleteAccount, updateAccountPassword,
} from "./db.js";

// ─── Auth helpers (browser-side SHA-256, same as original) ────────────────────
async function sha256(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}
export async function hashPassword(password) {
  const salt = Array.from(crypto.getRandomValues(new Uint8Array(16))).map(b => b.toString(16).padStart(2, "0")).join("");
  return salt + ":" + await sha256(salt + password);
}
export async function verifyPassword(password, stored) {
  if (!stored) return false;
  const [salt, hash] = stored.split(":");
  return await sha256(salt + password) === hash;
}

// ─── Session store (in-memory, localStorage for persistence) ─────────────────
function loadSession() {
  try { return JSON.parse(localStorage.getItem("om_session") || "null"); } catch { return null; }
}
function persistSession(s) {
  if (s) localStorage.setItem("om_session", JSON.stringify(s));
  else localStorage.removeItem("om_session");
}

// ─── Excel import (unchanged from original) ───────────────────────────────────
async function importFromExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const pcSheet = wb.Sheets["PCs"] || wb.Sheets[wb.SheetNames[0]];
        let pcRows = XLSX.utils.sheet_to_json(pcSheet, { defval: "" });
        if (pcRows.length > 0 && pcRows[0]["pc_id"] === undefined)
          pcRows = XLSX.utils.sheet_to_json(pcSheet, { defval: "", range: 1 });
        pcRows = pcRows.filter(r => r.pc_id && r.pc_id !== "pc_id");

        const userSheet = wb.Sheets["Users"] || wb.Sheets[wb.SheetNames[1]];
        let userRows = userSheet ? XLSX.utils.sheet_to_json(userSheet, { defval: "" }) : [];
        if (userRows.length > 0 && userRows[0]["pc_id"] === undefined)
          userRows = XLSX.utils.sheet_to_json(userSheet, { defval: "", range: 1 });
        userRows = userRows.filter(r => r.pc_id && r.pc_id !== "pc_id");

        const userMap = {};
        for (const u of userRows) {
          const key = String(u.pc_id || u.user_name || "").trim();
          if (!key) continue;
          const hash = u.outlook_password && !String(u.outlook_password).includes("[HASHED")
            ? await hashPassword(String(u.outlook_password)) : null;
          userMap[key] = { mail: u.mail || "", passwordHash: hash };
        }

        const pcs = {};
        for (const row of pcRows) {
          if (!row.pc_id) continue;
          const id = String(row.pc_id).trim();
          const userEntry = userMap[id] || userMap[String(row.user_name || "").trim()] || { mail: "", passwordHash: null };
          pcs[id] = {
            pcId: id, userName: row.user_name || "", autocadVersion: String(row.autocad_version || ""),
            licenseType: row.license_type || "", serialNumber: row.serial_number || "",
            productKey: row.product_key || "", operatingSystem: row.operating_system || "",
            cpu: { manufacturer: row.cpu_manufacturer || "", model: row.cpu_model || "", cores: row.cpu_cores || "", ghz: row.cpu_ghz || "" },
            gpu: { manufacturer: row.gpu_manufacturer || "", model: row.gpu_model || "", vramGb: row.gpu_vram_gb || "" },
            ssd: { brand: row.ssd_brand || "", capacityGb: row.ssd_capacity_gb || "", type: row.ssd_type || "" },
            ram: { sizeGb: row.ram_size_gb || "", speedMhz: row.ram_speed_mhz || "", sticks: row.ram_sticks || "" },
            teken: Number(row.teken) || 0, comments: row.comments || "", user: userEntry,
          };
        }
        resolve(pcs);
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function exportToExcel(pcs, desks) {
  const wb = XLSX.utils.book_new();
  const pcRows = Object.values(pcs).map(pc => ({
    pc_id: pc.pcId, user_name: pc.userName, autocad_version: pc.autocadVersion,
    license_type: pc.licenseType, serial_number: pc.serialNumber, product_key: pc.productKey,
    operating_system: pc.operatingSystem, cpu_manufacturer: pc.cpu?.manufacturer, cpu_model: pc.cpu?.model,
    cpu_cores: pc.cpu?.cores, cpu_ghz: pc.cpu?.ghz, gpu_manufacturer: pc.gpu?.manufacturer,
    gpu_model: pc.gpu?.model, gpu_vram_gb: pc.gpu?.vramGb, ssd_brand: pc.ssd?.brand,
    ssd_capacity_gb: pc.ssd?.capacityGb, ssd_type: pc.ssd?.type, ram_size_gb: pc.ram?.sizeGb,
    ram_speed_mhz: pc.ram?.speedMhz, ram_sticks: pc.ram?.sticks, teken: pc.teken, comments: pc.comments,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pcRows), "PCs");
  const userRows = Object.values(pcs).map(pc => ({
    pc_id: pc.pcId, user_name: pc.userName, mail: pc.user?.mail || "",
    outlook_password: pc.user?.passwordHash ? "[HASHED - do not edit]" : "",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(userRows), "Users");
  const deskRows = desks.map(d => ({ desk_id: d.id, label: d.label, table_group: d.tableGroup, color: d.color || "", pc_id: d.pcId || "", x_pct: d.x?.toFixed(2), y_pct: d.y?.toFixed(2) }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deskRows), "Desks");
  XLSX.writeFile(wb, `office_manager_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ─── Themes ───────────────────────────────────────────────────────────────────
const THEMES = {
  dark: {
    bg: "#181b27", panel: "#20243a", card: "#272c45", border: "#363d60",
    accent: "#6aa3fc", accent2: "#fbb360", green: "#5de89a", red: "#fc6b6b",
    yellow: "#fde96a", text: "#e8ecfa", muted: "#8892b8", dim: "#4a527a",
    inputBg: "#20243a", shadow: "0 4px 24px #0007",
  },
  light: {
    bg: "#f0f3fa", panel: "#ffffff", card: "#f6f8ff", border: "#d0d8f0",
    accent: "#2563eb", accent2: "#d97706", green: "#16a34a", red: "#dc2626",
    yellow: "#ca8a04", text: "#1a2040", muted: "#4a5280", dim: "#9aa0c0",
    inputBg: "#ffffff", shadow: "0 2px 12px #0002",
  },
};

// ─── Tiny components ──────────────────────────────────────────────────────────
const Tag = ({ color, children }) => (
  <span style={{ background: color + "22", color, border: `1px solid ${color}55`, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700 }}>{children}</span>
);

const Field = ({ label, value, mono, C }) => (
  <div style={{ marginBottom: 10 }}>
    <div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase", letterSpacing: 1, marginBottom: 2 }}>{label}</div>
    <div style={{ color: C.text, fontSize: 13, fontFamily: mono ? "monospace" : "inherit", background: mono ? C.bg : "transparent", padding: mono ? "3px 7px" : 0, borderRadius: 4 }}>
      {value || <span style={{ color: C.dim }}>—</span>}
    </div>
  </div>
);

const InputField = ({ label, value, onChange, type = "text", placeholder, C, error }) => (
  <div style={{ marginBottom: 12 }}>
    {label && <label style={{ display: "block", color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 4 }}>{label}</label>}
    <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
      style={{ width: "100%", background: C.inputBg, border: `1px solid ${error ? C.red : C.border}`, borderRadius: 7, padding: "9px 13px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none", boxSizing: "border-box" }}
      onFocus={e => e.target.style.borderColor = C.accent}
      onBlur={e => e.target.style.borderColor = error ? C.red : C.border} />
    {error && <div style={{ color: C.red, fontSize: 11, marginTop: 3 }}>⚠ {error}</div>}
  </div>
);

const Btn = ({ children, onClick, variant = "primary", style: s = {}, disabled }) => {
  const colors = { primary: "#6aa3fc", danger: "#fc6b6b", success: "#5de89a", neutral: "#8892b8", warning: "#fde96a" };
  const c = colors[variant] || colors.primary;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ background: c + "22", color: c, border: `1px solid ${c}55`, borderRadius: 7, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", fontFamily: "inherit", opacity: disabled ? 0.5 : 1, transition: "all 0.15s", ...s }}
      onMouseEnter={e => !disabled && (e.currentTarget.style.background = c + "44")}
      onMouseLeave={e => !disabled && (e.currentTarget.style.background = c + "22")}>
      {children}
    </button>
  );
};

// ─── Login Screen ─────────────────────────────────────────────────────────────
function LoginScreen({ onLogin, C }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    if (!email || !password) { setError("Please enter email and password"); return; }
    setLoading(true); setError("");
    try {
      const account = await getAccountByEmail(email);
      if (!account) { setError("Invalid email or password"); setLoading(false); return; }

      // First-time ADMIN_SEED — set password now
      if (account.passwordHash === "ADMIN_SEED") {
        const hash = await hashPassword(password);
        await updateAccountPassword(account.id, hash);
        onLogin({ email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
        setLoading(false); return;
      }

      const ok = await verifyPassword(password, account.passwordHash);
      if (!ok) { setError("Invalid email or password"); setLoading(false); return; }
      onLogin({ email: account.email, isAdmin: account.isAdmin, displayName: account.displayName });
    } catch (e) { setError("Error: " + e.message); }
    setLoading(false);
  };

  return (
    <div style={{ background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Mono','Courier New',monospace", padding: 20 }}>
      <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 18, padding: "clamp(28px,5vw,56px)", width: "min(420px, 100%)", boxShadow: C.shadow }}>
        <div style={{ textAlign: "center", marginBottom: "clamp(20px,4vh,36px)" }}>
          <div style={{ fontSize: "clamp(28px,4vw,44px)", marginBottom: 10 }}>⚙</div>
          <div style={{ color: C.accent, fontWeight: 800, fontSize: "clamp(16px,2.5vw,22px)", letterSpacing: 2 }}>OFFICE MANAGER</div>
          <div style={{ color: C.muted, fontSize: "clamp(11px,1.5vw,13px)", marginTop: 6 }}>Sign in to continue</div>
        </div>
        <InputField label="Email" value={email} onChange={setEmail} type="email" placeholder="your@email.com" C={C} />
        <InputField label="Password" value={password} onChange={setPassword} type="password" placeholder="••••••••" C={C} />
        {error && <div style={{ color: C.red, fontSize: 12, marginBottom: 14, padding: "10px 14px", background: C.red + "18", borderRadius: 8, border: `1px solid ${C.red}44` }}>✗ {error}</div>}
        <button onClick={handleLogin} onKeyDown={e => e.key === "Enter" && handleLogin()} disabled={loading}
          style={{ width: "100%", background: loading ? C.accent + "88" : C.accent, color: "#fff", border: "none", borderRadius: 9, padding: "clamp(10px,2vh,14px)", fontSize: "clamp(13px,1.8vw,15px)", fontWeight: 800, cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit", marginTop: 4 }}>
          {loading ? "Signing in..." : "Sign In →"}
        </button>
      </div>
    </div>
  );
}

// ─── Admin Panel ──────────────────────────────────────────────────────────────
function AdminPanel({ C }) {
  const [accounts, setAccounts] = useState([]);
  const [form, setForm] = useState({ email: "", displayName: "", password: "", isAdmin: false });
  const [editId, setEditId] = useState(null);
  const [msg, setMsg] = useState(null);

  const load = async () => { setAccounts(await getAccounts()); };
  useEffect(() => { load(); }, []);

  const save = async () => {
    try {
      if (editId) {
        const update = { displayName: form.displayName, isAdmin: form.isAdmin };
        if (form.password) update.passwordHash = await hashPassword(form.password);
        await updateAccount(editId, update);
      } else {
        if (!form.email || !form.password) { setMsg({ text: "Email and password required", type: "error" }); return; }
        await createAccount({ email: form.email, displayName: form.displayName, passwordHash: await hashPassword(form.password), isAdmin: form.isAdmin });
      }
      setMsg({ text: editId ? "Updated!" : "Account created!", type: "success" });
      setForm({ email: "", displayName: "", password: "", isAdmin: false }); setEditId(null);
      load();
    } catch (e) { setMsg({ text: e.message, type: "error" }); }
    setTimeout(() => setMsg(null), 3000);
  };

  const del = async (id, email) => {
    if (!confirm(`Delete account ${email}?`)) return;
    await deleteAccount(id); load();
  };

  return (
    <div style={{ padding: "clamp(14px,2vw,28px)" }}>
      <h3 style={{ color: C.text, margin: "0 0 20px", fontSize: "clamp(14px,1.8vw,18px)" }}>👥 User Accounts</h3>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", marginBottom: 24 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto auto", gap: 12, padding: "10px 16px", background: C.border + "44", color: C.muted, fontSize: 11, textTransform: "uppercase" }}>
          <span>Email</span><span>Name</span><span>Role</span><span></span>
        </div>
        {accounts.map(acc => (
          <div key={acc.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto auto", gap: 12, alignItems: "center", padding: "12px 16px", borderTop: `1px solid ${C.border}` }}>
            <span style={{ color: C.text, fontSize: 13 }}>{acc.email}</span>
            <span style={{ color: C.muted, fontSize: 13 }}>{acc.display_name}</span>
            <Tag color={acc.is_admin ? C.accent : C.muted}>{acc.is_admin ? "Admin" : "Viewer"}</Tag>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn onClick={() => { setEditId(acc.id); setForm({ email: acc.email, displayName: acc.display_name, password: "", isAdmin: acc.is_admin }); }} variant="primary" style={{ padding: "4px 10px", fontSize: 11 }}>Edit</Btn>
              <Btn onClick={() => del(acc.id, acc.email)} variant="danger" style={{ padding: "4px 10px", fontSize: 11 }}>✕</Btn>
            </div>
          </div>
        ))}
      </div>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "clamp(14px,2vw,22px)" }}>
        <h4 style={{ color: C.accent, margin: "0 0 16px", fontSize: 14 }}>{editId ? "✏️ Edit Account" : "➕ New Account"}</h4>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 12 }}>
          <InputField label="Email" value={form.email} onChange={v => setForm(f => ({ ...f, email: v }))} type="email" C={C} />
          <InputField label="Display Name" value={form.displayName} onChange={v => setForm(f => ({ ...f, displayName: v }))} C={C} />
          <InputField label={editId ? "New Password (blank = keep)" : "Password"} value={form.password} onChange={v => setForm(f => ({ ...f, password: v }))} type="password" C={C} />
          <div style={{ marginBottom: 12 }}>
            <label style={{ display: "block", color: C.muted, fontSize: 11, textTransform: "uppercase", marginBottom: 8 }}>Role</label>
            <div style={{ display: "flex", gap: 8 }}>
              {["Viewer", "Admin"].map(role => (
                <button key={role} onClick={() => setForm(f => ({ ...f, isAdmin: role === "Admin" }))}
                  style={{ flex: 1, padding: "9px", borderRadius: 7, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, background: (role === "Admin") === form.isAdmin ? C.accent + "33" : C.bg, color: (role === "Admin") === form.isAdmin ? C.accent : C.muted, border: `1px solid ${(role === "Admin") === form.isAdmin ? C.accent : C.border}` }}>
                  {role}
                </button>
              ))}
            </div>
          </div>
        </div>
        {msg && <div style={{ color: msg.type === "error" ? C.red : C.green, fontSize: 12, marginBottom: 12, padding: "8px 12px", background: (msg.type === "error" ? C.red : C.green) + "18", borderRadius: 6 }}>{msg.text}</div>}
        <div style={{ display: "flex", gap: 8 }}>
          <Btn onClick={save} variant="success">{editId ? "💾 Update" : "➕ Create"}</Btn>
          {editId && <Btn onClick={() => { setEditId(null); setForm({ email: "", displayName: "", password: "", isAdmin: false }); }} variant="neutral">Cancel</Btn>}
        </div>
      </div>
    </div>
  );
}

// ─── PC Swap Modal ────────────────────────────────────────────────────────────
function SwapModal({ pc, pcs, onSave, onCancel, C }) {
  const [newName, setNewName] = useState(pc.userName || "");
  const [saving, setSaving] = useState(false);
  const handle = async () => {
    setSaving(true);
    await swapPC(pc.pcId, newName);
    onSave(newName);
    setSaving(false);
  };
  return (
    <div style={{ background: C.card, borderRadius: 14, padding: 28, minWidth: 340 }}>
      <h3 style={{ color: C.text, margin: "0 0 6px", fontSize: 17 }}>🔄 Reassign PC</h3>
      <div style={{ color: C.muted, fontSize: 12, marginBottom: 20 }}>PC <strong style={{ color: C.accent }}>{pc.pcId}</strong> — email & password stay with the PC</div>
      <div style={{ background: C.bg, borderRadius: 8, padding: "10px 14px", marginBottom: 16, border: `1px solid ${C.border}` }}>
        <div style={{ color: C.muted, fontSize: 10, marginBottom: 3, textTransform: "uppercase" }}>Current User</div>
        <div style={{ color: C.text, fontWeight: 700 }}>{pc.userName || "—"}</div>
      </div>
      <InputField label="New User Name" value={newName} onChange={setNewName} C={C} />
      <div style={{ color: C.muted, fontSize: 11, marginBottom: 14 }}>Or pick from existing:</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 20 }}>
        {Object.values(pcs).filter(p => p.pcId !== pc.pcId && p.userName).map(p => (
          <button key={p.pcId} onClick={() => setNewName(p.userName)}
            style={{ background: newName === p.userName ? C.accent + "33" : C.bg, border: `1px solid ${newName === p.userName ? C.accent : C.border}`, borderRadius: 6, padding: "4px 10px", color: newName === p.userName ? C.accent : C.muted, fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>
            {p.userName}
          </button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <Btn onClick={handle} variant="success" style={{ flex: 1 }} disabled={saving}>{saving ? "Saving..." : "✓ Confirm Swap"}</Btn>
        <Btn onClick={onCancel} variant="neutral">Cancel</Btn>
      </div>
    </div>
  );
}

// ─── PC Form ──────────────────────────────────────────────────────────────────
function PCForm({ initial, onSave, onCancel, C }) {
  const blank = { pcId: "", userName: "", autocadVersion: "", licenseType: "", serialNumber: "", productKey: "", operatingSystem: "", comments: "", teken: 0, gpu: { manufacturer: "", model: "", vramGb: "" }, cpu: { manufacturer: "", model: "", cores: "", ghz: "" }, ssd: { brand: "", capacityGb: "", type: "" }, ram: { sizeGb: "", speedMhz: "", sticks: "" }, user: { mail: "", password: "" } };
  const [form, setForm] = useState(initial ? { ...blank, ...initial, user: { mail: initial.user?.mail || "", password: "" } } : blank);
  const [errors, setErrors] = useState({});
  const [submitError, setSubmitError] = useState(null);
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setN = (g, k, v) => setForm(f => ({ ...f, [g]: { ...f[g], [k]: v } }));

  const handleSave = async () => {
    if (!form.pcId.trim()) { setErrors({ pcId: "PC ID is required" }); return; }
    setSaving(true); setSubmitError(null);
    try {
      const pc = { ...form };
      if (form.user.password) pc.user = { mail: form.user.mail, passwordHash: await hashPassword(form.user.password) };
      else pc.user = { mail: form.user.mail, passwordHash: initial?.user?.passwordHash || null };
      delete pc.user.password;
      await onSave(pc);
    } catch (err) { setSubmitError(err.message || "Unknown error saving PC"); }
    setSaving(false);
  };

  const sec = (label, color) => (
    <div style={{ color, fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, marginBottom: 10, marginTop: 6, display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ flex: 1, height: 1, background: color + "44" }} />{label}<div style={{ flex: 1, height: 1, background: color + "44" }} />
    </div>
  );

  return (
    <div style={{ background: C.card, borderRadius: 14, padding: "clamp(18px,3vw,28px)" }}>
      <h3 style={{ color: C.text, margin: "0 0 18px", fontSize: "clamp(14px,1.8vw,17px)" }}>{initial ? `✏️ Edit PC — ${initial.pcId}` : "➕ Add New PC"}</h3>
      {sec("Identity", C.accent)}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 10 }}>
        <InputField label="PC ID *" value={form.pcId} onChange={v => set("pcId", v)} C={C} error={errors.pcId} />
        <InputField label="User Name" value={form.userName} onChange={v => set("userName", v)} C={C} />
        <InputField label="AutoCAD Version" value={form.autocadVersion} onChange={v => set("autocadVersion", v)} C={C} />
        <InputField label="License Type" value={form.licenseType} onChange={v => set("licenseType", v)} C={C} />
        <InputField label="Serial Number" value={form.serialNumber} onChange={v => set("serialNumber", v)} C={C} />
        <InputField label="Product Key" value={form.productKey} onChange={v => set("productKey", v)} C={C} />
        <InputField label="Operating System" value={form.operatingSystem} onChange={v => set("operatingSystem", v)} C={C} />
      </div>
      {sec("CPU", C.accent2)}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10 }}>
        <InputField label="Manufacturer" value={form.cpu.manufacturer} onChange={v => setN("cpu", "manufacturer", v)} C={C} />
        <InputField label="Model" value={form.cpu.model} onChange={v => setN("cpu", "model", v)} C={C} />
        <InputField label="Cores" value={form.cpu.cores} onChange={v => setN("cpu", "cores", v)} C={C} />
        <InputField label="GHz" value={form.cpu.ghz} onChange={v => setN("cpu", "ghz", v)} C={C} />
      </div>
      {sec("GPU", C.accent2)}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10 }}>
        <InputField label="Manufacturer" value={form.gpu.manufacturer} onChange={v => setN("gpu", "manufacturer", v)} C={C} />
        <InputField label="Model" value={form.gpu.model} onChange={v => setN("gpu", "model", v)} C={C} />
        <InputField label="VRAM GB" value={form.gpu.vramGb} onChange={v => setN("gpu", "vramGb", v)} C={C} />
      </div>
      {sec("Storage & RAM", C.accent2)}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10 }}>
        <InputField label="SSD Brand" value={form.ssd.brand} onChange={v => setN("ssd", "brand", v)} C={C} />
        <InputField label="SSD Capacity GB" value={form.ssd.capacityGb} onChange={v => setN("ssd", "capacityGb", v)} C={C} />
        <InputField label="SSD Type" value={form.ssd.type} onChange={v => setN("ssd", "type", v)} C={C} />
        <InputField label="RAM GB" value={form.ram.sizeGb} onChange={v => setN("ram", "sizeGb", v)} C={C} />
        <InputField label="RAM MHz" value={form.ram.speedMhz} onChange={v => setN("ram", "speedMhz", v)} C={C} />
        <InputField label="RAM Sticks" value={form.ram.sticks} onChange={v => setN("ram", "sticks", v)} C={C} />
      </div>
      {sec("User / Email", C.green)}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <InputField label="Email" value={form.user.mail} onChange={v => setN("user", "mail", v)} C={C} />
        <InputField label={initial?.user?.passwordHash ? "New Password (blank = keep)" : "Outlook Password"} value={form.user.password} onChange={v => setN("user", "password", v)} type="password" C={C} />
      </div>
      {initial?.user?.passwordHash && <div style={{ color: C.green, fontSize: 11, marginTop: -6, marginBottom: 12 }}>✓ Password already set</div>}
      <div style={{ marginTop: 8 }}>
        <label style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", marginBottom: 6, display: "block" }}>Comments</label>
        <textarea value={form.comments} onChange={e => set("comments", e.target.value)} style={{ width: "100%", background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 7, padding: "9px 13px", color: C.text, fontSize: 13, fontFamily: "inherit", resize: "vertical", minHeight: 60, boxSizing: "border-box" }} />
        <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, cursor: "pointer", color: C.muted, fontSize: 13 }}>
          <input type="checkbox" checked={!!form.teken} onChange={e => set("teken", e.target.checked ? 1 : 0)} style={{ width: 16, height: 16, accentColor: C.green }} />Teken+
        </label>
      </div>
      {submitError && <div style={{ marginTop: 14, padding: "10px 14px", background: C.red + "18", border: `1px solid ${C.red}55`, borderRadius: 8, color: C.red, fontSize: 13, fontWeight: 600 }}>✗ Error: {submitError}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
        <Btn onClick={handleSave} variant="success" disabled={saving}>{saving ? "Saving..." : "💾 Save"}</Btn>
        <Btn onClick={onCancel} variant="neutral">Cancel</Btn>
      </div>
    </div>
  );
}

// ─── PC Detail ────────────────────────────────────────────────────────────────
function PCDetail({ pc, onEdit, onDelete, onSwap, isAdmin, C }) {
  const [verifyPwd, setVerifyPwd] = useState("");
  const [verifyResult, setVerifyResult] = useState(null);
  const handleVerify = async () => {
    if (!pc.user?.passwordHash) { setVerifyResult("no-hash"); return; }
    setVerifyResult(await verifyPassword(verifyPwd, pc.user.passwordHash) ? "match" : "mismatch");
  };
  const hw = (label, line1, line2) => (
    <div style={{ background: C.bg, borderRadius: 8, padding: "10px 12px", border: `1px solid ${C.border}` }}>
      <div style={{ color: C.muted, fontSize: 10, marginBottom: 3, textTransform: "uppercase" }}>{label}</div>
      <div style={{ color: C.text, fontSize: 12, fontWeight: 700 }}>{line1 || "—"}</div>
      {line2 && <div style={{ color: C.muted, fontSize: 11, marginTop: 2 }}>{line2}</div>}
    </div>
  );
  return (
    <div style={{ background: C.card, borderRadius: 12, padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        <div>
          <div style={{ color: C.text, fontSize: "clamp(14px,1.8vw,18px)", fontWeight: 800 }}>{pc.userName || "Unassigned"}</div>
          <div style={{ color: C.muted, fontSize: 12, fontFamily: "monospace" }}>{pc.pcId}</div>
        </div>
        {isAdmin && (
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
            {pc.teken ? <Tag color={C.green}>Teken+</Tag> : null}
            <Btn onClick={onSwap} variant="warning" style={{ padding: "5px 9px", fontSize: 11 }}>🔄</Btn>
            <Btn onClick={onEdit} variant="primary" style={{ padding: "5px 9px", fontSize: 11 }}>✏️</Btn>
            <Btn onClick={onDelete} variant="danger" style={{ padding: "5px 9px", fontSize: 11 }}>🗑️</Btn>
          </div>
        )}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0 }}>
        <Field label="AutoCAD" value={pc.autocadVersion} C={C} />
        <Field label="License" value={pc.licenseType} C={C} />
        <Field label="OS" value={pc.operatingSystem} C={C} />
        <Field label="Serial #" value={pc.serialNumber} mono C={C} />
      </div>
      <Field label="Product Key" value={pc.productKey} mono C={C} />
      <div style={{ color: C.accent2, fontSize: 10, fontWeight: 700, textTransform: "uppercase", marginBottom: 8, marginTop: 4 }}>Hardware</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginBottom: 14 }}>
        {hw("CPU", `${pc.cpu?.manufacturer || ""} ${pc.cpu?.model || ""}`.trim(), `${pc.cpu?.cores || "?"}  cores · ${pc.cpu?.ghz || "?"}GHz`)}
        {hw("GPU", `${pc.gpu?.manufacturer || ""} ${pc.gpu?.model || ""}`.trim(), `${pc.gpu?.vramGb || "?"}GB VRAM`)}
        {hw("SSD", `${pc.ssd?.capacityGb || "?"}GB ${pc.ssd?.type || ""}`.trim(), pc.ssd?.brand)}
        {hw("RAM", `${pc.ram?.sizeGb || "?"}GB`, `${pc.ram?.speedMhz || "?"}MHz · ${pc.ram?.sticks || "?"}×`)}
      </div>
      <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
        <div style={{ color: C.green, fontSize: 10, fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>User Info</div>
        <Field label="Email" value={pc.user?.mail} C={C} />
        <div style={{ color: C.muted, fontSize: 11, marginBottom: 8 }}>Password: {pc.user?.passwordHash ? <span style={{ color: C.green }}>✓ Set</span> : <span style={{ color: C.red }}>✗ Not set</span>}</div>
        {isAdmin && pc.user?.passwordHash && (
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input type="password" value={verifyPwd} onChange={e => setVerifyPwd(e.target.value)} placeholder="Verify password..."
              style={{ background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12, fontFamily: "inherit", outline: "none", flex: 1 }} />
            <Btn onClick={handleVerify} style={{ padding: "7px 12px", fontSize: 11 }}>Check</Btn>
            {verifyResult === "match" && <span style={{ color: C.green, fontSize: 14 }}>✓</span>}
            {verifyResult === "mismatch" && <span style={{ color: C.red, fontSize: 14 }}>✗</span>}
          </div>
        )}
        {pc.comments && <div style={{ marginTop: 12, color: C.muted, fontSize: 12, fontStyle: "italic", background: C.bg, padding: "8px 10px", borderRadius: 6 }}>💬 {pc.comments}</div>}
      </div>
    </div>
  );
}

// ─── Desk Form ────────────────────────────────────────────────────────────────
function DeskForm({ desk, tableGroups, pcs, onSave, onCancel, onDelete, C }) {
  const [label, setLabel] = useState(desk?.label || "");
  const [groupName, setGroupName] = useState(desk?.tableGroup || (tableGroups[0]?.name || ""));
  const [pcId, setPcId] = useState(desk?.pcId || "");
  const selectedGroup = tableGroups.find(g => g.name === groupName);
  const groupColor = selectedGroup?.color || "#6aa3fc";
  return (
    <div style={{ background: C.card, borderRadius: 14, padding: 24, minWidth: 300 }}>
      <h3 style={{ color: C.text, margin: "0 0 16px", fontSize: 15 }}>{desk ? `Edit Desk ${desk.label}` : "New Desk"}</h3>
      <InputField label="Desk Label" value={label} onChange={setLabel} placeholder="A1, B3..." C={C} />
      <div style={{ marginBottom: 14 }}>
        <label style={{ display: "block", color: C.muted, fontSize: 11, textTransform: "uppercase", marginBottom: 6 }}>Table Group</label>
        {tableGroups.length === 0 ? (
          <div style={{ color: C.red, fontSize: 12, padding: "8px 12px", background: C.red + "18", borderRadius: 7, border: `1px solid ${C.red}44` }}>⚠ No table groups yet — create one in the 🗂️ Tables tab first</div>
        ) : (
          <select value={groupName} onChange={e => setGroupName(e.target.value)} style={{ width: "100%", background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 7, padding: "9px 12px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none" }}>
            {tableGroups.map(g => <option key={g.name} value={g.name}>{g.name}</option>)}
          </select>
        )}
      </div>
      <div style={{ marginBottom: 16 }}>
        <label style={{ display: "block", color: C.muted, fontSize: 11, textTransform: "uppercase", marginBottom: 4 }}>Assign PC</label>
        <select value={pcId} onChange={e => setPcId(e.target.value)} style={{ width: "100%", background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 7, padding: "9px 12px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none" }}>
          <option value="">— Empty Seat —</option>
          {Object.values(pcs).map(pc => <option key={pc.pcId} value={pc.pcId}>{pc.userName} ({pc.pcId})</option>)}
        </select>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <Btn onClick={() => onSave({ label, tableGroup: groupName, pcId: pcId || null, color: groupColor })} variant="success" disabled={tableGroups.length === 0}>💾 Save</Btn>
        {desk && <Btn onClick={onDelete} variant="danger">🗑️ Remove</Btn>}
        <Btn onClick={onCancel} variant="neutral">Cancel</Btn>
      </div>
    </div>
  );
}

// ─── Compare ──────────────────────────────────────────────────────────────────
function ComparePCs({ pcs, C }) {
  const [selected, setSelected] = useState([]);
  const toggle = id => setSelected(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  const chosen = selected.map(id => pcs[id]).filter(Boolean);
  const CompRow = ({ label, fn }) => {
    const vals = chosen.map(pc => fn(pc));
    const nums = vals.map(v => parseFloat(v));
    const maxNum = Math.max(...nums.filter(n => !isNaN(n)));
    return (
      <tr>
        <td style={{ color: C.muted, fontSize: 12, padding: "8px 14px", whiteSpace: "nowrap", borderBottom: `1px solid ${C.border}`, fontWeight: 600 }}>{label}</td>
        {chosen.map((pc, i) => {
          const isMax = !isNaN(nums[i]) && nums[i] === maxNum && chosen.length > 1;
          return <td key={pc.pcId} style={{ padding: "8px 14px", color: isMax ? C.green : C.text, fontWeight: isMax ? 700 : 400, fontSize: 12, background: isMax ? C.green + "15" : "transparent", borderBottom: `1px solid ${C.border}` }}>{vals[i] || "—"}</td>;
        })}
      </tr>
    );
  };
  return (
    <div>
      <h3 style={{ color: C.text, margin: "0 0 16px", fontSize: "clamp(14px,1.8vw,18px)" }}>⚖️ Compare PCs</h3>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 20 }}>
        {Object.values(pcs).map(pc => (
          <button key={pc.pcId} onClick={() => toggle(pc.pcId)}
            style={{ background: selected.includes(pc.pcId) ? C.accent + "33" : C.card, border: `1px solid ${selected.includes(pc.pcId) ? C.accent : C.border}`, color: selected.includes(pc.pcId) ? C.accent : C.muted, borderRadius: 7, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>
            {pc.userName || pc.pcId}
          </button>
        ))}
      </div>
      {chosen.length >= 2 ? (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", background: C.card, borderRadius: 10, overflow: "hidden" }}>
            <thead>
              <tr style={{ background: C.border + "55" }}>
                <th style={{ color: C.muted, fontSize: 11, padding: "10px 14px", textAlign: "left", textTransform: "uppercase" }}>Spec</th>
                {chosen.map(pc => <th key={pc.pcId} style={{ color: C.accent, fontSize: 13, padding: "10px 14px", textAlign: "left" }}>{pc.userName}<br /><span style={{ color: C.muted, fontWeight: 400, fontSize: 10 }}>{pc.pcId}</span></th>)}
              </tr>
            </thead>
            <tbody>
              <CompRow label="OS" fn={p => p.operatingSystem} />
              <CompRow label="AutoCAD" fn={p => p.autocadVersion} />
              <CompRow label="CPU" fn={p => `${p.cpu?.manufacturer || ""} ${p.cpu?.model || ""}`.trim()} />
              <CompRow label="CPU Cores" fn={p => p.cpu?.cores} />
              <CompRow label="GPU" fn={p => `${p.gpu?.manufacturer || ""} ${p.gpu?.model || ""}`.trim()} />
              <CompRow label="VRAM GB" fn={p => p.gpu?.vramGb} />
              <CompRow label="RAM GB" fn={p => p.ram?.sizeGb} />
              <CompRow label="SSD" fn={p => `${p.ssd?.brand || ""} ${p.ssd?.capacityGb || ""}GB`} />
            </tbody>
          </table>
          <div style={{ color: C.green, fontSize: 11, marginTop: 8 }}>✦ Green = best value</div>
        </div>
      ) : (
        <div style={{ color: C.muted, background: C.card, borderRadius: 10, padding: 32, textAlign: "center", border: `1px solid ${C.border}` }}>Select at least 2 PCs to compare</div>
      )}
    </div>
  );
}

// ─── Table Group Manager ──────────────────────────────────────────────────────
const TABLE_COLORS = ["#6aa3fc", "#5de89a", "#fc6b6b", "#fde96a", "#fbb360", "#c084fc", "#38bdf8", "#f472b6", "#a3e635", "#fb923c"];

function TableGroupView({ desks, pcs, tableGroups, onSelectDesk, onGroupsChange, isAdmin, C }) {
  const [editingGroup, setEditingGroup] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [newName, setNewName] = useState("");
  const [newColor, setNewColor] = useState(TABLE_COLORS[0]);

  const saveNewGroup = async () => {
    if (!newName.trim()) return;
    await onGroupsChange([...tableGroups, { name: newName.trim(), color: newColor }]);
    setNewName(""); setNewColor(TABLE_COLORS[0]); setShowAdd(false);
  };
  const saveEditGroup = async () => {
    await onGroupsChange(tableGroups.map(g => g.name === editingGroup.original ? { name: editingGroup.name, color: editingGroup.color } : g));
    setEditingGroup(null);
  };
  const deleteGroup = async (name) => {
    if (!confirm(`Delete table group "${name}"?`)) return;
    await onGroupsChange(tableGroups.filter(g => g.name !== name));
  };

  const ColorPicker = ({ selected, onSelect }) => (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", marginTop: 8 }}>
      {TABLE_COLORS.map(c => (
        <button key={c} onClick={() => onSelect(c)} style={{ width: 24, height: 24, borderRadius: "50%", background: c, border: selected === c ? `3px solid white` : "2px solid transparent", cursor: "pointer", transform: selected === c ? "scale(1.2)" : "scale(1)", transition: "all 0.15s", flexShrink: 0 }} />
      ))}
      <input type="color" value={selected} onChange={e => onSelect(e.target.value)} style={{ width: 28, height: 28, border: "none", background: "none", cursor: "pointer", borderRadius: "50%" }} title="Custom color" />
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
        <h3 style={{ color: C.text, margin: 0, fontSize: "clamp(14px,1.8vw,18px)" }}>🗂️ Table Groups</h3>
        {isAdmin && <Btn onClick={() => setShowAdd(p => !p)} variant="success" style={{ padding: "7px 14px", fontSize: 12 }}>{showAdd ? "✕ Cancel" : "+ New Table Group"}</Btn>}
      </div>
      {showAdd && isAdmin && (
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <InputField label="Group Name" value={newName} onChange={setNewName} placeholder="e.g. Table A, Drafting Room..." C={C} />
          <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", marginBottom: 4 }}>Color</div>
          <ColorPicker selected={newColor} onSelect={setNewColor} />
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={saveNewGroup} variant="success" disabled={!newName.trim()}>💾 Create Group</Btn>
            <Btn onClick={() => setShowAdd(false)} variant="neutral">Cancel</Btn>
          </div>
        </div>
      )}
      {tableGroups.length === 0 && !showAdd && (
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 40, textAlign: "center" }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>🗂️</div>
          <div style={{ color: C.muted, fontSize: 14 }}>No table groups yet</div>
        </div>
      )}
      {tableGroups.map(group => {
        const gd = desks.filter(d => d.tableGroup === group.name);
        const isEditing = editingGroup?.original === group.name;
        return (
          <div key={group.name} style={{ background: C.card, border: `1px solid ${C.border}`, borderLeft: `4px solid ${group.color}`, borderRadius: 12, marginBottom: 16, overflow: "hidden" }}>
            {isEditing ? (
              <div style={{ padding: "14px 18px", background: C.border + "22" }}>
                <input value={editingGroup.name} onChange={e => setEditingGroup(g => ({ ...g, name: e.target.value }))} style={{ width: "100%", background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 7, padding: "7px 12px", color: C.text, fontSize: 14, fontFamily: "inherit", outline: "none", marginBottom: 10 }} />
                <ColorPicker selected={editingGroup.color} onSelect={c => setEditingGroup(g => ({ ...g, color: c }))} />
                <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                  <Btn onClick={saveEditGroup} variant="success" style={{ padding: "6px 12px", fontSize: 11 }}>✓ Save</Btn>
                  <Btn onClick={() => setEditingGroup(null)} variant="neutral" style={{ padding: "6px 12px", fontSize: 11 }}>Cancel</Btn>
                  <Btn onClick={() => deleteGroup(group.name)} variant="danger" style={{ padding: "6px 12px", fontSize: 11 }}>🗑️ Delete</Btn>
                </div>
              </div>
            ) : (
              <div style={{ padding: "14px 18px", display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 14, height: 14, borderRadius: "50%", background: group.color, flexShrink: 0 }} />
                <span style={{ color: C.text, fontWeight: 700, fontSize: 15, flex: 1 }}>{group.name}</span>
                <span style={{ color: C.muted, fontSize: 12 }}>{gd.length} seat{gd.length !== 1 ? "s" : ""}</span>
                {isAdmin && <Btn onClick={() => setEditingGroup({ original: group.name, name: group.name, color: group.color })} variant="primary" style={{ padding: "4px 10px", fontSize: 11 }}>✏️ Edit</Btn>}
              </div>
            )}
            {gd.length > 0 && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 8, padding: "0 16px 16px" }}>
                {gd.map(desk => {
                  const pc = desk.pcId ? pcs[desk.pcId] : null;
                  return (
                    <div key={desk.id} onClick={() => onSelectDesk(desk)} style={{ background: C.bg, border: `1.5px solid ${pc ? group.color + "88" : C.border}`, borderRadius: 8, padding: "10px 12px", cursor: "pointer" }}>
                      <div style={{ color: group.color, fontSize: 10, fontWeight: 700, marginBottom: 4 }}>DESK {desk.label}</div>
                      {pc ? <><div style={{ color: C.text, fontSize: 13, fontWeight: 700 }}>{pc.userName}</div><div style={{ color: C.muted, fontSize: 11 }}>{pc.pcId}</div></> : <div style={{ color: C.dim, fontSize: 12 }}>Empty</div>}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Floor Map ────────────────────────────────────────────────────────────────
function FloorMapFull({ desks, pcs, onDeskClick, onAddDesk, editMode, floorImage, C }) {
  const containerRef = useRef(null);
  const [dragging, setDragging] = useState(null);
  const [localDesks, setLocalDesks] = useState(desks);
  useEffect(() => setLocalDesks(desks), [desks]);

  const handleMapClick = useCallback((e) => {
    if (!editMode || dragging) return;
    const rect = containerRef.current.getBoundingClientRect();
    onAddDesk({ x: ((e.clientX - rect.left) / rect.width) * 100, y: ((e.clientY - rect.top) / rect.height) * 100 });
  }, [editMode, dragging, onAddDesk]);

  const handleMouseMove = useCallback((e) => {
    if (!dragging || !containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    const x = Math.max(1, Math.min(99, ((e.clientX - rect.left) / rect.width) * 100));
    const y = Math.max(1, Math.min(99, ((e.clientY - rect.top) / rect.height) * 100));
    setLocalDesks(ds => ds.map(d => d.id === dragging ? { ...d, x, y } : d));
  }, [dragging]);

  const handleMouseUp = useCallback(() => {
    if (dragging) { const moved = localDesks.find(d => d.id === dragging); if (moved) onDeskClick(moved, true); setDragging(null); }
  }, [dragging, localDesks, onDeskClick]);

  return (
    <div ref={containerRef} onClick={handleMapClick} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp}
      style={{ position: "relative", width: "100%", height: "100%", background: C.panel, borderRadius: 10, overflow: "hidden", cursor: editMode ? "crosshair" : "default", userSelect: "none", border: `1px solid ${C.border}` }}>
      {floorImage
        ? <img src={floorImage} alt="floor" style={{ position: "absolute", inset: 0, width: "100%", height: "100%", objectFit: "contain", opacity: 0.92 }} />
        : <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
            <div style={{ color: C.dim, textAlign: "center" }}>
              <div style={{ fontSize: 52, marginBottom: 12 }}>🏢</div>
              <div style={{ fontSize: 15, color: C.muted }}>Upload a floor plan image</div>
            </div>
          </div>
      }
      {localDesks.map(desk => {
        const pc = desk.pcId ? pcs[desk.pcId] : null;
        const color = desk.color || (pc ? C.accent : C.dim);
        return (
          <div key={desk.id}
            onMouseDown={e => { if (!editMode) return; e.stopPropagation(); setDragging(desk.id); }}
            onClick={e => { e.stopPropagation(); if (!dragging) onDeskClick(desk, false); }}
            style={{ position: "absolute", left: `${desk.x}%`, top: `${desk.y}%`, transform: "translate(-50%,-50%)", width: 36, height: 36, borderRadius: "50%", background: color + "33", border: `2.5px solid ${color}`, display: "flex", alignItems: "center", justifyContent: "center", cursor: editMode ? "grab" : "pointer", zIndex: dragging === desk.id ? 10 : 2, boxShadow: `0 0 10px ${color}66` }}>
            <span style={{ color, fontSize: 9, fontWeight: 800, textAlign: "center" }}>{desk.label}</span>
          </div>
        );
      })}
      {editMode && <div style={{ position: "absolute", top: 12, right: 12, background: C.accent + "22", border: `1px solid ${C.accent}66`, borderRadius: 8, padding: "6px 14px", color: C.accent, fontSize: 12, fontWeight: 700, pointerEvents: "none" }}>✏️ Edit — Click to add · Drag to move</div>}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(loadSession);
  const [themeKey, setThemeKey] = useState("dark");
  const C = THEMES[themeKey];
  const isAdmin = session?.isAdmin;

  const [pcs, setPcs] = useState({});
  const [desks, setDesks] = useState([]);
  const [tableGroups, setTableGroups] = useState([]);
  const [selectedDesk, setSelectedDesk] = useState(null);
  const [editMode, setEditMode] = useState(false);
  const [activeTab, setActiveTab] = useState("map");
  const [modal, setModal] = useState(null);
  const [floorImage, setFloorImage] = useState(null);
  const [search, setSearch] = useState("");
  const [toast, setToast] = useState(null);
  const [importing, setImporting] = useState(false);
  const [loading, setLoading] = useState(false);

  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 4500); };

  const loadAll = async () => {
    setLoading(true);
    try {
      const [pcsData, desksData, floorData, tgData] = await Promise.all([
        getPCs(), getDesks(), getFloorPlan(), getTableGroups(),
      ]);
      setPcs(pcsData);
      setDesks(desksData);
      if (floorData) setFloorImage(floorData);
      setTableGroups(tgData);
    } catch (e) { showToast("Failed to load data: " + e.message, "error"); }
    setLoading(false);
  };

  const handleLogin = (data) => { persistSession(data); setSession(data); };
  useEffect(() => { if (session) loadAll(); }, [session]);

  const handleLogout = () => { persistSession(null); setSession(null); setPcs({}); setDesks([]); setFloorImage(null); setSelectedDesk(null); };

  const handleImportExcel = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setImporting(true);
    try {
      const imported = await importFromExcel(file);
      const pcArray = Object.values(imported);
      if (pcArray.length === 0) { showToast("No PCs found — check column names", "error"); setImporting(false); e.target.value = ""; return; }
      const result = await bulkSavePCs(pcArray);
      const fresh = await getPCs();
      setPcs(fresh);
      if (result.errors?.length) showToast(`⚠️ ${result.count} updated/created, ${result.errors.length} failed`, "error");
      else showToast(`✓ ${result.count} PCs updated/created`);
    } catch (err) { showToast(`✗ ${err.message}`, "error"); }
    setImporting(false); e.target.value = "";
  };

  const handleFloorUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      const imageData = ev.target.result;
      setFloorImage(imageData);
      await saveFloorPlan(imageData);
      showToast("✓ Floor plan saved");
    };
    reader.readAsDataURL(file);
  };

  const handleGroupsChange = async (updated) => {
    setTableGroups(updated);
    await saveTableGroups(updated);
    setDesks(ds => ds.map(d => { const g = updated.find(g => g.name === d.tableGroup); return g ? { ...d, color: g.color } : d; }));
  };

  const handleDeskClick = (desk, isPos) => {
    if (isPos) {
      setDesks(ds => ds.map(d => d.id === desk.id ? { ...d, x: desk.x, y: desk.y } : d));
      saveDesk(desk);
      return;
    }
    if (editMode) setModal({ type: "editDesk", data: desk });
    else setSelectedDesk(desk);
  };

  const handleAddDesk = ({ x, y }) => {
    const firstGroup = tableGroups[0];
    setModal({ type: "editDesk", data: { id: `desk-${Date.now()}`, x, y, pcId: null, label: `D${desks.length + 1}`, tableGroup: firstGroup?.name || "", color: firstGroup?.color || "#6aa3fc" }, isNew: true });
  };

  const saveDeskModal = async (updated) => {
    const group = tableGroups.find(g => g.name === updated.tableGroup);
    const withColor = { ...updated, color: group?.color || "#6aa3fc" };
    const desk = modal.isNew ? { ...modal.data, ...withColor } : { ...desks.find(d => d.id === modal.data.id), ...withColor };
    if (modal.isNew) { setDesks(ds => [...ds, desk]); await saveDesk(desk); }
    else { setDesks(ds => ds.map(d => d.id === desk.id ? desk : d)); await saveDesk(desk); }
    setModal(null);
  };

  const handleDeleteDesk = async () => {
    setDesks(ds => ds.filter(d => d.id !== modal.data.id));
    await deleteDesk(modal.data.id);
    setSelectedDesk(null); setModal(null);
  };

  const handleSavePC = async (pc) => {
    await savePC(pc);
    setPcs(p => ({ ...p, [pc.pcId]: pc }));
    setModal(null);
    showToast(`✓ PC ${pc.pcId} saved`);
  };

  const handleDeletePC = async (pcId) => {
    setPcs(p => { const n = { ...p }; delete n[pcId]; return n; });
    setDesks(ds => ds.map(d => d.pcId === pcId ? { ...d, pcId: null } : d));
    await deletePC(pcId);
    setSelectedDesk(null);
  };

  const handleSwap = async (pc, newUserName) => {
    setPcs(p => ({ ...p, [pc.pcId]: { ...p[pc.pcId], userName: newUserName } }));
    setModal(null);
    showToast(`✓ ${pc.pcId} reassigned to ${newUserName}`);
  };

  const filteredPCs = Object.values(pcs).filter(pc =>
    !search || pc.userName?.toLowerCase().includes(search.toLowerCase()) ||
    pc.pcId?.toLowerCase().includes(search.toLowerCase()) ||
    pc.user?.mail?.toLowerCase().includes(search.toLowerCase())
  );

  const selectedPC = selectedDesk?.pcId ? pcs[selectedDesk.pcId] : null;

  if (!session) return <LoginScreen onLogin={handleLogin} C={C} />;

  return (
    <div style={{ background: C.bg, height: "100vh", width: "100vw", fontFamily: "'DM Mono','Courier New',monospace", color: C.text, display: "flex", flexDirection: "column", overflow: "hidden" }}>
      {toast && <div style={{ position: "fixed", top: 16, right: 16, zIndex: 999, background: toast.type === "error" ? C.red + "22" : C.green + "22", border: `1px solid ${toast.type === "error" ? C.red : C.green}88`, borderRadius: 10, padding: "12px 20px", color: toast.type === "error" ? C.red : C.green, fontSize: 13, fontWeight: 700, boxShadow: C.shadow, maxWidth: 360 }}>{toast.msg}</div>}

      {/* Header */}
      <div style={{ background: C.panel, borderBottom: `1px solid ${C.border}`, padding: "0 clamp(12px,2vw,24px)", display: "flex", alignItems: "center", justifyContent: "space-between", height: "clamp(48px,6vh,60px)", flexShrink: 0, boxShadow: C.shadow, gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "clamp(8px,1.5vw,18px)" }}>
          <div style={{ background: C.accent + "22", border: `1px solid ${C.accent}55`, borderRadius: 8, padding: "5px clamp(8px,1.5vw,14px)", color: C.accent, fontWeight: 800, fontSize: "clamp(11px,1.4vw,14px)", letterSpacing: 1, whiteSpace: "nowrap" }}>⚙ OFFICE MGR</div>
          <span style={{ color: C.text, fontSize: "clamp(11px,1.2vw,13px)" }}><b style={{ color: C.accent }}>{Object.keys(pcs).length}</b> PCs</span>
          <span style={{ color: C.dim }}>·</span>
          <span style={{ color: C.text, fontSize: "clamp(11px,1.2vw,13px)" }}><b style={{ color: C.accent }}>{desks.length}</b> Desks</span>
          {loading && <span style={{ color: C.muted, fontSize: 12 }}>⏳ Loading...</span>}
        </div>
        <div style={{ display: "flex", gap: "clamp(4px,0.8vw,10px)", alignItems: "center" }}>
          <button onClick={() => setThemeKey(k => k === "dark" ? "light" : "dark")} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 7, padding: "6px 10px", color: C.muted, fontSize: 15, cursor: "pointer" }}>{themeKey === "dark" ? "☀️" : "🌙"}</button>
          {isAdmin && <>
            <label style={{ background: C.dim + "22", border: `1px solid ${C.border}`, borderRadius: 7, padding: "6px clamp(8px,1vw,14px)", color: C.muted, fontSize: "clamp(11px,1.1vw,13px)", cursor: "pointer", fontWeight: 700, whiteSpace: "nowrap" }}>
              🖼️ Plan <input type="file" accept="image/*" onChange={handleFloorUpload} style={{ display: "none" }} />
            </label>
            <label style={{ background: C.yellow + "18", border: `1px solid ${C.yellow}55`, borderRadius: 7, padding: "6px clamp(8px,1vw,14px)", color: C.yellow, fontSize: "clamp(11px,1.1vw,13px)", cursor: "pointer", fontWeight: 700, whiteSpace: "nowrap" }}>
              {importing ? "⏳" : "📥 Import"} <input type="file" accept=".xlsx,.xls" onChange={handleImportExcel} style={{ display: "none" }} disabled={importing} />
            </label>
            <button onClick={() => exportToExcel(pcs, desks)} style={{ background: C.dim + "22", border: `1px solid ${C.border}`, borderRadius: 7, padding: "6px clamp(8px,1vw,12px)", color: C.muted, fontSize: "clamp(11px,1.1vw,13px)", cursor: "pointer", fontFamily: "inherit", fontWeight: 700, whiteSpace: "nowrap" }}>📤 Export</button>
            <button onClick={() => setModal({ type: "addPC" })} style={{ background: C.green + "22", border: `1px solid ${C.green}55`, borderRadius: 7, padding: "6px clamp(8px,1vw,14px)", color: C.green, fontSize: "clamp(11px,1.1vw,13px)", cursor: "pointer", fontFamily: "inherit", fontWeight: 800, whiteSpace: "nowrap" }}>+ PC</button>
          </>}
          <div style={{ display: "flex", alignItems: "center", gap: 8, borderLeft: `1px solid ${C.border}`, paddingLeft: "clamp(6px,1vw,12px)" }}>
            <div>
              <div style={{ color: C.text, fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>{session.displayName || session.email}</div>
              <div style={{ color: isAdmin ? C.accent : C.muted, fontSize: 10 }}>{isAdmin ? "Admin" : "Viewer"}</div>
            </div>
            <button onClick={handleLogout} style={{ background: C.red + "18", border: `1px solid ${C.red}44`, borderRadius: 6, padding: "5px 10px", color: C.red, fontSize: "clamp(10px,1vw,12px)", cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>Sign out</button>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: C.panel, borderBottom: `1px solid ${C.border}`, padding: "0 clamp(10px,1.5vw,24px)", display: "flex", gap: 2, flexShrink: 0, overflowX: "auto" }}>
        {[{ id: "map", label: "🗺️ Map" }, { id: "tables", label: "🗂️ Tables" }, { id: "list", label: "📋 All PCs" }, { id: "compare", label: "⚖️ Compare" }, ...(isAdmin ? [{ id: "admin", label: "👥 Users" }] : [])].map(tab => (
          <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? `2.5px solid ${C.accent}` : "2.5px solid transparent", color: activeTab === tab.id ? C.accent : C.muted, padding: "clamp(7px,1.2vh,11px) clamp(10px,1.5vw,20px)", fontSize: "clamp(11px,1.2vw,13px)", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap", marginBottom: -1 }}>{tab.label}</button>
        ))}
      </div>

      {/* Body */}
      <div style={{ flex: 1, overflow: "hidden", position: "relative" }}>
        <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column" }}>
          {activeTab === "map" && (
            <div style={{ flex: 1, padding: "clamp(10px,1.5vh,16px) clamp(12px,1.5vw,22px)", display: "flex", flexDirection: "column", overflow: "hidden" }}>
              <div style={{ display: "flex", gap: 10, marginBottom: 10, alignItems: "center", flexShrink: 0, flexWrap: "wrap" }}>
                <div style={{ display: "flex", gap: "clamp(8px,1.5vw,16px)", alignItems: "center", flexWrap: "wrap" }}>
                  {tableGroups.map(g => (
                    <div key={g.name} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                      <div style={{ width: 11, height: 11, borderRadius: "50%", background: g.color, boxShadow: `0 0 6px ${g.color}99`, flexShrink: 0 }} />
                      <span style={{ color: C.muted, fontSize: 12, whiteSpace: "nowrap" }}>{g.name}</span>
                    </div>
                  ))}
                </div>
                <div style={{ flex: 1 }} />
                {isAdmin && <button onClick={() => setEditMode(m => !m)} style={{ background: editMode ? C.red + "22" : C.dim + "22", border: `1px solid ${editMode ? C.red + "88" : C.border}`, borderRadius: 7, padding: "6px 14px", color: editMode ? C.red : C.muted, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{editMode ? "✓ Done Editing" : "✏️ Edit Map"}</button>}
              </div>
              <div style={{ flex: 1, minHeight: 0 }}>
                <FloorMapFull desks={desks} pcs={pcs} onDeskClick={handleDeskClick} onAddDesk={handleAddDesk} editMode={editMode && isAdmin} floorImage={floorImage} C={C} />
              </div>
            </div>
          )}
          {activeTab === "tables" && (
            <div style={{ flex: 1, overflowY: "auto", padding: "clamp(14px,2vh,22px) clamp(14px,2vw,28px)" }}>
              <TableGroupView desks={desks} pcs={pcs} tableGroups={tableGroups} onSelectDesk={d => { setSelectedDesk(d); setActiveTab("map"); }} onGroupsChange={handleGroupsChange} isAdmin={isAdmin} C={C} />
            </div>
          )}
          {activeTab === "list" && (
            <div style={{ flex: 1, overflowY: "auto", padding: "clamp(14px,2vh,22px) clamp(14px,2vw,28px)" }}>
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍  Search by name, PC ID, email..." style={{ background: C.inputBg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "clamp(8px,1.2vh,11px) 16px", color: C.text, fontSize: "clamp(12px,1.3vw,14px)", fontFamily: "inherit", outline: "none", width: "100%", boxSizing: "border-box", marginBottom: 18 }} />
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px,1fr))", gap: 14 }}>
                {filteredPCs.map(pc => (
                  <div key={pc.pcId} onClick={() => setSelectedDesk({ pcId: pc.pcId, label: "—", tableGroup: "—" })} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: "clamp(12px,1.5vw,18px)", cursor: "pointer", transition: "all 0.15s" }} onMouseEnter={e => { e.currentTarget.style.borderColor = C.accent + "88"; }} onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; }}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                      <div>
                        <div style={{ color: C.text, fontWeight: 800, fontSize: "clamp(13px,1.4vw,15px)" }}>{pc.userName || <span style={{ color: C.dim }}>Unassigned</span>}</div>
                        <div style={{ color: C.muted, fontSize: 12, fontFamily: "monospace" }}>{pc.pcId}</div>
                      </div>
                      {pc.teken ? <Tag color={C.green}>Teken+</Tag> : null}
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
                      <div style={{ color: C.muted, fontSize: 12 }}>🖥 {pc.cpu?.model}</div>
                      <div style={{ color: C.muted, fontSize: 12 }}>🎮 {pc.gpu?.model}</div>
                      <div style={{ color: C.muted, fontSize: 12 }}>💾 {pc.ram?.sizeGb}GB RAM</div>
                      <div style={{ color: C.muted, fontSize: 12 }}>💿 {pc.ssd?.capacityGb}GB {pc.ssd?.type}</div>
                    </div>
                    {pc.user?.mail && <div style={{ color: C.dim, fontSize: 11, marginTop: 8 }}>✉ {pc.user.mail}</div>}
                  </div>
                ))}
              </div>
            </div>
          )}
          {activeTab === "compare" && <div style={{ flex: 1, overflowY: "auto", padding: "clamp(14px,2vh,22px) clamp(14px,2vw,28px)" }}><ComparePCs pcs={pcs} C={C} /></div>}
          {activeTab === "admin" && isAdmin && <div style={{ flex: 1, overflowY: "auto" }}><AdminPanel C={C} /></div>}
        </div>

        {/* Sidebar */}
        <div style={{ position: "absolute", top: 0, right: 0, bottom: 0, width: selectedDesk ? "clamp(300px,28vw,420px)" : 0, borderLeft: selectedDesk ? `1px solid ${C.border}` : "none", background: C.panel, overflowY: "auto", overflowX: "hidden", transition: "width 0.22s cubic-bezier(0.4,0,0.2,1)", zIndex: 20, boxShadow: selectedDesk ? "-4px 0 20px #0004" : "none" }}>
          {selectedDesk && (
            <div style={{ width: "clamp(300px,28vw,420px)", padding: "clamp(12px,1.5vw,20px)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase", letterSpacing: 1 }}>Desk {selectedDesk.label}</div>
                  <div style={{ color: C.dim, fontSize: 11 }}>{selectedDesk.tableGroup}</div>
                </div>
                <button onClick={() => setSelectedDesk(null)} style={{ background: C.dim + "33", border: "none", color: C.muted, cursor: "pointer", fontSize: 18, width: 30, height: 30, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
              </div>
              {selectedPC ? (
                <PCDetail pc={selectedPC} isAdmin={isAdmin} C={C} onEdit={() => setModal({ type: "editPC", data: selectedPC })} onDelete={() => handleDeletePC(selectedPC.pcId)} onSwap={() => setModal({ type: "swap", data: selectedPC })} />
              ) : (
                <div style={{ background: C.card, borderRadius: 12, padding: "clamp(20px,3vh,36px)", textAlign: "center", border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 44, marginBottom: 12 }}>🖥️</div>
                  <div style={{ color: C.muted, fontSize: 14, marginBottom: 20 }}>Empty desk</div>
                  {isAdmin && <button onClick={() => setModal({ type: "editDesk", data: selectedDesk })} style={{ background: C.accent + "22", border: `1px solid ${C.accent}55`, borderRadius: 8, padding: "10px 24px", color: C.accent, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Assign PC</button>}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Modals */}
      {modal && (
        <div style={{ position: "fixed", inset: 0, background: "#0009", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: "clamp(16px,3vw,40px)" }} onClick={e => { if (e.target === e.currentTarget) setModal(null); }}>
          <div style={{ maxWidth: 680, width: "100%", maxHeight: "90vh", overflowY: "auto", borderRadius: 14 }}>
            {(modal.type === "addPC" || modal.type === "editPC") && <PCForm initial={modal.data} onSave={handleSavePC} onCancel={() => setModal(null)} C={C} />}
            {modal.type === "editDesk" && <DeskForm desk={modal.isNew ? null : modal.data} tableGroups={tableGroups} pcs={pcs} onSave={saveDeskModal} onCancel={() => setModal(null)} onDelete={handleDeleteDesk} C={C} />}
            {modal.type === "swap" && <SwapModal pc={modal.data} pcs={pcs} onSave={name => handleSwap(modal.data, name)} onCancel={() => setModal(null)} C={C} />}
          </div>
        </div>
      )}
    </div>
  );
}
