import { useState, useEffect, useRef } from "react";

const API = "http://127.0.0.1:5001/api";

// ── tiny SVG icon helper ──────────────────────────────────────────────────────
function Ic({ d, size = 18, color = "currentColor" }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
      stroke={color} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
      <path d={d} />
    </svg>
  );
}
const I = {
  upload:   "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12",
  play:     "M5 3l14 9-14 9V3z",
  dl:       "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3",
  check:    "M20 6L9 17l-5-5",
  warn:     "M12 9v4M12 17h.01M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z",
  hosp:     "M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2zM9 22V12h6v10M12 6v4M10 8h4",
  file:     "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6",
  user:     "M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8z",
  log:      "M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01",
  refresh:  "M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15",
};

// ── shared styles ─────────────────────────────────────────────────────────────
const card = {
  background: "#fff", borderRadius: 14,
  border: "1px solid #E2E8F0",
  padding: "24px 28px",
  boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
};
const btn = (bg, color = "#fff") => ({
  background: bg, color, border: "none",
  borderRadius: 9, padding: "11px 22px",
  fontWeight: 700, fontSize: 14, cursor: "pointer",
  display: "flex", alignItems: "center", gap: 8,
  transition: "opacity .15s",
});
const th = { padding: "9px 12px", color: "#fff", fontWeight: 700, fontSize: 12, textAlign: "left" };
const td = { padding: "8px 12px", fontSize: 13, borderBottom: "1px solid #EDF2F7" };

// ── Step indicator ────────────────────────────────────────────────────────────
function Steps({ current }) {
  const steps = ["Upload Excel", "Configure", "Processing", "Download"];
  return (
    <div style={{ display: "flex", alignItems: "center", marginBottom: 32 }}>
      {steps.map((s, i) => {
        const done   = current > i + 1;
        const active = current === i + 1;
        return (
          <div key={i} style={{ display: "flex", alignItems: "center", flex: i < 3 ? 1 : "none" }}>
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 5 }}>
              <div style={{
                width: 30, height: 30, borderRadius: "50%", fontWeight: 700, fontSize: 13,
                display: "flex", alignItems: "center", justifyContent: "center",
                background: done ? "#1D9E75" : active ? "#1F4E79" : "rgba(255,255,255,.2)",
                color: "#fff",
              }}>
                {done ? <Ic d={I.check} size={15} color="#fff" /> : i + 1}
              </div>
              <span style={{ fontSize: 11, color: active || done ? "#fff" : "rgba(255,255,255,.45)", whiteSpace: "nowrap" }}>
                {s}
              </span>
            </div>
            {i < 3 && (
              <div style={{
                flex: 1, height: 2, margin: "0 8px", marginBottom: 18,
                background: done ? "#1D9E75" : "rgba(255,255,255,.15)",
              }} />
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep]     = useState(1);
  const [preview, setPreview] = useState(null);  // { count, preview[] }
  const [month, setMonth]   = useState("");
  const [drag, setDrag]     = useState(false);
  const [err, setErr]       = useState("");
  const [job, setJob]       = useState(null);
  const [logs, setLogs]     = useState([]);
  const fileRef  = useRef();
  const logRef   = useRef();
  const pollRef  = useRef();

  // Auto month label
  useEffect(() => {
    const d = new Date();
    setMonth(`${d.toLocaleString("default",{month:"long"})} ${String(d.getFullYear()).slice(-2)}`);
  }, []);

  // Scroll logs
  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [logs]);

  // Poll while running
  useEffect(() => {
    if (step === 3) {
      pollRef.current = setInterval(async () => {
        try {
          const r = await fetch(`${API}/status`);
          const s = await r.json();
          setJob(s);
          setLogs(s.logs || []);
          if (s.finished) { clearInterval(pollRef.current); setStep(4); }
        } catch (_) {}
      }, 1500);
    }
    return () => clearInterval(pollRef.current);
  }, [step]);

  // ── Handlers ────────────────────────────────────────────────────────────────
  async function handleFile(f) {
    if (!f || !f.name.endsWith(".xlsx")) { setErr("Please upload a .xlsx file"); return; }
    setErr("");
    const form = new FormData(); form.append("file", f);
    try {
      const r = await fetch(`${API}/upload`, { method: "POST", body: form });
      const d = await r.json();
      if (d.error) { setErr(d.error); return; }
      setPreview(d); setStep(2);
    } catch (_) {
      setErr("Cannot connect to backend. Make sure 'python server.py' is running.");
    }
  }

  async function handleStart() {
    try {
      const r = await fetch(`${API}/start`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ month }),
      });
      const d = await r.json();
      if (d.error) { alert(d.error); return; }
      setStep(3);
    } catch (_) { alert("Cannot connect to backend."); }
  }

  const progress = job ? Math.round((job.done / Math.max(job.total, 1)) * 100) : 0;

  // ── Render ───────────────────────────────────────────────────────────────────
  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(140deg,#0B1F35 0%,#1F4E79 55%,#2E75B6 100%)",
      fontFamily: "'Segoe UI',Arial,sans-serif",
      paddingBottom: 60,
    }}>

      {/* Header */}
      <div style={{
        padding: "16px 40px", display: "flex", alignItems: "center", gap: 12,
        borderBottom: "1px solid rgba(255,255,255,.1)",
        background: "rgba(0,0,0,.15)",
      }}>
        <Ic d={I.hosp} size={26} color="#7EC8E3" />
        <div>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 17 }}>HBS General Hospital</div>
          <div style={{ color: "#7EC8E3", fontSize: 11 }}>Lab Bill Automation System By Muhammad Imran — hbs.chealth.pk</div>
        </div>
        <div style={{ marginLeft: "auto", background: "rgba(255,255,255,.12)",
          borderRadius: 99, padding: "4px 14px", color: "#7EC8E3", fontSize: 11, fontWeight: 600 }}>
          AI Powered
        </div>
      </div>

      <div style={{ maxWidth: 760, margin: "36px auto", padding: "0 20px" }}>
        <Steps current={step} />

        {/* ── STEP 1: UPLOAD ── */}
        {step === 1 && (
          <div style={card}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
              <Ic d={I.upload} color="#1F4E79" size={22} />
              <h2 style={{ margin: 0, fontSize: 18, color: "#1F4E79" }}>Upload Admission Excel</h2>
            </div>
            <p style={{ color: "#555", fontSize: 14, marginBottom: 22 }}>
              Upload your monthly <strong>May26.xlsx</strong> (or similar). The system reads
              <strong> Column B</strong> (MRN) and <strong>Column C</strong> (Patient Name).
            </p>

            <div
              onDragOver={e => { e.preventDefault(); setDrag(true); }}
              onDragLeave={() => setDrag(false)}
              onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
              onClick={() => fileRef.current.click()}
              style={{
                border: `2px dashed ${drag ? "#1F4E79" : "#CBD5E0"}`,
                borderRadius: 12, padding: "52px 24px", textAlign: "center",
                cursor: "pointer", background: drag ? "#EBF4FF" : "#FAFBFC",
                transition: "all .2s",
              }}>
              <Ic d={I.file} size={44} color={drag ? "#1F4E79" : "#A0AEC0"} />
              <p style={{ margin: "12px 0 4px", fontWeight: 600, color: "#2D3748", fontSize: 15 }}>
                Drop your Excel file here
              </p>
              <p style={{ margin: 0, fontSize: 13, color: "#999" }}>or click to browse — .xlsx only</p>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx" style={{ display: "none" }}
              onChange={e => handleFile(e.target.files[0])} />

            {err && (
              <div style={{
                marginTop: 14, padding: "12px 14px", background: "#FAECE7",
                borderRadius: 8, color: "#4A1B0C", fontSize: 13,
                display: "flex", gap: 8, alignItems: "flex-start",
              }}>
                <Ic d={I.warn} size={16} color="#C00000" /> {err}
              </div>
            )}
          </div>
        )}

        {/* ── STEP 2: CONFIGURE ── */}
        {step === 2 && preview && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

            {/* Preview */}
            <div style={card}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
                <Ic d={I.user} color="#1D9E75" size={22} />
                <h2 style={{ margin: 0, fontSize: 18, color: "#1F4E79" }}>Excel Loaded</h2>
                <div style={{ marginLeft: "auto", background: "#E1F5EE", color: "#04342C",
                  borderRadius: 99, padding: "3px 14px", fontWeight: 700, fontSize: 12 }}>
                  {preview.count} Patients
                </div>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "#1F4E79" }}>
                      {["Sr.", "MRN", "Patient Name"].map(h => (
                        <th key={h} style={th}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.preview.map((p, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? "#F7FAFC" : "#fff" }}>
                        <td style={td}>{i + 1}</td>
                        <td style={td}>{p.mrn}</td>
                        <td style={td}>{p.name}</td>
                      </tr>
                    ))}
                    {preview.count > 8 && (
                      <tr>
                        <td colSpan={3} style={{ ...td, textAlign: "center", color: "#999" }}>
                          … and {preview.count - 8} more patients
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Config */}
            <div style={card}>
              <h3 style={{ margin: "0 0 14px", color: "#1F4E79", fontSize: 16 }}>Report Settings</h3>
              <label style={{ fontSize: 13, fontWeight: 600, color: "#4A5568" }}>
                Month Label (shown in Excel header)
              </label>
              <input value={month} onChange={e => setMonth(e.target.value)}
                style={{
                  display: "block", width: "100%", marginTop: 7,
                  padding: "10px 13px", borderRadius: 8,
                  border: "1.5px solid #CBD5E0", fontSize: 14, outline: "none",
                  boxSizing: "border-box",
                }} />

              {/* Workflow summary */}
              <div style={{
                marginTop: 18, padding: "14px 16px", background: "#EBF8F4",
                borderRadius: 9, fontSize: 13, color: "#04342C", lineHeight: 1.7,
              }}>
                <strong>What will happen automatically:</strong><br />
                1. Login to hbs.chealth.pk (user: 7786)<br />
                2. FrontDesk → Final Settlement<br />
                3. Click <em>Patients</em> → search each patient by <strong>MRN</strong> (unique per patient)<br />
                4. Confirm the MRN appears on the matched card before opening it<br />
                5. Click Print → Unpaid → scrape the real Admission No from the invoice<br />
                6. If no Lab Investigation section exists → mark as "No Labs", skip to next patient<br />
                7. Extract Lab Investigation tests — duplicate tests are merged (Qty combined)<br />
                8. Generate Lab Bill Excel with one table per patient
              </div>

              <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
                <button onClick={() => setStep(1)} style={{
                  ...btn("#fff", "#4A5568"), border: "1.5px solid #CBD5E0",
                }}>← Back</button>
                <button onClick={handleStart} style={{
                  ...btn("linear-gradient(135deg,#1F4E79,#2E75B6)"), flex: 1,
                  justifyContent: "center", fontSize: 15,
                }}>
                  <Ic d={I.play} size={16} color="#fff" />
                  Start Automation — {preview.count} Patients
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── STEP 3: RUNNING ── */}
        {step === 3 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

            {/* Progress card */}
            <div style={card}>
              <div style={{ textAlign: "center", marginBottom: 24 }}>
                <div style={{
                  width: 60, height: 60, borderRadius: "50%", background: "#EBF4FF",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  margin: "0 auto 14px",
                  animation: "spin 2s linear infinite",
                }}>
                  <Ic d={I.refresh} size={30} color="#2E75B6" />
                </div>
                <h2 style={{ margin: "0 0 5px", color: "#1F4E79", fontSize: 20 }}>Processing...</h2>
                <p style={{ margin: 0, color: "#666", fontSize: 13 }}>
                  Keep this window open. The browser is running in the background.
                </p>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 6 }}>
                <span style={{ color: "#4A5568", fontWeight: 600 }}>
                  {job?.done || 0} / {job?.total || "?"} patients
                </span>
                <span style={{ color: "#2E75B6", fontWeight: 700 }}>{progress}%</span>
              </div>
              <div style={{ height: 12, background: "#E8EDF2", borderRadius: 99, overflow: "hidden", marginBottom: 16 }}>
                <div style={{
                  height: "100%", borderRadius: 99,
                  width: `${progress}%`,
                  background: "linear-gradient(90deg,#1F4E79,#1D9E75)",
                  transition: "width .6s ease",
                }} />
              </div>

              {job?.current && (
                <div style={{
                  padding: "9px 14px", background: "#F7FAFC", borderRadius: 8,
                  fontSize: 13, color: "#2D3748", display: "flex", alignItems: "center", gap: 8,
                }}>
                  <Ic d={I.user} size={15} color="#2E75B6" />
                  Searching: <strong>{job.current}</strong>
                </div>
              )}
            </div>

            {/* Live log */}
            <div style={card}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <Ic d={I.log} size={16} color="#4A5568" />
                <span style={{ fontWeight: 600, fontSize: 13, color: "#4A5568" }}>Live Log</span>
              </div>
              <div ref={logRef} style={{
                background: "#0F1C2A", borderRadius: 8, padding: "12px 14px",
                height: 200, overflowY: "auto", fontFamily: "monospace",
                fontSize: 11.5, color: "#A8D8EA", lineHeight: 1.7,
              }}>
                {logs.length === 0 && <span style={{ color: "#555" }}>Starting...</span>}
                {logs.map((l, i) => (
                  <div key={i} style={{
                    color: l.includes("❌") ? "#FF6B6B"
                         : l.includes("✅") ? "#A8F0C6"
                         : l.includes("⚠") ? "#FFD93D"
                         : "#A8D8EA"
                  }}>{l}</div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── STEP 4: DONE ── */}
        {step === 4 && job && (
          <div style={card}>
            <div style={{ textAlign: "center", marginBottom: 26 }}>
              <div style={{
                width: 68, height: 68, borderRadius: "50%", background: "#E1F5EE",
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 14px",
              }}>
                <Ic d={I.check} size={34} color="#1D9E75" />
              </div>
              <h2 style={{ margin: "0 0 5px", color: "#1F4E79", fontSize: 22 }}>Lab Bill Ready!</h2>
              <p style={{ margin: 0, color: "#666", fontSize: 14 }}>
                All {job.total} patients processed. Your Excel is ready.
              </p>
            </div>

            {/* Stats */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 22 }}>
              {[
                { label: "With Labs",   val: job.success, color: "#1D9E75" },
                { label: "No Labs",     val: job.no_labs, color: "#C05621" },
                { label: "Failed / Skip", val: job.failed,  color: "#C00000" },
              ].map((s, i) => (
                <div key={i} style={{
                  padding: 14, background: "#F7FAFC", borderRadius: 10, textAlign: "center",
                }}>
                  <div style={{ fontSize: 26, fontWeight: 700, color: s.color }}>{s.val}</div>
                  <div style={{ fontSize: 11, color: "#718096", marginTop: 3 }}>{s.label}</div>
                </div>
              ))}
            </div>

            <button onClick={() => window.open(`${API}/download`,"_blank")}
              style={{ ...btn("linear-gradient(135deg,#1D9E75,#2E75B6)"), width:"100%", justifyContent:"center", fontSize:15, padding:"14px" }}>
              <Ic d={I.dl} size={20} color="#fff" />
              Download Lab Bill Excel
            </button>

            <button onClick={() => { setStep(1); setPreview(null); setJob(null); setLogs([]); }}
              style={{ ...btn("#fff","#4A5568"), width:"100%", justifyContent:"center",
                marginTop:10, border:"1.5px solid #CBD5E0" }}>
              ← Process Another Month
            </button>
          </div>
        )}
      </div>

      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
