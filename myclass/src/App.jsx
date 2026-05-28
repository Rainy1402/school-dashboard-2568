import { useState, useEffect, useCallback, useRef } from "react"
import * as XLSX from "xlsx"
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts"
import {
  LayoutDashboard, School, Users, FileSpreadsheet, BookOpen,
  BarChart3, Settings, Archive, RotateCcw, Plus, X, Check,
  Clock, Upload, Trash2, RefreshCw, GraduationCap, AlertCircle,
  CheckCircle, WifiOff, Search, Save, ChevronDown,
  Star, CalendarDays, ClipboardList, Download
} from "lucide-react"

// ── IN-MEMORY STORE ───────────────────────────────────────────
const _store = {}
const store = { get:k=>_store[k]??null, set:(k,v)=>{_store[k]=v}, remove:k=>{delete _store[k]} }

// ── CUSTOM SUPABASE REST CLIENT ───────────────────────────────
function createClient(baseUrl, apiKey) {
  const url = baseUrl.replace(/\/$/, "")
  const h = { apikey:apiKey, Authorization:`Bearer ${apiKey}`, "Content-Type":"application/json" }

  function qb(table) {
    let _sel="*", _filters=[], _orders=[], _lim=null
    let _countMode=null, _isHead=false, _single=false

    const q = {
      select(fields="*", opts={}) {
        _sel = fields.replace(/\s+/g,"")
        if (opts.count) _countMode = opts.count
        if (opts.head)  _isHead  = opts.head
        return q
      },
      eq(col,val)    { _filters.push({col,op:"eq",val:String(val)}); return q },
      neq(col,val)   { _filters.push({col,op:"neq",val:String(val)}); return q },
      ilike(col,pat) { _filters.push({col,op:"ilike",val:pat}); return q },
      order(col,opts={}) { _orders.push(`${col}.${opts.ascending===false?"desc":"asc"}`); return q },
      limit(n)  { _lim=n; return q },
      maybeSingle() { _single=true; _lim=1; return q },
      async then(resolve, reject) {
        try {
          let ep = `${url}/rest/v1/${table}?select=${_sel}`
          _filters.forEach(f => { ep += `&${f.col}=${f.op}.${encodeURIComponent(f.val)}` })
          if (_orders.length) ep += `&order=${_orders.join(",")}`
          if (_lim) ep += `&limit=${_lim}`
          const rh = {...h}
          if (_countMode==="exact") rh["Prefer"] = "count=exact"
          const method = _isHead ? "HEAD" : "GET"
          const res = await fetch(ep, { method, headers:rh })
          if (_isHead) {
            const rng = res.headers.get("Content-Range") || "*/0"
            const parts = rng.split("/")
            resolve({ count: parseInt(parts[1]||"0")||0, data:null, error:null })
            return
          }
          const body = await res.json().catch(()=>[])
          if (!res.ok) { resolve({ data:null, error:body, count:null }); return }
          let count = null
          if (_countMode==="exact") {
            const rng = res.headers.get("Content-Range")||""
            if (rng.includes("/")) count = parseInt(rng.split("/")[1])||0
          }
          const dataOut = Array.isArray(body) ? body : (body ? [body] : [])
          if (_single) resolve({ data: dataOut[0]??null, error:null })
          else resolve({ data:dataOut, error:null, count })
        } catch(e) { reject(e) }
      }
    }
    return q
  }

  return {
    from(table) {
      return {
        select(fields="*",opts={}) { return qb(table).select(fields,opts) },
        async insert(data) {
          try {
            const res = await fetch(`${url}/rest/v1/${table}`, {
              method:"POST",
              headers:{...h, Prefer:"return=minimal"},
              body:JSON.stringify(data)
            })
            if (res.status===201||res.status===204) return { data:null, error:null }
            const body = await res.json().catch(()=>({}))
            return res.ok ? { data:body, error:null } : { data:null, error:body }
          } catch(e) { return { data:null, error:{message:e.message} } }
        },
        update(data) {
          const _f = []
          const u = {
            eq(col,val) { _f.push(`${col}=eq.${encodeURIComponent(val)}`); return u },
            async then(resolve) {
              try {
                const ep = `${url}/rest/v1/${table}${_f.length?"?"+_f.join("&"):""}`
                const res = await fetch(ep, {
                  method:"PATCH",
                  headers:{...h, Prefer:"return=minimal"},
                  body:JSON.stringify(data)
                })
                if (res.status===204||res.status===200) { resolve({ data:null, error:null }); return }
                const body = await res.json().catch(()=>({}))
                resolve(res.ok ? { data:body, error:null } : { data:null, error:body })
              } catch(e) { resolve({ data:null, error:{message:e.message} }) }
            }
          }
          return u
        },
        delete() {
          const _f = []
          const d = {
            eq(col,val) { _f.push(`${col}=eq.${encodeURIComponent(val)}`); return d },
            async then(resolve) {
              if (!_f.length) { resolve({ error:{message:"ต้องระบุ filter ก่อนลบ"} }); return }
              try {
                const res = await fetch(`${url}/rest/v1/${table}?${_f.join("&")}`, {
                  method:"DELETE", headers:h
                })
                if (res.ok) { resolve({ error:null }); return }
                const body = await res.json().catch(()=>({}))
                resolve({ error:body })
              } catch(e) { resolve({ error:{message:e.message} }) }
            }
          }
          return d
        }
      }
    }
  }
}

// ── THEME & STYLES ────────────────────────────────────────────
const T = {
  purple:"#7C3AED", purpleL:"#EDE9FE",
  blue:"#2563EB",   blueL:"#DBEAFE",
  green:"#16A34A",  greenL:"#DCFCE7",
  orange:"#EA580C", orangeL:"#FFEDD5",
  red:"#DC2626",    redL:"#FEE2E2",
  yellow:"#D97706", yellowL:"#FEF3C7",
  slate:"#F8FAFC",  border:"#E2E8F0",
  text:"#1E293B",   muted:"#94A3B8",
}
const CLASS_COLORS = ["#7C3AED","#2563EB","#16A34A","#D97706","#DC2626","#0891B2","#BE185D","#0F766E"]
const ss = {
  app:    { position:"relative", height:"100vh", overflow:"hidden", display:"flex", fontFamily:"'Prompt','Sarabun',sans-serif", background:T.slate },
  sidebar:{ width:220, minWidth:220, background:"#fff", borderRight:`1px solid ${T.border}`, display:"flex", flexDirection:"column", overflowY:"auto" },
  main:   { flex:1, display:"flex", flexDirection:"column", overflow:"hidden" },
  topbar: { background:"#fff", borderBottom:`1px solid ${T.border}`, padding:"0 24px", display:"flex", alignItems:"center", justifyContent:"space-between", height:58, flexShrink:0 },
  content:{ flex:1, overflowY:"auto", padding:"24px" },
  overlay:{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", zIndex:60, display:"flex", alignItems:"center", justifyContent:"center", padding:16 },
  nav:    { flex:1, padding:"8px" },
  navI:   (a,d)=>({ display:"flex", alignItems:"center", gap:10, padding:"9px 10px", borderRadius:10, cursor:"pointer", fontSize:13.5, fontWeight:a?600:500, marginBottom:2, background:a?T.purple:"transparent", color:a?"#fff":d?T.red:"#475569", transition:"all 0.15s" }),
  addBtn: { margin:"12px 10px 16px", padding:"10px 0", background:T.purple, color:"#fff", border:"none", borderRadius:10, cursor:"pointer", fontFamily:"inherit", fontSize:13.5, fontWeight:600, display:"flex", alignItems:"center", justifyContent:"center", gap:6 },
  card:   { background:"#fff", borderRadius:14, border:`1px solid ${T.border}`, overflow:"hidden" },
  cHead:  { padding:"14px 18px 12px", display:"flex", alignItems:"center", justifyContent:"space-between" },
  cTitle: { fontSize:14.5, fontWeight:700, color:T.text },
  table:  { width:"100%", borderCollapse:"collapse" },
  th:     { padding:"8px 14px", fontSize:12.5, color:T.muted, fontWeight:600, textAlign:"left", borderBottom:`1px solid #F1F5F9`, whiteSpace:"nowrap" },
  td:     { padding:"10px 14px", fontSize:13.5, color:"#334155", borderBottom:`1px solid #F8FAFC` },
  input:  { width:"100%", border:`1px solid ${T.border}`, borderRadius:10, padding:"9px 12px", fontSize:13.5, fontFamily:"inherit", outline:"none", boxSizing:"border-box" },
  label:  { display:"block", fontSize:12.5, fontWeight:600, color:"#64748B", marginBottom:5 },
  fRow:   { marginBottom:14 },
  textarea:{ width:"100%", border:`1px solid ${T.border}`, borderRadius:10, padding:"9px 12px", fontSize:13.5, fontFamily:"inherit", outline:"none", resize:"vertical", minHeight:72, boxSizing:"border-box" },
  btnP:   (c=T.purple)=>({ padding:"9px 18px", background:c, color:"#fff", border:"none", borderRadius:10, cursor:"pointer", fontFamily:"inherit", fontSize:13.5, fontWeight:600, display:"inline-flex", alignItems:"center", gap:6 }),
  btnO:   { padding:"9px 18px", background:"#fff", color:"#64748B", border:`1px solid ${T.border}`, borderRadius:10, cursor:"pointer", fontFamily:"inherit", fontSize:13.5, fontWeight:600, display:"inline-flex", alignItems:"center", gap:6 },
  pill:   (c,bg)=>({ padding:"3px 10px", borderRadius:99, fontSize:12, fontWeight:600, background:bg, color:c, display:"inline-flex", alignItems:"center", gap:4 }),
  modal:  { background:"#fff", borderRadius:20, width:500, maxWidth:"calc(100vw - 32px)", maxHeight:"90vh", display:"flex", flexDirection:"column", overflow:"hidden", boxShadow:"0 24px 60px rgba(0,0,0,0.25)" },
  mHead:  { padding:"16px 20px", borderBottom:`1px solid ${T.border}`, display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 },
  mBody:  { padding:"20px", overflowY:"auto", flex:1 },
}

// ── MICRO COMPONENTS ──────────────────────────────────────────
function Spin({ size=18, color=T.purple }) {
  return <div style={{ width:size, height:size, border:`2.5px solid #E2E8F0`, borderTopColor:color, borderRadius:"50%", animation:"mcSpin 0.7s linear infinite", flexShrink:0 }}/>
}
function PageLoad({ text="กำลังโหลด..." }) {
  return <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", height:280, gap:12 }}>
    <Spin size={34}/><span style={{ color:T.muted, fontSize:14 }}>{text}</span>
  </div>
}
function EmptyState({ icon="📭", msg="ไม่มีข้อมูล", btn }) {
  return <div style={{ textAlign:"center", padding:"52px 24px" }}>
    <div style={{ fontSize:48, marginBottom:12 }}>{icon}</div>
    <div style={{ fontSize:15, fontWeight:600, color:T.muted }}>{msg}</div>
    {btn && <div style={{ marginTop:16 }}>{btn}</div>}
  </div>
}
function Toast({ list }) {
  if (!list.length) return null
  return <div style={{ position:"fixed", top:12, right:12, zIndex:300, display:"flex", flexDirection:"column", gap:8 }}>
    {list.map(t => (
      <div key={t.id} style={{ background:t.type==="err"?T.red:T.green, color:"#fff", borderRadius:12, padding:"11px 16px", fontSize:13.5, fontWeight:600, display:"flex", alignItems:"center", gap:8, fontFamily:"inherit", boxShadow:"0 4px 20px rgba(0,0,0,0.2)", maxWidth:320, animation:"mcSlide 0.2s ease" }}>
        {t.type==="err" ? <AlertCircle size={15}/> : <CheckCircle size={15}/>}{t.msg}
      </div>
    ))}
  </div>
}
function useToast() {
  const [list, setList] = useState([])
  const add = useCallback((msg, type="ok") => {
    const id = Date.now()
    setList(p => [...p, { id, msg, type }])
    setTimeout(() => setList(p => p.filter(t => t.id !== id)), 3200)
  }, [])
  return { list, ok:m=>add(m,"ok"), err:m=>add(m,"err") }
}

// ── ERROR BOUNDARY ────────────────────────────────────────────
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError:false, error:null } }
  static getDerivedStateFromError(e) { return { hasError:true, error:e } }
  render() {
    if (!this.state.hasError) return this.props.children
    return <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", background:"#FEF2F2", fontFamily:"'Prompt','Sarabun',sans-serif" }}>
      <div style={{ background:"#fff", borderRadius:16, padding:"32px 36px", maxWidth:460, textAlign:"center", boxShadow:"0 8px 32px rgba(0,0,0,0.1)" }}>
        <div style={{ fontSize:48, marginBottom:12 }}>⚠️</div>
        <div style={{ fontSize:18, fontWeight:700, color:T.text, marginBottom:8 }}>เกิดข้อผิดพลาด</div>
        <div style={{ fontSize:13, color:T.muted, marginBottom:20, wordBreak:"break-word" }}>{this.state.error?.message}</div>
        <button onClick={()=>window.location.reload()} style={{ ...ss.btnP(), width:"100%", justifyContent:"center" }}>
          🔄 โหลดหน้าใหม่
        </button>
      </div>
    </div>
  }
}

// ── DATA HOOKS ────────────────────────────────────────────────
function useQ(fn, deps=[]) {
  const [data, setData]       = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError]     = useState(null)
  const reload = useCallback(async () => {
    setLoading(true); setError(null)
    try { setData(await fn()) }
    catch(e) { setError(e.message); setData(null) }
    finally { setLoading(false) }
  }, deps)
  useEffect(() => { reload() }, [reload])
  return { data, loading, error, reload }
}

function useDash(sb) {
  return useQ(async () => {
    const [sc, cc, sesc, att] = await Promise.all([
      sb.from("students").select("id", { count:"exact", head:true }),
      sb.from("classes").select("id", { count:"exact", head:true }),
      sb.from("teaching_sessions").select("id", { count:"exact", head:true }),
      sb.from("attendance").select("status"),
    ])
    const a = att.data || []
    return {
      students: sc.count || 0,
      classes:  cc.count || 0,
      sessions: sesc.count || 0,
      att: {
        present: a.filter(x=>x.status==="present").length,
        absent:  a.filter(x=>x.status==="absent").length,
        late:    a.filter(x=>x.status==="late").length,
        leave:   a.filter(x=>x.status==="leave").length,
        total:   a.length
      }
    }
  }, [sb])
}

function useRecentSessions(sb) {
  return useQ(async () => {
    const { data, error } = await sb.from("teaching_sessions")
      .select("*,classes(class_name,subject_name,color)")
      .order("created_at", { ascending:false }).limit(6)
    if (error) throw new Error(error.message || JSON.stringify(error))
    return data || []
  }, [sb])
}

function useClasses(sb) {
  return useQ(async () => {
    // Get classes + all their student links (count client-side)
    const { data, error } = await sb.from("classes")
      .select("*,class_students(*)")
      .order("created_at", { ascending:false })
    if (error) throw new Error(error.message || JSON.stringify(error))
    return (data || []).map(c => ({
      ...c,
      _studentCount: Array.isArray(c.class_students) ? c.class_students.length : 0
    }))
  }, [sb])
}

function useStudents(sb, search, level) {
  return useQ(async () => {
    let q = sb.from("students").select("*")
      .order("level").order("room").order("student_no")
    if (level && level !== "ทั้งหมด") q = q.eq("level", level)
    if (search && search.trim()) q = q.ilike("full_name", `%${search.trim()}%`)
    const { data, error } = await q
    if (error) throw new Error(error.message || JSON.stringify(error))
    return data || []
  }, [sb, search, level])
}

function useAllSessions(sb) {
  return useQ(async () => {
    const { data, error } = await sb.from("teaching_sessions")
      .select("*,classes(class_name,subject_name,color)")
      .order("created_at", { ascending:false })
    if (error) throw new Error(error.message || JSON.stringify(error))
    return data || []
  }, [sb])
}

function useProfile(sb) {
  return useQ(async () => {
    const { data } = await sb.from("profiles").select("*").limit(1).maybeSingle()
    return data
  }, [sb])
}

// ── SETUP SCREEN ──────────────────────────────────────────────
function SetupScreen({ onReady }) {
  const [url,     setUrl]     = useState(store.get("mc_url") || "")
  const [key,     setKey]     = useState(store.get("mc_key") || "")
  const [testing, setTesting] = useState(false)
  const [err,     setErr]     = useState("")

  const connect = async () => {
    if (!url.trim() || !key.trim()) { setErr("กรุณากรอกทั้ง URL และ Key"); return }
    setTesting(true); setErr("")
    try {
      const client = createClient(url.trim(), key.trim())
      const { error } = await client.from("profiles").select("id").limit(1)
      if (error && error.code !== "PGRST116") {
        throw new Error(error.message || error.hint || JSON.stringify(error))
      }
      store.set("mc_url", url.trim()); store.set("mc_key", key.trim())
      onReady(url.trim(), key.trim())
    } catch(e) { setErr("เชื่อมต่อไม่ได้: " + e.message) }
    finally { setTesting(false) }
  }

  return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", background:"linear-gradient(135deg,#1E1B4B 0%,#312E81 50%,#1E3A5F 100%)", fontFamily:"'Prompt','Sarabun',sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');@keyframes mcSpin{to{transform:rotate(360deg)}}*{box-sizing:border-box}body{margin:0}`}</style>
      <div style={{ background:"#fff", borderRadius:24, padding:"36px 40px", width:460, maxWidth:"calc(100vw - 32px)", boxShadow:"0 32px 80px rgba(0,0,0,0.3)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:28 }}>
          <div style={{ width:52, height:52, borderRadius:16, background:T.purple, display:"flex", alignItems:"center", justifyContent:"center" }}>
            <GraduationCap size={26} color="#fff"/>
          </div>
          <div>
            <div style={{ fontSize:20, fontWeight:700, color:T.text }}>MyClass 🍎</div>
            <div style={{ fontSize:13, color:T.muted }}>เชื่อมต่อ Supabase เพื่อเริ่มใช้งาน</div>
          </div>
        </div>
        <div style={{ background:"#F0F9FF", borderRadius:12, padding:"12px 16px", marginBottom:22, border:"1px solid #BAE6FD" }}>
          <div style={{ fontSize:13, color:"#0369A1", lineHeight:1.7 }}>
            <b>วิธีรับค่า:</b> ไปที่ <b>supabase.com</b> → เลือก project<br/>
            → <b>Settings → API</b> → คัดลอก URL และ anon key
          </div>
        </div>
        <div style={ss.fRow}>
          <label style={ss.label}>Supabase Project URL</label>
          <input value={url} onChange={e=>setUrl(e.target.value)} style={ss.input} placeholder="https://xxxx.supabase.co" type="url" onKeyDown={e=>e.key==="Enter"&&connect()}/>
        </div>
        <div style={ss.fRow}>
          <label style={ss.label}>Anon Public Key</label>
          <input value={key} onChange={e=>setKey(e.target.value)} style={{ ...ss.input, fontFamily:"monospace", fontSize:12 }} placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..." onKeyDown={e=>e.key==="Enter"&&connect()}/>
        </div>
        {err && (
          <div style={{ background:T.redL, borderRadius:10, padding:"10px 14px", fontSize:13, color:T.red, marginBottom:16, display:"flex", gap:8, alignItems:"flex-start" }}>
            <AlertCircle size={15} style={{ flexShrink:0, marginTop:1 }}/>{err}
          </div>
        )}
        <button onClick={connect} disabled={testing} style={{ ...ss.btnP(), width:"100%", justifyContent:"center", padding:"13px 0", fontSize:14.5, opacity:testing?0.75:1 }}>
          {testing ? <><Spin size={17} color="#fff"/>กำลังเชื่อมต่อ...</> : <><CheckCircle size={17}/>เชื่อมต่อและเริ่มใช้งาน</>}
        </button>
      </div>
    </div>
  )
}

// ── ATTENDANCE CHART ──────────────────────────────────────────
const PIE_LABEL = ({ cx,cy,midAngle,innerRadius,outerRadius,percent }) => {
  if (percent < 0.04) return null
  const R = Math.PI/180, r = innerRadius + (outerRadius-innerRadius)*0.55
  return <text x={cx+r*Math.cos(-midAngle*R)} y={cy+r*Math.sin(-midAngle*R)} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={11} fontWeight={700}>{`${(percent*100).toFixed(0)}%`}</text>
}
function AttChart({ att }) {
  const items = [
    { name:"มาเรียน",  value:att.present, color:T.green  },
    { name:"ขาดเรียน", value:att.absent,  color:T.red    },
    { name:"สาย",      value:att.late,    color:T.yellow },
    { name:"ลา",       value:att.leave,   color:T.blue   },
  ].filter(d => d.value > 0)
  if (!att.total) return <div style={{ padding:"20px 18px", color:"#CBD5E1", fontSize:13, textAlign:"center" }}>ยังไม่มีข้อมูลการเช็คชื่อ</div>
  return (
    <div style={{ display:"flex", alignItems:"center", gap:12, padding:"10px 18px 18px" }}>
      <div style={{ width:160, height:160, flexShrink:0 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie data={items} cx="50%" cy="50%" innerRadius={42} outerRadius={70} dataKey="value" labelLine={false} label={PIE_LABEL}>
              {items.map((d,i) => <Cell key={i} fill={d.color}/>)}
            </Pie>
            <Tooltip contentStyle={{ fontFamily:"Prompt,sans-serif", fontSize:13, borderRadius:10 }}/>
          </PieChart>
        </ResponsiveContainer>
      </div>
      <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
        {items.map((d,i) => (
          <div key={i} style={{ display:"flex", alignItems:"center", gap:7, fontSize:13 }}>
            <span style={{ width:10, height:10, borderRadius:"50%", background:d.color, flexShrink:0 }}/>
            <span style={{ color:"#334155", fontWeight:500 }}>{d.name}</span>
            <span style={{ color:T.muted, fontSize:12 }}>{d.value}คน ({Math.round(d.value/att.total*100)}%)</span>
          </div>
        ))}
      </div>
    </div>
  )
}

// ── MODAL SHELL ───────────────────────────────────────────────
function ModalShell({ title, sub, onClose, footer, children, wide }) {
  // Close on Escape key
  useEffect(() => {
    const handler = e => { if (e.key === "Escape") onClose() }
    document.addEventListener("keydown", handler)
    return () => document.removeEventListener("keydown", handler)
  }, [onClose])

  return (
    <div style={ss.overlay} onClick={onClose}>
      <div style={{ ...ss.modal, width:wide||480 }} onClick={e => e.stopPropagation()}>
        <div style={ss.mHead}>
          <div>
            <div style={{ fontSize:15, fontWeight:700, color:T.text }}>{title}</div>
            {sub && <div style={{ fontSize:12, color:T.muted }}>{sub}</div>}
          </div>
          <button onClick={onClose} style={{ border:"none", background:"none", cursor:"pointer", padding:4, borderRadius:8 }}>
            <X size={18} color={T.muted}/>
          </button>
        </div>
        <div style={ss.mBody}>{children}</div>
        {footer && <div style={{ padding:"14px 20px", borderTop:`1px solid ${T.border}`, display:"flex", gap:10, flexShrink:0 }}>{footer}</div>}
      </div>
    </div>
  )
}

// ── ADD CLASS MODAL ───────────────────────────────────────────
function AddClassModal({ sb, onClose, onDone, toast }) {
  const [form,   setForm]   = useState({ class_name:"", subject_name:"", color:CLASS_COLORS[0] })
  const [saving, setSaving] = useState(false)
  const f = (k,v) => setForm(p => ({ ...p, [k]:v }))

  const save = async () => {
    if (!form.class_name.trim() || !form.subject_name.trim()) { toast.err("กรุณากรอกข้อมูลให้ครบ"); return }
    setSaving(true)
    try {
      const { error } = await sb.from("classes").insert({ class_name:form.class_name.trim(), subject_name:form.subject_name.trim(), color:form.color })
      if (error) throw new Error(error.message || JSON.stringify(error))
      toast.ok(`เพิ่มชั้นเรียน ${form.class_name} แล้ว`); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="เพิ่มชั้นเรียน" onClose={onClose}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={save} disabled={saving} style={{ ...ss.btnP(), flex:2, justifyContent:"center", opacity:saving?0.7:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังบันทึก...</> : <><Check size={15}/>เพิ่มชั้นเรียน</>}
        </button>
      </>}>
      <div style={ss.fRow}><label style={ss.label}>ชื่อห้อง *</label><input value={form.class_name} onChange={e=>f("class_name",e.target.value)} style={ss.input} placeholder="ม.3/1" autoFocus/></div>
      <div style={ss.fRow}><label style={ss.label}>วิชา *</label><input value={form.subject_name} onChange={e=>f("subject_name",e.target.value)} style={ss.input} placeholder="คณิตศาสตร์"/></div>
      <div style={ss.fRow}>
        <label style={ss.label}>สีประจำห้อง</label>
        <div style={{ display:"flex", gap:9, marginTop:5 }}>
          {CLASS_COLORS.map(c => <button key={c} onClick={()=>f("color",c)} style={{ width:30, height:30, borderRadius:"50%", background:c, border:form.color===c?"3px solid #1E293B":"3px solid transparent", cursor:"pointer", transition:"all 0.15s" }}/>)}
        </div>
      </div>
    </ModalShell>
  )
}

// ── ADD STUDENT MODAL ─────────────────────────────────────────
function AddStudentModal({ sb, onClose, onDone, toast }) {
  const [form,   setForm]   = useState({ student_no:"", full_name:"", level:"", room:"" })
  const [saving, setSaving] = useState(false)
  const f = (k,v) => setForm(p => ({ ...p, [k]:v }))

  const save = async () => {
    if (!form.full_name.trim()) { toast.err("กรุณากรอกชื่อ"); return }
    setSaving(true)
    try {
      const payload = {
        full_name:  form.full_name.trim(),
        level:      form.level.trim(),
        room:       form.room.trim(),
        student_no: form.student_no ? parseInt(form.student_no) : null,
      }
      const { error } = await sb.from("students").insert(payload)
      if (error) throw new Error(error.message || JSON.stringify(error))
      toast.ok(`เพิ่ม ${form.full_name} แล้ว`); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="เพิ่มนักเรียน" onClose={onClose}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={save} disabled={saving} style={{ ...ss.btnP(), flex:2, justifyContent:"center", opacity:saving?0.7:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังบันทึก...</> : <><Plus size={15}/>เพิ่มนักเรียน</>}
        </button>
      </>}>
      <div style={{ display:"grid", gridTemplateColumns:"80px 1fr", gap:12 }}>
        <div style={ss.fRow}><label style={ss.label}>เลขที่</label><input type="number" min="1" value={form.student_no} onChange={e=>f("student_no",e.target.value)} style={ss.input} placeholder="1"/></div>
        <div style={ss.fRow}><label style={ss.label}>ชื่อ-นามสกุล *</label><input value={form.full_name} onChange={e=>f("full_name",e.target.value)} style={ss.input} placeholder="ชื่อ นามสกุล" autoFocus/></div>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        <div style={ss.fRow}><label style={ss.label}>ชั้น</label><input value={form.level} onChange={e=>f("level",e.target.value)} style={ss.input} placeholder="ม.3"/></div>
        <div style={ss.fRow}><label style={ss.label}>ห้อง</label><input value={form.room} onChange={e=>f("room",e.target.value)} style={ss.input} placeholder="3/1"/></div>
      </div>
    </ModalShell>
  )
}

// ── PASTE MODAL ───────────────────────────────────────────────
function PasteModal({ sb, onClose, onDone, toast }) {
  const [text,   setText]   = useState("")
  const [saving, setSaving] = useState(false)

  const parsed = (() => {
    if (!text.trim()) return []
    try {
      return text.trim().split("\n")
        .map((line, i) => {
          if (!line.trim()) return null
          try {
            const p = line.trim().split(/\t|,/)
            const firstTrimmed = (p[0] || "").trim()
            const isNum = firstTrimmed !== "" && !isNaN(firstTrimmed) && !firstTrimmed.startsWith("เด็ก") && !firstTrimmed.startsWith("นาย") && !firstTrimmed.startsWith("นาง") && !firstTrimmed.startsWith("ด.")
            return {
              student_no: isNum ? parseInt(firstTrimmed) : i + 1,
              full_name:  (isNum ? (p[1]||"") : (p[0]||"")).trim(),
              level:      (isNum ? (p[2]||"") : (p[1]||"")).trim(),
              room:       (isNum ? (p[3]||"") : (p[2]||"")).trim(),
            }
          } catch { return null }
        })
        .filter(s => s && s.full_name && s.full_name.trim().length > 0)
    } catch { return [] }
  })()

  const save = async () => {
    if (!parsed.length) return
    setSaving(true)
    try {
      const { error } = await sb.from("students").insert(parsed)
      if (error) throw new Error(error.message || JSON.stringify(error))
      toast.ok(`นำเข้า ${parsed.length} คนแล้ว`); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="วางรายชื่อนักเรียน" onClose={onClose}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={save} disabled={!parsed.length||saving} style={{ ...ss.btnP(), flex:2, justifyContent:"center", opacity:(!parsed.length||saving)?0.5:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังนำเข้า...</> : <><Upload size={15}/>นำเข้า {parsed.length} คน</>}
        </button>
      </>}>
      <div style={{ background:T.yellowL, borderRadius:10, padding:"10px 14px", marginBottom:14, fontSize:12.5, color:"#78350F", lineHeight:1.7 }}>
        <b>รูปแบบ (แต่ละบรรทัด = 1 คน):</b><br/>
        วางชื่อเดี่ยว: <code>เด็กหญิงกมลชนก วิชัย</code><br/>
        หรือมีเลขที่: <code>1 [Tab] ชื่อ [Tab] ม.3 [Tab] 3/1</code>
      </div>
      <textarea
        value={text}
        onChange={e => setText(e.target.value)}
        rows={8}
        placeholder={"เด็กหญิงกมลชนก วิชัย\nเด็กชายขวัญฤดี มีสุข\n\nหรือ\n1\tเด็กหญิงกมลชนก วิชัย\tม.3\t3/1"}
        style={{ ...ss.textarea, fontFamily:"monospace", fontSize:12 }}
      />
      {parsed.length > 0 && (
        <div style={{ fontSize:13, color:T.green, fontWeight:600, marginTop:6, display:"flex", gap:6, alignItems:"center" }}>
          <CheckCircle size={14}/>พบ {parsed.length} คน — พร้อมนำเข้า
        </div>
      )}
    </ModalShell>
  )
}

// ── IMPORT EXCEL MODAL ────────────────────────────────────────
function ImportExcelModal({ sb, onClose, onDone, toast }) {
  const [dragging, setDragging] = useState(false)
  const [preview,  setPreview]  = useState([])
  const [saving,   setSaving]   = useState(false)
  const ref = useRef()

  const parse = file => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb   = XLSX.read(e.target.result, { type:"array" })
        const ws   = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" })
        const result = rows.slice(1)
          .filter(r => r[1] && String(r[1]).trim())
          .map(r => ({
            student_no: parseInt(r[0]) || null,
            full_name:  String(r[1] || "").trim(),
            level:      String(r[2] || "").trim(),
            room:       String(r[3] || "").trim(),
          }))
        if (result.length === 0) { toast.err("ไม่พบข้อมูล ตรวจสอบรูปแบบไฟล์"); return }
        setPreview(result)
      } catch(err) { toast.err("อ่านไฟล์ไม่ได้: " + err.message) }
    }
    reader.onerror = () => toast.err("อ่านไฟล์ไม่ได้")
    reader.readAsArrayBuffer(file)
  }

  const doImport = async () => {
    setSaving(true)
    try {
      const { error } = await sb.from("students").insert(preview)
      if (error) throw new Error(error.message || JSON.stringify(error))
      toast.ok(`นำเข้าสำเร็จ ${preview.length} คน`); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="นำเข้าจาก Excel" onClose={onClose}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={doImport} disabled={!preview.length||saving} style={{ ...ss.btnP(T.green), flex:2, justifyContent:"center", opacity:(!preview.length||saving)?0.5:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังนำเข้า...</> : <><Upload size={15}/>นำเข้า {preview.length} คน</>}
        </button>
      </>}>
      <div
        onDragEnter={()=>setDragging(true)} onDragLeave={()=>setDragging(false)}
        onDrop={e=>{e.preventDefault();setDragging(false);parse(e.dataTransfer.files[0])}}
        onDragOver={e=>e.preventDefault()}
        onClick={()=>ref.current?.click()}
        style={{ border:`2px dashed ${dragging?T.purple:"#CBD5E1"}`, borderRadius:14, padding:"32px 20px", textAlign:"center", marginBottom:14, background:dragging?T.purpleL:"#F8FAFC", cursor:"pointer", transition:"all 0.15s" }}>
        <FileSpreadsheet size={40} color={T.green} style={{ margin:"0 auto 10px" }}/>
        <div style={{ fontSize:14, fontWeight:600, color:"#334155" }}>ลากไฟล์ Excel มาวาง</div>
        <div style={{ fontSize:12.5, color:T.muted, marginTop:4 }}>หรือคลิกเพื่อเลือก .xlsx, .xls</div>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }} onChange={e=>e.target.files[0]&&parse(e.target.files[0])}/>
      </div>
      <div style={{ background:"#F8FAFC", borderRadius:10, padding:"10px 14px", fontSize:12.5, color:"#64748B", marginBottom:preview.length?12:0 }}>
        <b>รูปแบบ column:</b> A=เลขที่ | B=ชื่อ-นามสกุล | C=ชั้น | D=ห้อง (row 1 เป็น header)
      </div>
      {preview.length > 0 && (
        <div>
          <div style={{ fontSize:13, color:T.green, fontWeight:600, marginBottom:6, display:"flex", gap:6, alignItems:"center" }}>
            <CheckCircle size={14}/>พบ {preview.length} คน
          </div>
          <div style={{ maxHeight:160, overflowY:"auto", border:`1px solid ${T.border}`, borderRadius:10 }}>
            {preview.slice(0, 8).map((s,i) => (
              <div key={i} style={{ padding:"6px 12px", fontSize:12.5, borderBottom:`1px solid #F8FAFC`, color:"#475569" }}>
                {s.student_no}. {s.full_name} <span style={{ color:T.muted }}>({s.level} {s.room})</span>
              </div>
            ))}
            {preview.length > 8 && <div style={{ padding:"6px 12px", fontSize:12, color:T.muted }}>...และอีก {preview.length-8} คน</div>}
          </div>
        </div>
      )}
    </ModalShell>
  )
}

// ── SESSION MODAL ─────────────────────────────────────────────
function SessionModal({ sb, onClose, onDone, classes, toast }) {
  const [tab, setTab] = useState("a")
  const [cls, setCls] = useState(classes[0]?.id || "")
  const [f, setF] = useState({
    teach_date:   new Date().toISOString().split("T")[0],
    topic:"", objective:"", activities:"",
    highlights:"", problems:"", improvements:"", reflection:""
  })
  const sf = (k,v) => setF(p => ({ ...p, [k]:v }))
  const [saving, setSaving] = useState(false)

  const save = async () => {
    if (!cls) { toast.err("เลือกชั้นเรียนก่อน"); return }
    if (!f.teach_date) { toast.err("กรุณาระบุวันที่สอน"); return }
    setSaving(true)
    try {
      const { error } = await sb.from("teaching_sessions").insert({ class_id:cls, ...f, status:"saved" })
      if (error) throw new Error(error.message || JSON.stringify(error))
      toast.ok("บันทึกการสอนแล้ว ✅"); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="บันทึกการสอน" onClose={onClose} wide={520}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={save} disabled={saving||!cls} style={{ ...ss.btnP(), flex:2, justifyContent:"center", opacity:(saving||!cls)?0.6:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังบันทึก...</> : <><Save size={15}/>บันทึก</>}
        </button>
      </>}>
      <div style={{ display:"flex", borderBottom:`1px solid ${T.border}`, marginBottom:16 }}>
        {[{id:"a",l:"ข้อมูลหลัก"},{id:"b",l:"บันทึกหลังสอน"}].map(t => (
          <button key={t.id} onClick={()=>setTab(t.id)} style={{ flex:1, padding:"9px 0", border:"none", background:"none", cursor:"pointer", fontFamily:"inherit", fontSize:13, fontWeight:600, color:tab===t.id?T.purple:T.muted, borderBottom:tab===t.id?`2px solid ${T.purple}`:"2px solid transparent" }}>
            {t.l}
          </button>
        ))}
      </div>
      {tab === "a" && <>
        <div style={ss.fRow}>
          <label style={ss.label}>ชั้นเรียน *</label>
          <select value={cls} onChange={e=>setCls(e.target.value)} style={ss.input}>
            <option value="">-- เลือกชั้นเรียน --</option>
            {classes.map(c => <option key={c.id} value={c.id}>{c.class_name} · {c.subject_name}</option>)}
          </select>
        </div>
        {!classes.length && <div style={{ fontSize:13, color:T.orange, marginBottom:12 }}>⚠️ ยังไม่มีชั้นเรียน กรุณาเพิ่มชั้นเรียนก่อน</div>}
        <div style={ss.fRow}><label style={ss.label}>วันที่สอน *</label><input type="date" value={f.teach_date} onChange={e=>sf("teach_date",e.target.value)} style={ss.input}/></div>
        <div style={ss.fRow}><label style={ss.label}>เรื่องที่สอน</label><input value={f.topic} onChange={e=>sf("topic",e.target.value)} style={ss.input} placeholder="ระบบสมการเชิงเส้น 2 ตัวแปร"/></div>
        <div style={ss.fRow}><label style={ss.label}>จุดประสงค์การเรียนรู้</label><textarea value={f.objective} onChange={e=>sf("objective",e.target.value)} style={ss.textarea} placeholder="นักเรียนสามารถ..."/></div>
        <div style={ss.fRow}><label style={ss.label}>กิจกรรมการเรียนรู้</label><textarea value={f.activities} onChange={e=>sf("activities",e.target.value)} style={ss.textarea} placeholder="วิธีการสอน / กิจกรรม..."/></div>
      </>}
      {tab === "b" && <>
        <div style={ss.fRow}><label style={ss.label}>✨ จุดเด่นของคาบ</label><textarea value={f.highlights} onChange={e=>sf("highlights",e.target.value)} style={ss.textarea} placeholder="สิ่งที่น่าประทับใจ..."/></div>
        <div style={ss.fRow}><label style={ss.label}>⚠️ ปัญหาที่พบ</label><textarea value={f.problems} onChange={e=>sf("problems",e.target.value)} style={ss.textarea} placeholder="ปัญหาที่เกิดขึ้นในห้อง..."/></div>
        <div style={ss.fRow}><label style={ss.label}>🔧 สิ่งที่ต้องปรับปรุง</label><textarea value={f.improvements} onChange={e=>sf("improvements",e.target.value)} style={ss.textarea} placeholder="จะปรับอะไรในครั้งหน้า..."/></div>
        <div style={ss.fRow}><label style={ss.label}>💭 Reflection ของครู</label><textarea value={f.reflection} onChange={e=>sf("reflection",e.target.value)} style={{ ...ss.textarea, minHeight:90 }} placeholder="ความรู้สึกและสิ่งที่เรียนรู้จากคาบนี้..."/></div>
      </>}
    </ModalShell>
  )
}

// ── PROFILE MODAL ─────────────────────────────────────────────
function ProfileModal({ sb, onClose, profile, onDone, toast, onDisconnect }) {
  const [form,   setForm]   = useState({ teacher_name:profile?.teacher_name||"", school_name:profile?.school_name||"", subject_group:profile?.subject_group||"" })
  const [saving, setSaving] = useState(false)
  const sf = (k,v) => setForm(p => ({ ...p, [k]:v }))

  const save = async () => {
    setSaving(true)
    try {
      if (profile?.id) { await sb.from("profiles").update(form).eq("id", profile.id) }
      else { await sb.from("profiles").insert(form) }
      toast.ok("บันทึกข้อมูลแล้ว"); onDone(); onClose()
    } catch(e) { toast.err(e.message) } finally { setSaving(false) }
  }

  return (
    <ModalShell title="ตั้งค่า" onClose={onClose}
      footer={<>
        <button onClick={onClose} style={{ ...ss.btnO, flex:1, justifyContent:"center" }}>ยกเลิก</button>
        <button onClick={save} disabled={saving} style={{ ...ss.btnP(), flex:2, justifyContent:"center", opacity:saving?0.7:1 }}>
          {saving ? <><Spin size={15} color="#fff"/>กำลังบันทึก...</> : <><Save size={15}/>บันทึก</>}
        </button>
      </>}>
      <div style={{ fontSize:12, fontWeight:700, color:"#64748B", marginBottom:10, textTransform:"uppercase", letterSpacing:0.5 }}>ข้อมูลครู</div>
      <div style={ss.fRow}><label style={ss.label}>ชื่อ-นามสกุลครู</label><input value={form.teacher_name} onChange={e=>sf("teacher_name",e.target.value)} style={ss.input} placeholder="ครูจิราภรณ์ รักเด็ก" autoFocus/></div>
      <div style={ss.fRow}><label style={ss.label}>โรงเรียน</label><input value={form.school_name} onChange={e=>sf("school_name",e.target.value)} style={ss.input} placeholder="โรงเรียนบ้านดอน"/></div>
      <div style={{ ...ss.fRow, marginBottom:20 }}><label style={ss.label}>กลุ่มสาระ</label><input value={form.subject_group} onChange={e=>sf("subject_group",e.target.value)} style={ss.input} placeholder="คณิตศาสตร์"/></div>
      <div style={{ borderTop:`1px solid ${T.border}`, paddingTop:16 }}>
        <div style={{ background:"#F8FAFC", borderRadius:10, padding:"10px 14px", marginBottom:10, fontSize:12.5 }}>
          <span style={{ color:T.muted }}>ฐานข้อมูล: </span>
          <span style={{ fontFamily:"monospace", fontSize:11, color:"#475569", wordBreak:"break-all" }}>
            {(store.get("mc_url")||"").replace("https://","").slice(0,40)}...
          </span>
        </div>
        <button onClick={onDisconnect} style={{ ...ss.btnO, width:"100%", justifyContent:"center", color:T.red, borderColor:"#FECACA" }}>
          <WifiOff size={15}/>เปลี่ยน Supabase Project
        </button>
      </div>
    </ModalShell>
  )
}

// ── MENU ──────────────────────────────────────────────────────
const MENU = [
  { id:"dashboard", icon:LayoutDashboard, label:"หน้าหลัก" },
  { id:"classes",   icon:School,          label:"ชั้นเรียนของฉัน" },
  { id:"students",  icon:Users,           label:"นักเรียน" },
  { id:"session",   icon:BookOpen,        label:"บันทึกการสอน" },
  { id:"reports",   icon:BarChart3,       label:"รายงานสรุป" },
  { id:"settings",  icon:Settings,        label:"ตั้งค่า", action:"profile" },
  { id:"backup",    icon:Archive,         label:"สำรองข้อมูล", action:"backup" },
  { id:"reset",     icon:RotateCcw,       label:"รีเซทระบบ", danger:true, action:"reset" },
]

// ── DASHBOARD ─────────────────────────────────────────────────
function Dashboard({ sb, menu, setMenu, search, setSearch, filterLv, setFilterLv, toast, modal, setModal, onDisconnect }) {
  const dashQ    = useDash(sb)
  const sessQ    = useRecentSessions(sb)
  const clsQ     = useClasses(sb)
  const stuQ     = useStudents(sb, search, filterLv)
  const allSessQ = useAllSessions(sb)
  const profQ    = useProfile(sb)

  const stats    = dashQ.data  || { students:0, classes:0, sessions:0, att:{present:0,absent:0,late:0,leave:0,total:0} }
  const sessions = sessQ.data  || []
  const allSess  = allSessQ.data || []
  const classes  = clsQ.data   || []
  const students = stuQ.data   || []
  const profile  = profQ.data

  const reloadAll = () => { dashQ.reload(); sessQ.reload(); clsQ.reload(); stuQ.reload(); allSessQ.reload(); profQ.reload() }
  const thaiDate  = () => new Date().toLocaleDateString("th-TH", { year:"numeric", month:"long", day:"numeric", weekday:"long" })

  // Backup handler
  const handleBackup = async () => {
    try {
      toast.ok("กำลังเตรียมข้อมูล...")
      const [{ data:stuData }, { data:clsData }, { data:sessData }] = await Promise.all([
        sb.from("students").select("*").order("level").order("room").order("student_no"),
        sb.from("classes").select("*").order("created_at"),
        sb.from("teaching_sessions").select("*,classes(class_name,subject_name)").order("created_at", { ascending:false }),
      ])
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stuData||[]),  "นักเรียน")
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(clsData||[]),  "ชั้นเรียน")
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sessData||[]), "บันทึกการสอน")
      const d = new Date().toLocaleDateString("th-TH",{year:"numeric",month:"2-digit",day:"2-digit"}).replace(/\//g,"-")
      XLSX.writeFile(wb, `MyClass_backup_${d}.xlsx`)
      toast.ok("ดาวน์โหลดข้อมูลสำเร็จ ✅")
    } catch(e) { toast.err("สำรองข้อมูลไม่ได้: " + e.message) }
  }

  // Reset handler
  const handleReset = async () => {
    const confirmed = window.confirm("⚠️ จะลบข้อมูลทั้งหมด!\n\nนักเรียน, ชั้นเรียน และบันทึกการสอนทั้งหมดจะหายไปถาวร\n\nยืนยันหรือไม่?")
    if (!confirmed) return
    const confirmed2 = window.confirm("ยืนยันอีกครั้ง — ข้อมูลทั้งหมดจะถูกลบถาวร ไม่สามารถกู้คืนได้")
    if (!confirmed2) return
    try {
      toast.ok("กำลังลบข้อมูล...")
      // Delete in order (cascade will handle children, but let's be explicit)
      await sb.from("attendance").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("behavior_logs").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("assessments").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("session_images").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("teaching_sessions").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("class_students").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("classes").delete().neq("id","00000000-0000-0000-0000-000000000000")
      await sb.from("students").delete().neq("id","00000000-0000-0000-0000-000000000000")
      toast.ok("รีเซทข้อมูลเรียบร้อย")
      setMenu("dashboard")
      reloadAll()
    } catch(e) { toast.err("รีเซทไม่สำเร็จ: " + e.message) }
  }

  const handleMenuClick = (m) => {
    if (m.action === "profile") { setModal("profile"); return }
    if (m.action === "backup")  { handleBackup(); return }
    if (m.action === "reset")   { handleReset(); return }
    setMenu(m.id)
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');
        @keyframes mcSpin{to{transform:rotate(360deg)}}
        @keyframes mcSlide{from{transform:translateX(24px);opacity:0}to{transform:translateX(0);opacity:1}}
        *{box-sizing:border-box} body{margin:0}
        tr:hover td{background:#FAFAFA}
      `}</style>
      <Toast list={toast.list}/>
      <div style={ss.app}>
        {/* Sidebar */}
        <aside style={ss.sidebar}>
          <div style={{ display:"flex", alignItems:"center", gap:10, padding:"18px 16px 14px", borderBottom:`1px solid ${T.border}` }}>
            <div style={{ width:34, height:34, borderRadius:10, background:T.purple, display:"flex", alignItems:"center", justifyContent:"center" }}>
              <GraduationCap size={20} color="#fff"/>
            </div>
            <span style={{ fontWeight:700, fontSize:17, color:T.text }}>MyClass 🍎</span>
          </div>
          <nav style={ss.nav}>
            {MENU.map(m => (
              <div key={m.id} style={ss.navI(menu===m.id && !m.action, m.danger)} onClick={()=>handleMenuClick(m)}>
                <m.icon size={17}/>{m.label}
              </div>
            ))}
          </nav>
          <button style={ss.addBtn} onClick={()=>setModal("addClass")}>
            <Plus size={16}/>เพิ่มชั้นเรียน
          </button>
        </aside>

        {/* Main */}
        <div style={ss.main}>
          {/* Topbar */}
          <div style={ss.topbar}>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <div style={{ width:8, height:8, borderRadius:"50%", background:T.green }}/>
              <span style={{ fontSize:12.5, color:T.muted, fontWeight:500 }}>Supabase Connected</span>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <button onClick={reloadAll} title="รีเฟรชข้อมูล" style={{ border:"none", background:"none", cursor:"pointer", padding:6, borderRadius:8 }}>
                <RefreshCw size={17} color={T.muted}/>
              </button>
              <button onClick={()=>setModal("profile")} style={{ display:"flex", alignItems:"center", gap:8, border:"none", background:"none", cursor:"pointer", padding:"4px 8px", borderRadius:10 }}>
                <div style={{ width:34, height:34, borderRadius:"50%", background:`linear-gradient(135deg,${T.purple},${T.blue})`, display:"flex", alignItems:"center", justifyContent:"center" }}>
                  <span style={{ color:"#fff", fontSize:13, fontWeight:700 }}>{profile?.teacher_name?.[0] || "ค"}</span>
                </div>
                <div style={{ textAlign:"left" }}>
                  <div style={{ fontSize:13.5, fontWeight:600, color:T.text, lineHeight:1.2 }}>{profile?.teacher_name || "ตั้งค่าชื่อครู"}</div>
                  <div style={{ fontSize:11.5, color:T.muted }}>{profile?.school_name || "กดเพื่อตั้งค่า"}</div>
                </div>
                <ChevronDown size={14} color={T.muted}/>
              </button>
            </div>
          </div>

          {/* Content */}
          <div style={ss.content}>

            {/* ── DASHBOARD ── */}
            {menu === "dashboard" && <>
              <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:12 }}>
                <div>
                  <div style={{ fontSize:22, fontWeight:700, color:T.text, marginBottom:2 }}>ยินดีต้อนรับ {profile?.teacher_name || "คุณครู"} 👋</div>
                  <div style={{ fontSize:13, color:T.muted }}>{thaiDate()}</div>
                </div>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
                  <button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>สร้างบันทึกการสอน</button>
                  <button style={ss.btnP(T.blue)} onClick={()=>setMenu("classes")}><School size={15}/>ชั้นเรียนของฉัน</button>
                </div>
              </div>
              {/* Stats */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20 }}>
                {[
                  { icon:p=><Users {...p}/>,       label:"นักเรียนทั้งหมด", v:stats.students, u:"คน",   bg:T.blueL,   ic:T.blue   },
                  { icon:p=><School {...p}/>,       label:"ชั้นเรียน",       v:stats.classes,  u:"ห้อง",  bg:T.yellowL, ic:T.yellow },
                  { icon:p=><CalendarDays {...p}/>, label:"คาบที่สอน",       v:stats.sessions, u:"คาบ",   bg:T.purpleL, ic:T.purple },
                  { icon:p=><Star {...p}/>,         label:"เข้าเรียน",       v:stats.att.total?`${Math.round(stats.att.present/stats.att.total*100)}%`:"—", u:"เฉลี่ย", bg:T.greenL, ic:T.green },
                ].map((it,i) => (
                  <div key={i} style={{ background:"#fff", borderRadius:14, padding:"16px", border:`1px solid ${T.border}` }}>
                    <div style={{ width:42, height:42, borderRadius:12, background:it.bg, display:"flex", alignItems:"center", justifyContent:"center", marginBottom:10 }}>
                      {it.icon({ size:22, color:it.ic })}
                    </div>
                    {dashQ.loading ? <Spin size={24}/> : <>
                      <div style={{ fontSize:28, fontWeight:700, color:T.text, lineHeight:1 }}>{it.v}</div>
                      <div style={{ fontSize:12.5, color:T.muted, marginTop:4 }}>{it.label} · {it.u}</div>
                    </>}
                  </div>
                ))}
              </div>
              {/* 2-col */}
              <div style={{ display:"grid", gridTemplateColumns:"1.4fr 1fr", gap:16, marginBottom:20 }}>
                <div style={ss.card}>
                  <div style={ss.cHead}><span style={ss.cTitle}>บันทึกการสอนล่าสุด</span>{sessQ.loading && <Spin size={16}/>}</div>
                  {sessQ.error ? <EmptyState icon="❌" msg={sessQ.error}/> : !sessions.length ? <EmptyState icon="📋" msg="ยังไม่มีบันทึกการสอน"/> : (
                    <table style={ss.table}>
                      <thead><tr style={{ background:"#F8FAFC" }}>{["ชั้น/วิชา","หัวข้อ","วันที่","สถานะ"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead>
                      <tbody>{sessions.map(s => (
                        <tr key={s.id}>
                          <td style={ss.td}>
                            <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                              <div style={{ width:8, height:8, borderRadius:"50%", background:s.classes?.color||T.purple, flexShrink:0 }}/>
                              <div>
                                <div style={{ fontWeight:600, fontSize:13 }}>{s.classes?.class_name||"—"}</div>
                                <div style={{ color:T.muted, fontSize:11.5 }}>{s.classes?.subject_name}</div>
                              </div>
                            </div>
                          </td>
                          <td style={{ ...ss.td, color:"#64748B", fontSize:13, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{s.topic||"—"}</td>
                          <td style={{ ...ss.td, fontSize:12, color:T.muted, whiteSpace:"nowrap" }}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short"}):"—"}</td>
                          <td style={ss.td}>{s.status==="saved"?<span style={ss.pill("#15803D",T.greenL)}><Check size={11}/>สมบูรณ์</span>:<span style={ss.pill("#C2410C",T.orangeL)}><Clock size={11}/>ร่าง</span>}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  )}
                </div>
                <div style={ss.card}>
                  <div style={ss.cHead}><span style={ss.cTitle}>สรุปการเข้าเรียน</span>{dashQ.loading && <Spin size={16}/>}</div>
                  {!dashQ.loading && <AttChart att={stats.att}/>}
                </div>
              </div>
              {/* Class cards */}
              {classes.length > 0 && <>
                <div style={{ fontSize:14.5, fontWeight:700, color:T.text, marginBottom:12 }}>ชั้นเรียนของฉัน</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:14 }}>
                  {classes.slice(0,6).map(c => (
                    <div key={c.id} style={{ background:"#fff", borderRadius:14, border:`1px solid ${T.border}`, overflow:"hidden", cursor:"pointer" }} onClick={()=>setMenu("classes")}>
                      <div style={{ height:5, background:c.color||T.purple }}/>
                      <div style={{ padding:"14px" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
                          <div>
                            <div style={{ fontSize:15, fontWeight:700, color:T.text }}>{c.class_name}</div>
                            <div style={{ fontSize:12.5, color:T.muted }}>{c.subject_name}</div>
                          </div>
                          <div style={{ width:36, height:36, borderRadius:10, background:(c.color||T.purple)+"22", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700, color:c.color||T.purple, fontSize:15 }}>{c.class_name?.[0]}</div>
                        </div>
                        <div style={{ fontSize:12.5, color:"#64748B" }}><Users size={12} style={{ verticalAlign:"middle" }}/> {c._studentCount||0} คน</div>
                      </div>
                    </div>
                  ))}
                </div>
              </>}
            </>}

            {/* ── CLASSES ── */}
            {menu === "classes" && <>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:18 }}>
                <div>
                  <div style={{ fontSize:18, fontWeight:700, color:T.text }}>ชั้นเรียนของฉัน</div>
                  <div style={{ fontSize:13, color:T.muted }}>{clsQ.loading?"กำลังโหลด...":`${classes.length} ชั้นเรียน`}</div>
                </div>
                <button style={ss.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>เพิ่มชั้นเรียน</button>
              </div>
              {clsQ.loading ? <PageLoad/> : clsQ.error ? <EmptyState icon="❌" msg={clsQ.error}/> : !classes.length ? (
                <EmptyState icon="🏫" msg="ยังไม่มีชั้นเรียน" btn={<button style={ss.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>สร้างชั้นเรียนแรก</button>}/>
              ) : (
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:16 }}>
                  {classes.map(c => (
                    <div key={c.id} style={{ background:"#fff", borderRadius:16, border:`1px solid ${T.border}`, overflow:"hidden" }}>
                      <div style={{ height:6, background:c.color||T.purple }}/>
                      <div style={{ padding:"16px" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
                          <div>
                            <div style={{ fontSize:16, fontWeight:700, color:T.text }}>{c.class_name}</div>
                            <div style={{ fontSize:13, color:T.muted }}>{c.subject_name}</div>
                          </div>
                          <div style={{ width:40, height:40, borderRadius:12, background:(c.color||T.purple)+"22", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16, fontWeight:700, color:c.color||T.purple }}>{c.class_name?.[0]}</div>
                        </div>
                        <div style={{ fontSize:12.5, color:"#64748B", marginBottom:14 }}><Users size={12} style={{ verticalAlign:"middle" }}/> {c._studentCount||0} คน</div>
                        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                          <button onClick={()=>setModal("session")} style={{ padding:"8px 0", background:c.color||T.purple, color:"#fff", border:"none", borderRadius:10, cursor:"pointer", fontSize:12.5, fontWeight:600, fontFamily:"inherit" }}>▶ บันทึกสอน</button>
                          <button onClick={async()=>{
                            if (!window.confirm(`ลบ ${c.class_name}?`)) return
                            const { error } = await sb.from("classes").delete().eq("id", c.id)
                            if (error) { toast.err(error.message||"ลบไม่ได้"); return }
                            toast.ok("ลบแล้ว"); clsQ.reload(); dashQ.reload()
                          }} style={{ padding:"8px 0", background:T.redL, color:T.red, border:"none", borderRadius:10, cursor:"pointer", fontSize:12.5, fontWeight:600, fontFamily:"inherit" }}>ลบ</button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>}


            {/* STUDENTS */}
            {menu==="students"&&<>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
                <div><div style={{fontSize:18,fontWeight:700,color:T.text}}>นักเรียน (ฐานข้อมูลกลาง)</div><div style={{fontSize:13,color:T.muted}}>{stuQ.loading?"กำลังโหลด...":`${students.length} คน`}</div></div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <button style={ss.btnO} onClick={()=>setModal("paste")}><ClipboardList size={15}/>วางรายชื่อ</button>
                  <button style={ss.btnP(T.green)} onClick={()=>setModal("excel")}><Upload size={15}/>Import Excel</button>
                  <button style={ss.btnP()} onClick={()=>setModal("addStu")}><Plus size={15}/>เพิ่มเอง</button>
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
                <div style={{flex:1,minWidth:200,position:"relative"}}>
                  <Search size={15} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:T.muted}}/>
                  <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาชื่อ..." style={{...ss.input,paddingLeft:34}}/>
                </div>
                {["ทั้งหมด","ม.1","ม.2","ม.3","ม.4","ม.5","ม.6"].map(l=>(
                  <button key={l} onClick={()=>setFilterLv(l)} style={{padding:"7px 12px",border:"1px solid",borderColor:filterLv===l?T.purple:T.border,borderRadius:10,cursor:"pointer",background:filterLv===l?T.purpleL:"#fff",color:filterLv===l?T.purple:"#64748B",fontSize:12.5,fontWeight:600,fontFamily:"inherit"}}>{l}</button>
                ))}
              </div>
              {stuQ.loading?<PageLoad/>:stuQ.error?<EmptyState icon="❌" msg={stuQ.error}/>:!students.length?(
                <EmptyState icon="👩‍🎓" msg="ไม่พบนักเรียน" btn={<button style={ss.btnP()} onClick={()=>setModal("addStu")}><Plus size={15}/>เพิ่มนักเรียน</button>}/>
              ):(
                <div style={ss.card}>
                  <table style={ss.table}>
                    <thead><tr style={{background:"#F8FAFC"}}>{["#","ชื่อ-นามสกุล","ชั้น","ห้อง",""].map((h,i)=><th key={i} style={ss.th}>{h}</th>)}</tr></thead>
                    <tbody>{students.map((stu,i)=>(
                      <tr key={stu.id}>
                        <td style={{...ss.td,color:T.muted,fontSize:12.5,width:40}}>{stu.student_no||i+1}</td>
                        <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:30,height:30,borderRadius:"50%",background:T.blueL,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:T.blue,flexShrink:0}}>{stu.full_name?.[0]}</div>{stu.full_name}</div></td>
                        <td style={ss.td}>{stu.level}</td>
                        <td style={ss.td}>{stu.room}</td>
                        <td style={{...ss.td,width:50}}>
                          <button onClick={async()=>{
                            if(!window.confirm("ลบ "+stu.full_name+"?"))return
                            const{error}=await sb.from("students").delete().eq("id",stu.id)
                            if(error){toast.err(error.message||"ลบไม่ได้");return}
                            toast.ok("ลบแล้ว");stuQ.reload();dashQ.reload()
                          }} style={{padding:"5px 8px",border:"1px solid #FECACA",borderRadius:8,cursor:"pointer",background:T.redL,color:T.red}}>
                            <Trash2 size={13}/>
                          </button>
                        </td>
                      </tr>
                    ))}</tbody>
                  </table>
                  <div style={{padding:"10px 14px",background:"#F8FAFC",fontSize:12.5,color:T.muted,borderTop:"1px solid #F1F5F9"}}>แสดง {students.length} คน</div>
                </div>
              )}
            </>}

            {menu==="session"&&<>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
                <div><div style={{fontSize:18,fontWeight:700,color:T.text}}>บันทึกการสอน</div><div style={{fontSize:13,color:T.muted}}>{allSessQ.loading?"กำลังโหลด...":`${allSess.length} บันทึก`}</div></div>
                <button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>สร้างบันทึกใหม่</button>
              </div>
              {allSessQ.loading?<PageLoad/>:allSessQ.error?<EmptyState icon="❌" msg={allSessQ.error}/>:!allSess.length?(
                <EmptyState icon="📋" msg="ยังไม่มีบันทึกการสอน" btn={<button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>สร้างบันทึกแรก</button>}/>
              ):(
                <div style={ss.card}>
                  <table style={ss.table}>
                    <thead><tr style={{background:"#F8FAFC"}}>{["ชั้น/วิชา","เรื่องที่สอน","วันที่","จุดประสงค์","สถานะ"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead>
                    <tbody>{allSess.map(s=>(
                      <tr key={s.id}>
                        <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:8,height:8,borderRadius:"50%",background:s.classes?.color||T.purple,flexShrink:0}}/><div><div style={{fontWeight:600,fontSize:13}}>{s.classes?.class_name||"—"}</div><div style={{color:T.muted,fontSize:11.5}}>{s.classes?.subject_name}</div></div></div></td>
                        <td style={{...ss.td,maxWidth:160}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:13}}>{s.topic||"—"}</div></td>
                        <td style={{...ss.td,whiteSpace:"nowrap",color:T.muted,fontSize:12}}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short",year:"2-digit"}):"—"}</td>
                        <td style={{...ss.td,maxWidth:160}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:12,color:"#64748B"}}>{s.objective||"—"}</div></td>
                        <td style={ss.td}>{s.status==="saved"?<span style={ss.pill("#15803D",T.greenL)}><Check size={11}/>สมบูรณ์</span>:<span style={ss.pill("#C2410C",T.orangeL)}><Clock size={11}/>ร่าง</span>}</td>
                      </tr>
                    ))}</tbody>
                  </table>
                </div>
              )}
            </>}

            {menu==="reports"&&(
              <div style={{textAlign:"center",paddingTop:60}}>
                <div style={{fontSize:60,marginBottom:16}}>📊</div>
                <div style={{fontSize:18,fontWeight:600,color:T.muted,marginBottom:8}}>รายงานสรุป</div>
                <div style={{fontSize:13,color:"#CBD5E1",marginBottom:20}}>อยู่ระหว่างพัฒนา</div>
                <button style={ss.btnP()} onClick={handleBackup}><Download size={15}/>ดาวน์โหลดข้อมูลทั้งหมด (Excel)</button>
              </div>
            )}

            {!["dashboard","classes","students","session","reports"].includes(menu)&&(
              <div style={{textAlign:"center",paddingTop:60}}><div style={{fontSize:60,marginBottom:16}}>🚧</div><div style={{fontSize:18,fontWeight:600,color:T.muted}}>{MENU.find(m=>m.id===menu)?.label}</div></div>
            )}

          </div>
        </div>

        {modal==="addClass"&&<AddClassModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{clsQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="addStu"&&<AddStudentModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="paste"&&<PasteModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="excel"&&<ImportExcelModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="session"&&<SessionModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{allSessQ.reload();sessQ.reload();dashQ.reload()}} classes={classes} toast={toast}/>}
        {modal==="profile"&&<ProfileModal sb={sb} onClose={()=>setModal(null)} profile={profQ.data} onDone={profQ.reload} toast={toast} onDisconnect={onDisconnect}/>}
      </div>
    </>
  )
}

import React from "react"

export default function App(){
  const[connected,setConnected]=useState(false)
  const[sb,setSb]=useState(null)
  const[menu,setMenu]=useState("dashboard")
  const[search,setSearch]=useState("")
  const[filterLv,setFilterLv]=useState("ทั้งหมด")
  const[modal,setModal]=useState(null)
  const toast=useToast()
  useEffect(()=>{const u=store.get("mc_url")||"";const k=store.get("mc_key")||"";if(u&&k){setSb(createClient(u,k));setConnected(true)}},[])
  const handleReady=(url,key)=>{setSb(createClient(url,key));setConnected(true)}
  const handleDisconnect=()=>{store.remove("mc_url");store.remove("mc_key");setSb(null);setConnected(false)}
  if(!connected||!sb)return <SetupScreen onReady={handleReady}/>
  return <ErrorBoundary><Dashboard sb={sb} menu={menu} setMenu={setMenu} search={search} setSearch={setSearch} filterLv={filterLv} setFilterLv={setFilterLv} toast={toast} modal={modal} setModal={setModal} onDisconnect={handleDisconnect}/></ErrorBoundary>
}
