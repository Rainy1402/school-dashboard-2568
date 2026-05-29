import { useState, useEffect, useCallback, useRef } from "react"
import React from "react"
import * as XLSX from "xlsx"
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid } from "recharts"
import {
  LayoutDashboard, School, Users, BookOpen, BarChart3, Settings, Archive,
  RotateCcw, Plus, X, Check, Clock, Upload, Trash2, RefreshCw, GraduationCap,
  AlertCircle, CheckCircle, WifiOff, Search, Save, ChevronDown, Star,
  CalendarDays, ClipboardList, Download, ChevronRight, ChevronLeft,
  UserCheck, Award, ThumbsUp, ThumbsDown, Minus, UserPlus, Eye
} from "lucide-react"

// ── STORE ─────────────────────────────────────────────────────
const _s = {}
const store = { get:k=>_s[k]??null, set:(k,v)=>{_s[k]=v}, remove:k=>{delete _s[k]} }

// ── SUPABASE CLIENT ───────────────────────────────────────────
function createClient(baseUrl, apiKey) {
  const url = baseUrl.replace(/\/$/, "")
  const h = { apikey:apiKey, Authorization:`Bearer ${apiKey}`, "Content-Type":"application/json" }

  function qb(table) {
    let _sel="*", _filters=[], _orders=[], _lim=null, _countMode=null, _isHead=false, _single=false
    const q = {
      select(f="*",o={}){_sel=f.replace(/\s+/g,"");if(o.count)_countMode=o.count;if(o.head)_isHead=o.head;return q},
      eq(c,v){_filters.push({c,op:"eq",v:String(v)});return q},
      neq(c,v){_filters.push({c,op:"neq",v:String(v)});return q},
      ilike(c,v){_filters.push({c,op:"ilike",v});return q},
      order(c,o={}){_orders.push(`${c}.${o.ascending===false?"desc":"asc"}`);return q},
      limit(n){_lim=n;return q},
      maybeSingle(){_single=true;_lim=1;return q},
      async then(res,rej){
        try{
          let ep=`${url}/rest/v1/${table}?select=${_sel}`
          _filters.forEach(f=>{ep+=`&${f.c}=${f.op}.${encodeURIComponent(f.v)}`})
          if(_orders.length)ep+=`&order=${_orders.join(",")}`
          if(_lim)ep+=`&limit=${_lim}`
          const rh={...h};if(_countMode==="exact")rh["Prefer"]="count=exact"
          const r=await fetch(ep,{method:_isHead?"HEAD":"GET",headers:rh})
          if(_isHead){const rng=r.headers.get("Content-Range")||"*/0";res({count:parseInt(rng.split("/")[1]||"0")||0,data:null,error:null});return}
          const b=await r.json().catch(()=>[])
          if(!r.ok){res({data:null,error:b,count:null});return}
          let count=null
          if(_countMode==="exact"){const rng=r.headers.get("Content-Range")||"";if(rng.includes("/"))count=parseInt(rng.split("/")[1])||0}
          const d=Array.isArray(b)?b:(b?[b]:[])
          if(_single)res({data:d[0]??null,error:null})
          else res({data:d,error:null,count})
        }catch(e){rej(e)}
      }
    }
    return q
  }
  return {
    from(table){
      return{
        select(f="*",o={}){return qb(table).select(f,o)},
        async insert(data){
          try{
            const r=await fetch(`${url}/rest/v1/${table}`,{method:"POST",headers:{...h,Prefer:"return=minimal"},body:JSON.stringify(data)})
            if(r.status===201||r.status===204)return{data:null,error:null}
            const b=await r.json().catch(()=>({}))
            return r.ok?{data:b,error:null}:{data:null,error:b}
          }catch(e){return{data:null,error:{message:e.message}}}
        },
        async insertReturning(data){
          try{
            const r=await fetch(`${url}/rest/v1/${table}`,{method:"POST",headers:{...h,Prefer:"return=representation"},body:JSON.stringify(data)})
            const b=await r.json().catch(()=>({}))
            if(!r.ok)return{data:null,error:b}
            const item=Array.isArray(b)?(b[0]??null):(b??null)
            return{data:item,error:null}
          }catch(e){return{data:null,error:{message:e.message}}}
        },
        update(data){
          const _f=[]
          const u={
            eq(c,v){_f.push(`${c}=eq.${encodeURIComponent(v)}`);return u},
            async then(res){
              try{
                const ep=`${url}/rest/v1/${table}${_f.length?"?"+_f.join("&"):""}`
                const r=await fetch(ep,{method:"PATCH",headers:{...h,Prefer:"return=minimal"},body:JSON.stringify(data)})
                if(r.status===204||r.status===200){res({data:null,error:null});return}
                const b=await r.json().catch(()=>({}))
                res(r.ok?{data:b,error:null}:{data:null,error:b})
              }catch(e){res({data:null,error:{message:e.message}})}
            }
          }
          return u
        },
        delete(){
          const _f=[]
          const d={
            eq(c,v){_f.push(`${c}=eq.${encodeURIComponent(v)}`);return d},
            async then(res){
              if(!_f.length){res({error:{message:"ต้องระบุ filter"}});return}
              try{
                const r=await fetch(`${url}/rest/v1/${table}?${_f.join("&")}`,{method:"DELETE",headers:h})
                if(r.ok){res({error:null});return}
                const b=await r.json().catch(()=>({}));res({error:b})
              }catch(e){res({error:{message:e.message}})}
            }
          }
          return d
        }
      }
    }
  }
}

// ── THEME ─────────────────────────────────────────────────────
const T={
  purple:"#7C3AED",purpleL:"#EDE9FE",
  blue:"#2563EB",blueL:"#DBEAFE",
  green:"#16A34A",greenL:"#DCFCE7",
  red:"#DC2626",redL:"#FEE2E2",
  yellow:"#D97706",yellowL:"#FEF3C7",
  orange:"#EA580C",orangeL:"#FFEDD5",
  slate:"#F8FAFC",border:"#E2E8F0",
  text:"#1E293B",muted:"#94A3B8",
}
const CLASS_COLORS=["#7C3AED","#2563EB","#16A34A","#D97706","#DC2626","#0891B2","#BE185D","#0F766E"]
const ATT_CFG={
  present:{label:"มา",color:T.green,bg:T.greenL},
  absent: {label:"ขาด",color:T.red,bg:T.redL},
  late:   {label:"สาย",color:T.yellow,bg:T.yellowL},
  leave:  {label:"ลา",color:T.blue,bg:T.blueL},
}
const ASS_CFG={
  good:    {label:"ดีมาก",icon:"⭐",color:"#7C3AED",bg:"#EDE9FE"},
  pass:    {label:"ผ่าน",icon:"✓",color:T.green,bg:T.greenL},
  improve: {label:"ปรับปรุง",icon:"⚠️",color:T.orange,bg:T.orangeL},
}
const ss={
  app:{position:"relative",height:"100vh",overflow:"hidden",display:"flex",fontFamily:"'Prompt','Sarabun',sans-serif",background:T.slate},
  sidebar:{width:220,minWidth:220,background:"#fff",borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",overflowY:"auto"},
  main:{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"},
  topbar:{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:58,flexShrink:0},
  content:{flex:1,overflowY:"auto",padding:"24px"},
  overlay:{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center",padding:16},
  nav:{flex:1,padding:"8px"},
  navI:(a,d)=>({display:"flex",alignItems:"center",gap:10,padding:"9px 10px",borderRadius:10,cursor:"pointer",fontSize:13.5,fontWeight:a?600:500,marginBottom:2,background:a?T.purple:"transparent",color:a?"#fff":d?T.red:"#475569",transition:"all 0.15s"}),
  addBtn:{margin:"12px 10px 16px",padding:"10px 0",background:T.purple,color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13.5,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6},
  card:{background:"#fff",borderRadius:14,border:`1px solid ${T.border}`,overflow:"hidden"},
  cHead:{padding:"14px 18px 12px",display:"flex",alignItems:"center",justifyContent:"space-between"},
  cTitle:{fontSize:14.5,fontWeight:700,color:T.text},
  table:{width:"100%",borderCollapse:"collapse"},
  th:{padding:"8px 14px",fontSize:12.5,color:T.muted,fontWeight:600,textAlign:"left",borderBottom:`1px solid #F1F5F9`,whiteSpace:"nowrap"},
  td:{padding:"9px 14px",fontSize:13.5,color:"#334155",borderBottom:`1px solid #F8FAFC`},
  input:{width:"100%",border:`1px solid ${T.border}`,borderRadius:10,padding:"9px 12px",fontSize:13.5,fontFamily:"inherit",outline:"none",boxSizing:"border-box"},
  label:{display:"block",fontSize:12.5,fontWeight:600,color:"#64748B",marginBottom:5},
  fRow:{marginBottom:14},
  textarea:{width:"100%",border:`1px solid ${T.border}`,borderRadius:10,padding:"9px 12px",fontSize:13.5,fontFamily:"inherit",outline:"none",resize:"vertical",minHeight:72,boxSizing:"border-box"},
  btnP:(c=T.purple)=>({padding:"9px 18px",background:c,color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13.5,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6}),
  btnO:{padding:"9px 18px",background:"#fff",color:"#64748B",border:`1px solid ${T.border}`,borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13.5,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6},
  pill:(c,bg)=>({padding:"3px 10px",borderRadius:99,fontSize:12,fontWeight:600,background:bg,color:c,display:"inline-flex",alignItems:"center",gap:4}),
  modal:{background:"#fff",borderRadius:20,width:500,maxWidth:"calc(100vw - 32px)",maxHeight:"90vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 24px 60px rgba(0,0,0,0.25)"},
  mHead:{padding:"16px 20px",borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0},
  mBody:{padding:"20px",overflowY:"auto",flex:1},
}

// ── MICRO COMPONENTS ──────────────────────────────────────────
function Spin({size=18,color=T.purple}){return <div style={{width:size,height:size,border:`2.5px solid #E2E8F0`,borderTopColor:color,borderRadius:"50%",animation:"mcSpin 0.7s linear infinite",flexShrink:0}}/>}
function PageLoad({text="กำลังโหลด..."}){return <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:280,gap:12}}><Spin size={34}/><span style={{color:T.muted,fontSize:14}}>{text}</span></div>}
function EmptyState({icon="📭",msg="ไม่มีข้อมูล",btn}){return <div style={{textAlign:"center",padding:"52px 24px"}}><div style={{fontSize:48,marginBottom:12}}>{icon}</div><div style={{fontSize:15,fontWeight:600,color:T.muted}}>{msg}</div>{btn&&<div style={{marginTop:16}}>{btn}</div>}</div>}
function Toast({list}){if(!list.length)return null;return <div style={{position:"fixed",top:12,right:12,zIndex:300,display:"flex",flexDirection:"column",gap:8}}>{list.map(t=><div key={t.id} style={{background:t.type==="err"?T.red:T.green,color:"#fff",borderRadius:12,padding:"11px 16px",fontSize:13.5,fontWeight:600,display:"flex",alignItems:"center",gap:8,fontFamily:"inherit",boxShadow:"0 4px 20px rgba(0,0,0,0.2)",maxWidth:320}}>{t.type==="err"?<AlertCircle size={15}/>:<CheckCircle size={15}/>}{t.msg}</div>)}</div>}
function useToast(){const[list,setList]=useState([]);const add=useCallback((msg,type="ok")=>{const id=Date.now();setList(p=>[...p,{id,msg,type}]);setTimeout(()=>setList(p=>p.filter(t=>t.id!==id)),3200)},[]);return{list,ok:m=>add(m,"ok"),err:m=>add(m,"err")}}

// ── ERROR BOUNDARY ────────────────────────────────────────────
class ErrorBoundary extends React.Component{
  constructor(p){super(p);this.state={err:null}}
  static getDerivedStateFromError(e){return{err:e}}
  render(){
    if(!this.state.err)return this.props.children
    return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:"#FEF2F2",fontFamily:"'Prompt','Sarabun',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"32px 36px",maxWidth:460,textAlign:"center",boxShadow:"0 8px 32px rgba(0,0,0,0.1)"}}>
        <div style={{fontSize:48,marginBottom:12}}>⚠️</div>
        <div style={{fontSize:18,fontWeight:700,color:T.text,marginBottom:8}}>เกิดข้อผิดพลาด</div>
        <div style={{fontSize:13,color:T.muted,marginBottom:20,wordBreak:"break-word"}}>{this.state.err?.message}</div>
        <button onClick={()=>window.location.reload()} style={{...ss.btnP(),width:"100%",justifyContent:"center"}}>🔄 โหลดหน้าใหม่</button>
      </div>
    </div>
  }
}

// ── DATA HOOKS ────────────────────────────────────────────────
function useQ(fn,deps=[]){
  const[data,setData]=useState(null);const[loading,setLoading]=useState(true);const[error,setError]=useState(null)
  const reload=useCallback(async()=>{setLoading(true);setError(null);try{setData(await fn())}catch(e){setError(e.message);setData(null)}finally{setLoading(false)}},deps)
  useEffect(()=>{reload()},[reload]);return{data,loading,error,reload}
}
function useDash(sb){return useQ(async()=>{
  const[sc,cc,sesc,att]=await Promise.all([
    sb.from("students").select("id",{count:"exact",head:true}),
    sb.from("classes").select("id",{count:"exact",head:true}),
    sb.from("teaching_sessions").select("id",{count:"exact",head:true}),
    sb.from("attendance").select("status"),
  ])
  const a=att.data||[]
  return{students:sc.count||0,classes:cc.count||0,sessions:sesc.count||0,
    att:{present:a.filter(x=>x.status==="present").length,absent:a.filter(x=>x.status==="absent").length,
      late:a.filter(x=>x.status==="late").length,leave:a.filter(x=>x.status==="leave").length,total:a.length}}
},[sb])}
function useRecentSessions(sb){return useQ(async()=>{const{data,error}=await sb.from("teaching_sessions").select("*,classes(class_name,subject_name,color)").order("created_at",{ascending:false}).limit(6);if(error)throw new Error(error.message);return data||[]},[sb])}
function useClasses(sb){return useQ(async()=>{const{data,error}=await sb.from("classes").select("*,class_students(*)").order("created_at",{ascending:false});if(error)throw new Error(error.message);return(data||[]).map(c=>({...c,_studentCount:Array.isArray(c.class_students)?c.class_students.length:0}))},[sb])}
function useStudents(sb,search,level){return useQ(async()=>{let q=sb.from("students").select("*").order("level").order("room").order("student_no");if(level&&level!=="ทั้งหมด")q=q.eq("level",level);if(search&&search.trim())q=q.ilike("full_name",`%${search.trim()}%`);const{data,error}=await q;if(error)throw new Error(error.message);return data||[]},[sb,search,level])}
function useAllSessions(sb){return useQ(async()=>{const{data,error}=await sb.from("teaching_sessions").select("*,classes(class_name,subject_name,color)").order("created_at",{ascending:false});if(error)throw new Error(error.message);return data||[]},[sb])}
function useProfile(sb){return useQ(async()=>{const{data}=await sb.from("profiles").select("*").limit(1).maybeSingle();return data},[sb])}

// ── SETUP SCREEN ──────────────────────────────────────────────
function SetupScreen({onReady}){
  const[url,setUrl]=useState(store.get("mc_url")||"");const[key,setKey]=useState(store.get("mc_key")||"")
  const[testing,setTesting]=useState(false);const[err,setErr]=useState("")
  const connect=async()=>{
    if(!url.trim()||!key.trim()){setErr("กรุณากรอกทั้ง URL และ Key");return}
    setTesting(true);setErr("")
    try{
      const client=createClient(url.trim(),key.trim())
      const{error}=await client.from("profiles").select("id").limit(1)
      if(error&&error.code!=="PGRST116")throw new Error(error.message||JSON.stringify(error))
      store.set("mc_url",url.trim());store.set("mc_key",key.trim());onReady(url.trim(),key.trim())
    }catch(e){setErr("เชื่อมต่อไม่ได้: "+e.message)}finally{setTesting(false)}
  }
  return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:"linear-gradient(135deg,#1E1B4B 0%,#312E81 50%,#1E3A5F 100%)",fontFamily:"'Prompt','Sarabun',sans-serif"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');@keyframes mcSpin{to{transform:rotate(360deg)}}*{box-sizing:border-box}body{margin:0}`}</style>
      <div style={{background:"#fff",borderRadius:24,padding:"36px 40px",width:460,maxWidth:"calc(100vw - 32px)",boxShadow:"0 32px 80px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:28}}>
          <div style={{width:52,height:52,borderRadius:16,background:T.purple,display:"flex",alignItems:"center",justifyContent:"center"}}><GraduationCap size={26} color="#fff"/></div>
          <div><div style={{fontSize:20,fontWeight:700,color:T.text}}>MyClass 🍎</div><div style={{fontSize:13,color:T.muted}}>เชื่อมต่อ Supabase เพื่อเริ่มใช้งาน</div></div>
        </div>
        <div style={ss.fRow}><label style={ss.label}>Supabase Project URL</label><input value={url} onChange={e=>setUrl(e.target.value)} style={ss.input} placeholder="https://xxxx.supabase.co" type="url" onKeyDown={e=>e.key==="Enter"&&connect()}/></div>
        <div style={ss.fRow}><label style={ss.label}>Anon Public Key</label><input value={key} onChange={e=>setKey(e.target.value)} style={{...ss.input,fontFamily:"monospace",fontSize:12}} placeholder="eyJhbGci..." onKeyDown={e=>e.key==="Enter"&&connect()}/></div>
        {err&&<div style={{background:T.redL,borderRadius:10,padding:"10px 14px",fontSize:13,color:T.red,marginBottom:16,display:"flex",gap:8}}><AlertCircle size={15} style={{flexShrink:0,marginTop:1}}/>{err}</div>}
        <button onClick={connect} disabled={testing} style={{...ss.btnP(),width:"100%",justifyContent:"center",padding:"13px 0",fontSize:14.5,opacity:testing?0.75:1}}>
          {testing?<><Spin size={17} color="#fff"/>กำลังเชื่อมต่อ...</>:<><CheckCircle size={17}/>เชื่อมต่อและเริ่มใช้งาน</>}
        </button>
      </div>
    </div>
  )
}

// ── ATT CHART ─────────────────────────────────────────────────
const PIE_LABEL=({cx,cy,midAngle,innerRadius,outerRadius,percent})=>{if(percent<0.04)return null;const R=Math.PI/180,r=innerRadius+(outerRadius-innerRadius)*0.55;return <text x={cx+r*Math.cos(-midAngle*R)} y={cy+r*Math.sin(-midAngle*R)} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={11} fontWeight={700}>{`${(percent*100).toFixed(0)}%`}</text>}
function AttChart({att}){
  const items=[{name:"มาเรียน",value:att.present,color:T.green},{name:"ขาดเรียน",value:att.absent,color:T.red},{name:"สาย",value:att.late,color:T.yellow},{name:"ลา",value:att.leave,color:T.blue}].filter(d=>d.value>0)
  if(!att.total)return <div style={{padding:"20px 18px",color:"#CBD5E1",fontSize:13,textAlign:"center"}}>ยังไม่มีข้อมูลการเช็คชื่อ</div>
  return <div style={{display:"flex",alignItems:"center",gap:12,padding:"10px 18px 18px"}}><div style={{width:160,height:160,flexShrink:0}}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={items} cx="50%" cy="50%" innerRadius={42} outerRadius={70} dataKey="value" labelLine={false} label={PIE_LABEL}>{items.map((d,i)=><Cell key={i} fill={d.color}/>)}</Pie><Tooltip contentStyle={{fontFamily:"Prompt,sans-serif",fontSize:13,borderRadius:10}}/></PieChart></ResponsiveContainer></div><div style={{display:"flex",flexDirection:"column",gap:7}}>{items.map((d,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:7,fontSize:13}}><span style={{width:10,height:10,borderRadius:"50%",background:d.color,flexShrink:0}}/><span style={{color:"#334155",fontWeight:500}}>{d.name}</span><span style={{color:T.muted,fontSize:12}}>{d.value}คน ({Math.round(d.value/att.total*100)}%)</span></div>)}</div></div>
}

// ── MODAL SHELL ───────────────────────────────────────────────
function ModalShell({title,sub,onClose,footer,children,wide}){
  useEffect(()=>{const h=e=>{if(e.key==="Escape")onClose()};document.addEventListener("keydown",h);return()=>document.removeEventListener("keydown",h)},[onClose])
  return <div style={ss.overlay} onClick={onClose}><div style={{...ss.modal,width:wide||480}} onClick={e=>e.stopPropagation()}><div style={ss.mHead}><div><div style={{fontSize:15,fontWeight:700,color:T.text}}>{title}</div>{sub&&<div style={{fontSize:12,color:T.muted}}>{sub}</div>}</div><button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",padding:4,borderRadius:8}}><X size={18} color={T.muted}/></button></div><div style={ss.mBody}>{children}</div>{footer&&<div style={{padding:"14px 20px",borderTop:`1px solid ${T.border}`,display:"flex",gap:10,flexShrink:0}}>{footer}</div>}</div></div>
}

// ── LOAD STUDENTS FOR CLASS ───────────────────────────────────
async function loadStudentsForClass(sb, classId, className) {
  // 1. Try class_students table
  const { data: links } = await sb.from("class_students")
    .select("student_id,students(*)")
    .eq("class_id", classId)
  if (links && links.length > 0) {
    return links.map(l => l.students).filter(Boolean).sort((a,b) => (a.student_no||999)-(b.student_no||999))
  }
  // 2. Fallback: parse class_name like "ม.1/1"
  const m = (className||"").match(/^(.+)\/(\d+)$/)
  if (m) {
    const { data } = await sb.from("students").select("*")
      .eq("level", m[1]).eq("room", m[2]).order("student_no")
    if (data?.length) return data
  }
  return []
}

// ── SESSION WIZARD ────────────────────────────────────────────
const WIZARD_STEPS = [
  { id:1, label:"ข้อมูลหลัก", icon:"📋" },
  { id:2, label:"เช็คชื่อ",   icon:"✅" },
  { id:3, label:"คะแนน",      icon:"⭐" },
  { id:4, label:"ประเมินผล",  icon:"📊" },
  { id:5, label:"บันทึกหลังสอน", icon:"✍️" },
]

const BEHAVIOR_TAGS = [
  { label:"ตั้งใจเรียน",   pts:1,  icon:"🌟", color:T.green  },
  { label:"ส่งงาน",         pts:1,  icon:"📝", color:T.blue   },
  { label:"ช่วยเพื่อน",    pts:2,  icon:"🤝", color:T.purple },
  { label:"ไม่ตั้งใจ",    pts:-1, icon:"😴", color:T.orange },
  { label:"รบกวนเพื่อน",   pts:-2, icon:"😤", color:T.red    },
]

function SessionWizard({ sb, onClose, onDone, classes, toast, editId }) {
  const [step, setStep]         = useState(1)
  const [cls, setCls]           = useState(classes[0]?.id || "")
  const [info, setInfo]         = useState({ teach_date:new Date().toISOString().split("T")[0], topic:"", objective:"", activities:"" })
  const [students, setStudents] = useState([])
  const [loadingStu, setLoadingStu] = useState(false)
  const [attendance, setAttendance] = useState({})  // {id: 'present'|'absent'|'late'|'leave'}
  const [scores, setScores]     = useState({})       // {id: number}
  const [behaviors, setBehaviors] = useState({})     // {id: string}
  const [assessments, setAssessments] = useState({}) // {id: 'good'|'pass'|'improve'}
  const [post, setPost]         = useState({ highlights:"", problems:"", improvements:"", reflection:"" })
  const [saving, setSaving]     = useState(false)

  const selectedClass = classes.find(c => c.id === cls)

  // Load students when class changes
  useEffect(() => {
    if (!cls) { setStudents([]); return }
    setLoadingStu(true)
    loadStudentsForClass(sb, cls, selectedClass?.class_name)
      .then(studs => {
        setStudents(studs)
        // Default all to present
        const att = {}
        studs.forEach(s => { att[s.id] = "present" })
        setAttendance(att)
      })
      .catch(() => setStudents([]))
      .finally(() => setLoadingStu(false))
  }, [cls])

  const setAtt = (id, val) => setAttendance(p => ({ ...p, [id]:val }))
  const addScore = (id, pts) => setScores(p => ({ ...p, [id]:(p[id]||0)+pts }))
  const setScore = (id, val) => setScores(p => ({ ...p, [id]:val }))
  const setBehavior = (id, val) => setBehaviors(p => ({ ...p, [id]:val }))
  const setAssessment = (id, val) => setAssessments(p => ({ ...p, [id]: p[id]===val ? null : val }))

  // Attendance summary counts
  const attCounts = Object.values(attendance).reduce((acc, v) => { acc[v]=(acc[v]||0)+1; return acc }, {})

  const canNext = () => {
    if (step === 1 && !cls) return false
    return true
  }

  const save = async () => {
    if (!cls) { toast.err("เลือกชั้นเรียนก่อน"); return }
    setSaving(true)
    try {
      // 1. Create session
      const { data: session, error: sErr } = await sb.from("teaching_sessions").insertReturning({
        class_id:cls, ...info,
        highlights:post.highlights, problems:post.problems,
        improvements:post.improvements, reflection:post.reflection,
        status:"saved"
      })
      if (sErr) throw new Error(sErr.message || JSON.stringify(sErr))
      const sessionId = session?.id
      if (!sessionId) throw new Error("ไม่ได้รับ session ID")

      // 2. Attendance
      const attRecs = Object.entries(attendance).map(([sid, status]) => ({ session_id:sessionId, student_id:sid, status }))
      if (attRecs.length) await sb.from("attendance").insert(attRecs)

      // 3. Behavior logs
      const bhRecs = []
      students.forEach(s => {
        const pts = scores[s.id] || 0
        const note = behaviors[s.id] || ""
        if (pts !== 0 || note) {
          bhRecs.push({ session_id:sessionId, student_id:s.id, points:pts, behavior:note, icon:pts > 0 ? "⭐" : pts < 0 ? "⚠️" : "📝" })
        }
      })
      if (bhRecs.length) await sb.from("behavior_logs").insert(bhRecs)

      // 4. Assessments
      const assRecs = Object.entries(assessments).filter(([,v])=>v).map(([sid,level]) => ({ session_id:sessionId, student_id:sid, level }))
      if (assRecs.length) await sb.from("assessments").insert(assRecs)

      toast.ok(`บันทึกการสอนเสร็จสมบูรณ์ ✅ (${students.length} คน)`)
      onDone(); onClose()
    } catch(e) { toast.err("บันทึกไม่สำเร็จ: " + e.message) }
    finally { setSaving(false) }
  }

  const wStyle = {
    position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", zIndex:70,
    display:"flex", alignItems:"center", justifyContent:"center", padding:12
  }
  const boxStyle = {
    background:"#fff", borderRadius:20, display:"flex", flexDirection:"column",
    width:"min(900px,calc(100vw - 24px))", height:"min(660px,calc(100vh - 24px))",
    boxShadow:"0 32px 80px rgba(0,0,0,0.3)", overflow:"hidden"
  }

  return (
    <div style={wStyle} onClick={onClose}>
      <div style={boxStyle} onClick={e=>e.stopPropagation()}>
        {/* Header */}
        <div style={{padding:"16px 24px 0", borderBottom:`1px solid ${T.border}`, flexShrink:0}}>
          <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12}}>
            <div style={{fontSize:16, fontWeight:700, color:T.text}}>
              📖 บันทึกการสอน — {selectedClass ? `${selectedClass.class_name} · ${selectedClass.subject_name||""}` : "เลือกชั้นเรียน"}
            </div>
            <button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",padding:4,borderRadius:8}}><X size={18} color={T.muted}/></button>
          </div>
          {/* Progress steps */}
          <div style={{display:"flex", gap:0}}>
            {WIZARD_STEPS.map((s,i) => {
              const active = step === s.id
              const done = step > s.id
              return (
                <div key={s.id} style={{flex:1, display:"flex", flexDirection:"column", alignItems:"center", paddingBottom:10, cursor:done?"pointer":"default"}}
                  onClick={()=>done&&setStep(s.id)}>
                  <div style={{display:"flex", alignItems:"center", width:"100%"}}>
                    {i > 0 && <div style={{flex:1, height:2, background:done?T.purple:T.border, transition:"all 0.3s"}}/>}
                    <div style={{width:28,height:28,borderRadius:"50%",background:active?T.purple:done?T.purpleL:"#F1F5F9",border:`2px solid ${active||done?T.purple:T.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:active?"#fff":done?T.purple:T.muted,flexShrink:0,transition:"all 0.3s"}}>
                      {done ? <Check size={13}/> : s.id}
                    </div>
                    {i < WIZARD_STEPS.length-1 && <div style={{flex:1, height:2, background:step>s.id?T.purple:T.border, transition:"all 0.3s"}}/>}
                  </div>
                  <div style={{fontSize:11,color:active?T.purple:done?T.purple:T.muted,fontWeight:active?700:500,marginTop:4,textAlign:"center"}}>{s.icon} {s.label}</div>
                </div>
              )
            })}
          </div>
        </div>

        {/* Body */}
        <div style={{flex:1, overflowY:"auto", padding:"16px 24px"}}>
          {/* STEP 1: ข้อมูลหลัก */}
          {step===1 && (
            <div>
              <div style={ss.fRow}>
                <label style={ss.label}>ชั้นเรียน *</label>
                <select value={cls} onChange={e=>setCls(e.target.value)} style={ss.input}>
                  <option value="">-- เลือกชั้นเรียน --</option>
                  {classes.map(c=><option key={c.id} value={c.id}>{c.class_name}{c.subject_name?` · ${c.subject_name}`:""}</option>)}
                </select>
                {!classes.length && <div style={{fontSize:12,color:T.orange,marginTop:4}}>⚠️ ยังไม่มีชั้นเรียน กรุณาเพิ่มก่อน</div>}
              </div>
              <div style={ss.fRow}><label style={ss.label}>วันที่สอน *</label><input type="date" value={info.teach_date} onChange={e=>setInfo(p=>({...p,teach_date:e.target.value}))} style={ss.input}/></div>
              <div style={ss.fRow}><label style={ss.label}>เรื่องที่สอน</label><input value={info.topic} onChange={e=>setInfo(p=>({...p,topic:e.target.value}))} style={ss.input} placeholder="ระบบสมการเชิงเส้น" autoFocus/></div>
              <div style={ss.fRow}><label style={ss.label}>จุดประสงค์การเรียนรู้</label><textarea value={info.objective} onChange={e=>setInfo(p=>({...p,objective:e.target.value}))} style={ss.textarea} placeholder="นักเรียนสามารถ..."/></div>
              <div style={ss.fRow}><label style={ss.label}>กิจกรรมการเรียนรู้</label><textarea value={info.activities} onChange={e=>setInfo(p=>({...p,activities:e.target.value}))} style={ss.textarea} placeholder="วิธีการสอน / กิจกรรม..."/></div>
            </div>
          )}

          {/* STEP 2: เช็คชื่อ */}
          {step===2 && (
            <div>
              {/* Summary bar */}
              <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
                {Object.entries(ATT_CFG).map(([k,cfg])=>(
                  <div key={k} style={{padding:"5px 12px",borderRadius:99,background:cfg.bg,color:cfg.color,fontSize:12.5,fontWeight:700}}>
                    {cfg.label}: {attCounts[k]||0}
                  </div>
                ))}
                <div style={{padding:"5px 12px",borderRadius:99,background:"#F1F5F9",color:T.muted,fontSize:12.5,fontWeight:600}}>
                  รวม: {students.length} คน
                </div>
              </div>
              {/* Quick actions */}
              <div style={{display:"flex",gap:8,marginBottom:12}}>
                <button onClick={()=>{const a={};students.forEach(s=>{a[s.id]="present"});setAttendance(a)}}
                  style={{...ss.btnP(T.green),padding:"6px 14px",fontSize:12.5}}>✓ มาทั้งหมด</button>
              </div>
              {loadingStu ? <PageLoad text="กำลังโหลดรายชื่อ..."/> : !students.length ? (
                <EmptyState icon="👤" msg="ไม่พบนักเรียนในห้องนี้" btn={<div style={{fontSize:13,color:T.muted}}>ระบบจะค้นหาจากชั้น/ห้องใน class_name (เช่น ม.1/1)</div>}/>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:4}}>
                  {students.map(s => {
                    const status = attendance[s.id] || "present"
                    return (
                      <div key={s.id} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 12px",borderRadius:10,background:ATT_CFG[status].bg+"44",border:`1px solid ${ATT_CFG[status].bg}`}}>
                        <div style={{width:22,fontSize:12.5,color:T.muted,textAlign:"right",flexShrink:0}}>{s.student_no||""}</div>
                        <div style={{flex:1,fontSize:13.5,fontWeight:500,color:T.text,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.full_name}</div>
                        <div style={{display:"flex",gap:4,flexShrink:0}}>
                          {Object.entries(ATT_CFG).map(([k,cfg])=>(
                            <button key={k} onClick={()=>setAtt(s.id,k)}
                              style={{padding:"4px 10px",border:`1.5px solid ${status===k?cfg.color:T.border}`,borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,background:status===k?cfg.color:"#fff",color:status===k?"#fff":cfg.color,transition:"all 0.12s"}}>
                              {cfg.label}
                            </button>
                          ))}
                        </div>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          )}

          {/* STEP 3: คะแนน & พฤติกรรม */}
          {step===3 && (
            <div>
              <div style={{fontSize:13,color:T.muted,marginBottom:12}}>ให้/ลดคะแนน และบันทึกพฤติกรรม (เฉพาะนักเรียนที่ต้องการ)</div>
              {/* Quick behavior tags */}
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14,padding:"10px 12px",background:"#F8FAFC",borderRadius:10}}>
                <span style={{fontSize:12,color:T.muted,alignSelf:"center"}}>แท็กด่วน:</span>
                {BEHAVIOR_TAGS.map(t=>(
                  <span key={t.label} style={{padding:"3px 10px",borderRadius:99,background:t.pts>0?T.greenL:T.redL,color:t.pts>0?T.green:T.red,fontSize:12,fontWeight:600,cursor:"default"}}>
                    {t.icon} {t.label} ({t.pts>0?"+":""}{t.pts})
                  </span>
                ))}
              </div>
              {!students.length ? <EmptyState icon="👤" msg="ไม่มีนักเรียนในห้องนี้"/> : (
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {students.map(s => {
                    const score = scores[s.id] || 0
                    const behavior = behaviors[s.id] || ""
                    const isAbsent = attendance[s.id] === "absent"
                    return (
                      <div key={s.id} style={{padding:"10px 12px",borderRadius:10,border:`1px solid ${T.border}`,background:isAbsent?"#FAFAFA":"#fff",opacity:isAbsent?0.5:1}}>
                        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                          <div style={{width:22,fontSize:12,color:T.muted,textAlign:"right",flexShrink:0}}>{s.student_no||""}</div>
                          <div style={{flex:1,fontSize:13.5,fontWeight:600,color:T.text}}>{s.full_name}</div>
                          {isAbsent && <span style={{fontSize:11,color:T.muted,fontStyle:"italic"}}>ขาดเรียน</span>}
                        </div>
                        {!isAbsent && (
                          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",paddingLeft:32}}>
                            {/* Score buttons */}
                            <div style={{display:"flex",gap:4,alignItems:"center"}}>
                              {[-3,-1,1,3,5].map(pts=>(
                                <button key={pts} onClick={()=>addScore(s.id,pts)}
                                  style={{width:34,height:34,border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12.5,fontWeight:700,background:pts>0?T.greenL:T.redL,color:pts>0?T.green:T.red}}>
                                  {pts>0?`+${pts}`:pts}
                                </button>
                              ))}
                              <div style={{padding:"5px 10px",borderRadius:8,background:score>0?T.greenL:score<0?T.redL:"#F1F5F9",color:score>0?T.green:score<0?T.red:T.muted,fontWeight:700,fontSize:13,minWidth:36,textAlign:"center"}}>
                                {score>0?`+${score}`:score}
                              </div>
                              {score!==0&&<button onClick={()=>setScore(s.id,0)} style={{border:"none",background:"none",cursor:"pointer",color:T.muted,padding:2}}><X size={13}/></button>}
                            </div>
                            {/* Quick tags */}
                            <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                              {BEHAVIOR_TAGS.map(t=>(
                                <button key={t.label} onClick={()=>{addScore(s.id,t.pts);setBehavior(s.id,(behaviors[s.id]||"")+(behaviors[s.id]?", ":"")+t.label)}}
                                  style={{padding:"3px 8px",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:11.5,background:t.pts>0?T.greenL:T.redL,color:t.pts>0?T.green:T.red}}>
                                  {t.icon}{t.label}
                                </button>
                              ))}
                            </div>
                            {/* Behavior text */}
                            <input value={behavior} onChange={e=>setBehavior(s.id,e.target.value)}
                              placeholder="บันทึกพฤติกรรม..."
                              style={{flex:1,minWidth:140,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 10px",fontSize:12.5,fontFamily:"inherit",outline:"none"}}/>
                          </div>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          )}

          {/* STEP 4: ประเมินผล */}
          {step===4 && (
            <div>
              <div style={{fontSize:13,color:T.muted,marginBottom:12}}>ประเมินผลการเรียนรู้ของนักเรียนแต่ละคน (ไม่บังคับ)</div>
              {/* Quick all-pass */}
              <div style={{display:"flex",gap:8,marginBottom:12}}>
                <button onClick={()=>{const a={};students.filter(s=>attendance[s.id]!=="absent").forEach(s=>{a[s.id]="pass"});setAssessments(a)}}
                  style={{...ss.btnP(T.green),padding:"6px 14px",fontSize:12.5}}><Check size={13}/>ผ่านทั้งหมด</button>
                <button onClick={()=>{const a={};students.filter(s=>attendance[s.id]!=="absent").forEach(s=>{a[s.id]="good"});setAssessments(a)}}
                  style={{...ss.btnP(T.purple),padding:"6px 14px",fontSize:12.5}}>⭐ ดีมากทั้งหมด</button>
                <button onClick={()=>setAssessments({})} style={{...ss.btnO,padding:"6px 14px",fontSize:12.5}}><X size={13}/>ล้างทั้งหมด</button>
              </div>
              {!students.length ? <EmptyState icon="👤" msg="ไม่มีนักเรียน"/> : (
                <div style={{display:"flex",flexDirection:"column",gap:5}}>
                  {students.map(s => {
                    const val = assessments[s.id]
                    const isAbsent = attendance[s.id] === "absent"
                    return (
                      <div key={s.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 12px",borderRadius:10,border:`1px solid ${val?ASS_CFG[val]?.bg:T.border}`,background:val?(ASS_CFG[val]?.bg+"44"):"#fff",opacity:isAbsent?0.4:1}}>
                        <div style={{width:22,fontSize:12,color:T.muted,textAlign:"right",flexShrink:0}}>{s.student_no||""}</div>
                        <div style={{flex:1,fontSize:13.5,fontWeight:500,color:T.text}}>{s.full_name}</div>
                        {isAbsent ? <span style={{fontSize:11,color:T.muted}}>ขาดเรียน</span> : (
                          <div style={{display:"flex",gap:5}}>
                            {Object.entries(ASS_CFG).map(([k,cfg])=>(
                              <button key={k} onClick={()=>!isAbsent&&setAssessment(s.id,k)}
                                style={{padding:"5px 12px",border:`1.5px solid ${val===k?cfg.color:T.border}`,borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,background:val===k?cfg.color:"#fff",color:val===k?"#fff":cfg.color,transition:"all 0.12s"}}>
                                {cfg.icon} {cfg.label}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          )}

          {/* STEP 5: บันทึกหลังสอน */}
          {step===5 && (
            <div>
              {/* Summary chips */}
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16,padding:"12px 14px",background:"#F8FAFC",borderRadius:12}}>
                <div style={{fontSize:13,fontWeight:600,color:T.text,width:"100%",marginBottom:4}}>สรุปคาบเรียน</div>
                {Object.entries(ATT_CFG).map(([k,c])=>attCounts[k]>0&&(
                  <span key={k} style={ss.pill(c.color,c.bg)}>{c.label}: {attCounts[k]}</span>
                ))}
                {Object.values(scores).some(v=>v!==0)&&<span style={ss.pill(T.yellow,T.yellowL)}>⭐ มีบันทึกคะแนน {Object.values(scores).filter(v=>v!==0).length} คน</span>}
                {Object.values(assessments).filter(Boolean).length>0&&<span style={ss.pill(T.purple,T.purpleL)}>📊 ประเมินแล้ว {Object.values(assessments).filter(Boolean).length} คน</span>}
              </div>
              <div style={ss.fRow}><label style={ss.label}>✨ จุดเด่นของคาบ</label><textarea value={post.highlights} onChange={e=>setPost(p=>({...p,highlights:e.target.value}))} style={ss.textarea} placeholder="สิ่งที่น่าประทับใจ นักเรียนตอบสนองดี..."/></div>
              <div style={ss.fRow}><label style={ss.label}>⚠️ ปัญหาที่พบ</label><textarea value={post.problems} onChange={e=>setPost(p=>({...p,problems:e.target.value}))} style={ss.textarea} placeholder="ปัญหาที่เกิดขึ้น..."/></div>
              <div style={ss.fRow}><label style={ss.label}>🔧 สิ่งที่ต้องปรับปรุง</label><textarea value={post.improvements} onChange={e=>setPost(p=>({...p,improvements:e.target.value}))} style={ss.textarea} placeholder="จะปรับอะไรในครั้งหน้า..."/></div>
              <div style={ss.fRow}><label style={ss.label}>💭 Reflection</label><textarea value={post.reflection} onChange={e=>setPost(p=>({...p,reflection:e.target.value}))} style={{...ss.textarea,minHeight:80}} placeholder="ความรู้สึกและสิ่งที่เรียนรู้..."/></div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{padding:"14px 24px",borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,background:"#FAFAFA"}}>
          <button onClick={step===1?onClose:()=>setStep(s=>s-1)} style={{...ss.btnO,gap:6}}>
            {step===1?<><X size={15}/>ยกเลิก</>:<><ChevronLeft size={15}/>ย้อนกลับ</>}
          </button>
          <span style={{fontSize:12.5,color:T.muted}}>ขั้นที่ {step} จาก 5</span>
          {step<5 ? (
            <button onClick={()=>canNext()&&setStep(s=>s+1)} disabled={!canNext()} style={{...ss.btnP(),gap:6,opacity:canNext()?1:0.5}}>
              ถัดไป <ChevronRight size={15}/>
            </button>
          ) : (
            <button onClick={save} disabled={saving||!cls} style={{...ss.btnP(T.green),gap:6,opacity:(saving||!cls)?0.6:1}}>
              {saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><Save size={15}/>บันทึกทั้งหมด</>}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}

// ── SESSION DETAIL MODAL ──────────────────────────────────────
function SessionDetailModal({ sb, session, onClose }) {
  const [att, setAtt] = useState([])
  const [beh, setBeh] = useState([])
  const [ass, setAss] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!session?.id) return
    Promise.all([
      sb.from("attendance").select("*,students(student_no,full_name)").eq("session_id", session.id),
      sb.from("behavior_logs").select("*,students(student_no,full_name)").eq("session_id", session.id),
      sb.from("assessments").select("*,students(student_no,full_name)").eq("session_id", session.id),
    ]).then(([a, b, s]) => {
      setAtt(a.data||[]); setBeh(b.data||[]); setAss(s.data||[])
    }).finally(() => setLoading(false))
  }, [session?.id])

  const attCounts = att.reduce((a,v)=>{a[v.status]=(a[v.status]||0)+1;return a},{})
  const assGroups = ass.reduce((a,v)=>{a[v.level]=[...(a[v.level]||[]),v];return a},{})

  return (
    <ModalShell title="รายละเอียดบันทึกการสอน" onClose={onClose} wide={600}
      sub={`${session?.classes?.class_name||""} · ${session?.teach_date?new Date(session.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"long",year:"numeric"}):""}`}>
      {loading ? <PageLoad/> : (
        <div>
          {/* Basic info */}
          {session.topic && <div style={{marginBottom:12}}><span style={{fontWeight:700,color:T.text}}>หัวข้อ: </span>{session.topic}</div>}
          {session.objective && <div style={{marginBottom:12,padding:"10px 14px",background:T.blueL,borderRadius:10,fontSize:13}}><b>จุดประสงค์:</b> {session.objective}</div>}

          {/* Attendance */}
          {att.length > 0 && (
            <div style={{marginBottom:16}}>
              <div style={{fontWeight:700,color:T.text,marginBottom:8}}>การเข้าเรียน ({att.length} คน)</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                {Object.entries(ATT_CFG).map(([k,c])=>attCounts[k]>0&&<span key={k} style={ss.pill(c.color,c.bg)}>{c.label}: {attCounts[k]}</span>)}
              </div>
              {att.filter(a=>a.status!=="present").length > 0 && (
                <div style={{background:"#FFF7ED",borderRadius:10,padding:"10px 14px",fontSize:13}}>
                  <b>ไม่มาเรียน:</b> {att.filter(a=>a.status!=="present").map(a=>`${a.students?.full_name} (${ATT_CFG[a.status]?.label})`).join(", ")}
                </div>
              )}
            </div>
          )}

          {/* Behavior */}
          {beh.length > 0 && (
            <div style={{marginBottom:16}}>
              <div style={{fontWeight:700,color:T.text,marginBottom:8}}>พฤติกรรม/คะแนน ({beh.length} รายการ)</div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {beh.map((b,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 12px",borderRadius:8,background:b.points>0?T.greenL:b.points<0?T.redL:"#F8FAFC",fontSize:13}}>
                    <span>{b.icon||"📝"}</span>
                    <span style={{flex:1}}>{b.students?.student_no}. {b.students?.full_name}</span>
                    {b.points!==0&&<span style={{fontWeight:700,color:b.points>0?T.green:T.red}}>{b.points>0?"+":""}{b.points}</span>}
                    {b.behavior&&<span style={{color:T.muted}}>{b.behavior}</span>}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Assessment */}
          {ass.length > 0 && (
            <div style={{marginBottom:16}}>
              <div style={{fontWeight:700,color:T.text,marginBottom:8}}>ผลการประเมิน ({ass.length} คน)</div>
              {Object.entries(assGroups).map(([level, list])=>(
                <div key={level} style={{marginBottom:6}}>
                  <span style={ss.pill(ASS_CFG[level]?.color||T.muted, ASS_CFG[level]?.bg||"#F8FAFC")}>{ASS_CFG[level]?.icon} {ASS_CFG[level]?.label||level}: {list.length} คน</span>
                </div>
              ))}
            </div>
          )}

          {/* Post teaching */}
          {(session.highlights||session.problems||session.improvements||session.reflection) && (
            <div style={{borderTop:`1px solid ${T.border}`,paddingTop:14}}>
              <div style={{fontWeight:700,color:T.text,marginBottom:10}}>บันทึกหลังสอน</div>
              {[["✨ จุดเด่น",session.highlights],["⚠️ ปัญหา",session.problems],["🔧 ปรับปรุง",session.improvements],["💭 Reflection",session.reflection]].map(([k,v])=>v&&(
                <div key={k} style={{marginBottom:8,padding:"8px 12px",background:"#F8FAFC",borderRadius:8,fontSize:13}}>
                  <b style={{color:T.text}}>{k}: </b>{v}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </ModalShell>
  )
}

// ── OTHER MODALS ──────────────────────────────────────────────
function AddClassModal({sb,onClose,onDone,toast}){
  const[form,setForm]=useState({class_name:"",subject_name:"",color:CLASS_COLORS[0]});const[saving,setSaving]=useState(false)
  const f=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    if(!form.class_name.trim()||!form.subject_name.trim()){toast.err("กรุณากรอกข้อมูลให้ครบ");return}
    setSaving(true)
    try{const{error}=await sb.from("classes").insert({class_name:form.class_name.trim(),subject_name:form.subject_name.trim(),color:form.color});if(error)throw new Error(error.message||JSON.stringify(error));toast.ok(`เพิ่ม ${form.class_name} แล้ว`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <ModalShell title="เพิ่มชั้นเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...ss.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={saving} style={{...ss.btnP(),flex:2,justifyContent:"center",opacity:saving?0.7:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><Check size={15}/>เพิ่มชั้นเรียน</>}</button></>}>
    <div style={ss.fRow}><label style={ss.label}>ชื่อห้อง *</label><input value={form.class_name} onChange={e=>f("class_name",e.target.value)} style={ss.input} placeholder="ม.3/1" autoFocus/></div>
    <div style={ss.fRow}><label style={ss.label}>วิชา *</label><input value={form.subject_name} onChange={e=>f("subject_name",e.target.value)} style={ss.input} placeholder="คณิตศาสตร์"/></div>
    <div style={ss.fRow}><label style={ss.label}>สีประจำห้อง</label><div style={{display:"flex",gap:9,marginTop:5}}>{CLASS_COLORS.map(c=><button key={c} onClick={()=>f("color",c)} style={{width:30,height:30,borderRadius:"50%",background:c,border:form.color===c?"3px solid #1E293B":"3px solid transparent",cursor:"pointer",transition:"all 0.15s"}}/>)}</div></div>
  </ModalShell>
}

function AddStudentModal({sb,onClose,onDone,toast}){
  const[form,setForm]=useState({student_no:"",full_name:"",level:"",room:""});const[saving,setSaving]=useState(false)
  const f=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    if(!form.full_name.trim()){toast.err("กรุณากรอกชื่อ");return}
    setSaving(true)
    try{const{error}=await sb.from("students").insert({full_name:form.full_name.trim(),level:form.level.trim(),room:form.room.trim(),student_no:form.student_no?parseInt(form.student_no):null});if(error)throw new Error(error.message||JSON.stringify(error));toast.ok(`เพิ่ม ${form.full_name} แล้ว`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <ModalShell title="เพิ่มนักเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...ss.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={saving} style={{...ss.btnP(),flex:2,justifyContent:"center",opacity:saving?0.7:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><Plus size={15}/>เพิ่มนักเรียน</>}</button></>}>
    <div style={{display:"grid",gridTemplateColumns:"80px 1fr",gap:12}}><div style={ss.fRow}><label style={ss.label}>เลขที่</label><input type="number" value={form.student_no} onChange={e=>f("student_no",e.target.value)} style={ss.input} placeholder="1"/></div><div style={ss.fRow}><label style={ss.label}>ชื่อ-นามสกุล *</label><input value={form.full_name} onChange={e=>f("full_name",e.target.value)} style={ss.input} placeholder="ชื่อ นามสกุล" autoFocus/></div></div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}><div style={ss.fRow}><label style={ss.label}>ชั้น</label><input value={form.level} onChange={e=>f("level",e.target.value)} style={ss.input} placeholder="ม.3"/></div><div style={ss.fRow}><label style={ss.label}>ห้อง</label><input value={form.room} onChange={e=>f("room",e.target.value)} style={ss.input} placeholder="3/1"/></div></div>
  </ModalShell>
}

function PasteModal({sb,onClose,onDone,toast}){
  const[text,setText]=useState("");const[saving,setSaving]=useState(false)
  const parsed=(()=>{
    if(!text.trim())return[]
    try{return text.trim().split("\n").map((line,i)=>{if(!line.trim())return null;try{const p=line.trim().split(/\t|,/);const ft=(p[0]||"").trim();const isNum=ft!==""&&!isNaN(ft)&&!/^[ก-๙]/.test(ft);return{student_no:isNum?parseInt(ft):i+1,full_name:(isNum?(p[1]||""):(p[0]||"")).trim(),level:(isNum?(p[2]||""):(p[1]||"")).trim(),room:(isNum?(p[3]||""):(p[2]||"")).trim()}}catch{return null}}).filter(s=>s&&s.full_name&&s.full_name.trim().length>0)}catch{return[]}
  })()
  const save=async()=>{
    if(!parsed.length)return;setSaving(true)
    try{const{error}=await sb.from("students").insert(parsed);if(error)throw new Error(error.message||JSON.stringify(error));toast.ok(`นำเข้า ${parsed.length} คนแล้ว`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <ModalShell title="วางรายชื่อนักเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...ss.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={!parsed.length||saving} style={{...ss.btnP(),flex:2,justifyContent:"center",opacity:(!parsed.length||saving)?0.5:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังนำเข้า...</>:<><Upload size={15}/>นำเข้า {parsed.length} คน</>}</button></>}>
    <div style={{background:T.yellowL,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12.5,color:"#78350F",lineHeight:1.7}}><b>รูปแบบ:</b> วางชื่อเดี่ยว หรือ <code>เลขที่ [Tab] ชื่อ [Tab] ชั้น [Tab] ห้อง</code></div>
    <textarea value={text} onChange={e=>setText(e.target.value)} rows={8} placeholder={"เด็กหญิงกมลชนก วิชัย\nเด็กชายขวัญฤดี มีสุข"} style={{...ss.textarea,fontFamily:"monospace",fontSize:12}}/>
    {parsed.length>0&&<div style={{fontSize:13,color:T.green,fontWeight:600,marginTop:6,display:"flex",gap:6,alignItems:"center"}}><CheckCircle size={14}/>พบ {parsed.length} คน</div>}
  </ModalShell>
}

function ImportExcelModal({sb,onClose,onDone,toast}){
  const[dragging,setDragging]=useState(false);const[preview,setPreview]=useState([]);const[saving,setSaving]=useState(false);const ref=useRef()
  const parse=file=>{if(!file)return;const reader=new FileReader();reader.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});const result=rows.slice(1).filter(r=>r[1]&&String(r[1]).trim()).map(r=>({student_no:parseInt(r[0])||null,full_name:String(r[1]||"").trim(),level:String(r[2]||"").trim(),room:String(r[3]||"").trim()}));if(!result.length){toast.err("ไม่พบข้อมูล");return}setPreview(result)}catch(err){toast.err("อ่านไฟล์ไม่ได้")}};reader.onerror=()=>toast.err("อ่านไฟล์ไม่ได้");reader.readAsArrayBuffer(file)}
  const doImport=async()=>{setSaving(true);try{const{error}=await sb.from("students").insert(preview);if(error)throw new Error(error.message||JSON.stringify(error));toast.ok(`นำเข้า ${preview.length} คน`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}}
  return <ModalShell title="นำเข้าจาก Excel" onClose={onClose} footer={<><button onClick={onClose} style={{...ss.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={doImport} disabled={!preview.length||saving} style={{...ss.btnP(T.green),flex:2,justifyContent:"center",opacity:(!preview.length||saving)?0.5:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังนำเข้า...</>:<><Upload size={15}/>นำเข้า {preview.length} คน</>}</button></>}>
    <div onDragEnter={()=>setDragging(true)} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);parse(e.dataTransfer.files[0])}} onDragOver={e=>e.preventDefault()} onClick={()=>ref.current?.click()} style={{border:`2px dashed ${dragging?T.purple:"#CBD5E1"}`,borderRadius:14,padding:"32px 20px",textAlign:"center",marginBottom:14,background:dragging?T.purpleL:"#F8FAFC",cursor:"pointer"}}>
      <Upload size={40} color={T.green} style={{margin:"0 auto 10px"}}/><div style={{fontSize:14,fontWeight:600}}>ลากไฟล์ Excel มาวาง</div><div style={{fontSize:12.5,color:T.muted,marginTop:4}}>หรือคลิกเพื่อเลือก .xlsx, .xls</div>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&parse(e.target.files[0])}/>
    </div>
    {preview.length>0&&<div><div style={{fontSize:13,color:T.green,fontWeight:600,marginBottom:6,display:"flex",gap:6}}><CheckCircle size={14}/>พบ {preview.length} คน</div><div style={{maxHeight:120,overflowY:"auto",border:`1px solid ${T.border}`,borderRadius:10}}>{preview.slice(0,6).map((s,i)=><div key={i} style={{padding:"5px 12px",fontSize:12.5,borderBottom:`1px solid #F8FAFC`}}>{s.student_no}. {s.full_name}</div>)}{preview.length>6&&<div style={{padding:"5px 12px",fontSize:12,color:T.muted}}>...และอีก {preview.length-6} คน</div>}</div></div>}
  </ModalShell>
}

function ProfileModal({sb,onClose,profile,onDone,toast,onDisconnect}){
  const[form,setForm]=useState({teacher_name:profile?.teacher_name||"",school_name:profile?.school_name||"",subject_group:profile?.subject_group||""});const[saving,setSaving]=useState(false)
  const sf=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    setSaving(true)
    try{if(profile?.id){await sb.from("profiles").update(form).eq("id",profile.id)}else{await sb.from("profiles").insert(form)};toast.ok("บันทึกข้อมูลแล้ว");onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <ModalShell title="ตั้งค่า" onClose={onClose} footer={<><button onClick={onClose} style={{...ss.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={saving} style={{...ss.btnP(),flex:2,justifyContent:"center",opacity:saving?0.7:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><Save size={15}/>บันทึก</>}</button></>}>
    <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>ข้อมูลครู</div>
    <div style={ss.fRow}><label style={ss.label}>ชื่อ-นามสกุลครู</label><input value={form.teacher_name} onChange={e=>sf("teacher_name",e.target.value)} style={ss.input} placeholder="ครูจิราภรณ์ รักเด็ก" autoFocus/></div>
    <div style={ss.fRow}><label style={ss.label}>โรงเรียน</label><input value={form.school_name} onChange={e=>sf("school_name",e.target.value)} style={ss.input} placeholder="โรงเรียนบ้านดอน"/></div>
    <div style={{...ss.fRow,marginBottom:20}}><label style={ss.label}>กลุ่มสาระ</label><input value={form.subject_group} onChange={e=>sf("subject_group",e.target.value)} style={ss.input} placeholder="คณิตศาสตร์"/></div>
    <div style={{borderTop:`1px solid ${T.border}`,paddingTop:16}}>
      <button onClick={onDisconnect} style={{...ss.btnO,width:"100%",justifyContent:"center",color:T.red,borderColor:"#FECACA"}}><WifiOff size={15}/>เปลี่ยน Supabase Project</button>
    </div>
  </ModalShell>
}

// ── MENU ──────────────────────────────────────────────────────
const MENU=[
  {id:"dashboard",icon:LayoutDashboard,label:"หน้าหลัก"},
  {id:"classes",  icon:School,         label:"ชั้นเรียนของฉัน"},
  {id:"students", icon:Users,          label:"นักเรียน"},
  {id:"session",  icon:BookOpen,       label:"บันทึกการสอน"},
  {id:"reports",  icon:BarChart3,      label:"รายงานสรุป"},
  {id:"settings", icon:Settings,       label:"ตั้งค่า",     action:"profile"},
  {id:"backup",   icon:Archive,        label:"สำรองข้อมูล", action:"backup"},
  {id:"reset",    icon:RotateCcw,      label:"รีเซทระบบ",  danger:true, action:"reset"},
]

// ── DASHBOARD ─────────────────────────────────────────────────
function Dashboard({sb,menu,setMenu,search,setSearch,filterLv,setFilterLv,toast,modal,setModal,onDisconnect}){
  const dashQ=useDash(sb); const sessQ=useRecentSessions(sb); const clsQ=useClasses(sb)
  const stuQ=useStudents(sb,search,filterLv); const allSessQ=useAllSessions(sb); const profQ=useProfile(sb)
  const [viewSession, setViewSession] = useState(null)

  const stats=dashQ.data||{students:0,classes:0,sessions:0,att:{present:0,absent:0,late:0,leave:0,total:0}}
  const sessions=sessQ.data||[]; const allSess=allSessQ.data||[]
  const classes=clsQ.data||[]; const students=stuQ.data||[]; const profile=profQ.data
  const reloadAll=()=>{dashQ.reload();sessQ.reload();clsQ.reload();stuQ.reload();allSessQ.reload();profQ.reload()}
  const thaiDate=()=>new Date().toLocaleDateString("th-TH",{year:"numeric",month:"long",day:"numeric",weekday:"long"})

  const handleBackup=async()=>{
    try{
      toast.ok("กำลังเตรียมข้อมูล...")
      const[{data:s},{data:c},{data:se}]=await Promise.all([sb.from("students").select("*").order("level").order("room").order("student_no"),sb.from("classes").select("*"),sb.from("teaching_sessions").select("*,classes(class_name)").order("created_at",{ascending:false})])
      const wb=XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(s||[]),"นักเรียน")
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(c||[]),"ชั้นเรียน")
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(se||[]),"บันทึกการสอน")
      const d=new Date().toLocaleDateString("th-TH",{year:"numeric",month:"2-digit",day:"2-digit"}).replace(/\//g,"-")
      XLSX.writeFile(wb,`MyClass_backup_${d}.xlsx`)
      toast.ok("ดาวน์โหลดสำเร็จ ✅")
    }catch(e){toast.err("สำรองข้อมูลไม่ได้: "+e.message)}
  }

  const handleReset=async()=>{
    if(!window.confirm("⚠️ ลบข้อมูลทั้งหมด?\n\nยืนยันหรือไม่?"))return
    if(!window.confirm("ยืนยันอีกครั้ง — ข้อมูลจะหายถาวร"))return
    try{
      toast.ok("กำลังลบ...")
      for(const t of["attendance","behavior_logs","assessments","session_images","teaching_sessions","class_students","classes","students"]){
        await sb.from(t).delete().neq("id","00000000-0000-0000-0000-000000000000")
      }
      toast.ok("รีเซทเรียบร้อย"); setMenu("dashboard"); reloadAll()
    }catch(e){toast.err("รีเซทไม่สำเร็จ: "+e.message)}
  }

  const handleMenuClick=m=>{
    if(m.action==="profile"){setModal("profile");return}
    if(m.action==="backup"){handleBackup();return}
    if(m.action==="reset"){handleReset();return}
    setMenu(m.id)
  }

  return(
    <>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');@keyframes mcSpin{to{transform:rotate(360deg)}}@keyframes mcSlide{from{transform:translateX(20px);opacity:0}to{transform:translateX(0);opacity:1}}*{box-sizing:border-box}body{margin:0}tr:hover td{background:#FAFAFA}select{appearance:none}`}</style>
      <Toast list={toast.list}/>
      <div style={ss.app}>
        <aside style={ss.sidebar}>
          <div style={{display:"flex",alignItems:"center",gap:10,padding:"18px 16px 14px",borderBottom:`1px solid ${T.border}`}}>
            <div style={{width:34,height:34,borderRadius:10,background:T.purple,display:"flex",alignItems:"center",justifyContent:"center"}}><GraduationCap size={20} color="#fff"/></div>
            <span style={{fontWeight:700,fontSize:17,color:T.text}}>MyClass 🍎</span>
          </div>
          <nav style={ss.nav}>{MENU.map(m=><div key={m.id} style={ss.navI(menu===m.id&&!m.action,m.danger)} onClick={()=>handleMenuClick(m)}><m.icon size={17}/>{m.label}</div>)}</nav>
          <button style={ss.addBtn} onClick={()=>setModal("addClass")}><Plus size={16}/>เพิ่มชั้นเรียน</button>
        </aside>

        <div style={ss.main}>
          <div style={ss.topbar}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:T.green}}/>
              <span style={{fontSize:12.5,color:T.muted,fontWeight:500}}>Supabase Connected</span>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <button onClick={reloadAll} style={{border:"none",background:"none",cursor:"pointer",padding:6,borderRadius:8}}><RefreshCw size={17} color={T.muted}/></button>
              <button onClick={()=>setModal("profile")} style={{display:"flex",alignItems:"center",gap:8,border:"none",background:"none",cursor:"pointer",padding:"4px 8px",borderRadius:10}}>
                <div style={{width:34,height:34,borderRadius:"50%",background:`linear-gradient(135deg,${T.purple},${T.blue})`,display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{color:"#fff",fontSize:13,fontWeight:700}}>{profile?.teacher_name?.[0]||"ค"}</span></div>
                <div style={{textAlign:"left"}}><div style={{fontSize:13.5,fontWeight:600,color:T.text,lineHeight:1.2}}>{profile?.teacher_name||"ตั้งค่าชื่อครู"}</div><div style={{fontSize:11.5,color:T.muted}}>{profile?.school_name||"กดเพื่อตั้งค่า"}</div></div>
                <ChevronDown size={14} color={T.muted}/>
              </button>
            </div>
          </div>

          <div style={ss.content}>

            {/* DASHBOARD */}
            {menu==="dashboard"&&<>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
                <div><div style={{fontSize:22,fontWeight:700,color:T.text,marginBottom:2}}>ยินดีต้อนรับ {profile?.teacher_name||"คุณครู"} 👋</div><div style={{fontSize:13,color:T.muted}}>{thaiDate()}</div></div>
                <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  <button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>เริ่มบันทึกการสอน</button>
                  <button style={ss.btnP(T.blue)} onClick={()=>setMenu("classes")}><School size={15}/>ชั้นเรียนของฉัน</button>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
                {[
                  {icon:p=><Users {...p}/>,label:"นักเรียนทั้งหมด",v:stats.students,u:"คน",bg:T.blueL,ic:T.blue},
                  {icon:p=><School {...p}/>,label:"ชั้นเรียน",v:stats.classes,u:"ห้อง",bg:T.yellowL,ic:T.yellow},
                  {icon:p=><CalendarDays {...p}/>,label:"คาบที่สอน",v:stats.sessions,u:"คาบ",bg:T.purpleL,ic:T.purple},
                  {icon:p=><Star {...p}/>,label:"เข้าเรียน",v:stats.att.total?`${Math.round(stats.att.present/stats.att.total*100)}%`:"—",u:"เฉลี่ย",bg:T.greenL,ic:T.green},
                ].map((it,i)=><div key={i} style={{background:"#fff",borderRadius:14,padding:"16px",border:`1px solid ${T.border}`}}>
                  <div style={{width:42,height:42,borderRadius:12,background:it.bg,display:"flex",alignItems:"center",justifyContent:"center",marginBottom:10}}>{it.icon({size:22,color:it.ic})}</div>
                  {dashQ.loading?<Spin size={24}/>:<><div style={{fontSize:28,fontWeight:700,color:T.text,lineHeight:1}}>{it.v}</div><div style={{fontSize:12.5,color:T.muted,marginTop:4}}>{it.label} · {it.u}</div></>}
                </div>)}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1.4fr 1fr",gap:16,marginBottom:20}}>
                <div style={ss.card}>
                  <div style={ss.cHead}><span style={ss.cTitle}>บันทึกการสอนล่าสุด</span><button onClick={()=>setMenu("session")} style={{border:"none",background:"none",cursor:"pointer",fontSize:12.5,color:T.purple,fontFamily:"inherit",fontWeight:600}}>ดูทั้งหมด →</button></div>
                  {!sessions.length?<EmptyState icon="📋" msg="ยังไม่มีบันทึก"/>:(
                    <table style={ss.table}><thead><tr style={{background:"#F8FAFC"}}>{["ชั้น/วิชา","หัวข้อ","วันที่",""].map((h,i)=><th key={i} style={ss.th}>{h}</th>)}</tr></thead>
                    <tbody>{sessions.map(s=><tr key={s.id} style={{cursor:"pointer"}} onClick={()=>setViewSession(s)}>
                      <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:8,height:8,borderRadius:"50%",background:s.classes?.color||T.purple,flexShrink:0}}/><div><div style={{fontWeight:600,fontSize:13}}>{s.classes?.class_name||"—"}</div><div style={{color:T.muted,fontSize:11.5}}>{s.classes?.subject_name}</div></div></div></td>
                      <td style={{...ss.td,color:"#64748B",fontSize:13,maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.topic||"—"}</td>
                      <td style={{...ss.td,fontSize:12,color:T.muted,whiteSpace:"nowrap"}}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short"}):"—"}</td>
                      <td style={ss.td}><Eye size={14} color={T.muted}/></td>
                    </tr>)}</tbody></table>
                  )}
                </div>
                <div style={ss.card}><div style={ss.cHead}><span style={ss.cTitle}>สรุปการเข้าเรียน</span></div>{!dashQ.loading&&<AttChart att={stats.att}/>}</div>
              </div>
              {classes.length>0&&<><div style={{fontSize:14.5,fontWeight:700,color:T.text,marginBottom:12}}>ชั้นเรียนของฉัน</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14}}>{classes.slice(0,6).map(c=><div key={c.id} style={{background:"#fff",borderRadius:14,border:`1px solid ${T.border}`,overflow:"hidden",cursor:"pointer"}} onClick={()=>setMenu("classes")}><div style={{height:5,background:c.color||T.purple}}/><div style={{padding:"14px"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div><div style={{fontSize:15,fontWeight:700,color:T.text}}>{c.class_name}</div><div style={{fontSize:12.5,color:T.muted}}>{c.subject_name}</div></div><div style={{width:36,height:36,borderRadius:10,background:(c.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,color:c.color||T.purple,fontSize:15}}>{c.class_name?.[0]}</div></div><div style={{fontSize:12.5,color:"#64748B"}}><Users size={12} style={{verticalAlign:"middle"}}/> {c._studentCount||0} คน</div></div></div>)}</div></>}
            </>}

            {/* CLASSES */}
            {menu==="classes"&&<>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
                <div><div style={{fontSize:18,fontWeight:700,color:T.text}}>ชั้นเรียนของฉัน</div><div style={{fontSize:13,color:T.muted}}>{clsQ.loading?"กำลังโหลด...":`${classes.length} ชั้นเรียน`}</div></div>
                <button style={ss.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>เพิ่มชั้นเรียน</button>
              </div>
              {clsQ.loading?<PageLoad/>:!classes.length?<EmptyState icon="🏫" msg="ยังไม่มีชั้นเรียน" btn={<button style={ss.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>สร้างชั้นเรียนแรก</button>}/>:(
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                  {classes.map(c=>(
                    <div key={c.id} style={{background:"#fff",borderRadius:16,border:`1px solid ${T.border}`,overflow:"hidden"}}>
                      <div style={{height:6,background:c.color||T.purple}}/>
                      <div style={{padding:"16px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                          <div><div style={{fontSize:16,fontWeight:700,color:T.text}}>{c.class_name}</div><div style={{fontSize:13,color:T.muted}}>{c.subject_name}</div></div>
                          <div style={{width:40,height:40,borderRadius:12,background:(c.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,color:c.color||T.purple}}>{c.class_name?.[0]}</div>
                        </div>
                        <div style={{fontSize:12.5,color:"#64748B",marginBottom:14}}><Users size={12} style={{verticalAlign:"middle"}}/> {c._studentCount||0} คน</div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                          <button onClick={()=>setModal("session")} style={{padding:"8px 0",background:c.color||T.purple,color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontSize:12.5,fontWeight:600,fontFamily:"inherit"}}>▶ สอน</button>
                          <button onClick={async()=>{if(!window.confirm(`ลบ ${c.class_name}?`))return;const{error}=await sb.from("classes").delete().eq("id",c.id);if(error){toast.err(error.message);return}toast.ok("ลบแล้ว");clsQ.reload();dashQ.reload()}} style={{padding:"8px 0",background:T.redL,color:T.red,border:"none",borderRadius:10,cursor:"pointer",fontSize:12.5,fontWeight:600,fontFamily:"inherit"}}>ลบ</button>
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
                <div style={{flex:1,minWidth:200,position:"relative"}}><Search size={15} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:T.muted}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาชื่อ..." style={{...ss.input,paddingLeft:34}}/></div>
                {["ทั้งหมด","ม.1","ม.2","ม.3","ม.4","ม.5","ม.6"].map(l=><button key={l} onClick={()=>setFilterLv(l)} style={{padding:"7px 12px",border:"1px solid",borderColor:filterLv===l?T.purple:T.border,borderRadius:10,cursor:"pointer",background:filterLv===l?T.purpleL:"#fff",color:filterLv===l?T.purple:"#64748B",fontSize:12.5,fontWeight:600,fontFamily:"inherit"}}>{l}</button>)}
              </div>
              {stuQ.loading?<PageLoad/>:!students.length?<EmptyState icon="👩‍🎓" msg="ไม่พบนักเรียน" btn={<button style={ss.btnP()} onClick={()=>setModal("addStu")}><Plus size={15}/>เพิ่มนักเรียน</button>}/>:(
                <div style={ss.card}><table style={ss.table}><thead><tr style={{background:"#F8FAFC"}}>{["#","ชื่อ-นามสกุล","ชั้น","ห้อง",""].map((h,i)=><th key={i} style={ss.th}>{h}</th>)}</tr></thead>
                <tbody>{students.map((stu,i)=><tr key={stu.id}><td style={{...ss.td,color:T.muted,fontSize:12.5,width:40}}>{stu.student_no||i+1}</td><td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:30,height:30,borderRadius:"50%",background:T.blueL,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:T.blue,flexShrink:0}}>{stu.full_name?.[0]}</div>{stu.full_name}</div></td><td style={ss.td}>{stu.level}</td><td style={ss.td}>{stu.room}</td><td style={{...ss.td,width:50}}><button onClick={async()=>{if(!window.confirm(`ลบ ${stu.full_name}?`))return;const{error}=await sb.from("students").delete().eq("id",stu.id);if(error){toast.err(error.message);return}toast.ok("ลบแล้ว");stuQ.reload();dashQ.reload()}} style={{padding:"5px 8px",border:"1px solid #FECACA",borderRadius:8,cursor:"pointer",background:T.redL,color:T.red}}><Trash2 size={13}/></button></td></tr>)}</tbody></table>
                <div style={{padding:"10px 14px",background:"#F8FAFC",fontSize:12.5,color:T.muted,borderTop:`1px solid #F1F5F9`}}>แสดง {students.length} คน</div></div>
              )}
            </>}

            {/* SESSION LIST */}
            {menu==="session"&&<>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
                <div><div style={{fontSize:18,fontWeight:700,color:T.text}}>บันทึกการสอน</div><div style={{fontSize:13,color:T.muted}}>{allSessQ.loading?"กำลังโหลด...":`${allSess.length} บันทึก`}</div></div>
                <button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>เริ่มบันทึกการสอน</button>
              </div>
              {allSessQ.loading?<PageLoad/>:!allSess.length?<EmptyState icon="📋" msg="ยังไม่มีบันทึก" btn={<button style={ss.btnP()} onClick={()=>setModal("session")}><Plus size={15}/>สร้างบันทึกแรก</button>}/>:(
                <div style={ss.card}><table style={ss.table}>
                  <thead><tr style={{background:"#F8FAFC"}}>{["ชั้น/วิชา","เรื่องที่สอน","วันที่","จุดประสงค์","สถานะ",""].map((h,i)=><th key={i} style={ss.th}>{h}</th>)}</tr></thead>
                  <tbody>{allSess.map(s=>(
                    <tr key={s.id} style={{cursor:"pointer"}} onClick={()=>setViewSession(s)}>
                      <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:8,height:8,borderRadius:"50%",background:s.classes?.color||T.purple,flexShrink:0}}/><div><div style={{fontWeight:600,fontSize:13}}>{s.classes?.class_name||"—"}</div><div style={{color:T.muted,fontSize:11.5}}>{s.classes?.subject_name}</div></div></div></td>
                      <td style={{...ss.td,maxWidth:160}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:13}}>{s.topic||"—"}</div></td>
                      <td style={{...ss.td,whiteSpace:"nowrap",color:T.muted,fontSize:12}}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short",year:"2-digit"}):"—"}</td>
                      <td style={{...ss.td,maxWidth:160}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:12,color:"#64748B"}}>{s.objective||"—"}</div></td>
                      <td style={ss.td}>{s.status==="saved"?<span style={ss.pill("#15803D",T.greenL)}><Check size={11}/>สมบูรณ์</span>:<span style={ss.pill("#C2410C",T.orangeL)}><Clock size={11}/>ร่าง</span>}</td>
                      <td style={ss.td}><Eye size={14} color={T.muted}/></td>
                    </tr>
                  ))}</tbody>
                </table></div>
              )}
            </>}

            {menu==="reports"&&<div style={{textAlign:"center",paddingTop:60}}>
              <div style={{fontSize:60,marginBottom:16}}>📊</div>
              <div style={{fontSize:18,fontWeight:600,color:T.muted,marginBottom:8}}>รายงานสรุป</div>
              <div style={{fontSize:13,color:"#CBD5E1",marginBottom:20}}>อยู่ระหว่างพัฒนา</div>
              <button style={ss.btnP()} onClick={handleBackup}><Download size={15}/>ดาวน์โหลดข้อมูล (Excel)</button>
            </div>}

            {!["dashboard","classes","students","session","reports"].includes(menu)&&<div style={{textAlign:"center",paddingTop:60}}><div style={{fontSize:60,marginBottom:16}}>🚧</div><div style={{fontSize:18,fontWeight:600,color:T.muted}}>{MENU.find(m=>m.id===menu)?.label}</div></div>}

          </div>
        </div>

        {/* Modals */}
        {modal==="addClass"  &&<AddClassModal   sb={sb} onClose={()=>setModal(null)} onDone={()=>{clsQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="addStu"    &&<AddStudentModal  sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="paste"     &&<PasteModal       sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="excel"     &&<ImportExcelModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
        {modal==="session"   &&<SessionWizard    sb={sb} onClose={()=>setModal(null)} onDone={()=>{allSessQ.reload();sessQ.reload();dashQ.reload()}} classes={classes} toast={toast}/>}
        {modal==="profile"   &&<ProfileModal     sb={sb} onClose={()=>setModal(null)} profile={profQ.data} onDone={profQ.reload} toast={toast} onDisconnect={onDisconnect}/>}
        {viewSession         &&<SessionDetailModal sb={sb} session={viewSession} onClose={()=>setViewSession(null)}/>}
      </div>
    </>
  )
}

// ── AUTO-CONNECT ──────────────────────────────────────────────
const ENV_URL = import.meta.env.VITE_SUPABASE_URL || ""
const ENV_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || ""

export default function App(){
  const initUrl = ENV_URL || store.get("mc_url") || ""
  const initKey = ENV_KEY || store.get("mc_key") || ""
  const autoConn = !!(initUrl && initKey)
  const[connected,setConnected]=useState(autoConn)
  const[sb,setSb]=useState(autoConn?createClient(initUrl,initKey):null)
  const[menu,setMenu]=useState("dashboard")
  const[search,setSearch]=useState("")
  const[filterLv,setFilterLv]=useState("ทั้งหมด")
  const[modal,setModal]=useState(null)
  const toast=useToast()
  const handleReady=(url,key)=>{store.set("mc_url",url);store.set("mc_key",key);setSb(createClient(url,key));setConnected(true)}
  const handleDisconnect=()=>{if(ENV_URL&&ENV_KEY){window.location.reload();return}store.remove("mc_url");store.remove("mc_key");setSb(null);setConnected(false)}
  if(!connected||!sb)return <SetupScreen onReady={handleReady}/>
  return <ErrorBoundary><Dashboard sb={sb} menu={menu} setMenu={setMenu} search={search} setSearch={setSearch} filterLv={filterLv} setFilterLv={setFilterLv} toast={toast} modal={modal} setModal={setModal} onDisconnect={handleDisconnect}/></ErrorBoundary>
}