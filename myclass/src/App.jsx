import { useState, useEffect, useCallback, useRef, useMemo } from "react"
import React from "react"
import * as XLSX from "xlsx"
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts"
import {
  LayoutDashboard, School, Users, BookOpen, BarChart3, Settings, Archive,
  RotateCcw, Plus, X, Check, Clock, Upload, Trash2, RefreshCw, GraduationCap,
  AlertCircle, CheckCircle, WifiOff, Search, Save, ChevronDown, Star,
  CalendarDays, ClipboardList, Download, ChevronRight, ChevronLeft, Bell,
  UserCheck, Award, Zap, Eye, Home, BookMarked, LogOut, ChevronUp
} from "lucide-react"

const _s={}
const store={get:k=>_s[k]??null,set:(k,v)=>{_s[k]=v},remove:k=>{delete _s[k]}}
/* ─── MOBILE HOOK ─────────────────────────────────────────── */
function useIsMobile(){
  const[m,setM]=useState(()=>typeof window!=="undefined"&&window.innerWidth<768)
  useEffect(()=>{
    const h=()=>setM(window.innerWidth<768)
    window.addEventListener("resize",h)
    return()=>window.removeEventListener("resize",h)
  },[])
  return m
}


// ── SUPABASE CLIENT ───────────────────────────────────────────
function createClient(baseUrl,apiKey){
  const url=baseUrl.replace(/\/$/,"")
  const h={apikey:apiKey,Authorization:`Bearer ${apiKey}`,"Content-Type":"application/json"}
  function qb(table){
    let _sel="*",_filters=[],_orders=[],_lim=null,_countMode=null,_isHead=false,_single=false
    const q={
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
  return{
    from(table){return{
      select(f="*",o={}){return qb(table).select(f,o)},
      async insert(data){
        try{
          const r=await fetch(`${url}/rest/v1/${table}`,{method:"POST",headers:{...h,Prefer:"return=minimal"},body:JSON.stringify(data)})
          if(r.status===201||r.status===204)return{data:null,error:null}
          const b=await r.json().catch(()=>({}));return r.ok?{data:b,error:null}:{data:null,error:b}
        }catch(e){return{data:null,error:{message:e.message}}}
      },
      async insertReturning(data){
        try{
          const r=await fetch(`${url}/rest/v1/${table}`,{method:"POST",headers:{...h,Prefer:"return=representation"},body:JSON.stringify(data)})
          const b=await r.json().catch(()=>({}))
          if(!r.ok)return{data:null,error:b}
          return{data:Array.isArray(b)?(b[0]??null):(b??null),error:null}
        }catch(e){return{data:null,error:{message:e.message}}}
      },
      update(data){
        const _f=[]
        const u={
          eq(c,v){_f.push(`${c}=eq.${encodeURIComponent(v)}`);return u},
          async then(res){
            try{
              const r=await fetch(`${url}/rest/v1/${table}${_f.length?"?"+_f.join("&"):""}`,{method:"PATCH",headers:{...h,Prefer:"return=minimal"},body:JSON.stringify(data)})
              if(r.status===204||r.status===200){res({data:null,error:null});return}
              const b=await r.json().catch(()=>({}));res(r.ok?{data:b,error:null}:{data:null,error:b})
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
    }}
  }
}

// ── CONSTANTS ─────────────────────────────────────────────────
const T={
  purple:"#7C3AED",purpleL:"#EDE9FE",purpleD:"#5B21B6",
  blue:"#2563EB",blueL:"#DBEAFE",
  green:"#16A34A",greenL:"#DCFCE7",
  red:"#DC2626",redL:"#FEE2E2",
  yellow:"#D97706",yellowL:"#FEF3C7",
  orange:"#EA580C",orangeL:"#FFEDD5",
  teal:"#0891B2",tealL:"#CFFAFE",
  slate:"#F8FAFC",border:"#E2E8F0",
  text:"#1E293B",muted:"#94A3B8",sub:"#64748B",
}
const AVATAR_COLORS=["#EF4444","#F97316","#EAB308","#22C55E","#06B6D4","#8B5CF6","#EC4899","#3B82F6"]
const CLASS_COLORS=["#7C3AED","#2563EB","#16A34A","#D97706","#DC2626","#0891B2","#BE185D","#0F766E"]
const ATT_CFG={
  present:{label:"มา",short:"มา",color:T.green,bg:T.greenL},
  absent:{label:"ขาด",short:"ขาด",color:T.red,bg:T.redL},
  late:{label:"สาย",short:"สาย",color:T.yellow,bg:T.yellowL},
  leave:{label:"ลา",short:"ลา",color:T.blue,bg:T.blueL},
}
const POSITIVE_BEHAVIORS=[
  {id:"attention",label:"ตั้งใจเรียน",icon:"⭐",pts:1,color:"#16A34A"},
  {id:"creative",label:"ความคิดดี",icon:"💡",pts:2,color:"#2563EB"},
  {id:"excellent",label:"ผลงานดี",icon:"🏆",pts:3,color:"#D97706"},
  {id:"homework",label:"ส่งงาน",icon:"📚",pts:1,color:"#7C3AED"},
  {id:"helping",label:"ช่วยเพื่อน",icon:"🤝",pts:2,color:"#0891B2"},
  {id:"answer",label:"ตอบถูก",icon:"🎯",pts:1,color:"#BE185D"},
  {id:"leader",label:"ผู้นำ",icon:"👑",pts:2,color:"#EA580C"},
  {id:"effort",label:"พยายาม",icon:"💪",pts:1,color:"#65A30D"},
]
const NEGATIVE_BEHAVIORS=[
  {id:"sleepy",label:"ไม่ตั้งใจ",icon:"😴",pts:-1,color:"#D97706"},
  {id:"phone",label:"โทรศัพท์",icon:"📵",pts:-1,color:"#DC2626"},
  {id:"disturb",label:"รบกวน",icon:"😤",pts:-2,color:"#B91C1C"},
  {id:"nohw",label:"ไม่ส่งงาน",icon:"✏️",pts:-1,color:"#9A3412"},
  {id:"disrespect",label:"ขาดวินัย",icon:"⚠️",pts:-2,color:"#7F1D1D"},
]
const ASSESS_CFG={
  excellent:{label:"ดีเยี่ยม",icon:"⭐⭐⭐",pass:true,color:T.purple,bg:T.purpleL},
  good:{label:"พอใช้",icon:"⭐⭐",pass:true,color:T.green,bg:T.greenL},
  improve:{label:"ต้องพัฒนา",icon:"⭐",pass:false,color:T.orange,bg:T.orangeL},
}

// ── BASE STYLES ───────────────────────────────────────────────
const b={
  btnP:(c=T.purple)=>({padding:"10px 18px",background:c,color:"#fff",border:"none",borderRadius:12,cursor:"pointer",fontFamily:"inherit",fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6,transition:"all 0.15s"}),
  btnO:{padding:"10px 18px",background:"#fff",color:T.sub,border:`1px solid ${T.border}`,borderRadius:12,cursor:"pointer",fontFamily:"inherit",fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6},
  input:{width:"100%",border:`1px solid ${T.border}`,borderRadius:10,padding:"10px 12px",fontSize:14,fontFamily:"inherit",outline:"none",boxSizing:"border-box"},
  label:{display:"block",fontSize:12.5,fontWeight:700,color:T.sub,marginBottom:5,textTransform:"uppercase",letterSpacing:0.3},
  textarea:{width:"100%",border:`1px solid ${T.border}`,borderRadius:10,padding:"10px 12px",fontSize:14,fontFamily:"inherit",outline:"none",resize:"vertical",minHeight:80,boxSizing:"border-box"},
  card:{background:"#fff",borderRadius:16,border:`1px solid ${T.border}`},
  pill:(c,bg)=>({padding:"3px 10px",borderRadius:99,fontSize:12,fontWeight:700,background:bg,color:c,display:"inline-flex",alignItems:"center",gap:4}),
  fRow:{marginBottom:16},
  modal:{background:"#fff",borderRadius:20,maxWidth:"calc(100vw - 32px)",maxHeight:"90vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 24px 60px rgba(0,0,0,0.25)"},
}

// ── UTILS ─────────────────────────────────────────────────────
function avatarColor(idx){return AVATAR_COLORS[idx%8]}
function avatarLetter(name){return(name||"?")[0]}
function thaiDate(d){return new Date(d||Date.now()).toLocaleDateString("th-TH",{year:"numeric",month:"long",day:"numeric"})}
function todayStr(){return new Date().toISOString().split("T")[0]}

// ── MICRO COMPONENTS ──────────────────────────────────────────
function Spin({size=18,color=T.purple}){return <div style={{width:size,height:size,border:`2.5px solid #E2E8F0`,borderTopColor:color,borderRadius:"50%",animation:"mcSpin 0.7s linear infinite",flexShrink:0}}/>}
function PageLoad({text="กำลังโหลด..."}){return <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:240,gap:12}}><Spin size={34}/><span style={{color:T.muted,fontSize:14}}>{text}</span></div>}
function EmptyState({icon="📭",msg="ไม่มีข้อมูล",sub,btn}){return <div style={{textAlign:"center",padding:"48px 24px"}}><div style={{fontSize:52,marginBottom:12}}>{icon}</div><div style={{fontSize:16,fontWeight:700,color:T.text,marginBottom:sub?6:0}}>{msg}</div>{sub&&<div style={{fontSize:13,color:T.muted,marginBottom:16}}>{sub}</div>}{btn&&<div style={{marginTop:16}}>{btn}</div>}</div>}

function Toast({list}){
  if(!list.length)return null
  return <div style={{position:"fixed",top:12,right:12,zIndex:400,display:"flex",flexDirection:"column",gap:8}}>
    {list.map(t=><div key={t.id} style={{background:t.type==="err"?T.red:T.green,color:"#fff",borderRadius:14,padding:"12px 18px",fontSize:14,fontWeight:600,display:"flex",alignItems:"center",gap:8,boxShadow:"0 8px 32px rgba(0,0,0,0.2)",maxWidth:340,animation:"mcSlide 0.2s ease"}}>
      {t.type==="err"?<AlertCircle size={16}/>:<CheckCircle size={16}/>}{t.msg}
    </div>)}
  </div>
}
function useToast(){
  const[list,setList]=useState([])
  const add=useCallback((msg,type="ok")=>{const id=Date.now();setList(p=>[...p,{id,msg,type}]);setTimeout(()=>setList(p=>p.filter(t=>t.id!==id)),3500)},[])
  return{list,ok:m=>add(m,"ok"),err:m=>add(m,"err")}
}

class ErrorBoundary extends React.Component{
  constructor(p){super(p);this.state={err:null}}
  static getDerivedStateFromError(e){return{err:e}}
  render(){
    if(!this.state.err)return this.props.children
    return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:"#FEF2F2",fontFamily:"'Prompt','Sarabun',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:"36px",maxWidth:460,textAlign:"center",boxShadow:"0 8px 40px rgba(0,0,0,0.1)"}}>
        <div style={{fontSize:60,marginBottom:12}}>⚠️</div>
        <div style={{fontSize:20,fontWeight:700,color:T.text,marginBottom:8}}>เกิดข้อผิดพลาด</div>
        <div style={{fontSize:13,color:T.muted,marginBottom:24,wordBreak:"break-word"}}>{this.state.err?.message}</div>
        <button onClick={()=>window.location.reload()} style={{...b.btnP(),width:"100%",justifyContent:"center",padding:"14px"}}>🔄 โหลดหน้าใหม่</button>
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
  const[sc,cc,sesc,att,drafts]=await Promise.all([
    sb.from("students").select("id",{count:"exact",head:true}),
    sb.from("classes").select("id",{count:"exact",head:true}),
    sb.from("teaching_sessions").select("id",{count:"exact",head:true}),
    sb.from("attendance").select("status"),
    sb.from("teaching_sessions").select("id,teach_date,classes(class_name)").eq("status","draft"),
  ])
  const a=att.data||[]
  return{students:sc.count||0,classes:cc.count||0,sessions:sesc.count||0,
    drafts:drafts.data||[],
    att:{present:a.filter(x=>x.status==="present").length,absent:a.filter(x=>x.status==="absent").length,late:a.filter(x=>x.status==="late").length,leave:a.filter(x=>x.status==="leave").length,total:a.length}}
},[sb])}
function useClasses(sb){return useQ(async()=>{
  const{data,error}=await sb.from("classes").select("*,class_students(*)").order("created_at",{ascending:false})
  if(error)throw new Error(error.message)
  return(data||[]).map(c=>({...c,_studentCount:Array.isArray(c.class_students)?c.class_students.length:0}))
},[sb])}
function useStudents(sb,search,level){return useQ(async()=>{
  let q=sb.from("students").select("*").order("level").order("room").order("student_no")
  if(level&&level!=="ทั้งหมด")q=q.eq("level",level)
  if(search&&search.trim())q=q.ilike("full_name",`%${search.trim()}%`)
  const{data,error}=await q;if(error)throw new Error(error.message);return data||[]
},[sb,search,level])}
function useAllSessions(sb){return useQ(async()=>{
  const{data,error}=await sb.from("teaching_sessions").select("*,classes(class_name,subject_name,color)").order("created_at",{ascending:false})
  if(error)throw new Error(error.message);return data||[]
},[sb])}
function useProfile(sb){return useQ(async()=>{const{data}=await sb.from("profiles").select("*").limit(1).maybeSingle();return data},[sb])}

async function loadClassStudents(sb,classId,className){
  const{data:links}=await sb.from("class_students").select("student_id,students(*)").eq("class_id",classId)
  if(links&&links.length>0)return links.map(l=>l.students).filter(Boolean).sort((a,c)=>(a.student_no||999)-(c.student_no||999))
  const m=(className||"").match(/^(.+)\/(\d+)$/)
  if(m){const{data}=await sb.from("students").select("*").eq("level",m[1]).eq("room",m[2]).order("student_no");if(data?.length)return data}
  return[]
}

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
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');@keyframes mcSpin{to{transform:rotate(360deg)}}@keyframes mcSlide{from{transform:translateY(-10px);opacity:0}to{transform:translateY(0);opacity:1}}*{box-sizing:border-box}body{margin:0}`}</style>
      <div style={{background:"#fff",borderRadius:24,padding:"40px",width:460,maxWidth:"calc(100vw - 32px)",boxShadow:"0 40px 80px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:28}}>
          <div style={{width:56,height:56,borderRadius:18,background:`linear-gradient(135deg,${T.purple},${T.blue})`,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 8px 24px "+T.purple+"44"}}><GraduationCap size={28} color="#fff"/></div>
          <div><div style={{fontSize:22,fontWeight:700,color:T.text}}>MyClass 🍎</div><div style={{fontSize:13,color:T.muted}}>ระบบจัดการห้องเรียนสำหรับครู</div></div>
        </div>
        <div style={b.fRow}><label style={b.label}>Supabase URL</label><input value={url} onChange={e=>setUrl(e.target.value)} style={b.input} placeholder="https://xxxx.supabase.co" type="url" onKeyDown={e=>e.key==="Enter"&&connect()}/></div>
        <div style={b.fRow}><label style={b.label}>Anon Key</label><input value={key} onChange={e=>setKey(e.target.value)} style={{...b.input,fontFamily:"monospace",fontSize:12}} placeholder="eyJhbGci..." onKeyDown={e=>e.key==="Enter"&&connect()}/></div>
        {err&&<div style={{background:T.redL,borderRadius:12,padding:"10px 14px",fontSize:13,color:T.red,marginBottom:16,display:"flex",gap:8}}><AlertCircle size={15} style={{flexShrink:0,marginTop:1}}/>{err}</div>}
        <button onClick={connect} disabled={testing} style={{...b.btnP(),width:"100%",justifyContent:"center",padding:"14px",fontSize:15,opacity:testing?0.75:1}}>
          {testing?<><Spin size={18} color="#fff"/>กำลังเชื่อมต่อ...</>:<><CheckCircle size={18}/>เชื่อมต่อและเริ่มใช้งาน</>}
        </button>
      </div>
    </div>
  )
}

// ── MODAL SHELL ───────────────────────────────────────────────
function Modal({title,sub,onClose,footer,children,wide,fullMobile}){
  const m=useIsMobile()
  useEffect(()=>{const h=e=>{if(e.key==="Escape")onClose()};document.addEventListener("keydown",h);return()=>document.removeEventListener("keydown",h)},[onClose])
  const isBottomSheet=m&&!fullMobile
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:100,display:"flex",alignItems:isBottomSheet?"flex-end":"center",justifyContent:"center",padding:isBottomSheet?"0":"16px"}} onClick={onClose}>
      <div style={{background:"#fff",borderRadius:isBottomSheet?"20px 20px 0 0":"20px",width:"100%",maxWidth:m?"100%":(wide||500),maxHeight:m?"92vh":"92vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 24px 60px rgba(0,0,0,0.25)"}} onClick={e=>e.stopPropagation()}>
        <div style={{padding:"16px 18px",borderBottom:"1px solid "+T.border,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <div><div style={{fontSize:m?17:16,fontWeight:700,color:T.text}}>{title}</div>{sub&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{sub}</div>}</div>
          <button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",padding:8,borderRadius:8,color:T.muted,minWidth:40,minHeight:40,display:"flex",alignItems:"center",justifyContent:"center"}}><X size={20}/></button>
        </div>
        <div style={{padding:"18px",overflowY:"auto",flex:1}}>{children}</div>
        {footer&&<div style={{padding:"14px 18px",borderTop:"1px solid "+T.border,display:"flex",gap:10,flexShrink:0,background:"#FAFAFA"}}>{footer}</div>}
      </div>
    </div>
  )
}

function AddClassModal({sb,onClose,onDone,toast}){
  const[form,setForm]=useState({class_name:"",subject_name:"",color:CLASS_COLORS[0]});const[saving,setSaving]=useState(false)
  const f=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    if(!form.class_name.trim()||!form.subject_name.trim()){toast.err("กรุณากรอกชื่อห้องและวิชา");return}
    setSaving(true)
    try{const{error}=await sb.from("classes").insert({class_name:form.class_name.trim(),subject_name:form.subject_name.trim(),color:form.color});if(error)throw new Error(error.message);toast.ok(`เพิ่ม ${form.class_name} แล้ว`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <Modal title="เพิ่มชั้นเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...b.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={saving} style={{...b.btnP(),flex:2,justifyContent:"center",opacity:saving?0.7:1}}>{saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><Check size={15}/>เพิ่มชั้นเรียน</>}</button></>}>
    <div style={b.fRow}><label style={b.label}>ชื่อห้อง *</label><input value={form.class_name} onChange={e=>f("class_name",e.target.value)} style={b.input} placeholder="ม.3/1" autoFocus/></div>
    <div style={b.fRow}><label style={b.label}>วิชา *</label><input value={form.subject_name} onChange={e=>f("subject_name",e.target.value)} style={b.input} placeholder="คณิตศาสตร์"/></div>
    <div style={b.fRow}><label style={b.label}>สีประจำห้อง</label><div style={{display:"flex",gap:10,marginTop:6}}>{CLASS_COLORS.map(c=><button key={c} onClick={()=>f("color",c)} style={{width:32,height:32,borderRadius:"50%",background:c,border:form.color===c?"3px solid #1E293B":"3px solid transparent",cursor:"pointer"}}/>)}</div></div>
  </Modal>
}

function AddStudentModal({sb,onClose,onDone,toast}){
  const[form,setForm]=useState({student_no:"",full_name:"",level:"",room:""});const[saving,setSaving]=useState(false)
  const f=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    if(!form.full_name.trim()){toast.err("กรุณากรอกชื่อ");return}
    setSaving(true)
    try{const{error}=await sb.from("students").insert({full_name:form.full_name.trim(),level:form.level.trim(),room:form.room.trim(),student_no:form.student_no?parseInt(form.student_no):null});if(error)throw new Error(error.message);toast.ok(`เพิ่ม ${form.full_name} แล้ว`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return <Modal title="เพิ่มนักเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...b.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={saving} style={{...b.btnP(),flex:2,justifyContent:"center",opacity:saving?0.7:1}}>{saving?<><Spin size={15} color="#fff"/>กำลัง...</>:<><Plus size={15}/>เพิ่ม</>}</button></>}>
    <div style={{display:"grid",gridTemplateColumns:"80px 1fr",gap:12}}><div style={b.fRow}><label style={b.label}>เลขที่</label><input type="number" value={form.student_no} onChange={e=>f("student_no",e.target.value)} style={b.input} placeholder="1"/></div><div style={b.fRow}><label style={b.label}>ชื่อ-นามสกุล *</label><input value={form.full_name} onChange={e=>f("full_name",e.target.value)} style={b.input} placeholder="ชื่อ นามสกุล" autoFocus/></div></div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}><div style={b.fRow}><label style={b.label}>ชั้น</label><input value={form.level} onChange={e=>f("level",e.target.value)} style={b.input} placeholder="ม.3"/></div><div style={b.fRow}><label style={b.label}>ห้อง</label><input value={form.room} onChange={e=>f("room",e.target.value)} style={b.input} placeholder="1"/></div></div>
  </Modal>
}

function PasteModal({sb,onClose,onDone,toast}){
  const[text,setText]=useState("");const[saving,setSaving]=useState(false)
  const parsed=(()=>{if(!text.trim())return[];try{return text.trim().split("\n").map((line,i)=>{if(!line.trim())return null;try{const p=line.trim().split(/\t|,/);const ft=(p[0]||"").trim();const isNum=ft!==""&&!isNaN(ft)&&!/^[ก-๙]/.test(ft);return{student_no:isNum?parseInt(ft):i+1,full_name:(isNum?(p[1]||""):(p[0]||"")).trim(),level:(isNum?(p[2]||""):(p[1]||"")).trim(),room:(isNum?(p[3]||""):(p[2]||"")).trim()}}catch{return null}}).filter(s=>s&&s.full_name&&s.full_name.trim().length>0)}catch{return[]}})()
  const save=async()=>{if(!parsed.length)return;setSaving(true);try{const{error}=await sb.from("students").insert(parsed);if(error)throw new Error(error.message);toast.ok(`นำเข้า ${parsed.length} คน`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}}
  return <Modal title="วางรายชื่อนักเรียน" onClose={onClose} footer={<><button onClick={onClose} style={{...b.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={save} disabled={!parsed.length||saving} style={{...b.btnP(),flex:2,justifyContent:"center",opacity:(!parsed.length||saving)?0.5:1}}>{saving?<><Spin size={15} color="#fff"/>นำเข้า...</>:<><Upload size={15}/>นำเข้า {parsed.length} คน</>}</button></>}>
    <div style={{background:T.yellowL,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12.5,color:"#78350F",lineHeight:1.7}}><b>รูปแบบ:</b> วางชื่อเดี่ยว หรือ เลขที่ [Tab] ชื่อ [Tab] ชั้น [Tab] ห้อง</div>
    <textarea value={text} onChange={e=>setText(e.target.value)} rows={8} placeholder={"เด็กหญิงกมลชนก วิชัย\nเด็กชายขวัญฤดี มีสุข"} style={{...b.textarea,fontFamily:"monospace",fontSize:12}}/>
    {parsed.length>0&&<div style={{fontSize:13,color:T.green,fontWeight:600,marginTop:6,display:"flex",gap:6,alignItems:"center"}}><CheckCircle size={14}/>พบ {parsed.length} คน</div>}
  </Modal>
}

function ImportExcelModal({sb,onClose,onDone,toast}){
  const[dragging,setDragging]=useState(false);const[preview,setPreview]=useState([]);const[saving,setSaving]=useState(false);const ref=useRef()
  const parse=file=>{if(!file)return;const reader=new FileReader();reader.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});const result=rows.slice(1).filter(r=>r[1]&&String(r[1]).trim()).map(r=>({student_no:parseInt(r[0])||null,full_name:String(r[1]||"").trim(),level:String(r[2]||"").trim(),room:String(r[3]||"").trim()}));if(!result.length){toast.err("ไม่พบข้อมูล");return}setPreview(result)}catch{toast.err("อ่านไฟล์ไม่ได้")}};reader.onerror=()=>toast.err("อ่านไฟล์ไม่ได้");reader.readAsArrayBuffer(file)}
  const doImport=async()=>{setSaving(true);try{const{error}=await sb.from("students").insert(preview);if(error)throw new Error(error.message);toast.ok(`นำเข้า ${preview.length} คน`);onDone();onClose()}catch(e){toast.err(e.message)}finally{setSaving(false)}}
  return <Modal title="นำเข้าจาก Excel" onClose={onClose} footer={<><button onClick={onClose} style={{...b.btnO,flex:1,justifyContent:"center"}}>ยกเลิก</button><button onClick={doImport} disabled={!preview.length||saving} style={{...b.btnP(T.green),flex:2,justifyContent:"center",opacity:(!preview.length||saving)?0.5:1}}>{saving?<><Spin size={15} color="#fff"/>นำเข้า...</>:<><Upload size={15}/>นำเข้า {preview.length} คน</>}</button></>}>
    <div onDragEnter={()=>setDragging(true)} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);parse(e.dataTransfer.files[0])}} onDragOver={e=>e.preventDefault()} onClick={()=>ref.current?.click()} style={{border:`2px dashed ${dragging?T.purple:"#CBD5E1"}`,borderRadius:14,padding:"28px",textAlign:"center",marginBottom:14,background:dragging?T.purpleL:"#F8FAFC",cursor:"pointer"}}>
      <Upload size={36} color={T.green} style={{margin:"0 auto 8px"}}/><div style={{fontSize:14,fontWeight:600}}>ลากไฟล์ Excel มาวาง</div>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&parse(e.target.files[0])}/>
    </div>
    {preview.length>0&&<div style={{fontSize:13,color:T.green,fontWeight:600,display:"flex",gap:6}}><CheckCircle size={14}/>พบ {preview.length} คน</div>}
  </Modal>
}

// ── BEHAVIOR BOTTOM SHEET ─────────────────────────────────────
function BehaviorSheet({student,score,logs,onBehavior,onClose}){
  const recent=logs.filter(l=>l.studentId===student?.id).slice(-3)
  return(
    <div style={{position:"fixed",inset:0,zIndex:200,background:"rgba(0,0,0,0.6)"}} onClick={onClose}>
      <div style={{position:"absolute",bottom:0,left:0,right:0,background:"#fff",borderRadius:"24px 24px 0 0",padding:"20px 16px 28px",maxHeight:"85vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
        {/* Handle */}
        <div style={{width:40,height:4,background:"#E2E8F0",borderRadius:99,margin:"0 auto 16px"}}/>
        {/* Student header */}
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16}}>
          <div style={{width:52,height:52,borderRadius:"50%",background:avatarColor(student?.student_no||0),display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,fontWeight:700,color:"#fff",flexShrink:0}}>{avatarLetter(student?.full_name)}</div>
          <div style={{flex:1}}>
            <div style={{fontSize:16,fontWeight:700,color:T.text}}>{student?.full_name}</div>
            <div style={{fontSize:13,color:T.muted}}>เลขที่ {student?.student_no||"—"}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:28,fontWeight:800,color:score>0?T.green:score<0?T.red:T.muted,lineHeight:1}}>{score>0?`+${score}`:score===0?"0":score}</div>
            <div style={{fontSize:11,color:T.muted}}>คะแนนคาบนี้</div>
          </div>
        </div>
        {/* Recent */}
        {recent.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14}}>
          {recent.map((l,i)=><span key={i} style={{padding:"2px 8px",borderRadius:99,background:l.pts>0?T.greenL:T.redL,color:l.pts>0?T.green:T.red,fontSize:11.5,fontWeight:600}}>{l.icon} {l.label}</span>)}
        </div>}
        {/* Positive */}
        <div style={{fontSize:13,fontWeight:700,color:T.green,marginBottom:10,display:"flex",alignItems:"center",gap:6}}><ChevronUp size={14}/>พฤติกรรมเชิงบวก</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:16}}>
          {POSITIVE_BEHAVIORS.map(bv=>(
            <button key={bv.id} onClick={()=>onBehavior(bv)} style={{padding:"10px 4px",border:`2px solid ${bv.color}20`,borderRadius:14,cursor:"pointer",fontFamily:"inherit",background:bv.color+"15",display:"flex",flexDirection:"column",alignItems:"center",gap:4,transition:"all 0.12s"}}>
              <span style={{fontSize:24}}>{bv.icon}</span>
              <span style={{fontSize:11,fontWeight:700,color:bv.color,textAlign:"center",lineHeight:1.2}}>{bv.label}</span>
              <span style={{fontSize:12,fontWeight:800,color:T.green}}>+{bv.pts}</span>
            </button>
          ))}
        </div>
        {/* Negative */}
        <div style={{fontSize:13,fontWeight:700,color:T.red,marginBottom:10,display:"flex",alignItems:"center",gap:6}}><ChevronDown size={14}/>พฤติกรรมที่ต้องปรับปรุง</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8}}>
          {NEGATIVE_BEHAVIORS.map(bv=>(
            <button key={bv.id} onClick={()=>onBehavior(bv)} style={{padding:"10px 4px",border:"2px solid #FEE2E2",borderRadius:14,cursor:"pointer",fontFamily:"inherit",background:"#FFF5F5",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
              <span style={{fontSize:22}}>{bv.icon}</span>
              <span style={{fontSize:10,fontWeight:700,color:T.red,textAlign:"center",lineHeight:1.2}}>{bv.label}</span>
              <span style={{fontSize:12,fontWeight:800,color:T.red}}>{bv.pts}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  )
}

// ── CLASSROOM SESSION ─────────────────────────────────────────
const STEPS=["select","attend","live","assess","log"]
const STEP_LABELS=["เลือกห้อง","เช็คชื่อ","สอน","ประเมิน","บันทึก"]
const STEP_ICONS=["🏫","✅","⚡","📊","✍️"]

function ClassroomSession({sb,classes,toast,onExit,onReload}){
  const[step,setStep]=useState("select")
  const[selectedClass,setSelectedClass]=useState(null)
  const[students,setStudents]=useState([])
  const[loadingStu,setLoadingStu]=useState(false)
  const[attendance,setAttendance]=useState({})
  const[behaviorLogs,setBehaviorLogs]=useState([])
  const[assessments,setAssessments]=useState({})
  const[assessNotes,setAssessNotes]=useState({})
  const[activeStudent,setActiveStudent]=useState(null)
  const[teachLog,setTeachLog]=useState({teach_date:todayStr(),topic:"",objective:"",activities:"",highlights:"",problems:"",improvements:"",reflection:""})
  const[saving,setSaving]=useState(false)
  const[logMode,setLogMode]=useState(null) // 'draft' | 'full'

  const scores=useMemo(()=>behaviorLogs.reduce((acc,l)=>{acc[l.studentId]=(acc[l.studentId]||0)+l.pts;return acc},{}),[behaviorLogs])
  const attCounts=useMemo(()=>Object.values(attendance).reduce((a,v)=>{a[v]=(a[v]||0)+1;return a},{}),[attendance])
  const passCounts=useMemo(()=>{
    const pass=Object.values(assessments).filter(v=>v&&ASSESS_CFG[v]?.pass).length
    const fail=Object.values(assessments).filter(v=>v&&!ASSESS_CFG[v]?.pass).length
    return{pass,fail,total:pass+fail}
  },[assessments])

  const loadStudents=async(cls)=>{
    setLoadingStu(true)
    try{
      const studs=await loadClassStudents(sb,cls.id,cls.class_name)
      setStudents(studs)
      const att={}
      studs.forEach(s=>{att[s.id]="present"})
      setAttendance(att)
    }catch{setStudents([])}
    finally{setLoadingStu(false)}
  }

  const selectClass=async(cls)=>{
    setSelectedClass(cls)
    await loadStudents(cls)
    setStep("attend")
  }

  const addBehavior=(student,bv)=>{
    setBehaviorLogs(p=>[...p,{studentId:student.id,behavior_id:bv.id,label:bv.label,icon:bv.icon,pts:bv.pts,color:bv.color,time:Date.now()}])
  }

  const handleBehavior=bv=>{
    if(activeStudent){addBehavior(activeStudent,bv)}
  }

  const canProceed=()=>{
    if(step==="log"){
      // Check: need reason for improve
      const needReason=Object.entries(assessments).filter(([,v])=>v==="improve").some(([sid])=>!assessNotes[sid]?.trim())
      return !needReason
    }
    return true
  }

  const saveSession=async(status="saved")=>{
    setSaving(true)
    try{
      const{data:session,error}=await sb.from("teaching_sessions").insertReturning({
        class_id:selectedClass.id,...teachLog,status
      })
      if(error)throw new Error(error.message||JSON.stringify(error))
      const sessionId=session?.id
      if(!sessionId)throw new Error("ไม่ได้รับ session ID")

      const attRecs=Object.entries(attendance).map(([sid,st])=>({session_id:sessionId,student_id:sid,status:st}))
      if(attRecs.length)await sb.from("attendance").insert(attRecs)

      const bhMap={}
      behaviorLogs.forEach(l=>{
        if(!bhMap[l.studentId])bhMap[l.studentId]={pts:0,labels:[],icons:[]}
        bhMap[l.studentId].pts+=l.pts
        bhMap[l.studentId].labels.push(l.label)
        bhMap[l.studentId].icons.push(l.icon)
      })
      const bhRecs=Object.entries(bhMap).map(([sid,d])=>({session_id:sessionId,student_id:sid,points:d.pts,behavior:d.labels.join(", "),icon:d.icons[0]||"📝"}))
      if(bhRecs.length)await sb.from("behavior_logs").insert(bhRecs)

      const assRecs=Object.entries(assessments).filter(([,v])=>v).map(([sid,lv])=>({session_id:sessionId,student_id:sid,level:lv,note:assessNotes[sid]||""}))
      if(assRecs.length)await sb.from("assessments").insert(assRecs)

      toast.ok(status==="saved"?"บันทึกการสอนเสร็จสมบูรณ์ ✅":"บันทึกร่างแล้ว")
      onReload()
      onExit()
    }catch(e){toast.err("บันทึกไม่สำเร็จ: "+e.message)}
    finally{setSaving(false)}
  }

  const stepIdx=STEPS.indexOf(step)

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100%",background:T.slate}}>
      {/* Step indicator */}
      <div style={{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"10px 16px",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
          <div style={{fontSize:15,fontWeight:700,color:T.text}}>
            {selectedClass?`${selectedClass.class_name} · ${selectedClass.subject_name||""}`:step==="select"?"เลือกชั้นเรียน":"ห้องเรียน"}
          </div>
          <button onClick={()=>{if(!window.confirm("ออกจากการสอน? ข้อมูลที่กรอกจะหาย"))return;onExit()}} style={{...b.btnO,padding:"6px 12px",fontSize:12.5,gap:5}}><LogOut size={13}/>ออก</button>
        </div>
        <div style={{display:"flex",gap:0}}>
          {STEPS.map((s,i)=>{
            const done=stepIdx>i;const active=stepIdx===i
            return(
              <div key={s} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",cursor:done?"pointer":"default"}} onClick={()=>done&&setStep(s)}>
                <div style={{display:"flex",alignItems:"center",width:"100%"}}>
                  {i>0&&<div style={{flex:1,height:2,background:done?T.purple:T.border}}/>}
                  <div style={{width:26,height:26,borderRadius:"50%",background:active?T.purple:done?T.purpleL:"#F1F5F9",border:`2px solid ${active||done?T.purple:T.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:active?"#fff":done?T.purple:T.muted,flexShrink:0}}>
                    {done?<Check size={12}/>:i+1}
                  </div>
                  {i<STEPS.length-1&&<div style={{flex:1,height:2,background:stepIdx>i?T.purple:T.border}}/>}
                </div>
                <div style={{fontSize:10,color:active?T.purple:done?T.purple:T.muted,fontWeight:active?700:500,marginTop:3}}>{STEP_ICONS[i]}</div>
              </div>
            )
          })}
        </div>
      </div>

      {/* Content */}
      <div style={{flex:1,overflowY:"auto"}}>

        {/* STEP: select */}
        {step==="select"&&(
          <div style={{padding:"20px 16px"}}>
            <div style={{fontSize:16,fontWeight:700,color:T.text,marginBottom:4}}>เลือกชั้นเรียนที่จะสอน</div>
            <div style={{fontSize:13,color:T.muted,marginBottom:16}}>กดเพื่อเริ่มบันทึกคาบเรียน</div>
            {!classes.length?<EmptyState icon="🏫" msg="ยังไม่มีชั้นเรียน"/>:(
              <div style={{display:"flex",flexDirection:"column",gap:12}}>
                {classes.map(cls=>(
                  <button key={cls.id} onClick={()=>selectClass(cls)} style={{background:"#fff",border:`2px solid ${T.border}`,borderRadius:16,padding:"16px",cursor:"pointer",fontFamily:"inherit",textAlign:"left",display:"flex",alignItems:"center",gap:14,transition:"all 0.15s"}}>
                    <div style={{width:52,height:52,borderRadius:14,background:(cls.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:800,color:cls.color||T.purple,flexShrink:0}}>{cls.class_name?.[0]}</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:18,fontWeight:700,color:T.text}}>{cls.class_name}</div>
                      <div style={{fontSize:13,color:T.muted}}>{cls.subject_name} · {cls._studentCount||0} คน</div>
                    </div>
                    <div style={{width:36,height:36,borderRadius:"50%",background:cls.color||T.purple,display:"flex",alignItems:"center",justifyContent:"center"}}><ChevronRight size={18} color="#fff"/></div>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        {/* STEP: attend */}
        {step==="attend"&&(
          <div style={{padding:"16px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
              <div><div style={{fontSize:15,fontWeight:700,color:T.text}}>เช็คชื่อผู้เรียน</div><div style={{fontSize:12.5,color:T.muted}}>กดเปลี่ยนสถานะได้เลย</div></div>
              <div style={{display:"flex",gap:6}}>
                {Object.entries(ATT_CFG).map(([k,c])=>(
                  <div key={k} style={{padding:"4px 10px",borderRadius:99,background:c.bg,color:c.color,fontSize:12,fontWeight:700}}>{c.short}: {attCounts[k]||0}</div>
                ))}
              </div>
            </div>
            <button onClick={()=>{const a={};students.forEach(s=>{a[s.id]="present"});setAttendance(a)}} style={{...b.btnP(T.green),padding:"7px 14px",fontSize:12.5,marginBottom:12}}><Check size={13}/>มาทั้งหมด</button>
            {loadingStu?<PageLoad text="กำลังโหลดรายชื่อ..."/>:!students.length?<EmptyState icon="👤" msg="ไม่พบนักเรียน"/>:(
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {students.map(s=>{
                  const status=attendance[s.id]||"present"
                  const cfg=ATT_CFG[status]
                  return(
                    <div key={s.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:12,background:cfg.bg+"55",border:`1.5px solid ${cfg.bg}`}}>
                      <div style={{width:36,height:36,borderRadius:"50%",background:avatarColor(s.student_no||0),display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:"#fff",flexShrink:0}}>{avatarLetter(s.full_name)}</div>
                      <div style={{flex:1,fontSize:13.5,fontWeight:600,color:T.text}}>{s.student_no?`${s.student_no}. `:""}{s.full_name}</div>
                      <div style={{display:"flex",gap:5}}>
                        {Object.entries(ATT_CFG).map(([k,c])=>(
                          <button key={k} onClick={()=>setAttendance(p=>({...p,[s.id]:k}))} style={{padding:"5px 10px",border:`1.5px solid ${status===k?c.color:T.border}`,borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700,background:status===k?c.color:"#fff",color:status===k?"#fff":c.color,transition:"all 0.1s"}}>{c.short}</button>
                        ))}
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
          </div>
        )}

        {/* STEP: live */}
        {step==="live"&&(
          <div style={{padding:"12px"}}>
            {/* Live stats bar */}
            <div style={{background:"#fff",borderRadius:14,padding:"10px 14px",marginBottom:12,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{fontSize:12.5,fontWeight:700,color:T.text}}>⚡ Live</span>
              {Object.entries(ATT_CFG).map(([k,c])=>attCounts[k]>0&&<span key={k} style={b.pill(c.color,c.bg)}>{c.short} {attCounts[k]}</span>)}
              {behaviorLogs.length>0&&<span style={b.pill(T.purple,T.purpleL)}>⭐ {behaviorLogs.length} รายการ</span>}
              <span style={{marginLeft:"auto",fontSize:12,color:T.muted}}>แตะนักเรียนเพื่อบันทึก</span>
            </div>
            {/* Student grid */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(110px,1fr))",gap:10}}>
              {students.map((s,i)=>{
                const isAbsent=attendance[s.id]==="absent"
                const score=scores[s.id]||0
                const myLogs=behaviorLogs.filter(l=>l.studentId===s.id)
                return(
                  <button key={s.id} onClick={()=>!isAbsent&&setActiveStudent(s)} style={{background:"#fff",border:`2px solid ${score>0?"#BBF7D0":score<0?"#FECACA":T.border}`,borderRadius:14,padding:"12px 8px",cursor:isAbsent?"default":"pointer",opacity:isAbsent?0.45:1,textAlign:"center",transition:"all 0.15s",fontFamily:"inherit",position:"relative"}}>
                    <div style={{width:48,height:48,borderRadius:"50%",background:avatarColor(i),margin:"0 auto 8px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:700,color:"#fff",position:"relative"}}>
                      {avatarLetter(s.full_name)}
                      {myLogs.length>0&&<div style={{position:"absolute",top:-4,right:-4,width:18,height:18,borderRadius:"50%",background:score>0?T.green:T.red,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:"#fff"}}>{myLogs.length}</div>}
                    </div>
                    <div style={{fontSize:11.5,fontWeight:600,color:T.text,lineHeight:1.3,marginBottom:4}}>{s.full_name?.split(" ").slice(0,2).join(" ")}</div>
                    {isAbsent?<div style={{fontSize:11,color:T.red}}>ขาด</div>:(
                      <div style={{fontSize:14,fontWeight:800,color:score>0?T.green:score<0?T.red:T.muted}}>{score>0?`+${score}`:score===0?"—":score}</div>
                    )}
                  </button>
                )
              })}
            </div>
          </div>
        )}

        {/* STEP: assess */}
        {step==="assess"&&(
          <div style={{padding:"16px"}}>
            <div style={{background:"#fff",borderRadius:14,padding:"10px 14px",marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{fontSize:13,fontWeight:700,color:T.text}}>📊 ประเมินผล</span>
              {Object.entries(ASSESS_CFG).map(([k,c])=>{const cnt=Object.values(assessments).filter(v=>v===k).length;return cnt>0&&<span key={k} style={b.pill(c.color,c.bg)}>{c.label}: {cnt}</span>})}
              <span style={{fontSize:12,color:T.muted}}>ผ่าน: {passCounts.pass} | ไม่ผ่าน: {passCounts.fail}</span>
              <div style={{display:"flex",gap:6,marginLeft:"auto"}}>
                <button onClick={()=>{const a={};students.filter(s=>attendance[s.id]!=="absent").forEach(s=>{a[s.id]="good"});setAssessments(a)}} style={{...b.btnP(T.green),padding:"5px 12px",fontSize:12}}>✓ พอใช้ทั้งหมด</button>
                <button onClick={()=>{const a={};students.filter(s=>attendance[s.id]!=="absent").forEach(s=>{a[s.id]="excellent"});setAssessments(a)}} style={{...b.btnP(T.purple),padding:"5px 12px",fontSize:12}}>⭐ ดีเยี่ยมทั้งหมด</button>
              </div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {students.map(s=>{
                const isAbsent=attendance[s.id]==="absent"
                const val=assessments[s.id]
                const cfg=val?ASSESS_CFG[val]:null
                return(
                  <div key={s.id} style={{background:"#fff",borderRadius:12,border:`1.5px solid ${cfg?cfg.bg:T.border}`,overflow:"hidden"}}>
                    <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:cfg?cfg.bg+"33":"transparent"}}>
                      <div style={{width:34,height:34,borderRadius:"50%",background:avatarColor(s.student_no||0),display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:"#fff",flexShrink:0}}>{avatarLetter(s.full_name)}</div>
                      <div style={{flex:1,fontSize:13.5,fontWeight:600,color:isAbsent?T.muted:T.text}}>{s.student_no?`${s.student_no}. `:""}{s.full_name}</div>
                      {isAbsent?<span style={{fontSize:12,color:T.red}}>ขาด</span>:(
                        <div style={{display:"flex",gap:5}}>
                          {Object.entries(ASSESS_CFG).map(([k,c])=>(
                            <button key={k} onClick={()=>setAssessments(p=>({...p,[s.id]:p[s.id]===k?null:k}))} style={{padding:"5px 10px",border:`1.5px solid ${val===k?c.color:T.border}`,borderRadius:9,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700,background:val===k?c.color:"#fff",color:val===k?"#fff":c.color,transition:"all 0.12s"}}>{c.icon} {c.label}</button>
                          ))}
                        </div>
                      )}
                    </div>
                    {val==="improve"&&!isAbsent&&(
                      <div style={{padding:"8px 12px",borderTop:`1px solid ${T.orangeL}`}}>
                        <input value={assessNotes[s.id]||""} onChange={e=>setAssessNotes(p=>({...p,[s.id]:e.target.value}))} placeholder="ระบุเหตุผล/แนวทางพัฒนา (จำเป็น) *" style={{...b.input,border:`1.5px solid ${assessNotes[s.id]?.trim()?T.border:T.orange}`,fontSize:13}}/>
                        {!assessNotes[s.id]?.trim()&&<div style={{fontSize:11.5,color:T.orange,marginTop:4}}>⚠️ กรุณาระบุเหตุผล</div>}
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          </div>
        )}

        {/* STEP: log */}
        {step==="log"&&(
          <div style={{padding:"16px"}}>
            {/* Summary */}
            <div style={{background:`linear-gradient(135deg,${T.purple}22,${T.blue}11)`,borderRadius:14,padding:"14px",marginBottom:16,border:`1px solid ${T.purpleL}`}}>
              <div style={{fontSize:13,fontWeight:700,color:T.purpleD,marginBottom:8}}>สรุปคาบเรียน</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {Object.entries(ATT_CFG).map(([k,c])=>attCounts[k]>0&&<span key={k} style={b.pill(c.color,c.bg)}>{c.label}: {attCounts[k]}</span>)}
                {passCounts.total>0&&<><span style={b.pill(T.green,T.greenL)}>ผ่าน: {passCounts.pass}</span><span style={b.pill(T.orange,T.orangeL)}>ต้องพัฒนา: {passCounts.fail}</span></>}
                {behaviorLogs.length>0&&<span style={b.pill(T.purple,T.purpleL)}>⭐ บันทึกพฤติกรรม {behaviorLogs.length} รายการ</span>}
              </div>
            </div>
            <div style={b.fRow}><label style={b.label}>วันที่สอน *</label><input type="date" value={teachLog.teach_date} onChange={e=>setTeachLog(p=>({...p,teach_date:e.target.value}))} style={b.input}/></div>
            <div style={b.fRow}><label style={b.label}>เรื่องที่สอน</label><input value={teachLog.topic} onChange={e=>setTeachLog(p=>({...p,topic:e.target.value}))} style={b.input} placeholder="หน่วยที่ / บทที่ / เรื่อง..." autoFocus/></div>
            <div style={b.fRow}><label style={b.label}>จุดประสงค์การเรียนรู้</label><textarea value={teachLog.objective} onChange={e=>setTeachLog(p=>({...p,objective:e.target.value}))} style={b.textarea} placeholder="นักเรียนสามารถ..."/></div>
            <div style={b.fRow}><label style={b.label}>กิจกรรมที่ใช้ในคาบ</label><textarea value={teachLog.activities} onChange={e=>setTeachLog(p=>({...p,activities:e.target.value}))} style={b.textarea} placeholder="วิธีสอน / กิจกรรม Active Learning..."/></div>
            <div style={b.fRow}><label style={b.label}>✨ จุดเด่นของคาบเรียน</label><textarea value={teachLog.highlights} onChange={e=>setTeachLog(p=>({...p,highlights:e.target.value}))} style={{...b.textarea,minHeight:64}} placeholder="นักเรียนมีส่วนร่วม..."/></div>
            <div style={b.fRow}><label style={b.label}>⚠️ ปัญหาที่พบ</label><textarea value={teachLog.problems} onChange={e=>setTeachLog(p=>({...p,problems:e.target.value}))} style={{...b.textarea,minHeight:64}} placeholder="สิ่งที่ยังไม่ได้ผล..."/></div>
            <div style={b.fRow}><label style={b.label}>🔧 สิ่งที่ต้องปรับปรุง</label><textarea value={teachLog.improvements} onChange={e=>setTeachLog(p=>({...p,improvements:e.target.value}))} style={{...b.textarea,minHeight:64}} placeholder="จะปรับอะไรในครั้งหน้า..."/></div>
            <div style={b.fRow}><label style={b.label}>💭 Reflection ของครู</label><textarea value={teachLog.reflection} onChange={e=>setTeachLog(p=>({...p,reflection:e.target.value}))} style={{...b.textarea,minHeight:80}} placeholder="ความรู้สึก / สิ่งที่ได้เรียนรู้..."/></div>
          </div>
        )}

      </div>

      {/* Bottom Navigation */}
      <div style={{background:"#fff",borderTop:`1px solid ${T.border}`,padding:"12px 16px",display:"flex",gap:10,flexShrink:0}}>
        {step!=="select"&&<button onClick={()=>setStep(STEPS[stepIdx-1])} style={{...b.btnO,gap:6,padding:"12px 16px"}}><ChevronLeft size={16}/>ย้อนกลับ</button>}
        {step==="log"?(
          <div style={{display:"flex",gap:8,flex:1}}>
            <button onClick={()=>saveSession("draft")} disabled={saving} style={{...b.btnO,flex:1,justifyContent:"center",gap:6,padding:"12px",borderColor:T.yellow,color:T.yellow}}><Save size={15}/>บันทึกร่าง</button>
            <button onClick={()=>canProceed()?saveSession("saved"):toast.err("กรุณาระบุเหตุผลสำหรับนักเรียนที่ต้องพัฒนา")} disabled={saving} style={{...b.btnP(T.purple),flex:2,justifyContent:"center",gap:6,padding:"12px",opacity:(saving||!canProceed())?0.6:1}}>
              {saving?<><Spin size={16} color="#fff"/>กำลังบันทึก...</>:<><CheckCircle size={16}/>บันทึกหลังสอน</>}
            </button>
          </div>
        ):(
          <button onClick={()=>step!=="select"&&setStep(STEPS[stepIdx+1])} disabled={step==="select"} style={{...b.btnP(),flex:1,justifyContent:"center",gap:6,padding:"12px",opacity:step==="select"?0.4:1}}>
            {step==="live"?"สรุปคาบเรียน →":"ถัดไป"} {step!=="live"&&<ChevronRight size={16}/>}
          </button>
        )}
      </div>

      {/* Behavior sheet */}
      {activeStudent&&<BehaviorSheet student={activeStudent} score={scores[activeStudent.id]||0} logs={behaviorLogs} onBehavior={(bv)=>{addBehavior(activeStudent,bv)}} onClose={()=>setActiveStudent(null)}/>}
    </div>
  )
}


/* ─── SCORE SHEET ─────────────────────────────────────────── */
const MAX_SCORE_COLS=20

function ScoreSheet({sb,classId,className,toast}){
  const m=useIsMobile()
  const[students,setStudents]=useState([])
  const[cols,setCols]=useState([])
  const[scores,setScores]=useState({})
  const[loading,setLoading]=useState(true)
  const[saving,setSaving]=useState(false)
  const[activeCols,setActiveCols]=useState(3)

  useEffect(()=>{
    if(!classId)return
    setLoading(true)
    Promise.all([
      loadClassStudents(sb,classId,className),
      sb.from("score_columns").select("*").eq("class_id",classId).order("col_index"),
      sb.from("student_scores").select("*").eq("class_id",classId),
    ]).then(([studs,colRes,scoreRes])=>{
      setStudents(studs)
      const existCols=colRes.data||[]
      const allCols=Array.from({length:MAX_SCORE_COLS},(_,i)=>{
        const ex=existCols.find(c=>c.col_index===i+1)
        return ex||{col_index:i+1,title:"",max_score:10,class_id:classId}
      })
      setCols(allCols)
      const sm={}
      ;(scoreRes.data||[]).forEach(s=>{sm[s.student_id+"_"+s.col_index]=s.score})
      setScores(sm)
      const maxUsed=Math.max(...(colRes.data||[]).map(c=>c.col_index),3)
      setActiveCols(Math.min(maxUsed+1,MAX_SCORE_COLS))
    }).finally(()=>setLoading(false))
  },[classId])

  const setScore=(sid,ci,val)=>setScores(p=>({...p,[sid+"_"+ci]:val===""?null:parseFloat(val)||0}))
  const getScore=(sid,ci)=>scores[sid+"_"+ci]??""
  const getTotal=(sid)=>Array.from({length:activeCols},(_,i)=>scores[sid+"_"+(i+1)]||0).reduce((a,b)=>a+b,0)
  const getMax=()=>cols.slice(0,activeCols).reduce((a,c)=>a+(parseFloat(c.max_score)||0),0)
  const updateCol=(ci,field,val)=>setCols(p=>p.map(c=>c.col_index===ci?{...c,[field]:val}:c))

  const saveAll=async()=>{
    setSaving(true)
    try{
      const colsToSave=cols.slice(0,activeCols).map(c=>({class_id:classId,col_index:c.col_index,title:c.title||"",max_score:parseFloat(c.max_score)||10}))
      const{error:ce}=await sb.from("score_columns").upsert(colsToSave,{onConflict:"class_id,col_index"})
      if(ce)throw new Error(ce.message)
      const scoreRecs=[]
      students.forEach(st=>{
        for(let i=1;i<=activeCols;i++){
          const key=st.id+"_"+i
          if(scores[key]!==undefined&&scores[key]!==null&&scores[key]!==""){
            scoreRecs.push({class_id:classId,student_id:st.id,col_index:i,score:parseFloat(scores[key])||0})
          }
        }
      })
      if(scoreRecs.length){const{error:se}=await sb.from("student_scores").upsert(scoreRecs,{onConflict:"class_id,student_id,col_index"});if(se)throw new Error(se.message)}
      toast.ok("บันทึกคะแนนแล้ว ✅")
    }catch(e){toast.err("บันทึกไม่ได้: "+e.message)}finally{setSaving(false)}
  }

  if(loading)return <PageLoad text="กำลังโหลดคะแนน..."/>
  if(!students.length)return <EmptyState icon="👥" msg="ไม่พบนักเรียน" sub="ตรวจสอบชื่อห้องเรียน (เช่น ม.1/1)"/>

  const visibleCols=cols.slice(0,activeCols)
  const maxTotal=getMax()

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,flex:1}}>{students.length} คน · {activeCols} ช่อง · รวม {maxTotal} คะแนน</div>
        <div style={{display:"flex",gap:6}}>
          {activeCols<MAX_SCORE_COLS&&<button onClick={()=>setActiveCols(p=>Math.min(p+1,MAX_SCORE_COLS))} style={{padding:"7px 12px",border:"1px solid "+T.border,borderRadius:9,cursor:"pointer",background:"#fff",fontFamily:"inherit",fontSize:13,color:T.purple,display:"flex",alignItems:"center",gap:4}}><Plus size={13}/>เพิ่มช่อง</button>}
          {activeCols>1&&<button onClick={()=>setActiveCols(p=>Math.max(p-1,1))} style={{padding:"7px 12px",border:"1px solid "+T.border,borderRadius:9,cursor:"pointer",background:"#fff",fontFamily:"inherit",fontSize:13,color:T.red,display:"flex",alignItems:"center",gap:4}}><Minus size={13}/>ลด</button>}
          <button onClick={saveAll} disabled={saving} style={{padding:"7px 14px",background:T.green,color:"#fff",border:"none",borderRadius:9,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:5,opacity:saving?0.7:1}}>
            {saving?<><Spin size={14} color="#fff"/>บันทึก...</>:<><Save size={14}/>บันทึก</>}
          </button>
        </div>
      </div>
      <div style={{overflowX:"auto",borderRadius:14,border:"1px solid "+T.border,background:"#fff"}}>
        <table style={{borderCollapse:"collapse",minWidth:"100%"}}>
          <thead>
            <tr style={{background:"#F8FAFC"}}>
              <th style={{padding:"8px 12px",fontSize:12.5,color:T.muted,fontWeight:700,textAlign:"left",borderBottom:"1px solid "+T.border,borderRight:"1px solid "+T.border,position:"sticky",left:0,background:"#F8FAFC",minWidth:m?130:170,zIndex:2}}>ชื่อ-สกุล</th>
              {visibleCols.map(c=>(
                <th key={c.col_index} style={{padding:"6px 8px",fontSize:12,color:T.text,fontWeight:600,textAlign:"center",borderBottom:"1px solid "+T.border,borderRight:"1px solid #F1F5F9",minWidth:76}}>
                  <input value={c.title} onChange={e=>updateCol(c.col_index,"title",e.target.value)} placeholder={"งาน "+c.col_index} style={{width:68,border:"1px solid "+T.border,borderRadius:6,padding:"3px 5px",fontSize:11.5,fontFamily:"inherit",outline:"none",textAlign:"center",background:"transparent"}}/>
                </th>
              ))}
              <th style={{padding:"8px 10px",fontSize:12.5,color:T.purple,fontWeight:700,textAlign:"center",borderBottom:"1px solid "+T.border,minWidth:68,background:T.purpleL}}>รวม</th>
            </tr>
            <tr style={{background:"#FEFEFE"}}>
              <td style={{padding:"5px 12px",fontSize:11.5,color:T.muted,borderBottom:"1px solid "+T.border,borderRight:"1px solid "+T.border,position:"sticky",left:0,background:"#FEFEFE",zIndex:2}}>เต็ม</td>
              {visibleCols.map(c=>(
                <td key={c.col_index} style={{padding:"4px 8px",textAlign:"center",borderBottom:"1px solid "+T.border,borderRight:"1px solid #F1F5F9"}}>
                  <input type="number" value={c.max_score} onChange={e=>updateCol(c.col_index,"max_score",e.target.value)} style={{width:50,border:"1px solid "+T.border,borderRadius:6,padding:"3px 5px",fontSize:12,fontFamily:"inherit",outline:"none",textAlign:"center"}}/>
                </td>
              ))}
              <td style={{padding:"5px 10px",textAlign:"center",borderBottom:"1px solid "+T.border,fontSize:12.5,fontWeight:700,color:T.purple,background:T.purpleL+"44"}}>{maxTotal}</td>
            </tr>
          </thead>
          <tbody>
            {students.map((st,idx)=>{
              const total=getTotal(st.id)
              const pct=maxTotal>0?Math.round(total/maxTotal*100):0
              return(
                <tr key={st.id} style={{background:idx%2===0?"#fff":"#FAFAFA"}}>
                  <td style={{padding:"8px 12px",fontSize:m?13:13.5,borderBottom:"1px solid #F8FAFC",borderRight:"1px solid "+T.border,position:"sticky",left:0,background:idx%2===0?"#fff":"#FAFAFA",zIndex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:28,height:28,borderRadius:"50%",background:AVATAR_COLORS[idx%8],display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff",flexShrink:0}}>{(st.full_name||"?")[0]}</div>
                      <div>
                        <div style={{fontWeight:600,color:T.text,fontSize:m?12.5:13,whiteSpace:"nowrap"}}>{st.full_name}</div>
                        <div style={{fontSize:11,color:T.muted}}>เลขที่ {st.student_no||"—"}</div>
                      </div>
                    </div>
                  </td>
                  {visibleCols.map(c=>(
                    <td key={c.col_index} style={{padding:"5px 7px",textAlign:"center",borderBottom:"1px solid #F8FAFC",borderRight:"1px solid #F1F5F9"}}>
                      <input type="number" min="0" max={c.max_score} value={getScore(st.id,c.col_index)} onChange={e=>setScore(st.id,c.col_index,e.target.value)}
                        style={{width:50,border:"1px solid "+T.border,borderRadius:8,padding:"6px 4px",fontSize:13,fontFamily:"inherit",outline:"none",textAlign:"center",background:"transparent"}}
                        onFocus={e=>e.target.style.borderColor=T.purple} onBlur={e=>e.target.style.borderColor=T.border}/>
                    </td>
                  ))}
                  <td style={{padding:"7px 10px",textAlign:"center",borderBottom:"1px solid #F8FAFC",background:pct>=80?(T.greenL+"99"):pct>=50?(T.yellowL+"99"):(T.redL+"99")}}>
                    <div style={{fontSize:14,fontWeight:800,color:pct>=80?T.green:pct>=50?T.yellow:T.red}}>{total}</div>
                    <div style={{fontSize:10.5,color:T.muted}}>{pct}%</div>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <div style={{fontSize:12,color:T.muted,marginTop:8}}>💡 คลิกชื่อช่องเพื่อตั้งชื่องาน · แก้คะแนนเต็มได้ · กด บันทึก เมื่อเสร็จ</div>
    </div>
  )
}

/* ─── CLASS DETAIL ─────────────────────────────────────────── */
function ClassDetail({sb,cls,toast,onBack,onEnterClass}){
  const m=useIsMobile()
  const[activeTab,setActiveTab]=useState("sessions")
  const sessQ=useQ(async()=>{
    if(!cls?.id)return[]
    const{data,error}=await sb.from("teaching_sessions").select("*").eq("class_id",cls.id).order("teach_date",{ascending:false})
    if(error)throw new Error(error.message);return data||[]
  },[sb,cls?.id])
  const sessions=sessQ.data||[]

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100%",background:T.slate}}>
      <div style={{background:"#fff",borderBottom:"1px solid "+T.border,flexShrink:0}}>
        <div style={{padding:"12px 16px",display:"flex",alignItems:"center",gap:12}}>
          <button onClick={onBack} style={{border:"none",background:"none",cursor:"pointer",padding:8,borderRadius:8,color:T.sub,minWidth:36,minHeight:36,display:"flex",alignItems:"center",justifyContent:"center"}}><ChevronLeft size={22}/></button>
          <div style={{width:42,height:42,borderRadius:13,background:(cls.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:800,color:cls.color||T.purple,flexShrink:0}}>{cls.class_name?.[0]}</div>
          <div style={{flex:1}}>
            <div style={{fontSize:18,fontWeight:700,color:T.text}}>{cls.class_name}</div>
            <div style={{fontSize:13,color:T.muted}}>{cls.subject_name} · {cls._studentCount||0} คน</div>
          </div>
          <button onClick={()=>onEnterClass(cls)} style={{background:cls.color||T.purple,color:"#fff",border:"none",borderRadius:12,padding:"9px 14px",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:5,flexShrink:0}}>⚡ สอน</button>
        </div>
        <div style={{display:"flex",padding:"0 16px"}}>
          {[{id:"sessions",label:"📋 บันทึกการสอน"},{id:"scores",label:"📊 คะแนน"}].map(t=>(
            <button key={t.id} onClick={()=>setActiveTab(t.id)} style={{flex:1,padding:"10px 0",border:"none",background:"none",cursor:"pointer",fontFamily:"inherit",fontSize:14,fontWeight:600,color:activeTab===t.id?cls.color||T.purple:T.muted,borderBottom:activeTab===t.id?"2.5px solid "+(cls.color||T.purple):"2.5px solid transparent",transition:"all 0.15s"}}>{t.label}</button>
          ))}
        </div>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"16px"}}>
        {activeTab==="sessions"&&<>
          {sessions.length>0&&<div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
            {[{label:"คาบทั้งหมด",v:sessions.length,c:T.purple,bg:T.purpleL},{label:"บันทึกแล้ว",v:sessions.filter(s=>s.status==="saved").length,c:T.green,bg:T.greenL},{label:"ยังเป็นร่าง",v:sessions.filter(s=>s.status==="draft").length,c:T.orange,bg:T.orangeL}].map((it,i)=>(
              <div key={i} style={{background:"#fff",borderRadius:12,border:"1px solid "+T.border,padding:"12px",textAlign:"center"}}>
                <div style={{fontSize:22,fontWeight:800,color:it.c}}>{it.v}</div>
                <div style={{fontSize:11.5,color:T.muted,marginTop:2}}>{it.label}</div>
              </div>
            ))}
          </div>}
          {sessQ.loading?<PageLoad/>:!sessions.length?<EmptyState icon="📋" msg="ยังไม่มีบันทึก" sub="กด สอน เพื่อเริ่ม" btn={<button onClick={()=>onEnterClass(cls)} style={{background:cls.color||T.purple,color:"#fff",border:"none",borderRadius:12,padding:"10px 20px",cursor:"pointer",fontFamily:"inherit",fontSize:14,fontWeight:600}}>⚡ เริ่มสอนเลย</button>}/>:(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              {sessions.map(sess=>(
                <div key={sess.id} style={{background:"#fff",borderRadius:14,border:"1px solid "+T.border,padding:"14px 16px",display:"flex",alignItems:"center",gap:12}}>
                  <div style={{width:46,height:46,borderRadius:12,background:(cls.color||T.purple)+"15",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                    <div style={{fontSize:15,fontWeight:800,color:cls.color||T.purple,lineHeight:1}}>{sess.teach_date?new Date(sess.teach_date).toLocaleDateString("th-TH",{day:"numeric"}):"—"}</div>
                    <div style={{fontSize:10,color:T.muted}}>{sess.teach_date?new Date(sess.teach_date).toLocaleDateString("th-TH",{month:"short"}):""}</div>
                  </div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:14,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{sess.topic||"ไม่ระบุหัวข้อ"}</div>
                    <div style={{fontSize:12,color:T.muted,marginTop:2}}>{sess.teach_date?new Date(sess.teach_date).toLocaleDateString("th-TH",{weekday:"short",month:"long",day:"numeric"}):""}</div>
                  </div>
                  <span style={{padding:"4px 10px",borderRadius:99,fontSize:11.5,fontWeight:700,background:sess.status==="saved"?T.greenL:T.yellowL,color:sess.status==="saved"?T.green:T.yellow,flexShrink:0}}>{sess.status==="saved"?"✓":"ร่าง"}</span>
                </div>
              ))}
            </div>
          )}
        </>}
        {activeTab==="scores"&&<ScoreSheet sb={sb} classId={cls.id} className={cls.class_name} toast={toast}/>}
      </div>
    </div>
  )
}

// ── TEACHER ROOM ──────────────────────────────────────────────
// ── SESSION DETAIL MODAL ──────────────────────────────────────
function SessionDetailModal({sb,session,onClose,toast,onReload}){
  const isDraft = session?.status === "draft"
  const[att,setAtt]=useState([])
  const[beh,setBeh]=useState([])
  const[ass,setAss]=useState([])
  const[imgs,setImgs]=useState([])
  const[loading,setLoading]=useState(true)
  const[uploading,setUploading]=useState(false)
  const[saving,setSaving]=useState(false)
  const[editMode,setEditMode]=useState(isDraft) // ร่างเปิด edit ทันที
  const[form,setForm]=useState({
    topic:session?.topic||"",
    objective:session?.objective||"",
    activities:session?.activities||"",
    highlights:session?.highlights||"",
    problems:session?.problems||"",
    improvements:session?.improvements||"",
    reflection:session?.reflection||"",
    teach_date:session?.teach_date||todayStr(),
  })
  const imgRef=useRef()

  const supabaseUrl=(store.get("mc_url")||ENV_URL||"").replace(/\/$/,"")
  const supabaseKey=store.get("mc_key")||ENV_KEY||""

  useEffect(()=>{
    if(!session?.id)return
    Promise.all([
      sb.from("attendance").select("*,students(student_no,full_name)").eq("session_id",session.id),
      sb.from("behavior_logs").select("*,students(student_no,full_name)").eq("session_id",session.id),
      sb.from("assessments").select("*,students(student_no,full_name)").eq("session_id",session.id),
      sb.from("session_images").select("*").eq("session_id",session.id),
    ]).then(([a,bh,as,im])=>{
      setAtt(a.data||[]);setBeh(bh.data||[]);setAss(as.data||[]);setImgs(im.data||[])
    }).finally(()=>setLoading(false))
  },[session?.id])

  const attCounts=att.reduce((a,v)=>{a[v.status]=(a[v.status]||0)+1;return a},{})
  const assGroups=ass.reduce((a,v)=>{a[v.level]=[...(a[v.level]||[]),v];return a},{})
  const sf=(k,v)=>setForm(p=>({...p,[k]:v}))

  const saveLog=async(status)=>{
    setSaving(true)
    try{
      const{error}=await sb.from("teaching_sessions").update({...form,status}).eq("id",session.id)
      if(error)throw new Error(error.message||JSON.stringify(error))
      toast.ok(status==="saved"?"บันทึกการสอนสมบูรณ์ ✅":"บันทึกร่างแล้ว")
      onReload(); onClose()
    }catch(e){toast.err("บันทึกไม่สำเร็จ: "+e.message)}
    finally{setSaving(false)}
  }

  const uploadImage=async(file)=>{
    if(!file)return
    if(file.size>5*1024*1024){toast.err("ไฟล์ใหญ่เกิน 5MB");return}
    if(imgs.length>=3){toast.err("แนบได้สูงสุด 3 รูป");return}
    setUploading(true)
    try{
      const ext=file.name.split(".").pop().toLowerCase()
      const fname=`${session.id}_${Date.now()}.${ext}`
      const uploadUrl=`${supabaseUrl}/storage/v1/object/activity-images/${fname}`
      const r=await fetch(uploadUrl,{
        method:"POST",
        headers:{apikey:supabaseKey,Authorization:`Bearer ${supabaseKey}`,"Content-Type":file.type,"x-upsert":"true"},
        body:file
      })
      if(!r.ok){
        const reader=new FileReader()
        reader.onload=async e=>{
          const{error}=await sb.from("session_images").insert({session_id:session.id,image_url:e.target.result})
          if(error){toast.err("บันทึกรูปไม่ได้: "+error.message);setUploading(false);return}
          const{data}=await sb.from("session_images").select("*").eq("session_id",session.id)
          setImgs(data||[]);toast.ok("แนบรูปแล้ว ✅");setUploading(false)
        }
        reader.readAsDataURL(file); return
      }
      const pubUrl=`${supabaseUrl}/storage/v1/object/public/activity-images/${fname}`
      const{error}=await sb.from("session_images").insert({session_id:session.id,image_url:pubUrl})
      if(error)throw new Error(error.message)
      const{data}=await sb.from("session_images").select("*").eq("session_id",session.id)
      setImgs(data||[]);toast.ok("แนบรูปแล้ว ✅")
    }catch(e){toast.err("อัปโหลดไม่ได้: "+e.message)}
    finally{setUploading(false)}
  }

  const deleteImage=async(img)=>{
    if(!window.confirm("ลบรูปนี้?"))return
    await sb.from("session_images").delete().eq("id",img.id)
    setImgs(p=>p.filter(i=>i.id!==img.id));toast.ok("ลบรูปแล้ว")
  }

  const Section=({title,children})=>(
    <div style={{marginBottom:18}}>
      <div style={{fontSize:12,fontWeight:700,color:T.sub,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>{title}</div>
      {children}
    </div>
  )

  // Footer buttons
  const footer = editMode ? (
    <>
      <button onClick={()=>saveLog("draft")} disabled={saving} style={{...b.btnO,flex:1,justifyContent:"center",borderColor:T.yellow,color:T.yellow,gap:6}}>
        <Save size={15}/>บันทึกร่าง
      </button>
      <button onClick={()=>saveLog("saved")} disabled={saving} style={{...b.btnP(T.green),flex:2,justifyContent:"center",gap:6,opacity:saving?0.7:1}}>
        {saving?<><Spin size={15} color="#fff"/>กำลังบันทึก...</>:<><CheckCircle size={15}/>บันทึกหลังสอน</>}
      </button>
    </>
  ) : (
    <>
      <button onClick={()=>setEditMode(true)} style={{...b.btnO,flex:1,justifyContent:"center",gap:6}}>
        ✏️ แก้ไข
      </button>
      <button onClick={onClose} style={{...b.btnP(),flex:1,justifyContent:"center",gap:6}}>
        <Check size={15}/>ปิด
      </button>
    </>
  )

  return(
    <Modal title={isDraft?"📝 ร่างบันทึกการสอน — แก้ไขและบันทึก":"บันทึกการสอน"} wide={640} onClose={onClose}
      sub={`${session?.classes?.class_name||""} · ${session?.teach_date?thaiDate(session.teach_date):""} · ${isDraft?"ร่าง — ยังไม่สมบูรณ์":"✓ สมบูรณ์"}`}
      footer={footer}>
      {loading?<PageLoad/>:(
        <div>
          {/* Draft banner */}
          {isDraft&&<div style={{background:`linear-gradient(135deg,${T.yellowL},${T.orangeL})`,borderRadius:12,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:10,border:`1px solid ${T.yellow}`}}>
            <Bell size={18} color={T.yellow}/>
            <div><div style={{fontSize:13.5,fontWeight:700,color:"#92400E"}}>คาบนี้ยังเป็นร่าง</div><div style={{fontSize:12.5,color:"#B45309"}}>กรอกบันทึกหลังสอนแล้วกด "บันทึกหลังสอน" เพื่อทำให้สมบูรณ์</div></div>
          </div>}

          {/* Edit mode: full form */}
          {editMode ? (
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
                <div><label style={b.label}>วันที่สอน</label><input type="date" value={form.teach_date} onChange={e=>sf("teach_date",e.target.value)} style={b.input}/></div>
                <div><label style={b.label}>ชั้นเรียน</label><input value={session?.classes?.class_name||"—"} disabled style={{...b.input,background:"#F8FAFC",color:T.muted}}/></div>
              </div>
              <div style={b.fRow}><label style={b.label}>เรื่องที่สอน</label><input value={form.topic} onChange={e=>sf("topic",e.target.value)} style={b.input} placeholder="หัวข้อ / บทที่..."/></div>
              <div style={b.fRow}><label style={b.label}>จุดประสงค์การเรียนรู้</label><textarea value={form.objective} onChange={e=>sf("objective",e.target.value)} style={b.textarea} placeholder="นักเรียนสามารถ..."/></div>
              <div style={b.fRow}><label style={b.label}>กิจกรรมที่ใช้ในคาบ</label><textarea value={form.activities} onChange={e=>sf("activities",e.target.value)} style={b.textarea} placeholder="วิธีสอน / กิจกรรม Active Learning..."/></div>
              <div style={{borderTop:`1px dashed ${T.border}`,paddingTop:16,marginBottom:14,fontSize:12.5,fontWeight:700,color:T.purple}}>บันทึกหลังสอน</div>
              <div style={b.fRow}><label style={b.label}>✨ จุดเด่นของคาบ</label><textarea value={form.highlights} onChange={e=>sf("highlights",e.target.value)} style={{...b.textarea,minHeight:64}} placeholder="สิ่งที่ดี นักเรียนมีส่วนร่วม..."/></div>
              <div style={b.fRow}><label style={b.label}>⚠️ ปัญหาที่พบ</label><textarea value={form.problems} onChange={e=>sf("problems",e.target.value)} style={{...b.textarea,minHeight:64}} placeholder="สิ่งที่ยังไม่ได้ผล..."/></div>
              <div style={b.fRow}><label style={b.label}>🔧 สิ่งที่ต้องปรับปรุง</label><textarea value={form.improvements} onChange={e=>sf("improvements",e.target.value)} style={{...b.textarea,minHeight:64}} placeholder="จะปรับอะไรครั้งหน้า..."/></div>
              <div style={b.fRow}><label style={b.label}>💭 Reflection ของครู</label><textarea value={form.reflection} onChange={e=>sf("reflection",e.target.value)} style={{...b.textarea,minHeight:80}} placeholder="ความรู้สึก / สิ่งที่ได้เรียนรู้..."/></div>
            </div>
          ) : (
            <div>
              {/* View mode */}
              {session.topic&&<div style={{background:`linear-gradient(135deg,${T.purpleL},${T.blueL})`,borderRadius:12,padding:"12px 16px",marginBottom:16}}>
                <div style={{fontSize:16,fontWeight:700,color:T.purpleD,marginBottom:4}}>{session.topic}</div>
                {session.objective&&<div style={{fontSize:13,color:T.sub}}>{session.objective}</div>}
              </div>}
              {session.activities&&<Section title="กิจกรรมการเรียนรู้"><div style={{fontSize:13.5,color:T.text,lineHeight:1.8,background:"#F8FAFC",borderRadius:10,padding:"10px 14px"}}>{session.activities}</div></Section>}

              {att.length>0&&<Section title={`การเข้าเรียน (${att.length} คน)`}>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                  {Object.entries(ATT_CFG).map(([k,c])=>attCounts[k]>0&&<span key={k} style={b.pill(c.color,c.bg)}>{c.label}: {attCounts[k]}</span>)}
                </div>
                {att.filter(a=>a.status!=="present").length>0&&(
                  <div style={{background:T.orangeL,borderRadius:10,padding:"10px 14px",fontSize:13}}>
                    <b>ไม่มาเรียน: </b>{att.filter(a=>a.status!=="present").map(a=>`${a.students?.full_name||"—"} (${ATT_CFG[a.status]?.label||a.status})`).join(", ")}
                  </div>
                )}
              </Section>}

              {ass.length>0&&<Section title={`ผลการประเมิน (${ass.length} คน)`}>
                {Object.entries(assGroups).map(([lv,list])=>(
                  <div key={lv} style={{marginBottom:6}}>
                    <span style={b.pill(ASSESS_CFG[lv]?.color||T.muted,ASSESS_CFG[lv]?.bg||"#F8FAFC")}>{ASSESS_CFG[lv]?.icon} {ASSESS_CFG[lv]?.label||lv}: {list.length} คน</span>
                    {list.some(a=>a.note)&&<div style={{fontSize:12,color:T.sub,marginTop:4,paddingLeft:8}}>{list.filter(a=>a.note).map(a=>`${a.students?.full_name}: ${a.note}`).join(" | ")}</div>}
                  </div>
                ))}
                <div style={{fontSize:13,color:T.sub,marginTop:6}}>
                  ผ่าน: <b style={{color:T.green}}>{ass.filter(a=>ASSESS_CFG[a.level]?.pass).length}</b> &nbsp;|&nbsp;
                  ต้องพัฒนา: <b style={{color:T.orange}}>{ass.filter(a=>!ASSESS_CFG[a.level]?.pass).length}</b> คน
                </div>
              </Section>}

              {beh.length>0&&<Section title={`พฤติกรรม/คะแนน (${beh.length} รายการ)`}>
                <div style={{display:"flex",flexDirection:"column",gap:5}}>
                  {beh.map((bh,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 12px",borderRadius:8,background:bh.points>0?T.greenL:bh.points<0?T.redL:"#F8FAFC",fontSize:13}}>
                      <span>{bh.icon||"📝"}</span>
                      <span style={{flex:1}}>{bh.students?.student_no?`${bh.students.student_no}. `:""}{bh.students?.full_name||"—"}</span>
                      {bh.points!==0&&<span style={{fontWeight:700,color:bh.points>0?T.green:T.red}}>{bh.points>0?"+":""}{bh.points}</span>}
                      {bh.behavior&&<span style={{color:T.muted,fontSize:12}}>{bh.behavior}</span>}
                    </div>
                  ))}
                </div>
              </Section>}

              {(session.highlights||session.problems||session.improvements||session.reflection)&&(
                <Section title="บันทึกหลังสอน">
                  {[["✨ จุดเด่น",session.highlights,T.greenL],["⚠️ ปัญหา",session.problems,T.redL],["🔧 ปรับปรุง",session.improvements,T.yellowL],["💭 Reflection",session.reflection,T.blueL]].map(([k,v,bg])=>v&&(
                    <div key={k} style={{marginBottom:8,padding:"10px 14px",background:bg,borderRadius:10,fontSize:13.5}}>
                      <b style={{color:T.text}}>{k}: </b>{v}
                    </div>
                  ))}
                </Section>
              )}
            </div>
          )}

          {/* Images — always shown */}
          <Section title="ภาพหลักฐานการจัดกิจกรรม">
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:8}}>
              {imgs.map((img,i)=>(
                <div key={img.id} style={{position:"relative",aspectRatio:"4/3",borderRadius:12,overflow:"hidden",border:`1px solid ${T.border}`,background:"#F8FAFC"}}>
                  <img src={img.image_url} alt={`รูปที่ ${i+1}`} style={{width:"100%",height:"100%",objectFit:"cover"}} onError={e=>{e.target.style.display="none"}}/>
                  <button onClick={()=>deleteImage(img)} style={{position:"absolute",top:5,right:5,width:24,height:24,borderRadius:"50%",background:"rgba(0,0,0,0.6)",border:"none",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
                    <X size={12} color="#fff"/>
                  </button>
                </div>
              ))}
              {imgs.length<3&&(
                <div onClick={()=>!uploading&&imgRef.current?.click()} style={{aspectRatio:"4/3",borderRadius:12,border:`2px dashed ${T.border}`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:6,cursor:uploading?"wait":"pointer",background:"#F8FAFC"}}>
                  {uploading?<Spin size={22}/>:<><Upload size={22} color={T.muted}/><span style={{fontSize:12,color:T.muted,fontWeight:600}}>แนบรูป</span><span style={{fontSize:11,color:T.muted}}>{3-imgs.length} รูปที่เหลือ</span></>}
                </div>
              )}
            </div>
            <input ref={imgRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{if(e.target.files[0]){uploadImage(e.target.files[0]);e.target.value=""}}}/>
            <div style={{fontSize:12,color:T.muted}}>JPG, PNG, WEBP · สูงสุด 3 รูป · ไม่เกิน 5MB</div>
          </Section>
        </div>
      )}
    </Modal>
  )
}

function TeacherRoom({sb,toast,onEnterClass,profile:profileProp,onOpenProfile}){
  const m=useIsMobile()
  const[tab,setTab]=useState("home")
  const[modal,setModal]=useState(null)
  const[viewSession,setViewSession]=useState(null)
  const[selectedClass,setSelectedClass]=useState(null)
  const[search,setSearch]=useState("")
  const[filterLv,setFilterLv]=useState("ทั้งหมด")

  const dashQ=useDash(sb);const clsQ=useClasses(sb);const stuQ=useStudents(sb,search,filterLv)
  const allSessQ=useAllSessions(sb);const profQ=useProfile(sb)
  const stats=dashQ.data||{students:0,classes:0,sessions:0,drafts:[],att:{present:0,absent:0,late:0,leave:0,total:0}}
  const classes=clsQ.data||[];const students=stuQ.data||[];const allSess=allSessQ.data||[];const profile=profileProp||profQ.data

  const reloadAll=()=>{dashQ.reload();clsQ.reload();stuQ.reload();allSessQ.reload();profQ.reload()}
  if(selectedClass)return <ClassDetail sb={sb} cls={selectedClass} toast={toast} onBack={()=>setSelectedClass(null)} onEnterClass={onEnterClass}/>

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
    if(!window.confirm("ยืนยันอีกครั้ง"))return
    try{
      for(const t of["attendance","behavior_logs","assessments","session_images","teaching_sessions","class_students","classes","students"]){
        await sb.from(t).delete().neq("id","00000000-0000-0000-0000-000000000000")
      }
      toast.ok("รีเซทเรียบร้อย");reloadAll()
    }catch(e){toast.err("รีเซทไม่สำเร็จ: "+e.message)}
  }

  const TABS=[
    {id:"home",icon:Home,label:"หน้าหลัก"},
    {id:"classes",icon:School,label:"ชั้นเรียน"},
    {id:"students",icon:Users,label:"นักเรียน"},
    {id:"sessions",icon:BookOpen,label:"บันทึกการสอน"},
    {id:"settings",icon:Settings,label:"ตั้งค่า"},
  ]

  return(
    <div style={{display:"flex",height:"100%"}}>
      {/* Sidebar */}
      <aside style={{width:220,minWidth:220,background:"#fff",borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",overflowY:"auto"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"18px 16px 14px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{width:36,height:36,borderRadius:11,background:`linear-gradient(135deg,${T.purple},${T.blue})`,display:"flex",alignItems:"center",justifyContent:"center"}}><GraduationCap size={20} color="#fff"/></div>
          <div><div style={{fontWeight:700,fontSize:16,color:T.text}}>MyClass 🍎</div><div style={{fontSize:11,color:T.muted}}>ห้องครู</div></div>
        </div>
        <nav style={{flex:1,padding:"8px"}}>
          {TABS.map(t=>(
            <div key={t.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:11,cursor:"pointer",fontSize:13.5,fontWeight:tab===t.id?700:500,marginBottom:2,background:tab===t.id?T.purple:"transparent",color:tab===t.id?"#fff":"#475569",transition:"all 0.15s"}} onClick={()=>setTab(t.id)}>
              <t.icon size={16}/>{t.label}
              {t.id==="sessions"&&stats.drafts?.length>0&&<span style={{marginLeft:"auto",background:T.orange,color:"#fff",borderRadius:99,width:18,height:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700}}>{stats.drafts.length}</span>}
            </div>
          ))}
          <div style={{marginTop:8,padding:"0 4px"}}>
            <button onClick={handleBackup} style={{...b.btnO,width:"100%",justifyContent:"center",fontSize:12.5,padding:"8px 12px",marginBottom:6}}><Download size={14}/>สำรองข้อมูล</button>
            <button onClick={handleReset} style={{...b.btnO,width:"100%",justifyContent:"center",fontSize:12.5,padding:"8px 12px",color:T.red,borderColor:"#FECACA"}}><RotateCcw size={14}/>รีเซท</button>
          </div>
        </nav>
        <div style={{padding:"12px 10px 16px"}}>
          <button onClick={()=>setModal("profile")} style={{display:"flex",alignItems:"center",gap:8,width:"100%",border:"none",background:"none",cursor:"pointer",padding:"8px 10px",borderRadius:10}}>
            <div style={{width:34,height:34,borderRadius:"50%",background:`linear-gradient(135deg,${T.purple},${T.blue})`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><span style={{color:"#fff",fontSize:13,fontWeight:700}}>{profile?.teacher_name?.[0]||"ค"}</span></div>
            <div style={{textAlign:"left"}}><div style={{fontSize:13,fontWeight:600,color:T.text,lineHeight:1.2}}>{profile?.teacher_name||"ตั้งค่าชื่อครู"}</div><div style={{fontSize:11,color:T.muted}}>{profile?.school_name||"กดเพื่อตั้งค่า"}</div></div>
          </button>
        </div>
      </aside>

      {/* Main content */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:58,flexShrink:0}}>
          <div style={{fontSize:16,fontWeight:700,color:T.text}}>{TABS.find(t=>t.id===tab)?.label}</div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={reloadAll} style={{border:"none",background:"none",cursor:"pointer",padding:8,borderRadius:8}}><RefreshCw size={16} color={T.muted}/></button>
          </div>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"20px 24px"}}>

          {/* HOME */}
          {tab==="home"&&<>
            {stats.drafts?.length>0&&(
              <div style={{background:`linear-gradient(135deg,${T.orange}22,${T.yellow}11)`,border:`1.5px solid ${T.orange}`,borderRadius:14,padding:"14px 18px",marginBottom:18,display:"flex",alignItems:"center",gap:12}}>
                <Bell size={22} color={T.orange}/>
                <div style={{flex:1}}><div style={{fontSize:14,fontWeight:700,color:T.orange}}>มี {stats.drafts.length} คาบที่ยังไม่บันทึกการสอน</div><div style={{fontSize:12.5,color:T.sub,marginTop:2}}>{stats.drafts.map(d=>`${d.classes?.class_name||"—"} ${d.teach_date?thaiDate(d.teach_date):""}`).join(" | ")}</div></div>
                <button onClick={()=>setTab("sessions")} style={{...b.btnP(T.orange),padding:"7px 14px",fontSize:12.5,flexShrink:0}}>ดูรายการ →</button>
              </div>
            )}
            <div style={{fontSize:16,fontWeight:700,color:T.text,marginBottom:4}}>สวัสดี {profile?.teacher_name||"คุณครู"} 👋</div>
            <div style={{fontSize:13,color:T.muted,marginBottom:18}}>{thaiDate()}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
              {[
                {icon:p=><Users {...p}/>,label:"นักเรียนทั้งหมด",v:stats.students,u:"คน",bg:T.blueL,ic:T.blue},
                {icon:p=><School {...p}/>,label:"ชั้นเรียน",v:stats.classes,u:"ห้อง",bg:T.yellowL,ic:T.yellow},
                {icon:p=><CalendarDays {...p}/>,label:"คาบที่สอน",v:stats.sessions,u:"คาบ",bg:T.purpleL,ic:T.purple},
                {icon:p=><Star {...p}/>,label:"เข้าเรียน",v:stats.att.total?`${Math.round(stats.att.present/stats.att.total*100)}%`:"—",u:"เฉลี่ย",bg:T.greenL,ic:T.green},
              ].map((it,i)=><div key={i} style={{...b.card,padding:"16px"}}>
                <div style={{width:40,height:40,borderRadius:12,background:it.bg,display:"flex",alignItems:"center",justifyContent:"center",marginBottom:10}}>{it.icon({size:20,color:it.ic})}</div>
                {dashQ.loading?<Spin size={22}/>:<><div style={{fontSize:26,fontWeight:800,color:T.text,lineHeight:1}}>{it.v}</div><div style={{fontSize:12,color:T.muted,marginTop:3}}>{it.label} · {it.u}</div></>}
              </div>)}
            </div>
            {/* Recent sessions */}
            <div style={{...b.card,marginBottom:20}}>
              <div style={{padding:"14px 18px 12px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <span style={{fontSize:14.5,fontWeight:700,color:T.text}}>บันทึกการสอนล่าสุด</span>
                <button onClick={()=>setTab("sessions")} style={{border:"none",background:"none",cursor:"pointer",fontSize:12.5,color:T.purple,fontFamily:"inherit",fontWeight:600}}>ดูทั้งหมด →</button>
              </div>
              {allSess.slice(0,5).length===0?<EmptyState icon="📋" msg="ยังไม่มีบันทึก"/>:(
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <tbody>{allSess.slice(0,5).map(s=>(
                    <tr key={s.id}><td style={{padding:"10px 18px",fontSize:13.5,color:"#334155",borderBottom:`1px solid #F8FAFC`}}>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <div style={{width:8,height:8,borderRadius:"50%",background:s.classes?.color||T.purple,flexShrink:0}}/>
                        <div style={{flex:1}}><div style={{fontWeight:600}}>{s.classes?.class_name||"—"}</div><div style={{fontSize:12,color:T.muted}}>{s.topic||"ไม่ระบุหัวข้อ"}</div></div>
                        <div style={{fontSize:12,color:T.muted}}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short"}):"—"}</div>
                        <span style={{padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:700,background:s.status==="saved"?T.greenL:T.yellowL,color:s.status==="saved"?T.green:T.yellow}}>{s.status==="saved"?"✓ บันทึกแล้ว":"ร่าง"}</span>
                      </div>
                    </td></tr>
                  ))}</tbody>
                </table>
              )}
            </div>
            {/* Class cards */}
            {classes.length>0&&<><div style={{fontSize:14.5,fontWeight:700,color:T.text,marginBottom:12}}>ชั้นเรียนของฉัน</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14}}>
              {classes.slice(0,6).map(c=><div key={c.id} style={{...b.card,overflow:"hidden",cursor:"pointer"}} onClick={()=>setTab("classes")}><div style={{height:5,background:c.color||T.purple}}/><div style={{padding:"14px"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div><div style={{fontSize:15,fontWeight:700,color:T.text}}>{c.class_name}</div><div style={{fontSize:12.5,color:T.muted}}>{c.subject_name}</div></div><div style={{width:36,height:36,borderRadius:10,background:(c.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,color:c.color||T.purple,fontSize:15}}>{c.class_name?.[0]}</div></div><div style={{fontSize:12.5,color:"#64748B"}}><Users size={12} style={{verticalAlign:"middle"}}/> {c._studentCount||0} คน</div></div></div>)}
            </div></>}
          </>}

          {/* CLASSES */}
          {tab==="classes"&&<>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
              <div style={{fontSize:14,color:T.muted}}>{clsQ.loading?"กำลังโหลด...":`${classes.length} ชั้นเรียน`}</div>
              <button style={b.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>เพิ่มชั้นเรียน</button>
            </div>
            {clsQ.loading?<PageLoad/>:!classes.length?<EmptyState icon="🏫" msg="ยังไม่มีชั้นเรียน" btn={<button style={b.btnP()} onClick={()=>setModal("addClass")}><Plus size={15}/>สร้างชั้นเรียนแรก</button>}/>:(
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                {classes.map(c=><div key={c.id} style={{...b.card,overflow:"hidden"}}>
                  <div style={{height:6,background:c.color||T.purple}}/>
                  <div style={{padding:"16px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                      <div><div style={{fontSize:17,fontWeight:700,color:T.text}}>{c.class_name}</div><div style={{fontSize:13,color:T.muted}}>{c.subject_name}</div></div>
                      <div style={{width:42,height:42,borderRadius:12,background:(c.color||T.purple)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,fontWeight:700,color:c.color||T.purple}}>{c.class_name?.[0]}</div>
                    </div>
                    <div style={{fontSize:13,color:"#64748B",marginBottom:14}}><Users size={12} style={{verticalAlign:"middle"}}/> {c._studentCount||0} คน</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                      <button onClick={()=>onEnterClass(c)} style={{padding:"9px 0",background:c.color||T.purple,color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit"}}>⚡ สอนเลย</button>
                      <button onClick={async()=>{if(!window.confirm(`ลบ ${c.class_name}?`))return;const{error}=await sb.from("classes").delete().eq("id",c.id);if(error){toast.err(error.message);return}toast.ok("ลบแล้ว");clsQ.reload();dashQ.reload()}} style={{padding:"9px 0",background:T.redL,color:T.red,border:"none",borderRadius:10,cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit"}}>ลบ</button>
                    </div>
                  </div>
                </div>)}
              </div>
            )}
          </>}

          {/* STUDENTS */}
          {tab==="students"&&<>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <div style={{fontSize:14,color:T.muted}}>{stuQ.loading?"กำลังโหลด...":`${students.length} คน`}</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button style={b.btnO} onClick={()=>setModal("paste")}><ClipboardList size={15}/>วางรายชื่อ</button>
                <button style={b.btnP(T.green)} onClick={()=>setModal("excel")}><Upload size={15}/>Import Excel</button>
                <button style={b.btnP()} onClick={()=>setModal("addStu")}><Plus size={15}/>เพิ่มเอง</button>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:180,position:"relative"}}><Search size={15} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:T.muted}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาชื่อ..." style={{...b.input,paddingLeft:34}}/></div>
              {["ทั้งหมด","ม.1","ม.2","ม.3","ม.4","ม.5","ม.6"].map(l=><button key={l} onClick={()=>setFilterLv(l)} style={{padding:"7px 12px",border:"1px solid",borderColor:filterLv===l?T.purple:T.border,borderRadius:10,cursor:"pointer",background:filterLv===l?T.purpleL:"#fff",color:filterLv===l?T.purple:"#64748B",fontSize:12.5,fontWeight:600,fontFamily:"inherit"}}>{l}</button>)}
            </div>
            {stuQ.loading?<PageLoad/>:!students.length?<EmptyState icon="👩‍🎓" msg="ไม่พบนักเรียน" btn={<button style={b.btnP()} onClick={()=>setModal("addStu")}><Plus size={15}/>เพิ่มนักเรียน</button>}/>:(
              <div style={{...b.card}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr style={{background:"#F8FAFC"}}>{["#","ชื่อ-นามสกุล","ชั้น","ห้อง",""].map((h,i)=><th key={i} style={{padding:"8px 14px",fontSize:12.5,color:T.muted,fontWeight:600,textAlign:"left",borderBottom:`1px solid #F1F5F9`,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                  <tbody>{students.map((stu,i)=><tr key={stu.id}>
                    <td style={{padding:"10px 14px",fontSize:12.5,color:T.muted,width:40}}>{stu.student_no||i+1}</td>
                    <td style={{padding:"10px 14px",fontSize:13.5,color:"#334155"}}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:30,height:30,borderRadius:"50%",background:avatarColor(i),display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:"#fff",flexShrink:0}}>{avatarLetter(stu.full_name)}</div>{stu.full_name}</div></td>
                    <td style={{padding:"10px 14px",fontSize:13.5,color:"#334155"}}>{stu.level}</td>
                    <td style={{padding:"10px 14px",fontSize:13.5,color:"#334155"}}>{stu.room}</td>
                    <td style={{padding:"10px 14px",width:50}}><button onClick={async()=>{if(!window.confirm(`ลบ ${stu.full_name}?`))return;const{error}=await sb.from("students").delete().eq("id",stu.id);if(error){toast.err(error.message);return}toast.ok("ลบแล้ว");stuQ.reload();dashQ.reload()}} style={{padding:"5px 8px",border:"1px solid #FECACA",borderRadius:8,cursor:"pointer",background:T.redL,color:T.red}}><Trash2 size={13}/></button></td>
                  </tr>)}</tbody>
                </table>
                <div style={{padding:"10px 14px",background:"#F8FAFC",fontSize:12.5,color:T.muted,borderTop:"1px solid #F1F5F9"}}>แสดง {students.length} คน</div>
              </div>
            )}
          </>}

          {/* SESSIONS */}
          {tab==="sessions"&&<>
            <div style={{fontSize:14,color:T.muted,marginBottom:14}}>{allSessQ.loading?"กำลังโหลด...":`${allSess.length} บันทึก · คลิกแถวเพื่อดูรายละเอียด`}</div>
            {allSessQ.loading?<PageLoad/>:!allSess.length?<EmptyState icon="📋" msg="ยังไม่มีบันทึกการสอน"/>:(
              <div style={b.card}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr style={{background:"#F8FAFC"}}>{["ชั้น/วิชา","หัวข้อ","วันที่","สถานะ",""].map(h=><th key={h} style={{padding:"8px 14px",fontSize:12.5,color:T.muted,fontWeight:600,textAlign:"left",borderBottom:"1px solid #F1F5F9"}}>{h}</th>)}</tr></thead>
                  <tbody>{allSess.map(s=>(
                    <tr key={s.id} onClick={()=>setViewSession(s)} style={{cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background=""}>
                      <td style={{padding:"11px 14px",fontSize:13.5,color:"#334155",borderBottom:"1px solid #F8FAFC"}}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:8,height:8,borderRadius:"50%",background:s.classes?.color||T.purple,flexShrink:0}}/><div><div style={{fontWeight:600,fontSize:13}}>{s.classes?.class_name||"—"}</div><div style={{color:T.muted,fontSize:11.5}}>{s.classes?.subject_name}</div></div></div></td>
                      <td style={{padding:"11px 14px",fontSize:13,color:"#64748B",maxWidth:160,borderBottom:"1px solid #F8FAFC"}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.topic||"ไม่ระบุ"}</div></td>
                      <td style={{padding:"11px 14px",fontSize:12,color:T.muted,whiteSpace:"nowrap",borderBottom:"1px solid #F8FAFC"}}>{s.teach_date?new Date(s.teach_date).toLocaleDateString("th-TH",{day:"numeric",month:"short",year:"2-digit"}):"—"}</td>
                      <td style={{padding:"11px 14px",borderBottom:"1px solid #F8FAFC"}}><span style={{padding:"3px 10px",borderRadius:99,fontSize:12,fontWeight:700,background:s.status==="saved"?T.greenL:T.yellowL,color:s.status==="saved"?T.green:T.yellow}}>{s.status==="saved"?"✓ สมบูรณ์":"ร่าง"}</span></td>
                      <td style={{padding:"11px 14px",borderBottom:"1px solid #F8FAFC",color:T.muted}}><Eye size={14}/></td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            )}
          </>}

          {/* SETTINGS */}
          {tab==="settings"&&(
            <div style={{maxWidth:480}}>
              <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:12,textTransform:"uppercase",letterSpacing:0.5}}>ข้อมูลครูและโรงเรียน</div>
              {profQ.loading?<PageLoad/>:<ProfileForm sb={sb} profile={profQ.data} onDone={profQ.reload} toast={toast}/>}
            </div>
          )}

        </div>
      </div>

      {/* Modals */}

        {/* Mobile bottom nav */}
        {m&&<div style={{background:"#fff",borderTop:"1px solid "+T.border,display:"grid",gridTemplateColumns:"repeat(4,1fr)",flexShrink:0}}>
          {[{id:"home",icon:Home,label:"หน้าหลัก"},{id:"classes",icon:School,label:"ชั้นเรียน"},{id:"students",icon:Users,label:"นักเรียน"},{id:"sessions",icon:BookOpen,label:"บันทึก"}].map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 4px 8px",border:"none",background:"none",cursor:"pointer",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
              <t.icon size={20} color={tab===t.id?T.purple:T.muted}/>
              <span style={{fontSize:10,fontWeight:600,color:tab===t.id?T.purple:T.muted}}>{t.label}</span>
            </button>
          ))}
        </div>}

      {modal==="addClass"&&<AddClassModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{clsQ.reload();dashQ.reload()}} toast={toast}/>}
      {modal==="addStu"&&<AddStudentModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
      {modal==="paste"&&<PasteModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
      {modal==="excel"&&<ImportExcelModal sb={sb} onClose={()=>setModal(null)} onDone={()=>{stuQ.reload();dashQ.reload()}} toast={toast}/>}
      {modal==="profile"&&<Modal title="ตั้งค่าโปรไฟล์" onClose={()=>setModal(null)}><ProfileForm sb={sb} profile={profQ.data} onDone={()=>{profQ.reload();setModal(null)}} toast={toast}/></Modal>}
      {viewSession&&<SessionDetailModal sb={sb} session={viewSession} onClose={()=>setViewSession(null)} toast={toast} onReload={allSessQ.reload}/>}
    </div>
  )
}

// ── PROFILE FORM ──────────────────────────────────────────────
function ProfileForm({sb,profile,onDone,toast}){
  const[form,setForm]=useState({teacher_name:profile?.teacher_name||"",school_name:profile?.school_name||"",subject_group:profile?.subject_group||""})
  const[saving,setSaving]=useState(false)
  const sf=(k,v)=>setForm(p=>({...p,[k]:v}))
  const save=async()=>{
    setSaving(true)
    try{if(profile?.id){await sb.from("profiles").update(form).eq("id",profile.id)}else{await sb.from("profiles").insert(form)};toast.ok("บันทึกแล้ว");onDone()}catch(e){toast.err(e.message)}finally{setSaving(false)}
  }
  return(
    <div>
      <div style={b.fRow}><label style={b.label}>ชื่อ-นามสกุลครู</label><input value={form.teacher_name} onChange={e=>sf("teacher_name",e.target.value)} style={b.input} placeholder="ครูชื่อ นามสกุล"/></div>
      <div style={b.fRow}><label style={b.label}>โรงเรียน</label><input value={form.school_name} onChange={e=>sf("school_name",e.target.value)} style={b.input} placeholder="โรงเรียน..."/></div>
      <div style={{...b.fRow,marginBottom:20}}><label style={b.label}>กลุ่มสาระ</label><input value={form.subject_group} onChange={e=>sf("subject_group",e.target.value)} style={b.input} placeholder="คณิตศาสตร์"/></div>
      <button onClick={save} disabled={saving} style={{...b.btnP(),width:"100%",justifyContent:"center",padding:"12px"}}>{saving?<><Spin size={16} color="#fff"/>กำลังบันทึก...</>:<><Save size={16}/>บันทึกข้อมูล</>}</button>
    </div>
  )
}

// ── ROOT APP ──────────────────────────────────────────────────
const ENV_URL=import.meta.env.VITE_SUPABASE_URL||""
const ENV_KEY=import.meta.env.VITE_SUPABASE_ANON_KEY||""

export default function App(){
  const initUrl=ENV_URL||store.get("mc_url")||""
  const initKey=ENV_KEY||store.get("mc_key")||""
  const autoConn=!!(initUrl&&initKey)
  const[connected,setConnected]=useState(autoConn)
  const[sb,setSb]=useState(autoConn?createClient(initUrl,initKey):null)
  const[mode,setMode]=useState("teacher") // "teacher" | "classroom"
  const[preselectedClass,setPreselectedClass]=useState(null)
  const toast=useToast()

  const handleReady=(url,key)=>{store.set("mc_url",url);store.set("mc_key",key);setSb(createClient(url,key));setConnected(true)}

  const enterClassroom=(cls=null)=>{
    setPreselectedClass(cls)
    setMode("classroom")
  }

  const clsQ=useClasses(sb)
  const classes=clsQ.data||[]

  if(!connected||!sb)return <SetupScreen onReady={handleReady}/>

  return(
    <ErrorBoundary>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');@keyframes mcSpin{to{transform:rotate(360deg)}}@keyframes mcSlide{from{transform:translateY(-8px);opacity:0}to{transform:translateY(0);opacity:1}}*{box-sizing:border-box}body{margin:0}button:active{transform:scale(0.97)}`}</style>
      <Toast list={toast.list}/>
      <div style={{height:"100vh",display:"flex",flexDirection:"column",fontFamily:"'Prompt','Sarabun',sans-serif",background:T.slate,overflow:"hidden"}}>
        {/* Top mode toggle */}
        <div style={{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"0 20px",display:"flex",alignItems:"center",justifyContent:"space-between",height:52,flexShrink:0,zIndex:10}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:30,height:30,borderRadius:9,background:`linear-gradient(135deg,${T.purple},${T.blue})`,display:"flex",alignItems:"center",justifyContent:"center"}}><GraduationCap size={17} color="#fff"/></div>
            <span style={{fontWeight:700,fontSize:15,color:T.text}}>MyClass</span>
          </div>
          <div style={{display:"flex",background:"#F1F5F9",borderRadius:12,padding:3,gap:2}}>
            <button onClick={()=>setMode("teacher")} style={{padding:"6px 16px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,background:mode==="teacher"?"#fff":"transparent",color:mode==="teacher"?T.purple:T.muted,boxShadow:mode==="teacher"?"0 2px 8px rgba(0,0,0,0.08)":"none",transition:"all 0.2s"}}>🏫 ห้องครู</button>
            <button onClick={()=>enterClassroom()} style={{padding:"6px 16px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,background:mode==="classroom"?T.purple:"transparent",color:mode==="classroom"?"#fff":T.muted,boxShadow:mode==="classroom"?"0 2px 8px "+T.purple+"44":"none",transition:"all 0.2s"}}>⚡ เริ่มสอน</button>
          </div>
          <div style={{width:80}}/>
        </div>

        {/* Content */}
        <div style={{flex:1,overflow:"hidden"}}>
          {mode==="teacher"&&<TeacherRoom sb={sb} toast={toast} onEnterClass={enterClassroom} profile={profQ?.data} onOpenProfile={()=>setShowProfile&&setShowProfile(true)}/>}
          {mode==="classroom"&&<ClassroomSession sb={sb} classes={classes} toast={toast} onExit={()=>setMode("teacher")} onReload={clsQ.reload}/>}
        </div>
      </div>
    </ErrorBoundary>
  )
}