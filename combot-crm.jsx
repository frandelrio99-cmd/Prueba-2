import { useState, useEffect, useMemo, useRef } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, CartesianGrid, LineChart, Line } from "recharts";
import * as XLSX from "xlsx";

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const ESTADOS = [
  { id:"nuevo",       label:"Nuevo",       color:"#3b82f6" },
  { id:"contactado",  label:"Contactado",  color:"#06b6d4" },
  { id:"recontactar", label:"Recontactar", color:"#f59e0b" },
  { id:"vendido",     label:"Vendido",     color:"#10b981" },
  { id:"perdido",     label:"Perdido",     color:"#ef4444" },
  { id:"sin_asignar", label:"No asignado", color:"#f97316" },
];
const EM = Object.fromEntries(ESTADOS.map(e=>[e.id,e]));
const CATALOGADOS = ["recontactar","vendido","perdido","sin_asignar"];
const PRODUCTOS   = ["DJI","EcoFlow","SwellPro","Chasing","Otro"];
const FUENTES     = ["Zoho Chat","WhatsApp","Instagram","Web","Teléfono","Otro"];
const CIUDADES    = ["CABA","GBA Norte","GBA Sur","GBA Oeste","Córdoba","Rosario","Mendoza","Interior","Otro"];
const DISPOSITIVOS= ["Móvil","Desktop","Tablet"];
const DIAS_SEMANA = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];

const USERS_KEY = "td_users_v6";
const defaultUsers = {
  admin:     { pass:"admin2024",     role:"admin",     nombre:"Federico",  activo:true },
  francisco: { pass:"francisco2024", role:"comercial", nombre:"Francisco", activo:true },
  nacho:     { pass:"nacho2024",     role:"comercial", nombre:"Nacho",     activo:true },
};

const db = {
  getLeads:  ()=>{ try{return JSON.parse(localStorage.getItem("td_v6")||"[]")}catch{return[]} },
  setLeads:  d=>localStorage.setItem("td_v6",JSON.stringify(d)),
  getSess:   ()=>{ try{return JSON.parse(localStorage.getItem("td_s6")||"null")}catch{return null} },
  setSess:   s=>localStorage.setItem("td_s6",JSON.stringify(s)),
  clearSess: ()=>localStorage.removeItem("td_s6"),
  getUsers:  ()=>{ try{return JSON.parse(localStorage.getItem(USERS_KEY)||"null")||defaultUsers}catch{return defaultUsers} },
  setUsers:  u=>localStorage.setItem(USERS_KEY,JSON.stringify(u)),
};

const todayStr  = ()=>new Date().toISOString().split("T")[0];
const fmtDate   = iso=>new Date(iso+"T12:00:00").toLocaleDateString("es-AR",{day:"2-digit",month:"2-digit",year:"2-digit"});
const fmtTime   = iso=>new Date(iso).toLocaleTimeString("es-AR",{hour:"2-digit",minute:"2-digit"});
const fmtShort  = iso=>new Date(iso+"T12:00:00").toLocaleDateString("es-AR",{day:"2-digit",month:"2-digit"});
const uid       = ()=>Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const horasDesde= iso=>(Date.now()-new Date(iso).getTime())/36e5;
const getLast7  = ()=>{const d=[];for(let i=6;i>=0;i--){const x=new Date();x.setDate(x.getDate()-i);d.push(x.toISOString().split("T")[0]);}return d;};
const getLast30 = ()=>{const d=[];for(let i=29;i>=0;i--){const x=new Date();x.setDate(x.getDate()-i);d.push(x.toISOString().split("T")[0]);}return d;};

const autoEscalar = leads=>leads.map(l=>
  l.estado==="nuevo"&&horasDesde(l.createdAt)>=24
    ?{...l,estado:"sin_asignar",autoEsc:true,catalogadoAt:l.catalogadoAt||new Date(new Date(l.createdAt).getTime()+24*36e5).toISOString()}
    :l
);

const now0 = new Date();
const h = h=>new Date(now0-h*36e5).toISOString();
const SEED = [
  {id:"s1",nombre:"Martín García",    tel:"5491155556666",fecha:todayStr(),estado:"nuevo",     producto:"DJI",     fuente:"Zoho Chat",asesor:"Francisco",ciudad:"GBA Norte", dispositivo:"Móvil",  pagina:"/dji-mini-4-pro", createdAt:h(2),  notas:[],nota_rapida:"",historial:[]},
  {id:"s2",nombre:"Sofía Gómez",      tel:"5491199990000",fecha:todayStr(),estado:"nuevo",     producto:"EcoFlow", fuente:"Instagram", asesor:"Francisco",ciudad:"CABA",      dispositivo:"Móvil",  pagina:"/ecoflow",        createdAt:h(1),  notas:[],nota_rapida:"",historial:[]},
  {id:"s3",nombre:"Laura Pérez",      tel:"5491144445555",fecha:todayStr(),estado:"recontactar",producto:"EcoFlow",fuente:"Zoho Chat",asesor:"Francisco",ciudad:"CABA",      dispositivo:"Desktop",pagina:"/ecoflow-delta",  createdAt:h(5),  notas:[{texto:"Espera cobrar el viernes. Muy interesada en DELTA 2. Llamar el viernes a la tarde.",ts:h(4),autor:"Francisco"}],nota_rapida:"Llamar viernes tarde",historial:["recontactar"],catalogadoAt:h(4)},
  {id:"s4",nombre:"Carlos Rodríguez", tel:"5491133334444",fecha:todayStr(),estado:"vendido",   producto:"DJI",     fuente:"Zoho Chat",asesor:"Francisco",ciudad:"GBA Norte", dispositivo:"Desktop",pagina:"/mavic-3",         createdAt:h(6),  notas:[{texto:"Cerró Mavic 3 Classic. Pago en efectivo. Muy buena experiencia.",ts:h(3),autor:"Francisco"}],nota_rapida:"",historial:["vendido"],catalogadoAt:h(3)},
  {id:"s5",nombre:"Diego Fernández",  tel:"5491111112222",fecha:todayStr(),estado:"sin_asignar",producto:"DJI",   fuente:"Web",       asesor:"Francisco",ciudad:"Córdoba",   dispositivo:"Móvil",  pagina:"/dji-avata",      createdAt:h(26), notas:[],nota_rapida:"",historial:[],autoEsc:true,catalogadoAt:h(2)},
  {id:"s6",nombre:"Ana Martínez",     tel:"5491122223333",fecha:todayStr(),estado:"perdido",   producto:"SwellPro",fuente:"WhatsApp", asesor:"Nacho",    ciudad:"GBA Sur",   dispositivo:"Móvil",  pagina:"/swellpro",       createdAt:h(4),  notas:[{texto:"Presupuesto alto. Dijo que lo piensa. No respondió más.",ts:h(2),autor:"Nacho"}],nota_rapida:"",historial:["perdido"],catalogadoAt:h(2)},
  {id:"s7",nombre:"Roberto Suárez",   tel:"5491166667777",fecha:todayStr(),estado:"recontactar",producto:"DJI",  fuente:"Instagram", asesor:"Nacho",    ciudad:"CABA",      dispositivo:"Móvil",  pagina:"/dji-avata-2",    createdAt:h(3),  notas:[{texto:"Muy interesado en Avata 2. No responde en las últimas 2hs. Insistir mañana.",ts:h(1),autor:"Nacho"}],nota_rapida:"Insistir mañana AM",historial:["recontactar"],catalogadoAt:h(1)},
  {id:"s8",nombre:"Valeria Torres",   tel:"5491177778888",fecha:todayStr(),estado:"nuevo",     producto:"EcoFlow",fuente:"Zoho Chat",asesor:"Nacho",    ciudad:"Rosario",   dispositivo:"Desktop",pagina:"/ecoflow-river",  createdAt:h(0.5),notas:[],nota_rapida:"",historial:[]},
];

// ─── ICONS ───────────────────────────────────────────────────────────────────
const Sv=({size=16,children,...p})=><svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" {...p}>{children}</svg>;
const IcoX       = ({s=15})=><Sv size={s}><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></Sv>;
const IcoChevD   = ({s=13})=><Sv size={s}><polyline points="6 9 12 15 18 9"/></Sv>;
const IcoWA      = ({s=14})=><svg width={s} height={s} viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.19-.008-.398-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413z"/></svg>;
const IcoNote    = ({s=13})=><Sv size={s}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></Sv>;
const IcoEdit    = ({s=13})=><Sv size={s}><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></Sv>;
const IcoTrash   = ({s=13})=><Sv size={s}><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></Sv>;
const IcoSync    = ({s=13})=><Sv size={s}><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></Sv>;
const IcoDl      = ({s=13})=><Sv size={s}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></Sv>;
const IcoPlus    = ({s=13})=><Sv size={s} strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></Sv>;
const IcoCheck   = ({s=12})=><Sv size={s} strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></Sv>;
const IcoBell    = ({s=13})=><Sv size={s}><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></Sv>;
const IcoSend    = ({s=13})=><Sv size={s}><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></Sv>;
const IcoLogout  = ({s=13})=><Sv size={s}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></Sv>;
const IcoEye     = ({s=14})=><Sv size={s}><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></Sv>;
const IcoEyeOff  = ({s=14})=><Sv size={s}><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></Sv>;
const IcoUsers   = ({s=13})=><Sv size={s}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></Sv>;
const IcoChart   = ({s=13})=><Sv size={s}><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></Sv>;
const IcoClock   = ({s=13})=><Sv size={s}><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></Sv>;
const IcoMap     = ({s=13})=><Sv size={s}><polygon points="3 6 9 3 15 6 21 3 21 18 15 21 9 18 3 21"/><line x1="9" y1="3" x2="9" y2="18"/><line x1="15" y1="6" x2="15" y2="21"/></Sv>;
const IcoSidebar = ({s=14})=><Sv size={s}><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="9" y1="3" x2="9" y2="21"/></Sv>;

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function doExport(leads,fname){
  const rows=leads.map(l=>({
    Fecha:fmtDate(l.fecha),Hora:fmtTime(l.createdAt),
    Nombre:l.nombre,Telefono:l.tel,Asesor:l.asesor,
    Producto:l.producto,Fuente:l.fuente,Ciudad:l.ciudad||"",
    Dispositivo:l.dispositivo||"",Pagina:l.pagina||"",
    Estado:EM[l.estado]?.label||l.estado,
    "Nota rapida":l.nota_rapida||"",
    Notas:(l.notas||[]).map(n=>n.texto).join(" | "),
    "Tiempo hasta asignacion (hs)":l.catalogadoAt?((new Date(l.catalogadoAt)-new Date(l.createdAt))/36e5).toFixed(1):"",
    Creado:new Date(l.createdAt).toLocaleString("es-AR"),
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  ws["!cols"]=[10,8,22,14,12,12,12,12,10,20,14,22,40,14,18].map(w=>({wch:w}));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Leads");
  XLSX.writeFile(wb,fname+".xlsx");
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function Login({onLogin}){
  const [u,setU]=useState(""); const [p,setP]=useState("");
  const [show,setShow]=useState(false); const [err,setErr]=useState(""); const [load,setLoad]=useState(false);
  const go=()=>{
    setLoad(true);
    setTimeout(()=>{
      const users=db.getUsers();
      const key=Object.keys(users).find(k=>k===u.toLowerCase());
      const usr=key?users[key]:null;
      if(usr&&usr.pass===p&&usr.activo!==false) onLogin({username:key,...usr});
      else{setErr("Usuario o contraseña incorrectos");setLoad(false);}
    },400);
  };
  const inp={width:"100%",padding:"10px 14px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"8px",color:"#0f172a",fontSize:"14px",fontFamily:"inherit",outline:"none",boxSizing:"border-box"};
  return(
    <div style={{minHeight:"100vh",background:"#f8fafc",display:"flex",alignItems:"center",justifyContent:"center",padding:"24px"}}>
      <div style={{width:"100%",maxWidth:"360px"}}>
        <div style={{textAlign:"center",marginBottom:"28px"}}>
          <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"6px",marginBottom:"8px"}}>
            <img src="data:image/png;base64,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dnlFuh3Kq4JPBwu39Vd7epo3t6mi4H4733/BO/4C+IPgR8SfhL46+JXiPxHr/wAY77TLnxr4t1BYf7T1S20uSIpYhIbaK2gtDCjW4jiiUrFLJhtxDD334x/st/s0/Gv9q/4T/tdeJdUng1z4SR6jHp9nBGBa3Yvowkfn5iL/AOitukgCMoDsd2RgV+he9vU0b29TRdisfi94O/YP134ZeOvGPjD4VftQeN/Ddv478VXni3WLG10vRJo5ry+dS6CW50uaYRrFHHBGof5IkGPmyx7X4j/saW2rftE+Ov2iPgN8fPFfwsvfidDp0fie00iw029S4fSrf7LbyW8moWFy9s6w8fJkbiWx2r9bN7epo3t6mi4zwj9mf9nz4XfssfAzw98B/g1BND4e0KBhA9zKZ7meS4dp5rieRgC808sjyyNgZdjgAYA91oopAFFFFABRRRQAUUUUAFFFFAH/0f76KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9L++iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/T/voooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD/1P76KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9X++iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/Z" alt="TodoDrones" style={{width:"72px",height:"72px",objectFit:"contain",filter:"drop-shadow(0 2px 4px rgba(0,0,0,0.1))"}}/>
            <span style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"24px",fontWeight:800,color:"#0f172a",letterSpacing:"2px"}}>COMBOT</span>
          </div>
          <p style={{fontSize:"12px",color:"#64748b"}}>Gestión comercial interna</p>
        </div>
        <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"12px",padding:"24px",boxShadow:"0 4px 16px rgba(0,0,0,0.08)"}}>
          <div style={{marginBottom:"12px"}}>
            <label style={{display:"block",fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"5px"}}>Usuario</label>
            <input style={inp} value={u} onChange={e=>{setU(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&go()} placeholder="Ingresá tu usuario"/>
          </div>
          <div style={{marginBottom:"18px"}}>
            <label style={{display:"block",fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"5px"}}>Contraseña</label>
            <div style={{position:"relative"}}>
              <input type={show?"text":"password"} style={{...inp,paddingRight:"40px"}} value={p} onChange={e=>{setP(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&go()} placeholder="••••••••"/>
              <button onClick={()=>setShow(v=>!v)} style={{position:"absolute",right:"10px",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#64748b",cursor:"pointer",display:"flex",padding:"2px"}}>
                {show?<IcoEyeOff s={14}/>:<IcoEye s={14}/>}
              </button>
            </div>
          </div>
          {err&&<div style={{background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.18)",borderRadius:"6px",padding:"8px 12px",fontSize:"12px",color:"#f87171",marginBottom:"12px"}}>{err}</div>}
          <button onClick={go} disabled={load} style={{width:"100%",padding:"10px",borderRadius:"8px",border:"none",background:load?"#1B87C966":"#1B87C9",color:"#fff",fontSize:"14px",fontWeight:700,cursor:load?"default":"pointer",fontFamily:"'Barlow Condensed',sans-serif"}}>
            {load?"Verificando...":"Ingresar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── ESTADO DROPDOWN ─────────────────────────────────────────────────────────
function DropOption({est,isActive,onClick}){
  const [hov,setHov]=useState(false);
  return(
    <button
      onClick={onClick}
      onMouseEnter={()=>setHov(true)}
      onMouseLeave={()=>setHov(false)}
      style={{display:"flex",alignItems:"center",gap:"8px",width:"100%",padding:"8px 12px",borderRadius:"6px",border:"none",
        background: hov ? est.color+"18" : isActive ? est.color+"12" : "transparent",
        color: isActive ? est.color : hov ? est.color : "#334155",
        fontSize:"13px",fontWeight:500,cursor:"pointer",textAlign:"left",
        transition:"background 0.1s,color 0.1s",
        boxShadow: hov ? "inset 0 0 0 1px "+est.color+"30" : "none"}}>
      <span style={{width:"7px",height:"7px",borderRadius:"50%",background:est.color,flexShrink:0}}/>{est.label}
      {isActive&&<span style={{marginLeft:"auto",color:est.color}}><IcoCheck s={12}/></span>}
    </button>
  );
}

function EstadoDrop({lead,onUpdate}){
  const [open,setOpen]=useState(false);
  const e=EM[lead.estado]||EM.nuevo;
  return(
    <div style={{position:"relative",display:"inline-block",zIndex:open?200:1}}>
      <button
        onClick={ev=>{ev.stopPropagation();setOpen(o=>!o);}}
        style={{padding:"3px 9px 3px 7px",borderRadius:"20px",background:e.color+"15",color:e.color,fontSize:"11px",fontWeight:600,border:"1px solid "+e.color+"30",cursor:"pointer",display:"flex",alignItems:"center",gap:"4px",whiteSpace:"nowrap"}}>
        <span style={{width:"5px",height:"5px",borderRadius:"50%",background:e.color,flexShrink:0}}/>
        {e.label}<IcoChevD s={9}/>
      </button>
      {open&&(
        <>
          {/* backdrop */}
          <div onClick={()=>setOpen(false)} style={{position:"fixed",inset:0,zIndex:198}}/>
          {/* menu */}
          <div style={{position:"absolute",top:"calc(100% + 4px)",right:0,background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",zIndex:199,minWidth:"160px",padding:"4px",boxShadow:"0 8px 24px rgba(0,0,0,0.12)"}}>
            {ESTADOS.map(est=>(
              <DropOption key={est.id} est={est} isActive={lead.estado===est.id}
                onClick={ev=>{
                  ev.stopPropagation();
                  const now=new Date().toISOString();
                  onUpdate({...lead,estado:est.id,catalogadoAt:now,historial:[...(lead.historial||[]),est.id],autoEsc:false});
                  setOpen(false);
                }}
              />
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ─── NOMBRE EDITABLE ─────────────────────────────────────────────────────────
function NombreEdit({lead,onUpdate}){
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState(lead.nombre);
  if(editing) return(
    <input value={val}
      onChange={e=>setVal(e.target.value)}
      onBlur={()=>{onUpdate({...lead,nombre:val});setEditing(false);}}
      onKeyDown={e=>{if(e.key==="Enter"){onUpdate({...lead,nombre:val});setEditing(false);}if(e.key==="Escape")setEditing(false);}}
      autoFocus
      style={{background:"#ffffff",border:"1px solid #1B87C9",borderRadius:"5px",color:"#0f172a",fontSize:"13px",fontWeight:600,padding:"2px 6px",outline:"none",fontFamily:"inherit",width:"160px",boxShadow:"0 0 0 2px rgba(27,135,201,0.15)"}}
    />
  );
  return(
    <span onClick={()=>setEditing(true)}
      style={{fontSize:"13px",fontWeight:600,color:"#0f172a",cursor:"text",borderBottom:"1px dashed #cbd5e1",paddingBottom:"1px"}}>
      {lead.nombre}
    </span>
  );
}

// ─── PANEL NOTAS ─────────────────────────────────────────────────────────────
function PanelNotas({lead,onUpdate,onClose,session}){
  const [texto,setTexto]=useState("");
  const [notaR,setNotaR]=useState(lead.nota_rapida||"");
  const notas=lead.notas||[];
  const guardarR=()=>onUpdate({...lead,nota_rapida:notaR});
  const agregar=()=>{
    if(!texto.trim())return;
    const n={texto:texto.trim(),ts:new Date().toISOString(),autor:session.nombre};
    onUpdate({...lead,notas:[...notas,n]});
    setTexto("");
  };
  return(
    <div style={{display:"flex",flexDirection:"column",height:"100%",background:"#f8fafc"}}>
      <div style={{padding:"16px 20px",borderBottom:"1px solid #e2e8f0",flexShrink:0}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"10px"}}>
          <div>
            <div style={{fontSize:"16px",fontWeight:700,color:"#0f172a",marginBottom:"4px"}}>{lead.nombre}</div>
            <div style={{display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center"}}>
              {lead.tel&&<a href={"https://wa.me/"+lead.tel.replace(/\D/g,"")} target="_blank" rel="noreferrer"
                style={{display:"flex",alignItems:"center",gap:"3px",color:"#25d366",fontSize:"12px",textDecoration:"none",background:"rgba(37,211,102,0.08)",padding:"2px 7px",borderRadius:"5px"}}>
                <IcoWA s={11}/>{lead.tel}
              </a>}
              <span style={{fontSize:"11px",color:"#64748b"}}>{lead.producto}</span>
              <span style={{fontSize:"11px",color:"#64748b"}}>{lead.asesor}</span>
            </div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",padding:"2px"}}><IcoX s={14}/></button>
        </div>
        <div>
          <label style={{display:"block",fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"5px"}}>Resumen del cliente</label>
          <div style={{display:"flex",gap:"6px"}}>
            <input value={notaR} onChange={e=>setNotaR(e.target.value)} onKeyDown={e=>e.key==="Enter"&&guardarR()}
              placeholder="Ej: Espera cobrar, muy interesado en Mini 4 Pro"
              style={{flex:1,padding:"7px 10px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"12px",outline:"none",fontFamily:"inherit"}}/>
            <button onClick={guardarR} style={{padding:"7px 12px",background:"#1B87C9",border:"none",borderRadius:"6px",color:"#fff",fontSize:"11px",fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>Guardar</button>
          </div>
          {lead.nota_rapida&&<div style={{marginTop:"5px",fontSize:"11px",color:"#64748b",fontStyle:"italic"}}>"{lead.nota_rapida}"</div>}
        </div>
      </div>
      <div style={{flex:1,overflow:"auto",padding:"14px 20px"}}>
        <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px"}}>Historial de seguimiento</div>
        {notas.length===0?(
          <div style={{textAlign:"center",padding:"28px 0",color:"#cbd5e1"}}>
            <div style={{fontSize:"12px",marginBottom:"4px"}}>Sin notas todavía</div>
            <div style={{fontSize:"11px",color:"#e2e8f0"}}>Anotá lo que hablaron para no perder el hilo</div>
          </div>
        ):(
          <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
            {notas.map((n,i)=>(
              <div key={i} style={{background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"8px",padding:"10px 12px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px"}}>
                  <span style={{fontSize:"11px",fontWeight:600,color:"#1B87C9"}}>{n.autor}</span>
                  <span style={{fontSize:"10px",color:"#94a3b8"}}>{new Date(n.ts).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}</span>
                </div>
                <p style={{fontSize:"12px",color:"#334155",lineHeight:1.5,margin:0}}>{n.texto}</p>
              </div>
            ))}
          </div>
        )}
      </div>
      <div style={{padding:"12px 20px",borderTop:"1px solid #e2e8f0",flexShrink:0}}>
        <textarea value={texto} onChange={e=>setTexto(e.target.value)}
          placeholder="Qué dijo, cuándo llamar, qué frena la venta..."
          style={{width:"100%",minHeight:"64px",padding:"8px 10px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"7px",color:"#0f172a",fontSize:"12px",outline:"none",resize:"none",fontFamily:"inherit",boxSizing:"border-box",lineHeight:1.5}}
          onKeyDown={e=>e.key==="Enter"&&e.metaKey&&agregar()}
        />
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:"6px"}}>
          <button onClick={agregar} disabled={!texto.trim()}
            style={{padding:"6px 14px",background:texto.trim()?"#1B87C9":"#e2e8f0",border:"none",borderRadius:"6px",color:texto.trim()?"#fff":"#94a3b8",fontSize:"12px",fontWeight:600,cursor:texto.trim()?"pointer":"default",display:"flex",alignItems:"center",gap:"5px"}}>
            <IcoSend s={12}/>Agregar nota
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── SECCION COLAPSABLE ───────────────────────────────────────────────────────
function SeccionLead({titulo,leads,color,onUpdate,onEdit,onNota,isAdmin,defaultOpen=true}){
  const [open,setOpen]=useState(defaultOpen);
  if(!leads.length) return null;
  return(
    <div style={{marginBottom:"8px",border:"1px solid #e2e8f0",borderRadius:"10px",overflow:"visible"}}>
      <button onClick={()=>setOpen(o=>!o)}
        style={{width:"100%",display:"flex",alignItems:"center",gap:"8px",padding:"10px 14px",background:"#ffffff",border:"none",cursor:"pointer",textAlign:"left"}}>
        <span style={{width:"7px",height:"7px",borderRadius:"50%",background:color,flexShrink:0}}/>
        <span style={{fontSize:"13px",fontWeight:600,color:"#0f172a"}}>{titulo}</span>
        <span style={{background:color+"18",color:color,fontSize:"10px",fontWeight:700,padding:"1px 7px",borderRadius:"10px"}}>{leads.length}</span>
        <span style={{marginLeft:"auto",color:"#94a3b8",transform:open?"rotate(180deg)":"none",transition:"transform 0.15s"}}><IcoChevD s={12}/></span>
      </button>
      {open&&(
        <div style={{borderTop:"1px solid #e2e8f0"}}>
          {leads.map(l=>(
            <div key={l.id} style={{display:"flex",alignItems:"center",gap:"10px",padding:"9px 14px",borderBottom:"1px solid #f1f5f9",position:"relative"}}
              onMouseEnter={e=>e.currentTarget.style.background="#f1f5f9"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:"8px",flexWrap:"wrap"}}>
                  <NombreEdit lead={l} onUpdate={onUpdate}/>
                  {l.tel&&<a href={"https://wa.me/"+l.tel.replace(/\D/g,"")} target="_blank" rel="noreferrer"
                    style={{display:"flex",alignItems:"center",gap:"3px",color:"#25d366",fontSize:"11px",textDecoration:"none"}}>
                    <IcoWA s={11}/>{l.tel}
                  </a>}
                  {isAdmin&&<span style={{fontSize:"11px",color:"#64748b",fontWeight:500}}>{l.asesor}</span>}
                  <span style={{fontSize:"11px",color:"#334155",fontWeight:600,background:"#f1f5f9",padding:"1px 7px",borderRadius:"5px"}}>{l.producto}</span>
                  <span style={{fontSize:"11px",color:"#94a3b8"}}>·</span>
                  <span style={{fontSize:"11px",color:"#475569",fontWeight:500}}>{fmtDate(l.fecha)}</span>
                </div>
                {l.nota_rapida&&<div style={{fontSize:"11px",color:"#64748b",marginTop:"2px",fontStyle:"italic"}}>{l.nota_rapida}</div>}
              </div>
              <EstadoDrop lead={l} onUpdate={onUpdate}/>
              <button onClick={()=>onNota(l)}
                style={{display:"flex",alignItems:"center",gap:"3px",padding:"3px 7px",background:(l.notas||[]).length>0?"rgba(27,135,201,0.1)":"#f8fafc",border:"1px solid "+(l.notas||[]).length>0?"rgba(27,135,201,0.25)":"#e2e8f0",borderRadius:"5px",color:(l.notas||[]).length>0?"#1B87C9":"#64748b",cursor:"pointer",fontSize:"10px",fontWeight:600,whiteSpace:"nowrap"}}>
                <IcoNote s={11}/>{(l.notas||[]).length>0?l.notas.length:""}
              </button>
              <button onClick={()=>onEdit(l)} style={{background:"#f8fafc",border:"1px solid #e2e8f0",color:"#64748b",borderRadius:"5px",padding:"4px",cursor:"pointer",display:"flex"}}><IcoEdit s={11}/></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── VISTA NUEVOS ─────────────────────────────────────────────────────────────
function VistaNuevos({leads,onUpdate,onEdit,onNota,isAdmin,dateRange,setDateRange,customFrom,setCustomFrom,customTo,setCustomTo}){
  const nuevos=leads.filter(l=>l.estado==="nuevo");
  return(
    <div>
      <div style={{marginBottom:"16px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:"10px",marginBottom:"4px"}}>
          <h2 style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"22px",fontWeight:800,color:"#0f172a"}}>Leads nuevos</h2>
          <div style={{display:"flex",alignItems:"center",gap:"6px",flexWrap:"wrap"}}>
            <div style={{display:"flex",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"8px",padding:"3px",gap:"2px"}}>
              {[["today","Hoy"],["week","7 días"],["month","30 días"],["custom","Personalizado"]].map(([v,lb])=>(
                <button key={v} onClick={()=>setDateRange(v)}
                  style={{padding:"4px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"11px",fontWeight:600,whiteSpace:"nowrap",
                    background:dateRange===v?"#1B87C9":"transparent",
                    color:dateRange===v?"#fff":"#64748b",transition:"all 0.15s"}}>
                  {lb}
                </button>
              ))}
            </div>
            {dateRange==="custom"&&(
              <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
                <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)}
                  style={{padding:"4px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
                <span style={{fontSize:"11px",color:"#94a3b8"}}>—</span>
                <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)}
                  style={{padding:"4px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
              </div>
            )}
          </div>
        </div>
        <p style={{fontSize:"12px",color:"#64748b"}}>{new Date().toLocaleDateString("es-AR",{weekday:"long",day:"numeric",month:"long"})} · {nuevos.length} sin asignar · desaparecen a las 24hs</p>
      </div>
      {nuevos.length===0?(
        <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",padding:"48px",textAlign:"center",color:"#94a3b8"}}>
          <div style={{fontSize:"13px",marginBottom:"4px"}}>Sin leads nuevos hoy</div>
          <div style={{fontSize:"11px"}}>Los leads del día aparecen acá cuando entran por Zoho</div>
        </div>
      ):(
        <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",overflow:"visible"}}>
          {nuevos.map((l,i)=>(
            <div key={l.id} style={{display:"flex",alignItems:"center",gap:"12px",padding:"11px 16px",borderBottom:i<nuevos.length-1?"1px solid #f8fafc":"none",transition:"background 0.1s"}}
              onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              {/* Urgencia */}
              <div style={{width:"6px",height:"6px",borderRadius:"50%",flexShrink:0,background:horasDesde(l.createdAt)>8?"#ef4444":"#e2e8f0",boxShadow:horasDesde(l.createdAt)>8?"0 0 5px #ef4444":"none"}}/>
              {/* Hora */}
              <div style={{fontSize:"11px",color:"#94a3b8",whiteSpace:"nowrap",width:"38px",flexShrink:0}}>{fmtTime(l.createdAt)}</div>
              {/* Nombre editable */}
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:"8px",flexWrap:"wrap"}}>
                  <NombreEdit lead={l} onUpdate={onUpdate}/>
                  {l.tel&&(
                    <a href={"https://wa.me/"+l.tel.replace(/\D/g,"")} target="_blank" rel="noreferrer"
                      style={{display:"inline-flex",alignItems:"center",gap:"3px",background:"rgba(37,211,102,0.1)",border:"1px solid rgba(37,211,102,0.2)",color:"#25d366",fontSize:"11px",textDecoration:"none",padding:"2px 8px",borderRadius:"5px",whiteSpace:"nowrap"}}>
                      <IcoWA s={11}/> WhatsApp
                    </a>
                  )}
                  <span style={{fontSize:"11px",color:"#94a3b8"}}>{l.producto}</span>
                  {isAdmin&&<span style={{fontSize:"11px",color:"#94a3b8"}}>{l.asesor}</span>}
                </div>
                {l.nota_rapida&&<div style={{fontSize:"11px",color:"#64748b",marginTop:"2px",fontStyle:"italic"}}>{l.nota_rapida}</div>}
              </div>
              <EstadoDrop lead={l} onUpdate={onUpdate}/>
              <button onClick={()=>onNota(l)}
                style={{display:"flex",alignItems:"center",gap:"3px",padding:"4px 8px",background:(l.notas||[]).length>0?"rgba(27,135,201,0.1)":"#f8fafc",border:"1px solid "+(l.notas||[]).length>0?"rgba(27,135,201,0.25)":"#e2e8f0",borderRadius:"6px",color:(l.notas||[]).length>0?"#1B87C9":"#64748b",cursor:"pointer",fontSize:"11px",fontWeight:600}}>
                <IcoNote s={12}/>{(l.notas||[]).length>0?" "+l.notas.length:""}
              </button>
              <button onClick={()=>onEdit(l)} style={{background:"#f8fafc",border:"1px solid #e2e8f0",color:"#64748b",borderRadius:"6px",padding:"5px",cursor:"pointer",display:"flex"}}><IcoEdit s={11}/></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── VISTA SEGUIMIENTO ────────────────────────────────────────────────────────
function VistaSeguimiento({leads,onUpdate,onEdit,onNota,isAdmin}){
  const rc=leads.filter(l=>l.estado==="recontactar");
  const pe=leads.filter(l=>l.estado==="perdido");
  return(
    <div>
      <div style={{marginBottom:"14px"}}>
        <h2 style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"22px",fontWeight:800,color:"#0f172a",marginBottom:"2px"}}>Seguimiento</h2>
        <p style={{fontSize:"12px",color:"#64748b"}}>{rc.length} para recontactar · {pe.length} perdidos</p>
      </div>
      <SeccionLead titulo="Recontactar" leads={rc} color="#f59e0b" onUpdate={onUpdate} onEdit={onEdit} onNota={onNota} isAdmin={isAdmin}/>
      <SeccionLead titulo="Perdidos" leads={pe} color="#ef4444" onUpdate={onUpdate} onEdit={onEdit} onNota={onNota} isAdmin={isAdmin} defaultOpen={true}/>
    </div>
  );
}

// ─── VISTA TODOS ──────────────────────────────────────────────────────────────
function VistaTodos({leads,allLeads,onUpdate,onEdit,onNota,isAdmin,dateRange,setDateRange,customFrom,setCustomFrom,customTo,setCustomTo}){
  // Recontactar: always show ALL regardless of date filter (sticky)
  const recontactarLeads = allLeads.filter(l=>l.estado==="recontactar");
  // All other states: use the already-filtered leads
  const otrosEstados = ESTADOS.filter(e=>e.id!=="recontactar");
  return(
    <div>
      <div style={{marginBottom:"16px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:"10px",marginBottom:"4px"}}>
          <h2 style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"22px",fontWeight:800,color:"#0f172a"}}>Todos los leads</h2>
          <div style={{display:"flex",alignItems:"center",gap:"6px",flexWrap:"wrap"}}>
            <div style={{display:"flex",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"8px",padding:"3px",gap:"2px"}}>
              {[["today","Hoy"],["week","7 días"],["month","30 días"],["custom","Personalizado"]].map(([v,lb])=>(
                <button key={v} onClick={()=>setDateRange(v)}
                  style={{padding:"4px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"11px",fontWeight:600,whiteSpace:"nowrap",
                    background:dateRange===v?"#1B87C9":"transparent",
                    color:dateRange===v?"#fff":"#64748b",transition:"all 0.15s"}}>
                  {lb}
                </button>
              ))}
            </div>
            {dateRange==="custom"&&(
              <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
                <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)}
                  style={{padding:"4px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
                <span style={{fontSize:"11px",color:"#94a3b8"}}>—</span>
                <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)}
                  style={{padding:"4px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
              </div>
            )}
          </div>
        </div>
        <p style={{fontSize:"12px",color:"#64748b"}}>{leads.length} leads en el período · {recontactarLeads.length} para recontactar (siempre visibles)</p>
      </div>

      {/* Recontactar: sticky, ignora filtro de fecha */}
      {recontactarLeads.length>0&&(
        <div style={{marginBottom:"8px"}}>

          <SeccionLead titulo="Recontactar" leads={recontactarLeads}
            color="#f59e0b" onUpdate={onUpdate} onEdit={onEdit} onNota={onNota} isAdmin={isAdmin}
            defaultOpen={true}/>
        </div>
      )}

      {/* Resto de estados con filtro de fecha aplicado */}
      {otrosEstados.map(e=>(
        <SeccionLead key={e.id} titulo={e.label} leads={leads.filter(l=>l.estado===e.id)}
          color={e.color} onUpdate={onUpdate} onEdit={onEdit} onNota={onNota} isAdmin={isAdmin}
          defaultOpen={true}/>
      ))}
    </div>
  );
}

// ─── GESTIÓN DE EQUIPO ────────────────────────────────────────────────────────
function GestionEquipo({onClose}){
  const [users,setUsers]=useState(()=>db.getUsers());
  const [newU,setNewU]=useState({username:"",nombre:"",pass:"",role:"comercial"});
  const [showPw,setShowPw]=useState({});
  const save=(updated)=>{db.setUsers(updated);setUsers(updated);};
  const addUser=()=>{
    if(!newU.username||!newU.nombre||!newU.pass)return;
    const updated={...users,[newU.username.toLowerCase()]:{pass:newU.pass,role:newU.role,nombre:newU.nombre,activo:true}};
    save(updated);setNewU({username:"",nombre:"",pass:"",role:"comercial"});
  };
  const toggleActivo=(key)=>save({...users,[key]:{...users[key],activo:!users[key].activo}});
  const inp={width:"100%",padding:"8px 10px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"13px",fontFamily:"inherit",outline:"none",boxSizing:"border-box"};
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",backdropFilter:"blur(4px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"14px",width:"100%",maxWidth:"520px",maxHeight:"85vh",overflow:"auto"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 20px",borderBottom:"1px solid #e2e8f0"}}>
          <span style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"17px",fontWeight:700,color:"#0f172a"}}>Gestión de equipo</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer"}}><IcoX s={14}/></button>
        </div>
        <div style={{padding:"20px"}}>
          {/* Usuarios existentes */}
          <div style={{marginBottom:"20px"}}>
            <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px"}}>Equipo actual</div>
            {Object.entries(users).map(([key,u])=>(
              <div key={key} style={{display:"flex",alignItems:"center",gap:"10px",padding:"10px 12px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"8px",marginBottom:"6px"}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:"7px"}}>
                    <span style={{fontSize:"13px",fontWeight:600,color:u.activo!==false?"#0f172a":"#94a3b8"}}>{u.nombre}</span>
                    <span style={{fontSize:"10px",color:u.role==="admin"?"#1B87C9":"#64748b",background:u.role==="admin"?"rgba(27,135,201,0.12)":"#e2e8f0",padding:"1px 6px",borderRadius:"4px"}}>{u.role}</span>
                    {u.activo===false&&<span style={{fontSize:"10px",color:"#ef4444",background:"rgba(239,68,68,0.1)",padding:"1px 6px",borderRadius:"4px"}}>Inactivo</span>}
                  </div>
                  <div style={{fontSize:"11px",color:"#94a3b8",marginTop:"2px"}}>@{key}</div>
                </div>
                <div style={{position:"relative"}}>
                  <input type={showPw[key]?"text":"password"} defaultValue={u.pass} readOnly
                    style={{...inp,width:"110px",fontSize:"12px",paddingRight:"28px",color:"#64748b"}}/>
                  <button onClick={()=>setShowPw(p=>({...p,[key]:!p[key]}))}
                    style={{position:"absolute",right:"6px",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#94a3b8",cursor:"pointer",display:"flex",padding:"1px"}}>
                    {showPw[key]?<IcoEyeOff s={12}/>:<IcoEye s={12}/>}
                  </button>
                </div>
                {key!=="admin"&&(
                  <button onClick={()=>toggleActivo(key)}
                    style={{padding:"4px 10px",background:u.activo!==false?"rgba(239,68,68,0.08)":"rgba(16,185,129,0.08)",border:"1px solid "+( u.activo!==false?"rgba(239,68,68,0.2)":"rgba(16,185,129,0.2)"),borderRadius:"6px",color:u.activo!==false?"#ef4444":"#10b981",cursor:"pointer",fontSize:"11px",fontWeight:600,whiteSpace:"nowrap"}}>
                    {u.activo!==false?"Desactivar":"Activar"}
                  </button>
                )}
              </div>
            ))}
          </div>
          {/* Agregar nuevo */}
          <div style={{borderTop:"1px solid #e2e8f0",paddingTop:"16px"}}>
            <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px"}}>Agregar comercial</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
              <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"4px"}}>Usuario (para login)</label><input style={inp} value={newU.username} onChange={e=>setNewU(p=>({...p,username:e.target.value}))} placeholder="Ej: pablo"/></div>
              <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"4px"}}>Nombre visible</label><input style={inp} value={newU.nombre} onChange={e=>setNewU(p=>({...p,nombre:e.target.value}))} placeholder="Ej: Pablo"/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"12px"}}>
              <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"4px"}}>Contraseña</label><input style={inp} value={newU.pass} onChange={e=>setNewU(p=>({...p,pass:e.target.value}))} placeholder="Contraseña"/></div>
              <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"4px"}}>Rol</label>
                <select style={inp} value={newU.role} onChange={e=>setNewU(p=>({...p,role:e.target.value}))}>
                  <option value="comercial">Comercial</option>
                  <option value="admin">Admin</option>
                </select>
              </div>
            </div>
            <button onClick={addUser} style={{width:"100%",padding:"9px",background:"#1B87C9",border:"none",borderRadius:"7px",color:"#fff",fontSize:"13px",fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}>
              <IcoPlus s={13}/>Agregar al equipo
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── FORM LEAD ────────────────────────────────────────────────────────────────
function FormLead({lead,onSave,onClose,session,users}){
  const isAdmin=session.role==="admin";
  const comerciales=Object.entries(users).filter(([,u])=>u.activo!==false).map(([,u])=>u.nombre);
  const [f,setF]=useState(lead||{nombre:"",tel:"",fecha:todayStr(),estado:"nuevo",producto:"DJI",fuente:"Zoho Chat",asesor:isAdmin?comerciales[0]:session.nombre,ciudad:"CABA",dispositivo:"Móvil",pagina:"",notas:[],nota_rapida:""});
  const s=(k,v)=>setF(x=>({...x,[k]:v}));
  const inp={width:"100%",padding:"8px 10px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"13px",fontFamily:"inherit",outline:"none",boxSizing:"border-box"};
  const lbl={display:"block",fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"4px"};
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",backdropFilter:"blur(4px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"14px",width:"100%",maxWidth:"480px",boxShadow:"0 8px 32px rgba(0,0,0,0.12)",maxHeight:"88vh",overflow:"auto"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 20px",borderBottom:"1px solid #e2e8f0"}}>
          <span style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"17px",fontWeight:700,color:"#0f172a"}}>{lead?"Editar lead":"Nuevo lead"}</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer"}}><IcoX s={14}/></button>
        </div>
        <div style={{padding:"20px",display:"flex",flexDirection:"column",gap:"10px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"}}>
            <div><label style={lbl}>Nombre</label><input style={inp} value={f.nombre} onChange={e=>s("nombre",e.target.value)} placeholder="Nombre del cliente"/></div>
            <div><label style={lbl}>Telefono</label><input style={inp} value={f.tel} onChange={e=>s("tel",e.target.value)} placeholder="+54 11..."/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"}}>
            <div><label style={lbl}>Fecha</label><input type="date" style={inp} value={f.fecha} onChange={e=>s("fecha",e.target.value)}/></div>
            <div><label style={lbl}>Asesor</label>
              {isAdmin?<select style={inp} value={f.asesor} onChange={e=>s("asesor",e.target.value)}>{comerciales.map(a=><option key={a}>{a}</option>)}</select>
              :<input style={{...inp,color:"#64748b"}} value={f.asesor} readOnly/>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"}}>
            <div><label style={lbl}>Producto</label><select style={inp} value={f.producto} onChange={e=>s("producto",e.target.value)}>{PRODUCTOS.map(p=><option key={p}>{p}</option>)}</select></div>
            <div><label style={lbl}>Fuente</label><select style={inp} value={f.fuente} onChange={e=>s("fuente",e.target.value)}>{FUENTES.map(x=><option key={x}>{x}</option>)}</select></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"}}>
            <div><label style={lbl}>Ciudad</label><select style={inp} value={f.ciudad||"CABA"} onChange={e=>s("ciudad",e.target.value)}>{CIUDADES.map(c=><option key={c}>{c}</option>)}</select></div>
            <div><label style={lbl}>Dispositivo</label><select style={inp} value={f.dispositivo||"Móvil"} onChange={e=>s("dispositivo",e.target.value)}>{DISPOSITIVOS.map(d=><option key={d}>{d}</option>)}</select></div>
          </div>
          <div><label style={lbl}>Pagina de origen (URL)</label><input style={inp} value={f.pagina||""} onChange={e=>s("pagina",e.target.value)} placeholder="/dji-mini-4-pro"/></div>
          <div>
            <label style={lbl}>Estado</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:"5px"}}>
              {ESTADOS.map(e=><button key={e.id} onClick={()=>s("estado",e.id)}
                style={{padding:"3px 10px",borderRadius:"20px",border:"1px solid "+(f.estado===e.id?e.color:"#e2e8f0"),background:f.estado===e.id?e.color+"12":"transparent",color:f.estado===e.id?e.color:"#64748b",fontSize:"11px",cursor:"pointer"}}>
                {e.label}
              </button>)}
            </div>
          </div>
          <div><label style={lbl}>Resumen del cliente</label><input style={inp} value={f.nota_rapida||""} onChange={e=>s("nota_rapida",e.target.value)} placeholder="Ej: Espera cobrar, muy interesado en Mini 4 Pro"/></div>
          <div style={{display:"flex",gap:"8px",justifyContent:"flex-end",paddingTop:"4px"}}>
            <button onClick={onClose} style={{padding:"8px 16px",borderRadius:"6px",border:"1px solid #e2e8f0",background:"transparent",color:"#64748b",cursor:"pointer",fontSize:"13px"}}>Cancelar</button>
            <button onClick={()=>{if(!f.nombre)return;onSave({...f,id:f.id||uid(),createdAt:f.createdAt||new Date().toISOString(),historial:f.historial||[]});}}
              style={{padding:"8px 16px",borderRadius:"6px",border:"none",background:"#1B87C9",color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:700}}>Guardar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── DASHBOARD MKT ───────────────────────────────────────────────────────────
function DashMKT({leads,users}){
  const [vista,     setVista]    = useState("total");
  const [periodo,   setPeriodo]  = useState("month");
  const [comSel,    setComSel]   = useState("");
  const [periodoOpen,setPeriodoOpen] = useState(false);
  const [mktFrom,   setMktFrom]  = useState("");
  const [mktTo,     setMktTo]    = useState("");

  const last7  = getLast7();
  const last30 = getLast30();

  const filterPeriodo = arr => {
    if(periodo==="today") return arr.filter(l=>l.fecha===todayStr());
    if(periodo==="week")  return arr.filter(l=>last7.includes(l.fecha));
    if(periodo==="month") return arr.filter(l=>last30.includes(l.fecha));
    if(periodo==="custom"&&mktFrom&&mktTo) return arr.filter(l=>l.fecha>=mktFrom&&l.fecha<=mktTo);
    return arr;
  };
  const PERIODO_LABELS = {"today":"Hoy","week":"7 días","month":"30 días","custom":"Personalizado","all":"Todo"};

  const comerciales = Object.entries(users||{})
    .filter(([,u])=>u.role==="comercial"&&u.activo!==false)
    .map(([,u])=>u.nombre);

  // Base de datos según vista
  const base = filterPeriodo(
    vista==="comercial"&&comSel
      ? leads.filter(l=>l.asesor===comSel)
      : leads
  );

  const total       = base.length;
  const vendidos    = base.filter(l=>l.estado==="vendido").length;
  const perdidos    = base.filter(l=>l.estado==="perdido").length;
  const recontactar = base.filter(l=>l.estado==="recontactar").length;
  const sinAsignar  = base.filter(l=>l.estado==="sin_asignar").length;
  const conv        = total?Math.round(vendidos/total*100):0;
  const tasaPerd    = total?Math.round(perdidos/total*100):0;

  const conCat   = base.filter(l=>l.catalogadoAt&&l.createdAt);
  const tProm    = conCat.length?((conCat.reduce((a,l)=>a+(new Date(l.catalogadoAt)-new Date(l.createdAt)),0)/conCat.length)/36e5).toFixed(1):null;
  const rapidos  = conCat.length?Math.round(conCat.filter(l=>(new Date(l.catalogadoAt)-new Date(l.createdAt))/36e5<2).length/conCat.length*100):0;
  const rcTotal  = base.filter(l=>(l.historial||[]).includes("recontactar")).length;
  const rcVend   = base.filter(l=>(l.historial||[]).includes("recontactar")&&l.estado==="vendido").length;
  const tasaRC   = rcTotal?Math.round(rcVend/rcTotal*100):0;

  // Volumen 30d
  const vol30    = getLast30().map(d=>({d:fmtShort(d),n:base.filter(l=>l.fecha===d).length}));
  const vol7     = last7.map(d=>({d:fmtShort(d),Nuevos:base.filter(l=>l.fecha===d&&l.estado==="nuevo").length,Vendidos:base.filter(l=>l.fecha===d&&l.estado==="vendido").length,Perdidos:base.filter(l=>l.fecha===d&&l.estado==="perdido").length}));
  const porDiaSem= DIAS_SEMANA.map((dia,i)=>({dia,n:base.filter(l=>new Date(l.createdAt+"").getDay()===i).length}));
  const porHora  = Array.from({length:24},(_,h)=>({hora:`${h}h`,n:base.filter(l=>new Date(l.createdAt).getHours()===h).length})).filter(h=>h.n>0);
  const porProd  = PRODUCTOS.map(p=>({name:p,total:base.filter(l=>l.producto===p).length,vendidos:base.filter(l=>l.producto===p&&l.estado==="vendido").length})).filter(p=>p.total>0);
  const porFuente= FUENTES.map(f=>({name:f,total:base.filter(l=>l.fuente===f).length,vendidos:base.filter(l=>l.fuente===f&&l.estado==="vendido").length})).filter(f=>f.total>0);
  const porCiudad= CIUDADES.map(c=>({name:c,n:base.filter(l=>l.ciudad===c).length})).filter(c=>c.n>0).sort((a,b)=>b.n-a.n).slice(0,6);
  const porDisp  = DISPOSITIVOS.map(d=>({name:d,value:base.filter(l=>l.dispositivo===d).length})).filter(d=>d.value>0);
  const pags={};base.forEach(l=>{if(l.pagina){pags[l.pagina]=(pags[l.pagina]||0)+1;}});
  const topPags  = Object.entries(pags).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([p,n])=>({p,n}));

  // Por comercial (siempre total, para tabla comparativa)
  const statsCom = comerciales.map(a=>{
    const al=filterPeriodo(leads.filter(l=>l.asesor===a));
    const av=al.filter(l=>l.estado==="vendido").length;
    const ap=al.filter(l=>l.estado==="perdido").length;
    const ar=al.filter(l=>l.estado==="recontactar").length;
    const at=al.length?Math.round(av/al.length*100):0;
    const cat=al.filter(l=>l.catalogadoAt&&l.createdAt);
    const tp=cat.length?((cat.reduce((a,l)=>a+(new Date(l.catalogadoAt)-new Date(l.createdAt)),0)/cat.length)/36e5).toFixed(1):"-";
    return {nombre:a,total:al.length,vendidos:av,perdidos:ap,recontactar:ar,conv:at,tResp:tp};
  });

  const COLS=["#1B87C9","#10b981","#f59e0b","#a78bfa","#f97316","#06b6d4"];
  const CARD={background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",padding:"16px",boxShadow:"0 1px 3px rgba(0,0,0,0.04)"};
  const SK={fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px",display:"block"};

  return(
    <div style={{padding:"20px",overflow:"auto",height:"100%"}}>

      {/* Header con controles */}
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:"20px",flexWrap:"wrap",gap:"12px"}}>
        <div>
          <h2 style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"22px",fontWeight:800,color:"#0f172a",marginBottom:"2px"}}>Dashboard de Marketing</h2>
          <p style={{fontSize:"12px",color:"#64748b"}}>Métricas agregadas para análisis de campañas — sin datos de clientes</p>
        </div>
        <div style={{display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center"}}>
          {/* Período — dropdown */}
          <div style={{position:"relative"}}>
            <button onClick={()=>setPeriodoOpen(o=>!o)}
              style={{display:"flex",alignItems:"center",gap:"6px",padding:"6px 12px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"8px",cursor:"pointer",fontSize:"12px",fontWeight:600,color:"#0f172a",boxShadow:"0 1px 2px rgba(0,0,0,0.04)"}}>
              {PERIODO_LABELS[periodo]||"Período"} <IcoChevD s={11}/>
            </button>
            {periodoOpen&&(
              <>
                <div onClick={()=>setPeriodoOpen(false)} style={{position:"fixed",inset:0,zIndex:98}}/>
                <div style={{position:"absolute",top:"calc(100%+4px)",left:0,background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",zIndex:99,minWidth:"160px",padding:"4px",boxShadow:"0 8px 24px rgba(0,0,0,0.1)"}}>
                  {[["today","Hoy"],["week","7 días"],["month","30 días"],["all","Todo"],["custom","Personalizado"]].map(([v,lb])=>(
                    <button key={v} onClick={()=>{setPeriodo(v);setPeriodoOpen(false);}}
                      style={{display:"flex",alignItems:"center",gap:"8px",width:"100%",padding:"8px 12px",borderRadius:"6px",border:"none",
                        background:periodo===v?"#f0f7ff":"transparent",
                        color:periodo===v?"#1B87C9":"#334155",
                        fontSize:"13px",fontWeight:periodo===v?600:400,cursor:"pointer",textAlign:"left"}}>
                      {lb}{periodo===v&&<span style={{marginLeft:"auto"}}><IcoCheck s={11}/></span>}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
          {periodo==="custom"&&(
            <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
              <input type="date" value={mktFrom} onChange={e=>setMktFrom(e.target.value)}
                style={{padding:"5px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
              <span style={{fontSize:"11px",color:"#94a3b8"}}>—</span>
              <input type="date" value={mktTo} onChange={e=>setMktTo(e.target.value)}
                style={{padding:"5px 8px",background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"11px",outline:"none",fontFamily:"inherit"}}/>
            </div>
          )}
          {/* Vista */}
          <div style={{display:"flex",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"8px",padding:"3px",gap:"2px"}}>
            <button onClick={()=>setVista("total")}
              style={{padding:"4px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"11px",fontWeight:600,
                background:vista==="total"?"#0f172a":"transparent",color:vista==="total"?"#fff":"#64748b"}}>
              Total empresa
            </button>
            <button onClick={()=>{setVista("comercial");if(!comSel&&comerciales[0])setComSel(comerciales[0]);}}
              style={{padding:"4px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"11px",fontWeight:600,
                background:vista==="comercial"?"#0f172a":"transparent",color:vista==="comercial"?"#fff":"#64748b"}}>
              Por comercial
            </button>
          </div>
          {/* Selector comercial */}
          {vista==="comercial"&&(
            <div style={{display:"flex",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"8px",padding:"3px",gap:"2px"}}>
              {comerciales.map(a=>(
                <button key={a} onClick={()=>setComSel(a)}
                  style={{padding:"4px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"11px",fontWeight:600,
                    background:comSel===a?"#1B87C9":"transparent",color:comSel===a?"#fff":"#64748b"}}>
                  {a}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* KPIs fila 1 */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"10px",marginBottom:"12px"}}>
        {[["Total leads",total,"#3b82f6","Período seleccionado"],["Tasa de cierre",conv+"%","#10b981","Leads que compraron"],["Leads perdidos",tasaPerd+"%","#ef4444","No avanzaron"],["Recontactar",recontactar,"#f59e0b","Pipeline activo"]].map(([lb,v,c,sub])=>(
          <div key={lb} style={{...CARD,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:"3px",background:c}}/>
            <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"4px"}}>{lb}</div>
            <div style={{fontSize:"34px",fontWeight:800,color:c,fontFamily:"'Barlow Condensed',sans-serif",lineHeight:1}}>{v}</div>
            <div style={{fontSize:"10px",color:"#94a3b8",marginTop:"3px"}}>{sub}</div>
          </div>
        ))}
      </div>

      {/* KPIs fila 2 */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"10px",marginBottom:"16px"}}>
        {[["Sin asignar",sinAsignar,"#f97316","Sin respuesta +24hs"],["Tasa rescate RC",tasaRC+"%","#a78bfa","Recontactar → Venta"],["Resp. en <2hs",rapidos+"%","#06b6d4","Tasa de resp. rápida"],["T. prom. resp.",tProm?tProm+"hs":"—","#f59e0b","Tiempo hasta asignar"]].map(([lb,v,c,sub])=>(
          <div key={lb} style={{...CARD,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:"3px",background:c}}/>
            <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"4px"}}>{lb}</div>
            <div style={{fontSize:"34px",fontWeight:800,color:c,fontFamily:"'Barlow Condensed',sans-serif",lineHeight:1}}>{v}</div>
            <div style={{fontSize:"10px",color:"#94a3b8",marginTop:"3px"}}>{sub}</div>
          </div>
        ))}
      </div>

      {/* Tabla comparativa por comercial — siempre visible */}
      <div style={{...CARD,marginBottom:"16px"}}>
        <span style={SK}>Comparativa por comercial — período seleccionado</span>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{borderBottom:"1px solid #f1f5f9"}}>
              {["Comercial","Leads","Vendidos","Perdidos","Rcont.","Conv.","T. resp."].map(h=>(
                <th key={h} style={{padding:"6px 12px",textAlign:"left",fontSize:"10px",fontWeight:700,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.8px"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {statsCom.map((s,i)=>(
              <tr key={s.nombre} style={{borderBottom:"1px solid #f8fafc"}}>
                <td style={{padding:"10px 12px",fontSize:"13px",fontWeight:600,color:"#0f172a"}}>{s.nombre}</td>
                <td style={{padding:"10px 12px",fontSize:"13px",color:"#334155"}}>{s.total}</td>
                <td style={{padding:"10px 12px"}}><span style={{fontSize:"13px",fontWeight:700,color:"#10b981"}}>{s.vendidos}</span></td>
                <td style={{padding:"10px 12px"}}><span style={{fontSize:"13px",color:"#ef4444"}}>{s.perdidos}</span></td>
                <td style={{padding:"10px 12px"}}><span style={{fontSize:"13px",color:"#f59e0b"}}>{s.recontactar}</span></td>
                <td style={{padding:"10px 12px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                    <span style={{fontSize:"13px",fontWeight:700,color:s.conv>20?"#10b981":s.conv>10?"#f59e0b":"#ef4444"}}>{s.conv}%</span>
                    <div style={{flex:1,background:"#f1f5f9",borderRadius:"99px",height:"5px",overflow:"hidden",minWidth:"60px"}}>
                      <div style={{height:"100%",background:s.conv>20?"#10b981":s.conv>10?"#f59e0b":"#ef4444",width:s.conv+"%",borderRadius:"99px"}}/>
                    </div>
                  </div>
                </td>
                <td style={{padding:"10px 12px",fontSize:"12px",color:"#64748b"}}>{s.tResp}hs</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Gráficos */}
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:"12px",marginBottom:"12px"}}>
        <div style={CARD}>
          <span style={SK}>Volumen de leads — últimos 30 días</span>
          <ResponsiveContainer width="100%" height={160}>
            <LineChart data={vol30} margin={{top:0,right:0,left:-20,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false}/>
              <XAxis dataKey="d" tick={{fill:"#94a3b8",fontSize:9}} axisLine={false} tickLine={false} interval={4}/>
              <YAxis tick={{fill:"#94a3b8",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/>
              <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",fontSize:"11px",color:"#0f172a"}}/>
              <Line type="monotone" dataKey="n" stroke="#1B87C9" strokeWidth={2} dot={false} name="Leads"/>
            </LineChart>
          </ResponsiveContainer>
        </div>
        <div style={CARD}>
          <span style={SK}>Por dispositivo</span>
          {porDisp.length===0?<div style={{color:"#94a3b8",fontSize:"12px",textAlign:"center",padding:"30px 0"}}>Sin datos</div>:(
            <>
              <ResponsiveContainer width="100%" height={110}>
                <PieChart>
                  <Pie data={porDisp} cx="50%" cy="50%" innerRadius={30} outerRadius={50} paddingAngle={3} dataKey="value">
                    {porDisp.map((_,i)=><Cell key={i} fill={COLS[i%COLS.length]}/>)}
                  </Pie>
                  <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",fontSize:"11px"}}/>
                </PieChart>
              </ResponsiveContainer>
              {porDisp.map((d,i)=>(
                <div key={d.name} style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:"5px"}}>
                    <div style={{width:"6px",height:"6px",borderRadius:"50%",background:COLS[i%COLS.length]}}/>
                    <span style={{fontSize:"11px",color:"#64748b"}}>{d.name}</span>
                  </div>
                  <span style={{fontSize:"11px",fontWeight:600,color:"#334155"}}>{d.value} ({total?Math.round(d.value/total*100):0}%)</span>
                </div>
              ))}
            </>
          )}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
        <div style={CARD}>
          <span style={SK}>Leads por día de la semana</span>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={porDiaSem} barSize={22} margin={{top:0,right:0,left:-20,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false}/>
              <XAxis dataKey="dia" tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#94a3b8",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/>
              <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",fontSize:"11px",color:"#0f172a"}} cursor={{fill:"#f8fafc"}}/>
              <Bar dataKey="n" name="Leads" radius={[3,3,0,0]}>
                {porDiaSem.map((_,i)=><Cell key={i} fill={i===0||i===6?"#cbd5e1":"#1B87C9"}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <div style={{fontSize:"10px",color:"#94a3b8",marginTop:"6px"}}>Gris = fin de semana · Distribuí el presupuesto según los picos</div>
        </div>
        <div style={CARD}>
          <span style={SK}>Leads por hora del día</span>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={porHora} barSize={8} margin={{top:0,right:0,left:-20,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false}/>
              <XAxis dataKey="hora" tick={{fill:"#94a3b8",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#94a3b8",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/>
              <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",fontSize:"11px",color:"#0f172a"}} cursor={{fill:"#f8fafc"}}/>
              <Bar dataKey="n" fill="#a78bfa" radius={[2,2,0,0]} name="Leads"/>
            </BarChart>
          </ResponsiveContainer>
          <div style={{fontSize:"10px",color:"#94a3b8",marginTop:"6px"}}>Picos = cuándo concentrar presupuesto en Meta/Google</div>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
        <div style={CARD}>
          <span style={SK}>Producto — leads vs cierres</span>
          {porProd.length===0?<div style={{color:"#94a3b8",fontSize:"12px",padding:"20px 0",textAlign:"center"}}>Sin datos</div>:
            porProd.map(p=>{
              const tasa=p.total?Math.round(p.vendidos/p.total*100):0;
              return(
                <div key={p.name} style={{marginBottom:"10px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}>
                    <span style={{fontSize:"12px",color:"#334155",fontWeight:500}}>{p.name}</span>
                    <span style={{fontSize:"11px",color:"#64748b"}}>{p.vendidos}/{p.total} · <span style={{color:"#10b981",fontWeight:600}}>{tasa}%</span></span>
                  </div>
                  <div style={{background:"#f1f5f9",borderRadius:"99px",height:"5px",overflow:"hidden"}}>
                    <div style={{height:"100%",background:"#10b981",width:tasa+"%",borderRadius:"99px"}}/>
                  </div>
                </div>
              );
            })
          }
        </div>
        <div style={CARD}>
          <span style={SK}>Fuente del lead — calidad por canal</span>
          {porFuente.length===0?<div style={{color:"#94a3b8",fontSize:"12px",padding:"20px 0",textAlign:"center"}}>Sin datos</div>:(
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {["Canal","Leads","Vendidos","Conv."].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",fontSize:"10px",color:"#94a3b8",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {porFuente.sort((a,b)=>b.total-a.total).map(f=>{
                  const tasa=f.total?Math.round(f.vendidos/f.total*100):0;
                  return(
                    <tr key={f.name} style={{borderTop:"1px solid #f8fafc"}}>
                      <td style={{padding:"6px 8px",fontSize:"12px",color:"#334155"}}>{f.name}</td>
                      <td style={{padding:"6px 8px",fontSize:"12px",color:"#64748b"}}>{f.total}</td>
                      <td style={{padding:"6px 8px",fontSize:"12px",color:"#10b981",fontWeight:600}}>{f.vendidos}</td>
                      <td style={{padding:"6px 8px",fontSize:"12px",color:tasa>20?"#10b981":tasa>10?"#f59e0b":"#64748b",fontWeight:700}}>{tasa}%</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
        <div style={CARD}>
          <span style={SK}>Ciudad / región</span>
          {porCiudad.length===0?<div style={{color:"#94a3b8",fontSize:"12px",padding:"20px 0",textAlign:"center"}}>Sin datos</div>:
            porCiudad.map((c,i)=>(
              <div key={c.name} style={{marginBottom:"7px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                  <span style={{fontSize:"12px",color:"#64748b"}}>{c.name}</span>
                  <span style={{fontSize:"11px",fontWeight:600,color:"#334155"}}>{c.n}</span>
                </div>
                <div style={{background:"#f1f5f9",borderRadius:"99px",height:"4px",overflow:"hidden"}}>
                  <div style={{height:"100%",background:COLS[i%COLS.length],width:(total?Math.round(c.n/total*100):0)+"%",borderRadius:"99px"}}/>
                </div>
              </div>
            ))
          }
          <div style={{fontSize:"10px",color:"#94a3b8",marginTop:"6px"}}>Segmentación geográfica para Meta Ads</div>
        </div>
        <div style={CARD}>
          <span style={SK}>Páginas que generan más leads</span>
          {topPags.length===0?<div style={{color:"#94a3b8",fontSize:"12px",padding:"20px 0",textAlign:"center"}}>Sin datos de página de origen</div>:
            topPags.map((p,i)=>(
              <div key={p.p} style={{marginBottom:"8px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                  <span style={{fontSize:"11px",color:"#64748b",fontFamily:"monospace"}}>{p.p}</span>
                  <span style={{fontSize:"11px",fontWeight:600,color:"#334155"}}>{p.n}</span>
                </div>
                <div style={{background:"#f1f5f9",borderRadius:"99px",height:"3px",overflow:"hidden"}}>
                  <div style={{height:"100%",background:COLS[i%COLS.length],width:(total?Math.round(p.n/total*100):0)+"%",borderRadius:"99px"}}/>
                </div>
              </div>
            ))
          }
        </div>
      </div>

      <div style={CARD}>
        <span style={SK}>Últimos 7 días — nuevos vs vendidos vs perdidos</span>
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={vol7} barSize={14} margin={{top:0,right:0,left:-20,bottom:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false}/>
            <XAxis dataKey="d" tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fill:"#94a3b8",fontSize:10}} axisLine={false} tickLine={false} allowDecimals={false}/>
            <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"6px",fontSize:"11px",color:"#0f172a"}} cursor={{fill:"#f8fafc"}}/>
            <Bar dataKey="Nuevos" fill="#3b82f6" radius={[2,2,0,0]}/>
            <Bar dataKey="Vendidos" fill="#10b981" radius={[2,2,0,0]}/>
            <Bar dataKey="Perdidos" fill="#ef4444" radius={[2,2,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}


// ─── APP ──────────────────────────────────────────────────────────────────────

function SearchBox({search,setSearch,leads,setTab,setNotaL}){
  const [open,setOpen]=useState(false);
  const ref=useRef(null);

  const results=useMemo(()=>{
    if(!search||search.length<2) return [];
    const q=search.toLowerCase();
    return leads.filter(l=>
      l.nombre?.toLowerCase().includes(q)||l.tel?.includes(q)
    ).slice(0,8);
  },[search,leads]);

  useEffect(()=>{
    const close=ev=>{ if(ref.current&&!ref.current.contains(ev.target)) setOpen(false); };
    document.addEventListener("mousedown",close);
    return()=>document.removeEventListener("mousedown",close);
  },[]);

  const goToLead=(lead)=>{
    const estado=lead.estado;
    if(estado==="nuevo") setTab("nuevos");
    else if(estado==="recontactar"||estado==="perdido") setTab("seguimiento");
    else setTab("todos");
    setSearch("");
    setOpen(false);
    setTimeout(()=>setNotaL(lead),100);
  };

  return(
    <div ref={ref} style={{position:"relative"}}>
      <div style={{position:"relative",display:"flex",alignItems:"center"}}>
        <input
          value={search}
          onChange={e=>{setSearch(e.target.value);setOpen(true);}}
          onFocus={()=>setOpen(true)}
          placeholder="Buscar nombre o número..."
          style={{padding:"4px 9px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"12px",outline:"none",width:"200px",fontFamily:"inherit"}}
        />
        {search&&<button onClick={()=>{setSearch("");setOpen(false);}} style={{position:"absolute",right:"6px",background:"none",border:"none",color:"#94a3b8",cursor:"pointer",padding:"0",display:"flex"}}><IcoX s={11}/></button>}
      </div>
      {open&&results.length>0&&(
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",zIndex:9999,minWidth:"280px",padding:"4px",boxShadow:"0 8px 24px rgba(0,0,0,0.12)"}}>
          <div style={{fontSize:"10px",fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.8px",padding:"6px 10px 4px"}}>Resultados</div>
          {results.map(l=>{
            const e=EM[l.estado]||EM.nuevo;
            return(
              <SearchResult key={l.id} lead={l} e={e} onClick={()=>goToLead(l)}/>
            );
          })}
        </div>
      )}
      {open&&search.length>=2&&results.length===0&&(
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"10px",zIndex:9999,minWidth:"240px",padding:"14px 12px",boxShadow:"0 8px 24px rgba(0,0,0,0.12)",fontSize:"12px",color:"#94a3b8",textAlign:"center"}}>
          Sin resultados para "{search}"
        </div>
      )}
    </div>
  );
}

function SearchResult({lead,e,onClick}){
  const [hov,setHov]=useState(false);
  return(
    <button
      onClick={onClick}
      onMouseEnter={()=>setHov(true)}
      onMouseLeave={()=>setHov(false)}
      style={{display:"flex",alignItems:"center",gap:"10px",width:"100%",padding:"8px 10px",borderRadius:"7px",border:"none",
        background:hov?"#f8fafc":"transparent",cursor:"pointer",textAlign:"left",transition:"background 0.1s"}}>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:"13px",fontWeight:600,color:"#0f172a",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{lead.nombre}</div>
        <div style={{fontSize:"11px",color:"#64748b",marginTop:"1px"}}>{lead.tel} · {lead.producto} · {fmtDate(lead.fecha)}</div>
      </div>
      <span style={{padding:"2px 8px",borderRadius:"12px",background:e.color+"15",color:e.color,fontSize:"10px",fontWeight:700,whiteSpace:"nowrap",border:"1px solid "+e.color+"25",flexShrink:0}}>
        {e.label}
      </span>
    </button>
  );
}

export default function App(){
  const [session,   setSess]     = useState(()=>db.getSess());
  const [leads,     setLeads]    = useState(()=>{const s=db.getLeads();if(!s.length){db.setLeads(SEED);return SEED;}return autoEscalar(s);});
  const [users,     setUsersS]   = useState(()=>db.getUsers());
  const [tab,       setTab]      = useState(session?.role==="admin"?"mkt":"nuevos");
  const [showForm,  setShowForm] = useState(false);
  const [editL,     setEditL]    = useState(null);
  const [notaL,     setNotaL]    = useState(null);
  const [showEquipo,setShowEquipo]=useState(false);
  const [sideOpen,  setSideOpen] = useState(true);
  const [search,    setSearch]   = useState("");
  const [syncing,   setSyncing]  = useState(false);
  const [dateRange, setDateRange] = useState("today"); // today | week | month | custom
  const [customFrom,setCustomFrom]= useState("");
  const [customTo,  setCustomTo]  = useState("");

  useEffect(()=>{db.setLeads(leads);},[leads]);
  useEffect(()=>{const iv=setInterval(()=>setLeads(p=>{const n=autoEscalar(p);return n.some((l,i)=>l.estado!==p[i].estado)?n:p;}),5*60*1000);return()=>clearInterval(iv);},[]);

  const login =u=>{db.setSess(u);setSess(u);};
  const logout=()=>{db.clearSess();setSess(null);};
  const isAdmin=session?.role==="admin";
  const myLeads=isAdmin?leads:leads.filter(l=>l.asesor===session?.nombre);

  const filterByDate = (arr) => {
    const t = todayStr();
    const last7 = getLast7();
    const last30= getLast30();
    if(dateRange==="today")  return arr.filter(l=>l.fecha===t);
    if(dateRange==="week")   return arr.filter(l=>last7.includes(l.fecha));
    if(dateRange==="month")  return arr.filter(l=>last30.includes(l.fecha));
    if(dateRange==="custom"&&customFrom&&customTo) return arr.filter(l=>l.fecha>=customFrom&&l.fecha<=customTo);
    return arr;
  };
  const filteredByDate = filterByDate(myLeads);
  const filtrados=search?filteredByDate.filter(l=>l.nombre?.toLowerCase().includes(search.toLowerCase())||l.tel?.includes(search)):filteredByDate;

  const saveLead=l=>{setLeads(p=>{const i=p.findIndex(x=>x.id===l.id);if(i>=0){const n=[...p];n[i]=l;return n;}return [l,...p];});setShowForm(false);setEditL(null);};
  const updateLead=l=>setLeads(p=>p.map(x=>x.id===l.id?l:x));
  const updateUsers=u=>{db.setUsers(u);setUsersS(u);};

  const doSync=()=>{
    setSyncing(true);
    setTimeout(()=>{
      const comerciales=Object.values(users).filter(u=>u.role==="comercial"&&u.activo!==false).map(u=>u.nombre);
      const nl=[
        {id:uid(),nombre:"Roberto Suárez",tel:"5491166667777",fecha:todayStr(),estado:"nuevo",producto:"DJI",fuente:"Zoho Chat",asesor:comerciales[0]||"Francisco",ciudad:"CABA",dispositivo:"Móvil",pagina:"/dji-avata-2",createdAt:new Date().toISOString(),notas:[],nota_rapida:"",historial:[]},
        {id:uid(),nombre:"Valeria Torres",tel:"5491177778888",fecha:todayStr(),estado:"nuevo",producto:"EcoFlow",fuente:"Zoho Chat",asesor:comerciales[1]||"Nacho",ciudad:"GBA Norte",dispositivo:"Móvil",pagina:"/ecoflow-delta",createdAt:new Date().toISOString(),notas:[],nota_rapida:"",historial:[]},
      ];
      setLeads(p=>[...nl,...p]);setSyncing(false);
    },1400);
  };

  const sinAsignar=myLeads.filter(l=>l.estado==="sin_asignar").length;
  const nuevosHoy=myLeads.filter(l=>l.fecha===todayStr()&&l.estado==="nuevo").length;
  const rcCount=myLeads.filter(l=>l.estado==="recontactar").length;

  const TABS= isAdmin
    ? [
        {id:"mkt",   label:"DASHBOARD", badge:0,            alert:false, icon:null},
        {id:"equipo",label:"Equipo",    badge:0,            alert:false, icon:<IcoUsers s={12}/>},
      ]
    : [
        {id:"nuevos",      label:"Nuevos",      badge:nuevosHoy,      alert:myLeads.filter(l=>l.estado==="nuevo"&&horasDesde(l.createdAt)>8).length>0},
        {id:"seguimiento", label:"Seguimiento", badge:rcCount,        alert:false},
        {id:"todos",       label:"Todos",        badge:myLeads.length, alert:false},
      ];

  // Sidebar stats mini
  const last7=getLast7();
  const vendidosS=myLeads.filter(l=>l.estado==="vendido").length;
  const convS=myLeads.length?Math.round(vendidosS/myLeads.length*100):0;

  if(!session) return <Login onLogin={login}/>;

  return(
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        body{background:#f8fafc;font-family:'Inter',sans-serif;font-size:14px;color:#f9fafb;}
        ::-webkit-scrollbar{width:4px;height:4px;}::-webkit-scrollbar-track{background:#ffffff;}::-webkit-scrollbar-thumb{background:#e2e8f0;border-radius:2px;}
        select option{background:#ffffff;}
        @keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
        .spin{animation:spin 0.8s linear infinite;display:inline-block;}
      `}</style>

      <div style={{height:"100vh",background:"#f1f5f9",display:"flex",flexDirection:"column",overflow:"hidden"}}>

        {/* NAV */}
        <header style={{background:"#ffffff",borderBottom:"1px solid #e2e8f0",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",padding:"0 16px",display:"flex",alignItems:"center",gap:"12px",height:"50px",flexShrink:0,zIndex:200}}>
          <div style={{display:"flex",alignItems:"center",gap:"8px",minWidth:"fit-content"}}>
            <div style={{width:"36px",height:"36px",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
              <img src="data:image/png;base64,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dnlFuh3Kq4JPBwu39Vd7epo3t6mi4H4733/BO/4C+IPgR8SfhL46+JXiPxHr/wAY77TLnxr4t1BYf7T1S20uSIpYhIbaK2gtDCjW4jiiUrFLJhtxDD334x/st/s0/Gv9q/4T/tdeJdUng1z4SR6jHp9nBGBa3Yvowkfn5iL/AOitukgCMoDsd2RgV+he9vU0b29TRdisfi94O/YP134ZeOvGPjD4VftQeN/Ddv478VXni3WLG10vRJo5ry+dS6CW50uaYRrFHHBGof5IkGPmyx7X4j/saW2rftE+Ov2iPgN8fPFfwsvfidDp0fie00iw029S4fSrf7LbyW8moWFy9s6w8fJkbiWx2r9bN7epo3t6mi4zwj9mf9nz4XfssfAzw98B/g1BND4e0KBhA9zKZ7meS4dp5rieRgC808sjyyNgZdjgAYA91oopAFFFFABRRRQAUUUUAFFFFAH/0f76KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9L++iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/T/voooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD/1P76KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9X++iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/Z" alt="TodoDrones" style={{width:"36px",height:"36px",objectFit:"contain"}}/>
            </div>
            <span style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"14px",fontWeight:700,color:"#0f172a",whiteSpace:"nowrap"}}>COMBOT</span>
          </div>

          <div style={{display:"flex",gap:"1px",flex:1,overflowX:"auto"}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)}
                style={{padding:"5px 10px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"12px",fontWeight:500,display:"flex",alignItems:"center",gap:"4px",whiteSpace:"nowrap",flexShrink:0,position:"relative",
                  background:tab===t.id?"#f0f7ff":"transparent",
                  color:tab===t.id?"#0f172a":"#64748b",
                  borderBottom:tab===t.id?"2px solid #1B87C9":"2px solid transparent"}}>
                {t.icon}{t.label}
                {t.badge>0&&<span style={{background:tab===t.id?"#1B87C9":"#e2e8f0",color:tab===t.id?"#fff":"#64748b",fontSize:"9px",fontWeight:700,padding:"1px 4px",borderRadius:"8px",minWidth:"16px",textAlign:"center"}}>{t.badge}</span>}
                {t.alert&&<span style={{position:"absolute",top:"3px",right:"3px",width:"5px",height:"5px",borderRadius:"50%",background:"#ef4444"}}/>}
              </button>
            ))}
          </div>

          <div style={{display:"flex",gap:"5px",alignItems:"center",flexShrink:0}}>
            {sinAsignar>0&&(
              <button onClick={()=>setTab("seguimiento")}
                style={{display:"flex",alignItems:"center",gap:"4px",padding:"3px 8px",background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.2)",borderRadius:"5px",color:"#f87171",cursor:"pointer",fontSize:"11px",fontWeight:600}}>
                <IcoBell s={11}/>{sinAsignar}
              </button>
            )}
            <SearchBox search={search} setSearch={setSearch} leads={myLeads} setTab={setTab} setNotaL={setNotaL}/>
            {isAdmin&&<button onClick={doSync} disabled={syncing}
              style={{padding:"4px 9px",borderRadius:"5px",border:"1px solid rgba(27,135,201,0.2)",background:"rgba(27,135,201,0.06)",color:"#1B87C9",cursor:syncing?"default":"pointer",fontSize:"11px",fontWeight:600,display:"flex",alignItems:"center",gap:"4px"}}>
              <span className={syncing?"spin":""}><IcoSync s={11}/></span>{syncing?"Sync...":"Zoho"}
            </button>}
            <button onClick={()=>{setEditL(null);setShowForm(true);}}
              style={{padding:"4px 9px",borderRadius:"5px",border:"none",background:"#1B87C9",color:"#fff",cursor:"pointer",fontSize:"11px",fontWeight:600,display:"flex",alignItems:"center",gap:"3px"}}>
              <IcoPlus s={11}/>Nuevo
            </button>
            <button onClick={()=>doExport(myLeads,"leads_"+todayStr())}
              style={{padding:"4px 7px",borderRadius:"5px",border:"1px solid #e2e8f0",background:"#f8fafc",color:"#64748b",cursor:"pointer",display:"flex"}}>
              <IcoDl s={11}/>
            </button>
            <button onClick={()=>setSideOpen(o=>!o)}
              style={{padding:"4px 7px",borderRadius:"5px",border:"1px solid #e2e8f0",background:sideOpen?"#f8fafc":"transparent",color:sideOpen?"#1B87C9":"#64748b",cursor:"pointer",display:"flex"}}>
              <IcoSidebar s={11}/>
            </button>
            <div style={{display:"flex",alignItems:"center",gap:"5px",padding:"3px 8px",background:"#f8fafc",borderRadius:"5px",border:"1px solid #e2e8f0"}}>
              <span style={{fontSize:"11px",color:"#475569"}}>{session.nombre}</span>
              <button onClick={logout} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",padding:0,display:"flex"}}><IcoLogout s={11}/></button>
            </div>
          </div>
        </header>

        {/* BODY */}
        <div style={{display:"flex",flex:1,overflow:"hidden"}}>

          {/* MAIN */}
          <div style={{flex:1,overflow:"auto",padding:"16px"}}>
            {tab==="nuevos"&&<VistaNuevos leads={filtrados} onUpdate={updateLead} onEdit={l=>{setEditL(l);setShowForm(true);}} onNota={l=>setNotaL(l)} isAdmin={isAdmin}
              dateRange={dateRange} setDateRange={setDateRange} customFrom={customFrom} setCustomFrom={setCustomFrom} customTo={customTo} setCustomTo={setCustomTo}/>}
            {tab==="seguimiento"&&<VistaSeguimiento leads={filtrados} onUpdate={updateLead} onEdit={l=>{setEditL(l);setShowForm(true);}} onNota={l=>setNotaL(l)} isAdmin={isAdmin}/>}
            {tab==="todos"&&<VistaTodos leads={filtrados} allLeads={myLeads} onUpdate={updateLead} onEdit={l=>{setEditL(l);setShowForm(true);}} onNota={l=>setNotaL(l)} isAdmin={isAdmin}
              dateRange={dateRange} setDateRange={setDateRange} customFrom={customFrom} setCustomFrom={setCustomFrom} customTo={customTo} setCustomTo={setCustomTo}/>}
            {tab==="mkt"&&isAdmin&&<DashMKT leads={leads} users={users}/>}
            {tab==="equipo"&&isAdmin&&(
              <div style={{padding:"4px"}}>
                <div style={{marginBottom:"14px"}}>
                  <h2 style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"22px",fontWeight:800,color:"#0f172a",marginBottom:"2px"}}>Equipo</h2>
                  <p style={{fontSize:"12px",color:"#64748b"}}>Agregá o desactivá comerciales — los leads e historial se conservan siempre</p>
                </div>
                <GestionEquipo onClose={()=>{}} inline users={users} onUpdate={updateUsers}/>
              </div>
            )}
          </div>

          {/* PANEL NOTAS (slide) */}
          {notaL&&(
            <div style={{width:"340px",flexShrink:0,borderLeft:"1px solid #e2e8f0",overflow:"hidden",height:"100%"}}>
              <PanelNotas lead={notaL} session={session}
                onUpdate={l=>{updateLead(l);setNotaL(l);}}
                onClose={()=>setNotaL(null)}/>
            </div>
          )}

          {/* SIDEBAR STATS */}
          {!notaL&&(
            <div style={{width:sideOpen?"220px":"36px",flexShrink:0,background:"#f8fafc",borderLeft:"1px solid #e2e8f0",transition:"width 0.2s",overflow:"hidden",display:"flex",flexDirection:"column"}}>
              {/* Toggle button */}
              <div style={{padding:"10px 8px",borderBottom:"1px solid #e2e8f0",display:"flex",justifyContent:sideOpen?"flex-end":"center",flexShrink:0}}>
                <button onClick={()=>setSideOpen(o=>!o)}
                  style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8",display:"flex",alignItems:"center",padding:"2px",borderRadius:"4px"}}
                  title={sideOpen?"Cerrar panel":"Abrir panel"}>
                  <IcoSidebar s={13}/>
                </button>
              </div>
              {sideOpen&&<div style={{padding:"14px",overflowY:"auto",display:"flex",flexDirection:"column",gap:"14px",flex:1}}>
              <div>
                <span style={{fontSize:"9px",fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.8px",display:"block",marginBottom:"8px"}}>Resumen</span>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px"}}>
                  {[["Hoy",myLeads.filter(l=>l.fecha===todayStr()).length,"#3b82f6"],["Semana",myLeads.filter(l=>getLast7().includes(l.fecha)).length,"#06b6d4"],["Vendidos",vendidosS,"#10b981"],["Conv.",convS+"%","#10b981"],["Rcont.",rcCount,"#f59e0b"],["S/asig.",sinAsignar,"#f97316"]].map(([lb,v,c])=>(
                    <div key={lb} style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"7px",padding:"8px 10px",boxShadow:"0 1px 2px rgba(0,0,0,0.04)"}}>
                      <div style={{fontSize:"8px",fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:"2px"}}>{lb}</div>
                      <div style={{fontSize:"20px",fontWeight:800,color:c,fontFamily:"'Barlow Condensed',sans-serif"}}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div>
                <div style={{background:"#e2e8f0",borderRadius:"99px",height:"5px",overflow:"hidden",marginBottom:"4px"}}>
                  <div style={{height:"100%",background:"linear-gradient(90deg,#10b981,#34d399)",width:convS+"%",borderRadius:"99px"}}/>
                </div>
                <div style={{fontSize:"10px",color:"#94a3b8"}}>{vendidosS} vendidos de {myLeads.length} leads</div>
              </div>
              {isAdmin&&(
                <div>
                  <span style={{fontSize:"9px",fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.8px",display:"block",marginBottom:"8px"}}>Por comercial</span>
                  {Object.entries(users).filter(([,u])=>u.role==="comercial"&&u.activo!==false).map(([key,u])=>{
                    const al=leads.filter(l=>l.asesor===u.nombre);
                    const av=al.filter(l=>l.estado==="vendido").length;
                    const at=al.length?Math.round(av/al.length*100):0;
                    return(
                      <div key={key} style={{marginBottom:"8px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                          <span style={{fontSize:"11px",color:"#475569"}}>{u.nombre}</span>
                          <span style={{fontSize:"11px",fontWeight:600,color:"#10b981"}}>{av} ventas · {at}%</span>
                        </div>
                        <div style={{background:"#e2e8f0",borderRadius:"99px",height:"3px",overflow:"hidden"}}>
                          <div style={{height:"100%",background:"#10b981",width:at+"%",borderRadius:"99px"}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
              <div>
                <span style={{fontSize:"9px",fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.8px",display:"block",marginBottom:"8px"}}>7 dias</span>
                <ResponsiveContainer width="100%" height={70}>
                  <BarChart data={last7.map(d=>({d:fmtShort(d),n:myLeads.filter(l=>l.fecha===d).length}))} barSize={12} margin={{top:0,right:0,left:-20,bottom:0}}>
                    <XAxis dataKey="d" tick={{fill:"#94a3b8",fontSize:8}} axisLine={false} tickLine={false}/>
                    <Tooltip contentStyle={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"5px",fontSize:"10px",color:"#0f172a"}} cursor={{fill:"#f1f5f9"}}/>
                    <Bar dataKey="n" fill="#1B87C9" radius={[2,2,0,0]} name="Leads"/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              </div>}
            </div>
          )}
        </div>
      </div>

      {showForm&&<FormLead lead={editL} onSave={saveLead} onClose={()=>{setShowForm(false);setEditL(null);}} session={session} users={users}/>}
      {showEquipo&&<GestionEquipo onClose={()=>setShowEquipo(false)}/>}
    </>
  );
}

// Inline gestión de equipo (para tab equipo)
function GestionEquipo({onClose,inline,users:propUsers,onUpdate}){
  const [users,setUsers]=useState(()=>propUsers||db.getUsers());
  const [newU,setNewU]=useState({username:"",nombre:"",pass:"",role:"comercial"});
  const [showPw,setShowPw]=useState({});
  const save=(u)=>{
    db.setUsers(u);setUsers(u);
    if(onUpdate)onUpdate(u);
  };
  const addUser=()=>{
    if(!newU.username||!newU.nombre||!newU.pass)return;
    const updated={...users,[newU.username.toLowerCase()]:{pass:newU.pass,role:newU.role,nombre:newU.nombre,activo:true}};
    save(updated);setNewU({username:"",nombre:"",pass:"",role:"comercial"});
  };
  const toggleActivo=(key)=>save({...users,[key]:{...users[key],activo:!users[key].activo}});
  const inp={width:"100%",padding:"8px 10px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"6px",color:"#0f172a",fontSize:"13px",fontFamily:"inherit",outline:"none",boxSizing:"border-box"};
  const content=(
    <div style={{padding:inline?"0":"20px"}}>
      <div style={{marginBottom:"16px"}}>
        <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px"}}>Equipo actual</div>
        {Object.entries(users).map(([key,u])=>(
          <div key={key} style={{display:"flex",alignItems:"center",gap:"10px",padding:"10px 12px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:"8px",marginBottom:"6px"}}>
            <div style={{flex:1}}>
              <div style={{display:"flex",alignItems:"center",gap:"6px"}}>
                <span style={{fontSize:"13px",fontWeight:600,color:u.activo!==false?"#0f172a":"#94a3b8"}}>{u.nombre}</span>
                <span style={{fontSize:"9px",color:u.role==="admin"?"#1B87C9":"#64748b",background:u.role==="admin"?"rgba(27,135,201,0.12)":"#e2e8f0",padding:"1px 5px",borderRadius:"4px"}}>{u.role}</span>
                {u.activo===false&&<span style={{fontSize:"9px",color:"#ef4444",background:"rgba(239,68,68,0.1)",padding:"1px 5px",borderRadius:"4px"}}>Inactivo</span>}
              </div>
              <div style={{fontSize:"11px",color:"#94a3b8",marginTop:"1px"}}>@{key}</div>
            </div>
            <div style={{position:"relative"}}>
              <input type={showPw[key]?"text":"password"} defaultValue={u.pass} readOnly
                style={{...inp,width:"110px",fontSize:"11px",paddingRight:"26px",color:"#64748b"}}/>
              <button onClick={()=>setShowPw(p=>({...p,[key]:!p[key]}))}
                style={{position:"absolute",right:"6px",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#94a3b8",cursor:"pointer",display:"flex",padding:"1px"}}>
                {showPw[key]?<IcoEyeOff s={11}/>:<IcoEye s={11}/>}
              </button>
            </div>
            {key!=="admin"&&<button onClick={()=>toggleActivo(key)}
              style={{padding:"3px 8px",background:u.activo!==false?"rgba(239,68,68,0.07)":"rgba(16,185,129,0.07)",border:"1px solid "+(u.activo!==false?"rgba(239,68,68,0.2)":"rgba(16,185,129,0.2)"),borderRadius:"5px",color:u.activo!==false?"#ef4444":"#10b981",cursor:"pointer",fontSize:"10px",fontWeight:600,whiteSpace:"nowrap"}}>
              {u.activo!==false?"Desactivar":"Activar"}
            </button>}
          </div>
        ))}
      </div>
      <div style={{borderTop:"1px solid #e2e8f0",paddingTop:"14px"}}>
        <div style={{fontSize:"10px",fontWeight:600,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:"10px"}}>Agregar comercial</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
          <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"3px"}}>Usuario (login)</label><input style={inp} value={newU.username} onChange={e=>setNewU(p=>({...p,username:e.target.value}))} placeholder="pablo"/></div>
          <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"3px"}}>Nombre visible</label><input style={inp} value={newU.nombre} onChange={e=>setNewU(p=>({...p,nombre:e.target.value}))} placeholder="Pablo"/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"10px"}}>
          <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"3px"}}>Contraseña</label><input style={inp} value={newU.pass} onChange={e=>setNewU(p=>({...p,pass:e.target.value}))} placeholder="Contraseña"/></div>
          <div><label style={{fontSize:"10px",color:"#64748b",display:"block",marginBottom:"3px"}}>Rol</label>
            <select style={inp} value={newU.role} onChange={e=>setNewU(p=>({...p,role:e.target.value}))}>
              <option value="comercial">Comercial</option><option value="admin">Admin</option>
            </select>
          </div>
        </div>
        <button onClick={addUser} style={{width:"100%",padding:"9px",background:"#1B87C9",border:"none",borderRadius:"7px",color:"#fff",fontSize:"13px",fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"5px"}}>
          <IcoPlus s={12}/>Agregar al equipo
        </button>
      </div>
    </div>
  );
  if(inline) return content;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",backdropFilter:"blur(4px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:"14px",width:"100%",maxWidth:"500px",maxHeight:"85vh",overflow:"auto"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 20px",borderBottom:"1px solid #e2e8f0"}}>
          <span style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:"17px",fontWeight:700,color:"#0f172a"}}>Gestión de equipo</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer"}}><IcoX s={14}/></button>
        </div>
        {content}
      </div>
    </div>
  );
}
