import { useState, useEffect, useCallback, useMemo } from "react";
import { supabase } from "./supabase";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const ROLES = { admin:"Admin", operator:"Operador", viewer:"Visualizador" };
const CATEGORY_ICONS = { tirzepatida:"💉", material:"🧊", embalagem:"📦", seringa:"🩺", outro:"📋" };
const PAYMENT_METHODS = ["Dinheiro","PIX","Cartão Crédito","Cartão Débito","Transferência"];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt    = v => new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0);
const fmtN   = v => new Intl.NumberFormat("pt-BR").format(v||0);
const fmtPct = v => `${(v||0).toFixed(1)}%`;
const uid    = () => Math.random().toString(36).slice(2,10);
const todayStr = () => new Date().toLocaleDateString("pt-BR");
const nowISO = () => new Date().toISOString();
const hashPass = s => btoa(unescape(encodeURIComponent(s+"|caixapro2026")));
const calcM  = (cost,price) => {
  const c=parseFloat(cost)||0, p=parseFloat(price)||0;
  return {markup:c>0?((p-c)/c)*100:0, margin:p>0?((p-c)/p)*100:0, profit:p-c};
};
const daysUntil = d => { if(!d) return null; return Math.ceil((new Date(d)-new Date())/86400000); };
const expColor  = d => { if(d===null) return "var(--sub)"; if(d<0) return "#f56565"; if(d<=30) return "#f59e0b"; return "#10b981"; };
const stkColor  = q => q<=0?"#f56565":q<=5?"#f59e0b":"#10b981";

// ─── ICONS ────────────────────────────────────────────────────────────────────
const Icon = ({name,size=18}) => {
  const I = {
    dashboard: <path d="M3 3h8v8H3zm10 0h8v8h-8zM3 13h8v8H3zm10 0h8v8h-8z"/>,
    cash:      <><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M12 12m-2 0a2 2 0 1 0 4 0 2 2 0 1 0-4 0"/></>,
    product:   <><path d="M20.91 8.84L8.56 2.23a2 2 0 0 0-2.37.46L2.46 8.84A2 2 0 0 0 2 10v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V10a2 2 0 0 0-.46-1.16zM12 20v-6"/><path d="M2 10h20"/></>,
    client:    <><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></>,
    users:     <><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></>,
    stock:     <><path d="M5 8h14M5 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4zm14 0a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM5 8v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8"/><path d="M10 12h4"/></>,
    sales:     <><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 0 1-8 0"/></>,
    supplier:  <><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></>,
    plus:      <><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></>,
    trend_up:  <><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></>,
    trend_dn:  <><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></>,
    trash:     <><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></>,
    edit:      <><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></>,
    close:     <><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></>,
    sync:      <><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></>,
    warn:      <><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></>,
    search:    <><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></>,
    logout:    <><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></>,
    eye:       <><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></>,
    eye_off:   <><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></>,
    lock:      <><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></>,
    user:      <><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></>,
    shield:    <><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></>,
    check:     <><polyline points="20 6 9 17 4 12"/></>,
    save:      <><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></>,
    arrow_up:  <><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></>,
    arrow_dn:  <><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></>,
    key:       <><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></>,
    chart:     <><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></>,
    syringe:   <><path d="M18 2l4 4m-4-4L7 13m11-11l-4 4M3 21l4.5-4.5"/><path d="M9 15l-3 3"/><path d="M14.5 9.5l-5 5"/></>,
    sun:       <><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></>,
    moon:      <><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></>,
    sun:       <><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></>,
    moon:      <><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></>,
    filter:    <><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></>,
    info:      <><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{I[name]}</svg>;
};

// ─── PRIMITIVES ───────────────────────────────────────────────────────────────
const IS = {width:"100%",background:"var(--bg4)",border:"1px solid var(--border2)",borderRadius:".45rem",padding:".55rem .8rem",color:"var(--text)",fontSize:".84rem",fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box"};

const Modal = ({title,onClose,children,wide,icon}) => (
  <div style={{position:"fixed",inset:0,background:"var(--shadow)",backdropFilter:"blur(8px)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:".75rem"}}>
    <div style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"1.1rem",width:"100%",maxWidth:wide?700:500,maxHeight:"94vh",overflowY:"auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"1rem 1.4rem",borderBottom:"1px solid var(--border)",position:"sticky",top:0,background:"var(--bg3)",zIndex:1}}>
        <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
          {icon&&<span style={{color:"#4f5ef0"}}><Icon name={icon} size={15}/></span>}
          <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".95rem",color:"var(--text)"}}>{title}</span>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",color:"var(--text4)",cursor:"pointer",padding:".2rem"}}><Icon name="close" size={16}/></button>
      </div>
      <div style={{padding:"1.25rem"}}>{children}</div>
    </div>
  </div>
);

const Field = ({label,hint,children}) => (
  <div style={{marginBottom:".85rem"}}>
    <div style={{display:"flex",justifyContent:"space-between",marginBottom:".3rem"}}>
      <label style={{fontSize:".68rem",color:"#777a9a",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"'DM Sans',sans-serif"}}>{label}</label>
      {hint&&<span style={{fontSize:".65rem",color:"var(--text5)"}}>{hint}</span>}
    </div>
    {children}
  </div>
);
const Inp = ({label,hint,...p}) => <Field label={label} hint={hint}><input {...p} style={{...IS,...p.style}}/></Field>;
const Sel = ({label,hint,children,...p}) => <Field label={label} hint={hint}><select {...p} style={{...IS,...p.style}}>{children}</select></Field>;
const Row = ({children,gap=".65rem",cols}) => <div style={{display:"grid",gridTemplateColumns:cols||`repeat(${Array.isArray(children)?children.filter(Boolean).length:1},1fr)`,gap}}>{children}</div>;

const Btn = ({children,onClick,v="primary",sm,disabled,full}) => {
  const vs={primary:{background:"linear-gradient(135deg,#4f5ef0,#8b44f0)",color:"#fff",border:"none"},success:{background:"linear-gradient(135deg,#10b981,#059669)",color:"#fff",border:"none"},danger:{background:"#1e1010",color:"#f56565",border:"1px solid #3a1515"},ghost:{background:"transparent",color:"#777a9a",border:"1px solid var(--border2)"},info:{background:"linear-gradient(135deg,#0891b2,#0e7490)",color:"#fff",border:"none"},warn:{background:"#1e1500",color:"#f59e0b",border:"1px solid #3a2800"}};
  return <button onClick={onClick} disabled={disabled} style={{display:"inline-flex",alignItems:"center",gap:".35rem",padding:sm?".35rem .7rem":".55rem 1rem",borderRadius:".45rem",cursor:disabled?"not-allowed":"pointer",fontSize:sm?".75rem":".84rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,opacity:disabled?.5:1,width:full?"100%":undefined,justifyContent:full?"center":undefined,...vs[v]}}>{children}</button>;
};

const KCard = ({label,value,sub,color="#4f5ef0",icon}) => (
  <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",padding:".9rem 1rem",position:"relative",overflow:"hidden"}}>
    <div style={{position:"absolute",top:"-1rem",right:"-1rem",width:"4.5rem",height:"4.5rem",borderRadius:"50%",background:color,opacity:.07}}/>
    {icon&&<div style={{position:"absolute",top:".85rem",right:".85rem",color,opacity:.5}}><Icon name={icon} size={14}/></div>}
    <div style={{fontSize:".62rem",color:"var(--text4)",textTransform:"uppercase",letterSpacing:".07em",marginBottom:".3rem"}}>{label}</div>
    <div style={{fontSize:"1.15rem",fontWeight:700,color:"var(--text)",fontFamily:"'Syne',sans-serif",letterSpacing:"-.02em"}}>{value}</div>
    {sub&&<div style={{fontSize:".68rem",color,marginTop:".2rem"}}>{sub}</div>}
  </div>
);

const Badge = ({children,color="#4f5ef0",sm}) => (
  <span style={{fontSize:sm?".6rem":".65rem",fontWeight:600,color,background:color+"18",borderRadius:"99px",padding:sm?".1rem .4rem":".15rem .55rem",border:`1px solid ${color}30`,whiteSpace:"nowrap"}}>{children}</span>
);

const Pill = ({label,value,color}) => (
  <div style={{background:"var(--bg)",borderRadius:".4rem",padding:".4rem .6rem"}}>
    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".15rem"}}>{label}</div>
    <div style={{fontSize:".8rem",fontWeight:700,fontFamily:"'Syne',sans-serif",color}}>{value}</div>
  </div>
);

// ─── LOGIN ────────────────────────────────────────────────────────────────────
const LoginScreen = ({onLogin}) => {
  const [user,setUser]=useState(""), [pass,setPass]=useState("");
  const [showPw,setShowPw]=useState(false), [error,setError]=useState("");
  const [shake,setShake]=useState(false), [loading,setLoading]=useState(false);

  const attempt = async () => {
    if(!user.trim()||!pass){setError("Preencha usuário e senha.");return;}
    setLoading(true);
    await new Promise(r=>setTimeout(r,350));
    const {data,error:err}=await supabase.from("app_users").select("*").eq("username",user.trim().toLowerCase()).eq("active",true).single();
    if(err||!data){setError("Usuário não encontrado.");setShake(true);setTimeout(()=>setShake(false),500);setLoading(false);return;}
    if(data.password_hash!==hashPass(pass)){setError("Senha incorreta.");setShake(true);setTimeout(()=>setShake(false),500);setLoading(false);return;}
    await supabase.from("app_users").update({last_login:nowISO()}).eq("id",data.id);
    onLogin(data); setLoading(false);
  };

  const li={...IS,background:"#0d0f1c",border:"1px solid #252848",padding:".7rem 1rem",fontSize:".9rem"};
  return (
    <div style={{minHeight:"100vh",background:"#07080f",display:"flex",alignItems:"center",justifyContent:"center",padding:"1.25rem",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:"-15%",left:"-10%",width:"50vw",height:"50vw",borderRadius:"50%",background:"radial-gradient(circle,var(--glow1),transparent 70%)",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:"-15%",right:"-10%",width:"45vw",height:"45vw",borderRadius:"50%",background:"radial-gradient(circle,var(--glow2),transparent 70%)",pointerEvents:"none"}}/>
      <div style={{width:"100%",maxWidth:380,background:"var(--login-card)",border:"1px solid #252848",borderRadius:"1.25rem",padding:"2rem 1.75rem",boxShadow:"0 30px 80px rgba(0,0,0,.6)",animation:shake?"shake .4s ease":"none"}}>
        <div style={{textAlign:"center",marginBottom:"1.75rem"}}>
          <div style={{width:56,height:56,borderRadius:"1rem",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",marginBottom:".85rem",boxShadow:"0 8px 30px #4f5ef040"}}>
            <Icon name="syringe" size={28}/>
          </div>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.3rem",color:"var(--text)",letterSpacing:"-.02em"}}>CaixaPro</div>
          <div style={{fontSize:".73rem",color:"var(--text5)",marginTop:".25rem"}}>Gestão · Tirzepatida · v5.0</div>
        </div>
        <div style={{marginBottom:".85rem"}}>
          <label style={{display:"block",fontSize:".68rem",color:"var(--text4)",marginBottom:".3rem",textTransform:"uppercase",letterSpacing:".07em"}}>Usuário</label>
          <div style={{position:"relative"}}><div style={{position:"absolute",left:".8rem",top:"50%",transform:"translateY(-50%)",color:"var(--text5)"}}><Icon name="user" size={14}/></div><input value={user} onChange={e=>{setUser(e.target.value);setError("");}} onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="Digite seu usuário" style={{...li,paddingLeft:"2.3rem"}}/></div>
        </div>
        <div style={{marginBottom:"1.25rem"}}>
          <label style={{display:"block",fontSize:".68rem",color:"var(--text4)",marginBottom:".3rem",textTransform:"uppercase",letterSpacing:".07em"}}>Senha</label>
          <div style={{position:"relative"}}>
            <div style={{position:"absolute",left:".8rem",top:"50%",transform:"translateY(-50%)",color:"var(--text5)"}}><Icon name="lock" size={14}/></div>
            <input type={showPw?"text":"password"} value={pass} onChange={e=>{setPass(e.target.value);setError("");}} onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="••••••••••" style={{...li,paddingLeft:"2.3rem",paddingRight:"2.6rem"}}/>
            <button onClick={()=>setShowPw(v=>!v)} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--text5)",padding:0}}><Icon name={showPw?"eye_off":"eye"} size={14}/></button>
          </div>
        </div>
        {error&&<div style={{background:"#f5656512",border:"1px solid #f5656530",borderRadius:".45rem",padding:".55rem .85rem",marginBottom:".85rem",fontSize:".78rem",color:"#f56565",display:"flex",alignItems:"center",gap:".35rem"}}><Icon name="warn" size={13}/>{error}</div>}
        <button onClick={attempt} disabled={loading} style={{width:"100%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",color:"#fff",border:"none",borderRadius:".5rem",padding:".78rem",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",cursor:loading?"not-allowed":"pointer",opacity:loading?.7:1}}>
          {loading?"Verificando...":"Entrar"}
        </button>
        <div style={{textAlign:"center",marginTop:"1.25rem",fontSize:".65rem",color:"var(--text6)"}}>CaixaPro © 2026</div>
      </div>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-6px)}80%{transform:translateX(6px)}}`}</style>
    </div>
  );
};

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function CaixaPro() {
  const [cu,setCU]          = useState(null);
  const [isDark, setIsDark]   = useState(() => {
    try { return localStorage.getItem('cpro:theme') !== 'light'; } catch { return true; }
  });
  const toggleTheme = () => {
    setIsDark(v => {
      const next = !v;
      try { localStorage.setItem('cpro:theme', next ? 'dark' : 'light'); } catch {}
      if (next) { document.body.classList.remove('light'); }
      else       { document.body.classList.add('light'); }
      return next;
    });
  };
  const [tab,setTab]        = useState("dashboard");
  const [products,setProds] = useState([]);
  const [sales,setSales]    = useState([]);
  const [cashTx,setCashTx]  = useState([]);
  const [clients,setClients]= useState([]);
  const [suppliers,setSuppliers] = useState([]);
  const [appUsers,setAppUsers]   = useState([]);
  const [modal,setModal]    = useState(null);
  const [editing,setEditing]= useState(null);
  const [syncing,setSyncing]= useState(false);
  const [lastSync,setLastSync]=useState(null);
  const [logoutC,setLogoutC]= useState(false);
  const [toast,setToast]    = useState(null);
  const [search,setSearch]  = useState("");
  const [pwShow,setPwShow]  = useState({});
  const [filterCat,setFilterCat] = useState("all");

  // Forms
  const EP={code:"",name:"",description:"",category:"tirzepatida",unit:"ampola",cost_per_unit:"",price_per_unit:"",units_per_pack:1,batch:"",expiry:"",stock_qty:"0",supplier_id:"",min_stock:"5"};
  const ES={product_id:"",product_name:"",client_id:"",client_name:"",quantity:"1",unit_price:"",total_price:"",notes:"",payment_method:"PIX"};
  const EC={name:"",email:"",phone:"",notes:""};
  const EU={username:"",display_name:"",role:"operator",password:"",password2:""};
  const ESup={name:"",contact:"",phone:"",email:"",notes:""};
  const ECash={description:"",value:"",type:"saida",category:"",product_id:"",quantity:"1"};

  const [pf,setPf]=useState(EP);
  const [sf,setSf]=useState(ES);
  const [cf,setCf]=useState(EC);
  const [uf,setUf]=useState(EU);
  const [supf,setSupf]=useState(ESup);
  const [cashF,setCashF]=useState(ECash);
  const [stockF,setStockF]=useState({product_id:"",qty:"",cost_total:"",notes:""});

  const toast$ = (msg,color="#10b981")=>{ setToast({msg,color}); setTimeout(()=>setToast(null),3500); };

  // Session
  useEffect(()=>{ try{const s=sessionStorage.getItem("cp5:session");if(s)setCU(JSON.parse(s));}catch{} },[]);
  const login  = u=>{setCU(u);try{sessionStorage.setItem("cp5:session",JSON.stringify(u));}catch{}};
  const logout = ()=>{setCU(null);setLogoutC(false);try{sessionStorage.removeItem("cp5:session");}catch{}};

  // Load
  const load = useCallback(async()=>{
    if(!cu) return;
    setSyncing(true);
    try {
      const [a,b,c,d,e,f]=await Promise.all([
        supabase.from("products").select("*").order("name"),
        supabase.from("sales").select("*").order("created_at",{ascending:false}),
        supabase.from("cash_transactions").select("*").order("created_at",{ascending:false}),
        supabase.from("clients").select("*").order("name"),
        supabase.from("suppliers").select("*").order("name"),
        supabase.from("app_users").select("id,username,display_name,role,active,last_login,created_at").order("created_at",{ascending:false}),
      ]);
      if(a.data)setProds(a.data); if(b.data)setSales(b.data); if(c.data)setCashTx(c.data);
      if(d.data)setClients(d.data); if(e.data)setSuppliers(e.data); if(f.data)setAppUsers(f.data);
      setLastSync(new Date().toLocaleTimeString("pt-BR"));
    }catch(e){console.error(e);}
    setSyncing(false);
  },[cu]);

  useEffect(()=>{
    if(!cu) return;
    load();
    const tabs=["products","sales","cash_transactions","clients","suppliers","app_users"];
    const subs=tabs.map((t,i)=>supabase.channel(`ch${i}`).on("postgres_changes",{event:"*",schema:"public",table:t},load).subscribe());
    return ()=>subs.forEach(s=>s.unsubscribe());
  },[cu,load]);

  // ── METRICS ──
  const cashIn    = useMemo(()=>cashTx.filter(t=>t.type==="entrada").reduce((a,t)=>a+t.value,0),[cashTx]);
  const cashOut   = useMemo(()=>cashTx.filter(t=>t.type==="saida").reduce((a,t)=>a+t.value,0),[cashTx]);
  const netProfit = cashIn-cashOut;
  const avgMarkup = useMemo(()=>{
    const saleItems=cashTx.filter(t=>t.type==="entrada"&&t.sale_id);
    if(!saleItems.length||!cashIn) return 0;
    return (netProfit/cashOut)*100;
  },[cashTx,cashIn,cashOut,netProfit]);
  const totalSalesRev = useMemo(()=>sales.reduce((a,s)=>a+s.total_price,0),[sales]);
  const totalUnitsSold= useMemo(()=>sales.reduce((a,s)=>a+s.quantity,0),[sales]);
  const lowStock  = useMemo(()=>products.filter(p=>p.stock_qty>0&&p.stock_qty<=(p.min_stock||5)),[products]);
  const zeroStock = useMemo(()=>products.filter(p=>p.stock_qty<=0),[products]);
  const totalStockValue = useMemo(()=>products.reduce((a,p)=>a+p.stock_qty*(p.cost_per_unit||0),0),[products]);

  const isAdmin  = cu?.role==="admin";
  const canEdit  = cu?.role==="admin"||cu?.role==="operator";

  // ── SALE ──
  const handleSaleProd = (id)=>{
    const p=products.find(x=>x.id===id);
    if(!p){setSf(f=>({...f,product_id:"",product_name:"",unit_price:"",total_price:""}));return;}
    const up=p.price_per_unit||0;
    setSf(f=>({...f,product_id:id,product_name:p.name,unit_price:String(up),total_price:String(up*(parseInt(f.quantity)||1))}));
  };
  const handleSaleQty=(q)=>{
    const qty=Math.max(1,parseInt(q)||1);
    setSf(f=>({...f,quantity:String(qty),total_price:String((parseFloat(f.unit_price)||0)*qty)}));
  };
  const handleSalePrice=(p)=>{
    setSf(f=>({...f,unit_price:p,total_price:String((parseFloat(p)||0)*(parseInt(f.quantity)||1))}));
  };

  const registerSale=async()=>{
    if(!sf.product_id||!sf.quantity||!sf.unit_price){toast$("Preencha produto, quantidade e preço.","#f56565");return;}
    const prod=products.find(p=>p.id===sf.product_id);
    if(!prod){toast$("Produto não encontrado.","#f56565");return;}
    const qty=parseInt(sf.quantity)||1;
    const up=parseFloat(sf.unit_price)||0;
    const total=up*qty;
    if(prod.stock_qty<qty){toast$(`Estoque insuficiente! Disponível: ${prod.stock_qty} ${prod.unit||"un"}.`,"#f56565");return;}
    const sid=uid();
    const {error:sErr}=await supabase.from("sales").insert([{id:sid,product_id:sf.product_id,product_name:sf.product_name,client_id:sf.client_id||null,client_name:sf.client_name||null,quantity:qty,unit_price:up,total_price:total,notes:sf.notes,payment_method:sf.payment_method,added_by:cu.display_name,date:todayStr()}]);
    if(sErr){toast$("Erro ao registrar venda.","#f56565");return;}
    await supabase.from("products").update({stock_qty:prod.stock_qty-qty}).eq("id",sf.product_id);
    await supabase.from("cash_transactions").insert([{id:uid(),description:`Venda · ${sf.product_name}${sf.client_name?` · ${sf.client_name}`:""}`,value:total,type:"entrada",category:"Venda",sale_id:sid,product_name:sf.product_name,added_by:cu.display_name,date:todayStr()}]);
    toast$(`✅ Venda de ${qty} ${prod.unit||"un"} registrada!`);
    setModal(null); setSf(ES);
  };

  // ── STOCK ENTRY ──
  const registerStockEntry=async()=>{
    if(!stockF.product_id||!stockF.qty){toast$("Selecione produto e quantidade.","#f56565");return;}
    const prod=products.find(p=>p.id===stockF.product_id);
    if(!prod)return;
    const qty=parseInt(stockF.qty)||0;
    if(qty<=0){toast$("Quantidade inválida.","#f56565");return;}
    const cost=parseFloat(stockF.cost_total)||0;
    await supabase.from("products").update({stock_qty:prod.stock_qty+qty}).eq("id",stockF.product_id);
    if(cost>0){
      await supabase.from("cash_transactions").insert([{id:uid(),description:`Entrada estoque · ${prod.name} · ${qty} ${prod.unit||"un"}`,value:cost,type:"saida",category:"Compra/Estoque",product_name:prod.name,added_by:cu.display_name,date:todayStr()}]);
    }
    toast$(`✅ +${qty} ${prod.unit||"un"} adicionadas ao estoque de ${prod.name}!`);
    setModal(null); setStockF({product_id:"",qty:"",cost_total:"",notes:""});
  };

  // ── CASH TX MANUAL ──
  const addCashTx=async()=>{
    if(!cashF.description||!cashF.value){toast$("Preencha descrição e valor.","#f56565");return;}
    // Se tem produto vinculado e é saída, baixa estoque
    if(cashF.product_id&&cashF.type==="saida"){
      const prod=products.find(p=>p.id===cashF.product_id);
      const qty=parseInt(cashF.quantity)||0;
      if(prod&&qty>0){
        if(prod.stock_qty<qty){toast$(`Estoque insuficiente! ${prod.stock_qty} ${prod.unit||"un"} disponíveis.`,"#f56565");return;}
        await supabase.from("products").update({stock_qty:prod.stock_qty-qty}).eq("id",cashF.product_id);
      }
    }
    await supabase.from("cash_transactions").insert([{id:uid(),description:cashF.description,value:parseFloat(cashF.value),type:cashF.type,category:cashF.category,product_name:cashF.product_id?products.find(p=>p.id===cashF.product_id)?.name:null,added_by:cu.display_name,date:todayStr()}]);
    toast$("Lançamento adicionado!"); setModal(null); setCashF(ECash);
  };

  // ── PRODUCT CRUD ──
  const addProduct=async()=>{
    if(!pf.name||!pf.cost_per_unit||!pf.price_per_unit){toast$("Preencha nome, custo e preço.","#f56565");return;}
    const {markup,margin,profit}=calcM(pf.cost_per_unit,pf.price_per_unit);
    const sup=suppliers.find(s=>s.id===pf.supplier_id);
    await supabase.from("products").insert([{id:uid(),code:pf.code||`PRD-${uid().slice(0,4).toUpperCase()}`,name:pf.name,description:pf.description,category:pf.category,unit:pf.unit||"un",cost_per_unit:parseFloat(pf.cost_per_unit),price_per_unit:parseFloat(pf.price_per_unit),units_per_pack:parseInt(pf.units_per_pack)||1,batch:pf.batch,expiry:pf.expiry||null,stock_qty:parseInt(pf.stock_qty)||0,min_stock:parseInt(pf.min_stock)||5,supplier_id:pf.supplier_id||null,supplier_name:sup?.name||null,markup,margin,profit,added_by:cu.display_name}]);
    toast$("Produto cadastrado!"); setModal(null); setPf(EP);
  };
  const saveProduct=async()=>{
    const {markup,margin,profit}=calcM(editing.cost_per_unit,editing.price_per_unit);
    const sup=suppliers.find(s=>s.id===editing.supplier_id);
    await supabase.from("products").update({code:editing.code,name:editing.name,description:editing.description,category:editing.category,unit:editing.unit,cost_per_unit:parseFloat(editing.cost_per_unit),price_per_unit:parseFloat(editing.price_per_unit),units_per_pack:parseInt(editing.units_per_pack)||1,batch:editing.batch,expiry:editing.expiry||null,stock_qty:parseInt(editing.stock_qty),min_stock:parseInt(editing.min_stock)||5,supplier_id:editing.supplier_id||null,supplier_name:sup?.name||null,markup,margin,profit}).eq("id",editing.id);
    toast$("Produto atualizado!"); setModal(null); setEditing(null);
  };
  const delProduct=async id=>{await supabase.from("products").delete().eq("id",id);toast$("Removido.","#f59e0b");setModal(null);setEditing(null);};

  // ── SUPPLIER CRUD ──
  const addSupplier=async()=>{
    if(!supf.name){toast$("Nome é obrigatório.","#f56565");return;}
    await supabase.from("suppliers").insert([{id:uid(),...supf,added_by:cu.display_name}]);
    toast$("Fornecedor cadastrado!"); setModal(null); setSupf(ESup);
  };
  const saveSupplier=async()=>{
    await supabase.from("suppliers").update({name:editing.name,contact:editing.contact,phone:editing.phone,email:editing.email,notes:editing.notes}).eq("id",editing.id);
    toast$("Fornecedor atualizado!"); setModal(null); setEditing(null);
  };

  // ── CLIENT CRUD ──
  const addClient=async()=>{
    if(!cf.name)return;
    await supabase.from("clients").insert([{id:uid(),...cf,added_by:cu.display_name}]);
    toast$("Cliente cadastrado!"); setModal(null); setCf(EC);
  };
  const saveClient=async()=>{
    await supabase.from("clients").update({name:editing.name,email:editing.email,phone:editing.phone,notes:editing.notes}).eq("id",editing.id);
    toast$("Cliente atualizado!"); setModal(null); setEditing(null);
  };

  // ── USER CRUD ──
  const addUser=async()=>{
    if(!uf.username||!uf.display_name||!uf.password)return;
    if(uf.password!==uf.password2){toast$("Senhas não coincidem.","#f56565");return;}
    if(uf.password.length<6){toast$("Mínimo 6 caracteres.","#f59e0b");return;}
    const {error}=await supabase.from("app_users").insert([{id:uid(),username:uf.username.trim().toLowerCase(),display_name:uf.display_name,role:uf.role,password_hash:hashPass(uf.password),active:true}]);
    if(!error){toast$("Usuário criado!"); setModal(null); setUf(EU);}else toast$("Erro. Usuário já existe?","#f56565");
  };
  const saveUser=async()=>{
    const updates={display_name:editing.display_name,role:editing.role,active:editing.active};
    if(editing.new_password){
      if(editing.new_password.length<6){toast$("Mínimo 6 caracteres.","#f59e0b");return;}
      if(editing.new_password!==editing.new_password2){toast$("Senhas não coincidem.","#f56565");return;}
      updates.password_hash=hashPass(editing.new_password);
    }
    await supabase.from("app_users").update(updates).eq("id",editing.id);
    toast$("Usuário atualizado!"); setModal(null); setEditing(null);
  };

  // ── FILTERS ──
  const cats=["all",...new Set(products.map(p=>p.category).filter(Boolean))];
  const filteredProds=products.filter(p=>(filterCat==="all"||p.category===filterCat)&&(p.name?.toLowerCase().includes(search.toLowerCase())||p.code?.toLowerCase().includes(search.toLowerCase())));

  const nav=[
    {id:"dashboard",label:"Dashboard",icon:"dashboard"},
    {id:"vendas",   label:"Vendas",   icon:"sales"},
    {id:"estoque",  label:"Estoque",  icon:"stock"},
    {id:"caixa",    label:"Caixa",    icon:"cash"},
    {id:"clientes", label:"Clientes", icon:"client"},
    {id:"produtos", label:"Produtos", icon:"product"},
    ...(isAdmin?[{id:"usuarios",label:"Usuários",icon:"users"}]:[]),
  ];

  if(!cu) return (<><style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');*{margin:0;padding:0;box-sizing:border-box;}input:focus{border-color:#4f5ef0!important;}`}</style><LoginScreen onLogin={login}/></>);

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');
        @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        @keyframes shake{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-6px)}80%{transform:translateX(6px)}}
        input[type=date]::-webkit-calendar-picker-indicator{filter:${isDark?"invert(.5)":"none"};}
      `}</style>

      {toast&&<div style={{position:"fixed",bottom:"1.25rem",right:"1.25rem",zIndex:999,background:"var(--bg3)",border:`1px solid ${toast.color}50`,borderRadius:".65rem",padding:".65rem 1.1rem",display:"flex",alignItems:"center",gap:".45rem",fontSize:".82rem",color:toast.color,fontFamily:"'DM Sans',sans-serif",boxShadow:"0 8px 30px rgba(0,0,0,.5)",animation:"fadeUp .3s ease",maxWidth:320}}><Icon name="check" size={14}/>{toast.msg}</div>}

      <div style={{minHeight:"100vh",background:"var(--bg)",fontFamily:"'DM Sans',sans-serif",color:"var(--text)"}}>

        {/* TOP BAR */}
        <div style={{background:"var(--bg2)",borderBottom:"1px solid #181a28",padding:".65rem 1.1rem",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:50}}>
          <div style={{display:"flex",alignItems:"center",gap:".45rem"}}>
            <div style={{width:26,height:26,borderRadius:".4rem",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",color:"#fff"}}><Icon name="syringe" size={15}/></div>
            <div>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".88rem",lineHeight:1}}>CaixaPro</div>
              <div style={{fontSize:".55rem",color:"var(--text5)",lineHeight:1.2}}>Tirzepatida v5.0</div>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:".75rem"}}>
            {(zeroStock.length>0||lowStock.length>0)&&(
              <div style={{display:"flex",alignItems:"center",gap:".25rem",fontSize:".65rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".2rem .55rem",border:"1px solid #f59e0b30"}}>
                <Icon name="warn" size={10}/>{zeroStock.length>0?`${zeroStock.length} zerado(s)`:`${lowStock.length} baixo`}
              </div>
            )}
            <div style={{display:"flex",alignItems:"center",gap:".3rem",fontSize:".65rem",color:syncing?"#8b44f0":"var(--text6)"}}><Icon name="sync" size={10}/>{syncing?"...":lastSync||"—"}</div>
            <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
              <div style={{width:24,height:24,borderRadius:"50%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:".68rem",fontWeight:700,color:"#fff",fontFamily:"'Syne',sans-serif"}}>{cu.display_name.charAt(0).toUpperCase()}</div>
              <button
                onClick={toggleTheme}
                title={isDark?"☀️ Modo claro":"🌙 Modo escuro"}
                style={{background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",padding:".2rem .3rem",borderRadius:".35rem",color:isDark?"#f59e0b":"#4f5ef0",transition:"color .2s"}}
              ><Icon name={isDark?"sun":"moon"} size={16}/></button>
              <button onClick={()=>setLogoutC(true)} style={{background:"none",border:"none",cursor:"pointer",color:"var(--text5)",display:"flex"}}><Icon name="logout" size={13}/></button>
            </div>
          </div>
        </div>

        {/* NAV */}
        <div style={{background:"var(--bg2)",borderBottom:"1px solid #181a28",padding:"0 .5rem",display:"flex",overflowX:"auto"}}>
          {nav.map(n=>(
            <button key={n.id} onClick={()=>{setTab(n.id);setSearch("");setFilterCat("all");}} style={{display:"flex",alignItems:"center",gap:".28rem",padding:".6rem .7rem",background:"none",border:"none",cursor:"pointer",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,color:tab===n.id?"#4f5ef0":"var(--text4)",borderBottom:tab===n.id?"2px solid #4f5ef0":"2px solid transparent",transition:"all .2s",whiteSpace:"nowrap"}}>
              <Icon name={n.icon} size={12}/>{n.label}
              {n.id==="estoque"&&zeroStock.length>0&&<span style={{background:"#f56565",color:"#fff",borderRadius:"99px",fontSize:".55rem",fontWeight:700,padding:"0 .28rem",lineHeight:"1.5"}}>{zeroStock.length}</span>}
            </button>
          ))}
        </div>

        <div style={{maxWidth:960,margin:"0 auto",padding:"1rem .85rem"}}>

          {/* ════ DASHBOARD ════ */}
          {tab==="dashboard"&&(
            <>
              {/* KPIs */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".65rem",marginBottom:"1rem"}}>
                <KCard label="Receita" value={fmt(cashIn)} color="#10b981" icon="trend_up"/>
                <KCard label="Custos" value={fmt(cashOut)} color="#f56565" icon="trend_dn"/>
                <KCard label="Lucro" value={fmt(netProfit)} sub={cashIn>0?fmtPct((netProfit/cashIn)*100)+" margem":""} color={netProfit>=0?"#4f5ef0":"#f56565"}/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".65rem",marginBottom:"1rem"}}>
                <KCard label="Vendas" value={fmtN(sales.length)} sub={`${fmtN(totalUnitsSold)} un vendidas`} color="#8b44f0" icon="sales"/>
                <KCard label="Estoque" value={products.length} sub={zeroStock.length>0?`${zeroStock.length} zerado(s)`:lowStock.length>0?`${lowStock.length} baixo(s)`:"ok"} color={zeroStock.length>0?"#f56565":lowStock.length>0?"#f59e0b":"#10b981"} icon="stock"/>
                <KCard label="Clientes" value={clients.length} color="#10b981" icon="client"/>
              </div>

              {/* Alertas */}
              {(zeroStock.length>0||lowStock.length>0)&&(
                <div style={{background:"var(--bg3)",border:"1px solid #f59e0b30",borderRadius:".75rem",padding:".85rem 1rem",marginBottom:"1rem"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".35rem",color:"#f59e0b",fontWeight:700,fontSize:".8rem",marginBottom:".4rem"}}><Icon name="warn" size={14}/> Alertas de Estoque</div>
                  {zeroStock.length>0&&<p style={{fontSize:".75rem",color:"#f56565",marginBottom:".2rem"}}>🔴 Zerados: {zeroStock.map(p=>p.name).join(", ")}</p>}
                  {lowStock.length>0&&<p style={{fontSize:".75rem",color:"#f59e0b"}}>🟡 Baixo: {lowStock.map(p=>`${p.name}(${p.stock_qty})`).join(", ")}</p>}
                </div>
              )}

              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".85rem"}}>
                {/* Últimas vendas */}
                <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",padding:"1rem"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem"}}>
                    <div style={{display:"flex",alignItems:"center",gap:".35rem",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--text2)"}}><Icon name="sales" size={13}/> Últimas Vendas</div>
                    {canEdit&&<Btn sm v="primary" onClick={()=>setModal("sale")}><Icon name="plus" size={11}/> Nova</Btn>}
                  </div>
                  {sales.length===0?<p style={{color:"var(--text6)",fontSize:".78rem",textAlign:"center",padding:"1.25rem 0"}}>Nenhuma venda.</p>
                  :sales.slice(0,5).map(s=>(
                    <div key={s.id} style={{padding:".45rem 0",borderBottom:"1px solid var(--sep)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div style={{minWidth:0}}>
                        <div style={{fontSize:".78rem",color:"var(--text2)",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.product_name}</div>
                        <div style={{fontSize:".65rem",color:"var(--text5)"}}>{s.date} · {s.quantity}un{s.client_name&&` · ${s.client_name}`}</div>
                      </div>
                      <span style={{color:"#10b981",fontWeight:700,fontFamily:"'Syne',sans-serif",fontSize:".8rem",flexShrink:0,marginLeft:".5rem"}}>{fmt(s.total_price)}</span>
                    </div>
                  ))}
                </div>

                {/* Estoque resumo */}
                <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",padding:"1rem"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem"}}>
                    <div style={{display:"flex",alignItems:"center",gap:".35rem",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--text2)"}}><Icon name="stock" size={13}/> Estoque</div>
                    {canEdit&&<Btn sm v="info" onClick={()=>setModal("stockEntry")}><Icon name="arrow_up" size={11}/> Entrada</Btn>}
                  </div>
                  {products.length===0?<p style={{color:"var(--text6)",fontSize:".78rem",textAlign:"center",padding:"1.25rem 0"}}>Nenhum produto.</p>
                  :products.slice(0,6).map(p=>(
                    <div key={p.id} style={{padding:".45rem 0",borderBottom:"1px solid var(--sep)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div style={{minWidth:0}}>
                        <div style={{fontSize:".78rem",color:"var(--text2)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{CATEGORY_ICONS[p.category]||"📋"} {p.name}</div>
                        <div style={{fontSize:".65rem",color:"var(--text5)"}}>{p.code}</div>
                      </div>
                      <span style={{color:stkColor(p.stock_qty),fontWeight:700,fontFamily:"'Syne',sans-serif",fontSize:".8rem",flexShrink:0,marginLeft:".5rem"}}>{p.stock_qty} {p.unit||"un"}</span>
                    </div>
                  ))}
                  <div style={{marginTop:".65rem",paddingTop:".65rem",borderTop:"1px solid var(--border)",display:"flex",justifyContent:"space-between",fontSize:".72rem"}}>
                    <span style={{color:"var(--text5)"}}>Valor total estoque</span>
                    <span style={{color:"#f59e0b",fontWeight:700,fontFamily:"'Syne',sans-serif"}}>{fmt(totalStockValue)}</span>
                  </div>
                </div>
              </div>
            </>
          )}

          {/* ════ VENDAS ════ */}
          {tab==="vendas"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Vendas</h2>
                {canEdit&&<Btn onClick={()=>setModal("sale")}><Icon name="plus" size={13}/> Nova venda</Btn>}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:".6rem",marginBottom:".85rem"}}>
                <KCard label="Total" value={fmtN(sales.length)} color="#4f5ef0"/>
                <KCard label="Receita" value={fmt(totalSalesRev)} color="#10b981"/>
                <KCard label="Unidades" value={fmtN(totalUnitsSold)} color="#8b44f0"/>
                <KCard label="Ticket médio" value={sales.length>0?fmt(totalSalesRev/sales.length):"—"} color="#f59e0b"/>
              </div>
              <div style={{position:"relative",marginBottom:".75rem"}}>
                <div style={{position:"absolute",left:".75rem",top:"50%",transform:"translateY(-50%)",color:"var(--text5)"}}><Icon name="search" size={14}/></div>
                <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por produto ou cliente..." style={{...IS,paddingLeft:"2.1rem",width:"100%",background:"var(--bg3)"}}/>
              </div>
              <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem"}}>
                {sales.filter(s=>s.product_name?.toLowerCase().includes(search.toLowerCase())||s.client_name?.toLowerCase().includes(search.toLowerCase())).length===0
                  ?<p style={{color:"var(--text5)",textAlign:"center",padding:"2.5rem 0",fontSize:".8rem"}}>Nenhuma venda.</p>
                  :sales.filter(s=>s.product_name?.toLowerCase().includes(search.toLowerCase())||s.client_name?.toLowerCase().includes(search.toLowerCase())).map(s=>(
                  <div key={s.id} style={{padding:".75rem 1rem",borderBottom:"1px solid var(--sep)",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:".4rem",flexWrap:"wrap",marginBottom:".2rem"}}>
                        <span style={{fontWeight:700,fontSize:".85rem",color:"var(--text)"}}>{s.product_name}</span>
                        <Badge color="#8b44f0" sm>{s.quantity} un</Badge>
                        <Badge color="#0891b2" sm>{s.payment_method}</Badge>
                      </div>
                      <div style={{fontSize:".68rem",color:"var(--text5)"}}>{s.date}{s.client_name&&` · 👤 ${s.client_name}`}{s.notes&&` · ${s.notes}`}</div>
                    </div>
                    <div style={{flexShrink:0,marginLeft:"1rem",textAlign:"right"}}>
                      <div style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:"#10b981",fontSize:".9rem"}}>{fmt(s.total_price)}</div>
                      <div style={{fontSize:".65rem",color:"var(--text5)"}}>{fmt(s.unit_price)}/un</div>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ════ ESTOQUE ════ */}
          {tab==="estoque"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Estoque</h2>
                {canEdit&&<div style={{display:"flex",gap:".5rem"}}>
                  <Btn v="info" sm onClick={()=>setModal("stockEntry")}><Icon name="arrow_up" size={13}/> Entrada</Btn>
                  <Btn sm onClick={()=>setModal("produto")}><Icon name="plus" size={13}/> Produto</Btn>
                </div>}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:".6rem",marginBottom:".85rem"}}>
                <KCard label="Itens" value={products.length} color="#4f5ef0"/>
                <KCard label="Unidades" value={fmtN(products.reduce((a,p)=>a+p.stock_qty,0))} color="#8b44f0"/>
                <KCard label="Valor estoque" value={fmt(totalStockValue)} color="#f59e0b"/>
                {zeroStock.length>0&&<KCard label="Zerados" value={zeroStock.length} color="#f56565"/>}
              </div>
              {/* Filtro por categoria */}
              <div style={{display:"flex",gap:".4rem",marginBottom:".75rem",overflowX:"auto",paddingBottom:".25rem"}}>
                {cats.map(c=>(
                  <button key={c} onClick={()=>setFilterCat(c)} style={{padding:".3rem .7rem",borderRadius:"99px",border:`1px solid ${filterCat===c?"#4f5ef0":"var(--border2)"}`,background:filterCat===c?"#4f5ef018":"transparent",color:filterCat===c?"#4f5ef0":"var(--text4)",cursor:"pointer",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,whiteSpace:"nowrap"}}>
                    {c==="all"?"Todos":`${CATEGORY_ICONS[c]||"📋"} ${c}`}
                  </button>
                ))}
              </div>
              <div style={{display:"grid",gap:".65rem"}}>
                {filteredProds.map(p=>{
                  const days=p.expiry?daysUntil(p.expiry):null;
                  return (
                    <div key={p.id} style={{background:"var(--bg3)",border:`1px solid ${p.stock_qty<=0?"#f5656525":p.stock_qty<=(p.min_stock||5)?"#f59e0b25":"var(--border)"}`,borderRadius:".75rem",padding:".85rem 1rem"}}>
                      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".6rem"}}>
                        <div style={{flex:1}}>
                          <div style={{display:"flex",alignItems:"center",gap:".4rem",flexWrap:"wrap",marginBottom:".25rem"}}>
                            <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".88rem",color:"var(--text)"}}>{CATEGORY_ICONS[p.category]||"📋"} {p.name}</span>
                            <Badge color={stkColor(p.stock_qty)} sm>{p.stock_qty<=0?"Zerado":p.stock_qty<=(p.min_stock||5)?"Baixo":"OK"}</Badge>
                            {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                          </div>
                          <div style={{fontSize:".67rem",color:"var(--text5)"}}>
                            {p.code}{p.batch&&` · Lote: ${p.batch}`}
                            {p.expiry&&<span style={{color:expColor(days),marginLeft:".35rem"}}>· Vcto {new Date(p.expiry+"T00:00:00").toLocaleDateString("pt-BR")}{days!==null&&days<=30&&<span> ⚠️</span>}</span>}
                          </div>
                        </div>
                        <div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".75rem"}}>
                          {canEdit&&<button onClick={()=>{setStockF({product_id:p.id,qty:"",cost_total:"",notes:""});setModal("stockEntry");}} style={{background:"var(--info-box)",border:"1px solid #10b98130",borderRadius:".4rem",padding:".3rem .55rem",cursor:"pointer",color:"#10b981",display:"flex",alignItems:"center",gap:".2rem",fontSize:".68rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}><Icon name="arrow_up" size={11}/></button>}
                          {(isAdmin||canEdit)&&<button onClick={()=>{setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5)});setModal("editProd");}} style={{background:"none",border:"none",cursor:"pointer",color:"#4f5ef0"}}><Icon name="edit" size={14}/></button>}
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(72px,1fr))",gap:".4rem"}}>
                        {[{l:"Estoque",v:`${p.stock_qty} ${p.unit||"un"}`,c:stkColor(p.stock_qty)},{l:"Mínimo",v:`${p.min_stock||5} ${p.unit||"un"}`,c:"var(--sub)"},{l:"Custo",v:fmt(p.cost_per_unit),c:"var(--text3)"},{l:"Preço",v:fmt(p.price_per_unit),c:"#10b981"},{l:"Markup",v:fmtPct(p.markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(p.margin),c:"#8b44f0"}].map(m=>(
                          <Pill key={m.l} label={m.l} value={m.v} color={m.c}/>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {/* ════ CAIXA ════ */}
          {tab==="caixa"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Caixa</h2>
                {canEdit&&<Btn sm onClick={()=>setModal("cashTx")}><Icon name="plus" size={13}/> Lançamento</Btn>}
              </div>
              {/* Resumo financeiro completo */}
              <div style={{background:"var(--analytics-card)",border:"1px solid var(--analytics-border)",borderRadius:".75rem",padding:"1rem",marginBottom:".85rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"#8b44f0",marginBottom:".75rem",display:"flex",alignItems:"center",gap:".35rem"}}><Icon name="chart" size={14}/> Resumo Financeiro</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem",marginBottom:".65rem"}}>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".6rem"}}>
                    <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".2rem"}}>Receita (100%)</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(cashIn)}</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".6rem"}}>
                    <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".2rem"}}>Custos</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:"#f56565",fontFamily:"'Syne',sans-serif"}}>{fmt(cashOut)}</div>
                    <div style={{fontSize:".62rem",color:"#f56565"}}>{cashIn>0?fmtPct((cashOut/cashIn)*100):"0%"} da receita</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".6rem"}}>
                    <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".2rem"}}>Lucro líquido</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:netProfit>=0?"#4f5ef0":"#f56565",fontFamily:"'Syne',sans-serif"}}>{fmt(netProfit)}</div>
                    <div style={{fontSize:".62rem",color:netProfit>=0?"#4f5ef0":"#f56565"}}>{cashIn>0?fmtPct((netProfit/cashIn)*100):"0%"} margem</div>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem"}}>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".5rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>Markup médio</div>
                    <div style={{fontSize:".88rem",fontWeight:700,color:"#f59e0b",fontFamily:"'Syne',sans-serif"}}>{cashOut>0?fmtPct(avgMarkup):"—"}</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".5rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>Receita vendas</div>
                    <div style={{fontSize:".88rem",fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(totalSalesRev)}</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--info-box)",borderRadius:".5rem",padding:".5rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>Valor estoque</div>
                    <div style={{fontSize:".88rem",fontWeight:700,color:"#8b44f0",fontFamily:"'Syne',sans-serif"}}>{fmt(totalStockValue)}</div>
                  </div>
                </div>
              </div>
              <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem"}}>
                {cashTx.length===0?<p style={{color:"var(--text5)",fontSize:".8rem",textAlign:"center",padding:"2.5rem 0"}}>Nenhum lançamento.</p>
                :cashTx.map(t=>(
                  <div key={t.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".75rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:".84rem",color:"var(--text2)",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.description}</div>
                      <div style={{fontSize:".67rem",color:"var(--text5)",marginTop:".1rem"}}>{t.date}{t.category&&` · ${t.category}`}{t.added_by&&` · ${t.added_by}`}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:".6rem",flexShrink:0,marginLeft:".85rem"}}>
                      <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:t.type==="entrada"?"#10b981":"#f56565",fontSize:".88rem"}}>{t.type==="entrada"?"+":"-"}{fmt(t.value)}</span>
                      {canEdit&&!t.sale_id&&<button onClick={async()=>{await supabase.from("cash_transactions").delete().eq("id",t.id);toast$("Removido.","#f59e0b");}} style={{background:"none",border:"none",cursor:"pointer",color:"var(--text6)"}}><Icon name="trash" size={13}/></button>}
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ════ CLIENTES ════ */}
          {tab==="clientes"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Clientes</h2>
                {canEdit&&<Btn sm onClick={()=>setModal("cliente")}><Icon name="plus" size={13}/> Novo</Btn>}
              </div>
              {clients.length===0?<p style={{color:"var(--text5)",textAlign:"center",padding:"3rem 0",fontSize:".8rem"}}>Nenhum cliente.</p>
              :<div style={{display:"grid",gap:".6rem"}}>
                {clients.map(c=>{
                  const cs=sales.filter(s=>s.client_id===c.id||s.client_name===c.name);
                  return (
                    <div key={c.id} style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",padding:".8rem 1rem",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div style={{display:"flex",alignItems:"center",gap:".75rem"}}>
                        <div style={{width:34,height:34,borderRadius:"50%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".85rem",color:"#fff",flexShrink:0}}>{c.name.charAt(0).toUpperCase()}</div>
                        <div>
                          <div style={{fontWeight:700,fontSize:".86rem",color:"var(--text)"}}>{c.name}</div>
                          <div style={{fontSize:".68rem",color:"var(--text5)"}}>{[c.phone,c.email].filter(Boolean).join(" · ")||"Sem contato"}</div>
                          {c.notes&&<div style={{fontSize:".67rem",color:"#8b44f0"}}>📝 {c.notes}</div>}
                        </div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:".65rem"}}>
                        {cs.length>0&&<div style={{textAlign:"right"}}>
                          <div style={{fontWeight:700,color:"#10b981",fontSize:".82rem",fontFamily:"'Syne',sans-serif"}}>{fmt(cs.reduce((a,s)=>a+s.total_price,0))}</div>
                          <div style={{fontSize:".62rem",color:"var(--text5)"}}>{cs.length} compra{cs.length!==1?"s":""}</div>
                        </div>}
                        {canEdit&&<div style={{display:"flex",gap:".35rem"}}>
                          <button onClick={()=>{setEditing({...c});setModal("editCliente");}} style={{background:"none",border:"none",cursor:"pointer",color:"#4f5ef0"}}><Icon name="edit" size={13}/></button>
                          <button onClick={async()=>{await supabase.from("clients").delete().eq("id",c.id);toast$("Removido.","#f59e0b");}} style={{background:"none",border:"none",cursor:"pointer",color:"var(--text6)"}}><Icon name="trash" size={13}/></button>
                        </div>}
                      </div>
                    </div>
                  );
                })}
              </div>}
            </>
          )}

          {/* ════ PRODUTOS ════ */}
          {tab==="produtos"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Produtos</h2>
                <div style={{display:"flex",gap:".4rem"}}>
                  <Btn sm v="ghost" onClick={()=>setModal("supplier")}><Icon name="supplier" size={12}/> Fornecedor</Btn>
                  {canEdit&&<Btn sm onClick={()=>setModal("produto")}><Icon name="plus" size={13}/> Produto</Btn>}
                </div>
              </div>
              <div style={{display:"grid",gap:".65rem"}}>
                {products.map(p=>(
                  <div key={p.id} style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",padding:".85rem 1rem"}}>
                    <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".6rem"}}>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:".4rem",flexWrap:"wrap",marginBottom:".2rem"}}>
                          <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",color:"var(--text)"}}>{CATEGORY_ICONS[p.category]||"📋"} {p.name}</span>
                          {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                        </div>
                        <div style={{fontSize:".67rem",color:"var(--text5)"}}>Cód: {p.code}{p.batch&&` · Lote: ${p.batch}`} · {p.units_per_pack>1?`${p.units_per_pack} ${p.unit||"un"}/embalagem`:p.unit||"un"}</div>
                        {p.description&&<div style={{fontSize:".7rem",color:"var(--text4)",marginTop:".15rem"}}>{p.description}</div>}
                      </div>
                      {(isAdmin||canEdit)&&<div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".75rem"}}>
                        <button onClick={()=>{setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5),units_per_pack:String(p.units_per_pack||1)});setModal("editProd");}} style={{background:"none",border:"none",cursor:"pointer",color:"#4f5ef0"}}><Icon name="edit" size={14}/></button>
                        {isAdmin&&<button onClick={()=>delProduct(p.id)} style={{background:"none",border:"none",cursor:"pointer",color:"var(--text6)"}}><Icon name="trash" size={14}/></button>}
                      </div>}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(80px,1fr))",gap:".4rem"}}>
                      {[{l:"Estoque",v:`${p.stock_qty} ${p.unit||"un"}`,c:stkColor(p.stock_qty)},{l:"Custo",v:fmt(p.cost_per_unit),c:"var(--text3)"},{l:"Preço",v:fmt(p.price_per_unit),c:"#10b981"},{l:"Markup",v:fmtPct(p.markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(p.margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(p.profit),c:"#10b981"}].map(m=>(
                        <Pill key={m.l} label={m.l} value={m.v} color={m.c}/>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ════ USUÁRIOS ════ */}
          {tab==="usuarios"&&isAdmin&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".85rem"}}>
                <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem"}}>Usuários</h2>
                <Btn sm onClick={()=>setModal("addUser")}><Icon name="plus" size={13}/> Novo</Btn>
              </div>
              <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:".75rem",overflow:"hidden"}}>
                {appUsers.map(u=>(
                  <div key={u.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".85rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                    <div style={{display:"flex",alignItems:"center",gap:".75rem"}}>
                      <div style={{width:34,height:34,borderRadius:"50%",background:u.active?"linear-gradient(135deg,#4f5ef0,#10b981)":"var(--border2)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".85rem",color:"#fff",flexShrink:0}}>{u.display_name.charAt(0).toUpperCase()}</div>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
                          <span style={{fontWeight:700,fontSize:".86rem",color:u.active?"var(--text)":"var(--text5)"}}>{u.display_name}</span>
                          <Badge color={u.role==="admin"?"#f59e0b":u.role==="operator"?"#4f5ef0":"var(--text5)"} sm>{ROLES[u.role]}</Badge>
                          {!u.active&&<Badge color="#f56565" sm>Inativo</Badge>}
                          {u.id===cu.id&&<Badge color="#10b981" sm>Você</Badge>}
                        </div>
                        <div style={{fontSize:".67rem",color:"var(--text5)"}}>@{u.username}{u.last_login&&` · ${new Date(u.last_login).toLocaleDateString("pt-BR")}`}</div>
                      </div>
                    </div>
                    <button onClick={()=>{setEditing({...u,new_password:"",new_password2:""});setModal("editUser");}} style={{background:"none",border:"none",cursor:"pointer",color:"#4f5ef0"}}><Icon name="edit" size={14}/></button>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </div>

      {/* ════ MODAIS ════ */}

      {/* NOVA VENDA */}
      {modal==="sale"&&(
        <Modal title="Registrar Venda" onClose={()=>setModal(null)} icon="sales" wide>
          <div style={{background:"var(--bg)",borderRadius:".5rem",padding:".6rem .85rem",marginBottom:".85rem",fontSize:".75rem",color:"#4f5ef0",display:"flex",gap:".35rem",alignItems:"center"}}><Icon name="info" size={12}/> Venda baixa estoque e lança automaticamente no caixa</div>
          <Sel label="Produto *" value={sf.product_id} onChange={e=>handleSaleProd(e.target.value)}>
            <option value="">Selecione o produto...</option>
            {products.map(p=><option key={p.id} value={p.id}>{CATEGORY_ICONS[p.category]||"📋"} {p.name} — Estoque: {p.stock_qty} {p.unit||"un"}</option>)}
          </Sel>
          <Row>
            <Inp label="Quantidade *" type="number" min="1" value={sf.quantity} onChange={e=>handleSaleQty(e.target.value)}/>
            <Inp label="Preço unitário (R$) *" type="number" min="0" step="0.01" value={sf.unit_price} onChange={e=>handleSalePrice(e.target.value)}/>
          </Row>
          {sf.product_id&&(
            <div style={{background:"var(--bg)",borderRadius:".5rem",padding:".75rem",marginBottom:".85rem",display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem",textAlign:"center"}}>
              {[{l:"Quantidade",v:`${sf.quantity} un`,c:"var(--text)"},{l:"Preço unit.",v:fmt(parseFloat(sf.unit_price)||0),c:"#4f5ef0"},{l:"Total",v:fmt(parseFloat(sf.total_price)||0),c:"#10b981"}].map(m=>(
                <div key={m.l}><div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div><div style={{fontSize:"1rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div></div>
              ))}
            </div>
          )}
          <Sel label="Cliente" hint="opcional" value={sf.client_id} onChange={e=>{const c=clients.find(x=>x.id===e.target.value);setSf(f=>({...f,client_id:e.target.value,client_name:c?.name||""}));}}>
            <option value="">Sem cliente</option>
            {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
          </Sel>
          <Row>
            <Sel label="Pagamento" value={sf.payment_method} onChange={e=>setSf(f=>({...f,payment_method:e.target.value}))}>
              {PAYMENT_METHODS.map(m=><option key={m} value={m}>{m}</option>)}
            </Sel>
            <Inp label="Observações" hint="opcional" placeholder="Ex: Dose 2.5mg..." value={sf.notes} onChange={e=>setSf(f=>({...f,notes:e.target.value}))}/>
          </Row>
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn v="success" onClick={registerSale} disabled={!sf.product_id||!sf.unit_price}><Icon name="save" size={13}/> Confirmar venda</Btn>
          </div>
        </Modal>
      )}

      {/* ENTRADA ESTOQUE */}
      {modal==="stockEntry"&&(
        <Modal title="Entrada de Estoque" onClose={()=>setModal(null)} icon="stock">
          <div style={{background:"var(--bg)",borderRadius:".5rem",padding:".6rem .85rem",marginBottom:".85rem",fontSize:".75rem",color:"#10b981",display:"flex",gap:".35rem",alignItems:"center"}}><Icon name="arrow_up" size={12}/> Entrada gera saída no caixa (custo de compra)</div>
          <Sel label="Produto *" value={stockF.product_id} onChange={e=>setStockF(s=>({...s,product_id:e.target.value}))}>
            <option value="">Selecione o produto...</option>
            {products.map(p=><option key={p.id} value={p.id}>{CATEGORY_ICONS[p.category]||"📋"} {p.name} — Atual: {p.stock_qty} {p.unit||"un"}</option>)}
          </Sel>
          <Row>
            <Inp label="Quantidade *" type="number" min="1" placeholder={`Ex: 10 ${stockF.product_id?products.find(p=>p.id===stockF.product_id)?.unit||"un":"un"}`} value={stockF.qty} onChange={e=>setStockF(s=>({...s,qty:e.target.value}))}/>
            <Inp label="Custo total (R$)" hint="opcional" type="number" min="0" step="0.01" placeholder="0,00" value={stockF.cost_total} onChange={e=>setStockF(s=>({...s,cost_total:e.target.value}))}/>
          </Row>
          <Inp label="Observações" hint="opcional" placeholder="Ex: NF 1234, Fornecedor X..." value={stockF.notes} onChange={e=>setStockF(s=>({...s,notes:e.target.value}))}/>
          {stockF.product_id&&stockF.qty&&(()=>{
            const prod=products.find(p=>p.id===stockF.product_id);
            const qty=parseInt(stockF.qty)||0;
            return qty>0&&<div style={{background:"var(--info-box)",borderRadius:".5rem",padding:".6rem .85rem",marginBottom:".85rem",fontSize:".78rem",color:"#10b981"}}>✅ +{qty} {prod?.unit||"un"} → novo estoque: {(prod?.stock_qty||0)+qty} {prod?.unit||"un"}</div>;
          })()}
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn v="success" onClick={registerStockEntry}><Icon name="save" size={13}/> Confirmar entrada</Btn>
          </div>
        </Modal>
      )}

      {/* LANÇAMENTO CAIXA */}
      {modal==="cashTx"&&(
        <Modal title="Lançamento Manual" onClose={()=>setModal(null)} icon="cash">
          <Sel label="Tipo" value={cashF.type} onChange={e=>setCashF(f=>({...f,type:e.target.value}))}>
            <option value="entrada">📈 Entrada</option>
            <option value="saida">📉 Saída</option>
          </Sel>
          <Inp label="Descrição *" placeholder="Ex: Pagamento fornecedor..." value={cashF.description} onChange={e=>setCashF(f=>({...f,description:e.target.value}))}/>
          <Inp label="Valor (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={cashF.value} onChange={e=>setCashF(f=>({...f,value:e.target.value}))}/>
          <Inp label="Categoria" placeholder="Ex: Despesa fixa, Marketing..." value={cashF.category} onChange={e=>setCashF(f=>({...f,category:e.target.value}))}/>
          {/* Produto vinculado com baixa de estoque */}
          {cashF.type==="saida"&&(
            <>
              <Sel label="Produto vinculado" hint="opcional — baixa estoque" value={cashF.product_id} onChange={e=>setCashF(f=>({...f,product_id:e.target.value}))}>
                <option value="">Nenhum</option>
                {products.map(p=><option key={p.id} value={p.id}>{p.name} (estoque: {p.stock_qty})</option>)}
              </Sel>
              {cashF.product_id&&<Inp label="Quantidade" type="number" min="1" value={cashF.quantity} onChange={e=>setCashF(f=>({...f,quantity:e.target.value}))}/>}
            </>
          )}
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn v="success" onClick={addCashTx} disabled={!cashF.description||!cashF.value}><Icon name="save" size={13}/> Salvar</Btn>
          </div>
        </Modal>
      )}

      {/* FORNECEDOR */}
      {modal==="supplier"&&(
        <Modal title="Fornecedores" onClose={()=>setModal(null)} icon="supplier" wide>
          <div style={{display:"flex",justifyContent:"flex-end",marginBottom:"1rem"}}>
            <Btn sm onClick={()=>setModal("addSupplier")}><Icon name="plus" size={12}/> Novo fornecedor</Btn>
          </div>
          {suppliers.length===0?<p style={{color:"var(--text5)",textAlign:"center",padding:"2rem 0",fontSize:".82rem"}}>Nenhum fornecedor cadastrado.</p>
          :<div style={{display:"grid",gap:".6rem"}}>
            {suppliers.map(s=>(
              <div key={s.id} style={{background:"var(--bg4)",border:"1px solid var(--border2)",borderRadius:".6rem",padding:".75rem 1rem",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <div>
                  <div style={{fontWeight:700,fontSize:".88rem",color:"var(--text)"}}>🏭 {s.name}</div>
                  <div style={{fontSize:".7rem",color:"var(--text5)",marginTop:".15rem"}}>{[s.contact,s.phone,s.email].filter(Boolean).join(" · ")||"Sem contato"}</div>
                  {s.notes&&<div style={{fontSize:".68rem",color:"#8b44f0",marginTop:".1rem"}}>📝 {s.notes}</div>}
                </div>
                {isAdmin&&<div style={{display:"flex",gap:".35rem"}}>
                  <button onClick={()=>{setEditing({...s});setModal("editSupplier");}} style={{background:"none",border:"none",cursor:"pointer",color:"#4f5ef0"}}><Icon name="edit" size={13}/></button>
                  <button onClick={async()=>{await supabase.from("suppliers").delete().eq("id",s.id);toast$("Fornecedor removido.","#f59e0b");}} style={{background:"none",border:"none",cursor:"pointer",color:"var(--text6)"}}><Icon name="trash" size={13}/></button>
                </div>}
              </div>
            ))}
          </div>}
        </Modal>
      )}

      {/* NOVO FORNECEDOR */}
      {modal==="addSupplier"&&(
        <Modal title="Novo Fornecedor" onClose={()=>setModal(null)} icon="supplier">
          <Inp label="Nome *" placeholder="Ex: Farmácia Distribuidora X" value={supf.name} onChange={e=>setSupf(f=>({...f,name:e.target.value}))}/>
          <Inp label="Contato" placeholder="Nome do responsável" value={supf.contact} onChange={e=>setSupf(f=>({...f,contact:e.target.value}))}/>
          <Row>
            <Inp label="Telefone / WhatsApp" type="tel" placeholder="(66) 99999-9999" value={supf.phone} onChange={e=>setSupf(f=>({...f,phone:e.target.value}))}/>
            <Inp label="E-mail" type="email" placeholder="contato@fornecedor.com" value={supf.email} onChange={e=>setSupf(f=>({...f,email:e.target.value}))}/>
          </Row>
          <Inp label="Observações" hint="opcional" placeholder="Ex: prazo de entrega, condições..." value={supf.notes} onChange={e=>setSupf(f=>({...f,notes:e.target.value}))}/>
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>setModal("supplier")}>Voltar</Btn>
            <Btn onClick={addSupplier}><Icon name="save" size={13}/> Salvar</Btn>
          </div>
        </Modal>
      )}

      {/* EDITAR FORNECEDOR */}
      {modal==="editSupplier"&&editing&&(
        <Modal title="Editar Fornecedor" onClose={()=>{setModal("supplier");setEditing(null);}} icon="edit">
          <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
          <Inp label="Contato" value={editing.contact||""} onChange={e=>setEditing(v=>({...v,contact:e.target.value}))}/>
          <Row>
            <Inp label="Telefone" type="tel" value={editing.phone||""} onChange={e=>setEditing(v=>({...v,phone:e.target.value}))}/>
            <Inp label="E-mail" type="email" value={editing.email||""} onChange={e=>setEditing(v=>({...v,email:e.target.value}))}/>
          </Row>
          <Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/>
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>{setModal("supplier");setEditing(null);}}>Cancelar</Btn>
            <Btn onClick={saveSupplier}><Icon name="save" size={13}/> Salvar</Btn>
          </div>
        </Modal>
      )}

      {/* NOVO PRODUTO */}
      {modal==="produto"&&(
        <Modal title="Novo Produto" onClose={()=>setModal(null)} icon="product" wide>
          <Row>
            <Inp label="Código" hint="auto se vazio" placeholder="PRD-001" value={pf.code} onChange={e=>setPf(f=>({...f,code:e.target.value}))}/>
            <Sel label="Categoria" value={pf.category} onChange={e=>setPf(f=>({...f,category:e.target.value}))}>
              <option value="tirzepatida">💉 Tirzepatida</option>
              <option value="seringa">🩺 Seringa/Agulha</option>
              <option value="material">🧊 Material (gelo, isopor...)</option>
              <option value="embalagem">📦 Embalagem</option>
              <option value="outro">📋 Outro</option>
            </Sel>
          </Row>
          <Inp label="Nome *" placeholder="Ex: Tirzepatida 2.5mg, Caixa isopor..." value={pf.name} onChange={e=>setPf(f=>({...f,name:e.target.value}))}/>
          <Inp label="Descrição" hint="opcional" placeholder="Ex: Caneta injetável, unidade..." value={pf.description} onChange={e=>setPf(f=>({...f,description:e.target.value}))}/>
          <Row>
            <Sel label="Fornecedor" hint="opcional" value={pf.supplier_id} onChange={e=>setPf(f=>({...f,supplier_id:e.target.value}))}>
              <option value="">Nenhum</option>
              {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </Sel>
            <Inp label="Unidade" hint="ex: un, ampola, cx..." placeholder="un" value={pf.unit} onChange={e=>setPf(f=>({...f,unit:e.target.value}))}/>
          </Row>
          <Row>
            <Inp label="Lote" placeholder="L2025001" value={pf.batch} onChange={e=>setPf(f=>({...f,batch:e.target.value}))}/>
            <Inp label="Vencimento" type="date" value={pf.expiry} onChange={e=>setPf(f=>({...f,expiry:e.target.value}))}/>
          </Row>
          <Row>
            <Inp label="Custo/unidade (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={pf.cost_per_unit} onChange={e=>setPf(f=>({...f,cost_per_unit:e.target.value}))}/>
            <Inp label="Preço venda/unidade (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={pf.price_per_unit} onChange={e=>setPf(f=>({...f,price_per_unit:e.target.value}))}/>
          </Row>
          {pf.cost_per_unit&&pf.price_per_unit&&(()=>{
            const {markup,margin,profit}=calcM(pf.cost_per_unit,pf.price_per_unit);
            return <div style={{background:"var(--bg)",borderRadius:".5rem",padding:".75rem",marginBottom:".85rem",display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem",textAlign:"center"}}>
              {[{l:"Markup",v:fmtPct(markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(profit),c:"#10b981"}].map(m=><div key={m.l}><div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div><div style={{fontSize:".9rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div></div>)}
            </div>;
          })()}
          <Row>
            <Inp label="Estoque inicial" type="number" min="0" value={pf.stock_qty} onChange={e=>setPf(f=>({...f,stock_qty:e.target.value}))}/>
            <Inp label="Estoque mínimo" hint="alerta abaixo desse valor" type="number" min="0" value={pf.min_stock} onChange={e=>setPf(f=>({...f,min_stock:e.target.value}))}/>
          </Row>
          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between",alignItems:"center"}}>
            <Btn v="ghost" sm onClick={()=>setModal("addSupplier")}><Icon name="supplier" size={12}/> + Fornecedor</Btn>
            <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addProduct}><Icon name="save" size={13}/> Salvar</Btn></div>
          </div>
        </Modal>
      )}

      {/* EDITAR PRODUTO */}
      {modal==="editProd"&&editing&&(
        <Modal title="Editar Produto" onClose={()=>{setModal(null);setEditing(null);}} icon="edit" wide>
          <Row>
            <Inp label="Código" value={editing.code||""} onChange={e=>setEditing(v=>({...v,code:e.target.value}))}/>
            <Sel label="Categoria" value={editing.category||"outro"} onChange={e=>setEditing(v=>({...v,category:e.target.value}))}>
              <option value="tirzepatida">💉 Tirzepatida</option>
              <option value="seringa">🩺 Seringa/Agulha</option>
              <option value="material">🧊 Material</option>
              <option value="embalagem">📦 Embalagem</option>
              <option value="outro">📋 Outro</option>
            </Sel>
          </Row>
          <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
          <Inp label="Descrição" value={editing.description||""} onChange={e=>setEditing(v=>({...v,description:e.target.value}))}/>
          <Row>
            <Sel label="Fornecedor" value={editing.supplier_id||""} onChange={e=>setEditing(v=>({...v,supplier_id:e.target.value}))}>
              <option value="">Nenhum</option>
              {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </Sel>
            <Inp label="Unidade" value={editing.unit||"un"} onChange={e=>setEditing(v=>({...v,unit:e.target.value}))}/>
          </Row>
          <Row>
            <Inp label="Lote" value={editing.batch||""} onChange={e=>setEditing(v=>({...v,batch:e.target.value}))}/>
            <Inp label="Vencimento" type="date" value={editing.expiry||""} onChange={e=>setEditing(v=>({...v,expiry:e.target.value}))}/>
          </Row>
          <Row>
            <Inp label="Custo/unidade (R$)" type="number" min="0" step="0.01" value={editing.cost_per_unit} onChange={e=>setEditing(v=>({...v,cost_per_unit:e.target.value}))}/>
            <Inp label="Preço venda/unidade (R$)" type="number" min="0" step="0.01" value={editing.price_per_unit} onChange={e=>setEditing(v=>({...v,price_per_unit:e.target.value}))}/>
          </Row>
          {editing.cost_per_unit&&editing.price_per_unit&&(()=>{
            const {markup,margin,profit}=calcM(editing.cost_per_unit,editing.price_per_unit);
            return <div style={{background:"var(--bg)",borderRadius:".5rem",padding:".65rem",marginBottom:".85rem",display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".4rem",textAlign:"center"}}>
              {[{l:"Markup",v:fmtPct(markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(profit),c:"#10b981"}].map(m=><div key={m.l}><div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".12rem"}}>{m.l}</div><div style={{fontSize:".85rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div></div>)}
            </div>;
          })()}
          <Row>
            <Inp label="Quantidade em estoque" type="number" min="0" value={editing.stock_qty} onChange={e=>setEditing(v=>({...v,stock_qty:e.target.value}))}/>
            <Inp label="Estoque mínimo" type="number" min="0" value={editing.min_stock} onChange={e=>setEditing(v=>({...v,min_stock:e.target.value}))}/>
          </Row>
          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
            {isAdmin&&<Btn v="danger" sm onClick={()=>delProduct(editing.id)}><Icon name="trash" size={12}/> Excluir</Btn>}
            <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveProduct}><Icon name="save" size={13}/> Salvar</Btn></div>
          </div>
        </Modal>
      )}

      {/* CLIENTE */}
      {modal==="cliente"&&(
        <Modal title="Novo Cliente" onClose={()=>setModal(null)} icon="client">
          <Inp label="Nome *" placeholder="Nome completo" value={cf.name} onChange={e=>setCf(f=>({...f,name:e.target.value}))}/>
          <Row>
            <Inp label="Telefone / WhatsApp" type="tel" placeholder="(66) 99999-9999" value={cf.phone} onChange={e=>setCf(f=>({...f,phone:e.target.value}))}/>
            <Inp label="E-mail" type="email" placeholder="email@exemplo.com" value={cf.email} onChange={e=>setCf(f=>({...f,email:e.target.value}))}/>
          </Row>
          <Inp label="Observações" hint="ex: dosagem, periodicidade..." placeholder="Ex: Usa 2.5mg a cada 7 dias" value={cf.notes} onChange={e=>setCf(f=>({...f,notes:e.target.value}))}/>
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addClient}><Icon name="save" size={13}/> Salvar</Btn></div>
        </Modal>
      )}

      {/* EDITAR CLIENTE */}
      {modal==="editCliente"&&editing&&(
        <Modal title="Editar Cliente" onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
          <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
          <Row>
            <Inp label="Telefone" type="tel" value={editing.phone||""} onChange={e=>setEditing(v=>({...v,phone:e.target.value}))}/>
            <Inp label="E-mail" type="email" value={editing.email||""} onChange={e=>setEditing(v=>({...v,email:e.target.value}))}/>
          </Row>
          <Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/>
          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
            {isAdmin&&<Btn v="danger" sm onClick={async()=>{await supabase.from("clients").delete().eq("id",editing.id);toast$("Removido.","#f59e0b");setModal(null);setEditing(null);}}><Icon name="trash" size={12}/> Excluir</Btn>}
            <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveClient}><Icon name="save" size={13}/> Salvar</Btn></div>
          </div>
        </Modal>
      )}

      {/* NOVO USUÁRIO */}
      {modal==="addUser"&&(
        <Modal title="Novo Usuário" onClose={()=>setModal(null)} icon="users">
          <Row>
            <Inp label="Login *" placeholder="ex: maria.silva" value={uf.username} onChange={e=>setUf(f=>({...f,username:e.target.value.toLowerCase().replace(/\s/g,".")}))}/>
            <Inp label="Nome de exibição *" placeholder="ex: Maria Silva" value={uf.display_name} onChange={e=>setUf(f=>({...f,display_name:e.target.value}))}/>
          </Row>
          <Sel label="Nível" value={uf.role} onChange={e=>setUf(f=>({...f,role:e.target.value}))}>
            <option value="admin">Administrador — acesso total</option>
            <option value="operator">Operador — adiciona e edita</option>
            <option value="viewer">Visualizador — somente leitura</option>
          </Sel>
          <Row>
            <Field label="Senha *" hint="mín. 6 caracteres">
              <div style={{position:"relative"}}><input type={pwShow.a?"text":"password"} value={uf.password} onChange={e=>setUf(f=>({...f,password:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwShow(s=>({...s,a:!s.a}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--text5)",padding:0}}><Icon name={pwShow.a?"eye_off":"eye"} size={13}/></button></div>
            </Field>
            <Field label="Confirmar senha">
              <div style={{position:"relative"}}><input type={pwShow.b?"text":"password"} value={uf.password2} onChange={e=>setUf(f=>({...f,password2:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwShow(s=>({...s,b:!s.b}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--text5)",padding:0}}><Icon name={pwShow.b?"eye_off":"eye"} size={13}/></button></div>
            </Field>
          </Row>
          {uf.password&&uf.password2&&uf.password!==uf.password2&&<div style={{fontSize:".75rem",color:"#f56565",marginBottom:".75rem",display:"flex",gap:".3rem",alignItems:"center"}}><Icon name="warn" size={12}/> Senhas não coincidem</div>}
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addUser} disabled={!uf.username||!uf.display_name||!uf.password||uf.password!==uf.password2}><Icon name="save" size={13}/> Criar</Btn></div>
        </Modal>
      )}

      {/* EDITAR USUÁRIO */}
      {modal==="editUser"&&editing&&(
        <Modal title={`Editar · ${editing.display_name}`} onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
          <div style={{background:"var(--bg)",borderRadius:".45rem",padding:".55rem .85rem",marginBottom:".85rem",fontSize:".75rem",color:"var(--text5)"}}>Login: <b style={{color:"var(--text4)"}}>@{editing.username}</b></div>
          <Row>
            <Inp label="Nome de exibição *" value={editing.display_name} onChange={e=>setEditing(v=>({...v,display_name:e.target.value}))}/>
            <Sel label="Nível" value={editing.role} onChange={e=>setEditing(v=>({...v,role:e.target.value}))} disabled={editing.id===cu.id}>
              <option value="admin">Administrador</option>
              <option value="operator">Operador</option>
              <option value="viewer">Visualizador</option>
            </Sel>
          </Row>
          <Field label="Status">
            <div style={{display:"flex",gap:".5rem"}}>
              {[{v:true,l:"✅ Ativo"},{v:false,l:"🚫 Inativo"}].map(o=>(
                <button key={String(o.v)} onClick={()=>setEditing(v=>({...v,active:o.v}))} disabled={editing.id===cu.id} style={{flex:1,padding:".55rem",borderRadius:".45rem",border:`1px solid ${editing.active===o.v?"#4f5ef0":"var(--border2)"}`,background:editing.active===o.v?"#4f5ef018":"transparent",color:editing.active===o.v?"#4f5ef0":"var(--text5)",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:".82rem"}}>{o.l}</button>
              ))}
            </div>
          </Field>
          <div style={{borderTop:"1px solid var(--border)",paddingTop:".85rem",marginTop:".25rem",marginBottom:".85rem"}}>
            <div style={{fontSize:".72rem",color:"#4f5ef0",marginBottom:".65rem",display:"flex",gap:".3rem",alignItems:"center"}}><Icon name="key" size={12}/> Nova senha (deixe vazio para manter)</div>
            <Row>
              <Field label="Nova senha"><div style={{position:"relative"}}><input type={pwShow.c?"text":"password"} value={editing.new_password||""} onChange={e=>setEditing(v=>({...v,new_password:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwShow(s=>({...s,c:!s.c}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--text5)",padding:0}}><Icon name={pwShow.c?"eye_off":"eye"} size={13}/></button></div></Field>
              <Field label="Confirmar"><div style={{position:"relative"}}><input type={pwShow.d?"text":"password"} value={editing.new_password2||""} onChange={e=>setEditing(v=>({...v,new_password2:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwShow(s=>({...s,d:!s.d}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--text5)",padding:0}}><Icon name={pwShow.d?"eye_off":"eye"} size={13}/></button></div></Field>
            </Row>
          </div>
          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
            {editing.id!==cu.id?<Btn v="danger" sm onClick={async()=>{await supabase.from("app_users").delete().eq("id",editing.id);toast$("Removido.","#f59e0b");setModal(null);setEditing(null);}}><Icon name="trash" size={12}/> Excluir</Btn>:<div/>}
            <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveUser}><Icon name="save" size={13}/> Salvar</Btn></div>
          </div>
        </Modal>
      )}

      {/* LOGOUT */}
      {logoutC&&(
        <div style={{position:"fixed",inset:0,background:"var(--shadow)",backdropFilter:"blur(8px)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:"1.25rem"}}>
          <div style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"1rem",padding:"1.75rem",maxWidth:300,width:"100%",textAlign:"center"}}>
            <div style={{color:"#4f5ef0",marginBottom:".85rem"}}><Icon name="logout" size={32}/></div>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".95rem",color:"var(--text)",marginBottom:".4rem"}}>Sair do sistema?</div>
            <div style={{fontSize:".78rem",color:"var(--text5)",marginBottom:"1.25rem"}}>Precisará fazer login novamente.</div>
            <div style={{display:"flex",gap:".65rem",justifyContent:"center"}}><Btn v="ghost" onClick={()=>setLogoutC(false)}>Cancelar</Btn><Btn v="danger" onClick={logout}><Icon name="logout" size={13}/> Sair</Btn></div>
          </div>
        </div>
      )}
    </>
  );
}
