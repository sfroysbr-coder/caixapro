import React,{useState,useEffect,useCallback,useMemo}from"react";
import{supabase}from"./supabase";

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0);
const fmtN=v=>new Intl.NumberFormat("pt-BR").format(v||0);
const fmtPct=v=>`${(v||0).toFixed(1)}%`;
const uid=()=>Math.random().toString(36).slice(2,10);
const today=()=>new Date().toLocaleDateString("pt-BR");
const nowISO=()=>new Date().toISOString();
const hashPw=s=>btoa(unescape(encodeURIComponent(s+"|caixapro2026")));
const calcM=(c,p)=>{const cv=parseFloat(c)||0,pv=parseFloat(p)||0;return{markup:cv>0?((pv-cv)/cv)*100:0,margin:pv>0?((pv-cv)/pv)*100:0,profit:pv-cv};};
const daysUntil=d=>d?Math.ceil((new Date(d)-new Date())/86400000):null;
const expColor=d=>d===null?"var(--tx5)":d<0?"#f56565":d<=30?"#f59e0b":"#10b981";
const stkColor=q=>q<=0?"#f56565":q<=5?"#f59e0b":"#10b981";

const ROLES={admin:"Admin",operator:"Operador",viewer:"Visualizador"};
const CATS={tirzepatida:"💉",material:"🧊",embalagem:"📦",seringa:"🩺",outro:"📋"};
const PAYS=["Dinheiro","PIX","Cartão Crédito","Cartão Débito","Transferência"];

// Tema
const getTheme=()=>{try{return localStorage.getItem("cpro:theme")!=="light"}catch{return true}};
const setTheme=dark=>{try{localStorage.setItem("cpro:theme",dark?"dark":"light")}catch{}};

// Export
const exportXLS=(rows,name,sheet="Dados")=>{
  if(!window.XLSX){alert("Aguarde carregar a página.");return;}
  const ws=window.XLSX.utils.json_to_sheet(rows);
  const wb=window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb,ws,sheet);
  window.XLSX.writeFile(wb,name+".xlsx");
};
const exportPDF=(cols,rows,name,title)=>{
  if(!window.jspdf){alert("Aguarde carregar a página.");return;}
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});
  doc.setFillColor(13,15,26);doc.rect(0,0,297,18,"F");
  doc.setTextColor(232,234,246);doc.setFontSize(11);doc.setFont("helvetica","bold");
  doc.text("CaixaPro · Tirzepatida",14,12);
  doc.setFontSize(9);doc.setFont("helvetica","normal");doc.setTextColor(160,160,200);
  doc.text(title,110,12);doc.text(new Date().toLocaleString("pt-BR"),240,12);
  doc.autoTable({head:[cols.map(c=>c.l)],body:rows.map(r=>cols.map(c=>String(r[c.k]??""))),startY:22,
    styles:{fontSize:7,cellPadding:2,textColor:[40,40,70]},
    headStyles:{fillColor:[79,94,240],textColor:[255,255,255],fontStyle:"bold"},
    alternateRowStyles:{fillColor:[245,247,255]},margin:{left:12,right:12}});
  doc.save(name+".pdf");
};

// ─── ÍCONES ───────────────────────────────────────────────────────────────────
const Ic=({n,s=18})=>{
  const P={
    dashboard:<path d="M3 3h8v8H3zm10 0h8v8h-8zM3 13h8v8H3zm10 0h8v8h-8z"/>,
    cash:<><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M12 12m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0"/></>,
    product:<><path d="M20.91 8.84L8.56 2.23a2 2 0 0 0-2.37.46L2.46 8.84A2 2 0 0 0 2 10v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V10a2 2 0 0 0-.46-1.16zM12 20v-6"/><path d="M2 10h20"/></>,
    client:<><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></>,
    users:<><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></>,
    stock:<><path d="M5 8h14M5 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4zm14 0a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM5 8v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8"/><path d="M10 12h4"/></>,
    sales:<><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 0 1-8 0"/></>,
    supplier:<><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></>,
    analytics:<><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/><line x1="2" y1="20" x2="22" y2="20"/></>,
    plus:<><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></>,
    up:<><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></>,
    dn:<><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></>,
    trash:<><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></>,
    edit:<><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></>,
    close:<><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></>,
    sync:<><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></>,
    warn:<><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></>,
    search:<><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></>,
    logout:<><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></>,
    eye:<><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></>,
    eyeoff:<><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></>,
    lock:<><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></>,
    user:<><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></>,
    save:<><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></>,
    arrup:<><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></>,
    key:<><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></>,
    syringe:<><path d="M18 2l4 4m-4-4L7 13m11-11l-4 4M3 21l4.5-4.5"/><path d="M9 15l-3 3"/><path d="M14.5 9.5l-5 5"/></>,
    info:<><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></>,
    cal:<><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></>,
    dollar:<><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></>,
    award:<><circle cx="12" cy="8" r="6"/><path d="M15.477 12.89L17 22l-5-3-5 3 1.523-9.11"/></>,
    pkg:<><line x1="16.5" y1="9.4" x2="7.5" y2="4.21"/><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></>,
    sun:<><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></>,
    moon:<><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></>,
    xls:<><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="m6 13 3 4m0-4-3 4"/></>,
    pdf:<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></>,
  };
  return <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{P[n]||<circle cx="12" cy="12" r="10"/>}</svg>;
};

// ─── UI PRIMITIVES ────────────────────────────────────────────────────────────
const IS={width:"100%",background:"var(--inp)",border:"1px solid var(--bdr2)",borderRadius:".45rem",padding:".55rem .8rem",color:"var(--tx)",fontSize:".84rem",fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box"};

const Btn=({children,onClick,v="p",sm,disabled,full})=>{
  const S={
    p:{background:"linear-gradient(135deg,#4f5ef0,#8b44f0)",color:"#fff",border:"none"},
    ok:{background:"linear-gradient(135deg,#10b981,#059669)",color:"#fff",border:"none"},
    del:{background:"#1e1010",color:"#f56565",border:"1px solid #3a1515"},
    ghost:{background:"transparent",color:"var(--tx4)",border:"1px solid var(--bdr2)"},
    info:{background:"linear-gradient(135deg,#0891b2,#0e7490)",color:"#fff",border:"none"},
    warn:{background:"#1e1500",color:"#f59e0b",border:"1px solid #3a2800"},
    xls:{background:"#14532d",color:"#4ade80",border:"1px solid #166834"},
    pdf:{background:"#450a0a",color:"#f87171",border:"1px solid #7f1d1d"},
  };
  return <button onClick={onClick} disabled={disabled} style={{display:"inline-flex",alignItems:"center",gap:".35rem",padding:sm?".32rem .65rem":".52rem .95rem",borderRadius:".45rem",fontSize:sm?".73rem":".83rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,opacity:disabled?.5:1,width:full?"100%":undefined,justifyContent:full?"center":undefined,...(S[v]||S.p)}}>{children}</button>;
};

const Badge=({children,color="#4f5ef0",sm})=><span style={{fontSize:sm?".6rem":".65rem",fontWeight:600,color,background:color+"20",borderRadius:"99px",padding:sm?".1rem .4rem":".12rem .5rem",border:`1px solid ${color}35`,whiteSpace:"nowrap"}}>{children}</span>;

const KCard=({label,value,sub,color="#4f5ef0",icon})=>(
  <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".85rem 1rem",position:"relative",overflow:"hidden"}}>
    <div style={{position:"absolute",top:"-1rem",right:"-1rem",width:"4rem",height:"4rem",borderRadius:"50%",background:color,opacity:.08}}/>
    {icon&&<div style={{position:"absolute",top:".8rem",right:".8rem",color,opacity:.45}}><Ic n={icon} s={14}/></div>}
    <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".07em",marginBottom:".28rem"}}>{label}</div>
    <div style={{fontSize:"1.05rem",fontWeight:700,color:"var(--tx)",fontFamily:"'Syne',sans-serif",letterSpacing:"-.02em"}}>{value}</div>
    {sub&&<div style={{fontSize:".67rem",color,marginTop:".18rem"}}>{sub}</div>}
  </div>
);

const Pill=({label,value,color})=>(
  <div style={{background:"var(--pill)",borderRadius:".4rem",padding:".38rem .55rem"}}>
    <div style={{fontSize:".56rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".12rem"}}>{label}</div>
    <div style={{fontSize:".78rem",fontWeight:700,fontFamily:"'Syne',sans-serif",color}}>{value}</div>
  </div>
);

const Modal=({title,onClose,children,wide,full,icon})=>(
  <div style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,.78)",backdropFilter:"blur(6px)",display:"flex",alignItems:"center",justifyContent:"center",padding:".75rem"}}>
    <div style={{background:"var(--bg3)",border:"1px solid var(--bdr2)",borderRadius:"1rem",width:"100%",maxWidth:full?"98vw":wide?700:500,maxHeight:"93vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 64px var(--shadow)"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"1rem 1.4rem",borderBottom:"1px solid var(--bdr)",position:"sticky",top:0,background:"var(--bg3)",borderRadius:"1rem 1rem 0 0",zIndex:1}}>
        <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
          {icon&&<span style={{color:"#4f5ef0"}}><Ic n={icon} s={15}/></span>}
          <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".95rem",color:"var(--tx)"}}>{title}</span>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",color:"var(--tx4)",padding:".2rem"}}><Ic n="close" s={16}/></button>
      </div>
      <div style={{padding:"1.25rem",overflowY:"auto",flex:1}}>{children}</div>
    </div>
  </div>
);

const Field=({label,hint,children})=>(
  <div style={{marginBottom:".8rem"}}>
    <div style={{display:"flex",justifyContent:"space-between",marginBottom:".3rem"}}>
      <label style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"'DM Sans',sans-serif"}}>{label}</label>
      {hint&&<span style={{fontSize:".65rem",color:"var(--tx5)"}}>{hint}</span>}
    </div>
    {children}
  </div>
);
const Inp=({label,hint,...p})=><Field label={label} hint={hint}><input {...p} style={{...IS,...p.style}}/></Field>;
const Sel=({label,hint,children,...p})=><Field label={label} hint={hint}><select {...p} style={{...IS,...p.style}}>{children}</select></Field>;
const R2=({children,gap=".65rem"})=>{const k=Array.isArray(children)?children.filter(Boolean):[children];return <div style={{display:"grid",gridTemplateColumns:`repeat(${k.length},1fr)`,gap}}>{children}</div>;};

// Botões Export
const XBtn=({rows,name,sheet})=><button onClick={()=>exportXLS(rows,name,sheet)} style={{display:"inline-flex",alignItems:"center",gap:".3rem",padding:".32rem .65rem",borderRadius:".4rem",background:"#14532d",color:"#4ade80",border:"1px solid #166834",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}><Ic n="xls" s={11}/>Excel</button>;
const PBtn=({cols,rows,name,title})=><button onClick={()=>exportPDF(cols,rows,name,title)} style={{display:"inline-flex",alignItems:"center",gap:".3rem",padding:".32rem .65rem",borderRadius:".4rem",background:"#450a0a",color:"#f87171",border:"1px solid #7f1d1d",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}><Ic n="pdf" s={11}/>PDF</button>;

// Sparkline
const Spark=({data=[],color="#4f5ef0",h=40})=>{
  const max=Math.max(...data,1);
  return <div style={{display:"flex",alignItems:"flex-end",gap:"2px",height:h}}>{data.map((v,i)=><div key={i} style={{flex:1,background:color,borderRadius:"2px 2px 0 0",height:`${(v/max)*100}%`,minHeight:2,opacity:.6+.4*(v/max)}}/>)}</div>;
};

// Donut simples
const Donut=({segs=[],size=80})=>{
  const total=segs.reduce((a,s)=>a+s.v,0);
  if(!total)return <div style={{width:size,height:size,borderRadius:"50%",background:"var(--bdr)",flexShrink:0}}/>;
  const r=26,cx=size/2,cy=size/2,circ=2*Math.PI*r;
  let off=0;
  return <svg width={size} height={size} style={{flexShrink:0,transform:"rotate(-90deg)"}}>
    {segs.map((s,i)=>{const pct=s.v/total,dash=pct*circ,gap=circ-dash;const el=<circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={s.c} strokeWidth={size/5} strokeDasharray={`${dash} ${gap}`} strokeDashoffset={-off*circ}/>;off+=pct;return el;})}
  </svg>;
};

// ─── TELA LOGIN ───────────────────────────────────────────────────────────────
function LoginScreen({onLogin,dark}){
  const[user,setUser]   = useState("");
  const[pw_,setPw]      = useState("");
  const[show,setShow]   = useState(false);
  const[error,setErr]   = useState("");
  const[shk,setShk]     = useState(false);
  const[loading,setLoad]= useState(false);

  const attempt=async()=>{
    if(!user.trim()||!pw_){setErr("Preencha usuário e senha.");return;}
    setLoad(true);await new Promise(r=>setTimeout(r,400));
    try{
      const{data,error:e}=await supabase.from("app_users").select("*").eq("username",user.trim().toLowerCase()).eq("active",true).single();
      if(e||!data){setErr("Usuário não encontrado.");setShk(true);setTimeout(()=>setShk(false),500);setLoad(false);return;}
      if(data.password_hash!==hashPw(pw_)){setErr("Senha incorreta.");setShk(true);setTimeout(()=>setShk(false),500);setLoad(false);return;}
      await supabase.from("app_users").update({last_login:nowISO()}).eq("id",data.id);
      onLogin(data);
    }catch{setErr("Erro de conexão.");}
    setLoad(false);
  };

  return(
    <div style={{minHeight:"100vh",background:"var(--bg)",display:"flex",alignItems:"center",justifyContent:"center",padding:"1.25rem",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:"-15%",left:"-10%",width:"50vw",height:"50vw",borderRadius:"50%",background:"radial-gradient(circle,#4f5ef020,transparent 70%)",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:"-15%",right:"-10%",width:"45vw",height:"45vw",borderRadius:"50%",background:"radial-gradient(circle,#10b98115,transparent 70%)",pointerEvents:"none"}}/>
      <div style={{width:"100%",maxWidth:380,background:"var(--login-bg)",border:"1px solid var(--bdr2)",borderRadius:"1.25rem",padding:"2rem 1.75rem",boxShadow:"0 30px 80px var(--shadow)",animation:shk?"loginShake .4s ease":"none",transition:"background .3s"}}>
        <div style={{textAlign:"center",marginBottom:"1.5rem"}}>
          <div style={{width:54,height:54,borderRadius:"1rem",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",marginBottom:".75rem",boxShadow:"0 8px 30px #4f5ef040"}}><Ic n="syringe" s={26}/></div>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.28rem",color:"var(--tx)",letterSpacing:"-.02em"}}>CaixaPro</div>
          <div style={{fontSize:".72rem",color:"var(--tx5)",marginTop:".2rem"}}>Gestão · Tirzepatida · v6.0</div>
        </div>
        <Field label="Usuário">
          <div style={{position:"relative"}}>
            <div style={{position:"absolute",left:".8rem",top:"50%",transform:"translateY(-50%)",color:"var(--tx5)"}}><Ic n="user" s={14}/></div>
            <input value={user} onChange={e=>{setUser(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="Digite seu usuário" style={{...IS,paddingLeft:"2.3rem"}}/>
          </div>
        </Field>
        <Field label="Senha">
          <div style={{position:"relative"}}>
            <div style={{position:"absolute",left:".8rem",top:"50%",transform:"translateY(-50%)",color:"var(--tx5)"}}><Ic n="lock" s={14}/></div>
            <input type={show?"text":"password"} value={pw_} onChange={e=>{setPw(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="••••••••••" style={{...IS,paddingLeft:"2.3rem",paddingRight:"2.5rem"}}/>
            <button onClick={()=>setShow(v=>!v)} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"var(--tx5)",padding:0}}><Ic n={show?"eyeoff":"eye"} s={14}/></button>
          </div>
        </Field>
        {error&&<div style={{background:"#f5656512",border:"1px solid #f5656530",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".8rem",fontSize:".78rem",color:"#f56565",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="warn" s={13}/>{error}</div>}
        <button onClick={attempt} disabled={loading} style={{width:"100%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",color:"#fff",border:"none",borderRadius:".5rem",padding:".75rem",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",opacity:loading?.7:1,marginBottom:".5rem"}}>
          {loading?"Verificando...":"Entrar no sistema"}
        </button>
        <div style={{textAlign:"center",fontSize:".63rem",color:"var(--tx6)"}}>CaixaPro © 2026 · Tirzepatida</div>
      </div>
    </div>
  );
}

// ─── ANALYTICS ────────────────────────────────────────────────────────────────
function Analytics({onClose,sales,cashTx,products,clients,dark}){
  const[range,setRange]=useState("30d");
  const[cf,setCf]=useState(""),ct=useState("");
  const[ct_,setCt]=ct;

  const fd=useCallback(items=>{
    if(range==="custom"&&cf&&ct_){const a=new Date(cf),b=new Date(ct_+"T23:59:59");return items.filter(i=>new Date(i.created_at)>=a&&new Date(i.created_at)<=b);}
    const days={"1d":1,"7d":7,"30d":30,"3m":90,"6m":180,"12m":365}[range];
    if(!days)return items;
    const from=new Date();from.setDate(from.getDate()-days);
    return items.filter(i=>new Date(i.created_at)>=from);
  },[range,cf,ct_]);

  const fS=useMemo(()=>fd(sales),[fd,sales]);
  const fC=useMemo(()=>fd(cashTx),[fd,cashTx]);

  const rev=fC.filter(t=>t.type==="entrada").reduce((a,t)=>a+t.value,0);
  const cost=fC.filter(t=>t.type==="saida").reduce((a,t)=>a+t.value,0);
  const profit=rev-cost;
  const margin=rev>0?(profit/rev)*100:0;
  const markup=cost>0?(profit/cost)*100:0;
  const units=fS.reduce((a,s)=>a+s.quantity,0);
  const ticket=fS.length>0?fS.reduce((a,s)=>a+s.total_price,0)/fS.length:0;
  const stockVal=products.reduce((a,p)=>a+p.stock_qty*(p.cost_per_unit||0),0);

  // Sparkline 7 dias
  const spark=useMemo(()=>{
    const b=Array(7).fill(0);const now=new Date();
    fC.filter(x=>x.type==="entrada").forEach(x=>{const d=Math.floor((now-new Date(x.created_at))/86400000);if(d<7)b[6-d]+=x.value;});
    return b;
  },[fC]);

  // Produtos ranking
  const prods=useMemo(()=>{
    const m={};fS.forEach(s=>{if(!s.product_name)return;if(!m[s.product_name])m[s.product_name]={n:s.product_name,r:0,q:0};m[s.product_name].r+=s.total_price;m[s.product_name].q+=s.quantity;});
    return Object.values(m).sort((a,b)=>b.r-a.r);
  },[fS]);
  const totalPR=prods.reduce((a,p)=>a+p.r,0);

  // Pagamentos
  const pays=useMemo(()=>{
    const m={};const colors=["#4f5ef0","#10b981","#f59e0b","#8b44f0","#f56565"];
    fS.forEach(s=>{const pm=s.payment_method||"Outros";m[pm]=(m[pm]||0)+s.total_price;});
    return Object.entries(m).map(([k,v],i)=>({n:k,v,c:colors[i%5]}));
  },[fS]);
  const totalPay=pays.reduce((a,p)=>a+p.v,0);

  const RANGES=[{k:"1d",l:"Hoje"},{k:"7d",l:"7 dias"},{k:"30d",l:"30 dias"},{k:"3m",l:"3 meses"},{k:"6m",l:"6 meses"},{k:"12m",l:"12 meses"},{k:"custom",l:"Personalizado"}];
  const PCOLS=["#f59e0b","#9ca3af","#cd7c2f","#4f5ef0","#10b981","#8b44f0","#f56565","#0891b2"];

  return(
    <Modal title="Relatório Gerencial · Analytics" onClose={onClose} icon="analytics" full>
      {/* Filtros */}
      <div style={{display:"flex",gap:".35rem",flexWrap:"wrap",marginBottom:"1rem",alignItems:"center"}}>
        <Ic n="cal" s={14}/>
        {RANGES.map(r=>(
          <button key={r.k} onClick={()=>setRange(r.k)} style={{padding:".28rem .65rem",borderRadius:"99px",border:`1px solid ${range===r.k?"#4f5ef0":"var(--bdr2)"}`,background:range===r.k?"#4f5ef020":"transparent",color:range===r.k?"#4f5ef0":"var(--navoff)",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>{r.l}</button>
        ))}
        {range==="custom"&&<div style={{display:"flex",gap:".4rem",alignItems:"center"}}>
          <input type="date" value={cf} onChange={e=>setCf(e.target.value)} style={{...IS,width:130,fontSize:".75rem",padding:".3rem .6rem"}}/>
          <span style={{color:"var(--tx5)",fontSize:".8rem"}}>até</span>
          <input type="date" value={ct_} onChange={e=>setCt(e.target.value)} style={{...IS,width:130,fontSize:".75rem",padding:".3rem .6rem"}}/>
        </div>}
      </div>

      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:".6rem",marginBottom:"1rem"}}>
        <KCard label="Receita" value={fmt(rev)} sub={`${fC.filter(x=>x.type==="entrada").length} lançamentos`} color="#10b981" icon="up"/>
        <KCard label="Custos" value={fmt(cost)} color="#f56565" icon="dn"/>
        <KCard label="Lucro líquido" value={fmt(profit)} sub={fmtPct(margin)+" margem"} color={profit>=0?"#4f5ef0":"#f56565"}/>
        <KCard label="Markup empresa" value={fmtPct(markup)} sub="lucro/custo" color="#8b44f0"/>
        <KCard label="Vendas" value={fmtN(fS.length)} sub={`${fmtN(units)} unidades`} color="#f59e0b" icon="sales"/>
        <KCard label="Ticket médio" value={fmt(ticket)} color="#4f5ef0"/>
        <KCard label="Valor estoque" value={fmt(stockVal)} color="#8b44f0" icon="stock"/>
        <KCard label="Clientes" value={clients.length} color="#10b981" icon="client"/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".85rem",marginBottom:".85rem"}}>
        {/* Sparkline */}
        <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem"}}>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="analytics" s={13}/>Receita · últimos 7 dias</div>
          <Spark data={spark} color="#10b981" h={55}/>
          <div style={{display:"flex",marginTop:".3rem"}}>
            {["D-6","D-5","D-4","D-3","D-2","D-1","Hoje"].map(d=><span key={d} style={{flex:1,fontSize:".58rem",color:"var(--tx6)",textAlign:"center"}}>{d}</span>)}
          </div>
        </div>
        {/* Pagamentos */}
        <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem"}}>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="dollar" s={13}/>Formas de Pagamento</div>
          {pays.length===0?<p style={{color:"var(--tx5)",fontSize:".78rem",textAlign:"center",padding:"1rem 0"}}>Sem dados</p>:(
            <div style={{display:"flex",alignItems:"center",gap:".85rem"}}>
              <Donut segs={pays} size={80}/>
              <div style={{flex:1}}>
                {pays.map(p=><div key={p.n} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".3rem"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".35rem"}}><div style={{width:8,height:8,borderRadius:"50%",background:p.c,flexShrink:0}}/><span style={{fontSize:".72rem",color:"var(--tx2)"}}>{p.n}</span></div>
                  <span style={{fontSize:".72rem",color:p.c,fontWeight:700}}>{totalPay>0?fmtPct((p.v/totalPay)*100):"—"}</span>
                </div>)}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Ranking produtos */}
      <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".85rem"}}>
        <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".85rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="award" s={13}/>Share de Produtos</div>
        {prods.length===0?<p style={{color:"var(--tx5)",fontSize:".78rem",textAlign:"center",padding:"1rem 0"}}>Sem vendas no período</p>
        :prods.slice(0,8).map((p,i)=>{
          const share=totalPR>0?(p.r/totalPR)*100:0;
          const prod=products.find(x=>x.name===p.n);
          return(
            <div key={p.n} style={{marginBottom:".7rem"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:".22rem"}}>
                <div style={{display:"flex",alignItems:"center",gap:".4rem",minWidth:0}}>
                  <span style={{fontSize:".7rem",fontWeight:700,color:PCOLS[i],flexShrink:0,fontFamily:"'Syne',sans-serif"}}>#{i+1}</span>
                  <span style={{fontSize:".78rem",color:"var(--tx2)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.n}</span>
                </div>
                <div style={{display:"flex",gap:".3rem",flexShrink:0,marginLeft:".5rem"}}>
                  <Badge color="#10b981" sm>{fmt(p.r)}</Badge>
                  <Badge color="#8b44f0" sm>{p.q}un</Badge>
                  <Badge color="#f59e0b" sm>{fmtPct(share)}</Badge>
                  {prod&&<Badge color="#4f5ef0" sm>mk {fmtPct(prod.markup)}</Badge>}
                </div>
              </div>
              <div style={{height:5,background:"var(--bdr)",borderRadius:3}}>
                <div style={{height:"100%",width:`${share}%`,background:`linear-gradient(90deg,${PCOLS[i]},${PCOLS[i]}88)`,borderRadius:3,transition:"width .6s"}}/>
              </div>
            </div>
          );
        })}
      </div>

      {/* Métricas avançadas */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:".85rem",marginBottom:".85rem"}}>
        {[
          {title:"Saúde Financeira",color:"#8b44f0",icon:"dollar",rows:[
            {l:"ROI (lucro/custo)",v:cost>0?fmtPct((profit/cost)*100):"—",c:profit>=0?"#10b981":"#f56565"},
            {l:"Margem líquida",v:fmtPct(margin),c:"#4f5ef0"},
            {l:"Markup empresa",v:fmtPct(markup),c:"#8b44f0"},
            {l:"Break-even/dia",v:range==="30d"&&cost>0?fmt(cost/30):"—",c:"#f59e0b"},
          ]},
          {title:"Estoque & Giro",color:"#f59e0b",icon:"pkg",rows:[
            {l:"Valor total estoque",v:fmt(stockVal),c:"#f59e0b"},
            {l:"Total unidades",v:fmtN(products.reduce((a,p)=>a+p.stock_qty,0))+" un",c:"var(--tx)"},
            {l:"Produtos cadastrados",v:products.length,c:"#4f5ef0"},
            {l:"Estoque zerado",v:products.filter(p=>p.stock_qty<=0).length+" item(s)",c:"#f56565"},
          ]},
          {title:"Performance Vendas",color:"#10b981",icon:"sales",rows:[
            {l:"Total de vendas",v:fS.length,c:"#10b981"},
            {l:"Unidades vendidas",v:fmtN(units)+" un",c:"var(--tx)"},
            {l:"Ticket médio",v:fmt(ticket),c:"#4f5ef0"},
            {l:"Clientes ativos",v:clients.length,c:"#8b44f0"},
          ]},
        ].map(block=>(
          <div key={block.title} style={{background:"var(--acard)",border:"1px solid var(--abdr)",borderRadius:".75rem",padding:"1rem"}}>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".78rem",color:block.color,marginBottom:".65rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n={block.icon} s={13}/>{block.title}</div>
            {block.rows.map(row=>(
              <div key={row.l} style={{display:"flex",justifyContent:"space-between",padding:".32rem 0",borderBottom:"1px solid var(--sep)"}}>
                <span style={{fontSize:".73rem",color:"var(--tx5)"}}>{row.l}</span>
                <span style={{fontSize:".76rem",fontWeight:700,color:row.c,fontFamily:"'Syne',sans-serif"}}>{row.v}</span>
              </div>
            ))}
          </div>
        ))}
      </div>

      {/* Export analytics */}
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
        <XBtn rows={[
          {Métrica:"Receita",Valor:fmt(rev)},{Métrica:"Custos",Valor:fmt(cost)},
          {Métrica:"Lucro líquido",Valor:fmt(profit)},{Métrica:"Margem",Valor:fmtPct(margin)},
          {Métrica:"Markup",Valor:fmtPct(markup)},{Métrica:"Vendas",Valor:fS.length},
          {Métrica:"Unidades",Valor:units},{Métrica:"Ticket médio",Valor:fmt(ticket)},
          ...prods.map(p=>({Métrica:`Share ${p.n}`,Valor:fmtPct(totalPR>0?(p.r/totalPR)*100:0)}))
        ]} name="analytics-caixapro" sheet="Analytics"/>
        <PBtn cols={[{k:"Métrica",l:"Métrica"},{k:"Valor",l:"Valor"}]}
          rows={[{Métrica:"Receita",Valor:fmt(rev)},{Métrica:"Custos",Valor:fmt(cost)},{Métrica:"Lucro",Valor:fmt(profit)},{Métrica:"Margem",Valor:fmtPct(margin)},{Métrica:"Markup",Valor:fmtPct(markup)},{Métrica:"Vendas",Valor:String(fS.length)},{Métrica:"Ticket",Valor:fmt(ticket)},...prods.map(p=>({Métrica:`Produto: ${p.n}`,Valor:`${fmt(p.r)} · ${p.q}un · ${fmtPct(totalPR>0?(p.r/totalPR)*100:0)} share`}))]}
          name="analytics-caixapro" title="Relatório Gerencial CaixaPro"/>
      </div>
    </Modal>
  );
}


// ─── TELA DE NOTÍCIAS ────────────────────────────────────────────────────────
// ─── APP PRINCIPAL ────────────────────────────────────────────────────────────
export default function App(){
  // Tema
  const[dark,setDark]=useState(getTheme);
  // Sincroniza classe do body com o estado — garante persistência no F5
  useEffect(()=>{
    document.body.classList.toggle("light",!dark);
  },[dark]);
  const toggle=()=>setDark(v=>{const nv=!v;setTheme(nv);return nv;});

  // Auth
  const[cu,setCU]=useState(null);
  const[logoutC,setLogoutC]=useState(false);
  useEffect(()=>{try{const s=sessionStorage.getItem("cpro:s");if(s)setCU(JSON.parse(s));}catch{};},[]);
  const login=u=>{setCU(u);try{sessionStorage.setItem("cpro:s",JSON.stringify(u));}catch{}};
  const logout=()=>{setCU(null);setLogoutC(false);try{sessionStorage.removeItem("cpro:s");}catch{}};

  // Dados
  const[products,setProds]=useState([]);
  const[sales,setSales]=useState([]);
  const[cashTx,setCash]=useState([]);
  const[clients,setClients]=useState([]);
  const[suppliers,setSupp]=useState([]);
  const[appUsers,setUsers]=useState([]);
  const[syncing,setSyncing]=useState(false);
  const[lastSync,setLastSync]=useState(null);

  // UI
  const[tab,setTab]=useState("dashboard");
  const[modal,setModal]=useState(null);
  const[editing,setEditing]=useState(null);
  const[search,setSearch]=useState("");
  const[fcat,setFcat]=useState("all");
  const[showA,setShowA]=useState(false);
  const[toast,setToast]=useState(null);
  const[pwv,setPwv]=useState({});
  // ── Novos estados v8 ──
  const[monthGoal,setMonthGoal]=useState(()=>{try{return localStorage.getItem("cpro:goal")||""}catch{return ""}});
  const[showReceipt,setShowReceipt]=useState(null);       // comprovante venda
  const[showClientHist,setShowClientHist]=useState(null); // histórico cliente
  const[showReceivables,setShowReceivables]=useState(false);
  const[receivables,setReceivables]=useState([]);
  const[stockMoves,setStockMoves]=useState([]);
  const[showImportCalc,setShowImportCalc]=useState(false);
  const[importCalc,setImportCalc]=useState({totalCost:"",qty:"",extras:""});
  // desconto no carrinho
  const[cartDiscount,setCartDiscount]=useState("");
  const[cartDiscountType,setCartDiscountType]=useState("fixed"); // "fixed" | "percent" 

  const toast$=(msg,color="#10b981")=>{setToast({msg,color});setTimeout(()=>setToast(null),3500);};

  // Forms defaults
  const FP={code:"",name:"",description:"",category:"tirzepatida",unit:"ampola",cost_per_unit:"",price_per_unit:"",units_per_pack:"1",batch:"",expiry:"",stock_qty:"0",min_stock:"5",supplier_id:""};
  // ── CART STATE ──
  const newCartItem=()=>({key:uid(),product_id:"",product_name:"",unit:"un",quantity:1,unit_price:0});
  const[cartItems,setCartItems]=useState([newCartItem()]);
  const[cartClient,setCartClient]=useState({id:"",name:""});
  const[cartPayment,setCartPayment]=useState("PIX");
  const[cartNotes,setCartNotes]=useState("");
  const[cartFreight,setCartFreight]=useState("");
  const[cartDelivery,setCartDelivery]=useState(false);   // flag entregador solicitado
  const[cartDeliveryCost,setCartDeliveryCost]=useState(""); // custo pago ao entregador
  // legado (para editSale)
  const FS={product_id:"",product_name:"",client_id:"",client_name:"",quantity:"1",unit_price:"",total_price:"",notes:"",payment_method:"PIX"};
  const FC={name:"",email:"",phone:"",notes:""};
  const FU={username:"",display_name:"",role:"operator",password:"",password2:""};
  const FSup={name:"",contact:"",phone:"",email:"",notes:""};
  const FCash={description:"",value:"",type:"saida",category:"",product_id:"",quantity:"1"};
  const FSt={product_id:"",qty:"",cost_total:"",notes:""};

  const[pf,setPf]=useState(FP);
  const[sf,setSf]=useState(FS); // usado em editSale
  const[cf,setCf]=useState(FC);
  const[uf,setUf]=useState(FU);
  const[supf,setSupf]=useState(FSup);
  const[cashF,setCashF]=useState(FCash);
  const[stF,setStF]=useState(FSt);

  // Load dados
  const load=useCallback(async()=>{
    if(!cu)return;
    setSyncing(true);
    try{
      const[a,b,c,d,e,f]=await Promise.all([
        supabase.from("products").select("*").order("name"),
        supabase.from("sales").select("*").order("created_at",{ascending:false}),
        supabase.from("cash_transactions").select("*").order("created_at",{ascending:false}),
        supabase.from("clients").select("*").order("name"),
        supabase.from("suppliers").select("*").order("name"),
        supabase.from("app_users").select("id,username,display_name,role,active,last_login,created_at").order("created_at",{ascending:false}),
        supabase.from("receivables").select("*").order("due_date",{ascending:true}),
      ]);
      if(a.data)setProds(a.data);if(b.data)setSales(b.data);if(c.data)setCash(c.data);
      if(d.data)setClients(d.data);if(e.data)setSupp(e.data);if(f.data)setUsers(f.data);
      const{data:rv}=await supabase.from("receivables").select("*").order("due_date",{ascending:true});
      if(rv)setReceivables(rv);
      setLastSync(new Date().toLocaleTimeString("pt-BR"));
    }catch(e){console.error(e);}
    setSyncing(false);
  },[cu]);

  useEffect(()=>{
    if(!cu)return;load();
    const tbls=["products","sales","cash_transactions","clients","suppliers","app_users","receivables","stock_movements"];
    const subs=tbls.map((t,i)=>supabase.channel(`ch${i}`).on("postgres_changes",{event:"*",schema:"public",table:t},load).subscribe());
    return()=>subs.forEach(s=>s.unsubscribe());
  },[cu,load]);

  // Métricas
  const cashIn=useMemo(()=>cashTx.filter(t=>t.type==="entrada").reduce((a,t)=>a+t.value,0),[cashTx]);
  const cashOut=useMemo(()=>cashTx.filter(t=>t.type==="saida").reduce((a,t)=>a+t.value,0),[cashTx]);
  const net=cashIn-cashOut;
  const margin=cashIn>0?(net/cashIn)*100:0;
  const mrkp=cashOut>0?(net/cashOut)*100:0;
  const stockVal=useMemo(()=>products.reduce((a,p)=>a+p.stock_qty*(p.cost_per_unit||0),0),[products]);
  const zeroStk=useMemo(()=>products.filter(p=>p.stock_qty<=0),[products]);
  const lowStk=useMemo(()=>products.filter(p=>p.stock_qty>0&&p.stock_qty<=(p.min_stock||5)),[products]);
  const totalSalesRev=useMemo(()=>sales.reduce((a,s)=>a+s.total_price,0),[sales]);
  const totalUnits=useMemo(()=>sales.reduce((a,s)=>a+s.quantity,0),[sales]);

  const isAdmin=cu?.role==="admin";
  const canEdit=cu?.role==="admin"||cu?.role==="operator";

  // Sale handlers
  // ── CART HELPERS ──
  const cartSetProd=(key,pid)=>{
    const p=products.find(x=>x.id===pid);
    setCartItems(items=>items.map(i=>i.key!==key?i:{...i,
      product_id:pid,
      product_name:p?p.name:"",
      unit:p?p.unit||"un":"un",
      unit_price:p?parseFloat(p.price_per_unit)||0:0,
    }));
  };
  const cartSetQty=(key,val)=>{
    const qty=Math.max(1,parseInt(val)||1);
    setCartItems(items=>items.map(i=>i.key!==key?i:{...i,quantity:qty}));
  };
  const cartSetPrice=(key,val)=>{
    setCartItems(items=>items.map(i=>i.key!==key?i:{...i,unit_price:parseFloat(val)||0}));
  };
  const cartAddLine=()=>setCartItems(items=>[...items,newCartItem()]);
  const cartRemoveLine=key=>setCartItems(items=>items.length>1?items.filter(i=>i.key!==key):items);
  const cartReset=()=>{setCartItems([newCartItem()]);setCartClient({id:"",name:""});setCartPayment("PIX");setCartNotes("");setCartFreight("");setCartDelivery(false);setCartDeliveryCost("");setCartDiscount("");setCartDiscountType("fixed");};

  const cartSubtotal=cartItems.reduce((a,i)=>a+i.unit_price*i.quantity,0);
  const cartFreightVal=parseFloat(cartFreight)||0;
  const cartDeliveryCostVal=cartDelivery?(parseFloat(cartDeliveryCost)||0):0;
  const cartDiscountVal=cartDiscountType==="percent"
    ? cartSubtotal*(parseFloat(cartDiscount)||0)/100
    : (parseFloat(cartDiscount)||0);
  const cartTotal=Math.max(0,cartSubtotal-cartDiscountVal)+cartFreightVal; // total cobrado do cliente

  // CRUD
  const registerSale=async()=>{
    // Aceita preço 0 (doação/cortesia) — filtra só por produto selecionado + quantidade
    const validItems=cartItems.filter(i=>i.product_id&&i.quantity>0);
    if(validItems.length===0){toast$("Selecione pelo menos 1 produto.","#f56565");return;}

    // Verificar estoque de todos os itens
    for(const item of validItems){
      const prod=products.find(p=>p.id===item.product_id);
      if(!prod){toast$(`Produto não encontrado: ${item.product_name}`,"#f56565");return;}
      if(prod.stock_qty<item.quantity){
        toast$(`Estoque insuficiente: ${prod.name} (disponível: ${prod.stock_qty} ${prod.unit||"un"})`,"#f56565");return;
      }
    }

    const batchId=uid(); // ID do grupo desta venda
    const clientName=cartClient.name||null;
    const clientId=cartClient.id||null;
    const freight=parseFloat(cartFreight)||0;
    const subtotal=validItems.reduce((a,i)=>a+i.unit_price*i.quantity,0);
    const discount=cartDiscountType==="percent"
      ? subtotal*(parseFloat(cartDiscount)||0)/100
      : (parseFloat(cartDiscount)||0);
    const grandTotal=Math.max(0,subtotal-discount)+freight;
    const desc=validItems.map(i=>`${i.product_name}(${i.quantity})`).join(", ");

    try{
      // 1. Inserir cada item como registro de venda
      const saleRecords=validItems.map(i=>({
        id:uid(),
        product_id:i.product_id,
        product_name:i.product_name,
        client_id:clientId,
        client_name:clientName,
        quantity:i.quantity,
        unit_price:i.unit_price,
        total_price:i.unit_price*i.quantity,
        notes:cartNotes||null,
        payment_method:cartPayment,
        discount:discount>0?discount:null,
        added_by:cu.display_name,
        date:today(),
        // batch_id para agrupar itens da mesma venda
        batch_id:batchId,
      }));

      const{error:e1}=await supabase.from("sales").insert(saleRecords);
      if(e1){toast$("Erro ao registrar venda: "+e1.message,"#f56565");return;}

      // 2. Baixar estoque de cada produto
      for(const item of validItems){
        const prod=products.find(p=>p.id===item.product_id);
        if(prod){
          await supabase.from("products")
            .update({stock_qty:prod.stock_qty-item.quantity})
            .eq("id",item.product_id);
        }
      }

      // 3. Montar lançamentos de caixa
      const cashInserts=[];

      // Separar itens pagos (preço > 0) e cortesia (preço = 0)
      const paidItems=validItems.filter(i=>i.unit_price>0);
      const freeItems=validItems.filter(i=>i.unit_price===0);

      // 3a. Entrada de receita para itens com preço
      if(grandTotal>0){
        cashInserts.push({
          id:uid(),
          description:`Venda${clientName?` · ${clientName}`:""}${freight>0?" (c/ frete)":""} · ${desc}`,
          value:grandTotal,
          type:"entrada",
          category:"Venda",
          sale_id:batchId,
          product_name:validItems.map(i=>i.product_name).join(", "),
          added_by:cu.display_name,
          date:today(),
        });
      }

      // 3b. Saída de custo para itens CORTESIA (preço = 0)
      // Representa o custo do produto doado/dado sem custo ao cliente
      for(const item of freeItems){
        const prod=products.find(p=>p.id===item.product_id);
        const custoProd=(prod?.cost_per_unit||0)*item.quantity;
        if(custoProd>0){
          cashInserts.push({
            id:uid(),
            description:`Cortesia · ${item.product_name}${clientName?` · ${clientName}`:""}`,
            value:custoProd,
            type:"saida",
            category:"Cortesia/Brinde",
            sale_id:batchId,
            product_name:item.product_name,
            added_by:cu.display_name,
            date:today(),
          });
        } else {
          // Custo zero — registra R$0 apenas para rastreio
          cashInserts.push({
            id:uid(),
            description:`Cortesia · ${item.product_name}${clientName?` · ${clientName}`:""}`,
            value:0,
            type:"saida",
            category:"Cortesia/Brinde",
            sale_id:batchId,
            product_name:item.product_name,
            added_by:cu.display_name,
            date:today(),
          });
        }
      }

      // 3c. Se não tem nenhum item pago e não tem custo de cortesia,
      //     registra ao menos 1 lançamento simbólico de R$0
      if(cashInserts.length===0){
        cashInserts.push({
          id:uid(),
          description:`Venda cortesia · ${desc}${clientName?` · ${clientName}`:""}`,
          value:0,
          type:"entrada",
          category:"Cortesia",
          sale_id:batchId,
          product_name:validItems.map(i=>i.product_name).join(", "),
          added_by:cu.display_name,
          date:today(),
        });
      }

      // 4. Se entregador solicitado e tem custo, lança saída do custo
      if(cartDelivery&&cartDeliveryCostVal>0){
        cashInserts.push({
          id:uid(),
          description:`Custo Entregador${clientName?` · ${clientName}`:""}`,
          value:cartDeliveryCostVal,
          type:"saida",
          category:"Custo Entregador",
          sale_id:batchId,
          product_name:null,
          added_by:cu.display_name,
          date:today(),
        });
      }

      // 5. Inserir todos os lançamentos de caixa
      const{error:e2}=await supabase.from("cash_transactions").insert(cashInserts);
      if(e2){toast$("Aviso: venda salva mas erro no caixa: "+e2.message,"#f59e0b");}

      const freeCount=freeItems.length;
      const paidCount=paidItems.length;
      const deliveryInfo=cartDelivery&&cartDeliveryCostVal>0?` · 🛵 Entregador: ${fmt(cartDeliveryCostVal)}`:"";
      const msg=freeCount>0
        ?`✅ Venda registrada! ${paidCount} pago(s) · ${freeCount} cortesia · Total: ${fmt(grandTotal)}${deliveryInfo}`
        :`✅ Venda registrada! ${validItems.length} produto(s) · Total: ${fmt(grandTotal)}${deliveryInfo}`;
      toast$(msg);
      setModal(null);
      cartReset();
    }catch(ex){toast$("Erro de conexão: "+ex.message,"#f56565");}
  };

  // ── RECEBÍVEIS ──
  const loadReceivables=async()=>{
    const{data}=await supabase.from("receivables").select("*").order("due_date",{ascending:true});
    if(data)setReceivables(data);
  };
  const addReceivable=async(rec)=>{
    await supabase.from("receivables").insert([{id:uid(),...rec,added_by:cu.display_name,created_at:nowISO()}]);
    loadReceivables();toast$("Conta a receber registrada!");
  };
  const payReceivable=async(id)=>{
    await supabase.from("receivables").update({paid:true,paid_date:today()}).eq("id",id);
    // lança entrada no caixa
    const rec=receivables.find(r=>r.id===id);
    if(rec){
      await supabase.from("cash_transactions").insert([{id:uid(),description:`Recebimento · ${rec.description}${rec.client_name?` · ${rec.client_name}`:""}`,value:rec.value,type:"entrada",category:"Recebimento",added_by:cu.display_name,date:today()}]);
    }
    loadReceivables();toast$("✅ Recebimento registrado no caixa!");
  };
  const deleteReceivable=async(id)=>{
    await supabase.from("receivables").delete().eq("id",id);
    loadReceivables();toast$("Removido.","#f59e0b");
  };

  // ── MOVIMENTAÇÃO ESTOQUE ──
  const loadStockMoves=async(productId)=>{
    const{data}=await supabase.from("stock_movements")
      .select("*").eq("product_id",productId)
      .order("created_at",{ascending:false}).limit(30);
    if(data)setStockMoves(data);
  };

  // ── COMPROVANTE ──
  const printReceipt=(saleGroup)=>{setShowReceipt(saleGroup);};

  const updateSale=async()=>{  const updateSale=async()=>{
    if(!editing)return;
    await supabase.from("sales").update({product_name:editing.product_name,client_name:editing.client_name||null,quantity:parseInt(editing.quantity)||1,unit_price:parseFloat(editing.unit_price)||0,total_price:(parseFloat(editing.unit_price)||0)*(parseInt(editing.quantity)||1),notes:editing.notes,payment_method:editing.payment_method}).eq("id",editing.id);
    toast$("Venda atualizada!");setModal(null);setEditing(null);
  };
  const deleteSale=async(id)=>{
    try{
      // 1. Buscar a venda clicada
      const sale=sales.find(s=>s.id===id);
      if(!sale){toast$("Venda não encontrada.","#f56565");return;}

      // 2. Encontrar TODOS os itens do mesmo batch (ou só este se não tem batch)
      const batchKey=sale.batch_id||id;
      const batchSales=sale.batch_id
        ? sales.filter(s=>s.batch_id===batchKey)
        : [sale];

      // 3. Reverter estoque de CADA item do batch
      for(const s of batchSales){
        const prod=products.find(p=>p.id===s.product_id||p.name===s.product_name);
        if(prod&&s.quantity>0){
          await supabase.from("products")
            .update({stock_qty:(prod.stock_qty||0)+s.quantity})
            .eq("id",prod.id);
        }
      }

      // 4. Excluir TODOS os registros de venda do batch
      if(sale.batch_id){
        await supabase.from("sales").delete().eq("batch_id",batchKey);
      } else {
        await supabase.from("sales").delete().eq("id",id);
      }

      // 5. Excluir TODOS os lançamentos de caixa vinculados ao batch
      // (Entrada de receita + Saída de custo de cortesia + Frete)
      await supabase.from("cash_transactions").delete().eq("sale_id",batchKey);

      // 6. Feedback detalhado
      const nItens=batchSales.length;
      const totalRevertido=batchSales.reduce((a,s)=>a+s.total_price,0);
      toast$(
        nItens>1
          ? `🔄 Venda excluída · ${nItens} itens revertidos · Estoque e caixa restaurados`
          : `🔄 Venda excluída · Estoque e caixa revertidos`,
        "#f59e0b"
      );
    }catch(ex){toast$("Erro ao excluir: "+ex.message,"#f56565");}
  };

  const registerStock=async()=>{
    if(!stF.product_id||!stF.qty){toast$("Selecione produto e quantidade.","#f56565");return;}
    const prod=products.find(p=>p.id===stF.product_id);if(!prod)return;
    const qty=parseInt(stF.qty)||0,cost=parseFloat(stF.cost_total)||0;
    if(qty<=0){toast$("Quantidade inválida.","#f56565");return;}
    await supabase.from("products").update({stock_qty:prod.stock_qty+qty}).eq("id",stF.product_id);
    if(cost>0)await supabase.from("cash_transactions").insert([{id:uid(),description:`Entrada estoque · ${prod.name} · ${qty} ${prod.unit||"un"}`,value:cost,type:"saida",category:"Compra/Estoque",product_name:prod.name,added_by:cu.display_name,date:today()}]);
    toast$(`✅ +${qty} ${prod.unit||"un"} no estoque!`);setModal(null);setStF(FSt);
  };

  const addCash=async()=>{
    if(!cashF.description||!cashF.value){toast$("Preencha descrição e valor.","#f56565");return;}
    if(cashF.product_id&&cashF.type==="saida"){
      const prod=products.find(p=>p.id===cashF.product_id);const qty=parseInt(cashF.quantity)||0;
      if(prod&&qty>0){if(prod.stock_qty<qty){toast$("Estoque insuficiente!","#f56565");return;}await supabase.from("products").update({stock_qty:prod.stock_qty-qty}).eq("id",cashF.product_id);}
    }
    try{
      const{error:e}=await supabase.from("cash_transactions").insert([{id:uid(),description:cashF.description,value:parseFloat(cashF.value),type:cashF.type,category:cashF.category||null,product_name:cashF.product_id?products.find(p=>p.id===cashF.product_id)?.name:null,added_by:cu.display_name,date:today()}]);
      if(e){toast$("Erro ao salvar: "+e.message,"#f56565");return;}
      toast$("Lançamento adicionado!");setModal(null);setCashF(FCash);
    }catch(ex){toast$("Erro de conexão.","#f56565");}
  };
  const updateCash=async()=>{
    if(!editing)return;
    await supabase.from("cash_transactions").update({description:editing.description,value:parseFloat(editing.value)||0,type:editing.type,category:editing.category||""}).eq("id",editing.id);
    toast$("Atualizado!");setModal(null);setEditing(null);
  };
  const deleteCash=async id=>{await supabase.from("cash_transactions").delete().eq("id",id);toast$("Removido.","#f59e0b");};

  const addProduct=async()=>{
    if(!pf.name||!pf.cost_per_unit||!pf.price_per_unit){toast$("Preencha nome, custo e preço.","#f56565");return;}
    const{markup,margin,profit}=calcM(pf.cost_per_unit,pf.price_per_unit);
    const sup=suppliers.find(s=>s.id===pf.supplier_id);
    try{
      const{error:e}=await supabase.from("products").insert([{id:uid(),code:pf.code||`PRD-${uid().slice(0,4).toUpperCase()}`,name:pf.name,description:pf.description||null,category:pf.category,unit:pf.unit||"un",cost_per_unit:parseFloat(pf.cost_per_unit),price_per_unit:parseFloat(pf.price_per_unit),units_per_pack:parseInt(pf.units_per_pack)||1,batch:pf.batch||null,expiry:pf.expiry||null,stock_qty:parseInt(pf.stock_qty)||0,min_stock:parseInt(pf.min_stock)||5,supplier_id:pf.supplier_id||null,supplier_name:sup?.name||null,markup,margin,profit,added_by:cu.display_name}]);
      if(e){toast$("Erro ao salvar: "+e.message,"#f56565");return;}
      toast$("Produto cadastrado!");setModal(null);setPf(FP);
    }catch(ex){toast$("Erro de conexão: "+ex.message,"#f56565");}
  };
  const saveProduct=async()=>{
    const{markup,margin,profit}=calcM(editing.cost_per_unit,editing.price_per_unit);
    const sup=suppliers.find(s=>s.id===editing.supplier_id);
    await supabase.from("products").update({code:editing.code,name:editing.name,description:editing.description,category:editing.category,unit:editing.unit,cost_per_unit:parseFloat(editing.cost_per_unit),price_per_unit:parseFloat(editing.price_per_unit),units_per_pack:parseInt(editing.units_per_pack)||1,batch:editing.batch,expiry:editing.expiry||null,stock_qty:parseInt(editing.stock_qty),min_stock:parseInt(editing.min_stock)||5,supplier_id:editing.supplier_id||null,supplier_name:sup?.name||null,markup,margin,profit}).eq("id",editing.id);
    toast$("Produto atualizado!");setModal(null);setEditing(null);
  };
  const delProduct=async id=>{await supabase.from("products").delete().eq("id",id);toast$("Removido.","#f59e0b");setModal(null);setEditing(null);};

  const addClient=async()=>{
    if(!cf.name){toast$("Nome é obrigatório.","#f56565");return;}
    try{
      const{error:e}=await supabase.from("clients").insert([{id:uid(),name:cf.name,email:cf.email||null,phone:cf.phone||null,notes:cf.notes||null,added_by:cu.display_name}]);
      if(e){toast$("Erro ao salvar: "+e.message,"#f56565");return;}
      toast$("Cliente cadastrado!");setModal(null);setCf(FC);
    }catch(ex){toast$("Erro de conexão.","#f56565");}
  };
  const saveClient=async()=>{await supabase.from("clients").update({name:editing.name,email:editing.email,phone:editing.phone,notes:editing.notes,dose:editing.dose||null,interval_days:parseInt(editing.interval_days)||7,treatment_start:editing.treatment_start||null,treatment_notes:editing.treatment_notes||null}).eq("id",editing.id);toast$("Cliente atualizado!");setModal(null);setEditing(null);};
  const delClient=async id=>{await supabase.from("clients").delete().eq("id",id);toast$("Removido.","#f59e0b");};

  const addSupp=async()=>{
    if(!supf.name){toast$("Nome obrigatório.","#f56565");return;}
    try{
      const{error:e}=await supabase.from("suppliers").insert([{id:uid(),name:supf.name,contact:supf.contact||null,phone:supf.phone||null,email:supf.email||null,notes:supf.notes||null,added_by:cu.display_name}]);
      if(e){toast$("Erro ao salvar: "+e.message,"#f56565");return;}
      toast$("Fornecedor cadastrado!");setModal(null);setSupf(FSup);
    }catch(ex){toast$("Erro de conexão.","#f56565");}
  };
  const saveSupp=async()=>{await supabase.from("suppliers").update({name:editing.name,contact:editing.contact,phone:editing.phone,email:editing.email,notes:editing.notes}).eq("id",editing.id);toast$("Fornecedor atualizado!");setModal(null);setEditing(null);};

  const addUser=async()=>{
    if(!uf.username||!uf.display_name||!uf.password)return;
    if(uf.password!==uf.password2){toast$("Senhas não coincidem.","#f56565");return;}
    if(uf.password.length<6){toast$("Mínimo 6 caracteres.","#f59e0b");return;}
    try{
      const{error:e}=await supabase.from("app_users").insert([{id:uid(),username:uf.username.trim().toLowerCase(),display_name:uf.display_name,role:uf.role,password_hash:hashPw(uf.password),active:true}]);
      if(e){toast$(e.message.includes("duplicate")?"Usuário já existe!":"Erro: "+e.message,"#f56565");return;}
      toast$("Usuário criado!");setModal(null);setUf(FU);
    }catch(ex){toast$("Erro de conexão.","#f56565");}
  };
  const saveUser=async()=>{
    const updates={display_name:editing.display_name,role:editing.role,active:editing.active};
    if(editing.new_password){if(editing.new_password.length<6){toast$("Mínimo 6 caracteres.","#f59e0b");return;}if(editing.new_password!==editing.new_password2){toast$("Senhas não coincidem.","#f56565");return;}updates.password_hash=hashPw(editing.new_password);}
    await supabase.from("app_users").update(updates).eq("id",editing.id);
    toast$("Usuário atualizado!");setModal(null);setEditing(null);
  };

  // Filtros
  const cats=["all",...new Set(products.map(p=>p.category).filter(Boolean))];
  const fProds=products.filter(p=>(fcat==="all"||p.category===fcat)&&(p.name?.toLowerCase().includes(search.toLowerCase())||p.code?.toLowerCase().includes(search.toLowerCase())));
  const fSales=sales.filter(s=>s.product_name?.toLowerCase().includes(search.toLowerCase())||s.client_name?.toLowerCase().includes(search.toLowerCase()));
  const fCash=cashTx.filter(x=>x.description?.toLowerCase().includes(search.toLowerCase())||(x.category||"").toLowerCase().includes(search.toLowerCase()));

  const nav=[
    {id:"dashboard",l:"Dashboard",n:"dashboard"},
    {id:"vendas",l:"Vendas",n:"sales"},
    {id:"estoque",l:"Estoque",n:"stock"},
    {id:"caixa",l:"Caixa",n:"cash"},
    {id:"clientes",l:"Clientes",n:"client"},
    {id:"produtos",l:"Produtos",n:"product"},
    {id:"recebiveis",l:"A Receber",n:"dollar"},
    ...(isAdmin?[{id:"usuarios",l:"Usuários",n:"users"}]:[]),
  ];

  // Helper cálculo preview
  const MPreview=({cost,price})=>{
    if(!cost||!price||parseFloat(cost)<=0||parseFloat(price)<=0)return null;
    const{markup,margin,profit}=calcM(cost,price);
    return <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem",marginBottom:".8rem",display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".4rem",textAlign:"center"}}>
      {[{l:"Markup",v:fmtPct(markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(profit),c:"#10b981"}].map(m=><div key={m.l}><div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".1rem"}}>{m.l}</div><div style={{fontSize:".88rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div></div>)}
    </div>;
  };

  // ── RENDER LOGIN ──
  if(!cu)return(<>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');@keyframes loginShake{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-6px)}80%{transform:translateX(6px)}}`}</style>
    <LoginScreen onLogin={login} dark={dark}/>
  </>);

  // ── RENDER APP ──
  return(<>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}input[type=date]::-webkit-calendar-picker-indicator{filter:${dark?"invert(.5)":"none"}}`}</style>

    {/* TOAST */}
    {toast&&<div style={{position:"fixed",bottom:"1.25rem",right:"1.25rem",zIndex:999,background:"var(--card)",border:`1px solid ${toast.color}50`,borderRadius:".65rem",padding:".65rem 1.1rem",display:"flex",alignItems:"center",gap:".45rem",fontSize:".82rem",color:toast.color,fontFamily:"'DM Sans',sans-serif",boxShadow:"0 8px 30px var(--shadow)",animation:"fadeUp .3s ease",maxWidth:320}}><Ic n="save" s={14}/>{toast.msg}</div>}

    <div style={{minHeight:"100vh",background:"var(--bg)",fontFamily:"'DM Sans',sans-serif",color:"var(--tx)"}}>

      {/* ── TOP BAR ── */}
      <div style={{background:"var(--bg2)",borderBottom:"1px solid var(--bdr)",padding:".6rem 1rem",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:50,transition:"background .3s"}}>
        <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
          <div style={{width:24,height:24,borderRadius:".4rem",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",color:"#fff"}}><Ic n="syringe" s={14}/></div>
          <div>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".85rem",lineHeight:1,color:"var(--tx)"}}>CaixaPro</div>
            <div style={{fontSize:".52rem",color:"var(--tx5)",lineHeight:1.2}}>Tirzepatida v6.0</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:".6rem"}}>
          {(zeroStk.length>0||lowStk.length>0)&&<div style={{fontSize:".62rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".18rem .5rem",border:"1px solid #f59e0b30",display:"flex",alignItems:"center",gap:".25rem"}}><Ic n="warn" s={10}/>{zeroStk.length>0?`${zeroStk.length} zerado(s)`:`${lowStk.length} baixo`}</div>}
          <div style={{fontSize:".62rem",color:syncing?"#8b44f0":"var(--tx6)",display:"flex",alignItems:"center",gap:".25rem"}}><Ic n="sync" s={10}/>{syncing?"...":lastSync||"—"}</div>
          <button onClick={toggle} title={dark?"Modo claro":"Modo escuro"} style={{background:"none",border:"none",padding:".2rem",borderRadius:".35rem",color:dark?"#f59e0b":"#4f5ef0",display:"flex",alignItems:"center"}}>
            <Ic n={dark?"sun":"moon"} s={16}/>
          </button>
          <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
            <div style={{width:22,height:22,borderRadius:"50%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:".65rem",fontWeight:700,color:"#fff"}}>{cu.display_name.charAt(0).toUpperCase()}</div>
            <span style={{fontSize:".72rem",color:"var(--tx4)",display:"none"}}>{cu.display_name}</span>
            <button onClick={()=>setLogoutC(true)} style={{background:"none",border:"none",color:"var(--tx5)",display:"flex"}}><Ic n="logout" s={13}/></button>
          </div>
        </div>
      </div>

      {/* ── NAV ── */}
      <div style={{background:"var(--bg2)",borderBottom:"1px solid var(--bdr)",padding:"0 .4rem",display:"flex",overflowX:"auto",transition:"background .3s"}}>
        {nav.map(item=>(
          <button key={item.id} onClick={()=>{setTab(item.id);setSearch("");setFcat("all");}} style={{display:"flex",alignItems:"center",gap:".28rem",padding:".58rem .7rem",background:"none",border:"none",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,color:tab===item.id?"var(--navon)":"var(--navoff)",borderBottom:tab===item.id?"2px solid #4f5ef0":"2px solid transparent",transition:"all .2s",whiteSpace:"nowrap"}}>
            <Ic n={item.n} s={12}/>{item.l}
            {item.id==="estoque"&&zeroStk.length>0&&<span style={{background:"#f56565",color:"#fff",borderRadius:"99px",fontSize:".55rem",fontWeight:700,padding:"0 .28rem",lineHeight:"1.5"}}>{zeroStk.length}</span>}
          </button>
        ))}
      </div>

      {/* ── CONTEÚDO ── */}
      <div style={{maxWidth:960,margin:"0 auto",padding:".9rem .8rem"}}>

        {/* ══ DASHBOARD ══ */}
        {tab==="dashboard"&&(<>
          <div style={{display:"flex",justifyContent:"flex-end",marginBottom:".75rem"}}>
            <Btn v="info" onClick={()=>setShowA(true)}><Ic n="analytics" s={14}/> Relatório Gerencial</Btn>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Receita" value={fmt(cashIn)} color="#10b981" icon="up"/>
            <KCard label="Custos" value={fmt(cashOut)} color="#f56565" icon="dn"/>
            <KCard label="Lucro" value={fmt(net)} sub={fmtPct(margin)+" margem"} color={net>=0?"#4f5ef0":"#f56565"}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Vendas" value={fmtN(sales.length)} sub={`${fmtN(totalUnits)} un`} color="#8b44f0" icon="sales"/>
            <KCard label="Estoque" value={products.length} sub={zeroStk.length>0?`${zeroStk.length} zerado(s)`:lowStk.length>0?`${lowStk.length} baixo(s)`:"ok"} color={zeroStk.length>0?"#f56565":lowStk.length>0?"#f59e0b":"#10b981"} icon="stock"/>
            <KCard label="Markup empresa" value={fmtPct(mrkp)} sub="lucro/custo" color="#f59e0b"/>
          </div>
          {(zeroStk.length>0||lowStk.length>0)&&(
            <div style={{background:"var(--card)",border:"1px solid #f59e0b30",borderRadius:".75rem",padding:".8rem 1rem",marginBottom:".7rem"}}>
              <div style={{display:"flex",alignItems:"center",gap:".35rem",color:"#f59e0b",fontWeight:700,fontSize:".78rem",marginBottom:".35rem"}}><Ic n="warn" s={13}/>Alertas de Estoque</div>
              {zeroStk.length>0&&<p style={{fontSize:".73rem",color:"#f56565",marginBottom:".18rem"}}>🔴 Zerados: {zeroStk.map(p=>p.name).join(", ")}</p>}
              {lowStk.length>0&&<p style={{fontSize:".73rem",color:"#f59e0b"}}>🟡 Baixo: {lowStk.map(p=>`${p.name}(${p.stock_qty})`).join(", ")}</p>}
            </div>
          )}

          {/* Meta mensal */}
          {(()=>{
            const goal=parseFloat(monthGoal)||0;
            const now=new Date();
            const startM=new Date(now.getFullYear(),now.getMonth(),1);
            const monthRev=cashTx.filter(t=>t.type==="entrada"&&new Date(t.created_at)>=startM).reduce((a,t)=>a+t.value,0);
            const pct=goal>0?Math.min((monthRev/goal)*100,100):0;
            const overdue=receivables.filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)<new Date());
            const overdueVal=overdue.reduce((a,r)=>a+r.value,0);
            return(<>
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".9rem 1rem",marginBottom:".75rem"}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".5rem"}}>
                  <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",display:"flex",alignItems:"center",gap:".3rem"}}>
                    🎯 Meta Mensal
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:".5rem"}}>
                    <span style={{fontSize:".75rem",color:"#10b981",fontWeight:700}}>{fmt(monthRev)}</span>
                    <span style={{fontSize:".7rem",color:"var(--tx5)"}}>de</span>
                    <input
                      type="number" min="0" step="100"
                      value={monthGoal}
                      onChange={e=>{setMonthGoal(e.target.value);try{localStorage.setItem("cpro:goal",e.target.value)}catch{}}}
                      placeholder="definir meta..."
                      style={{width:110,background:"var(--inp)",border:"1px solid var(--bdr2)",borderRadius:".35rem",padding:".25rem .5rem",color:"var(--tx)",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",outline:"none"}}
                    />
                  </div>
                </div>
                <div style={{height:8,background:"var(--bdr)",borderRadius:4}}>
                  <div style={{height:"100%",width:`${pct}%`,background:pct>=100?"linear-gradient(90deg,#10b981,#059669)":pct>=70?"linear-gradient(90deg,#f59e0b,#d97706)":"linear-gradient(90deg,#4f5ef0,#8b44f0)",borderRadius:4,transition:"width .6s"}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",marginTop:".3rem",fontSize:".65rem",color:"var(--sub)"}}>
                  <span>{fmtPct(pct)} atingido</span>
                  {goal>0&&<span style={{color:pct>=100?"#10b981":"var(--sub)"}}>faltam {fmt(Math.max(0,goal-monthRev))}</span>}
                </div>
              </div>
              {overdue.length>0&&(
                <div onClick={()=>setShowReceivables(true)} style={{background:"var(--card)",border:"1px solid #f5656540",borderRadius:".75rem",padding:".75rem 1rem",marginBottom:".75rem",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".5rem",color:"#f56565"}}>
                    <Ic n="warn" s={14}/>
                    <span style={{fontWeight:700,fontSize:".8rem"}}>{overdue.length} conta{overdue.length>1?"s":""} vencida{overdue.length>1?"s":""}</span>
                  </div>
                  <span style={{fontWeight:700,color:"#f56565",fontFamily:"'Syne',sans-serif",fontSize:".88rem"}}>{fmt(overdueVal)}</span>
                </div>
              )}
            </>);
          })()}

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".75rem"}}>
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".9rem"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".65rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",display:"flex",alignItems:"center",gap:".3rem"}}><Ic n="sales" s={12}/>Últimas Vendas</div>
                {canEdit&&<Btn sm onClick={()=>setModal("sale")}><Ic n="plus" s={11}/>Nova</Btn>}
              </div>
              {sales.length===0?<p style={{color:"var(--tx5)",fontSize:".78rem",textAlign:"center",padding:"1.25rem 0"}}>Nenhuma venda.</p>
              :sales.slice(0,5).map(s=>(
                <div key={s.id} style={{padding:".4rem 0",borderBottom:"1px solid var(--sep)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{minWidth:0}}><div style={{fontSize:".76rem",color:"var(--tx2)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.product_name}</div><div style={{fontSize:".63rem",color:"var(--tx5)"}}>{s.date} · {s.quantity}un</div></div>
                  <span style={{color:"#10b981",fontWeight:700,fontFamily:"'Syne',sans-serif",fontSize:".78rem",flexShrink:0,marginLeft:".5rem"}}>{fmt(s.total_price)}</span>
                </div>
              ))}
            </div>
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".9rem"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".65rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",display:"flex",alignItems:"center",gap:".3rem"}}><Ic n="stock" s={12}/>Estoque</div>
                {canEdit&&<Btn sm v="info" onClick={()=>setModal("stockEntry")}><Ic n="arrup" s={11}/>Entrada</Btn>}
              </div>
              {products.slice(0,6).map(p=>(
                <div key={p.id} style={{padding:".4rem 0",borderBottom:"1px solid var(--sep)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{fontSize:".76rem",color:"var(--tx2)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{CATS[p.category]||"📋"} {p.name}</div>
                  <span style={{color:stkColor(p.stock_qty),fontWeight:700,fontSize:".78rem",flexShrink:0,marginLeft:".5rem"}}>{p.stock_qty} {p.unit||"un"}</span>
                </div>
              ))}
              <div style={{marginTop:".55rem",paddingTop:".55rem",borderTop:"1px solid var(--bdr)",display:"flex",justifyContent:"space-between",fontSize:".68rem"}}>
                <span style={{color:"var(--tx5)"}}>Valor total estoque</span>
                <span style={{color:"#f59e0b",fontWeight:700}}>{fmt(stockVal)}</span>
              </div>
            </div>
          </div>
        </>)}

        {/* ══ VENDAS ══ */}
        {tab==="vendas"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Vendas</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={fSales.map(s=>({Data:s.date,Produto:s.product_name,Cliente:s.client_name||"—",Qtd:s.quantity,Preço:fmt(s.unit_price),Total:fmt(s.total_price),Pagamento:s.payment_method,Obs:s.notes||""}))} name="vendas-caixapro" sheet="Vendas"/>
              <PBtn cols={[{k:"Data",l:"Data"},{k:"Produto",l:"Produto"},{k:"Cliente",l:"Cliente"},{k:"Qtd",l:"Qtd"},{k:"Total",l:"Total"},{k:"Pagamento",l:"Pagamento"}]} rows={fSales.map(s=>({Data:s.date,Produto:s.product_name,Cliente:s.client_name||"—",Qtd:s.quantity,Total:fmt(s.total_price),Pagamento:s.payment_method}))} name="vendas-caixapro" title="Relatório de Vendas"/>
              {canEdit&&<Btn sm onClick={()=>setModal("sale")}><Ic n="plus" s={12}/>Nova</Btn>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Total" value={fmtN(sales.length)} color="#4f5ef0"/>
            <KCard label="Receita" value={fmt(totalSalesRev)} color="#10b981"/>
            <KCard label="Unidades" value={fmtN(totalUnits)} color="#8b44f0"/>
            <KCard label="Ticket médio" value={sales.length>0?fmt(totalSalesRev/sales.length):"—"} color="#f59e0b"/>
          </div>
          <div style={{position:"relative",marginBottom:".65rem"}}>
            <div style={{position:"absolute",left:".75rem",top:"50%",transform:"translateY(-50%)",color:"var(--tx5)"}}><Ic n="search" s={14}/></div>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar produto ou cliente..." style={{...IS,paddingLeft:"2.1rem"}}/>
          </div>
          <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem"}}>
            {fSales.length===0?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2.5rem 0",fontSize:".8rem"}}>Nenhuma venda.</p>
            :fSales.map(s=>(
              <div key={s.id} style={{padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:".35rem",flexWrap:"wrap",marginBottom:".18rem"}}>
                    <span style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)"}}>{s.product_name}</span>
                    <Badge color="#8b44f0" sm>{s.quantity} un</Badge>
                    {s.unit_price===0&&<Badge color="#f59e0b" sm>Cortesia</Badge>}
                    {s.batch_id&&<Badge color="#44475a" sm>Lote</Badge>}
                    <Badge color="#0891b2" sm>{s.payment_method}</Badge>
                  </div>
                  <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{s.date}{s.client_name&&` · 👤 ${s.client_name}`}{s.notes&&` · ${s.notes}`}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:".5rem",flexShrink:0,marginLeft:".75rem"}}>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:700,color:"#10b981",fontSize:".88rem",fontFamily:"'Syne',sans-serif"}}>{fmt(s.total_price)}</div>
                    <div style={{fontSize:".63rem",color:"var(--tx5)"}}>{fmt(s.unit_price)}/un</div>
                  </div>
                  {canEdit&&<div style={{display:"flex",gap:".2rem"}}>
                    <button onClick={()=>{
                      const batch=s.batch_id?sales.filter(x=>x.batch_id===s.batch_id):[s];
                      setShowReceipt(batch);
                    }} style={{background:"none",border:"none",color:"#8b44f0",padding:".2rem"}} title="Comprovante"><Ic n="pdf" s={13}/></button>
                    {s.client_name&&<button onClick={()=>{
                      const msg=`Olá ${s.client_name}! ✅ Sua compra de *${s.product_name}* foi registrada.%0AValor: *${fmt(s.total_price)}*%0APagamento: ${s.payment_method}%0AData: ${s.date}%0AObrigado! 🙏`;
                      window.open(`https://wa.me/?text=${msg}`,"_blank");
                    }} style={{background:"none",border:"none",color:"#25d366",padding:".2rem"}} title="WhatsApp"><span style={{fontSize:13}}>📱</span></button>}
                    <button onClick={()=>{setEditing({...s,quantity:String(s.quantity),unit_price:String(s.unit_price)});setModal("editSale");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}} title="Editar"><Ic n="edit" s={13}/></button>
                    {isAdmin&&<button onClick={()=>deleteSale(s.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}} title="Excluir"><Ic n="trash" s={13}/></button>}
                  </div>}
                </div>
              </div>
            ))}
          </div>
        </>)}

        {/* ══ ESTOQUE ══ */}
        {tab==="estoque"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Estoque</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={fProds.map(p=>({Código:p.code,Nome:p.name,Cat:p.category,Unidade:p.unit||"un",Estoque:p.stock_qty,"Est.Mín":p.min_stock||5,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Margem:fmtPct(p.margin),Fornecedor:p.supplier_name||"—"}))} name="estoque-caixapro" sheet="Estoque"/>
              <PBtn cols={[{k:"Nome",l:"Produto"},{k:"Cat",l:"Cat."},{k:"Estoque",l:"Estoque"},{k:"Custo",l:"Custo"},{k:"Preço",l:"Preço"},{k:"Markup",l:"Markup"},{k:"Fornecedor",l:"Fornecedor"}]} rows={fProds.map(p=>({Nome:p.name,Cat:p.category,Estoque:`${p.stock_qty} ${p.unit||"un"}`,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Fornecedor:p.supplier_name||"—"}))} name="estoque-caixapro" title="Relatório de Estoque"/>
              {canEdit&&<><Btn v="info" sm onClick={()=>setModal("stockEntry")}><Ic n="arrup" s={12}/>Entrada</Btn><Btn sm onClick={()=>setModal("produto")}><Ic n="plus" s={12}/>Produto</Btn></>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Itens" value={products.length} color="#4f5ef0"/>
            <KCard label="Unidades" value={fmtN(products.reduce((a,p)=>a+p.stock_qty,0))} color="#8b44f0"/>
            <KCard label="Valor estoque" value={fmt(stockVal)} color="#f59e0b"/>
            {zeroStk.length>0&&<KCard label="Zerados" value={zeroStk.length} color="#f56565"/>}
          </div>
          {/* Alerta vencimentos */}
          {(()=>{
            const soon=products.filter(p=>{const d=daysUntil(p.expiry);return d!==null&&d<=30;}).sort((a,b)=>daysUntil(a.expiry)-daysUntil(b.expiry));
            if(!soon.length)return null;
            return <div style={{background:"var(--card)",border:"1px solid #f59e0b40",borderRadius:".65rem",padding:".75rem 1rem",marginBottom:".65rem"}}>
              <div style={{fontWeight:700,fontSize:".78rem",color:"#f59e0b",marginBottom:".4rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="warn" s={13}/>⏰ Vencimentos nos próximos 30 dias</div>
              <div style={{display:"flex",gap:".45rem",flexWrap:"wrap"}}>
                {soon.map(p=>{const d=daysUntil(p.expiry);return <span key={p.id} style={{fontSize:".7rem",color:d<0?"#f56565":d<=7?"#f59e0b":"var(--tx3)",background:d<0?"#f5656515":"#f59e0b15",borderRadius:"99px",padding:".15rem .55rem",border:`1px solid ${d<0?"#f5656530":"#f59e0b30"}`}}>{p.name} · {d<0?"Vencido":d===0?"Hoje":d+"d"}</span>;})}
              </div>
            </div>;
          })()}
          <div style={{display:"flex",gap:".35rem",marginBottom:".65rem",overflowX:"auto"}}>
            {cats.map(c=><button key={c} onClick={()=>setFcat(c)} style={{padding:".28rem .65rem",borderRadius:"99px",border:`1px solid ${fcat===c?"#4f5ef0":"var(--bdr2)"}`,background:fcat===c?"#4f5ef020":"transparent",color:fcat===c?"#4f5ef0":"var(--navoff)",fontSize:".7rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,whiteSpace:"nowrap"}}>{c==="all"?"Todos":`${CATS[c]||"📋"} ${c}`}</button>)}
          </div>
          <div style={{display:"grid",gap:".6rem"}}>
            {fProds.map(p=>{
              const days=p.expiry?daysUntil(p.expiry):null;
              return(
                <div key={p.id} style={{background:"var(--card)",border:`1px solid ${p.stock_qty<=0?"#f5656530":p.stock_qty<=(p.min_stock||5)?"#f59e0b30":"var(--bdr)"}`,borderRadius:".75rem",padding:".85rem 1rem"}}>
                  <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".55rem"}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:".38rem",flexWrap:"wrap",marginBottom:".2rem"}}>
                        <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".88rem",color:"var(--tx)"}}>{CATS[p.category]||"📋"} {p.name}</span>
                        <Badge color={stkColor(p.stock_qty)} sm>{p.stock_qty<=0?"Zerado":p.stock_qty<=(p.min_stock||5)?"Baixo":"OK"}</Badge>
                        {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                        {days!==null&&days<=30&&<Badge color={expColor(days)} sm>{days<0?"Vencido":`Vcto ${days}d`}</Badge>}
                      </div>
                      <div style={{fontSize:".65rem",color:"var(--tx5)"}}>{p.code}{p.batch&&` · Lote: ${p.batch}`}</div>
                    </div>
                    <div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".65rem"}}>
                      {canEdit&&<><button onClick={()=>{setStF({product_id:p.id,qty:"",cost_total:"",notes:""});setModal("stockEntry");}} style={{background:"#0e1e0e",border:"1px solid #10b98130",borderRadius:".4rem",padding:".28rem .5rem",color:"#10b981",display:"flex",alignItems:"center"}}><Ic n="arrup" s={11}/></button>
                      <button onClick={()=>{loadStockMoves(p.id);setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5),showMoves:true});setModal("editProd");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}} title="Editar / Histórico"><Ic n="edit" s={13}/></button>
                      {isAdmin&&<button onClick={()=>delProduct(p.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}</>}
                    </div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(70px,1fr))",gap:".38rem"}}>
                    {[{l:"Estoque",v:`${p.stock_qty} ${p.unit||"un"}`,c:stkColor(p.stock_qty)},{l:"Mínimo",v:String(p.min_stock||5),c:"var(--sub)"},{l:"Custo",v:fmt(p.cost_per_unit),c:"var(--tx3)"},{l:"Preço",v:fmt(p.price_per_unit),c:"#10b981"},{l:"Markup",v:fmtPct(p.markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(p.margin),c:"#8b44f0"}].map(m=><Pill key={m.l} label={m.l} value={m.v} color={m.c}/>)}
                  </div>
                </div>
              );
            })}
          </div>
        </>)}

        {/* ══ CAIXA ══ */}
        {tab==="caixa"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Caixa</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={fCash.map(x=>({Data:x.date,Descrição:x.description,Tipo:x.type==="entrada"?"Entrada":"Saída",Valor:fmt(x.value),Categoria:x.category||"",Produto:x.product_name||""}))} name="caixa-caixapro" sheet="Caixa"/>
              <PBtn cols={[{k:"Data",l:"Data"},{k:"Descrição",l:"Descrição"},{k:"Tipo",l:"Tipo"},{k:"Valor",l:"Valor"},{k:"Categoria",l:"Categoria"}]} rows={fCash.map(x=>({Data:x.date,Descrição:x.description,Tipo:x.type==="entrada"?"Entrada":"Saída",Valor:fmt(x.value),Categoria:x.category||""}))} name="caixa-caixapro" title="Relatório de Caixa"/>
              {canEdit&&<Btn sm onClick={()=>setModal("cashTx")}><Ic n="plus" s={12}/>Lançamento</Btn>}
            </div>
          </div>
          {/* Resumo financeiro */}
          <div style={{background:"var(--acard)",border:"1px solid var(--abdr)",borderRadius:".75rem",padding:".9rem",marginBottom:".75rem"}}>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".78rem",color:"#8b44f0",marginBottom:".65rem"}}>📊 Resumo Financeiro</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem",marginBottom:".5rem"}}>
              {[{l:"Receita (100%)",v:fmt(cashIn),c:"#10b981",s:fmtPct(100)},{l:"Custos",v:fmt(cashOut),c:"#f56565",s:cashIn>0?fmtPct((cashOut/cashIn)*100)+" da receita":""},{l:"Lucro líquido",v:fmt(net),c:net>=0?"#4f5ef0":"#f56565",s:fmtPct(margin)+" margem"}].map(m=>(
                <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".55rem",border:"1px solid var(--bdr)"}}>
                  <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div>
                  <div style={{fontSize:".9rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                  {m.s&&<div style={{fontSize:".6rem",color:m.c,opacity:.7}}>{m.s}</div>}
                </div>
              ))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem"}}>
              {[{l:"Markup empresa",v:fmtPct(mrkp),c:"#f59e0b"},{l:"Receita vendas",v:fmt(totalSalesRev),c:"#10b981"},{l:"Valor estoque",v:fmt(stockVal),c:"#8b44f0"}].map(m=>(
                <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".5rem",border:"1px solid var(--bdr)"}}>
                  <div style={{fontSize:".57rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".12rem"}}>{m.l}</div>
                  <div style={{fontSize:".82rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={{position:"relative",marginBottom:".65rem"}}>
            <div style={{position:"absolute",left:".75rem",top:"50%",transform:"translateY(-50%)",color:"var(--tx5)"}}><Ic n="search" s={14}/></div>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar lançamento..." style={{...IS,paddingLeft:"2.1rem"}}/>
          </div>
          <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem"}}>
            {fCash.length===0?<p style={{color:"var(--tx5)",fontSize:".8rem",textAlign:"center",padding:"2.5rem 0"}}>Nenhum lançamento.</p>
            :fCash.map(x=>(
              <div key={x.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:".83rem",color:"var(--tx)",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{x.description}</div>
                  <div style={{fontSize:".66rem",color:"var(--tx5)",marginTop:".1rem"}}>{x.date}{x.category&&` · ${x.category}`}{x.added_by&&` · ${x.added_by}`}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:".5rem",flexShrink:0,marginLeft:".75rem"}}>
                  <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:x.type==="entrada"?"#10b981":"#f56565",fontSize:".88rem"}}>{x.type==="entrada"?"+":"-"}{fmt(x.value)}</span>
                  {canEdit&&!x.sale_id&&<div style={{display:"flex",gap:".2rem"}}>
                    <button onClick={()=>{setEditing({...x,value:String(x.value)});setModal("editCash");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>
                    {isAdmin&&<button onClick={()=>deleteCash(x.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}
                  </div>}
                </div>
              </div>
            ))}
          </div>
        </>)}

        {/* ══ CLIENTES ══ */}
        {tab==="clientes"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Clientes</h2>
            <div style={{display:"flex",gap:".4rem",alignItems:"center"}}>
              <XBtn rows={clients.map(c=>({Nome:c.name,Telefone:c.phone||"",Email:c.email||"",Obs:c.notes||"",Compras:sales.filter(s=>s.client_id===c.id||s.client_name===c.name).length,Total:fmt(sales.filter(s=>s.client_id===c.id||s.client_name===c.name).reduce((a,s)=>a+s.total_price,0))}))} name="clientes-caixapro" sheet="Clientes"/>
              <PBtn cols={[{k:"Nome",l:"Nome"},{k:"Telefone",l:"Telefone"},{k:"Email",l:"Email"},{k:"Compras",l:"Compras"},{k:"Total",l:"Total"}]} rows={clients.map(c=>({Nome:c.name,Telefone:c.phone||"",Email:c.email||"",Compras:sales.filter(s=>s.client_id===c.id||s.client_name===c.name).length,Total:fmt(sales.filter(s=>s.client_id===c.id||s.client_name===c.name).reduce((a,s)=>a+s.total_price,0))}))} name="clientes-caixapro" title="Relatório de Clientes"/>
              {canEdit&&<Btn sm onClick={()=>setModal("cliente")}><Ic n="plus" s={12}/>Novo</Btn>}
            </div>
          </div>
          {clients.length===0?<p style={{color:"var(--tx5)",textAlign:"center",padding:"3rem 0",fontSize:".8rem"}}>Nenhum cliente cadastrado.</p>
          :<div style={{display:"grid",gap:".55rem"}}>
            {clients.map(c=>{
              const cs=sales.filter(s=>s.client_id===c.id||s.client_name===c.name);
              return(
                <div key={c.id} style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".78rem 1rem",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".7rem"}}>
                    <div style={{width:32,height:32,borderRadius:"50%",background:"linear-gradient(135deg,#4f5ef0,#10b981)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".82rem",color:"#fff",flexShrink:0}}>{c.name.charAt(0).toUpperCase()}</div>
                    <div>
                      <div style={{fontWeight:700,fontSize:".85rem",color:"var(--tx)"}}>{c.name}</div>
                      <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{[c.phone,c.email].filter(Boolean).join(" · ")||"Sem contato"}</div>
                      {c.notes&&<div style={{fontSize:".65rem",color:"#8b44f0"}}>📝 {c.notes}</div>}
                      {c.dose&&(()=>{
                        const daysSince=c.treatment_start?Math.floor((new Date()-new Date(c.treatment_start))/86400000):null;
                        const interval=parseInt(c.interval_days)||7;
                        const nextDose=daysSince!==null?interval-((daysSince)%interval):null;
                        return <div style={{fontSize:".64rem",color:nextDose!==null&&nextDose<=2?"#f59e0b":"#10b981",marginTop:".1rem",display:"flex",alignItems:"center",gap:".3rem"}}>
                          💉 Dose: {c.dose} · Intervalo: {interval}d
                          {nextDose!==null&&<span style={{fontWeight:700}}>{nextDose<=0?"🔔 Próxima dose HOJE!":nextDose<=2?`⚠️ Próxima dose em ${nextDose}d`:`Próxima em ${nextDose}d`}</span>}
                        </div>;
                      })()}
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:".6rem"}}>
                    {cs.length>0&&<div style={{textAlign:"right"}}>
                      <div style={{fontWeight:700,color:"#10b981",fontSize:".82rem",fontFamily:"'Syne',sans-serif"}}>{fmt(cs.reduce((a,s)=>a+s.total_price,0))}</div>
                      <div style={{fontSize:".62rem",color:"var(--tx5)"}}>{cs.length} compra{cs.length!==1?"s":""}</div>
                    </div>}
                    <div style={{display:"flex",gap:".3rem"}}>
                      <button onClick={()=>setShowClientHist(c)} style={{background:"none",border:"none",color:"#8b44f0",padding:".2rem"}} title="Histórico"><Ic n="analytics" s={13}/></button>
                      {c.phone&&<button onClick={()=>{const msg=`Olá ${c.name}! Como está o seu tratamento? Estamos à disposição. 💉`;window.open(`https://wa.me/55${c.phone.replace(/\D/g,"")}?text=${encodeURIComponent(msg)}`,"_blank");}} style={{background:"none",border:"none",color:"#25d366",padding:".2rem"}} title="WhatsApp"><span style={{fontSize:13}}>📱</span></button>}
                      {canEdit&&<><button onClick={()=>{setEditing({...c});setModal("editCliente");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>
                      {isAdmin&&<button onClick={()=>delClient(c.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}</>}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>}
        </>)}

        {/* ══ PRODUTOS ══ */}
        {tab==="produtos"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Produtos</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={products.map(p=>({Código:p.code,Nome:p.name,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Margem:fmtPct(p.margin),Estoque:p.stock_qty,Fornecedor:p.supplier_name||"—"}))} name="produtos-caixapro" sheet="Produtos"/>
              <PBtn cols={[{k:"Nome",l:"Produto"},{k:"Custo",l:"Custo"},{k:"Preço",l:"Preço"},{k:"Markup",l:"Markup"},{k:"Estoque",l:"Estoque"},{k:"Fornecedor",l:"Fornecedor"}]} rows={products.map(p=>({Nome:p.name,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Estoque:`${p.stock_qty} ${p.unit||"un"}`,Fornecedor:p.supplier_name||"—"}))} name="produtos-caixapro" title="Relatório de Produtos"/>
              <Btn sm v="ghost" onClick={()=>setModal("fornecedores")}><Ic n="supplier" s={12}/>Fornecedores</Btn>
              {canEdit&&<Btn sm onClick={()=>setModal("produto")}><Ic n="plus" s={12}/>Produto</Btn>}
            </div>
          </div>
          <div style={{display:"grid",gap:".6rem"}}>
            {products.map(p=>(
              <div key={p.id} style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".82rem 1rem"}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".55rem"}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:".38rem",flexWrap:"wrap",marginBottom:".18rem"}}>
                      <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",color:"var(--tx)"}}>{CATS[p.category]||"📋"} {p.name}</span>
                      {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                    </div>
                    <div style={{fontSize:".65rem",color:"var(--tx5)"}}>Cód: {p.code}{p.batch&&` · Lote: ${p.batch}`} · {p.unit||"un"}</div>
                    {p.description&&<div style={{fontSize:".68rem",color:"var(--tx4)",marginTop:".12rem"}}>{p.description}</div>}
                  </div>
                  {canEdit&&<div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".65rem"}}>
                    <button onClick={()=>{setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5),units_per_pack:String(p.units_per_pack||1)});setModal("editProd");}} style={{background:"none",border:"none",color:"#4f5ef0"}}><Ic n="edit" s={13}/></button>
                    {isAdmin&&<button onClick={()=>delProduct(p.id)} style={{background:"none",border:"none",color:"var(--tx6)"}}><Ic n="trash" s={13}/></button>}
                  </div>}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(78px,1fr))",gap:".38rem"}}>
                  {[{l:"Estoque",v:`${p.stock_qty} ${p.unit||"un"}`,c:stkColor(p.stock_qty)},{l:"Custo/un",v:fmt(p.cost_per_unit),c:"var(--tx3)"},{l:"Preço/un",v:fmt(p.price_per_unit),c:"#10b981"},{l:"Markup",v:fmtPct(p.markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(p.margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(p.profit),c:"#10b981"}].map(m=><Pill key={m.l} label={m.l} value={m.v} color={m.c}/>)}
                </div>
              </div>
            ))}
          </div>
        </>)}

        {/* ══ USUÁRIOS ══ */}
        {/* ══ A RECEBER ══ */}
        {tab==="recebiveis"&&(()=>{
          const pendentes=receivables.filter(r=>!r.paid);
          const vencidos=pendentes.filter(r=>r.due_date&&new Date(r.due_date)<new Date());
          const totalPend=pendentes.reduce((a,r)=>a+r.value,0);
          const[recForm,setRecForm]=React.useState({client_id:"",client_name:"",description:"",value:"",due_date:"",notes:""});
          return(<>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>💰 Contas a Receber</h2>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:".75rem"}}>
              <KCard label="Em aberto" value={fmtN(pendentes.length)} sub={fmt(totalPend)} color="#4f5ef0"/>
              <KCard label="Vencidas" value={fmtN(vencidos.length)} sub={fmt(vencidos.reduce((a,r)=>a+r.value,0))} color="#f56565"/>
              <KCard label="Recebido (mês)" value={fmt(receivables.filter(r=>r.paid&&r.paid_date===today()).reduce((a,r)=>a+r.value,0))} color="#10b981"/>
            </div>
            {/* Form nova conta */}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>➕ Registrar nova conta a receber</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".5rem",marginBottom:".5rem"}}>
                <Field label="Cliente" t={{sub:"var(--sub)",text5:"var(--tx5)"}}>
                  <select value={recForm.client_id} onChange={e=>{const c=clients.find(x=>x.id===e.target.value);setRecForm(f=>({...f,client_id:e.target.value,client_name:c?.name||""}));}} style={IS}>
                    <option value="">Selecione...</option>
                    {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
                </Field>
                <Field label="Vencimento" t={{sub:"var(--sub)",text5:"var(--tx5)"}}>
                  <input type="date" value={recForm.due_date} onChange={e=>setRecForm(f=>({...f,due_date:e.target.value}))} style={IS}/>
                </Field>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:".5rem",marginBottom:".5rem"}}>
                <Field label="Descrição" t={{sub:"var(--sub)",text5:"var(--tx5)"}}>
                  <input value={recForm.description} onChange={e=>setRecForm(f=>({...f,description:e.target.value}))} placeholder="Ex: Venda parcelada, fiado..." style={IS}/>
                </Field>
                <Field label="Valor (R$)" t={{sub:"var(--sub)",text5:"var(--tx5)"}}>
                  <input type="number" min="0" step="0.01" value={recForm.value} onChange={e=>setRecForm(f=>({...f,value:e.target.value}))} placeholder="0,00" style={IS}/>
                </Field>
              </div>
              <Btn v="ok" onClick={async()=>{
                if(!recForm.description||!recForm.value){toast$("Preencha descrição e valor.","#f56565");return;}
                await addReceivable(recForm);
                setRecForm({client_id:"",client_name:"",description:"",value:"",due_date:"",notes:""});
              }}><Ic n="save" s={13}/>Registrar</Btn>
            </div>
            {/* Lista */}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem"}}>
              {receivables.length===0?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2.5rem 0",fontSize:".8rem"}}>Nenhuma conta registrada.</p>
              :receivables.map(r=>{
                const overdue=!r.paid&&r.due_date&&new Date(r.due_date)<new Date();
                return(
                  <div key={r.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",background:overdue?"#f5656506":"transparent"}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".15rem"}}>
                        <span style={{fontWeight:700,fontSize:".83rem",color:r.paid?"var(--tx5)":"var(--tx)",textDecoration:r.paid?"line-through":"none"}}>{r.description}</span>
                        {r.paid&&<Badge color="#10b981" sm>✅ Recebido</Badge>}
                        {overdue&&<Badge color="#f56565" sm>⚠️ Vencido</Badge>}
                      </div>
                      <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{r.client_name||"Sem cliente"}{r.due_date&&` · Vcto: ${new Date(r.due_date+"T00:00:00").toLocaleDateString("pt-BR")}`}{r.paid_date&&` · Recebido: ${r.paid_date}`}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:".6rem",flexShrink:0,marginLeft:".75rem"}}>
                      <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:r.paid?"#10b981":"#f59e0b",fontSize:".9rem"}}>{fmt(r.value)}</span>
                      {!r.paid&&<button onClick={()=>payReceivable(r.id)} style={{background:"#10b98115",border:"1px solid #10b98130",borderRadius:".4rem",padding:".28rem .6rem",color:"#10b981",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>✅ Recebido</button>}
                      <button onClick={()=>deleteReceivable(r.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem",cursor:"pointer"}}><Ic n="trash" s={13}/></button>
                    </div>
                  </div>
                );
              })}
            </div>
          </>);
        })()}

        {tab==="usuarios"&&isAdmin&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Usuários</h2>
            <Btn sm onClick={()=>setModal("addUser")}><Ic n="plus" s={12}/>Novo</Btn>
          </div>
          <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
            {appUsers.map(u=>(
              <div key={u.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".82rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                <div style={{display:"flex",alignItems:"center",gap:".7rem"}}>
                  <div style={{width:32,height:32,borderRadius:"50%",background:u.active?"linear-gradient(135deg,#4f5ef0,#10b981)":"var(--bdr2)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:".82rem",color:"#fff",flexShrink:0}}>{u.display_name.charAt(0).toUpperCase()}</div>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:".38rem"}}>
                      <span style={{fontWeight:700,fontSize:".85rem",color:u.active?"var(--tx)":"var(--tx5)"}}>{u.display_name}</span>
                      <Badge color={u.role==="admin"?"#f59e0b":u.role==="operator"?"#4f5ef0":"var(--tx5)"} sm>{ROLES[u.role]}</Badge>
                      {!u.active&&<Badge color="#f56565" sm>Inativo</Badge>}
                      {u.id===cu.id&&<Badge color="#10b981" sm>Você</Badge>}
                    </div>
                    <div style={{fontSize:".65rem",color:"var(--tx5)"}}>@{u.username}{u.last_login&&` · ${new Date(u.last_login).toLocaleDateString("pt-BR")}`}</div>
                  </div>
                </div>
                <button onClick={()=>{setEditing({...u,new_password:"",new_password2:""});setModal("editUser");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>
              </div>
            ))}
          </div>
        </>)}
      </div>
    </div>

    {/* ══ ANALYTICS ══ */}
    {showA&&<Analytics onClose={()=>setShowA(false)} sales={sales} cashTx={cashTx} products={products} clients={clients} dark={dark}/>}

    {/* ══ MODAIS ══ */}

    {/* Nova venda — CARRINHO */}
    {modal==="sale"&&(
      <Modal title="Nova Venda" onClose={()=>{setModal(null);cartReset();}} icon="sales" wide>

        {/* Info banner */}
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:"1rem",fontSize:".72rem",color:"#4f5ef0",display:"flex",gap:".35rem",alignItems:"center"}}><Ic n="info" s={12}/>Múltiplos produtos · Preço 0 = cortesia (registra custo) · Baixa estoque automática</div>

        {/* ── ITENS DO CARRINHO ── */}
        <div style={{marginBottom:"1rem"}}>
          <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".5rem",fontWeight:700}}>
            🛒 Itens da Venda
          </div>

          {/* Cabeçalho colunas */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 80px 110px 80px",gap:".4rem",marginBottom:".35rem",padding:"0 .1rem"}}>
            {["Produto","Qtd","Preço unit.",""].map(h=>(
              <div key={h} style={{fontSize:".62rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em"}}>{h}</div>
            ))}
          </div>

          {/* Linhas de produto */}
          {cartItems.map((item,idx)=>(
            <div key={item.key} style={{display:"grid",gridTemplateColumns:"1fr 80px 110px 36px",gap:".4rem",marginBottom:".4rem",alignItems:"center"}}>

              {/* Seletor de produto */}
              <select
                value={item.product_id}
                onChange={e=>cartSetProd(item.key,e.target.value)}
                style={{...IS,fontSize:".8rem",padding:".45rem .6rem"}}
              >
                <option value="">Selecione...</option>
                {products.map(p=>(
                  <option key={p.id} value={p.id} disabled={p.stock_qty<=0}>
                    {CATS[p.category]||"📋"} {p.name} {p.stock_qty<=0?"(zerado)":`(${p.stock_qty} ${p.unit||"un"})`}
                  </option>
                ))}
              </select>

              {/* Quantidade */}
              <input
                type="number" min="1"
                value={item.quantity}
                onChange={e=>cartSetQty(item.key,e.target.value)}
                style={{...IS,fontSize:".82rem",padding:".45rem .5rem",textAlign:"center"}}
              />

              {/* Preço unitário */}
              <input
                type="number" min="0" step="0.01"
                value={item.unit_price||"" }
                onChange={e=>cartSetPrice(item.key,e.target.value)}
                placeholder="0,00"
                style={{...IS,fontSize:".82rem",padding:".45rem .5rem"}}
              />

              {/* Remover linha */}
              <button
                onClick={()=>cartRemoveLine(item.key)}
                disabled={cartItems.length===1}
                style={{background:cartItems.length===1?"transparent":"#1e1010",border:`1px solid ${cartItems.length===1?"var(--bdr)":"#3a1515"}`,borderRadius:".4rem",padding:".4rem",color:cartItems.length===1?"var(--bdr2)":"#f56565",display:"flex",alignItems:"center",justifyContent:"center",height:"100%"}}
                title="Remover item"
              ><Ic n="trash" s={13}/></button>
            </div>
          ))}

          {/* Subtotal por linha */}
          {cartItems.some(i=>i.product_id)&&(
            <div style={{display:"grid",gridTemplateColumns:"1fr 80px 110px 36px",gap:".4rem",marginBottom:".4rem",padding:"0 .1rem"}}>
              {cartItems.map(i=>(
                <React.Fragment key={i.key+"_sub"}>
                  <div style={{fontSize:".7rem",color:"var(--tx4)",alignSelf:"center",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{i.product_name||"—"}</div>
                  <div style={{fontSize:".7rem",color:"var(--sub)",textAlign:"center"}}>{i.quantity>0?i.quantity:"—"}</div>
                  <div style={{fontSize:".75rem",fontWeight:700,color:i.unit_price>0?"#10b981":"var(--sub)",fontFamily:"'Syne',sans-serif"}}>{i.unit_price>0?fmt(i.unit_price*i.quantity):"—"}</div>
                  <div/>
                </React.Fragment>
              ))}
            </div>
          )}

          {/* Botão adicionar produto */}
          <button
            onClick={cartAddLine}
            style={{display:"flex",alignItems:"center",gap:".35rem",padding:".4rem .85rem",borderRadius:".45rem",border:"1px dashed #4f5ef060",background:"#4f5ef008",color:"#4f5ef0",fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,width:"100%",justifyContent:"center",marginTop:".25rem"}}
          >
            <Ic n="plus" s={13}/> Adicionar produto
          </button>
        </div>

        {/* ── ENTREGA ── */}
        <div style={{background:"var(--pill)",border:`1px solid ${cartDelivery?"#f59e0b40":"var(--bdr)"}`,borderRadius:".6rem",padding:".85rem",marginBottom:"1rem",transition:"border-color .25s"}}>

          {/* Toggle flag entregador */}
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:cartDelivery?".85rem":0}}>
            <div style={{display:"flex",alignItems:"center",gap:".55rem"}}>
              <span style={{fontSize:"1.2rem"}}>🛵</span>
              <div>
                <div style={{fontSize:".8rem",fontWeight:700,color:"var(--tx)",fontFamily:"'DM Sans',sans-serif"}}>Entregador Solicitado</div>
                <div style={{fontSize:".65rem",color:"var(--tx5)"}}>Ativa cobrança de frete + registro de custo do entregador</div>
              </div>
            </div>
            {/* Switch toggle */}
            <button
              onClick={()=>setCartDelivery(v=>{if(v){setCartFreight("");setCartDeliveryCost("");}return !v;})}
              style={{
                width:44,height:24,borderRadius:12,padding:2,border:"none",cursor:"pointer",
                background:cartDelivery?"linear-gradient(135deg,#f59e0b,#d97706)":"var(--bdr2)",
                transition:"background .25s",position:"relative",flexShrink:0
              }}
            >
              <div style={{
                width:20,height:20,borderRadius:"50%",background:"#fff",
                position:"absolute",top:2,
                left:cartDelivery?22:2,
                transition:"left .25s",
                boxShadow:"0 1px 4px rgba(0,0,0,.3)"
              }}/>
            </button>
          </div>

          {/* Campos visíveis só quando flag ativa */}
          {cartDelivery&&(
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".65rem"}}>
              {/* Taxa cobrada do cliente */}
              <div>
                <div style={{fontSize:".65rem",color:"#10b981",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".3rem",fontWeight:700,display:"flex",alignItems:"center",gap:".3rem"}}>
                  📈 Taxa cobrada do cliente
                </div>
                <input
                  type="number" min="0" step="0.01"
                  value={cartFreight}
                  onChange={e=>setCartFreight(e.target.value)}
                  placeholder="R$ 0,00"
                  style={{...IS,fontSize:".83rem",borderColor:cartFreightVal>0?"#10b98150":"var(--bdr2)"}}
                />
                {cartFreightVal>0&&<div style={{fontSize:".65rem",color:"#10b981",marginTop:".2rem"}}>+ {fmt(cartFreightVal)} na receita</div>}
              </div>

              {/* Custo do entregador */}
              <div>
                <div style={{fontSize:".65rem",color:"#f56565",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".3rem",fontWeight:700,display:"flex",alignItems:"center",gap:".3rem"}}>
                  📉 Custo pago ao entregador
                </div>
                <input
                  type="number" min="0" step="0.01"
                  value={cartDeliveryCost}
                  onChange={e=>setCartDeliveryCost(e.target.value)}
                  placeholder="R$ 0,00"
                  style={{...IS,fontSize:".83rem",borderColor:cartDeliveryCostVal>0?"#f5656550":"var(--bdr2)"}}
                />
                {cartDeliveryCostVal>0&&<div style={{fontSize:".65rem",color:"#f56565",marginTop:".2rem"}}>- {fmt(cartDeliveryCostVal)} no caixa</div>}
              </div>

              {/* Lucro líquido da entrega */}
              {(cartFreightVal>0||cartDeliveryCostVal>0)&&(
                <div style={{gridColumn:"1/-1",background:cartFreightVal-cartDeliveryCostVal>=0?"#10b98110":"#f5656510",borderRadius:".4rem",padding:".5rem .75rem",display:"flex",justifyContent:"space-between",alignItems:"center",border:`1px solid ${cartFreightVal-cartDeliveryCostVal>=0?"#10b98130":"#f5656530"}`}}>
                  <span style={{fontSize:".73rem",color:"var(--tx4)",fontWeight:600}}>
                    {cartFreightVal-cartDeliveryCostVal>=0?"💰 Lucro líquido da entrega":"⚠️ Prejuízo na entrega"}
                  </span>
                  <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".88rem",color:cartFreightVal-cartDeliveryCostVal>=0?"#10b981":"#f56565"}}>
                    {fmt(Math.abs(cartFreightVal-cartDeliveryCostVal))}
                  </span>
                </div>
              )}
            </div>
          )}
        </div>

        {/* ── DESCONTO ── */}
        {cartSubtotal>0&&(
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:"1rem"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".45rem"}}>
              <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",fontWeight:700}}>🏷️ Desconto</div>
              <div style={{display:"flex",gap:".25rem"}}>
                {["fixed","percent"].map(t=>(
                  <button key={t} onClick={()=>setCartDiscountType(t)} style={{padding:".2rem .55rem",borderRadius:".3rem",fontSize:".68rem",fontWeight:600,fontFamily:"'DM Sans',sans-serif",border:`1px solid ${cartDiscountType===t?"#4f5ef0":"var(--bdr2)"}`,background:cartDiscountType===t?"#4f5ef020":"transparent",color:cartDiscountType===t?"#4f5ef0":"var(--navoff)"}}>
                    {t==="fixed"?"R$":"%"}
                  </button>
                ))}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:".5rem",alignItems:"center"}}>
              <input type="number" min="0" step="0.01" value={cartDiscount} onChange={e=>setCartDiscount(e.target.value)} placeholder={cartDiscountType==="percent"?"ex: 10 (%)":"ex: 20,00 (R$)"} style={{...IS,fontSize:".83rem"}}/>
              <div style={{fontSize:".75rem"}}>
                {cartDiscountVal>0
                  ?<span style={{color:"#4f5ef0",fontWeight:600}}>- {fmt(cartDiscountVal)} no total</span>
                  :<span style={{color:"var(--tx5)"}}>Sem desconto</span>}
              </div>
            </div>
          </div>
        )}

        {/* ── CLIENTE + PAGAMENTO ── */}
        <R2>
          <Field label="Cliente" hint="opcional">
            <select
              value={cartClient.id}
              onChange={e=>{const c=clients.find(x=>x.id===e.target.value);setCartClient({id:e.target.value,name:c?.name||""});}}
              style={IS}
            >
              <option value="">Sem cliente</option>
              {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="Forma de Pagamento">
            <select value={cartPayment} onChange={e=>setCartPayment(e.target.value)} style={IS}>
              {PAYS.map(m=><option key={m} value={m}>{m}</option>)}
            </select>
          </Field>
        </R2>
        <Inp label="Observações" hint="opcional" placeholder="Ex: Entregar no período da tarde..." value={cartNotes} onChange={e=>setCartNotes(e.target.value)}/>

        {/* ── RESUMO TOTAL ── */}
        <div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98115)",border:"1px solid #4f5ef030",borderRadius:".65rem",padding:"1rem",marginBottom:"1rem"}}>
          <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".6rem",fontWeight:700}}>💰 Resumo da Venda</div>
          <div style={{display:"flex",flexDirection:"column",gap:".3rem"}}>
            {cartItems.filter(i=>i.product_id).map(i=>(
              <div key={i.key+"_res"} style={{display:"flex",justifyContent:"space-between",fontSize:".78rem"}}>
                <span style={{color:"var(--tx4)"}}>{i.product_name} × {i.quantity}{i.unit_price===0?" 🎁":""}</span>
                <span style={{color:i.unit_price===0?"#f59e0b":"var(--tx3)",fontWeight:600}}>
                  {i.unit_price===0?"Cortesia":fmt(i.unit_price*i.quantity)}
                </span>
              </div>
            ))}
            {cartDelivery&&cartFreightVal>0&&(
              <div style={{display:"flex",justifyContent:"space-between",fontSize:".78rem"}}>
                <span style={{color:"var(--tx4)"}}>🛵 Frete cobrado do cliente</span>
                <span style={{color:"#10b981",fontWeight:600}}>+ {fmt(cartFreightVal)}</span>
              </div>
            )}
            {cartDelivery&&cartDeliveryCostVal>0&&(
              <div style={{display:"flex",justifyContent:"space-between",fontSize:".78rem"}}>
                <span style={{color:"var(--tx4)"}}>📉 Custo entregador (sai do caixa)</span>
                <span style={{color:"#f56565",fontWeight:600}}>- {fmt(cartDeliveryCostVal)}</span>
              </div>
            )}
            {cartDiscountVal>0&&(
              <div style={{display:"flex",justifyContent:"space-between",fontSize:".78rem"}}>
                <span style={{color:"var(--tx4)"}}>🏷️ Desconto {cartDiscountType==="percent"?`(${cartDiscount}%)`:"(fixo)"}</span>
                <span style={{color:"#4f5ef0",fontWeight:600}}>- {fmt(cartDiscountVal)}</span>
              </div>
            )}
            <div style={{borderTop:"1px solid #4f5ef030",marginTop:".4rem",paddingTop:".4rem",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx)"}}>TOTAL</span>
              <span style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.2rem",color:"#10b981"}}>{fmt(cartTotal)}</span>
            </div>
            {cartItems.filter(i=>i.product_id).length>0&&(
              <div style={{fontSize:".67rem",color:"var(--sub)",textAlign:"right"}}>
                {cartItems.filter(i=>i.product_id).length} produto(s) · {cartItems.filter(i=>i.product_id).reduce((a,i)=>a+i.quantity,0)} unidade(s)
                {cartDelivery&&" · 🛵 Com entrega"}
              </div>
            )}
          </div>
        </div>

        {/* ── AÇÕES ── */}
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between",alignItems:"center"}}>
          <button onClick={cartReset} style={{background:"none",border:"none",color:"var(--tx5)",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",cursor:"pointer",display:"flex",alignItems:"center",gap:".3rem"}}>
            <Ic n="trash" s={12}/> Limpar
          </button>
          <div style={{display:"flex",gap:".5rem"}}>
            <Btn v="ghost" onClick={()=>{setModal(null);cartReset();}}>Cancelar</Btn>
            <Btn v="ok" onClick={registerSale} disabled={cartItems.every(i=>!i.product_id)}>
              <Ic n="save" s={13}/>Finalizar Venda · {fmt(cartTotal)}
            </Btn>
          </div>
        </div>
      </Modal>
    )}

    {/* Editar venda */}
    {modal==="editSale"&&editing&&(
      <Modal title="Editar Venda" onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
        <Inp label="Produto" value={editing.product_name} onChange={e=>setEditing(v=>({...v,product_name:e.target.value}))}/>
        <R2><Inp label="Quantidade" type="number" min="1" value={editing.quantity} onChange={e=>setEditing(v=>({...v,quantity:e.target.value}))}/><Inp label="Preço unit. (R$)" type="number" min="0" step="0.01" value={editing.unit_price} onChange={e=>setEditing(v=>({...v,unit_price:e.target.value}))}/></R2>
        <Inp label="Cliente" value={editing.client_name||""} onChange={e=>setEditing(v=>({...v,client_name:e.target.value}))}/>
        <R2><Sel label="Pagamento" value={editing.payment_method||"PIX"} onChange={e=>setEditing(v=>({...v,payment_method:e.target.value}))}>{PAYS.map(m=><option key={m} value={m}>{m}</option>)}</Sel><Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/></R2>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {isAdmin&&<Btn v="del" sm onClick={()=>{deleteSale(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn v="ok" onClick={updateSale}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Entrada estoque */}
    {modal==="stockEntry"&&(
      <Modal title="Entrada de Estoque" onClose={()=>setModal(null)} icon="stock">
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".55rem .8rem",marginBottom:".8rem",fontSize:".73rem",color:"#10b981",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="arrup" s={12}/>Entrada gera saída no caixa (custo de compra)</div>
        <Sel label="Produto *" value={stF.product_id} onChange={e=>setStF(s=>({...s,product_id:e.target.value}))}>
          <option value="">Selecione...</option>{products.map(p=><option key={p.id} value={p.id}>{CATS[p.category]||"📋"} {p.name} — Atual: {p.stock_qty} {p.unit||"un"}</option>)}
        </Sel>
        <R2><Inp label={`Quantidade * (${stF.product_id?products.find(p=>p.id===stF.product_id)?.unit||"un":"un"})`} type="number" min="1" placeholder="Ex: 10" value={stF.qty} onChange={e=>setStF(s=>({...s,qty:e.target.value}))}/><Inp label="Custo total (R$)" hint="opcional" type="number" min="0" step="0.01" placeholder="0,00" value={stF.cost_total} onChange={e=>setStF(s=>({...s,cost_total:e.target.value}))}/></R2>
        {stF.product_id&&stF.qty&&parseInt(stF.qty)>0&&(()=>{const p=products.find(x=>x.id===stF.product_id);return <div style={{background:"#0e1e0e",borderRadius:".45rem",padding:".55rem .8rem",marginBottom:".8rem",fontSize:".76rem",color:"#10b981"}}>✅ +{stF.qty} {p?.unit||"un"} → novo estoque: {(p?.stock_qty||0)+parseInt(stF.qty)}</div>})()}
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn v="ok" onClick={registerStock}><Ic n="save" s={13}/>Confirmar</Btn></div>
      </Modal>
    )}

    {/* Lançamento caixa */}
    {modal==="cashTx"&&(
      <Modal title="Lançamento Manual" onClose={()=>setModal(null)} icon="cash">
        <Sel label="Tipo" value={cashF.type} onChange={e=>setCashF(f=>({...f,type:e.target.value}))}><option value="entrada">📈 Entrada</option><option value="saida">📉 Saída</option></Sel>
        <Inp label="Descrição *" placeholder="Ex: Pagamento fornecedor..." value={cashF.description} onChange={e=>setCashF(f=>({...f,description:e.target.value}))}/>
        <Inp label="Valor (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={cashF.value} onChange={e=>setCashF(f=>({...f,value:e.target.value}))}/>
        <Inp label="Categoria" placeholder="Ex: Despesa fixa..." value={cashF.category} onChange={e=>setCashF(f=>({...f,category:e.target.value}))}/>
        {cashF.type==="saida"&&<>
          <Sel label="Produto vinculado" hint="opcional — baixa estoque" value={cashF.product_id} onChange={e=>setCashF(f=>({...f,product_id:e.target.value}))}>
            <option value="">Nenhum</option>{products.map(p=><option key={p.id} value={p.id}>{p.name} (estoque: {p.stock_qty})</option>)}
          </Sel>
          {cashF.product_id&&<Inp label="Quantidade" type="number" min="1" value={cashF.quantity} onChange={e=>setCashF(f=>({...f,quantity:e.target.value}))}/>}
        </>}
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn v="ok" onClick={addCash} disabled={!cashF.description||!cashF.value}><Ic n="save" s={13}/>Salvar</Btn></div>
      </Modal>
    )}

    {/* Editar caixa */}
    {modal==="editCash"&&editing&&(
      <Modal title="Editar Lançamento" onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
        <Sel label="Tipo" value={editing.type} onChange={e=>setEditing(v=>({...v,type:e.target.value}))}><option value="entrada">📈 Entrada</option><option value="saida">📉 Saída</option></Sel>
        <Inp label="Descrição *" value={editing.description} onChange={e=>setEditing(v=>({...v,description:e.target.value}))}/>
        <Inp label="Valor (R$) *" type="number" min="0" step="0.01" value={editing.value} onChange={e=>setEditing(v=>({...v,value:e.target.value}))}/>
        <Inp label="Categoria" value={editing.category||""} onChange={e=>setEditing(v=>({...v,category:e.target.value}))}/>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {isAdmin&&<Btn v="del" sm onClick={()=>{deleteCash(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn v="ok" onClick={updateCash}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Novo produto */}
    {modal==="produto"&&(
      <Modal title="Novo Produto" onClose={()=>setModal(null)} icon="product" wide>
        <R2><Inp label="Código" hint="auto" placeholder="PRD-001" value={pf.code} onChange={e=>setPf(f=>({...f,code:e.target.value}))}/><Sel label="Categoria" value={pf.category} onChange={e=>setPf(f=>({...f,category:e.target.value}))}><option value="tirzepatida">💉 Tirzepatida</option><option value="seringa">🩺 Seringa/Agulha</option><option value="material">🧊 Material</option><option value="embalagem">📦 Embalagem</option><option value="outro">📋 Outro</option></Sel></R2>
        <Inp label="Nome *" placeholder="Ex: Tirzepatida 2.5mg, Isopor 10L..." value={pf.name} onChange={e=>setPf(f=>({...f,name:e.target.value}))}/>
        <R2><Sel label="Fornecedor" hint="opcional" value={pf.supplier_id} onChange={e=>setPf(f=>({...f,supplier_id:e.target.value}))}><option value="">Nenhum</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</Sel><Inp label="Unidade" hint="ex: ampola, un, cx..." placeholder="ampola" value={pf.unit} onChange={e=>setPf(f=>({...f,unit:e.target.value}))}/></R2>
        <R2><Inp label="Lote" placeholder="L2025001" value={pf.batch} onChange={e=>setPf(f=>({...f,batch:e.target.value}))}/><Inp label="Vencimento" type="date" value={pf.expiry} onChange={e=>setPf(f=>({...f,expiry:e.target.value}))}/></R2>
        <R2><Inp label="Custo/un (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={pf.cost_per_unit} onChange={e=>setPf(f=>({...f,cost_per_unit:e.target.value}))}/><Inp label="Preço venda/un (R$) *" type="number" min="0" step="0.01" placeholder="0,00" value={pf.price_per_unit} onChange={e=>setPf(f=>({...f,price_per_unit:e.target.value}))}/></R2>
        <MPreview cost={pf.cost_per_unit} price={pf.price_per_unit}/>
        <R2><Inp label="Estoque inicial" type="number" min="0" value={pf.stock_qty} onChange={e=>setPf(f=>({...f,stock_qty:e.target.value}))}/><Inp label="Estoque mínimo" hint="alerta" type="number" min="0" value={pf.min_stock} onChange={e=>setPf(f=>({...f,min_stock:e.target.value}))}/></R2>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between",alignItems:"center"}}>
          <Btn v="ghost" sm onClick={()=>setModal("addSupp")}><Ic n="supplier" s={12}/>+ Fornecedor</Btn>
          <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addProduct}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Editar produto */}
    {modal==="editProd"&&editing&&(
      <Modal title="Editar Produto" onClose={()=>{setModal(null);setEditing(null);}} icon="edit" wide>
        <R2><Inp label="Código" value={editing.code||""} onChange={e=>setEditing(v=>({...v,code:e.target.value}))}/><Sel label="Categoria" value={editing.category||"outro"} onChange={e=>setEditing(v=>({...v,category:e.target.value}))}><option value="tirzepatida">💉 Tirzepatida</option><option value="seringa">🩺 Seringa/Agulha</option><option value="material">🧊 Material</option><option value="embalagem">📦 Embalagem</option><option value="outro">📋 Outro</option></Sel></R2>
        <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
        <R2><Sel label="Fornecedor" value={editing.supplier_id||""} onChange={e=>setEditing(v=>({...v,supplier_id:e.target.value}))}><option value="">Nenhum</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</Sel><Inp label="Unidade" value={editing.unit||"un"} onChange={e=>setEditing(v=>({...v,unit:e.target.value}))}/></R2>
        <R2><Inp label="Lote" value={editing.batch||""} onChange={e=>setEditing(v=>({...v,batch:e.target.value}))}/><Inp label="Vencimento" type="date" value={editing.expiry||""} onChange={e=>setEditing(v=>({...v,expiry:e.target.value}))}/></R2>
        <R2><Inp label="Custo/un (R$)" type="number" min="0" step="0.01" value={editing.cost_per_unit} onChange={e=>setEditing(v=>({...v,cost_per_unit:e.target.value}))}/><Inp label="Preço venda/un (R$)" type="number" min="0" step="0.01" value={editing.price_per_unit} onChange={e=>setEditing(v=>({...v,price_per_unit:e.target.value}))}/></R2>
        <MPreview cost={editing.cost_per_unit} price={editing.price_per_unit}/>
        <R2><Inp label="Estoque atual" type="number" min="0" value={editing.stock_qty} onChange={e=>setEditing(v=>({...v,stock_qty:e.target.value}))}/><Inp label="Estoque mínimo" type="number" min="0" value={editing.min_stock} onChange={e=>setEditing(v=>({...v,min_stock:e.target.value}))}/></R2>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {isAdmin&&<Btn v="del" sm onClick={()=>delProduct(editing.id)}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveProduct}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Fornecedores */}
    {modal==="fornecedores"&&(
      <Modal title="Fornecedores" onClose={()=>setModal(null)} icon="supplier" wide>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:".85rem"}}><Btn sm onClick={()=>setModal("addSupp")}><Ic n="plus" s={12}/>Novo</Btn></div>
        {suppliers.length===0?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2rem 0",fontSize:".8rem"}}>Nenhum fornecedor cadastrado.</p>
        :<div style={{display:"grid",gap:".55rem"}}>
          {suppliers.map(s=>(
            <div key={s.id} style={{background:"var(--inp)",border:"1px solid var(--bdr2)",borderRadius:".6rem",padding:".72rem 1rem",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div><div style={{fontWeight:700,fontSize:".85rem",color:"var(--tx)"}}>🏭 {s.name}</div><div style={{fontSize:".7rem",color:"var(--tx5)",marginTop:".12rem"}}>{[s.contact,s.phone,s.email].filter(Boolean).join(" · ")||"Sem contato"}</div>{s.notes&&<div style={{fontSize:".68rem",color:"#8b44f0"}}>📝 {s.notes}</div>}</div>
              {isAdmin&&<div style={{display:"flex",gap:".3rem"}}>
                <button onClick={()=>{setEditing({...s});setModal("editSupp");}} style={{background:"none",border:"none",color:"#4f5ef0"}}><Ic n="edit" s={13}/></button>
                <button onClick={async()=>{await supabase.from("suppliers").delete().eq("id",s.id);toast$("Removido.","#f59e0b");}} style={{background:"none",border:"none",color:"var(--tx6)"}}><Ic n="trash" s={13}/></button>
              </div>}
            </div>
          ))}
        </div>}
      </Modal>
    )}
    {modal==="addSupp"&&(<Modal title="Novo Fornecedor" onClose={()=>setModal(null)} icon="supplier">
      <Inp label="Nome *" placeholder="Ex: Indufar Paraguai" value={supf.name} onChange={e=>setSupf(f=>({...f,name:e.target.value}))}/>
      <Inp label="Contato" placeholder="Nome do responsável" value={supf.contact} onChange={e=>setSupf(f=>({...f,contact:e.target.value}))}/>
      <R2><Inp label="Telefone / WhatsApp" type="tel" placeholder="(66) 99999-9999" value={supf.phone} onChange={e=>setSupf(f=>({...f,phone:e.target.value}))}/><Inp label="E-mail" type="email" placeholder="contato@..." value={supf.email} onChange={e=>setSupf(f=>({...f,email:e.target.value}))}/></R2>
      <Inp label="Observações" hint="opcional" placeholder="Ex: prazo entrega..." value={supf.notes} onChange={e=>setSupf(f=>({...f,notes:e.target.value}))}/>
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addSupp}><Ic n="save" s={13}/>Salvar</Btn></div>
    </Modal>)}
    {modal==="editSupp"&&editing&&(<Modal title="Editar Fornecedor" onClose={()=>{setModal("fornecedores");setEditing(null);}} icon="edit">
      <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
      <Inp label="Contato" value={editing.contact||""} onChange={e=>setEditing(v=>({...v,contact:e.target.value}))}/>
      <R2><Inp label="Telefone" type="tel" value={editing.phone||""} onChange={e=>setEditing(v=>({...v,phone:e.target.value}))}/><Inp label="E-mail" type="email" value={editing.email||""} onChange={e=>setEditing(v=>({...v,email:e.target.value}))}/></R2>
      <Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/>
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>{setModal("fornecedores");setEditing(null);}}>Cancelar</Btn><Btn onClick={saveSupp}><Ic n="save" s={13}/>Salvar</Btn></div>
    </Modal>)}

    {/* Cliente */}
    {modal==="cliente"&&(<Modal title="Novo Cliente" onClose={()=>setModal(null)} icon="client">
      <Inp label="Nome *" placeholder="Nome completo" value={cf.name} onChange={e=>setCf(f=>({...f,name:e.target.value}))}/>
      <R2><Inp label="Telefone / WhatsApp" type="tel" placeholder="(66) 99999-9999" value={cf.phone} onChange={e=>setCf(f=>({...f,phone:e.target.value}))}/><Inp label="E-mail" type="email" placeholder="email@..." value={cf.email} onChange={e=>setCf(f=>({...f,email:e.target.value}))}/></R2>
      <Inp label="Observações" hint="ex: dosagem, periodicidade..." placeholder="Ex: Usa 2.5mg a cada 7 dias" value={cf.notes} onChange={e=>setCf(f=>({...f,notes:e.target.value}))}/>
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addClient}><Ic n="save" s={13}/>Salvar</Btn></div>
    </Modal>)}
    {modal==="editCliente"&&editing&&(<Modal title="Editar Cliente" onClose={()=>{setModal(null);setEditing(null);}} icon="edit" wide>
      <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
      <R2><Inp label="Telefone" type="tel" value={editing.phone||""} onChange={e=>setEditing(v=>({...v,phone:e.target.value}))}/><Inp label="E-mail" type="email" value={editing.email||""} onChange={e=>setEditing(v=>({...v,email:e.target.value}))}/></R2>
      <Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/>
      <div style={{borderTop:"1px solid var(--bdr)",paddingTop:".85rem",marginTop:".25rem"}}>
        <div style={{fontWeight:700,fontSize:".78rem",color:"#4f5ef0",marginBottom:".65rem",display:"flex",gap:".35rem",alignItems:"center"}}><Ic n="syringe" s={13}/>💉 Protocolo de Tratamento (opcional)</div>
        <R2>
          <Inp label="Dose atual" placeholder="Ex: 2.5mg, 5mg, 7.5mg..." value={editing.dose||""} onChange={e=>setEditing(v=>({...v,dose:e.target.value}))}/>
          <Inp label="Intervalo (dias)" type="number" min="1" placeholder="7" value={editing.interval_days||""} onChange={e=>setEditing(v=>({...v,interval_days:e.target.value}))}/>
        </R2>
        <Inp label="Data início tratamento" type="date" value={editing.treatment_start||""} onChange={e=>setEditing(v=>({...v,treatment_start:e.target.value}))}/>
        <Inp label="Notas do tratamento" placeholder="Ex: Responde bem à dose 5mg, próxima revisão em..." value={editing.treatment_notes||""} onChange={e=>setEditing(v=>({...v,treatment_notes:e.target.value}))}/>
      </div>
      <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
        {isAdmin&&<Btn v="del" sm onClick={()=>{delClient(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
        <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveClient}><Ic n="save" s={13}/>Salvar</Btn></div>
      </div>
    </Modal>)}

    {/* Usuários */}
    {modal==="addUser"&&(<Modal title="Novo Usuário" onClose={()=>setModal(null)} icon="users">
      <R2><Inp label="Login *" placeholder="ex: maria.silva" value={uf.username} onChange={e=>setUf(f=>({...f,username:e.target.value.toLowerCase().replace(/\s/g,".")}))}/><Inp label="Nome *" placeholder="ex: Maria Silva" value={uf.display_name} onChange={e=>setUf(f=>({...f,display_name:e.target.value}))}/></R2>
      <Sel label="Nível" value={uf.role} onChange={e=>setUf(f=>({...f,role:e.target.value}))}><option value="admin">Administrador — acesso total</option><option value="operator">Operador — adiciona e edita</option><option value="viewer">Visualizador — somente leitura</option></Sel>
      <R2>
        <Field label="Senha *" hint="mín. 6 caracteres"><div style={{position:"relative"}}><input type={pwv.a?"text":"password"} value={uf.password} onChange={e=>setUf(f=>({...f,password:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwv(s=>({...s,a:!s.a}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"var(--tx5)",padding:0}}><Ic n={pwv.a?"eyeoff":"eye"} s={13}/></button></div></Field>
        <Field label="Confirmar"><div style={{position:"relative"}}><input type={pwv.b?"text":"password"} value={uf.password2} onChange={e=>setUf(f=>({...f,password2:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwv(s=>({...s,b:!s.b}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"var(--tx5)",padding:0}}><Ic n={pwv.b?"eyeoff":"eye"} s={13}/></button></div></Field>
      </R2>
      {uf.password&&uf.password2&&uf.password!==uf.password2&&<div style={{fontSize:".73rem",color:"#f56565",marginBottom:".7rem",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="warn" s={12}/>Senhas não coincidem</div>}
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addUser} disabled={!uf.username||!uf.display_name||!uf.password||uf.password!==uf.password2}><Ic n="save" s={13}/>Criar</Btn></div>
    </Modal>)}
    {modal==="editUser"&&editing&&(<Modal title={`Editar · ${editing.display_name}`} onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
      <div style={{background:"var(--pill)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".8rem",fontSize:".73rem",color:"var(--tx5)"}}>Login: <b style={{color:"var(--tx4)"}}>@{editing.username}</b></div>
      <R2><Inp label="Nome *" value={editing.display_name} onChange={e=>setEditing(v=>({...v,display_name:e.target.value}))}/><Sel label="Nível" value={editing.role} onChange={e=>setEditing(v=>({...v,role:e.target.value}))} disabled={editing.id===cu.id}><option value="admin">Administrador</option><option value="operator">Operador</option><option value="viewer">Visualizador</option></Sel></R2>
      <Field label="Status"><div style={{display:"flex",gap:".5rem"}}>{[{v:true,l:"✅ Ativo"},{v:false,l:"🚫 Inativo"}].map(o=><button key={String(o.v)} onClick={()=>setEditing(v=>({...v,active:o.v}))} disabled={editing.id===cu.id} style={{flex:1,padding:".5rem",borderRadius:".45rem",border:`1px solid ${editing.active===o.v?"#4f5ef0":"var(--bdr2)"}`,background:editing.active===o.v?"#4f5ef020":"transparent",color:editing.active===o.v?"#4f5ef0":"var(--tx5)",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:".82rem"}}>{o.l}</button>)}</div></Field>
      <div style={{borderTop:"1px solid var(--bdr)",paddingTop:".8rem",marginBottom:".8rem"}}>
        <div style={{fontSize:".7rem",color:"#4f5ef0",marginBottom:".6rem",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="key" s={12}/>Nova senha (vazio = manter atual)</div>
        <R2>
          <Field label="Nova senha"><div style={{position:"relative"}}><input type={pwv.c?"text":"password"} value={editing.new_password||""} onChange={e=>setEditing(v=>({...v,new_password:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwv(s=>({...s,c:!s.c}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"var(--tx5)",padding:0}}><Ic n={pwv.c?"eyeoff":"eye"} s={13}/></button></div></Field>
          <Field label="Confirmar"><div style={{position:"relative"}}><input type={pwv.d?"text":"password"} value={editing.new_password2||""} onChange={e=>setEditing(v=>({...v,new_password2:e.target.value}))} placeholder="••••••••" style={IS}/><button onClick={()=>setPwv(s=>({...s,d:!s.d}))} style={{position:"absolute",right:".7rem",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"var(--tx5)",padding:0}}><Ic n={pwv.d?"eyeoff":"eye"} s={13}/></button></div></Field>
        </R2>
      </div>
      <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
        {editing.id!==cu.id?<Btn v="del" sm onClick={async()=>{await supabase.from("app_users").delete().eq("id",editing.id);toast$("Removido.","#f59e0b");setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>:<div/>}
        <div style={{display:"flex",gap:".5rem"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn onClick={saveUser}><Ic n="save" s={13}/>Salvar</Btn></div>
      </div>
    </Modal>)}


    {/* ═══ COMPROVANTE ═══ */}
    {showReceipt&&(
      <Modal title="Comprovante de Venda" onClose={()=>setShowReceipt(null)} icon="pdf" wide>
        <div id="receipt-content" style={{fontFamily:"'DM Sans',sans-serif"}}>
          <div style={{textAlign:"center",marginBottom:"1.25rem",paddingBottom:"1rem",borderBottom:"1px solid var(--bdr)"}}>
            <div style={{fontSize:"1.2rem",fontWeight:800,fontFamily:"'Syne',sans-serif",color:"var(--tx)"}}>CaixaPro · Tirzepatida</div>
            <div style={{fontSize:".75rem",color:"var(--tx5)",marginTop:".2rem"}}>Comprovante de Venda</div>
            <div style={{fontSize:".72rem",color:"var(--sub)",marginTop:".1rem"}}>{showReceipt[0]?.date}</div>
          </div>
          {showReceipt[0]?.client_name&&<div style={{marginBottom:"1rem",padding:".65rem .85rem",background:"var(--pill)",borderRadius:".5rem"}}>
            <div style={{fontSize:".65rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>Cliente</div>
            <div style={{fontWeight:700,color:"var(--tx)",fontSize:".88rem"}}>{showReceipt[0].client_name}</div>
          </div>}
          <table style={{width:"100%",borderCollapse:"collapse",marginBottom:"1rem"}}>
            <thead><tr style={{borderBottom:"1px solid var(--bdr)"}}>
              {["Produto","Qtd","Preço unit.","Total"].map(h=><th key={h} style={{textAlign:"left",padding:".4rem .3rem",fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",fontWeight:600}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {showReceipt.map((s,i)=><tr key={i} style={{borderBottom:"1px solid var(--sep)"}}>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx)"}}>{s.product_name}{s.unit_price===0&&" 🎁"}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx3)"}}>{s.quantity}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx3)"}}>{s.unit_price===0?"Cortesia":fmt(s.unit_price)}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",fontWeight:700,color:"#10b981"}}>{s.unit_price===0?"R$ 0,00":fmt(s.total_price)}</td>
              </tr>)}
            </tbody>
          </table>
          <div style={{display:"flex",flexDirection:"column",gap:".3rem",padding:"1rem",background:"var(--pill)",borderRadius:".5rem",marginBottom:"1rem"}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:".82rem"}}>
              <span style={{color:"var(--tx4)"}}>Subtotal</span>
              <span style={{color:"var(--tx3)",fontWeight:600}}>{fmt(showReceipt.reduce((a,s)=>a+s.total_price,0))}</span>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid var(--bdr)",paddingTop:".3rem"}}>
              <span style={{fontWeight:700,color:"var(--tx)",fontFamily:"'Syne',sans-serif"}}>TOTAL</span>
              <span style={{fontWeight:800,fontSize:"1.1rem",color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(showReceipt.reduce((a,s)=>a+s.total_price,0))}</span>
            </div>
            {showReceipt[0]?.payment_method&&<div style={{fontSize:".73rem",color:"var(--sub)",textAlign:"right"}}>Pagamento: {showReceipt[0].payment_method}</div>}
          </div>
          <div style={{textAlign:"center",fontSize:".68rem",color:"var(--tx6)"}}>Obrigado pela compra! · CaixaPro</div>
        </div>
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end",marginTop:"1rem",paddingTop:"1rem",borderTop:"1px solid var(--bdr)"}}>
          <Btn v="ghost" onClick={()=>setShowReceipt(null)}>Fechar</Btn>
          <Btn onClick={()=>window.print()}><Ic n="pdf" s={13}/>Imprimir</Btn>
        </div>
      </Modal>
    )}

    {/* ═══ HISTÓRICO CLIENTE ═══ */}
    {showClientHist&&(
      <Modal title={`Histórico · ${showClientHist.name}`} onClose={()=>setShowClientHist(null)} icon="analytics" wide>
        {(()=>{
          const cs=sales.filter(s=>s.client_id===showClientHist.id||s.client_name===showClientHist.name);
          const total=cs.reduce((a,s)=>a+s.total_price,0);
          const prods={};cs.forEach(s=>{prods[s.product_name]=(prods[s.product_name]||0)+s.quantity;});
          const lastBuy=cs[0]?.date||null;
          const daysSince=lastBuy?Math.floor((new Date()-new Date(cs[0].created_at))/86400000):null;
          return(<>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:"1rem"}}>
              <KCard label="Total gasto" value={fmt(total)} color="#10b981"/>
              <KCard label="Compras" value={cs.length} color="#4f5ef0"/>
              <KCard label="Última compra" value={lastBuy||"—"} sub={daysSince!==null?`${daysSince}d atrás`:""} color={daysSince!==null&&daysSince>30?"#f59e0b":"#10b981"}/>
            </div>
            {showClientHist.dose&&<div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98115)",border:"1px solid #4f5ef030",borderRadius:".65rem",padding:".85rem",marginBottom:"1rem"}}>
              <div style={{fontWeight:700,fontSize:".8rem",color:"#4f5ef0",marginBottom:".5rem"}}>💉 Protocolo de Tratamento</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem"}}>
                {[{l:"Dose atual",v:showClientHist.dose,c:"#4f5ef0"},{l:"Intervalo",v:`${showClientHist.interval_days||7} dias`,c:"#8b44f0"},{l:"Início",v:showClientHist.treatment_start?new Date(showClientHist.treatment_start+"T00:00:00").toLocaleDateString("pt-BR"):"—",c:"#10b981"}].map(m=>(
                  <div key={m.l} style={{textAlign:"center"}}>
                    <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div>
                    <div style={{fontWeight:700,fontSize:".88rem",color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                  </div>
                ))}
              </div>
              {showClientHist.treatment_notes&&<div style={{fontSize:".73rem",color:"var(--tx4)",marginTop:".5rem"}}>📝 {showClientHist.treatment_notes}</div>}
            </div>}
            {Object.keys(prods).length>0&&<div style={{background:"var(--pill)",borderRadius:".6rem",padding:".75rem",marginBottom:"1rem"}}>
              <div style={{fontWeight:700,fontSize:".78rem",color:"var(--tx2)",marginBottom:".5rem"}}>📦 Produtos mais comprados</div>
              {Object.entries(prods).sort((a,b)=>b[1]-a[1]).map(([n,q])=><div key={n} style={{display:"flex",justifyContent:"space-between",padding:".3rem 0",borderBottom:"1px solid var(--sep)",fontSize:".78rem"}}><span style={{color:"var(--tx)"}}>{n}</span><span style={{color:"#4f5ef0",fontWeight:700}}>{q} un</span></div>)}
            </div>}
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".5rem"}}>🛒 Últimas compras</div>
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".65rem"}}>
              {cs.length===0?<p style={{color:"var(--tx5)",textAlign:"center",padding:"1.5rem",fontSize:".8rem"}}>Nenhuma compra.</p>
              :cs.map(s=>(
                <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:".6rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                  <div><div style={{fontSize:".82rem",color:"var(--tx)",fontWeight:600}}>{s.product_name}</div><div style={{fontSize:".65rem",color:"var(--tx5)"}}>{s.date} · {s.quantity}un · {s.payment_method}</div></div>
                  <span style={{fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif",fontSize:".85rem"}}>{fmt(s.total_price)}</span>
                </div>
              ))}
            </div>
          </>);
        })()}
      </Modal>
    )}

    {/* ═══ CALCULADORA IMPORTAÇÃO ═══ */}
    {showImportCalc&&(
      <Modal title="🧮 Calculadora de Custo de Importação" onClose={()=>setShowImportCalc(false)}>
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".55rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="info" s={12}/>Calcule o custo real por unidade considerando frete e taxas</div>
        <Inp label="Custo total da remessa (R$) *" type="number" min="0" step="0.01" placeholder="Ex: 3500,00" value={importCalc.totalCost} onChange={e=>setImportCalc(v=>({...v,totalCost:e.target.value}))}/>
        <Inp label="Quantidade de unidades *" type="number" min="1" placeholder="Ex: 20 ampolas" value={importCalc.qty} onChange={e=>setImportCalc(v=>({...v,qty:e.target.value}))}/>
        <Inp label="Extras (frete, taxas, câmbio) R$" hint="opcional" type="number" min="0" step="0.01" placeholder="Ex: 250,00" value={importCalc.extras} onChange={e=>setImportCalc(v=>({...v,extras:e.target.value}))}/>
        {importCalc.totalCost&&importCalc.qty&&(()=>{
          const total=(parseFloat(importCalc.totalCost)||0)+(parseFloat(importCalc.extras)||0);
          const qty=parseInt(importCalc.qty)||1;
          const custo=total/qty;
          const sugeridoMark2=(custo*2).toFixed(2);
          const sugeridoMark3=(custo*3).toFixed(2);
          return <div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98115)",border:"1px solid #4f5ef030",borderRadius:".6rem",padding:"1rem",marginTop:".5rem"}}>
            <div style={{fontWeight:700,fontSize:".82rem",color:"#4f5ef0",marginBottom:".65rem",fontFamily:"'Syne',sans-serif"}}>📊 Resultado</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:".5rem",textAlign:"center"}}>
              {[{l:"Custo/unidade",v:fmt(custo),c:"#f56565"},{l:"Preço c/ 100% markup",v:fmt(custo*2),c:"#f59e0b"},{l:"Preço c/ 200% markup",v:fmt(custo*3),c:"#10b981"}].map(m=>(
                <div key={m.l} style={{background:"var(--pill)",borderRadius:".5rem",padding:".6rem .5rem"}}>
                  <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div>
                  <div style={{fontWeight:800,fontSize:".95rem",color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                </div>
              ))}
            </div>
            <div style={{fontSize:".7rem",color:"var(--tx5)",marginTop:".65rem",textAlign:"center"}}>
              Total remessa: {fmt(total)} · {qty} unidades
            </div>
          </div>;
        })()}
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:"1rem"}}>
          <Btn v="ghost" onClick={()=>{setShowImportCalc(false);setImportCalc({totalCost:"",qty:"",extras:""});}}>Fechar</Btn>
        </div>
      </Modal>
    )}

    {/* Logout */}
    {logoutC&&(
      <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",backdropFilter:"blur(6px)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:"1.25rem"}}>
        <div style={{background:"var(--bg3)",border:"1px solid var(--bdr2)",borderRadius:"1rem",padding:"1.65rem",maxWidth:290,width:"100%",textAlign:"center",boxShadow:"0 24px 64px var(--shadow)"}}>
          <div style={{color:"#4f5ef0",marginBottom:".75rem"}}><Ic n="logout" s={30}/></div>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".92rem",color:"var(--tx)",marginBottom:".4rem"}}>Sair do sistema?</div>
          <div style={{fontSize:".77rem",color:"var(--tx5)",marginBottom:"1.1rem"}}>Precisará fazer login novamente.</div>
          <div style={{display:"flex",gap:".6rem",justifyContent:"center"}}>
            <Btn v="ghost" onClick={()=>setLogoutC(false)}>Cancelar</Btn>
            <Btn v="del" onClick={logout}><Ic n="logout" s={13}/>Sair</Btn>
          </div>
        </div>
      </div>
    )}
  </>);
}
