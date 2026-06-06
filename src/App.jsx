import React,{useState,useEffect,useCallback,useMemo}from"react";
import{supabase}from"./supabase";

//  HELPERS 
const fmt=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(Number(v)||0);
const fmtN=v=>new Intl.NumberFormat("pt-BR").format(Number(v)||0);
const fmtPct=v=>`${(Number(v)||0).toFixed(1)}%`;
const uid=()=>Math.random().toString(36).slice(2,10);
const today=()=>new Date().toLocaleDateString("pt-BR");
const nowISO=()=>new Date().toISOString();
// Soma N dias úteis (pula fim de semana) e devolve no formato BR DD/MM/AAAA (igual ao today())
const addBizDays=(n)=>{const d=new Date();let added=0;while(added<(parseInt(n)||0)){d.setDate(d.getDate()+1);const wd=d.getDay();if(wd!==0&&wd!==6)added++;}return d.toLocaleDateString("pt-BR");};
// Mesma soma, mas devolve ISO YYYY-MM-DD (para campos due_date de recebíveis)
const addBizDaysISO=(n)=>{const d=new Date();let added=0;while(added<(parseInt(n)||0)){d.setDate(d.getDate()+1);const wd=d.getDay();if(wd!==0&&wd!==6)added++;}return d.toISOString().slice(0,10);};
// Lê com segurança os itens (JSON) de um pedido; nunca quebra a tela se o dado vier corrompido
const parseItems=(o)=>{try{const v=JSON.parse((o&&o.items)||"[]");return Array.isArray(v)?v:[];}catch(e){return [];}};
const hashPw=s=>btoa(unescape(encodeURIComponent(s+"|caixapro2026")));
const calcM=(c,p)=>{const cv=parseFloat(c)||0,pv=parseFloat(p)||0;return{markup:cv>0?((pv-cv)/cv)*100:0,margin:pv>0?((pv-cv)/pv)*100:0,profit:pv-cv};};
const daysUntil=d=>d?Math.ceil((new Date(d)-new Date())/86400000):null;
const expColor=d=>d===null?"var(--tx5)":d<0?"#f56565":d<=30?"#f59e0b":"#10b981";
const stkColor=q=>q<=0?"#f56565":q<=5?"#f59e0b":"#10b981";

const ROLES={admin:"Admin",operator:"Operador",viewer:"Visualizador"};
const DEFAULT_CATS=[
  {key:"tirzepatida",label:"Tirzepatida",icon:"💉"},
  {key:"seringa",label:"Seringas",icon:"🩺"},
  {key:"caneta",label:"Canetas",icon:"✒️"},
  {key:"acessorio",label:"Acessórios",icon:"📦"},
  {key:"outro",label:"Outros",icon:"📋"},
];
const CATS={tirzepatida:"💉",seringa:"🩺",caneta:"✒️",acessorio:"📦",outro:"📋"};
const PAYS_SIMPLES=["Dinheiro","PIX","Transferência","Promissória"];
// Liquidação padrão por forma de pagamento. type: instant | days(+days) | duedate(promissória) | parcelado
const DEFAULT_SETTLE={"Dinheiro":{type:"instant"},"PIX":{type:"instant"},"Transferência":{type:"instant"},"Débito":{type:"days",days:1},"Crédito à Vista":{type:"days",days:1},"Crédito Parcelado":{type:"parcelado"},"Promissória":{type:"duedate"}};
const CARD_BRANDS=["Visa","Mastercard","Elo","Hipercard","AmEx"];
const CARD_MODES=[
  {key:"debito",    label:"Débito"},
  {key:"vista",     label:"Crédito à Vista"},
  {key:"parc2a6",   label:"Parcelado 2-6x"},
  {key:"parc7a12",  label:"Parcelado 7-12x"},
];
const DEFAULT_TAXES={
  Visa:      {debito:1.99,vista:2.99,parc2a6:3.49,parc7a12:3.99,antecip:1.99},
  Mastercard:{debito:1.89,vista:2.89,parc2a6:3.39,parc7a12:3.89,antecip:1.99},
  Elo:       {debito:2.09,vista:3.09,parc2a6:3.59,parc7a12:4.09,antecip:2.09},
  Hipercard: {debito:2.19,vista:3.19,parc2a6:3.69,parc7a12:4.19,antecip:2.19},
  AmEx:      {debito:2.49,vista:3.49,parc2a6:3.99,parc7a12:4.49,antecip:2.49},
  __opts:    {antecip:false},
};
const PAYS=["Dinheiro","PIX","Transferência","Débito","Crédito à Vista","Crédito Parcelado"];
// Módulos liberáveis por usuário (operador/visualizador). Admin sempre vê tudo.
const PERM_MODULES=[
  {id:"dashboard",l:"📊 Dashboard"},
  {id:"vendas",l:"🛒 Vendas"},
  {id:"estoque",l:"📦 Estoque, Produtos e Pedidos"},
  {id:"financeiro",l:"💵 Financeiro (caixa, contas, taxas)"},
  {id:"clientes",l:"👥 Clientes"},
];

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

//  ÍCONES 
// Error Boundary - catches render crashes
class ErrorBoundary extends React.Component{
  constructor(p){super(p);this.state={err:null};}
  static getDerivedStateFromError(e){return{err:e};}
  componentDidCatch(e){console.error("App crash:",e);}
  render(){
    if(this.state.err){
      const s={padding:"2rem",background:"#0a0000",color:"#f56565",minHeight:"100vh",fontFamily:"monospace"};
      const ps={background:"#1a0000",padding:"1rem",borderRadius:".5rem",fontSize:".75rem",overflow:"auto",whiteSpace:"pre-wrap",maxHeight:"60vh"};
      const bs={marginTop:"1rem",padding:".5rem 1.2rem",background:"#4f5ef0",color:"#fff",border:"none",borderRadius:".4rem",cursor:"pointer"};
      return React.createElement("div",{style:s},
        React.createElement("h2",{style:{color:"#f59e0b",marginBottom:"1rem"}},"Erro — envie este detalhe:"),
        React.createElement("pre",{style:ps},String(this.state.err)+"  "+String(this.state.err&&this.state.err.stack||"")),
        React.createElement("button",{style:bs,onClick:()=>window.location.reload()},"Recarregar")
      );
    }
    return this.props.children;
  }
}


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

//  UI PRIMITIVES 
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
  return <button onClick={onClick} disabled={disabled} style={{display:"inline-flex",alignItems:"center",gap:".35rem",padding:sm?".34rem .7rem":".55rem 1rem",borderRadius:".6rem",fontSize:sm?".74rem":".84rem",fontFamily:"'DM Sans',sans-serif",fontWeight:700,letterSpacing:".01em",opacity:disabled?.5:1,width:full?"100%":undefined,justifyContent:full?"center":undefined,boxShadow:disabled?undefined:({p:"0 2px 10px rgba(79,94,240,.28)",ok:"0 2px 10px rgba(16,185,129,.26)",info:"0 2px 10px rgba(8,145,178,.26)"}[v]||undefined),...(S[v]||S.p)}}>{children}</button>;
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
    <div style={{background:"var(--bg3)",border:"1px solid var(--bdr2)",borderRadius:window.innerWidth<640?".65rem":"1rem",width:"100%",maxWidth:full?"99vw":wide?Math.min(700,window.innerWidth-16):Math.min(500,window.innerWidth-16),maxHeight:"93vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 64px var(--shadow)"}}>
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
const R2=({children,gap=".65rem",noStack})=>{const k=Array.isArray(children)?children.filter(Boolean):[children];const mob=!noStack&&window.innerWidth<640;return <div style={{display:"grid",gridTemplateColumns:mob?"1fr":`repeat(${k.length},1fr)`,gap}}>{children}</div>;};

// Botões Export
const XBtn=({rows,name,sheet})=><button onClick={()=>exportXLS(rows,name,sheet)} style={{display:"inline-flex",alignItems:"center",gap:".3rem",padding:".32rem .65rem",borderRadius:".4rem",background:"#14532d",color:"#4ade80",border:"1px solid #166834",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}><Ic n="xls" s={11}/>Excel</button>;
const exportOrderPDF=(order)=>{
  if(!window.jspdf){alert("Aguarde carregar a página.");return;}
  const{jsPDF}=window.jspdf;
  const items=parseItems(order);
  const doc=new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  const W=210,margin=14;
  const fmt2=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0);

  //  Status atual espelhado 
  const statusLabel=order.status==="recebido"?"RECEBIDO":order.status==="perdido"?"PERDIDO":order.status==="parcial"?"PARCIAL (em andamento)":"PENDENTE";
  const statusColor=order.status==="recebido"?[0,150,100]:order.status==="perdido"?[200,60,60]:order.status==="parcial"?[8,145,178]:[200,140,0];

  //  Cabeçalho 
  const showComp=typeof companyInfo!=="undefined"&&companyInfo.showInPDF!==false;
  const cname=showComp&&companyInfo.name?companyInfo.name:"CaixaPro · Tirzepatida";
  doc.setFillColor(13,15,26);
  doc.rect(0,0,W,22,"F");
  doc.setTextColor(232,234,246);
  doc.setFontSize(14);doc.setFont("helvetica","bold");
  doc.text(cname,margin,10);
  doc.setFontSize(9);doc.setFont("helvetica","normal");doc.setTextColor(160,160,200);
  doc.text("Pedido de Compra",margin,16);
  if(showComp&&companyInfo.phone)doc.text(companyInfo.phone,W-margin,16,{align:"right"});
  else doc.text(new Date().toLocaleString("pt-BR"),W-margin,16,{align:"right"});

  //  Badge de status 
  doc.setFillColor(...statusColor);
  doc.roundedRect(W-margin-50,26,50,10,2,2,"F");
  doc.setTextColor(255,255,255);
  doc.setFontSize(8);doc.setFont("helvetica","bold");
  doc.text(statusLabel,W-margin-25,32,{align:"center"});

  //  Info do pedido 
  doc.setFontSize(11);doc.setFont("helvetica","bold");doc.setTextColor(40,40,70);
  doc.text("PEDIDO #"+order.id.slice(0,8).toUpperCase(),margin,33);
  doc.setFontSize(9);doc.setFont("helvetica","normal");
  const infoRows=[
    ["Fornecedor:",order.supplier_name||"—"],
    ["Data do pedido:",order.order_date||"—"],
    ["Recebimento:",order.received_date||"Aguardando"],
  ];
  let y=42;
  infoRows.forEach(([l,v])=>{
    doc.setFont("helvetica","bold");doc.setTextColor(100,100,140);doc.text(l,margin,y);
    doc.setFont("helvetica","normal");doc.setTextColor(40,40,70);doc.text(v,55,y);
    y+=7;
  });
  if(order.notes){
    doc.setFont("helvetica","italic");doc.setTextColor(120,120,160);
    doc.text("Obs: "+order.notes,margin,y);y+=8;
  }

  //  Tabela de itens COM STATUS DE RECEBIMENTO 
  y+=3;
  doc.setFontSize(10);doc.setFont("helvetica","bold");doc.setTextColor(79,94,240);
  doc.text("ITENS DO PEDIDO",margin,y);y+=4;
  doc.autoTable({
    head:[["Status","Produto","Un","Qtd ped.","Qtd rec.","Custo/un","Total"]],
    body:items.map(i=>[
      i.received?"Recebido":"Pendente",
      i.product_name,
      i.unit||"un",
      String(i.qty),
      i.received?String(i.received_qty||i.qty):"—",
      fmt2(i.unit_cost),
      fmt2(i.total),
    ]),
    startY:y,
    styles:{fontSize:8,cellPadding:2.5,textColor:[40,40,70]},
    headStyles:{fillColor:[79,94,240],textColor:[255,255,255],fontStyle:"bold"},
    alternateRowStyles:{fillColor:[248,249,255]},
    columnStyles:{
      0:{halign:"center",cellWidth:22,fontStyle:"bold"},
      3:{halign:"center"},4:{halign:"center"},
      5:{halign:"right"},6:{halign:"right"},
    },
    didParseCell:(data)=>{
      if(data.column.index===0&&data.section==="body"){
        data.cell.styles.textColor=data.cell.raw==="Recebido"?[0,150,100]:[200,140,0];
      }
    },
    margin:{left:margin,right:margin},
  });
  y=doc.lastAutoTable.finalY+8;

  //  Resumo financeiro ATUALIZADO 
  const totalPaid=(order.initial_value||0)+(order.remaining_paid||0);
  const stillOwed=Math.max(0,(order.total_value||0)-totalPaid);
  doc.setFontSize(10);doc.setFont("helvetica","bold");doc.setTextColor(79,94,240);
  doc.text("RESUMO FINANCEIRO",margin,y);y+=4;
  doc.setFillColor(245,247,255);
  doc.roundedRect(margin,y-2,W-margin*2,48,2,2,"F");
  const finRows=[
    ["Total do pedido:",            fmt2(order.total_value),           [40,40,70]],
    ["Sinal pago ("+order.initial_pct+"%):", fmt2(order.initial_value),[200,140,0]],
    ["Pagamentos no recebimento:",  fmt2(order.remaining_paid||0),     [79,94,240]],
    ["Total pago até agora:",       fmt2(totalPaid),                   [0,150,100]],
    ["Saldo a pagar:",              fmt2(stillOwed),                   stillOwed>0?[200,60,60]:[0,150,100]],
  ];
  finRows.forEach(([l,v,color])=>{
    doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor(80,80,120);
    doc.text(l,margin+4,y+5);
    doc.setFont("helvetica","bold");doc.setFontSize(9);doc.setTextColor(...color);
    doc.text(v,W-margin-4,y+5,{align:"right"});
    y+=9;
  });

  //  Progresso itens recebidos 
  y+=6;
  const recCount=items.filter(i=>i.received).length;
  const totalCount=items.length;
  const pct=totalCount>0?Math.round((recCount/totalCount)*100):0;
  doc.setFontSize(9);doc.setFont("helvetica","bold");doc.setTextColor(80,80,120);
  doc.text("Progresso do recebimento: "+recCount+"/"+totalCount+" item(s) ("+pct+"%)",margin,y);y+=5;
  doc.setFillColor(220,220,240);doc.rect(margin,y,W-margin*2,4,"F");
  doc.setFillColor(...statusColor);doc.rect(margin,y,(W-margin*2)*(pct/100),4,"F");
  y+=12;

  //  Rodapé 
  doc.setFontSize(7.5);doc.setFont("helvetica","normal");doc.setTextColor(180,180,200);
  doc.line(margin,282,W-margin,282);
  const footerTxt=showComp&&companyInfo.name?companyInfo.name+" · "+new Date().toLocaleString("pt-BR"):"Documento gerado em "+new Date().toLocaleString("pt-BR");
  doc.text(footerTxt,W/2,287,{align:"center"});
  doc.save("pedido-"+order.id.slice(0,8).toUpperCase()+"-"+order.status.toUpperCase()+".pdf");
};

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

//  TELA LOGIN 
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
    <div id="app-root" style={{minHeight:"100vh",background:dark?"#080a14":"#f0f4ff",display:"flex",alignItems:"center",justifyContent:"center",padding:"1.25rem",position:"relative",overflow:"hidden"}}>
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

//  ANALYTICS 
function Analytics({onClose,sales,cashTx,products,clients,dark,receivables=[],orders=[],payables=[]}){
  // Local batchRevenue (same logic as App scope, needed here as Analytics is a separate component)
  const batchRevenue=(arr)=>{
    if(!arr||!arr.length)return 0;
    const b={};
    arr.forEach(s=>{const k=s.batch_id||s.id||"x";if(!b[k])b[k]={total:0,disc:parseFloat(s.discount)||0};b[k].total+=parseFloat(s.total_price)||0;});
    return Object.values(b).reduce((a,v)=>a+Math.max(0,v.total-v.disc),0);
  };
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
  const ticket=fS.length>0?batchRevenue(fS)/fS.length:0;
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
  // Contas em aberto (saldo atual, não filtrado por período) + saldo projetado
  const aReceberOpen=(receivables||[]).filter(r=>!r.paid).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
  const aPagarOpen=(payables||[]).filter(p=>!p.paid).reduce((a,p)=>a+Math.max(0,(parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0)),0)+(orders||[]).filter(o=>o.status!=="cancelado"&&o.status!=="perdido").reduce((a,o)=>a+Math.max(0,(parseFloat(o.remaining_value)||0)-(parseFloat(o.remaining_paid)||0)),0);
  const saldoCaixaTot=(cashTx||[]).reduce((a,t)=>a+(t.type==="entrada"?(parseFloat(t.value)||0):-(parseFloat(t.value)||0)),0);
  const saldoProj=saldoCaixaTot+aReceberOpen-aPagarOpen;
  // Prejuízo real de um pedido perdido = tudo que foi pago − valor das mercadorias recebidas (a custo).
  // As unidades recebidas ficam no estoque (recuperáveis na venda); a perda é só a parte sem contrapartida.
  const orderLoss=(o)=>{
    const paid=(parseFloat(o.initial_value)||0)+(parseFloat(o.remaining_paid)||0);
    let received=0;
    try{parseItems(o).forEach(it=>{received+=(parseFloat(it.received_qty)||0)*(parseFloat(it.unit_cost)||0);});}catch(e){}
    return Math.max(0,paid-received);
  };

  const RANGES=[{k:"1d",l:"Hoje"},{k:"7d",l:"7 dias"},{k:"30d",l:"30 dias"},{k:"3m",l:"3 meses"},{k:"6m",l:"6 meses"},{k:"12m",l:"12 meses"},{k:"custom",l:"Personalizado"}];
  const PCOLS=["#f59e0b","#9ca3af","#cd7c2f","#4f5ef0","#10b981","#8b44f0","#f56565","#0891b2"];

  return(
    <Modal title="Relatório Gerencial · Analytics" onClose={onClose} icon="analytics" full>
      {/* Filtros */}
      <div style={{display:"flex",gap:".35rem",flexWrap:"wrap",marginBottom:"1rem",alignItems:"center",overflowX:"auto",paddingBottom:".25rem"}}>
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".6rem",marginBottom:"1rem"}}>
        <KCard label="Receita" value={fmt(rev)} sub={`${fC.filter(x=>x.type==="entrada").length} lançamentos`} color="#10b981" icon="up"/>
        <KCard label="Custos" value={fmt(cost)} color="#f56565" icon="dn"/>
        <KCard label="Lucro líquido" value={fmt(profit)} sub={fmtPct(margin)+" margem"} color={profit>=0?"#4f5ef0":"#f56565"}/>
        <KCard label="Markup empresa" value={fmtPct(markup)} sub="lucro/custo" color="#8b44f0"/>
        <KCard label="Vendas" value={fmtN(fS.length)} sub={`${fmtN(units)} unidades`} color="#f59e0b" icon="sales"/>
        <KCard label="Ticket médio" value={fmt(ticket)} color="#4f5ef0"/>
        <KCard label="Valor estoque" value={fmt(stockVal)} color="#8b44f0" icon="stock"/>
        <KCard label="Clientes" value={clients.length} color="#10b981" icon="client"/>
        <KCard label="A Receber" value={fmt((receivables||[]).filter(r=>!r.paid).reduce((a,r)=>a+r.value,0))} sub={(receivables||[]).filter(r=>!r.paid).length+" pendente(s)"} color="#f59e0b" icon="dollar"/>
        <KCard label="A Pagar" value={fmt(aPagarOpen)} sub="contas + pedidos" color="#f56565" icon="dollar"/>
      </div>

      {/* Saldo projetado */}
      <div style={{background:saldoProj>=0?"linear-gradient(135deg,#10b98118,#10b98106)":"linear-gradient(135deg,#f5656518,#f5656506)",border:`1px solid ${saldoProj>=0?"#10b98140":"#f5656540"}`,borderRadius:".75rem",padding:".85rem 1rem",marginBottom:"1rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:".5rem"}}>
        <div>
          <div style={{fontSize:".66rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".04em"}}>Saldo projetado</div>
          <div style={{fontSize:".64rem",color:"var(--tx6)"}}>caixa {fmt(saldoCaixaTot)} + a receber {fmt(aReceberOpen)} − a pagar {fmt(aPagarOpen)}</div>
        </div>
        <div style={{fontSize:"1.5rem",fontWeight:800,fontFamily:"'Syne',sans-serif",color:saldoProj>=0?"#10b981":"#f56565"}}>{fmt(saldoProj)}</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".85rem",marginBottom:".85rem"}}>
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:".85rem",marginBottom:".85rem"}}>
        {[
          {title:"Saúde Financeira",color:"#8b44f0",icon:"dollar",rows:[
            {l:"ROI (lucro/custo)",v:cost>0?fmtPct((profit/cost)*100):"—",c:profit>=0?"#10b981":"#f56565"},
            {l:"Margem líquida",v:fmtPct(margin),c:"#4f5ef0"},
            {l:"Markup empresa",v:fmtPct(markup),c:"#8b44f0"},
            {l:"Pendências a receber",v:fmt((receivables||[]).filter(r=>!r.paid).reduce((a,r)=>a+r.value,0)),c:"#f59e0b"},
            {l:"Receita projetada",v:fmt(rev+(receivables||[]).filter(r=>!r.paid).reduce((a,r)=>a+r.value,0)),c:"#10b981"},
          ]},
          {title:"Estoque & Giro",color:"#f59e0b",icon:"pkg",rows:[
            {l:"Valor total estoque",v:fmt(stockVal),c:"#f59e0b"},
            {l:"Total unidades",v:fmtN(products.reduce((a,p)=>a+p.stock_qty,0))+" un",c:"var(--tx)"},
            {l:"Produtos cadastrados",v:products.length,c:"#4f5ef0"},
            {l:"Estoque zerado",v:products.filter(p=>p.stock_qty<=0).length+" item(s)",c:"#f56565"},
            {l:"Vencendo em 30d",v:products.filter(p=>p.expiry&&Math.ceil((new Date(p.expiry)-new Date())/86400000)<=30).length+" produto(s)",c:"#f59e0b"},
          ]},
          {title:"Performance Vendas",color:"#10b981",icon:"sales",rows:[
            {l:"Total de vendas",v:fS.length,c:"#10b981"},
            {l:"Unidades vendidas",v:fmtN(units)+" un",c:"var(--tx)"},
            {l:"Ticket médio",v:fmt(ticket),c:"#4f5ef0"},
            {l:"Clientes ativos",v:clients.length,c:"#8b44f0"},
            {l:"Parcelas a receber",v:fmtN(receivables.filter(r=>!r.paid&&r.description&&r.description.includes("Parcela")).length),c:"#8b44f0"},
          ]},
          {title:"Pedidos & Compras",color:"#f59e0b",icon:"pkg",rows:[
            {l:"Pedidos pendentes",v:(orders||[]).filter(o=>o.status==="pendente"||o.status==="parcial").length+" pedido(s)",c:"#f59e0b"},
            {l:"A pagar (restante)",v:fmt((orders||[]).filter(o=>o.status==="pendente").reduce((a,o)=>a+o.remaining_value,0)),c:"#f56565"},
            {l:"Pedidos recebidos",v:(orders||[]).filter(o=>o.status==="recebido").length+" pedido(s)",c:"#10b981"},
            {l:"Total investido",v:fmt((orders||[]).reduce((a,o)=>a+(o.initial_value||0)+(o.remaining_paid||0),0)),c:"#4f5ef0"},
            {l:"💀 Prejuízo (pago − recebido)",v:fmt((orders||[]).filter(o=>o.status==="perdido").reduce((a,o)=>a+orderLoss(o),0)),c:"#f56565"},
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


//  TELA DE NOTÍCIAS 
//  APP PRINCIPAL 
export default function App(){
  // Tema
  const[dark,setDark]=useState(getTheme);
  useEffect(()=>{document.body.classList.toggle("light",!dark);},[dark]);
  const toggle=()=>setDark(v=>{const nv=!v;setTheme(nv);return nv;});

  //  Responsividade 
  const[vw,setVw]=useState(()=>typeof window!=="undefined"?window.innerWidth:1024);
  useEffect(()=>{
    const h=()=>setVw(window.innerWidth);
    window.addEventListener("resize",h,{passive:true});
    return()=>window.removeEventListener("resize",h);
  },[]);
  const isMobile=vw<640;
  const isTablet=vw<900;
  // Helpers de responsividade
  const cols=(mobile,tablet,desktop)=>isMobile?mobile:isTablet?tablet:desktop;

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
  const[treatments,setTreatments]=useState([]);
  const[showTreatments,setShowTreatments]=useState(false);
  const[editingTreatment,setEditingTreatment]=useState(null);
  const[showCashFlow,setShowCashFlow]=useState(false);
  const[showRanking,setShowRanking]=useState(false);
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
  const[monthGoal,setMonthGoal]=useState("");
  const[showReceipt,setShowReceipt]=useState(null);
  const[showClientHist,setShowClientHist]=useState(null);
  const[showImportCalc,setShowImportCalc]=useState(false);
  const[importCalc,setImportCalc]=useState({totalCost:"",qty:"",extras:""});
  const[receivables,setReceivables]=useState([]);
  const[recForm,setRecForm]=useState({client_id:"",client_name:"",description:"",value:"",due_date:""});
  // Contas a pagar (Financeiro)
  const[payables,setPayables]=useState([]);
  const[payForm,setPayForm]=useState({description:"",value:"",category:"aluguel",supplier_name:"",due_date:"",recurring:false,installments:1});
  const[payTarget,setPayTarget]=useState(null); // parcela selecionada para registrar pagamento
  const[payAmount,setPayAmount]=useState("");
  const[payQuitar,setPayQuitar]=useState(false); // liquidar a conta com este pagamento (quita saldo restante)
  const[finTab,setFinTab]=useState("resumo"); // sub-guia da tela Financeiro
  const[estTab,setEstTab]=useState("estoque"); // sub-guia da tela Estoque
  //  Config global 
  const[configSection,setConfigSection]=useState("usuarios");
  const[dynCats,setDynCats]=useState(DEFAULT_CATS);
  const[dynPays,setDynPays]=useState([]);
  const[dynSettle,setDynSettle]=useState({});
  const[companyInfo,setCompanyInfo]=useState({name:"CaixaPro · Tirzepatida",phone:"",address:"",cnpj:"",obs:"",showInPDF:true});
  const[newCatName,setNewCatName]=useState("");
  const[newCatIcon,setNewCatIcon]=useState("📦");
  const[newPayName,setNewPayName]=useState("");
  const[localCI,setLocalCI]=useState(null);
  // Pagamentos ativos (dinâmicos ou default)
  const activePays=dynPays.length>0?dynPays:[...PAYS_SIMPLES,"Débito","Crédito à Vista","Crédito Parcelado"];
  const activeSimplePays=dynPays.length>0?dynPays.filter(p=>!["Débito","Crédito à Vista","Crédito Parcelado"].includes(p)):PAYS_SIMPLES;
  // Liquidação configurada (ou padrão) de cada forma de pagamento
  const getSettle=(method)=>dynSettle[method]||DEFAULT_SETTLE[method]||{type:"instant"};
  // Data (BR) em que o dinheiro entra no caixa conforme a liquidação da forma
  const settleDate=(method)=>{const s=getSettle(method);return s.type==="days"&&(parseInt(s.days)||0)>0?addBizDays(s.days):today();};
  const settleLabel=(s)=>s.type==="instant"?"Instantânea":s.type==="days"?("Após "+(parseInt(s.days)||1)+" dia"+((parseInt(s.days)||1)>1?"s":"")+" útil"+((parseInt(s.days)||1)>1?"eis":"")):s.type==="duedate"?"Promissória (data de vencimento)":s.type==="parcelado"?"Parcelado (mensal)":"Instantânea";
  const activeCats=dynCats.length>0?dynCats:DEFAULT_CATS;

  //  Pedidos 
  const[orders,setOrders]=useState([]);
  const[showOrderModal,setShowOrderModal]=useState(false);
  const[showReceiveModal,setShowReceiveModal]=useState(null); // order object
  const[editingOrder,setEditingOrder]=useState(null);
  const newOrderItem=()=>({key:uid(),product_id:"",product_name:"",unit:"un",qty:1,unit_cost:0});
  const[orderItems,setOrderItems]=useState([newOrderItem()]);
  const[orderSupplier,setOrderSupplier]=useState({id:"",name:""});
  const[orderPct,setOrderPct]=useState("50");
  const[orderNotes,setOrderNotes]=useState("");
  const[receiveExtra,setReceiveExtra]=useState(""); // extra payment on receipt
  const[receiveChecked,setReceiveChecked]=useState({});
  //  Frete 
  const defaultFreteConfig={
    origens:[
      {id:"1",name:"Ponto 1",address:"",lat:null,lon:null},
      {id:"2",name:"Ponto 2",address:"",lat:null,lon:null},
      {id:"3",name:"Ponto 3",address:"",lat:null,lon:null},
    ],
    base:5,ratePerKm:2.5,minFee:10,maxFee:0,
    zonas:[]
  };
  const[freteConfig,setFreteConfig]=useState(defaultFreteConfig);
  const[localFrete,setLocalFrete]=useState(null);
  const[showFreteConfig,setShowFreteConfig]=useState(false);
  const[freteOrigem,setFreteOrigem]=useState(()=>freteConfig?.origens?.find(o=>o.lat)?.id||"1");
  const[freteDestino,setFreteDestino]=useState("");
  const[zonaForm,setZonaForm]=useState({name:"",fee:""}); // zonas de entrega taxa fixa
  const[freteResult,setFreteResult]=useState(null);
  const[freteLoading,setFreteLoading]=useState(false);
  const[freteGeoList,setFreteGeoList]=useState([]);
  const[showFreteCalcInCart,setShowFreteCalcInCart]=useState(false);
  const[freteCalcDestino,setFreteCalcDestino]=useState("");
  const[freteCalcOrigem,setFreteCalcOrigem]=useState("1");
  const[freteCalcResult,setFreteCalcResult]=useState(null);
  const[freteCalcLoading,setFreteCalcLoading]=useState(false);
  const[freteCalcGeoList,setFreteCalcGeoList]=useState([]); // {itemIndex: {checked, qty}}
  const[receivePayment,setReceivePayment]=useState(""); // payment for selected items
  const[cartDiscount,setCartDiscount]=useState("");
  const[cartParcelas,setCartParcelas]=useState(2);
  const[cartCardBrand,setCartCardBrand]=useState("Visa");
  const[cardTaxes,setCardTaxes]=useState(DEFAULT_TAXES);
  const[showCardConfig,setShowCardConfig]=useState(false);
  const[showGoals,setShowGoals]=useState(false);
  const[monthlyGoals,setMonthlyGoals]=useState({});
  const[goalYear,setGoalYear]=useState(new Date().getFullYear());
  const[localTaxes,setLocalTaxes]=useState(()=>JSON.parse(JSON.stringify(DEFAULT_TAXES)));
  const updateLocalTax=(brand,mode,val)=>setLocalTaxes(t=>({...t,[brand]:{...t[brand],[mode]:val}}));
  const[cartDiscountType,setCartDiscountType]=useState("fixed");

  const toast$=(msg,color="#10b981")=>{setToast({msg,color});setTimeout(()=>setToast(null),3500);};

  // Forms defaults
  const FP={code:"",name:"",description:"",category:"tirzepatida",unit:"ampola",cost_per_unit:"",price_per_unit:"",units_per_pack:"1",batch:"",expiry:"",stock_qty:"0",min_stock:"5",supplier_id:""};
  //  CART STATE 
  const newCartItem=()=>({key:uid(),product_id:"",product_name:"",unit:"un",quantity:1,unit_price:0});
  const[cartItems,setCartItems]=useState([newCartItem()]);
  const[cartClient,setCartClient]=useState({id:"",name:""});
  const[cartIsDirect,setCartIsDirect]=useState(false);
  const[cartAutoOrder,setCartAutoOrder]=useState(true);// gerar pedido automático ao fornecedor na venda direta
  const[cartSupplierName,setCartSupplierName]=useState("");
  const[cartPayment,setCartPayment]=useState("PIX");
  const[cartDueDate,setCartDueDate]=useState("");// vencimento da venda promissória/fiado (ISO)
  const[cartNotes,setCartNotes]=useState("");
  const[cartFreight,setCartFreight]=useState("");
  const[cartDelivery,setCartDelivery]=useState(false);   // flag entregador solicitado
  const[cartDeliveryCost,setCartDeliveryCost]=useState(""); // custo pago ao entregador
  // PAGAMENTO DIVIDIDO (split) — mais de uma forma de pagamento na mesma venda
  const newSplit=()=>({key:uid(),method:"PIX",amount:"",brand:"Visa",parcelas:2});
  const[cartSplitEnabled,setCartSplitEnabled]=useState(false);
  const[cartSplits,setCartSplits]=useState([]);
  const updateSplit=(key,patch)=>setCartSplits(arr=>arr.map(s=>s.key===key?{...s,...patch}:s));
  const removeSplit=(key)=>setCartSplits(arr=>arr.filter(s=>s.key!==key));
  const addSplit=()=>setCartSplits(arr=>[...arr,newSplit()]);
  // legado (para editSale)
  const FS={product_id:"",product_name:"",client_id:"",client_name:"",quantity:"1",unit_price:"",total_price:"",notes:"",payment_method:"PIX"};
  const FC={name:"",email:"",phone:"",notes:"",credit_limit:""};
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
        supabase.from("app_users").select("id,username,display_name,role,active,last_login,created_at,permissions").order("created_at",{ascending:false}),
      ]);
      if(a.data)setProds(a.data);if(b.data)setSales(b.data);if(c.data)setCash(c.data);
      if(d.data)setClients(d.data);if(e.data)setSupp(e.data);if(f.data)setUsers(f.data);
      const{data:rv}=await supabase.from("receivables").select("*").order("due_date",{ascending:true});
      if(rv)setReceivables(rv);
      const{data:pyb}=await supabase.from("payables").select("*").order("due_date",{ascending:true});
      if(pyb)setPayables(pyb);
      const{data:od}=await supabase.from("orders").select("*").order("created_at",{ascending:false});
      if(od)setOrders(od);
      const{data:tr}=await supabase.from("treatments").select("*").order("next_purchase",{ascending:true});
      if(tr)setTreatments(tr);
      // Carregar configurações do sistema (taxas cartão + metas)
      const{data:settings}=await supabase.from("app_settings").select("*");
      if(settings&&settings.length>0){
        const taxRow=settings.find(s=>s.key==="cardtaxes");
        if(taxRow){try{setCardTaxes(JSON.parse(taxRow.value));}catch{}}
        const freteRow=settings.find(s=>s.key==="freteconfig");
        if(freteRow){try{setFreteConfig(JSON.parse(freteRow.value));}catch{}}
        const catsRow=settings.find(s=>s.key==="categories");
        if(catsRow){try{setDynCats(JSON.parse(catsRow.value));}catch{}}
        const paysRow=settings.find(s=>s.key==="payments");
        if(paysRow){try{setDynPays(JSON.parse(paysRow.value));}catch{}}
        const settleRow=settings.find(s=>s.key==="paysettle");
        if(settleRow){try{setDynSettle(JSON.parse(settleRow.value));}catch{}}
        const compRow=settings.find(s=>s.key==="companyinfo");
        if(compRow){try{setCompanyInfo(JSON.parse(compRow.value));}catch{}}
        const goals={};
        settings.filter(s=>s.key.startsWith("goal_")).forEach(s=>{
          goals[s.key.replace("goal_","")]=parseFloat(s.value)||0;
        });
        if(Object.keys(goals).length>0)setMonthlyGoals(goals);
        const mgRow=settings.find(s=>s.key==="monthgoal");
        if(mgRow)setMonthGoal(mgRow.value||"");
      }
      setLastSync(new Date().toLocaleTimeString("pt-BR"));
    }catch(e){console.error(e);}
    setSyncing(false);
  },[cu]);

  useEffect(()=>{
    if(!cu)return;load();
    const tbls=["products","sales","cash_transactions","clients","suppliers","app_users","receivables","payables","app_settings","orders","treatments"];
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
  // Helper: calcula receita real descontando desconto UMA vez por batch
  const batchRevenue=(arr)=>{
    if(!arr||!arr.length)return 0;
    const b={};
    arr.forEach(s=>{
      const k=s.batch_id||s.id||Math.random().toString();
      if(!b[k])b[k]={total:0,disc:parseFloat(s.discount)||0};
      b[k].total+=parseFloat(s.total_price)||0;
    });
    return Object.values(b).reduce((a,v)=>a+Math.max(0,v.total-v.disc),0);
  };
  const totalSalesRev=useMemo(()=>batchRevenue(sales),[sales]);
  const totalUnits=useMemo(()=>sales.reduce((a,s)=>a+s.quantity,0),[sales]);

  const isAdmin=cu?.role==="admin";
  const canEdit=cu?.role==="admin"||cu?.role==="operator";
  // Pode excluir registros: admin sempre; operador só se liberado nas permissões
  const canDelete=isAdmin||!!(cu?.permissions&&cu.permissions.candelete);
  // Pedidos a fornecedor (pendente/parcial) refletidos como contas a pagar futuras
  const orderPayables=useMemo(()=>(orders||[]).filter(o=>o.status!=="cancelado"&&o.status!=="perdido"&&((parseFloat(o.remaining_value)||0)-(parseFloat(o.remaining_paid)||0))>0.005).map(o=>{
    let drop=false;try{drop=parseItems(o).some(i=>i.dropship);}catch(e){}
    return {
      id:"ord_"+o.id,_order:true,_orderId:o.id,
      description:(drop?"🔄 Venda direta · ":"📦 Pedido · ")+(o.supplier_name||"fornecedor"),
      value:parseFloat(o.remaining_value)||0,        // total a pagar do pedido (após sinal)
      paid_amount:parseFloat(o.remaining_paid)||0,   // já pago do restante (acumulado)
      category:"fornecedor",supplier_name:o.supplier_name||null,
      due_date:o.expected_date||null,paid:false,_dropship:drop,
    };
  }),[orders]);

  // Detecta pedidos antigos com recebimento parcial gravado como concluído (para mostrar o botão de correção)
  const needsPartialFix=useMemo(()=>(orders||[]).some(o=>{try{return parseItems(o).some(it=>it.received===true&&(parseFloat(it.received_qty)||0)>0&&(parseFloat(it.received_qty)||0)<(parseFloat(it.qty)||0));}catch(e){return false;}}),[orders]);

  // Sale handlers
  //  CART HELPERS 
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
  const cartReset=()=>{setCartItems([newCartItem()]);setCartClient({id:"",name:""});setCartIsDirect(false);setCartAutoOrder(true);setCartSupplierName("");setCartPayment("PIX");setCartDueDate("");setCartNotes("");setCartFreight("");setCartDelivery(false);setCartDeliveryCost("");setCartDiscount("");setCartDiscountType("fixed");setCartParcelas(2);setCartCardBrand("Visa");setCartSplitEnabled(false);setCartSplits([]);};

  const cartSubtotal=cartItems.reduce((a,i)=>a+i.unit_price*i.quantity,0);
  const cartFreightVal=parseFloat(cartFreight)||0;
  const cartDeliveryCostVal=cartDelivery?(parseFloat(cartDeliveryCost)||0):0;
  const cartDiscountVal=cartDiscountType==="percent"?cartSubtotal*(parseFloat(cartDiscount)||0)/100:(parseFloat(cartDiscount)||0);
  const cartTotal=Math.max(0,cartSubtotal-cartDiscountVal)+cartFreightVal;
  // Soma e validação do pagamento dividido
  const cartSplitSum=Math.round(cartSplits.reduce((a,s)=>a+(parseFloat(s.amount)||0),0)*100)/100;
  const cartSplitRemaining=Math.round((cartTotal-cartSplitSum)*100)/100;
  const cartSplitValid=cartSplitEnabled?(cartSplits.length>0&&cartSplits.every(s=>s.method&&(parseFloat(s.amount)||0)>0)&&Math.abs(cartSplitRemaining)<0.01):true;

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
    const discount=cartDiscountType==="percent"?subtotal*(parseFloat(cartDiscount)||0)/100:(parseFloat(cartDiscount)||0);
    const grandTotal=Math.max(0,subtotal-discount)+freight;
    const desc=validItems.map(i=>`${i.product_name}(${i.quantity})`).join(", ");

    // Validação do pagamento dividido
    let splitMethodLabel=null;
    if(cartSplitEnabled){
      const activeSplits=cartSplits.filter(s=>(parseFloat(s.amount)||0)>0);
      if(activeSplits.length===0){toast$("Adicione ao menos uma forma de pagamento com valor.","#f56565");return;}
      const sum=Math.round(cartSplits.reduce((a,s)=>a+(parseFloat(s.amount)||0),0)*100)/100;
      if(Math.abs(Math.round((grandTotal-sum)*100)/100)>=0.01){
        toast$("A soma dos pagamentos ("+fmt(sum)+") deve ser igual ao total ("+fmt(grandTotal)+").","#f56565");return;
      }
      splitMethodLabel="Dividido: "+activeSplits.map(s=>{const c=["Débito","Crédito à Vista","Crédito Parcelado"].includes(s.method);return c?(s.method==="Crédito Parcelado"?"Crédito "+(parseInt(s.parcelas)||2)+"x":s.method):s.method;}).join(" + ");
    }

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
        payment_method:cartSplitEnabled?splitMethodLabel:cartPayment,
        added_by:cu.display_name,
        date:today(),
        // batch_id para agrupar itens da mesma venda
        batch_id:batchId,is_direct:cartIsDirect,direct_status:cartIsDirect?"Encomenda":null,supplier_name:cartIsDirect?cartSupplierName:null,
      }));

      const{error:e1}=await supabase.from("sales").insert(saleRecords);
      if(e1){toast$("Erro ao registrar venda: "+e1.message,"#f56565");return;}

      // 2. Baixar estoque de cada produto (PULA se for venda direta)
      if(!cartIsDirect){
        for(const item of validItems){
          await deductStock(item.product_id,item.quantity);
        }
      }

      // 3. Montar lançamentos de caixa
      const cashInserts=[];

      // Separar itens pagos e cortesia
      const paidItems=validItems.filter(i=>i.unit_price>0);
      const freeItems=validItems.filter(i=>i.unit_price===0);

      // Calcular taxa de cartão
      const isCard=["Débito","Crédito à Vista","Crédito Parcelado"].includes(cartPayment);
      const cardMode=cartPayment==="Débito"?"debito":cartPayment==="Crédito à Vista"?"vista":cartParcelas<=6?"parc2a6":"parc7a12";
      const taxRate=isCard?getCardTaxRate(cartCardBrand,cardMode):0;
      const taxAmount=grandTotal*(taxRate/100);
      const netTotal=grandTotal-taxAmount;
      const payLabel=isCard?(cartPayment==="Crédito Parcelado"?cartPayment+" "+cartParcelas+"x · "+cartCardBrand:cartPayment+" · "+cartCardBrand):cartPayment;

      // 3a. Para parcelado: NÃO lança no caixa agora (será lançado ao marcar parcelas pagas)
      //     Para débito/crédito à vista: lança entrada líquida + saída da taxa
      if(cartSplitEnabled&&grandTotal>0){
        // PAGAMENTO DIVIDIDO: processa cada forma de pagamento separadamente
        for(const sp of cartSplits){
          const amt=parseFloat(sp.amount)||0;
          if(amt<=0)continue;
          const spIsCard=["Débito","Crédito à Vista","Crédito Parcelado"].includes(sp.method);
          const spMode=sp.method==="Débito"?"debito":sp.method==="Crédito à Vista"?"vista":((parseInt(sp.parcelas)||2)<=6?"parc2a6":"parc7a12");
          const spTaxRate=spIsCard?getCardTaxRate(sp.brand,spMode):0;
          const spTax=Math.round(amt*(spTaxRate/100)*100)/100;
          const spNet=Math.round((amt-spTax)*100)/100;
          const spLabel=spIsCard?(sp.method==="Crédito Parcelado"?sp.method+" "+(parseInt(sp.parcelas)||2)+"x · "+sp.brand:sp.method+" · "+sp.brand):sp.method;
          if(sp.method==="Crédito Parcelado"){
            // Parte parcelada → cria parcelas em A Receber (valor líquido)
            const nP=parseInt(sp.parcelas)||2;
            const pNet=Math.round((spNet/nP)*100)/100;
            const recIns=[];
            for(let p=0;p<nP;p++){
              const dd=new Date();dd.setMonth(dd.getMonth()+p+1);
              recIns.push({id:uid(),client_id:clientId,client_name:clientName,description:"Parcela "+(p+1)+"/"+nP+" · "+sp.brand+" · "+desc,value:pNet,due_date:dd.toISOString().slice(0,10),paid:false,sale_id:batchId,added_by:cu.display_name,created_at:nowISO()});
            }
            await supabase.from("receivables").insert(recIns);
          } else if(getSettle(sp.method).type==="duedate"){
            // Parte fiado/promissória → recebível com vencimento escolhido
            const dd=cartDueDate||addBizDaysISO(30);
            await supabase.from("receivables").insert([{id:uid(),client_id:clientId,client_name:clientName,description:"📝 "+sp.method+" · "+desc,value:amt,due_date:dd,paid:false,sale_id:batchId,added_by:cu.display_name,created_at:nowISO()}]);
          } else {
            // PIX / Dinheiro / Transferência / Débito / Crédito à Vista → entra no caixa (na data de liquidação)
            if(spNet>0)cashInserts.push({id:uid(),description:"Venda"+(clientName?" · "+clientName:"")+" · "+spLabel+" · "+desc,value:spNet,type:"entrada",category:"Venda",sale_id:batchId,product_name:validItems.map(i=>i.product_name).join(", "),added_by:cu.display_name,date:settleDate(sp.method)});
            if(spIsCard&&spTax>0)cashInserts.push({id:uid(),description:"Taxa "+spLabel+(clientName?" · "+clientName:""),value:spTax,type:"saida",category:"Taxa Cartão",sale_id:batchId,product_name:null,added_by:cu.display_name,date:settleDate(sp.method)});
          }
        }
        await loadReceivables();
      } else if(cartPayment==="Crédito Parcelado"&&grandTotal>0){
        // Cria recebíveis com valor líquido por parcela
        const nParcelas=parseInt(cartParcelas)||2;
        const parcelaNet=Math.round((netTotal/nParcelas)*100)/100;
        const recInserts=[];
        for(let p=0;p<nParcelas;p++){
          const dueDate=new Date();
          dueDate.setMonth(dueDate.getMonth()+p+1);
          recInserts.push({
            id:uid(),
            client_id:clientId,
            client_name:clientName,
            description:"Parcela "+(p+1)+"/"+nParcelas+" · "+cartCardBrand+" · "+desc,
            value:parcelaNet,
            due_date:dueDate.toISOString().slice(0,10),
            paid:false,
            sale_id:batchId,   // vincula ao batch para reversão
            added_by:cu.display_name,
            created_at:nowISO(),
          });
        }
        await supabase.from("receivables").insert(recInserts);
        await loadReceivables();
      } else if(getSettle(cartPayment).type==="duedate"&&grandTotal>0){
        // Venda fiado/promissória: cria 1 recebível com o vencimento escolhido (não entra no caixa agora)
        const dd=cartDueDate||addBizDaysISO(30);
        await supabase.from("receivables").insert([{id:uid(),client_id:clientId,client_name:clientName,description:"📝 "+cartPayment+" · "+desc,value:grandTotal,due_date:dd,paid:false,sale_id:batchId,added_by:cu.display_name,created_at:nowISO()}]);
        await loadReceivables();
      } else if(grandTotal>0){
        // Débito/Crédito à vista ou outros: lança no caixa
        if(netTotal>0){
          cashInserts.push({
            id:uid(),
            description:"Venda"+(clientName?" · "+clientName:"")+(freight>0?" (c/ frete)":"")+" · "+desc+" · "+payLabel,
            value:netTotal,
            type:"entrada",
            category:"Venda",
            sale_id:batchId,
            product_name:validItems.map(i=>i.product_name).join(", "),
            added_by:cu.display_name,
            date:settleDate(cartPayment),
          });
        }
        // Lança taxa do cartão como saída
        if(isCard&&taxAmount>0){
          cashInserts.push({
            id:uid(),
            description:"Taxa "+payLabel+(clientName?" · "+clientName:""),
            value:taxAmount,
            type:"saida",
            category:"Taxa Cartão",
            sale_id:batchId,
            product_name:null,
            added_by:cu.display_name,
            date:settleDate(cartPayment),
          });
        }
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

      // Pedido automático ao fornecedor (dropshipping) — só quando a venda é direta
      let autoOrderMsg="";
      if(cartIsDirect&&cartAutoOrder&&cartSupplierName.trim()&&validItems.length>0){
        try{
          const ordItems=validItems.map(it=>{
            const prod=products.find(p=>p.id===it.product_id);
            const cost=prod?parseFloat(prod.cost_per_unit)||0:0;
            const q=parseInt(it.quantity)||1;
            return {product_id:it.product_id,product_name:it.product_name,unit:it.unit||"un",qty:q,unit_cost:cost,total:cost*q,received:false,received_qty:0,dropship:true};
          });
          const ordTotal=Math.round(ordItems.reduce((a,i)=>a+i.total,0)*100)/100;
          const sup=suppliers.find(s=>s.name&&s.name.toLowerCase()===cartSupplierName.trim().toLowerCase());
          await supabase.from("orders").insert([{
            id:uid(),supplier_id:sup?sup.id:null,supplier_name:cartSupplierName.trim(),status:"pendente",
            items:JSON.stringify(ordItems),total_value:ordTotal,initial_pct:0,initial_value:0,remaining_value:ordTotal,remaining_paid:0,
            notes:"🔄 Dropshipping · gerado da venda direta "+batchId.slice(0,8).toUpperCase(),
            origin_sale_id:batchId,
            order_date:today(),added_by:cu.display_name,created_at:nowISO(),
          }]);
          await loadOrders();
          autoOrderMsg=" · 🔄 Pedido a "+cartSupplierName.trim()+" gerado ("+fmt(ordTotal)+")";
        }catch(ordEx){autoOrderMsg=" · ⚠️ venda ok, mas falhou gerar o pedido automático";}
      }

      const freeCount=freeItems.length;
      const paidCount=paidItems.length;
      const deliveryInfo=(cartDelivery&&cartDeliveryCostVal>0?" · 🛵 Entregador: "+fmt(cartDeliveryCostVal):"")+autoOrderMsg;

      if(cartSplitEnabled){
        const nForms=cartSplits.filter(s=>(parseFloat(s.amount)||0)>0).length;
        toast$("✅ Venda registrada! Pagamento dividido em "+nForms+" forma"+(nForms>1?"s":"")+" · Total: "+fmt(grandTotal)+deliveryInfo);
      } else if(getSettle(cartPayment).type==="duedate"){
        const dd=cartDueDate||addBizDaysISO(30);const pt=dd.split("-");
        toast$("📝 Venda "+cartPayment+" lançada em A Receber · "+fmt(grandTotal)+" · vence "+(pt[2]+"/"+pt[1]+"/"+pt[0])+deliveryInfo);
      } else if(cartPayment!=="Crédito Parcelado"){
        const msg=freeCount>0
          ?"✅ Venda registrada! "+paidCount+" pago(s) · "+freeCount+" cortesia · Total: "+fmt(grandTotal)+deliveryInfo
          :"✅ Venda registrada! "+validItems.length+" produto(s) · Total: "+fmt(grandTotal)+deliveryInfo;
        toast$(msg);
      } else {
        const nParcelas=parseInt(cartParcelas)||2;
        toast$("✅ "+cartCardBrand+" "+nParcelas+"x · "+nParcelas+" parcelas de "+fmt(cartInstallmentNet)+" (líq.) em A Receber!"+autoOrderMsg);
      }
      setModal(null);
      cartReset();
    }catch(ex){toast$("Erro de conexão: "+ex.message,"#f56565");}
  };

  //  RECEBÍVEIS 
  //  CARD TAX HELPERS 
  const getCardTaxRate=(brand,mode)=>{
    const t=cardTaxes[brand]||DEFAULT_TAXES[brand]||{};
    let r=parseFloat(t[mode])||0;
    const opts=cardTaxes.__opts||DEFAULT_TAXES.__opts||{};
    if(opts.antecip&&mode!=="debito")r+=parseFloat(t.antecip)||0;
    return r;
  };
  const saveCardTaxes=async(newTaxes)=>{
    setCardTaxes(newTaxes);
    try{
      await supabase.from("app_settings").upsert({
        key:"cardtaxes",
        value:JSON.stringify(newTaxes),
        updated_at:nowISO(),
        updated_by:cu.display_name
      },{onConflict:"key"});
      toast$("✅ Taxas de cartão salvas na nuvem!");
    }catch(e){toast$("Erro ao salvar taxas: "+e.message,"#f56565");}
  };
  const saveMonthlyGoal=async(yearMonth,value)=>{
    const numVal=parseFloat(value)||0;
    setMonthlyGoals(prev=>({...prev,[yearMonth]:numVal}));
    try{
      await supabase.from("app_settings").upsert({
        key:"goal_"+yearMonth,
        value:String(numVal),
        updated_at:nowISO(),
        updated_by:cu.display_name
      },{onConflict:"key"});
    }catch(e){console.error("Erro ao salvar meta:",e);}
  };

  // Computed — current cart card tax
  const cartIsCard=["Débito","Crédito à Vista","Crédito Parcelado"].includes(cartPayment);
  const cartCardMode=cartPayment==="Débito"?"debito":cartPayment==="Crédito à Vista"?"vista":cartParcelas<=6?"parc2a6":"parc7a12";
  const cartTaxRate=cartIsCard?getCardTaxRate(cartCardBrand,cartCardMode):0;
  const cartTaxAmount=cartTotal*(cartTaxRate/100);
  const cartNetTotal=cartTotal-cartTaxAmount;
  const cartInstallmentNet=cartPayment==="Crédito Parcelado"&&cartParcelas>0?cartNetTotal/cartParcelas:0;


  //  PEDIDOS HELPERS 
  const loadOrders=async()=>{
    const{data}=await supabase.from("orders").select("*").order("created_at",{ascending:false});
    if(data)setOrders(data);
  };
  const orderReset=()=>{
    setOrderItems([newOrderItem()]);
    setOrderSupplier({id:"",name:""});
    setOrderPct("50");
    setOrderNotes("");
    setReceiveExtra("");
  };
  const orderSetProduct=(key,pid)=>{
    const p=products.find(x=>x.id===pid);
    setOrderItems(items=>items.map(i=>i.key!==key?i:{
      ...i,product_id:pid,
      product_name:p?p.name:"",
      unit:p?p.unit||"un":"un",
      unit_cost:p?parseFloat(p.cost_per_unit)||0:0,
    }));
  };
  const createOrder=async()=>{
    const valid=orderItems.filter(i=>i.product_id&&i.qty>0&&i.unit_cost>0);
    if(valid.length===0){toast$("Adicione pelo menos 1 produto com custo.","#f56565");return;}
    if(!orderSupplier.id){toast$("Selecione o fornecedor.","#f56565");return;}
    const total=valid.reduce((a,i)=>a+i.unit_cost*i.qty,0);
    const pct=Math.min(100,Math.max(0,parseFloat(orderPct)||0));
    const initialVal=Math.round(total*(pct/100)*100)/100;
    const remainingVal=Math.round((total-initialVal)*100)/100;
    const orderId=uid();
    const itemsJson=JSON.stringify(valid.map(i=>({product_id:i.product_id,product_name:i.product_name,unit:i.unit,qty:parseInt(i.qty),unit_cost:parseFloat(i.unit_cost),total:parseFloat(i.unit_cost)*parseInt(i.qty)})));
    try{
      const{error:e}=await supabase.from("orders").insert([{
        id:orderId,
        supplier_id:orderSupplier.id,
        supplier_name:orderSupplier.name,
        status:"pendente",
        items:itemsJson,
        total_value:total,
        initial_pct:pct,
        initial_value:initialVal,
        remaining_value:remainingVal,
        remaining_paid:0,
        notes:orderNotes||null,
        order_date:today(),
        added_by:cu.display_name,
        created_at:nowISO(),
      }]);
      if(e){toast$("Erro ao criar pedido: "+e.message,"#f56565");return;}
      // Lança sinal no caixa (se houver)
      if(initialVal>0){
        await supabase.from("cash_transactions").insert([{
          id:uid(),
          description:"Sinal Pedido "+orderId.slice(0,8).toUpperCase()+" · "+orderSupplier.name+" ("+pct+"% de "+fmt(total)+")",
          value:initialVal,
          type:"saida",
          category:"Pedido Compra",
          sale_id:orderId,
          product_name:valid.map(i=>i.product_name).join(", "),
          added_by:cu.display_name,
          date:today(),
        }]);
      }
      toast$("✅ Pedido criado! Sinal: "+fmt(initialVal)+" · Restante: "+fmt(remainingVal));
      setShowOrderModal(false);
      orderReset();
    }catch(ex){toast$("Erro: "+ex.message,"#f56565");}
  };

  const confirmReceive=async(order,checkedItems,payment)=>{
    // checkedItems: [{index, qty, item}] — itens selecionados com qtd a receber
    if(!checkedItems||checkedItems.length===0){
      toast$("Selecione pelo menos 1 produto para receber.","#f56565");return;
    }
    const payVal=parseFloat(payment)||0;
    const allItems=parseItems(order);
    try{
      // 1. Dar entrada no estoque de cada item selecionado
      const receivedNames=[];
      for(const {index,qty,item} of checkedItems){
        const prod=products.find(p=>p.id===item.product_id||p.name===item.product_name);
        if(prod&&qty>0&&!item.dropship){
          // Adiciona ao estoque a quantidade que chegou agora (dropshipping NÃO entra no estoque)
          await restoreStock(prod.id,qty);
        }
        if(qty>0)receivedNames.push(item.product_name+" ("+qty+")"+(item.dropship?" 🔄 direto":""));
        // Acumula a quantidade recebida; só conclui a linha quando chega tudo
        const ordered=parseFloat(allItems[index].qty)||0;
        const prevRecv=parseFloat(allItems[index].received_qty)||0;
        const newRecv=prevRecv+qty;
        const fully=newRecv>=ordered;
        allItems[index]={...allItems[index],received:fully,received_qty:newRecv,received_date:fully?today():(allItems[index].received_date||null)};
      }
      // 2. Verificar se TODOS os itens foram recebidos
      const allReceived=allItems.every(i=>i.received);
      const newStatus=allReceived?"recebido":"parcial";
      // 3. Atualizar pedido
      const prevPaid=order.remaining_paid||0;
      await supabase.from("orders").update({
        status:newStatus,
        received_date:allReceived?today():(order.received_date||null),
        remaining_paid:prevPaid+payVal,
        items:JSON.stringify(allItems),
      }).eq("id",order.id);
      // 4. Lançar pagamento no caixa (se houver)
      if(payVal>0){
        await supabase.from("cash_transactions").insert([{
          id:uid(),
          description:"Pgto Pedido "+order.id.slice(0,8).toUpperCase()+" · "+order.supplier_name+" · "+receivedNames.join(", "),
          value:payVal,
          type:"saida",
          category:"Pedido Compra",
          sale_id:order.id,
          product_name:receivedNames.join(", "),
          added_by:cu.display_name,
          date:today(),
        }]);
      }
      const msg=allReceived
        ?"✅ Pedido totalmente recebido! Estoque e caixa atualizados."
        :"✅ "+checkedItems.length+" produto(s) recebido(s)! Pedido parcialmente concluído.";
      toast$(msg);
      setShowReceiveModal(null);
      setReceiveChecked({});
      setReceivePayment("");
    }catch(ex){toast$("Erro: "+ex.message,"#f56565");}
  };


  //  FRETE HELPERS 
  const saveToSettings=async(key,value,label="Configuração")=>{
    try{
      await supabase.from("app_settings").upsert({
        key,value:JSON.stringify(value),updated_at:nowISO(),updated_by:cu.display_name
      },{onConflict:"key"});
      toast$("✅ "+label+" salva!");
    }catch(e){toast$("Erro: "+e.message,"#f56565");}
  };

  //  ESTOQUE SIMPLES 
  const deductStock=async(productId,qty)=>{
    if(!productId||qty<=0)return;
    const prod=products.find(p=>p.id===productId);
    if(!prod)return;
    const newQty=Math.max(0,Math.round(((prod.stock_qty||0)-qty)*100)/100);
    await supabase.from("products").update({stock_qty:newQty}).eq("id",productId);
    setProds(prev=>prev.map(p=>p.id===productId?{...p,stock_qty:newQty}:p));
  };

  const restoreStock=async(productId,qty)=>{
    if(!productId||qty<=0)return;
    const{data}=await supabase.from("products").select("stock_qty").eq("id",productId).single();
    const newQty=Math.round(((parseFloat(data?.stock_qty)||0)+qty)*100)/100;
    await supabase.from("products").update({stock_qty:newQty}).eq("id",productId);
    setProds(prev=>prev.map(p=>p.id===productId?{...p,stock_qty:newQty}:p));
  };

  const saveCategories=async(cats)=>{setDynCats(cats);await saveToSettings("categories",cats,"Categorias");};
  const savePayments=async(pays)=>{setDynPays(pays);await saveToSettings("payments",pays,"Formas de pagamento");};
  const saveSettle=async(obj)=>{setDynSettle(obj);await saveToSettings("paysettle",obj,"Liquidação de pagamentos");};
  const setMethodSettle=(method,patch)=>{const cur=dynSettle[method]||DEFAULT_SETTLE[method]||{type:"instant"};saveSettle({...dynSettle,[method]:{...cur,...patch}});};
  const saveCompanyInfo=async(info)=>{setCompanyInfo(info);await saveToSettings("companyinfo",info,"Informações da empresa");};

  const saveFreteConfig=async(cfg)=>{
    try{
      await supabase.from("app_settings").upsert({
        key:"freteconfig",value:JSON.stringify(cfg),
        updated_at:nowISO(),updated_by:cu.display_name
      },{onConflict:"key"});
      setFreteConfig(cfg);
      toast$("✅ Configurações de frete salvas na nuvem!");
    }catch(e){toast$("Erro ao salvar: "+e.message,"#f56565");}
  };
  const addZona=()=>{
    if(!zonaForm.name||!zonaForm.fee){toast$("Informe o bairro/região e o valor.","#f56565");return;}
    const cfg={...freteConfig,zonas:[...(freteConfig.zonas||[]),{id:uid(),name:zonaForm.name,fee:parseFloat(zonaForm.fee)||0}]};
    setFreteConfig(cfg);saveFreteConfig(cfg);setZonaForm({name:"",fee:""});
  };
  const delZona=(id)=>{
    const cfg={...freteConfig,zonas:(freteConfig.zonas||[]).filter(z=>z.id!==id)};
    setFreteConfig(cfg);saveFreteConfig(cfg);
  };

  const geocodeAddr=async(address)=>{
    const url="https://nominatim.openstreetmap.org/search?q="+encodeURIComponent(address+" Brasil")+"&format=json&limit=5&accept-language=pt-BR";
    const res=await fetch(url);
    const data=await res.json();
    return data.map(d=>({display:d.display_name,lat:parseFloat(d.lat),lon:parseFloat(d.lon)}));
  };

  const calcFreteWithCoords=async(origem,dest)=>{
    const url="https://router.project-osrm.org/route/v1/driving/"+origem.lon+","+origem.lat+";"+dest.lon+","+dest.lat+"?overview=false";
    const res=await fetch(url);
    const data=await res.json();
    if(data.code!=="Ok")throw new Error("Rota não encontrada. Verifique os endereços.");
    const distKm=data.routes[0].distance/1000;
    const durMin=Math.round(data.routes[0].duration/60);
    const base=parseFloat(freteConfig.base)||0;
    const rate=parseFloat(freteConfig.ratePerKm)||0;
    const minF=parseFloat(freteConfig.minFee)||0;
    const maxF=parseFloat(freteConfig.maxFee)||0;
    let fee=base+distKm*rate;
    if(minF>0)fee=Math.max(fee,minF);
    if(maxF>0)fee=Math.min(fee,maxF);
    fee=Math.round(fee*100)/100;
    setFreteResult({distKm:distKm.toFixed(1),durMin,fee,destName:dest.display,origName:origem.name,calcDetail:{base,rate,distKm:distKm.toFixed(1),minF,maxF}});
    setFreteGeoList([]);
    setFreteLoading(false);
  };

  const calcFrete=async()=>{
    if(freteLoading)return;
    setFreteLoading(true);setFreteResult(null);setFreteGeoList([]);
    try{
      const origem=freteConfig.origens.find(o=>o.id===freteOrigem);
      if(!origem||!origem.lat){toast$("Configure e localize o ponto de origem antes.","#f56565");setFreteLoading(false);return;}
      if(!freteDestino.trim()){toast$("Informe o endereço de destino.","#f56565");setFreteLoading(false);return;}
      const results=await geocodeAddr(freteDestino);
      if(results.length===0){toast$("Endereço não encontrado. Inclua cidade/bairro.","#f56565");setFreteLoading(false);return;}
      if(results.length===1){await calcFreteWithCoords(origem,results[0]);return;}
      setFreteGeoList(results);setFreteLoading(false);
    }catch(e){toast$(e.message,"#f56565");setFreteLoading(false);}
  };

  const calcFreteInCart=async(destino,origemId)=>{
    setFreteCalcLoading(true);setFreteCalcResult(null);setFreteCalcGeoList([]);
    try{
      const origem=freteConfig.origens.find(o=>o.id===origemId);
      if(!origem||!origem.lat){toast$("Configure o ponto de origem na aba Frete primeiro.","#f56565");setFreteCalcLoading(false);return;}
      if(!destino.trim()){toast$("Informe o endereço de destino.","#f56565");setFreteCalcLoading(false);return;}
      const results=await geocodeAddr(destino);
      if(results.length===0){toast$("Endereço não encontrado. Inclua cidade/bairro.","#f56565");setFreteCalcLoading(false);return;}
      if(results.length>1){setFreteCalcGeoList(results);setFreteCalcLoading(false);return;}
      await calcFreteInCartWithCoords(origem,results[0]);
    }catch(e){toast$(e.message,"#f56565");setFreteCalcLoading(false);}
  };

  const calcFreteInCartWithCoords=async(origem,dest)=>{
    const url="https://router.project-osrm.org/route/v1/driving/"+origem.lon+","+origem.lat+";"+dest.lon+","+dest.lat+"?overview=false";
    const res=await fetch(url);
    const data=await res.json();
    if(data.code!=="Ok")throw new Error("Rota não encontrada.");
    const distKm=data.routes[0].distance/1000;
    const durMin=Math.round(data.routes[0].duration/60);
    const base=parseFloat(freteConfig.base)||0;
    const rate=parseFloat(freteConfig.ratePerKm)||0;
    const minF=parseFloat(freteConfig.minFee)||0;
    const maxF=parseFloat(freteConfig.maxFee)||0;
    let fee=base+distKm*rate;
    if(minF>0)fee=Math.max(fee,minF);
    if(maxF>0)fee=Math.min(fee,maxF);
    fee=Math.round(fee*100)/100;
    setFreteCalcResult({distKm:distKm.toFixed(1),durMin,fee,destName:dest.display,origName:origem.name});
    setFreteCalcGeoList([]);
    setFreteCalcLoading(false);
  };

  const markOrderLost=async(order)=>{
    const items=(()=>{try{return parseItems(order);}catch(e){return [];}})();
    let totRec=0,totMiss=0,recVal=0;
    items.forEach(it=>{const o=parseFloat(it.qty)||0;const r=parseFloat(it.received_qty)||0;totRec+=r;totMiss+=Math.max(0,o-r);recVal+=r*(parseFloat(it.unit_cost)||0);});
    const pago=(parseFloat(order.initial_value)||0)+(parseFloat(order.remaining_paid)||0);
    const prejuizo=Math.max(0,pago-recVal);
    if(!window.confirm(
      "Marcar o pedido de "+(order.supplier_name||"fornecedor")+" como PERDIDO?\n\n"+
      "• "+totRec+" unidade(s) já recebida(s) PERMANECEM no estoque ("+fmt(recVal)+").\n"+
      "• "+fmt(pago)+" já pago(s) PERMANECE(m) registrado(s) no caixa.\n"+
      "• "+totMiss+" unidade(s) que faltavam serão dadas como PERDIDAS (não serão recebidas nem pagas).\n\n"+
      "➡️ Prejuízo estimado: "+fmt(prejuizo)+" (o que foi pago além do valor já recebido).\n\n"+
      "O pedido sai da lista de contas a pagar. Deseja continuar?"
    )) return;
    try{
      await supabase.from("orders").update({
        status:"perdido",
        notes:(order.notes?order.notes+" | ":"")+"PERDIDO em "+today()+" · recebidas "+totRec+", perdidas "+totMiss+" · prejuízo "+fmt(prejuizo),
      }).eq("id",order.id);
      toast$("💀 Pedido perdido · prejuízo de "+fmt(prejuizo)+". "+totRec+" un. mantidas no estoque.","#f59e0b");
    }catch(ex){toast$("Erro: "+ex.message,"#f56565");}
  };

  const deleteOrder=async(order)=>{
    if(!window.confirm('Excluir pedido de '+order.supplier_name+'? Todo o fluxo será revertido.'))return;
    const items=parseItems(order);
    try{
      // 1. Reverter estoque pela quantidade efetivamente recebida (parcial ou total)
      if(order.status==="recebido"||order.status==="parcial"){
        for(const item of items){
          if(item.dropship)continue; // dropshipping nunca entrou no estoque
          const qtyToRevert=parseFloat(item.received_qty)||0;
          if(qtyToRevert<=0)continue;
          const oid=item.product_id||(products.find(p=>p.name===item.product_name)?.id);
          if(oid)await deductStock(oid,qtyToRevert);
        }
      }
      // 2. Reverter TODOS os pagamentos lançados no caixa (sinal + recebimentos)
      await supabase.from("cash_transactions").delete().eq("sale_id",order.id);
      // 3. Excluir o pedido
      await supabase.from("orders").delete().eq("id",order.id);

      // 4. Feedback detalhado
      const recItems=items.filter(i=>(parseFloat(i.received_qty)||0)>0);
      const totalPaid=(order.initial_value||0)+(order.remaining_paid||0);
      let msg="🔄 Pedido excluído";
      if(recItems.length>0)msg+=" · "+recItems.length+" item(s) removido(s) do estoque";
      if(totalPaid>0)msg+=" · "+fmt(totalPaid)+" revertido(s) do caixa";
      toast$(msg,"#f59e0b");
    }catch(ex){toast$("Erro: "+ex.message,"#f56565");}
  };

  // Corrige pedidos antigos: reabre linhas que receberam parcial mas foram marcadas como concluídas.
  // Mantém a quantidade já recebida (estoque NÃO é alterado); só o saldo restante volta a ficar disponível.
  const fixPartialOrders=async()=>{
    if(!isAdmin){toast$("Apenas administradores podem corrigir pedidos.","#f56565");return;}
    let affectedOrders=0,affectedItems=0;const updates=[];const nomes=[];
    for(const o of (orders||[])){
      let items;try{items=parseItems(o);}catch(e){continue;}
      let changed=false;
      items=items.map(it=>{
        const ordered=parseFloat(it.qty)||0;const recv=parseFloat(it.received_qty)||0;
        if(it.received===true&&recv>0&&recv<ordered){changed=true;affectedItems++;return{...it,received:false};}
        return it;
      });
      if(changed){
        affectedOrders++;if(o.supplier_name)nomes.push(o.supplier_name);
        const allRec=items.every(i=>i.received);
        updates.push({id:o.id,items:JSON.stringify(items),status:allRec?"recebido":"parcial",received_date:allRec?(o.received_date||today()):null});
      }
    }
    if(affectedOrders===0){toast$("✅ Nenhum pedido antigo precisa de correção.");return;}
    if(!window.confirm("Encontrei "+affectedItems+" item(ns) em "+affectedOrders+" pedido(s)"+(nomes.length?" ("+[...new Set(nomes)].join(", ")+")":"")+" que receberam quantidade PARCIAL mas foram marcados como concluídos.\n\nReabrir o saldo restante desses itens?\n\n• O estoque já lançado NÃO será alterado.\n• Só o saldo que ainda falta volta a ficar disponível para receber.")) return;
    try{
      for(const u of updates){await supabase.from("orders").update({items:u.items,status:u.status,received_date:u.received_date}).eq("id",u.id);}
      toast$("✅ "+affectedItems+" item(ns) reaberto(s) em "+affectedOrders+" pedido(s). Saldo restante disponível para receber.");
    }catch(ex){toast$("Erro: "+ex.message,"#f56565");}
  };



  const saveTreatment=async(t)=>{
    try{
      const rec={id:t.id||uid(),client_id:t.client_id||null,client_name:t.client_name,product_name:t.product_name||null,current_dose_mg:parseFloat(t.current_dose_mg)||null,start_date:t.start_date||null,next_purchase:t.next_purchase||null,frequency_days:parseInt(t.frequency_days)||30,status:t.status||"ativo",notes:t.notes||null,added_by:cu.display_name};
      if(t.id){await supabase.from("treatments").update(rec).eq("id",t.id);}
      else{await supabase.from("treatments").insert([rec]);}
      toast$("✅ Tratamento salvo!");
      setEditingTreatment(null);
      const{data}=await supabase.from("treatments").select("*").order("next_purchase",{ascending:true});
      if(data)setTreatments(data);
    }catch(e){toast$("Erro: "+e.message,"#f56565");}
  };

  const delTreatment=async(id)=>{
    if(!window.confirm("Excluir este tratamento?"))return;
    await supabase.from("treatments").delete().eq("id",id);
    setTreatments(prev=>prev.filter(t=>t.id!==id));
    toast$("Tratamento removido.","#f59e0b");
  };

  const advanceTreatment=async(t)=>{
    // Avança para próxima compra: soma frequency_days à data
    const next=new Date();next.setDate(next.getDate()+(parseInt(t.frequency_days)||30));
    await supabase.from("treatments").update({next_purchase:next.toISOString().slice(0,10)}).eq("id",t.id);
    setTreatments(prev=>prev.map(x=>x.id===t.id?{...x,next_purchase:next.toISOString().slice(0,10)}:x));
    toast$("✅ Próxima compra reagendada!");
  };

  const loadTreatments=async()=>{
    const{data}=await supabase.from("treatments").select("*").order("next_purchase",{ascending:true});
    if(data)setTreatments(data);
  };

  const loadReceivables=async()=>{
    const{data}=await supabase.from("receivables").select("*").order("due_date",{ascending:true});
    if(data)setReceivables(data);
  };
  const addReceivable=async()=>{
    if(!recForm.description||!recForm.value){toast$("Preencha descrição e valor.","#f56565");return;}
    await supabase.from("receivables").insert([{id:uid(),client_id:recForm.client_id||null,client_name:recForm.client_name||null,description:recForm.description,value:parseFloat(recForm.value),due_date:recForm.due_date||null,paid:false,parent_product_id:pf.parent_product_id||null,qty_per_parent:parseFloat(pf.qty_per_parent)||1,total_mg:parseFloat(pf.total_mg)||null,dose_mg:parseFloat(pf.dose_mg)||null,added_by:cu.display_name,created_at:nowISO()}]);
    setRecForm({client_id:"",client_name:"",description:"",value:"",due_date:""});
    loadReceivables();toast$("Conta registrada!");
  };
  // Lembrete de cobrança via WhatsApp (usa o telefone do cliente, se houver)
  // Situação financeira do cliente: saldo em aberto, atrasos atuais e pontualidade histórica
  const clientFin=(cid,cname)=>{
    const recs=(receivables||[]).filter(r=>(cid&&r.client_id===cid)||(cname&&r.client_name===cname));
    const aberto=recs.filter(r=>!r.paid).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
    const vencidasAbertas=recs.filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)<new Date()).length;
    const pagas=recs.filter(r=>r.paid);
    const paidLate=pagas.filter(r=>{if(!r.paid_date||!r.due_date)return false;const pp=String(r.paid_date).split("/");if(pp.length!==3)return false;const pd=new Date(+pp[2],+pp[1]-1,+pp[0]);return pd>new Date(r.due_date+"T23:59:59");}).length;
    const pontualidade=pagas.length>0?((pagas.length-paidLate)/pagas.length)*100:null;
    return {aberto,vencidasAbertas,pagas:pagas.length,paidLate,pontualidade};
  };
  // Backup completo: baixa todos os dados num único arquivo JSON
  const exportBackup=()=>{
    try{
      const data={app:"CaixaPro",version:"7.0",exported_at:new Date().toISOString(),
        sales,products,orders,receivables,payables,clients,suppliers,cash:cashTx,
        settings:{company:companyInfo,payments:dynPays,paysettle:dynSettle,cardtaxes:cardTaxes,frete:freteConfig}};
      const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
      const url=URL.createObjectURL(blob);
      const a=document.createElement("a");a.href=url;a.download="caixapro-backup-"+new Date().toISOString().slice(0,10)+".json";document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);
      const tot=(sales||[]).length+(products||[]).length+(orders||[]).length+(receivables||[]).length+(payables||[]).length+(clients||[]).length+(cashTx||[]).length;
      toast$("✅ Backup gerado ("+tot+" registros). Guarde o arquivo em local seguro.");
    }catch(ex){toast$("Erro ao gerar backup: "+ex.message,"#f56565");}
  };
  // Recibo / Promissória em PDF a partir de uma conta a receber (fiado)
  const exportReceivablePDF=(rec)=>{
    if(!window.jspdf){toast$("Aguarde a página carregar e tente de novo.","#f59e0b");return;}
    const{jsPDF}=window.jspdf;
    const doc=new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
    const W=210,m=18;let y=22;
    const money=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0);
    const cli=clients.find(c=>c.id===rec.client_id||c.name===rec.client_name);
    const venc=rec.due_date?new Date(rec.due_date+"T00:00:00").toLocaleDateString("pt-BR"):"à vista";
    const empresa=(companyInfo&&companyInfo.name)||"CaixaPro · Tirzepatida";
    doc.setFont("helvetica","bold");doc.setFontSize(13);doc.setTextColor(40,40,40);doc.text(empresa,m,y);y+=5;
    doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor(110,110,110);
    const infoLine=[companyInfo&&companyInfo.cnpj?"CNPJ: "+companyInfo.cnpj:"",companyInfo&&companyInfo.phone?companyInfo.phone:"",companyInfo&&companyInfo.address?companyInfo.address:""].filter(Boolean).join("   ·   ");
    if(infoLine){doc.text(infoLine,m,y);y+=5;}
    y+=2;doc.setDrawColor(210);doc.line(m,y,W-m,y);y+=11;
    doc.setFont("helvetica","bold");doc.setFontSize(17);doc.setTextColor(30,30,30);doc.text("RECIBO / PROMISSÓRIA",W/2,y,{align:"center"});y+=10;
    doc.setDrawColor(180);doc.setFillColor(245,245,245);doc.roundedRect(m,y,W-2*m,20,2,2,"FD");
    doc.setFont("helvetica","normal");doc.setFontSize(8);doc.setTextColor(120,120,120);
    doc.text("VALOR",m+6,y+7);doc.text("VENCIMENTO",W/2-10,y+7);doc.text("Nº",W-m-34,y+7);
    doc.setFont("helvetica","bold");doc.setFontSize(13);doc.setTextColor(30,30,30);
    doc.text(money(rec.value),m+6,y+15);doc.text(venc,W/2-10,y+15);doc.text((rec.id||"").slice(0,8).toUpperCase(),W-m-34,y+15);
    y+=30;
    doc.setFont("helvetica","normal");doc.setFontSize(11);doc.setTextColor(40,40,40);
    const devedor=rec.client_name||"__________________________";
    const corpo="Eu, "+devedor+", reconheço dever e prometo pagar a "+empresa+", ou à sua ordem, a quantia de "+money(rec.value)+", referente a \u201C"+(rec.description||"venda")+"\u201D, com vencimento em "+venc+".";
    const linhas=doc.splitTextToSize(corpo,W-2*m);doc.text(linhas,m,y);y+=linhas.length*6+10;
    doc.setFontSize(10);doc.setTextColor(80,80,80);
    doc.text("Devedor(a): "+(rec.client_name||"—")+(cli&&cli.phone?"     Tel: "+cli.phone:""),m,y);y+=16;
    doc.text("Local e data: ____________________________, "+new Date().toLocaleDateString("pt-BR"),m,y);y+=22;
    doc.setDrawColor(120);doc.line(m+22,y,W-m-22,y);y+=5;
    doc.setFontSize(10);doc.setTextColor(60,60,60);doc.text(rec.client_name||"Assinatura do(a) devedor(a)",W/2,y,{align:"center"});
    doc.setFontSize(8);doc.setTextColor(150,150,150);doc.text(empresa+" · Documento gerado em "+new Date().toLocaleString("pt-BR"),W/2,287,{align:"center"});
    doc.save("promissoria-"+String(rec.client_name||"cliente").replace(/\s+/g,"_")+"-"+(rec.id||"").slice(0,6)+".pdf");
    toast$("📄 Recibo/promissória gerado em PDF.");
  };
  const cobrarWhatsApp=(r)=>{
    const cli=clients.find(c=>c.id===r.client_id||c.name===r.client_name);
    let digits=((cli&&cli.phone)||"").replace(/\D/g,"");
    if(digits.length>0&&digits.length<=11)digits="55"+digits;
    const venc=r.due_date?new Date(r.due_date+"T00:00:00").toLocaleDateString("pt-BR"):"";
    const ov=r.due_date&&new Date(r.due_date)<new Date();
    const msg="Olá"+(r.client_name?" "+r.client_name:"")+"! "+(ov?"Passando para lembrar que ":"Lembrete amigável: ")+"o valor de "+fmt(r.value)+" referente a “"+r.description+"”"+(venc?(ov?" venceu em "+venc:" vence em "+venc):"")+". "+(ov?"Consegue me confirmar quando puder acertar? Obrigado!":"Qualquer dúvida, estou à disposição. Obrigado!");
    window.open("https://wa.me/"+digits+"?text="+encodeURIComponent(msg),"_blank");
  };
  const payReceivable=async(id)=>{
    const rec=receivables.find(r=>r.id===id);
    await supabase.from("receivables").update({paid:true,paid_date:today()}).eq("id",id);
    if(rec)await supabase.from("cash_transactions").insert([{id:uid(),description:"Recebimento · "+rec.description+(rec.client_name?" · "+rec.client_name:""),value:rec.value,type:"entrada",category:"Recebimento",added_by:cu.display_name,date:today()}]);
    loadReceivables();toast$("✅ Recebido e lançado no caixa!");
  };
  const deleteReceivable=async(id)=>{
    if(!window.confirm('Excluir esta conta a receber?'))return;
    await supabase.from("receivables").delete().eq("id",id);
    loadReceivables();toast$("Conta removida.","#f59e0b");
  };
  // ── CONTAS A PAGAR ──
  const loadPayables=async()=>{
    const{data}=await supabase.from("payables").select("*").order("due_date",{ascending:true});
    if(data)setPayables(data);
  };
  const addPayable=async()=>{
    if(!payForm.description||!payForm.value){toast$("Preencha descrição e valor.","#f56565");return;}
    const val=parseFloat(payForm.value)||0;
    const n=Math.max(1,Math.min(60,parseInt(payForm.installments)||1));
    const baseDue=payForm.due_date?new Date(payForm.due_date+"T00:00:00"):null;
    const rows=[];
    for(let i=0;i<n;i++){
      let due=null;
      if(baseDue){const d=new Date(baseDue);d.setMonth(d.getMonth()+i);due=d.toISOString().slice(0,10);}
      rows.push({id:uid(),description:payForm.description+(n>1?` (${i+1}/${n})`:""),value:val,category:payForm.category||"outro",supplier_name:payForm.supplier_name||null,due_date:due,paid:false,recurring:n===1&&!!payForm.recurring,added_by:cu.display_name,created_at:nowISO()});
    }
    await supabase.from("payables").insert(rows);
    setPayForm({description:"",value:"",category:"aluguel",supplier_name:"",due_date:"",recurring:false,installments:1});
    loadPayables();toast$(n>1?`${n} parcelas registradas em A Pagar!`:"Conta a pagar registrada!");
  };
  const confirmPayPayable=async()=>{
    const p=payTarget; if(!p)return;
    const total=parseFloat(p.value)||0;
    const already=parseFloat(p.paid_amount)||0;
    const falta=Math.max(0,total-already);
    const paidVal=parseFloat(payAmount);
    if(isNaN(paidVal)||paidVal<=0){toast$("Informe um valor de pagamento válido.","#f56565");return;}
    const eff=Math.min(paidVal,falta>0?falta:paidVal); // não paga mais que o saldo
    const newPaid=already+eff;
    const liquida=payQuitar||newPaid>=total-0.005;
    const eco=payQuitar?Math.max(0,falta-eff):0;
    try{
      if(p._order){
        // Pedido (fornecedor): acumula remaining_paid; quitar marca o restante como pago
        const novoRP=payQuitar?total:newPaid;
        await supabase.from("orders").update({remaining_paid:novoRP}).eq("id",p._orderId);
        await supabase.from("cash_transactions").insert([{id:uid(),description:"Pagamento pedido · "+(p.supplier_name||"fornecedor"),value:eff,type:"saida",category:"Pedido Compra",sale_id:p._orderId,added_by:cu.display_name,date:today()}]);
        await loadOrders();
      } else {
        // Conta a pagar normal
        const upd={paid_amount:newPaid,paid:liquida};
        if(liquida)upd.paid_date=today();
        await supabase.from("payables").update(upd).eq("id",p.id);
        await supabase.from("cash_transactions").insert([{id:uid(),description:"Pagamento · "+p.description+(p.supplier_name?" · "+p.supplier_name:""),value:eff,type:"saida",category:"Pagamento",added_by:cu.display_name,date:today()}]);
        if(liquida&&p.recurring){
          let nextDue=null;
          if(p.due_date){const d=new Date(p.due_date+"T00:00:00");d.setMonth(d.getMonth()+1);nextDue=d.toISOString().slice(0,10);}
          await supabase.from("payables").insert([{id:uid(),description:p.description,value:p.value,category:p.category||"outro",supplier_name:p.supplier_name||null,due_date:nextDue,paid:false,recurring:true,added_by:cu.display_name,created_at:nowISO()}]);
        }
        loadPayables();
      }
    }catch(ex){toast$("Erro ao registrar pagamento: "+ex.message,"#f56565");return;}
    setPayTarget(null);setPayAmount("");setPayQuitar(false);
    if(liquida)toast$(eco>0.005?"✅ Conta liquidada! Economia de "+fmt(eco)+" no saldo.":"✅ Conta liquidada! Pago "+fmt(eff)+".","#f59e0b");
    else toast$("💸 Pagamento parcial de "+fmt(eff)+" registrado · falta "+fmt(Math.max(0,total-newPaid))+".","#f59e0b");
  };
  const deletePayable=async(id)=>{
    if(!window.confirm('Excluir esta conta a pagar?'))return;
    await supabase.from("payables").delete().eq("id",id);
    loadPayables();toast$("Conta removida.","#f59e0b");
  };
  const saveClient=async()=>{
    await supabase.from("clients").update({name:editing.name,email:editing.email,phone:editing.phone,notes:editing.notes,credit_limit:parseFloat(editing.credit_limit)||0,dose:editing.dose||null,interval_days:parseInt(editing.interval_days)||7,treatment_start:editing.treatment_start||null,treatment_notes:editing.treatment_notes||null}).eq("id",editing.id);
    toast$("Cliente atualizado!");setModal(null);setEditing(null);
  };

  const updateSale=async()=>{
    if(!editing)return;
    await supabase.from("sales").update({product_name:editing.product_name,client_name:editing.client_name||null,quantity:parseInt(editing.quantity)||1,unit_price:parseFloat(editing.unit_price)||0,total_price:(parseFloat(editing.unit_price)||0)*(parseInt(editing.quantity)||1),notes:editing.notes,payment_method:editing.payment_method}).eq("id",editing.id);
    toast$("Venda atualizada!");setModal(null);setEditing(null);
  };
  const deleteSale=async(id)=>{
    if(!window.confirm('Excluir esta venda? Estoque (se aplicável) e lançamentos serão revertidos.'))return;
    try{
      // 1. Buscar a venda clicada
      const sale=sales.find(s=>s.id===id);
      if(!sale){toast$("Venda não encontrada.","#f56565");return;}

      // 2. Encontrar TODOS os itens do mesmo batch (ou só este se não tem batch)
      const batchKey=sale.batch_id||id;
      const batchSales=sale.batch_id
        ? sales.filter(s=>s.batch_id===batchKey)
        : [sale];

      // 3. Reverter estoque de CADA item do batch (PULA vendas diretas)
      for(const s of batchSales){
        if(s.is_direct)continue; // venda direta não mexeu no estoque
        const dpid2=s.product_id||(products.find(p=>p.name===s.product_name)?.id);
        if(dpid2&&s.quantity>0)await restoreStock(dpid2,s.quantity);
      }

      // 4. Excluir TODOS os registros de venda do batch
      if(sale.batch_id){
        await supabase.from("sales").delete().eq("batch_id",batchKey);
      } else {
        await supabase.from("sales").delete().eq("id",id);
      }

      // 5. Excluir lançamentos de caixa vinculados ao batch
      await supabase.from("cash_transactions").delete().eq("sale_id",batchKey);

      // 6. Reverter parcelas de A Receber vinculadas ao batch
      const{data:recsToDelete}=await supabase
        .from("receivables")
        .select("id,paid,value")
        .eq("sale_id",batchKey);

      let parcelasRevertidas=0;
      let parcelasPagas=0;
      if(recsToDelete&&recsToDelete.length>0){
        parcelasRevertidas=recsToDelete.length;
        parcelasPagas=recsToDelete.filter(r=>r.paid).length;
        await supabase.from("receivables").delete().eq("sale_id",batchKey);
        await loadReceivables();
      }

      // 7. Reverter pedido(s) automático(s) gerados por esta venda direta (dropshipping)
      let pedidosRevertidos=0;
      const{data:linkedOrders}=await supabase.from("orders").select("id").eq("origin_sale_id",batchKey);
      if(linkedOrders&&linkedOrders.length>0){
        for(const o of linkedOrders){
          // pedido dropship não mexeu no estoque; reverte só os pagamentos lançados no caixa e remove o pedido
          await supabase.from("cash_transactions").delete().eq("sale_id",o.id);
          await supabase.from("orders").delete().eq("id",o.id);
          pedidosRevertidos++;
        }
        await loadOrders();
      }

      // 8. Feedback completo
      const nItens=batchSales.length;
      let msg="🔄 Venda excluída · Estoque revertido";
      if(parcelasRevertidas>0){
        msg+=" · "+parcelasRevertidas+" parcela"+(parcelasRevertidas>1?"s":"")+" removida"+(parcelasRevertidas>1?"s":"")+" de A Receber";
        if(parcelasPagas>0){
          msg+=" (⚠️ "+parcelasPagas+" já havia"+(parcelasPagas>1?"m":"")+" sido recebida"+(parcelasPagas>1?"s":"")+")";
        }
      }
      if(pedidosRevertidos>0)msg+=" · "+pedidosRevertidos+" pedido"+(pedidosRevertidos>1?"s":"")+" automático"+(pedidosRevertidos>1?"s":"")+" removido"+(pedidosRevertidos>1?"s":"");
      toast$(msg,"#f59e0b");
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
  const deleteCash=async id=>{
    if(!window.confirm('Excluir este lançamento do caixa?'))return;
    await supabase.from("cash_transactions").delete().eq("id",id);
    toast$("Lançamento removido.","#f59e0b");
  };

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
    // Detectar se vínculo pai mudou

    const sup=suppliers.find(s=>s.id===editing.supplier_id);
    await supabase.from("products").update({code:editing.code,name:editing.name,description:editing.description,category:editing.category,unit:editing.unit,cost_per_unit:parseFloat(editing.cost_per_unit),price_per_unit:parseFloat(editing.price_per_unit),units_per_pack:parseInt(editing.units_per_pack)||1,batch:editing.batch,expiry:editing.expiry||null,stock_qty:parseInt(editing.stock_qty),min_stock:parseInt(editing.min_stock)||5,supplier_id:editing.supplier_id||null,supplier_name:sup?.name||null,markup,margin,profit,parent_product_id:editing.parent_product_id||null,qty_per_parent:parseFloat(editing.qty_per_parent)||1,total_mg:parseFloat(editing.total_mg)||null,dose_mg:parseFloat(editing.dose_mg)||null}).eq("id",editing.id);
    toast$("Produto atualizado!");setModal(null);setEditing(null);
  };
  const delProduct=async id=>{
    if(!window.confirm('Excluir produto?'))return;
    await supabase.from("products").delete().eq("id",id);
    toast$("Produto removido.","#f59e0b");
    setModal(null);setEditing(null);
  };

  const addClient=async()=>{
    if(!cf.name){toast$("Nome é obrigatório.","#f56565");return;}
    try{
      const{error:e}=await supabase.from("clients").insert([{id:uid(),name:cf.name,email:cf.email||null,phone:cf.phone||null,notes:cf.notes||null,credit_limit:parseFloat(cf.credit_limit)||0,added_by:cu.display_name}]);
      if(e){toast$("Erro ao salvar: "+e.message,"#f56565");return;}
      toast$("Cliente cadastrado!");setModal(null);setCf(FC);
    }catch(ex){toast$("Erro de conexão.","#f56565");}
  };

  const delSupp=async id=>{
    if(!window.confirm('Excluir fornecedor?'))return;
    await supabase.from("suppliers").delete().eq("id",id);
    setSupp(prev=>prev.filter(s=>s.id!==id));
    toast$("Fornecedor removido.","#f59e0b");
  };

  const delClient=async id=>{
    if(!window.confirm('Excluir este cliente?'))return;
    await supabase.from("clients").delete().eq("id",id);
    toast$("Cliente removido.","#f59e0b");
  };

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
  const toggleUser=async(id,active)=>{
    await supabase.from("app_users").update({active}).eq("id",id);
    setUsers(prev=>prev.map(u=>u.id===id?{...u,active}:u));
    toast$(active?"Usuário ativado.":"Usuário desativado.","#10b981");
  };

  const saveUser=async()=>{
    const updates={display_name:editing.display_name,role:editing.role,active:editing.active,permissions:editing.role==="admin"?{}:(editing.permissions||{})};
    if(editing.new_password){if(editing.new_password.length<6){toast$("Mínimo 6 caracteres.","#f59e0b");return;}if(editing.new_password!==editing.new_password2){toast$("Senhas não coincidem.","#f56565");return;}updates.password_hash=hashPw(editing.new_password);}
    await supabase.from("app_users").update(updates).eq("id",editing.id);
    toast$("Usuário atualizado!");setModal(null);setEditing(null);
  };
  const deleteUser=async(id)=>{
    if(!isAdmin){toast$("Apenas administradores podem excluir usuários.","#f56565");return;}
    const u=appUsers.find(x=>x.id===id);
    if(u&&u.username===cu.username){toast$("Você não pode excluir o próprio usuário.","#f56565");return;}
    if(!window.confirm("Excluir o usuário "+(u?u.display_name:"")+"? Esta ação é permanente."))return;
    await supabase.from("app_users").delete().eq("id",id);
    setUsers(prev=>prev.filter(x=>x.id!==id));
    toast$("Usuário excluído.","#f59e0b");
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
    {id:"financeiro",l:"Financeiro",n:"dollar"},
    {id:"clientes",l:"Clientes",n:"client"},
    ...(isAdmin?[{id:"config",l:"Config",n:"settings"}]:[]),
  ].filter(it=>isAdmin||it.id==="config"||((cu?.permissions||{})[it.id]!==false));
  const allowedTabs=nav.map(n=>n.id).join(",");
  useEffect(()=>{
    if(!cu)return;
    const allowed=allowedTabs.split(",").filter(Boolean);
    if(allowed.length&&!allowed.includes(tab)&&tab!=="frete")setTab(allowed[0]);
  },[cu,tab,allowedTabs]);

  // Helper cálculo preview
  const MPreview=({cost,price})=>{
    if(!cost||!price||parseFloat(cost)<=0||parseFloat(price)<=0)return null;
    const{markup,margin,profit}=calcM(cost,price);
    return <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem",marginBottom:".8rem",display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".4rem",textAlign:"center"}}>
      {[{l:"Markup",v:fmtPct(markup),c:"#4f5ef0"},{l:"Margem",v:fmtPct(margin),c:"#8b44f0"},{l:"Lucro/un",v:fmt(profit),c:"#10b981"}].map(m=><div key={m.l}><div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".1rem"}}>{m.l}</div><div style={{fontSize:".88rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div></div>)}
    </div>;
  };

  //  RENDER LOGIN 
  if(!cu)return(<>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');@keyframes loginShake{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-6px)}80%{transform:translateX(6px)}}`}</style>
    <LoginScreen onLogin={login} dark={dark}/>
  </>);

  //  RENDER APP 
  return(<ErrorBoundary><>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}input[type=date]::-webkit-calendar-picker-indicator{filter:${dark?"invert(.5)":"none"}}`}</style>

    {/* TOAST */}
    {toast&&<div style={{position:"fixed",bottom:"1.25rem",right:"1.25rem",zIndex:999,background:"var(--card)",border:`1px solid ${toast.color}50`,borderRadius:".65rem",padding:".65rem 1.1rem",display:"flex",alignItems:"center",gap:".45rem",fontSize:".82rem",color:toast.color,fontFamily:"'DM Sans',sans-serif",boxShadow:"0 8px 30px var(--shadow)",animation:"fadeUp .3s ease",maxWidth:320}}><Ic n="save" s={14}/>{toast.msg}</div>}

    <div id="app-root" style={{minHeight:"100vh",background:dark?"#080a14":"#f0f4ff",fontFamily:"'DM Sans',sans-serif",color:"var(--tx)"}}>

      {/*  TOP BAR  */}
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
          {!isMobile&&<div style={{fontSize:".62rem",color:syncing?"#8b44f0":"var(--tx6)",display:"flex",alignItems:"center",gap:".25rem"}}><Ic n="sync" s={10}/>{syncing?"...":lastSync||"—"}</div>}
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

      {/*  NAV  */}
      <div style={{background:"var(--bg2)",borderBottom:"1px solid var(--bdr)",padding:"0 .4rem",display:"flex",overflowX:"auto",transition:"background .3s"}}>
        {nav.map(item=>(
          <button key={item.id} onClick={()=>{setTab(item.id);setSearch("");setFcat("all");}} style={{display:"flex",alignItems:"center",gap:".28rem",padding:isMobile?".5rem .5rem":".58rem .7rem",background:"none",border:"none",fontSize:isMobile?".68rem":".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,color:tab===item.id?"var(--navon)":"var(--navoff)",borderBottom:tab===item.id?"2px solid #4f5ef0":"2px solid transparent",transition:"all .2s",whiteSpace:"nowrap"}}>
            <Ic n={item.n} s={12}/>{isMobile?item.l.split(" ")[0]:item.l}
            {item.id==="estoque"&&zeroStk.length>0&&<span style={{background:"#f56565",color:"#fff",borderRadius:"99px",fontSize:".55rem",fontWeight:700,padding:"0 .28rem",lineHeight:"1.5"}}>{zeroStk.length}</span>}
          </button>
        ))}
      </div>

      {/*  CONTEÚDO  */}
      <div style={{maxWidth:960,margin:"0 auto",padding:isMobile?".6rem .5rem":".9rem .8rem"}}>

        {/* ══ DASHBOARD ══ */}
        {tab==="dashboard"&&(<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
              <Btn sm v="ghost" onClick={()=>setShowGoals(true)}>📅 Metas & Histórico</Btn>
              <Btn sm v="warn" onClick={()=>setShowOrderModal(true)}><Ic n="pkg" s={12}/>📦 Pedido de Compra</Btn>
              <Btn sm v="ghost" onClick={()=>setShowTreatments(true)}>💊 Tratamentos</Btn>
              <Btn sm v="ghost" onClick={()=>setShowRanking(true)}>🏆 Ranking</Btn>
              <Btn sm v="ghost" onClick={()=>setShowCashFlow(true)}>💰 Fluxo Caixa</Btn>
              {isAdmin&&<Btn sm v="ghost" onClick={()=>{setTab("config");setConfigSection("usuarios");}}>⚙️ Config</Btn>}
            </div>
            <Btn v="info" onClick={()=>setShowA(true)}><Ic n="analytics" s={14}/> Relatório Gerencial</Btn>
          </div>
          {/* ══ ALERTAS INTELIGENTES ══ */}
          {(()=>{
            const hoje=new Date();
            const l3=new Date();l3.setDate(l3.getDate()+3);
            const l30=new Date();l30.setDate(l30.getDate()+30);
            const venRec=(receivables||[]).filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)<hoje);
            const venPag=(payables||[]).filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)<hoje);
            const dueRec=(receivables||[]).filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)>=hoje&&new Date(r.due_date)<=l3);
            const duePag=(payables||[]).filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)>=hoje&&new Date(p.due_date)<=l3);
            const pag30=(payables||[]).filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)<=l30).reduce((a,p)=>a+(parseFloat(p.value)||0),0)+(orders||[]).filter(o=>o.status==="pendente"||o.status==="parcial").reduce((a,o)=>a+(parseFloat(o.remaining_value)||0),0);
            const recompras=(treatments||[]).filter(t=>t.status==="ativo"&&t.next_purchase&&new Date(t.next_purchase)<=l3);
            const al=[];
            if(venPag.length)al.push({c:"#f56565",i:"⚠️",t:venPag.length+" conta(s) a pagar VENCIDA(S) · "+fmt(venPag.reduce((a,p)=>a+p.value,0)),go:()=>{setTab("financeiro");setFinTab("pagar");}});
            if(venRec.length)al.push({c:"#f56565",i:"⚠️",t:venRec.length+" recebível(is) em atraso · "+fmt(venRec.reduce((a,r)=>a+r.value,0)),go:()=>{setTab("financeiro");setFinTab("receber");}});
            if(duePag.length)al.push({c:"#f59e0b",i:"⏰",t:duePag.length+" conta(s) a pagar vencem em ≤3 dias",go:()=>{setTab("financeiro");setFinTab("pagar");}});
            if(dueRec.length)al.push({c:"#f59e0b",i:"⏰",t:dueRec.length+" recebível(is) vencem em ≤3 dias",go:()=>{setTab("financeiro");setFinTab("receber");}});
            if(net<pag30)al.push({c:"#f56565",i:"💸",t:"Caixa ("+fmt(net)+") não cobre as contas dos próximos 30 dias ("+fmt(pag30)+")",go:()=>{setTab("financeiro");setFinTab("resumo");}});
            if(cashIn>0&&margin<10)al.push({c:"#f59e0b",i:"📉",t:"Margem baixa: "+fmtPct(margin)+" — revise preços/custos",go:()=>{setTab("financeiro");setFinTab("saude");}});
            if(zeroStk.length)al.push({c:"#f56565",i:"📦",t:zeroStk.length+" produto(s) zerado(s)",go:()=>{setTab("estoque");setEstTab("estoque");}});
            else if(lowStk.length)al.push({c:"#f59e0b",i:"📦",t:lowStk.length+" produto(s) com estoque baixo",go:()=>{setTab("estoque");setEstTab("estoque");}});
            if(recompras.length)al.push({c:"#8b44f0",i:"🔄",t:recompras.length+" recompra(s) de tratamento previstas",go:()=>setShowTreatments(true)});
            if(!al.length)return(
              <div style={{background:"linear-gradient(135deg,#10b98115,#10b98105)",border:"1px solid #10b98130",borderRadius:".75rem",padding:".7rem 1rem",marginBottom:".7rem",display:"flex",alignItems:"center",gap:".5rem"}}>
                <span style={{fontSize:"1.1rem"}}>✅</span><span style={{fontSize:".82rem",fontWeight:600,color:"#10b981"}}>Tudo em dia — nenhum alerta no momento.</span>
              </div>
            );
            return(
              <div style={{marginBottom:".7rem",display:"flex",flexDirection:"column",gap:".4rem"}}>
                {al.map((a,i)=>(
                  <button key={i} onClick={a.go} style={{display:"flex",alignItems:"center",gap:".5rem",textAlign:"left",width:"100%",background:a.c+"12",border:"1px solid "+a.c+"35",borderRadius:".6rem",padding:".6rem .8rem",cursor:"pointer"}}>
                    <span style={{fontSize:"1rem",flexShrink:0}}>{a.i}</span>
                    <span style={{fontSize:".8rem",fontWeight:600,color:a.c,flex:1}}>{a.t}</span>
                    <span style={{fontSize:".8rem",color:a.c,opacity:.7,flexShrink:0}}>→</span>
                  </button>
                ))}
              </div>
            );
          })()}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Receita" value={fmt(cashIn)} color="#10b981" icon="up"/>
            <KCard label="Custos" value={fmt(cashOut)} color="#f56565" icon="dn"/>
            <KCard label="Lucro" value={fmt(net)} sub={fmtPct(margin)+" margem"} color={net>=0?"#4f5ef0":"#f56565"}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Vendas" value={fmtN(sales.length)} sub={fmtN(totalUnits)+" un"} color="#8b44f0" icon="sales"/>
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
          {/* ⚠️ Clientes em risco de churn */}
          {(()=>{
            const thirtyDaysAgo=new Date();thirtyDaysAgo.setDate(thirtyDaysAgo.getDate()-30);
            const churnRisk=(clients||[]).filter(c=>{
              const lastSale=(sales||[]).filter(s=>s.client_id===c.id||s.client_name===c.name)
                .sort((a,b)=>new Date(b.created_at)-new Date(a.created_at))[0];
              if(!lastSale)return false;
              return new Date(lastSale.created_at)<thirtyDaysAgo;
            });
            if(churnRisk.length===0)return null;
            return(
              <div style={{background:"linear-gradient(135deg,#f5656508,#f59e0b08)",border:"1px solid #f59e0b30",borderRadius:".75rem",padding:".75rem 1rem",marginBottom:".65rem"}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".35rem"}}>
                  <div style={{fontWeight:700,fontSize:".78rem",color:"#f59e0b",display:"flex",alignItems:"center",gap:".35rem"}}>⚠️ Clientes sem compra +30 dias</div>
                  <button onClick={()=>setTab("clientes")} style={{fontSize:".65rem",color:"#f59e0b",background:"none",border:"none",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>ver clientes →</button>
                </div>
                <div style={{display:"flex",gap:".35rem",flexWrap:"wrap"}}>
                  {churnRisk.slice(0,4).map(c=>(
                    <span key={c.id} style={{fontSize:".7rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".15rem .55rem",border:"1px solid #f59e0b30"}}>{c.name?.split(" ")[0]}</span>
                  ))}
                  {churnRisk.length>4&&<span style={{fontSize:".7rem",color:"var(--sub)"}}>+{churnRisk.length-4} mais</span>}
                </div>
              </div>
            );
          })()}
          {/* 💊 Recompras de tratamento próximas */}
          {(()=>{
            const ativos=(treatments||[]).filter(t=>t.status==="ativo"&&t.next_purchase);
            const proximas=ativos.map(t=>({...t,dias:Math.ceil((new Date(t.next_purchase)-new Date())/864e5)})).filter(t=>t.dias<=7).sort((a,b)=>a.dias-b.dias);
            if(proximas.length===0)return null;
            return(
              <div style={{background:"var(--card)",border:"1px solid #8b44f030",borderRadius:".75rem",padding:".8rem 1rem",marginBottom:".7rem"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:".5rem"}}>
                  <span style={{fontWeight:700,fontSize:".82rem",color:"#8b44f0"}}>💊 Recompras próximas ({proximas.length})</span>
                  <button onClick={()=>setShowTreatments(true)} style={{background:"none",border:"none",color:"#8b44f0",fontSize:".68rem",fontWeight:600,cursor:"pointer"}}>Ver todos →</button>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:".35rem"}}>
                  {proximas.slice(0,5).map(t=>(
                    <div key={t.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:".73rem"}}>
                      <span style={{color:"var(--tx3)",fontWeight:600}}>{t.client_name}{t.current_dose_mg?` · ${t.current_dose_mg}mg`:""}</span>
                      <span style={{fontWeight:700,color:t.dias<0?"#f56565":t.dias<=2?"#f59e0b":"#10b981"}}>{t.dias<0?`Atrasado ${Math.abs(t.dias)}d`:t.dias===0?"HOJE":`${t.dias}d`}</span>
                    </div>
                  ))}
                  {proximas.length>5&&<span style={{fontSize:".68rem",color:"var(--sub)"}}>+{proximas.length-5} mais</span>}
                </div>
              </div>
            );
          })()}

          {/* Pedidos pendentes */}
          {(()=>{
            const pedPend=(orders||[]).filter(o=>o.status==="pendente"||o.status==="parcial");
            const perdidos=(orders||[]).filter(o=>o.status==="perdido");
            if(pedPend.length===0&&perdidos.length===0)return null;
            const allItems=pedPend.flatMap(o=>parseItems(o));
            const prodMap={};
            allItems.forEach(i=>{prodMap[i.product_name]=(prodMap[i.product_name]||0)+i.qty;});
            return(
              <div style={{background:"var(--card)",border:"1px solid #f59e0b30",borderRadius:".75rem",padding:".75rem 1rem",marginBottom:".65rem"}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".45rem"}}>
                  <div style={{fontWeight:700,fontSize:".78rem",color:"#f59e0b",display:"flex",alignItems:"center",gap:".35rem"}}>
                    <Ic n="pkg" s={13}/>📦 Pedidos em aberto
                  </div>
                  <div style={{display:"flex",gap:".5rem",alignItems:"center"}}>
                    {perdidos.length>0&&<span style={{fontSize:".65rem",color:"#f56565",background:"#f5656515",borderRadius:"99px",padding:".12rem .45rem",border:"1px solid #f5656530"}}>💀 {perdidos.length} perdido{perdidos.length>1?"s":""}</span>}
                    <button onClick={()=>{setTab("estoque");setEstTab("pedidos");}} style={{fontSize:".65rem",color:"#f59e0b",background:"none",border:"none",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>ver todos →</button>
                  </div>
                </div>
                <div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
                  {Object.entries(prodMap).slice(0,5).map(([name,qty])=>(
                    <span key={name} style={{fontSize:".7rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".15rem .55rem",border:"1px solid #f59e0b30"}}>
                      {name}: +{qty}
                    </span>
                  ))}
                  {Object.keys(prodMap).length>5&&<span style={{fontSize:".7rem",color:"var(--sub)"}}>+{Object.keys(prodMap).length-5} produto(s)</span>}
                  {Object.keys(prodMap).length===0&&pedPend.length>0&&<span style={{fontSize:".73rem",color:"var(--tx5)"}}>Aguardando {pedPend.length} pedido{pedPend.length>1?"s":""}</span>}
                </div>
              </div>
            );
          })()}

          {/* Informativos — A Receber + Meta + Alertas */}
          {(()=>{
            const now=new Date();
            const monthStart=new Date(now.getFullYear(),now.getMonth(),1);
            const monthRev=(cashTx||[]).filter(t=>t.type==="entrada"&&new Date(t.created_at)>=monthStart).reduce((a,t)=>a+t.value,0);
            const goal=parseFloat(monthGoal)||0;
            const pend=(receivables||[]).filter(r=>!r.paid);
            const pendVal=pend.reduce((a,r)=>a+r.value,0);
            const overdue=pend.filter(r=>r.due_date&&new Date(r.due_date)<new Date());
            return(<>
              {/* Meta Mensal — espelho do Goals modal, mês vigente */}
              {(()=>{
                const MONTHS_ABR=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
                const mName=MONTHS_ABR[now.getMonth()];
                const key=now.getFullYear()+"-"+(now.getMonth()+1).toString().padStart(2,"0");
                const mGoal=monthlyGoals[key]||parseFloat(monthGoal)||0;
                const mReal=(sales||[]).filter(s=>{const d=new Date(s.created_at||Date.now());return d>=monthStart&&d<=new Date(now.getFullYear(),now.getMonth()+1,0,23,59,59);}).reduce((a,s)=>a+s.total_price,0); // discount applied in batchRevenue
                const mPct=mGoal>0?(mReal/mGoal)*100:0;
                return(
                  <div style={{background:"linear-gradient(135deg,#4f5ef010,#10b98108)",border:"1px solid #4f5ef040",borderRadius:".65rem",padding:".7rem .85rem",marginBottom:".65rem"}}>
                    <div style={{display:"flex",alignItems:"center",gap:".75rem",marginBottom:".35rem"}}>
                      <span style={{fontWeight:700,fontSize:".82rem",color:"#4f5ef0",minWidth:30}}>{mName}</span>
                      <span style={{fontSize:".6rem",background:"#4f5ef020",color:"#4f5ef0",borderRadius:"99px",padding:".1rem .4rem",fontWeight:700}}>atual</span>
                      <div style={{flex:1}}/>
                      <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
                        <span style={{fontSize:".75rem",color:mReal>=mGoal&&mGoal>0?"#10b981":"var(--tx3)",fontWeight:600}}>{fmt(mReal)}</span>
                        <span style={{fontSize:".68rem",color:"var(--tx5)"}}>de</span>
                        <span style={{fontSize:".78rem",color:"var(--tx4)",fontWeight:600,fontFamily:"'Syne',sans-serif"}}>{mGoal>0?fmt(mGoal):"—"}</span>
                        <span style={{fontSize:".72rem",fontWeight:700,color:mPct>=100?"#10b981":mPct>=70?"#f59e0b":"var(--tx5)",minWidth:42,textAlign:"right"}}>{mGoal>0?fmtPct(mPct):"—"}{mPct>100?" 🚀":mPct>=100?" 🏆":""}</span>
                      </div>
                    </div>
                    <div style={{height:6,background:"var(--bdr)",borderRadius:3}}>
                      <div style={{height:"100%",width:Math.min(100,mPct)+"%",background:mPct>=100?"linear-gradient(90deg,#10b981,#059669)":mPct>=70?"linear-gradient(90deg,#f59e0b,#d97706)":"linear-gradient(90deg,#4f5ef0,#8b44f0)",borderRadius:3,transition:"width .5s"}}/>
                    </div>
                  </div>
                );
              })()}
              {/* A Receber + Alertas */}
              <div style={{display:"grid",gridTemplateColumns:pend.length>0&&!isMobile?"repeat(auto-fit,minmax(180px,1fr))":"1fr",gap:".65rem",marginBottom:".65rem"}}>
                {pend.length>0&&(
                  <div onClick={()=>{setTab("financeiro");setFinTab("receber");}} style={{background:"var(--card)",border:"1px solid #4f5ef030",borderRadius:".75rem",padding:".65rem .9rem",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div>
                      <div style={{fontSize:".72rem",fontWeight:700,color:"var(--tx)",display:"flex",alignItems:"center",gap:".3rem"}}><Ic n="dollar" s={12}/>A Receber</div>
                      <div style={{fontSize:".63rem",color:"var(--tx5)",marginTop:".1rem"}}>{pend.length} conta{pend.length!==1?"s":""} em aberto</div>
                    </div>
                    <span style={{fontWeight:800,color:"#4f5ef0",fontFamily:"'Syne',sans-serif",fontSize:".9rem"}}>{fmt(pendVal)}</span>
                  </div>
                )}
                {overdue.length>0&&(
                  <div onClick={()=>{setTab("financeiro");setFinTab("receber");}} style={{background:"var(--card)",border:"1px solid #f5656540",borderRadius:".75rem",padding:".65rem .9rem",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div>
                      <div style={{fontSize:".72rem",fontWeight:700,color:"#f56565",display:"flex",alignItems:"center",gap:".3rem"}}><Ic n="warn" s={12}/>{overdue.length} Vencida{overdue.length!==1?"s":""}</div>
                      <div style={{fontSize:".63rem",color:"var(--tx5)",marginTop:".1rem"}}>clique para ver</div>
                    </div>
                    <span style={{fontWeight:800,color:"#f56565",fontFamily:"'Syne',sans-serif",fontSize:".9rem"}}>{fmt(overdue.reduce((a,r)=>a+r.value,0))}</span>
                  </div>
                )}
              </div>
            </>);
          })()}

          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".75rem"}}>
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
                  <div style={{fontSize:".76rem",color:"var(--tx2)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name}</div>
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
                <Btn sm v="ghost" onClick={()=>setTab("frete")}><span style={{fontSize:11}}>🛵</span> Frete</Btn>
                {canEdit&&<Btn sm onClick={()=>setModal("sale")}><Ic n="plus" s={12}/>Nova</Btn>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".6rem",marginBottom:".7rem"}}>
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
            {fSales.length===0?<div style={{textAlign:"center",padding:"2.5rem 1rem"}}><div style={{fontSize:"2rem",marginBottom:".5rem"}}>🛒</div><p style={{color:"var(--tx5)",fontSize:".8rem",marginBottom:".75rem"}}>Nenhuma venda registrada.</p>{canEdit&&<Btn sm onClick={()=>setModal("sale")}><Ic n="plus" s={12}/>Registrar venda</Btn>}</div>
            :fSales.map(s=>(
              <div key={s.id} style={{padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:".35rem",flexWrap:"wrap",marginBottom:".18rem"}}>
                    <span style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)"}}>{s.product_name}</span>
                    <Badge color="#8b44f0" sm>{s.quantity} un</Badge>
                    {s.unit_price===0&&<Badge color="#f59e0b" sm>Cortesia</Badge>}
                    {s.batch_id&&<Badge color="#44475a" sm>Lote</Badge>}
                    <Badge color={(s.payment_method||"").startsWith("Dividido")?"#8b44f0":s.payment_method==="Crédito Parcelado"?"#8b44f0":s.payment_method==="Débito"?"#0891b2":s.payment_method==="PIX"?"#10b981":s.payment_method==="Dinheiro"?"#f59e0b":"#0891b2"} sm>{s.payment_method}</Badge>
                    {(parseFloat(s.discount)||0)>0&&sales.filter(x=>x.batch_id===s.batch_id)[0]?.id===s.id&&<Badge color="#4f5ef0" sm>🏷️ -{fmt(parseFloat(s.discount)||0)}</Badge>}
                  </div>
                  <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{s.date}{s.client_name&&` · 👤 ${s.client_name}`}{s.notes&&` · ${s.notes}`}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:".5rem",flexShrink:0,marginLeft:".75rem"}}>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:700,color:"#10b981",fontSize:".88rem",fontFamily:"'Syne',sans-serif"}}>{fmt(Math.max(0,s.total_price-(s.batch_id?0:(parseFloat(s.discount)||0))))}</div>
                    <div style={{fontSize:".63rem",color:"var(--tx5)"}}>{fmt(s.unit_price)}/un</div>
                  </div>
                  {canEdit&&<div style={{display:"flex",gap:".2rem"}}>
                    <button onClick={()=>{const batch=s.batch_id?sales.filter(x=>x.batch_id===s.batch_id):[s];setShowReceipt(batch);}} style={{background:"none",border:"none",color:"#8b44f0",padding:".2rem"}} title="Comprovante"><Ic n="pdf" s={13}/></button>
                    {canEdit&&<button onClick={()=>{
                      const batch=s.batch_id?sales.filter(x=>x.batch_id===s.batch_id&&x.product_id):[s];
                      cartReset();
                      const itens=batch.map(b=>({key:uid(),product_id:b.product_id||"",product_name:b.product_name,unit_price:String(b.unit_price||0),qty:b.quantity||1,unit:b.unit||"un"}));
                      setCartItems(itens.length>0?itens:[newCartItem()]);
                      if(s.client_name)setCartClient({id:s.client_id||"",name:s.client_name});
                      setModal("sale");toast$("🔄 Carrinho preenchido com a venda anterior!");
                    }} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}} title="Repetir venda"><Ic n="refresh" s={13}/></button>}
                    {s.client_name&&<button onClick={()=>{const m="Olá "+s.client_name+"! ✅ Venda de *"+s.product_name+"* registrada. Total: *"+fmt(Math.max(0,s.total_price-(parseFloat(s.discount)||0)))+"*. Pag: "+s.payment_method+". Obrigado!";window.open("https://wa.me/?text="+encodeURIComponent(m),"_blank");}} style={{background:"none",border:"none",color:"#25d366",padding:".2rem"}} title="WhatsApp"><span style={{fontSize:13}}>📱</span></button>}
                    <button onClick={()=>{setEditing({...s,quantity:String(s.quantity),unit_price:String(s.unit_price)});setModal("editSale");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}} title="Editar"><Ic n="edit" s={13}/></button>
                    {canDelete&&<button onClick={()=>deleteSale(s.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}} title="Excluir"><Ic n="trash" s={13}/></button>}
                  </div>}
                </div>
              </div>
            ))}
          </div>
        </>)}

        {/* ══ ESTOQUE ══ */}
        {/* ══════════ ESTOQUE — sub-navegação ══════════ */}
        {tab==="estoque"&&(
          <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",marginBottom:"1rem",borderBottom:"1px solid var(--bdr)",paddingBottom:".75rem"}}>
            {[{id:"estoque",l:"📦 Estoque"},{id:"produtos",l:"🏷️ Produtos"},{id:"pedidos",l:"🚚 Pedidos"}].map(g=>(
              <button key={g.id} onClick={()=>setEstTab(g.id)} style={{padding:".42rem .85rem",borderRadius:".5rem",fontSize:".82rem",fontFamily:"'DM Sans',sans-serif",fontWeight:700,border:"none",cursor:"pointer",background:estTab===g.id?"linear-gradient(135deg,#4f5ef0,#8b44f0)":"var(--pill)",color:estTab===g.id?"#fff":"var(--navoff)",transition:"all .2s",whiteSpace:"nowrap"}}>{g.l}</button>
            ))}
          </div>
        )}
        {tab==="estoque"&&estTab==="estoque"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Estoque</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={fProds.map(p=>({Código:p.code,Nome:p.name,Cat:p.category,Unidade:p.unit||"un",Estoque:p.stock_qty,"Est.Mín":p.min_stock||5,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Margem:fmtPct(p.margin),Fornecedor:p.supplier_name||"—"}))} name="estoque-caixapro" sheet="Estoque"/>
              <PBtn cols={[{k:"Nome",l:"Produto"},{k:"Cat",l:"Cat."},{k:"Estoque",l:"Estoque"},{k:"Custo",l:"Custo"},{k:"Preço",l:"Preço"},{k:"Markup",l:"Markup"},{k:"Fornecedor",l:"Fornecedor"}]} rows={fProds.map(p=>({Nome:p.name,Cat:p.category,Estoque:`${p.stock_qty} ${p.unit||"un"}`,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Fornecedor:p.supplier_name||"—"}))} name="estoque-caixapro" title="Relatório de Estoque"/>
              {canEdit&&<><Btn v="info" sm onClick={()=>setModal("stockEntry")}><Ic n="arrup" s={12}/>Entrada</Btn><Btn sm onClick={()=>setModal("produto")}><Ic n="plus" s={12}/>Produto</Btn></>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".6rem",marginBottom:".7rem"}}>
            <KCard label="Itens" value={products.length} color="#4f5ef0"/>
            <KCard label="Unidades" value={fmtN(products.reduce((a,p)=>a+p.stock_qty,0))} color="#8b44f0"/>
            <KCard label="Valor estoque" value={fmt(stockVal)} color="#f59e0b"/>
            {zeroStk.length>0&&<KCard label="Zerados" value={zeroStk.length} color="#f56565"/>}
          </div>
          {(()=>{
            const soon=(products||[]).filter(p=>p.expiry&&daysUntil(p.expiry)!==null&&daysUntil(p.expiry)<=30).sort((a,b)=>daysUntil(a.expiry)-daysUntil(b.expiry));
            if(!soon.length)return null;
            return(
              <div style={{background:"var(--card)",border:"1px solid #f59e0b40",borderRadius:".65rem",padding:".7rem 1rem",marginBottom:".65rem"}}>
                <div style={{fontWeight:700,fontSize:".78rem",color:"#f59e0b",marginBottom:".35rem",display:"flex",alignItems:"center",gap:".35rem"}}><Ic n="warn" s={13}/>⏰ Vencimentos nos próximos 30 dias</div>
                <div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
                  {soon.map(p=>{
                    const d=daysUntil(p.expiry);
                    const col=d<0?"#f56565":d<=7?"#f59e0b":"var(--tx3)";
                    const bg=d<0?"#f5656515":"#f59e0b15";
                    const bdr=d<0?"#f5656530":"#f59e0b30";
                    return <span key={p.id} style={{fontSize:".7rem",color:col,background:bg,borderRadius:"99px",padding:".15rem .55rem",border:"1px solid "+bdr}}>{p.name} · {d<0?"Vencido":d===0?"Hoje":d+"d"}</span>;
                  })}
                </div>
              </div>
            );
          })()}
          <div style={{display:"flex",gap:".35rem",marginBottom:".65rem",overflowX:"auto"}}>
            {cats.map(c=><button key={c} onClick={()=>setFcat(c)} style={{padding:".28rem .65rem",borderRadius:"99px",border:`1px solid ${fcat===c?"#4f5ef0":"var(--bdr2)"}`,background:fcat===c?"#4f5ef020":"transparent",color:fcat===c?"#4f5ef0":"var(--navoff)",fontSize:".7rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,whiteSpace:"nowrap"}}>{c==="all"?"Todos":`${(activeCats.find(x=>x.key===c)||{icon:"📋",label:c}).icon} ${(activeCats.find(x=>x.key===c)||{label:c}).label}`}</button>)}
          </div>
          <div style={{display:"grid",gap:".6rem"}}>
            {fProds.map(p=>{
              const days=p.expiry?daysUntil(p.expiry):null;
              return(
                <div key={p.id} style={{background:"var(--card)",border:`1px solid ${p.stock_qty<=0?"#f5656530":p.stock_qty<=(p.min_stock||5)?"#f59e0b30":"var(--bdr)"}`,borderRadius:".75rem",padding:".85rem 1rem"}}>
                  <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".55rem"}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:".38rem",flexWrap:"wrap",marginBottom:".2rem"}}>
                        <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".88rem",color:"var(--tx)"}}>{(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name}</span>
                        
                      
                      {p.category==="tirzepatida"&&p.total_mg&&<Badge color="#8b44f0" sm>💊 {(p.stock_qty*parseFloat(p.total_mg)).toFixed(0)}mg</Badge>}
                      {p.dose_mg&&<Badge color="#8b44f0" sm>💊 {p.dose_mg}mg dose</Badge>}
                        <Badge color={stkColor(p.stock_qty)} sm>{p.stock_qty<=0?"Zerado":p.stock_qty<=(p.min_stock||5)?"Baixo":"OK"}</Badge>
                        {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                        {days!==null&&days<=30&&<Badge color={expColor(days)} sm>{days<0?"Vencido":`Vcto ${days}d`}</Badge>}
                      </div>
                      <div style={{fontSize:".65rem",color:"var(--tx5)"}}>
                      {p.code}{p.batch&&` · Lote: ${p.batch}`}
                      {(()=>{
                        const sold30=(sales||[]).filter(s=>{try{const d=new Date(s.created_at);return(s.product_id===p.id||s.product_name===p.name)&&d>new Date(Date.now()-30*864e5);}catch{return false;}}).reduce((a,s)=>a+(s.quantity||0),0);
                        if(sold30<=0||p.stock_qty<=0)return null;
                        const daysLeft=Math.floor((p.stock_qty/sold30)*30);
                        const col=daysLeft<7?"#f56565":daysLeft<15?"#f59e0b":"#10b981";
                        return<span style={{marginLeft:".4rem",color:col,fontWeight:600}}>· ~{daysLeft}d estoque</span>;
                      })()}
                    </div>
                    </div>
                    <div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".65rem"}}>
                      {canEdit&&<><button onClick={()=>{setStF({product_id:p.id,qty:"",cost_total:"",notes:""});setModal("stockEntry");}} style={{background:"#0e1e0e",border:"1px solid #10b98130",borderRadius:".4rem",padding:".28rem .5rem",color:"#10b981",display:"flex",alignItems:"center"}}><Ic n="arrup" s={11}/></button>
                      <button onClick={()=>{setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5)});setModal("editProd");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>
                      {canDelete&&<button onClick={()=>delProduct(p.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}</>}
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

        {/* ══════════ FINANCEIRO — sub-navegação ══════════ */}
        {tab==="financeiro"&&(
          <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",marginBottom:"1rem",borderBottom:"1px solid var(--bdr)",paddingBottom:".75rem"}}>
            {[{id:"resumo",l:"📊 Resumo"},{id:"projecao",l:"📈 Projeção"},{id:"saude",l:"💚 Saúde"},{id:"receber",l:"💰 A Receber"},{id:"pagar",l:"📄 A Pagar"},{id:"caixa",l:"💵 Caixa"},{id:"cartao",l:"💳 Taxas Cartão"}].map(g=>(
              <button key={g.id} onClick={()=>setFinTab(g.id)} style={{padding:".42rem .85rem",borderRadius:".5rem",fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"none",cursor:"pointer",background:finTab===g.id?"linear-gradient(135deg,#4f5ef0,#8b44f0)":"var(--pill)",color:finTab===g.id?"#fff":"var(--navoff)",transition:"all .2s",whiteSpace:"nowrap"}}>{g.l}</button>
            ))}
          </div>
        )}

        {/* ══ FINANCEIRO · RESUMO (compilado a pagar + a receber + caixa) ══ */}
        {tab==="financeiro"&&finTab==="resumo"&&(()=>{
          const saldoCaixa=net;
          const abertoRec=(receivables||[]).filter(r=>!r.paid);
          const abertoPag=(payables||[]).filter(p=>!p.paid);
          const abertoPagAll=[...abertoPag,...orderPayables];
          const totalRec=abertoRec.reduce((a,r)=>a+(parseFloat(r.value)||0),0);
          const totalPag=abertoPagAll.reduce((a,p)=>a+Math.max(0,(parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0)),0);
          const hoje=new Date();
          const venRec=abertoRec.filter(r=>r.due_date&&new Date(r.due_date)<hoje);
          const venPag=abertoPag.filter(p=>p.due_date&&new Date(p.due_date)<hoje);
          const projetado=saldoCaixa+totalRec-totalPag;
          const proximos=[...abertoRec.map(r=>({...r,_t:"receber"})),...abertoPagAll.map(p=>({...p,_t:"pagar"}))].filter(x=>x.due_date).sort((a,b)=>new Date(a.due_date)-new Date(b.due_date)).slice(0,8);
          return(
            <div style={{animation:"fadeUp .4s ease"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)",marginBottom:".75rem"}}>📊 Resumo Financeiro</h2>
              <div style={{background:projetado>=0?"linear-gradient(135deg,#10b98120,#10b98108)":"linear-gradient(135deg,#f5656520,#f5656508)",border:`1px solid ${projetado>=0?"#10b98140":"#f5656540"}`,borderRadius:".75rem",padding:"1rem",marginBottom:".75rem",textAlign:"center"}}>
                <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".5px"}}>Saldo projetado (caixa + a receber − a pagar)</div>
                <div style={{fontSize:"1.7rem",fontWeight:800,color:projetado>=0?"#10b981":"#f56565",fontFamily:"'Syne',sans-serif"}}>{fmt(projetado)}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:".6rem",marginBottom:".75rem"}}>
                <KCard label="Saldo em caixa" value={fmt(saldoCaixa)} color={saldoCaixa>=0?"#10b981":"#f56565"}/>
                <KCard label="Total a receber" value={fmt(totalRec)} sub={fmtN(abertoRec.length)+" conta(s)"} color="#4f5ef0"/>
                <KCard label="Total a pagar" value={fmt(totalPag)} sub={fmtN(abertoPagAll.length)+" conta(s)"} color="#f59e0b"/>
                <KCard label="Vencidas" value={fmtN(venRec.length+venPag.length)} sub={"rec "+fmt(venRec.reduce((a,r)=>a+r.value,0))+" · pag "+fmt(venPag.reduce((a,p)=>a+p.value,0))} color="#f56565"/>
              </div>
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
                <div style={{padding:".7rem 1rem",borderBottom:"1px solid var(--sep)",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)"}}>📅 Próximos vencimentos</div>
                {proximos.length===0
                  ?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2rem 0",fontSize:".8rem"}}>Nenhum vencimento futuro registrado.</p>
                  :proximos.map(x=>{
                    const ov=x.due_date&&new Date(x.due_date)<hoje;
                    return(
                      <div key={x._t+x.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".65rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
                            <Badge color={x._t==="receber"?"#10b981":"#f59e0b"} sm>{x._t==="receber"?"↗ Receber":"↘ Pagar"}</Badge>
                            <span style={{fontWeight:600,fontSize:".8rem",color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{x.description}</span>
                          </div>
                          <div style={{fontSize:".66rem",color:ov?"#f56565":"var(--tx5)",marginTop:".15rem"}}>{(x.client_name||x.supplier_name||"—")+" · Vcto "+new Date(x.due_date+"T00:00:00").toLocaleDateString("pt-BR")+(ov?" · VENCIDO":"")}</div>
                        </div>
                        <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",fontSize:".9rem",color:x._t==="receber"?"#10b981":"#f59e0b",marginLeft:".75rem"}}>{(x._t==="receber"?"+":"−")+fmt(x.value)}</span>
                      </div>
                    );
                  })
                }
              </div>
              <div style={{fontSize:".68rem",color:"var(--tx6)",marginTop:".6rem",textAlign:"center"}}>Compilado de contas a pagar, a receber e saldo do caixa. Para projeção por período (30/60/90 dias), use 💰 Fluxo Caixa no Dashboard.</div>
            </div>
          );
        })()}

        {/* ══ FINANCEIRO · PROJEÇÃO DE FLUXO DE CAIXA NO TEMPO ══ */}
        {tab==="financeiro"&&finTab==="projecao"&&(()=>{
          const hoje=new Date();hoje.setHours(0,0,0,0);
          const dias=(dateStr)=>{if(!dateStr)return 0;const d=new Date(dateStr+(dateStr.length<=10?"T00:00:00":""));if(isNaN(d))return 0;return Math.floor((d-hoje)/86400000);};
          // Entradas previstas: recebíveis em aberto (por vencimento)
          const recsAb=(receivables||[]).filter(r=>!r.paid).map(r=>({d:dias(r.due_date),v:parseFloat(r.value)||0}));
          // Saídas previstas: contas a pagar em aberto (saldo restante, por vencimento) + pedidos a fornecedor (saldo restante)
          const paysAb=(payables||[]).filter(p=>!p.paid).map(p=>({d:dias(p.due_date),v:Math.max(0,(parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0))}));
          const ordsAb=orderPayables.map(o=>({d:dias(o.due_date),v:Math.max(0,(parseFloat(o.value)||0)-(parseFloat(o.paid_amount)||0))}));
          const saidasAb=[...paysAb,...ordsAb];
          const faixas=[{l:"Em atraso",min:-99999,max:-1},{l:"Próx. 7 dias",min:0,max:7},{l:"8-15 dias",min:8,max:15},{l:"16-30 dias",min:16,max:30},{l:"31-60 dias",min:31,max:60},{l:"61-90 dias",min:61,max:90},{l:"+90 dias",min:91,max:99999}];
          const inFx=(min,max)=>recsAb.filter(x=>x.d>=min&&x.d<=max).reduce((a,x)=>a+x.v,0);
          const outFx=(min,max)=>saidasAb.filter(x=>x.d>=min&&x.d<=max).reduce((a,x)=>a+x.v,0);
          let saldo=net; // começa do caixa atual
          const linhas=faixas.map(f=>{const ent=inFx(f.min,f.max);const sai=outFx(f.min,f.max);const liq=ent-sai;saldo+=liq;return {...f,ent,sai,liq,saldo};});
          const totEnt=recsAb.reduce((a,x)=>a+x.v,0),totSai=saidasAb.reduce((a,x)=>a+x.v,0);
          const menorSaldo=Math.min(net,...linhas.map(l=>l.saldo));
          const faixaNeg=linhas.find(l=>l.saldo<0);
          return(
          <div style={{animation:"fadeUp .4s ease"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)",marginBottom:".25rem"}}>📈 Projeção de Fluxo de Caixa</h2>
            <p style={{fontSize:".72rem",color:"var(--tx5)",marginBottom:".75rem"}}>Partindo do caixa atual, o que deve entrar (recebíveis) e sair (contas + pedidos) por período de vencimento.</p>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".6rem",marginBottom:".75rem"}}>
              <KCard label="Caixa hoje" value={fmt(net)} color={net>=0?"#10b981":"#f56565"}/>
              <KCard label="A entrar" value={fmt(totEnt)} color="#4f5ef0"/>
              <KCard label="A sair" value={fmt(totSai)} color="#f59e0b"/>
              <KCard label="Saldo final previsto" value={fmt(net+totEnt-totSai)} color={(net+totEnt-totSai)>=0?"#10b981":"#f56565"}/>
            </div>
            {menorSaldo<0&&faixaNeg&&(
              <div style={{background:"#f5656512",border:"1px solid #f5656540",borderRadius:".6rem",padding:".7rem .9rem",marginBottom:".75rem",fontSize:".76rem",color:"#f56565",fontWeight:600}}>
                ⚠️ Alerta de caixa: pela projeção, seu saldo fica negativo ({fmt(menorSaldo)}) no período “{faixaNeg.l}”. Antecipe recebimentos ou renegocie pagamentos desse período.
              </div>
            )}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
              <div style={{display:"grid",gridTemplateColumns:"1.3fr 1fr 1fr 1fr 1.1fr",gap:".3rem",padding:".55rem .85rem",borderBottom:"1px solid var(--bdr)",background:"var(--pill)",fontSize:".62rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".03em",fontWeight:700}}>
                <span>Período</span><span style={{textAlign:"right"}}>Entradas</span><span style={{textAlign:"right"}}>Saídas</span><span style={{textAlign:"right"}}>Líquido</span><span style={{textAlign:"right"}}>Saldo acum.</span>
              </div>
              {linhas.map(l=>(
                <div key={l.l} style={{display:"grid",gridTemplateColumns:"1.3fr 1fr 1fr 1fr 1.1fr",gap:".3rem",padding:".55rem .85rem",borderBottom:"1px solid var(--sep)",fontSize:".75rem",alignItems:"center",background:l.l==="Em atraso"&&(l.ent>0||l.sai>0)?"#f5656506":"transparent"}}>
                  <span style={{fontWeight:600,color:"var(--tx2)"}}>{l.l}</span>
                  <span style={{textAlign:"right",color:l.ent>0?"#10b981":"var(--tx6)",fontWeight:l.ent>0?600:400}}>{l.ent>0?fmt(l.ent):"—"}</span>
                  <span style={{textAlign:"right",color:l.sai>0?"#f59e0b":"var(--tx6)",fontWeight:l.sai>0?600:400}}>{l.sai>0?fmt(l.sai):"—"}</span>
                  <span style={{textAlign:"right",fontWeight:700,color:l.liq>0?"#10b981":l.liq<0?"#f56565":"var(--tx6)"}}>{l.liq!==0?(l.liq>0?"+":"")+fmt(l.liq):"—"}</span>
                  <span style={{textAlign:"right",fontWeight:800,fontFamily:"'Syne',sans-serif",color:l.saldo>=0?"var(--tx)":"#f56565"}}>{fmt(l.saldo)}</span>
                </div>
              ))}
            </div>
            <p style={{fontSize:".66rem",color:"var(--tx6)",marginTop:".6rem"}}>💡 A coluna “Saldo acum.” mostra como o caixa evolui somando cada período. Se ficar vermelho, é onde o caixa aperta. Itens sem data de vencimento entram em “Próx. 7 dias”.</p>
          </div>
          );
        })()}

        {/* ══ FINANCEIRO · SAÚDE FINANCEIRA / DRE ══ */}
        {tab==="financeiro"&&finTab==="saude"&&(()=>{
          const hoje=new Date();
          const lim30=new Date();lim30.setDate(lim30.getDate()+30);
          // Despesas por categoria (saídas do caixa)
          const saidasPorCat={};
          (cashTx||[]).filter(t=>t.type==="saida").forEach(t=>{const k=t.category||"Outros";saidasPorCat[k]=(saidasPorCat[k]||0)+(parseFloat(t.value)||0);});
          const catList=Object.entries(saidasPorCat).sort((a,b)=>b[1]-a[1]);
          const maxCat=catList.length?catList[0][1]:1;
          // Custos fixos mensais estimados: recorrentes (dedup por descrição) + categorias fixas em aberto
          const fixCats=["aluguel","salario","servico","imposto"];
          const recMap={};
          (payables||[]).filter(p=>p.recurring).forEach(p=>{recMap[p.description]=parseFloat(p.value)||0;});
          const custoFixoRec=Object.values(recMap).reduce((a,v)=>a+v,0);
          const custoFixoOutros=(payables||[]).filter(p=>!p.recurring&&fixCats.includes(p.category)&&!p.paid).reduce((a,p)=>a+(parseFloat(p.value)||0),0);
          const custoFixo=custoFixoRec+custoFixoOutros;
          const pag30=(payables||[]).filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)<=lim30).reduce((a,p)=>a+(parseFloat(p.value)||0),0);
          const vencidasRec=(receivables||[]).filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)<hoje).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
          // Qualidade do recebimento: inadimplência, DSO (prazo médio) e carteira de inadimplentes
          const recAbertos=(receivables||[]).filter(r=>!r.paid);
          const totalReceber=recAbertos.reduce((a,r)=>a+(parseFloat(r.value)||0),0);
          const inadimplencia=totalReceber>0?(vencidasRec/totalReceber)*100:0;
          const _d90=new Date();_d90.setDate(_d90.getDate()-90);
          const vendasPrazo90=(receivables||[]).filter(r=>r.created_at&&new Date(r.created_at)>=_d90).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
          const dso=vendasPrazo90>0?Math.round((totalReceber/vendasPrazo90)*90):null;
          const inadMap={};
          recAbertos.filter(r=>r.due_date&&new Date(r.due_date)<hoje).forEach(r=>{
            const key=r.client_name||r.client_id||"Sem cliente";
            if(!inadMap[key])inadMap[key]={name:r.client_name||"Sem cliente",client_id:r.client_id,total:0,count:0,maxDias:0,oldest:null};
            const o=inadMap[key];o.total+=parseFloat(r.value)||0;o.count++;
            const dd=Math.floor((hoje-new Date(r.due_date+"T00:00:00"))/86400000);
            if(dd>=o.maxDias){o.maxDias=dd;o.oldest=r;}
          });
          const carteiraInad=Object.values(inadMap).sort((a,b)=>b.total-a.total);
          const breakEven=margin>0?custoFixo/(margin/100):0;
          const coberturaFixo=custoFixo>0?net/custoFixo:(net>0?99:0);
          const liquidezOk=net>=pag30;
          const margStatus=margin>=20?{c:"#10b981",t:"Saudável"}:margin>=10?{c:"#f59e0b",t:"Atenção"}:{c:"#f56565",t:"Baixa"};
          // score 0-100 — inclui qualidade do recebimento (inadimplência)
          let score=0;
          if(net>0)score+=25;
          if(margin>=20)score+=20; else if(margin>=10)score+=10;
          if(liquidezOk)score+=20;
          if(inadimplencia<=5)score+=20; else if(inadimplencia<=15)score+=12; else if(inadimplencia<=25)score+=5;
          if(coberturaFixo>=1)score+=15;
          score=Math.min(100,Math.round(score));
          const scoreColor=score>=70?"#10b981":score>=40?"#f59e0b":"#f56565";
          const scoreLabel=score>=70?"Saudável":score>=40?"Requer atenção":"Crítico";
          // Comparativo mês atual vs mês anterior (campo date = DD/MM/YYYY)
          const _now=new Date();
          const curMY=String(_now.getMonth()+1).padStart(2,"0")+"/"+_now.getFullYear();
          const _prev=new Date(_now.getFullYear(),_now.getMonth()-1,1);
          const prevMY=String(_prev.getMonth()+1).padStart(2,"0")+"/"+_prev.getFullYear();
          const monthAgg=(my)=>{
            const tx=(cashTx||[]).filter(t=>((t.date||"").split("/").slice(1).join("/"))===my);
            const r=tx.filter(t=>t.type==="entrada").reduce((a,t)=>a+(parseFloat(t.value)||0),0);
            const c=tx.filter(t=>t.type==="saida").reduce((a,t)=>a+(parseFloat(t.value)||0),0);
            return{r,c,p:r-c,m:r>0?((r-c)/r)*100:0};
          };
          const mCur=monthAgg(curMY),mPrev=monthAgg(prevMY);
          const delta=(a,b)=>b!==0?((a-b)/Math.abs(b))*100:(a>0?100:0);
          const cmp=[
            {l:"Receita",cur:mCur.r,prev:mPrev.r,money:true},
            {l:"Custos",cur:mCur.c,prev:mPrev.c,money:true,inv:true},
            {l:"Lucro",cur:mCur.p,prev:mPrev.p,money:true},
            {l:"Margem",cur:mCur.m,prev:mPrev.m,money:false},
          ];
          const dre=[
            {l:"(+) Receita bruta",v:cashIn,c:"#10b981"},
            {l:"(−) Custos e despesas",v:-cashOut,c:"#f56565"},
            {l:"(=) Resultado líquido",v:net,c:net>=0?"#4f5ef0":"#f56565",bold:true},
          ];
          return(
            <div style={{animation:"fadeUp .4s ease"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)",marginBottom:".75rem"}}>💚 Saúde Financeira</h2>
              {/* Score geral */}
              <div style={{background:`linear-gradient(135deg,${scoreColor}18,${scoreColor}06)`,border:`1px solid ${scoreColor}40`,borderRadius:".75rem",padding:"1rem",marginBottom:".75rem",display:"flex",alignItems:"center",gap:"1rem",flexWrap:"wrap"}}>
                <div style={{width:64,height:64,borderRadius:"50%",border:`5px solid ${scoreColor}`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                  <span style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.2rem",color:scoreColor}}>{score}</span>
                </div>
                <div style={{flex:1,minWidth:140}}>
                  <div style={{fontWeight:800,fontSize:"1rem",color:scoreColor,fontFamily:"'Syne',sans-serif"}}>{scoreLabel}</div>
                  <div style={{fontSize:".7rem",color:"var(--tx5)"}}>Índice de saúde com base em lucro, margem, liquidez e inadimplência.</div>
                </div>
              </div>
              {/* KPIs */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:".6rem",marginBottom:".75rem"}}>
                <KCard label="Receita" value={fmt(cashIn)} color="#10b981"/>
                <KCard label="Custos" value={fmt(cashOut)} sub={cashIn>0?fmtPct((cashOut/cashIn)*100)+" da receita":""} color="#f56565"/>
                <KCard label="Lucro líquido" value={fmt(net)} color={net>=0?"#4f5ef0":"#f56565"}/>
                <KCard label="Margem" value={fmtPct(margin)} sub={margStatus.t} color={margStatus.c}/>
                <KCard label="Markup" value={fmtPct(mrkp)} color="#f59e0b"/>
                <KCard label="Inadimplência" value={fmtPct(inadimplencia)} sub={inadimplencia<=15?"saudável (<15%)":inadimplencia<=25?"atenção":"crítico"} color={inadimplencia<=15?"#10b981":inadimplencia<=25?"#f59e0b":"#f56565"}/>
                <KCard label="Prazo médio (DSO)" value={dso!==null?dso+"d":"—"} sub={dso!==null?(dso<=30?"saudável (<30d)":dso<=45?"atenção":"alto"):"sem vendas a prazo"} color={dso===null?"#8890b0":dso<=30?"#10b981":dso<=45?"#f59e0b":"#f56565"}/>
              </div>
              {carteiraInad.length>0&&(
                <div style={{background:"var(--card)",border:"1px solid #f5656540",borderRadius:".75rem",overflow:"hidden",marginBottom:".75rem"}}>
                  <div style={{padding:".7rem 1rem",borderBottom:"1px solid var(--sep)",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:".4rem"}}>
                    <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"#f56565"}}>🚨 Carteira de Inadimplentes</span>
                    <span style={{fontSize:".68rem",color:"var(--tx5)"}}>{carteiraInad.length} cliente(s) · {fmt(carteiraInad.reduce((a,c)=>a+c.total,0))} em atraso</span>
                  </div>
                  {carteiraInad.map((c,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".6rem 1rem",borderBottom:"1px solid var(--sep)",gap:".5rem"}}>
                      <div style={{minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:".8rem",color:"var(--tx)"}}>{c.name}</div>
                        <div style={{fontSize:".66rem",color:"#f56565"}}>{c.count} conta(s) vencida(s) · atraso de até {c.maxDias} dia{c.maxDias!==1?"s":""}</div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:".5rem",flexShrink:0}}>
                        <span style={{fontWeight:800,fontFamily:"'Syne',sans-serif",color:"#f56565",fontSize:".88rem"}}>{fmt(c.total)}</span>
                        {c.oldest&&<button onClick={()=>cobrarWhatsApp(c.oldest)} title="Cobrar no WhatsApp" style={{background:"#25D36615",border:"1px solid #25D36640",borderRadius:".4rem",padding:".25rem .5rem",color:"#1ba34e",fontSize:".7rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>📲 Cobrar</button>}
                      </div>
                    </div>
                  ))}
                  <div style={{padding:".5rem 1rem",fontSize:".64rem",color:"var(--tx6)"}}>Clientes com contas a receber vencidas. Quitar ou negociar essas contas melhora direto seu índice de saúde.</div>
                </div>
              )}
              {/* Break-even */}
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx2)",marginBottom:".6rem"}}>⚖️ Ponto de Equilíbrio (Break-even)</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:".5rem"}}>
                  <div style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".55rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase"}}>Custos fixos/mês (estim.)</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:"#f59e0b",fontFamily:"'Syne',sans-serif"}}>{fmt(custoFixo)}</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".55rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase"}}>Faturamento p/ equilíbrio</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:"#8b44f0",fontFamily:"'Syne',sans-serif"}}>{margin>0?fmt(breakEven):"—"}</div>
                  </div>
                  <div style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".55rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase"}}>Cobertura dos fixos</div>
                    <div style={{fontSize:".95rem",fontWeight:700,color:coberturaFixo>=1?"#10b981":"#f56565",fontFamily:"'Syne',sans-serif"}}>{custoFixo>0?coberturaFixo.toFixed(1)+"x":"—"}</div>
                  </div>
                </div>
                <div style={{fontSize:".66rem",color:"var(--tx6)",marginTop:".5rem"}}>Custos fixos = contas a pagar recorrentes + aluguel/salário/serviço/imposto em aberto. Equilíbrio = custos fixos ÷ margem.</div>
              </div>
              {/* Comparativo mensal */}
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx2)",marginBottom:".65rem"}}>📈 Este mês vs mês anterior</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:".5rem"}}>
                  {cmp.map(x=>{
                    const d=delta(x.cur,x.prev);
                    const good=x.inv?d<=0:d>=0;
                    const arrow=d>0.5?"▲":d<-0.5?"▼":"▬";
                    return(
                      <div key={x.l} style={{background:"var(--sumbox)",borderRadius:".5rem",padding:".6rem",border:"1px solid var(--bdr)"}}>
                        <div style={{fontSize:".62rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{x.l}</div>
                        <div style={{fontSize:".95rem",fontWeight:700,color:"var(--tx)",fontFamily:"'Syne',sans-serif"}}>{x.money?fmt(x.cur):fmtPct(x.cur)}</div>
                        <div style={{fontSize:".64rem",color:good?"#10b981":"#f56565",fontWeight:600,marginTop:".1rem"}}>{arrow} {Math.abs(d).toFixed(0)}% <span style={{color:"var(--tx6)"}}>vs {x.money?fmt(x.prev):fmtPct(x.prev)}</span></div>
                      </div>
                    );
                  })}
                </div>
              </div>
              {/* DRE simplificado */}
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden",marginBottom:".75rem"}}>
                <div style={{padding:".7rem 1rem",borderBottom:"1px solid var(--sep)",fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)"}}>📑 DRE Simplificado</div>
                {dre.map(d=>(
                  <div key={d.l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:".6rem 1rem",borderBottom:"1px solid var(--sep)",background:d.bold?"var(--sumbox)":"transparent"}}>
                    <span style={{fontSize:".8rem",color:"var(--tx3)",fontWeight:d.bold?700:500}}>{d.l}</span>
                    <span style={{fontSize:".88rem",fontWeight:700,color:d.c,fontFamily:"'Syne',sans-serif"}}>{fmt(d.v)}</span>
                  </div>
                ))}
              </div>
              {/* Despesas por categoria */}
              <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>📉 Despesas por categoria</div>
                {catList.length===0
                  ?<p style={{color:"var(--tx5)",fontSize:".78rem",textAlign:"center",padding:"1rem 0"}}>Sem saídas registradas ainda.</p>
                  :catList.map(([k,v])=>(
                    <div key={k} style={{marginBottom:".5rem"}}>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:".72rem",marginBottom:".2rem"}}>
                        <span style={{color:"var(--tx3)",fontWeight:600}}>{k}</span>
                        <span style={{color:"var(--tx4)",fontWeight:700}}>{fmt(v)} · {cashOut>0?fmtPct((v/cashOut)*100):"0%"}</span>
                      </div>
                      <div style={{height:6,background:"var(--sumbox)",borderRadius:3,overflow:"hidden"}}>
                        <div style={{height:"100%",width:`${maxCat>0?(v/maxCat)*100:0}%`,background:"linear-gradient(90deg,#f59e0b,#f56565)",borderRadius:3}}/>
                      </div>
                    </div>
                  ))
                }
              </div>
              {/* Indicadores de saúde */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:".6rem"}}>
                <div style={{background:"var(--card)",border:`1px solid ${liquidezOk?"#10b98140":"#f5656540"}`,borderRadius:".6rem",padding:".75rem"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".25rem"}}><Badge color={liquidezOk?"#10b981":"#f56565"} sm>{liquidezOk?"✓ OK":"⚠ Risco"}</Badge><span style={{fontSize:".78rem",fontWeight:700,color:"var(--tx2)"}}>Liquidez (30 dias)</span></div>
                  <div style={{fontSize:".7rem",color:"var(--tx5)"}}>Caixa {fmt(net)} {liquidezOk?"cobre":"NÃO cobre"} as contas a pagar dos próximos 30 dias ({fmt(pag30)}).</div>
                </div>
                <div style={{background:"var(--card)",border:`1px solid ${vencidasRec===0?"#10b98140":"#f5656540"}`,borderRadius:".6rem",padding:".75rem"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".25rem"}}><Badge color={vencidasRec===0?"#10b981":"#f56565"} sm>{vencidasRec===0?"✓ Em dia":"⚠ Vencido"}</Badge><span style={{fontSize:".78rem",fontWeight:700,color:"var(--tx2)"}}>Inadimplência</span></div>
                  <div style={{fontSize:".7rem",color:"var(--tx5)"}}>{vencidasRec===0?"Nenhum recebível vencido. ":fmt(vencidasRec)+" a receber em atraso."}</div>
                </div>
                <div style={{background:"var(--card)",border:`1px solid ${margStatus.c}40`,borderRadius:".6rem",padding:".75rem"}}>
                  <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".25rem"}}><Badge color={margStatus.c} sm>{margStatus.t}</Badge><span style={{fontSize:".78rem",fontWeight:700,color:"var(--tx2)"}}>Margem de lucro</span></div>
                  <div style={{fontSize:".7rem",color:"var(--tx5)"}}>Margem atual de {fmtPct(margin)}. {margin>=20?"Excelente.":margin>=10?"Pode melhorar.":"Revise preços/custos."}</div>
                </div>
              </div>
              <div style={{fontSize:".66rem",color:"var(--tx6)",marginTop:".7rem",textAlign:"center"}}>Indicadores calculados a partir do caixa, contas a pagar/receber e estoque. Para projeção de fluxo use 💰 Fluxo Caixa no Dashboard.</div>
            </div>
          );
        })()}

        {/* ══ FINANCEIRO · CONTAS A PAGAR ══ */}
        {tab==="financeiro"&&finTab==="pagar"&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>📄 Contas a Pagar</h2>
              <XBtn rows={(payables||[]).map(p=>({Categoria:p.category||"—",Descrição:p.description,Fornecedor:p.supplier_name||"—",Valor:fmt(p.value),Vencimento:p.due_date||"—",Status:p.paid?"Pago":"Pendente"}))} name="contas-a-pagar-caixapro" sheet="A Pagar"/>
            </div>
            {(()=>{
              const pays=payables||[];
              const outP=p=>Math.max(0,(parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0));
              const paidP=pays.filter(p=>p.paid);
              const dividaPaga=paidP.reduce((a,p)=>a+(parseFloat(p.value)||0),0);
              const capitalPago=pays.reduce((a,p)=>a+(parseFloat(p.paid_amount)||0),0)+(orders||[]).filter(o=>o.status!=="cancelado"&&o.status!=="perdido").reduce((a,o)=>a+(parseFloat(o.initial_value)||0)+(parseFloat(o.remaining_paid)||0),0);
              const economia=paidP.reduce((a,p)=>{const v=parseFloat(p.value)||0;const pa=p.paid_amount!=null?(parseFloat(p.paid_amount)||0):v;return a+Math.max(0,v-pa);},0);
              const abertoSum=pays.filter(p=>!p.paid).reduce((a,p)=>a+outP(p),0)+orderPayables.reduce((a,o)=>a+outP(o),0);
              const vencList=pays.filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)<new Date());
              return(
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(108px,1fr))",gap:".6rem",marginBottom:".75rem"}}>
                <KCard label="Em aberto" value={fmtN(pays.filter(p=>!p.paid).length+orderPayables.length)} sub={fmt(abertoSum)} color="#f59e0b"/>
                <KCard label="Vencidas" value={fmtN(vencList.length)} sub={fmt(vencList.reduce((a,p)=>a+outP(p),0))} color="#f56565"/>
                <KCard label="Capital pago" value={fmt(capitalPago)} color="#10b981"/>
                <KCard label="Economia" value={fmt(economia)} sub={dividaPaga>0?fmtPct((economia/dividaPaga)*100)+" da dívida":"descontos"} color={economia>0.005?"#10b981":"#8890b0"}/>
              </div>
              );
            })()}
            {canEdit&&<div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>➕ Nova conta a pagar (boleto, aluguel, parcela...)</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:".5rem",marginBottom:".5rem"}}>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Categoria</div>
                  <select value={payForm.category} onChange={e=>setPayForm(f=>({...f,category:e.target.value}))} style={IS}>
                    <option value="aluguel">🏠 Aluguel</option>
                    <option value="emprestimo">🏦 Empréstimo/Financiamento</option>
                    <option value="compra">📦 Parcela de compra</option>
                    <option value="fornecedor">🚚 Fornecedor</option>
                    <option value="imposto">🧾 Imposto/Taxa</option>
                    <option value="salario">👤 Salário/Comissão</option>
                    <option value="servico">🔧 Serviço/Assinatura</option>
                    <option value="outro">📋 Outro</option>
                  </select>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Vencimento</div>
                  <input type="date" value={payForm.due_date} onChange={e=>setPayForm(f=>({...f,due_date:e.target.value}))} style={IS}/>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Fornecedor/Credor</div>
                  <input value={payForm.supplier_name||""} onChange={e=>setPayForm(f=>({...f,supplier_name:e.target.value}))} placeholder="opcional" style={IS}/>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:".5rem",marginBottom:".65rem"}}>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Descrição *</div>
                  <input value={payForm.description||""} onChange={e=>setPayForm(f=>({...f,description:e.target.value}))} placeholder="Ex: Aluguel maio, Parcela 3/12 empréstimo..." style={IS}/>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Valor (R$) *{payForm.installments>1?" — por parcela":""}</div>
                  <input type="number" min="0" step="0.01" value={payForm.value||""} onChange={e=>setPayForm(f=>({...f,value:e.target.value}))} placeholder="0,00" style={IS}/>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Parcelas</div>
                  <input type="number" min="1" max="60" step="1" value={payForm.installments} onChange={e=>setPayForm(f=>({...f,installments:e.target.value}))} placeholder="1" style={IS}/>
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:".75rem",flexWrap:"wrap",marginBottom:".65rem"}}>
                <label style={{display:"flex",alignItems:"center",gap:".4rem",cursor:payForm.installments>1?"not-allowed":"pointer",fontSize:".78rem",fontWeight:600,color:payForm.installments>1?"var(--tx6)":(payForm.recurring?"#8b44f0":"var(--tx3)"),opacity:payForm.installments>1?.5:1}}>
                  <input type="checkbox" disabled={payForm.installments>1} checked={!!payForm.recurring&&payForm.installments<=1} onChange={e=>setPayForm(f=>({...f,recurring:e.target.checked}))} style={{width:16,height:16,accentColor:"#8b44f0"}}/>
                  🔁 Recorrente mensal (gera a próxima ao pagar)
                </label>
                {payForm.installments>1&&<span style={{fontSize:".7rem",color:"#4f5ef0",fontWeight:600}}>📦 Cria {payForm.installments} parcelas mensais de {fmt(parseFloat(payForm.value)||0)}{(parseFloat(payForm.value)||0)>0?" · total "+fmt((parseFloat(payForm.value)||0)*(parseInt(payForm.installments)||1)):""}</span>}
              </div>
              <Btn v="ok" onClick={addPayable}><Ic n="save" s={13}/>{payForm.installments>1?`Registrar ${payForm.installments} parcelas`:"Registrar conta a pagar"}</Btn>
            </div>}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
              {orderPayables.map(o=>(
                <div key={o.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",background:"#0891b209"}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".15rem",flexWrap:"wrap"}}>
                      <span style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)"}}>{o.description}</span>
                      <Badge color="#0891b2" sm>Pedido a fornecedor</Badge>
                    </div>
                    <div style={{fontSize:".67rem",color:"var(--tx5)"}}>Pagamento na chegada da mercadoria · gerido em Estoque ▸ Pedidos</div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:".55rem",flexShrink:0,marginLeft:".75rem"}}>
                    <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:"#0891b2",fontSize:".9rem"}}>{fmt(o.value)}</span>
                    <button onClick={()=>{setTab("estoque");setEstTab("pedidos");}} style={{background:"#0891b215",border:"1px solid #0891b230",borderRadius:".4rem",padding:".28rem .6rem",color:"#0891b2",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>Ver pedido →</button>
                  </div>
                </div>
              ))}
              {(payables||[]).length===0&&orderPayables.length===0
                ?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2.5rem 0",fontSize:".8rem"}}>Nenhuma conta a pagar registrada. Use o formulário acima.</p>
                :[...(payables||[]),...orderPayables].map(p=>{
                  const total=parseFloat(p.value)||0;
                  const already=parseFloat(p.paid_amount)||0;
                  const falta=Math.max(0,total-already);
                  const partial=!p.paid&&already>0.005;
                  const ov=!p.paid&&p.due_date&&new Date(p.due_date)<new Date();
                  const dv=p.due_date?Math.ceil((new Date(p.due_date+"T23:59:59")-new Date())/86400000):null;
                  const catIcon=p._order?(p._dropship?"🔄":"📦"):({aluguel:"🏠",emprestimo:"🏦",compra:"📦",fornecedor:"🚚",imposto:"🧾",salario:"👤",servico:"🔧",outro:"📋"}[p.category]||"📋");
                  return(
                    <div key={p.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",background:ov?"#f5656508":partial?"#4f5ef008":"transparent"}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".15rem",flexWrap:"wrap"}}>
                          <span style={{fontSize:".9rem"}}>{catIcon}</span>
                          <span style={{fontWeight:700,fontSize:".83rem",color:p.paid?"var(--tx5)":"var(--tx)",textDecoration:p.paid?"line-through":"none"}}>{p.description}</span>
                          {p.paid&&<Badge color="#10b981" sm>✅ Pago</Badge>}
                          {partial&&<Badge color="#4f5ef0" sm>➗ Parcial</Badge>}
                          {p._order&&!partial&&<Badge color="#0891b2" sm>pedido</Badge>}
                          {ov&&<Badge color="#f56565" sm>⚠️ Vencido</Badge>}
                          {!p.paid&&!ov&&dv!==null&&dv<=3&&<Badge color="#f59e0b" sm>{"⏰ Vence em "+dv+"d"}</Badge>}
                          {p.recurring&&<Badge color="#8b44f0" sm>🔁 Recorrente</Badge>}
                        </div>
                        <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{p.supplier_name||"—"}{p.due_date&&" · Vcto: "+new Date(p.due_date+"T00:00:00").toLocaleDateString("pt-BR")}{p.paid_date&&" · Pago "+p.paid_date}</div>
                        {partial&&<div style={{fontSize:".66rem",color:"#4f5ef0",fontWeight:600,marginTop:".15rem"}}>Pago {fmt(already)} de {fmt(total)} · falta {fmt(falta)}</div>}
                        {p.paid&&p.paid_amount!=null&&((parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0))>0.005&&<div style={{fontSize:".66rem",color:"#10b981",fontWeight:600,marginTop:".1rem"}}>💚 Economia de {fmt((parseFloat(p.value)||0)-(parseFloat(p.paid_amount)||0))} (quitação com desconto)</div>}
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:".55rem",flexShrink:0,marginLeft:".75rem"}}>
                        <div style={{textAlign:"right"}}>
                          <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:p.paid?"#10b981":ov?"#f56565":"#f59e0b",fontSize:".9rem"}}>{fmt(p.paid?total:falta)}</span>
                          {(partial||(p._order&&already>0.005))&&<div style={{fontSize:".6rem",color:"var(--tx6)"}}>de {fmt(total)}</div>}
                        </div>
                        {canEdit&&!p.paid&&<button onClick={()=>{setPayTarget(p);setPayAmount(falta.toFixed(2));setPayQuitar(false);}} style={{background:"#f59e0b15",border:"1px solid #f59e0b30",borderRadius:".4rem",padding:".28rem .6rem",color:"#f59e0b",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>💸 Pagar</button>}
                        {canDelete&&!p._order&&<button onClick={()=>deletePayable(p.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem",cursor:"pointer"}}><Ic n="trash" s={13}/></button>}
                      </div>
                    </div>
                  );
                })
              }
            </div>
          </div>
        )}

        {/* ══ FINANCEIRO · TAXAS DE CARTÃO ══ */}
        {tab==="financeiro"&&finTab==="cartao"&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>💳 Taxas de Cartão</h2>
              {canEdit&&<Btn sm onClick={()=>{setLocalTaxes(JSON.parse(JSON.stringify(cardTaxes)));setShowCardConfig(true);}}><Ic n="edit" s={12}/>Editar Taxas</Btn>}
            </div>
            <div style={{fontSize:".72rem",color:"var(--sub)",marginBottom:".75rem"}}>Taxas aplicadas automaticamente nas vendas com cartão (débito, crédito à vista e parcelado). Clique em Editar para ajustar os percentuais.</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:".65rem"}}>
              {CARD_BRANDS.map(b=>{
                const t=cardTaxes[b]||DEFAULT_TAXES[b]||{};
                return(
                  <div key={b} style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".6rem",padding:".8rem"}}>
                    <div style={{fontWeight:700,fontSize:".85rem",color:"var(--tx)",marginBottom:".5rem"}}>{b}</div>
                    <div style={{display:"flex",flexDirection:"column",gap:".25rem",fontSize:".72rem"}}>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--tx5)"}}>Débito</span><span style={{fontWeight:600,color:"var(--tx2)"}}>{fmtPct(t.debito||0)}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--tx5)"}}>Crédito à vista</span><span style={{fontWeight:600,color:"var(--tx2)"}}>{fmtPct(t.vista||0)}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--tx5)"}}>Parcelado 2-6x</span><span style={{fontWeight:600,color:"var(--tx2)"}}>{fmtPct(t.parc2a6||0)}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--tx5)"}}>Parcelado 7-12x</span><span style={{fontWeight:600,color:"var(--tx2)"}}>{fmtPct(t.parc7a12||0)}</span></div>
                      {(cardTaxes.__opts||{}).antecip&&<div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid var(--sep)",paddingTop:".2rem",marginTop:".1rem"}}><span style={{color:"#8b44f0"}}>⚡ Antecipação</span><span style={{fontWeight:700,color:"#8b44f0"}}>+{fmtPct(t.antecip||0)}</span></div>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ══ CAIXA ══ */}
        {tab==="financeiro"&&finTab==="caixa"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Caixa</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={fCash.map(x=>({Data:x.date,Descrição:x.description,Tipo:x.type==="entrada"?"Entrada":"Saída",Valor:fmt(x.value),Categoria:x.category||"",Produto:x.product_name||""}))} name="caixa-caixapro" sheet="Caixa"/>
              <PBtn cols={[{k:"Data",l:"Data"},{k:"Descrição",l:"Descrição"},{k:"Tipo",l:"Tipo"},{k:"Valor",l:"Valor"},{k:"Categoria",l:"Categoria"}]} rows={fCash.map(x=>({Data:x.date,Descrição:x.description,Tipo:x.type==="entrada"?"Entrada":"Saída",Valor:fmt(x.value),Categoria:x.category||""}))} name="caixa-caixapro" title="Relatório de Caixa"/>
              <Btn sm v="ghost" onClick={()=>{setLocalTaxes(JSON.parse(JSON.stringify(cardTaxes)));setShowCardConfig(true);}}>💳 Taxas de Cartão</Btn>
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
            {fCash.length===0?<div style={{textAlign:"center",padding:"2.5rem 1rem"}}><div style={{fontSize:"2rem",marginBottom:".5rem"}}>💰</div><p style={{color:"var(--tx5)",fontSize:".8rem"}}>Nenhum lançamento no caixa.</p></div>
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
                    {canDelete&&<button onClick={()=>deleteCash(x.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}
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
              <XBtn rows={clients.map(c=>({Nome:c.name,Telefone:c.phone||"",Email:c.email||"",Obs:c.notes||"",Compras:sales.filter(s=>s.client_id===c.id||s.client_name===c.name).length,Total:fmt(batchRevenue(sales.filter(s=>s.client_id===c.id||s.client_name===c.name)))}))} name="clientes-caixapro" sheet="Clientes"/>
              <PBtn cols={[{k:"Nome",l:"Nome"},{k:"Telefone",l:"Telefone"},{k:"Email",l:"Email"},{k:"Compras",l:"Compras"},{k:"Total",l:"Total"}]} rows={clients.map(c=>({Nome:c.name,Telefone:c.phone||"",Email:c.email||"",Compras:sales.filter(s=>s.client_id===c.id||s.client_name===c.name).length,Total:fmt(batchRevenue(sales.filter(s=>s.client_id===c.id||s.client_name===c.name)))}))} name="clientes-caixapro" title="Relatório de Clientes"/>
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
                        const intv=parseInt(c.interval_days)||7;
                        const started=c.treatment_start?Math.floor((new Date()-new Date(c.treatment_start))/86400000):null;
                        const nxt=started!==null?intv-(started%intv):null;
                        const col=nxt!==null&&nxt<=2?"#f59e0b":"#10b981";
                        return <div style={{fontSize:".64rem",color:col,marginTop:".1rem"}}>💉 {c.dose} · {intv}d{nxt!==null?" · "+(nxt<=0?"🔔 Dose hoje!":nxt<=2?"⚠️ "+nxt+"d":"Próxima: "+nxt+"d"):""}</div>;
                      })()}
                      {(()=>{
                        const fin=clientFin(c.id,c.name);
                        const lim=parseFloat(c.credit_limit)||0;
                        if(fin.aberto<=0.005&&lim<=0&&fin.pagas===0)return null;
                        const over=lim>0&&fin.aberto>lim+0.005;
                        return <div style={{fontSize:".64rem",marginTop:".1rem",fontWeight:600,color:over?"#f56565":fin.vencidasAbertas>0?"#f59e0b":"#4f5ef0"}}>💳 Fiado em aberto: {fmt(fin.aberto)}{lim>0?" / limite "+fmt(lim):""}{over?" ⚠️ acima do limite":""}{fin.vencidasAbertas>0?" · "+fin.vencidasAbertas+" vencida(s)":""}{fin.pontualidade!==null?" · "+fmtPct(fin.pontualidade)+" em dia":""}</div>;
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
                      {c.phone&&<button onClick={()=>{const m="Olá "+c.name+"! Como está o tratamento? Estamos à disposição 💉";window.open("https://wa.me/55"+c.phone.replace(/[^0-9]/g,"")+"?text="+encodeURIComponent(m),"_blank");}} style={{background:"none",border:"none",color:"#25d366",padding:".2rem"}} title="WhatsApp"><span style={{fontSize:13}}>📱</span></button>}
                      {canEdit&&<><button onClick={()=>{setEditing({...c});setModal("editCliente");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>{canDelete&&<button onClick={()=>delClient(c.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>}</>}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>}
        </>)}

        {/* ══ PRODUTOS ══ */}
        {tab==="estoque"&&estTab==="produtos"&&(<>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>Produtos</h2>
            <div style={{display:"flex",gap:".4rem",flexWrap:"wrap",alignItems:"center"}}>
              <XBtn rows={products.map(p=>({Código:p.code,Nome:p.name,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Margem:fmtPct(p.margin),Estoque:p.stock_qty,Fornecedor:p.supplier_name||"—"}))} name="produtos-caixapro" sheet="Produtos"/>
              <PBtn cols={[{k:"Nome",l:"Produto"},{k:"Custo",l:"Custo"},{k:"Preço",l:"Preço"},{k:"Markup",l:"Markup"},{k:"Estoque",l:"Estoque"},{k:"Fornecedor",l:"Fornecedor"}]} rows={products.map(p=>({Nome:p.name,Custo:fmt(p.cost_per_unit),Preço:fmt(p.price_per_unit),Markup:fmtPct(p.markup),Estoque:`${p.stock_qty} ${p.unit||"un"}`,Fornecedor:p.supplier_name||"—"}))} name="produtos-caixapro" title="Relatório de Produtos"/>
              <Btn sm v="ghost" onClick={()=>setShowImportCalc(true)}>🧮 Custo Importação</Btn>
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
                      <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",color:"var(--tx)"}}>{(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name}</span>
                      {p.supplier_name&&<Badge color="#8b44f0" sm>🏭 {p.supplier_name}</Badge>}
                    </div>
                    <div style={{fontSize:".65rem",color:"var(--tx5)"}}>Cód: {p.code}{p.batch&&` · Lote: ${p.batch}`} · {p.unit||"un"}</div>
                    {p.description&&<div style={{fontSize:".68rem",color:"var(--tx4)",marginTop:".12rem"}}>{p.description}</div>}
                  </div>
                  {canEdit&&<div style={{display:"flex",gap:".35rem",flexShrink:0,marginLeft:".65rem"}}>
                    <button onClick={()=>{setEditing({...p,cost_per_unit:String(p.cost_per_unit),price_per_unit:String(p.price_per_unit),stock_qty:String(p.stock_qty),min_stock:String(p.min_stock||5),units_per_pack:String(p.units_per_pack||1)});setModal("editProd");}} style={{background:"none",border:"none",color:"#4f5ef0"}}><Ic n="edit" s={13}/></button>
                    {canDelete&&<button onClick={()=>delProduct(p.id)} style={{background:"none",border:"none",color:"var(--tx6)"}}><Ic n="trash" s={13}/></button>}
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

        {/* ══ FRETE ══ */}
        {tab==="frete"&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>🛵 Calculadora de Frete</h2>
              <Btn sm v="ghost" onClick={()=>{setLocalFrete(JSON.parse(JSON.stringify(freteConfig)));setShowFreteConfig(true);}}>⚙️ Configurar</Btn>
            </div>

            {/* Pré-visualização da fórmula */}
            <div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98110)",border:"1px solid #4f5ef030",borderRadius:".75rem",padding:".75rem 1rem",marginBottom:".85rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:".5rem"}}>
              <div style={{fontSize:".75rem",color:"var(--tx4)",display:"flex",alignItems:"center",gap:".35rem",flexWrap:"wrap"}}>
                <span style={{fontWeight:700,color:"#4f5ef0"}}>Fórmula:</span>
                <span>R$ {freteConfig.base||0} base</span>
                <span style={{color:"var(--sub)"}}>+</span>
                <span>({freteConfig.ratePerKm||0} × km)</span>
                {(freteConfig.minFee||0)>0&&<span style={{color:"#f59e0b"}}>· mín. R$ {freteConfig.minFee}</span>}
                {(freteConfig.maxFee||0)>0&&<span style={{color:"#10b981"}}>· máx. R$ {freteConfig.maxFee}</span>}
              </div>
              <span style={{fontSize:".7rem",color:"var(--tx5)"}}>Ex: 10km → {fmt((parseFloat(freteConfig.base)||0)+(10*(parseFloat(freteConfig.ratePerKm)||0)))}</span>
            </div>

            {/* Zonas de entrega (taxa fixa) */}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".85rem 1rem",marginBottom:".85rem"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".55rem",flexWrap:"wrap",gap:".4rem"}}>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)"}}>🗺️ Zonas de entrega (taxa fixa)</div>
                <span style={{fontSize:".66rem",color:"var(--tx5)"}}>Valor fixo por bairro/região — alternativa ao cálculo por km</span>
              </div>
              {(freteConfig.zonas||[]).length>0
                ?<div style={{display:"flex",gap:".4rem",flexWrap:"wrap",marginBottom:canEdit?".65rem":0}}>
                    {(freteConfig.zonas||[]).map(z=>(
                      <div key={z.id} style={{display:"flex",alignItems:"center",gap:".4rem",background:"#10b98112",border:"1px solid #10b98130",borderRadius:".5rem",padding:".3rem .6rem"}}>
                        <span style={{fontSize:".76rem",fontWeight:600,color:"var(--tx2)"}}>{z.name}</span>
                        <span style={{fontSize:".78rem",fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(z.fee)}</span>
                        {canEdit&&<button onClick={()=>delZona(z.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:0,cursor:"pointer",display:"flex"}}><Ic n="trash" s={12}/></button>}
                      </div>
                    ))}
                  </div>
                :<p style={{fontSize:".72rem",color:"var(--tx5)",marginBottom:canEdit?".65rem":0}}>Nenhuma zona cadastrada. {canEdit?"Adicione abaixo (ex: Centro · R$ 8).":""}</p>
              }
              {canEdit&&<div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
                <input value={zonaForm.name} onChange={e=>setZonaForm(f=>({...f,name:e.target.value}))} placeholder="Bairro / região" style={{...IS,flex:1,minWidth:140,fontSize:".82rem"}}/>
                <input type="number" min="0" step="0.01" value={zonaForm.fee} onChange={e=>setZonaForm(f=>({...f,fee:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addZona()} placeholder="R$" style={{...IS,width:90,fontSize:".82rem"}}/>
                <Btn sm v="ok" onClick={addZona}><Ic n="plus" s={12}/>Adicionar</Btn>
              </div>}
            </div>

            {/* Origens configuradas */}
            {freteConfig.origens.some(o=>o.lat)&&(
              <div style={{display:"flex",gap:".35rem",marginBottom:".85rem",flexWrap:"wrap"}}>
                {freteConfig.origens.filter(o=>o.name&&o.lat).map(o=>(
                  <button key={o.id} onClick={()=>setFreteOrigem(o.id)}
                    style={{padding:".35rem .75rem",borderRadius:".45rem",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"1px solid "+(freteOrigem===o.id?"#4f5ef0":"var(--bdr2)"),background:freteOrigem===o.id?"#4f5ef020":"transparent",color:freteOrigem===o.id?"#4f5ef0":"var(--navoff)"}}>
                    📍 {o.name}
                  </button>
                ))}
              </div>
            )}
            {!freteConfig.origens.some(o=>o.lat)&&(
              <div style={{background:"#f59e0b10",border:"1px solid #f59e0b30",borderRadius:".6rem",padding:".65rem .85rem",marginBottom:".85rem",fontSize:".75rem",color:"#f59e0b",display:"flex",alignItems:"center",gap:".45rem"}}>
                <Ic n="warn" s={13}/>Nenhum ponto de origem configurado. Clique em ⚙️ Configurar para cadastrar.
              </div>
            )}

            {/* Calculadora */}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
              <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".5rem",fontWeight:700}}>📍 Endereço de Destino</div>
              <div style={{display:"flex",gap:".5rem",marginBottom:".5rem",flexWrap:"wrap"}}>
                <input
                  value={freteDestino}
                  onChange={e=>setFreteDestino(e.target.value)}
                  onKeyDown={e=>e.key==="Enter"&&calcFrete()}
                  placeholder="Ex: Rua das Flores, 123, Bairro, Cidade"
                  style={{...IS,flex:1,minWidth:200,fontSize:".85rem"}}
                />
                <Btn onClick={calcFrete} disabled={freteLoading||!freteConfig.origens.some(o=>o.lat)}>
                  {freteLoading?<><span style={{display:"inline-block",animation:"spin 1s linear infinite"}}>⏳</span> Calculando...</>:<>🗺️ Calcular</>}
                </Btn>
              </div>

              {/* Múltiplos resultados de geocoding */}
              {freteGeoList.length>0&&(
                <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".5rem"}}>
                  <div style={{fontSize:".7rem",color:"#f59e0b",marginBottom:".45rem",fontWeight:600}}>⚠️ Encontramos vários endereços. Selecione o correto:</div>
                  {freteGeoList.map((g,i)=>(
                    <button key={i} onClick={async()=>{
                      setFreteGeoList([]);setFreteLoading(true);
                      const origem=freteConfig.origens.find(o=>o.id===freteOrigem);
                      if(origem&&origem.lat)await calcFreteWithCoords(origem,g);
                      else setFreteLoading(false);
                    }} style={{display:"block",width:"100%",textAlign:"left",padding:".5rem .65rem",borderRadius:".4rem",fontSize:".75rem",color:"var(--tx)",background:"transparent",border:"1px solid var(--bdr2)",marginBottom:".3rem",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",lineHeight:1.4}}>
                      📍 {g.display}
                    </button>
                  ))}
                </div>
              )}

              {/* Resultado */}
              {freteResult&&(
                <div style={{background:"linear-gradient(135deg,#10b98115,#4f5ef010)",border:"1px solid #10b98130",borderRadius:".65rem",padding:"1rem",marginTop:".65rem"}}>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".5rem",marginBottom:".75rem"}}>
                    {[{l:"Distância",v:freteResult.distKm+" km",c:"#4f5ef0"},{l:"Tempo estimado",v:freteResult.durMin+" min",c:"var(--tx)"},{l:"Taxa de entrega",v:fmt(freteResult.fee),c:"#10b981"}].map(m=>(
                      <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".5rem",padding:".6rem .5rem"}}>
                        <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div>
                        <div style={{fontSize:".95rem",fontWeight:800,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{fontSize:".67rem",color:"var(--tx5)",marginBottom:".65rem",lineHeight:1.5}}>
                    🏁 <strong>{freteResult.origName}</strong> → {freteResult.destName.slice(0,80)}{freteResult.destName.length>80?"...":""}
                  </div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",marginBottom:".65rem",background:"var(--pill)",padding:".4rem .65rem",borderRadius:".4rem"}}>
                    💡 Cálculo: R$ {freteResult.calcDetail.base} base + ({freteResult.calcDetail.distKm}km × R$ {freteResult.calcDetail.rate}/km)
                    {freteResult.calcDetail.minF>0&&` · mín. R$ ${freteResult.calcDetail.minF}`}
                    {freteResult.calcDetail.maxF>0&&` · máx. R$ ${freteResult.calcDetail.maxF}`}
                  </div>
                  <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end",flexWrap:"wrap"}}>
                    <Btn v="ghost" onClick={()=>setFreteResult(null)}>Limpar</Btn>
                    <Btn v="ok" onClick={()=>{
                      setCartFreight(String(freteResult.fee));
                      setCartDelivery(false);
                      toast$("🛵 Taxa de R$ "+fmt(freteResult.fee)+" aplicada ao carrinho!");
                      setTab("vendas");
                      setModal("sale");
                    }}>✅ Aplicar à Venda (R$ {fmt(freteResult.fee)})</Btn>
                  </div>
                </div>
              )}
            </div>

            {/* Histórico de cálculos */}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".75rem 1rem"}}>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".78rem",color:"var(--tx2)",marginBottom:".5rem"}}>📊 Últimas entregas registradas</div>
              {(()=>{
                const entregas=cashTx.filter(t=>t.category==="Frete"||t.description?.toLowerCase().includes("frete")||t.description?.toLowerCase().includes("entregador")).slice(0,5);
                if(entregas.length===0)return <p style={{color:"var(--tx5)",fontSize:".75rem",textAlign:"center",padding:"1rem"}}>Nenhuma entrega no caixa ainda.</p>;
                return entregas.map((t,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",padding:".4rem 0",borderBottom:"1px solid var(--sep)",fontSize:".75rem"}}>
                    <div><span style={{color:"var(--tx)"}}>{t.description?.slice(0,50)}</span><span style={{color:"var(--tx5)",fontSize:".65rem"}}> · {t.date}</span></div>
                    <span style={{fontWeight:700,color:"#f59e0b",fontFamily:"'Syne',sans-serif"}}>{fmt(t.value)}</span>
                  </div>
                ));
              })()}
            </div>
          </div>
        )}

        {/* ══ PEDIDOS ══ */}
        {tab==="estoque"&&estTab==="pedidos"&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>📦 Pedidos de Compra</h2>
              <div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
                <XBtn rows={orders.map(o=>({Data:o.order_date,Fornecedor:o.supplier_name,Status:o.status,Total:fmt(o.total_value),"Sinal (%)":o.initial_pct+"%","Sinal R$":fmt(o.initial_value),Restante:fmt(o.remaining_value),Recebimento:o.received_date||"—"}))} name="pedidos-caixapro" sheet="Pedidos"/>
                {isAdmin&&needsPartialFix&&<Btn sm v="ghost" onClick={fixPartialOrders} title="Reabre o saldo de pedidos antigos recebidos parcialmente">🔧 Corrigir parciais</Btn>}
                <Btn sm onClick={()=>setShowOrderModal(true)}><Ic n="plus" s={12}/>Pedido de Compra</Btn>
              </div>
            </div>
            {/* KPIs pedidos */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".6rem",marginBottom:".75rem"}}>
              <KCard label="Pendentes" value={fmtN(orders.filter(o=>o.status==="pendente"||o.status==="parcial").length)} sub={fmt((orders||[]).filter(o=>o.status==="pendente"||o.status==="parcial").reduce((a,o)=>a+o.remaining_value,0))+" restante"} color="#f59e0b"/>
              <KCard label="Recebidos" value={fmtN((orders||[]).filter(o=>o.status==="recebido").length)} color="#10b981"/>
              <KCard label="Total pago" value={fmt((orders||[]).reduce((a,o)=>a+(o.initial_value||0)+(o.remaining_paid||0),0))} color="#4f5ef0"/>
              <KCard label="A pagar" value={fmt((orders||[]).filter(o=>o.status==="pendente").reduce((a,o)=>a+o.remaining_value,0))} color="#f56565"/>
            </div>
            {/* Lista pedidos */}
            {orders.length===0
              ?<div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"3rem",textAlign:"center"}}>
                <div style={{fontSize:"2rem",marginBottom:".75rem"}}>📦</div>
                <div style={{color:"var(--tx5)",fontSize:".85rem",marginBottom:".75rem"}}>Nenhum pedido registrado.</div>
                <Btn onClick={()=>setShowOrderModal(true)}><Ic n="plus" s={13}/>Criar primeiro pedido</Btn>
               </div>
              :orders.map(order=>{
                const items=parseItems(order);
                const isDropship=items.some(i=>i.dropship);
                const statusColor=order.status==="recebido"?"#10b981":order.status==="perdido"?"#f56565":order.status==="parcial"?"#0891b2":order.status==="cancelado"?"#666a88":"#f59e0b";
                const statusLabel=order.status==="recebido"?"✅ Recebido":order.status==="cancelado"?"❌ Cancelado":order.status==="perdido"?"💀 Perdido":"🟡 Pendente";
                const pctPaid=order.total_value>0?((order.initial_value+(order.remaining_paid||0))/order.total_value)*100:0;
                return(
                  <div key={order.id} style={{background:"var(--card)",border:"1px solid "+(order.status==="pendente"?"#f59e0b30":order.status==="recebido"?"#10b98130":order.status==="parcial"?"#0891b230":"var(--bdr)"),borderRadius:".75rem",padding:".9rem 1rem",marginBottom:".65rem"}}>
                    <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:".55rem"}}>
                      <div style={{flex:1}}>
                        <div style={{display:"flex",alignItems:"center",gap:".45rem",flexWrap:"wrap",marginBottom:".2rem"}}>
                          <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".9rem",color:"var(--tx)"}}>🏭 {order.supplier_name}</span>
                          <Badge color={statusColor} sm>{statusLabel}</Badge>
                          <Badge color="#4f5ef0" sm>{order.order_date}</Badge>
                          {isDropship&&<Badge color="#8b44f0" sm>🔄 Venda direta</Badge>}
                        </div>
                        <div style={{fontSize:".73rem",color:"var(--tx4)",marginBottom:".3rem"}}>
                          {items.slice(0,3).map((i,idx)=><span key={idx}>{idx>0?" · ":""}{i.product_name} ({i.qty} {i.unit})</span>)}
                          {items.length>3&&<span> +{items.length-3} item(s)</span>}
                        </div>
                        {order.notes&&<div style={{fontSize:".68rem",color:"var(--sub)"}}>📝 {order.notes}</div>}
                      </div>
                      <div style={{display:"flex",gap:".3rem",flexShrink:0,marginLeft:".75rem"}}>
                        <button onClick={()=>exportOrderPDF(order)} title="Gerar PDF do pedido" style={{background:"#450a0a",border:"1px solid #7f1d1d",borderRadius:".4rem",padding:".28rem .5rem",color:"#f87171",display:"flex",alignItems:"center",cursor:"pointer"}}><Ic n="pdf" s={13}/></button>
                        {(order.status==="pendente"||order.status==="parcial")&&(
                          <button onClick={()=>{
                            const itsAll=parseItems(order);
                            const initChecked={};
                            itsAll.forEach((it,i)=>{if(!it.received)initChecked[i]={checked:false,qty:it.qty};});
                            setReceiveChecked(initChecked);
                            setReceivePayment("");
                            setShowReceiveModal(order);
                          }} style={{background:"#10b98115",border:"1px solid #10b98130",borderRadius:".45rem",padding:".32rem .65rem",color:"#10b981",fontSize:".73rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                            {order.status==="parcial"?"📦 Continuar Recebimento":"✅ Receber"}
                          </button>
                        )}
                        {(order.status==="pendente"||order.status==="parcial")&&isAdmin&&(
                          <button onClick={()=>markOrderLost(order)} title="Marcar como perdido" style={{background:"#1e1010",border:"1px solid #3a1515",borderRadius:".4rem",padding:".28rem .55rem",color:"#f59e0b",fontSize:".7rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>💀</button>
                        )}
                        {canEdit&&(order.status==="pendente"||order.status==="parcial")&&<button onClick={()=>setEditingOrder(order)} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>}
                        {canDelete&&<button onClick={()=>{if(window.confirm("Excluir pedido e reverter todo o fluxo? Esta ação não pode ser desfeita."))deleteOrder(order);}} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}} title="Excluir e reverter tudo"><Ic n="trash" s={13}/></button>}
                      </div>
                    </div>
                    {/* Financeiro */}
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(100px,1fr))",gap:".38rem",marginBottom:".5rem"}}>
                      {[{l:"Total pedido",v:fmt(order.total_value),c:"var(--tx)"},{l:"Sinal pago ("+order.initial_pct+"%)",v:fmt(order.initial_value),c:"#f59e0b"},{l:"Restante",v:fmt(order.remaining_value),c:"#f56565"},{l:order.status==="recebido"?"Pago recebim.":order.status==="parcial"?"Pago parcial":"A pagar",v:fmt(order.status==="recebido"||order.status==="parcial"?order.remaining_paid||0:order.remaining_value),c:order.status==="recebido"?"#10b981":order.status==="parcial"?"#0891b2":"#f56565"}].map(m=>(
                        <Pill key={m.l} label={m.l} value={m.v} color={m.c}/>
                      ))}
                    </div>
                    {/* Barras de progresso — itens + pagamento */}
                    <div style={{display:"grid",gap:".3rem"}}>
                      {/* Barra itens recebidos */}
                      {(()=>{
                        const recCount=items.filter(i=>i.received).length;
                        const pctItems=items.length>0?(recCount/items.length)*100:0;
                        return(<>
                          <div style={{height:5,background:"var(--bdr)",borderRadius:3}}>
                            <div style={{height:"100%",width:Math.min(100,pctItems)+"%",background:pctItems>=100?"linear-gradient(90deg,#10b981,#059669)":"linear-gradient(90deg,#0891b2,#0e7490)",borderRadius:3,transition:"width .5s"}}/>
                          </div>
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:".6rem",color:"var(--sub)"}}>
                            <span style={{color:"#0891b2"}}>📦 {recCount}/{items.length} produto(s) recebido(s)</span>
                            <span style={{color:pctPaid>=100?"#10b981":"var(--sub)"}}>💰 {fmtPct(pctPaid)} pago</span>
                          </div>
                        </>);
                      })()}
                    </div>
                    {/* Itens detalhados */}
                    {items.length>0&&(
                      <div style={{borderTop:"1px solid var(--sep)",marginTop:".5rem",paddingTop:".5rem",display:"grid",gap:".25rem"}}>
                        {items.map((it,idx)=>(
                          <div key={idx} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:".73rem"}}>
                            <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
                              {it.received
                                ?<span style={{fontSize:".65rem",color:"#10b981",background:"#10b98115",borderRadius:"99px",padding:".08rem .4rem",border:"1px solid #10b98130",flexShrink:0}}>✅ {it.received_qty||it.qty}</span>
                                :(parseFloat(it.received_qty)||0)>0
                                  ?<span style={{fontSize:".65rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".08rem .4rem",border:"1px solid #f59e0b30",flexShrink:0,fontWeight:700}}>📦 {it.received_qty}/{it.qty}</span>
                                  :<span style={{fontSize:".65rem",color:"#f59e0b",background:"#f59e0b15",borderRadius:"99px",padding:".08rem .4rem",border:"1px solid #f59e0b30",flexShrink:0}}>⏳ aguardando</span>
                              }
                              <span style={{color:it.received?"var(--tx5)":"var(--tx3)",textDecoration:it.received?"line-through":""}}>{it.product_name} × {it.qty} {it.unit}</span>
                            </div>
                            <span style={{color:"var(--tx4)",fontWeight:600}}>{fmt(it.total)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })
            }
          </div>
        )}

        {/* ══ A RECEBER ══ */}
        {tab==="financeiro"&&finTab==="receber"&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".75rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>💰 Contas a Receber</h2>
              <XBtn rows={receivables.map(r=>({Cliente:r.client_name||"—",Descrição:r.description,Valor:fmt(r.value),Vencimento:r.due_date||"—",Status:r.paid?"Recebido":"Pendente"}))} name="recebiveis-caixapro" sheet="A Receber"/>
            </div>
            {(()=>{
              const abertos=(receivables||[]).filter(r=>!r.paid);
              const totalAberto=abertos.reduce((a,r)=>a+(parseFloat(r.value)||0),0);
              const hoje=new Date();hoje.setHours(0,0,0,0);
              const dOver=r=>r.due_date?Math.floor((hoje-new Date(r.due_date+"T00:00:00"))/86400000):0;
              const fx=(min,max)=>abertos.filter(r=>{const d=dOver(r);return d>=min&&(max===null||d<=max);}).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
              const aVencer=abertos.filter(r=>dOver(r)<=0).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
              const b1=fx(1,30),b2=fx(31,60),b3=fx(61,90),b4=fx(91,null);
              const vencidoTot=b1+b2+b3+b4;
              const inad=totalAberto>0?(vencidoTot/totalAberto)*100:0;
              const recebidoTot=(receivables||[]).filter(r=>r.paid).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
              const buckets=[{l:"A vencer",v:aVencer,c:"#10b981"},{l:"1-30d",v:b1,c:"#f59e0b"},{l:"31-60d",v:b2,c:"#f97316"},{l:"61-90d",v:b3,c:"#ef4444"},{l:"+90d",v:b4,c:"#b91c1c"}];
              return(<>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:".6rem"}}>
                  <KCard label="Em aberto" value={fmtN(abertos.length)} sub={fmt(totalAberto)} color="#4f5ef0"/>
                  <KCard label="Vencido" value={fmt(vencidoTot)} sub={fmtN(abertos.filter(r=>dOver(r)>0).length)+" conta(s)"} color="#f56565"/>
                  <KCard label="Inadimplência" value={fmtPct(inad)} sub={inad<=15?"saudável (<15%)":inad<=25?"atenção":"crítico"} color={inad<=15?"#10b981":inad<=25?"#f59e0b":"#f56565"}/>
                </div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:".75rem .9rem",marginBottom:".75rem"}}>
                  <div style={{fontSize:".7rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".04em",marginBottom:".5rem",fontWeight:700}}>📅 Envelhecimento dos recebíveis</div>
                  {totalAberto<=0
                    ?<div style={{fontSize:".75rem",color:"var(--tx5)"}}>Nenhum valor em aberto.</div>
                    :<>
                      <div style={{display:"flex",height:10,borderRadius:"99px",overflow:"hidden",marginBottom:".55rem",background:"var(--pill)"}}>
                        {buckets.map(bk=>bk.v>0&&<div key={bk.l} style={{width:(bk.v/totalAberto*100)+"%",background:bk.c}} title={bk.l+": "+fmt(bk.v)}/>)}
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(92px,1fr))",gap:".4rem"}}>
                        {buckets.map(bk=>(
                          <div key={bk.l} style={{textAlign:"center",padding:".35rem",borderRadius:".4rem",background:"var(--pill)"}}>
                            <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:".25rem",fontSize:".62rem",color:"var(--tx5)"}}><span style={{width:7,height:7,borderRadius:"50%",background:bk.c,display:"inline-block"}}/>{bk.l}</div>
                            <div style={{fontSize:".76rem",fontWeight:700,fontFamily:"'Syne',sans-serif",color:bk.v>0?bk.c:"var(--tx6)"}}>{fmt(bk.v)}</div>
                          </div>
                        ))}
                      </div>
                    </>
                  }
                </div>
              </>);
            })()}
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem",marginBottom:".75rem"}}>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>➕ Nova conta a receber</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".5rem",marginBottom:".5rem"}}>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Cliente</div>
                  <select value={recForm.client_id} onChange={e=>{const c=clients.find(x=>x.id===e.target.value);setRecForm(f=>({...f,client_id:e.target.value,client_name:c?c.name:""}));}} style={IS}>
                    <option value="">Selecione...</option>
                    {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Vencimento</div>
                  <input type="date" value={recForm.due_date} onChange={e=>setRecForm(f=>({...f,due_date:e.target.value}))} style={IS}/>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:".5rem",marginBottom:".65rem"}}>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Descrição *</div>
                  <input value={recForm.description||""} onChange={e=>setRecForm(f=>({...f,description:e.target.value}))} placeholder="Ex: Venda parcelada, fiado..." style={IS}/>
                </div>
                <div>
                  <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".3rem"}}>Valor (R$) *</div>
                  <input type="number" min="0" step="0.01" value={recForm.value||""} onChange={e=>setRecForm(f=>({...f,value:e.target.value}))} placeholder="0,00" style={IS}/>
                </div>
              </div>
              <Btn v="ok" onClick={addReceivable}><Ic n="save" s={13}/>Registrar</Btn>
            </div>
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
              {receivables.length===0
                ?<p style={{color:"var(--tx5)",textAlign:"center",padding:"2.5rem 0",fontSize:".8rem"}}>Nenhuma conta registrada. Use o formulário acima para registrar.</p>
                :receivables.map(r=>{
                  const ov=!r.paid&&r.due_date&&new Date(r.due_date)<new Date();
                  const dv=r.due_date?Math.ceil((new Date(r.due_date+"T23:59:59")-new Date())/86400000):null;
                  return(
                    <div key={r.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)",background:ov?"#f5656508":"transparent"}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".15rem",flexWrap:"wrap"}}>
                          <span style={{fontWeight:700,fontSize:".83rem",color:r.paid?"var(--tx5)":"var(--tx)",textDecoration:r.paid?"line-through":"none"}}>{r.description}</span>
                          {r.paid&&<Badge color="#10b981" sm>✅ Recebido</Badge>}
                          {ov&&<Badge color="#f56565" sm>⚠️ Vencido</Badge>}
                          {!r.paid&&!ov&&dv!==null&&dv<=3&&<Badge color="#f59e0b" sm>{"⏰ Vence em "+dv+"d"}</Badge>}
                        </div>
                        <div style={{fontSize:".67rem",color:"var(--tx5)"}}>{r.client_name||"Sem cliente"}{r.due_date&&" · Vcto: "+new Date(r.due_date+"T00:00:00").toLocaleDateString("pt-BR")}{r.paid_date&&" · Pago: "+r.paid_date}</div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:".55rem",flexShrink:0,marginLeft:".75rem"}}>
                        <span style={{fontWeight:700,fontFamily:"'Syne',sans-serif",color:r.paid?"#10b981":ov?"#f56565":"#f59e0b",fontSize:".9rem"}}>{fmt(r.value)}</span>
                        {!r.paid&&<button onClick={()=>cobrarWhatsApp(r)} title="Enviar lembrete de cobrança no WhatsApp" style={{background:"#25D36615",border:"1px solid #25D36640",borderRadius:".4rem",padding:".28rem .55rem",color:"#1ba34e",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>📲 Cobrar</button>}
                        <button onClick={()=>exportReceivablePDF(r)} title="Gerar recibo/promissória em PDF" style={{background:"#450a0a",border:"1px solid #7f1d1d",borderRadius:".4rem",padding:".28rem .5rem",color:"#f87171",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>📄</button>
                        {!r.paid&&<button onClick={()=>payReceivable(r.id)} style={{background:"#10b98115",border:"1px solid #10b98130",borderRadius:".4rem",padding:".28rem .6rem",color:"#10b981",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>✅ Recebido</button>}
                        {canDelete&&<button onClick={()=>deleteReceivable(r.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem",cursor:"pointer"}}><Ic n="trash" s={13}/></button>}
                      </div>
                    </div>
                  );
                })
              }
            </div>
          </div>
        )}

        {tab==="config"&&isAdmin&&(
          <div style={{animation:"fadeUp .4s ease"}}>
            {/* Config header */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"1rem",flexWrap:"wrap",gap:".5rem"}}>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)"}}>⚙️ Configurações do Sistema</h2>
            </div>

            {/* Config Nav Pills */}
            <div style={{display:"flex",gap:".35rem",flexWrap:"wrap",marginBottom:"1rem",background:"var(--card)",borderRadius:".75rem",padding:".5rem"}}>
              {[
                {id:"usuarios",l:"👥 Usuários"},
                {id:"fornecedores",l:"🏭 Fornecedores"},
                {id:"cartoes",l:"💳 Taxas Cartão"},
                {id:"frete",l:"📍 Frete"},
                {id:"categorias",l:"🏷️ Categorias"},
                {id:"pagamentos",l:"💰 Pagamentos"},
                {id:"metas",l:"🎯 Metas"},
                {id:"empresa",l:"🏢 Empresa"},
              ].map(s=>(
                <button key={s.id} onClick={()=>{if(s.id==="empresa"&&!localCI)setLocalCI({...companyInfo});if(s.id==="cartoes")setLocalTaxes(JSON.parse(JSON.stringify(cardTaxes)));if(s.id==="frete")setLocalFrete(JSON.parse(JSON.stringify(freteConfig)));setConfigSection(s.id);}}
                  style={{padding:".38rem .75rem",borderRadius:".5rem",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"none",background:configSection===s.id?"linear-gradient(135deg,#4f5ef0,#8b44f0)":"transparent",color:configSection===s.id?"#fff":"var(--navoff)",transition:"all .2s",whiteSpace:"nowrap"}}>
                  {s.l}
                </button>
              ))}
            </div>

            {/*  USUÁRIOS  */}
            {configSection==="usuarios"&&(
              <div>
                <div style={{display:"flex",justifyContent:"flex-end",marginBottom:".65rem"}}>
                  <Btn sm onClick={()=>setModal("addUser")}><Ic n="plus" s={12}/>Novo Usuário</Btn>
                </div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
                  {appUsers.map(u=>(
                    <div key={u.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                      <div>
                        <div style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)",display:"flex",alignItems:"center",gap:".4rem",flexWrap:"wrap"}}>
                          {u.display_name}
                          <Badge color={u.role==="admin"?"#8b44f0":(u.role==="operator"||u.role==="operador")?"#4f5ef0":"#10b981"} sm>{u.role==="admin"?"Administrador":(u.role==="operator"||u.role==="operador")?"Operador":"Visualizador"}</Badge>
                          {!u.active&&<Badge color="#f56565" sm>inativo</Badge>}
                          {u.role!=="admin"&&u.permissions&&u.permissions.candelete&&<Badge color="#f56565" sm>🗑️ pode excluir</Badge>}
                        </div>
                        <div style={{fontSize:".68rem",color:"var(--tx5)"}}>@{u.username} · último acesso: {u.last_login||"—"}</div>
                      </div>
                      <div style={{display:"flex",gap:".3rem",alignItems:"center"}}>
                        {isAdmin&&<button onClick={()=>{setEditing({...u,permissions:u.permissions||{}});setModal("editUser");}} style={{background:"#4f5ef015",border:"1px solid #4f5ef030",borderRadius:".4rem",padding:".28rem .55rem",color:"#4f5ef0",cursor:"pointer",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,display:"flex",alignItems:"center",gap:".25rem"}}><Ic n="edit" s={12}/>Editar</button>}
                        {u.username!==cu.username&&<button onClick={()=>toggleUser(u.id,!u.active)} style={{background:"none",border:"none",color:u.active?"#f59e0b":"#10b981",padding:".2rem",cursor:"pointer",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>{u.active?"Desativar":"Ativar"}</button>}
                        {isAdmin&&u.username!==cu.username&&<button onClick={()=>deleteUser(u.id)} title="Excluir usuário" style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem",cursor:"pointer"}}><Ic n="trash" s={13}/></button>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/*  FORNECEDORES  */}
            {configSection==="fornecedores"&&(
              <div>
                <div style={{display:"flex",justifyContent:"flex-end",marginBottom:".65rem"}}>
                  <Btn sm onClick={()=>setModal("addSupp")}><Ic n="plus" s={12}/>Novo Fornecedor</Btn>
                </div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",overflow:"hidden"}}>
                  {suppliers.length===0
                    ?<p style={{textAlign:"center",color:"var(--tx5)",padding:"2rem",fontSize:".8rem"}}>Nenhum fornecedor cadastrado.</p>
                    :suppliers.map(s=>(
                    <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".72rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                      <div>
                        <div style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)"}}>{s.name}</div>
                        <div style={{fontSize:".68rem",color:"var(--tx5)"}}>{[s.phone,s.email,s.address].filter(Boolean).join(" · ")}</div>
                      </div>
                      <div style={{display:"flex",gap:".3rem"}}>
                        <button onClick={()=>{setEditing({...s});setModal("editSupp");}} style={{background:"none",border:"none",color:"#4f5ef0",padding:".2rem"}}><Ic n="edit" s={13}/></button>
                        <button onClick={()=>delSupp(s.id)} style={{background:"none",border:"none",color:"var(--tx6)",padding:".2rem"}}><Ic n="trash" s={13}/></button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/*  TAXAS CARTÃO  */}
            {configSection==="cartoes"&&(()=>{
              const lt=localTaxes;const setLt=setLocalTaxes;
              return(
                <div>
                  <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .85rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0"}}>💡 Taxas por bandeira e modalidade. Afetam o valor líquido nas vendas.</div>
                  <div style={{overflowX:"auto",marginBottom:"1rem"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:".78rem"}}>
                      <thead><tr style={{borderBottom:"2px solid var(--bdr)"}}>
                        <th style={{textAlign:"left",padding:".5rem .6rem",color:"var(--sub)",fontSize:".68rem",textTransform:"uppercase"}}>Bandeira</th>
                        {CARD_MODES.map(m=><th key={m.key} style={{textAlign:"center",padding:".5rem .4rem",color:"var(--sub)",fontSize:".68rem",textTransform:"uppercase"}}>{m.label}</th>)}
                      </tr></thead>
                      <tbody>
                        {CARD_BRANDS.map(brand=>(
                          <tr key={brand} style={{borderBottom:"1px solid var(--sep)"}}>
                            <td style={{padding:".5rem .6rem",fontWeight:700,color:"var(--tx)"}}><span style={{background:"#4f5ef020",color:"#4f5ef0",borderRadius:".3rem",padding:".15rem .5rem",fontSize:".75rem"}}>{brand}</span></td>
                            {CARD_MODES.map(mode=>(
                              <td key={mode.key} style={{padding:".4rem"}}>
                                <div style={{display:"flex",alignItems:"center",gap:".2rem"}}>
                                  <input type="number" min="0" max="20" step="0.01" value={lt[brand]?.[mode.key]||""} onChange={e=>setLt(t=>({...t,[brand]:{...t[brand],[mode.key]:e.target.value}}))} style={{...IS,textAlign:"center",width:72,fontSize:".8rem"}}/>
                                  <span style={{fontSize:".7rem",color:"var(--tx5)"}}>%</span>
                                </div>
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{background:"var(--pill)",border:"1px solid var(--bdr)",borderRadius:".55rem",padding:".7rem .85rem",marginBottom:"1rem"}}>
                    <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",marginBottom:lt.__opts?.antecip?".6rem":0}}>
                      <input type="checkbox" checked={!!lt.__opts?.antecip} onChange={e=>setLt(t=>({...t,__opts:{...(t.__opts||{}),antecip:e.target.checked}}))} style={{width:18,height:18,accentColor:"#8b44f0"}}/>
                      <span style={{fontWeight:700,fontSize:".82rem",color:lt.__opts?.antecip?"#8b44f0":"var(--tx3)"}}>⚡ Habilitar taxa de antecipação (recebimento na hora)</span>
                    </label>
                    {lt.__opts?.antecip&&<div>
                      <div style={{fontSize:".68rem",color:"var(--tx5)",marginBottom:".5rem"}}>% adicional somado às vendas no crédito (à vista e parcelado). Débito não é afetado.</div>
                      <div style={{display:"flex",gap:".5rem",flexWrap:"wrap"}}>
                        {CARD_BRANDS.map(brand=>(
                          <div key={brand} style={{display:"flex",flexDirection:"column",alignItems:"center"}}>
                            <span style={{fontSize:".62rem",color:"var(--sub)",marginBottom:".15rem"}}>{brand}</span>
                            <div style={{display:"flex",alignItems:"center",gap:".15rem"}}>
                              <input type="number" min="0" max="20" step="0.01" value={lt[brand]?.antecip||""} onChange={e=>setLt(t=>({...t,[brand]:{...t[brand],antecip:e.target.value}}))} style={{...IS,textAlign:"center",width:64,fontSize:".78rem"}}/>
                              <span style={{fontSize:".68rem",color:"var(--tx5)"}}>%</span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>}
                  </div>
                  <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
                    <Btn v="ghost" onClick={()=>setLt(JSON.parse(JSON.stringify(DEFAULT_TAXES)))}>↩ Padrões</Btn>
                    <Btn v="ok" onClick={()=>saveCardTaxes(lt)}><Ic n="save" s={13}/>Salvar Taxas</Btn>
                  </div>
                </div>
              );
            })()}

            {/*  FRETE  */}
            {configSection==="frete"&&(()=>{
              const lf=localFrete||freteConfig;const setLf=setLocalFrete;
              return(
                <div>
                  <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".85rem",color:"var(--tx)",marginBottom:".75rem"}}>📍 Pontos de Origem</div>
                  {lf.origens.map((origem,oi)=>(
                    <div key={origem.id} style={{background:"var(--pill)",borderRadius:".6rem",padding:".75rem .85rem",marginBottom:".6rem"}}>
                      <div style={{display:"flex",gap:".4rem",marginBottom:".45rem",alignItems:"center"}}>
                        <span style={{fontWeight:700,fontSize:".75rem",color:"#4f5ef0",minWidth:55}}>Ponto {oi+1}</span>
                        <input value={origem.name} onChange={e=>setLf(f=>{const o=[...f.origens];o[oi]={...o[oi],name:e.target.value};return{...f,origens:o};})} placeholder="Nome do ponto" style={{...IS,flex:1,fontSize:".8rem"}}/>
                        {origem.lat&&<span style={{fontSize:".65rem",color:"#10b981",background:"#10b98115",borderRadius:"99px",padding:".1rem .45rem",border:"1px solid #10b98130",whiteSpace:"nowrap"}}>✅ ok</span>}
                      </div>
                      <div style={{display:"flex",gap:".4rem",flexWrap:"wrap"}}>
                        <input value={origem.address} onChange={e=>setLf(f=>{const o=[...f.origens];o[oi]={...o[oi],address:e.target.value,lat:null,lon:null};return{...f,origens:o};})} placeholder="Endereço completo, Cidade" style={{...IS,flex:1,fontSize:".78rem",minWidth:180}}/>
                        <button onClick={async()=>{
                          if(!origem.address.trim()){toast$("Informe o endereço.","#f56565");return;}
                          toast$("🔍 Localizando...","#4f5ef0");
                          try{
                            const results=await geocodeAddr(origem.address);
                            if(!results.length){toast$("Não encontrado. Inclua a cidade.","#f56565");return;}
                            const first=results[0];
                            setLf(f=>{const o=[...f.origens];o[oi]={...o[oi],lat:first.lat,lon:first.lon,address:first.display.split(",").slice(0,3).join(",").trim()};return{...f,origens:o};});
                            toast$("✅ "+origem.name+" localizado!");
                          }catch(e){toast$("Erro: "+e.message,"#f56565");}
                        }} style={{background:"#4f5ef020",border:"1px solid #4f5ef040",borderRadius:".4rem",padding:".35rem .7rem",color:"#4f5ef0",fontSize:".74rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>🔍 Localizar</button>
                      </div>
                    </div>
                  ))}
                  <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".85rem",color:"var(--tx)",marginTop:"1rem",marginBottom:".65rem"}}>💰 Valores de Cobrança</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:".65rem",marginBottom:"1rem"}}>
                    {[{l:"Base (R$)",k:"base",h:"Cobrado sempre"},{l:"Por km (R$/km)",k:"ratePerKm",h:"× distância"},{l:"Mínimo (R$)",k:"minFee",h:"0 = sem mínimo"},{l:"Máximo (R$)",k:"maxFee",h:"0 = sem teto"}].map(f=>(
                      <div key={f.k}>
                        <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>{f.l}</div>
                        <input type="number" min="0" step="0.5" value={lf[f.k]||""} onChange={e=>setLf(v=>({...v,[f.k]:e.target.value}))} placeholder="0" style={IS}/>
                        <div style={{fontSize:".6rem",color:"var(--tx6)",marginTop:".15rem"}}>{f.h}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{background:"var(--pill)",borderRadius:".45rem",padding:".5rem .85rem",marginBottom:"1rem",display:"flex",justifyContent:"space-between",fontSize:".75rem"}}>
                    <span style={{color:"var(--tx4)"}}>Prévia 10km:</span>
                    <span style={{fontWeight:700,color:"#10b981"}}>{fmt(Math.max(parseFloat(lf.minFee)||0,Math.min(parseFloat(lf.maxFee)||999999,(parseFloat(lf.base)||0)+(10*(parseFloat(lf.ratePerKm)||0)))))}</span>
                  </div>
                  <div style={{display:"flex",justifyContent:"flex-end"}}>
                    <Btn v="ok" onClick={()=>saveFreteConfig(lf)}><Ic n="save" s={13}/>Salvar Configurações de Frete</Btn>
                  </div>
                </div>
              );
            })()}

            {/*  CATEGORIAS  */}
            {configSection==="categorias"&&(
              <div>
                <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .85rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0"}}>🏷️ Categorias aparecem no cadastro de produtos e nos filtros de estoque.</div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",marginBottom:".75rem",overflow:"hidden"}}>
                  {activeCats.map((cat,i)=>(
                    <div key={cat.key} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:".65rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                      <div style={{display:"flex",alignItems:"center",gap:".6rem"}}>
                        <span style={{fontSize:"1.1rem"}}>{cat.icon}</span>
                        <div>
                          <div style={{fontWeight:700,fontSize:".83rem",color:"var(--tx)"}}>{cat.label}</div>
                          <div style={{fontSize:".65rem",color:"var(--tx5)"}}>chave: {cat.key}</div>
                        </div>
                      </div>
                      {i>=5&&<button onClick={()=>saveCategories(activeCats.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:"var(--tx6)",cursor:"pointer",padding:".2rem"}}><Ic n="trash" s={13}/></button>}
                      {i<5&&<span style={{fontSize:".65rem",color:"var(--tx6)"}}>padrão</span>}
                    </div>
                  ))}
                </div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem"}}>
                  <div style={{fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>➕ Nova Categoria</div>
                  <div style={{display:"grid",gridTemplateColumns:"80px 1fr",gap:".5rem",marginBottom:".5rem"}}>
                    <div>
                      <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Ícone (emoji)</div>
                      <input value={newCatIcon} onChange={e=>setNewCatIcon(e.target.value)} placeholder="📦" style={{...IS,fontSize:"1.2rem",textAlign:"center"}} maxLength={2}/>
                    </div>
                    <div>
                      <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Nome da categoria</div>
                      <input value={newCatName} onChange={e=>setNewCatName(e.target.value)} placeholder="Ex: Vitaminas, Canetas..." style={{...IS,fontSize:".85rem"}}/>
                    </div>
                  </div>
                  <Btn v="ok" onClick={()=>{
                    if(!newCatName.trim()){toast$("Informe o nome.","#f56565");return;}
                    const key=newCatName.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"");
                    if(activeCats.find(c=>c.key===key)){toast$("Categoria já existe.","#f56565");return;}
                    saveCategories([...activeCats,{key,label:newCatName.trim(),icon:newCatIcon||"📦"}]);
                    setNewCatName("");setNewCatIcon("📦");
                  }}><Ic n="plus" s={13}/>Adicionar Categoria</Btn>
                </div>
              </div>
            )}

            {/*  FORMAS DE PAGAMENTO  */}
            {configSection==="pagamentos"&&(
              <div>
                <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .85rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0"}}>💰 Configure a <b>liquidação</b> de cada forma: ⚡ instantânea (cai no caixa na hora), 📅 após X dias úteis (ex.: débito), ou 📝 promissória/fiado (vai para <b>A Receber</b> com vencimento definido na venda).</div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",marginBottom:".75rem",overflow:"hidden"}}>
                  {activePays.map((pay)=>{
                    const isParc=pay==="Crédito Parcelado";
                    const isFixed=["Débito","Crédito à Vista","Crédito Parcelado"].includes(pay);
                    const st=getSettle(pay);
                    const selSt={background:"var(--inp)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".25rem .4rem",color:"var(--tx)",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",cursor:"pointer"};
                    return(
                    <div key={pay} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:".5rem",padding:".6rem 1rem",borderBottom:"1px solid var(--sep)",flexWrap:"wrap"}}>
                      <span style={{fontWeight:600,fontSize:".83rem",color:"var(--tx)",minWidth:110}}>{pay}</span>
                      <div style={{display:"flex",alignItems:"center",gap:".4rem",flexWrap:"wrap"}}>
                        {isParc
                          ?<span style={{fontSize:".68rem",color:"#4f5ef0",background:"#4f5ef015",borderRadius:"99px",padding:".15rem .55rem",border:"1px solid #4f5ef030"}}>Parcelado mensal</span>
                          :<>
                            <select value={st.type} onChange={e=>setMethodSettle(pay,{type:e.target.value})} style={selSt}>
                              <option value="instant">⚡ Instantânea</option>
                              <option value="days">📅 Após dias úteis</option>
                              <option value="duedate">📝 Promissória</option>
                            </select>
                            {st.type==="days"&&<select value={st.days||1} onChange={e=>setMethodSettle(pay,{days:parseInt(e.target.value)})} style={selSt}>{[1,2,3,5,7,10,15,30].map(d=><option key={d} value={d}>{d} dia{d>1?"s":""} úteis</option>)}</select>}
                          </>
                        }
                        {isFixed
                          ?<span style={{fontSize:".65rem",color:"var(--tx6)"}}>fixo</span>
                          :<button onClick={()=>savePayments(activePays.filter(p=>p!==pay))} style={{background:"none",border:"none",color:"var(--tx6)",cursor:"pointer",padding:".2rem"}}><Ic n="trash" s={13}/></button>
                        }
                      </div>
                    </div>
                    );
                  })}
                </div>
                <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".75rem",padding:"1rem"}}>
                  <div style={{fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".65rem"}}>➕ Nova Forma de Pagamento</div>
                  <div style={{display:"flex",gap:".5rem"}}>
                    <input value={newPayName} onChange={e=>setNewPayName(e.target.value)} placeholder="Ex: Cheque, Boleto, Crediário..." style={{...IS,flex:1,fontSize:".85rem"}}/>
                    <Btn v="ok" onClick={()=>{
                      if(!newPayName.trim()){toast$("Informe o nome.","#f56565");return;}
                      if(activePays.includes(newPayName.trim())){toast$("Já existe.","#f56565");return;}
                      savePayments([...activePays,newPayName.trim()]);
                      setNewPayName("");
                    }}><Ic n="plus" s={13}/>Adicionar</Btn>
                  </div>
                </div>
              </div>
            )}

            {/*  METAS  */}
            {configSection==="metas"&&(()=>{
              const MONTHS=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
              const gYear=goalYear;const setGYear=setGoalYear;
              const rows=MONTHS.map((mName,mi)=>{
                const key=gYear+"-"+(mi+1).toString().padStart(2,"0");
                const goal=monthlyGoals[key]||0;
                const mStart=new Date(gYear,mi,1);const mEnd=new Date(gYear,mi+1,0,23,59,59);
                const real=batchRevenue(sales.filter(s=>{const d=new Date(s.created_at);return d>=mStart&&d<=mEnd;}));
                const pct=goal>0?(real/goal)*100:0;
                const isCurrent=new Date().getFullYear()===gYear&&new Date().getMonth()===mi;
                const isPast=new Date(gYear,mi+1,1)<new Date();
                return{key,mName,mi,goal,real,pct,isCurrent,isPast};
              });
              const totalGoal=rows.reduce((a,r)=>a+r.goal,0);
              const totalReal=rows.reduce((a,r)=>a+r.real,0);
              return(
                <div>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"1rem"}}>
                    <div style={{display:"flex",alignItems:"center",gap:".5rem"}}>
                      <button onClick={()=>setGYear(y=>y-1)} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".3rem .65rem",color:"var(--tx)",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>‹</button>
                      <span style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1rem",color:"var(--tx)",minWidth:50,textAlign:"center"}}>{gYear}</span>
                      <button onClick={()=>setGYear(y=>y+1)} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".3rem .65rem",color:"var(--tx)",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>›</button>
                    </div>
                    <div style={{display:"flex",gap:".75rem",fontSize:".75rem"}}>
                      <div style={{textAlign:"right"}}><div style={{color:"var(--tx5)",fontSize:".62rem"}}>Meta anual</div><div style={{fontWeight:700,color:"#4f5ef0",fontFamily:"'Syne',sans-serif"}}>{fmt(totalGoal)}</div></div>
                      <div style={{textAlign:"right"}}><div style={{color:"var(--tx5)",fontSize:".62rem"}}>Realizado</div><div style={{fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(totalReal)}</div></div>
                    </div>
                  </div>
                  <div style={{display:"grid",gap:".5rem"}}>
                    {rows.map(r=>(
                      <div key={r.key} style={{background:r.isCurrent?"linear-gradient(135deg,#4f5ef010,#10b98108)":"var(--card)",border:"1px solid "+(r.isCurrent?"#4f5ef040":"var(--bdr)"),borderRadius:".6rem",padding:".65rem .85rem"}}>
                        <div style={{display:"flex",alignItems:"center",gap:".5rem",marginBottom:".3rem"}}>
                          <span style={{fontWeight:700,fontSize:".82rem",color:r.isCurrent?"#4f5ef0":"var(--tx)",minWidth:90}}>{r.mName}</span>
                          {r.isCurrent&&<span style={{fontSize:".6rem",background:"#4f5ef020",color:"#4f5ef0",borderRadius:"99px",padding:".1rem .4rem",fontWeight:700}}>atual</span>}
                          <div style={{flex:1}}/>
                          <span style={{fontSize:".75rem",color:r.real>=r.goal&&r.goal>0?"#10b981":"var(--tx3)",fontWeight:600}}>{fmt(r.real)}</span>
                          <span style={{fontSize:".65rem",color:"var(--tx5)"}}>de</span>
                          <input type="number" min="0" step="100" onFocus={e=>e.target.select()} value={r.goal||""} onChange={e=>saveMonthlyGoal(r.key,e.target.value)} placeholder="R$ meta..." style={{width:100,background:"var(--inp)",border:"1px solid var(--bdr2)",borderRadius:".35rem",padding:".3rem .5rem",color:"#4f5ef0",fontWeight:700,fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",outline:"none",textAlign:"right"}}/>
                          <span style={{fontSize:".72rem",fontWeight:700,color:r.pct>=100?"#10b981":r.pct>=70?"#f59e0b":r.isPast&&r.goal>0?"#f56565":"var(--tx5)",minWidth:42,textAlign:"right"}}>{r.goal>0?fmtPct(r.pct):"—"}{r.pct>100?" 🚀":r.pct>=100?" 🏆":""}</span>
                        </div>
                        <div style={{height:5,background:"var(--bdr)",borderRadius:3}}>
                          <div style={{height:"100%",width:Math.min(100,r.pct)+"%",background:r.pct>=100?"linear-gradient(90deg,#10b981,#059669)":r.pct>=70?"linear-gradient(90deg,#f59e0b,#d97706)":"linear-gradient(90deg,#4f5ef0,#8b44f0)",borderRadius:3,transition:"width .5s"}}/>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}


            {/*  EMPRESA  */}
            {configSection==="empresa"&&(()=>{
              const ci=localCI||companyInfo;const setCi=setLocalCI;
              return(
                <div>
                  <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .85rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0"}}>🏢 Informações aparecem nos PDFs gerados pelo sistema.</div>
                  <div style={{display:"grid",gap:".65rem",marginBottom:"1rem"}}>
                    <Inp label="Nome da Empresa *" value={ci.name} onChange={e=>setCi(v=>({...v,name:e.target.value}))} placeholder="Ex: CaixaPro · Tirzepatida"/>
                    <R2>
                      <Inp label="Telefone / WhatsApp" value={ci.phone||""} onChange={e=>setCi(v=>({...v,phone:e.target.value}))} placeholder="(xx) xxxxx-xxxx"/>
                      <Inp label="CNPJ / CPF" value={ci.cnpj||""} onChange={e=>setCi(v=>({...v,cnpj:e.target.value}))} placeholder="xx.xxx.xxx/xxxx-xx"/>
                    </R2>
                    <Inp label="Endereço" value={ci.address||""} onChange={e=>setCi(v=>({...v,address:e.target.value}))} placeholder="Rua, Nº, Bairro, Cidade - UF"/>
                    <Inp label="Observação no rodapé dos PDFs" hint="opcional" value={ci.obs||""} onChange={e=>setCi(v=>({...v,obs:e.target.value}))} placeholder="Ex: Atendemos apenas com agendamento..."/>
                  <div style={{display:"flex",alignItems:"center",gap:".6rem",padding:".65rem .85rem",background:"var(--pill)",borderRadius:".5rem"}}>
                    <button onClick={()=>setCi(v=>({...v,showInPDF:!v.showInPDF}))} style={{width:38,height:22,borderRadius:11,border:"none",background:ci.showInPDF!==false?"linear-gradient(135deg,#4f5ef0,#8b44f0)":"var(--bdr2)",transition:"background .2s",cursor:"pointer",position:"relative"}}>
                      <span style={{position:"absolute",top:3,left:ci.showInPDF!==false?18:3,width:16,height:16,borderRadius:"50%",background:"#fff",transition:"left .2s",display:"block"}}/>
                    </button>
                    <div>
                      <div style={{fontSize:".78rem",fontWeight:600,color:"var(--tx)"}}>Exibir dados da empresa nos PDFs</div>
                      <div style={{fontSize:".65rem",color:"var(--tx5)"}}>{ci.showInPDF!==false?"Cabeçalho com nome/telefone/endereço nos documentos":"PDFs sem identificação da empresa"}</div>
                    </div>
                  </div>
                  </div>
                  {ci.name&&(
                    <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".75rem",marginBottom:"1rem",fontSize:".75rem"}}>
                      <div style={{fontWeight:700,color:"var(--tx2)",marginBottom:".35rem",fontSize:".7rem"}}>Prévia do cabeçalho nos PDFs:</div>
                      <div style={{fontWeight:700,color:"var(--tx)"}}>{ci.name}</div>
                      {ci.phone&&<div style={{color:"var(--tx4)"}}>{ci.phone}</div>}
                      {ci.address&&<div style={{color:"var(--tx4)"}}>{ci.address}</div>}
                      {ci.cnpj&&<div style={{color:"var(--tx5)"}}>CNPJ: {ci.cnpj}</div>}
                    </div>
                  )}
                  <div style={{display:"flex",justifyContent:"flex-end"}}>
                    <Btn v="ok" onClick={()=>{saveCompanyInfo(ci);setLocalCI(null);}}><Ic n="save" s={13}/>Salvar Informações da Empresa</Btn>
                  </div>
                  <div style={{borderTop:"1px solid var(--bdr)",marginTop:"1.1rem",paddingTop:"1rem"}}>
                    <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx2)",marginBottom:".35rem"}}>💾 Backup dos dados</div>
                    <div style={{fontSize:".7rem",color:"var(--tx5)",marginBottom:".7rem"}}>Baixe uma cópia completa de tudo — vendas, estoque, pedidos, contas a pagar/receber, clientes, fornecedores, caixa e configurações — num único arquivo. Guarde periodicamente em local seguro (nuvem/pendrive).</div>
                    <Btn v="ghost" onClick={exportBackup}><Ic n="save" s={13}/>Baixar backup completo (.json)</Btn>
                  </div>
                </div>
              );
            })()}
          </div>
        )}


        {/* ══ A RECEBER ══ */}

      </div>
    </div>

    {/* ══ ANALYTICS ══ */}
    {showA&&<Analytics onClose={()=>setShowA(false)} sales={sales} cashTx={cashTx} products={products} clients={clients} dark={dark} receivables={receivables} orders={orders} payables={payables}/>}

    {/* ══ MODAIS ══ */}

    {/* Registrar pagamento de conta a pagar (com desconto/antecipação) */}
    {payTarget&&(()=>{
      const total=parseFloat(payTarget.value)||0;
      const already=parseFloat(payTarget.paid_amount)||0;
      const falta=Math.max(0,total-already);
      const pago=parseFloat(payAmount);
      const eff=isNaN(pago)?0:Math.min(pago,falta>0?falta:pago);
      const restante=Math.max(0,falta-eff);
      const liquida=payQuitar||(eff>0&&restante<=0.005);
      const eco=payQuitar?Math.max(0,falta-eff):0;
      const catIcon=payTarget._order?(payTarget._dropship?"🔄":"📦"):({aluguel:"🏠",emprestimo:"🏦",compra:"📦",fornecedor:"🚚",imposto:"🧾",salario:"👤",servico:"🔧",outro:"📋"}[payTarget.category]||"📋");
      return(
        <Modal title="Registrar Pagamento" onClose={()=>{setPayTarget(null);setPayAmount("");setPayQuitar(false);}} icon="dollar">
          <div style={{background:"var(--sumbox)",borderRadius:".6rem",padding:".8rem .9rem",marginBottom:".9rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".2rem"}}>
              <span style={{fontSize:"1rem"}}>{catIcon}</span>
              <span style={{fontWeight:700,fontSize:".88rem",color:"var(--tx)"}}>{payTarget.description}</span>
            </div>
            <div style={{fontSize:".68rem",color:"var(--tx5)"}}>{payTarget.supplier_name||"—"}{payTarget.due_date&&" · Vencimento "+new Date(payTarget.due_date+"T00:00:00").toLocaleDateString("pt-BR")}</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:".4rem",marginTop:".6rem",paddingTop:".55rem",borderTop:"1px solid var(--sep)"}}>
              {[{l:"Total",v:fmt(total),c:"var(--tx2)"},{l:"Já pago",v:fmt(already),c:already>0.005?"#10b981":"var(--tx5)"},{l:"Falta",v:fmt(falta),c:"#f59e0b"}].map(m=>(
                <div key={m.l} style={{textAlign:"center"}}>
                  <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".04em"}}>{m.l}</div>
                  <div style={{fontSize:".88rem",fontWeight:700,fontFamily:"'Syne',sans-serif",color:m.c}}>{m.v}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={{marginBottom:".7rem"}}>
            <div style={{fontSize:".7rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".35rem",fontWeight:700}}>💸 Valor a pagar agora</div>
            <input type="number" min="0" step="0.01" autoFocus value={payAmount} onChange={e=>setPayAmount(e.target.value)} onKeyDown={e=>e.key==="Enter"&&confirmPayPayable()} placeholder="0,00" style={{...IS,fontSize:"1.05rem",fontWeight:700,padding:".7rem .85rem"}}/>
            <div style={{display:"flex",gap:".4rem",marginTop:".5rem",flexWrap:"wrap"}}>
              <button onClick={()=>setPayAmount(falta.toFixed(2))} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".25rem .6rem",fontSize:".7rem",color:"var(--tx4)",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>Pagar tudo ({fmt(falta)})</button>
              {[25,50].map(pc=><button key={pc} onClick={()=>setPayAmount((falta*pc/100).toFixed(2))} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".25rem .6rem",fontSize:".7rem",color:"var(--tx4)",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>{pc}%</button>)}
              {!payTarget._order&&[5,10,15].map(d=><button key={d} onClick={()=>setPayAmount((falta*(1-d/100)).toFixed(2))} style={{background:"#10b98112",border:"1px solid #10b98130",borderRadius:".4rem",padding:".25rem .6rem",fontSize:".7rem",color:"#10b981",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>-{d}%</button>)}
            </div>
          </div>
          {eff>0&&(
            <div style={{background:liquida?"#10b98112":"#4f5ef012",border:`1px solid ${liquida?"#10b98130":"#4f5ef030"}`,borderRadius:".6rem",padding:".6rem .85rem",marginBottom:".7rem",fontSize:".75rem",fontWeight:600,color:liquida?"#10b981":"#4f5ef0"}}>
              {liquida
                ?(eco>0.005?"✅ Liquida a conta · economia de "+fmt(eco):"✅ Este pagamento liquida a conta")
                :"➗ Pagamento parcial · após este, faltará "+fmt(restante)}
            </div>
          )}
          {!payTarget._order&&falta>0&&eff>0&&eff<falta-0.005&&(
            <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",marginBottom:".8rem",background:payQuitar?"#10b98112":"transparent",border:"1px solid "+(payQuitar?"#10b98130":"var(--bdr2)"),borderRadius:".45rem",padding:".5rem .7rem"}}>
              <input type="checkbox" checked={payQuitar} onChange={e=>setPayQuitar(e.target.checked)} style={{width:15,height:15,accentColor:"#10b981"}}/>
              <span style={{fontSize:".74rem",fontWeight:600,color:payQuitar?"#10b981":"var(--tx4)"}}>✓ Quitar a conta com este valor (abate o saldo restante como desconto)</span>
            </label>
          )}
          <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
            <Btn v="ghost" onClick={()=>{setPayTarget(null);setPayAmount("");setPayQuitar(false);}}>Cancelar</Btn>
            <Btn v="ok" onClick={confirmPayPayable}><Ic n="save" s={13}/>{liquida?"Liquidar conta":"Registrar pagamento"}</Btn>
          </div>
        </Modal>
      );
    })()}

    {/* Nova venda — CARRINHO */}
    {modal==="sale"&&(
      <Modal title="Nova Venda" onClose={()=>{setModal(null);cartReset();}} icon="sales" wide>

        {/* Info banner */}
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:"1rem",fontSize:".72rem",color:"#4f5ef0",display:"flex",gap:".35rem",alignItems:"center"}}><Ic n="info" s={12}/>Múltiplos produtos · Preço 0 = cortesia (registra custo) · Baixa estoque automática</div>

        {/*  ITENS DO CARRINHO  */}
        <div style={{marginBottom:"1rem"}}>
          <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".5rem",fontWeight:700}}>
            🛒 Itens da Venda
          </div>

          {/* Cabeçalho colunas */}
          <div style={{overflowX:"auto",WebkitOverflowScrolling:"touch"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 80px 110px 80px",gap:".4rem",marginBottom:".35rem",padding:"0 .1rem",minWidth:340}}>
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
                    {(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name} {p.stock_qty<=0?"(zerado)":`(${p.stock_qty} ${p.unit||"un"})`}
                  </option>
                ))}
              </select>

              {/* Quantidade */}
              <input
                type="number" min="1"
                onFocus={e=>e.target.select()} value={item.quantity}
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

          </div>{/* fecha scroll wrapper overflowX */}

          {/* Botão adicionar produto */}
          <button
            onClick={cartAddLine}
            style={{display:"flex",alignItems:"center",gap:".35rem",padding:".4rem .85rem",borderRadius:".45rem",border:"1px dashed #4f5ef060",background:"#4f5ef008",color:"#4f5ef0",fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,width:"100%",justifyContent:"center",marginTop:".25rem"}}
          >
            <Ic n="plus" s={13}/> Adicionar produto
          </button>
        </div>

        {/*  ENTREGA  */}
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
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".65rem"}}>
              {/* Taxa cobrada do cliente */}
              <div>
                <div style={{fontSize:".65rem",color:"#10b981",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".3rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <span style={{display:"flex",alignItems:"center",gap:".3rem"}}>📈 Taxa cobrada do cliente</span>
                  <button onClick={()=>{
                    setFreteCalcDestino(cartClient.name?cartClient.name+" ":"");
                    setFreteCalcOrigem(freteConfig.origens.find(o=>o.lat)?.id||"1");
                    setFreteCalcResult(null);setFreteCalcGeoList([]);
                    setShowFreteCalcInCart(true);
                  }} style={{background:"#10b98120",border:"1px solid #10b98140",borderRadius:".35rem",padding:".15rem .5rem",color:"#10b981",fontSize:".65rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:".25rem"}}>
                    🗺️ Calcular
                  </button>
                </div>
                <input
                  type="number" min="0" step="0.01"
                  onFocus={e=>e.target.select()} value={cartFreight}
                  onChange={e=>setCartFreight(e.target.value)}
                  placeholder="R$ 0,00"
                  style={{...IS,fontSize:".83rem",borderColor:cartFreightVal>0?"#10b98150":"var(--bdr2)"}}
                />
                {cartFreightVal>0&&<div style={{fontSize:".65rem",color:"#10b981",marginTop:".2rem",display:"flex",alignItems:"center",gap:".3rem"}}>+ {fmt(cartFreightVal)} na receita
                  <button onClick={()=>setCartFreight("")} style={{background:"none",border:"none",color:"var(--tx6)",cursor:"pointer",fontSize:".6rem",padding:0}}>✕</button>
                </div>}
              </div>

              {/* Custo do entregador */}
              <div>
                <div style={{fontSize:".65rem",color:"#f56565",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".3rem",fontWeight:700,display:"flex",alignItems:"center",gap:".3rem"}}>
                  📉 Custo pago ao entregador
                  <span style={{fontSize:".6rem",color:"var(--tx6)",fontWeight:400,textTransform:"none"}}>(manual)</span>
                </div>
                <input
                  type="number" min="0" step="0.01"
                  onFocus={e=>e.target.select()} value={cartDeliveryCost}
                  onChange={e=>setCartDeliveryCost(e.target.value)}
                  placeholder="R$ 0,00"
                  style={{...IS,fontSize:".83rem",borderColor:cartDeliveryCostVal>0?"#f5656550":"var(--bdr2)"}}
                />
                {cartDeliveryCostVal>0&&(
                  <div style={{fontSize:".65rem",color:"#f56565",marginTop:".2rem",display:"flex",justifyContent:"space-between"}}>
                    <span>- {fmt(cartDeliveryCostVal)} no caixa</span>
                    {cartFreightVal>0&&cartDeliveryCostVal>0&&<span style={{color:cartFreightVal>cartDeliveryCostVal?"#10b981":"#f59e0b",fontWeight:600}}>
                      margem entrega: {fmt(cartFreightVal-cartDeliveryCostVal)}
                    </span>}
                  </div>
                )}
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

        {cartSubtotal>0&&(
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:"1rem"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:".45rem"}}>
              <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",fontWeight:700}}>🏷️ Desconto</div>
              <div style={{display:"flex",gap:".25rem"}}>
                {["fixed","percent"].map(dt=>(
                  <button key={dt} onClick={()=>setCartDiscountType(dt)} style={{padding:".2rem .55rem",borderRadius:".3rem",fontSize:".68rem",fontWeight:600,fontFamily:"'DM Sans',sans-serif",border:"1px solid "+(cartDiscountType===dt?"#4f5ef0":"var(--bdr2)"),background:cartDiscountType===dt?"#4f5ef020":"transparent",color:cartDiscountType===dt?"#4f5ef0":"var(--navoff)"}}>
                    {dt==="fixed"?"R$":"%"}
                  </button>
                ))}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".5rem",alignItems:"center"}}>
              <input type="number" min="0" step="0.01" value={cartDiscount} onChange={e=>setCartDiscount(e.target.value)} placeholder={cartDiscountType==="percent"?"ex: 10 (%)":"ex: 20,00 (R$)"} style={IS}/>
              <div style={{fontSize:".75rem",color:cartDiscountVal>0?"#4f5ef0":"var(--tx5)",fontWeight:cartDiscountVal>0?700:400}}>
                {cartDiscountVal>0?"- "+fmt(cartDiscountVal)+" no total":"Sem desconto"}
              </div>
            </div>
          </div>
        )}

        {/*  CLIENTE + PAGAMENTO  */}
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
          <div>
            <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".3rem"}}>Forma de Pagamento</div>
            <div style={{marginTop:".6rem",padding:".6rem .75rem",background:cartIsDirect?"#f59e0b15":"var(--pill)",border:`1px solid ${cartIsDirect?"#f59e0b40":"var(--bdr)"}`,borderRadius:".5rem"}}>
              <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",fontSize:".8rem",fontWeight:600,color:cartIsDirect?"#f59e0b":"var(--tx3)"}}>
                <input type="checkbox" checked={cartIsDirect} onChange={e=>setCartIsDirect(e.target.checked)} style={{width:16,height:16,accentColor:"#f59e0b"}}/>
                🚀 Venda Direta / Encomenda (sem baixa de estoque)
              </label>
              {cartIsDirect&&(
                <div style={{marginTop:".5rem"}}>
                  <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Fornecedor que enviará ao cliente</div>
                  <input value={cartSupplierName} onChange={e=>setCartSupplierName(e.target.value)} placeholder="Nome do fornecedor" list="supp-list" style={IS}/>
                  <datalist id="supp-list">{suppliers.map(s=><option key={s.id} value={s.name}/>)}</datalist>
                  <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",marginTop:".55rem",background:cartAutoOrder?"#4f5ef012":"transparent",border:"1px solid "+(cartAutoOrder?"#4f5ef040":"var(--bdr2)"),borderRadius:".45rem",padding:".45rem .6rem"}}>
                    <input type="checkbox" checked={cartAutoOrder} onChange={e=>setCartAutoOrder(e.target.checked)} style={{width:15,height:15,accentColor:"#4f5ef0"}}/>
                    <span style={{fontSize:".74rem",fontWeight:600,color:cartAutoOrder?"#4f5ef0":"var(--tx4)"}}>🔄 Gerar pedido de compra automático para o fornecedor</span>
                  </label>
                  <div style={{fontSize:".68rem",color:"#f59e0b",marginTop:".4rem"}}>💡 O lucro entra no caixa, mas o estoque não é alterado.{cartAutoOrder&&cartSupplierName.trim()?" Um pedido (a custo) será gerado para "+cartSupplierName.trim()+" em Pedidos, sem entrada no estoque ao receber.":""}</div>
                </div>
              )}
            </div>
            {!cartSplitEnabled&&(
            <select value={cartPayment} onChange={e=>setCartPayment(e.target.value)} style={IS}>
              <optgroup label="Sem taxa">
                {activeSimplePays.map(m=><option key={m} value={m}>{m}</option>)}
              </optgroup>
              <optgroup label="Cartão">
                <option value="Débito">Débito</option>
                <option value="Crédito à Vista">Crédito à Vista</option>
                <option value="Crédito Parcelado">Crédito Parcelado</option>
              </optgroup>
            </select>
            )}
            {cartSplitEnabled&&(
              <div style={{marginTop:".6rem",padding:".55rem .7rem",background:"#8b44f015",border:"1px solid #8b44f040",borderRadius:".5rem",fontSize:".74rem",color:"#8b44f0",fontWeight:600,display:"flex",alignItems:"center",gap:".35rem"}}>
                💳 Pagamento dividido ativo — defina as formas abaixo
              </div>
            )}
          </div>
        </R2>
        {!cartSplitEnabled&&cartIsCard&&(
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".5rem"}}>
            <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".45rem",fontWeight:700}}>💳 Bandeira</div>
            <div style={{display:"flex",gap:".35rem",flexWrap:"wrap",marginBottom:cartTaxRate>0?".55rem":0}}>
              {CARD_BRANDS.map(b=>(
                <button key={b} onClick={()=>setCartCardBrand(b)} style={{padding:".3rem .7rem",borderRadius:".4rem",fontSize:".76rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"1px solid "+(cartCardBrand===b?"#4f5ef0":"var(--bdr2)"),background:cartCardBrand===b?"#4f5ef0":"transparent",color:cartCardBrand===b?"#fff":"var(--navoff)"}}>
                  {b}
                </button>
              ))}
            </div>
            {cartTaxRate>0&&cartTotal>0&&(
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:".4rem"}}>
                {[{l:"Bruto",v:fmt(cartTotal),c:"var(--tx)"},{l:"Taxa "+fmtPct(cartTaxRate),v:"- "+fmt(cartTaxAmount),c:"#f56565"},{l:"Líquido",v:fmt(cartNetTotal),c:"#10b981"}].map(m=>(
                  <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".4rem",padding:".4rem .3rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".1rem"}}>{m.l}</div>
                    <div style={{fontSize:".82rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {!cartSplitEnabled&&cartPayment==="Crédito Parcelado"&&(
          <div style={{background:"#4f5ef015",border:"1px solid #4f5ef040",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".5rem"}}>
            <div style={{fontSize:".68rem",color:"#4f5ef0",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".45rem",fontWeight:700}}>Número de Parcelas</div>
            <div style={{display:"flex",gap:".3rem",flexWrap:"wrap",marginBottom:".45rem"}}>
              {[2,3,4,5,6,7,8,9,10,11,12].map(n=>(
                <button key={n} onClick={()=>setCartParcelas(n)} style={{padding:".28rem .6rem",borderRadius:".35rem",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",fontWeight:700,border:"1px solid "+(cartParcelas===n?"#4f5ef0":"var(--bdr2)"),background:cartParcelas===n?"#4f5ef0":"transparent",color:cartParcelas===n?"#fff":"var(--navoff)"}}>
                  {n}x
                </button>
              ))}
            </div>
            {cartNetTotal>0&&(
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:"var(--sumbox)",borderRadius:".4rem",padding:".45rem .7rem",marginBottom:".35rem"}}>
                <span style={{fontSize:".75rem",color:"var(--tx4)"}}>{cartParcelas}x de (líquido)</span>
                <span style={{fontSize:".88rem",fontWeight:700,color:"#4f5ef0",fontFamily:"'Syne',sans-serif"}}>{fmt(cartInstallmentNet)}</span>
              </div>
            )}
            <div style={{fontSize:".7rem",color:"#8b44f0",display:"flex",gap:".3rem",alignItems:"center"}}>
              <Ic n="info" s={11}/>Parcelas criadas em A Receber com valor líquido
            </div>
          </div>
        )}
        {!cartSplitEnabled&&getSettle(cartPayment).type==="duedate"&&(
          <div style={{background:"#f59e0b12",border:"1px solid #f59e0b40",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".5rem"}}>
            <div style={{fontSize:".68rem",color:"#f59e0b",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".45rem",fontWeight:700}}>📝 {cartPayment} · Data de vencimento</div>
            <input type="date" value={cartDueDate} onChange={e=>setCartDueDate(e.target.value)} style={{...IS,fontSize:".9rem"}}/>
            <div style={{display:"flex",gap:".35rem",marginTop:".45rem",flexWrap:"wrap"}}>
              {[{l:"+7 dias",d:7},{l:"+15 dias",d:15},{l:"+30 dias",d:30}].map(o=>(
                <button key={o.l} onClick={()=>{const dt=new Date();dt.setDate(dt.getDate()+o.d);setCartDueDate(dt.toISOString().slice(0,10));}} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".25rem .6rem",fontSize:".7rem",color:"var(--tx4)",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>{o.l}</button>
              ))}
            </div>
            <div style={{fontSize:".7rem",color:"#8b44f0",display:"flex",gap:".3rem",alignItems:"center",marginTop:".5rem"}}>
              <Ic n="info" s={11}/>{cartDueDate?"Vai para A Receber com vencimento em "+cartDueDate.split("-").reverse().join("/"):"Sem data → vence em 30 dias. Marque como recebido quando o cliente pagar."}
            </div>
            {cartClient.id&&(()=>{
              const cli=clients.find(c=>c.id===cartClient.id);const lim=parseFloat(cli&&cli.credit_limit)||0;
              if(lim<=0)return null;
              const fin=clientFin(cartClient.id,cartClient.name);
              const novoTotal=fin.aberto+cartTotal;
              const disp=lim-fin.aberto;
              if(novoTotal>lim+0.005)return <div style={{marginTop:".5rem",background:"#f5656518",border:"1px solid #f5656540",borderRadius:".45rem",padding:".5rem .7rem",fontSize:".72rem",color:"#f56565",fontWeight:600}}>⚠️ Limite de crédito estourado: {cartClient.name} já tem {fmt(fin.aberto)} em aberto + esta venda ({fmt(cartTotal)}) = {fmt(novoTotal)}, acima do limite de {fmt(lim)}. Você pode prosseguir, mas avalie o risco.</div>;
              return <div style={{marginTop:".5rem",fontSize:".68rem",color:"#10b981"}}>✅ Dentro do limite · disponível {fmt(disp)} de {fmt(lim)}{fin.vencidasAbertas>0?" · ⚠️ cliente tem "+fin.vencidasAbertas+" conta(s) vencida(s)":""}</div>;
            })()}
          </div>
        )}
        {/* ── PAGAMENTO DIVIDIDO (split) ── */}
        <div style={{background:cartSplitEnabled?"#8b44f010":"var(--pill)",border:`1px solid ${cartSplitEnabled?"#8b44f040":"var(--bdr)"}`,borderRadius:".55rem",padding:".7rem .85rem",marginBottom:".5rem"}}>
          <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",fontSize:".82rem",fontWeight:600,color:cartSplitEnabled?"#8b44f0":"var(--tx3)"}}>
            <input type="checkbox" checked={cartSplitEnabled} onChange={e=>{const on=e.target.checked;setCartSplitEnabled(on);if(on&&cartSplits.length===0){const m0=activeSimplePays[0]||"PIX";const m1=activeSimplePays[1]||m0;setCartSplits([{key:uid(),method:m0,amount:"",brand:"Visa",parcelas:2},{key:uid(),method:m1,amount:"",brand:"Visa",parcelas:2}]);}}} style={{width:16,height:16,accentColor:"#8b44f0"}}/>
            💳 Dividir em mais de uma forma de pagamento
          </label>
          {cartSplitEnabled&&(
            <div style={{marginTop:".7rem"}}>
              <div style={{fontSize:".66rem",color:"var(--sub)",marginBottom:".55rem"}}>Combine quantas formas quiser, em qualquer valor (ex: parte no PIX, parte no cartão, parte em dinheiro). A soma deve ser igual ao total da venda.</div>
              {cartSplits.map((sp)=>{
                const spIsCard=["Débito","Crédito à Vista","Crédito Parcelado"].includes(sp.method);
                return (
                <div key={sp.key} style={{background:"var(--card)",border:"1px solid var(--bdr2)",borderRadius:".5rem",padding:".55rem .6rem",marginBottom:".5rem"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 108px 30px",gap:".4rem",alignItems:"center"}}>
                    <select value={sp.method} onChange={e=>updateSplit(sp.key,{method:e.target.value})} style={{...IS,fontSize:".8rem"}}>
                      <optgroup label="Sem taxa">{activeSimplePays.map(m=><option key={m} value={m}>{m}</option>)}</optgroup>
                      <optgroup label="Cartão"><option value="Débito">Débito</option><option value="Crédito à Vista">Crédito à Vista</option><option value="Crédito Parcelado">Crédito Parcelado</option></optgroup>
                    </select>
                    <input type="number" min="0" step="0.01" onFocus={e=>e.target.select()} value={sp.amount} onChange={e=>updateSplit(sp.key,{amount:e.target.value})} placeholder="R$ 0,00" style={{...IS,fontSize:".82rem",textAlign:"right"}}/>
                    <button onClick={()=>removeSplit(sp.key)} disabled={cartSplits.length<=1} title="Remover forma" style={{background:"none",border:"none",color:cartSplits.length<=1?"var(--tx6)":"#f56565",cursor:cartSplits.length<=1?"default":"pointer",padding:".2rem",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="trash" s={14}/></button>
                  </div>
                  {spIsCard&&(
                    <div style={{display:"flex",gap:".3rem",flexWrap:"wrap",marginTop:".45rem"}}>
                      {CARD_BRANDS.map(b=>(
                        <button key={b} onClick={()=>updateSplit(sp.key,{brand:b})} style={{padding:".22rem .55rem",borderRadius:".35rem",fontSize:".68rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"1px solid "+(sp.brand===b?"#8b44f0":"var(--bdr2)"),background:sp.brand===b?"#8b44f0":"transparent",color:sp.brand===b?"#fff":"var(--navoff)"}}>{b}</button>
                      ))}
                    </div>
                  )}
                  {sp.method==="Crédito Parcelado"&&(
                    <div style={{display:"flex",gap:".25rem",flexWrap:"wrap",marginTop:".4rem",alignItems:"center"}}>
                      <span style={{fontSize:".66rem",color:"var(--sub)",marginRight:".2rem"}}>Parcelas:</span>
                      {[2,3,4,5,6,7,8,9,10,11,12].map(n=>(
                        <button key={n} onClick={()=>updateSplit(sp.key,{parcelas:n})} style={{padding:".2rem .45rem",borderRadius:".3rem",fontSize:".68rem",fontFamily:"'DM Sans',sans-serif",fontWeight:700,border:"1px solid "+((parseInt(sp.parcelas)||2)===n?"#4f5ef0":"var(--bdr2)"),background:(parseInt(sp.parcelas)||2)===n?"#4f5ef0":"transparent",color:(parseInt(sp.parcelas)||2)===n?"#fff":"var(--navoff)"}}>{n}x</button>
                      ))}
                    </div>
                  )}
                </div>
                );
              })}
              <button onClick={addSplit} style={{width:"100%",padding:".5rem",borderRadius:".45rem",border:"1px dashed #8b44f060",background:"#8b44f008",color:"#8b44f0",fontSize:".76rem",fontWeight:600,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:".35rem",marginBottom:".55rem"}}>
                <Ic n="plus" s={13}/> Adicionar forma de pagamento
              </button>
              <div style={{display:"flex",flexDirection:"column",gap:".25rem",background:"var(--sumbox)",borderRadius:".45rem",padding:".55rem .7rem"}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:".76rem"}}>
                  <span style={{color:"var(--tx4)"}}>Soma das formas</span>
                  <span style={{fontWeight:700,color:"var(--tx2)"}}>{fmt(cartSplitSum)}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:".76rem"}}>
                  <span style={{color:"var(--tx4)"}}>Total da venda</span>
                  <span style={{fontWeight:700,color:"var(--tx2)"}}>{fmt(cartTotal)}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderTop:"1px solid var(--bdr2)",marginTop:".15rem",paddingTop:".3rem",fontSize:".78rem"}}>
                  {Math.abs(cartSplitRemaining)<0.01
                    ? <><span style={{color:"#10b981",fontWeight:700,display:"flex",alignItems:"center",gap:".3rem"}}><Ic n="check" s={13}/>Valores conferem</span><span style={{color:"#10b981",fontWeight:700}}>OK</span></>
                    : <><span style={{color:cartSplitRemaining>0?"#f59e0b":"#f56565",fontWeight:700}}>{cartSplitRemaining>0?"Falta distribuir":"Passou do total"}</span><span style={{color:cartSplitRemaining>0?"#f59e0b":"#f56565",fontWeight:700}}>{(cartSplitRemaining>0?"":"+")+fmt(Math.abs(cartSplitRemaining))}</span></>
                  }
                </div>
              </div>
            </div>
          )}
        </div>
        <Inp label="Observações" hint="opcional" placeholder="Ex: Entregar no período da tarde..." value={cartNotes} onChange={e=>setCartNotes(e.target.value)}/>

        {/*  RESUMO TOTAL  */}
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
                <span style={{color:"var(--tx4)"}}>🏷️ Desconto</span>
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

        {/*  AÇÕES  */}
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between",alignItems:"center"}}>
          <button onClick={cartReset} style={{background:"none",border:"none",color:"var(--tx5)",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",cursor:"pointer",display:"flex",alignItems:"center",gap:".3rem"}}>
            <Ic n="trash" s={12}/> Limpar
          </button>
          <div style={{display:"flex",gap:".5rem"}}>
            <Btn v="ghost" onClick={()=>{setModal(null);cartReset();}}>Cancelar</Btn>
            <Btn v="ok" onClick={registerSale} disabled={cartItems.every(i=>!i.product_id)||!cartSplitValid}>
              <Ic n="save" s={13}/>{cartSplitEnabled&&!cartSplitValid?"Ajuste os pagamentos":"Finalizar Venda · "+fmt(cartTotal)}
            </Btn>
          </div>
        </div>
      </Modal>
    )}

    {/* Editar venda */}
    {/* ── EDITAR PRODUTO ── */}
    {modal==="editProd"&&editing&&(
      <Modal title={"✏️ "+editing.name} onClose={()=>{setModal(null);setEditing(null);}} icon="edit" wide>
        <R2>
          <Inp label="Código" value={editing.code||""} onChange={e=>setEditing(v=>({...v,code:e.target.value}))}/>
          <Sel label="Categoria" value={editing.category||"outro"} onChange={e=>setEditing(v=>({...v,category:e.target.value}))}>
            {activeCats.map(c=><option key={c.key} value={c.key}>{c.icon} {c.label}</option>)}
          </Sel>
        </R2>
        <Inp label="Nome *" value={editing.name||""} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
        <R2>
          <Sel label="Fornecedor" hint="opcional" value={editing.supplier_id||""} onChange={e=>setEditing(v=>({...v,supplier_id:e.target.value}))}>
            <option value="">Nenhum</option>
            {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </Sel>
          <Inp label="Unidade" value={editing.unit||"un"} onChange={e=>setEditing(v=>({...v,unit:e.target.value}))}/>
        </R2>
        <R2>
          <Inp label="Lote" value={editing.batch||""} onChange={e=>setEditing(v=>({...v,batch:e.target.value}))}/>
          <Inp label="Vencimento" type="date" value={editing.expiry||""} onChange={e=>setEditing(v=>({...v,expiry:e.target.value}))}/>
        </R2>
        <R2>
          <Inp label="Custo/un (R$)" type="number" min="0" step="0.01" onFocus={e=>e.target.select()} value={editing.cost_per_unit||""} onChange={e=>setEditing(v=>({...v,cost_per_unit:e.target.value}))}/>
          <Inp label="Preço venda/un (R$)" type="number" min="0" step="0.01" onFocus={e=>e.target.select()} value={editing.price_per_unit||""} onChange={e=>setEditing(v=>({...v,price_per_unit:e.target.value}))}/>
        </R2>
        <R2>
          <Inp label="Estoque atual" type="number" min="0" onFocus={e=>e.target.select()} value={editing.stock_qty||""} onChange={e=>setEditing(v=>({...v,stock_qty:e.target.value}))}/>
          <Inp label="Estoque mínimo" type="number" min="0" onFocus={e=>e.target.select()} value={editing.min_stock||""} onChange={e=>setEditing(v=>({...v,min_stock:e.target.value}))}/>
        </R2>
        {editing.category==="tirzepatida"&&(
          <div style={{display:"flex",gap:".4rem",alignItems:"center"}}>
            <span style={{fontSize:".65rem",color:"var(--sub)"}}>💊 Total mg/unidade:</span>
            <input type="number" min="0" step="0.001" onFocus={e=>e.target.select()} value={editing.total_mg||""} onChange={e=>setEditing(v=>({...v,total_mg:e.target.value}))} placeholder="ex: 15" style={{...IS,width:80,color:"#8b44f0",fontWeight:700}}/>
            <span style={{fontSize:".8rem",color:"var(--tx5)"}}>mg</span>
          </div>
        )}
        <Inp label="Descrição" hint="opcional" value={editing.description||""} onChange={e=>setEditing(v=>({...v,description:e.target.value}))}/>
        {(()=>{const{markup,margin}=calcM(editing.cost_per_unit,editing.price_per_unit);return markup>0&&(<div style={{background:"var(--pill)",borderRadius:".45rem",padding:".45rem .75rem",fontSize:".73rem",display:"flex",gap:"1rem"}}><span style={{color:"var(--tx4)"}}>Markup: <strong style={{color:"#f59e0b"}}>{fmtPct(markup)}</strong></span><span style={{color:"var(--tx4)"}}>Margem: <strong style={{color:"#10b981"}}>{fmtPct(margin)}</strong></span></div>);})()}
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end",marginTop:".5rem"}}>
          <Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn>
          <Btn v="ok" onClick={saveProduct}><Ic n="save" s={13}/>Salvar</Btn>
        </div>
      </Modal>
    )}
    {modal==="editSale"&&editing&&(
      <Modal title="Editar Venda" onClose={()=>{setModal(null);setEditing(null);}} icon="edit">
        <Inp label="Produto" value={editing.product_name} onChange={e=>setEditing(v=>({...v,product_name:e.target.value}))}/>
        <R2><Inp label="Quantidade" type="number" min="1" value={editing.quantity} onChange={e=>setEditing(v=>({...v,quantity:e.target.value}))}/><Inp label="Preço unit. (R$)" type="number" min="0" step="0.01" value={editing.unit_price} onChange={e=>setEditing(v=>({...v,unit_price:e.target.value}))}/></R2>
        <Inp label="Cliente" value={editing.client_name||""} onChange={e=>setEditing(v=>({...v,client_name:e.target.value}))}/>
        <R2><Sel label="Pagamento" value={editing.payment_method||"PIX"} onChange={e=>setEditing(v=>({...v,payment_method:e.target.value}))}>{PAYS.map(m=><option key={m} value={m}>{m}</option>)}</Sel><Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/></R2>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {canDelete&&<Btn v="del" sm onClick={()=>{deleteSale(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn v="ok" onClick={updateSale}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Entrada estoque */}
    {modal==="stockEntry"&&(
      <Modal title="Entrada de Estoque" onClose={()=>setModal(null)} icon="stock">
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".55rem .8rem",marginBottom:".8rem",fontSize:".73rem",color:"#10b981",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="arrup" s={12}/>Entrada gera saída no caixa (custo de compra)</div>
        <Sel label="Produto *" value={stF.product_id} onChange={e=>setStF(s=>({...s,product_id:e.target.value}))}>
          <option value="">Selecione...</option>{products.map(p=><option key={p.id} value={p.id}>{(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name} — Atual: {p.stock_qty} {p.unit||"un"}</option>)}
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
        <Sel label="Categoria" value={editing.category||"outro"} onChange={e=>setEditing(v=>({...v,category:e.target.value}))}>{activeCats.map(c=><option key={c.key} value={c.key}>{c.icon} {c.label}</option>)}</Sel>
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {canDelete&&<Btn v="del" sm onClick={()=>{deleteCash(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>{setModal(null);setEditing(null);}}>Cancelar</Btn><Btn v="ok" onClick={updateCash}><Ic n="save" s={13}/>Salvar</Btn></div>
        </div>
      </Modal>
    )}

    {/* Novo produto */}
    {modal==="produto"&&(
      <Modal title="Novo Produto" onClose={()=>setModal(null)} icon="product" wide>
        <R2><Inp label="Código" hint="auto" placeholder="PRD-001" value={pf.code} onChange={e=>setPf(f=>({...f,code:e.target.value}))}/><Sel label="Categoria" value={pf.category} onChange={e=>setPf(f=>({...f,category:e.target.value}))}>{activeCats.map(c=><option key={c.key} value={c.key}>{c.icon} {c.label}</option>)}</Sel></R2>
                <Inp label="Nome *" value={pf.name} onChange={e=>setPf(f=>({...f,name:e.target.value}))}/>
        <R2><Sel label="Fornecedor" value={pf.supplier_id||""} onChange={e=>setPf(f=>({...f,supplier_id:e.target.value}))}><option value="">Nenhum</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</Sel><Inp label="Unidade" value={pf.unit||"un"} onChange={e=>setPf(f=>({...f,unit:e.target.value}))}/></R2>
        <R2><Inp label="Lote" value={pf.batch||""} onChange={e=>setPf(f=>({...f,batch:e.target.value}))}/><Inp label="Vencimento" type="date" value={pf.expiry||""} onChange={e=>setPf(f=>({...f,expiry:e.target.value}))}/></R2>
        <R2><Inp label="Custo/un (R$)" type="number" min="0" step="0.01" value={pf.cost_per_unit} onChange={e=>setPf(f=>({...f,cost_per_unit:e.target.value}))}/><Inp label="Preço venda/un (R$)" type="number" min="0" step="0.01" value={pf.price_per_unit} onChange={e=>setPf(f=>({...f,price_per_unit:e.target.value}))}/></R2>
        <MPreview cost={pf.cost_per_unit} price={pf.price_per_unit}/>
        <R2><Inp label="Estoque atual" type="number" min="0" onFocus={e=>e.target.select()} value={pf.stock_qty} onChange={e=>setPf(f=>({...f,stock_qty:e.target.value}))}/><Inp label="Estoque mínimo" type="number" min="0" value={pf.min_stock} onChange={e=>setPf(f=>({...f,min_stock:e.target.value}))}/></R2>
        {pf.category==="tirzepatida"&&(
          <div>
            <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Total mg / unidade <span style={{color:"var(--tx6)"}}>(opcional)</span></div>
            <div style={{display:"flex",gap:".4rem",alignItems:"center"}}>
              <input type="number" min="0" step="0.001" value={pf.total_mg||""} onChange={e=>setPf(f=>({...f,total_mg:e.target.value}))} placeholder="Ex: 15" style={{...IS,width:90,color:"#8b44f0",fontWeight:700}}/>
              <span style={{fontSize:".8rem",color:"var(--tx5)"}}>mg / unidade</span>
            </div>
          </div>
        )}
        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          {canDelete&&<Btn v="del" sm onClick={()=>delProduct(pf.id)}><Ic n="trash" s={12}/>Excluir</Btn>}
          <div style={{display:"flex",gap:".5rem",marginLeft:"auto"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addProduct}><Ic n="save" s={13}/>Salvar</Btn></div>
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
      <Inp label="Limite de crédito p/ fiado (R$)" hint="0 = sem limite definido" type="number" placeholder="0,00" value={cf.credit_limit} onChange={e=>setCf(f=>({...f,credit_limit:e.target.value}))}/>
      <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={addClient}><Ic n="save" s={13}/>Salvar</Btn></div>
    </Modal>)}
    {modal==="editCliente"&&editing&&(<Modal title="Editar Cliente" onClose={()=>{setModal(null);setEditing(null);}} icon="edit" wide>
      <Inp label="Nome *" value={editing.name} onChange={e=>setEditing(v=>({...v,name:e.target.value}))}/>
      <R2><Inp label="Telefone" type="tel" value={editing.phone||""} onChange={e=>setEditing(v=>({...v,phone:e.target.value}))}/><Inp label="E-mail" type="email" value={editing.email||""} onChange={e=>setEditing(v=>({...v,email:e.target.value}))}/></R2>
      <Inp label="Observações" value={editing.notes||""} onChange={e=>setEditing(v=>({...v,notes:e.target.value}))}/>
      <Inp label="Limite de crédito p/ fiado (R$)" hint="0 = sem limite" type="number" value={editing.credit_limit||""} onChange={e=>setEditing(v=>({...v,credit_limit:e.target.value}))}/>
      <div style={{borderTop:"1px solid var(--bdr)",paddingTop:".85rem",marginTop:".25rem"}}>
        <div style={{fontWeight:700,fontSize:".78rem",color:"#4f5ef0",marginBottom:".65rem"}}>💉 Protocolo de Tratamento (opcional)</div>
        <R2><Inp label="Dose atual" placeholder="Ex: 2.5mg, 5mg..." value={editing.dose||""} onChange={e=>setEditing(v=>({...v,dose:e.target.value}))}/><Inp label="Intervalo (dias)" type="number" min="1" placeholder="7" value={editing.interval_days||""} onChange={e=>setEditing(v=>({...v,interval_days:e.target.value}))}/></R2>
        <Inp label="Início do tratamento" type="date" value={editing.treatment_start||""} onChange={e=>setEditing(v=>({...v,treatment_start:e.target.value}))}/>
        <Inp label="Notas do protocolo" placeholder="Ex: Responde bem à dose 5mg..." value={editing.treatment_notes||""} onChange={e=>setEditing(v=>({...v,treatment_notes:e.target.value}))}/>
      </div>
      <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
        {canDelete&&<Btn v="del" sm onClick={()=>{delClient(editing.id);setModal(null);setEditing(null);}}><Ic n="trash" s={12}/>Excluir</Btn>}
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
      {editing.role!=="admin"&&<div style={{borderTop:"1px solid var(--bdr)",paddingTop:".8rem",marginBottom:".8rem"}}>
        <div style={{fontSize:".72rem",color:"#8b44f0",marginBottom:".5rem",display:"flex",gap:".3rem",alignItems:"center",fontWeight:700}}><Ic n="key" s={12}/>Permissões de acesso</div>
        <div style={{fontSize:".64rem",color:"var(--tx5)",marginBottom:".3rem",textTransform:"uppercase",letterSpacing:".04em"}}>Aplicar perfil pronto (De ▸ Para)</div>
        <div style={{display:"flex",gap:".35rem",flexWrap:"wrap",marginBottom:".7rem"}}>
          {[{l:"🟢 Acesso total",p:{dashboard:true,vendas:true,estoque:true,financeiro:true,clientes:true,candelete:true}},{l:"🔵 Operador padrão",p:{dashboard:true,vendas:true,estoque:true,financeiro:true,clientes:true,candelete:false}},{l:"🟡 Básico (vendas)",p:{dashboard:true,vendas:true,estoque:false,financeiro:false,clientes:true,candelete:false}}].map(t=>(
            <button key={t.l} onClick={()=>setEditing(v=>({...v,permissions:{...t.p}}))} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".45rem",padding:".32rem .62rem",fontSize:".71rem",color:"var(--tx3)",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer"}}>{t.l}</button>
          ))}
        </div>
        <div style={{fontSize:".64rem",color:"var(--tx5)",marginBottom:".3rem",textTransform:"uppercase",letterSpacing:".04em"}}>Telas liberadas (De: tela ▸ Para: acesso)</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr",gap:".4rem"}}>
          {PERM_MODULES.map(mod=>{
            const perms=editing.permissions||{};
            const on=perms[mod.id]!==false;
            return(
              <label key={mod.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:".5rem",cursor:"pointer",background:on?"#8b44f012":"var(--pill)",border:`1px solid ${on?"#8b44f040":"var(--bdr2)"}`,borderRadius:".5rem",padding:".5rem .75rem"}}>
                <span style={{fontSize:".8rem",fontWeight:600,color:on?"var(--tx2)":"var(--tx5)"}}>{mod.l}</span>
                <span style={{display:"flex",alignItems:"center",gap:".45rem"}}>
                  <span style={{fontSize:".64rem",fontWeight:700,color:on?"#10b981":"var(--tx6)"}}>{on?"LIBERADO":"BLOQUEADO"}</span>
                  <input type="checkbox" checked={on} onChange={e=>setEditing(v=>({...v,permissions:{...(v.permissions||{}),[mod.id]:e.target.checked}}))} style={{width:18,height:18,accentColor:"#8b44f0",flexShrink:0}}/>
                </span>
              </label>
            );
          })}
        </div>
        <div style={{fontSize:".64rem",color:"var(--tx5)",margin:".6rem 0 .3rem",textTransform:"uppercase",letterSpacing:".04em"}}>Permissão especial</div>
        {(()=>{const cd=!!(editing.permissions||{}).candelete;return(
          <label style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:".5rem",cursor:"pointer",background:cd?"#f5656512":"var(--pill)",border:`1px solid ${cd?"#f5656540":"var(--bdr2)"}`,borderRadius:".5rem",padding:".5rem .75rem"}}>
            <span style={{fontSize:".8rem",fontWeight:600,color:cd?"#f56565":"var(--tx3)"}}>🗑️ Pode excluir registros do sistema</span>
            <input type="checkbox" checked={cd} onChange={e=>setEditing(v=>({...v,permissions:{...(v.permissions||{}),candelete:e.target.checked}}))} style={{width:18,height:18,accentColor:"#f56565",flexShrink:0}}/>
          </label>
        );})()}
        <div style={{fontSize:".66rem",color:"var(--tx6)",marginTop:".5rem"}}>Config é sempre exclusivo de administradores. Sem "pode excluir", os botões de exclusão ficam ocultos para este usuário.</div>
      </div>}
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

    {/* ══ COMPROVANTE ══ */}
    {showReceipt&&(
      <Modal title="Comprovante de Venda" onClose={()=>setShowReceipt(null)} icon="pdf" wide>
        <div style={{textAlign:"center",marginBottom:"1.25rem",paddingBottom:"1rem",borderBottom:"1px solid var(--bdr)"}}>
          <div style={{fontSize:"1.2rem",fontWeight:800,fontFamily:"'Syne',sans-serif",color:"var(--tx)"}}>{companyInfo.showInPDF!==false?(companyInfo.name||"CaixaPro · Tirzepatida"):"Comprovante de Venda"}</div>
          <div style={{fontSize:".75rem",color:"var(--tx5)",marginTop:".15rem"}}>Comprovante de Venda · {showReceipt[0]&&showReceipt[0].date}</div>
        </div>
        {showReceipt[0]&&showReceipt[0].client_name&&(
          <div style={{marginBottom:"1rem",padding:".6rem .85rem",background:"var(--pill)",borderRadius:".5rem"}}>
            <div style={{fontSize:".65rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".12rem"}}>Cliente</div>
            <div style={{fontWeight:700,color:"var(--tx)"}}>{showReceipt[0].client_name}</div>
          </div>
        )}
        <table style={{width:"100%",borderCollapse:"collapse",marginBottom:"1rem"}}>
          <thead>
            <tr style={{borderBottom:"1px solid var(--bdr)"}}>
              {["Produto","Qtd","Preço","Total"].map(h=><th key={h} style={{textAlign:"left",padding:".4rem .3rem",fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase"}}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {showReceipt.map((s,i)=>(
              <tr key={i} style={{borderBottom:"1px solid var(--sep)"}}>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx)"}}>{s.product_name}{s.unit_price===0&&" 🎁"}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx3)"}}>{s.quantity}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",color:"var(--tx3)"}}>{s.unit_price===0?"Cortesia":fmt(s.unit_price)}</td>
                <td style={{padding:".5rem .3rem",fontSize:".82rem",fontWeight:700,color:"#10b981"}}>{fmt(s.total_price)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{background:"var(--pill)",borderRadius:".5rem",padding:"1rem",marginBottom:"1rem"}}>
          {(()=>{
            const subtotal=showReceipt.reduce((a,s)=>a+s.total_price,0);
            const disc=parseFloat(showReceipt[0]?.discount)||0;
            const total=Math.max(0,subtotal-disc);
            return(<>
              {disc>0&&(
                <div style={{display:"flex",justifyContent:"space-between",fontSize:".8rem",paddingTop:".3rem"}}>
                  <span style={{color:"var(--tx4)"}}>Subtotal</span>
                  <span style={{color:"var(--tx3)"}}>{fmt(subtotal)}</span>
                </div>
              )}
              {disc>0&&(
                <div style={{display:"flex",justifyContent:"space-between",fontSize:".8rem"}}>
                  <span style={{color:"#4f5ef0"}}>🏷️ Desconto aplicado</span>
                  <span style={{color:"#4f5ef0",fontWeight:600}}>- {fmt(disc)}</span>
                </div>
              )}
              <div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid var(--bdr)",paddingTop:".5rem",marginTop:".25rem"}}>
                <span style={{fontWeight:700,color:"var(--tx)",fontFamily:"'Syne',sans-serif"}}>TOTAL</span>
                <span style={{fontWeight:800,fontSize:"1.1rem",color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(total)}</span>
              </div>
              {showReceipt[0]&&<div style={{fontSize:".73rem",color:"var(--sub)",textAlign:"right",marginTop:".2rem"}}>Pagamento: {showReceipt[0].payment_method}</div>}
            </>);
          })()}
        </div>
        <div style={{textAlign:"center",fontSize:".68rem",color:"var(--tx6)",marginBottom:"1rem"}}>{companyInfo.showInPDF!==false?"Obrigado pela compra! · "+(companyInfo.name||"CaixaPro")+(companyInfo.phone?" · "+companyInfo.phone:""):"Obrigado pela compra!"}</div>
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
          <Btn v="ghost" onClick={()=>setShowReceipt(null)}>Fechar</Btn>
          <Btn onClick={()=>window.print()}><Ic n="pdf" s={13}/>Imprimir</Btn>
        </div>
      </Modal>
    )}

    {/* ══ HISTÓRICO CLIENTE ══ */}
    {showClientHist&&(
      <Modal title={"Histórico · "+(showClientHist.name||"")} onClose={()=>setShowClientHist(null)} icon="analytics" wide>
        {(()=>{
          const cs=sales.filter(s=>s.client_id===showClientHist.id||s.client_name===showClientHist.name);
          const total=batchRevenue(cs);
          const prodMap={};
          cs.forEach(s=>{prodMap[s.product_name]=(prodMap[s.product_name]||0)+s.quantity;});
          const lastBuy=cs.length>0?cs[0].date:null;
          const daysSince=cs.length>0?Math.floor((new Date()-new Date(cs[0].created_at))/86400000):null;
          return(<>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".6rem",marginBottom:"1rem"}}>
              <KCard label="Total gasto" value={fmt(total)} color="#10b981"/>
              <KCard label="Compras" value={cs.length} color="#4f5ef0"/>
              <KCard label="Última compra" value={lastBuy||"—"} sub={daysSince!==null?daysSince+"d atrás":""} color={daysSince!==null&&daysSince>30?"#f59e0b":"#10b981"}/>
            </div>
            {showClientHist.dose&&(
              <div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98115)",border:"1px solid #4f5ef030",borderRadius:".65rem",padding:".85rem",marginBottom:"1rem"}}>
                <div style={{fontWeight:700,fontSize:".8rem",color:"#4f5ef0",marginBottom:".5rem"}}>💉 Protocolo de Tratamento</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".5rem",textAlign:"center"}}>
                  {[{l:"Dose",v:showClientHist.dose,c:"#4f5ef0"},{l:"Intervalo",v:(showClientHist.interval_days||7)+"d",c:"#8b44f0"},{l:"Início",v:showClientHist.treatment_start?new Date(showClientHist.treatment_start+"T00:00:00").toLocaleDateString("pt-BR"):"—",c:"#10b981"}].map(m=>(
                    <div key={m.l}>
                      <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".12rem"}}>{m.l}</div>
                      <div style={{fontWeight:700,fontSize:".88rem",color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {Object.keys(prodMap).length>0&&(
              <div style={{background:"var(--pill)",borderRadius:".6rem",padding:".75rem",marginBottom:"1rem"}}>
                <div style={{fontWeight:700,fontSize:".78rem",color:"var(--tx2)",marginBottom:".5rem"}}>📦 Produtos mais comprados</div>
                {Object.entries(prodMap).sort((a,b)=>b[1]-a[1]).map(([n,q])=>(
                  <div key={n} style={{display:"flex",justifyContent:"space-between",padding:".3rem 0",borderBottom:"1px solid var(--sep)",fontSize:".78rem"}}>
                    <span style={{color:"var(--tx)"}}>{n}</span>
                    <span style={{color:"#4f5ef0",fontWeight:700}}>{q} un</span>
                  </div>
                ))}
              </div>
            )}
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".8rem",color:"var(--tx2)",marginBottom:".5rem"}}>🛒 Últimas compras</div>
            <div style={{background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".65rem"}}>
              {cs.length===0
                ?<p style={{color:"var(--tx5)",textAlign:"center",padding:"1.5rem",fontSize:".8rem"}}>Nenhuma compra.</p>
                :cs.map(s=>(
                  <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:".6rem 1rem",borderBottom:"1px solid var(--sep)"}}>
                    <div>
                      <div style={{fontSize:".82rem",color:"var(--tx)",fontWeight:600}}>{s.product_name}</div>
                      <div style={{fontSize:".65rem",color:"var(--tx5)"}}>{s.date} · {s.quantity}un · {s.payment_method}</div>
                    </div>
                    <span style={{fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif",fontSize:".85rem"}}>{fmt(s.total_price)}</span>
                  </div>
                ))
              }
            </div>
          </>);
        })()}
      </Modal>
    )}

    {/* ══ CALCULADORA IMPORTAÇÃO ══ */}
    {showImportCalc&&(
      <Modal title="🧮 Calculadora de Custo de Importação" onClose={()=>setShowImportCalc(false)}>
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".55rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="info" s={12}/>Calcule o custo real por unidade incluindo frete e taxas</div>
        <Inp label="Custo total da remessa (R$) *" type="number" min="0" step="0.01" placeholder="Ex: 3500,00" value={importCalc.totalCost} onChange={e=>setImportCalc(v=>({...v,totalCost:e.target.value}))}/>
        <Inp label="Quantidade de unidades *" type="number" min="1" placeholder="Ex: 20 ampolas" value={importCalc.qty} onChange={e=>setImportCalc(v=>({...v,qty:e.target.value}))}/>
        <Inp label="Extras — frete, taxas, câmbio (R$)" hint="opcional" type="number" min="0" step="0.01" placeholder="Ex: 250,00" value={importCalc.extras} onChange={e=>setImportCalc(v=>({...v,extras:e.target.value}))}/>
        {importCalc.totalCost&&importCalc.qty&&parseInt(importCalc.qty)>0&&(()=>{
          const tot=(parseFloat(importCalc.totalCost)||0)+(parseFloat(importCalc.extras)||0);
          const qty=parseInt(importCalc.qty)||1;
          const custo=tot/qty;
          return(
            <div style={{background:"linear-gradient(135deg,#4f5ef015,#10b98115)",border:"1px solid #4f5ef030",borderRadius:".6rem",padding:"1rem",marginTop:".5rem"}}>
              <div style={{fontWeight:700,fontSize:".82rem",color:"#4f5ef0",marginBottom:".65rem",fontFamily:"'Syne',sans-serif"}}>📊 Resultado</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:".5rem",textAlign:"center"}}>
                {[{l:"Custo/unidade",v:fmt(custo),c:"#f56565"},{l:"Preço 100% markup",v:fmt(custo*2),c:"#f59e0b"},{l:"Preço 200% markup",v:fmt(custo*3),c:"#10b981"}].map(m=>(
                  <div key={m.l} style={{background:"var(--pill)",borderRadius:".5rem",padding:".6rem .5rem"}}>
                    <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".15rem"}}>{m.l}</div>
                    <div style={{fontWeight:800,fontSize:".95rem",color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                  </div>
                ))}
              </div>
              <div style={{fontSize:".7rem",color:"var(--tx5)",marginTop:".65rem",textAlign:"center"}}>Total: {fmt(tot)} · {qty} unidades</div>
            </div>
          );
        })()}
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:"1rem"}}>
          <Btn v="ghost" onClick={()=>{setShowImportCalc(false);setImportCalc({totalCost:"",qty:"",extras:""});}}>Fechar</Btn>
        </div>
      </Modal>
    )}


    {/* ══ CARD CONFIG ══ */}
    {showCardConfig&&(()=>{
      return(
        <Modal title="💳 Taxas de Cartão" onClose={()=>setShowCardConfig(false)} icon="dollar" wide>
          <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0",display:"flex",gap:".3rem",alignItems:"center"}}>
            <Ic n="info" s={12}/>Taxas cobradas pelas operadoras. Afetam o valor líquido nas vendas com cartão.
          </div>
          <div style={{overflowX:"auto",marginBottom:"1rem"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:".78rem"}}>
              <thead>
                <tr style={{borderBottom:"2px solid var(--bdr)"}}>
                  <th style={{textAlign:"left",padding:".5rem .6rem",color:"var(--sub)",fontSize:".68rem",textTransform:"uppercase",fontWeight:600}}>Bandeira</th>
                  {CARD_MODES.map(m=><th key={m.key} style={{textAlign:"center",padding:".5rem .4rem",color:"var(--sub)",fontSize:".68rem",textTransform:"uppercase",fontWeight:600}}>{m.label}</th>)}
                </tr>
              </thead>
              <tbody>
                {CARD_BRANDS.map(brand=>(
                  <tr key={brand} style={{borderBottom:"1px solid var(--sep)"}}>
                    <td style={{padding:".55rem .6rem",fontWeight:700,color:"var(--tx)",whiteSpace:"nowrap"}}>
                      <span style={{background:"#4f5ef020",color:"#4f5ef0",borderRadius:".3rem",padding:".15rem .5rem",fontSize:".75rem"}}>{brand}</span>
                    </td>
                    {CARD_MODES.map(mode=>(
                      <td key={mode.key} style={{padding:".4rem"}}>
                        <div style={{position:"relative",display:"flex",alignItems:"center"}}>
                          <input type="number" min="0" max="20" step="0.01"
                            value={localTaxes[brand]?.[mode.key]||""}
                            onChange={e=>updateLocalTax(brand,mode.key,e.target.value)}
                            style={{...IS,textAlign:"center",paddingRight:"1.4rem",fontSize:".8rem",width:80}}
                          />
                          <span style={{position:"absolute",right:".45rem",fontSize:".7rem",color:"var(--tx5)"}}>%</span>
                        </div>
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:"1rem",fontSize:".73rem",color:"var(--tx4)"}}>
            💡 Exemplo: Venda R$ 100 com Visa Débito a 1.99% → taxa R$ 1,99 → líquido R$ 98,01
          </div>
          <div style={{background:"var(--pill)",border:"1px solid var(--bdr)",borderRadius:".5rem",padding:".7rem .85rem",marginBottom:"1rem"}}>
            <label style={{display:"flex",alignItems:"center",gap:".5rem",cursor:"pointer",marginBottom:localTaxes.__opts?.antecip?".6rem":0}}>
              <input type="checkbox" checked={!!localTaxes.__opts?.antecip} onChange={e=>setLocalTaxes(t=>({...t,__opts:{...(t.__opts||{}),antecip:e.target.checked}}))} style={{width:18,height:18,accentColor:"#8b44f0"}}/>
              <span style={{fontWeight:700,fontSize:".82rem",color:localTaxes.__opts?.antecip?"#8b44f0":"var(--tx3)"}}>⚡ Habilitar taxa de antecipação (recebimento na hora)</span>
            </label>
            {localTaxes.__opts?.antecip&&<div>
              <div style={{fontSize:".68rem",color:"var(--tx5)",marginBottom:".5rem"}}>% adicional somado às vendas no crédito (à vista e parcelado). Débito não é afetado.</div>
              <div style={{display:"flex",gap:".5rem",flexWrap:"wrap"}}>
                {CARD_BRANDS.map(brand=>(
                  <div key={brand} style={{display:"flex",flexDirection:"column",alignItems:"center"}}>
                    <span style={{fontSize:".62rem",color:"var(--sub)",marginBottom:".15rem"}}>{brand}</span>
                    <div style={{display:"flex",alignItems:"center",gap:".15rem"}}>
                      <input type="number" min="0" max="20" step="0.01" value={localTaxes[brand]?.antecip||""} onChange={e=>setLocalTaxes(t=>({...t,[brand]:{...t[brand],antecip:e.target.value}}))} style={{...IS,textAlign:"center",width:64,fontSize:".78rem"}}/>
                      <span style={{fontSize:".68rem",color:"var(--tx5)"}}>%</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>}
          </div>
          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
            <Btn v="ghost" onClick={()=>setLocalTaxes(JSON.parse(JSON.stringify(DEFAULT_TAXES)))}> Restaurar padrões</Btn>
            <div style={{display:"flex",gap:".5rem"}}>
              <Btn v="ghost" onClick={()=>setShowCardConfig(false)}>Cancelar</Btn>
              <Btn v="ok" onClick={()=>{saveCardTaxes(localTaxes);setShowCardConfig(false);}}><Ic n="save" s={13}/>Salvar Taxas</Btn>
            </div>
          </div>
        </Modal>
      );
    })()}

    {/* ══ GOALS SCREEN ══ */}
    {showGoals&&(()=>{
      const MONTHS=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
      const now=new Date();
      const rows=MONTHS.map((mName,mi)=>{
        const key=goalYear+"-"+(mi+1).toString().padStart(2,"0");
        const goal=monthlyGoals[key]||0;
        const monthStart=new Date(goalYear,mi,1);
        const monthEnd=new Date(goalYear,mi+1,0,23,59,59);
        const mSales=sales.filter(s=>{const sd=new Date(s.created_at);return sd>=monthStart&&sd<=monthEnd;});const real=batchRevenue(mSales); // discount applied in batchRevenue
        const pct=goal>0?(real/goal)*100:0;
        const isCurrent=now.getFullYear()===goalYear&&now.getMonth()===mi;
        return{key,mName,mi,goal,real,pct,isCurrent};
      });
      const totalGoal=rows.reduce((a,r)=>a+r.goal,0);
      const totalReal=rows.reduce((a,r)=>a+r.real,0);
      return(
        <Modal title="📅 Metas & Histórico de Vendas" onClose={()=>setShowGoals(false)} icon="analytics" wide>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"1rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:".5rem"}}>
              <button onClick={()=>setGoalYear(y=>y-1)} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".3rem .65rem",color:"var(--tx)",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:".82rem"}}>‹</button>
              <span style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:"1.05rem",color:"var(--tx)",minWidth:60,textAlign:"center"}}>{goalYear}</span>
              <button onClick={()=>setGoalYear(y=>y+1)} style={{background:"var(--pill)",border:"1px solid var(--bdr2)",borderRadius:".4rem",padding:".3rem .65rem",color:"var(--tx)",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:".82rem"}}>›</button>
            </div>
            <div style={{display:"flex",gap:".75rem",fontSize:".75rem"}}>
              <div style={{textAlign:"right"}}>
                <div style={{color:"var(--tx5)",fontSize:".65rem",textTransform:"uppercase",marginBottom:".08rem"}}>Meta anual</div>
                <div style={{fontWeight:700,color:"#4f5ef0",fontFamily:"'Syne',sans-serif"}}>{fmt(totalGoal)}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{color:"var(--tx5)",fontSize:".65rem",textTransform:"uppercase",marginBottom:".08rem"}}>Realizado</div>
                <div style={{fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(totalReal)}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{color:"var(--tx5)",fontSize:".65rem",textTransform:"uppercase",marginBottom:".08rem"}}>% anual</div>
                <div style={{fontWeight:700,color:totalGoal>0&&totalReal>=totalGoal?"#10b981":"#f59e0b",fontFamily:"'Syne',sans-serif"}}>{totalGoal>0?fmtPct((totalReal/totalGoal)*100):"—"}</div>
              </div>
            </div>
          </div>
          <div style={{display:"grid",gap:".5rem",marginBottom:"1rem"}}>
            {rows.map(r=>(
              <div key={r.key} style={{background:r.isCurrent?"linear-gradient(135deg,#4f5ef010,#10b98108)":"var(--card)",border:"1px solid "+(r.isCurrent?"#4f5ef040":"var(--bdr)"),borderRadius:".65rem",padding:".7rem .85rem"}}>
                <div style={{display:"flex",alignItems:"center",gap:".75rem",marginBottom:".35rem"}}>
                  <span style={{fontWeight:700,fontSize:".82rem",color:r.isCurrent?"#4f5ef0":"var(--tx)",minWidth:30}}>{r.mName}</span>
                  {r.isCurrent&&<span style={{fontSize:".6rem",background:"#4f5ef020",color:"#4f5ef0",borderRadius:"99px",padding:".1rem .4rem",fontWeight:700}}>atual</span>}
                  <div style={{flex:1}}/>
                  <div style={{display:"flex",alignItems:"center",gap:".4rem"}}>
                    <span style={{fontSize:".75rem",color:r.real>=r.goal&&r.goal>0?"#10b981":"var(--tx3)",fontWeight:600}}>{fmt(r.real)}</span>
                    <span style={{fontSize:".68rem",color:"var(--tx5)"}}>de</span>
                    <div style={{fontSize:".85rem",fontWeight:700,color:"#4f5ef0",fontFamily:"'Syne',sans-serif",minWidth:90,textAlign:"right"}}>{r.goal>0?fmt(r.goal):<span style={{color:"var(--tx6)",fontSize:".72rem"}}>—</span>}</div>
                    <span style={{fontSize:".72rem",fontWeight:700,color:r.pct>=100?"#10b981":r.pct>=70?"#f59e0b":r.isPast&&r.goal>0?"#f56565":"var(--tx5)",minWidth:42,textAlign:"right"}}>{r.goal>0?fmtPct(r.pct):"—"}{r.pct>100?" 🚀":r.pct>=100?" 🏆":r.isPast&&r.goal>0&&r.pct<70?" ❌":""}</span>
                  </div>
                </div>
                <div style={{height:6,background:"var(--bdr)",borderRadius:3}}>
                  <div style={{height:"100%",width:Math.min(100,r.pct)+"%",background:r.pct>=100?"linear-gradient(90deg,#10b981,#059669)":r.pct>=70?"linear-gradient(90deg,#f59e0b,#d97706)":"linear-gradient(90deg,#4f5ef0,#8b44f0)",borderRadius:3,transition:"width .5s"}}/>
                </div>
              </div>
            ))}
          </div>
          <div style={{textAlign:"center",fontSize:".7rem",color:"var(--tx5)",display:"flex",alignItems:"center",justifyContent:"center",gap:".4rem"}}><span>Visualização das metas · Para editar acesse</span><button onClick={()=>{setShowGoals(false);setTab("config");setConfigSection("metas");}} style={{background:"none",border:"none",color:"#4f5ef0",fontSize:".72rem",fontFamily:"'DM Sans',sans-serif",fontWeight:700,cursor:"pointer",padding:0}}>⚙️ Config → 🎯 Metas</button></div>
        </Modal>
      );
    })()}


    {/* ══ NOVO PEDIDO ══ */}
    {showOrderModal&&(
      <Modal title="📦 Novo Pedido de Compra" onClose={()=>{setShowOrderModal(false);orderReset();}} icon="pkg" wide>
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#f59e0b",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="info" s={12}/>Estoque só é atualizado ao confirmar o recebimento</div>

        {/* Fornecedor */}
        <Field label="Fornecedor *">
          <select value={orderSupplier.id} onChange={e=>{const s=suppliers.find(x=>x.id===e.target.value);setOrderSupplier({id:e.target.value,name:s?s.name:"Outro"});}} style={IS}>
            <option value="">Selecione o fornecedor...</option>
            {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>

        {/* Itens do pedido */}
        <div style={{marginBottom:"1rem"}}>
          <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".5rem",fontWeight:700}}>🛒 Itens do Pedido</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 70px 100px 70px",gap:".4rem",marginBottom:".3rem",padding:"0 .1rem"}}>
            {["Produto","Qtd","Custo/un",""].map(h=><div key={h} style={{fontSize:".62rem",color:"var(--sub)",textTransform:"uppercase"}}>{h}</div>)}
          </div>
          {orderItems.map((item,idx)=>(
            <div key={item.key} style={{display:"grid",gridTemplateColumns:"1fr 60px 90px 32px",gap:".35rem",marginBottom:".4rem",alignItems:"center"}}>
              <select value={item.product_id} onChange={e=>orderSetProduct(item.key,e.target.value)} style={{...IS,fontSize:".8rem",padding:".45rem .6rem"}}>
                <option value="">Selecione...</option>
                {products.map(p=><option key={p.id} value={p.id}>{(activeCats.find(c=>c.key===p.category)||{icon:"📋"}).icon} {p.name}</option>)}
              </select>
              <input type="number" min="1" onFocus={e=>e.target.select()} value={item.qty} onChange={e=>setOrderItems(items=>items.map(i=>i.key!==item.key?i:{...i,qty:parseInt(e.target.value)||1}))} style={{...IS,fontSize:".82rem",textAlign:"center",padding:".45rem .4rem"}}/>
              <input type="number" min="0" step="0.01" value={item.unit_cost||""} onChange={e=>setOrderItems(items=>items.map(i=>i.key!==item.key?i:{...i,unit_cost:parseFloat(e.target.value)||0}))} placeholder="0,00" style={{...IS,fontSize:".82rem",padding:".45rem .5rem"}}/>
              <button onClick={()=>setOrderItems(items=>items.length>1?items.filter(i=>i.key!==item.key):items)} disabled={orderItems.length===1} style={{background:orderItems.length===1?"transparent":"#1e1010",border:"1px solid "+(orderItems.length===1?"var(--bdr)":"#3a1515"),borderRadius:".4rem",padding:".4rem",color:orderItems.length===1?"var(--bdr2)":"#f56565",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="trash" s={13}/></button>
            </div>
          ))}
          <button onClick={()=>setOrderItems(items=>[...items,newOrderItem()])} style={{display:"flex",alignItems:"center",gap:".35rem",padding:".4rem .85rem",borderRadius:".45rem",border:"1px dashed #f59e0b60",background:"#f59e0b08",color:"#f59e0b",fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,width:"100%",justifyContent:"center",marginTop:".25rem"}}>
            <Ic n="plus" s={13}/>Adicionar produto
          </button>
        </div>

        {/* Pagamento */}
        <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".75rem .85rem",marginBottom:"1rem"}}>
          <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".6rem",fontWeight:700}}>💰 Pagamento</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:".5rem",alignItems:"center",marginBottom:".5rem"}}>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Sinal / Entrada (%)</div>
              <div style={{display:"flex",alignItems:"center",gap:".35rem"}}>
                <input type="number" min="0" max="100" step="5" value={orderPct} onChange={e=>setOrderPct(e.target.value)} style={{...IS,width:80,textAlign:"center"}}/>
                <span style={{fontSize:".8rem",color:"var(--tx4)"}}>%</span>
              </div>
              <div style={{display:"flex",gap:".3rem",marginTop:".35rem",flexWrap:"wrap"}}>
                {[0,25,50,75,100].map(p=>(
                  <button key={p} onClick={()=>setOrderPct(String(p))} style={{padding:".18rem .45rem",borderRadius:".3rem",fontSize:".68rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"1px solid "+(orderPct===String(p)?"#f59e0b":"var(--bdr2)"),background:orderPct===String(p)?"#f59e0b20":"transparent",color:orderPct===String(p)?"#f59e0b":"var(--navoff)"}}>
                    {p}%
                  </button>
                ))}
              </div>
            </div>
            <div>
              {(()=>{
                const total=orderItems.reduce((a,i)=>a+i.unit_cost*i.qty,0);
                const pct=Math.min(100,parseFloat(orderPct)||0);
                const sinal=total*(pct/100);
                const restante=total-sinal;
                return(
                  <div style={{display:"flex",flexDirection:"column",gap:".3rem"}}>
                    {[{l:"Total do pedido",v:fmt(total),c:"var(--tx)"},{l:"Sinal agora ("+pct+"%)",v:fmt(sinal),c:"#f59e0b"},{l:"Restante no recebimento",v:fmt(restante),c:"#f56565"}].map(m=>(
                      <div key={m.l} style={{display:"flex",justifyContent:"space-between",fontSize:".75rem"}}>
                        <span style={{color:"var(--tx5)"}}>{m.l}</span>
                        <span style={{fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</span>
                      </div>
                    ))}
                  </div>
                );
              })()}
            </div>
          </div>
        </div>

        <Inp label="Observações" hint="opcional" placeholder="Ex: Aguardar confirmação, entrega prevista..." value={orderNotes} onChange={e=>setOrderNotes(e.target.value)}/>
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end"}}>
          <Btn v="ghost" onClick={()=>{setShowOrderModal(false);orderReset();}}>Cancelar</Btn>
          <Btn v="warn" onClick={createOrder} disabled={orderItems.every(i=>!i.product_id)||!orderSupplier.id}><Ic n="save" s={13}/>Criar Pedido</Btn>
        </div>
      </Modal>
    )}

    {/* ══ CONFIRMAR RECEBIMENTO ══ */}
    {showReceiveModal&&(()=>{
      const allItems=parseItems(showReceiveModal);
      const isDropshipOrder=allItems.some(i=>i.dropship);
      const pendingItems=allItems.map((it,i)=>({...it,_idx:i})).filter(it=>!it.received);
      const checkedList=Object.entries(receiveChecked)
        .filter(([,v])=>v&&v.checked)
        .map(([k,v])=>{const idx=parseInt(k);const it=allItems[idx];if(!it)return null;const rem=(parseFloat(it.qty)||0)-(parseFloat(it.received_qty)||0);const q=Math.min(rem||1,Math.max(1,parseInt(v.qty)||rem||1));return{index:idx,qty:q,item:it};})
        .filter(Boolean);
      const selectedTotal=checkedList.reduce((a,{item,qty})=>a+(parseFloat(item?.unit_cost||item?.unit_price||0)*qty),0);
      const orderTotal=parseFloat(showReceiveModal.total_value)||0;
      const alreadyPaid=(parseFloat(showReceiveModal.initial_value)||0)+(parseFloat(showReceiveModal.remaining_paid)||0);
      const proportionalPayment=orderTotal>0?Math.round(((selectedTotal/orderTotal)*(parseFloat(showReceiveModal.remaining_value)||0))*100)/100:0;
      const payAmt=receivePayment!==""&&receivePayment!==null?parseFloat(receivePayment)||0:proportionalPayment;
      return(
        <Modal title={"📦 Receber Produtos — "+showReceiveModal.supplier_name} onClose={()=>{setShowReceiveModal(null);setReceiveChecked({});setReceivePayment("");}} icon="arrup" wide>
          {isDropshipOrder
            ?<div style={{background:"#8b44f015",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#8b44f0",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="info" s={12}/>🔄 Pedido de <b>venda direta (dropshipping)</b>: confirmar o recebimento <b>não dá entrada no estoque</b> — serve só para registrar que o fornecedor enviou ao cliente e controlar o pagamento.</div>
            :<div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#10b981",display:"flex",gap:".3rem",alignItems:"center"}}><Ic n="info" s={12}/>Selecione os produtos que chegaram. A entrada no estoque e o pagamento são feitos por item selecionado.</div>
          }

          {/* Lista de itens do pedido */}
          <div style={{marginBottom:"1rem"}}>
            <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:".5rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <span>Selecione os produtos recebidos</span>
              <button onClick={()=>{
                const newChecked={};
                pendingItems.forEach(it=>{const rem=(parseFloat(it.qty)||0)-(parseFloat(it.received_qty)||0);newChecked[it._idx]={checked:true,qty:rem>0?rem:1};});
                setReceiveChecked(newChecked);
              }} style={{fontSize:".65rem",color:"#4f5ef0",background:"#4f5ef015",border:"1px solid #4f5ef030",borderRadius:".3rem",padding:".15rem .5rem",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>Todos</button>
            </div>

            <div style={{display:"grid",gap:".45rem",maxHeight:"45vh",overflowY:"auto",paddingRight:".3rem"}}>
              {allItems.map((it,i)=>{
                const isRec=it.received;
                const ordered=parseFloat(it.qty)||0;
                const recv=parseFloat(it.received_qty)||0;
                const remaining=Math.max(0,ordered-recv);
                const ch=receiveChecked[i]||{checked:false,qty:remaining||1};
                const partial=!isRec&&recv>0;
                return(
                  <div key={i} style={{display:"grid",gridTemplateColumns:"auto 1fr auto auto",gap:".6rem",alignItems:"center",padding:".6rem .85rem",borderRadius:".55rem",border:"1px solid "+(isRec?"#10b98130":partial?"#f59e0b40":ch.checked?"#4f5ef040":"var(--bdr)"),background:isRec?"#10b98108":partial?"#f59e0b08":ch.checked?"#4f5ef008":"transparent",transition:"all .2s"}}>
                    {/* Checkbox */}
                    <button
                      onClick={()=>{if(!isRec)setReceiveChecked(prev=>({...prev,[i]:{...prev[i],checked:!ch.checked,qty:prev[i]?.qty||remaining||1}}));}}
                      disabled={isRec}
                      style={{width:22,height:22,borderRadius:".35rem",border:"2px solid "+(isRec?"#10b98160":ch.checked?"#4f5ef0":"var(--bdr2)"),background:isRec?"#10b98120":ch.checked?"#4f5ef0":"transparent",display:"flex",alignItems:"center",justifyContent:"center",cursor:isRec?"default":"pointer",flexShrink:0,transition:"all .2s"}}
                    >
                      {(isRec||ch.checked)&&<svg width="12" height="12" viewBox="0 0 12 12" fill="none"><polyline points="2,6 5,9 10,3" stroke={isRec?"#10b981":"#fff"} strokeWidth="2" strokeLinecap="round"/></svg>}
                    </button>
                    {/* Produto info */}
                    <div>
                      <div style={{fontSize:".83rem",fontWeight:600,color:isRec?"var(--tx5)":"var(--tx)",textDecoration:isRec?"line-through":""}}>{it.product_name}</div>
                      <div style={{fontSize:".67rem",color:"var(--sub)"}}>Pedido: {ordered} {it.unit} · {fmt(it.unit_cost)}/un{partial?" · já recebido "+recv+", faltam "+remaining:""}</div>
                    </div>
                    {/* Quantidade */}
                    {isRec?(
                      <span style={{fontSize:".75rem",color:"#10b981",fontWeight:700,background:"#10b98115",borderRadius:"99px",padding:".15rem .55rem",border:"1px solid #10b98130",whiteSpace:"nowrap"}}>✅ {recv||ordered} {it.unit}</span>
                    ):(
                      <div style={{display:"flex",alignItems:"center",gap:".3rem"}}>
                        <span style={{fontSize:".68rem",color:"var(--sub)"}}>Recebendo:</span>
                        <input type="number" min="1" max={remaining} onFocus={e=>e.target.select()} value={ch.qty||remaining}
                          onChange={e=>setReceiveChecked(prev=>({...prev,[i]:{...prev[i],qty:Math.min(remaining,Math.max(1,parseInt(e.target.value)||1)),checked:prev[i]?.checked||false}}))}
                          disabled={!ch.checked}
                          style={{width:55,background:"var(--inp)",border:"1px solid "+(ch.checked?"#4f5ef0":"var(--bdr2)"),borderRadius:".35rem",padding:".22rem .4rem",color:"var(--tx)",fontSize:".78rem",fontFamily:"'DM Sans',sans-serif",outline:"none",textAlign:"center",opacity:ch.checked?1:.5}}/>
                        <span style={{fontSize:".68rem",color:"var(--sub)"}}>/ {remaining} {it.unit}</span>
                      </div>
                    )}
                    {/* Valor */}
                    <span style={{fontSize:".8rem",fontWeight:700,color:isRec?"#10b981":ch.checked?"#4f5ef0":"var(--tx5)",fontFamily:"'Syne',sans-serif",minWidth:70,textAlign:"right"}}>
                      {isRec?fmt(it.unit_cost*(recv||ordered)):ch.checked?fmt(it.unit_cost*(ch.qty||remaining)):"—"}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Pagamento */}
          {checkedList.length>0&&(
            <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".75rem .85rem",marginBottom:"1rem"}}>
              <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".05em",marginBottom:".6rem",fontWeight:700}}>💰 Pagamento por este recebimento</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:".5rem",marginBottom:".65rem"}}>
                {[
                  {l:"Valor itens selecionados",v:fmt(selectedTotal),c:"var(--tx)"},
                  {l:"Pagamento proporcional",v:fmt(proportionalPayment),c:"#f59e0b"},
                  {l:"Já pago (sinal+anteriores)",v:fmt(alreadyPaid),c:"#10b981"},
                ].map(m=>(
                  <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".4rem",padding:".5rem .3rem"}}>
                    <div style={{fontSize:".57rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".1rem"}}>{m.l}</div>
                    <div style={{fontSize:".82rem",fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                  </div>
                ))}
              </div>
              <div>
                <div style={{fontSize:".68rem",color:"var(--sub)",marginBottom:".3rem"}}>Valor a pagar agora (R$) <span style={{color:"var(--tx6)"}}>— editável</span></div>
                <input type="number" min="0" step="0.01" value={receivePayment!==""?receivePayment:String(proportionalPayment)}
                  onChange={e=>setReceivePayment(e.target.value)}
                  style={{...IS,fontSize:".9rem",fontWeight:700}}/>
                {receivePayment!==""&&Math.abs(parseFloat(receivePayment)-proportionalPayment)>0.01&&(
                  <div style={{fontSize:".68rem",color:"#f59e0b",marginTop:".28rem",display:"flex",gap:".3rem",alignItems:"center"}}>
                    <Ic n="warn" s={11}/>Diferente do proporcional calculado ({fmt(proportionalPayment)})
                  </div>
                )}
              </div>
            </div>
          )}

          {checkedList.length===0&&pendingItems.length>0&&(
            <div style={{background:"#f59e0b10",border:"1px solid #f59e0b30",borderRadius:".45rem",padding:".55rem .85rem",marginBottom:".85rem",fontSize:".75rem",color:"#f59e0b",display:"flex",gap:".35rem",alignItems:"center"}}>
              <Ic n="warn" s={13}/>Selecione pelo menos 1 produto para registrar o recebimento
            </div>
          )}

          <div style={{display:"flex",gap:".5rem",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{fontSize:".72rem",color:"var(--tx5)"}}>
              {checkedList.length>0?checkedList.length+" produto(s) selecionado(s)":"Nenhum selecionado"}
            </div>
            <div style={{display:"flex",gap:".5rem"}}>
              <Btn v="ghost" onClick={()=>{setShowReceiveModal(null);setReceiveChecked({});setReceivePayment("");}}>Cancelar</Btn>
              <Btn v="ok" onClick={()=>confirmReceive(showReceiveModal,checkedList,receivePayment!==""?receivePayment:String(proportionalPayment))} disabled={checkedList.length===0}>
                <Ic n="save" s={13}/>Confirmar Recebimento
              </Btn>
            </div>
          </div>
        </Modal>
      );
    })()}


    {/* ══ FRETE CONFIG ══ */}
    {showFreteConfig&&localFrete&&(
      <Modal title="⚙️ Configurar Frete" onClose={()=>setShowFreteConfig(false)} icon="delivery" wide>
        <div style={{background:"var(--infobox)",borderRadius:".45rem",padding:".5rem .8rem",marginBottom:".85rem",fontSize:".73rem",color:"#4f5ef0",display:"flex",gap:".3rem",alignItems:"center"}}>
          <Ic n="info" s={12}/>Configurações salvas na nuvem — acessíveis em qualquer dispositivo
        </div>

        {/* ORIGENS */}
        <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx)",marginBottom:".75rem"}}>📍 Pontos de Origem</div>
        {localFrete.origens.map((origem,oi)=>(
          <div key={origem.id} style={{background:"var(--pill)",borderRadius:".6rem",padding:".75rem .85rem",marginBottom:".6rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:".4rem",marginBottom:".5rem"}}>
              <span style={{fontWeight:700,fontSize:".78rem",color:"#4f5ef0",minWidth:60}}>Ponto {oi+1}</span>
              <input
                value={origem.name}
                onChange={e=>setLocalFrete(lf=>{const o=[...lf.origens];o[oi]={...o[oi],name:e.target.value};return{...lf,origens:o};})}
                placeholder={"Nome do ponto (ex: Depósito Centro)"}
                style={{...IS,flex:1,fontSize:".8rem"}}
              />
              {origem.lat&&<span style={{fontSize:".65rem",color:"#10b981",background:"#10b98115",borderRadius:"99px",padding:".1rem .45rem",border:"1px solid #10b98130",whiteSpace:"nowrap"}}>✅ localizado</span>}
            </div>
            <div style={{display:"flex",gap:".4rem",alignItems:"center",flexWrap:"wrap"}}>
              <input
                value={origem.address}
                onChange={e=>setLocalFrete(lf=>{const o=[...lf.origens];o[oi]={...o[oi],address:e.target.value,lat:null,lon:null};return{...lf,origens:o};})}
                placeholder="Endereço completo: Rua, Nº, Bairro, Cidade"
                style={{...IS,flex:1,fontSize:".78rem",minWidth:200}}
              />
              <button
                onClick={async()=>{
                  if(!origem.address.trim()){toast$("Informe o endereço.","#f56565");return;}
                  toast$("🔍 Localizando...","#4f5ef0");
                  try{
                    const results=await geocodeAddr(origem.address);
                    if(results.length===0){toast$("Endereço não encontrado. Tente incluir a cidade.","#f56565");return;}
                    const first=results[0];
                    setLocalFrete(lf=>{const o=[...lf.origens];o[oi]={...o[oi],lat:first.lat,lon:first.lon,address:first.display.split(",").slice(0,3).join(",").trim()};return{...lf,origens:o};});
                    toast$("✅ Ponto "+origem.name+" localizado!");
                  }catch(e){toast$("Erro: "+e.message,"#f56565");}
                }}
                style={{background:"#4f5ef020",border:"1px solid #4f5ef040",borderRadius:".4rem",padding:".38rem .75rem",color:"#4f5ef0",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}
              >🔍 Localizar</button>
              {origem.lat&&<button
                onClick={()=>setLocalFrete(lf=>{const o=[...lf.origens];o[oi]={...o[oi],lat:null,lon:null};return{...lf,origens:o};})}
                style={{background:"none",border:"none",color:"var(--tx6)",cursor:"pointer",padding:".2rem"}}
                title="Remover localização"
              ><Ic n="trash" s={12}/></button>}
            </div>
            {origem.lat&&<div style={{fontSize:".65rem",color:"var(--tx5)",marginTop:".3rem"}}>📌 {origem.lat?.toFixed(5)}, {origem.lon?.toFixed(5)}</div>}
          </div>
        ))}

        {/* VALORES */}
        <div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:".82rem",color:"var(--tx)",marginBottom:".65rem",marginTop:"1rem"}}>💰 Cálculo da Taxa</div>
        <div style={{background:"var(--pill)",borderRadius:".65rem",padding:".85rem",marginBottom:"1rem"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:".65rem",marginBottom:".65rem"}}>
            {[
              {l:"Valor base (R$)",k:"base",hint:"Cobrado em qualquer entrega"},
              {l:"Taxa por km (R$/km)",k:"ratePerKm",hint:"Multiplicado pela distância"},
              {l:"Mínimo total (R$)",k:"minFee",hint:"Piso mínimo da taxa (0 = sem mínimo)"},
              {l:"Máximo total (R$)",k:"maxFee",hint:"Teto máximo da taxa (0 = sem teto)"},
            ].map(f=>(
              <div key={f.k}>
                <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>{f.l}</div>
                <input type="number" min="0" step="0.5"
                  value={localFrete[f.k]||""}
                  onChange={e=>setLocalFrete(lf=>({...lf,[f.k]:e.target.value}))}
                  placeholder="0,00"
                  style={IS}/>
                <div style={{fontSize:".6rem",color:"var(--tx6)",marginTop:".18rem"}}>{f.hint}</div>
              </div>
            ))}
          </div>
          {/* Preview */}
          <div style={{background:"var(--card)",borderRadius:".5rem",padding:".6rem .85rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:".35rem"}}>
            <span style={{fontSize:".73rem",color:"var(--tx4)"}}>💡 Prévia para 10 km:</span>
            <span style={{fontWeight:700,color:"#10b981",fontFamily:"'Syne',sans-serif",fontSize:".88rem"}}>
              {fmt(Math.max(
                parseFloat(localFrete.minFee)||0,
                Math.min(
                  parseFloat(localFrete.maxFee)||999999,
                  (parseFloat(localFrete.base)||0)+(10*(parseFloat(localFrete.ratePerKm)||0))
                )
              ))}
            </span>
          </div>
        </div>

        <div style={{display:"flex",gap:".5rem",justifyContent:"space-between"}}>
          <Btn v="ghost" onClick={()=>setLocalFrete(JSON.parse(JSON.stringify(defaultFreteConfig)))}>↩ Restaurar padrão</Btn>
          <div style={{display:"flex",gap:".5rem"}}>
            <Btn v="ghost" onClick={()=>setShowFreteConfig(false)}>Cancelar</Btn>
            <Btn v="ok" onClick={()=>{saveFreteConfig(localFrete);setShowFreteConfig(false);}}><Ic n="save" s={13}/>Salvar Configurações</Btn>
          </div>
        </div>
      </Modal>
    )}


    {/* ══ CALCULAR FRETE IN-CART ══ */}
    {showFreteCalcInCart&&(
      <Modal title="🗺️ Calcular Frete para Entrega" onClose={()=>{setShowFreteCalcInCart(false);setFreteCalcResult(null);setFreteCalcGeoList([]);}} icon="delivery">
        {/* Seletor de origem */}
        {freteConfig.origens.filter(o=>o.lat).length>1&&(
          <div style={{marginBottom:".75rem"}}>
            <div style={{fontSize:".65rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".35rem",fontWeight:700}}>Ponto de Origem</div>
            <div style={{display:"flex",gap:".35rem",flexWrap:"wrap"}}>
              {freteConfig.origens.filter(o=>o.lat).map(o=>(
                <button key={o.id} onClick={()=>setFreteCalcOrigem(o.id)}
                  style={{padding:".32rem .65rem",borderRadius:".4rem",fontSize:".75rem",fontFamily:"'DM Sans',sans-serif",fontWeight:600,border:"1px solid "+(freteCalcOrigem===o.id?"#4f5ef0":"var(--bdr2)"),background:freteCalcOrigem===o.id?"#4f5ef020":"transparent",color:freteCalcOrigem===o.id?"#4f5ef0":"var(--navoff)"}}>
                  📍 {o.name}
                </button>
              ))}
            </div>
          </div>
        )}
        {freteConfig.origens.filter(o=>o.lat).length===0&&(
          <div style={{background:"#f59e0b10",border:"1px solid #f59e0b30",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".75rem",fontSize:".75rem",color:"#f59e0b"}}>
            ⚠️ Nenhum ponto de origem configurado. Acesse a aba <strong>Frete</strong> para cadastrar.
          </div>
        )}
        {/* Destino */}
        <div style={{marginBottom:".65rem"}}>
          <div style={{fontSize:".65rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".35rem",fontWeight:700}}>Endereço do Cliente</div>
          <div style={{display:"flex",gap:".5rem"}}>
            <input
              value={freteCalcDestino}
              onChange={e=>setFreteCalcDestino(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&calcFreteInCart(freteCalcDestino,freteCalcOrigem)}
              placeholder="Rua, Nº, Bairro, Cidade"
              style={{...IS,flex:1,fontSize:".85rem"}}
              autoFocus
            />
            <Btn onClick={()=>calcFreteInCart(freteCalcDestino,freteCalcOrigem)} disabled={freteCalcLoading||!freteConfig.origens.some(o=>o.lat)}>
              {freteCalcLoading?"⏳":"🗺️"}
            </Btn>
          </div>
        </div>
        {/* Múltiplos endereços */}
        {freteCalcGeoList.length>0&&(
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem",marginBottom:".65rem"}}>
            <div style={{fontSize:".7rem",color:"#f59e0b",marginBottom:".45rem",fontWeight:600}}>⚠️ Selecione o endereço correto:</div>
            {freteCalcGeoList.map((g,i)=>(
              <button key={i} onClick={async()=>{
                setFreteCalcGeoList([]);setFreteCalcLoading(true);
                const origem=freteConfig.origens.find(o=>o.id===freteCalcOrigem);
                if(origem&&origem.lat)await calcFreteInCartWithCoords(origem,g);
                else setFreteCalcLoading(false);
              }} style={{display:"block",width:"100%",textAlign:"left",padding:".45rem .65rem",borderRadius:".4rem",fontSize:".75rem",color:"var(--tx)",background:"transparent",border:"1px solid var(--bdr2)",marginBottom:".3rem",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",lineHeight:1.4}}>
                📍 {g.display}
              </button>
            ))}
          </div>
        )}
        {/* Resultado */}
        {freteCalcResult&&(
          <div style={{background:"linear-gradient(135deg,#10b98115,#4f5ef010)",border:"1px solid #10b98130",borderRadius:".65rem",padding:"1rem",marginBottom:".75rem"}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:".45rem",marginBottom:".65rem"}}>
              {[{l:"Distância",v:freteCalcResult.distKm+" km",c:"#4f5ef0"},{l:"Tempo",v:freteCalcResult.durMin+" min",c:"var(--tx)"},{l:"Taxa calculada",v:fmt(freteCalcResult.fee),c:"#10b981"}].map(m=>(
                <div key={m.l} style={{textAlign:"center",background:"var(--sumbox)",borderRadius:".45rem",padding:".5rem .3rem"}}>
                  <div style={{fontSize:".58rem",color:"var(--sub)",textTransform:"uppercase",marginBottom:".1rem"}}>{m.l}</div>
                  <div style={{fontSize:".88rem",fontWeight:800,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                </div>
              ))}
            </div>
            <div style={{fontSize:".67rem",color:"var(--tx5)",marginBottom:".65rem",lineHeight:1.4}}>
              🏁 {freteCalcResult.origName} → {freteCalcResult.destName.slice(0,70)}{freteCalcResult.destName.length>70?"...":""}
            </div>
            <div style={{background:"#4f5ef010",borderRadius:".4rem",padding:".4rem .65rem",fontSize:".68rem",color:"var(--sub)",marginBottom:".65rem"}}>
              💡 R$ {freteConfig.base} base + ({freteCalcResult.distKm}km × R$ {freteConfig.ratePerKm}/km) = {fmt(freteCalcResult.fee)}
            </div>
            <Btn v="ok" style={{width:"100%",justifyContent:"center"}} onClick={()=>{
              setCartFreight(String(freteCalcResult.fee));
              setCartDelivery(true);
              setShowFreteCalcInCart(false);
              setFreteCalcResult(null);
              setFreteCalcGeoList([]);
              toast$("✅ Taxa de "+fmt(freteCalcResult.fee)+" aplicada ao campo de frete!");
            }}><Ic n="save" s={13}/>✅ Aplicar R$ {fmt(freteCalcResult.fee)} em "Taxa cobrada do cliente"</Btn>
          </div>
        )}
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:freteCalcResult?0:".5rem"}}>
          <Btn v="ghost" onClick={()=>{setShowFreteCalcInCart(false);setFreteCalcResult(null);setFreteCalcGeoList([]);}}>Cancelar</Btn>
        </div>
      </Modal>
    )}


    {/* ══ EDITAR PEDIDO ══ */}
    {editingOrder&&(
      <Modal title={"✏️ Editar Pedido — "+editingOrder.supplier_name} onClose={()=>setEditingOrder(null)} icon="edit">
        <div style={{display:"grid",gap:".65rem"}}>
          <div>
            <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Fornecedor</div>
            <select value={editingOrder.supplier_id||""} onChange={e=>{
              const s=suppliers.find(x=>x.id===e.target.value);
              setEditingOrder(v=>({...v,supplier_id:e.target.value,supplier_name:s?.name||v.supplier_name}));
            }} style={IS}>
              <option value="">Selecione...</option>
              {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <R2>
            <Inp label="Data do pedido" type="date" value={editingOrder.order_date||""} onChange={e=>setEditingOrder(v=>({...v,order_date:e.target.value}))}/>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".28rem"}}>Sinal (%)</div>
              <input type="number" min="0" max="100" step="1" onFocus={e=>e.target.select()} value={editingOrder.initial_pct||""} onChange={e=>{
                const pct=parseFloat(e.target.value)||0;
                const initVal=Math.round(editingOrder.total_value*(pct/100)*100)/100;
                const remVal=Math.max(0,Math.round((editingOrder.total_value-initVal)*100)/100);
                setEditingOrder(v=>({...v,initial_pct:e.target.value,initial_value:initVal,remaining_value:remVal}));
              }} style={IS}/>
            </div>
          </R2>
          <Inp label="Observações" hint="opcional" value={editingOrder.notes||""} onChange={e=>setEditingOrder(v=>({...v,notes:e.target.value}))}/>
          <div style={{background:"var(--pill)",borderRadius:".5rem",padding:".65rem .85rem"}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:".5rem"}}>
              {[{l:"Total",v:fmt(editingOrder.total_value),c:"var(--tx)"},{l:"Sinal",v:fmt(editingOrder.initial_value),c:"#f59e0b"},{l:"Restante",v:fmt(editingOrder.remaining_value),c:"#f56565"}].map(m=>(
                <div key={m.l} style={{textAlign:"center"}}>
                  <div style={{fontSize:".6rem",color:"var(--sub)",textTransform:"uppercase"}}>{m.l}</div>
                  <div style={{fontWeight:700,color:m.c,fontFamily:"'Syne',sans-serif"}}>{m.v}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end",marginTop:"1rem"}}>
          <Btn v="ghost" onClick={()=>setEditingOrder(null)}>Cancelar</Btn>
          <Btn v="ok" onClick={async()=>{
            try{
              await supabase.from("orders").update({
                supplier_id:editingOrder.supplier_id||null,
                supplier_name:editingOrder.supplier_name,
                order_date:editingOrder.order_date,
                initial_pct:parseFloat(editingOrder.initial_pct)||0,
                initial_value:parseFloat(editingOrder.initial_value)||0,
                remaining_value:parseFloat(editingOrder.remaining_value)||0,
                notes:editingOrder.notes||null,
              }).eq("id",editingOrder.id);
              toast$("✅ Pedido atualizado!");
              setEditingOrder(null);
            }catch(e){toast$("Erro: "+e.message,"#f56565");}
          }}><Ic n="save" s={13}/>Salvar</Btn>
        </div>
      </Modal>
    )}

    {/* Logout */}
    {/* ══ RANKING DE CLIENTES ══ */}
    {showRanking&&(()=>{
      const stats={};
      (sales||[]).forEach(s=>{
        const k=s.client_name||"Sem cliente";
        if(!stats[k])stats[k]={name:k,revenue:0,cost:0,count:0,lastDate:null};
        const rev=(parseFloat(s.total_price)||0)-(parseFloat(s.discount)||0);
        stats[k].revenue+=rev;
        const prod=(products||[]).find(p=>p.id===s.product_id||p.name===s.product_name);
        stats[k].cost+=(parseFloat(prod?.cost_per_unit)||0)*(s.quantity||0);
        stats[k].count+=1;
        const d=new Date(s.created_at);
        if(!stats[k].lastDate||d>stats[k].lastDate)stats[k].lastDate=d;
      });
      const ranked=Object.values(stats).map(c=>({...c,profit:c.revenue-c.cost,margin:c.revenue>0?((c.revenue-c.cost)/c.revenue)*100:0})).sort((a,b)=>b.revenue-a.revenue);
      return(
        <Modal title="🏆 Ranking de Clientes" onClose={()=>setShowRanking(false)} icon="dollar" wide>
          <div style={{fontSize:".72rem",color:"var(--sub)",marginBottom:".75rem"}}>Ordenado por receita gerada · {ranked.length} cliente(s)</div>
          <div style={{display:"grid",gap:".5rem"}}>
            {ranked.slice(0,20).map((c,i)=>(
              <div key={c.name} style={{display:"flex",alignItems:"center",gap:".6rem",padding:".6rem .75rem",background:i<3?"linear-gradient(135deg,#f59e0b12,#10b98108)":"var(--card)",border:`1px solid ${i<3?"#f59e0b30":"var(--bdr)"}`,borderRadius:".55rem"}}>
                <span style={{fontSize:"1rem",fontWeight:800,color:i===0?"#f59e0b":i===1?"#94a3b8":i===2?"#b45309":"var(--tx6)",minWidth:28}}>{i===0?"🥇":i===1?"🥈":i===2?"🥉":"#"+(i+1)}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:".82rem",color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
                  <div style={{fontSize:".65rem",color:"var(--tx5)"}}>{c.count} compra(s) · margem {c.margin.toFixed(0)}%</div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontWeight:700,fontSize:".85rem",color:"#10b981",fontFamily:"'Syne',sans-serif"}}>{fmt(c.revenue)}</div>
                  <div style={{fontSize:".62rem",color:"var(--tx5)"}}>lucro {fmt(c.profit)}</div>
                </div>
              </div>
            ))}
            {ranked.length===0&&<div style={{textAlign:"center",padding:"2rem",color:"var(--tx6)"}}>Nenhuma venda registrada ainda.</div>}
          </div>
        </Modal>
      );
    })()}

    {/* ══ FLUXO DE CAIXA PROJETADO ══ */}
    {showCashFlow&&(()=>{
      const hoje=new Date();
      const saldoAtual=(cashTx||[]).reduce((a,t)=>a+(t.type==="entrada"?(parseFloat(t.value)||0):-(parseFloat(t.value)||0)),0);
      const periodos=[{label:"30 dias",dias:30},{label:"60 dias",dias:60},{label:"90 dias",dias:90}];
      const calc=(dias)=>{
        const limite=new Date();limite.setDate(limite.getDate()+dias);
        const aReceber=(receivables||[]).filter(r=>!r.paid&&r.due_date&&new Date(r.due_date)<=limite).reduce((a,r)=>a+(parseFloat(r.value)||0),0);
        const aPagarPedidos=(orders||[]).filter(o=>(o.status==="pendente"||o.status==="parcial")).reduce((a,o)=>a+(parseFloat(o.remaining_value)||0),0);
        const aPagarContas=(payables||[]).filter(p=>!p.paid&&p.due_date&&new Date(p.due_date)<=limite).reduce((a,p)=>a+(parseFloat(p.value)||0),0);
        const aPagar=aPagarPedidos+aPagarContas;
        const recompras=(treatments||[]).filter(t=>t.status==="ativo"&&t.next_purchase&&new Date(t.next_purchase)<=limite).length;
        return{aReceber,aPagar,recompras,projetado:saldoAtual+aReceber-aPagar};
      };
      return(
        <Modal title="💰 Fluxo de Caixa Projetado" onClose={()=>setShowCashFlow(false)} icon="dollar" wide>
          <div style={{background:"var(--pill)",borderRadius:".55rem",padding:".75rem 1rem",marginBottom:"1rem",textAlign:"center"}}>
            <div style={{fontSize:".68rem",color:"var(--sub)",textTransform:"uppercase",letterSpacing:".5px"}}>Saldo atual em caixa</div>
            <div style={{fontSize:"1.5rem",fontWeight:800,color:saldoAtual>=0?"#10b981":"#f56565",fontFamily:"'Syne',sans-serif"}}>{fmt(saldoAtual)}</div>
          </div>
          <div style={{display:"grid",gap:".65rem"}}>
            {periodos.map(p=>{
              const r=calc(p.dias);
              return(
                <div key={p.label} style={{padding:".75rem 1rem",background:"var(--card)",border:"1px solid var(--bdr)",borderRadius:".55rem"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:".5rem"}}>
                    <span style={{fontWeight:700,fontSize:".85rem",color:"#4f5ef0"}}>Próximos {p.label}</span>
                    <span style={{fontWeight:800,fontSize:"1rem",color:r.projetado>=0?"#10b981":"#f56565",fontFamily:"'Syne',sans-serif"}}>{fmt(r.projetado)}</span>
                  </div>
                  <div style={{display:"flex",gap:"1rem",flexWrap:"wrap",fontSize:".7rem"}}>
                    <span style={{color:"#10b981"}}>↗ A receber: {fmt(r.aReceber)}</span>
                    <span style={{color:"#f56565"}}>↘ A pagar: {fmt(r.aPagar)}</span>
                    {r.recompras>0&&<span style={{color:"#8b44f0"}}>🔄 {r.recompras} recompra(s) prevista(s)</span>}
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{fontSize:".68rem",color:"var(--tx6)",marginTop:".75rem",textAlign:"center"}}>Projeção baseada em contas a receber, contas a pagar (boletos), pedidos a fornecedores e recompras de tratamentos ativos.</div>
          <div style={{marginTop:".6rem",textAlign:"center"}}>
            <Btn sm v="ghost" onClick={()=>{setShowCashFlow(false);setTab("financeiro");setFinTab("resumo");}}>📊 Abrir Financeiro completo</Btn>
          </div>
        </Modal>
      );
    })()}

    {/* ══ PROTOCOLO DE TRATAMENTO ══ */}
    {showTreatments&&(
      <Modal title="💊 Protocolos de Tratamento" onClose={()=>setShowTreatments(false)} icon="product" wide>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:".75rem"}}>
          <span style={{fontSize:".72rem",color:"var(--sub)"}}>{(treatments||[]).filter(t=>t.status==="ativo").length} ativo(s)</span>
          <Btn sm onClick={()=>setEditingTreatment({client_name:"",product_name:"",current_dose_mg:"",start_date:new Date().toISOString().slice(0,10),next_purchase:"",frequency_days:"30",status:"ativo",notes:""})}><Ic n="plus" s={12}/>Novo Tratamento</Btn>
        </div>
        <div style={{display:"grid",gap:".55rem"}}>
          {(treatments||[]).map(t=>{
            const daysUntil=t.next_purchase?Math.ceil((new Date(t.next_purchase)-new Date())/864e5):null;
            const urgent=daysUntil!==null&&daysUntil<=5;
            const overdue=daysUntil!==null&&daysUntil<0;
            return(
              <div key={t.id} style={{padding:".7rem .85rem",background:overdue?"#f5656512":urgent?"#f59e0b12":"var(--card)",border:`1px solid ${overdue?"#f5656540":urgent?"#f59e0b40":"var(--bdr)"}`,borderRadius:".55rem"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:".5rem"}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:".85rem",color:"var(--tx)"}}>{t.client_name}</div>
                    <div style={{fontSize:".68rem",color:"var(--tx5)",marginTop:".15rem"}}>
                      {t.product_name||"—"}{t.current_dose_mg?` · ${t.current_dose_mg}mg`:""} · a cada {t.frequency_days}d
                    </div>
                    {t.next_purchase&&(
                      <div style={{fontSize:".7rem",marginTop:".3rem",fontWeight:600,color:overdue?"#f56565":urgent?"#f59e0b":"#10b981"}}>
                        {overdue?`⚠️ Atrasado ${Math.abs(daysUntil)}d`:daysUntil===0?"🔔 Recompra HOJE":`📅 Próxima em ${daysUntil}d (${new Date(t.next_purchase).toLocaleDateString("pt-BR")})`}
                      </div>
                    )}
                  </div>
                  <div style={{display:"flex",gap:".25rem"}}>
                    <button onClick={()=>advanceTreatment(t)} title="Comprou - reagendar" style={{background:"#10b98120",border:"none",borderRadius:".35rem",padding:".3rem .5rem",color:"#10b981",cursor:"pointer",fontSize:".7rem",fontWeight:600}}>✓ Comprou</button>
                    <button onClick={()=>setEditingTreatment(t)} style={{background:"none",border:"none",color:"#4f5ef0",cursor:"pointer",padding:".3rem"}}><Ic n="edit" s={13}/></button>
                    {canDelete&&<button onClick={()=>delTreatment(t.id)} style={{background:"none",border:"none",color:"var(--tx6)",cursor:"pointer",padding:".3rem"}}><Ic n="trash" s={13}/></button>}
                  </div>
                </div>
              </div>
            );
          })}
          {(treatments||[]).length===0&&<div style={{textAlign:"center",padding:"2rem",color:"var(--tx6)"}}>Nenhum tratamento cadastrado. Crie o primeiro!</div>}
        </div>
      </Modal>
    )}

    {/* ══ EDITAR/NOVO TRATAMENTO ══ */}
    {editingTreatment&&(
      <Modal title={editingTreatment.id?"✏️ Editar Tratamento":"💊 Novo Tratamento"} onClose={()=>setEditingTreatment(null)} icon="product">
        <div style={{display:"grid",gap:".6rem"}}>
          <div>
            <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Cliente *</div>
            <input value={editingTreatment.client_name} onChange={e=>{const c=clients.find(x=>x.name===e.target.value);setEditingTreatment(v=>({...v,client_name:e.target.value,client_id:c?.id||v.client_id}));}} list="treat-clients" placeholder="Nome do cliente" style={IS}/>
            <datalist id="treat-clients">{(clients||[]).map(c=><option key={c.id} value={c.name}/>)}</datalist>
          </div>
          <R2>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Produto</div>
              <input value={editingTreatment.product_name||""} onChange={e=>setEditingTreatment(v=>({...v,product_name:e.target.value}))} list="treat-prods" placeholder="Tirzepatida..." style={IS}/>
              <datalist id="treat-prods">{(products||[]).map(p=><option key={p.id} value={p.name}/>)}</datalist>
            </div>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Dose atual (mg)</div>
              <input type="number" step="0.1" value={editingTreatment.current_dose_mg||""} onChange={e=>setEditingTreatment(v=>({...v,current_dose_mg:e.target.value}))} placeholder="2.5" style={IS}/>
            </div>
          </R2>
          <R2>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Próxima compra</div>
              <input type="date" value={editingTreatment.next_purchase||""} onChange={e=>setEditingTreatment(v=>({...v,next_purchase:e.target.value}))} style={IS}/>
            </div>
            <div>
              <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Frequência (dias)</div>
              <input type="number" value={editingTreatment.frequency_days||"30"} onChange={e=>setEditingTreatment(v=>({...v,frequency_days:e.target.value}))} placeholder="30" style={IS}/>
            </div>
          </R2>
          <div>
            <div style={{fontSize:".65rem",color:"var(--sub)",marginBottom:".25rem"}}>Status</div>
            <select value={editingTreatment.status||"ativo"} onChange={e=>setEditingTreatment(v=>({...v,status:e.target.value}))} style={IS}>
              <option value="ativo">Ativo</option>
              <option value="pausado">Pausado</option>
              <option value="concluido">Concluído</option>
            </select>
          </div>
          <Inp label="Observações" hint="opcional" value={editingTreatment.notes||""} onChange={e=>setEditingTreatment(v=>({...v,notes:e.target.value}))}/>
        </div>
        <div style={{display:"flex",gap:".5rem",justifyContent:"flex-end",marginTop:"1rem"}}>
          <Btn v="ghost" onClick={()=>setEditingTreatment(null)}>Cancelar</Btn>
          <Btn v="ok" onClick={()=>{if(!editingTreatment.client_name){toast$("Informe o cliente.","#f56565");return;}saveTreatment(editingTreatment);}}><Ic n="save" s={13}/>Salvar</Btn>
        </div>
      </Modal>
    )}

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
  </></ErrorBoundary>
  );
}
