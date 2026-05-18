import { useState, useEffect, useRef } from "react";
import { db } from "./firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, writeBatch } from "firebase/firestore";
import * as XLSX from "xlsx";

const USERS = [
  { id:"sofi",  name:"Sofi",  avatar:"S", color:"#8b5cf6" },
  { id:"andre", name:"André", avatar:"A", color:"#6366f1" },
  { id:"axel",  name:"Axel",  avatar:"X", color:"#0ea5e9" },
  { id:"joaco", name:"Joaco", avatar:"J", color:"#10b981" },
];

const PALETTE = {
  indigo: { bg:"#eef2ff", fg:"#4338ca", border:"#c7d2fe" },
  amber:  { bg:"#fffbeb", fg:"#b45309", border:"#fcd34d" },
  green:  { bg:"#f0fdf4", fg:"#166534", border:"#86efac" },
  red:    { bg:"#fef2f2", fg:"#b91c1c", border:"#fca5a5" },
  sky:    { bg:"#f0f9ff", fg:"#0369a1", border:"#7dd3fc" },
  purple: { bg:"#faf5ff", fg:"#7e22ce", border:"#d8b4fe" },
  rose:   { bg:"#fff1f2", fg:"#be123c", border:"#fda4af" },
  teal:   { bg:"#f0fdfa", fg:"#115e59", border:"#5eead4" },
  orange: { bg:"#fff7ed", fg:"#c2410c", border:"#fdba74" },
  gray:   { bg:"#f9fafb", fg:"#374151", border:"#d1d5db" },
};
const LICEO_PAL = ["indigo","amber","green","red","sky","purple","rose","teal","orange","gray"];
const TYPE_PAL  = { contacto:"sky", derivacion:"purple", rastreo:"amber" };
const PRIORITY_META = {
  alta:  { label:"Alta",  bg:"#fef2f2", fg:"#b91c1c", border:"#fca5a5", icon:"ti-alert-triangle" },
  media: { label:"Media", bg:"#fffbeb", fg:"#b45309", border:"#fcd34d", icon:"ti-minus" },
  baja:  { label:"Baja",  bg:"#f0f9ff", fg:"#0369a1", border:"#7dd3fc", icon:"ti-arrow-down" },
};

const LICEOS_DEFAULT = ["Aleph","Areteia","Aletheia","San Felipe","Crandon","Elbio Fernandez","Santa Elena","Alternativo","Ánima","IEP","London","Gabriela Mistral","Eduschool"];
const ITEM_TYPES = { contacto:"Contacto al alumno", derivacion:"Derivación", rastreo:"Rastreo de información" };
const TYPE_ICON  = { contacto:"ti-phone", derivacion:"ti-corner-up-right", rastreo:"ti-search" };

const genId = () => Date.now().toString(36) + Math.random().toString(36).slice(2);
const today = () => new Date().toISOString().slice(0,10);
const fmt   = d => d ? new Date(d+"T12:00:00").toLocaleDateString("es-UY",{day:"2-digit",month:"short",year:"numeric"}) : "—";
const fmtS  = d => d ? new Date(d+"T12:00:00").toLocaleDateString("es-UY",{day:"2-digit",month:"short"}) : "—";
const liceoPal = liceo => { const i=[...(liceo||"")].reduce((a,c)=>a+c.charCodeAt(0),0); return PALETTE[LICEO_PAL[i%LICEO_PAL.length]]; };
const typePal  = type  => PALETTE[TYPE_PAL[type]||"gray"];

const uCol  = (uid,col)    => collection(db,`users/${uid}/${col}`);
const uDoc  = (uid,col,id) => doc(db,`users/${uid}/${col}/${id}`);
const uCfg  = uid          => doc(db,`users/${uid}/config/main`);
const fbSet    = async (uid,col,id,data) => { try { await setDoc(uDoc(uid,col,id),data); } catch(e){console.error(e);} };
const fbSetCfg = async (uid,data)        => { try { await setDoc(uCfg(uid),data); }       catch(e){console.error(e);} };

/* ── Shared students collection (all users) ── */
const studentsCol = () => collection(db,"students");
const studentDoc  = id => doc(db,`students/${id}`);
const fbSetStudent = async (id,data) => { try { await setDoc(studentDoc(id),data); } catch(e){console.error(e);} };
const fbDelStudent = async id => { try { await deleteDoc(studentDoc(id)); } catch(e){console.error(e);} };

async function compressImage(file) {
  return new Promise(resolve => {
    const r = new FileReader();
    r.onload = e => {
      const img = new Image();
      img.onload = () => {
        const MAX=700, s=Math.min(1,MAX/Math.max(img.width,img.height));
        const c=document.createElement("canvas");
        c.width=Math.round(img.width*s); c.height=Math.round(img.height*s);
        c.getContext("2d").drawImage(img,0,0,c.width,c.height);
        resolve(c.toDataURL("image/jpeg",0.68));
      };
      img.src=e.target.result;
    };
    r.readAsDataURL(file);
  });
}

function Badge({label,palette,icon,size="md"}) {
  const p=palette||PALETTE.gray, pad=size==="sm"?"1px 7px":"3px 10px";
  return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:p.bg,color:p.fg,border:`1px solid ${p.border}`,fontSize:11,fontWeight:600,padding:pad,borderRadius:20,whiteSpace:"nowrap"}}>
    {icon&&<i className={`ti ${icon}`} style={{fontSize:11}}/>}{label}
  </span>;
}
const TypeBadge     = ({type,size="md"})     => <Badge label={ITEM_TYPES[type]||type} palette={typePal(type)} icon={TYPE_ICON[type]||"ti-tag"} size={size}/>;
const LiceoChip     = ({liceo,size="md"})    => <Badge label={liceo} palette={liceoPal(liceo)} size={size}/>;
const PriorityBadge = ({priority,size="md"}) => { const m=PRIORITY_META[priority]||PRIORITY_META.media; return <Badge label={m.label} palette={{bg:m.bg,fg:m.fg,border:m.border}} icon={m.icon} size={size}/>; };

function StatCard({label,value,icon,palette,onClick}) {
  const p=palette||PALETTE.gray;
  return <div onClick={onClick} style={{background:p.bg,border:`1px solid ${p.border}`,borderRadius:12,padding:"16px 18px",cursor:onClick?"pointer":"default"}}>
    <div style={{fontSize:12,color:p.fg,fontWeight:500,marginBottom:8,display:"flex",alignItems:"center",gap:5}}><i className={`ti ${icon}`} style={{fontSize:14}}/>{label}</div>
    <div style={{fontSize:28,fontWeight:700,color:p.fg}}>{value}</div>
  </div>;
}

function Lightbox({src,onClose}) {
  useEffect(()=>{ const esc=e=>{if(e.key==="Escape")onClose();}; document.addEventListener("keydown",esc); return ()=>document.removeEventListener("keydown",esc); },[onClose]);
  return <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <button onClick={onClose} style={{position:"absolute",top:16,right:16,width:36,height:36,borderRadius:"50%",background:"rgba(255,255,255,.15)",border:"none",cursor:"pointer",color:"#fff",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
    <img src={src} alt="" onClick={e=>e.stopPropagation()} style={{maxWidth:"100%",maxHeight:"90vh",borderRadius:10,objectFit:"contain"}}/>
  </div>;
}

function PhotoPicker({photos=[],onChange,compact=false}) {
  const ref=useRef(); const [lb,setLb]=useState(null); const sz=compact?52:68;
  async function handleFiles(e) {
    const files=Array.from(e.target.files); if(!files.length)return;
    const compressed=await Promise.all(files.map(compressImage));
    onChange([...photos,...compressed.map((d,i)=>({id:genId(),dataUrl:d,name:files[i].name}))]);
    e.target.value="";
  }
  return <>
    {lb&&<Lightbox src={lb} onClose={()=>setLb(null)}/>}
    <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
      {photos.map(p=><div key={p.id} style={{position:"relative",width:sz,height:sz,flexShrink:0}}>
        <img src={p.dataUrl} alt={p.name} onClick={()=>setLb(p.dataUrl)} style={{width:"100%",height:"100%",objectFit:"cover",borderRadius:8,cursor:"zoom-in",border:"1px solid #e5e7eb"}}/>
        <button onClick={()=>onChange(photos.filter(x=>x.id!==p.id))} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:"50%",background:"#111",color:"#fff",border:"none",cursor:"pointer",fontSize:9,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
      </div>)}
      <label style={{width:sz,height:sz,border:"1.5px dashed #d1d5db",borderRadius:8,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",color:"#9ca3af",gap:2,flexShrink:0,background:"#f9fafb"}}>
        <i className="ti ti-camera-plus" style={{fontSize:compact?16:20}}/>{!compact&&<span style={{fontSize:10}}>Foto</span>}
        <input ref={ref} type="file" accept="image/*" multiple onChange={handleFiles} style={{display:"none"}}/>
      </label>
    </div>
  </>;
}
function PhotoStrip({photos=[]}) {
  const [lb,setLb]=useState(null);
  if(!photos?.length)return null;
  return <>
    {lb&&<Lightbox src={lb} onClose={()=>setLb(null)}/>}
    <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:8}}>
      {photos.map(p=><img key={p.id} src={p.dataUrl} alt={p.name} onClick={()=>setLb(p.dataUrl)} style={{width:56,height:56,objectFit:"cover",borderRadius:8,cursor:"zoom-in",border:"1px solid #e5e7eb"}}/>)}
    </div>
  </>;
}

function BulletNotes({bullets,onChange,placeholder="Agregar punto…"}) {
  const refs=useRef([]);
  const update=(i,v)=>{const n=[...bullets];n[i]=v;onChange(n);};
  const addAfter=i=>{const n=[...bullets];n.splice(i+1,0,"");onChange(n);setTimeout(()=>refs.current[i+1]?.focus(),30);};
  const removeAt=i=>{if(bullets.length===1){onChange([""]);return;}const n=[...bullets];n.splice(i,1);onChange(n);setTimeout(()=>refs.current[Math.max(0,i-1)]?.focus(),30);};
  const onKey=(e,i)=>{if(e.key==="Enter"){e.preventDefault();addAfter(i);}if(e.key==="Backspace"&&bullets[i]===""&&bullets.length>1){e.preventDefault();removeAt(i);}};
  return <div style={{background:"#f9fafb",borderRadius:8,border:"1px solid #e5e7eb",padding:"10px 12px"}}>
    {bullets.map((b,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:i<bullets.length-1?6:0}}>
      <span style={{color:"#8b5cf6",fontSize:16,userSelect:"none",flexShrink:0}}>•</span>
      <input ref={el=>refs.current[i]=el} value={b} onChange={e=>update(i,e.target.value)} onKeyDown={e=>onKey(e,i)}
        placeholder={i===0?placeholder:""} style={{flex:1,background:"transparent",border:"none",outline:"none",fontSize:13,color:"#111827",fontFamily:"inherit",padding:"2px 0"}}/>
      {bullets.length>1&&<button onClick={()=>removeAt(i)} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:12,padding:"0 2px",flexShrink:0}}>✕</button>}
    </div>)}
    <button onClick={()=>addAfter(bullets.length-1)} style={{marginTop:8,fontSize:11,color:"#8b5cf6",background:"none",border:"none",cursor:"pointer",padding:"2px 0",display:"flex",alignItems:"center",gap:4,fontWeight:500}}>
      <i className="ti ti-plus" style={{fontSize:11}}/>Agregar punto
    </button>
  </div>;
}
function BulletDisplay({bullets}) {
  const items=(bullets||[]).filter(b=>b.trim());
  if(!items.length)return null;
  return <ul style={{margin:"6px 0 0",padding:"0 0 0 14px",listStyleType:"none"}}>
    {items.map((b,i)=><li key={i} style={{fontSize:13,color:"#4b5563",lineHeight:1.6,display:"flex",gap:8,alignItems:"baseline"}}>
      <span style={{color:"#8b5cf6",fontSize:14,flexShrink:0}}>•</span>{b}
    </li>)}
  </ul>;
}

function VisitCommentThread({visit,onAddComment,onDeleteComment,onResolve}) {
  const [text,setText]=useState("");
  const comments=visit.comments||[];
  const submit=()=>{if(!text.trim())return;onAddComment(visit.id,text);setText("");};
  return <div style={{borderTop:"1px solid #f3f4f6",paddingTop:10,marginTop:10}}>
    {comments.length>0&&<div style={{marginBottom:10}}>
      {comments.map(c=><div key={c.id} style={{display:"flex",gap:8,marginBottom:6,alignItems:"flex-start"}}>
        <div style={{width:3,flexShrink:0,background:"#ede9fe",borderRadius:2,marginTop:2}}/>
        <div style={{flex:1}}>
          <span style={{fontSize:11,color:"#9ca3af",marginRight:8}}>{fmt(c.date)}</span>
          <span style={{fontSize:13,color:"#374151"}}>{c.text}</span>
        </div>
        {onDeleteComment&&<button onClick={()=>onDeleteComment(visit.id,c.id)} title="Eliminar"
          style={{background:"none",border:"none",cursor:"pointer",color:"#d1d5db",fontSize:12,padding:"0 2px",flexShrink:0,lineHeight:1}}>✕</button>}
      </div>)}
    </div>}
    <div style={{display:"flex",gap:8,marginBottom:8}}>
      <input value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>e.key==="Enter"&&submit()}
        placeholder="Agregar comentario a las notas…"
        style={{flex:1,fontSize:13,padding:"6px 10px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontFamily:"inherit"}}/>
      <button onClick={submit} disabled={!text.trim()}
        style={{padding:"6px 12px",fontSize:12,fontWeight:600,background:text.trim()?"#8b5cf6":"#f3f4f6",color:text.trim()?"#fff":"#9ca3af",border:"none",borderRadius:8,cursor:text.trim()?"pointer":"not-allowed"}}>Enviar</button>
    </div>
    {!visit.notesResolved&&<button onClick={()=>onResolve(visit.id)}
      style={{width:"100%",padding:"7px",fontSize:12,fontWeight:600,background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:8,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
      <i className="ti ti-check" style={{fontSize:13}}/>Marcar notas como resueltas
    </button>}
  </div>;
}

function CommentThread({item,onAddComment,onDeleteComment,onResolve}) {
  const [text,setText]=useState("");
  const comments=item.comments||[];
  const submit=()=>{if(!text.trim())return;onAddComment(item.id,text);setText("");};
  return <div style={{borderTop:"1px solid #f3f4f6",paddingTop:12,marginTop:12}}>
    {comments.length>0&&<div style={{marginBottom:12}}>
      {comments.map(c=><div key={c.id} style={{display:"flex",gap:10,marginBottom:8,alignItems:"flex-start"}}>
        <div style={{width:3,flexShrink:0,background:"#ede9fe",borderRadius:2,marginTop:3}}/>
        <div style={{flex:1}}>
          <div style={{fontSize:11,color:"#9ca3af",marginBottom:2}}>{fmt(c.date)}</div>
          <div style={{fontSize:13,color:"#374151",lineHeight:1.55}}>{c.text}</div>
        </div>
        {onDeleteComment&&<button onClick={()=>onDeleteComment(item.id,c.id)} title="Eliminar"
          style={{background:"none",border:"none",cursor:"pointer",color:"#d1d5db",fontSize:12,padding:"2px",flexShrink:0}}>✕</button>}
      </div>)}
    </div>}
    <div style={{display:"flex",gap:8}}>
      <input value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>e.key==="Enter"&&submit()}
        placeholder="Agregar seguimiento o comentario…"
        style={{flex:1,fontSize:13,padding:"7px 11px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontFamily:"inherit"}}/>
      <button onClick={submit} disabled={!text.trim()}
        style={{padding:"7px 14px",fontSize:12,fontWeight:600,background:text.trim()?"#8b5cf6":"#f3f4f6",color:text.trim()?"#fff":"#9ca3af",border:"none",borderRadius:8,cursor:text.trim()?"pointer":"not-allowed"}}>Enviar</button>
    </div>
    {item.status==="abierto"&&onResolve&&<button onClick={()=>onResolve(item.id)}
      style={{marginTop:10,width:"100%",padding:"8px",fontSize:12,fontWeight:600,background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:8,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
      <i className="ti ti-check" style={{fontSize:13}}/>Marcar como resuelto
    </button>}
  </div>;
}

function ItemCard({item,studentName,visitDate,onResolve,onAddComment,onDeleteComment,onEditItem,onDelete,showStudent=true}) {
  const displayStudent=item.studentName||studentName;
  const [expanded,setExpanded]=useState(false);
  const [editing,setEditing]=useState(false);
  const [editDesc,setEditDesc]=useState(item.description);
  const comments=item.comments||[]; const p=typePal(item.type);
  const isOverdue=item.alertDate&&item.alertDate<=today()&&item.status==="abierto";
  const saveEdit=()=>{onEditItem(item.id,editDesc);setEditing(false);};
  return <div style={{background:"#fff",border:`1px solid ${isOverdue?"#fca5a5":item.status==="resuelto"?"#f3f4f6":p.border}`,borderRadius:12,padding:"14px 16px",marginBottom:8,opacity:item.status==="resuelto"?0.65:1,borderLeft:`3px solid ${isOverdue?"#ef4444":item.status==="resuelto"?"#e5e7eb":p.fg}`}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
      <div style={{display:"flex",flexWrap:"wrap",gap:6,alignItems:"center"}}>
        <TypeBadge type={item.type}/>
        {item.liceo&&<LiceoChip liceo={item.liceo}/>}
        {item.priority&&item.priority!=="media"&&<PriorityBadge priority={item.priority}/>}
        {item.alertDate&&item.status==="abierto"&&<span style={{fontSize:11,color:isOverdue?"#b91c1c":"#6b7280",background:isOverdue?"#fef2f2":"#f9fafb",padding:"1px 8px",borderRadius:8,border:`1px solid ${isOverdue?"#fca5a5":"#e5e7eb"}`,display:"flex",alignItems:"center",gap:3}}>
          <i className="ti ti-bell" style={{fontSize:11}}/>{isOverdue?"Vencida":fmt(item.alertDate)}
        </span>}
        {item.assignedBy&&<Badge label={`Asignado por ${item.assignedBy}`} palette={PALETTE.purple} icon="ti-user-share" size="sm"/>}
        {item.status==="resuelto"&&<Badge label="Resuelto" palette={PALETTE.green} icon="ti-check" size="sm"/>}
      </div>
      <span style={{fontSize:11,color:"#9ca3af",whiteSpace:"nowrap",flexShrink:0}}>{fmtS(visitDate)}</span>
    </div>
    {editing?<div style={{marginBottom:10}}>
      <input value={editDesc} onChange={e=>setEditDesc(e.target.value)}
        onKeyDown={e=>{if(e.key==="Enter")saveEdit();if(e.key==="Escape")setEditing(false);}}
        autoFocus style={{width:"100%",fontSize:14,padding:"7px 11px",borderRadius:8,border:"2px solid #8b5cf6",background:"#f9fafb",color:"#111827",fontFamily:"inherit",boxSizing:"border-box"}}/>
      <div style={{display:"flex",gap:6,marginTop:6,justifyContent:"flex-end"}}>
        <button onClick={()=>setEditing(false)} style={{fontSize:11,padding:"4px 12px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:6,cursor:"pointer"}}>Cancelar</button>
        <button onClick={saveEdit} style={{fontSize:11,padding:"4px 12px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:6,cursor:"pointer",fontWeight:600}}>✓ Guardar</button>
      </div>
    </div>:<div style={{display:"flex",alignItems:"flex-start",gap:6,marginBottom:6}}>
      <p style={{margin:0,fontSize:14,color:"#111827",lineHeight:1.55,flex:1}}>{item.description}</p>
      {item.status==="abierto"&&onEditItem&&<button onClick={()=>{setEditDesc(item.description);setEditing(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",padding:"2px",flexShrink:0}}>
        <i className="ti ti-pencil" style={{fontSize:13}}/>
      </button>}
      {onDelete&&<button onClick={()=>onDelete(item.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#fca5a5",padding:"2px",flexShrink:0}}>
        <i className="ti ti-trash" style={{fontSize:13}}/>
      </button>}
    </div>}
    <PhotoStrip photos={item.photos}/>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}>
      <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
        {showStudent&&displayStudent&&<span style={{fontSize:12,color:"#6b7280",display:"flex",alignItems:"center",gap:4}}><i className="ti ti-user" style={{fontSize:12}}/>{displayStudent}</span>}
        <span style={{fontSize:12,color:"#9ca3af",display:"flex",alignItems:"center",gap:4}}><i className="ti ti-calendar" style={{fontSize:12}}/>{fmt(visitDate)}</span>
      </div>
      <button onClick={()=>setExpanded(e=>!e)} style={{display:"flex",alignItems:"center",gap:4,fontSize:12,background:"none",border:"none",cursor:"pointer",color:"#8b5cf6",padding:0,fontWeight:500}}>
        <i className="ti ti-message-circle" style={{fontSize:13}}/>
        {comments.length>0?`${comments.length} comentario${comments.length>1?"s":""}` : "Comentar"}
        <i className={`ti ${expanded?"ti-chevron-up":"ti-chevron-down"}`} style={{fontSize:12}}/>
      </button>
    </div>
    {expanded&&<CommentThread item={item} onAddComment={onAddComment} onDeleteComment={onDeleteComment} onResolve={onResolve}/>}
  </div>;
}

function NavBtn({icon,label,active,badge,onClick,sub,color}) {
  return <button onClick={onClick} style={{display:"flex",alignItems:"center",gap:9,width:"100%",padding:sub?"7px 16px 7px 34px":"9px 16px",background:active?"#f5f3ff":"transparent",border:"none",borderLeft:active?`3px solid ${color||"#8b5cf6"}`:"3px solid transparent",cursor:"pointer",textAlign:"left",color:active?"#6d28d9":"#6b7280",fontSize:sub?13:14,fontWeight:active?600:400}}>
    <i className={`ti ${icon}`} style={{fontSize:sub?14:16,color:active?(color||"#8b5cf6"):"inherit"}}/>
    <span style={{flex:1}}>{label}</span>
    {badge>0&&<span style={{fontSize:11,background:"#fef3c7",color:"#92400e",padding:"1px 8px",borderRadius:10,fontWeight:700,border:"1px solid #fcd34d"}}>{badge}</span>}
  </button>;
}

function NewStudentModal({onConfirm,onCancel}) {
  const [nombre,setNombre]=useState(""); const [ci,setCi]=useState(""); const [liceo,setLiceo]=useState(""); const [cohorte,setCohorte]=useState(""); const ok=nombre.trim().length>1;
  const inp={width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #e5e7eb",fontSize:14,fontFamily:"inherit",color:"#111827",boxSizing:"border-box",outline:"none"};
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,padding:"24px",width:"100%",maxWidth:420,boxShadow:"0 8px 40px rgba(0,0,0,.18)"}}>
      <div style={{fontSize:16,fontWeight:700,color:"#111827",marginBottom:16}}>Nuevo alumno</div>
      <div style={{marginBottom:12}}>
        <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Nombre completo *</label>
        <input autoFocus value={nombre} onChange={e=>setNombre(e.target.value)} placeholder="Ej: Juan Pérez" style={inp}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <div>
          <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>CI (opcional)</label>
          <input value={ci} onChange={e=>setCi(e.target.value)} placeholder="Ej: 5.234.567-8" style={inp}/>
        </div>
        <div>
          <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Cohorte (opcional)</label>
          <input value={cohorte} onChange={e=>setCohorte(e.target.value)} placeholder="Ej: 2024" style={inp}/>
        </div>
      </div>
      <div style={{marginBottom:20}}>
        <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Liceo (opcional)</label>
        <input value={liceo} onChange={e=>setLiceo(e.target.value)} placeholder="Ej: Crandon" style={inp}/>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={onCancel} style={{flex:1,padding:"9px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:9,cursor:"pointer",fontSize:14}}>Cancelar</button>
        <button onClick={()=>ok&&onConfirm(nombre.trim(),ci.trim(),liceo.trim(),cohorte.trim())} disabled={!ok}
          style={{flex:2,padding:"9px",background:ok?"#8b5cf6":"#f3f4f6",color:ok?"#fff":"#9ca3af",border:"none",borderRadius:9,cursor:ok?"pointer":"not-allowed",fontSize:14,fontWeight:700}}>Agregar estudiante</button>
      </div>
    </div>
  </div>;
}

function BulkImportModal({onConfirm,onCancel}) {
  const [preview,setPreview]=useState([]);
  const [error,setError]=useState("");
  const [fileName,setFileName]=useState("");
  const [loading,setLoading]=useState(false);

  async function handleFile(e) {
    const file=e.target.files[0]; if(!file)return;
    setFileName(file.name); setError(""); setPreview([]); setLoading(true);
    try {
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
      if(!rows.length){setError("El archivo está vacío.");setLoading(false);return;}

      // Normalize headers: find columns by partial/case-insensitive match
      const norm=s=>String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
      const sampleRow=rows[0];
      const headers=Object.keys(sampleRow);
      const findCol=(keywords)=>headers.find(h=>keywords.some(k=>norm(h).includes(k)))||null;

      const colNombre =findCol(["nombre completo","nombre","name"]);
      const colLiceo  =findCol(["liceo","institucion","colegio","centro"]);
      const colCI     =findCol(["cedula","ci","documento","doc"]);
      const colCohorte=findCol(["cohorte","promocion","año","año de ingreso","generacion"]);

      if(!colNombre){setError("No se encontró la columna 'Nombre completo'. Verificá los encabezados.");setLoading(false);return;}

      const results=[]; const errs=[];
      rows.forEach((row,i)=>{
        const nombre  =String(row[colNombre]  ||"").trim();
        const liceo   =colLiceo  ?String(row[colLiceo]  ||"").trim():"";
        const ci      =colCI     ?String(row[colCI]     ||"").trim():"";
        const cohorte =colCohorte?String(row[colCohorte]||"").trim():"";
        if(!nombre||nombre.length<2){errs.push(`Fila ${i+2}: nombre vacío o muy corto`);return;}
        results.push({nombre,liceo,ci,cohorte});
      });
      if(errs.length&&!results.length){setError(errs.join(" · "));setLoading(false);return;}
      if(errs.length) setError(`${errs.length} fila${errs.length>1?"s":""} omitida${errs.length>1?"s":""}: ${errs.slice(0,3).join(" · ")}${errs.length>3?"…":""}`);
      setPreview(results);
    } catch(err) {
      setError("No se pudo leer el archivo. Asegurate de que sea un .xlsx o .xls válido.");
      console.error(err);
    }
    setLoading(false);
    e.target.value="";
  }

  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,padding:"24px",width:"100%",maxWidth:540,boxShadow:"0 8px 40px rgba(0,0,0,.18)",maxHeight:"90vh",display:"flex",flexDirection:"column",gap:14}}>
      <div>
        <div style={{fontSize:16,fontWeight:700,color:"#111827",marginBottom:4}}>Carga masiva desde Excel</div>
        <div style={{fontSize:13,color:"#6b7280"}}>El archivo debe tener estas columnas (en cualquier orden):</div>
        <div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>
          {["Nombre completo","Nombre de Liceo","Cédula","Cohorte"].map(col=><span key={col} style={{fontSize:12,background:"#f5f3ff",color:"#6d28d9",padding:"2px 10px",borderRadius:10,border:"1px solid #ddd6fe",fontWeight:500}}>{col}</span>)}
        </div>
        <div style={{fontSize:12,color:"#9ca3af",marginTop:6}}>Liceo, Cédula y Cohorte son opcionales.</div>
      </div>

      {/* File drop zone */}
      <label style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:8,padding:"28px 16px",border:"2px dashed #ddd6fe",borderRadius:12,cursor:"pointer",background:"#faf5ff",transition:"border-color .15s"}}>
        <i className="ti ti-file-spreadsheet" style={{fontSize:32,color:"#8b5cf6"}}/>
        <div style={{fontSize:14,fontWeight:600,color:"#374151"}}>{fileName||"Hacé click para elegir el archivo"}</div>
        <div style={{fontSize:12,color:"#9ca3af"}}>.xlsx o .xls</div>
        <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
      </label>

      {loading&&<div style={{fontSize:13,color:"#8b5cf6",display:"flex",alignItems:"center",gap:6}}><i className="ti ti-loader-2 ti-spin" style={{fontSize:15}}/>Procesando…</div>}
      {error&&<div style={{fontSize:12,color:"#b91c1c",background:"#fef2f2",padding:"8px 12px",borderRadius:8,border:"1px solid #fca5a5"}}>{error}</div>}

      {preview.length>0&&(
        <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column",gap:8}}>
          <div style={{fontSize:12,color:"#166534",background:"#f0fdf4",padding:"6px 12px",borderRadius:8,border:"1px solid #86efac",fontWeight:500}}>
            ✓ {preview.length} alumno{preview.length!==1?"s":""} listos para importar
          </div>
          <div style={{overflowY:"auto",maxHeight:180,border:"1px solid #e5e7eb",borderRadius:8}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:"#f9fafb",position:"sticky",top:0}}>
                <th style={{padding:"6px 10px",textAlign:"left",color:"#6b7280",fontWeight:600,borderBottom:"1px solid #e5e7eb"}}>Nombre</th>
                <th style={{padding:"6px 10px",textAlign:"left",color:"#6b7280",fontWeight:600,borderBottom:"1px solid #e5e7eb"}}>Liceo</th>
                <th style={{padding:"6px 10px",textAlign:"left",color:"#6b7280",fontWeight:600,borderBottom:"1px solid #e5e7eb"}}>CI</th>
                <th style={{padding:"6px 10px",textAlign:"left",color:"#6b7280",fontWeight:600,borderBottom:"1px solid #e5e7eb"}}>Cohorte</th>
              </tr></thead>
              <tbody>{preview.map((r,i)=><tr key={i} style={{borderBottom:"1px solid #f3f4f6"}}>
                <td style={{padding:"5px 10px",color:"#111827"}}>{r.nombre}</td>
                <td style={{padding:"5px 10px",color:"#6b7280"}}>{r.liceo||"—"}</td>
                <td style={{padding:"5px 10px",color:"#6b7280"}}>{r.ci||"—"}</td>
                <td style={{padding:"5px 10px",color:"#6b7280"}}>{r.cohorte||"—"}</td>
              </tr>)}</tbody>
            </table>
          </div>
        </div>
      )}

      <div style={{display:"flex",gap:8,marginTop:"auto"}}>
        <button onClick={onCancel} style={{flex:1,padding:"9px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:9,cursor:"pointer",fontSize:14}}>Cancelar</button>
        <button onClick={()=>preview.length>0&&onConfirm(preview)} disabled={preview.length===0}
          style={{flex:2,padding:"9px",background:preview.length>0?"#8b5cf6":"#f3f4f6",color:preview.length>0?"#fff":"#9ca3af",border:"none",borderRadius:9,cursor:preview.length>0?"pointer":"not-allowed",fontSize:14,fontWeight:700}}>
          Importar {preview.length>0?`${preview.length} alumnos`:""}
        </button>
      </div>
    </div>
  </div>;
}

function EditCIModal({student,onConfirm,onCancel}) {
  const [ci,setCi]=useState(student.ci||"");
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,padding:"24px",width:"100%",maxWidth:360,boxShadow:"0 8px 40px rgba(0,0,0,.18)"}}>
      <div style={{fontSize:15,fontWeight:700,color:"#111827",marginBottom:4}}>Editar CI</div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:16}}>{student.name}</div>
      <input autoFocus value={ci} onChange={e=>setCi(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onConfirm(ci.trim())}
        placeholder="Ej: 5.234.567-8"
        style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #8b5cf6",fontSize:14,fontFamily:"inherit",color:"#111827",boxSizing:"border-box",outline:"none",marginBottom:16}}/>
      <div style={{display:"flex",gap:8}}>
        <button onClick={onCancel} style={{flex:1,padding:"8px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer",fontSize:13}}>Cancelar</button>
        <button onClick={()=>onConfirm(ci.trim())} style={{flex:2,padding:"8px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:700}}>Guardar</button>
      </div>
    </div>
  </div>;
}

function InlineTA({type,setType,liceo,setLiceo,allFormLiceos,allTypeOptions,liceoAdding,setLiceoAdding,liceoNewVal,setLiceoNewVal,confirmNewLiceo,typeAdding,setTypeAdding,typeNewVal,setTypeNewVal,confirmNewType}) {
  const inp={padding:"7px 11px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontSize:13,fontFamily:"inherit",width:"100%",boxSizing:"border-box"};
  return <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
    <div>
      <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Tipo</label>
      {typeAdding?<div style={{display:"flex",gap:4}}>
        <input autoFocus value={typeNewVal} onChange={e=>setTypeNewVal(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter"){const v=confirmNewType();if(v)setType(v);}if(e.key==="Escape"){setTypeAdding(false);setTypeNewVal("");}}} placeholder="Nuevo tipo…" style={{...inp,flex:1}}/>
        <button onClick={()=>{const v=confirmNewType();if(v)setType(v);}} style={{padding:"7px 10px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:8,cursor:"pointer"}}>✓</button>
        <button onClick={()=>{setTypeAdding(false);setTypeNewVal("");}} style={{padding:"7px 10px",background:"transparent",border:"1px solid #e5e7eb",color:"#6b7280",borderRadius:8,cursor:"pointer"}}>✕</button>
      </div>:<select value={type} onChange={e=>{if(e.target.value==="__new__")setTypeAdding(true);else setType(e.target.value);}} style={inp}>
        {allTypeOptions.map(t=><option key={t.key} value={t.key}>{t.label}</option>)}
        <option value="__new__">＋ Agregar tipo nuevo…</option>
      </select>}
    </div>
    <div>
      <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Liceo (si aplica)</label>
      {liceoAdding?<div style={{display:"flex",gap:4}}>
        <input autoFocus value={liceoNewVal} onChange={e=>setLiceoNewVal(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter"){const v=confirmNewLiceo();if(v)setLiceo(v);}if(e.key==="Escape"){setLiceoAdding(false);setLiceoNewVal("");}}} placeholder="Nuevo liceo…" style={{...inp,flex:1}}/>
        <button onClick={()=>{const v=confirmNewLiceo();if(v)setLiceo(v);}} style={{padding:"7px 10px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:8,cursor:"pointer"}}>✓</button>
        <button onClick={()=>{setLiceoAdding(false);setLiceoNewVal("");}} style={{padding:"7px 10px",background:"transparent",border:"1px solid #e5e7eb",color:"#6b7280",borderRadius:8,cursor:"pointer"}}>✕</button>
      </div>:<select value={liceo} onChange={e=>{if(e.target.value==="__new__")setLiceoAdding(true);else setLiceo(e.target.value);}} style={inp}>
        <option value="">Sin liceo</option>
        {allFormLiceos.map(a=><option key={a} value={a}>{a}</option>)}
        <option value="__new__">＋ Agregar liceo nuevo…</option>
      </select>}
    </div>
  </div>;
}

function InlineAddItem({allFormLiceos,allTypeOptions,onAdd,onCancel,showCancel=false,currentUserId,...shared}) {
  const [type,setType]=useState("contacto"); const [liceo,setLiceo]=useState(""); const [desc,setDesc]=useState("");
  const [photos,setPhotos]=useState([]); const [priority,setPriority]=useState("media"); const [alertDate,setAlertDate]=useState(""); const [showExtra,setShowExtra]=useState(false);
  const [assignedTo,setAssignedTo]=useState(currentUserId||"");
  const inp={padding:"7px 11px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontSize:13,fontFamily:"inherit",width:"100%",boxSizing:"border-box"};
  const handleAdd=()=>{
    if(!desc.trim())return;
    onAdd({type,liceo,description:desc,photos,priority,alertDate:alertDate||null,assignedTo:assignedTo||currentUserId});
    setType("contacto");setLiceo("");setDesc("");setPhotos([]);setPriority("media");setAlertDate("");setShowExtra(false);setAssignedTo(currentUserId||"");
  };
  return <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px",boxShadow:"0 1px 3px rgba(0,0,0,.06)"}}>
    <InlineTA type={type} setType={setType} liceo={liceo} setLiceo={setLiceo} allFormLiceos={allFormLiceos} allTypeOptions={allTypeOptions} {...shared}/>
    <div style={{marginBottom:10}}><input value={desc} onChange={e=>setDesc(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleAdd()} placeholder="Descripción de la tarea…" style={inp}/></div>

    {/* Asignar a */}
    <div style={{marginBottom:10}}>
      <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Asignar a</label>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {USERS.map(u=>{
          const active=assignedTo===u.id;
          return <button key={u.id} onClick={()=>setAssignedTo(u.id)}
            style={{display:"flex",alignItems:"center",gap:5,padding:"4px 10px",borderRadius:20,border:`1.5px solid ${active?u.color:"#e5e7eb"}`,background:active?u.color+"18":"transparent",color:active?u.color:"#6b7280",cursor:"pointer",fontSize:12,fontWeight:active?700:400}}>
            <span style={{width:16,height:16,borderRadius:"50%",background:u.color,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:"#fff",flexShrink:0}}>{u.avatar}</span>
            {u.name}{u.id===currentUserId?" (yo)":""}
          </button>;
        })}
      </div>
    </div>

    <button onClick={()=>setShowExtra(e=>!e)} style={{fontSize:12,color:"#6b7280",background:"none",border:"none",cursor:"pointer",padding:"0 0 8px",display:"flex",alignItems:"center",gap:4}}>
      <i className={`ti ${showExtra?"ti-chevron-up":"ti-chevron-down"}`} style={{fontSize:11}}/>{showExtra?"Ocultar opciones":"＋ Prioridad y alerta"}
    </button>
    {showExtra&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
      <div>
        <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Prioridad</label>
        <select value={priority} onChange={e=>setPriority(e.target.value)} style={inp}>
          <option value="alta">🔴 Alta</option><option value="media">🟡 Media</option><option value="baja">🔵 Baja</option>
        </select>
      </div>
      <div>
        <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Alerta para el día</label>
        <input type="date" value={alertDate} onChange={e=>setAlertDate(e.target.value)} style={inp}/>
      </div>
    </div>}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
      <PhotoPicker photos={photos} onChange={setPhotos} compact/>
      <div style={{display:"flex",gap:8,flexShrink:0}}>
        {showCancel&&<button onClick={onCancel} style={{padding:"7px 14px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer",fontSize:13}}>Cancelar</button>}
        <button onClick={handleAdd} style={{padding:"7px 16px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:6}}>
          <i className="ti ti-plus"/>Agregar
        </button>
      </div>
    </div>
  </div>;
}

const bInp={padding:"8px 12px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontSize:14,width:"100%",boxSizing:"border-box",fontFamily:"inherit"};
const sInp={...bInp,fontSize:13,padding:"7px 11px"};

function LoginScreen({onSelect}) {
  return <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#f5f3ff 0%,#e0f2fe 100%)",fontFamily:"system-ui,sans-serif",padding:24}}>
    <div style={{marginBottom:40,textAlign:"center"}}>
      <div style={{width:56,height:56,borderRadius:14,background:"#8b5cf6",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px",boxShadow:"0 4px 14px rgba(139,92,246,.35)"}}>
        <i className="ti ti-school" style={{fontSize:26,color:"#fff"}}/>
      </div>
      <div style={{fontSize:11,fontWeight:600,letterSpacing:"0.14em",textTransform:"uppercase",color:"#8b5cf6",marginBottom:6}}>Proyecto 3F</div>
      <div style={{fontSize:28,fontWeight:700,color:"#111827",marginBottom:6}}>Gestión de Alumnos</div>
      <div style={{fontSize:15,color:"#6b7280"}}>¿Quién está ingresando?</div>
    </div>
    <div style={{display:"flex",gap:16,flexWrap:"wrap",justifyContent:"center"}}>
      {USERS.map(u=><button key={u.id} onClick={()=>onSelect(u)}
        style={{display:"flex",flexDirection:"column",alignItems:"center",gap:14,padding:"28px 36px",background:"#fff",border:"1px solid #e5e7eb",borderRadius:16,cursor:"pointer",minWidth:130,boxShadow:"0 2px 8px rgba(0,0,0,.06)"}}>
        <div style={{width:60,height:60,borderRadius:"50%",background:u.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:700,color:"#fff",boxShadow:`0 4px 12px ${u.color}66`}}>{u.avatar}</div>
        <div style={{fontSize:17,fontWeight:600,color:"#111827"}}>{u.name}</div>
      </button>)}
    </div>
  </div>;
}

function VisitCard({visit,items,onStartEdit,onDeleteVisit,onAddVisitComment,onDeleteVisitComment,onResolveVisitNotes,onResolveItem,onAddComment,onDeleteComment,onEditItemText}) {
  const [showVC,setShowVC]=useState(false);
  const vis=items.filter(i=>i.visitId===visit.id);
  const oc=vis.filter(i=>i.status==="abierto").length;
  const bl=visit.bullets?.length?visit.bullets.filter(b=>b.trim()):(visit.notes?[visit.notes]:[]);
  return <div style={{background:"#fff",border:"1px solid #f3f4f6",borderRadius:12,padding:"14px 16px",marginBottom:10,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:vis.length?10:0}}>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:14,fontWeight:700,color:"#111827",marginBottom:bl.length?4:0}}>{fmt(visit.date)}</div>
        <BulletDisplay bullets={bl}/>
        <PhotoStrip photos={visit.photos}/>
        <div style={{marginTop:8,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>setShowVC(e=>!e)}
            style={{fontSize:12,color:"#8b5cf6",background:"none",border:"none",cursor:"pointer",padding:0,display:"flex",alignItems:"center",gap:4,fontWeight:500}}>
            <i className="ti ti-message-circle" style={{fontSize:13}}/>
            {(visit.comments||[]).length>0?`${visit.comments.length} comentario${visit.comments.length>1?"s":""}`:"Comentar notas"}
            <i className={`ti ${showVC?"ti-chevron-up":"ti-chevron-down"}`} style={{fontSize:11}}/>
          </button>
          {visit.notesResolved&&<span style={{fontSize:11,color:"#166534",fontWeight:600,display:"flex",alignItems:"center",gap:3}}><i className="ti ti-check" style={{fontSize:11}}/>Notas resueltas</span>}
        </div>
        {showVC&&<VisitCommentThread visit={visit} onAddComment={onAddVisitComment} onDeleteComment={onDeleteVisitComment} onResolve={onResolveVisitNotes}/>}
      </div>
      <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0,marginLeft:12}}>
        {vis.length>0&&<span style={{fontSize:11,fontWeight:600,color:oc>0?"#b45309":"#166534",background:oc>0?"#fffbeb":"#f0fdf4",padding:"2px 8px",borderRadius:8,border:`1px solid ${oc>0?"#fcd34d":"#86efac"}`}}>{oc>0?`${oc} abierta${oc!==1?"s":""}` :"✓ Todo ok"}</span>}
        <button onClick={()=>onStartEdit(visit.id)} style={{display:"flex",alignItems:"center",gap:4,fontSize:12,padding:"5px 11px",border:"1px solid #e5e7eb",background:"#f9fafb",color:"#6b7280",borderRadius:7,cursor:"pointer",fontWeight:500}}>
          <i className="ti ti-pencil" style={{fontSize:12}}/> Editar
        </button>
        {onDeleteVisit&&<button onClick={()=>onDeleteVisit(visit.id)} style={{display:"flex",alignItems:"center",gap:4,fontSize:12,padding:"5px 9px",border:"1px solid #fca5a5",background:"#fef2f2",color:"#b91c1c",borderRadius:7,cursor:"pointer"}}>
          <i className="ti ti-trash" style={{fontSize:12}}/>
        </button>}
      </div>
    </div>
    {vis.map(item=><div key={item.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderTop:"1px solid #f9fafb",flexWrap:"wrap"}}>
      <TypeBadge type={item.type} size="sm"/>{item.liceo&&<LiceoChip liceo={item.liceo} size="sm"/>}
      {item.priority&&item.priority!=="media"&&<PriorityBadge priority={item.priority} size="sm"/>}
      <span style={{fontSize:13,color:"#374151",flex:1,minWidth:120}}>{item.description}</span>
      {item.alertDate&&<span style={{fontSize:11,color:item.alertDate<=today()?"#b91c1c":"#6b7280",display:"flex",alignItems:"center",gap:3}}><i className="ti ti-bell" style={{fontSize:11}}/>{fmtS(item.alertDate)}</span>}
      {(item.photos||[]).length>0&&<span style={{fontSize:11,color:"#9ca3af",display:"flex",alignItems:"center",gap:3}}><i className="ti ti-photo" style={{fontSize:12}}/>{item.photos.length}</span>}
      {(item.comments||[]).length>0&&<span style={{fontSize:11,color:"#9ca3af",display:"flex",alignItems:"center",gap:3}}><i className="ti ti-message-circle" style={{fontSize:12}}/>{item.comments.length}</span>}
      {item.status==="abierto"?<button onClick={()=>onResolveItem(item.id)} style={{fontSize:11,padding:"2px 9px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:6,cursor:"pointer",flexShrink:0,fontWeight:600}}>✓ Resolver</button>:<span style={{fontSize:11,color:"#166534",fontWeight:700}}>✓ Resuelto</span>}
    </div>)}
  </div>;
}

export default function App() {
  const [currentUser,setCurrentUser] = useState(null);
  const [loaded,setLoaded]           = useState(false);
  const unsubRef = useRef([]);

  const [students,setStudents]         = useState([]);
  const [visits,setVisits]             = useState([]);
  const [items,setItems]               = useState([]);
  const [customLiceos,setCustomLiceos] = useState([]);
  const [customTypes,setCustomTypes]   = useState([]);
  const [assignedOutItems,setAssignedOutItems] = useState([]); // tasks I assigned to others

  const [view,setView]               = useState("dashboard");
  const [selStudId,setSelStudId]     = useState(null);
  const [showNewStudentModal,setShowNewStudentModal] = useState(false);
  const [showBulkImport,setShowBulkImport]           = useState(false);
  const [editingStudId,setEditingStudId] = useState(null);
  const [editStudName,setEditStudName]   = useState("");
  const [editCIModal,setEditCIModal]     = useState(null);

  const [fLiceo,setFLiceo]       = useState("");
  const [fType,setFType]         = useState("");
  const [fStud,setFStud]         = useState("");
  const [fPriority,setFPriority] = useState("");
  const [showResolved,setShowResolved] = useState(false);

  const [nvStudId,setNvStudId]   = useState("");
  const [nvDate,setNvDate]       = useState(today());
  const [nvBullets,setNvBullets] = useState([""]);
  const [nvPhotos,setNvPhotos]   = useState([]);
  const [nvItems,setNvItems]     = useState([]);

  const [liceoAdding,setLiceoAdding] = useState(false);
  const [liceoNewVal,setLiceoNewVal] = useState("");
  const [typeAdding,setTypeAdding]   = useState(false);
  const [typeNewVal,setTypeNewVal]   = useState("");

  const [editVid,setEditVid]           = useState(null);
  const [editDate,setEditDate]         = useState("");
  const [editBullets,setEditBullets]   = useState([""]);
  const [editPhotos,setEditPhotos]     = useState([]);
  const [editItems,setEditItems]       = useState([]);
  const [editDelIds,setEditDelIds]     = useState(new Set());
  const [editItemId,setEditItemId]     = useState(null);
  const [eiType,setEiType]             = useState("");
  const [eiLiceo,setEiLiceo]           = useState("");
  const [eiDesc,setEiDesc]             = useState("");
  const [eiPhotos,setEiPhotos]         = useState([]);
  const [eiPriority,setEiPriority]     = useState("media");
  const [eiAlertDate,setEiAlertDate]   = useState("");
  const [showAddInEdit,setShowAddInEdit] = useState(false);

  const [sidebarOpen,setSidebarOpen] = useState(true);
  const [isMobile,setIsMobile]       = useState(false);
  const [searchStudents,setSearchStudents] = useState("");

  useEffect(()=>{
    const check=()=>{ const m=window.innerWidth<640; setIsMobile(m); if(!m)setSidebarOpen(true); };
    check(); window.addEventListener("resize",check); return ()=>window.removeEventListener("resize",check);
  },[]);
  useEffect(()=>()=>unsubRef.current.forEach(fn=>fn()),[]);

  function selectUser(u) {
    unsubRef.current.forEach(fn=>fn());
    setCurrentUser(u); setStudents([]); setVisits([]); setItems([]); setCustomLiceos([]); setCustomTypes([]); setAssignedOutItems([]);
    setLoaded(true); setView("dashboard");
    const uid=u.id;
    // Listen to own data (shared students + per-user visits/items)
    const ownUnsubs=[
      onSnapshot(studentsCol(), s=>setStudents(s.docs.map(d=>d.data())), e=>console.error(e)),
      onSnapshot(uCol(uid,"visits"),    s=>setVisits(s.docs.map(d=>d.data())),   e=>console.error(e)),
      onSnapshot(uCol(uid,"items"),     s=>setItems(s.docs.map(d=>d.data())),    e=>console.error(e)),
      onSnapshot(uCfg(uid), s=>{ if(s.exists()){const d=s.data();setCustomLiceos(d.customAreas||[]);setCustomTypes(d.customTypes||[]);} }, e=>console.error(e)),
    ];
    // Listen to items assigned to others by me (read from other users' collections)
    const otherUsers=USERS.filter(x=>x.id!==uid);
    const outMap={};
    const outUnsubs=otherUsers.map(ou=>{
      return onSnapshot(uCol(ou.id,"items"), s=>{
        const mine=s.docs.map(d=>d.data()).filter(it=>it.assignedByUid===uid||it.assignedBy===u.name);
        outMap[ou.id]=mine;
        setAssignedOutItems(Object.values(outMap).flat());
      }, e=>console.error(e));
    });
    unsubRef.current=[...ownUnsubs,...outUnsubs];
  }
  function logout() {
    unsubRef.current.forEach(fn=>fn()); unsubRef.current=[];
    setCurrentUser(null); setLoaded(false); setStudents([]); setVisits([]); setItems([]); setCustomLiceos([]); setCustomTypes([]); setAssignedOutItems([]);
  }
  function navAndClose(fn) { fn(); if(isMobile) setSidebarOpen(false); }

  const uid=currentUser?.id;
  const allFormLiceos=[...LICEOS_DEFAULT,...customLiceos.filter(l=>!LICEOS_DEFAULT.includes(l))];
  const allTypeOptions=[...Object.entries(ITEM_TYPES).map(([k,v])=>({key:k,label:v})),...customTypes.map(t=>({key:t,label:t}))];
  const allItemTypes=[...new Set(items.map(i=>i.type))];

  function confirmNewLiceo(){
    const val=liceoNewVal.trim(); setLiceoAdding(false); setLiceoNewVal(""); if(!val)return null;
    const n=[...LICEOS_DEFAULT,...customLiceos].includes(val)?customLiceos:[...customLiceos,val];
    setCustomLiceos(n); fbSetCfg(uid,{customAreas:n,customTypes}); return val;
  }
  function confirmNewType(){
    const val=typeNewVal.trim(); setTypeAdding(false); setTypeNewVal(""); if(!val)return null;
    const n=[...Object.keys(ITEM_TYPES),...customTypes].includes(val)?customTypes:[...customTypes,val];
    setCustomTypes(n); fbSetCfg(uid,{customAreas:customLiceos,customTypes:n}); return val;
  }
  const sharedTA={allFormLiceos,allTypeOptions,liceoAdding,setLiceoAdding,liceoNewVal,setLiceoNewVal,confirmNewLiceo,typeAdding,setTypeAdding,typeNewVal,setTypeNewVal,confirmNewType};

  function addStudent(nombre,ci,liceo,cohorte){ const s={id:genId(),name:nombre,ci:ci||"",liceo:liceo||"",cohorte:cohorte||""}; fbSetStudent(s.id,s); setShowNewStudentModal(false); }
  async function bulkAddStudents(list){
    try {
      const batch=writeBatch(db);
      list.forEach(({nombre,ci,liceo,cohorte})=>{
        const s={id:genId(),name:nombre,ci:ci||"",liceo:liceo||"",cohorte:cohorte||""};
        batch.set(studentDoc(s.id),s);
      });
      await batch.commit();
      setShowBulkImport(false);
    } catch(e){
      console.error(e);
      alert("Error al importar: " + (e.message||"verificá las reglas de Firestore y tu conexión."));
    }
  }
  function saveStudName(id){ if(!editStudName.trim())return; const s=students.find(s=>s.id===id); if(!s)return; const u={...s,name:editStudName.trim()}; setStudents(p=>p.map(x=>x.id===id?u:x)); fbSetStudent(id,u); setEditingStudId(null); }
  function saveStudCI(id,ci){ const s=students.find(s=>s.id===id); if(!s)return; const u={...s,ci}; setStudents(p=>p.map(x=>x.id===id?u:x)); fbSetStudent(id,u); setEditCIModal(null); }

  async function deleteVisit(visitId){
    if(!window.confirm("¿Eliminar esta gestión y todas sus tareas?"))return;
    const visitItems=items.filter(i=>i.visitId===visitId);
    const batch=writeBatch(db);
    batch.delete(doc(db,`users/${uid}/visits/${visitId}`));
    visitItems.forEach(it=>batch.delete(doc(db,`users/${uid}/items/${it.id}`)));
    await batch.commit();
    setVisits(p=>p.filter(v=>v.id!==visitId));
    setItems(p=>p.filter(i=>i.visitId!==visitId));
    if(view==="editvisit") setView("student");
  }

  async function deleteAllStudents(){
    if(!window.confirm(`¿Eliminar los ${students.length} alumnos y TODOS sus datos (gestiones y tareas)? Esta acción no se puede deshacer.`))return;
    const batch=writeBatch(db);
    students.forEach(s=>batch.delete(studentDoc(s.id)));
    visits.forEach(v=>batch.delete(doc(db,`users/${uid}/visits/${v.id}`)));
    items.forEach(i=>batch.delete(doc(db,`users/${uid}/items/${i.id}`)));
    await batch.commit();
    setStudents([]); setVisits([]); setItems([]);
  }
  async function deleteStudent(studId){
    if(!window.confirm("¿Eliminar este alumno, todas sus gestiones y tareas? Esta acción no se puede deshacer."))return;
    const studVisitIds=visits.filter(v=>v.locationId===studId).map(v=>v.id);
    const studItems=items.filter(i=>i.locationId===studId);
    const batch=writeBatch(db);
    batch.delete(studentDoc(studId));
    studVisitIds.forEach(vid=>batch.delete(doc(db,`users/${uid}/visits/${vid}`)));
    studItems.forEach(it=>batch.delete(doc(db,`users/${uid}/items/${it.id}`)));
    await batch.commit();
    setStudents(p=>p.filter(s=>s.id!==studId));
    setVisits(p=>p.filter(v=>v.locationId!==studId));
    setItems(p=>p.filter(i=>i.locationId!==studId));
    setView("students");
  }

  async function submitVisit(){
    if(!nvStudId||!nvDate)return;
    const visitId=genId();
    const visit={id:visitId,locationId:nvStudId,date:nvDate,bullets:nvBullets,photos:nvPhotos,comments:[],notesResolved:false};
    const selStudName=students.find(s=>s.id===nvStudId)?.name||"";
    const newItems=nvItems.map(it=>{
      const targetUid=it.assignedTo||uid;
      const assignedBy=targetUid!==uid?currentUser.name:null;
      const assignedByUid=targetUid!==uid?uid:null;
      return {...it,id:genId(),visitId,locationId:nvStudId,createdAt:nvDate,status:"abierto",comments:[],priority:it.priority||"media",alertDate:it.alertDate||null,assignedTo:targetUid,assignedBy,assignedByUid,studentName:selStudName};
    });
    const batch=writeBatch(db);
    batch.set(doc(db,`users/${uid}/visits/${visitId}`),visit);
    newItems.forEach(it=>batch.set(doc(db,`users/${it.assignedTo}/items/${it.id}`),it));
    await batch.commit();
    setNvItems([]); setNvBullets([""]); setNvDate(today()); setNvPhotos([]);
    setSelStudId(nvStudId); setView("student");
  }
  function resolveItem(id){ const it=items.find(i=>i.id===id); if(!it)return; const u={...it,status:"resuelto",resolvedAt:today()}; setItems(p=>p.map(i=>i.id===id?u:i)); fbSet(uid,"items",id,u); }
  function addComment(itemId,text){ const it=items.find(i=>i.id===itemId); if(!it)return; const c={id:genId(),text:text.trim(),date:today()}; const u={...it,comments:[...(it.comments||[]),c]}; setItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(uid,"items",itemId,u); }
  function deleteComment(itemId,commentId){ const it=items.find(i=>i.id===itemId); if(!it)return; const u={...it,comments:(it.comments||[]).filter(c=>c.id!==commentId)}; setItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(uid,"items",itemId,u); }
  function editItemText(itemId,newDesc){ const it=items.find(i=>i.id===itemId); if(!it)return; const u={...it,description:newDesc}; setItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(uid,"items",itemId,u); }

  /* ── Handlers for items assigned OUT (write to assignee's collection) ── */
  function resolveAssignedItem(itemId){
    const it=assignedOutItems.find(i=>i.id===itemId); if(!it)return;
    const u={...it,status:"resuelto",resolvedAt:today()};
    setAssignedOutItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(it.assignedTo,"items",itemId,u);
  }
  function addCommentAssigned(itemId,text){
    const it=assignedOutItems.find(i=>i.id===itemId); if(!it)return;
    const c={id:genId(),text:text.trim(),date:today()};
    const u={...it,comments:[...(it.comments||[]),c]};
    setAssignedOutItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(it.assignedTo,"items",itemId,u);
  }
  function deleteCommentAssigned(itemId,commentId){
    const it=assignedOutItems.find(i=>i.id===itemId); if(!it)return;
    const u={...it,comments:(it.comments||[]).filter(c=>c.id!==commentId)};
    setAssignedOutItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(it.assignedTo,"items",itemId,u);
  }
  function editAssignedItemText(itemId,newDesc){
    const it=assignedOutItems.find(i=>i.id===itemId); if(!it)return;
    const u={...it,description:newDesc};
    setAssignedOutItems(p=>p.map(i=>i.id===itemId?u:i)); fbSet(it.assignedTo,"items",itemId,u);
  }
  async function deleteAssignedItem(itemId){
    const it=assignedOutItems.find(i=>i.id===itemId); if(!it)return;
    if(!window.confirm("¿Eliminar esta tarea asignada?"))return;
    setAssignedOutItems(p=>p.filter(i=>i.id!==itemId));
    try{ await deleteDoc(uDoc(it.assignedTo,"items",itemId)); }catch(e){console.error(e);}
  }
  function addVisitComment(visitId,text){ const v=visits.find(v=>v.id===visitId); if(!v)return; const c={id:genId(),text:text.trim(),date:today()}; const u={...v,comments:[...(v.comments||[]),c]}; setVisits(p=>p.map(x=>x.id===visitId?u:x)); fbSet(uid,"visits",visitId,u); }
  function deleteVisitComment(visitId,commentId){ const v=visits.find(v=>v.id===visitId); if(!v)return; const u={...v,comments:(v.comments||[]).filter(c=>c.id!==commentId)}; setVisits(p=>p.map(x=>x.id===visitId?u:x)); fbSet(uid,"visits",visitId,u); }
  function resolveVisitNotes(visitId){ const v=visits.find(v=>v.id===visitId); if(!v)return; const u={...v,notesResolved:true}; setVisits(p=>p.map(x=>x.id===visitId?u:x)); fbSet(uid,"visits",visitId,u); }

  function startEdit(visitId){
    const v=visits.find(v=>v.id===visitId); if(!v)return;
    setEditVid(visitId); setEditDate(v.date);
    setEditBullets(v.bullets?.length?v.bullets:[v.notes||""]);
    setEditPhotos(v.photos||[]);
    setEditItems(items.filter(i=>i.visitId===visitId).map(i=>({...i})));
    setEditDelIds(new Set()); setEditItemId(null); setShowAddInEdit(false); setView("editvisit");
  }
  function startEditItem(it){setEditItemId(it.id);setEiType(it.type);setEiLiceo(it.liceo||"");setEiDesc(it.description);setEiPhotos(it.photos||[]);setEiPriority(it.priority||"media");setEiAlertDate(it.alertDate||"");}
  function saveEditItem(id){setEditItems(p=>p.map(it=>it.id===id?{...it,type:eiType,liceo:eiLiceo,description:eiDesc,photos:eiPhotos,priority:eiPriority,alertDate:eiAlertDate||null}:it));setEditItemId(null);}
  function deleteEditItem(id){setEditDelIds(p=>new Set([...p,id]));if(editItemId===id)setEditItemId(null);}
  function addEditItem(item){setEditItems(p=>[...p,{...item,id:genId(),visitId:editVid,locationId:visits.find(v=>v.id===editVid)?.locationId,createdAt:editDate,status:"abierto",comments:[],priority:item.priority||"media",alertDate:item.alertDate||null,assignedTo:item.assignedTo||uid,assignedBy:item.assignedTo&&item.assignedTo!==uid?currentUser.name:null}]);setShowAddInEdit(false);}
  async function saveEdit(){
    const orig=visits.find(v=>v.id===editVid);
    const updV={...orig,date:editDate,bullets:editBullets,photos:editPhotos};
    const kept=editItems.filter(i=>!editDelIds.has(i.id));
    const batch=writeBatch(db);
    batch.set(doc(db,`users/${uid}/visits/${editVid}`),updV);
    kept.forEach(it=>batch.set(doc(db,`users/${it.assignedTo||uid}/items/${it.id}`),it));
    editDelIds.forEach(id=>batch.delete(doc(db,`users/${uid}/items/${id}`)));
    await batch.commit();
    setVisits(p=>p.map(v=>v.id===editVid?updV:v));
    setItems(p=>[...p.filter(i=>i.visitId!==editVid),...kept.filter(i=>!i.assignedTo||i.assignedTo===uid)]);
    setView("student");
  }

  const openItems=items.filter(i=>i.status==="abierto");
  const resolvedItems=items.filter(i=>i.status==="resuelto");
  const overdueItems=openItems.filter(i=>i.alertDate&&i.alertDate<=today());
  const highPriorityItems=openItems.filter(i=>i.priority==="alta");
  const allLiceos=[...new Set(openItems.filter(i=>i.liceo).map(i=>i.liceo))].sort();
  const allResolvedLiceos=[...new Set(resolvedItems.filter(i=>i.liceo).map(i=>i.liceo))].sort();
  const selStud=students.find(s=>s.id===selStudId);
  const studVisits=selStud?visits.filter(v=>v.locationId===selStudId).sort((a,b)=>b.date.localeCompare(a.date)):[];
  const studOpen=selStud?items.filter(i=>i.locationId===selStudId&&i.status==="abierto"):[];
  const prevOpen=nvStudId?items.filter(i=>i.locationId===nvStudId&&i.status==="abierto"):[];
  const editStud=editVid?students.find(s=>s.id===visits.find(v=>v.id===editVid)?.locationId):null;
  const pendFiltered=openItems.filter(i=>!fLiceo||i.liceo===fLiceo).filter(i=>!fType||i.type===fType).filter(i=>!fStud||i.locationId===fStud).filter(i=>!fPriority||i.priority===fPriority)
    .sort((a,b)=>{ const po={alta:0,media:1,baja:2}; const pa=po[a.priority]??1, pb=po[b.priority]??1; if(pa!==pb)return pa-pb; return (b.createdAt||"").localeCompare(a.createdAt||""); });
  const resolvedFiltered=resolvedItems.filter(i=>!fLiceo||i.liceo===fLiceo).filter(i=>!fStud||i.locationId===fStud).sort((a,b)=>(b.resolvedAt||b.createdAt||"").localeCompare(a.resolvedAt||a.createdAt||""));

  if(!currentUser) return <LoginScreen onSelect={selectUser}/>;

  const SL=({children})=><div style={{fontSize:10,fontWeight:700,letterSpacing:"0.1em",textTransform:"uppercase",color:"#9ca3af",padding:"8px 18px 4px"}}>{children}</div>;

  return (
    <div style={{display:"flex",minHeight:"100vh",fontFamily:"system-ui,sans-serif",background:"#f8f9fb",position:"relative"}}>
      {showNewStudentModal&&<NewStudentModal onConfirm={addStudent} onCancel={()=>setShowNewStudentModal(false)}/>}
      {showBulkImport&&<BulkImportModal onConfirm={bulkAddStudents} onCancel={()=>setShowBulkImport(false)}/>}
      {editCIModal&&<EditCIModal student={editCIModal} onConfirm={ci=>saveStudCI(editCIModal.id,ci)} onCancel={()=>setEditCIModal(null)}/>}

      {sidebarOpen&&isMobile&&<div onClick={()=>setSidebarOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.35)",zIndex:40}}/>}
      <button onClick={()=>setSidebarOpen(o=>!o)} style={{position:"fixed",top:12,left:12,zIndex:50,width:36,height:36,borderRadius:9,background:"#fff",border:"1px solid #e5e7eb",boxShadow:"0 1px 4px rgba(0,0,0,.1)",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",color:"#374151"}}>
        <i className={`ti ${sidebarOpen?"ti-x":"ti-menu-2"}`} style={{fontSize:17}}/>
      </button>

      <aside style={{width:228,background:"#fff",borderRight:"1px solid #f3f4f6",display:"flex",flexDirection:"column",flexShrink:0,boxShadow:"2px 0 8px rgba(0,0,0,.06)",position:isMobile?"fixed":"sticky",top:0,left:0,height:"100vh",zIndex:45,transform:sidebarOpen?"translateX(0)":"translateX(-100%)",transition:"transform .22s ease"}}>
        <div style={{padding:"20px 18px 16px",borderBottom:"1px solid #f3f4f6",paddingLeft:56}}>
          <div style={{fontSize:10,fontWeight:700,letterSpacing:"0.12em",textTransform:"uppercase",color:"#9ca3af",marginBottom:10}}>Proyecto 3F</div>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,borderRadius:"50%",background:currentUser.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:"#fff",flexShrink:0}}>{currentUser.avatar}</div>
            <div>
              <div style={{fontSize:15,fontWeight:700,color:"#111827"}}>{currentUser.name}</div>
              <button onClick={logout} style={{fontSize:11,color:"#9ca3af",background:"none",border:"none",cursor:"pointer",padding:0,display:"flex",alignItems:"center",gap:3,marginTop:1}}>
                <i className="ti ti-logout" style={{fontSize:11}}/>Cambiar usuario
              </button>
            </div>
          </div>
        </div>
        <nav style={{flex:1,paddingTop:6,overflowY:"auto"}}>
          <SL>General</SL>
          <NavBtn icon="ti-layout-dashboard" label="Inicio" active={view==="dashboard"} onClick={()=>navAndClose(()=>setView("dashboard"))} color={currentUser.color}/>
          <NavBtn icon="ti-users" label="Mis Alumnos" active={["students","student","editvisit"].includes(view)} onClick={()=>navAndClose(()=>setView("students"))} color={currentUser.color}/>
          <div style={{height:"1px",background:"#f3f4f6",margin:"8px 0"}}/>
          <SL>Seguimiento</SL>
          <NavBtn icon="ti-clock-exclamation" label="Tareas abiertas" active={view==="pending"&&!showResolved} badge={openItems.length}
            onClick={()=>navAndClose(()=>{setFLiceo("");setFType("");setFStud("");setFPriority("");setShowResolved(false);setView("pending");})} color={currentUser.color}/>
          <NavBtn icon="ti-user-share" label="Asignadas" active={view==="asignadas"} badge={assignedOutItems.filter(i=>i.status==="abierto").length+(items.filter(i=>i.assignedBy&&i.status==="abierto").length)}
            onClick={()=>navAndClose(()=>setView("asignadas"))} color="#7e22ce"/>
          <NavBtn icon="ti-alert-triangle" label="Alta prioridad" active={view==="priority"} badge={highPriorityItems.length}
            onClick={()=>navAndClose(()=>setView("priority"))} color="#b91c1c"/>
          <NavBtn icon="ti-bell" label="Alertas vencidas" active={view==="alerts"} badge={overdueItems.length}
            onClick={()=>navAndClose(()=>setView("alerts"))} color="#f59e0b"/>
          <NavBtn icon="ti-circle-check" label="Historial resueltos" active={view==="pending"&&showResolved} badge={resolvedItems.length}
            onClick={()=>navAndClose(()=>{setFLiceo("");setFType("");setFStud("");setFPriority("");setShowResolved(true);setView("pending");})} color="#10b981"/>
          {!showResolved&&allLiceos.slice(0,5).map(liceo=>{
            const n=openItems.filter(i=>i.liceo===liceo).length;
            return <NavBtn key={liceo} icon="ti-school" label={liceo} sub active={view==="pending"&&fLiceo===liceo&&!showResolved} badge={n}
              onClick={()=>navAndClose(()=>{setFLiceo(liceo);setFType("");setFStud("");setFPriority("");setShowResolved(false);setView("pending");})}/>;
          })}
          <div style={{height:"1px",background:"#f3f4f6",margin:"8px 0"}}/>
          <NavBtn icon="ti-calendar-plus" label="Nueva Gestión" active={view==="newvisit"} onClick={()=>navAndClose(()=>{setNvStudId("");setView("newvisit");})} color={currentUser.color}/>
        </nav>
        <div style={{padding:"12px 18px",borderTop:"1px solid #f3f4f6"}}>
          <div style={{fontSize:11,color:"#9ca3af"}}>{students.length} alumnos · {visits.length} gestiones</div>
        </div>
      </aside>

      <main style={{flex:1,overflow:"auto",paddingTop:isMobile?52:0}}>

        {view==="dashboard"&&(
          <div style={{padding:"28px",maxWidth:780}}>
            <div style={{marginBottom:24}}>
              <h1 style={{fontSize:24,fontWeight:700,margin:"0 0 4px",color:"#111827",paddingLeft:isMobile?0:48}}>Inicio</h1>
              <div style={{fontSize:14,color:"#6b7280",paddingLeft:isMobile?0:48}}>Bienvenido/a, {currentUser.name}</div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:28}}>
              <StatCard label="Alumnos activos" value={students.length} icon="ti-users" palette={PALETTE.purple}/>
              <StatCard label="Gestiones totales" value={visits.length} icon="ti-calendar-event" palette={PALETTE.sky}/>
              <StatCard label="Tareas abiertas" value={openItems.length} icon="ti-clock-exclamation" palette={openItems.length>0?PALETTE.amber:PALETTE.green} onClick={()=>{setFLiceo("");setFType("");setFStud("");setFPriority("");setShowResolved(false);setView("pending");}}/>
              <StatCard label="Alta prioridad" value={highPriorityItems.length} icon="ti-alert-triangle" palette={highPriorityItems.length>0?PALETTE.red:PALETTE.green} onClick={()=>setView("priority")}/>
              <StatCard label="Alertas vencidas" value={overdueItems.length} icon="ti-bell" palette={overdueItems.length>0?PALETTE.orange:PALETTE.green} onClick={()=>setView("alerts")}/>
            </div>
            {overdueItems.length>0&&<div style={{marginBottom:24,background:"#fff7ed",border:"1px solid #fdba74",borderRadius:12,padding:"16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#c2410c",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
                <i className="ti ti-bell" style={{fontSize:14}}/>Alertas vencidas hoy ({overdueItems.length})
              </div>
              {overdueItems.slice(0,3).map(item=>{const stud=students.find(s=>s.id===item.locationId);const v=visits.find(v=>v.id===item.visitId);return <ItemCard key={item.id} item={item} studentName={stud?.name} visitDate={v?.date} onResolve={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={editItemText}/>;})}</div>}
            {allLiceos.length>0&&<div style={{marginBottom:28}}>
              <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12}}>Tareas abiertas por liceo</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:8}}>
                {allLiceos.map(liceo=>{const n=openItems.filter(i=>i.liceo===liceo).length;const p=liceoPal(liceo);return(
                  <button key={liceo} onClick={()=>{setFLiceo(liceo);setFType("");setFStud("");setFPriority("");setShowResolved(false);setView("pending");}}
                    style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:p.bg,border:`1px solid ${p.border}`,borderRadius:10,cursor:"pointer",textAlign:"left"}}>
                    <span style={{fontSize:13,color:p.fg,fontWeight:600}}>{liceo}</span>
                    <span style={{background:"#fff",color:p.fg,fontSize:12,padding:"2px 10px",borderRadius:10,fontWeight:700,border:`1px solid ${p.border}`}}>{n}</span>
                  </button>
                );})}
              </div>
            </div>}
            {openItems.length>0&&<div>
              <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12}}>Últimas tareas abiertas</div>
              {[...openItems].sort((a,b)=>(b.createdAt||"").localeCompare(a.createdAt||"")).slice(0,4).map(item=>{const stud=students.find(s=>s.id===item.locationId);const v=visits.find(v=>v.id===item.visitId);return <ItemCard key={item.id} item={item} studentName={stud?.name} visitDate={v?.date} onResolve={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={editItemText}/>;})}</div>}
            {assignedOutItems.length>0&&<div style={{marginBottom:28}}>
              <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
                <i className="ti ti-user-share" style={{fontSize:13,color:"#7e22ce"}}/>Tareas asignadas a otros ({assignedOutItems.filter(i=>i.status==="abierto").length} abiertas)
                <button onClick={()=>setView("asignadas")} style={{marginLeft:"auto",fontSize:11,color:"#7e22ce",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>Ver todas ↗</button>
              </div>
              {assignedOutItems.filter(i=>i.status==="abierto").slice(0,3).map(item=>{
                const assigneeName=USERS.find(u=>u.id===item.assignedTo)?.name;
                return <ItemCard key={item.id} item={{...item,assignedBy:null}}
                  studentName={item.studentName} visitDate={item.createdAt}
                  onResolve={resolveAssignedItem} onAddComment={addCommentAssigned}
                  onDeleteComment={deleteCommentAssigned} onEditItem={editAssignedItemText}
                  onDelete={deleteAssignedItem}
                  showStudent={true}/>;
              })}
            </div>}
            {students.length===0&&<div style={{textAlign:"center",padding:"60px 24px",background:"#f5f3ff",borderRadius:16,border:"1px dashed #c4b5fd"}}>
              <i className="ti ti-users" style={{fontSize:40,display:"block",marginBottom:12,color:"#c4b5fd"}}/>
              <p style={{margin:"0 0 16px",fontSize:15,color:"#6b7280"}}>Agregá tu primer alumno para comenzar</p>
              <button onClick={()=>setView("students")} style={{padding:"9px 22px",borderRadius:9,border:"none",background:"#8b5cf6",color:"#fff",cursor:"pointer",fontSize:14,fontWeight:600}}>Agregar alumno ↗</button>
            </div>}
          </div>
        )}

        {view==="priority"&&(
          <div style={{padding:"28px",maxWidth:760}}>
            <h1 style={{fontSize:22,fontWeight:700,margin:"0 0 6px",color:"#111827"}}>Alta Prioridad</h1>
            <div style={{fontSize:13,color:"#6b7280",marginBottom:22}}>{highPriorityItems.length} tarea{highPriorityItems.length!==1?"s":""} de alta prioridad</div>
            {highPriorityItems.length===0?<div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>¡Sin tareas de alta prioridad! ✓</div>
            :highPriorityItems.sort((a,b)=>(b.createdAt||"").localeCompare(a.createdAt||"")).map(item=>{const stud=students.find(s=>s.id===item.locationId);const v=visits.find(v=>v.id===item.visitId);return <ItemCard key={item.id} item={item} studentName={stud?.name} visitDate={v?.date} onResolve={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={editItemText}/>;})}</div>
        )}

        {view==="alerts"&&(
          <div style={{padding:"28px",maxWidth:760}}>
            <h1 style={{fontSize:22,fontWeight:700,margin:"0 0 6px",color:"#111827"}}>Alertas Vencidas</h1>
            <div style={{fontSize:13,color:"#6b7280",marginBottom:22}}>{overdueItems.length} alerta{overdueItems.length!==1?"s":""} vencida{overdueItems.length!==1?"s":""}</div>
            {overdueItems.length===0?<div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>¡Sin alertas vencidas! ✓</div>
            :overdueItems.sort((a,b)=>(a.alertDate||"").localeCompare(b.alertDate||"")).map(item=>{const stud=students.find(s=>s.id===item.locationId);const v=visits.find(v=>v.id===item.visitId);return <ItemCard key={item.id} item={item} studentName={stud?.name} visitDate={v?.date} onResolve={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={editItemText}/>;})}</div>
        )}

        {view==="students"&&(
          <div style={{padding:"28px",maxWidth:740}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,paddingLeft:isMobile?0:48}}>
              <h1 style={{fontSize:22,fontWeight:700,margin:0,color:"#111827"}}>Mis Alumnos</h1>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",justifyContent:"flex-end"}}>
                {students.length>0&&<button onClick={deleteAllStudents} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 12px",background:"#fef2f2",color:"#b91c1c",border:"1px solid #fca5a5",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:500}}>
                  <i className="ti ti-trash" style={{fontSize:13}}/> Eliminar todos
                </button>}
                <button onClick={()=>setShowBulkImport(true)} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"#f5f3ff",color:"#6d28d9",border:"1px solid #ddd6fe",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>
                  <i className="ti ti-table-import" style={{fontSize:14}}/> Importar lista
                </button>
                <button onClick={()=>setShowNewStudentModal(true)} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 16px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>
                  <i className="ti ti-user-plus" style={{fontSize:14}}/> Agregar alumno
                </button>
              </div>
            </div>

            {/* Search bar */}
            <div style={{position:"relative",marginBottom:16}}>
              <i className="ti ti-search" style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",fontSize:15,color:"#9ca3af",pointerEvents:"none"}}/>
              <input value={searchStudents} onChange={e=>setSearchStudents(e.target.value)}
                placeholder={`Buscar entre ${students.length} alumnos…`}
                style={{width:"100%",padding:"9px 12px 9px 36px",borderRadius:10,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#111827",fontSize:14,fontFamily:"inherit",boxSizing:"border-box",outline:"none"}}/>
              {searchStudents&&<button onClick={()=>setSearchStudents("")}
                style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:16,lineHeight:1}}>✕</button>}
            </div>

            {(()=>{
              const norm=s=>s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
              const q=norm(searchStudents.trim());
              const filtered=q?students.filter(s=>norm(s.name).includes(q)||norm(s.ci||"").includes(q)||norm(s.liceo||"").includes(q)||norm(s.cohorte||"").includes(q)):students;
              const sorted=[...filtered].sort((a,b)=>a.name.localeCompare(b.name,"es"));
              if(students.length===0)return <div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>Aún no hay alumnos registrados.</div>;
              if(sorted.length===0)return <div style={{textAlign:"center",padding:"32px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12}}>Sin resultados para "{searchStudents}"</div>;
              return <>
                {q&&<div style={{fontSize:12,color:"#6b7280",marginBottom:10}}>{sorted.length} resultado{sorted.length!==1?"s":""}</div>}
                {sorted.map(stud=>{
              const sv=visits.filter(v=>v.locationId===stud.id);
              const so=openItems.filter(i=>i.locationId===stud.id).length;
              const sr=resolvedItems.filter(i=>i.locationId===stud.id).length;
              const last=[...sv].sort((a,b)=>b.date.localeCompare(a.date))[0];
              const ap=[...new Set(openItems.filter(i=>i.locationId===stud.id&&i.liceo).map(i=>i.liceo))];
              return <div key={stud.id} style={{background:"#fff",border:"1px solid #f3f4f6",borderRadius:12,padding:"14px 16px",marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
                {editingStudId===stud.id?<div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <input value={editStudName} onChange={e=>setEditStudName(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")saveStudName(stud.id);if(e.key==="Escape")setEditingStudId(null);}} autoFocus style={{...bInp,flex:1,fontSize:15,fontWeight:600}}/>
                  <button onClick={()=>saveStudName(stud.id)} style={{padding:"8px 16px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600,whiteSpace:"nowrap"}}>✓ Guardar</button>
                  <button onClick={()=>setEditingStudId(null)} style={{padding:"8px 10px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer"}}>✕</button>
                </div>:<div style={{display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer"}} onClick={()=>{setSelStudId(stud.id);setView("student");}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5,flexWrap:"wrap"}}>
                      <span style={{fontSize:15,fontWeight:700,color:"#111827"}}>{stud.name}</span>
                      {stud.ci&&<span style={{fontSize:11,color:"#8b5cf6",background:"#f5f3ff",padding:"1px 8px",borderRadius:8,border:"1px solid #ddd6fe",fontWeight:500}}>CI: {stud.ci}</span>}
                      {stud.cohorte&&<span style={{fontSize:11,color:"#0369a1",background:"#f0f9ff",padding:"1px 8px",borderRadius:8,border:"1px solid #7dd3fc",fontWeight:500}}>Cohorte {stud.cohorte}</span>}
                      {stud.liceo&&<LiceoChip liceo={stud.liceo} size="sm"/>}
                      <button onClick={e=>{e.stopPropagation();setEditingStudId(stud.id);setEditStudName(stud.name);}} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,padding:"2px 8px",borderRadius:6,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#6b7280",cursor:"pointer",fontWeight:500}}>
                        <i className="ti ti-pencil" style={{fontSize:11}}/>Renombrar
                      </button>
                      <button onClick={e=>{e.stopPropagation();setEditCIModal(stud);}} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,padding:"2px 8px",borderRadius:6,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#6b7280",cursor:"pointer",fontWeight:500}}>
                        <i className="ti ti-id-badge" style={{fontSize:11}}/>{stud.ci?"Editar CI":"+ CI"}
                      </button>
                    </div>
                    <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                      <span style={{fontSize:12,color:"#9ca3af"}}>{sv.length} gestión{sv.length!==1?"es":""}{last?` · última: ${fmt(last.date)}`:" · sin gestiones"}</span>
                      {sr>0&&<span style={{fontSize:11,color:"#166534",background:"#f0fdf4",padding:"1px 7px",borderRadius:10,border:"1px solid #86efac",fontWeight:600}}>✓ {sr} resueltos</span>}
                      {ap.map(a=><LiceoChip key={a} liceo={a} size="sm"/>)}
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
                    {so>0&&<span style={{fontSize:12,background:"#fef3c7",color:"#92400e",padding:"3px 10px",borderRadius:10,fontWeight:700,border:"1px solid #fcd34d"}}>{so}</span>}
                    <button onClick={e=>{e.stopPropagation();deleteStudent(stud.id);}} style={{padding:"5px 8px",background:"#fef2f2",color:"#b91c1c",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer",display:"flex",alignItems:"center"}}>
                      <i className="ti ti-trash" style={{fontSize:13}}/>
                    </button>
                    <i className="ti ti-chevron-right" style={{fontSize:16,color:"#d1d5db"}}/>
                  </div>
                </div>}
              </div>;
            })}
              </>;
            })()}
          </div>
        )}

        {view==="student"&&selStud&&(
          <div style={{padding:"28px",maxWidth:740}}>
            <button onClick={()=>setView("students")} style={{background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:13,padding:"0 0 14px",display:"flex",alignItems:"center",gap:4,fontWeight:500,paddingLeft:isMobile?0:48}}>
              <i className="ti ti-arrow-left" style={{fontSize:13}}/> Alumnos
            </button>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:22}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:4,flexWrap:"wrap"}}>
                  <h1 style={{fontSize:22,fontWeight:700,margin:0,color:"#111827"}}>{selStud.name}</h1>
                  {selStud.ci&&<span style={{fontSize:12,color:"#8b5cf6",background:"#f5f3ff",padding:"2px 10px",borderRadius:8,border:"1px solid #ddd6fe",fontWeight:500}}>CI: {selStud.ci}</span>}
                  {selStud.cohorte&&<span style={{fontSize:12,color:"#0369a1",background:"#f0f9ff",padding:"2px 10px",borderRadius:8,border:"1px solid #7dd3fc",fontWeight:500}}>Cohorte {selStud.cohorte}</span>}
                  {selStud.liceo&&<LiceoChip liceo={selStud.liceo}/>}
                  <button onClick={()=>setEditCIModal(selStud)} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,padding:"3px 8px",borderRadius:6,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#6b7280",cursor:"pointer",fontWeight:500}}>
                    <i className="ti ti-id-badge" style={{fontSize:11}}/>{selStud.ci?"Editar CI":"+ CI"}
                  </button>
                </div>
                <div style={{fontSize:13,color:"#9ca3af"}}>{studVisits.length} gestión{studVisits.length!==1?"es":""} · {studOpen.length} tarea{studOpen.length!==1?"s":""} abierta{studOpen.length!==1?"s":""}</div>
              </div>
              <div style={{display:"flex",gap:8,flexShrink:0}}>
                <button onClick={()=>deleteStudent(selStud.id)} style={{display:"flex",alignItems:"center",gap:5,padding:"8px 12px",background:"#fef2f2",color:"#b91c1c",border:"1px solid #fca5a5",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:500}}>
                  <i className="ti ti-trash" style={{fontSize:13}}/> Eliminar alumno
                </button>
                <button onClick={()=>{setNvStudId(selStud.id);setView("newvisit");}} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 16px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>
                  <i className="ti ti-plus" style={{fontSize:14}}/> Nueva gestión
                </button>
              </div>
            </div>
            {studOpen.length>0&&<div style={{marginBottom:28}}>
              <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12}}>Tareas abiertas ({studOpen.length})</div>
              {studOpen.map(item=>{const v=visits.find(v=>v.id===item.visitId);return <ItemCard key={item.id} item={item} visitDate={v?.date} onResolve={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={editItemText} showStudent={false}/>;})}</div>}
            <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12}}>Historial de gestiones</div>
            {studVisits.length===0&&<div style={{textAlign:"center",padding:"32px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>Aún no hay gestiones registradas.</div>}
            {studVisits.map(visit=><VisitCard key={visit.id} visit={visit} items={items} onStartEdit={startEdit} onDeleteVisit={deleteVisit} onAddVisitComment={addVisitComment} onResolveVisitNotes={resolveVisitNotes} onResolveItem={resolveItem} onAddComment={addComment} onDeleteComment={deleteComment} onDeleteVisitComment={deleteVisitComment} onEditItemText={editItemText}/>)}
          </div>
        )}

        {view==="newvisit"&&(
          <div style={{padding:"28px",maxWidth:680}}>
            <h1 style={{fontSize:22,fontWeight:700,margin:"0 0 22px",color:"#111827",paddingLeft:isMobile?0:48}}>Nueva Gestión</h1>
            <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"20px",marginBottom:20,boxShadow:"0 1px 4px rgba(0,0,0,.05)"}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
                <div>
                  <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Alumno *</label>
                  <select value={nvStudId} onChange={e=>setNvStudId(e.target.value)} style={bInp}>
                    <option value="">Seleccioná un alumno</option>
                    {students.map(s=><option key={s.id} value={s.id}>{s.name}{s.ci?` (${s.ci})`:""}</option>)}
                  </select>
                </div>
                <div>
                  <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Fecha *</label>
                  <input type="date" value={nvDate} onChange={e=>setNvDate(e.target.value)} style={bInp}/>
                </div>
              </div>
              <div style={{marginBottom:14}}>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Notas generales</label>
                <BulletNotes bullets={nvBullets} onChange={setNvBullets}/>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Fotos</label>
                <PhotoPicker photos={nvPhotos} onChange={setNvPhotos}/>
              </div>
            </div>
            {prevOpen.length>0&&<div style={{background:"#fffbeb",borderRadius:12,padding:"16px",marginBottom:20,border:"1px solid #fcd34d"}}>
              <div style={{fontSize:12,fontWeight:600,color:"#92400e",marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
                <i className="ti ti-history" style={{fontSize:14}}/>Tareas anteriores — marcá las resueltas hoy
              </div>
              {prevOpen.map(item=>{const v=visits.find(v=>v.id===item.visitId);return(
                <div key={item.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderTop:"1px solid #fde68a",flexWrap:"wrap"}}>
                  <TypeBadge type={item.type} size="sm"/>{item.liceo&&<LiceoChip liceo={item.liceo} size="sm"/>}
                  {item.priority&&item.priority!=="media"&&<PriorityBadge priority={item.priority} size="sm"/>}
                  <span style={{fontSize:13,flex:1,minWidth:120,color:"#374151"}}>{item.description}</span>
                  <span style={{fontSize:11,color:"#9ca3af"}}>{fmt(v?.date)}</span>
                  <button onClick={()=>resolveItem(item.id)} style={{fontSize:11,padding:"3px 10px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:6,cursor:"pointer",flexShrink:0,fontWeight:600}}>✓ Resolver</button>
                </div>
              );})}
            </div>}
            <div style={{fontSize:14,fontWeight:700,color:"#374151",marginBottom:10}}>Agregar tareas</div>
            <InlineAddItem {...sharedTA} currentUserId={uid} onAdd={item=>setNvItems(p=>[...p,item])}/>
            {nvItems.length>0&&<div style={{marginTop:10,marginBottom:16}}>
              {nvItems.map((item,idx)=><div key={item.id||idx} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:"10px 14px",marginBottom:6}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  <TypeBadge type={item.type} size="sm"/>{item.liceo&&<LiceoChip liceo={item.liceo} size="sm"/>}
                  {item.priority&&item.priority!=="media"&&<PriorityBadge priority={item.priority} size="sm"/>}
                  <span style={{fontSize:13,flex:1,minWidth:120,color:"#374151"}}>{item.description}</span>
                  {item.alertDate&&<span style={{fontSize:11,color:"#6b7280",display:"flex",alignItems:"center",gap:3}}><i className="ti ti-bell" style={{fontSize:11}}/>{fmtS(item.alertDate)}</span>}
                  <button onClick={()=>setNvItems(p=>p.filter((_,i)=>i!==idx))} style={{background:"none",border:"none",cursor:"pointer",color:"#d1d5db",fontSize:14,padding:2,flexShrink:0}}><i className="ti ti-x"/></button>
                </div>
                <PhotoStrip photos={item.photos}/>
              </div>)}
            </div>}
            <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:16}}>
              <button onClick={()=>setView("dashboard")} style={{padding:"9px 18px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer",fontSize:14,fontWeight:500}}>Cancelar</button>
              <button onClick={submitVisit} disabled={!nvStudId} style={{padding:"9px 24px",background:nvStudId?"#8b5cf6":"#f3f4f6",color:nvStudId?"#fff":"#9ca3af",border:"none",borderRadius:8,cursor:nvStudId?"pointer":"not-allowed",fontSize:14,fontWeight:700}}>Guardar gestión</button>
            </div>
          </div>
        )}

        {view==="editvisit"&&editVid&&(
          <div style={{padding:"28px",maxWidth:700}}>
            <button onClick={()=>setView("student")} style={{background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:13,padding:"0 0 14px",display:"flex",alignItems:"center",gap:4,fontWeight:500}}>
              <i className="ti ti-arrow-left" style={{fontSize:13}}/> {editStud?.name||"Volver"}
            </button>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:22}}>
              <h1 style={{fontSize:22,fontWeight:700,margin:0,color:"#111827"}}>Editar Gestión</h1>
              <div style={{display:"flex",gap:8,flexShrink:0}}>
                <button onClick={()=>setView("student")} style={{padding:"8px 16px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer",fontSize:13}}>Cancelar</button>
                <button onClick={saveEdit} style={{padding:"8px 18px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:700}}>Guardar cambios</button>
              </div>
            </div>
            <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"20px",marginBottom:20}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
                <div>
                  <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Fecha</label>
                  <input type="date" value={editDate} onChange={e=>setEditDate(e.target.value)} style={bInp}/>
                </div>
                <div style={{display:"flex",alignItems:"flex-end",paddingBottom:2}}>
                  <div style={{fontSize:12,color:"#9ca3af"}}>Alumno: <span style={{color:"#374151",fontWeight:600}}>{editStud?.name}</span></div>
                </div>
              </div>
              <div style={{marginBottom:14}}>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Notas generales</label>
                <BulletNotes bullets={editBullets} onChange={setEditBullets}/>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Fotos</label>
                <PhotoPicker photos={editPhotos} onChange={setEditPhotos}/>
              </div>
            </div>
            <div style={{fontSize:12,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:12}}>Tareas ({editItems.filter(i=>!editDelIds.has(i.id)).length})</div>
            {editItems.filter(i=>!editDelIds.has(i.id)).map(item=>(
              <div key={item.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:"12px 14px",marginBottom:8}}>
                {editItemId===item.id?<div>
                  <InlineTA type={eiType} setType={setEiType} liceo={eiLiceo} setLiceo={setEiLiceo} allFormLiceos={allFormLiceos} allTypeOptions={allTypeOptions} {...sharedTA}/>
                  <input value={eiDesc} onChange={e=>setEiDesc(e.target.value)} onKeyDown={e=>e.key==="Enter"&&saveEditItem(item.id)} style={{...sInp,marginBottom:10}} placeholder="Descripción…"/>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                    <div>
                      <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Prioridad</label>
                      <select value={eiPriority} onChange={e=>setEiPriority(e.target.value)} style={sInp}>
                        <option value="alta">🔴 Alta</option><option value="media">🟡 Media</option><option value="baja">🔵 Baja</option>
                      </select>
                    </div>
                    <div>
                      <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>Alerta para el día</label>
                      <input type="date" value={eiAlertDate} onChange={e=>setEiAlertDate(e.target.value)} style={sInp}/>
                    </div>
                  </div>
                  <div style={{marginBottom:10}}>
                    <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:5,fontWeight:500}}>Fotos</label>
                    <PhotoPicker photos={eiPhotos} onChange={setEiPhotos} compact/>
                  </div>
                  <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                    <button onClick={()=>setEditItemId(null)} style={{padding:"5px 12px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:7,cursor:"pointer",fontSize:12}}>Cancelar</button>
                    <button onClick={()=>saveEditItem(item.id)} style={{padding:"5px 14px",background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Guardar tarea</button>
                  </div>
                </div>:<div>
                  <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                    <TypeBadge type={item.type} size="sm"/>{item.liceo&&<LiceoChip liceo={item.liceo} size="sm"/>}
                    {item.priority&&item.priority!=="media"&&<PriorityBadge priority={item.priority} size="sm"/>}
                    <span style={{fontSize:13,flex:1,minWidth:120,color:"#374151"}}>{item.description}</span>
                    {item.alertDate&&<span style={{fontSize:11,color:"#6b7280",display:"flex",alignItems:"center",gap:3}}><i className="ti ti-bell" style={{fontSize:11}}/>{fmtS(item.alertDate)}</span>}
                    {item.status==="resuelto"&&<span style={{fontSize:11,color:"#166534",fontWeight:700}}>✓ Resuelto</span>}
                    <button onClick={()=>startEditItem(item)} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:13,padding:"2px 4px"}}><i className="ti ti-pencil" style={{fontSize:13}}/></button>
                    <button onClick={()=>deleteEditItem(item.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#fca5a5",fontSize:13,padding:"2px 4px"}}><i className="ti ti-trash" style={{fontSize:13}}/></button>
                  </div>
                  <PhotoStrip photos={item.photos}/>
                </div>}
              </div>
            ))}
            {showAddInEdit?<div style={{marginTop:8}}><InlineAddItem {...sharedTA} currentUserId={uid} onAdd={addEditItem} onCancel={()=>setShowAddInEdit(false)} showCancel/></div>
            :<button onClick={()=>setShowAddInEdit(true)} style={{width:"100%",padding:"10px",marginTop:8,border:"1.5px dashed #ddd6fe",background:"#f5f3ff",color:"#6d28d9",borderRadius:10,cursor:"pointer",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
              <i className="ti ti-plus" style={{fontSize:14}}/> Agregar tarea
            </button>}
            <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,paddingTop:16,borderTop:"1px solid #f3f4f6"}}>
              <button onClick={()=>setView("student")} style={{padding:"9px 18px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:8,cursor:"pointer",fontSize:14}}>Cancelar</button>
              <button onClick={saveEdit} style={{padding:"9px 24px",background:"#8b5cf6",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:14,fontWeight:700}}>Guardar cambios</button>
            </div>
          </div>
        )}

        {view==="pending"&&(
          <div style={{padding:"28px",maxWidth:760}}>
            <div style={{display:"flex",gap:0,marginBottom:22,background:"#f3f4f6",borderRadius:10,padding:4,width:"fit-content"}}>
              <button onClick={()=>setShowResolved(false)} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:!showResolved?"#fff":"transparent",color:!showResolved?"#6d28d9":"#9ca3af",boxShadow:!showResolved?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
                <i className="ti ti-clock" style={{fontSize:12,marginRight:5}}/>Abiertas ({openItems.length})
              </button>
              <button onClick={()=>setShowResolved(true)} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:showResolved?"#fff":"transparent",color:showResolved?"#166534":"#9ca3af",boxShadow:showResolved?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
                <i className="ti ti-circle-check" style={{fontSize:12,marginRight:5}}/>Resueltas ({resolvedItems.length})
              </button>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
              <div style={{fontSize:13,color:"#6b7280",display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                {(showResolved?resolvedFiltered:pendFiltered).length} tarea{(showResolved?resolvedFiltered:pendFiltered).length!==1?"s":""}
                {fLiceo&&<><span>·</span><LiceoChip liceo={fLiceo} size="sm"/></>}
                {fStud&&<><span>·</span><span>{students.find(s=>s.id===fStud)?.name}</span></>}
                {fPriority&&<><span>·</span><PriorityBadge priority={fPriority} size="sm"/></>}
              </div>
              {(fLiceo||fType||fStud||fPriority)&&<button onClick={()=>{setFLiceo("");setFType("");setFStud("");setFPriority("");}}
                style={{fontSize:12,padding:"5px 12px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",borderRadius:7,cursor:"pointer",display:"flex",alignItems:"center",gap:4,fontWeight:500}}>
                <i className="ti ti-x" style={{fontSize:12}}/> Limpiar
              </button>}
            </div>
            {(showResolved?allResolvedLiceos:allLiceos).length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12,paddingBottom:12,borderBottom:"1px solid #f3f4f6"}}>
              <button onClick={()=>setFLiceo("")} style={{fontSize:12,padding:"4px 13px",borderRadius:20,border:`1px solid ${!fLiceo?"#8b5cf6":"#e5e7eb"}`,background:!fLiceo?"#f5f3ff":"transparent",color:!fLiceo?"#6d28d9":"#6b7280",cursor:"pointer",fontWeight:!fLiceo?700:400}}>Todos</button>
              {(showResolved?allResolvedLiceos:allLiceos).map(liceo=>{const p=liceoPal(liceo);const active=fLiceo===liceo;return(
                <button key={liceo} onClick={()=>setFLiceo(active?"":liceo)} style={{fontSize:12,padding:"4px 13px",borderRadius:20,border:`1px solid ${active?p.border:"#e5e7eb"}`,background:active?p.bg:"transparent",color:active?p.fg:"#6b7280",cursor:"pointer",fontWeight:active?700:400}}>{liceo}</button>
              );})}
            </div>}
            {!showResolved&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:20}}>
              <select value={fType} onChange={e=>setFType(e.target.value)} style={sInp}>
                <option value="">Todos los tipos</option>
                {allTypeOptions.filter(t=>allItemTypes.includes(t.key)).map(t=><option key={t.key} value={t.key}>{t.label}</option>)}
              </select>
              <select value={fStud} onChange={e=>setFStud(e.target.value)} style={sInp}>
                <option value="">Todos los alumnos</option>
                {students.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <select value={fPriority} onChange={e=>setFPriority(e.target.value)} style={sInp}>
                <option value="">Toda prioridad</option>
                <option value="alta">🔴 Alta</option><option value="media">🟡 Media</option><option value="baja">🔵 Baja</option>
              </select>
            </div>}
            {showResolved&&<div style={{marginBottom:20}}>
              <select value={fStud} onChange={e=>setFStud(e.target.value)} style={{...sInp,maxWidth:280}}>
                <option value="">Todos los alumnos</option>
                {students.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </div>}
            {(showResolved?resolvedFiltered:pendFiltered).length===0?
              <div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>
                {showResolved?"No hay tareas resueltas con estos filtros.":(openItems.length===0?"¡Sin tareas abiertas! Todo al día. ✓":"Sin resultados con los filtros aplicados.")}
              </div>
            :(showResolved?resolvedFiltered:pendFiltered).map(item=>{
              const stud=students.find(s=>s.id===item.locationId);const v=visits.find(v=>v.id===item.visitId);
              return <ItemCard key={item.id} item={item} studentName={stud?.name} visitDate={v?.date}
                onResolve={!showResolved?resolveItem:undefined} onAddComment={addComment} onDeleteComment={deleteComment} onEditItem={!showResolved?editItemText:undefined}/>;
            })}
          </div>
        )}

        {/* ASIGNADAS VIEW */}
        {view==="asignadas"&&(()=>{
          const assignedToMe=items.filter(i=>i.assignedBy);
          const assignedByMe=assignedOutItems;
          const [tab,setTab]=useState("porme");
          return <div style={{padding:"28px",maxWidth:760}}>
            <h1 style={{fontSize:22,fontWeight:700,margin:"0 0 20px",color:"#111827"}}>Tareas Asignadas</h1>

            {/* Tabs */}
            <div style={{display:"flex",gap:0,marginBottom:24,background:"#f3f4f6",borderRadius:10,padding:4,width:"fit-content"}}>
              <button onClick={()=>setTab("porme")} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab==="porme"?"#fff":"transparent",color:tab==="porme"?"#6d28d9":"#9ca3af",boxShadow:tab==="porme"?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
                <i className="ti ti-arrow-up-right" style={{fontSize:12,marginRight:5}}/>Por mí ({assignedByMe.length})
              </button>
              <button onClick={()=>setTab("ami")} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab==="ami"?"#fff":"transparent",color:tab==="ami"?"#6d28d9":"#9ca3af",boxShadow:tab==="ami"?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
                <i className="ti ti-arrow-down-left" style={{fontSize:12,marginRight:5}}/>A mí ({assignedToMe.length})
              </button>
            </div>

            {/* Asignadas POR mí */}
            {tab==="porme"&&(<>
              {assignedByMe.length===0?<div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>No asignaste tareas a otros todavía.</div>
              :USERS.filter(u=>u.id!==uid).map(u=>{
                const uItems=assignedByMe.filter(i=>i.assignedTo===u.id);
                if(!uItems.length)return null;
                const openN=uItems.filter(i=>i.status==="abierto").length;
                return <div key={u.id} style={{marginBottom:24}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                    <div style={{width:28,height:28,borderRadius:"50%",background:u.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff"}}>{u.avatar}</div>
                    <span style={{fontSize:15,fontWeight:700,color:"#111827"}}>{u.name}</span>
                    {openN>0&&<span style={{fontSize:11,background:"#fef3c7",color:"#92400e",padding:"1px 8px",borderRadius:10,fontWeight:700,border:"1px solid #fcd34d"}}>{openN} pendiente{openN!==1?"s":""}</span>}
                    {uItems.filter(i=>i.status==="resuelto").length>0&&<span style={{fontSize:11,background:"#f0fdf4",color:"#166534",padding:"1px 8px",borderRadius:10,fontWeight:700,border:"1px solid #86efac"}}>✓ {uItems.filter(i=>i.status==="resuelto").length} resuelto{uItems.filter(i=>i.status==="resuelto").length!==1?"s":""}</span>}
                  </div>
                  {uItems.sort((a,b)=>{const po={alta:0,media:1,baja:2};return (po[a.priority]??1)-(po[b.priority]??1);}).map(item=>(
                    <ItemCard key={item.id} item={{...item,assignedBy:null}}
                      studentName={item.studentName} visitDate={item.createdAt}
                      onResolve={resolveAssignedItem} onAddComment={addCommentAssigned}
                      onDeleteComment={deleteCommentAssigned} onEditItem={editAssignedItemText}
                      onDelete={deleteAssignedItem}/>
                  ))}
                </div>;
              })}
            </>)}

            {/* Asignadas A mí */}
            {tab==="ami"&&(<>
              {assignedToMe.length===0?<div style={{textAlign:"center",padding:"48px",color:"#9ca3af",fontSize:14,background:"#f9fafb",borderRadius:12,border:"1px dashed #e5e7eb"}}>Nadie te asignó tareas todavía.</div>
              :[...new Set(assignedToMe.map(i=>i.assignedBy))].map(assignerName=>{
                const aItems=assignedToMe.filter(i=>i.assignedBy===assignerName);
                const assigner=USERS.find(u=>u.name===assignerName);
                return <div key={assignerName} style={{marginBottom:24}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                    {assigner&&<div style={{width:28,height:28,borderRadius:"50%",background:assigner.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff"}}>{assigner.avatar}</div>}
                    <span style={{fontSize:15,fontWeight:700,color:"#111827"}}>De {assignerName}</span>
                    <span style={{fontSize:11,color:"#6b7280"}}>{aItems.filter(i=>i.status==="abierto").length} pendiente{aItems.filter(i=>i.status==="abierto").length!==1?"s":""}</span>
                  </div>
                  {aItems.sort((a,b)=>{const po={alta:0,media:1,baja:2};return (po[a.priority]??1)-(po[b.priority]??1);}).map(item=>{
                    const v=visits.find(v=>v.id===item.visitId);
                    return <ItemCard key={item.id} item={item}
                      studentName={item.studentName||students.find(s=>s.id===item.locationId)?.name}
                      visitDate={v?.date||item.createdAt}
                      onResolve={resolveItem} onAddComment={addComment}
                      onDeleteComment={deleteComment} onEditItem={editItemText}/>;
                  })}
                </div>;
              })}
            </>)}
          </div>;
        })()}

      </main>
    </div>
  );
}
