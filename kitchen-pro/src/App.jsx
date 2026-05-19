import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import Holidays from 'date-holidays'


// ═══ TOAST SİSTEMİ ═══
window._toastQueue=window._toastQueue||[];
window._toastListeners=window._toastListeners||[];
window.toast=(message,type="info",duration=3000)=>{
  const id=Date.now()+Math.random();
  const t={id,message,type,duration};
  window._toastQueue.push(t);
  window._toastListeners.forEach(fn=>fn(t));
};
window.toast.success=(m,d)=>window.toast(m,"success",d);
window.toast.error=(m,d)=>window.toast(m,"error",d||4000);
window.toast.info=(m,d)=>window.toast(m,"info",d);

const ToastContainer=()=>{
  const[items,setItems]=useState([]);
  useEffect(()=>{
    const fn=(t)=>{setItems(p=>[...p,t]);setTimeout(()=>setItems(p=>p.filter(i=>i.id!==t.id)),t.duration);};
    window._toastListeners.push(fn);
    return()=>{window._toastListeners=window._toastListeners.filter(f=>f!==fn);};
  },[]);
  if(!items.length)return null;
  const colors={success:{bg:"#d1fae5",bo:"#10b981",fg:"#065f46",icon:"✓"},error:{bg:"#fee2e2",bo:"#dc2626",fg:"#991b1b",icon:"✕"},info:{bg:"#e0e7ff",bo:"#6366f1",fg:"#3730a3",icon:"ℹ"}};
  return <div style={{position:"fixed",top:"max(20px,env(safe-area-inset-top))",left:"50%",transform:"translateX(-50%)",zIndex:10000,display:"flex",flexDirection:"column",gap:8,maxWidth:"calc(100% - 32px)",width:"min(440px,calc(100% - 32px))",pointerEvents:"none"}}>
    {items.map(t=>{const c=colors[t.type]||colors.info;return <div key={t.id} style={{background:c.bg,border:`1px solid ${c.bo}`,color:c.fg,padding:"12px 16px",borderRadius:12,fontSize:13,fontWeight:500,boxShadow:"0 4px 16px rgba(0,0,0,0.12)",display:"flex",alignItems:"center",gap:10,pointerEvents:"auto",animation:"toastIn 0.25s ease-out"}}>
      <span style={{fontSize:16,fontWeight:700}}>{c.icon}</span>
      <span style={{flex:1,lineHeight:1.4,wordBreak:"break-word"}}>{t.message}</span>
    </div>;})}

  </div>;
};



// ═══ STORAGE ═══
const LS={get:(k,d)=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(e){return d}},set:(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}};
// ═══ PRO STORAGE PREFIX ═══
(()=>{
  // Eski tk_ key'leri kmp_'ye taşı
  const toMigrate=["tk_recipes","tk_stock","tk_invoices","tk_lang","tk_dark",
    "tk_menus","tk_expenses","tk_storage","tk_productions","tk_reportcats",
    "tk_profile","tk_traceability","tk_lots","tk_trackedings","tk_resethour",
    "tk_organizations","tk_storagechecks","tk_menutemplates","tk_notifsettings",
    "tk_botmessages","tk_caloriedb","tk_printers"];
  for(const k of toMigrate){
    const val=localStorage.getItem(k);
    if(val!==null){
      const newKey="kmp_"+k.slice(3);
      if(!localStorage.getItem(newKey))localStorage.setItem(newKey,val);
      localStorage.removeItem(k);
    }
  }
  // km_ ve kmc_ verilerini temizle — ama kmp_ ve sb- verilerine dokunma
  const allKeys=[];
  for(let i=0;i<localStorage.length;i++){
    const k=localStorage.key(i);
    if(k)allKeys.push(k);
  }
  allKeys.forEach(k=>{
    if((k.startsWith("km_")||k.startsWith("kmc_"))&&!k.startsWith("kmp_"))localStorage.removeItem(k);
  });
})();

const SK={
  key:"kmp_apikey",recipes:"kmp_recipes",stock:"kmp_stock",invoices:"kmp_invoices",
  lang:"kmp_lang",dark:"kmp_dark",logs:"kmp_logs",menus:"kmp_menus",
  expenses:"kmp_expenses",storage:"kmp_storage",productions:"kmp_productions",
  reportCats:"kmp_reportcats",profile:"kmp_profile",traceability:"kmp_traceability",
  lots:"kmp_lots",trackedIngs:"kmp_trackedings",resetHour:"kmp_resethour",
  organizations:"kmp_organizations",storageChecks:"kmp_storagechecks",
  menuTemplates:"kmp_menutemplates",conversations:"kmp_conversations",
  activeConvId:"kmp_activeconv",notifSettings:"kmp_notifsettings",
  botMessages:"kmp_botmessages",calorieDB:"kmp_caloriedb",printers:"kmp_printers"
};

// Default organizasyonlar (numune etiketi için)
const DEF_ORGANIZATIONS=["Restoran","Ballroom","Bellini","Gazebo","Personel Yemekhanesi"];

const DEF_PROFILE={fullName:"",workplace:"",department:"",role:"",country:"",sector:"",unitSystem:"metric"};

// Sektör ve ülke listeleri (Pro modülünde de kullanılır)
const SECTORS=[
  {id:"hotel",emoji:"🏨",tr:"Otel",en:"Hotel",ru:"Отель",es:"Hotel",de:"Hotel",fr:"Hôtel",zh:"酒店",ar:"فندق"},
  {id:"restaurant",emoji:"🍽",tr:"Restoran",en:"Restaurant",ru:"Ресторан",es:"Restaurante",de:"Restaurant",fr:"Restaurant",zh:"餐厅",ar:"مطعم"},
  {id:"cafe",emoji:"☕",tr:"Kafe",en:"Café",ru:"Кафе",es:"Café",de:"Café",fr:"Café",zh:"咖啡馆",ar:"مقهى"},
  {id:"bakery",emoji:"🥐",tr:"Pastane / Fırın",en:"Bakery / Pastry",ru:"Пекарня",es:"Panadería",de:"Bäckerei",fr:"Boulangerie",zh:"面包店",ar:"مخبز"},
  {id:"catering",emoji:"🍱",tr:"Catering",en:"Catering",ru:"Кейтеринг",es:"Catering",de:"Catering",fr:"Traiteur",zh:"餐饮服务",ar:"تموين"},
  {id:"cloudkitchen",emoji:"☁️",tr:"Bulut Mutfak",en:"Cloud Kitchen",ru:"Облачная кухня",es:"Cocina en la Nube",de:"Cloud-Küche",fr:"Cuisine Cloud",zh:"云厨房",ar:"مطبخ سحابي"},
  {id:"foodtruck",emoji:"🚚",tr:"Food Truck",en:"Food Truck",ru:"Фудтрак",es:"Food Truck",de:"Food Truck",fr:"Food Truck",zh:"餐车",ar:"شاحنة طعام"},
  {id:"canteen",emoji:"🍴",tr:"Yemekhane / Toplu",en:"Canteen / Mass",ru:"Столовая",es:"Comedor",de:"Kantine",fr:"Cantine",zh:"食堂",ar:"مطعم جماعي"},
  {id:"hospital",emoji:"🏥",tr:"Hastane Mutfağı",en:"Hospital Kitchen",ru:"Больничная кухня",es:"Cocina Hospitalaria",de:"Krankenhausküche",fr:"Cuisine Hôpital",zh:"医院厨房",ar:"مطبخ مستشفى"},
  {id:"school",emoji:"🏫",tr:"Okul Mutfağı",en:"School Kitchen",ru:"Школьная кухня",es:"Cocina Escolar",de:"Schulküche",fr:"Cuisine Scolaire",zh:"学校厨房",ar:"مطبخ مدرسة"},
  {id:"chain",emoji:"🏢",tr:"Zincir / Franchise",en:"Chain / Franchise",ru:"Сеть",es:"Cadena",de:"Kette",fr:"Chaîne",zh:"连锁",ar:"سلسلة"},
  {id:"other",emoji:"❓",tr:"Diğer",en:"Other",ru:"Другое",es:"Otro",de:"Andere",fr:"Autre",zh:"其他",ar:"أخرى"}
];

const COUNTRIES=[
  {code:"TR",name:"Türkiye",flag:"🇹🇷"},{code:"US",name:"United States",flag:"🇺🇸"},
  {code:"GB",name:"United Kingdom",flag:"🇬🇧"},{code:"DE",name:"Deutschland",flag:"🇩🇪"},
  {code:"FR",name:"France",flag:"🇫🇷"},{code:"IT",name:"Italia",flag:"🇮🇹"},
  {code:"ES",name:"España",flag:"🇪🇸"},{code:"NL",name:"Nederland",flag:"🇳🇱"},
  {code:"BE",name:"Belgique",flag:"🇧🇪"},{code:"CH",name:"Schweiz",flag:"🇨🇭"},
  {code:"AT",name:"Österreich",flag:"🇦🇹"},{code:"SE",name:"Sverige",flag:"🇸🇪"},
  {code:"NO",name:"Norge",flag:"🇳🇴"},{code:"DK",name:"Danmark",flag:"🇩🇰"},
  {code:"FI",name:"Suomi",flag:"🇫🇮"},{code:"PL",name:"Polska",flag:"🇵🇱"},
  {code:"GR",name:"Ελλάδα",flag:"🇬🇷"},{code:"PT",name:"Portugal",flag:"🇵🇹"},
  {code:"RU",name:"Россия",flag:"🇷🇺"},{code:"UA",name:"Україна",flag:"🇺🇦"},
  {code:"AZ",name:"Azərbaycan",flag:"🇦🇿"},{code:"KZ",name:"Қазақстан",flag:"🇰🇿"},
  {code:"AE",name:"الإمارات",flag:"🇦🇪"},{code:"SA",name:"السعودية",flag:"🇸🇦"},
  {code:"QA",name:"قطر",flag:"🇶🇦"},{code:"KW",name:"الكويت",flag:"🇰🇼"},
  {code:"EG",name:"مصر",flag:"🇪🇬"},{code:"MA",name:"المغرب",flag:"🇲🇦"},
  {code:"JP",name:"日本",flag:"🇯🇵"},{code:"KR",name:"한국",flag:"🇰🇷"},
  {code:"CN",name:"中国",flag:"🇨🇳"},{code:"IN",name:"भारत",flag:"🇮🇳"},
  {code:"AU",name:"Australia",flag:"🇦🇺"},{code:"CA",name:"Canada",flag:"🇨🇦"},
  {code:"BR",name:"Brasil",flag:"🇧🇷"},{code:"MX",name:"México",flag:"🇲🇽"},
  {code:"AR",name:"Argentina",flag:"🇦🇷"},{code:"CL",name:"Chile",flag:"🇨🇱"},
  {code:"ZA",name:"South Africa",flag:"🇿🇦"},{code:"OTHER",name:"Other / Diğer",flag:"🌍"}
];

// Default rapor kategorileri
const DEF_REPORT_CATS=[
  {id:"production",name:"Üretim Raporları",icon:"📋"},
  {id:"fire",name:"Fire Kayıtları",icon:"🗑"},
  {id:"samples",name:"Numuneler",icon:"🧪"}
];

// Default işletme giderleri
const DEF_EXPENSES={fixed:[],personnel:[],monthlyPortions:1000};

// Default depolama alanları
const DEF_STORAGE=[
  {id:"fridge1",name:"Buzdolabı",type:"fridge",temp:4,capacity:null},
  {id:"freezer1",name:"Dondurucu",type:"freezer",temp:-18,capacity:null},
  {id:"dry1",name:"Kuru Depo",type:"dry",temp:20,capacity:null},
  {id:"bm1",name:"Bain-Marie",type:"hot",temp:65,capacity:null}
];

const STORAGE_TYPES=[
  {id:"fridge",l:"Buzdolabı",tr:"Buzdolabı",en:"Refrigerator",ru:"Холодильник",es:"Nevera",de:"Kühlschrank",fr:"Réfrigérateur",zh:"冰箱",ar:"ثلاجة",icon:"🧊",defaultTemp:4},
  {id:"freezer",l:"Dondurucu",tr:"Dondurucu",en:"Freezer",ru:"Морозильник",es:"Congelador",de:"Tiefkühlschrank",fr:"Congélateur",zh:"冷冻柜",ar:"فريزر",icon:"❄️",defaultTemp:-18},
  {id:"dry",l:"Kuru Depo",tr:"Kuru Depo",en:"Dry Storage",ru:"Сухой склад",es:"Almacén seco",de:"Trockenlager",fr:"Réserve sèche",zh:"干货仓",ar:"مخزن جاف",icon:"📦",defaultTemp:20},
  {id:"hot",l:"Sıcak Tutma",tr:"Sıcak Tutma",en:"Hot Holding",ru:"Горячее хранение",es:"Mantenimiento caliente",de:"Warmhaltung",fr:"Maintien chaud",zh:"保温",ar:"حفظ ساخن",icon:"🔥",defaultTemp:65},
  {id:"room",l:"Oda Sıcaklığı",tr:"Oda Sıcaklığı",en:"Room Temperature",ru:"Комнатная температура",es:"Temperatura ambiente",de:"Raumtemperatur",fr:"Température ambiante",zh:"室温",ar:"درجة حرارة الغرفة",icon:"🌡",defaultTemp:22}
];
const storageTypeL=(s,lang)=>s&&(s[lang]||s.tr||s.l)||"";

// ═══ ERROR LOG ═══
let LOGS=LS.get(SK.logs,[]);
const log=(type,detail,raw)=>{LOGS.push({ts:new Date().toISOString(),type,detail:String(detail).slice(0,500),raw:String(raw||"").slice(0,1500)});if(LOGS.length>100)LOGS=LOGS.slice(-100);LS.set(SK.logs,LOGS)};

// ═══ SUPABASE CONFIG ═══
const SUPABASE_URL="https://tbacctbscojfknqttfly.supabase.co";
const SUPABASE_ANON_KEY="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRiYWNjdGJzY29qZmtucXR0Zmx5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1MjQ0NTIsImV4cCI6MjA5NDEwMDQ1Mn0.HqqMzmJICuso2ZAdotRho0Jgn36yR_nVxdej-teGru0";

// ═══ INLINE QR ENCODER (CDN bağımsız) ═══
// Minimal QR Code generator - no external dependencies
const QREncoder=(()=>{
  const ALPHANUMERIC="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:";
  const isAlphanumeric=s=>s.split("").every(c=>ALPHANUMERIC.includes(c));
  
  // Galois field multiplication
  const gfMul=(a,b)=>{
    let r=0;
    for(let i=0;i<8;i++){
      if(b&1)r^=a;
      const hbs=a&0x80;
      a=(a<<1)&0xff;
      if(hbs)a^=0x1d;
      b>>=1;
    }
    return r;
  };
  
  // Reed-Solomon error correction
  const rsEC=(data,ecCount)=>{
    const gen=Array(ecCount).fill(0);
    gen[0]=1;
    for(let i=0;i<ecCount;i++){
      const a=1<<i; // 2^i in GF
      let pw=1;
      for(let j=0;j<i;j++)pw=gfMul(pw,2);
      for(let j=ecCount-1;j>=0;j--){
        gen[j]=gfMul(gen[j],pw);
        if(j>0)gen[j]^=gen[j-1];
      }
    }
    const msg=data.slice();
    for(let i=0;i<ecCount;i++)msg.push(0);
    for(let i=0;i<data.length;i++){
      const c=msg[i];
      if(c!==0)for(let j=0;j<ecCount;j++)msg[i+j+1]^=gfMul(gen[ecCount-1-j],c);
    }
    return msg.slice(data.length);
  };

  // Encode text to QR matrix (Version 2, ECC M)
  // Returns 2D boolean array
  const encode=(text)=>{
    // Use URL-safe approach: force numeric/alphanumeric if possible
    const upper=text.toUpperCase();
    const useAlpha=isAlphanumeric(upper)&&upper.length<=47;
    
    // For complex URLs we use byte mode
    const bytes=[];
    for(let i=0;i<text.length;i++)bytes.push(text.charCodeAt(i)&0xff);
    
    // Version 2 (25x25), ECC M: 28 data codewords, 16 EC codewords
    const size=25;
    const dataCapacity=28;
    const ecCount=16;
    
    // Build data bits
    const bits=[];
    const pushBits=(val,len)=>{for(let i=len-1;i>=0;i--)bits.push((val>>i)&1)};
    
    // Mode: byte (4 bits = 0100)
    pushBits(4,4);
    // Character count (8 bits for byte mode v2)
    const len=Math.min(bytes.length,dataCapacity-2);
    pushBits(len,8);
    // Data
    for(let i=0;i<len;i++)pushBits(bytes[i],8);
    // Terminator
    while(bits.length<dataCapacity*8&&bits.length%8!==0||bits.length<dataCapacity*8-4)bits.push(0);
    // Padding
    const pads=[0xEC,0x11];
    let pi=0;
    while(bits.length<dataCapacity*8){const p=pads[pi%2];pi++;pushBits(p,8);}
    
    // Convert to codewords
    const codewords=[];
    for(let i=0;i<dataCapacity;i++){
      let b=0;
      for(let j=0;j<8;j++)b=(b<<1)|(bits[i*8+j]||0);
      codewords.push(b);
    }
    
    // Error correction
    const ec=rsEC(codewords,ecCount);
    const allCW=[...codewords,...ec];
    
    // Build matrix
    const mat=Array(size).fill(null).map(()=>Array(size).fill(null));
    const func=Array(size).fill(null).map(()=>Array(size).fill(false));
    
    const setFunc=(r,c,v)=>{if(r>=0&&r<size&&c>=0&&c<size){mat[r][c]=v?1:0;func[r][c]=true;}};
    
    // Finder patterns
    const finder=(r,c)=>{
      for(let i=-1;i<=7;i++)for(let j=-1;j<=7;j++){
        if(i<0||i>6||j<0||j>6)setFunc(r+i,c+j,0);
        else if(i===0||i===6||j===0||j===6)setFunc(r+i,c+j,1);
        else if(i>=2&&i<=4&&j>=2&&j<=4)setFunc(r+i,c+j,1);
        else setFunc(r+i,c+j,0);
      }
    };
    finder(0,0);finder(0,size-7);finder(size-7,0);
    
    // Timing patterns
    for(let i=8;i<size-8;i++){setFunc(6,i,i%2===0);setFunc(i,6,i%2===0);}
    
    // Dark module
    setFunc(size-8,8,1);
    
    // Alignment pattern (v2: one at 18,18)
    const ap=(r,c)=>{
      for(let i=-2;i<=2;i++)for(let j=-2;j<=2;j++){
        if(Math.abs(i)===2||Math.abs(j)===2)setFunc(r+i,c+j,1);
        else if(i===0&&j===0)setFunc(r+i,c+j,1);
        else setFunc(r+i,c+j,0);
      }
    };
    ap(18,18);
    
    // Format info (mask 0, ECC M = 00)
    const fmt=[1,1,1,0,1,1,1,1,1,0,0,0,1,0,0];
    const fmtPos=[[8,0],[8,1],[8,2],[8,3],[8,4],[8,5],[8,7],[8,8],[7,8],[5,8],[4,8],[3,8],[2,8],[1,8],[0,8]];
    fmtPos.forEach(([r,c],i)=>setFunc(r,c,fmt[i]));
    setFunc(size-7,8,fmt[6]);
    for(let i=0;i<7;i++)setFunc(size-1-i,8,fmt[i]);
    for(let i=0;i<8;i++)setFunc(8,size-8+i,fmt[14-i]);
    
    // Data placement
    let bi=0;
    const allBits=[];
    for(const cw of allCW)for(let i=7;i>=0;i--)allBits.push((cw>>i)&1);
    
    let up=true;
    for(let col=size-1;col>0;col-=2){
      if(col===6)col=5;
      for(let row=0;row<size;row++){
        const r=up?size-1-row:row;
        for(let dc=0;dc<2;dc++){
          const c=col-dc;
          if(!func[r][c]){
            mat[r][c]=bi<allBits.length?allBits[bi]^0:0; // mask 0: (i+j)%2==0
            if((r+c)%2===0)mat[r][c]^=1;
            bi++;
          }
        }
      }
      up=!up;
    }
    
    // Fill unfilled
    for(let r=0;r<size;r++)for(let c=0;c<size;c++)if(mat[r][c]===null)mat[r][c]=0;
    
    return mat;
  };
  
  // Draw QR to canvas
  const draw=(canvas,text,fg="#000",bg="#fff")=>{
    try{
      const mat=encode(text);
      const size=mat.length;
      const scale=Math.floor(canvas.width/( size+8));
      const off=Math.floor((canvas.width-scale*(size+8))/2)+scale*4;
      const ctx=canvas.getContext("2d");
      ctx.fillStyle=bg;ctx.fillRect(0,0,canvas.width,canvas.height);
      ctx.fillStyle=fg;
      for(let r=0;r<size;r++)for(let c=0;c<size;c++)
        if(mat[r][c])ctx.fillRect(off+c*scale,off+r*scale,scale,scale);
      return true;
    }catch(e){
      // Fallback: takip no yaz
      const ctx=canvas.getContext("2d");
      ctx.fillStyle=bg;ctx.fillRect(0,0,canvas.width,canvas.height);
      ctx.fillStyle=fg;ctx.font=`bold ${Math.floor(canvas.width/12)}px monospace`;
      ctx.textAlign="center";
      const parts=text.split("?prod=");
      ctx.fillText(parts[1]||text.slice(-12),canvas.width/2,canvas.height/2);
      return false;
    }
  };
  
  return{draw};
})();
let supabase=null;
const initSupabase=()=>{
  if(supabase)return supabase;
  if(true){
    supabase=createClient(SUPABASE_URL,SUPABASE_ANON_KEY,{
      auth:{storageKey:"kmp-auth",storage:window.localStorage}
    });
    return supabase;
  }
  return null;
};


// ═══ STORAGE UPLOAD ═══
const uploadFile=async(file,teamId,folder,onProgress=null,userId=null)=>{
  const sb=initSupabase();if(!sb)throw new Error("Supabase yüklenemedi");
  const ext=file.name.split(".").pop().toLowerCase();
  const allowed=["jpg","jpeg","png","gif","webp","pdf","xlsx","xls","docx","txt","mp4","mov"];
  if(!allowed.includes(ext))throw new Error("Desteklenmeyen dosya türü: "+ext);
  if(file.size>50*1024*1024)throw new Error("Dosya 50MB'dan büyük olamaz");
  const ts=Date.now();
  const path=userId
    ?`users/${userId}/${folder}/${ts}_${file.name.replace(/[^a-zA-Z0-9._-]/g,"_")}`
    :`teams/${teamId}/${folder}/${ts}_${file.name.replace(/[^a-zA-Z0-9._-]/g,"_")}`;
  const{data,error}=await sb.storage.from("tulpar-storage").upload(path,file,{
    cacheControl:"3600",upsert:false,
    contentType:file.type||"application/octet-stream"
  });
  if(error)throw error;
  const{data:{publicUrl}}=sb.storage.from("tulpar-storage").getPublicUrl(path);
  return{path,url:publicUrl,name:file.name,size:file.size,type:file.type,ext};
};

const getFileUrl=(path)=>{
  const sb=initSupabase();if(!sb)return null;
  const{data:{publicUrl}}=sb.storage.from("tulpar-storage").getPublicUrl(path);
  return publicUrl;
};

const deleteFile=async(path)=>{
  const sb=initSupabase();if(!sb)return;
  await sb.storage.from("tulpar-storage").remove([path]);
};

const isImage=(ext)=>["jpg","jpeg","png","gif","webp"].includes((ext||"").toLowerCase());
const isVideo=(ext)=>["mp4","mov","avi"].includes((ext||"").toLowerCase());
const isPDF=(ext)=>ext==="pdf";

// ═══ AI MODELS ═══
const MODELS={
  sonnet:"claude-sonnet-4-5-20250929",  // Reçete OCR - kalite kritik
  haiku:"claude-haiku-4-5-20251001"      // Fatura OCR + Sohbet - hız/maliyet
};

// ═══ API ═══
// ═══ CLOUDFLARE WORKER PROXY ═══
const WORKER_URL="https://kitchen-manager-ai.aligny0.workers.dev";
const WORKER_AUTH_TOKEN="km_2026_x9k4n7j2p8r5t1w6";

async function callAI(apiKey,system,userContent,modelKey){
  // apiKey parametresi artık kullanılmıyor (geriye dönük uyumluluk için duruyor)
  // API key Cloudflare Worker'da güvenli şekilde saklı
  const model=MODELS[modelKey||"sonnet"]||MODELS.sonnet;
  const messages=[{role:"user",content:userContent}];
  let raw="";
  try{
    const res=await fetch(WORKER_URL,{
      method:"POST",
      headers:{
        "Content-Type":"application/json",
        "X-Auth-Token":WORKER_AUTH_TOKEN
      },
      body:JSON.stringify({model,max_tokens:2000,system,messages})
    });
    raw=await res.text();
    if(!res.ok){log("api-http",`HTTP ${res.status}`,raw);throw new Error(`HTTP ${res.status}: ${raw.slice(0,200)}`)}
    const data=JSON.parse(raw);
    if(data.error){log("api-err",data.error.message,raw);throw new Error(data.error.message)}
    const text=(data.content||[]).map(c=>c.text||"").join("");
    if(!text.trim())throw new Error("Boş yanıt");
    return text;
  }catch(e){log("api-call",e.message,raw);throw e}
}

function parseJSON(raw){
  let c=raw.replace(/```json\s*/gi,"").replace(/```\s*/gi,"").trim();
  let d=0,s=-1;
  for(let i=0;i<c.length;i++){
    if(c[i]==="{"){if(s===-1)s=i;d++}
    else if(c[i]==="}"){ d--;if(d===0&&s!==-1)return JSON.parse(c.slice(s,i+1))}
  }
  return JSON.parse(c);
}

// ═══ IMAGE ═══
function resizeImg(dataUrl){
  return new Promise(r=>{
    const img=new Image();
    img.onload=()=>{
      let w=img.width,h=img.height;const max=1200;
      if(w>max){h=h*max/w;w=max}
      const c=document.createElement("canvas");c.width=w;c.height=h;
      c.getContext("2d").drawImage(img,0,0,w,h);
      const out=c.toDataURL("image/jpeg",0.85);
      r({base64:out.split(",")[1],mediaType:"image/jpeg"});
    };
    img.onerror=()=>r({base64:dataUrl.split(",")[1],mediaType:"image/jpeg"});
    img.src=dataUrl;
  });
}

// ═══ DATA ═══
// Mutfak ünvanları
const KITCHEN_TITLES=[
  "Executive Chef","Head Chef","Sous Chef","Chef de Partie",
  "Pastry Chef","Demi Chef","Commis Chef","Komi",
  "Garde Manger","Saucier","Poissonnier","Rôtisseur",
  "Entremetier","Tournant","Aboyeur",
  "F&B Manager","Kitchen Manager","Barista","Bartender","Waiter"
];

// ═══ HİYERARŞİK ROL SİSTEMİ (Pro) ═══
// Kademe yüksek = otorite yüksek. Sadece daha düşük kademeye görev verilebilir.
const ROLE_HIERARCHY={
  // En üst yönetim — sadece Pro'da
  "executive_chef":{level:100,label:{tr:"Executive Chef (Baş Aşçı)",en:"Executive Chef"},icon:"👑",app:"pro"},
  "hotel_manager":{level:98,label:{tr:"Hotel Manager (Otel Müdürü)",en:"Hotel Manager"},icon:"🏨",app:"pro"},
  "fb_manager":{level:95,label:{tr:"F&B Manager (F&B Müdürü)",en:"F&B Manager"},icon:"💼",app:"pro"},
  "restaurant_manager":{level:92,label:{tr:"Restaurant Manager (Restoran Müdürü)",en:"Restaurant Manager"},icon:"🏛",app:"pro"},
  "bar_manager":{level:88,label:{tr:"Bar Manager (Bar Şefi)",en:"Bar Manager"},icon:"🍸",app:"pro"},
  "food_engineer":{level:85,label:{tr:"Food Engineer (Gıda Mühendisi)",en:"Food Engineer"},icon:"🔬",app:"pro"},
  // Orta kademe — Manager
  "chef_de_cuisine":{level:80,label:{tr:"Chef de Cuisine (Mutfak Şefi)",en:"Chef de Cuisine"},icon:"🎩",app:"manager"},
  "pastry_chef":{level:75,label:{tr:"Pastry Chef (Pastane Şefi)",en:"Pastry Chef"},icon:"🥐",app:"manager"},
  "butcher_chef":{level:72,label:{tr:"Butcher Chef (Kasaphane Şefi)",en:"Butcher Chef"},icon:"🔪",app:"manager"},
  "sous_chef":{level:70,label:{tr:"Sous Chef (Yardımcı Şef)",en:"Sous Chef"},icon:"🍳",app:"manager"},
  "chef_de_partie":{level:60,label:{tr:"Chef de Partie (Kısım Şefi)",en:"Chef de Partie"},icon:"👨‍🍳",app:"manager"},
  // Çalışan — Chef
  "demi_chef":{level:40,label:{tr:"Demi Chef (Yarı Şef)",en:"Demi Chef"},icon:"🔪",app:"chef"},
  "commis":{level:30,label:{tr:"Commis (Komi)",en:"Commis Chef"},icon:"🧑‍🍳",app:"chef"},
  "barista":{level:25,label:{tr:"Barista (Barista)",en:"Barista"},icon:"☕",app:"chef"},
  "stagiaire":{level:20,label:{tr:"Stagiaire (Stajyer)",en:"Trainee"},icon:"📚",app:"chef"},
  "waiter":{level:15,label:{tr:"Waiter (Garson)",en:"Waiter"},icon:"🍽",app:"chef"},
  "member":{level:10,label:{tr:"Member (Üye)",en:"Member"},icon:"👤",app:"chef"}
};

// Bir kullanıcı diğerine görev atayabilir mi?
const canAssignTo=(myRole,otherRole)=>{
  const me=ROLE_HIERARCHY[myRole];
  const them=ROLE_HIERARCHY[otherRole];
  if(!me||!them)return false;
  return me.level>them.level;
};

// Bir rol Pro/Manager/Chef hangi uygulama için uygun?
const getAppForRole=(role)=>ROLE_HIERARCHY[role]?.app||"chef";

// Pro yöneticileri (üst kademe — sadece Pro'da görünür/kullanılır)
const PRO_ROLES=Object.entries(ROLE_HIERARCHY).filter(([k,v])=>v.app==="pro").map(([k])=>k);
const MANAGER_ROLES=Object.entries(ROLE_HIERARCHY).filter(([k,v])=>v.app==="manager").map(([k])=>k);
const CHEF_ROLES=Object.entries(ROLE_HIERARCHY).filter(([k,v])=>v.app==="chef").map(([k])=>k);


const MAIN_CATS=[
  {id:"all",label:"Tümü",tr:"Tümü",en:"All",ru:"Все",es:"Todos",de:"Alle",fr:"Tout",zh:"全部",ar:"الكل",icon:"⊹"},
  {id:"pastry",label:"Pastane",tr:"Pastane",en:"Pastry",ru:"Кондитерская",es:"Pastelería",de:"Konditorei",fr:"Pâtisserie",zh:"糕点",ar:"المعجنات",icon:"🍰"},
  {id:"hot",label:"Sıcak Mutfak",tr:"Sıcak Mutfak",en:"Hot Kitchen",ru:"Горячий цех",es:"Cocina caliente",de:"Warmküche",fr:"Cuisine chaude",zh:"热菜",ar:"المطبخ الساخن",icon:"🍽"},
  {id:"cold",label:"Soğuk Mutfak",tr:"Soğuk Mutfak",en:"Cold Kitchen",ru:"Холодный цех",es:"Cocina fría",de:"Kaltküche",fr:"Cuisine froide",zh:"冷菜",ar:"المطبخ البارد",icon:"🥗"},
  {id:"breakfast",label:"Kahvaltı",tr:"Kahvaltı",en:"Breakfast",ru:"Завтрак",es:"Desayuno",de:"Frühstück",fr:"Petit-déjeuner",zh:"早餐",ar:"الإفطار",icon:"🍳"},
  {id:"bar",label:"Bar",tr:"Bar",en:"Bar",ru:"Бар",es:"Bar",de:"Bar",fr:"Bar",zh:"酒吧",ar:"البار",icon:"☕"}
];
const mainCatL=(c,lang)=>c&&(c[lang]||c.tr||c.label)||"";
const SUB_CATS={
  pastry:[{id:"milk",label:"Sütlü Tatlılar"},{id:"cake",label:"Kekler"},{id:"cookie",label:"Kurabiyeler"},{id:"pie",label:"Pasta & Tart"},{id:"dough",label:"Hamur İşleri"},{id:"sherbet",label:"Şerbetli"},{id:"ice",label:"Dondurma"}],
  hot:[{id:"main",label:"Ana Yemek"},{id:"soup",label:"Çorba"},{id:"sauce",label:"Sos"},{id:"garnish",label:"Garnitür"},{id:"rice",label:"Pilav"}],
  cold:[{id:"meze",label:"Meze"},{id:"salad",label:"Salata"},{id:"snack",label:"Aperatif"},{id:"coldsauce",label:"Soğuk Sos"}],
  breakfast:[{id:"egg",label:"Yumurta Çeşitleri"},{id:"pastryb",label:"Hamur İşi (Börek/Poğaça)"},{id:"spread",label:"Kahvaltılık (Reçel/Bal/Peynir)"},{id:"cereal",label:"Granola & Müsli"},{id:"brplate",label:"Tabak & Servis"}],
  bar:[{id:"hotdrink",label:"Sıcak İçecek"},{id:"colddrink",label:"Soğuk İçecek"},{id:"cocktail",label:"Kokteyl"}]
};

// VENUES — banket eklendi (madde 2)
const VENUES=[{id:"all",l:"Tümü",icon:"all"},{id:"buffet",l:"Büfe",icon:"buffet"},{id:"alacarte",l:"À la Carte",icon:"alacarte"},{id:"festival",l:"Festival",icon:"festival"},{id:"catering",l:"Catering",icon:"catering"},{id:"banquet",l:"Banket",icon:"banquet"}];
const PREPS=[{id:"all",l:"Tümü",icon:"all"},{id:"raw",l:"Çiğ",icon:"raw"},{id:"cooked",l:"Pişmiş",icon:"cooked"},{id:"chilled",l:"Soğutulan",icon:"chilled"}];
const CUISINES_DATA={
  tr:[{id:"all",l:"Tümü",icon:"all"},{id:"turkish",l:"Türk",icon:"turkish"},{id:"italian",l:"İtalyan",icon:"italian"},{id:"french",l:"Fransız",icon:"french"},{id:"japanese",l:"Japon",icon:"japanese"},{id:"mexican",l:"Meksika",icon:"mexican"},{id:"fusion",l:"Fusion",icon:"fusion"},{id:"other",l:"Diğer",icon:"other"}],
  en:[{id:"all",l:"All",icon:"all"},{id:"turkish",l:"Turkish",icon:"turkish"},{id:"italian",l:"Italian",icon:"italian"},{id:"french",l:"French",icon:"french"},{id:"japanese",l:"Japanese",icon:"japanese"},{id:"mexican",l:"Mexican",icon:"mexican"},{id:"fusion",l:"Fusion",icon:"fusion"},{id:"other",l:"Other",icon:"other"}],
  ru:[{id:"all",l:"Все",icon:"all"},{id:"turkish",l:"Турецкая",icon:"turkish"},{id:"italian",l:"Итальянская",icon:"italian"},{id:"french",l:"Французская",icon:"french"},{id:"japanese",l:"Японская",icon:"japanese"},{id:"mexican",l:"Мексиканская",icon:"mexican"},{id:"fusion",l:"Фьюжн",icon:"fusion"},{id:"other",l:"Другая",icon:"other"}],
  es:[{id:"all",l:"Todos",icon:"all"},{id:"turkish",l:"Turca",icon:"turkish"},{id:"italian",l:"Italiana",icon:"italian"},{id:"french",l:"Francesa",icon:"french"},{id:"japanese",l:"Japonesa",icon:"japanese"},{id:"mexican",l:"Mexicana",icon:"mexican"},{id:"fusion",l:"Fusión",icon:"fusion"},{id:"other",l:"Otra",icon:"other"}],
  de:[{id:"all",l:"Alle",icon:"all"},{id:"turkish",l:"Türkisch",icon:"turkish"},{id:"italian",l:"Italienisch",icon:"italian"},{id:"french",l:"Französisch",icon:"french"},{id:"japanese",l:"Japanisch",icon:"japanese"},{id:"mexican",l:"Mexikanisch",icon:"mexican"},{id:"fusion",l:"Fusion",icon:"fusion"},{id:"other",l:"Sonstige",icon:"other"}],
  fr:[{id:"all",l:"Tout",icon:"all"},{id:"turkish",l:"Turque",icon:"turkish"},{id:"italian",l:"Italienne",icon:"italian"},{id:"french",l:"Française",icon:"french"},{id:"japanese",l:"Japonaise",icon:"japanese"},{id:"mexican",l:"Mexicaine",icon:"mexican"},{id:"fusion",l:"Fusion",icon:"fusion"},{id:"other",l:"Autre",icon:"other"}],
  zh:[{id:"all",l:"全部",icon:"all"},{id:"turkish",l:"土耳其菜",icon:"turkish"},{id:"italian",l:"意大利菜",icon:"italian"},{id:"french",l:"法国菜",icon:"french"},{id:"japanese",l:"日本料理",icon:"japanese"},{id:"mexican",l:"墨西哥菜",icon:"mexican"},{id:"fusion",l:"融合菜",icon:"fusion"},{id:"other",l:"其他",icon:"other"}],
  ar:[{id:"all",l:"الكل",icon:"all"},{id:"turkish",l:"تركي",icon:"turkish"},{id:"italian",l:"إيطالي",icon:"italian"},{id:"french",l:"فرنسي",icon:"french"},{id:"japanese",l:"ياباني",icon:"japanese"},{id:"mexican",l:"مكسيكي",icon:"mexican"},{id:"fusion",l:"فيوجن",icon:"fusion"},{id:"other",l:"أخرى",icon:"other"}]
};
const getCuisines=(lang)=>CUISINES_DATA[lang]||CUISINES_DATA.tr;
const CUISINES=CUISINES_DATA.tr;
const DIFFS=[
  {id:"all",l:"Tümü",tr:"Tümü",en:"All",ru:"Все",es:"Todos",de:"Alle",fr:"Tout",zh:"全部",ar:"الكل",icon:"all"},
  {id:"easy",l:"Kolay",tr:"Kolay",en:"Easy",ru:"Легко",es:"Fácil",de:"Einfach",fr:"Facile",zh:"简单",ar:"سهل",icon:"easy"},
  {id:"medium",l:"Orta",tr:"Orta",en:"Medium",ru:"Средне",es:"Medio",de:"Mittel",fr:"Moyen",zh:"中等",ar:"متوسط",icon:"medium"},
  {id:"hard",l:"Zor",tr:"Zor",en:"Hard",ru:"Сложно",es:"Difícil",de:"Schwer",fr:"Difficile",zh:"困难",ar:"صعب",icon:"hard"}
];
const diffL=(d,lang)=>d&&(d[lang]||d.tr||d.l)||"";

// 21 ALERJEN (madde 1) — 14 AB + 7 ek
const ALLERGENS=[
  {id:"gluten",l:"Gluten",tr:"Gluten",en:"Gluten",ru:"Глютен",es:"Gluten",de:"Gluten",fr:"Gluten",zh:"麸质",ar:"الغلوتين",icon:"🌾",c:"#d97706"},
  {id:"milk",l:"Süt",tr:"Süt",en:"Milk",ru:"Молоко",es:"Leche",de:"Milch",fr:"Lait",zh:"牛奶",ar:"الحليب",icon:"🥛",c:"#3b82f6"},
  {id:"egg",l:"Yumurta",tr:"Yumurta",en:"Egg",ru:"Яйцо",es:"Huevo",de:"Ei",fr:"Œuf",zh:"鸡蛋",ar:"البيض",icon:"🥚",c:"#eab308"},
  {id:"soy",l:"Soya",tr:"Soya",en:"Soy",ru:"Соя",es:"Soja",de:"Soja",fr:"Soja",zh:"大豆",ar:"الصويا",icon:"🫘",c:"#16a34a"},
  {id:"peanut",l:"Yer Fıstığı",tr:"Yer Fıstığı",en:"Peanut",ru:"Арахис",es:"Maní",de:"Erdnuss",fr:"Cacahuète",zh:"花生",ar:"الفول السوداني",icon:"🥜",c:"#92400e"},
  {id:"treenut",l:"Ağaç Kuruyemiş",tr:"Ağaç Kuruyemiş",en:"Tree Nut",ru:"Орехи",es:"Frutos secos",de:"Nüsse",fr:"Fruits à coque",zh:"坚果",ar:"المكسرات",icon:"🌰",c:"#7c3aed"},
  {id:"sesame",l:"Susam",tr:"Susam",en:"Sesame",ru:"Кунжут",es:"Sésamo",de:"Sesam",fr:"Sésame",zh:"芝麻",ar:"السمسم",icon:"⚪",c:"#ca8a04"},
  {id:"fish",l:"Balık",tr:"Balık",en:"Fish",ru:"Рыба",es:"Pescado",de:"Fisch",fr:"Poisson",zh:"鱼",ar:"السمك",icon:"🐟",c:"#0ea5e9"},
  {id:"shellfish",l:"Kabuklu Deniz",tr:"Kabuklu Deniz",en:"Shellfish",ru:"Моллюски",es:"Mariscos",de:"Schalentiere",fr:"Crustacés",zh:"贝类",ar:"المحاريات",icon:"🦐",c:"#dc2626"},
  {id:"mollusk",l:"Yumuşakça",tr:"Yumuşakça",en:"Mollusk",ru:"Моллюски",es:"Moluscos",de:"Weichtiere",fr:"Mollusques",zh:"软体动物",ar:"الرخويات",icon:"🐙",c:"#9333ea"},
  {id:"mustard",l:"Hardal",tr:"Hardal",en:"Mustard",ru:"Горчица",es:"Mostaza",de:"Senf",fr:"Moutarde",zh:"芥末",ar:"الخردل",icon:"🌶",c:"#facc15"},
  {id:"celery",l:"Kereviz",tr:"Kereviz",en:"Celery",ru:"Сельдерей",es:"Apio",de:"Sellerie",fr:"Céleri",zh:"芹菜",ar:"الكرفس",icon:"🌿",c:"#65a30d"},
  {id:"lupin",l:"Lupin",tr:"Lupin",en:"Lupin",ru:"Люпин",es:"Lupino",de:"Lupinen",fr:"Lupin",zh:"羽扇豆",ar:"الترمس",icon:"🌱",c:"#84cc16"},
  {id:"sulfite",l:"Sülfit",tr:"Sülfit",en:"Sulphite",ru:"Сульфиты",es:"Sulfitos",de:"Sulfite",fr:"Sulfites",zh:"亚硫酸盐",ar:"الكبريتات",icon:"💧",c:"#a855f7"},
  {id:"goatmilk",l:"Keçi/Koyun Sütü",tr:"Keçi/Koyun Sütü",en:"Goat/Sheep Milk",ru:"Козье/овечье молоко",es:"Leche cabra/oveja",de:"Ziegenmilch",fr:"Lait chèvre/brebis",zh:"羊奶",ar:"حليب الماعز",icon:"🐐",c:"#78716c"},
  {id:"buckwheat",l:"Karabuğday",tr:"Karabuğday",en:"Buckwheat",ru:"Гречка",es:"Trigo sarraceno",de:"Buchweizen",fr:"Sarrasin",zh:"荞麦",ar:"الحنطة السوداء",icon:"🟤",c:"#a16207"},
  {id:"legume",l:"Baklagiller",tr:"Baklagiller",en:"Legumes",ru:"Бобовые",es:"Legumbres",de:"Hülsenfrüchte",fr:"Légumineuses",zh:"豆类",ar:"البقوليات",icon:"🫛",c:"#4ade80"},
  {id:"sugar",l:"Şeker",tr:"Şeker",en:"Sugar",ru:"Сахар",es:"Azúcar",de:"Zucker",fr:"Sucre",zh:"糖",ar:"السكر",icon:"🍬",c:"#f472b6"},
  {id:"highsalt",l:"Yüksek Tuz",tr:"Yüksek Tuz",en:"High Salt",ru:"Много соли",es:"Alto en sal",de:"Hoher Salz",fr:"Sel élevé",zh:"高盐",ar:"ملح مرتفع",icon:"🧂",c:"#94a3b8"},
  {id:"alcohol",l:"Alkol",tr:"Alkol",en:"Alcohol",ru:"Алкоголь",es:"Alcohol",de:"Alkohol",fr:"Alcool",zh:"酒精",ar:"الكحول",icon:"🍷",c:"#b91c1c"},
  {id:"pork",l:"Domuz",tr:"Domuz",en:"Pork",ru:"Свинина",es:"Cerdo",de:"Schwein",fr:"Porc",zh:"猪肉",ar:"لحم الخنزير",icon:"🐷",c:"#fb923c"}
];
const allergenL=(a,lang)=>a&&(a[lang]||a.tr||a.l)||"";

const DIETS=[
  {id:"vegan",l:"Vegan",tr:"Vegan",en:"Vegan",ru:"Веган",es:"Vegano",de:"Vegan",fr:"Végan",zh:"纯素",ar:"نباتي",icon:"🌱"},
  {id:"vegetarian",l:"Vejetaryen",tr:"Vejetaryen",en:"Vegetarian",ru:"Вегетарианец",es:"Vegetariano",de:"Vegetarisch",fr:"Végétarien",zh:"素食",ar:"نباتي",icon:"🥬"},
  {id:"glutenfree",l:"Glutensiz",tr:"Glutensiz",en:"Gluten-Free",ru:"Без глютена",es:"Sin gluten",de:"Glutenfrei",fr:"Sans gluten",zh:"无麸质",ar:"خالٍ من الغلوتين",icon:"🚫🌾"},
  {id:"lactosefree",l:"Laktozsuz",tr:"Laktozsuz",en:"Lactose-Free",ru:"Без лактозы",es:"Sin lactosa",de:"Laktosefrei",fr:"Sans lactose",zh:"无乳糖",ar:"خالٍ من اللاكتوز",icon:"🚫🥛"},
  {id:"sugarfree",l:"Şekersiz",tr:"Şekersiz",en:"Sugar-Free",ru:"Без сахара",es:"Sin azúcar",de:"Zuckerfrei",fr:"Sans sucre",zh:"无糖",ar:"خالٍ من السكر",icon:"🚫🍬"},
  {id:"kosher",l:"Koşer",tr:"Koşer",en:"Kosher",ru:"Кошер",es:"Kosher",de:"Koscher",fr:"Casher",zh:"洁食",ar:"كوشر",icon:"✡"},
  {id:"halal",l:"Helal",tr:"Helal",en:"Halal",ru:"Халяль",es:"Halal",de:"Halal",fr:"Halal",zh:"清真",ar:"حلال",icon:"☪"}
];
const dietL=(d,lang)=>d&&(d[lang]||d.tr||d.l)||"";

// Pişirme yöntemini dile göre çevir (eski reçetelerde TR ile saklı olabilir)
const cookMethodL=(method,L)=>{
  if(!method)return "";
  const m=method.toLowerCase().trim();
  const map={
    "fırın":L.methodOven,"firin":L.methodOven,"oven":L.methodOven,
    "kaynat":L.methodBoil,"haşla":L.methodBoil,"hasla":L.methodBoil,"boil":L.methodBoil,
    "kızart":L.methodFry,"kizart":L.methodFry,"fry":L.methodFry,
    "ızgara":L.methodGrill,"izgara":L.methodGrill,"grill":L.methodGrill,
    "benmari":L.methodBainMarie,"bain-marie":L.methodBainMarie,
    "buzdolabı":L.methodFridge,"buzdolabi":L.methodFridge,"refrigerator":L.methodFridge,"fridge":L.methodFridge,
    "dondurma":L.methodFreeze,"dondurucu":L.methodFreeze,"freeze":L.methodFreeze,"freezer":L.methodFreeze,
    "soğutma":L.methodCool,"sogutma":L.methodCool,"cool":L.methodCool,"cooling":L.methodCool,
    "pişirme yok":L.methodNone,"no cooking":L.methodNone
  };
  return map[m]||method;
};

// 10 STOK KATEGORİSİ (madde 7)
const STOCK_CATS=[
  {id:"dairy",l:"Süt Ürünleri",tr:"Süt Ürünleri",en:"Dairy",ru:"Молочные",es:"Lácteos",de:"Milchprodukte",fr:"Produits laitiers",zh:"乳制品",ar:"الألبان",icon:"🥛"},
  {id:"eggs",l:"Yumurta",tr:"Yumurta",en:"Eggs",ru:"Яйца",es:"Huevos",de:"Eier",fr:"Œufs",zh:"鸡蛋",ar:"البيض",icon:"🥚"},
  {id:"redmeat",l:"Kırmızı Et",tr:"Kırmızı Et",en:"Red Meat",ru:"Красное мясо",es:"Carne roja",de:"Rotfleisch",fr:"Viande rouge",zh:"红肉",ar:"اللحم الأحمر",icon:"🥩"},
  {id:"poultry",l:"Kümes Hayvanları",tr:"Kümes Hayvanları",en:"Poultry",ru:"Птица",es:"Aves",de:"Geflügel",fr:"Volaille",zh:"家禽",ar:"الدواجن",icon:"🍗"},
  {id:"seafood",l:"Su Ürünleri",tr:"Su Ürünleri",en:"Seafood",ru:"Морепродукты",es:"Mariscos",de:"Meeresfrüchte",fr:"Fruits de mer",zh:"海鲜",ar:"المأكولات البحرية",icon:"🐟"},
  {id:"vegetable",l:"Sebze",tr:"Sebze",en:"Vegetables",ru:"Овощи",es:"Verduras",de:"Gemüse",fr:"Légumes",zh:"蔬菜",ar:"الخضروات",icon:"🥦"},
  {id:"fruit",l:"Meyve",tr:"Meyve",en:"Fruits",ru:"Фрукты",es:"Frutas",de:"Obst",fr:"Fruits",zh:"水果",ar:"الفواكه",icon:"🍎"},
  {id:"legumes",l:"Bakliyat",tr:"Bakliyat",en:"Legumes",ru:"Бобовые",es:"Legumbres",de:"Hülsenfrüchte",fr:"Légumineuses",zh:"豆类",ar:"البقوليات",icon:"🫘"},
  {id:"grain",l:"Unlu & Tahıl",tr:"Unlu & Tahıl",en:"Grains & Flour",ru:"Зерновые",es:"Granos",de:"Getreide",fr:"Céréales",zh:"谷物",ar:"الحبوب",icon:"🌾"},
  {id:"spice",l:"Baharat & Ot",tr:"Baharat & Ot",en:"Spices & Herbs",ru:"Специи",es:"Especias",de:"Gewürze",fr:"Épices",zh:"香料",ar:"التوابل",icon:"🌶"},
  {id:"oils",l:"Yağ & Sirke",tr:"Yağ & Sirke",en:"Oils & Vinegar",ru:"Масла",es:"Aceites",de:"Öle",fr:"Huiles",zh:"油醋",ar:"الزيوت",icon:"🛢"},
  {id:"sweetener",l:"Şeker & Tatlandırıcı",tr:"Şeker & Tatlandırıcı",en:"Sugar & Sweetener",ru:"Сахар",es:"Azúcar",de:"Zucker",fr:"Sucre",zh:"糖",ar:"السكر",icon:"🍬"},
  {id:"canned",l:"Konserve & Sos",tr:"Konserve & Sos",en:"Canned & Sauces",ru:"Консервы",es:"Enlatados",de:"Konserven",fr:"Conserves",zh:"罐头",ar:"المعلبات",icon:"🥫"},
  {id:"pastry",l:"Pastane & Hamur",tr:"Pastane & Hamur",en:"Pastry & Dough",ru:"Кондитерские",es:"Pastelería",de:"Backwaren",fr:"Pâtisserie",zh:"糕点",ar:"المعجنات",icon:"🎂"},
  {id:"nuts",l:"Kuruyemiş & Kuru Meyve",tr:"Kuruyemiş & Kuru Meyve",en:"Nuts & Dried Fruit",ru:"Орехи",es:"Frutos secos",de:"Nüsse",fr:"Fruits secs",zh:"坚果",ar:"المكسرات",icon:"🥜"},
  {id:"coffee",l:"Çay & Kahve",tr:"Çay & Kahve",en:"Tea & Coffee",ru:"Чай и кофе",es:"Té y café",de:"Tee & Kaffee",fr:"Thé & café",zh:"茶咖啡",ar:"الشاي والقهوة",icon:"☕"},
  {id:"beverage",l:"İçecek",tr:"İçecek",en:"Beverages",ru:"Напитки",es:"Bebidas",de:"Getränke",fr:"Boissons",zh:"饮料",ar:"المشروبات",icon:"🥤"},
  {id:"alcohol",l:"Alkol & Bar",tr:"Alkol & Bar",en:"Alcohol & Bar",ru:"Алкоголь",es:"Alcohol",de:"Alkohol",fr:"Alcool",zh:"酒类",ar:"الكحول",icon:"🍷"},
  {id:"frozen",l:"Dondurulmuş",tr:"Dondurulmuş",en:"Frozen",ru:"Замороженные",es:"Congelados",de:"Tiefkühl",fr:"Surgelés",zh:"冷冻",ar:"المجمدات",icon:"🧊"},
  {id:"cleaning",l:"Temizlik",tr:"Temizlik",en:"Cleaning",ru:"Уборка",es:"Limpieza",de:"Reinigung",fr:"Nettoyage",zh:"清洁",ar:"التنظيف",icon:"🧹"},
  {id:"packaging",l:"Ambalaj & Sarf",tr:"Ambalaj & Sarf",en:"Packaging",ru:"Упаковка",es:"Embalaje",de:"Verpackung",fr:"Emballage",zh:"包装",ar:"التعبئة",icon:"📦"},
  {id:"other",l:"Diğer",tr:"Diğer",en:"Other",ru:"Другое",es:"Otros",de:"Sonstige",fr:"Autres",zh:"其他",ar:"أخرى",icon:"🗂"}
];
const stockCatL=(c,lang)=>c&&(c[lang]||c.tr||c.l)||"";

// Keyword → kategori eşlemesi (zengin, profesyonel)
const CAT_KEYWORDS={
  dairy:["süt","krema","peynir","yoğurt","tereyağ","kaşar","lor","labne","kefir","ayran","mozzarella","parmesan","ricotta","cheddar","feta","keçi peyniri","süzme","kaymak","eritme peyniri","beyaz peynir","dil peyniri","örgü peynir","hellim","halloumi","burrata","mascarpone","gorgonzola","roquefort","brie","camembert","manda"],
  eggs:["yumurta","egg","bıldırcın yumurta"],
  redmeat:["dana","sığır","kuzu","koyun","bonfile","antrikot","kontrfile","pirzola","kaburga","biftek","rosto","kıyma","köfte","sucuk","salam","pastırma","jambon","şnitzel","sığır fileto","dana kuşbaşı","dana kıyma","kuzu but","kuzu pirzola"],
  poultry:["tavuk","hindi","ördek","kaz","piliç","bıldırcın","tavuk göğüs","tavuk but","tavuk kanat","tavuk baget","tavuk filetosu","hindi göğüs","hindi but","chicken","turkey","duck"],
  seafood:["balık","karides","ahtapot","kalamar","levrek","çipura","somon","salmon","alabalık","hamsi","palamut","lüfer","uskumru","sardalya","barbunya balık","istavrit","mersin","ton","tuna","midye","istiridye","yengeç","ıstakoz","deniz mahsulleri","octopus","shrimp"],
  vegetable:["domates","salatalık","patlıcan","biber","soğan","sarımsak","patates","havuç","enginar","ıspanak","marul","kabak","mantar","kereviz","lahana","turp","brokoli","karnabahar","roka","maydanoz","dereotu","pazı","pırasa","bamya","brüksel","şalgam","taze soğan","pancar","rezene","kuşkonmaz","fasulye taze","bezelye taze","semizotu","tere","ebegümeci","kenger","endivye","radika","cherry domates","arpacık"],
  fruit:["elma","limon","portakal","muz","kiraz","karpuz","kavun","şeftali","kayısı","erik","üzüm","incir","ayva","mandalina","greyfurt","nar","avokado","ananas","mango","kivi","çilek","böğürtlen","ahududu","yaban mersini","blueberry","strawberry","pineapple"],
  legumes:["mercimek","nohut","fasulye","barbunya","bezelye","bakla","soya","ful","lentil","bean","chickpea","börülce","siyah fasulye","kuru fasulye","yeşil mercimek","kırmızı mercimek"],
  grain:["un","pirinç","bulgur","makarna","irmik","nişasta","yulaf","mısır unu","ekmek","pasta","bisküvi","kraker","granola","çavdar","arpa","buğday","kuskus","spagetti","penne","fusilli","makarn","vermicelli","lasagna","galeta","gofret","pilavlık","arborio","basmati","jasmine","noodle","şehriye","tel şehriye"],
  spice:["karabiber","kimyon","zerdeçal","pul biber","kırmızı biber","tarçın","kakule","zencefil","rezene","kekik","nane","biberiye","fesleğen","kişniş","köri","safran","sumak","tuz","deniz tuzu","kaya tuzu","baharat","defne","karanfil","anason","hindistan cevizi baharat","muskat","paprika","cayenne","curry","garam masala","za'atar","oregano","thyme","rosemary","basil"],
  oils:["zeytinyağ","ayçiçek yağ","mısır yağ","kanola","susam yağ","sirke","balzamik","elma sirkesi","üzüm sirkesi","margarin","hindistan cevizi yağı","avokado yağı","truf yağı","aspir yağı","palmiye yağı","kızartma yağ"],
  sweetener:["şeker","bal","pekmez","stevia","sakarin","fruktoz","glikoz","şurup","pudra şekeri","esmer şeker","akçaağaç","tatlandırıcı","karamel","kesme şeker","sıvı şeker"],
  canned:["salça","ketçap","hardal","mayonez","konserve","turşu","zeytin","kapari","sos","barbekü","soya sosu","pesto","tahini","tahin","reçel","marmelat","fıstık ezmesi","sriracha","tabasco","worcestershire","hoisin","teriyaki"],
  pastry:["vanilya","kakao","maya","kabartma tozu","karbonat","jelatin","agar","fondan","krem şanti","pasta kremi","sprinkle","süsleme","pasta malzemesi","yufka","baklavalık","börek yufkası","milföy","puf pastry","galeta unu","badem unu","nişasta hamur","krem peynir","çikolata parçaları","kakao tozu","beyaz çikolata","bitter çikolata","sütlü çikolata","tavasim","glaze","buttercream"],
  nuts:["ceviz","badem","fındık","fıstık","yer fıstığı","antep fıstığı","kaju","macadamia","pekan","brazil","çam fıstığı","ay çekirdeği","kabak çekirdeği","kuru üzüm","kuru kayısı","kuru incir","kuru erik","kuru hurma","hurma","lokum","kestane","chia","keten tohumu","susam","haşhaş"],
  coffee:["kahve","çay","filtre kahve","espresso","nescafe","türk kahvesi","yeşil çay","siyah çay","bitki çayı","oolong","matcha","earl grey","chai","coffee","tea","kaffee"],
  beverage:["su","maden suyu","meyve suyu","gazoz","kola","limonata","içecek","soda","cola","pepsi","sprite","fanta","redbull","enerji içeceği","ice tea","smoothie","shake","ayran içecek"],
  alcohol:["şarap","bira","rakı","votka","viski","cin","rom","tekila","likör","bitters","vermut","şampanya","prosecco","cognac","sake","wine","beer","whiskey","gin","rum","vodka","champagne"],
  frozen:["dondurulmuş","dondurma","frozen","buz","ice cream","sorbet","donmuş","pronto","dondurulmus"],
  cleaning:["deterjan","sabun","temizlik","bez","sünger","çamaşır","bulaşık","dezenfektan","eldiven","çöp","yumuşatıcı","arap sabunu","amonyak","klor","çamaşır suyu"],
  packaging:["folyo","streç","poşet","torba","kutu","naylon","ambalaj","etiket","tek kullanımlık","peçete","havlu","kürdan","pipet","karton","stretch","alüminyum folyo","pişirme kağıdı","parşömen","pastry bag","sıkma torbası"]
};

// Bilinen Türk gıda markaları
const TR_BRANDS=["Pınar","Sek","İçim","Ülker","Eti","Bim","Migros","Dost","Tikveşli","Monea",
  "Torku","Sütas","Gönen","Yörsan","Yörem","Altınkılıç","Algida","Carte D'Or","Nestlé","Nestle",
  "Danone","Activia","President","Elle&Vire","Elle Vire","Lurpak","Anchor","Kerrygold",
  "Aytaç","Namet","Gürer","Polonez","Banvit","Beypazarı","Misir","Pastavilla","Barilla",
  "Nescafe","Lipton","Çaykur","Doğuş","Filiz","Oba","Reis","Sade","Knorr","Maggi",
  "Heinz","Delmonte","Del Monte","Hero","Olives","Komili","Kristal","Yudum","Botanik",
  "Arifoğlu","Öncü","Sera","Tat","Tamek","Bağdat","Burgaz","Bizim","Lüks",
  "For Kitchen","ForKitchen","Pastacı","Royal","Dr.Oetker","Dr Oetker","Oetker",
  "Jacobs","Nespresso","Lavazza","İlhan","Kahve Dünyası","Starbucks"];

function parseProductName(name){
  if(!name)return{cleanName:name,brand:"",qty:0,unit:"g"};

  // Marka ayrıştır
  let brand="";
  let workName=name;
  for(const b of TR_BRANDS){
    const re=new RegExp("\\b"+b.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i");
    if(re.test(workName)){
      brand=b;
      workName=workName.replace(re,"").trim().replace(/\s+/g," ");
      break;
    }
  }

  // Miktar/birim ayrıştır
  const patterns=[
    {re:/(\d+(?:[.,]\d+)?)\s*(lt|litre|l\b)/i, unit:"ml", mult:1000},
    {re:/(\d+(?:[.,]\d+)?)\s*(ml|mililitre)/i, unit:"ml", mult:1},
    {re:/(\d+(?:[.,]\d+)?)\s*(kg|kilogram)/i, unit:"g", mult:1000},
    {re:/(\d+(?:[.,]\d+)?)\s*(gr|gram|g\b)/i, unit:"g", mult:1},
    {re:/(\d+(?:[.,]\d+)?)\s*(adet|pcs|piece|ad\.?)/i, unit:"adet", mult:1},
    {re:/x(\d+)/i, unit:"adet", mult:1},
  ];
  for(const p of patterns){
    const m=workName.match(p.re);
    if(m){
      const qty=parseFloat(m[1].replace(",","."))*p.mult;
      const cleanName=workName.replace(m[0],"").trim().replace(/\s+/g," ");
      return{cleanName:cleanName||workName,brand,qty,unit:p.unit};
    }
  }
  return{cleanName:workName||name,brand,qty:0,unit:"g"};
}


function guessStockCat(name){
  const n=name.toLowerCase();
  for(const[cat,kws]of Object.entries(CAT_KEYWORDS)){
    if(kws.some(k=>n.includes(k)))return cat;
  }
  return "other";
}

// Migration: eski kategorileri yeni kategorilere taşı (v1.8.0 → v1.8.0)
const OLD_TO_NEW_CAT={
  "meatfish":null,    // Ad kontrolüyle redmeat/poultry/seafood
  "produce":null,     // Ad kontrolüyle vegetable/fruit
  "dryfood":"grain",  // eski v1.1 uyumluluk
  "stationery":"packaging"
};
function migrateStockCat(item){
  if(!item.cat||item.cat==="other"){
    return{...item,cat:guessStockCat(item.name)};
  }
  if(OLD_TO_NEW_CAT[item.cat]===null){
    return{...item,cat:guessStockCat(item.name)};
  }
  if(OLD_TO_NEW_CAT[item.cat]){
    return{...item,cat:OLD_TO_NEW_CAT[item.cat]};
  }
  return item;
}

// ═══ FUZZY MATCH — Genel kelimeler düşük ağırlık, ayırt edici sıfatlar kritik ═══
const STOP_WORDS=["ve","ile","için","bir","tam","yağlı","yağsız","light","g","ml","kg","l","adet","gr","lt","cl"];
// Genel kategori kelimeleri — tek başına eşleşme SAYILMAZ, ayırt edici değil
const GENERIC_WORDS=["peynir","peyniri","et","sos","sosu","süt","sütü","yağ","yağı","un","unu","şeker","şekeri","biber","biberi","salça","salçası","krema","kreması","makarna","makarnası","turşu","turşusu","reçel","reçeli","konserve","conserve","çikolata","çikolatası"];
const tokenize=(s)=>String(s||"").toLowerCase().replace(/[^a-zçğıöşü0-9\s]/gi," ").split(/\s+/).filter(w=>w.length>1&&!STOP_WORDS.includes(w));
const fuzzyMatch=(a,b)=>{
  const ta=tokenize(a),tb=tokenize(b);
  if(!ta.length||!tb.length)return 0;
  const common=ta.filter(w=>tb.some(x=>x===w||(x.length>3&&w.length>3&&(x.includes(w)||w.includes(x)))));
  const all=new Set([...ta,...tb]).size;
  // Eğer ortak kelimelerin TAMAMI generic ise → eşleşme sayma
  const nonGenericCommon=common.filter(w=>!GENERIC_WORDS.includes(w));
  if(common.length>0&&nonGenericCommon.length===0)return 0;
  const jaccard=common.length/all;
  if(ta.length>=2&&tb.length>=2&&common.length<2)return 0;
  return jaccard;
};

// Birim formatlama: 1500 g → "1,5 kg", 250 ml → "250 ml"
const fmtQty=(qty,unit)=>{
  if(qty===null||qty===undefined||isNaN(qty))return "0 "+unit;
  if(unit==="g"&&qty>=1000)return (qty/1000).toFixed(qty%1000===0?0:2).replace(".",",")+" kg";
  if(unit==="ml"&&qty>=1000)return (qty/1000).toFixed(qty%1000===0?0:2).replace(".",",")+" l";
  if(qty%1===0)return qty+" "+unit;
  return qty.toFixed(2).replace(".",",")+" "+unit;
};

const DEF_RECIPES=[
  {id:1,name:"Crème Brûlée",mainCat:"pastry",subCat:"milk",venue:"alacarte",prep:"chilled",cuisine:"french",difficulty:"medium",
    allergens:["milk","egg"],diets:[],yield:"8 porsiyon",servings:8,photo:null,calories:320,
    ingredients:[{name:"Krema %35",amount:"500 ml"},{name:"Yumurta sarısı",amount:"6 adet"},{name:"Toz şeker",amount:"120 g"},{name:"Vanilya",amount:"1 çubuk"}],
    notes:"180°C benmari 35-40 dk.",created:"2026-04-01"},
  {id:2,name:"Zeytinyağlı Enginar",mainCat:"cold",subCat:"meze",venue:"alacarte",prep:"cooked",cuisine:"turkish",difficulty:"easy",
    allergens:[],diets:["vegan","vegetarian","glutenfree","lactosefree"],yield:"4 porsiyon",servings:4,photo:null,calories:110,
    ingredients:[{name:"Enginar",amount:"4 adet"},{name:"Zeytinyağı",amount:"100 ml"},{name:"Limon",amount:"2 adet"}],
    notes:"Yavaş pişirme 45 dk.",created:"2026-04-02"}
];
const DEF_STOCK=[
  {id:1,name:"Krema %35",unit:"ml",qty:5000,ppu:0.045,upd:"2026-04-01",low:1000,cat:"dairy"},
  {id:2,name:"Yumurta",unit:"adet",qty:120,ppu:3.5,upd:"2026-04-01",low:30,cat:"other"},
  {id:3,name:"Toz Şeker",unit:"g",qty:10000,ppu:0.012,upd:"2026-04-01",low:2000,cat:"sweetener"}
];

// 8 DİL (madde 10) — görsel iskelet
const LANGS=[
  {code:"tr",flag:"🇹🇷",label:"TR",fullName:"Türkçe"},
  {code:"en",flag:"🇬🇧",label:"EN",fullName:"English"},
  {code:"de",flag:"🇩🇪",label:"DE",fullName:"Deutsch"},
  {code:"fr",flag:"🇫🇷",label:"FR",fullName:"Français"},
  {code:"ar",flag:"🇸🇦",label:"AR",fullName:"العربية"},
  {code:"ru",flag:"🇷🇺",label:"RU",fullName:"Русский"},
  {code:"zh",flag:"🇨🇳",label:"ZH",fullName:"中文"},
  {code:"es",flag:"🇪🇸",label:"ES",fullName:"Español"}
];

// ═══ ÇEVİRİ SİSTEMİ ═══
const I18N={
  tr:{
    // Sekmeler
    tabRecipes:"Reçeteler",tabStock:"Stok",tabProduction:"Üretim",tabReports:"Raporlar",tabMenus:"Menüler",tabAssistant:"Asistan",tabSettings:"Ayarlar",tabTodo:"Görevler",tabTeam:"Ekip",tabTeamChat:"Ekip Chat",
    // Genel
    save:"Kaydet",cancel:"İptal",delete:"Sil",edit:"Düzenle",close:"Kapat",add:"Ekle",back:"Geri",next:"İleri",search:"Ara",filter:"Filtrele",all:"Tümü",yes:"Evet",no:"Hayır",confirm:"Onayla",remove:"Kaldır",update:"Güncelle",name:"Ad",amount:"Miktar",unit:"Birim",price:"Fiyat",date:"Tarih",time:"Saat",note:"Not",optional:"opsiyonel",required:"zorunlu",loading:"Yükleniyor...",error:"Hata",success:"Başarılı",warning:"Uyarı",
    // Reçeteler
    recipes:"Reçeteler",newRecipe:"+ Yeni Reçete",recipePhoto:"📷 Reçete (AI)",recipeManual:"+ Manuel",recipeDetail:"Reçete Detayı",ingredients:"Malzemeler",allergens:"Alerjenler",cookMethod:"Pişirme",cookTemp:"Sıcaklık",cookTime:"Süre",servings:"Porsiyon",totalWeight:"Toplam Ağırlık",portionG:"Porsiyon Gramajı",calories:"Kalori",yieldL:"Verim",recipeName:"Reçete Adı",deductStock:"📦 Stoktan Düş & Üret",sampleLabel:"🧪 Numune Şahit Etiketi",editPhoto:"📷 Fotoğrafı Değiştir",addPhoto:"📷 Fotoğraf Ekle",costAnalysis:"MALİYET ANALİZİ",material:"Hammadde",overhead:"Sabit+Personel payı",realCost:"Gerçek Maliyet",perPortion:"porsiyon",costHint:"💡 Sabit gider ve personel giderlerini Ayarlar'a ekleyerek gerçek maliyeti görebilirsiniz.",
    // Stok
    stock:"Stok",stockMaterial:"📦 Hammadde",stockLots:"🏷 Parti Numaraları",invoice:"📄 Fatura (AI)",manualAdd:"+ Manuel",lowStock:"⚠ Düşük Stok",products:"ürün",currentLots:"Güncel Parti Numaraları",noTrackedIngs:"Takip edilecek hammadde yok",noTrackedIngsHint:"Ayarlar → İzlenebilirlik → Takip Edilecek Hammaddeler ekleyin",lotMissing:"Parti no girilmedi",lotToday:"bugün",lotDaysAgo:"gün önce",lotEntered:"girildi",lotHint:"💡 Sabah ambalajların parti no'larını bu listeye işle. 🔴 eksik · 🔵 güncel · 🟡 eski (3+ gün).",
    // Üretim
    production:"Üretim",productionRecord:"🍱 Üretim Kaydı",productionReports:"📊 Raporlar",storage:"Depolama Yeri",expiryDays:"Tüketim Tarihi (gün)",daysLater:"gün sonra",risk:"Risk",riskHigh:"Yüksek",riskMedium:"Orta",riskLow:"Düşük",riskHint:"Otomatik hesap, elle değiştirebilirsin",noteHint:"Örn: Akşam servisi için",produceSaveBtn:"✓ Üret & Stoktan Düş",portions:"porsiyon",lot:"Lot",produced:"Üretim",expires:"Tüketim",consumed:"tüketildi",active:"aktif",fire:"Fire",consumeAll:"✓ Tümü Tüketildi",printLabel:"🏷 Etiket Bas",partialConsume:"KISMI TÜKETİM",decrease:"− Düş",markFire:"× Fire Olarak İşaretle",gotoReport:"📊 Rapora Git",noProduction:"Henüz üretim yok",noProductionHint:"Reçete detayından \"Stoktan Düş\" ile üretim kaydı ekle",consumedFull:"{name} ({portions} porsiyon) tüketildi olarak işaretlenecek. Onaylıyor musun?",invalidAmount:"Geçersiz miktar",howManyFire:"Kaç porsiyon fire? (toplam {total})",ingredientLotsTitle:"🏷 HAMMADDE PARTİ NUMARALARI",lotMissingHint:"💡 Eksik parti no'ları için Stok → Parti Numaraları sekmesine git.",
    // Numune
    sampleTitle:"🧪 Numune Şahit Etiketi",sampleDesc:"ISO 22000 uyumlu numune kaydı — 72 saat saklama önerilir",sampleName:"Numune Adı",sampleTaker:"Numuneyi Alan",sampleOrgLabel:"Organizasyon",sampleOrgSelect:"Seçin...",samplePax:"Kişi Sayısı",sampleLocation:"Ek Bilgi / Yer",sampleLocationHint:"Örn: Ahmet-Ayşe düğünü, Salon A, VIP masa",sampleDateTime:"Alım Saati ve Tarihi (tabaklama anı)",sampleTemp:"Numune Sıcaklığı (°C)",sampleTempHint:"Tabaklama anı",sampleHours:"Saklama Süresi (saat)",samplePrintSave:"🏷 Etiket Bas & Kaydet",sampleOrgOrLoc:"Organizasyon veya yer girin",sampleProfileWarn:"⚠ Ayarlar → Kullanıcı Profili'ne ad girin",
    // Raporlar
    reports:"Raporlar",reportsProduction:"📋 Üretim Raporları",reportsStorage:"🗄 Depo Kontrolü",noReports:"Henüz rapor yok",checkTotal:"kontrol",checkComplete:"✓ Tam",addCheck:"+ Kontrol Ekle",noChecksYet:"Henüz kontrol girilmedi.",storageCheckHint:"💡 Her depo/dolap için günde 3 kez sıcaklık kontrolü yapılır (09:00, 15:00, 21:00).",
    productionsCap:"ÜRETİMLER",samplesCap:"NUMUNELER",firesCap:"FİRE",
    // Depo Kontrol Modal
    storageCheckTitle:"🗄 Depo Kontrolü",controlTime:"Kontrol Saati",controlParams:"KONTROL PARAMETRELERİ",paramEtiket:"Etiket Kontrol",paramAcikGida:"Üzeri Açık Gıda",paramYerdeDuran:"Yerde Duran Gıda",paramCigPismis:"Çiğ/Pişmiş Ayrımı",paramTemizlik:"Temizlik",paramAlerjen:"Alerjen Raf Ayrımı",paramSkt:"SKT/Raf Ömrü",ok:"✓ Uygun",notOk:"✗ Uygun Değil",productTemp:"Ürün Sıcaklığı (°C)",gaugeValue:"Gösterge Değeri (°C)",explanation:"Açıklama (opsiyonel)",explanationHint:"Düzeltici faaliyet vs.",atLeastOneTemp:"En az bir sıcaklık değeri girin",
    // Menüler
    menus:"Menü Kartları",newMenu:"+ Yeni Menü",chooseMenuType:"Menü Tipi Seç",menuEdit:"Menü Düzenle",preview:"Önizleme",menuName:"Menü Adı",font:"Font",theme:"Tema",customSizes:"📐 ÖZEL ÖLÇÜLER (mm)",or:"veya",showPrice:"Fiyat göster",showCalorie:"Kalori göster",showAllergen:"Alerjen göster",bgImage:"🖼 Arka Plan Fotoğrafı (opsiyonel)",chooseBg:"📷 Seç",changeBg:"📷 Değiştir",cropBg:"✂ Kırp",bgOpacity:"Opaklık",bgHint:"💡 Mermer, kağıt, ahşap doku gibi arka plan ekleyebilirsin. Yükledikten sonra kırpma editörü açılır.",menuTemplates:"📑 MENÜ ŞABLONU",loadTemplate:"Şablon yükle...",saveTemplate:"💾 Şablon Kaydet",templateHint:"💡 Font, tema, boyut, arka plan ve görüntü seçeneklerini şablon olarak kaydet.",templateName:"Şablon ismi:",defaultTemplateName:"{menuName} Şablonu",templateSaveConfirm:"\"{name}\" şablonu yüklensin mi? (Bölümler ve ürünler korunur, tasarım değişir)",templateSaved:"\"{name}\" şablon olarak kaydedildi.",templateDeletePrompt:"Silinecek şablon ismi?",templateNotFound:"Bulunamadı",
    // Crop modal
    cropTitle:"✂ Fotoğrafı Kırp",cropDesc:"Sürükleyerek konumu değiştir, slider ile zoom. Çerçeve içindeki alan menü arka planı olur.",cropZoom:"🔍 Zoom",cropApply:"✓ Kırp & Kullan",
    // Ayarlar
    settings:"Ayarlar",apiKey:"API Anahtarı",language:"Dil",darkMode:"Koyu Tema",userProfile:"👤 Kullanıcı Profili",fullName:"Ad Soyad",workplace:"Çalıştığınız Yer",department:"Departman",role:"Görev",profileHint:"💡 Bu bilgiler numune şahit etiketleri ve ileride üretim kayıtlarında kullanılır.",traceabilityTitle:"🔍 İzlenebilirlik (ISO 22000)",traceabilityMode:"İzlenebilirlik Modu",traceabilityDesc:"Parti numarası takibi, ISO 22000 raporları (FR.06, FR.12) ve PDF çıktılarını etkinleştirir",resetHourLabel:"Gün Sonu Parti No Sıfırlama Saati",resetHourDesc:"her gün bu saatte parti numaraları sıfırlanır",trackedIngsLabel:"Takip Edilecek Hammaddeler",trackedIngsPh:"Hammadde ekle (örn: Un, Süt, Yumurta)",trackedIngsHint:"💡 Her sabah bu hammaddelerin parti numaralarını güncellersin. FR.06 formunda otomatik kullanılır.",organizationsTitle:"🏢 Organizasyonlar (Numune için)",organizationsPh:"Organizasyon ekle (örn: Düğün 200 pax, Gala, VIP Dinner)",organizationsHint:"💡 Şahit numune etiketinde bu liste dropdown olarak kullanılır.",expensesTitle:"💼 İşletme Giderleri",fixedExpenses:"SABİT GİDERLER (aylık)",personnel:"PERSONEL (aylık maaş)",addFixed:"+ Sabit Gider Ekle",addPersonnel:"+ Personel Ekle",monthlyPortions:"AYLIK TAHMİNİ ÜRETİM (porsiyon)",portionHint:"Reçete maliyetinde gider payı hesabı için kullanılır.",fixedTotal:"Sabit gider toplamı:",personnelTotal:"Personel toplamı:",monthlyTotal:"Aylık toplam gider:",perPortionShare:"Porsiyon başı gider payı:",storageAreas:"🗄 Depolama Alanları",addStorage:"+ Depolama Alanı Ekle",storageHint:"💡 Kendi mutfağınızdaki dolap/buzdolabı/depo isimlerini buraya ekleyin.",backup:"💾 Yedekleme",backupBtn:"⬇ Yedekle",restoreBtn:"⬆ Geri Yükle",developer:"🔧 Geliştirici",copyLogs:"📋 Logları Kopyala",
    // FR.06 PDF
    fr06Title:"GIDA HAZIRLIK KONTROL FORMU",fr06ProductCtrl:"Ürün Hazırlık Kontrolü",fr06ProductIng:"Ürün Adı · Malzemeler · Parti No",fr06PrepStart:"Başl.",fr06PrepEnd:"Bitiş",fr06ProductTemp:"Ürün °C",fr06AmbientTemp:"Ortam °C",fr06Controller:"Kontrol Eden",fr06CookCtrl:"Pişirme / Yeniden Isıtma Sıcaklık Kontrolü",fr06ProductDef:"Ürün Tanımı",fr06PartyNo:"Parti No",fr06Cook:"Pişirme (°C)",fr06Reheat:"Yeniden Isıtma (°C)",fr06CorrectiveAction:"Düzeltici Faaliyet",fr06FreezingCtrl:"Donuk Çözündürme Kontrolü",fr06ThawStart:"Çözdürme Başlangıç",fr06ThawEnd:"Çözdürme Bitiş",fr06Temp:"Sıcaklık",fr06FastCoolCtrl:"Hızlı Soğutma / Dondurma Kontrolü",fr06Entry:"Giriş (Saat-°C)",fr06Exit:"Çıkış (Saat-°C)",fr06Chiller:"Chiller",fr06IceWater:"Buzlu Su/-18°C",fr06PlateHold:"Tabakta Bekletme",fr06ProductOrg:"Ürün / Org.",fr06PlateTime:"Tabağa Konma",fr06Service:"Servis",fr06InnerTemp:"İç Isı",fr06Ctrl:"Kontrol",fr06FosterHold:"Foster/Kurutma Bekletme",fr06Notes:"Uygulama Açıklamaları",fr06NoteRisk:"Riskli Gıda Hazırlık:",fr06NoteRiskText:"Ortam sıcaklığı +24°C'yi geçmemelidir. Yüksek riskli (et-tavuk-balık) gıdalarda max 16°C. Hazırlık süresi: +20°C'de max 45 dk, +25°C'de max 10 dk.",fr06NoteCool:"Soğutma/Dondurma:",fr06NoteCoolText:"90 dakika içinde max 10°C'ye. 5°C'ye ulaşmak için asla 4 saati aşma. Şok dondurma: 4 saat içinde -18°C.",fr06NoteCook:"Pişirme/Yeniden Isıtma:",fr06NoteCookText:"Pişme iç sıcaklık 75°C, dış 65°C. Yeniden ısıtma min 82°C.",fr06NotePlate:"Tabakta Bekletme:",fr06NotePlateText:"15 dk üzeri bekletme yapılmamalıdır.",fr06NoteThaw:"Çözündürme:",fr06NoteThawText:"Buzlu suda çözdürme yapılabilir. Donuk ürünler 1 gün önceden çıkarılır.",fr06DocNo:"Doküman No",fr06DateLbl:"Tarih",fr06DefaultDept:"Mutfak / Bölüm",
    // FR.12 PDF
    fr12Title:"ŞAHİT NUMUNE FORMU",fr12SampleOrg:"Numune Adı / Organizasyon Adı",fr12Pax:"Kişi Sayısı",fr12DateTime:"Alım Tarihi ve Saati",fr12Temp:"Numune Sıcaklığı",fr12Taker:"Numune Alan Kişi",fr12Signature:"İmza",fr12Explanation:"Açıklama",fr12Note:"Şahit Numune:",fr12NoteText:"20 kişi üzerinde hazırlığı yapılan organizasyonların (kahvaltı, personel yemekhanesi dahil) tümü için şahit numuneler alınmalı ve max 5°C, 72 saat saklanmalıdır. Şahit numune olarak belirlenen ürün minimum 200 gram olarak alım sağlanmalıdır.",
    // FR.05 PDF
    fr05Title:"GÜNLÜK DEPO / DOLAP KONTROL FORMU",fr05Dept:"DEPARTMAN:",fr05Cabinet:"DOLAP/DEPO:",fr05Target:"Hedef:",fr05CtrlParams:"KONTROL PARAMETRELERİ",fr05TempCtrl:"SICAKLIK KONTROL (°C)",fr05CtrlTime:"Kontrol Saati",fr05NoteDry:"Kuru Depolama:",fr05NoteDryText:"Sıcaklık <25°C ve bağıl nem ≤%65 değerlerinde olmalıdır.",fr05NoteCold:"Soğuk/Donuk Depolama:",fr05NoteColdText:"Soğuk alanlar 5°C veya altında, donuk depolarda -18°C veya altında muhafaza sağlanmalıdır.",fr05NoteApp:"Uygulama Notu:",fr05NoteAppText:"Denetim sırasında UYGUN görülen durumlar için ilgili kolona (+) yazılacaktır. UYGUN OLMADIĞI görülen durumlar için (-) yazılacaktır.",
    // Label
    labelProduction:"ÜRETİM ETİKETİ",labelFreezer:"DONDURUCU ETİKETİ",labelProductName:"Product Name",labelProductionDate:"Production Date",labelOpeningDate:"Opening Date",labelExpiryDate:"Expiry Date",labelLotNumber:"Lot Number",labelFreezingDate:"Freezing Date",labelFreezingTime:"Freezing Time",labelThawingDate:"Thawing Date",labelThawingTime:"Thawing Time",sampleWarn:"⚠ NUMUNE ŞAHİT — TÜKETİLMEZ ⚠",sampleNameLbl:"Numune Adı:",sampleOrgLbl:"Organizasyon / Yer:",sampleTakerLbl:"Numuneyi Alan Kişi:",sampleDateTimeLbl:"Alım Tarihi ve Saati:",sampleStorageLbl:"Saklama Süresi:",sampleEndLbl:"bitiş:",sampleTempLbl:"Numune Sıcaklığı:",
    // Toplu
    select:"Seç",none:"Yok",today:"Bugün",fresh:"taze",ok:"uygun",old:"eski",
    // Stok
    stockAddManual:"Manuel Stok Ekle",stockEditTitle:"Stok Düzenle",stockCategory:"Kategori",stockMinLevel:"Minimum Stok",stockCurrentLevel:"Mevcut Stok",stockAddBtn:"+ Ekle",stockDetail:"Stok Detayı",stockAllCats:"Tüm Kategoriler",stockSearch:"Ürün ara...",noStock:"Henüz stok yok",invoiceTitle:"Fatura AI",invoiceUpload:"📷 Fatura Fotoğrafı Yükle",invoiceAnalyze:"AI Analiz Et",invoiceItems:"Ürünler",invoiceTotalExtracted:"{count} ürün tespit edildi",addToStock:"+ Stoğa Ekle",packageUnit:"Paket/Birim",packageQty:"Paket Adedi",unitQty:"Birim Başı Miktar",invoiceSaveBtn:"✓ Tümünü Stoğa Ekle",
    // Reçete
    recipeAdd:"Reçete Ekle",recipeEdit:"Reçete Düzenle",recipeSave:"✓ Kaydet",recipeAddIngredient:"+ Malzeme Ekle",recipeMainCat:"Ana Kategori",recipeSubCat:"Alt Kategori",recipeSearch:"Reçete ara...",noRecipes:"Henüz reçete yok",recipePhotoCapture:"📷 Fotoğraf Çek / Yükle",recipeAnalyzing:"AI analiz ediyor...",recipeCostDetail:"Detaylı Maliyet",recipeDelete:"× Sil",recipeDeleteConfirm:"{name} silinecek. Onaylıyor musun?",
    // Menü detay
    menuCreateTitle:"Menü Kartları",menuTypeCocktail:"Kokteyl",menuTypeBreakfast:"Kahvaltı",menuTypeLunch:"Öğle Yemeği",menuTypeDinner:"Akşam Yemeği",menuTypeGala:"Gala / Düğün",menuSaved:"Menü kaydedildi",menuDeleteConfirm:"\"{name}\" silinecek. Onaylıyor musun?",menuAddSection:"+ Bölüm Ekle",menuAddItem:"+ Öğe",menuSectionTitle:"Bölüm Başlığı",menuItemTitle:"Öğe Adı",menuItemDesc:"Açıklama",menuItemCalorie:"Kalori (kcal/100g)",menuItemPrice:"Fiyat",menuAutoLoad:"Reçeteden yükle",menuPreviewBtn:"👁 Önizle",menuEditBtn:"Düzenle",menuPrintBtn:"🖨",menuPageAuto:"Otomatik",menuPageSingle:"Tek sayfa",menuPageMulti:"Çok sayfa",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Paylaş",
    // Ayarlar
    settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"API anahtarı Claude'a OCR ve analiz için gereklidir. https://console.anthropic.com adresinden alabilirsiniz.",settingsProfileNamePh:"Adınızı girin",settingsProfileWorkplacePh:"Örn: Çırağan Sarayı",settingsProfileDeptPh:"Örn: Pastane",settingsProfileRolePh:"Örn: Pastane Şefi",settingsBackupDesc:"Tüm verileri JSON olarak indir veya geri yükle.",settingsRestoreConfirm:"Mevcut veriler üzerine yazılacak. Onaylıyor musun?",settingsRestoreSuccess:"Geri yükleme tamamlandı",settingsDevDesc:"Hata ayıklama için",settingsExpenseName:"Gider adı",settingsExpenseAmount:"Aylık tutar (₺)",settingsPersonnelName:"Kişi adı",settingsPersonnelSalary:"Aylık maaş (₺)",settingsStorageName:"Depo/dolap adı",settingsStorageType:"Tip",settingsStorageTemp:"Hedef sıcaklık (°C)",settingsAddIngredient:"Takip edilecek hammadde adı",settingsAddOrg:"Organizasyon adı",
    // Stok durum
    stockEmpty:"Stok boşaltılıyor...",stockAdded:"Stoğa eklendi",stockDeducted:"Stoktan düşüldü",insufficientStock:"Yetersiz stok: {name}",
    // Boş durumlar
    emptyStateRecipes:"İlk reçeteni oluştur. 📷 AI ile fotoğraftan tanıyabilir veya manuel ekleyebilirsin.",emptyStateStock:"Stoğuna ürün eklemek için 📄 Fatura (AI) veya + Manuel butonlarını kullan.",emptyStateMenus:"İlk menünü oluştur. Reçetelerinden otomatik yükleyebilir veya manuel ekleyebilirsin.",
    // Profil
    profileRequired:"Kullanıcı profili eksik. Ayarlar'a git.",goToSettings:"Ayarlara Git",
    // Sohbet
    newChat:"Yeni Sohbet",msgCount:"mesaj",archived:"Arşivlenenler",archive:"Arşivle",unarchive:"Geri Al",deleteConvConfirm:"Bu sohbet silinecek. Onaylıyor musun?",chatEmptyHint:"Menü planla, maliyet sor, alışveriş listesi çıkar...",askQuestion:"Soru sor...",
    // Arşiv
    archiveTab:"📅 Arşiv",archiveTitle:"Aylık Rapor Arşivi",archiveEmpty:"Henüz arşivlenmiş rapor yok",archiveHint:"💡 Her ay sonunda raporlar otomatik olarak burada arşivlenir. Geçmiş aylara buradan ulaşabilirsin.",archiveExport:"📥 Ay PDF",archiveTotal:"toplam kayıt",
    // Bildirim ayarları
    notificationsTitle:"🔔 Bildirimler",notifEnabled:"Bot bildirimleri aktif",notifDesc:"Bildirimler sohbet ekranına düşer, pop-up olmaz",notifStorage:"Depo kontrolü hatırlatma (09:00, 15:00, 21:00)",notifExpired:"SKT geçen ürün uyarısı",notifLow:"Düşük stok uyarısı",notifLot:"Parti no hatırlatma (22:55)",testNotif:"Test mesajı gönder",
    // Yazıcı
    printersTitle:"🖨 Etiket Yazıcıları",addPrinter:"+ Yazıcı Ekle",printerName:"Yazıcı adı",printerType:"Tür",printerProtocol:"Bağlantı",printerAddress:"Adres (IP/MAC)",printerArea:"Üretim alanı",printerColor:"Etiket rengi",printerTest:"Test",printerWifi:"WiFi",printerBt:"Bluetooth",labelBlue:"Mavi (şok donuk)",labelWhite:"Beyaz (standart)",labelCustomer:"Müşteri QR",
    // QR
    customerQRBtn:"📱 Müşteri QR",customerQRTitle:"Müşteri QR Kodu",customerQRDesc:"Bu QR ürün üzerine yapıştırılabilir. Müşteri okutunca içindekiler, alerjenler ve SKT görünür.",kitchenQR:"Mutfak QR",shareRecipeQR:"📤 Reçete Paylaş",
    // Error mesajları
    errNoPhoto:"Fotoğraf seçin",errNoText:"Metin girin",errNoAPIKey:"API anahtarı yok",errEmptyResponse:"Boş yanıt",errInvalidResponse:"Geçersiz yanıt",
    // Reçete detay
    detailMethod:"YÖNTEM",detailTemp:"SICAKLIK",detailDuration:"SÜRE",detailMultiplier:"Çarpan",detailPortion:"Porsiyon (g)",detailTotalWeight:"Toplam Ağırlık",detailPortionCount:"PORSIYON ADEDİ",detailCalorie:"KALORİ",detailCostAnalysis:"MALİYET ANALİZİ",detailRawMaterial:"Hammadde:",detailFixedShare:"Sabit+Personel payı:",detailRealCost:"Gerçek Maliyet:",detailIngredients:"Malzemeler",detailAddPhoto:"📷 Fotoğraf Ekle",detailPhotoSelect:"Fotoğraf seç veya çek",detailPhotoOption:"Galeri veya kamera",
    // Genel placeholder'lar
    placeholderWorkplace:"Otel / Restoran",placeholderDept:"Departman",placeholderRole:"Pozisyon",placeholderName:"Adınızı girin",
    // Yeni menü
    newMenuName:"Yeni {type} Menüsü",
    // Referans lot
    refLotTitle:"🏷 Referans Lot",refLotFromPhoto:"📷 Fotoğraftan Oku",refLotFromQR:"📷 QR Tara",refLotHint:"Ambalajın fotoğrafını çek veya ürün QR kodunu okut, parti numarası otomatik dolsun",refLotNotFound:"Parti numarası bulunamadı",refLotFound:"Parti no: {lot}",
    // Yeni bölüm
    newSection:"Yeni Bölüm",newSectionPrompt:"Yeni bölüm başlığı:",
    // Pişirme yöntemleri
    methodOven:"Fırın",methodBoil:"Kaynat / Haşla",methodFry:"Kızart",methodGrill:"Izgara",methodBainMarie:"Benmari",methodFridge:"Buzdolabı",methodFreeze:"Dondurma",methodCool:"Soğutma",methodNone:"Pişirme yok"
  },
  en:{
    tabRecipes:"Recipes",tabStock:"Stock",tabProduction:"Production",tabReports:"Reports",tabMenus:"Menus",tabAssistant:"Assistant",tabSettings:"Settings",tabTodo:"Tasks",tabTeam:"Team",tabTeamChat:"Team Chat",
    save:"Save",cancel:"Cancel",delete:"Delete",edit:"Edit",close:"Close",add:"Add",back:"Back",next:"Next",search:"Search",filter:"Filter",all:"All",yes:"Yes",no:"No",confirm:"Confirm",remove:"Remove",update:"Update",name:"Name",amount:"Amount",unit:"Unit",price:"Price",date:"Date",time:"Time",note:"Note",optional:"optional",required:"required",loading:"Loading...",error:"Error",success:"Success",warning:"Warning",
    recipes:"Recipes",newRecipe:"+ New Recipe",recipePhoto:"📷 Recipe (AI)",recipeManual:"+ Manual",recipeDetail:"Recipe Detail",ingredients:"Ingredients",allergens:"Allergens",cookMethod:"Cooking",cookTemp:"Temperature",cookTime:"Time",servings:"Servings",totalWeight:"Total Weight",portionG:"Portion Grams",calories:"Calories",yieldL:"Yield",recipeName:"Recipe Name",deductStock:"📦 Deduct Stock & Produce",sampleLabel:"🧪 Sample Witness Label",editPhoto:"📷 Change Photo",addPhoto:"📷 Add Photo",costAnalysis:"COST ANALYSIS",material:"Materials",overhead:"Fixed+Personnel share",realCost:"Real Cost",perPortion:"portion",costHint:"💡 Add fixed expenses and personnel costs in Settings to see the real cost.",
    stock:"Stock",stockMaterial:"📦 Materials",stockLots:"🏷 Lot Numbers",invoice:"📄 Invoice (AI)",manualAdd:"+ Manual",lowStock:"⚠ Low Stock",products:"items",currentLots:"Current Lot Numbers",noTrackedIngs:"No tracked ingredients",noTrackedIngsHint:"Settings → Traceability → Add Tracked Ingredients",lotMissing:"Lot not entered",lotToday:"today",lotDaysAgo:"days ago",lotEntered:"entered",lotHint:"💡 Scan lot numbers from packages each morning. 🔴 missing · 🔵 current · 🟡 old (3+ days).",
    production:"Production",productionRecord:"🍱 Production Record",productionReports:"📊 Reports",storage:"Storage Location",expiryDays:"Expiry (days)",daysLater:"days later",risk:"Risk",riskHigh:"High",riskMedium:"Medium",riskLow:"Low",riskHint:"Auto-calculated, editable",noteHint:"E.g. For evening service",produceSaveBtn:"✓ Produce & Deduct",portions:"portions",lot:"Lot",produced:"Produced",expires:"Expires",consumed:"consumed",active:"active",fire:"Waste",consumeAll:"✓ All Consumed",printLabel:"🏷 Print Label",partialConsume:"PARTIAL CONSUMPTION",decrease:"− Deduct",markFire:"× Mark as Waste",gotoReport:"📊 Go to Report",noProduction:"No production yet",noProductionHint:"Use \"Deduct Stock\" from recipe detail to create a production record",consumedFull:"{name} ({portions} portions) will be marked as consumed. Confirm?",invalidAmount:"Invalid amount",howManyFire:"How many portions waste? (total {total})",ingredientLotsTitle:"🏷 INGREDIENT LOT NUMBERS",lotMissingHint:"💡 Go to Stock → Lot Numbers for missing lots.",
    sampleTitle:"🧪 Sample Witness Label",sampleDesc:"ISO 22000 compliant sample record — 72 hours storage recommended",sampleName:"Sample Name",sampleTaker:"Sampled By",sampleOrgLabel:"Organization",sampleOrgSelect:"Select...",samplePax:"Pax",sampleLocation:"Extra Info / Location",sampleLocationHint:"E.g. Wedding, VIP table",sampleDateTime:"Collection Time (plating moment)",sampleTemp:"Sample Temperature (°C)",sampleTempHint:"Plating moment",sampleHours:"Storage Duration (hours)",samplePrintSave:"🏷 Print & Save",sampleOrgOrLoc:"Enter organization or location",sampleProfileWarn:"⚠ Enter name in Settings → User Profile",
    reports:"Reports",reportsProduction:"📋 Production Reports",reportsStorage:"🗄 Storage Control",noReports:"No reports yet",checkTotal:"checks",checkComplete:"✓ Complete",addCheck:"+ Add Check",noChecksYet:"No checks recorded yet.",storageCheckHint:"💡 Temperature control 3 times daily per storage (09:00, 15:00, 21:00).",
    productionsCap:"PRODUCTIONS",samplesCap:"SAMPLES",firesCap:"WASTE",
    storageCheckTitle:"🗄 Storage Check",controlTime:"Check Time",controlParams:"CHECK PARAMETERS",paramEtiket:"Label Check",paramAcikGida:"Open Food",paramYerdeDuran:"Floor-standing Food",paramCigPismis:"Raw/Cooked Separation",paramTemizlik:"Cleanliness",paramAlerjen:"Allergen Shelf Separation",paramSkt:"Expiry/Shelf Life",ok:"✓ OK",notOk:"✗ Not OK",productTemp:"Product Temperature (°C)",gaugeValue:"Gauge Value (°C)",explanation:"Explanation (optional)",explanationHint:"Corrective action etc.",atLeastOneTemp:"Enter at least one temperature value",
    menus:"Menu Cards",newMenu:"+ New Menu",chooseMenuType:"Choose Menu Type",menuEdit:"Edit Menu",preview:"Preview",menuName:"Menu Name",font:"Font",theme:"Theme",customSizes:"📐 CUSTOM SIZES (mm)",or:"or",showPrice:"Show Price",showCalorie:"Show Calories",showAllergen:"Show Allergens",bgImage:"🖼 Background Image (optional)",chooseBg:"📷 Choose",changeBg:"📷 Change",cropBg:"✂ Crop",bgOpacity:"Opacity",bgHint:"💡 Add marble, paper, wood texture backgrounds. Crop editor opens after upload.",menuTemplates:"📑 MENU TEMPLATE",loadTemplate:"Load template...",saveTemplate:"💾 Save Template",templateHint:"💡 Save font, theme, size, background options as template.",templateName:"Template name:",defaultTemplateName:"{menuName} Template",templateSaveConfirm:"Load template \"{name}\"? (Sections and items preserved, design changes)",templateSaved:"\"{name}\" saved as template.",templateDeletePrompt:"Template name to delete?",templateNotFound:"Not found",
    cropTitle:"✂ Crop Photo",cropDesc:"Drag to reposition, zoom slider. Area inside the frame becomes menu background.",cropZoom:"🔍 Zoom",cropApply:"✓ Crop & Apply",
    settings:"Settings",apiKey:"API Key",language:"Language",darkMode:"Dark Theme",userProfile:"👤 User Profile",fullName:"Full Name",workplace:"Workplace",department:"Department",role:"Role",profileHint:"💡 Used in sample witness labels and future production records.",traceabilityTitle:"🔍 Traceability (ISO 22000)",traceabilityMode:"Traceability Mode",traceabilityDesc:"Enables lot tracking, ISO 22000 reports (FR.06, FR.12), and PDF outputs",resetHourLabel:"End-of-day Lot Reset Time",resetHourDesc:"lot numbers reset daily at this time",trackedIngsLabel:"Tracked Ingredients",trackedIngsPh:"Add ingredient (e.g. Flour, Milk, Eggs)",trackedIngsHint:"💡 Update these ingredient lot numbers every morning. Auto-used in FR.06.",organizationsTitle:"🏢 Organizations (For Samples)",organizationsPh:"Add organization (e.g. Wedding 200 pax, Gala)",organizationsHint:"💡 This list is used as dropdown in sample witness label.",expensesTitle:"💼 Business Expenses",fixedExpenses:"FIXED EXPENSES (monthly)",personnel:"PERSONNEL (monthly salary)",addFixed:"+ Add Fixed Expense",addPersonnel:"+ Add Personnel",monthlyPortions:"MONTHLY ESTIMATED PRODUCTION (portions)",portionHint:"Used to calculate overhead share in recipe costs.",fixedTotal:"Fixed expense total:",personnelTotal:"Personnel total:",monthlyTotal:"Monthly total expense:",perPortionShare:"Per-portion expense share:",storageAreas:"🗄 Storage Areas",addStorage:"+ Add Storage Area",storageHint:"💡 Add your kitchen's refrigerator/freezer/storage names here.",backup:"💾 Backup",backupBtn:"⬇ Backup",restoreBtn:"⬆ Restore",developer:"🔧 Developer",copyLogs:"📋 Copy Logs",
    fr06Title:"FOOD PREPARATION CONTROL FORM",fr06ProductCtrl:"Product Preparation Control",fr06ProductIng:"Product · Ingredients · Lot No",fr06PrepStart:"Start",fr06PrepEnd:"End",fr06ProductTemp:"Prod °C",fr06AmbientTemp:"Amb °C",fr06Controller:"Controller",fr06CookCtrl:"Cooking / Reheating Temperature Control",fr06ProductDef:"Product Definition",fr06PartyNo:"Lot No",fr06Cook:"Cooking (°C)",fr06Reheat:"Reheating (°C)",fr06CorrectiveAction:"Corrective Action",fr06FreezingCtrl:"Frozen Thawing Control",fr06ThawStart:"Thawing Start",fr06ThawEnd:"Thawing End",fr06Temp:"Temperature",fr06FastCoolCtrl:"Rapid Cooling / Freezing Control",fr06Entry:"Entry (Time-°C)",fr06Exit:"Exit (Time-°C)",fr06Chiller:"Chiller",fr06IceWater:"Ice Water/-18°C",fr06PlateHold:"Plate Holding",fr06ProductOrg:"Product / Org.",fr06PlateTime:"Plating Time",fr06Service:"Service",fr06InnerTemp:"Core Temp",fr06Ctrl:"Control",fr06FosterHold:"Foster/Drying Holding",fr06Notes:"Application Notes",fr06NoteRisk:"Risky Food Preparation:",fr06NoteRiskText:"Ambient temp should not exceed +24°C. For high-risk (meat-poultry-fish) max 16°C. Prep time: 45 min at +20°C, 10 min at +25°C.",fr06NoteCool:"Cooling/Freezing:",fr06NoteCoolText:"Max 10°C within 90 min. Never exceed 4 hours to reach 5°C. Blast freezing: -18°C within 4 hours.",fr06NoteCook:"Cooking/Reheating:",fr06NoteCookText:"Core temp 75°C, surface 65°C. Reheating min 82°C.",fr06NotePlate:"Plate Holding:",fr06NotePlateText:"No holding over 15 min allowed.",fr06NoteThaw:"Thawing:",fr06NoteThawText:"Ice water thawing allowed. Frozen items taken out 1 day in advance.",fr06DocNo:"Document No",fr06DateLbl:"Date",fr06DefaultDept:"Kitchen / Section",
    fr12Title:"SAMPLE WITNESS FORM",fr12SampleOrg:"Sample Name / Organization",fr12Pax:"Pax",fr12DateTime:"Collection Date and Time",fr12Temp:"Sample Temperature",fr12Taker:"Sampled By",fr12Signature:"Signature",fr12Explanation:"Explanation",fr12Note:"Sample Witness:",fr12NoteText:"Witness samples must be taken for all organizations serving over 20 people (including breakfast and staff cafeteria) and stored at max 5°C for 72 hours. Sample must be minimum 200 grams.",
    fr05Title:"DAILY STORAGE / CABINET CONTROL FORM",fr05Dept:"DEPARTMENT:",fr05Cabinet:"CABINET/STORAGE:",fr05Target:"Target:",fr05CtrlParams:"CONTROL PARAMETERS",fr05TempCtrl:"TEMPERATURE CONTROL (°C)",fr05CtrlTime:"Check Time",fr05NoteDry:"Dry Storage:",fr05NoteDryText:"Temperature <25°C and relative humidity ≤65%.",fr05NoteCold:"Cold/Frozen Storage:",fr05NoteColdText:"Cold areas at 5°C or below, frozen storage at -18°C or below.",fr05NoteApp:"Application Note:",fr05NoteAppText:"During inspection, write (+) for compliant situations, (-) for non-compliant.",
    labelProduction:"PRODUCTION LABEL",labelFreezer:"FREEZER LABEL",labelProductName:"Product Name",labelProductionDate:"Production Date",labelOpeningDate:"Opening Date",labelExpiryDate:"Expiry Date",labelLotNumber:"Lot Number",labelFreezingDate:"Freezing Date",labelFreezingTime:"Freezing Time",labelThawingDate:"Thawing Date",labelThawingTime:"Thawing Time",sampleWarn:"⚠ SAMPLE WITNESS — DO NOT CONSUME ⚠",sampleNameLbl:"Sample Name:",sampleOrgLbl:"Organization / Location:",sampleTakerLbl:"Sampled By:",sampleDateTimeLbl:"Collection Date and Time:",sampleStorageLbl:"Storage Duration:",sampleEndLbl:"end:",sampleTempLbl:"Sample Temperature:",
    select:"Select",none:"None",today:"Today",fresh:"fresh",ok:"ok",old:"old",
    stockAddManual:"Add Stock Manually",stockEditTitle:"Edit Stock",stockCategory:"Category",stockMinLevel:"Minimum Stock",stockCurrentLevel:"Current Stock",stockAddBtn:"+ Add",stockDetail:"Stock Detail",stockAllCats:"All Categories",stockSearch:"Search items...",noStock:"No stock yet",invoiceTitle:"Invoice AI",invoiceUpload:"📷 Upload Invoice Photo",invoiceAnalyze:"AI Analyze",invoiceItems:"Items",invoiceTotalExtracted:"{count} items detected",addToStock:"+ Add to Stock",packageUnit:"Package/Unit",packageQty:"Package Count",unitQty:"Amount Per Unit",invoiceSaveBtn:"✓ Add All to Stock",
    recipeAdd:"Add Recipe",recipeEdit:"Edit Recipe",recipeSave:"✓ Save",recipeAddIngredient:"+ Add Ingredient",recipeMainCat:"Main Category",recipeSubCat:"Sub Category",recipeSearch:"Search recipes...",noRecipes:"No recipes yet",recipePhotoCapture:"📷 Capture / Upload Photo",recipeAnalyzing:"AI analyzing...",recipeCostDetail:"Detailed Cost",recipeDelete:"× Delete",recipeDeleteConfirm:"{name} will be deleted. Confirm?",
    menuCreateTitle:"Menu Cards",menuTypeCocktail:"Cocktail",menuTypeBreakfast:"Breakfast",menuTypeLunch:"Lunch",menuTypeDinner:"Dinner",menuTypeGala:"Gala / Wedding",menuSaved:"Menu saved",menuDeleteConfirm:"\"{name}\" will be deleted. Confirm?",menuAddSection:"+ Add Section",menuAddItem:"+ Item",menuSectionTitle:"Section Title",menuItemTitle:"Item Name",menuItemDesc:"Description",menuItemCalorie:"Calorie (kcal/100g)",menuItemPrice:"Price",menuAutoLoad:"Load from recipe",menuPreviewBtn:"👁 Preview",menuEditBtn:"Edit",menuPrintBtn:"🖨",menuPageAuto:"Auto",menuPageSingle:"Single page",menuPageMulti:"Multi page",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Share",
    settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"API key is required for Claude OCR and analysis. Get one at https://console.anthropic.com",settingsProfileNamePh:"Enter your name",settingsProfileWorkplacePh:"E.g. Grand Hotel",settingsProfileDeptPh:"E.g. Pastry",settingsProfileRolePh:"E.g. Pastry Chef",settingsBackupDesc:"Download all data as JSON or restore.",settingsRestoreConfirm:"Existing data will be overwritten. Confirm?",settingsRestoreSuccess:"Restore complete",settingsDevDesc:"For debugging",settingsExpenseName:"Expense name",settingsExpenseAmount:"Monthly amount (₺)",settingsPersonnelName:"Person name",settingsPersonnelSalary:"Monthly salary (₺)",settingsStorageName:"Storage name",settingsStorageType:"Type",settingsStorageTemp:"Target temperature (°C)",settingsAddIngredient:"Ingredient name to track",settingsAddOrg:"Organization name",
    stockEmpty:"Clearing stock...",stockAdded:"Added to stock",stockDeducted:"Deducted from stock",insufficientStock:"Insufficient stock: {name}",
    emptyStateRecipes:"Create your first recipe. Use 📷 AI to detect from photo or add manually.",emptyStateStock:"Use 📄 Invoice (AI) or + Manual buttons to add items to stock.",emptyStateMenus:"Create your first menu. Auto-load from recipes or add manually.",
    profileRequired:"User profile missing. Go to Settings.",goToSettings:"Go to Settings",
    newChat:"New Chat",msgCount:"messages",archived:"Archived",archive:"Archive",unarchive:"Unarchive",deleteConvConfirm:"This chat will be deleted. Confirm?",chatEmptyHint:"Plan menu, ask about costs, create shopping list...",askQuestion:"Ask a question...",
    archiveTab:"📅 Archive",archiveTitle:"Monthly Report Archive",archiveEmpty:"No archived reports yet",archiveHint:"💡 Reports auto-archive at month end. Access past months here.",archiveExport:"📥 Month PDF",archiveTotal:"total records",
    notificationsTitle:"🔔 Notifications",notifEnabled:"Bot notifications active",notifDesc:"Notifications appear in chat, no pop-ups",notifStorage:"Storage check reminder (09:00, 15:00, 21:00)",notifExpired:"Expired product alert",notifLow:"Low stock alert",notifLot:"Lot number reminder (22:55)",testNotif:"Send test message",
    printersTitle:"🖨 Label Printers",addPrinter:"+ Add Printer",printerName:"Printer name",printerType:"Type",printerProtocol:"Connection",printerAddress:"Address (IP/MAC)",printerArea:"Production area",printerColor:"Label color",printerTest:"Test",printerWifi:"WiFi",printerBt:"Bluetooth",labelBlue:"Blue (frozen)",labelWhite:"White (standard)",labelCustomer:"Customer QR",
    customerQRBtn:"📱 Customer QR",customerQRTitle:"Customer QR Code",customerQRDesc:"This QR can be stuck on the product. Customers scan to see ingredients, allergens and expiry date.",kitchenQR:"Kitchen QR",shareRecipeQR:"📤 Share Recipe",
    errNoPhoto:"Select a photo",errNoText:"Enter text",errNoAPIKey:"No API key",errEmptyResponse:"Empty response",errInvalidResponse:"Invalid response",
    detailMethod:"METHOD",detailTemp:"TEMPERATURE",detailDuration:"DURATION",detailMultiplier:"Multiplier",detailPortion:"Portion (g)",detailTotalWeight:"Total Weight",detailPortionCount:"PORTIONS",detailCalorie:"CALORIES",detailCostAnalysis:"COST ANALYSIS",detailRawMaterial:"Raw materials:",detailFixedShare:"Fixed+Personnel share:",detailRealCost:"Real Cost:",detailIngredients:"Ingredients",detailAddPhoto:"📷 Add Photo",detailPhotoSelect:"Select or capture photo",detailPhotoOption:"Gallery or camera",
    placeholderWorkplace:"Hotel / Restaurant",placeholderDept:"Department",placeholderRole:"Position",placeholderName:"Enter your name",
    newMenuName:"New {type} Menu",
    refLotTitle:"🏷 Reference Lot",refLotFromPhoto:"📷 Read from Photo",refLotFromQR:"📷 Scan QR",refLotHint:"Take photo of package or scan product QR to auto-fill lot number",refLotNotFound:"Lot number not found",refLotFound:"Lot: {lot}",
    newSection:"New Section",newSectionPrompt:"New section title:",
    methodOven:"Oven",methodBoil:"Boil",methodFry:"Fry",methodGrill:"Grill",methodBainMarie:"Bain-marie",methodFridge:"Refrigerator",methodFreeze:"Freezing",methodCool:"Cooling",methodNone:"No cooking"
  }
};

// Diğer diller TR fallback ile çalışır (Tur 5b/c'de tamamlanacak: ru, es, de, fr, zh, ar)
I18N.ru={
  tabRecipes:"Рецепты",tabStock:"Склад",tabProduction:"Производство",tabReports:"Отчёты",tabMenus:"Меню",tabAssistant:"Ассистент",tabSettings:"Настройки",
  save:"Сохранить",cancel:"Отмена",delete:"Удалить",edit:"Редактировать",close:"Закрыть",add:"Добавить",back:"Назад",next:"Далее",search:"Поиск",filter:"Фильтр",all:"Все",yes:"Да",no:"Нет",confirm:"Подтвердить",remove:"Убрать",update:"Обновить",name:"Название",amount:"Кол-во",unit:"Ед.",price:"Цена",date:"Дата",time:"Время",note:"Заметка",optional:"необязательно",required:"обязательно",loading:"Загрузка...",error:"Ошибка",success:"Успех",warning:"Предупреждение",
  recipes:"Рецепты",newRecipe:"+ Новый рецепт",recipePhoto:"📷 Рецепт (AI)",recipeManual:"+ Вручную",recipeDetail:"Детали рецепта",ingredients:"Ингредиенты",allergens:"Аллергены",cookMethod:"Приготовление",cookTemp:"Температура",cookTime:"Время",servings:"Порции",totalWeight:"Общий вес",portionG:"Граммы порции",calories:"Калории",yieldL:"Выход",recipeName:"Название рецепта",deductStock:"📦 Списать со склада и произвести",sampleLabel:"🧪 Этикетка образца",editPhoto:"📷 Изменить фото",addPhoto:"📷 Добавить фото",costAnalysis:"АНАЛИЗ СТОИМОСТИ",material:"Материалы",overhead:"Постоянные+Персонал",realCost:"Реальная стоимость",perPortion:"порция",costHint:"💡 Добавьте постоянные расходы и персонал в Настройках для расчета реальной стоимости.",
  stock:"Склад",stockMaterial:"📦 Сырьё",stockLots:"🏷 Номера партий",invoice:"📄 Накладная (AI)",manualAdd:"+ Вручную",lowStock:"⚠ Низкий запас",products:"товаров",currentLots:"Текущие номера партий",noTrackedIngs:"Нет отслеживаемых ингредиентов",noTrackedIngsHint:"Настройки → Прослеживаемость → Добавить ингредиенты",lotMissing:"Партия не введена",lotToday:"сегодня",lotDaysAgo:"дн. назад",lotEntered:"введено",lotHint:"💡 Сканируйте номера партий с упаковок каждое утро. 🔴 нет · 🔵 актуально · 🟡 старое (3+ дня).",
  production:"Производство",productionRecord:"🍱 Запись о производстве",productionReports:"📊 Отчёты",storage:"Место хранения",expiryDays:"Срок годности (дней)",daysLater:"дней",risk:"Риск",riskHigh:"Высокий",riskMedium:"Средний",riskLow:"Низкий",riskHint:"Авторасчёт, можно изменить",noteHint:"Напр.: Для вечерней службы",produceSaveBtn:"✓ Произвести и списать",portions:"порций",lot:"Партия",produced:"Произведено",expires:"До",consumed:"использовано",active:"активно",fire:"Брак",consumeAll:"✓ Всё использовано",printLabel:"🏷 Печать этикетки",partialConsume:"ЧАСТИЧНОЕ ИСПОЛЬЗОВАНИЕ",decrease:"− Списать",markFire:"× Отметить как брак",gotoReport:"📊 К отчёту",noProduction:"Пока нет производства",noProductionHint:"Создайте запись через \"Списать со склада\" в деталях рецепта",consumedFull:"{name} ({portions} порций) будет отмечено как использованное. Подтвердить?",invalidAmount:"Неверное количество",howManyFire:"Сколько порций в брак? (всего {total})",ingredientLotsTitle:"🏷 НОМЕРА ПАРТИЙ СЫРЬЯ",lotMissingHint:"💡 Для отсутствующих партий перейдите в Склад → Номера партий.",
  sampleTitle:"🧪 Этикетка образца",sampleDesc:"Запись образца по ISO 22000 — рекомендуется хранение 72 часа",sampleName:"Название образца",sampleTaker:"Отбор произвёл",sampleOrgLabel:"Организация",sampleOrgSelect:"Выбрать...",samplePax:"Персон",sampleLocation:"Доп. информация / Место",sampleLocationHint:"Напр.: Свадьба, VIP стол",sampleDateTime:"Дата и время отбора",sampleTemp:"Температура образца (°C)",sampleTempHint:"В момент подачи",sampleHours:"Длительность хранения (часов)",samplePrintSave:"🏷 Печать и сохранить",sampleOrgOrLoc:"Введите организацию или место",sampleProfileWarn:"⚠ Введите имя в Настройки → Профиль",
  reports:"Отчёты",reportsProduction:"📋 Отчёты производства",reportsStorage:"🗄 Контроль хранения",noReports:"Отчётов пока нет",checkTotal:"проверок",checkComplete:"✓ Полно",addCheck:"+ Добавить проверку",noChecksYet:"Проверки ещё не записаны.",storageCheckHint:"💡 Контроль температуры 3 раза в день (09:00, 15:00, 21:00).",
  notificationsTitle:"🔔 Уведомления",notifEnabled:"Бот-уведомления активны",notifDesc:"Уведомления появляются в чате, без всплывающих окон",notifStorage:"Напоминание о контроле хранения (09:00, 15:00, 21:00)",notifExpired:"Предупреждение об истёкшем сроке",notifLow:"Предупреждение о низком запасе",testNotif:"Отправить тест",
  userProfile:"Профиль пользователя",fullName:"Полное имя",workplace:"Место работы",department:"Отдел",role:"Должность",profileHint:"Эти данные используются в формах FR.05 и FR.06",placeholderName:"Введите имя",placeholderWorkplace:"Отель / Ресторан",placeholderDept:"Отдел",placeholderRole:"Должность",tabTodo:"Задачи",
  productionsCap:"ПРОИЗВОДСТВО",samplesCap:"ОБРАЗЦЫ",firesCap:"БРАК",
  storageCheckTitle:"🗄 Проверка хранения",controlTime:"Время проверки",controlParams:"ПАРАМЕТРЫ ПРОВЕРКИ",paramEtiket:"Проверка этикеток",paramAcikGida:"Открытая еда",paramYerdeDuran:"Еда на полу",paramCigPismis:"Разделение сырое/готовое",paramTemizlik:"Чистота",paramAlerjen:"Разделение аллергенов",paramSkt:"Срок годности",ok:"✓ OK",notOk:"✗ Не OK",productTemp:"Температура продукта (°C)",gaugeValue:"Показание прибора (°C)",explanation:"Объяснение (необязательно)",explanationHint:"Корректирующие действия и т.д.",atLeastOneTemp:"Введите хотя бы одно значение температуры",
  menus:"Карты меню",newMenu:"+ Новое меню",chooseMenuType:"Выберите тип меню",menuEdit:"Редактировать меню",preview:"Предпросмотр",menuName:"Название меню",font:"Шрифт",theme:"Тема",customSizes:"📐 СВОИ РАЗМЕРЫ (мм)",or:"или",showPrice:"Показывать цену",showCalorie:"Показывать калории",showAllergen:"Показывать аллергены",bgImage:"🖼 Фоновое изображение (необязательно)",chooseBg:"📷 Выбрать",changeBg:"📷 Изменить",cropBg:"✂ Обрезать",bgOpacity:"Прозрачность",bgHint:"💡 Добавьте фоны с текстурой мрамора, бумаги, дерева. После загрузки откроется редактор обрезки.",menuTemplates:"📑 ШАБЛОН МЕНЮ",loadTemplate:"Загрузить шаблон...",saveTemplate:"💾 Сохранить шаблон",templateHint:"💡 Сохраните шрифт, тему, размер, фон как шаблон.",templateName:"Название шаблона:",defaultTemplateName:"Шаблон {menuName}",templateSaveConfirm:"Загрузить шаблон \"{name}\"? (Разделы и позиции сохраняются, дизайн меняется)",templateSaved:"\"{name}\" сохранено как шаблон.",templateDeletePrompt:"Название шаблона для удаления?",templateNotFound:"Не найдено",
  cropTitle:"✂ Обрезать фото",cropDesc:"Перетащите для позиционирования, ползунок для зума. Область внутри рамки станет фоном меню.",cropZoom:"🔍 Зум",cropApply:"✓ Обрезать и применить",
  settings:"Настройки",apiKey:"API ключ",language:"Язык",darkMode:"Тёмная тема",userProfile:"👤 Профиль пользователя",fullName:"ФИО",workplace:"Место работы",department:"Отдел",role:"Роль",profileHint:"💡 Используется в этикетках образцов и будущих записях производства.",traceabilityTitle:"🔍 Прослеживаемость (ISO 22000)",traceabilityMode:"Режим прослеживаемости",traceabilityDesc:"Включает отслеживание партий, отчёты ISO 22000 (FR.06, FR.12) и PDF вывод",resetHourLabel:"Время сброса партий в конце дня",resetHourDesc:"номера партий сбрасываются ежедневно в это время",trackedIngsLabel:"Отслеживаемые ингредиенты",trackedIngsPh:"Добавить ингредиент (напр.: Мука, Молоко, Яйца)",trackedIngsHint:"💡 Обновляйте номера партий этих ингредиентов каждое утро. Авто-использование в FR.06.",organizationsTitle:"🏢 Организации (для образцов)",organizationsPh:"Добавить организацию (напр.: Свадьба 200 чел., Гала)",organizationsHint:"💡 Этот список используется как выпадающий в этикетке образца.",expensesTitle:"💼 Бизнес-расходы",fixedExpenses:"ПОСТОЯННЫЕ РАСХОДЫ (в месяц)",personnel:"ПЕРСОНАЛ (месячная зарплата)",addFixed:"+ Добавить постоянный расход",addPersonnel:"+ Добавить персонал",monthlyPortions:"МЕСЯЧНОЕ ОЦЕНОЧНОЕ ПРОИЗВОДСТВО (порции)",portionHint:"Используется для расчёта доли накладных расходов в стоимости рецепта.",fixedTotal:"Итого постоянных расходов:",personnelTotal:"Итого персонала:",monthlyTotal:"Итого ежемесячных расходов:",perPortionShare:"Доля расходов на порцию:",storageAreas:"🗄 Места хранения",addStorage:"+ Добавить место хранения",storageHint:"💡 Добавьте сюда названия ваших холодильников/морозильников/складов.",backup:"💾 Резервная копия",backupBtn:"⬇ Резервная копия",restoreBtn:"⬆ Восстановить",developer:"🔧 Разработчик",copyLogs:"📋 Копировать логи",
  fr06Title:"ФОРМА КОНТРОЛЯ ПОДГОТОВКИ ПИЩИ",fr06ProductCtrl:"Контроль подготовки продукта",fr06ProductIng:"Название продукта · Ингредиенты · № партии",fr06PrepStart:"Нач. подгот.",fr06PrepEnd:"Оконч. подгот.",fr06ProductTemp:"Темп. продукта (°C)",fr06AmbientTemp:"Темп. среды (°C)",fr06Controller:"Контролёр",fr06CookCtrl:"Контроль температуры готовки / повторного нагрева",fr06ProductDef:"Название продукта",fr06PartyNo:"№ партии",fr06Cook:"Готовка (°C)",fr06Reheat:"Повт. нагрев (°C)",fr06CorrectiveAction:"Корректирующее действие",fr06FreezingCtrl:"Контроль разморозки",fr06ThawStart:"Начало разморозки",fr06ThawEnd:"Конец разморозки",fr06Temp:"Температура",fr06FastCoolCtrl:"Контроль быстрого охлаждения / заморозки",fr06Entry:"Вход (Время-°C)",fr06Exit:"Выход (Время-°C)",fr06Chiller:"Чиллер",fr06IceWater:"Ледяная вода/-18°C",fr06PlateHold:"Выдержка на тарелке",fr06ProductOrg:"Продукт / Орг.",fr06PlateTime:"Время выкладки",fr06Service:"Сервис",fr06InnerTemp:"Внутр. темп.",fr06Ctrl:"Контроль",fr06FosterHold:"Выдержка Foster/сушка",fr06Notes:"Замечания по применению",fr06NoteRisk:"Подготовка рискованной пищи:",fr06NoteRiskText:"Температура среды не должна превышать +24°C. Для высокорисковых (мясо-птица-рыба) макс. 16°C. Время подготовки: 45 мин при +20°C, 10 мин при +25°C.",fr06NoteCool:"Охлаждение/Заморозка:",fr06NoteCoolText:"Макс. 10°C за 90 минут. Никогда не превышать 4 часа до 5°C. Шоковая заморозка: -18°C за 4 часа.",fr06NoteCook:"Готовка/Повт. нагрев:",fr06NoteCookText:"Внутр. темп. 75°C, пов. 65°C. Повт. нагрев мин. 82°C.",fr06NotePlate:"Выдержка на тарелке:",fr06NotePlateText:"Не допускается выдержка более 15 мин.",fr06NoteThaw:"Разморозка:",fr06NoteThawText:"Разрешена разморозка в ледяной воде. Замороженные продукты достают за 1 день.",fr06DocNo:"№ документа",fr06DateLbl:"Дата",fr06DefaultDept:"Кухня / Отдел",
  fr12Title:"ФОРМА СВИДЕТЕЛЯ ОБРАЗЦА",fr12SampleOrg:"Название образца / Организация",fr12Pax:"Персон",fr12DateTime:"Дата и время отбора",fr12Temp:"Температура образца",fr12Taker:"Отбор произвёл",fr12Signature:"Подпись",fr12Explanation:"Объяснение",fr12Note:"Образец-свидетель:",fr12NoteText:"Для всех мероприятий свыше 20 человек (включая завтрак и столовую) должны быть взяты образцы-свидетели и храниться при макс. 5°C в течение 72 часов. Мин. 200 грамм.",
  fr05Title:"ЕЖЕДНЕВНАЯ ФОРМА КОНТРОЛЯ СКЛАДА / ШКАФА",fr05Dept:"ОТДЕЛ:",fr05Cabinet:"ШКАФ/СКЛАД:",fr05Target:"Цель:",fr05CtrlParams:"ПАРАМЕТРЫ КОНТРОЛЯ",fr05TempCtrl:"КОНТРОЛЬ ТЕМПЕРАТУРЫ (°C)",fr05CtrlTime:"Время проверки",fr05NoteDry:"Сухое хранение:",fr05NoteDryText:"Температура <25°C и относительная влажность ≤65%.",fr05NoteCold:"Холодное/замороженное хранение:",fr05NoteColdText:"Холодные зоны при 5°C или ниже, замороженные при -18°C или ниже.",fr05NoteApp:"Замечание по применению:",fr05NoteAppText:"При проверке пишите (+) для соответствующих ситуаций, (-) для несоответствующих.",
  labelProduction:"ЭТИКЕТКА ПРОИЗВОДСТВА",labelFreezer:"ЭТИКЕТКА МОРОЗИЛЬНИКА",labelProductName:"Название продукта",labelProductionDate:"Дата производства",labelOpeningDate:"Дата открытия",labelExpiryDate:"Срок годности",labelLotNumber:"№ партии",labelFreezingDate:"Дата заморозки",labelFreezingTime:"Время заморозки",labelThawingDate:"Дата разморозки",labelThawingTime:"Время разморозки",sampleWarn:"⚠ ОБРАЗЕЦ-СВИДЕТЕЛЬ — НЕ УПОТРЕБЛЯТЬ ⚠",sampleNameLbl:"Название образца:",sampleOrgLbl:"Организация / Место:",sampleTakerLbl:"Отбор произвёл:",sampleDateTimeLbl:"Дата и время отбора:",sampleStorageLbl:"Срок хранения:",sampleEndLbl:"конец:",sampleTempLbl:"Температура образца:",
  select:"Выбрать",none:"Нет",today:"Сегодня",fresh:"свежий",ok:"ок",old:"старый",
  stockAddManual:"Добавить вручную",stockEditTitle:"Редактировать",stockCategory:"Категория",stockMinLevel:"Мин. запас",stockCurrentLevel:"Текущий запас",stockAddBtn:"+ Добавить",stockDetail:"Детали запаса",stockAllCats:"Все категории",stockSearch:"Поиск товаров...",noStock:"Склад пуст",invoiceTitle:"Накладная AI",invoiceUpload:"📷 Загрузить фото накладной",invoiceAnalyze:"AI анализ",invoiceItems:"Позиции",invoiceTotalExtracted:"Найдено {count} позиций",addToStock:"+ На склад",packageUnit:"Упаковка/Ед.",packageQty:"Кол-во упаковок",unitQty:"Кол-во в единице",invoiceSaveBtn:"✓ Добавить всё на склад",
  recipeAdd:"Добавить рецепт",recipeEdit:"Редактировать рецепт",recipeSave:"✓ Сохранить",recipeAddIngredient:"+ Добавить ингредиент",recipeMainCat:"Главная категория",recipeSubCat:"Подкатегория",recipeSearch:"Поиск рецептов...",noRecipes:"Рецептов ещё нет",recipePhotoCapture:"📷 Сделать / загрузить фото",recipeAnalyzing:"AI анализирует...",recipeCostDetail:"Детальная стоимость",recipeDelete:"× Удалить",recipeDeleteConfirm:"{name} будет удалён. Подтвердить?",
  menuCreateTitle:"Карты меню",menuTypeCocktail:"Коктейль",menuTypeBreakfast:"Завтрак",menuTypeLunch:"Обед",menuTypeDinner:"Ужин",menuTypeGala:"Гала / Свадьба",menuSaved:"Меню сохранено",menuDeleteConfirm:"\"{name}\" будет удалено. Подтвердить?",menuAddSection:"+ Добавить раздел",menuAddItem:"+ Позиция",menuSectionTitle:"Название раздела",menuItemTitle:"Название позиции",menuItemDesc:"Описание",menuItemCalorie:"Калории (ккал/100г)",menuItemPrice:"Цена",menuAutoLoad:"Загрузить из рецепта",menuPreviewBtn:"👁 Предпросмотр",menuEditBtn:"Изменить",menuPrintBtn:"🖨",menuPageAuto:"Авто",menuPageSingle:"Одна страница",menuPageMulti:"Несколько страниц",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Поделиться",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"API ключ нужен для Claude OCR и анализа. Получите на https://console.anthropic.com",settingsProfileNamePh:"Введите имя",settingsProfileWorkplacePh:"Напр. Grand Hotel",settingsProfileDeptPh:"Напр. Кондитерская",settingsProfileRolePh:"Напр. Шеф-кондитер",settingsBackupDesc:"Скачать все данные в JSON или восстановить.",settingsRestoreConfirm:"Существующие данные будут перезаписаны. Подтвердить?",settingsRestoreSuccess:"Восстановление завершено",settingsDevDesc:"Для отладки",settingsExpenseName:"Название расхода",settingsExpenseAmount:"Сумма в месяц (₺)",settingsPersonnelName:"Имя сотрудника",settingsPersonnelSalary:"Месячная зарплата (₺)",settingsStorageName:"Название хранилища",settingsStorageType:"Тип",settingsStorageTemp:"Целевая температура (°C)",settingsAddIngredient:"Название ингредиента для отслеживания",settingsAddOrg:"Название организации",
  stockEmpty:"Очистка склада...",stockAdded:"Добавлено на склад",stockDeducted:"Списано со склада",insufficientStock:"Недостаточно: {name}",
  emptyStateRecipes:"Создайте первый рецепт. Используйте 📷 AI или добавьте вручную.",emptyStateStock:"Используйте 📄 Накладная (AI) или + Вручную для добавления товаров.",emptyStateMenus:"Создайте первое меню. Автозагрузка из рецептов или вручную.",
  profileRequired:"Профиль не заполнен. Перейдите в Настройки.",goToSettings:"К настройкам"
};
I18N.es={
  tabRecipes:"Recetas",tabStock:"Stock",tabProduction:"Producción",tabReports:"Informes",tabMenus:"Menús",tabAssistant:"Asistente",tabSettings:"Ajustes",
  save:"Guardar",cancel:"Cancelar",delete:"Eliminar",edit:"Editar",close:"Cerrar",add:"Añadir",back:"Atrás",next:"Siguiente",search:"Buscar",filter:"Filtro",all:"Todo",yes:"Sí",no:"No",confirm:"Confirmar",remove:"Quitar",update:"Actualizar",name:"Nombre",amount:"Cantidad",unit:"Unidad",price:"Precio",date:"Fecha",time:"Hora",note:"Nota",optional:"opcional",required:"requerido",loading:"Cargando...",error:"Error",success:"Éxito",warning:"Advertencia",
  recipes:"Recetas",newRecipe:"+ Nueva receta",recipePhoto:"📷 Receta (IA)",recipeManual:"+ Manual",recipeDetail:"Detalle de receta",ingredients:"Ingredientes",allergens:"Alérgenos",cookMethod:"Cocción",cookTemp:"Temperatura",cookTime:"Tiempo",servings:"Porciones",totalWeight:"Peso total",portionG:"Gramos por porción",calories:"Calorías",yieldL:"Rendimiento",recipeName:"Nombre de receta",deductStock:"📦 Deducir stock y producir",sampleLabel:"🧪 Etiqueta de muestra",editPhoto:"📷 Cambiar foto",addPhoto:"📷 Añadir foto",costAnalysis:"ANÁLISIS DE COSTES",material:"Materiales",overhead:"Fijos+Personal",realCost:"Coste real",perPortion:"porción",costHint:"💡 Añade gastos fijos y personal en Ajustes para ver coste real.",
  stock:"Stock",stockMaterial:"📦 Materiales",stockLots:"🏷 Números de lote",invoice:"📄 Factura (IA)",manualAdd:"+ Manual",lowStock:"⚠ Stock bajo",products:"artículos",currentLots:"Números de lote actuales",noTrackedIngs:"Sin ingredientes rastreados",noTrackedIngsHint:"Ajustes → Trazabilidad → Añadir ingredientes",lotMissing:"Lote no introducido",lotToday:"hoy",lotDaysAgo:"días atrás",lotEntered:"introducido",lotHint:"💡 Escanea los números de lote de los paquetes cada mañana. 🔴 falta · 🔵 actual · 🟡 antiguo (3+ días).",
  production:"Producción",productionRecord:"🍱 Registro de producción",productionReports:"📊 Informes",storage:"Ubicación de almacenamiento",expiryDays:"Caducidad (días)",daysLater:"días",risk:"Riesgo",riskHigh:"Alto",riskMedium:"Medio",riskLow:"Bajo",riskHint:"Auto-calculado, editable",noteHint:"Ej. Para servicio de tarde",produceSaveBtn:"✓ Producir y deducir",portions:"porciones",lot:"Lote",produced:"Producido",expires:"Caduca",consumed:"consumido",active:"activo",fire:"Desperdicio",consumeAll:"✓ Todo consumido",printLabel:"🏷 Imprimir etiqueta",partialConsume:"CONSUMO PARCIAL",decrease:"− Deducir",markFire:"× Marcar como desperdicio",gotoReport:"📊 Ir al informe",noProduction:"Sin producción todavía",noProductionHint:"Usa \"Deducir stock\" desde el detalle de receta para crear un registro",consumedFull:"{name} ({portions} porciones) se marcará como consumido. ¿Confirmar?",invalidAmount:"Cantidad inválida",howManyFire:"¿Cuántas porciones son desperdicio? (total {total})",ingredientLotsTitle:"🏷 NÚMEROS DE LOTE DE INGREDIENTES",lotMissingHint:"💡 Para lotes faltantes, ve a Stock → Números de lote.",
  sampleTitle:"🧪 Etiqueta de muestra",sampleDesc:"Registro de muestra conforme a ISO 22000 — se recomienda almacenamiento 72 horas",sampleName:"Nombre de muestra",sampleTaker:"Muestreado por",sampleOrgLabel:"Organización",sampleOrgSelect:"Seleccionar...",samplePax:"Personas",sampleLocation:"Información extra / Ubicación",sampleLocationHint:"Ej. Boda, mesa VIP",sampleDateTime:"Fecha y hora de toma",sampleTemp:"Temperatura de muestra (°C)",sampleTempHint:"Momento de emplatado",sampleHours:"Duración de almacenamiento (horas)",samplePrintSave:"🏷 Imprimir y guardar",sampleOrgOrLoc:"Introduce organización o ubicación",sampleProfileWarn:"⚠ Introduce nombre en Ajustes → Perfil",
  reports:"Informes",reportsProduction:"📋 Informes de producción",reportsStorage:"🗄 Control de almacenamiento",noReports:"Sin informes todavía",checkTotal:"controles",checkComplete:"✓ Completo",addCheck:"+ Añadir control",noChecksYet:"Aún no hay controles registrados.",storageCheckHint:"💡 Control de temperatura 3 veces al día (09:00, 15:00, 21:00).",
  notificationsTitle:"🔔 Notificaciones",notifEnabled:"Notificaciones de bot activas",notifDesc:"Las notificaciones aparecen en el chat, sin ventanas emergentes",notifStorage:"Recordatorio de control de almacenamiento (09:00, 15:00, 21:00)",notifExpired:"Alerta de producto caducado",notifLow:"Alerta de stock bajo",testNotif:"Enviar mensaje de prueba",
  userProfile:"Perfil de usuario",fullName:"Nombre completo",workplace:"Lugar de trabajo",department:"Departamento",role:"Cargo",profileHint:"Estos datos se usan en los formularios FR.05 y FR.06",placeholderName:"Ingrese su nombre",placeholderWorkplace:"Hotel / Restaurante",placeholderDept:"Departamento",placeholderRole:"Cargo",tabTodo:"Tareas",
  productionsCap:"PRODUCCIONES",samplesCap:"MUESTRAS",firesCap:"DESPERDICIO",
  storageCheckTitle:"🗄 Control de almacenamiento",controlTime:"Hora de control",controlParams:"PARÁMETROS DE CONTROL",paramEtiket:"Control de etiquetas",paramAcikGida:"Alimento abierto",paramYerdeDuran:"Alimento en el suelo",paramCigPismis:"Separación crudo/cocido",paramTemizlik:"Limpieza",paramAlerjen:"Separación de alérgenos",paramSkt:"Caducidad / Vida útil",ok:"✓ OK",notOk:"✗ No OK",productTemp:"Temperatura del producto (°C)",gaugeValue:"Valor del medidor (°C)",explanation:"Explicación (opcional)",explanationHint:"Acción correctiva etc.",atLeastOneTemp:"Introduce al menos un valor de temperatura",
  menus:"Tarjetas de menú",newMenu:"+ Nuevo menú",chooseMenuType:"Elegir tipo de menú",menuEdit:"Editar menú",preview:"Vista previa",menuName:"Nombre del menú",font:"Fuente",theme:"Tema",customSizes:"📐 TAMAÑOS PERSONALIZADOS (mm)",or:"o",showPrice:"Mostrar precio",showCalorie:"Mostrar calorías",showAllergen:"Mostrar alérgenos",bgImage:"🖼 Imagen de fondo (opcional)",chooseBg:"📷 Elegir",changeBg:"📷 Cambiar",cropBg:"✂ Recortar",bgOpacity:"Opacidad",bgHint:"💡 Añade fondos de mármol, papel, textura de madera. El editor se abre después de subir.",menuTemplates:"📑 PLANTILLA DE MENÚ",loadTemplate:"Cargar plantilla...",saveTemplate:"💾 Guardar plantilla",templateHint:"💡 Guarda fuente, tema, tamaño, fondo como plantilla.",templateName:"Nombre de plantilla:",defaultTemplateName:"Plantilla {menuName}",templateSaveConfirm:"¿Cargar plantilla \"{name}\"? (Secciones e items preservados, diseño cambia)",templateSaved:"\"{name}\" guardado como plantilla.",templateDeletePrompt:"¿Nombre de plantilla a eliminar?",templateNotFound:"No encontrado",
  cropTitle:"✂ Recortar foto",cropDesc:"Arrastra para reposicionar, control deslizante para zoom. El área dentro del marco será el fondo del menú.",cropZoom:"🔍 Zoom",cropApply:"✓ Recortar y aplicar",
  settings:"Ajustes",apiKey:"Clave API",language:"Idioma",darkMode:"Tema oscuro",userProfile:"👤 Perfil de usuario",fullName:"Nombre completo",workplace:"Lugar de trabajo",department:"Departamento",role:"Rol",profileHint:"💡 Usado en etiquetas de muestras y futuros registros de producción.",traceabilityTitle:"🔍 Trazabilidad (ISO 22000)",traceabilityMode:"Modo trazabilidad",traceabilityDesc:"Habilita seguimiento de lotes, informes ISO 22000 (FR.06, FR.12) y salidas PDF",resetHourLabel:"Hora de reinicio de lotes al final del día",resetHourDesc:"los números de lote se reinician diariamente a esta hora",trackedIngsLabel:"Ingredientes rastreados",trackedIngsPh:"Añadir ingrediente (ej. Harina, Leche, Huevos)",trackedIngsHint:"💡 Actualiza los números de lote cada mañana. Auto-usado en FR.06.",organizationsTitle:"🏢 Organizaciones (para muestras)",organizationsPh:"Añadir organización (ej. Boda 200 pax, Gala)",organizationsHint:"💡 Esta lista se usa como desplegable en la etiqueta de muestra.",expensesTitle:"💼 Gastos empresariales",fixedExpenses:"GASTOS FIJOS (mensuales)",personnel:"PERSONAL (salario mensual)",addFixed:"+ Añadir gasto fijo",addPersonnel:"+ Añadir personal",monthlyPortions:"PRODUCCIÓN MENSUAL ESTIMADA (porciones)",portionHint:"Usado para calcular la parte de gastos generales en costes de receta.",fixedTotal:"Total gastos fijos:",personnelTotal:"Total personal:",monthlyTotal:"Total gastos mensuales:",perPortionShare:"Parte de gastos por porción:",storageAreas:"🗄 Áreas de almacenamiento",addStorage:"+ Añadir área de almacenamiento",storageHint:"💡 Añade aquí los nombres de tus neveras/congeladores/almacenes.",backup:"💾 Copia de seguridad",backupBtn:"⬇ Copia",restoreBtn:"⬆ Restaurar",developer:"🔧 Desarrollador",copyLogs:"📋 Copiar logs",
  fr06Title:"FORMULARIO DE CONTROL DE PREPARACIÓN DE ALIMENTOS",fr06ProductCtrl:"Control de preparación de producto",fr06ProductIng:"Nombre del producto · Ingredientes · N° de lote",fr06PrepStart:"Inicio prep.",fr06PrepEnd:"Fin prep.",fr06ProductTemp:"Temp. producto (°C)",fr06AmbientTemp:"Temp. ambiente (°C)",fr06Controller:"Controlador",fr06CookCtrl:"Control de temperatura de cocción / recalentamiento",fr06ProductDef:"Definición del producto",fr06PartyNo:"N° de lote",fr06Cook:"Cocción (°C)",fr06Reheat:"Recalent. (°C)",fr06CorrectiveAction:"Acción correctiva",fr06FreezingCtrl:"Control de descongelación",fr06ThawStart:"Inicio descongelación",fr06ThawEnd:"Fin descongelación",fr06Temp:"Temperatura",fr06FastCoolCtrl:"Control de enfriamiento rápido / congelación",fr06Entry:"Entrada (Hora-°C)",fr06Exit:"Salida (Hora-°C)",fr06Chiller:"Enfriador",fr06IceWater:"Agua con hielo/-18°C",fr06PlateHold:"Espera en plato",fr06ProductOrg:"Producto / Org.",fr06PlateTime:"Hora de emplatado",fr06Service:"Servicio",fr06InnerTemp:"Temp. interna",fr06Ctrl:"Control",fr06FosterHold:"Espera Foster/secado",fr06Notes:"Notas de aplicación",fr06NoteRisk:"Preparación de alimentos de riesgo:",fr06NoteRiskText:"La temperatura ambiente no debe exceder +24°C. Para alimentos de alto riesgo (carne-pollo-pescado) máx 16°C. Tiempo de preparación: 45 min a +20°C, 10 min a +25°C.",fr06NoteCool:"Enfriamiento/Congelación:",fr06NoteCoolText:"Máx 10°C en 90 min. Nunca exceder 4 horas para alcanzar 5°C. Congelación rápida: -18°C en 4 horas.",fr06NoteCook:"Cocción/Recalentamiento:",fr06NoteCookText:"Temp. interna 75°C, superficie 65°C. Recalentamiento mín 82°C.",fr06NotePlate:"Espera en plato:",fr06NotePlateText:"No se permite espera superior a 15 min.",fr06NoteThaw:"Descongelación:",fr06NoteThawText:"Descongelación en agua con hielo permitida. Los productos congelados se sacan con 1 día de antelación.",fr06DocNo:"N° de documento",fr06DateLbl:"Fecha",fr06DefaultDept:"Cocina / Sección",
  fr12Title:"FORMULARIO DE TESTIGO DE MUESTRA",fr12SampleOrg:"Nombre de muestra / Organización",fr12Pax:"Personas",fr12DateTime:"Fecha y hora de recolección",fr12Temp:"Temperatura de muestra",fr12Taker:"Muestreado por",fr12Signature:"Firma",fr12Explanation:"Explicación",fr12Note:"Muestra testigo:",fr12NoteText:"Se deben tomar muestras testigo para todas las organizaciones que sirvan a más de 20 personas (incluyendo desayuno y comedor de personal) y almacenarse a máx 5°C durante 72 horas. Mínimo 200 gramos.",
  fr05Title:"FORMULARIO DE CONTROL DIARIO DE ALMACÉN / ARMARIO",fr05Dept:"DEPARTAMENTO:",fr05Cabinet:"ARMARIO/ALMACÉN:",fr05Target:"Objetivo:",fr05CtrlParams:"PARÁMETROS DE CONTROL",fr05TempCtrl:"CONTROL DE TEMPERATURA (°C)",fr05CtrlTime:"Hora de control",fr05NoteDry:"Almacenamiento seco:",fr05NoteDryText:"Temperatura <25°C y humedad relativa ≤65%.",fr05NoteCold:"Almacenamiento frío/congelado:",fr05NoteColdText:"Zonas frías a 5°C o menos, almacenamiento congelado a -18°C o menos.",fr05NoteApp:"Nota de aplicación:",fr05NoteAppText:"Durante la inspección, escribe (+) para situaciones conformes, (-) para no conformes.",
  labelProduction:"ETIQUETA DE PRODUCCIÓN",labelFreezer:"ETIQUETA DE CONGELADOR",labelProductName:"Nombre del producto",labelProductionDate:"Fecha de producción",labelOpeningDate:"Fecha de apertura",labelExpiryDate:"Fecha de caducidad",labelLotNumber:"N° de lote",labelFreezingDate:"Fecha de congelación",labelFreezingTime:"Hora de congelación",labelThawingDate:"Fecha de descongelación",labelThawingTime:"Hora de descongelación",sampleWarn:"⚠ MUESTRA TESTIGO — NO CONSUMIR ⚠",sampleNameLbl:"Nombre de muestra:",sampleOrgLbl:"Organización / Ubicación:",sampleTakerLbl:"Muestreado por:",sampleDateTimeLbl:"Fecha y hora de recolección:",sampleStorageLbl:"Duración de almacenamiento:",sampleEndLbl:"fin:",sampleTempLbl:"Temperatura de muestra:",
  select:"Seleccionar",none:"Ninguno",today:"Hoy",fresh:"fresco",ok:"ok",old:"antiguo",
  stockAddManual:"Añadir manualmente",stockEditTitle:"Editar stock",stockCategory:"Categoría",stockMinLevel:"Stock mínimo",stockCurrentLevel:"Stock actual",stockAddBtn:"+ Añadir",stockDetail:"Detalle de stock",stockAllCats:"Todas las categorías",stockSearch:"Buscar items...",noStock:"Sin stock todavía",invoiceTitle:"Factura IA",invoiceUpload:"📷 Subir foto de factura",invoiceAnalyze:"Analizar con IA",invoiceItems:"Artículos",invoiceTotalExtracted:"{count} items detectados",addToStock:"+ Al stock",packageUnit:"Paquete/Unidad",packageQty:"Cantidad de paquetes",unitQty:"Cantidad por unidad",invoiceSaveBtn:"✓ Añadir todo al stock",
  recipeAdd:"Añadir receta",recipeEdit:"Editar receta",recipeSave:"✓ Guardar",recipeAddIngredient:"+ Añadir ingrediente",recipeMainCat:"Categoría principal",recipeSubCat:"Subcategoría",recipeSearch:"Buscar recetas...",noRecipes:"Sin recetas todavía",recipePhotoCapture:"📷 Capturar / Subir foto",recipeAnalyzing:"IA analizando...",recipeCostDetail:"Coste detallado",recipeDelete:"× Eliminar",recipeDeleteConfirm:"{name} será eliminado. ¿Confirmar?",
  menuCreateTitle:"Tarjetas de menú",menuTypeCocktail:"Cóctel",menuTypeBreakfast:"Desayuno",menuTypeLunch:"Almuerzo",menuTypeDinner:"Cena",menuTypeGala:"Gala / Boda",menuSaved:"Menú guardado",menuDeleteConfirm:"\"{name}\" será eliminado. ¿Confirmar?",menuAddSection:"+ Añadir sección",menuAddItem:"+ Item",menuSectionTitle:"Título de sección",menuItemTitle:"Nombre del item",menuItemDesc:"Descripción",menuItemCalorie:"Calorías (kcal/100g)",menuItemPrice:"Precio",menuAutoLoad:"Cargar desde receta",menuPreviewBtn:"👁 Vista previa",menuEditBtn:"Editar",menuPrintBtn:"🖨",menuPageAuto:"Auto",menuPageSingle:"Una página",menuPageMulti:"Varias páginas",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Compartir",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"La clave API es necesaria para Claude OCR y análisis. Obténgala en https://console.anthropic.com",settingsProfileNamePh:"Introduce tu nombre",settingsProfileWorkplacePh:"Ej. Grand Hotel",settingsProfileDeptPh:"Ej. Repostería",settingsProfileRolePh:"Ej. Jefe de Repostería",settingsBackupDesc:"Descarga todos los datos como JSON o restaura.",settingsRestoreConfirm:"Los datos existentes serán sobrescritos. ¿Confirmar?",settingsRestoreSuccess:"Restauración completa",settingsDevDesc:"Para depuración",settingsExpenseName:"Nombre del gasto",settingsExpenseAmount:"Importe mensual (₺)",settingsPersonnelName:"Nombre de la persona",settingsPersonnelSalary:"Salario mensual (₺)",settingsStorageName:"Nombre de almacenamiento",settingsStorageType:"Tipo",settingsStorageTemp:"Temperatura objetivo (°C)",settingsAddIngredient:"Nombre de ingrediente a rastrear",settingsAddOrg:"Nombre de organización",
  stockEmpty:"Vaciando stock...",stockAdded:"Añadido al stock",stockDeducted:"Deducido del stock",insufficientStock:"Stock insuficiente: {name}",
  emptyStateRecipes:"Crea tu primera receta. Usa 📷 IA para detectar desde foto o añade manualmente.",emptyStateStock:"Usa 📄 Factura (IA) o + Manual para añadir items al stock.",emptyStateMenus:"Crea tu primer menú. Auto-carga desde recetas o añade manualmente.",
  profileRequired:"Falta perfil de usuario. Ve a Ajustes.",goToSettings:"Ir a Ajustes"
};
I18N.de={
  tabRecipes:"Rezepte",tabStock:"Bestand",tabProduction:"Produktion",tabReports:"Berichte",tabMenus:"Menüs",tabAssistant:"Assistent",tabSettings:"Einstellungen",
  save:"Speichern",cancel:"Abbrechen",delete:"Löschen",edit:"Bearbeiten",close:"Schließen",add:"Hinzufügen",back:"Zurück",next:"Weiter",search:"Suchen",filter:"Filter",all:"Alle",yes:"Ja",no:"Nein",confirm:"Bestätigen",remove:"Entfernen",update:"Aktualisieren",name:"Name",amount:"Menge",unit:"Einheit",price:"Preis",date:"Datum",time:"Zeit",note:"Notiz",optional:"optional",required:"erforderlich",loading:"Laden...",error:"Fehler",success:"Erfolg",warning:"Warnung",
  recipes:"Rezepte",newRecipe:"+ Neues Rezept",recipePhoto:"📷 Rezept (KI)",recipeManual:"+ Manuell",recipeDetail:"Rezeptdetail",ingredients:"Zutaten",allergens:"Allergene",cookMethod:"Kochen",cookTemp:"Temperatur",cookTime:"Zeit",servings:"Portionen",totalWeight:"Gesamtgewicht",portionG:"Gramm pro Portion",calories:"Kalorien",yieldL:"Ausbeute",recipeName:"Rezeptname",deductStock:"📦 Bestand abziehen & produzieren",sampleLabel:"🧪 Probezeuge-Etikett",editPhoto:"📷 Foto ändern",addPhoto:"📷 Foto hinzufügen",costAnalysis:"KOSTENANALYSE",material:"Material",overhead:"Fixkosten+Personal",realCost:"Reale Kosten",perPortion:"Portion",costHint:"💡 Fügen Sie Fixkosten und Personalkosten in Einstellungen hinzu, um reale Kosten zu sehen.",
  stock:"Bestand",stockMaterial:"📦 Material",stockLots:"🏷 Chargennummern",invoice:"📄 Rechnung (KI)",manualAdd:"+ Manuell",lowStock:"⚠ Niedriger Bestand",products:"Artikel",currentLots:"Aktuelle Chargennummern",noTrackedIngs:"Keine verfolgten Zutaten",noTrackedIngsHint:"Einstellungen → Rückverfolgbarkeit → Zutaten hinzufügen",lotMissing:"Charge nicht eingegeben",lotToday:"heute",lotDaysAgo:"Tage her",lotEntered:"eingegeben",lotHint:"💡 Scannen Sie morgens die Chargennummern der Verpackungen. 🔴 fehlt · 🔵 aktuell · 🟡 alt (3+ Tage).",
  production:"Produktion",productionRecord:"🍱 Produktionsaufzeichnung",productionReports:"📊 Berichte",storage:"Lagerort",expiryDays:"Verfall (Tage)",daysLater:"Tage später",risk:"Risiko",riskHigh:"Hoch",riskMedium:"Mittel",riskLow:"Niedrig",riskHint:"Automatisch berechnet, bearbeitbar",noteHint:"z.B. Für Abendservice",produceSaveBtn:"✓ Produzieren & abziehen",portions:"Portionen",lot:"Charge",produced:"Produziert",expires:"Verfällt",consumed:"verbraucht",active:"aktiv",fire:"Abfall",consumeAll:"✓ Alles verbraucht",printLabel:"🏷 Etikett drucken",partialConsume:"TEILVERBRAUCH",decrease:"− Abziehen",markFire:"× Als Abfall markieren",gotoReport:"📊 Zum Bericht",noProduction:"Noch keine Produktion",noProductionHint:"Verwenden Sie \"Bestand abziehen\" aus den Rezeptdetails, um einen Eintrag zu erstellen",consumedFull:"{name} ({portions} Portionen) wird als verbraucht markiert. Bestätigen?",invalidAmount:"Ungültige Menge",howManyFire:"Wie viele Portionen sind Abfall? (total {total})",ingredientLotsTitle:"🏷 CHARGENNUMMERN DER ZUTATEN",lotMissingHint:"💡 Für fehlende Chargen gehen Sie zu Bestand → Chargennummern.",
  sampleTitle:"🧪 Probezeuge-Etikett",sampleDesc:"ISO 22000-konforme Probeaufzeichnung — 72 Stunden Lagerung empfohlen",sampleName:"Probenname",sampleTaker:"Entnommen von",sampleOrgLabel:"Organisation",sampleOrgSelect:"Auswählen...",samplePax:"Personen",sampleLocation:"Zusatzinfo / Ort",sampleLocationHint:"z.B. Hochzeit, VIP-Tisch",sampleDateTime:"Entnahmedatum und -zeit",sampleTemp:"Probentemperatur (°C)",sampleTempHint:"Moment des Anrichtens",sampleHours:"Lagerungsdauer (Stunden)",samplePrintSave:"🏷 Drucken & speichern",sampleOrgOrLoc:"Organisation oder Ort eingeben",sampleProfileWarn:"⚠ Name in Einstellungen → Profil eingeben",
  reports:"Berichte",reportsProduction:"📋 Produktionsberichte",reportsStorage:"🗄 Lagerkontrolle",noReports:"Noch keine Berichte",checkTotal:"Kontrollen",checkComplete:"✓ Vollständig",addCheck:"+ Kontrolle hinzufügen",noChecksYet:"Noch keine Kontrollen aufgezeichnet.",storageCheckHint:"💡 Temperaturkontrolle 3-mal täglich pro Lager (09:00, 15:00, 21:00).",
  notificationsTitle:"🔔 Benachrichtigungen",notifEnabled:"Bot-Benachrichtigungen aktiv",notifDesc:"Benachrichtigungen erscheinen im Chat, keine Pop-ups",notifStorage:"Lagerprüfungs-Erinnerung (09:00, 15:00, 21:00)",notifExpired:"Ablaufdatum-Warnung",notifLow:"Niedrigbestand-Warnung",testNotif:"Testnachricht senden",
  userProfile:"Benutzerprofil",fullName:"Vollständiger Name",workplace:"Arbeitsort",department:"Abteilung",role:"Position",profileHint:"Diese Daten werden in FR.05 und FR.06 verwendet",placeholderName:"Name eingeben",placeholderWorkplace:"Hotel / Restaurant",placeholderDept:"Abteilung",placeholderRole:"Position",tabTodo:"Aufgaben",
  productionsCap:"PRODUKTIONEN",samplesCap:"PROBEN",firesCap:"ABFALL",
  storageCheckTitle:"🗄 Lagerkontrolle",controlTime:"Kontrollzeit",controlParams:"KONTROLLPARAMETER",paramEtiket:"Etikettenkontrolle",paramAcikGida:"Offene Lebensmittel",paramYerdeDuran:"Lebensmittel am Boden",paramCigPismis:"Roh/Gekocht-Trennung",paramTemizlik:"Sauberkeit",paramAlerjen:"Allergen-Regal-Trennung",paramSkt:"Verfall/Haltbarkeit",ok:"✓ OK",notOk:"✗ Nicht OK",productTemp:"Produkttemperatur (°C)",gaugeValue:"Anzeigewert (°C)",explanation:"Erklärung (optional)",explanationHint:"Korrigierende Maßnahme etc.",atLeastOneTemp:"Mindestens einen Temperaturwert eingeben",
  menus:"Menükarten",newMenu:"+ Neues Menü",chooseMenuType:"Menütyp wählen",menuEdit:"Menü bearbeiten",preview:"Vorschau",menuName:"Menüname",font:"Schrift",theme:"Thema",customSizes:"📐 BENUTZERDEFINIERTE GRÖSSEN (mm)",or:"oder",showPrice:"Preis anzeigen",showCalorie:"Kalorien anzeigen",showAllergen:"Allergene anzeigen",bgImage:"🖼 Hintergrundbild (optional)",chooseBg:"📷 Wählen",changeBg:"📷 Ändern",cropBg:"✂ Zuschneiden",bgOpacity:"Deckkraft",bgHint:"💡 Fügen Sie Marmor-, Papier-, Holzstrukturen als Hintergrund hinzu. Der Zuschneide-Editor öffnet sich nach dem Upload.",menuTemplates:"📑 MENÜVORLAGE",loadTemplate:"Vorlage laden...",saveTemplate:"💾 Vorlage speichern",templateHint:"💡 Speichern Sie Schrift, Thema, Größe, Hintergrund als Vorlage.",templateName:"Vorlagenname:",defaultTemplateName:"{menuName} Vorlage",templateSaveConfirm:"Vorlage \"{name}\" laden? (Abschnitte und Artikel bleiben erhalten, Design ändert sich)",templateSaved:"\"{name}\" als Vorlage gespeichert.",templateDeletePrompt:"Zu löschender Vorlagenname?",templateNotFound:"Nicht gefunden",
  cropTitle:"✂ Foto zuschneiden",cropDesc:"Ziehen zum Neupositionieren, Schieberegler für Zoom. Der Bereich innerhalb des Rahmens wird zum Menühintergrund.",cropZoom:"🔍 Zoom",cropApply:"✓ Zuschneiden & anwenden",
  settings:"Einstellungen",apiKey:"API-Schlüssel",language:"Sprache",darkMode:"Dunkles Thema",userProfile:"👤 Benutzerprofil",fullName:"Vollständiger Name",workplace:"Arbeitsplatz",department:"Abteilung",role:"Rolle",profileHint:"💡 Verwendet in Probezeugen-Etiketten und zukünftigen Produktionsaufzeichnungen.",traceabilityTitle:"🔍 Rückverfolgbarkeit (ISO 22000)",traceabilityMode:"Rückverfolgbarkeitsmodus",traceabilityDesc:"Aktiviert Chargenverfolgung, ISO 22000-Berichte (FR.06, FR.12) und PDF-Ausgaben",resetHourLabel:"Tagesende Chargen-Reset-Zeit",resetHourDesc:"Chargennummern werden täglich um diese Zeit zurückgesetzt",trackedIngsLabel:"Verfolgte Zutaten",trackedIngsPh:"Zutat hinzufügen (z.B. Mehl, Milch, Eier)",trackedIngsHint:"💡 Aktualisieren Sie diese Chargennummern jeden Morgen. Automatisch in FR.06 verwendet.",organizationsTitle:"🏢 Organisationen (für Proben)",organizationsPh:"Organisation hinzufügen (z.B. Hochzeit 200 Personen, Gala)",organizationsHint:"💡 Diese Liste wird als Dropdown im Probezeugen-Etikett verwendet.",expensesTitle:"💼 Geschäftsausgaben",fixedExpenses:"FIXKOSTEN (monatlich)",personnel:"PERSONAL (Monatsgehalt)",addFixed:"+ Fixkosten hinzufügen",addPersonnel:"+ Personal hinzufügen",monthlyPortions:"GESCHÄTZTE MONATLICHE PRODUKTION (Portionen)",portionHint:"Wird verwendet, um den Gemeinkostenanteil in den Rezeptkosten zu berechnen.",fixedTotal:"Fixkostensumme:",personnelTotal:"Personalsumme:",monthlyTotal:"Monatliche Gesamtausgaben:",perPortionShare:"Ausgabenanteil pro Portion:",storageAreas:"🗄 Lagerbereiche",addStorage:"+ Lagerbereich hinzufügen",storageHint:"💡 Fügen Sie hier die Namen Ihrer Kühlschränke/Gefrierschränke/Lager hinzu.",backup:"💾 Sicherung",backupBtn:"⬇ Sichern",restoreBtn:"⬆ Wiederherstellen",developer:"🔧 Entwickler",copyLogs:"📋 Logs kopieren",
  fr06Title:"FORMULAR ZUR KONTROLLE DER LEBENSMITTELZUBEREITUNG",fr06ProductCtrl:"Produktzubereitungskontrolle",fr06ProductIng:"Produktname · Zutaten · Chargennr.",fr06PrepStart:"Vorb. Start",fr06PrepEnd:"Vorb. Ende",fr06ProductTemp:"Produkttemp. (°C)",fr06AmbientTemp:"Umgebungstemp. (°C)",fr06Controller:"Kontrolleur",fr06CookCtrl:"Koch- / Aufwärmtemperaturkontrolle",fr06ProductDef:"Produktdefinition",fr06PartyNo:"Chargennr.",fr06Cook:"Kochen (°C)",fr06Reheat:"Aufwärmen (°C)",fr06CorrectiveAction:"Korrekturmaßnahme",fr06FreezingCtrl:"Auftaukontrolle",fr06ThawStart:"Auftaubeginn",fr06ThawEnd:"Auftauende",fr06Temp:"Temperatur",fr06FastCoolCtrl:"Schnellkühl- / Gefrierkontrolle",fr06Entry:"Eingang (Zeit-°C)",fr06Exit:"Ausgang (Zeit-°C)",fr06Chiller:"Kühler",fr06IceWater:"Eiswasser/-18°C",fr06PlateHold:"Tellerhaltung",fr06ProductOrg:"Produkt / Org.",fr06PlateTime:"Plattierungszeit",fr06Service:"Service",fr06InnerTemp:"Kerntemp.",fr06Ctrl:"Kontrolle",fr06FosterHold:"Foster/Trocknungshaltung",fr06Notes:"Anwendungshinweise",fr06NoteRisk:"Risikoreiche Lebensmittelzubereitung:",fr06NoteRiskText:"Umgebungstemperatur darf +24°C nicht überschreiten. Für Hochrisiko (Fleisch-Geflügel-Fisch) max. 16°C. Zubereitungszeit: 45 Min bei +20°C, 10 Min bei +25°C.",fr06NoteCool:"Kühlen/Gefrieren:",fr06NoteCoolText:"Max. 10°C innerhalb von 90 Min. Nie länger als 4 Stunden, um 5°C zu erreichen. Schockgefrieren: -18°C innerhalb von 4 Stunden.",fr06NoteCook:"Kochen/Aufwärmen:",fr06NoteCookText:"Kerntemperatur 75°C, Oberfläche 65°C. Aufwärmen mindestens 82°C.",fr06NotePlate:"Tellerhaltung:",fr06NotePlateText:"Keine Haltung über 15 Min erlaubt.",fr06NoteThaw:"Auftauen:",fr06NoteThawText:"Auftauen in Eiswasser erlaubt. Gefrorene Produkte 1 Tag im Voraus herausnehmen.",fr06DocNo:"Dokument-Nr.",fr06DateLbl:"Datum",fr06DefaultDept:"Küche / Abteilung",
  fr12Title:"PROBEZEUGEN-FORMULAR",fr12SampleOrg:"Probenname / Organisation",fr12Pax:"Personen",fr12DateTime:"Entnahmedatum und -zeit",fr12Temp:"Probentemperatur",fr12Taker:"Entnommen von",fr12Signature:"Unterschrift",fr12Explanation:"Erklärung",fr12Note:"Probezeuge:",fr12NoteText:"Probezeugen müssen für alle Organisationen mit über 20 Personen (einschließlich Frühstück und Personalkantine) entnommen und bei max. 5°C für 72 Stunden gelagert werden. Mindestens 200 Gramm.",
  fr05Title:"TÄGLICHES LAGER- / SCHRANKKONTROLLFORMULAR",fr05Dept:"ABTEILUNG:",fr05Cabinet:"SCHRANK/LAGER:",fr05Target:"Ziel:",fr05CtrlParams:"KONTROLLPARAMETER",fr05TempCtrl:"TEMPERATURKONTROLLE (°C)",fr05CtrlTime:"Kontrollzeit",fr05NoteDry:"Trockenlagerung:",fr05NoteDryText:"Temperatur <25°C und relative Luftfeuchtigkeit ≤65%.",fr05NoteCold:"Kühl-/Gefrierlagerung:",fr05NoteColdText:"Kühlbereiche bei 5°C oder darunter, Gefrierlagerung bei -18°C oder darunter.",fr05NoteApp:"Anwendungshinweis:",fr05NoteAppText:"Schreiben Sie bei der Inspektion (+) für konforme Situationen, (-) für nicht-konforme.",
  labelProduction:"PRODUKTIONSETIKETT",labelFreezer:"GEFRIERFACH-ETIKETT",labelProductName:"Produktname",labelProductionDate:"Produktionsdatum",labelOpeningDate:"Öffnungsdatum",labelExpiryDate:"Verfallsdatum",labelLotNumber:"Chargennummer",labelFreezingDate:"Gefrierdatum",labelFreezingTime:"Gefrierzeit",labelThawingDate:"Auftaudatum",labelThawingTime:"Auftauzeit",sampleWarn:"⚠ PROBEZEUGE — NICHT VERZEHREN ⚠",sampleNameLbl:"Probenname:",sampleOrgLbl:"Organisation / Ort:",sampleTakerLbl:"Entnommen von:",sampleDateTimeLbl:"Entnahmedatum und -zeit:",sampleStorageLbl:"Lagerungsdauer:",sampleEndLbl:"Ende:",sampleTempLbl:"Probentemperatur:",
  select:"Auswählen",none:"Keine",today:"Heute",fresh:"frisch",ok:"ok",old:"alt",
  stockAddManual:"Manuell hinzufügen",stockEditTitle:"Bestand bearbeiten",stockCategory:"Kategorie",stockMinLevel:"Min. Bestand",stockCurrentLevel:"Aktueller Bestand",stockAddBtn:"+ Hinzufügen",stockDetail:"Bestandsdetails",stockAllCats:"Alle Kategorien",stockSearch:"Artikel suchen...",noStock:"Noch kein Bestand",invoiceTitle:"Rechnung KI",invoiceUpload:"📷 Rechnungsfoto hochladen",invoiceAnalyze:"KI analysieren",invoiceItems:"Artikel",invoiceTotalExtracted:"{count} Artikel erkannt",addToStock:"+ Zum Bestand",packageUnit:"Packung/Einheit",packageQty:"Packungsanzahl",unitQty:"Menge pro Einheit",invoiceSaveBtn:"✓ Alle zum Bestand",
  recipeAdd:"Rezept hinzufügen",recipeEdit:"Rezept bearbeiten",recipeSave:"✓ Speichern",recipeAddIngredient:"+ Zutat hinzufügen",recipeMainCat:"Hauptkategorie",recipeSubCat:"Unterkategorie",recipeSearch:"Rezepte suchen...",noRecipes:"Noch keine Rezepte",recipePhotoCapture:"📷 Foto aufnehmen / hochladen",recipeAnalyzing:"KI analysiert...",recipeCostDetail:"Detaillierte Kosten",recipeDelete:"× Löschen",recipeDeleteConfirm:"{name} wird gelöscht. Bestätigen?",
  menuCreateTitle:"Menükarten",menuTypeCocktail:"Cocktail",menuTypeBreakfast:"Frühstück",menuTypeLunch:"Mittagessen",menuTypeDinner:"Abendessen",menuTypeGala:"Gala / Hochzeit",menuSaved:"Menü gespeichert",menuDeleteConfirm:"\"{name}\" wird gelöscht. Bestätigen?",menuAddSection:"+ Abschnitt hinzufügen",menuAddItem:"+ Artikel",menuSectionTitle:"Abschnittstitel",menuItemTitle:"Artikelname",menuItemDesc:"Beschreibung",menuItemCalorie:"Kalorien (kcal/100g)",menuItemPrice:"Preis",menuAutoLoad:"Aus Rezept laden",menuPreviewBtn:"👁 Vorschau",menuEditBtn:"Bearbeiten",menuPrintBtn:"🖨",menuPageAuto:"Auto",menuPageSingle:"Einzelseite",menuPageMulti:"Mehrere Seiten",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Teilen",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"API-Schlüssel wird für Claude OCR und Analyse benötigt. Erhältlich unter https://console.anthropic.com",settingsProfileNamePh:"Namen eingeben",settingsProfileWorkplacePh:"z.B. Grand Hotel",settingsProfileDeptPh:"z.B. Konditorei",settingsProfileRolePh:"z.B. Konditormeister",settingsBackupDesc:"Alle Daten als JSON herunterladen oder wiederherstellen.",settingsRestoreConfirm:"Vorhandene Daten werden überschrieben. Bestätigen?",settingsRestoreSuccess:"Wiederherstellung abgeschlossen",settingsDevDesc:"Zum Debuggen",settingsExpenseName:"Ausgabenname",settingsExpenseAmount:"Monatlicher Betrag (₺)",settingsPersonnelName:"Personenname",settingsPersonnelSalary:"Monatsgehalt (₺)",settingsStorageName:"Lagerbezeichnung",settingsStorageType:"Typ",settingsStorageTemp:"Zieltemperatur (°C)",settingsAddIngredient:"Zu verfolgender Zutatenname",settingsAddOrg:"Organisationsname",
  stockEmpty:"Bestand wird geleert...",stockAdded:"Zum Bestand hinzugefügt",stockDeducted:"Vom Bestand abgezogen",insufficientStock:"Unzureichender Bestand: {name}",
  emptyStateRecipes:"Erstellen Sie Ihr erstes Rezept. Verwenden Sie 📷 KI oder fügen Sie manuell hinzu.",emptyStateStock:"Verwenden Sie 📄 Rechnung (KI) oder + Manuell, um Artikel hinzuzufügen.",emptyStateMenus:"Erstellen Sie Ihr erstes Menü. Auto-Laden aus Rezepten oder manuell.",
  profileRequired:"Benutzerprofil fehlt. Gehen Sie zu Einstellungen.",goToSettings:"Zu Einstellungen"
};
I18N.fr={
  tabRecipes:"Recettes",tabStock:"Stock",tabProduction:"Production",tabReports:"Rapports",tabMenus:"Menus",tabAssistant:"Assistant",tabSettings:"Paramètres",
  save:"Enregistrer",cancel:"Annuler",delete:"Supprimer",edit:"Modifier",close:"Fermer",add:"Ajouter",back:"Retour",next:"Suivant",search:"Rechercher",filter:"Filtrer",all:"Tous",yes:"Oui",no:"Non",confirm:"Confirmer",remove:"Retirer",update:"Mettre à jour",name:"Nom",amount:"Quantité",unit:"Unité",price:"Prix",date:"Date",time:"Heure",note:"Note",optional:"optionnel",required:"requis",loading:"Chargement...",error:"Erreur",success:"Succès",warning:"Avertissement",
  recipes:"Recettes",newRecipe:"+ Nouvelle recette",recipePhoto:"📷 Recette (IA)",recipeManual:"+ Manuel",recipeDetail:"Détail de recette",ingredients:"Ingrédients",allergens:"Allergènes",cookMethod:"Cuisson",cookTemp:"Température",cookTime:"Temps",servings:"Portions",totalWeight:"Poids total",portionG:"Grammes par portion",calories:"Calories",yieldL:"Rendement",recipeName:"Nom de recette",deductStock:"📦 Déduire stock & produire",sampleLabel:"🧪 Étiquette témoin",editPhoto:"📷 Changer photo",addPhoto:"📷 Ajouter photo",costAnalysis:"ANALYSE DES COÛTS",material:"Matériaux",overhead:"Fixes+Personnel",realCost:"Coût réel",perPortion:"portion",costHint:"💡 Ajoutez les frais fixes et le personnel dans Paramètres pour voir le coût réel.",
  stock:"Stock",stockMaterial:"📦 Matières",stockLots:"🏷 Numéros de lot",invoice:"📄 Facture (IA)",manualAdd:"+ Manuel",lowStock:"⚠ Stock bas",products:"articles",currentLots:"Numéros de lot actuels",noTrackedIngs:"Aucun ingrédient suivi",noTrackedIngsHint:"Paramètres → Traçabilité → Ajouter des ingrédients",lotMissing:"Lot non saisi",lotToday:"aujourd'hui",lotDaysAgo:"jours",lotEntered:"saisi",lotHint:"💡 Scannez les numéros de lot des emballages chaque matin. 🔴 manquant · 🔵 actuel · 🟡 ancien (3+ jours).",
  production:"Production",productionRecord:"🍱 Enregistrement de production",productionReports:"📊 Rapports",storage:"Lieu de stockage",expiryDays:"Péremption (jours)",daysLater:"jours plus tard",risk:"Risque",riskHigh:"Élevé",riskMedium:"Moyen",riskLow:"Bas",riskHint:"Auto-calculé, modifiable",noteHint:"Ex. Pour service du soir",produceSaveBtn:"✓ Produire & déduire",portions:"portions",lot:"Lot",produced:"Produit",expires:"Expire",consumed:"consommé",active:"actif",fire:"Déchet",consumeAll:"✓ Tout consommé",printLabel:"🏷 Imprimer étiquette",partialConsume:"CONSOMMATION PARTIELLE",decrease:"− Déduire",markFire:"× Marquer comme déchet",gotoReport:"📊 Aller au rapport",noProduction:"Pas encore de production",noProductionHint:"Utilisez \"Déduire stock\" depuis le détail de la recette pour créer un enregistrement",consumedFull:"{name} ({portions} portions) sera marqué comme consommé. Confirmer?",invalidAmount:"Quantité invalide",howManyFire:"Combien de portions en déchet? (total {total})",ingredientLotsTitle:"🏷 NUMÉROS DE LOT DES INGRÉDIENTS",lotMissingHint:"💡 Pour les lots manquants, allez à Stock → Numéros de lot.",
  sampleTitle:"🧪 Étiquette témoin",sampleDesc:"Enregistrement d'échantillon conforme ISO 22000 — 72 heures de conservation recommandées",sampleName:"Nom de l'échantillon",sampleTaker:"Prélevé par",sampleOrgLabel:"Organisation",sampleOrgSelect:"Sélectionner...",samplePax:"Personnes",sampleLocation:"Info supplémentaire / Lieu",sampleLocationHint:"Ex. Mariage, table VIP",sampleDateTime:"Date et heure de prélèvement",sampleTemp:"Température de l'échantillon (°C)",sampleTempHint:"Moment du dressage",sampleHours:"Durée de conservation (heures)",samplePrintSave:"🏷 Imprimer & enregistrer",sampleOrgOrLoc:"Entrez organisation ou lieu",sampleProfileWarn:"⚠ Entrez le nom dans Paramètres → Profil",
  reports:"Rapports",reportsProduction:"📋 Rapports de production",reportsStorage:"🗄 Contrôle de stockage",noReports:"Pas encore de rapports",checkTotal:"contrôles",checkComplete:"✓ Complet",addCheck:"+ Ajouter contrôle",noChecksYet:"Aucun contrôle enregistré.",storageCheckHint:"💡 Contrôle de température 3 fois par jour par stockage (09:00, 15:00, 21:00).",
  notificationsTitle:"🔔 Notifications",notifEnabled:"Notifications bot actives",notifDesc:"Les notifications apparaissent dans le chat, sans pop-ups",notifStorage:"Rappel de contrôle de stockage (09:00, 15:00, 21:00)",notifExpired:"Alerte produit périmé",notifLow:"Alerte stock bas",testNotif:"Envoyer message test",
  userProfile:"Profil utilisateur",fullName:"Nom complet",workplace:"Lieu de travail",department:"Département",role:"Poste",profileHint:"Ces données sont utilisées dans les formulaires FR.05 et FR.06",placeholderName:"Entrez votre nom",placeholderWorkplace:"Hôtel / Restaurant",placeholderDept:"Département",placeholderRole:"Poste",tabTodo:"Tâches",
  productionsCap:"PRODUCTIONS",samplesCap:"ÉCHANTILLONS",firesCap:"DÉCHETS",
  storageCheckTitle:"🗄 Contrôle de stockage",controlTime:"Heure de contrôle",controlParams:"PARAMÈTRES DE CONTRÔLE",paramEtiket:"Contrôle d'étiquette",paramAcikGida:"Aliment ouvert",paramYerdeDuran:"Aliment au sol",paramCigPismis:"Séparation cru/cuit",paramTemizlik:"Propreté",paramAlerjen:"Séparation allergènes",paramSkt:"Péremption/DLC",ok:"✓ OK",notOk:"✗ Non OK",productTemp:"Température du produit (°C)",gaugeValue:"Valeur de la jauge (°C)",explanation:"Explication (optionnelle)",explanationHint:"Action corrective etc.",atLeastOneTemp:"Entrez au moins une valeur de température",
  menus:"Cartes de menu",newMenu:"+ Nouveau menu",chooseMenuType:"Choisir le type de menu",menuEdit:"Modifier le menu",preview:"Aperçu",menuName:"Nom du menu",font:"Police",theme:"Thème",customSizes:"📐 TAILLES PERSONNALISÉES (mm)",or:"ou",showPrice:"Afficher prix",showCalorie:"Afficher calories",showAllergen:"Afficher allergènes",bgImage:"🖼 Image de fond (optionnelle)",chooseBg:"📷 Choisir",changeBg:"📷 Changer",cropBg:"✂ Recadrer",bgOpacity:"Opacité",bgHint:"💡 Ajoutez des fonds de marbre, papier, texture bois. L'éditeur s'ouvre après téléversement.",menuTemplates:"📑 MODÈLE DE MENU",loadTemplate:"Charger modèle...",saveTemplate:"💾 Enregistrer modèle",templateHint:"💡 Enregistrez police, thème, taille, fond comme modèle.",templateName:"Nom du modèle:",defaultTemplateName:"Modèle {menuName}",templateSaveConfirm:"Charger le modèle \"{name}\"? (Sections et articles préservés, design change)",templateSaved:"\"{name}\" enregistré comme modèle.",templateDeletePrompt:"Nom du modèle à supprimer?",templateNotFound:"Non trouvé",
  cropTitle:"✂ Recadrer photo",cropDesc:"Glissez pour repositionner, curseur pour zoomer. La zone à l'intérieur du cadre devient le fond du menu.",cropZoom:"🔍 Zoom",cropApply:"✓ Recadrer & appliquer",
  settings:"Paramètres",apiKey:"Clé API",language:"Langue",darkMode:"Thème sombre",userProfile:"👤 Profil utilisateur",fullName:"Nom complet",workplace:"Lieu de travail",department:"Département",role:"Rôle",profileHint:"💡 Utilisé dans les étiquettes témoins et futurs enregistrements de production.",traceabilityTitle:"🔍 Traçabilité (ISO 22000)",traceabilityMode:"Mode traçabilité",traceabilityDesc:"Active le suivi des lots, rapports ISO 22000 (FR.06, FR.12) et sorties PDF",resetHourLabel:"Heure de réinitialisation des lots en fin de journée",resetHourDesc:"les numéros de lot se réinitialisent quotidiennement à cette heure",trackedIngsLabel:"Ingrédients suivis",trackedIngsPh:"Ajouter ingrédient (ex. Farine, Lait, Œufs)",trackedIngsHint:"💡 Mettez à jour ces numéros de lot chaque matin. Auto-utilisé dans FR.06.",organizationsTitle:"🏢 Organisations (pour échantillons)",organizationsPh:"Ajouter organisation (ex. Mariage 200 pers., Gala)",organizationsHint:"💡 Cette liste est utilisée comme liste déroulante dans l'étiquette témoin.",expensesTitle:"💼 Dépenses d'entreprise",fixedExpenses:"DÉPENSES FIXES (mensuelles)",personnel:"PERSONNEL (salaire mensuel)",addFixed:"+ Ajouter dépense fixe",addPersonnel:"+ Ajouter personnel",monthlyPortions:"PRODUCTION MENSUELLE ESTIMÉE (portions)",portionHint:"Utilisé pour calculer la part des frais généraux dans les coûts de recette.",fixedTotal:"Total dépenses fixes:",personnelTotal:"Total personnel:",monthlyTotal:"Total dépenses mensuelles:",perPortionShare:"Part des dépenses par portion:",storageAreas:"🗄 Zones de stockage",addStorage:"+ Ajouter zone de stockage",storageHint:"💡 Ajoutez ici les noms de vos réfrigérateurs/congélateurs/stocks.",backup:"💾 Sauvegarde",backupBtn:"⬇ Sauvegarder",restoreBtn:"⬆ Restaurer",developer:"🔧 Développeur",copyLogs:"📋 Copier les logs",
  fr06Title:"FORMULAIRE DE CONTRÔLE DE PRÉPARATION DES ALIMENTS",fr06ProductCtrl:"Contrôle de préparation du produit",fr06ProductIng:"Nom du produit · Ingrédients · N° de lot",fr06PrepStart:"Début prép.",fr06PrepEnd:"Fin prép.",fr06ProductTemp:"Temp. produit (°C)",fr06AmbientTemp:"Temp. ambiante (°C)",fr06Controller:"Contrôleur",fr06CookCtrl:"Contrôle de température de cuisson / réchauffage",fr06ProductDef:"Définition du produit",fr06PartyNo:"N° de lot",fr06Cook:"Cuisson (°C)",fr06Reheat:"Réchauffage (°C)",fr06CorrectiveAction:"Action corrective",fr06FreezingCtrl:"Contrôle de décongélation",fr06ThawStart:"Début décongélation",fr06ThawEnd:"Fin décongélation",fr06Temp:"Température",fr06FastCoolCtrl:"Contrôle de refroidissement rapide / congélation",fr06Entry:"Entrée (Heure-°C)",fr06Exit:"Sortie (Heure-°C)",fr06Chiller:"Refroidisseur",fr06IceWater:"Eau glacée/-18°C",fr06PlateHold:"Maintien en assiette",fr06ProductOrg:"Produit / Org.",fr06PlateTime:"Heure de dressage",fr06Service:"Service",fr06InnerTemp:"Temp. cœur",fr06Ctrl:"Contrôle",fr06FosterHold:"Maintien Foster/séchage",fr06Notes:"Notes d'application",fr06NoteRisk:"Préparation d'aliments à risque:",fr06NoteRiskText:"La température ambiante ne doit pas dépasser +24°C. Pour aliments à haut risque (viande-volaille-poisson) max 16°C. Temps de préparation: 45 min à +20°C, 10 min à +25°C.",fr06NoteCool:"Refroidissement/Congélation:",fr06NoteCoolText:"Max 10°C en 90 min. Ne jamais dépasser 4 heures pour atteindre 5°C. Congélation rapide: -18°C en 4 heures.",fr06NoteCook:"Cuisson/Réchauffage:",fr06NoteCookText:"Temp. cœur 75°C, surface 65°C. Réchauffage min 82°C.",fr06NotePlate:"Maintien en assiette:",fr06NotePlateText:"Aucun maintien supérieur à 15 min autorisé.",fr06NoteThaw:"Décongélation:",fr06NoteThawText:"Décongélation à l'eau glacée autorisée. Les produits congelés sortis 1 jour à l'avance.",fr06DocNo:"N° de document",fr06DateLbl:"Date",fr06DefaultDept:"Cuisine / Section",
  fr12Title:"FORMULAIRE D'ÉCHANTILLON TÉMOIN",fr12SampleOrg:"Nom d'échantillon / Organisation",fr12Pax:"Personnes",fr12DateTime:"Date et heure de prélèvement",fr12Temp:"Température d'échantillon",fr12Taker:"Prélevé par",fr12Signature:"Signature",fr12Explanation:"Explication",fr12Note:"Échantillon témoin:",fr12NoteText:"Des échantillons témoins doivent être prélevés pour toutes les organisations servant plus de 20 personnes (y compris petit-déjeuner et cafétéria du personnel) et stockés à max 5°C pendant 72 heures. Minimum 200 grammes.",
  fr05Title:"FORMULAIRE DE CONTRÔLE QUOTIDIEN DE STOCKAGE / ARMOIRE",fr05Dept:"DÉPARTEMENT:",fr05Cabinet:"ARMOIRE/STOCK:",fr05Target:"Cible:",fr05CtrlParams:"PARAMÈTRES DE CONTRÔLE",fr05TempCtrl:"CONTRÔLE DE TEMPÉRATURE (°C)",fr05CtrlTime:"Heure de contrôle",fr05NoteDry:"Stockage sec:",fr05NoteDryText:"Température <25°C et humidité relative ≤65%.",fr05NoteCold:"Stockage froid/congelé:",fr05NoteColdText:"Zones froides à 5°C ou moins, stockage congelé à -18°C ou moins.",fr05NoteApp:"Note d'application:",fr05NoteAppText:"Pendant l'inspection, écrire (+) pour les situations conformes, (-) pour les non-conformes.",
  labelProduction:"ÉTIQUETTE DE PRODUCTION",labelFreezer:"ÉTIQUETTE DE CONGÉLATEUR",labelProductName:"Nom du produit",labelProductionDate:"Date de production",labelOpeningDate:"Date d'ouverture",labelExpiryDate:"Date de péremption",labelLotNumber:"N° de lot",labelFreezingDate:"Date de congélation",labelFreezingTime:"Heure de congélation",labelThawingDate:"Date de décongélation",labelThawingTime:"Heure de décongélation",sampleWarn:"⚠ ÉCHANTILLON TÉMOIN — NE PAS CONSOMMER ⚠",sampleNameLbl:"Nom d'échantillon:",sampleOrgLbl:"Organisation / Lieu:",sampleTakerLbl:"Prélevé par:",sampleDateTimeLbl:"Date et heure de prélèvement:",sampleStorageLbl:"Durée de conservation:",sampleEndLbl:"fin:",sampleTempLbl:"Température d'échantillon:",
  select:"Sélectionner",none:"Aucun",today:"Aujourd'hui",fresh:"frais",ok:"ok",old:"ancien",
  stockAddManual:"Ajouter manuellement",stockEditTitle:"Modifier stock",stockCategory:"Catégorie",stockMinLevel:"Stock minimum",stockCurrentLevel:"Stock actuel",stockAddBtn:"+ Ajouter",stockDetail:"Détail du stock",stockAllCats:"Toutes les catégories",stockSearch:"Rechercher articles...",noStock:"Pas encore de stock",invoiceTitle:"Facture IA",invoiceUpload:"📷 Téléverser photo de facture",invoiceAnalyze:"Analyser avec IA",invoiceItems:"Articles",invoiceTotalExtracted:"{count} articles détectés",addToStock:"+ Au stock",packageUnit:"Emballage/Unité",packageQty:"Nombre d'emballages",unitQty:"Quantité par unité",invoiceSaveBtn:"✓ Tout ajouter au stock",
  recipeAdd:"Ajouter recette",recipeEdit:"Modifier recette",recipeSave:"✓ Enregistrer",recipeAddIngredient:"+ Ajouter ingrédient",recipeMainCat:"Catégorie principale",recipeSubCat:"Sous-catégorie",recipeSearch:"Rechercher recettes...",noRecipes:"Pas encore de recettes",recipePhotoCapture:"📷 Capturer / Téléverser photo",recipeAnalyzing:"IA en analyse...",recipeCostDetail:"Coût détaillé",recipeDelete:"× Supprimer",recipeDeleteConfirm:"{name} sera supprimé. Confirmer?",
  menuCreateTitle:"Cartes de menu",menuTypeCocktail:"Cocktail",menuTypeBreakfast:"Petit-déjeuner",menuTypeLunch:"Déjeuner",menuTypeDinner:"Dîner",menuTypeGala:"Gala / Mariage",menuSaved:"Menu enregistré",menuDeleteConfirm:"\"{name}\" sera supprimé. Confirmer?",menuAddSection:"+ Ajouter section",menuAddItem:"+ Article",menuSectionTitle:"Titre de section",menuItemTitle:"Nom de l'article",menuItemDesc:"Description",menuItemCalorie:"Calories (kcal/100g)",menuItemPrice:"Prix",menuAutoLoad:"Charger depuis recette",menuPreviewBtn:"👁 Aperçu",menuEditBtn:"Modifier",menuPrintBtn:"🖨",menuPageAuto:"Auto",menuPageSingle:"Une page",menuPageMulti:"Plusieurs pages",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 Partager",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"La clé API est requise pour Claude OCR et analyse. Obtenez-la sur https://console.anthropic.com",settingsProfileNamePh:"Entrez votre nom",settingsProfileWorkplacePh:"Ex. Grand Hôtel",settingsProfileDeptPh:"Ex. Pâtisserie",settingsProfileRolePh:"Ex. Chef Pâtissier",settingsBackupDesc:"Télécharger toutes les données en JSON ou restaurer.",settingsRestoreConfirm:"Les données existantes seront écrasées. Confirmer?",settingsRestoreSuccess:"Restauration terminée",settingsDevDesc:"Pour le débogage",settingsExpenseName:"Nom de la dépense",settingsExpenseAmount:"Montant mensuel (₺)",settingsPersonnelName:"Nom de la personne",settingsPersonnelSalary:"Salaire mensuel (₺)",settingsStorageName:"Nom du stockage",settingsStorageType:"Type",settingsStorageTemp:"Température cible (°C)",settingsAddIngredient:"Nom d'ingrédient à suivre",settingsAddOrg:"Nom de l'organisation",
  stockEmpty:"Vidage du stock...",stockAdded:"Ajouté au stock",stockDeducted:"Déduit du stock",insufficientStock:"Stock insuffisant: {name}",
  emptyStateRecipes:"Créez votre première recette. Utilisez 📷 IA ou ajoutez manuellement.",emptyStateStock:"Utilisez 📄 Facture (IA) ou + Manuel pour ajouter des articles au stock.",emptyStateMenus:"Créez votre premier menu. Auto-chargement depuis recettes ou manuellement.",
  profileRequired:"Profil utilisateur manquant. Allez dans Paramètres.",goToSettings:"Aux Paramètres"
};
I18N.zh={
  tabRecipes:"食谱",tabStock:"库存",tabProduction:"生产",tabReports:"报告",tabMenus:"菜单",tabAssistant:"助手",tabSettings:"设置",
  save:"保存",cancel:"取消",delete:"删除",edit:"编辑",close:"关闭",add:"添加",back:"返回",next:"下一个",search:"搜索",filter:"筛选",all:"全部",yes:"是",no:"否",confirm:"确认",remove:"移除",update:"更新",name:"名称",amount:"数量",unit:"单位",price:"价格",date:"日期",time:"时间",note:"备注",optional:"可选",required:"必需",loading:"加载中...",error:"错误",success:"成功",warning:"警告",
  recipes:"食谱",newRecipe:"+ 新食谱",recipePhoto:"📷 食谱 (AI)",recipeManual:"+ 手动",recipeDetail:"食谱详情",ingredients:"配料",allergens:"过敏原",cookMethod:"烹饪",cookTemp:"温度",cookTime:"时间",servings:"份数",totalWeight:"总重量",portionG:"每份克数",calories:"卡路里",yieldL:"产量",recipeName:"食谱名称",deductStock:"📦 扣除库存并生产",sampleLabel:"🧪 留样标签",editPhoto:"📷 更改照片",addPhoto:"📷 添加照片",costAnalysis:"成本分析",material:"原料",overhead:"固定+人员",realCost:"实际成本",perPortion:"份",costHint:"💡 在设置中添加固定开支和人员成本以查看实际成本。",
  stock:"库存",stockMaterial:"📦 原料",stockLots:"🏷 批次号",invoice:"📄 发票 (AI)",manualAdd:"+ 手动",lowStock:"⚠ 库存低",products:"项目",currentLots:"当前批次号",noTrackedIngs:"无跟踪的配料",noTrackedIngsHint:"设置 → 可追溯性 → 添加跟踪配料",lotMissing:"未输入批次",lotToday:"今天",lotDaysAgo:"天前",lotEntered:"已输入",lotHint:"💡 每天早晨扫描包装上的批次号。🔴 缺失 · 🔵 当前 · 🟡 旧 (3+ 天)。",
  production:"生产",productionRecord:"🍱 生产记录",productionReports:"📊 报告",storage:"存储位置",expiryDays:"保质期(天)",daysLater:"天后",risk:"风险",riskHigh:"高",riskMedium:"中",riskLow:"低",riskHint:"自动计算，可编辑",noteHint:"例如：用于晚间服务",produceSaveBtn:"✓ 生产并扣除",portions:"份数",lot:"批次",produced:"已生产",expires:"过期",consumed:"已消费",active:"活跃",fire:"废弃",consumeAll:"✓ 全部消费",printLabel:"🏷 打印标签",partialConsume:"部分消费",decrease:"− 扣除",markFire:"× 标记为废弃",gotoReport:"📊 转到报告",noProduction:"暂无生产",noProductionHint:"使用食谱详情中的\"扣除库存\"创建生产记录",consumedFull:"{name} ({portions} 份) 将被标记为已消费。确认？",invalidAmount:"无效数量",howManyFire:"多少份废弃？（共 {total}）",ingredientLotsTitle:"🏷 配料批次号",lotMissingHint:"💡 对于缺失的批次，请转到库存 → 批次号。",
  sampleTitle:"🧪 留样标签",sampleDesc:"符合 ISO 22000 的样品记录 — 建议储存 72 小时",sampleName:"样品名称",sampleTaker:"取样人",sampleOrgLabel:"组织",sampleOrgSelect:"选择...",samplePax:"人数",sampleLocation:"附加信息 / 位置",sampleLocationHint:"例如：婚礼，VIP 桌",sampleDateTime:"取样日期和时间",sampleTemp:"样品温度 (°C)",sampleTempHint:"摆盘时刻",sampleHours:"储存时间（小时）",samplePrintSave:"🏷 打印并保存",sampleOrgOrLoc:"输入组织或位置",sampleProfileWarn:"⚠ 在设置 → 用户资料中输入姓名",
  reports:"报告",reportsProduction:"📋 生产报告",reportsStorage:"🗄 存储控制",noReports:"暂无报告",checkTotal:"检查",checkComplete:"✓ 完成",addCheck:"+ 添加检查",noChecksYet:"尚未记录检查。",storageCheckHint:"💡 每天 3 次温度控制（09:00、15:00、21:00）。",
  notificationsTitle:"🔔 通知",notifEnabled:"机器人通知已激活",notifDesc:"通知出现在聊天中，无弹窗",notifStorage:"存储检查提醒（09:00、15:00、21:00）",notifExpired:"过期产品警告",notifLow:"低库存警告",testNotif:"发送测试消息",
  userProfile:"用户资料",fullName:"全名",workplace:"工作地点",department:"部门",role:"职位",profileHint:"这些数据用于FR.05和FR.06表格",placeholderName:"输入您的姓名",placeholderWorkplace:"酒店/餐厅",placeholderDept:"部门",placeholderRole:"职位",tabTodo:"任务",
  productionsCap:"生产",samplesCap:"样品",firesCap:"废弃",
  storageCheckTitle:"🗄 存储检查",controlTime:"检查时间",controlParams:"检查参数",paramEtiket:"标签检查",paramAcikGida:"开放食品",paramYerdeDuran:"地面食品",paramCigPismis:"生/熟分离",paramTemizlik:"清洁度",paramAlerjen:"过敏原架分离",paramSkt:"保质期",ok:"✓ 合格",notOk:"✗ 不合格",productTemp:"产品温度 (°C)",gaugeValue:"仪表值 (°C)",explanation:"说明（可选）",explanationHint:"纠正措施等",atLeastOneTemp:"至少输入一个温度值",
  menus:"菜单卡",newMenu:"+ 新菜单",chooseMenuType:"选择菜单类型",menuEdit:"编辑菜单",preview:"预览",menuName:"菜单名称",font:"字体",theme:"主题",customSizes:"📐 自定义尺寸（毫米）",or:"或",showPrice:"显示价格",showCalorie:"显示卡路里",showAllergen:"显示过敏原",bgImage:"🖼 背景图片（可选）",chooseBg:"📷 选择",changeBg:"📷 更改",cropBg:"✂ 裁剪",bgOpacity:"不透明度",bgHint:"💡 添加大理石、纸张、木质纹理背景。上传后打开裁剪编辑器。",menuTemplates:"📑 菜单模板",loadTemplate:"加载模板...",saveTemplate:"💾 保存模板",templateHint:"💡 将字体、主题、尺寸、背景保存为模板。",templateName:"模板名称：",defaultTemplateName:"{menuName} 模板",templateSaveConfirm:"加载模板\"{name}\"？（保留区域和项目，设计更改）",templateSaved:"\"{name}\" 已保存为模板。",templateDeletePrompt:"要删除的模板名称？",templateNotFound:"未找到",
  cropTitle:"✂ 裁剪照片",cropDesc:"拖动重新定位，滑块缩放。框内区域成为菜单背景。",cropZoom:"🔍 缩放",cropApply:"✓ 裁剪并应用",
  settings:"设置",apiKey:"API 密钥",language:"语言",darkMode:"深色主题",userProfile:"👤 用户资料",fullName:"全名",workplace:"工作单位",department:"部门",role:"角色",profileHint:"💡 用于留样标签和未来的生产记录。",traceabilityTitle:"🔍 可追溯性 (ISO 22000)",traceabilityMode:"可追溯性模式",traceabilityDesc:"启用批次跟踪、ISO 22000 报告 (FR.06, FR.12) 和 PDF 输出",resetHourLabel:"日终批次重置时间",resetHourDesc:"批次号每天在此时间重置",trackedIngsLabel:"跟踪的配料",trackedIngsPh:"添加配料（例如：面粉、牛奶、鸡蛋）",trackedIngsHint:"💡 每天早晨更新这些配料批次号。自动用于 FR.06。",organizationsTitle:"🏢 组织（用于样品）",organizationsPh:"添加组织（例如：婚礼 200 人、晚宴）",organizationsHint:"💡 此列表用作留样标签中的下拉菜单。",expensesTitle:"💼 业务费用",fixedExpenses:"固定费用（每月）",personnel:"人员（月薪）",addFixed:"+ 添加固定费用",addPersonnel:"+ 添加人员",monthlyPortions:"每月估计生产（份数）",portionHint:"用于计算食谱成本中的间接费用份额。",fixedTotal:"固定费用总计：",personnelTotal:"人员总计：",monthlyTotal:"每月总费用：",perPortionShare:"每份费用份额：",storageAreas:"🗄 存储区域",addStorage:"+ 添加存储区域",storageHint:"💡 在此添加您厨房的冰箱/冷冻柜/储存名称。",backup:"💾 备份",backupBtn:"⬇ 备份",restoreBtn:"⬆ 还原",developer:"🔧 开发者",copyLogs:"📋 复制日志",
  fr06Title:"食品制备控制表",fr06ProductCtrl:"产品制备控制",fr06ProductIng:"产品名称 · 配料 · 批次号",fr06PrepStart:"开始",fr06PrepEnd:"结束",fr06ProductTemp:"产品温度 (°C)",fr06AmbientTemp:"环境温度 (°C)",fr06Controller:"控制员",fr06CookCtrl:"烹饪/再加热温度控制",fr06ProductDef:"产品定义",fr06PartyNo:"批次号",fr06Cook:"烹饪 (°C)",fr06Reheat:"再加热 (°C)",fr06CorrectiveAction:"纠正措施",fr06FreezingCtrl:"解冻控制",fr06ThawStart:"开始解冻",fr06ThawEnd:"结束解冻",fr06Temp:"温度",fr06FastCoolCtrl:"快速冷却/冷冻控制",fr06Entry:"进入（时间-°C）",fr06Exit:"退出（时间-°C）",fr06Chiller:"冷却器",fr06IceWater:"冰水/-18°C",fr06PlateHold:"盘中保温",fr06ProductOrg:"产品/组织",fr06PlateTime:"摆盘时间",fr06Service:"服务",fr06InnerTemp:"内温",fr06Ctrl:"控制",fr06FosterHold:"Foster/干燥保温",fr06Notes:"应用说明",fr06NoteRisk:"高风险食品制备：",fr06NoteRiskText:"环境温度不得超过 +24°C。高风险（肉禽鱼）最高 16°C。制备时间：+20°C 时 45 分钟，+25°C 时 10 分钟。",fr06NoteCool:"冷却/冷冻：",fr06NoteCoolText:"90 分钟内最高 10°C。绝不超过 4 小时达到 5°C。急速冷冻：4 小时内 -18°C。",fr06NoteCook:"烹饪/再加热：",fr06NoteCookText:"内温 75°C，表面 65°C。再加热至少 82°C。",fr06NotePlate:"盘中保温：",fr06NotePlateText:"不允许超过 15 分钟的保温。",fr06NoteThaw:"解冻：",fr06NoteThawText:"允许冰水解冻。冷冻产品提前 1 天取出。",fr06DocNo:"文档编号",fr06DateLbl:"日期",fr06DefaultDept:"厨房/部门",
  fr12Title:"留样见证表",fr12SampleOrg:"样品名称/组织",fr12Pax:"人数",fr12DateTime:"采集日期和时间",fr12Temp:"样品温度",fr12Taker:"取样人",fr12Signature:"签名",fr12Explanation:"说明",fr12Note:"留样见证：",fr12NoteText:"对于所有服务超过 20 人的组织（包括早餐和员工食堂），必须取留样见证并在 5°C 以下储存 72 小时。最少 200 克。",
  fr05Title:"每日存储/橱柜控制表",fr05Dept:"部门：",fr05Cabinet:"橱柜/存储：",fr05Target:"目标：",fr05CtrlParams:"控制参数",fr05TempCtrl:"温度控制 (°C)",fr05CtrlTime:"检查时间",fr05NoteDry:"干燥存储：",fr05NoteDryText:"温度 <25°C，相对湿度 ≤65%。",fr05NoteCold:"冷藏/冷冻存储：",fr05NoteColdText:"冷藏区 5°C 或以下，冷冻存储 -18°C 或以下。",fr05NoteApp:"应用说明：",fr05NoteAppText:"检查期间，对合规情况写 (+)，不合规情况写 (-)。",
  labelProduction:"生产标签",labelFreezer:"冷冻柜标签",labelProductName:"产品名称",labelProductionDate:"生产日期",labelOpeningDate:"开启日期",labelExpiryDate:"保质期",labelLotNumber:"批次号",labelFreezingDate:"冷冻日期",labelFreezingTime:"冷冻时间",labelThawingDate:"解冻日期",labelThawingTime:"解冻时间",sampleWarn:"⚠ 留样见证 — 请勿食用 ⚠",sampleNameLbl:"样品名称：",sampleOrgLbl:"组织/位置：",sampleTakerLbl:"取样人：",sampleDateTimeLbl:"采集日期和时间：",sampleStorageLbl:"储存时间：",sampleEndLbl:"结束：",sampleTempLbl:"样品温度：",
  select:"选择",none:"无",today:"今天",fresh:"新鲜",ok:"合格",old:"旧",
  stockAddManual:"手动添加",stockEditTitle:"编辑库存",stockCategory:"类别",stockMinLevel:"最低库存",stockCurrentLevel:"当前库存",stockAddBtn:"+ 添加",stockDetail:"库存详情",stockAllCats:"所有类别",stockSearch:"搜索项目...",noStock:"暂无库存",invoiceTitle:"发票 AI",invoiceUpload:"📷 上传发票照片",invoiceAnalyze:"AI 分析",invoiceItems:"项目",invoiceTotalExtracted:"检测到 {count} 项",addToStock:"+ 添加到库存",packageUnit:"包装/单位",packageQty:"包装数量",unitQty:"每单位数量",invoiceSaveBtn:"✓ 全部添加到库存",
  recipeAdd:"添加食谱",recipeEdit:"编辑食谱",recipeSave:"✓ 保存",recipeAddIngredient:"+ 添加配料",recipeMainCat:"主类别",recipeSubCat:"子类别",recipeSearch:"搜索食谱...",noRecipes:"暂无食谱",recipePhotoCapture:"📷 拍照 / 上传",recipeAnalyzing:"AI 分析中...",recipeCostDetail:"详细成本",recipeDelete:"× 删除",recipeDeleteConfirm:"{name} 将被删除。确认？",
  menuCreateTitle:"菜单卡",menuTypeCocktail:"鸡尾酒",menuTypeBreakfast:"早餐",menuTypeLunch:"午餐",menuTypeDinner:"晚餐",menuTypeGala:"晚会 / 婚礼",menuSaved:"菜单已保存",menuDeleteConfirm:"\"{name}\" 将被删除。确认？",menuAddSection:"+ 添加部分",menuAddItem:"+ 项目",menuSectionTitle:"部分标题",menuItemTitle:"项目名称",menuItemDesc:"描述",menuItemCalorie:"卡路里 (千卡/100克)",menuItemPrice:"价格",menuAutoLoad:"从食谱加载",menuPreviewBtn:"👁 预览",menuEditBtn:"编辑",menuPrintBtn:"🖨",menuPageAuto:"自动",menuPageSingle:"单页",menuPageMulti:"多页",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 分享",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"Claude OCR 和分析需要 API 密钥。请在 https://console.anthropic.com 获取",settingsProfileNamePh:"输入您的姓名",settingsProfileWorkplacePh:"例如：Grand Hotel",settingsProfileDeptPh:"例如：糕点部",settingsProfileRolePh:"例如：糕点主厨",settingsBackupDesc:"以 JSON 格式下载所有数据或恢复。",settingsRestoreConfirm:"现有数据将被覆盖。确认？",settingsRestoreSuccess:"恢复完成",settingsDevDesc:"用于调试",settingsExpenseName:"开支名称",settingsExpenseAmount:"月金额 (₺)",settingsPersonnelName:"人员姓名",settingsPersonnelSalary:"月薪 (₺)",settingsStorageName:"存储名称",settingsStorageType:"类型",settingsStorageTemp:"目标温度 (°C)",settingsAddIngredient:"要跟踪的配料名称",settingsAddOrg:"组织名称",
  stockEmpty:"清空库存...",stockAdded:"已添加到库存",stockDeducted:"已从库存扣除",insufficientStock:"库存不足：{name}",
  emptyStateRecipes:"创建您的第一个食谱。使用 📷 AI 或手动添加。",emptyStateStock:"使用 📄 发票 (AI) 或 + 手动 按钮向库存添加项目。",emptyStateMenus:"创建您的第一个菜单。从食谱自动加载或手动添加。",
  profileRequired:"缺少用户资料。请转到设置。",goToSettings:"转到设置"
};
I18N.ar={
  tabRecipes:"الوصفات",tabStock:"المخزون",tabProduction:"الإنتاج",tabReports:"التقارير",tabMenus:"القوائم",tabAssistant:"المساعد",tabSettings:"الإعدادات",
  save:"حفظ",cancel:"إلغاء",delete:"حذف",edit:"تعديل",close:"إغلاق",add:"إضافة",back:"رجوع",next:"التالي",search:"بحث",filter:"تصفية",all:"الكل",yes:"نعم",no:"لا",confirm:"تأكيد",remove:"إزالة",update:"تحديث",name:"الاسم",amount:"الكمية",unit:"الوحدة",price:"السعر",date:"التاريخ",time:"الوقت",note:"ملاحظة",optional:"اختياري",required:"مطلوب",loading:"جاري التحميل...",error:"خطأ",success:"نجح",warning:"تحذير",
  recipes:"الوصفات",newRecipe:"+ وصفة جديدة",recipePhoto:"📷 وصفة (ذكاء اصطناعي)",recipeManual:"+ يدوي",recipeDetail:"تفاصيل الوصفة",ingredients:"المكونات",allergens:"المواد المسببة للحساسية",cookMethod:"الطبخ",cookTemp:"الحرارة",cookTime:"الوقت",servings:"الحصص",totalWeight:"الوزن الإجمالي",portionG:"جرام لكل حصة",calories:"السعرات الحرارية",yieldL:"المحصول",recipeName:"اسم الوصفة",deductStock:"📦 خصم من المخزون والإنتاج",sampleLabel:"🧪 ملصق العينة الشاهدة",editPhoto:"📷 تغيير الصورة",addPhoto:"📷 إضافة صورة",costAnalysis:"تحليل التكلفة",material:"المواد",overhead:"ثابت+موظفين",realCost:"التكلفة الحقيقية",perPortion:"حصة",costHint:"💡 أضف النفقات الثابتة وتكاليف الموظفين في الإعدادات لرؤية التكلفة الحقيقية.",
  stock:"المخزون",stockMaterial:"📦 المواد",stockLots:"🏷 أرقام الدفعات",invoice:"📄 الفاتورة (ذكاء اصطناعي)",manualAdd:"+ يدوي",lowStock:"⚠ مخزون منخفض",products:"عناصر",currentLots:"أرقام الدفعات الحالية",noTrackedIngs:"لا توجد مكونات متتبعة",noTrackedIngsHint:"الإعدادات ← قابلية التتبع ← إضافة مكونات",lotMissing:"لم يتم إدخال الدفعة",lotToday:"اليوم",lotDaysAgo:"أيام",lotEntered:"تم الإدخال",lotHint:"💡 امسح أرقام الدفعات من العبوات كل صباح. 🔴 مفقود · 🔵 حالي · 🟡 قديم (3+ أيام).",
  production:"الإنتاج",productionRecord:"🍱 سجل الإنتاج",productionReports:"📊 التقارير",storage:"موقع التخزين",expiryDays:"انتهاء الصلاحية (أيام)",daysLater:"أيام",risk:"المخاطر",riskHigh:"عالي",riskMedium:"متوسط",riskLow:"منخفض",riskHint:"محسوب تلقائيًا، قابل للتعديل",noteHint:"مثل: لخدمة المساء",produceSaveBtn:"✓ الإنتاج والخصم",portions:"حصص",lot:"دفعة",produced:"تم الإنتاج",expires:"ينتهي",consumed:"تم الاستهلاك",active:"نشط",fire:"نفايات",consumeAll:"✓ تم استهلاك الكل",printLabel:"🏷 طباعة الملصق",partialConsume:"استهلاك جزئي",decrease:"− خصم",markFire:"× وضع علامة كنفايات",gotoReport:"📊 الذهاب إلى التقرير",noProduction:"لا يوجد إنتاج بعد",noProductionHint:"استخدم \"خصم من المخزون\" من تفاصيل الوصفة لإنشاء سجل إنتاج",consumedFull:"سيتم وضع علامة {name} ({portions} حصص) كمستهلكة. تأكيد؟",invalidAmount:"كمية غير صالحة",howManyFire:"كم حصة نفايات؟ (الإجمالي {total})",ingredientLotsTitle:"🏷 أرقام دفعات المكونات",lotMissingHint:"💡 للدفعات المفقودة، اذهب إلى المخزون ← أرقام الدفعات.",
  sampleTitle:"🧪 ملصق العينة الشاهدة",sampleDesc:"سجل عينة متوافق مع ISO 22000 — يوصى بالتخزين لمدة 72 ساعة",sampleName:"اسم العينة",sampleTaker:"تم أخذ العينة بواسطة",sampleOrgLabel:"المنظمة",sampleOrgSelect:"اختر...",samplePax:"الأشخاص",sampleLocation:"معلومات إضافية / الموقع",sampleLocationHint:"مثل: حفل زفاف، طاولة VIP",sampleDateTime:"تاريخ ووقت جمع العينة",sampleTemp:"درجة حرارة العينة (°C)",sampleTempHint:"لحظة التقديم",sampleHours:"مدة التخزين (ساعات)",samplePrintSave:"🏷 طباعة وحفظ",sampleOrgOrLoc:"أدخل المنظمة أو الموقع",sampleProfileWarn:"⚠ أدخل الاسم في الإعدادات ← الملف الشخصي",
  reports:"التقارير",reportsProduction:"📋 تقارير الإنتاج",reportsStorage:"🗄 التحكم في التخزين",noReports:"لا توجد تقارير بعد",checkTotal:"فحوصات",checkComplete:"✓ مكتمل",addCheck:"+ إضافة فحص",noChecksYet:"لم يتم تسجيل أي فحوصات بعد.",storageCheckHint:"💡 التحكم في درجة الحرارة 3 مرات يوميًا لكل تخزين (09:00، 15:00، 21:00).",
  notificationsTitle:"🔔 الإشعارات",notifEnabled:"إشعارات الروبوت نشطة",notifDesc:"تظهر الإشعارات في الدردشة، بدون نوافذ منبثقة",notifStorage:"تذكير بفحص التخزين (09:00، 15:00، 21:00)",notifExpired:"تحذير انتهاء الصلاحية",notifLow:"تحذير المخزون المنخفض",testNotif:"إرسال رسالة اختبار",
  userProfile:"الملف الشخصي",fullName:"الاسم الكامل",workplace:"مكان العمل",department:"القسم",role:"المنصب",profileHint:"تُستخدم هذه البيانات في نماذج FR.05 وFR.06",placeholderName:"أدخل اسمك",placeholderWorkplace:"فندق / مطعم",placeholderDept:"القسم",placeholderRole:"المنصب",tabTodo:"المهام",
  productionsCap:"الإنتاج",samplesCap:"العينات",firesCap:"النفايات",
  storageCheckTitle:"🗄 فحص التخزين",controlTime:"وقت الفحص",controlParams:"معاملات الفحص",paramEtiket:"فحص الملصق",paramAcikGida:"طعام مفتوح",paramYerdeDuran:"طعام على الأرض",paramCigPismis:"فصل نيء/مطبوخ",paramTemizlik:"النظافة",paramAlerjen:"فصل رفوف المواد المسببة للحساسية",paramSkt:"الصلاحية/العمر التخزيني",ok:"✓ موافق",notOk:"✗ غير موافق",productTemp:"درجة حرارة المنتج (°C)",gaugeValue:"قيمة المقياس (°C)",explanation:"الشرح (اختياري)",explanationHint:"إجراء تصحيحي إلخ",atLeastOneTemp:"أدخل قيمة واحدة على الأقل من درجة الحرارة",
  menus:"بطاقات القائمة",newMenu:"+ قائمة جديدة",chooseMenuType:"اختر نوع القائمة",menuEdit:"تعديل القائمة",preview:"معاينة",menuName:"اسم القائمة",font:"الخط",theme:"المظهر",customSizes:"📐 أحجام مخصصة (مم)",or:"أو",showPrice:"إظهار السعر",showCalorie:"إظهار السعرات",showAllergen:"إظهار المواد المسببة للحساسية",bgImage:"🖼 صورة الخلفية (اختياري)",chooseBg:"📷 اختيار",changeBg:"📷 تغيير",cropBg:"✂ قص",bgOpacity:"الشفافية",bgHint:"💡 أضف خلفيات من الرخام أو الورق أو نسيج الخشب. يفتح محرر القص بعد التحميل.",menuTemplates:"📑 قالب القائمة",loadTemplate:"تحميل قالب...",saveTemplate:"💾 حفظ القالب",templateHint:"💡 احفظ الخط والمظهر والحجم والخلفية كقالب.",templateName:"اسم القالب:",defaultTemplateName:"قالب {menuName}",templateSaveConfirm:"تحميل القالب \"{name}\"؟ (يتم الحفاظ على الأقسام والعناصر، يتغير التصميم)",templateSaved:"تم حفظ \"{name}\" كقالب.",templateDeletePrompt:"اسم القالب للحذف؟",templateNotFound:"غير موجود",
  cropTitle:"✂ قص الصورة",cropDesc:"اسحب لإعادة التموضع، المنزلق للتكبير/التصغير. المنطقة داخل الإطار تصبح خلفية القائمة.",cropZoom:"🔍 تكبير",cropApply:"✓ قص وتطبيق",
  settings:"الإعدادات",apiKey:"مفتاح API",language:"اللغة",darkMode:"المظهر الداكن",userProfile:"👤 الملف الشخصي",fullName:"الاسم الكامل",workplace:"مكان العمل",department:"القسم",role:"الدور",profileHint:"💡 يُستخدم في ملصقات العينات الشاهدة وسجلات الإنتاج المستقبلية.",traceabilityTitle:"🔍 قابلية التتبع (ISO 22000)",traceabilityMode:"وضع قابلية التتبع",traceabilityDesc:"يُمكّن تتبع الدفعات وتقارير ISO 22000 (FR.06، FR.12) ومخرجات PDF",resetHourLabel:"وقت إعادة تعيين الدفعات في نهاية اليوم",resetHourDesc:"يتم إعادة تعيين أرقام الدفعات يوميًا في هذا الوقت",trackedIngsLabel:"المكونات المتتبعة",trackedIngsPh:"أضف مكونًا (مثل: الدقيق، الحليب، البيض)",trackedIngsHint:"💡 قم بتحديث أرقام دفعات هذه المكونات كل صباح. تُستخدم تلقائيًا في FR.06.",organizationsTitle:"🏢 المنظمات (للعينات)",organizationsPh:"أضف منظمة (مثل: حفل زفاف 200 شخص، حفل غالا)",organizationsHint:"💡 تُستخدم هذه القائمة كقائمة منسدلة في ملصق العينة الشاهدة.",expensesTitle:"💼 مصروفات الأعمال",fixedExpenses:"المصروفات الثابتة (شهريًا)",personnel:"الموظفون (راتب شهري)",addFixed:"+ إضافة مصروف ثابت",addPersonnel:"+ إضافة موظف",monthlyPortions:"الإنتاج الشهري المقدر (حصص)",portionHint:"يُستخدم لحساب حصة النفقات العامة في تكاليف الوصفات.",fixedTotal:"إجمالي المصروفات الثابتة:",personnelTotal:"إجمالي الموظفين:",monthlyTotal:"إجمالي المصروفات الشهرية:",perPortionShare:"حصة المصروفات لكل جزء:",storageAreas:"🗄 مناطق التخزين",addStorage:"+ إضافة منطقة تخزين",storageHint:"💡 أضف أسماء ثلاجاتك/مجمداتك/تخزينك هنا.",backup:"💾 النسخ الاحتياطي",backupBtn:"⬇ نسخ احتياطي",restoreBtn:"⬆ استعادة",developer:"🔧 المطور",copyLogs:"📋 نسخ السجلات",
  fr06Title:"نموذج مراقبة تحضير الطعام",fr06ProductCtrl:"مراقبة تحضير المنتج",fr06ProductIng:"اسم المنتج · المكونات · رقم الدفعة",fr06PrepStart:"بداية التحضير",fr06PrepEnd:"نهاية التحضير",fr06ProductTemp:"حرارة المنتج (°C)",fr06AmbientTemp:"الحرارة المحيطة (°C)",fr06Controller:"المراقب",fr06CookCtrl:"مراقبة حرارة الطبخ / إعادة التسخين",fr06ProductDef:"تعريف المنتج",fr06PartyNo:"رقم الدفعة",fr06Cook:"الطبخ (°C)",fr06Reheat:"إعادة التسخين (°C)",fr06CorrectiveAction:"إجراء تصحيحي",fr06FreezingCtrl:"مراقبة إذابة الثلج",fr06ThawStart:"بداية الإذابة",fr06ThawEnd:"نهاية الإذابة",fr06Temp:"درجة الحرارة",fr06FastCoolCtrl:"مراقبة التبريد السريع / التجميد",fr06Entry:"الدخول (الوقت-°C)",fr06Exit:"الخروج (الوقت-°C)",fr06Chiller:"المبرد",fr06IceWater:"ماء مثلج/-18°C",fr06PlateHold:"حفظ في الطبق",fr06ProductOrg:"المنتج / المنظمة",fr06PlateTime:"وقت التقديم",fr06Service:"الخدمة",fr06InnerTemp:"حرارة داخلية",fr06Ctrl:"مراقبة",fr06FosterHold:"Foster/تجفيف",fr06Notes:"ملاحظات التطبيق",fr06NoteRisk:"تحضير طعام خطير:",fr06NoteRiskText:"يجب ألا تتجاوز درجة الحرارة المحيطة +24°C. للأطعمة عالية الخطورة (لحم-دجاج-سمك) 16°C كحد أقصى. وقت التحضير: 45 دقيقة عند +20°C، 10 دقائق عند +25°C.",fr06NoteCool:"التبريد/التجميد:",fr06NoteCoolText:"10°C كحد أقصى خلال 90 دقيقة. لا تتجاوز أبدًا 4 ساعات للوصول إلى 5°C. التجميد السريع: -18°C خلال 4 ساعات.",fr06NoteCook:"الطبخ/إعادة التسخين:",fr06NoteCookText:"الحرارة الداخلية 75°C، السطح 65°C. إعادة التسخين 82°C على الأقل.",fr06NotePlate:"حفظ في الطبق:",fr06NotePlateText:"لا يُسمح بالحفظ لأكثر من 15 دقيقة.",fr06NoteThaw:"الإذابة:",fr06NoteThawText:"يُسمح بالإذابة في ماء مثلج. يُخرج المنتج المجمد قبل يوم واحد.",fr06DocNo:"رقم المستند",fr06DateLbl:"التاريخ",fr06DefaultDept:"المطبخ / القسم",
  fr12Title:"نموذج العينة الشاهدة",fr12SampleOrg:"اسم العينة / المنظمة",fr12Pax:"الأشخاص",fr12DateTime:"تاريخ ووقت الجمع",fr12Temp:"حرارة العينة",fr12Taker:"تم أخذ العينة بواسطة",fr12Signature:"التوقيع",fr12Explanation:"الشرح",fr12Note:"العينة الشاهدة:",fr12NoteText:"يجب أخذ عينات شاهدة لجميع المنظمات التي تخدم أكثر من 20 شخصًا (بما في ذلك الإفطار وكافتيريا الموظفين) وتخزينها عند 5°C كحد أقصى لمدة 72 ساعة. 200 جرام كحد أدنى.",
  fr05Title:"نموذج مراقبة التخزين / الخزانة اليومية",fr05Dept:"القسم:",fr05Cabinet:"الخزانة/التخزين:",fr05Target:"الهدف:",fr05CtrlParams:"معاملات المراقبة",fr05TempCtrl:"مراقبة درجة الحرارة (°C)",fr05CtrlTime:"وقت الفحص",fr05NoteDry:"التخزين الجاف:",fr05NoteDryText:"درجة الحرارة <25°C والرطوبة النسبية ≤65%.",fr05NoteCold:"التخزين البارد/المجمد:",fr05NoteColdText:"المناطق الباردة عند 5°C أو أقل، التخزين المجمد عند -18°C أو أقل.",fr05NoteApp:"ملاحظة التطبيق:",fr05NoteAppText:"أثناء التفتيش، اكتب (+) للحالات المطابقة، (-) لغير المطابقة.",
  labelProduction:"ملصق الإنتاج",labelFreezer:"ملصق المجمد",labelProductName:"اسم المنتج",labelProductionDate:"تاريخ الإنتاج",labelOpeningDate:"تاريخ الفتح",labelExpiryDate:"تاريخ انتهاء الصلاحية",labelLotNumber:"رقم الدفعة",labelFreezingDate:"تاريخ التجميد",labelFreezingTime:"وقت التجميد",labelThawingDate:"تاريخ الإذابة",labelThawingTime:"وقت الإذابة",sampleWarn:"⚠ عينة شاهدة — لا تستهلك ⚠",sampleNameLbl:"اسم العينة:",sampleOrgLbl:"المنظمة / الموقع:",sampleTakerLbl:"تم أخذ العينة بواسطة:",sampleDateTimeLbl:"تاريخ ووقت الجمع:",sampleStorageLbl:"مدة التخزين:",sampleEndLbl:"النهاية:",sampleTempLbl:"حرارة العينة:",
  select:"اختر",none:"لا شيء",today:"اليوم",fresh:"طازج",ok:"موافق",old:"قديم",
  stockAddManual:"إضافة يدوية",stockEditTitle:"تعديل المخزون",stockCategory:"الفئة",stockMinLevel:"الحد الأدنى",stockCurrentLevel:"المخزون الحالي",stockAddBtn:"+ إضافة",stockDetail:"تفاصيل المخزون",stockAllCats:"جميع الفئات",stockSearch:"البحث عن العناصر...",noStock:"لا يوجد مخزون بعد",invoiceTitle:"فاتورة AI",invoiceUpload:"📷 تحميل صورة الفاتورة",invoiceAnalyze:"تحليل AI",invoiceItems:"العناصر",invoiceTotalExtracted:"تم اكتشاف {count} عناصر",addToStock:"+ إلى المخزون",packageUnit:"عبوة/وحدة",packageQty:"عدد العبوات",unitQty:"الكمية لكل وحدة",invoiceSaveBtn:"✓ إضافة الكل إلى المخزون",
  recipeAdd:"إضافة وصفة",recipeEdit:"تعديل الوصفة",recipeSave:"✓ حفظ",recipeAddIngredient:"+ إضافة مكون",recipeMainCat:"الفئة الرئيسية",recipeSubCat:"الفئة الفرعية",recipeSearch:"البحث عن الوصفات...",noRecipes:"لا توجد وصفات بعد",recipePhotoCapture:"📷 التقاط / تحميل صورة",recipeAnalyzing:"AI يحلل...",recipeCostDetail:"التكلفة التفصيلية",recipeDelete:"× حذف",recipeDeleteConfirm:"سيتم حذف {name}. تأكيد؟",
  menuCreateTitle:"بطاقات القائمة",menuTypeCocktail:"كوكتيل",menuTypeBreakfast:"إفطار",menuTypeLunch:"غداء",menuTypeDinner:"عشاء",menuTypeGala:"حفل / زفاف",menuSaved:"تم حفظ القائمة",menuDeleteConfirm:"سيتم حذف \"{name}\". تأكيد؟",menuAddSection:"+ إضافة قسم",menuAddItem:"+ عنصر",menuSectionTitle:"عنوان القسم",menuItemTitle:"اسم العنصر",menuItemDesc:"الوصف",menuItemCalorie:"السعرات (كالوري/100غ)",menuItemPrice:"السعر",menuAutoLoad:"التحميل من الوصفة",menuPreviewBtn:"👁 معاينة",menuEditBtn:"تعديل",menuPrintBtn:"🖨",menuPageAuto:"تلقائي",menuPageSingle:"صفحة واحدة",menuPageMulti:"عدة صفحات",menuDownloadBtn:"📥 PDF",menuShareBtn:"📤 مشاركة",
  settingsApiKeyPh:"sk-ant-api03-...",settingsApiKeyHint:"مفتاح API مطلوب لـ Claude OCR والتحليل. احصل عليه من https://console.anthropic.com",settingsProfileNamePh:"أدخل اسمك",settingsProfileWorkplacePh:"مثل: Grand Hotel",settingsProfileDeptPh:"مثل: الحلويات",settingsProfileRolePh:"مثل: رئيس الحلويات",settingsBackupDesc:"قم بتنزيل جميع البيانات بصيغة JSON أو استعادتها.",settingsRestoreConfirm:"سيتم الكتابة فوق البيانات الحالية. تأكيد؟",settingsRestoreSuccess:"اكتمل الاستعادة",settingsDevDesc:"لتصحيح الأخطاء",settingsExpenseName:"اسم المصروف",settingsExpenseAmount:"المبلغ الشهري (₺)",settingsPersonnelName:"اسم الشخص",settingsPersonnelSalary:"الراتب الشهري (₺)",settingsStorageName:"اسم التخزين",settingsStorageType:"النوع",settingsStorageTemp:"درجة الحرارة المستهدفة (°C)",settingsAddIngredient:"اسم المكون المراد تتبعه",settingsAddOrg:"اسم المنظمة",
  stockEmpty:"جاري تفريغ المخزون...",stockAdded:"تمت الإضافة إلى المخزون",stockDeducted:"تم الخصم من المخزون",insufficientStock:"مخزون غير كافي: {name}",
  emptyStateRecipes:"أنشئ وصفتك الأولى. استخدم 📷 AI أو أضف يدويًا.",emptyStateStock:"استخدم 📄 فاتورة (AI) أو + يدوي لإضافة عناصر إلى المخزون.",emptyStateMenus:"أنشئ قائمتك الأولى. التحميل التلقائي من الوصفات أو يدويًا.",
  profileRequired:"الملف الشخصي مفقود. اذهب إلى الإعدادات.",goToSettings:"إلى الإعدادات"
};

// Dil yardımcı fonksiyonu
const tr_=(lang,key,vars)=>{
  const dict=I18N[lang]||I18N.tr;
  let s=dict[key]||I18N.tr[key]||key;
  if(vars){Object.keys(vars).forEach(k=>{s=s.replace(`{${k}}`,vars[k])});}
  return s;
};


// ═══ THEME ═══
const THEMES={
  light:{bg:"#faf8f3",card:"#fff",cardB:"#ebe5d8",text:"#1a1612",ts:"#6b6357",tm:"#a09684",accent:"#c8965a",acB:"#c8965a18",acBo:"#c8965a44",border:"#ebe5d8",inBg:"#f5f1e8",inBo:"#e0d9c7",danger:"#c75450",danBg:"#c7545012",danBo:"#c7545033",success:"#5a8a5e",sucBg:"#5a8a5e12",sucBo:"#5a8a5e33",warn:"#b8924a",waBg:"#b8924a14",waBo:"#b8924a33",glass:"rgba(255,253,247,0.85)",overlay:"rgba(26,22,18,0.4)",topBar:"rgba(250,248,243,0.92)",pBg:"#f0ebe0",pA:"#1a1612",pAT:"#fff",pT:"#6b6357",cardS:"0 1px 3px #00000008,0 4px 16px #0000000a",cardSH:"0 8px 28px #00000014"},
  dark:{bg:"#1a1612",card:"#251f1a",cardB:"#352e26",text:"#f0ebe0",ts:"#a09684",tm:"#5a5247",accent:"#d4a574",acB:"#d4a57418",acBo:"#d4a57444",border:"#352e26",inBg:"#2d2620",inBo:"#3f3830",danger:"#e07070",danBg:"#e0707012",danBo:"#e0707033",success:"#7ab880",sucBg:"#7ab88012",sucBo:"#7ab88033",warn:"#d4a84a",waBg:"#d4a84a14",waBo:"#d4a84a33",glass:"rgba(37,31,26,0.88)",overlay:"rgba(0,0,0,0.6)",topBar:"rgba(26,22,18,0.92)",pBg:"#2d2620",pA:"#d4a574",pAT:"#1a1612",pT:"#a09684",cardS:"0 1px 3px #00000030,0 4px 16px #00000020",cardSH:"0 8px 28px #00000040"}
};

// ═══ STYLE HELPERS ═══
const iSt=(t)=>({background:t.inBg,border:`1px solid ${t.inBo}`,borderRadius:12,padding:"12px 14px",fontSize:15,color:t.text,width:"100%",outline:"none",fontFamily:"'Inter Tight',sans-serif"});
const lSt=(t)=>({fontSize:11,fontWeight:700,color:t.tm,letterSpacing:"0.1em",display:"block",marginBottom:6,textTransform:"uppercase"});
const bSt=(v,t)=>({padding:"11px 20px",borderRadius:12,cursor:"pointer",fontSize:14,fontWeight:600,border:"none",transition:"all .2s",background:v==="p"?t.accent:v==="d"?t.danBg:t.pBg,color:v==="p"?(t.bg==="#1a1612"?"#1a1612":"#fff"):v==="d"?t.danger:t.ts,...(v==="d"?{border:`1px solid ${t.danBo}`}:{})});
const cSt=(t)=>({background:t.card,borderRadius:18,border:`1px solid ${t.cardB}`,boxShadow:t.cardS,transition:"all .25s"});
const mOv=(t)=>({position:"fixed",inset:0,background:t.overlay,backdropFilter:"blur(14px)",WebkitBackdropFilter:"blur(14px)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px 8px",overflowX:"hidden"});
const mPn=(t)=>({...cSt(t),padding:20,maxWidth:480,width:"100%",maxHeight:"92vh",overflowY:"auto",overflowX:"hidden",background:t.glass,backdropFilter:"blur(28px)",WebkitBackdropFilter:"blur(28px)"});

// ═══ LOGO — İki kavisli hançer (Osmanlı stili), ince daire çerçeve (madde 11) ═══
const Logo=({size=28,c="#c8965a"})=><svg width={size} height={size} viewBox="0 0 48 48" fill="none">
  <circle cx="24" cy="24" r="22" stroke={c} strokeWidth="1.2" opacity="0.5"/>
  {/* Sol hançer — sola eğik, kavisle */}
  <path d="M17 10 C13 16, 12 24, 16 36 C17 38, 18 38, 18.5 36 C15 26, 16 18, 20 12 Z" fill={c} opacity="0.9"/>
  <path d="M17 10 C16 8, 19 7, 20 8 C20 10, 20 12, 20 12 C18.5 11, 17 10, 17 10 Z" fill={c}/>
  {/* Sağ hançer — sağa eğik, karşılıklı */}
  <path d="M31 10 C35 16, 36 24, 32 36 C31 38, 30 38, 29.5 36 C33 26, 32 18, 28 12 Z" fill={c} opacity="0.9"/>
  <path d="M31 10 C32 8, 29 7, 28 8 C28 10, 28 12, 28 12 C29.5 11, 31 10, 31 10 Z" fill={c}/>
  {/* Kavisle bağlayan alt çizgi */}
  <path d="M18.5 36 Q24 40, 29.5 36" stroke={c} strokeWidth="1.5" strokeLinecap="round" fill="none" opacity="0.6"/>
</svg>;

// ═══ ALLERGEN BADGE ═══
const ABadge=({a,t,big})=>{const al=ALLERGENS.find(x=>x.id===a);if(!al)return null;return <span style={{background:al.c+"15",border:`1px solid ${al.c}40`,color:al.c,borderRadius:8,padding:big?"5px 12px":"3px 9px",fontSize:big?13:11,fontWeight:600,whiteSpace:"nowrap",display:"inline-flex",alignItems:"center",gap:4}}>{al.icon} {allergenL(al,t?.lang||"tr")}</span>};
const DBadge=({d,t})=>{const di=DIETS.find(x=>x.id===d);if(!di)return null;return <span style={{background:t.sucBg,border:`1px solid ${t.sucBo}`,color:t.success,borderRadius:8,padding:"3px 9px",fontSize:11,fontWeight:600,whiteSpace:"nowrap",display:"inline-flex",alignItems:"center",gap:4}}>{di.icon} {dietL(di,t?.lang||"tr")}</span>};

// ═══ PHOTO PICKER ═══
const PhotoPick=({onImg,t})=>{
  const ref=useRef();const[prev,setPrev]=useState(null);
  return <div>
    <input ref={ref} type="file" accept="image/*" style={{display:"none"}} onChange={async e=>{
      const f=e.target.files[0];if(!f)return;
      const reader=new FileReader();
      reader.onload=async ev=>{setPrev(ev.target.result);const{base64,mediaType}=await resizeImg(ev.target.result);onImg(base64,mediaType);};
      reader.readAsDataURL(f);
    }}/>
    {!prev?<div onClick={()=>ref.current.click()} style={{border:`2px dashed ${t.inBo}`,borderRadius:16,padding:"36px 20px",textAlign:"center",cursor:"pointer",background:t.inBg}}>
      <div style={{fontSize:38,marginBottom:10,opacity:0.6}}>📷</div>
      <div style={{fontWeight:600,color:t.text}}>{t.L.detailPhotoSelect||"Select or capture"}</div>
      <div style={{fontSize:12,color:t.tm,marginTop:4}}>{t.L.detailPhotoOption||"Gallery or camera"}</div>
    </div>:<div>
      <img src={prev} style={{width:"100%",borderRadius:14,maxHeight:200,objectFit:"cover",marginBottom:10}}/>
      <button onClick={()=>{setPrev(null);onImg(null,null);ref.current.value=""}} style={{...bSt("s",t),width:"100%",fontSize:13}}>↩ Değiştir</button>
    </div>}
  </div>;
};

// ═══ API KEY MODAL ═══
const KeyModal=({onSave,t,initial})=>{
  const[k,setK]=useState(initial||"");
  return <div style={mOv(t)}><div style={mPn(t)}>
    <Logo size={56} c={t.accent}/>
    <h2 style={{fontSize:24,marginTop:14,marginBottom:6,color:t.text}}>Kitchen Manager Pro</h2>
    <p style={{color:t.tm,fontSize:14,marginTop:0,marginBottom:20}}>Devam etmek için Anthropic API anahtarınızı girin. Anahtar sadece bu cihazda saklanır.</p>
    <input style={iSt(t)} value={k} onChange={e=>setK(e.target.value)} placeholder="sk-ant-api03-..." type="password"/>
    <button onClick={()=>{if(k.trim())onSave(k.trim())}} style={{...bSt("p",t),width:"100%",marginTop:14,padding:14}}>Kaydet ve Devam Et</button>
    <p style={{color:t.tm,fontSize:11,marginTop:14,marginBottom:0,textAlign:"center"}}>API anahtarı almak için: console.anthropic.com</p>
  </div></div>;
};

// ═══ OCR — Kısaltma sözlüğü (madde 5) ═══
const BIM_ABBR={
  "Y.YAĞLI":"Yarım Yağlı","TMYG":"Tam Yağlı Yoğurt","YYMK":"Yarım Yağlı Makarna",
  "GT":"g","GP":"g Paket","BRO":"Browni","PNR":"Peynir","DMTS":"Domates",
  "PTTS":"Patates","SRMSAK":"Sarımsak","ZYTNYĞI":"Zeytinyağı","KRM":"Krema",
  "YNMRK":"Yumurta","TRYBĞ":"Tereyağı","ÇKOLATA":"Çikolata","VNLYA":"Vanilya",
  "MKR":"Makarna","PRN":"Pirinç","NŞSTA":"Nişasta","TZ":"Toz","MHL":"Mısır Unu",
  "FLTR":"Filtre","HMSR":"Hamur Serbest","YRKDŞ":"Yürekdaş"
};
function applyAbbr(name){
  let n=name.replace(/\*/g,"").trim();
  for(const[k,v]of Object.entries(BIM_ABBR)){
    n=n.replace(new RegExp("\\b"+k+"\\b","gi"),v);
  }
  return n.trim();
}
// Normalizasyon: binlik/ondalık, TR harf, 4 ondalık destek
function normNum(s){
  if(!s)return 0;
  const str=String(s).replace(/\s/g,"");
  // son virgül kuralı: 1.250,00 → 1250.00
  const lastComma=str.lastIndexOf(",");
  const lastDot=str.lastIndexOf(".");
  let normalized=str;
  if(lastComma>lastDot){
    // virgül ondalık ayracı
    normalized=str.replace(/\./g,"").replace(",",".");
  }else if(lastDot>lastComma){
    // nokta ondalık ayracı
    normalized=str.replace(/,/g,"");
  }
  return parseFloat(normalized)||0;
}
function trNorm(s){return s.replace(/İ/g,"İ").replace(/ı/g,"ı").replace(/Ğ/g,"Ğ").replace(/ğ/g,"ğ").replace(/Ü/g,"Ü").replace(/ü/g,"ü").replace(/Ö/g,"Ö").replace(/ö/g,"ö").replace(/Ş/g,"Ş").replace(/ş/g,"ş").replace(/Ç/g,"Ç").replace(/ç/g,"ç")}

// Özet satır filtresi
const SUMMARY_KWS=["TOPKDV","TOPLAM","BANKA","NAKIT","KREDİ","KART","ONAY","REF","POS","GS","ETTN","VKN","IBAN","ÖDENECEK","MAL/HİZMET","FIŞTOPLAM","ARATOPLAM","ARATOP","KDV TOPLAM","GENEL TOPLAM"];
function isSummaryRow(s){const u=s.toUpperCase();return SUMMARY_KWS.some(k=>u.includes(k))||/^TR\d{2}/.test(s)||/^[A-Z0-9]{10,}$/.test(s.trim())}

// MOD A: BİM yazarkasa fişi parser (ön işleme, AI'ya da gönderilir)
function preParseModA(text){
  const lines=text.split("\n").map(l=>l.trim()).filter(l=>l.length>1);
  const items=[];
  for(let i=0;i<lines.length;i++){
    const l=lines[i];
    if(isSummaryRow(l))continue;
    // %KDV *FİYAT satırı tespiti
    const kdvMatch=l.match(/(%\d+)\s+\*?([\d.,]+)/);
    if(kdvMatch){
      // bir sonraki satırda N ad x BF olabilir
      const nextL=lines[i+1]||"";
      const countMatch=nextL.match(/^(\d+)\s+(.+?)\s+[xX]\s+([\d.,]+)/);
      if(countMatch){
        items.push({
          raw:l+" | "+nextL,
          kdv:parseInt(kdvMatch[1],10),
          fiyat:normNum(kdvMatch[2]),
          adet:parseInt(countMatch[1],10),
          ad:applyAbbr(countMatch[2]),
          bf:normNum(countMatch[3])
        });
        i++;continue;
      }
    }
    // Normal satır: ad fiyat
    const priceMatch=l.match(/^(.+?)\s+([\d.,]+)\s*₺?$/);
    if(priceMatch&&!isSummaryRow(priceMatch[1])){
      items.push({raw:l,ad:applyAbbr(priceMatch[1]),fiyat:normNum(priceMatch[2]),adet:1});
    }
  }
  return items;
}

// MOD B: tablo faturası parser (vgsepetim tipi)
function preParseModB(text){
  const lines=text.split("\n").map(l=>l.trim()).filter(l=>l.length>3);
  const items=[];
  for(const l of lines){
    if(isSummaryRow(l))continue;
    // Sütunlar: sıraNo | kod | ad | miktar | birimFiyat | toplam
    const cols=l.split(/\t|\s{2,}|\|/).map(c=>c.trim()).filter(c=>c);
    if(cols.length<4)continue;
    // Ad içinden ambalaj miktarı ayıkla: "Süt 1 lt", "Tereyağı 250 g", "Su 1500 ml"
    let adRaw=cols.find(c=>isNaN(normNum(c))&&c.length>1)||"";
    const pkgMatch=adRaw.match(/(\d+[\d.,]*)\s*(kg|gr|g|lt|ml|adet|cl)\b/i);
    let pkgAmt=1,pkgUnit="adet";
    if(pkgMatch){
      pkgAmt=normNum(pkgMatch[1]);
      pkgUnit=pkgMatch[2].toLowerCase().replace("gr","g").replace("lt","l");
      adRaw=adRaw.replace(pkgMatch[0],"").trim();
    }
    const nums=cols.filter(c=>!isNaN(normNum(c))&&normNum(c)>0).map(normNum);
    if(nums.length<2)continue;
    const adet=nums[0];
    const birimFiyat=nums[nums.length-2]||0;
    const toplam=nums[nums.length-1]||0;
    // Matematiksel doğrulama
    const hesap=Math.round(adet*birimFiyat*100)/100;
    const flag=Math.abs(hesap-toplam)>0.05;
    items.push({
      ad:applyAbbr(adRaw),adet,pkgAmt,pkgUnit,birimFiyat,toplam,flag,
      toplamStok:adet*pkgAmt
    });
  }
  return items;
}

// OCR AI PROMPT — few-shot örnek, JSON şeması, "uydurma" direktifi (madde 5)
const P_OCR_RECIPE=`Reçeteyi oku, SADECE JSON döndür:
{"name":"","mainCat":"pastry","subCat":"milk","venue":"alacarte","prep":"cooked","cuisine":"turkish","difficulty":"medium","yield":"","servings":4,"calories":null,"totalWeight":0,"ingredients":[{"name":"","amount":""}],"allergens":[],"diets":[],"notes":"","cookMethod":"","cookTemp":"","cookTime":""}

calories: 100 GRAM başına kcal olarak hesapla (porsiyon başına DEĞİL, 100 gram başına). Emin değilsen null.

totalWeight: tüm malzemelerin toplam tahmini ağırlığı GRAM cinsinden. Dönüşüm tablosu:
1 su bardağı un=140g, şeker=200g, kakao=100g, nişasta=120g, pudra şekeri=160g
1 su bardağı yoğurt/süt/krema/su=240g, bal/pekmez=320g, zeytinyağı=216g
1 su bardağı kuruyemiş(file)=100g, ceviz=90g
1 yk un=9g, şeker=13g, tereyağı=15g, kakao=6g, bal=20g
1 tk tuz=8g, kabartma tozu=4g, tarçın=3g
1 yumurta=55g, 1 sarı=20g, 1 ak=35g
1 vanilya çubuğu=3g, 1 defne yaprağı=0.3g, 1 diş sarımsak=6g
1 demet maydanoz=50g, 1ml su=1g, 1ml yağ=0.9g
totalWeight sıfır bırakma, yaklaşık hesapla.
KRİTİK KURAL - ingredients alanı:
ingredients dizisine SADECE ham malzemeler girer (un, şeker, yumurta, süt gibi gıda maddeleri).
ASLA şunları ingredients'a koyma:
- Pişirme talimatları ("180 derecede pişir", "fırına ver", "çırp", "yoğur")
- Sıcaklık değerleri → cookTemp alanına
- Süre bilgileri → cookTime alanına
- Pişirme yöntemi → cookMethod alanına
- Ekipman ("fırın tepsisi", "mikser") → notes alanına
Her ingredient.name sadece malzeme adı olmalı.

Pişirme bilgisi alanları:
- cookMethod: "fırın" | "dondurma" | "buzdolabı" | "kaynat" | "kızart" | "ızgara" | "haşla" | "benmari" | "soğutma" | "" (yok)
- cookTemp: sıcaklık metni, örn: "180°C", "-18°C", "4°C", "" (yok)
- cookTime: süre metni, örn: "35 dk", "2 saat", "24 saat", "" (yok)

Örnek doğru çıktı: reçete "180°C'de 35 dk pişir" diyorsa → cookMethod:"fırın", cookTemp:"180°C", cookTime:"35 dk"
Örnek doğru çıktı: dondurma reçetesi → cookMethod:"dondurma", cookTemp:"-18°C", cookTime:"4 saat"

mainCat: pastry|hot|cold|breakfast|bar
subCat: milk|cake|cookie|pie|dough|sherbet|ice|main|soup|sauce|garnish|rice|meze|salad|snack|coldsauce|egg|pastryb|spread|cereal|brplate|hotdrink|colddrink|cocktail
allergens: gluten,milk,egg,soy,peanut,treenut,sesame,fish,shellfish,mollusk,mustard,celery,lupin,sulfite,goatmilk,buckwheat,legume,sugar,highsalt,alcohol,pork
diets: vegan,vegetarian,glutenfree,lactosefree,sugarfree,kosher,halal`;

const P_OCR_INV=`Türk faturası/fişi analiz et. Mod tespiti:
MOD_A = yazarkasa fişi (BİM, A101, ŞOK tipi): %KDV *FİYAT satırları, kısaltmalar var
MOD_B = tablo faturası (vgsepetim, toptancı tipi): sıra no, stok kodu, ad, miktar, birim fiyat, toplam kolonları

SADECE JSON döndür:
{
  "mod": "A|B",
  "saglayici": "",
  "tarih": "YYYY-MM-DD",
  "urunler": [
    {
      "ad": "temiz ürün adı, kısaltmaları aç",
      "miktar": 0,
      "birim": "g|ml|adet|kg|l",
      "adet": 1,
      "birimFiyat": 0.00,
      "kdvOrani": 1,
      "toplam": 0.00,
      "kategori": "dairy|legumes|dryfood|produce|meatfish|spice|sweetener|cleaning|stationery|other"
    }
  ],
  "toplam": 0.00,
  "kdvToplam": 0.00
}

Kurallar:
- Binlik ayraç: nokta (1.250,00 → 1250.00); ondalık: virgül
- 4 ondalık basamağa kadar destekle
- Özet satırları (TOPLAM, KDV, BANKA, NAKIT, KART, IBAN, ETTN, VKN, ÖDENECEK) ürün listesine EKLEME
- Emin olmadığın alanı boş bırak, ASLA uydurma

KRİTİK KURAL - AMBALAJI KIR, BASE UNIT'E ÇEVİR:
Her ürünü MUTLAKA en küçük birime (g, ml, adet) çevir. Paket halinde kaydetme.

"Irmak Toz Şeker 5 kg - 1 Adet - 123,80 TL" nasıl okumalı:
→ ad: "Toz Şeker"
→ miktar: 5000 (gram olarak)
→ birim: "g"
→ birimFiyat: 0.02476 (1 gram fiyatı = 123.80 / 5000)
→ toplam: 123.80

"Egeden Sızma Zeytinyağı 5 lt - 1 Adet - 345,70 TL" nasıl okumalı:
→ ad: "Sızma Zeytinyağı"
→ miktar: 5000 (ml olarak)
→ birim: "ml"
→ birimFiyat: 0.06914 (1 ml fiyatı = 345.70 / 5000)
→ toplam: 345.70

"Barilla Spagetti 2 kg - 2 Adet - 87,90 TL" nasıl okumalı:
→ ad: "Spagetti"
→ miktar: 4000 (2 adet × 2000 g = 4000 g)
→ birim: "g"
→ birimFiyat: 0.02198 (87.90 / 2000 = 1 paket/gram)
→ toplam: 175.80

"Yumurta 30 LU - 1 adet - 200 TL" nasıl okumalı:
→ ad: "Yumurta"
→ miktar: 30 (adet olarak — yumurta GERÇEK adet, gram değil)
→ birim: "adet"
→ birimFiyat: 6.67 (200 / 30)
→ toplam: 200

Kurallar:
- Ambalaj bilgisi ad'dan ayıkla (5 kg, 1 lt, 500 gr, 250 ml, 30 lu vs)
- kg → g'a çevir (× 1000), l → ml'ye çevir (× 1000)
- Adet sayısı × paket boyutu = toplam miktar (gram/ml)
- birimFiyat = toplam / toplamMiktar (1 gram/ml/adet fiyatı)
- Yumurta, defne yaprağı, vanilya çubuğu, kahve kapsülü gibi gerçekten ADET olan ürünlerde birim "adet" kalır
- Deterjan, temizlik ürünleri adet olarak kalabilir


"ad" alanına SADECE jenerik ürün adı yaz. Marka adını kaldır.
Türk gıda markaları (bunlardan biri varsa MUTLAKA kaldır): Sütaş, İçim, Pınar, Eker, Danone, Yörsan, Tahsildaroğlu, Ülker, Eti, Torku, Knorr, Calve, Tadım, Tat, Bizim, Migros, Koska, Banvit, Şenpiliç, Beypiliç, Komili, Kırlangıç, Yudum, Ayçiçek, Luna, Sana, Marsa, Namet, Maret, Piyale, Nuh'un Ankara, Barilla, Kellogg's, Nestlé, Nescafé, Tchibo, Lipton, Doğuş, Çaykur, Caykur, Ferrero, Hero, Sek, Tikveşli, Muratbey, Tahsildaroğlu, Yayla, Tukaş, Tariş, Kavaklıdere, Doluca, Sek
Marka örnekleri:
- "Sütaş Tam Yağlı Süt 1L" → ad:"Tam Yağlı Süt"
- "İçim Beyaz Peynir 500g" → ad:"Beyaz Peynir"
- "Pınar Kaşar Peyniri" → ad:"Kaşar Peyniri"
- "Eker Süzme Yoğurt" → ad:"Süzme Yoğurt"
- "Knorr Tavuk Bulyon" → ad:"Tavuk Bulyon"
- "Torku Un 5kg" → ad:"Un"
ad alanı asla markayla başlamasın.

Few-shot MOD_A (BİM fişi):
Girdi:
%8 *45,90
3 TMYG YOĞURT x 15,30
%18 *12,50
SUTAS PEYNIR 200G
Çıktı: [{ad:"Tam Yağlı Yoğurt",miktar:3,birim:"adet",birimFiyat:15.30,kdvOrani:8,toplam:45.90},{ad:"Peynir 200g",miktar:1,birim:"adet",birimFiyat:12.50,kdvOrani:18,toplam:12.50}]

Few-shot MOD_B (vgsepetim):
Girdi:
1 | ST001 | Sütaş Tam Yağlı Süt 1 lt | 12 | 8,50 | 102,00
2 | ST002 | İçim Tereyağı 250 g | 6 | 45,00 | 270,00
Çıktı: [{ad:"Tam Yağlı Süt",miktar:12,birim:"l",adet:12,pkgAmt:1,birimFiyat:8.50,toplam:102.00},{ad:"Tereyağı",miktar:6,birim:"g",adet:6,pkgAmt:250,birimFiyat:45.00,toplam:270.00}]`;

// ═══ EDIT FORM ═══
// Ağırlık parse helper - "500 g", "1 kg", "2 adet" gibi metinleri gram'a çevirir
// Marka-agnostik isim normalize: "Sütaş Süt" → "Süt", "İçim Beyaz Peynir" → "Beyaz Peynir"
const BRAND_LIST=[
  "sütaş","icim","içim","pınar","eker","torku","ülker","ulker","knorr","bim","a101","şok","sok",
  "yayla","migros","carrefour","cw","başak","basak","dimes","cappy","tamek","aytaç","aytac",
  "koska","namet","banvit","beypiliç","beypilic","şenpiliç","senpilic","kent","eti","duru","filiz",
  "golf","bolca","tikveşli","tikvesli","danem","danone","lipton","nestle","nescafe","jacobs",
  "tat","cefis","kozmo","hayat","selpak","colgate","ipana","fa","nivea","dove","palmolive",
  "elidor","head&shoulders","pantene","dalin","organik","doğal","dogal","simply","fresh",
  "classic","premium","gold","light","lite","zero","ekstra","extra","özel","ozel","seçkin","seckin",
  "taze","fresh","natural","bio","organik","yayla","çamlıca","camlica","çağrı","cagri"
];
// Sayı input helper: 0 sabit kalmaz, boş silinebilir
const numVal=(v)=>v===0||v===null||v===undefined?"":String(v);
const numParse=(s,isFloat=false,min=null,fallback=0)=>{
  if(s===""||s===null||s===undefined)return fallback;
  const n=isFloat?parseFloat(s):parseInt(s,10);
  if(isNaN(n))return fallback;
  return min!==null?Math.max(min,n):n;
};

const normalizeName=(name)=>{
  if(!name)return "";
  let s=name.trim();
  // Paketleme ağırlık/hacim bilgisini sona kalacak şekilde bul
  const sizeMatch=s.match(/\b\d+(\.|,)?\d*\s*(g|gr|gram|kg|ml|lt|l|cc|adet|pcs|pack)\b/i);
  const size=sizeMatch?sizeMatch[0]:"";
  if(size)s=s.replace(sizeMatch[0],"").trim();
  // Yüzde/kalibre örneği: %3.5, %35
  s=s.replace(/%\d+(\.\d+)?/g,"").trim();
  // Markaları çıkar - sadece ismin başında/sonunda geçenleri
  const words=s.split(/\s+/);
  const filtered=words.filter((w,i)=>{
    const wl=w.toLowerCase().replace(/[.,'"]/g,"");
    // Marka listesinde varsa çıkar
    return !BRAND_LIST.includes(wl);
  });
  let cleaned=filtered.join(" ").trim();
  // Küçük temizlik
  cleaned=cleaned.replace(/\s+/g," ").trim();
  // Baş harfleri büyük (Türkçe uyumlu)
  cleaned=cleaned.split(" ").map(w=>w?w[0].toLocaleUpperCase("tr")+w.slice(1).toLocaleLowerCase("tr"):"").join(" ");
  return cleaned||name;
};

// ═══ CORRECTION LEARNING: Kalori DB ═══
// Malzeme isminden kalori tahminini DB'den alır veya AI'a sorar, sonucu DB'ye kaydeder
const getCalorieFromDB=(calorieDB,ingredientName)=>{
  if(!ingredientName||!calorieDB)return null;
  const key=normalizeName(ingredientName).toLowerCase().trim();
  if(calorieDB[key])return calorieDB[key];
  // Fuzzy lookup: sık geçen kelimeler
  for(const dbKey of Object.keys(calorieDB)){
    if(dbKey.length>3&&(key.includes(dbKey)||dbKey.includes(key))){
      return calorieDB[dbKey];
    }
  }
  return null;
};
const saveCalorieToDB=(setCalorieDB,ingredientName,kcalPer100g)=>{
  if(!ingredientName||!kcalPer100g)return;
  const key=normalizeName(ingredientName).toLowerCase().trim();
  setCalorieDB(p=>({...p,[key]:{kcal:kcalPer100g,learnedAt:new Date().toISOString()}}));
};
// Reçete için toplam kalori hesabı (kcal/100g) - DB'de olanlar için AI'a gitmeden
// Eksik malzemeleri döner, kullanıcı istediğinde AI ile tamamlanır
const estimateRecipeCalories=(recipe,calorieDB)=>{
  if(!recipe||!recipe.ingredients||recipe.ingredients.length===0)return {kcal:null,missing:[],known:[]};
  let totalKcal=0,totalWeight=0;
  const missing=[],known=[];
  for(const ing of recipe.ingredients){
    const weight=parseAmountToGram(ing.amount);
    if(weight<=0){missing.push(ing);continue;}
    const entry=getCalorieFromDB(calorieDB,ing.name);
    if(entry){
      totalKcal+=(entry.kcal*weight)/100;
      totalWeight+=weight;
      known.push({...ing,kcal:entry.kcal,weight});
    }else{
      missing.push(ing);
    }
  }
  if(totalWeight===0)return {kcal:null,missing,known};
  // 100g başı kcal
  const avgKcal=Math.round((totalKcal/totalWeight)*100);
  return {kcal:avgKcal,totalKcal:Math.round(totalKcal),totalWeight:Math.round(totalWeight),missing,known};
};
// Correction learning: OCR düzeltmelerini sakla
// Kullanıcı "Sütaş Süt" yerine "Süt" yazarsa → bir daha "Sütaş Süt" için "Süt" dönülür
const getCorrectionFromDB=(calorieDB,originalName)=>{
  if(!calorieDB._corrections)return null;
  const key=originalName.toLowerCase().trim();
  return calorieDB._corrections[key]||null;
};
const saveCorrectionToDB=(setCalorieDB,originalName,correctedName)=>{
  if(!originalName||!correctedName)return;
  const key=originalName.toLowerCase().trim();
  const correctedKey=correctedName.trim();
  if(key===correctedKey.toLowerCase())return; // aynıysa kaydetme
  setCalorieDB(p=>({
    ...p,
    _corrections:{
      ...(p._corrections||{}),
      [key]:{corrected:correctedKey,learnedAt:new Date().toISOString()}
    }
  }));
};

// ═══ ÖLÇÜ BİRİMİ VERİTABANI ═══
// Sıvı ölçüler: ml cinsinden
// Kuru ölçüler: gram cinsinden (malzemeye göre değişir — varsayılan yoğunluk)
const UNIT_DB={
  // Büyük hacim
  "l":1000,"litre":1000,"liter":1000,
  "dl":100,"desilitre":100,
  "cl":10,"santilitre":10,
  "ml":1,"mililitre":1,
  // Ağırlık
  "kg":1000,"kilogram":1000,"kilo":1000,
  "g":1,"gr":1,"gram":1,
  "mg":0.001,"miligram":0.001,
  // Pound / Ounce
  "lb":453.592,"pound":453.592,"libre":453.592,
  "oz":28.3495,"ounce":28.3495,
  // Bardak (cup) — su bazlı 240ml
  "cup":240,"cups":240,"c":240,
  "su bardağı":200,"su bard":200,
  "çay bardağı":100,"çay bard":100,
  "kahve fincanı":60,"fincan":60,
  // Kaşık
  "tbsp":15,"tablespoon":15,"yemek kaşığı":15,"yemek kaş":15,"yemek k":15,
  "tsp":5,"teaspoon":5,"tatlı kaşığı":5,"tatlı kaş":5,"tatlı k":5,
  "çay kaşığı":2,"çay kaş":2,"çay k":2,
  // Diğer
  "pinch":0.5,"tutam":0.5,
  "adet":0,"piece":0,"pcs":0,"pc":0,
  "demet":0,"bunch":0,
  "dilim":0,"slice":0,
  "diş":0,"clove":0,
  "yaprak":0,"leaf":0,
  "dal":0,"sprig":0
};

const parseAmountToGram=(amount)=>{
  if(!amount)return 0;
  const s=String(amount).toLowerCase().trim();
  const numMatch=s.match(/[\d.,]+/);
  if(!numMatch)return 0;
  const num=parseFloat(numMatch[0].replace(",","."));
  if(isNaN(num))return 0;
  // Uzun birim adından kısa olana doğru sırala (önce uzunu yakala)
  const sortedUnits=Object.keys(UNIT_DB).sort((a,b)=>b.length-a.length);
  for(const unit of sortedUnits){
    if(s.includes(unit)){
      const factor=UNIT_DB[unit];
      return factor===0?0:num*factor;
    }
  }
  // Sadece sayı varsa gram varsay
  return num;
};

// Birim dönüştürücü: gram → istenen birim
const gramToUnit=(gram,unit)=>{
  const u=unit.toLowerCase();
  const factor=UNIT_DB[u];
  if(!factor||factor===0)return gram;
  return Math.round((gram/factor)*100)/100;
};

// Mevcut reçete malzemelerini yeni birime çevir
const convertRecipeUnit=(recipe,fromUnit,toUnit)=>{
  if(!recipe||!recipe.ingredients)return recipe;
  const fromFactor=UNIT_DB[fromUnit.toLowerCase()]||1;
  const toFactor=UNIT_DB[toUnit.toLowerCase()]||1;
  if(fromFactor===toFactor)return recipe;
  return{
    ...recipe,
    ingredients:recipe.ingredients.map(ing=>{
      const numMatch=String(ing.amount).match(/[\d.,]+/);
      if(!numMatch)return ing;
      const num=parseFloat(numMatch[0].replace(",","."));
      const grams=num*fromFactor;
      const newNum=Math.round((grams/toFactor)*100)/100;
      return{...ing,amount:`${newNum} ${toUnit}`};
    })
  };
};
const EditForm=({init,onSave,onCancel,t})=>{
  const[f,sF]=useState(init||{name:"",mainCat:"pastry",subCat:"milk",venue:"alacarte",prep:"cooked",cuisine:"other",difficulty:"medium",portionG:100,calories:null,totalWeight:0,ingredients:[{name:"",amount:""}],allergens:[],diets:[],notes:"",photo:null,cookMethod:"",cookTemp:"",cookTime:""});
  const u=(k,v)=>sF(p=>({...p,[k]:v}));
  const subs=SUB_CATS[f.mainCat]||[];

  // Otomatik toplam ağırlık hesabı - malzeme değişince
  const autoWeight=(f.ingredients||[]).reduce((sum,ing)=>sum+parseAmountToGram(ing.amount),0);
  // Kullanıcı elle değiştirmediyse otomatik hesabı uygula
  useEffect(()=>{
    if(autoWeight>0&&(!f.totalWeight||f.totalWeight===0||f._autoWeight)){
      sF(p=>({...p,totalWeight:Math.round(autoWeight),_autoWeight:true}));
    }
  // eslint-disable-next-line
  },[autoWeight]);

  // Canlı porsiyon hesabı
  const portionG=f.portionG||100;
  const portionCount=portionG>0&&f.totalWeight>0?Math.floor(f.totalWeight/portionG):0;

  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div><label style={lSt(t)}>{t.L.recipeName}</label><input style={iSt(t)} value={f.name} onChange={e=>u("name",e.target.value)}/></div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <div><label style={lSt(t)}>{t.L.risk||"Difficulty"}</label><select style={iSt(t)} value={f.difficulty} onChange={e=>u("difficulty",e.target.value)}>{DIFFS.filter(v=>v.id!=="all").map(v=><option key={v.id} value={v.id}>{diffL(v,t.lang)}</option>)}</select></div>
      <div><label style={lSt(t)}>{t.L.cookMethod||"Method"}</label><select style={iSt(t)} value={f.prep} onChange={e=>u("prep",e.target.value)}>{PREPS.filter(v=>v.id!=="all").map(v=><option key={v.id} value={v.id}>{v.l}</option>)}</select></div>
    </div>
    {/* Porsiyon Gramajı + Canlı porsiyon adedi gösterimi */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <div>
        <label style={lSt(t)}>{t.L.portionG} (g)</label>
        <input style={iSt(t)} type="number" value={numVal(f.portionG)} onChange={e=>u("portionG",numParse(e.target.value,false,1,100))} placeholder="100"/>
      </div>
      <div>
        <label style={lSt(t)}>{t.L.totalWeight} (g)</label>
        <input style={iSt(t)} type="number" placeholder={autoWeight>0?String(Math.round(autoWeight)):"AI"} value={numVal(f.totalWeight)} onChange={e=>sF(p=>({...p,totalWeight:numParse(e.target.value,false,0,0),_autoWeight:false}))}/>
      </div>
    </div>
    {portionCount>0&&<div style={{background:t.acB,border:`1px solid ${t.acBo}`,borderRadius:10,padding:"8px 12px",fontSize:12,color:t.accent,fontWeight:600,textAlign:"center"}}>
      ≈ {portionCount} {t.L.portions} ({f.totalWeight}g ÷ {portionG}g)
    </div>}
    <div>
      <label style={lSt(t)}>{t.L.calories} (kcal/100g)</label>
      <input style={iSt(t)} type="number" placeholder="320" value={numVal(f.calories)} onChange={e=>u("calories",e.target.value===""?null:numParse(e.target.value,false,0,null))}/>
    </div>
    <div style={{background:t.waBg,border:`1px solid ${t.waBo}`,borderRadius:12,padding:"12px 14px"}}>
      <div style={{fontSize:11,fontWeight:700,color:t.warn,letterSpacing:"0.1em",marginBottom:8}}>{t.L.cookMethod.toUpperCase()} / {t.L.storage.toUpperCase()}</div>
      <div><label style={lSt(t)}>{t.L.cookMethod}</label><select style={iSt(t)} value={f.cookMethod||""} onChange={e=>u("cookMethod",e.target.value)}>
        <option value="">— {t.L.select||"Seç"} —</option>
        <option value="fırın">🔥 {t.L.methodOven||"Fırın"}</option>
        <option value="kaynat">♨️ {t.L.methodBoil||"Kaynat"}</option>
        <option value="kızart">🍳 {t.L.methodFry||"Kızart"}</option>
        <option value="ızgara">🔥 {t.L.methodGrill||"Izgara"}</option>
        <option value="benmari">🫕 {t.L.methodBainMarie||"Benmari"}</option>
        <option value="buzdolabı">❄️ {t.L.methodFridge||"Buzdolabı"}</option>
        <option value="dondurma">🧊 {t.L.methodFreeze||"Dondurma"}</option>
        <option value="soğutma">💧 {t.L.methodCool||"Soğutma"}</option>
        <option value="pişirme yok">— {t.L.methodNone||"Pişirme yok"}</option>
      </select></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:10}}>
        <div><label style={lSt(t)}>{t.L.cookTemp}</label><input style={iSt(t)} placeholder="180°C" value={f.cookTemp||""} onChange={e=>u("cookTemp",e.target.value)}/></div>
        <div><label style={lSt(t)}>{t.L.cookTime}</label><input style={iSt(t)} placeholder="35 min" value={f.cookTime||""} onChange={e=>u("cookTime",e.target.value)}/></div>
      </div>
    </div>
    <div><label style={lSt(t)}>{t.L.detailIngredients||t.L.ingredients}</label>
      {f.ingredients.map((ing,i)=>{
        // amount'tan sayı ve birimi ayır
        const amtMatch=String(ing.amount||"").match(/^([\d.,]+)\s*(.*)$/);
        const amtNum=amtMatch?amtMatch[1]:"";
        const amtUnit=amtMatch?amtMatch[2].trim():"g";
        return <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 60px 70px 28px",gap:4,marginBottom:6,alignItems:"center"}}>
          <input style={iSt(t)} placeholder={t.L.name} value={ing.name} onChange={e=>{const a=[...f.ingredients];a[i]={...a[i],name:e.target.value};u("ingredients",a)}}/>
          <input style={{...iSt(t),textAlign:"right"}} placeholder="0" value={amtNum} onChange={e=>{const a=[...f.ingredients];a[i]={...a[i],amount:(e.target.value+" "+amtUnit).trim()};u("ingredients",a)}}/>
          <select style={{...iSt(t),padding:"8px 4px",fontSize:12}} value={amtUnit||"g"} onChange={e=>{const a=[...f.ingredients];a[i]={...a[i],amount:(amtNum+" "+e.target.value).trim()};u("ingredients",a)}}>
            <optgroup label="Ağırlık / Weight">
              <option value="g">g</option>
              <option value="kg">kg</option>
              <option value="mg">mg</option>
              <option value="oz">oz</option>
              <option value="lb">lb</option>
            </optgroup>
            <optgroup label="Hacim / Volume">
              <option value="ml">ml</option>
              <option value="cl">cl</option>
              <option value="dl">dl</option>
              <option value="l">L</option>
              <option value="fl oz">fl oz</option>
            </optgroup>
            <optgroup label="Ölçü / Measure">
              <option value="cup">cup</option>
              <option value="tbsp">tbsp</option>
              <option value="tsp">tsp</option>
            </optgroup>
            <optgroup label="Diğer / Other">
              <option value="adet">adet</option>
              <option value="piece">pc</option>
              <option value="pinch">pinch</option>
              <option value="dilim">dilim</option>
              <option value="diş">diş</option>
              <option value="demet">demet</option>
            </optgroup>
          </select>
          <button onClick={()=>u("ingredients",f.ingredients.filter((_,j)=>j!==i))} style={{background:"none",border:`1px solid ${t.danger}`,borderRadius:8,color:t.danger,cursor:"pointer",padding:"6px 8px",fontSize:14,flexShrink:0}}>✕</button>
        </div>;
      })}
      <button onClick={()=>u("ingredients",[...f.ingredients,{name:"",amount:"100 g"}])} style={{...bSt("s",t),fontSize:13}}>{t.L.recipeAddIngredient}</button>
    </div>
    <div><label style={lSt(t)}>{t.L.allergens}</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {ALLERGENS.map(a=>{const on=f.allergens.includes(a.id);return <button key={a.id} onClick={()=>u("allergens",on?f.allergens.filter(x=>x!==a.id):[...f.allergens,a.id])} style={{padding:"5px 10px",borderRadius:8,fontSize:12,fontWeight:600,border:`1px solid ${a.c}44`,background:on?a.c+"20":t.inBg,color:on?a.c:t.tm,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4}}>{a.icon} {allergenL(a,t?.lang||"tr")}</button>})}
      </div>
    </div>
    <div><label style={lSt(t)}>{t.lang==="tr"?"Diyet Etiketleri":t.lang==="en"?"Diet Labels":t.lang==="ru"?"Диетические метки":t.lang==="es"?"Etiquetas de dieta":t.lang==="de"?"Diät-Labels":t.lang==="fr"?"Étiquettes régime":t.lang==="zh"?"饮食标签":"علامات النظام الغذائي"}</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {DIETS.map(d=>{const on=(f.diets||[]).includes(d.id);return <button key={d.id} onClick={()=>u("diets",on?(f.diets||[]).filter(x=>x!==d.id):[...(f.diets||[]),d.id])} style={{padding:"5px 10px",borderRadius:8,fontSize:12,fontWeight:600,border:`1px solid ${on?t.success:t.inBo}`,background:on?t.sucBg:t.inBg,color:on?t.success:t.tm,cursor:"pointer"}}>{d.icon} {dietL(d,t?.lang||"tr")}</button>})}
      </div>
    </div>
    <div><label style={lSt(t)}>{t.L.note}</label><textarea style={{...iSt(t),minHeight:70,resize:"vertical"}} value={f.notes} onChange={e=>u("notes",e.target.value)}/></div>
    <div style={{display:"flex",gap:10}}>
      <button onClick={onCancel} style={{...bSt("s",t),flex:1}}>{t.L.cancel}</button>
      <button onClick={()=>onSave(f)} style={{...bSt("p",t),flex:2}}>{t.L.save}</button>
    </div>
  </div>;
};

// ═══ ADD MODAL ═══
const AddModal=({onClose,onAdd,apiKey,t})=>{
  const[mode,setMode]=useState("photo");
  const[imgB,setIB]=useState(null);const[imgM,setIM]=useState("image/jpeg");
  const[txt,setTxt]=useState("");
  const[loading,setL]=useState(false);
  const[err,setE]=useState("");
  const[parsed,setP]=useState(null);

  const analyze=async()=>{
    setL(true);setE("");
    try{
      let raw;
      if(mode==="photo"){
        if(!imgB){setE("Önce fotoğraf seçin");setL(false);return}
        raw=await callAI(apiKey,P_OCR_RECIPE,[{type:"image",source:{type:"base64",media_type:imgM,data:imgB}},{type:"text",text:"Bu reçeteyi analiz et. SADECE JSON döndür."}]);
      }else{
        if(!txt.trim()){setE(t.L.errNoText||"Metin girin");setL(false);return}
        raw=await callAI(apiKey,P_OCR_RECIPE,txt);
      }
      const r=parseJSON(raw);
      if(!r.name)throw new Error("Geçersiz yanıt");
      setP(r);
    }catch(e){setE("Hata: "+e.message.slice(0,100))}
    setL(false);
  };

  return <div style={mOv(t)} onClick={onClose}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),maxWidth:480,width:"calc(100% - 32px)"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
      <h3 style={{fontSize:22,color:t.text}}>{t.L.newRecipe}</h3>
      <button onClick={onClose} style={{...bSt("s",t),padding:"6px 14px",fontSize:18}}>×</button>
    </div>
    {!parsed?<div>
      <div style={{display:"flex",gap:8,marginBottom:18}}>
        {[{id:"photo",l:"📷 "+t.L.addPhoto.replace("📷 ",""),icon:"photo"},{id:"text",l:"📝 "+t.L.note,icon:"text"},{id:"manual",l:"✏ "+t.L.manualAdd.replace("+ ",""),icon:"manual"}].map(m=><button key={m.id} onClick={()=>{setMode(m.id);setE("");if(m.id==="manual")setP({name:"",mainCat:"pastry",subCat:"milk",venue:"alacarte",prep:"cooked",cuisine:"other",difficulty:"medium",portionG:100,totalWeight:0,calories:null,ingredients:[{name:"",amount:""}],allergens:[],diets:[],notes:""})}} style={{...bSt(mode===m.id?"p":"s",t),flex:1,fontSize:13}}>{m.l}</button>)}
      </div>
      {mode==="photo"?<PhotoPick onImg={(b,m)=>{setIB(b);setIM(m||"image/jpeg")}} t={t}/>:
       mode==="text"?<textarea style={{...iSt(t),minHeight:140,resize:"vertical"}} value={txt} onChange={e=>setTxt(e.target.value)} placeholder={t.L.recipeName+"..."}/>:null}
      {err&&<div style={{marginTop:12,color:t.danger,fontSize:13,background:t.danBg,border:`1px solid ${t.danBo}`,borderRadius:10,padding:"10px 14px"}}>{err}</div>}
      {mode!=="manual"&&<button onClick={analyze} disabled={loading} style={{...bSt("p",t),width:"100%",marginTop:14,padding:14,opacity:loading?0.7:1}}>{loading?"⏳ "+t.L.recipeAnalyzing:"✦ "+t.L.invoiceAnalyze}</button>}
    </div>:<div>
      <div style={{background:t.sucBg,border:`1px solid ${t.sucBo}`,borderRadius:12,padding:"10px 14px",marginBottom:16,fontSize:13,color:t.success}}>✓ {t.L.recipeEdit}</div>
      <EditForm init={parsed} onSave={f=>{onAdd({...f,id:Date.now(),created:new Date().toISOString().split("T")[0],photo:null});onClose()}} onCancel={()=>setP(null)} t={t}/>
    </div>}
  </div></div>;
};

// ═══ DETAIL ═══
const Detail=({r,onClose,onDel,onEdit,onDeduct,stock,expenses,storageAreas,productions,setProductions,profile,organizations,lots,trackedIngs,traceability,team,onShareToChat,t})=>{
  const[mult,setMult]=useState(1);
  const[portionG,setPortionG]=useState(r.portionG||100);
  const[ded,setDed]=useState(false);
  const[photoMode,setPM]=useState(false);
  const[showShareModal,setShowShareModal]=useState(false);
  const[showProdModal,setShowProdModal]=useState(false);
  const[prodStorage,setProdStorage]=useState((storageAreas||[])[0]?.id||"");
  const[prodNote,setProdNote]=useState("");
  const[prodDays,setProdDays]=useState(0);
  const[unitSystem,setUnitSystem]=useState("metric"); // metric | imperial
  const lang=t.lang;

  // Birim sistemi dönüşüm helper
  const convertAmt=(amount)=>{
    if(unitSystem==="metric")return amount;
    const grams=parseAmountToGram(amount);
    if(grams<=0)return amount;
    if(unitSystem==="imperial"){
      if(grams>=453.6)return `${(grams/453.6).toFixed(2)} lb`;
      if(grams>=28.35)return `${(grams/28.35).toFixed(1)} oz`;
      return `${grams}g`;
    }
    if(unitSystem==="turkish"){
      if(grams>=1000)return `${(grams/1000).toFixed(2)} kg`;
      // Su bardağı = 200ml≈200g, yemek kaşığı=15g, çay kaşığı=5g
      if(grams%200===0&&grams<=1000)return `${grams/200} su bardağı`;
      if(grams>=15&&grams%15===0)return `${grams/15} yemek kaşığı`;
      return `${grams}g`;
    }
    return amount;
  };

  // Numune şahit etiketi
  const[showSampleModal,setShowSampleModal]=useState(false);
  const nowStr=()=>{const d=new Date();return `${String(d.getFullYear())}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}T${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`};
  const[sampleLocation,setSampleLocation]=useState("");
  const[sampleOrganization,setSampleOrganization]=useState("");
  const[samplePax,setSamplePax]=useState(1);
  const[sampleDateTime,setSampleDateTime]=useState(nowStr());
  const[sampleTemp,setSampleTemp]=useState("");
  const[sampleHours,setSampleHours]=useState(72);

  // Numune etiketi yazdır + rapora kaydet
  const printSampleLabel=()=>{
    if(!sampleLocation.trim()&&!sampleOrganization.trim()){window.toast.info("Organizasyon veya yer girin");return}
    const d=new Date(sampleDateTime);
    const fmtD=`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getFullYear()).slice(-2)}`;
    const fmtT=`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
    const endDate=new Date(d.getTime()+sampleHours*60*60*1000);
    const fmtED=`${String(endDate.getDate()).padStart(2,"0")}/${String(endDate.getMonth()+1).padStart(2,"0")}/${String(endDate.getFullYear()).slice(-2)} ${String(endDate.getHours()).padStart(2,"0")}:${String(endDate.getMinutes()).padStart(2,"0")}`;
    const takerName=(profile&&profile.fullName)?profile.fullName:"_______________";
    const orgLine=sampleOrganization?(sampleLocation?`${sampleOrganization} — ${sampleLocation}`:sampleOrganization):sampleLocation;
    const paxLine=samplePax>1?` (${samplePax} kişi)`:"";

    // Rapor kaydı
    const sampleRecord={
      id:Date.now(),
      recipeId:r.id,
      recipeName:r.name,
      sampleLocation:orgLine,
      sampleOrganization,
      samplePax,
      sampleDateTime:d.toISOString(),
      sampleTemp,
      sampleHours,
      sampleEndAt:endDate.toISOString(),
      takerName,
      status:"sample",
      producedAt:d.toISOString(),
      reportCat:"samples",
      lotNumber:fmtD
    };
    setProductions(p=>[sampleRecord,...p]);

    // Etiket yazdır
    const w=window.open("","","width=600,height=600");
    let html=`<html><head><style>
@page{size:10cm 6cm;margin:0}
*{margin:0;box-sizing:border-box;font-family:-apple-system,Arial,sans-serif}
body{background:#fff3cd;color:#000;padding:6mm;font-size:9px;line-height:1.35;border:2px solid #dc2626;width:10cm;height:6cm}
.warn{text-align:center;background:#dc2626;color:#fff;padding:2px;font-weight:700;letter-spacing:0.1em;margin-bottom:4px;font-size:9px}
.row{margin-bottom:2px;display:flex;gap:4px;align-items:baseline}
.lbl{font-weight:600;flex-shrink:0;min-width:80px;font-size:8.5px}
.val{border-bottom:1px dotted #000;flex:1;padding-left:3px;font-size:9px}
@media print{body{margin:0;padding:6mm}}
</style></head><body>
<div class="warn">⚠ NUMUNE ŞAHİT — TÜKETİLMEZ ⚠</div>
<div class="row"><span class="lbl">Numune:</span><span class="val">${r.name}${paxLine}</span></div>
<div class="row"><span class="lbl">Organizasyon:</span><span class="val">${orgLine}</span></div>
<div class="row"><span class="lbl">Alan Kişi:</span><span class="val">${takerName}</span></div>
<div class="row"><span class="lbl">Tarih / Saat:</span><span class="val">${fmtD} ${fmtT}</span></div>
<div class="row"><span class="lbl">Saklama:</span><span class="val">${sampleHours}s (bitiş: ${fmtED})</span></div>
<div class="row"><span class="lbl">Sıcaklık:</span><span class="val">${sampleTemp||"___"} °C</span></div>
</body></html>`;
    w.document.write(html);w.document.close();
    setTimeout(()=>w.print(),500);
    setShowSampleModal(false);
    setSampleLocation("");
    setSampleOrganization("");
    setSamplePax(1);
    setSampleTemp("");
  };

  // Risk sınıfı tespiti - keyword bazlı
  const detectRisk=()=>{
    const allText=r.ingredients.map(i=>i.name.toLowerCase()).join(" ");
    const highRisk=["yumurta","krema","süt","et","tavuk","balık","karides","kıyma","mayonez","kuzu","dana"];
    const lowRisk=["un","şeker","bal","kakao","tuz","baharat","pirinç","bulgur"];
    if(highRisk.some(k=>allText.includes(k)))return"high";
    if(lowRisk.every(k=>allText.includes(k)||true)&&!highRisk.some(k=>allText.includes(k)))return"low";
    return"medium";
  };

  // Tüketim tarihi hesabı - depo tipi + risk'e göre
  const calcExpiryDays=(storageType,risk)=>{
    if(storageType==="freezer")return risk==="high"?60:90; // 2-3 ay
    if(storageType==="fridge")return risk==="high"?1:risk==="medium"?3:5;
    if(storageType==="dry")return risk==="high"?1:risk==="medium"?7:14;
    if(storageType==="hot")return 0; // 4 saat (0 gün gösterilir)
    if(storageType==="room")return risk==="high"?0:risk==="medium"?2:7;
    return 3;
  };

  // Modal açılınca tahmini tarihi hesapla
  useEffect(()=>{
    if(showProdModal&&prodStorage){
      const s=(storageAreas||[]).find(x=>x.id===prodStorage);
      if(s){
        const risk=detectRisk();
        const days=calcExpiryDays(s.type,risk);
        setProdDays(days);
      }
    }
  },[showProdModal,prodStorage]);

  // Üretim kaydı oluştur
  const saveProduction=()=>{
    const s=(storageAreas||[]).find(x=>x.id===prodStorage);
    if(!s){window.toast.info(t.L.storage+"?");return}
    const typ=STORAGE_TYPES.find(x=>x.id===s.type);
    const now=new Date();
    const expires=new Date(now.getTime()+prodDays*24*60*60*1000);
    const tw=(r.totalWeight||0)*mult;
    const portions=portionG>0&&tw>0?Math.round(tw/portionG):Math.round((r.servings||1)*mult);
    const lotNumber=`${String(now.getDate()).padStart(2,"0")}/${String(now.getMonth()+1).padStart(2,"0")}/${String(now.getFullYear()).slice(-2)}`;

    // Hammadde lot izi — TÜM malzemeler (takip edilenlerin parti no'su var, diğerleri boş)
    const ingredientLots=[];
    if(r.ingredients){
      for(const ing of r.ingredients){
        const ingLower=ing.name.toLowerCase();
        // trackedIngs içinde eşleşen bul
        let matched=null;
        let lotNo=null;
        if(traceability&&lots){
          matched=(trackedIngs||[]).find(ti=>{
            const tLower=ti.toLowerCase();
            return ingLower.includes(tLower)||tLower.includes(ingLower);
          });
          if(matched)lotNo=lots[matched]?.no||null;
        }
        ingredientLots.push({
          name:ing.name,
          amount:ing.amount,
          tracked:matched||ing.name,
          lotNo:lotNo,
          lotUpdatedAt:matched?lots[matched]?.updatedAt:null
        });
      }
    }

    const newProd={
      id:Date.now(),
      recipeId:r.id,
      recipeName:r.name,
      portions:portions,
      portionG:portionG,
      totalWeight:tw,
      multiplier:mult,
      producedAt:now.toISOString(),
      expiresAt:expires.toISOString(),
      storageId:s.id,
      storageName:s.name,
      storageType:s.type,
      storageTemp:s.temp,
      storageIcon:typ?.icon||"📦",
      lotNumber:lotNumber,
      note:prodNote,
      allergens:r.allergens||[],
      risk:detectRisk(),
      status:"active",
      producedBy:(profile&&profile.fullName)||"",
      ingredientLots:ingredientLots,
      reportCat:"production"
    };
    setProductions(p=>[newProd,...p]);
    // Hammadde düş
    onDeduct(r,mult);
    setShowProdModal(false);
    setProdNote("");
    setDed(true);
    setTimeout(()=>setDed(false),2000);
  };

  const mc=MAIN_CATS.find(c=>c.id===r.mainCat);
  const sc=(SUB_CATS[r.mainCat]||[]).find(s=>s.id===r.subCat);
  const ven=VENUES.find(v=>v.id===r.venue);

  const cost=(()=>{let tot=0;for(const i of r.ingredients){const si=stock.find(s=>i.name.toLowerCase().includes(s.name.toLowerCase())||s.name.toLowerCase().includes(i.name.toLowerCase().split(" ")[0]));if(!si)continue;const nm=i.amount.match(/[\d.,]+/);if(!nm)continue;tot+=parseFloat(nm[0].replace(",","."))*si.ppu*mult}return tot})();
  const scale=(a)=>{if(mult===1)return a;const nm=a.match(/[\d.,]+/);if(!nm)return a;const n=parseFloat(nm[0].replace(",","."));return a.replace(nm[0],(n*mult).toFixed(n%1===0?0:1))};

  return (
    <>
    {showSampleModal&&<div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,zIndex:9999}} onClick={()=>setShowSampleModal(false)}><div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:18,padding:22,maxWidth:440,width:"100%",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.4)"}}>
      <h3 style={{fontSize:20,color:t.text,marginBottom:6}}>{t.L.sampleTitle}</h3>
      <div style={{fontSize:11,color:t.tm,marginBottom:14}}>{t.L.sampleDesc}</div>
      <div style={{background:t.inBg,borderRadius:12,padding:"10px 12px",marginBottom:12,fontSize:12,color:t.ts}}>
        <div><strong>{t.L.sampleName}:</strong> {r.name}</div>
        <div><strong>{t.L.sampleTaker}:</strong> {(profile&&profile.fullName)||<span style={{color:t.warn}}>{t.L.sampleProfileWarn}</span>}</div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10,marginBottom:10}}>
        <div>
          <label style={lSt(t)}>{t.L.sampleOrgLabel}</label>
          <select style={iSt(t)} value={sampleOrganization} onChange={e=>setSampleOrganization(e.target.value)}>
            <option value="">{t.L.sampleOrgSelect}</option>
            {(organizations||[]).map(o=><option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div>
          <label style={lSt(t)}>{t.L.samplePax}</label>
          <input type="number" min="1" style={iSt(t)} value={samplePax} onChange={e=>setSamplePax(Math.max(1,parseInt(e.target.value,10)||1))}/>
        </div>
      </div>
      <div style={{marginBottom:10}}>
        <label style={lSt(t)}>{t.L.sampleLocation}</label>
        <input style={iSt(t)} value={sampleLocation} onChange={e=>setSampleLocation(e.target.value)} placeholder={t.L.sampleLocationHint}/>
      </div>
      <div style={{marginBottom:10}}>
        <label style={lSt(t)}>{t.L.sampleDateTime}</label>
        <input type="datetime-local" style={iSt(t)} value={sampleDateTime} onChange={e=>setSampleDateTime(e.target.value)}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
        <div>
          <label style={lSt(t)}>{t.L.sampleTemp}</label>
          <input type="number" style={iSt(t)} value={sampleTemp} onChange={e=>setSampleTemp(e.target.value)} placeholder={t.L.sampleTempHint}/>
        </div>
        <div>
          <label style={lSt(t)}>{t.L.sampleHours}</label>
          <input type="number" min="1" style={iSt(t)} value={sampleHours} onChange={e=>setSampleHours(Math.max(1,parseInt(e.target.value,10)||72))}/>
        </div>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>setShowSampleModal(false)} style={{...bSt("s",t),flex:1}}>{t.L.cancel}</button>
        <button onClick={printSampleLabel} style={{...bSt("p",t),flex:2}}>{t.L.samplePrintSave}</button>
      </div>
    </div></div>}
    {showProdModal&&<div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,zIndex:9999}} onClick={()=>setShowProdModal(false)}><div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:18,padding:22,maxWidth:440,width:"100%",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.4)"}}>
      <h3 style={{fontSize:20,color:t.text,marginBottom:14}}>{t.L.productionRecord}</h3>
      <div style={{background:t.inBg,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
        <div style={{fontSize:14,fontWeight:700,color:t.text}}>{r.name} × {mult}</div>
        <div style={{fontSize:12,color:t.tm,marginTop:4}}>
          {(()=>{const tw=(r.totalWeight||0)*mult;const portions=portionG>0&&tw>0?Math.round(tw/portionG):Math.round((r.servings||1)*mult);return `${portions} porsiyon (${portionG}g) · Toplam ${tw>=1000?(tw/1000).toFixed(1)+" kg":Math.round(tw)+" g"}`})()}
        </div>
      </div>
      <div style={{marginBottom:12}}>
        <label style={lSt(t)}>{t.L.storage}</label>
        <select style={iSt(t)} value={prodStorage} onChange={e=>setProdStorage(e.target.value)}>
          {(storageAreas||[]).map(s=>{const typ=STORAGE_TYPES.find(x=>x.id===s.type);return <option key={s.id} value={s.id}>{typ?.icon} {s.name} ({s.temp}°C)</option>})}
        </select>
      </div>
      <div style={{marginBottom:12}}>
        <label style={lSt(t)}>{t.L.expiryDays}</label>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <input type="number" min="0" value={numVal(prodDays)} onChange={e=>setProdDays(e.target.value===""?0:Math.max(0,parseInt(e.target.value,10)||0))} style={{...iSt(t),flex:1,textAlign:"center",fontWeight:700}}/>
          <span style={{fontSize:12,color:t.tm}}>{t.L.daysLater}</span>
        </div>
        <div style={{fontSize:11,color:t.tm,marginTop:4}}>💡 {t.L.risk}: <span style={{fontWeight:600,color:detectRisk()==="high"?t.danger:detectRisk()==="medium"?t.warn:t.success}}>{detectRisk()==="high"?t.L.riskHigh:detectRisk()==="medium"?t.L.riskMedium:t.L.riskLow}</span> · {t.L.riskHint}</div>
      </div>
      {traceability&&(()=>{
        // Bu reçetedeki takip edilen hammaddeleri bul
        const usedTracked=[];
        (r.ingredients||[]).forEach(ing=>{
          const ingLower=ing.name.toLowerCase();
          const matched=(trackedIngs||[]).find(ti=>{
            const tLower=ti.toLowerCase();
            return ingLower.includes(tLower)||tLower.includes(ingLower);
          });
          if(matched&&!usedTracked.find(x=>x.tracked===matched))usedTracked.push({tracked:matched,ingName:ing.name});
        });
        if(usedTracked.length===0)return null;
        const missing=usedTracked.filter(u=>!lots[u.tracked]?.no);
        return <div style={{background:missing.length>0?t.waBg:t.inBg,border:missing.length>0?`1px solid ${t.waBo}`:"none",borderRadius:10,padding:"10px 12px",marginBottom:12}}>
          <div style={{fontSize:11,fontWeight:700,color:missing.length>0?t.warn:t.tm,marginBottom:6}}>{t.L.ingredientLotsTitle} ({usedTracked.length})</div>
          {usedTracked.map(u=>{
            const l=lots[u.tracked];
            return <div key={u.tracked} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,padding:"3px 0",color:t.ts}}>
              <span>{u.tracked}</span>
              {l?.no?<span style={{color:t.success,fontWeight:600}}>✓ {l.no}</span>:<span style={{color:t.danger,fontWeight:600}}>⚠ Eksik</span>}
            </div>;
          })}
          {missing.length>0&&<div style={{fontSize:10,color:t.warn,marginTop:6,lineHeight:1.4}}>{t.L.lotMissingHint}</div>}
        </div>;
      })()}
      <div style={{marginBottom:14}}>
        <label style={lSt(t)}>{t.L.note} ({t.L.optional})</label>
        <input style={iSt(t)} value={prodNote} onChange={e=>setProdNote(e.target.value)} placeholder={t.L.noteHint}/>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>setShowProdModal(false)} style={{...bSt("s",t),flex:1}}>{t.L.cancel}</button>
        <button onClick={saveProduction} style={{...bSt("p",t),flex:2}}>{t.L.produceSaveBtn}</button>
      </div>
    </div></div>}
    <div style={mOv(t)} onClick={onClose}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),padding:0}}>
      {r.photo&&<div style={{width:"100%",height:180,borderRadius:"18px 18px 0 0",overflow:"hidden"}}><img src={r.photo} style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>}
      <div style={{padding:"22px 26px 18px",borderBottom:`1px solid ${t.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,color:t.tm,marginBottom:6}}>{mc?.icon} {mainCatL(mc,t?.lang||"tr")} › {sc?.label} · {ven?.l}</div>
            <h2 style={{fontSize:26,color:t.text,marginBottom:10}}>{r.name}</h2>
            {r.allergens.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:6}}>{r.allergens.map(a=><ABadge key={a} a={a} t={t}/>)}</div>}
            {(r.diets||[]).length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:6}}>{(r.diets||[]).map(d=><DBadge key={d} d={d} t={t}/>)}</div>}
            {/* Kalori gösterimi — 100g baz */}
            {r.calories&&<div style={{display:"inline-flex",alignItems:"center",gap:5,background:t.waBg,border:`1px solid ${t.waBo}`,borderRadius:8,padding:"4px 10px",fontSize:12,color:t.warn,fontWeight:700}}>🔥 {r.calories} kcal/100g{portionG!==100&&<span style={{opacity:0.7,fontWeight:400}}> · {Math.round(r.calories*portionG/100)} kcal/{portionG}g</span>}</div>}
          </div>
          <div style={{display:"flex",gap:6,marginLeft:10,flexShrink:0}}>
            
            <button onClick={onClose} style={{...bSt("s",t),padding:"6px 12px",fontSize:18}}>×</button>
          </div>
        </div>
      </div>
      <div style={{padding:"18px 26px 26px"}}>
        {/* ÇARPAN — sadece slider + manuel */}
        <div style={{marginBottom:14}}><label style={lSt(t)}>{t.L.detailMultiplier}</label>
          <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
            <button onClick={()=>setMult(Math.max(0.1,Math.round((mult-0.5)*10)/10))} style={{padding:"8px 14px",borderRadius:10,fontSize:16,fontWeight:700,border:`1px solid ${t.inBo}`,background:t.pBg,color:t.text,cursor:"pointer"}}>−</button>
            <input type="number" step="0.5" min="0.1" value={numVal(mult)} onChange={e=>setMult(e.target.value===""?1:Math.max(0.1,parseFloat(e.target.value)||1))} style={{...iSt(t),flex:1,textAlign:"center",fontWeight:700,fontSize:16}}/>
            <button onClick={()=>setMult(Math.round((mult+0.5)*10)/10)} style={{padding:"8px 14px",borderRadius:10,fontSize:16,fontWeight:700,border:`1px solid ${t.inBo}`,background:t.pBg,color:t.text,cursor:"pointer"}}>+</button>
          </div>
          <input type="range" min="0.5" max="100" step="0.5" value={Math.min(100,mult)} onChange={e=>setMult(parseFloat(e.target.value))} style={{width:"100%",accentColor:t.accent}}/>
        </div>
        {/* PORSIYON GRAMAJI + BİLGİLER */}
        {(()=>{
          const tw=(r.totalWeight||0)*mult;
          const portionCount=portionG>0&&tw>0?(tw/portionG):0;
          const cal100=r.calories||0;
          const calPortion=portionG>0?(cal100*portionG/100):0;
          return <>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
              <div><label style={lSt(t)}>{t.L.detailPortion}</label><input type="number" min="10" step="10" value={numVal(portionG)} onChange={e=>setPortionG(e.target.value===""?100:Math.max(1,parseInt(e.target.value,10)||100))} style={{...iSt(t),textAlign:"center",fontWeight:600}}/></div>
              <div><label style={lSt(t)}>{t.L.detailTotalWeight}</label><div style={{...iSt(t),background:"transparent",border:"none",fontWeight:700,textAlign:"center",padding:"12px 0"}}>{tw>=1000?(tw/1000).toFixed(2).replace(".",",")+" kg":Math.round(tw)+" g"}</div></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
              <div style={{background:t.acB,borderRadius:12,padding:"10px 14px",textAlign:"center"}}>
                <div style={{fontSize:10,fontWeight:700,color:t.accent,letterSpacing:"0.1em"}}>{t.L.detailPortionCount}</div>
                <div style={{fontSize:20,fontWeight:700,color:t.text}}>{portionCount>0?portionCount.toFixed(1):"—"}</div>
              </div>
              <div style={{background:t.acB,borderRadius:12,padding:"10px 14px",textAlign:"center"}}>
                <div style={{fontSize:10,fontWeight:700,color:t.accent,letterSpacing:"0.1em"}}>{t.L.detailCalorie}</div>
                <div style={{fontSize:14,fontWeight:700,color:t.text}}>{cal100>0?cal100+" kcal/100g":"—"}</div>
                {cal100>0&&portionG!==100&&<div style={{fontSize:11,color:t.tm}}>{Math.round(calPortion)} kcal/{portionG}g</div>}
              </div>
            </div>
          </>;
        })()}
        {/* MALİYET */}
        {cost>0&&(()=>{
          const tw=(r.totalWeight||0)*mult;
          const portionCount=portionG>0&&tw>0?(tw/portionG):0;
          const portionCost=portionCount>0?(cost/portionCount):0;
          // Sabit gider + personel payı hesabı
          const exp=expenses||{fixed:[],personnel:[],monthlyPortions:1000};
          const totalFixed=(exp.fixed||[]).reduce((a,x)=>a+(x.amount||0),0);
          const totalPersonnel=(exp.personnel||[]).reduce((a,x)=>a+(x.salary||0),0);
          const totalExp=totalFixed+totalPersonnel;
          const mp=exp.monthlyPortions||1000;
          const overheadPerPortion=totalExp>0?(totalExp/mp):0;
          const overheadTotal=overheadPerPortion*(portionCount||1);
          const realCost=cost+overheadTotal;
          const realPerPortion=portionCount>0?(realCost/portionCount):realCost;
          return <div style={{background:t.waBg,border:`1px solid ${t.waBo}`,borderRadius:14,padding:"12px 16px",marginBottom:18}}>
            <div style={{fontSize:11,color:t.warn,fontWeight:700,marginBottom:6}}>{t.L.detailCostAnalysis}</div>
            <div style={{fontSize:13,color:t.warn,display:"flex",justifyContent:"space-between"}}><span>{t.L.detailRawMaterial}</span><span style={{fontWeight:600}}>₺{cost.toFixed(2)}{portionCount>0&&<span style={{opacity:0.7,fontWeight:400,fontSize:11}}> · ₺{portionCost.toFixed(2)}/por</span>}</span></div>
            {overheadPerPortion>0&&portionCount>0&&<div style={{fontSize:13,color:t.warn,display:"flex",justifyContent:"space-between"}}><span>{t.L.detailFixedShare}</span><span style={{fontWeight:600}}>₺{overheadTotal.toFixed(2)}<span style={{opacity:0.7,fontWeight:400,fontSize:11}}> · ₺{overheadPerPortion.toFixed(2)}/por</span></span></div>}
            {overheadPerPortion>0&&<div style={{fontSize:16,color:t.warn,display:"flex",justifyContent:"space-between",fontWeight:700,borderTop:`1px solid ${t.waBo}`,marginTop:6,paddingTop:6}}><span>{t.L.detailRealCost}</span><span>₺{realCost.toFixed(2)}{portionCount>0&&<span style={{fontSize:12,fontWeight:400,opacity:0.8}}> · ₺{realPerPortion.toFixed(2)}/por</span>}</span></div>}
            {overheadPerPortion===0&&portionCount>0&&<div style={{fontSize:10,color:t.tm,marginTop:4,fontStyle:"italic"}}>💡 Sabit gider ve personel giderlerini Ayarlar'a ekleyerek gerçek maliyeti görebilirsiniz.</div>}
          </div>;
        })()}
        <div style={{marginBottom:18}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <label style={lSt(t)}>{t.L.detailIngredients} {mult!==1&&<span style={{color:t.accent}}>×{mult}</span>}</label>
            <div style={{display:"flex",gap:4}}>
              {[["metric","g/ml"],["imperial","oz/lb"]].map(([sys,lbl])=>
                <button key={sys} onClick={()=>setUnitSystem(sys)} style={{
                  padding:"3px 8px",borderRadius:8,fontSize:10,fontWeight:600,border:"none",cursor:"pointer",
                  background:unitSystem===sys?t.accent:t.inBg,
                  color:unitSystem===sys?"#fff":t.tm
                }}>{lbl}</button>
              )}
            </div>
          </div>
          {r.ingredients.map((ing,i)=>{
            const scaledAmt=scale(ing.amount);
            const converted=convertAmt(scaledAmt);
            return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"10px 0",borderBottom:`1px solid ${t.border}`,fontSize:15}}>
              <span style={{color:t.ts}}>{ing.name}</span>
              <span style={{color:t.text,fontWeight:600}}>
                {converted}
                {converted!==scaledAmt&&<span style={{fontSize:10,color:t.tm,marginLeft:4}}>({scaledAmt})</span>}
              </span>
            </div>;
          })}
        </div>
        {(r.cookMethod||r.cookTemp||r.cookTime)&&<div style={{background:t.waBg,border:`1px solid ${t.waBo}`,borderRadius:12,padding:"12px 16px",marginBottom:18,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          {r.cookMethod&&<div><div style={{fontSize:10,fontWeight:700,color:t.warn,letterSpacing:"0.1em",marginBottom:2}}>{t.L.detailMethod}</div><div style={{fontSize:15,fontWeight:700,color:t.text}}>{cookMethodL(r.cookMethod,t.L)}</div></div>}
          {r.cookTemp&&<div><div style={{fontSize:10,fontWeight:700,color:t.warn,letterSpacing:"0.1em",marginBottom:2}}>{t.L.detailTemp}</div><div style={{fontSize:15,fontWeight:700,color:t.text}}>{r.cookTemp}</div></div>}
          {r.cookTime&&<div><div style={{fontSize:10,fontWeight:700,color:t.warn,letterSpacing:"0.1em",marginBottom:2}}>{t.L.detailDuration}</div><div style={{fontSize:15,fontWeight:700,color:t.text}}>{r.cookTime}</div></div>}
        </div>}
        {r.notes&&<div style={{marginBottom:18}}><label style={lSt(t)}>{t.L.notes||"Notes"}</label><p style={{margin:0,color:t.ts,fontSize:14,lineHeight:1.7}}>{r.notes}</p></div>}
        {!r.photo&&!photoMode&&<button onClick={()=>setPM(true)} style={{...bSt("s",t),width:"100%",marginBottom:10,fontSize:13}}>{t.L.detailAddPhoto}</button>}
        {photoMode&&<div style={{marginBottom:10}}><PhotoPick onImg={(b,m)=>{if(b){onEdit({...r,photo:"data:"+m+";base64,"+b,_photoOnly:true});setPM(false)}}} t={t}/></div>}
        {r.photo&&<button onClick={()=>setPM(true)} style={{...bSt("s",t),width:"100%",marginBottom:10,fontSize:13}}>{t.L.editPhoto||"📷"}</button>}
        <button onClick={()=>setShowProdModal(true)} style={{...bSt("p",t),width:"100%",marginBottom:8,padding:12}}>{t.L.deductStock} (×{mult})</button>
        <button onClick={()=>setShowSampleModal(true)} style={{...bSt("s",t),width:"100%",marginBottom:8,padding:10,fontSize:13}}>{t.L.sampleLabel}</button>
        {/* Tek Paylaş Butonu */}
        <button onClick={()=>setShowShareModal(s=>!s)} style={{...bSt("s",t),width:"100%",marginBottom:8,padding:11,fontSize:14,fontWeight:600}}>
          📤 {t.lang==="tr"?"Paylaş / Yazdır":"Share / Print"}
        </button>
        {showShareModal&&<div style={{...cSt(t),padding:12,marginBottom:10,borderRadius:14}}>
          {/* Ekip Sohbetine */}
          {team&&onShareToChat&&<button onClick={()=>{
            onShareToChat({
              type:"recipe",id:r.id,name:r.name,
              desc:`${(r.ingredients||[]).length} ${t.lang==="tr"?"malzeme":"ingredients"} · ${r.calories||"?"}kcal`,
              data:{name:r.name,ingredients:r.ingredients,steps:r.steps,allergens:r.allergens,calories:r.calories,difficulty:r.difficulty}
            });
            setShowShareModal(false);onClose();
          }} style={{...bSt("s",t),width:"100%",marginBottom:8,padding:"9px 12px",fontSize:13,textAlign:"left"}}>
            👥 {t.lang==="tr"?"Ekip Sohbetine Gönder":"Send to Team Chat"}
          </button>}
          {/* Link Paylaş */}
          <button onClick={()=>{
            const shareUrl=`${window.location.origin}${window.location.pathname}?share=recipe&id=${r.id}&name=${encodeURIComponent(r.name)}`;
            const shareText=t.lang==="tr"?`🍽 "${r.name}" reçetesi:\n${shareUrl}`:`🍽 "${r.name}" recipe:\n${shareUrl}`;
            if(navigator.share){
              navigator.share({title:r.name,text:shareText,url:shareUrl}).catch(()=>{});
            }else{
              // Safari fallback
              const ta=document.createElement("textarea");
              ta.value=shareUrl;ta.style.position="fixed";ta.style.opacity="0";
              document.body.appendChild(ta);ta.focus();ta.select();
              try{document.execCommand("copy");}catch{}
              document.body.removeChild(ta);
              window.toast.success(t.lang==="tr"?"🔗 Link kopyalandı":"🔗 Link copied");
            }
            setShowShareModal(false);
          }} style={{...bSt("s",t),width:"100%",marginBottom:8,padding:"9px 12px",fontSize:13,textAlign:"left"}}>
            🔗 {t.lang==="tr"?"Link Kopyala":"Copy Link"}
          </button>
          {/* Yazdır */}
          <button onClick={()=>{
            const lg=t.lang;
            const ings=(r.ingredients||[]).map(i=>`<div class="ing"><span>${i.name}</span><span>${i.amount}</span></div>`).join("");
            const steps=(r.steps||[]).map((s,i)=>`<div class="step"><span class="n">${i+1}.</span><span>${s}</span></div>`).join("");
            const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${r.name}</title><style>body{font-family:sans-serif;max-width:600px;margin:40px auto;padding:20px;color:#222}h1{font-size:22px;margin-bottom:4px}.meta{color:#888;font-size:12px;margin-bottom:12px}h2{font-size:14px;font-weight:700;border-bottom:1px solid #eee;padding-bottom:4px;margin:14px 0 8px}.ing{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f5f5f5;font-size:13px}.step{padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:13px;display:flex;gap:8px}.n{font-weight:700;color:#b45309;min-width:20px}.foot{margin-top:24px;font-size:10px;color:#aaa;text-align:center}@media print{button{display:none}}</style></head><body><h1>${r.name}</h1><div class="meta">${r.mainCat||""} ${r.portions?"· "+r.portions+(lg==="tr"?" porsiyon":" portions"):""} ${r.calories?"· "+r.calories+"kcal":""}</div>${r.allergens?.length?`<div class="meta">⚠️ ${r.allergens.join(", ")}</div>`:""}<h2>${lg==="tr"?"Malzemeler":"Ingredients"}</h2>${ings}<h2>${lg==="tr"?"Hazırlanışı":"Instructions"}</h2>${steps}<div class="foot">Kitchen Manager · Tulpar Kitchen Software</div><script>window.onload=()=>window.print()<\/script></body></html>`;
            const w=window.open("","_blank");
            if(w){w.document.write(html);w.document.close();}
            setShowShareModal(false);
          }} style={{...bSt("s",t),width:"100%",padding:"9px 12px",fontSize:13,textAlign:"left"}}>
            🖨 {t.lang==="tr"?"Yazdır":"Print"}
          </button>
        </div>}
                <div style={{display:"flex",gap:10}}>
          <button onClick={()=>onDel(r.id)} style={{...bSt("d",t),flex:1}}>🗑 {t.L.delete||"Delete"}</button>
          <button onClick={()=>onEdit(r)} style={{...bSt("s",t),flex:1}}>✏ {t.L.edit||"Edit"}</button>
        </div>
      </div>
    </div></div>
    </>
  );
};

// ═══ CARD — kalori gösterimi (madde 6) ═══
const RCard=({r,onClick,t})=>{
  const mc=MAIN_CATS.find(c=>c.id===r.mainCat);
  const sc=(SUB_CATS[r.mainCat]||[]).find(s=>s.id===r.subCat);
  const ven=VENUES.find(v=>v.id===r.venue);
  const[h,setH]=useState(false);
  return <div onClick={()=>onClick(r)} onMouseEnter={()=>setH(true)} onMouseLeave={()=>setH(false)} style={{...cSt(t),cursor:"pointer",overflow:"hidden",transform:h?"translateY(-3px)":"none",boxShadow:h?t.cardSH:t.cardS}}>
    {r.photo&&<div style={{width:"100%",height:120,overflow:"hidden"}}><img src={r.photo} style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>}
    <div style={{padding:"16px 18px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:11,color:t.tm,marginBottom:4}}>{mc?.icon} {mainCatL(mc,t?.lang||"tr")}{sc?` › ${sc.label}`:""}</div>
          <div style={{fontSize:17,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif",lineHeight:1.3}}>{r.name}</div>
        </div>
        <div style={{background:t.acB,border:`1px solid ${t.acBo}`,borderRadius:8,padding:"3px 8px",fontSize:10,color:t.accent,marginLeft:8,fontWeight:600}}>{ven?.l}</div>
      </div>
      <div style={{fontSize:12,color:t.tm,marginBottom:8,display:"flex",gap:8,flexWrap:"wrap"}}>
        <span>📊 {r.yield}</span>
        <span>🥄 {r.ingredients.length}</span>
        {r.calories&&<span style={{color:t.warn}}>🔥 {r.calories} kcal/100g</span>}
        {r.cookTemp&&<span style={{color:t.warn}}>🌡 {r.cookTemp}</span>}
        {r.cookTime&&<span style={{color:t.warn}}>⏱ {r.cookTime}</span>}
      </div>
      {r.allergens.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:4}}>{r.allergens.slice(0,4).map(a=><ABadge key={a} a={a} t={t}/>)}{r.allergens.length>4&&<span style={{fontSize:11,color:t.tm,padding:"3px 6px"}}>+{r.allergens.length-4}</span>}</div>}
      {(r.diets||[]).length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:4}}>{(r.diets||[]).map(d=><DBadge key={d} d={d} t={t}/>)}</div>}
    </div>
  </div>;
};

// ═══ STOCK TAB — tam OCR sistemi + kategori (madde 5, 7) ═══
const StockTab=({stock,setStock,invoices,setInvoices,apiKey,traceability,lots,setLots,trackedIngs,profile,calorieDB,setCalorieDB,t})=>{
  const[stockSubTab,setStockSubTab]=useState("stock");
  const[showInv,setSI]=useState(false);const[showAdd,setSA]=useState(false);
  const[imgB,setIB]=useState(null);const[imgM,setIM]=useState("image/jpeg");
  const[loading,setL]=useState(false);const[err,setE]=useState("");
  const[ocrResult,setOCR]=useState(null);
  const[editIdx,setEditIdx]=useState(-1);
  const[ns,setNS]=useState({name:"",brand:"",unit:"g",qty:0,ppu:0,low:100,cat:"other",skt:"",lot:""});
  const lows=stock.filter(s=>s.qty<=(s.low||100));
  const[catFilter,setCF]=useState("all");
  const[editStockId,setESI]=useState(null);
  const[refLotIng,setRefLotIng]=useState(null);
  const[refLotMode,setRefLotMode]=useState("photo");
  const[refLotLoading,setRefLotLoading]=useState(false);
  // Barkod okuma
  const[barcodeLoading,setBarcodeLoading]=useState(false);
  const[barcodeResult,setBarcodeResult]=useState(null);
  const[barcodeStep,setBarcodeStep]=useState("idle"); // idle | scanned | label | done
  const[labelLoading,setLabelLoading]=useState(false);

  const lookupBarcode=async(barcode)=>{
    setBarcodeLoading(true);setBarcodeResult(null);
    // Türk ürün DB (manuel)
    const TR_PRODUCTS={
      "8690642105106":{name:"Tikveşli Krema 200ml",brand:"Tikveşli",cat:"dairy"},
      "8681856049164":{name:"Monea Mascarpone 400g",brand:"Monea",cat:"dairy"},
    };
    if(TR_PRODUCTS[barcode]){
      const p=TR_PRODUCTS[barcode];
      const parsed=parseProductName(p.name);
      const r={name:parsed.cleanName,brand:parsed.brand||p.brand,unit:parsed.unit,qty:parsed.qty,ppu:0,low:100,cat:p.cat,barcode};
      setBarcodeResult(r);setNS(r);setSA(true);setBarcodeStep("scanned");setBarcodeLoading(false);return;
    }
    try{
      const res=await fetch(`https://world.openfoodfacts.org/api/v0/product/${barcode}.json`);
      const data=await res.json();
      if(data.status===1&&data.product){
        const p=data.product;
        const rawName=p.product_name_tr||p.product_name||p.generic_name||barcode;
        // Marka + isim birleştir, sonra parse et
        const fullName=p.brands?`${p.brands} ${rawName}`:rawName;
        const parsed=parseProductName(fullName);
        // Net weight'ten miktar al (daha güvenilir)
        let qty=parsed.qty;let unit=parsed.unit;
        if(p.product_quantity&&p.product_quantity_unit){
          const pq=parseFloat(p.product_quantity)||0;
          const pu=p.product_quantity_unit.toLowerCase();
          if(pu==="g"||pu==="gr"){qty=pq;unit="g";}
          else if(pu==="kg"){qty=pq*1000;unit="g";}
          else if(pu==="ml"){qty=pq;unit="ml";}
          else if(pu==="l"||pu==="lt"){qty=pq*1000;unit="ml";}
        }
        const result={
          name:parsed.cleanName||rawName,
          brand:parsed.brand||p.brands||"",
          unit,qty,ppu:0,low:100,
          cat:p.categories_tags?.includes("en:dairy")||p.categories_tags?.includes("en:milks")?"dairy":
              p.categories_tags?.includes("en:meats")?"meat":
              p.categories_tags?.includes("en:vegetables")?"vegetable":"other",
          barcode
        };
        setBarcodeResult(result);setNS(result);setSA(true);setBarcodeStep("scanned");
      }else if(apiKey){
        // Open Food Facts'te yok — AI ile dene
        const raw=await callAI(apiKey,
          `Barkod numarası: ${barcode}\nBu Türk ürününü tanımlayabilir misin? Sadece JSON döndür: {"name":"...", "brand":"...", "unit":"g veya ml", "cat":"dairy/meat/vegetable/other"}`,[],
          "haiku");
        try{
          const obj=JSON.parse(raw.replace(/```json|```/g,"").trim());
          const result={...obj,qty:0,ppu:0,low:100,barcode};
          setBarcodeResult(result);setNS(result);setSA(true);setBarcodeStep("scanned");
        }catch{setBarcodeResult({error:t.lang==="tr"?"Ürün tanımlanamadı":"Could not identify product"});}
      }else{
        setBarcodeResult({error:t.lang==="tr"?"Ürün bulunamadı (barkod: "+barcode+")":"Product not found (barcode: "+barcode+")"});
      }
    }catch(e){
      setBarcodeResult({error:e.message});
    }
    setBarcodeLoading(false);
  };

  const analyzeInv=async()=>{
    if(!imgB){setE(t.L.errNoPhoto||"Fotoğraf seçin");return}
    setL(true);setE("");setOCR(null);
    try{
      const raw=await callAI(apiKey,P_OCR_INV,[
        {type:"image",source:{type:"base64",media_type:imgM,data:imgB}},
        {type:"text",text:"Bu faturayı analiz et. SADECE JSON döndür."}
      ],"haiku");
      const inv=parseJSON(raw);
      if(!inv.urunler&&!inv.items)throw new Error("Geçersiz yanıt");
      // Eski format uyumluluğu
      const urunler=inv.urunler||inv.items.map(it=>({ad:it.name,miktar:it.qty,birim:it.unit,birimFiyat:it.ppu,adet:1,toplam:it.qty*it.ppu}));
      const saglayici=inv.saglayici||inv.supplier||"";
      const tarih=inv.tarih||inv.date||new Date().toISOString().split("T")[0];
      const faturaToplam=inv.toplam||inv.total||0;
      setOCR({saglayici,tarih,urunler,faturaToplam,mod:inv.mod||"?"});
    }catch(e){setE("Hata: "+e.message.slice(0,120))}
    setL(false);
  };

  const applyInvoice=(urunler,saglayici,tarih,faturaToplam)=>{
    const invObj={id:Date.now(),saglayici,tarih,toplam:faturaToplam,urunler};
    setInvoices(p=>[invObj,...p]);
    setStock(prev=>{
      let next=[...prev];
      for(const it of urunler){
        const rawAd=(it.ad||it.name||"").trim();
        if(!rawAd)continue;
        // CORRECTION LEARNING: Kullanıcı düzelttiyse kaydet
        if(it._origAd&&it._origAd!==rawAd){
          saveCorrectionToDB(setCalorieDB,it._origAd,rawAd);
        }
        // 1. Önce correction DB kontrol: bu isim daha önce düzeltildi mi?
        const correction=getCorrectionFromDB(calorieDB,rawAd);
        const preliminaryAd=correction?correction.corrected:rawAd;
        // 2. Marka-agnostik normalize
        const ad=normalizeName(preliminaryAd);
        let miktar=it.miktar||it.qty||0;
        let birim=it.birim||it.unit||"adet";
        let ppu=it.birimFiyat||it.ppu||0;
        const toplam=it.toplam||0;
        if(toplam>0&&miktar>0){
          const hesaplanan=ppu*miktar;
          if(Math.abs(hesaplanan-toplam)/Math.max(toplam,1)>0.2){
            ppu=toplam/miktar;
          }
        }
        const cat=it.kategori||guessStockCat(ad);
        const adLower=ad.toLowerCase();
        let idx=next.findIndex(s=>{
          // Hem exact hem de normalize-edilmiş karşılaştır
          const existingLower=s.name.toLowerCase().trim();
          if(existingLower===adLower)return true;
          // Mevcut stok adı da normalize edilsin
          const existingNorm=normalizeName(s.name).toLowerCase();
          return existingNorm===adLower;
        });
        if(idx<0){
          let bestScore=0,bestIdx=-1;
          for(let i=0;i<next.length;i++){
            const score=fuzzyMatch(ad,next[i].name);
            if(score>bestScore&&score>=0.6){bestScore=score;bestIdx=i;}
          }
          idx=bestIdx;
        }
        if(idx>=0){
          next=next.map((s,i)=>i===idx?{...s,qty:s.qty+miktar,ppu:ppu||s.ppu,upd:tarih,cat:cat||s.cat||"other"}:s);
        }else{
          next=[...next,{id:Date.now()+Math.random()+ad.length,name:ad,unit:birim,qty:miktar,ppu,upd:tarih,low:100,cat:cat||"other"}];
        }
      }
      return next;
    });
    setSI(false);setIB(null);setOCR(null);setEditIdx(-1);
  };

  const filteredStock=catFilter==="all"?stock:stock.filter(s=>s.cat===catFilter);

  // Parti No helpers
  const todayStr=()=>{const d=new Date();return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getFullYear()).slice(-2)}`};
  const lotDaysOld=(iso)=>{if(!iso)return 999;const d=new Date(iso);const now=new Date();return Math.floor((now-d)/(1000*60*60*24))};
  const updateLot=(ing,val)=>{
    setLots(prev=>({...prev,[ing]:{no:val,updatedAt:new Date().toISOString(),updatedBy:(profile&&profile.fullName)||""}}));
  };

  // Referans Lot: fotoğraftan AI ile parti no çıkar
  const analyzeRefLot=async()=>{
    if(!imgB){setE(t.L.errNoPhoto||"Fotoğraf seçin");return}
    setRefLotLoading(true);setE("");
    try{
      const prompt=`You are a product label / package reading expert. Look at this image and find the LOT NUMBER, BATCH NUMBER, or PARTI NO. These usually appear as:
- "LOT: 24/04/26" or "PARTI: A1234"
- "Batch: B2024-04" or "L: 123456"
- Near expiry date (SKT/EXP)
- Codes like "LOT NO: XYZ123"
Return ONLY JSON: {"lot":"found_lot_number","confidence":"high|medium|low"}
If no lot number found, return: {"lot":"","confidence":"none"}`;
      const raw=await callAI(apiKey,prompt,[
        {type:"image",source:{type:"base64",media_type:imgM,data:imgB}},
        {type:"text",text:"Find the lot/batch/parti number. Return JSON only."}
      ],"haiku");
      const result=parseJSON(raw);
      if(result.lot&&result.lot.trim()){
        updateLot(refLotIng,result.lot.trim());
        setRefLotIng(null);
        setIB(null);
      }else{
        setE(t.L.refLotNotFound||"Parti numarası bulunamadı");
      }
    }catch(e){
      setE(e.message);
    }
    setRefLotLoading(false);
  };

  // QR tarayıcı: kamera açıp QR tararız (jsQR ile)
  const scanQRForLot=()=>{
    const input=document.createElement("input");
    input.type="file";
    input.accept="image/*";
    input.capture="environment";
    input.onchange=async(e)=>{
      const file=e.target.files[0];if(!file)return;
      const reader=new FileReader();
      reader.onload=async(ev)=>{
        // QR kodunu çözümlemek için: basit fallback - AI'a gönder
        const base64=ev.target.result.split(",")[1];
        setRefLotLoading(true);
        try{
          const raw=await callAI(apiKey,"Extract QR code content or any printed lot/batch/parti number from this image. Return ONLY JSON: {\"lot\":\"value\"}",[
            {type:"image",source:{type:"base64",media_type:file.type,data:base64}},
            {type:"text",text:"Return JSON only with lot number."}
          ],"haiku");
          const result=parseJSON(raw);
          if(result.lot&&result.lot.trim()){
            updateLot(refLotIng,result.lot.trim());
            setRefLotIng(null);
          }else{
            setE(t.L.refLotNotFound||"Parti numarası bulunamadı");
          }
        }catch(err){setE(err.message)}
        setRefLotLoading(false);
      };
      reader.readAsDataURL(file);
    };
    input.click();
  };

  return <div>
    {traceability&&<div style={{display:"flex",gap:6,marginBottom:14,background:t.inBg,padding:4,borderRadius:12}}>
      <button onClick={()=>setStockSubTab("stock")} style={{flex:1,padding:"8px 12px",borderRadius:10,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:stockSubTab==="stock"?t.card:"transparent",color:stockSubTab==="stock"?t.text:t.tm}}>{t.L.stockMaterial}</button>
      <button onClick={()=>setStockSubTab("lots")} style={{flex:1,padding:"8px 12px",borderRadius:10,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:stockSubTab==="lots"?t.card:"transparent",color:stockSubTab==="lots"?t.text:t.tm}}>{t.L.stockLots}</button>
    </div>}

    {traceability&&stockSubTab==="lots"&&<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <h3 style={{fontSize:18,color:t.text,margin:0}}>{t.L.currentLots}</h3>
        <div style={{fontSize:11,color:t.tm}}>{todayStr()}</div>
      </div>
      {trackedIngs.length===0?<div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}>
        <div style={{fontSize:40,opacity:0.4,marginBottom:10}}>🏷</div>
        <div style={{fontSize:13}}>{t.L.noTrackedIngs}</div>
        <div style={{fontSize:11,marginTop:4}}>{t.L.noTrackedIngsHint}</div>
      </div>:<>
        {trackedIngs.map(ing=>{
          const lot=lots[ing];
          const age=lot?lotDaysOld(lot.updatedAt):999;
          const status=!lot?"missing":age===0?"fresh":age<=2?"ok":"old";
          const col=status==="missing"?t.danger:status==="fresh"?t.success:status==="ok"?t.accent:t.warn;
          const icon=status==="missing"?"🔴":status==="fresh"?"🔵":status==="ok"?"🔵":"🟡";
          return <div key={ing} style={{...cSt(t),padding:"12px 14px",marginBottom:8,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{fontSize:14}}>{icon}</span>
            <div style={{flex:"1 1 180px",minWidth:120}}>
              <div style={{fontSize:14,fontWeight:600,color:t.text}}>{ing}</div>
              <div style={{fontSize:11,color:t.tm,marginTop:2}}>
                {!lot?t.L.lotMissing:`${t.L.lot}: ${lot.no} · ${age===0?t.L.lotToday:age+" "+t.L.lotDaysAgo} ${t.L.lotEntered}${lot.updatedBy?" ("+lot.updatedBy+")":""}`}
              </div>
            </div>
            <input type="text" placeholder={t.L.lot+" no"} defaultValue={lot?.no||""} onBlur={e=>{if(e.target.value.trim()&&e.target.value.trim()!==lot?.no)updateLot(ing,e.target.value.trim())}} style={{...iSt(t),width:100,fontSize:12,padding:"6px 8px"}}/>
            <button onClick={()=>{setRefLotIng(ing);setRefLotMode("photo");setIB(null);setE("")}} style={{...bSt("s",t),padding:"6px 10px",fontSize:11}} title={t.L.refLotFromPhoto}>📷</button>
          </div>;
        })}
        <div style={{marginTop:14,padding:"10px 12px",background:t.inBg,borderRadius:10,fontSize:11,color:t.tm,lineHeight:1.6}}>
          {t.L.lotHint}
        </div>
      </>}
    </div>}

    {(!traceability||stockSubTab==="stock")&&<>
    {lows.length>0&&<div style={{background:t.danBg,border:`1px solid ${t.danBo}`,borderRadius:14,padding:"14px 18px",marginBottom:20}}>
      <div style={{fontWeight:700,color:t.danger,fontSize:14}}>{t.L.lowStock} — {lows.length} {t.L.products}</div>
      <div style={{fontSize:13,color:t.danger,opacity:0.8,marginTop:4}}>{lows.map(s=>s.name).join(", ")}</div>
    </div>}
    <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
      <button onClick={()=>setSI(true)} style={{...bSt("p",t),flex:"1 1 120px"}}>{t.L.invoice}</button>
      <button onClick={()=>setSA(true)} style={{...bSt("s",t),flex:"1 1 100px"}}>{t.L.manualAdd}</button>
      <label style={{...bSt("s",t),flex:"1 1 100px",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>
        📷 {t.lang==="tr"?"Barkod":"Barcode"}
        <input type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={async e=>{
          const file=e.target.files?.[0];if(!file)return;
          if(!apiKey){window.toast.error(t.lang==="tr"?"API key gerekli":"API key required");return;}
          setBarcodeLoading(true);setBarcodeStep("idle");
          const reader=new FileReader();
          reader.onload=async(ev)=>{
            try{
              const b64=ev.target.result.split(",")[1];
              const raw=await callAI(apiKey,
                "Extract the barcode/EAN/UPC number from this product image. Return ONLY the number, nothing else.",
                [{type:"image",source:{type:"base64",media_type:file.type,data:b64}}],
                "haiku"
              );
              const barcode=raw.trim().replace(/[^0-9]/g,"");
              if(barcode.length>=8)await lookupBarcode(barcode);
              else{setBarcodeResult({error:t.lang==="tr"?"Barkod okunamadı":"Could not read barcode"});setBarcodeLoading(false);}
            }catch(err){setBarcodeResult({error:err.message});setBarcodeLoading(false);}
          };
          reader.readAsDataURL(file);
          e.target.value="";
        }}/>
      </label>
    </div>

    {/* Adım 1: Barkod okundu → etiket çek */}
    {barcodeLoading&&<div style={{...cSt(t),padding:"14px 16px",marginBottom:8,display:"flex",alignItems:"center",gap:10}}>
      <div style={{fontSize:20}}>⏳</div>
      <div style={{fontSize:13,color:t.tm}}>{t.lang==="tr"?"Barkod aranıyor...":"Looking up barcode..."}</div>
    </div>}

    {barcodeResult?.error&&<div style={{color:t.danger,fontSize:12,marginBottom:8,padding:"10px 14px",background:t.danBg,borderRadius:10}}>⚠️ {barcodeResult.error}</div>}

    {barcodeStep==="scanned"&&barcodeResult&&!barcodeResult.error&&<div style={{...cSt(t),padding:"14px 16px",marginBottom:12}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
        <div style={{width:36,height:36,borderRadius:10,background:t.acB,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>✅</div>
        <div>
          <div style={{fontSize:14,fontWeight:700,color:t.text}}>{barcodeResult.name}</div>
          <div style={{fontSize:11,color:t.tm}}>{barcodeResult.brand&&barcodeResult.brand+" · "}{barcodeResult.qty>0?barcodeResult.qty+" "+barcodeResult.unit:""}</div>
        </div>
      </div>
      <div style={{fontSize:12,color:t.tm,marginBottom:10,padding:"8px 10px",background:t.inBg,borderRadius:8}}>
        📷 {t.lang==="tr"?"Şimdi SKT ve Lot numarasının yazılı olduğu etiketi çek:":"Now take a photo of the label with expiry date and lot number:"}
      </div>
      <div style={{display:"flex",gap:8}}>
        <label style={{...bSt("p",t),flex:2,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          📷 {t.lang==="tr"?"Etiketi Tara":"Scan Label"}
          <input type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={async e=>{
            const file=e.target.files?.[0];if(!file)return;
            if(!apiKey)return;
            setLabelLoading(true);
            const reader=new FileReader();
            reader.onload=async(ev)=>{
              try{
                const b64=ev.target.result.split(",")[1];
                const raw=await callAI(apiKey,
                  `Bu ürün etiketinde son kullanma tarihi (SKT, BBE, Exp, MHD) ve parti/lot numarasını bul.
Sadece JSON döndür, başka hiçbir şey yazma:
{"skt":"YYYY-MM-DD","lot":"..."}
SKT bulamazsan skt:"" yaz. Lot bulamazsan lot:"" yaz.
Tarih formatı mutlaka YYYY-MM-DD olsun.`,
                  [{type:"image",source:{type:"base64",media_type:file.type,data:b64}}],
                  "haiku"
                );
                const obj=JSON.parse(raw.replace(/```json|```/g,"").trim());
                setNS(s=>({...s,skt:obj.skt||s.skt,lot:obj.lot||s.lot}));
                setBarcodeStep("done");
              }catch(err){
                // Hata olsa bile devam et
                setBarcodeStep("done");
              }
              setLabelLoading(false);
              e.target.value="";
            };
            reader.readAsDataURL(file);
          }}/>
        </label>
        <button onClick={()=>setBarcodeStep("done")} style={{...bSt("s",t),flex:1,fontSize:12}}>
          {t.lang==="tr"?"Atla":"Skip"}
        </button>
      </div>
      {labelLoading&&<div style={{textAlign:"center",padding:"10px",color:t.tm,fontSize:12,marginTop:8}}>⏳ {t.lang==="tr"?"Etiket okunuyor...":"Reading label..."}</div>}
    </div>}

    {barcodeStep==="done"&&barcodeResult&&!barcodeResult.error&&<div style={{...cSt(t),padding:"12px 14px",marginBottom:12,borderLeft:`3px solid ${t.success}`}}>
      <div style={{fontSize:13,fontWeight:700,color:t.success,marginBottom:6}}>✓ {barcodeResult.name}</div>
      <div style={{display:"flex",gap:12,fontSize:12,color:t.tm,flexWrap:"wrap"}}>
        {barcodeResult.qty>0&&<span>📦 {barcodeResult.qty} {barcodeResult.unit}</span>}
        {ns.skt&&<span>📅 SKT: {ns.skt}</span>}
        {ns.lot&&<span>🏷 Lot: {ns.lot}</span>}
      </div>
      <button onClick={()=>{setBarcodeResult(null);setBarcodeStep("idle");}} style={{fontSize:11,color:t.tm,background:"none",border:"none",cursor:"pointer",marginTop:6}}>
        ↩ {t.lang==="tr"?"Yeni barkod tara":"Scan new barcode"}
      </button>
    </div>}

    {/* Kategori filtresi */}
    <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:6,marginBottom:14}}>
      <button onClick={()=>setCF("all")} style={{whiteSpace:"nowrap",padding:"5px 12px",borderRadius:16,fontSize:12,fontWeight:500,border:`1px solid ${catFilter==="all"?t.acBo:t.inBo}`,background:catFilter==="all"?t.acB:"transparent",color:catFilter==="all"?t.accent:t.tm,cursor:"pointer"}}>{t.L.all}</button>
      {STOCK_CATS.map(c=><button key={c.id} onClick={()=>setCF(c.id)} style={{whiteSpace:"nowrap",padding:"5px 12px",borderRadius:16,fontSize:12,fontWeight:500,border:`1px solid ${catFilter===c.id?t.acBo:t.inBo}`,background:catFilter===c.id?t.acB:"transparent",color:catFilter===c.id?t.accent:t.tm,cursor:"pointer"}}>{c.icon} {stockCatL(c,t?.lang||"tr")}</button>)}
    </div>

    {/* REFERANS LOT MODAL */}
    {refLotIng&&<div style={mOv(t)} onClick={()=>{setRefLotIng(null);setIB(null);setE("")}}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),maxWidth:440}}>
      <h3 style={{fontSize:20,marginBottom:6,color:t.text}}>{t.L.refLotTitle||"🏷 Referans Lot"}</h3>
      <div style={{fontSize:13,color:t.accent,fontWeight:700,marginBottom:4}}>{refLotIng}</div>
      <div style={{fontSize:12,color:t.tm,marginBottom:16,lineHeight:1.5}}>{t.L.refLotHint||"Ambalajın fotoğrafını çek veya ürün QR'ını okut"}</div>
      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <button onClick={()=>setRefLotMode("photo")} style={{...bSt(refLotMode==="photo"?"p":"s",t),flex:1,fontSize:13}}>{t.L.refLotFromPhoto||"📷 Fotoğraftan"}</button>
        <button onClick={()=>{setRefLotMode("qr");scanQRForLot()}} style={{...bSt(refLotMode==="qr"?"p":"s",t),flex:1,fontSize:13}}>{t.L.refLotFromQR||"📷 QR Tara"}</button>
      </div>
      {refLotMode==="photo"&&<>
        <PhotoPick onImg={(b,m)=>{setIB(b);setIM(m||"image/jpeg")}} t={t}/>
        {err&&<div style={{color:t.danger,fontSize:13,marginTop:10,background:t.danBg,borderRadius:10,padding:"10px 14px"}}>{err}</div>}
        <button onClick={analyzeRefLot} disabled={refLotLoading||!imgB} style={{...bSt("p",t),width:"100%",marginTop:14,padding:14,opacity:(refLotLoading||!imgB)?0.5:1}}>
          {refLotLoading?"⏳ "+(t.L.loading||"Okunuyor..."):"✦ "+(t.L.refLotFromPhoto||"AI ile Oku")}
        </button>
      </>}
      {refLotMode==="qr"&&refLotLoading&&<div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}>⏳ {t.L.loading||"Okunuyor..."}</div>}
      {refLotMode==="qr"&&err&&<div style={{color:t.danger,fontSize:13,marginTop:10,background:t.danBg,borderRadius:10,padding:"10px 14px"}}>{err}</div>}
      <button onClick={()=>{setRefLotIng(null);setIB(null);setE("")}} style={{...bSt("s",t),width:"100%",marginTop:10}}>{t.L.cancel||"İptal"}</button>
    </div></div>}

    {/* FATURA MODAL */}
    {showInv&&<div style={mOv(t)} onClick={()=>{setSI(false);setOCR(null);setIB(null);setEditIdx(-1)}}><div onClick={e=>e.stopPropagation()} style={mPn(t)}>
      <h3 style={{fontSize:20,marginBottom:18,color:t.text}}>{t.L.invoiceTitle}</h3>
      {!ocrResult?<>
        <PhotoPick onImg={(b,m)=>{setIB(b);setIM(m||"image/jpeg")}} t={t}/>
        {err&&<div style={{color:t.danger,fontSize:13,marginTop:10,background:t.danBg,borderRadius:10,padding:"10px 14px"}}>{err}</div>}
        <button onClick={analyzeInv} disabled={loading} style={{...bSt("p",t),width:"100%",marginTop:14,padding:14,opacity:loading?0.7:1}}>{loading?"⏳ "+t.L.loading:"✦ "+t.L.invoiceAnalyze}</button>
      </>:<>
        <div style={{background:t.sucBg,border:`1px solid ${t.sucBo}`,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:13,color:t.success}}>
          ✓ Mod {ocrResult.mod} · {ocrResult.saglayici||"?"} · {ocrResult.tarih} · {ocrResult.urunler.length} ürün
        </div>
        <div style={{maxHeight:320,overflowY:"auto",marginBottom:14}}>
          {ocrResult.urunler.map((u,i)=>{
            const cat=STOCK_CATS.find(c=>c.id===(u.kategori||guessStockCat(u.ad||"")));
            const isEdit=editIdx===i;
            const updateRow=(field,val)=>{
              const newUrunler=[...ocrResult.urunler];
              // İsim değişiyorsa orijinali sakla (correction learning için)
              if(field==="ad"&&!newUrunler[i]._origAd&&newUrunler[i].ad){
                newUrunler[i]={...newUrunler[i],_origAd:newUrunler[i].ad};
              }
              newUrunler[i]={...newUrunler[i],[field]:val};
              setOCR({...ocrResult,urunler:newUrunler});
            };
            const deleteRow=()=>{
              const newUrunler=ocrResult.urunler.filter((_,j)=>j!==i);
              setOCR({...ocrResult,urunler:newUrunler});
              setEditIdx(-1);
            };
            if(isEdit){
              return <div key={i} style={{padding:"10px 0",borderBottom:`1px solid ${t.border}`,background:t.acB,borderRadius:8,padding:"10px",marginBottom:4}}>
                <input style={{...iSt(t),fontSize:13,padding:"8px 10px",marginBottom:6}} value={u.ad||""} onChange={e=>updateRow("ad",e.target.value)} placeholder="Ürün adı"/>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:6}}>
                  <input style={{...iSt(t),fontSize:13,padding:"8px 10px"}} type="number" step="0.01" value={u.miktar||0} onChange={e=>updateRow("miktar",parseFloat(e.target.value)||0)} placeholder="Miktar"/>
                  <select style={{...iSt(t),fontSize:13,padding:"8px 10px"}} value={u.birim||"adet"} onChange={e=>updateRow("birim",e.target.value)}>
                    {["g","ml","adet","kg","l"].map(x=><option key={x} value={x}>{x}</option>)}
                  </select>
                  <input style={{...iSt(t),fontSize:13,padding:"8px 10px"}} type="number" step="0.01" value={u.birimFiyat||0} onChange={e=>updateRow("birimFiyat",parseFloat(e.target.value)||0)} placeholder="₺"/>
                </div>
                <select style={{...iSt(t),fontSize:13,padding:"8px 10px",marginBottom:6}} value={u.kategori||guessStockCat(u.ad||"")} onChange={e=>updateRow("kategori",e.target.value)}>
                  {STOCK_CATS.map(c=><option key={c.id} value={c.id}>{c.icon} {stockCatL(c,t?.lang||"tr")}</option>)}
                </select>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>setEditIdx(-1)} style={{...bSt("p",t),flex:1,padding:"8px",fontSize:12}}>✓ Tamam</button>
                  <button onClick={deleteRow} style={{...bSt("d",t),padding:"8px 14px",fontSize:12}}>🗑 Sil</button>
                </div>
              </div>;
            }
            return <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${t.border}`,fontSize:13,gap:6}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{color:t.text,fontWeight:600}}>{u.ad||"(adsız)"}</div>
                <div style={{color:t.tm,fontSize:11,marginTop:2}}>
                  {cat?.icon} {cat?.l} · {u.miktar} {u.birim} · ₺{(u.birimFiyat||0).toFixed(2)}
                  {u.flag&&<span style={{color:t.danger,marginLeft:6}}>⚠</span>}
                </div>
              </div>
              <button onClick={()=>setEditIdx(i)} style={{...bSt("s",t),padding:"6px 10px",fontSize:11}}>✏</button>
              <button onClick={deleteRow} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
            </div>;
          })}
          {ocrResult.urunler.length===0&&<div style={{textAlign:"center",padding:"20px",color:t.tm,fontSize:13}}>Tüm ürünler silindi. "Yeniden" ile tekrar tara.</div>}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{setOCR(null);setIB(null);setEditIdx(-1)}} style={{...bSt("s",t),flex:1}}>↩ {t.L.back}</button>
          <button disabled={ocrResult.urunler.length===0} onClick={()=>{setEditIdx(-1);applyInvoice(ocrResult.urunler,ocrResult.saglayici,ocrResult.tarih,ocrResult.faturaToplam)}} style={{...bSt("p",t),flex:2,opacity:ocrResult.urunler.length===0?0.5:1}}>{t.L.invoiceSaveBtn} ({ocrResult.urunler.length})</button>
        </div>
      </>}
    </div></div>}

    {/* MANUEL STOK MODAL — label düzeltmesi (madde 4) */}
    {showAdd&&<div style={mOv(t)} onClick={()=>setSA(false)}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),maxWidth:420}}>
      <h3 style={{fontSize:20,marginBottom:16,color:t.text}}>{t.L.stockAddManual}</h3>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <div><label style={lSt(t)}>{t.L.name}</label><input style={iSt(t)} placeholder={t.L.name} value={ns.name} onChange={e=>{
          const n=e.target.value;
          const parsed=parseProductName(n);
          setNS(s=>({...s,
            name:n,
            brand:s.brand||parsed.brand||"",
            cat:n?guessStockCat(n):"other",
            qty:s.qty===0&&parsed.qty>0?parsed.qty:s.qty,
            unit:s.unit==="g"&&parsed.unit!=="g"?parsed.unit:s.unit
          }));
        }}/></div>
        <div><label style={lSt(t)}>🏷 {t.lang==="tr"?"Marka":"Brand"}</label><input style={iSt(t)} placeholder={t.lang==="tr"?"Pınar, Sek, Ülker...":"Brand name..."} value={ns.brand||""} onChange={e=>setNS(s=>({...s,brand:e.target.value}))}/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lSt(t)}>{t.L.unit}</label><select style={iSt(t)} value={ns.unit} onChange={e=>setNS(s=>({...s,unit:e.target.value}))}>{["g","ml","adet","kg","l"].map(u=><option key={u}>{u}</option>)}</select></div>
          <div><label style={lSt(t)}>{t.L.stockCategory}</label><select style={iSt(t)} value={ns.cat} onChange={e=>setNS(s=>({...s,cat:e.target.value}))}>{STOCK_CATS.map(c=><option key={c.id} value={c.id}>{c.icon} {stockCatL(c,t?.lang||"tr")}</option>)}</select></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lSt(t)}>{t.L.amount}</label><input style={iSt(t)} type="number" placeholder="0" value={ns.qty} onChange={e=>setNS(s=>({...s,qty:parseFloat(e.target.value)||0}))}/></div>
          <div><label style={lSt(t)}>₺ / {t.L.unit}</label><input style={iSt(t)} type="number" step="0.001" placeholder="0.000" value={ns.ppu} onChange={e=>setNS(s=>({...s,ppu:parseFloat(e.target.value)||0}))}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lSt(t)}>📅 SKT</label><input style={iSt(t)} type="date" value={ns.skt||""} onChange={e=>setNS(s=>({...s,skt:e.target.value}))}/></div>
          <div><label style={lSt(t)}>🏷 Parti/Lot</label><input style={iSt(t)} placeholder="L240301" value={ns.lot||""} onChange={e=>setNS(s=>({...s,lot:e.target.value}))}/></div>
        </div>
        <div><label style={lSt(t)}>{t.L.stockMinLevel}</label><input style={iSt(t)} type="number" placeholder="100" value={ns.low} onChange={e=>setNS(s=>({...s,low:parseFloat(e.target.value)||100}))}/></div>
        <button onClick={()=>{
          if(!ns.name)return;
          const parsed=parseProductName(ns.name);
          const cleanName=parsed.cleanName||normalizeName(ns.name);
          const finalQty=ns.qty>0?ns.qty:parsed.qty;
          const finalUnit=parsed.unit!=="g"||ns.unit!=="g"?parsed.unit:ns.unit;
          setStock(p=>{
            const idx=p.findIndex(s=>normalizeName(s.name).toLowerCase()===cleanName.toLowerCase());
            if(idx>=0){
              return p.map((s,i)=>i===idx?{...s,qty:s.qty+finalQty,ppu:ns.ppu||s.ppu,skt:ns.skt||s.skt,lot:ns.lot||s.lot,upd:new Date().toISOString().split("T")[0]}:s);
            }
            return [...p,{...ns,name:cleanName,qty:finalQty,unit:finalUnit,id:Date.now(),upd:new Date().toISOString().split("T")[0]}];
          });
          setNS({name:"",brand:"",unit:"g",qty:0,ppu:0,low:100,cat:"other",skt:"",lot:""});
          setSA(false);
        }} style={{...bSt("p",t),padding:12}}>{t.L.stockAddBtn}</button>
      </div>
    </div></div>}

    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
      <h3 style={{fontSize:22,color:t.text,margin:0}}>{t.L.stock} <span style={{fontSize:14,color:t.tm,fontWeight:400}}>({filteredStock.length})</span></h3>
      <button onClick={()=>{
        // Mevcut stokta marka-agnostik birleştirme yap
        const groups={};
        for(const s of stock){
          const key=normalizeName(s.name).toLowerCase()+"|"+s.unit;
          if(!groups[key])groups[key]={...s,name:normalizeName(s.name)};
          else{
            groups[key].qty+=s.qty;
            // En yüksek ppu değerini al (daha güvenilir)
            if(s.ppu>groups[key].ppu)groups[key].ppu=s.ppu;
            // En yeni tarih
            if(s.upd>groups[key].upd)groups[key].upd=s.upd;
          }
        }
        const merged=Object.values(groups);
        if(merged.length<stock.length){
          if(window.confirm(`${stock.length-merged.length} ${t.lang==="tr"?"ürün markalardan ayrılıp birleştirilecek. Devam?":t.lang==="en"?"items will be merged by removing brands. Continue?":"işlem?"}`)){
            setStock(merged);
          }
        }else{
          window.toast.error(t.lang==="tr"?"Birleştirilecek ürün yok":t.lang==="en"?"No items to merge":"OK");
        }
      }} style={{...bSt("s",t),fontSize:11,padding:"6px 12px"}} title={t.lang==="tr"?"Marka-agnostik birleştirme":"Brand-agnostic merge"}>🔀 {t.lang==="tr"?"Markaları Birleştir":t.lang==="en"?"Merge Brands":"🔀"}</button>
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      {filteredStock.map(it=>{
        const isLow=it.qty<=(it.low||100);
        const catInfo=STOCK_CATS.find(c=>c.id===it.cat)||STOCK_CATS.find(c=>c.id==="other");
        const isEdit=editStockId===it.id;
        const updateField=(field,val)=>setStock(p=>p.map(s=>s.id===it.id?{...s,[field]:val}:s));
        if(isEdit){
          return <div key={it.id} style={{...cSt(t),padding:"14px 16px",background:t.acB}}>
            <div style={{fontSize:11,fontWeight:700,color:t.accent,letterSpacing:"0.1em",marginBottom:8}}>{t.L.edit.toUpperCase()}</div>
            <input style={{...iSt(t),fontSize:13,marginBottom:8}} value={it.name} onChange={e=>updateField("name",e.target.value)} placeholder={t.L.name}/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
              <select style={{...iSt(t),fontSize:13}} value={it.unit} onChange={e=>updateField("unit",e.target.value)}>
                {["g","ml","adet","kg","l"].map(x=><option key={x}>{x}</option>)}
              </select>
              <select style={{...iSt(t),fontSize:13}} value={it.cat||"other"} onChange={e=>updateField("cat",e.target.value)}>
                {STOCK_CATS.map(c=><option key={c.id} value={c.id}>{c.icon} {stockCatL(c,t?.lang||"tr")}</option>)}
              </select>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
              <input style={{...iSt(t),fontSize:13}} type="number" step="0.01" value={it.qty} onChange={e=>updateField("qty",parseFloat(e.target.value)||0)} placeholder="Miktar"/>
              <input style={{...iSt(t),fontSize:13}} type="number" step="0.001" value={it.ppu} onChange={e=>updateField("ppu",parseFloat(e.target.value)||0)} placeholder="₺/birim"/>
              <input style={{...iSt(t),fontSize:13}} type="number" value={it.low||100} onChange={e=>updateField("low",parseFloat(e.target.value)||100)} placeholder="Düşük eşik"/>
            </div>
            <div style={{display:"flex",gap:6}}>
              <button onClick={()=>setESI(null)} style={{...bSt("p",t),flex:1,padding:"8px",fontSize:12}}>✓ Tamam</button>
              <button onClick={()=>{setStock(p=>p.filter(s=>s.id!==it.id));setESI(null);}} style={{...bSt("d",t),padding:"8px 14px",fontSize:12}}>🗑 Sil</button>
            </div>
          </div>;
        }
        return <div key={it.id} style={{...cSt(t),padding:"14px 16px",display:"flex",alignItems:"center",gap:10,borderColor:isLow?t.danBo:t.cardB}}>
          {isLow&&<div style={{width:3,height:32,borderRadius:2,background:t.danger,flexShrink:0}}/>}
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:600,fontSize:14,color:t.text}}>{it.name}{it.brand&&<span style={{fontSize:11,color:t.tm,fontWeight:400}}> · {it.brand}</span>}</div>
            <div style={{fontSize:11,color:t.tm}}>{catInfo.icon} {catInfo.l} · {it.upd}</div>
            {(it.skt||it.lot)&&<div style={{fontSize:10,marginTop:2,display:"flex",gap:8}}>
              {it.skt&&<span style={{color:new Date(it.skt)<new Date()?t.danger:new Date(it.skt)<new Date(Date.now()+7*86400000)?"#f59e0b":t.tm}}>
                📅 {it.skt}{new Date(it.skt)<new Date()?" ⚠️":""}
              </span>}
              {it.lot&&<span style={{color:t.tm}}>🏷 {it.lot}</span>}
            </div>}
          </div>
          <div style={{textAlign:"right",marginRight:4}}>
            <div style={{fontWeight:700,fontSize:14,color:isLow?t.danger:t.text}}>{fmtQty(it.qty,it.unit)}</div>
            <div style={{fontSize:11,color:t.tm,display:"flex",alignItems:"center",gap:3,justifyContent:"flex-end"}}>
              {(()=>{const p=it.ppu||0,u=it.unit;const tooHigh=(u==="g"&&p>10)||(u==="ml"&&p>5)||(u==="kg"&&p>2000)||(u==="l"&&p>1000)||(u==="adet"&&p>500);return tooHigh?<span title="Bu fiyat beklenenden yüksek, kontrol et" style={{color:t.warn,fontSize:13}}>⚠</span>:null})()}
              ₺{(it.ppu||0).toFixed(3)}/{it.unit}
            </div>
          </div>
          <button onClick={()=>setESI(it.id)} style={{...bSt("s",t),padding:"8px 10px",fontSize:12}}>✏</button>
          <button onClick={()=>setStock(p=>p.filter(s=>s.id!==it.id))} style={{...bSt("d",t),padding:"8px 10px",fontSize:12}}>×</button>
        </div>;
      })}
    </div>
    </>}
  </div>;
};

// ═══ MENU KARTI OLUŞTURUCU ═══
const MENU_TYPES=[
  {id:"festival",l:"Festival",tr:"Festival",en:"Festival",ru:"Фестиваль",es:"Festival",de:"Festival",fr:"Festival",zh:"节日",ar:"مهرجان",icon:"🎪",sections:["Başlangıç","Ana Yemek","Tatlı","İçecek"],sectionsEn:["Starter","Main","Dessert","Drink"]},
  {id:"alacarte",l:"Restoran",tr:"Restoran",en:"Restaurant",ru:"Ресторан",es:"Restaurante",de:"Restaurant",fr:"Restaurant",zh:"餐厅",ar:"مطعم",icon:"🍷",sections:["Amuse-bouche","Başlangıç","Ara Sıcak","Ana Yemek","Tatlı"],sectionsEn:["Amuse-bouche","Starter","Mid-Course","Main","Dessert"]},
  {id:"banquet",l:"Ziyafet",tr:"Ziyafet",en:"Banquet",ru:"Банкет",es:"Banquete",de:"Bankett",fr:"Banquet",zh:"宴会",ar:"وليمة",icon:"🎩",sections:["Soğuk","Sıcak","Ana","Tatlı"],sectionsEn:["Cold","Hot","Main","Dessert"]},
  {id:"catering",l:"Catering",tr:"Catering",en:"Catering",ru:"Кейтеринг",es:"Catering",de:"Catering",fr:"Traiteur",zh:"外烩",ar:"تموين",icon:"🍱",sections:["Başlangıç","Ana Yemek","Yan Lezzetler","Tatlı","İçecek"],sectionsEn:["Starter","Main","Sides","Dessert","Drink"]},
  {id:"buffet",l:"Büfe",tr:"Büfe",en:"Buffet",ru:"Шведский стол",es:"Bufé",de:"Buffet",fr:"Buffet",zh:"自助餐",ar:"بوفيه",icon:"🍽",sections:["Soğuk Meze","Sıcak Yemek","Pilav & Makarna","Salata","Tatlı"],sectionsEn:["Cold Meze","Hot Dish","Rice & Pasta","Salad","Dessert"]},
  {id:"drinks",l:"İçecekler",tr:"İçecekler",en:"Drinks",ru:"Напитки",es:"Bebidas",de:"Getränke",fr:"Boissons",zh:"饮品",ar:"مشروبات",icon:"🍹",sections:["Sıcak İçecekler","Soğuk İçecekler","Kokteyller","Şaraplar","Alkolsüz"],sectionsEn:["Hot Drinks","Cold Drinks","Cocktails","Wines","Non-Alcoholic"]}
];
const menuTypeL=(m,lang)=>m&&(m[lang]||m.tr||m.l)||"";
const menuSectionsL=(m,lang)=>{
  if(!m)return [];
  if(lang==="en"&&m.sectionsEn)return m.sectionsEn;
  return m.sections||[];
};
const MENU_FONTS=["Fraunces","Playfair Display","Cormorant Garamond","EB Garamond","Inter Tight"];
const MENU_THEMES=[
  {id:"classic",l:"Klasik",tr:"Klasik",en:"Classic",ru:"Классика",es:"Clásico",de:"Klassisch",fr:"Classique",zh:"经典",ar:"كلاسيكي",bg:"#fff",text:"#1a1612",accent:"#333",border:"#ddd",icon:"classic"},
  {id:"warm",l:"Sıcak",tr:"Sıcak",en:"Warm",ru:"Тёплый",es:"Cálido",de:"Warm",fr:"Chaud",zh:"温暖",ar:"دافئ",bg:"#faf8f3",text:"#1a1612",accent:"#c8965a",border:"#ebe5d8",icon:"warm"},
  {id:"elegant",l:"Elegan",tr:"Elegan",en:"Elegant",ru:"Элегантный",es:"Elegante",de:"Elegant",fr:"Élégant",zh:"优雅",ar:"أنيق",bg:"#0d1b2a",text:"#e0e1dd",accent:"#c9a96e",border:"#1b263b",icon:"elegant"},
  {id:"modern",l:"Modern",tr:"Modern",en:"Modern",ru:"Современный",es:"Moderno",de:"Modern",fr:"Moderne",zh:"现代",ar:"حديث",bg:"#f8f9fa",text:"#212529",accent:"#0d6efd",border:"#dee2e6",icon:"modern"}
];
const themeL=(th,lang)=>th&&(th[lang]||th.tr||th.l)||"";

// ═══ DEPO KONTROL MODAL ═══
const StorageCheckModal=({storage,storageType,date,profile,onSave,onClose,t})=>{
  const nowTime=()=>{const d=new Date();const h=d.getHours();
    if(h<12)return "09.00";
    if(h<18)return "15.00";
    return "21.00";
  };
  const[time,setTime]=useState(nowTime());
  const[checks,setChecks]=useState({etiket:true,acikGida:true,yerdeDuran:true,cigPismis:true,temizlik:true,alerjen:true,skt:true});
  const[urunSicaklik,setUrunSicaklik]=useState("");
  const[gostergeDeger,setGostergeDeger]=useState("");
  const[aciklama,setAciklama]=useState("");

  const save=()=>{
    if(!urunSicaklik&&!gostergeDeger){window.toast.info(t.L.atLeastOneTemp);return}
    onSave({
      id:Date.now(),
      storageId:storage.id,
      storageName:storage.name,
      date,
      time,
      ...checks,
      urunSicaklik,
      gostergeDeger,
      aciklama,
      kontrolEden:(profile&&profile.fullName)||"",
      createdAt:new Date().toISOString()
    });
  };

  const CheckRow=({lbl,k})=>(
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${t.border}`}}>
      <span style={{fontSize:13,color:t.text}}>{lbl}</span>
      <div style={{display:"flex",gap:4}}>
        <button onClick={()=>setChecks({...checks,[k]:true})} style={{padding:"4px 10px",borderRadius:8,border:`1px solid ${checks[k]?t.success:t.border}`,background:checks[k]?t.success:"transparent",color:checks[k]?"#fff":t.tm,fontSize:11,fontWeight:600,cursor:"pointer"}}>{t.L.ok}</button>
        <button onClick={()=>setChecks({...checks,[k]:false})} style={{padding:"4px 10px",borderRadius:8,border:`1px solid ${!checks[k]?t.danger:t.border}`,background:!checks[k]?t.danger:"transparent",color:!checks[k]?"#fff":t.tm,fontSize:11,fontWeight:600,cursor:"pointer"}}>{t.L.notOk}</button>
      </div>
    </div>
  );

  return <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,zIndex:9999}} onClick={onClose}><div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:18,padding:22,maxWidth:480,width:"100%",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.4)"}}>
    <h3 style={{fontSize:18,color:t.text,marginBottom:4}}>{t.L.storageCheckTitle}</h3>
    <div style={{fontSize:12,color:t.tm,marginBottom:14}}>{storageType?.icon} {storage.name} · {storage.temp}°C · {date}</div>

    <div style={{marginBottom:12}}>
      <label style={lSt(t)}>{t.L.controlTime}</label>
      <div style={{display:"flex",gap:4}}>
        {["09.00","15.00","21.00"].map(h=><button key={h} onClick={()=>setTime(h)} style={{flex:1,padding:"8px",borderRadius:8,border:`1px solid ${time===h?t.accent:t.border}`,background:time===h?t.acB:"transparent",color:time===h?t.accent:t.tm,fontSize:13,fontWeight:600,cursor:"pointer"}}>{h}</button>)}
      </div>
    </div>

    <div style={{marginBottom:12}}>
      <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:4}}>{t.L.controlParams}</div>
      <CheckRow lbl={t.L.paramEtiket} k="etiket"/>
      <CheckRow lbl={t.L.paramAcikGida} k="acikGida"/>
      <CheckRow lbl={t.L.paramYerdeDuran} k="yerdeDuran"/>
      <CheckRow lbl={t.L.paramCigPismis} k="cigPismis"/>
      <CheckRow lbl={t.L.paramTemizlik} k="temizlik"/>
      <CheckRow lbl={t.L.paramAlerjen} k="alerjen"/>
      <CheckRow lbl={t.L.paramSkt} k="skt"/>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
      <div>
        <label style={lSt(t)}>{t.L.productTemp}</label>
        <input type="number" style={iSt(t)} value={urunSicaklik} onChange={e=>setUrunSicaklik(e.target.value)} placeholder="-17.5"/>
      </div>
      <div>
        <label style={lSt(t)}>{t.L.gaugeValue}</label>
        <input type="number" style={iSt(t)} value={gostergeDeger} onChange={e=>setGostergeDeger(e.target.value)} placeholder="-18"/>
      </div>
    </div>

    <div style={{marginBottom:14}}>
      <label style={lSt(t)}>{t.L.explanation}</label>
      <input style={iSt(t)} value={aciklama} onChange={e=>setAciklama(e.target.value)} placeholder={t.L.explanationHint}/>
    </div>

    <div style={{display:"flex",gap:8}}>
      <button onClick={onClose} style={{...bSt("s",t),flex:1}}>{t.L.cancel}</button>
      <button onClick={save} style={{...bSt("p",t),flex:2}}>{t.L.save}</button>
    </div>
  </div></div>;
};

// ═══ ÜRETİM ═══
const ProductionTab=({productions,setProductions,storageAreas,reportCats,setReportCats,profile,initialShowReports,traceability,setTab,storageChecks,setStorageChecks,recipes,getLabelSeq,t})=>{
  const[selectedStorage,setSelectedStorage]=useState("all");
  const[activeProd,setActiveProd]=useState(null);
  const[showReports,setShowReports]=useState(initialShowReports||false);
  const[reportView,setReportView]=useState(()=>{
    if(typeof window!=="undefined"&&LS.get("tk_reportscrollstorage",null))return "storage";
    return "production";
  });
  const[showStorageCheckModal,setShowStorageCheckModal]=useState(null);
  const[partialAmount,setPartialAmount]=useState(0);
  const[trackSearch,setTrackSearch]=useState(""); // Takip No arama

  // Deep link event'ini dinle
  useEffect(()=>{
    const handler=(e)=>{if(e.detail&&e.detail.prod){setActiveProd(e.detail.prod);setShowReports(false);}};
    window.addEventListener("km-open-prod",handler);
    return()=>window.removeEventListener("km-open-prod",handler);
  },[]);

  // Sadece aktif (tüketilmemiş) üretimler
  const activeList=productions.filter(p=>p.status==="active");
  // Takip no araması
  const filtered=trackSearch.trim()
    ? productions.filter(p=>p.labelSeq&&p.labelSeq.includes(trackSearch.trim().toUpperCase()))
    : selectedStorage==="all"?activeList:activeList.filter(p=>p.storageId===selectedStorage);

  // Tarih formatlama
  const fmtDate=(iso)=>{if(!iso)return"—";const d=new Date(iso);return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getFullYear()).slice(-2)}`};
  const fmtTime=(iso)=>{if(!iso)return"";const d=new Date(iso);return `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`};
  const daysLeft=(iso)=>{if(!iso)return null;const now=new Date();const d=new Date(iso);const diff=Math.round((d-now)/(1000*60*60*24));return diff};

  // Tüket
  const consumeAll=(prod)=>{
    if(!window.confirm(tr_(t.lang,"consumedFull",{name:prod.recipeName,portions:prod.portions})))return;
    setProductions(p=>p.map(x=>x.id===prod.id?{...x,status:"consumed",consumedAt:new Date().toISOString(),reportCat:"production"}:x));
    setActiveProd(null);
  };
  const consumePartial=(prod,amount)=>{
    if(amount<=0||amount>prod.portions){window.toast.error(t.L.invalidAmount);return}
    const remaining=prod.portions-amount;
    if(remaining===0){consumeAll(prod);return}
    setProductions(p=>p.map(x=>x.id===prod.id?{...x,portions:remaining}:x));
    setPartialAmount(0);
    setActiveProd(null);
  };
  const markFire=(prod)=>{
    const amt=parseInt(window.prompt(tr_(t.lang,"howManyFire",{total:prod.portions}),10));
    if(!amt||amt<=0||amt>prod.portions)return;
    if(amt===prod.portions){
      setProductions(p=>p.map(x=>x.id===prod.id?{...x,status:"fire",fireAt:new Date().toISOString(),fireAmount:amt,reportCat:"fire"}:x));
      setActiveProd(null);
    }else{
      // Kısmi fire: üretimi böl
      const fireRec={...prod,id:Date.now(),portions:amt,status:"fire",fireAt:new Date().toISOString(),fireAmount:amt,reportCat:"fire"};
      setProductions(p=>[...p.map(x=>x.id===prod.id?{...x,portions:prod.portions-amt}:x),fireRec]);
    }
  };

  // Etiket yazdır
  const printLabel=(prod)=>{
    const isFreezer=prod.storageType==="freezer";
    const bg=isFreezer?"#1a56c4":"#ffffff";
    const tc=isFreezer?"#ffffff":"#111111";
    const borderC=isFreezer?"rgba(255,255,255,0.25)":"rgba(0,0,0,0.12)";
    const labelSeq=getLabelSeq();
    setProductions(p=>p.map(pr=>pr.id===prod.id?{...pr,labelSeq}:pr));
    const baseUrl=window.location.origin+window.location.pathname;
    const deepLink=`${baseUrl}?prod=${labelSeq}`;
    const lang=t.lang;
    const L={
      production:{tr:"Üretim",en:"Production",ru:"Производство",es:"Producción",de:"Herst.",fr:"Prod.",zh:"生产",ar:"إنتاج"}[lang]||"Production",
      opening:{tr:"Açılış",en:"Opening",ru:"Открытие",es:"Apertura",de:"Öffnung",fr:"Ouvert.",zh:"开封",ar:"فتح"}[lang]||"Opening",
      expiry:{tr:"SKT",en:"Exp",ru:"Срок",es:"Cad",de:"MHD",fr:"Exp",zh:"效期",ar:"انتهاء"}[lang]||"Exp",
      lot:{tr:"Lot",en:"Lot",ru:"Лот",es:"Lote",de:"Lot",fr:"Lot",zh:"批次",ar:"دفعة"}[lang]||"Lot",
      frozen:{tr:"Dond",en:"Frozen",ru:"Заморозка",es:"Congelado",de:"Eingefroren",fr:"Congelé",zh:"冷冻",ar:"تجميد"}[lang]||"Frozen",
      thaw:{tr:"Çözd",en:"Thaw",ru:"Разморозка",es:"Descongelado",de:"Auftauen",fr:"Décongel.",zh:"解冻",ar:"إذابة"}[lang]||"Thaw",
      trackNo:{tr:"№",en:"№",ru:"№",es:"№",de:"№",fr:"№",zh:"№",ar:"№"}[lang]||"№"
    };
    const prodDate=fmtDate(prod.producedAt);
    const expDate=fmtDate(prod.expiresAt);
    const w=window.open("","_blank","width=640,height=380");
    if(!w)return;
    w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>
@page{size:8cm 4cm;margin:0}
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:8cm;height:4cm;overflow:hidden;background:${bg};color:${tc};font-family:-apple-system,Arial,sans-serif}
.wrap{width:8cm;height:4cm;display:flex;flex-direction:column}
.main{display:flex;flex:1;padding:2.5mm 2.5mm 1.5mm 2.5mm;gap:2.5mm;overflow:hidden}
.qr-col{flex-shrink:0;width:28mm;display:flex;flex-direction:column;align-items:center;justify-content:center}
#qrc{display:block;width:28mm;height:28mm;image-rendering:pixelated}
.info-col{flex:1;min-width:0;display:flex;flex-direction:column;justify-content:center;gap:2px}
.pname{font-size:11px;font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border-bottom:0.5px solid ${borderC};padding-bottom:2px;margin-bottom:1px}
.row{display:flex;align-items:baseline;gap:3px;line-height:1.2}
.lbl{font-size:6px;font-weight:700;letter-spacing:0.05em;opacity:0.65;white-space:nowrap;flex-shrink:0}
.val{font-size:8.5px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.val.bold{font-weight:700}
.footer{height:5mm;border-top:0.5px solid ${borderC};display:flex;align-items:center;justify-content:space-between;padding:0 2.5mm;flex-shrink:0;background:${isFreezer?"rgba(0,0,0,0.15)":"rgba(0,0,0,0.03)"}}
.tno{font-size:6px;font-weight:700;letter-spacing:0.04em;opacity:0.55;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
@media print{html,body{width:8cm;height:4cm}}
</style></head><body>
<div class="wrap">
  <div class="main">
    <div class="qr-col"><canvas id="qrc" width="112" height="112"></canvas></div>
    <div class="info-col">
      <div class="pname">${(prod.recipeName||"").replace(/</g,"&lt;")}</div>
      ${isFreezer?`
      <div class="row"><span class="lbl">${L.frozen}</span><span class="val">${prodDate} ${fmtTime(prod.producedAt)}</span></div>
      <div class="row"><span class="lbl">${L.thaw}</span><span class="val">__/__/__ __:__</span></div>
      <div class="row"><span class="lbl">${L.expiry}</span><span class="val bold">${expDate}</span></div>
      <div class="row"><span class="lbl">${L.lot}</span><span class="val">${(prod.lotNumber||"-").replace(/</g,"&lt;")}</span></div>
      `:`
      <div class="row"><span class="lbl">${L.production}</span><span class="val">${prodDate}</span></div>
      <div class="row"><span class="lbl">${L.opening}</span><span class="val">__/__/____</span></div>
      <div class="row"><span class="lbl">${L.expiry}</span><span class="val bold">${expDate}</span></div>
      <div class="row"><span class="lbl">${L.lot}</span><span class="val">${(prod.lotNumber||"-").replace(/</g,"&lt;")}</span></div>
      `}
    </div>
  </div>
  <div class="footer">
    <span class="tno">${L.trackNo} ${labelSeq}</span>
  </div>
</div>
<script>
// Inline QR - Reed-Solomon tabanlı minimal QR encoder
// URL kısa tutmak için sadece takip no ve base URL
(function(){
  var url=${JSON.stringify(deepLink)};
  var canvas=document.getElementById("qrc");
  var ctx=canvas.getContext("2d");
  // QRCode.js CDN dene - başarısız olursa fallback göster
  function drawFallback(){
    ctx.fillStyle="${bg}";ctx.fillRect(0,0,112,112);
    ctx.fillStyle="${tc}";
    // Manuel QR benzeri border pattern çiz (görsel)
    ctx.strokeStyle="${tc}";ctx.lineWidth=2;
    ctx.strokeRect(4,4,104,104);
    ctx.strokeRect(8,8,30,30);ctx.fillRect(12,12,22,22);
    ctx.strokeRect(74,8,30,30);ctx.fillRect(78,12,22,22);
    ctx.strokeRect(8,74,30,30);ctx.fillRect(12,78,22,22);
    // Takip no yaz
    ctx.font="bold 7px monospace";ctx.textAlign="center";
    ctx.fillText("${labelSeq.split("-")[0]}",56,55);
    ctx.font="bold 8px monospace";
    ctx.fillText("${labelSeq.split("-")[1]||""}",56,66);
  }
  var s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js";
  s.onload=function(){
    var tmp=document.createElement("div");
    tmp.style.cssText="position:absolute;visibility:hidden;width:112px;height:112px";
    document.body.appendChild(tmp);
    try{
      new QRCode(tmp,{text:url,width:112,height:112,correctLevel:QRCode.CorrectLevel.M});
      setTimeout(function(){
        var img=tmp.querySelector("img");
        var cv=tmp.querySelector("canvas");
        if(cv){ctx.drawImage(cv,0,0,112,112);}
        else if(img){var i=new Image();i.onload=function(){ctx.drawImage(i,0,0,112,112);};i.src=img.src;}
        else drawFallback();
        document.body.removeChild(tmp);
      },200);
    }catch(e){drawFallback();}
  };
  s.onerror=drawFallback;
  document.head.appendChild(s);
  setTimeout(function(){window.print();},1400);
})();
<\/script>
</body></html>`);
    w.document.close();
  };

  // MÜŞTERİ QR'ı - URL tabanlı, overflow yok
  const printCustomerQR=(prod,recipe)=>{
    const lang=t.lang;
    const CL={
      ingredients: {tr:"İçindekiler",en:"Ingredients",ru:"Состав",es:"Ingredientes",de:"Zutaten",fr:"Ingrédients",zh:"成分",ar:"المكونات"}[lang]||"Ingredients",
      allergens:   {tr:"Alerjenler",en:"Allergens",ru:"Аллергены",es:"Alérgenos",de:"Allergene",fr:"Allergènes",zh:"过敏原",ar:"مسببات الحساسية"}[lang]||"Allergens",
      expiry:      {tr:"Son Kullanma",en:"Best Before",ru:"Годен до",es:"Fecha cad",de:"MHD",fr:"À consommer avant",zh:"保质期",ar:"صالح حتى"}[lang]||"Best Before",
      producer:    {tr:"Üretici",en:"Producer",ru:"Производитель",es:"Productor",de:"Hersteller",fr:"Producteur",zh:"生产商",ar:"المنتج"}[lang]||"Producer",
      scanHint:    {tr:"QR'ı okutun — içerik, alerjenler ve SKT görüntülenir",en:"Scan QR to view ingredients, allergens & expiry",ru:"Сканируйте для просмотра состава",es:"Escanee para ver detalles",de:"Scannen für Produktdetails",fr:"Scannez pour voir les détails",zh:"扫描查看产品信息",ar:"امسح لعرض المكونات"}[lang]||"Scan to view",
      customer:    {tr:"MÜŞTERİ",en:"CUSTOMER",ru:"КЛИЕНТ",es:"CLIENTE",de:"KUNDE",fr:"CLIENT",zh:"客户",ar:"العميل"}[lang]||"CUSTOMER",
      noAllergen:  {tr:"Alerjen yok",en:"No allergens",ru:"Без аллергенов",es:"Sin alérgenos",de:"Keine Allergene",fr:"Sans allergènes",zh:"无过敏原",ar:"لا توجد مسببات"}[lang]||"No allergens"
    };
    const ingList=(recipe?.ingredients||[]).map(i=>i.name).filter(Boolean);
    const allergenList=(recipe?.allergens||[]);
    const allerNames=allergenList.map(a=>{const al=ALLERGENS.find(x=>x.id===a);return al?allergenL(al,lang):a;}).filter(Boolean);
    const producer=(profile&&profile.workplace)||"";

    // QR içeriği sadece URL — overflow yok, her cihazda çalışır
    const prodId=prod.id||prod.lot||Date.now();
    const qrContent=`https://tulparkitchen.com/qr?id=${prodId}`;

    const showModal=()=>{
      const allergenHTML=allerNames.length
        ?`<div style="background:#fff0f0;border:1px solid #ffcccc;border-radius:10px;padding:10px 14px;margin-bottom:10px">
            <div style="font-size:11px;font-weight:800;color:#c00;letter-spacing:0.08em;margin-bottom:5px">⚠️ ${CL.allergens.toUpperCase()}</div>
            <div style="font-size:13px;font-weight:600;color:#900;line-height:1.6">${allerNames.join(" · ")}</div>
          </div>`
        :`<div style="background:#f0fff4;border:1px solid #bbf7d0;border-radius:10px;padding:8px 14px;margin-bottom:10px;font-size:12px;color:#166534;font-weight:600">✓ ${CL.noAllergen}</div>`;
      const ingHTML=ingList.length
        ?`<div style="margin-bottom:10px">
            <div style="font-size:11px;font-weight:700;color:#555;letter-spacing:0.06em;margin-bottom:5px">${CL.ingredients.toUpperCase()}</div>
            <div style="font-size:12px;color:#333;line-height:1.7">${ingList.join(", ")}</div>
          </div>`:"";
      const producerHTML=producer?`<div style="font-size:11px;color:#777;margin-bottom:6px">${CL.producer}: <strong>${producer}</strong></div>`:"";

      const overlay=document.createElement("div");
      overlay.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,0.7);z-index:99999;display:flex;align-items:center;justify-content:center;padding:20px";
      overlay.innerHTML=`<div style="background:#fff;border-radius:16px;padding:24px;max-width:400px;width:100%;color:#111;font-family:-apple-system,Arial,sans-serif;max-height:90vh;overflow-y:auto">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
          <strong style="font-size:16px">📱 ${CL.customer} QR</strong>
          <button id="kmCloseQR" style="background:none;border:none;font-size:22px;cursor:pointer;color:#888;padding:0 8px">✕</button>
        </div>
        <div id="qrcode" style="width:220px;height:220px;margin:0 auto 14px;display:flex;align-items:center;justify-content:center;background:#f5f5f5;border-radius:8px"></div>
        <div style="font-size:15px;font-weight:700;text-align:center;margin-bottom:4px">${prod.recipeName}</div>
        <div style="font-size:12px;color:#c00;font-weight:700;text-align:center;margin-bottom:14px">📅 ${CL.expiry}: ${fmtDate(prod.expiresAt)}</div>
        <div style="border-top:1px solid #eee;padding-top:12px;margin-bottom:12px">
          ${allergenHTML}
          ${ingHTML}
          ${producerHTML}
        </div>
        <div style="font-size:10px;color:#aaa;text-align:center;margin-bottom:12px">${CL.scanHint}</div>
        <button id="kmPrintQR" style="width:100%;padding:12px;background:#c8965a;color:#fff;border:none;border-radius:10px;font-weight:700;cursor:pointer;font-size:14px">🖨 ${lang==="tr"?"Yazdır":"Print"}</button>
      </div>`;
      document.body.appendChild(overlay);
      overlay.addEventListener("click",(e)=>{if(e.target===overlay||e.target.id==="kmCloseQR")overlay.remove();});

      const renderQR=()=>{
        const target=overlay.querySelector("#qrcode");
        if(!target||!window.QRCode){target&&(target.innerHTML='<span style="font-size:11px;color:#888">QR kütüphanesi yüklenemedi</span>');return;}
        try{
          target.innerHTML="";
          new window.QRCode(target,{text:qrContent,width:220,height:220,correctLevel:window.QRCode.CorrectLevel.M});
        }catch(e){target.innerHTML='<span style="font-size:11px;color:#c00">QR oluşturulamadı: '+e.message+'</span>';}
      };
      if(window.QRCode){renderQR();}
      else{
        const s=document.createElement("script");
        s.src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js";
        s.onload=renderQR;
        s.onerror=()=>{const t2=overlay.querySelector("#qrcode");if(t2)t2.innerHTML='<span style="font-size:11px;color:#c00">İnternet yok</span>';};
        document.head.appendChild(s);
      }

      overlay.querySelector("#kmPrintQR").onclick=()=>{
        const qrEl=overlay.querySelector("#qrcode img,#qrcode canvas");
        if(!qrEl){window.toast.info(lang==="tr"?"QR henüz hazır değil, biraz bekleyin":"QR not ready yet");return;}
        let qrSrc="";
        try{qrSrc=qrEl.tagName==="IMG"?qrEl.src:qrEl.toDataURL("image/png");}
        catch(e){window.toast.info("QR alınamadı: "+e.message);return;}
        const allergenPrintHTML=allerNames.length
          ?`<div style="background:#fff0f0;border:1.5px solid #ffaaaa;border-radius:6px;padding:4px 7px;margin-bottom:3px"><span style="font-size:6.5px;font-weight:800;color:#c00">⚠ ${CL.allergens.toUpperCase()}: </span><span style="font-size:7px;font-weight:700;color:#900">${allerNames.join(" · ")}</span></div>`
          :`<div style="font-size:6.5px;color:#166534;font-weight:600">✓ ${CL.noAllergen}</div>`;
        const ingPrintHTML=ingList.length?`<div style="font-size:6.5px;color:#444;margin-top:3px;line-height:1.4"><span style="font-weight:700">${CL.ingredients}: </span>${ingList.join(", ")}</div>`:"";
        const printW=window.open("","_blank","width=600,height=400");
        if(!printW){
          const a=document.createElement("a");a.href=qrSrc;a.download=`qr-${prod.recipeName||"product"}.png`;a.click();
          return;
        }
        printW.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>@page{size:8cm 4cm;margin:0}*{margin:0;box-sizing:border-box;font-family:-apple-system,Arial,sans-serif}html,body{width:8cm;height:4cm;overflow:hidden;background:#fff;color:#111;font-size:8px;line-height:1.3}.wrap{display:flex;align-items:center;gap:3mm;padding:3mm;height:4cm}.qr-col{flex-shrink:0;display:flex;flex-direction:column;align-items:center;gap:1px}.qr-col img{width:28mm;height:28mm}.qr-lbl{font-size:5.5px;font-weight:700;letter-spacing:0.06em;margin-top:1px;opacity:0.7}.info{flex:1;min-width:0;display:flex;flex-direction:column;gap:2px}.name{font-size:10px;font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.exp{font-size:8.5px;font-weight:700;color:#c00}.hint{font-size:6px;color:#777;line-height:1.3;margin-top:2px}</style></head><body><div class="wrap"><div class="qr-col"><img src="${qrSrc}"/><div class="qr-lbl">📱 ${CL.customer}</div></div><div class="info"><div class="name">${prod.recipeName}</div><div class="exp">📅 ${CL.expiry}: ${fmtDate(prod.expiresAt)}</div>${allergenPrintHTML}${ingPrintHTML}<div class="hint">${CL.scanHint}</div></div></div><script>setTimeout(function(){window.print();},400);<\/script></body></html>`);
        printW.document.close();
      };
    };
    showModal();
  };

  // Raporlar görünümü
  if(showReports){
    // TÜM kayıtları günlere göre grupla (aktif + tamamlanan)
    const allRecords=productions;
    const byDay={};
    allRecords.forEach(p=>{
      const d=new Date(p.producedAt);
      const dayKey=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
      if(!byDay[dayKey])byDay[dayKey]={date:d,productions:[],samples:[],fires:[]};
      if(p.reportCat==="samples")byDay[dayKey].samples.push(p);
      else if(p.status==="fire")byDay[dayKey].fires.push(p);
      else byDay[dayKey].productions.push(p);
    });
    const sortedDays=Object.keys(byDay).sort().reverse();

    // FR.06 PDF oluştur
    const generateFR06=(dayKey)=>{
      const day=byDay[dayKey];
      const dateStr=fmtDate(day.date.toISOString());
      const workplace=(profile&&profile.workplace)||"";
      const dept=(profile&&profile.department)||t.L.fr06DefaultDept;
      const prodList=day.productions;
      const L=t.L;
      const w=window.open("","","width=1100,height=800");
      let html=`<html><head><meta charset="utf-8"><style>
@page{size:A4;margin:10mm}
*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
body{font-size:8px;color:#000;padding:8px}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;border:1px solid #000;padding:4px 8px}
.hleft{flex:1;font-size:10px;font-weight:700}
.hcenter{flex:2;text-align:center}
.hright{border:1px solid #000;font-size:7px}
.hright td{padding:2px 6px;border:1px solid #000}
h1{font-size:14px;font-weight:700}
table{width:100%;border-collapse:collapse;margin-bottom:6px;font-size:7px}
th,td{border:1px solid #000;padding:2px 3px;text-align:left;vertical-align:top}
th{background:#f0f0f0;font-weight:700;text-align:center;font-size:7px}
.section-title{background:#ddd;padding:3px;font-weight:700;font-size:9px;text-align:center;border:1px solid #000;border-bottom:none}
.empty-row td{height:16px}
.prod-name{font-weight:700;font-size:8px;margin-bottom:1px}
.ing-table{width:100%;border-collapse:collapse;margin-top:2px;font-size:6.5px;table-layout:fixed}
.ing-table td{border:1px dotted #999;padding:1px 2px;vertical-align:middle}
.ing-name{width:15%;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ing-lot{width:10%;text-align:center;font-weight:600;border-left:none!important}
.notes{font-size:6px;margin-top:8px;border-top:1px solid #000;padding-top:4px;line-height:1.3}
.notes-title{font-weight:700;font-size:7px;margin-bottom:2px}
@media print{body{margin:0;padding:5mm}}
</style></head><body>
<div class="header">
<div class="hleft">${dept}${workplace?" — "+workplace:""}</div>
<div class="hcenter"><h1>${L.fr06Title}</h1></div>
<div class="hright"><table>
<tr><td>${L.fr06DocNo}</td><td>FR.06</td></tr>
<tr><td>${L.fr06DateLbl}</td><td><strong>${dateStr}</strong></td></tr>
</table></div>
</div>

<div class="section-title">${L.fr06ProductCtrl}</div>
<table>
<thead><tr>
<th style="width:62%">${L.fr06ProductIng}</th>
<th style="width:7%">${L.fr06PrepStart}</th>
<th style="width:7%">${L.fr06PrepEnd}</th>
<th style="width:7%">${L.fr06ProductTemp}</th>
<th style="width:9%">${L.fr06AmbientTemp}</th>
<th style="width:11%">${L.fr06Controller}</th>
</tr></thead>
<tbody>
${prodList.map(p=>{
  const prodTime=new Date(p.producedAt);
  const timeStr=`${String(prodTime.getHours()).padStart(2,"0")}:${String(prodTime.getMinutes()).padStart(2,"0")}`;
  const allIngs=[];
  (p.ingredientLots||[]).forEach(il=>{
    allIngs.push({name:il.name||il.tracked,lot:il.lotNo||""});
  });
  const ingRows=[];
  for(let i=0;i<allIngs.length;i+=3){
    const a=allIngs[i],b=allIngs[i+1],c=allIngs[i+2];
    ingRows.push(`<tr>
<td class="ing-name">${a.name}</td><td class="ing-lot">${a.lot||"____"}</td>
<td class="ing-name">${b?b.name:""}</td><td class="ing-lot">${b?(b.lot||"____"):""}</td>
<td class="ing-name">${c?c.name:""}</td><td class="ing-lot">${c?(c.lot||"____"):""}</td>
</tr>`);
  }
  const ingHtml=allIngs.length>0?`<table class="ing-table"><tbody>${ingRows.join("")}</tbody></table>`:"";
  return `<tr>
<td>
<div class="prod-name">${p.recipeName}${p.portions?" ("+p.portions+" por.)":""} · <span style="font-weight:400;font-size:7px">${L.fr06PartyNo}: ${p.lotNumber||""}</span></div>
${ingHtml}
</td>
<td>${timeStr}</td>
<td></td>
<td></td>
<td></td>
<td>${p.producedBy||""}</td>
</tr>`;
}).join("")}
${Array.from({length:Math.max(0,6-prodList.length)}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}
</tbody>
</table>

<div class="section-title">${L.fr06CookCtrl}</div>
<table>
<thead><tr>
<th>${L.fr06ProductDef}</th><th>${L.fr06PartyNo}</th><th>${L.fr06Cook}</th><th>${L.fr06Reheat}</th><th>${L.fr06Controller}</th><th>${L.fr06CorrectiveAction}</th>
</tr></thead>
<tbody>
${Array.from({length:5}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}
</tbody>
</table>

<div class="section-title">${L.fr06FreezingCtrl}</div>
<table>
<thead><tr>
<th rowspan="2">${L.fr06ProductDef}</th>
<th colspan="3">${L.fr06ThawStart}</th>
<th colspan="3">${L.fr06ThawEnd}</th>
<th rowspan="2">${L.fr06Controller}</th>
<th rowspan="2">${L.fr06CorrectiveAction}</th>
</tr>
<tr><th>${L.fr06PartyNo}</th><th>${L.time}</th><th>${L.fr06Temp}</th><th>${L.date}</th><th>${L.time}</th><th>${L.fr06Temp}</th></tr>
</thead>
<tbody>${Array.from({length:3}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}</tbody>
</table>

<div class="section-title">${L.fr06FastCoolCtrl}</div>
<table>
<thead><tr>
<th>${L.fr06ProductDef}</th><th>${L.fr06Entry}</th><th>${L.fr06Exit}</th><th>${L.fr06Chiller}</th><th>${L.fr06IceWater}</th><th>${L.fr06Controller}</th><th>${L.fr06CorrectiveAction}</th>
</tr></thead>
<tbody>${Array.from({length:3}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}</tbody>
</table>

<div style="display:flex;gap:6px">
<div style="flex:1">
<div class="section-title">${L.fr06PlateHold}</div>
<table>
<thead><tr><th>${L.fr06ProductOrg}</th><th>${L.fr06PlateTime}</th><th>${L.fr06Service}</th><th>${L.fr06InnerTemp}</th><th>${L.fr06Ctrl}</th></tr></thead>
<tbody>${Array.from({length:3}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}</tbody>
</table>
</div>
<div style="flex:1">
<div class="section-title">${L.fr06FosterHold}</div>
<table>
<thead><tr><th>${L.recipeName}</th><th>${L.fr06Entry}</th><th>${L.fr06Exit}</th><th>${L.fr06Ctrl}</th></tr></thead>
<tbody>${Array.from({length:3}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td></tr>`).join("")}</tbody>
</table>
</div>
</div>

<div class="notes">
<div class="notes-title">${L.fr06Notes}</div>
<strong>${L.fr06NoteRisk}</strong> ${L.fr06NoteRiskText}<br>
<strong>${L.fr06NoteCool}</strong> ${L.fr06NoteCoolText}<br>
<strong>${L.fr06NoteCook}</strong> ${L.fr06NoteCookText}<br>
<strong>${L.fr06NotePlate}</strong> ${L.fr06NotePlateText}<br>
<strong>${L.fr06NoteThaw}</strong> ${L.fr06NoteThawText}
</div>
</body></html>`;
      w.document.write(html);w.document.close();
      setTimeout(()=>w.print(),600);
    };

    // FR.12 Şahit Numune PDF oluştur
    const generateFR12=(dayKey)=>{
      const day=byDay[dayKey];
      const dateStr=fmtDate(day.date.toISOString());
      const workplace=(profile&&profile.workplace)||"";
      const dept=(profile&&profile.department)||t.L.fr06DefaultDept;
      const samples=day.samples;
      const L=t.L;
      const w=window.open("","","width=1100,height=800");
      let html=`<html><head><meta charset="utf-8"><style>
@page{size:A4;margin:10mm}
*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
body{font-size:10px;color:#000;padding:8px}
.header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;border:1px solid #000;padding:6px}
.hleft{flex:1}
.hcenter{flex:2;text-align:center}
.hright{border:1px solid #000;padding:3px;font-size:8px}
.hright td{padding:2px 6px;border:1px solid #000}
h1{font-size:16px;font-weight:700}
.info{margin-bottom:8px;font-size:11px;font-weight:600}
table{width:100%;border-collapse:collapse;font-size:10px}
th,td{border:1px solid #000;padding:5px 6px;text-align:left;vertical-align:top}
th{background:#f0f0f0;font-weight:700;text-align:center}
.empty-row td{height:26px}
@media print{body{margin:0;padding:5mm}}
</style></head><body>
<div class="header">
<div class="hleft"><strong>${dept}</strong>${workplace?" — "+workplace:""}</div>
<div class="hcenter"><h1>${L.fr12Title}</h1></div>
<div class="hright"><table>
<tr><td>${L.fr06DocNo}</td><td>FR.12</td></tr>
<tr><td>${L.fr06DateLbl}</td><td><strong>${dateStr}</strong></td></tr>
</table></div>
</div>

<table>
<thead><tr>
<th style="width:20%">${L.fr12SampleOrg}</th>
<th style="width:8%">${L.fr12Pax}</th>
<th style="width:15%">${L.fr12DateTime}</th>
<th style="width:10%">${L.fr12Temp}</th>
<th style="width:15%">${L.fr12Taker}</th>
<th style="width:12%">${L.fr12Signature}</th>
<th style="width:20%">${L.fr12Explanation}</th>
</tr></thead>
<tbody>
${samples.map(s=>{
  const sd=new Date(s.sampleDateTime||s.producedAt);
  const timeStr=`${String(sd.getDate()).padStart(2,"0")}.${String(sd.getMonth()+1).padStart(2,"0")}.${String(sd.getFullYear()).slice(-2)} ${String(sd.getHours()).padStart(2,"0")}:${String(sd.getMinutes()).padStart(2,"0")}`;
  const orgDisplay=s.sampleOrganization||s.sampleLocation||"";
  return `<tr>
<td><strong>${s.recipeName}</strong>${orgDisplay?"<br><small>"+orgDisplay+"</small>":""}</td>
<td style="text-align:center">${s.samplePax||1}</td>
<td>${timeStr}</td>
<td style="text-align:center">+${s.sampleTemp||"__"} °C</td>
<td>${s.takerName||""}</td>
<td></td>
<td></td>
</tr>`;
}).join("")}
${Array.from({length:Math.max(0,12-samples.length)}).map(()=>`<tr class="empty-row"><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join("")}
</tbody>
</table>

<div style="font-size:8px;margin-top:10px;border-top:1px solid #000;padding-top:5px;line-height:1.5">
<strong>${L.fr12Note}</strong> ${L.fr12NoteText}
</div>
</body></html>`;
      w.document.write(html);w.document.close();
      setTimeout(()=>w.print(),600);
    };

    // FR.05 Depo Kontrol PDF oluştur
    const generateFR05=(dayKey,storageId)=>{
      const storage=(storageAreas||[]).find(s=>s.id===storageId);
      if(!storage)return;
      const typ=STORAGE_TYPES.find(x=>x.id===storage.type);
      const dateObj=new Date(dayKey);
      const dateStr=fmtDate(dateObj.toISOString());
      const workplace=(profile&&profile.workplace)||"";
      const dept=(profile&&profile.department)||t.L.fr06DefaultDept;
      const L=t.L;
      // O gün bu depoya ait tüm kontroller
      const checks=storageChecks.filter(c=>c.storageId===storageId&&c.date===dayKey).sort((a,b)=>a.time.localeCompare(b.time));
      const w=window.open("","","width=1100,height=800");
      let html=`<html><head><meta charset="utf-8"><style>
@page{size:A4;margin:10mm}
*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
body{font-size:9px;color:#000;padding:8px}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;border:1px solid #000;padding:5px 8px}
.hleft{flex:1;font-size:10px;font-weight:700}
.hcenter{flex:2;text-align:center}
.hright{border:1px solid #000;font-size:8px}
.hright td{padding:2px 6px;border:1px solid #000}
h1{font-size:15px;font-weight:700}
.info{margin-bottom:8px;font-size:11px;font-weight:600;background:#f0f0f0;padding:6px;border:1px solid #000}
table{width:100%;border-collapse:collapse;font-size:8px}
th,td{border:1px solid #000;padding:3px 4px;text-align:center;vertical-align:middle}
th{background:#f0f0f0;font-weight:700;font-size:8px}
.empty-row td{height:22px}
.tick{font-size:10px;font-weight:700}
.notes{font-size:7px;margin-top:8px;border-top:1px solid #000;padding-top:5px;line-height:1.4}
@media print{body{margin:0;padding:5mm}}
</style></head><body>
<div class="header">
<div class="hleft">${dept}${workplace?" — "+workplace:""}</div>
<div class="hcenter"><h1>${L.fr05Title}</h1></div>
<div class="hright"><table>
<tr><td>${L.fr06DocNo}</td><td>FR.05</td></tr>
<tr><td>${L.fr06DateLbl}</td><td><strong>${dateStr}</strong></td></tr>
</table></div>
</div>

<div class="info">${L.fr05Dept} ${dept} · ${L.fr05Cabinet} ${typ?.icon||""} <strong>${storage.name}</strong> (${storageTypeL(typ,t.lang)}) · ${L.fr05Target} ${storage.temp}°C</div>

<table>
<thead>
<tr>
<th colspan="8" style="background:#ddd">${L.fr05CtrlParams}</th>
<th colspan="3" style="background:#ddd">${L.fr05TempCtrl}</th>
<th rowspan="2">${L.fr12Explanation}</th>
<th rowspan="2">${L.fr06Controller}</th>
</tr>
<tr>
<th>${L.fr06DateLbl}</th>
<th>${L.paramEtiket}</th>
<th>${L.paramAcikGida}</th>
<th>${L.paramYerdeDuran}</th>
<th>${L.paramCigPismis}</th>
<th>${L.paramTemizlik}</th>
<th>${L.paramAlerjen}</th>
<th>${L.paramSkt}</th>
<th>${L.fr05CtrlTime}</th>
<th>${L.productTemp}</th>
<th>${L.gaugeValue}</th>
</tr>
</thead>
<tbody>
${["09.00","15.00","21.00"].map((tt,idx)=>{
  const c=checks.find(x=>x.time===tt);
  const v=(k)=>c?(c[k]?"<span class='tick'>✓</span>":"<span class='tick'>✗</span>"):"";
  const tD=(k)=>c?(c[k]||""):"";
  return `<tr>
<td>${idx===0?dateStr:""}</td>
<td>${v("etiket")}</td>
<td>${v("acikGida")}</td>
<td>${v("yerdeDuran")}</td>
<td>${v("cigPismis")}</td>
<td>${v("temizlik")}</td>
<td>${v("alerjen")}</td>
<td>${v("skt")}</td>
<td>${tt}</td>
<td>${tD("urunSicaklik")}</td>
<td>${tD("gostergeDeger")}</td>
<td>${tD("aciklama")}</td>
<td>${tD("kontrolEden")}</td>
</tr>`;
}).join("")}
</tbody>
</table>

<div class="notes">
<strong>${L.fr05NoteDry}</strong> ${L.fr05NoteDryText}<br>
<strong>${L.fr05NoteCold}</strong> ${L.fr05NoteColdText}<br>
<strong>${L.fr05NoteApp}</strong> ${L.fr05NoteAppText}
</div>
</body></html>`;
      w.document.write(html);w.document.close();
      setTimeout(()=>w.print(),600);
    };

    return <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h3 style={{fontSize:22,color:t.text,margin:0}}>📊 {t.L.reports}</h3>
        {!initialShowReports&&<button onClick={()=>setShowReports(false)} style={{...bSt("s",t),fontSize:13}}>← {t.L.back}</button>}
      </div>

      {/* Alt-tab switcher */}
      <div style={{display:"flex",gap:6,marginBottom:14,background:t.inBg,padding:4,borderRadius:12}}>
        <button onClick={()=>setReportView("production")} style={{flex:1,padding:"8px 12px",borderRadius:10,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:reportView==="production"?t.card:"transparent",color:reportView==="production"?t.text:t.tm}}>{t.L.reportsProduction}</button>
        <button onClick={()=>setReportView("storage")} style={{flex:1,padding:"8px 12px",borderRadius:10,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:reportView==="storage"?t.card:"transparent",color:reportView==="storage"?t.text:t.tm}}>{t.L.reportsStorage}</button>
        <button onClick={()=>setReportView("archive")} style={{flex:1,padding:"8px 12px",borderRadius:10,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:reportView==="archive"?t.card:"transparent",color:reportView==="archive"?t.text:t.tm}}>{t.L.archiveTab||"📅 Arşiv"}</button>
      </div>

      {/* ARŞİV VIEW - Aylık grupla geçmiş aylar */}
      {reportView==="archive"&&(()=>{
        // Tüm kayıtları ay-yıl bazında grupla
        const monthGroups={};
        productions.forEach(p=>{
          const d=new Date(p.producedAt);
          const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
          if(!monthGroups[key])monthGroups[key]={year:d.getFullYear(),month:d.getMonth()+1,prods:[],samples:[],fires:[]};
          if(p.reportCat==="samples")monthGroups[key].samples.push(p);
          else if(p.status==="fire")monthGroups[key].fires.push(p);
          else monthGroups[key].prods.push(p);
        });
        // Bu ayı dahil etme (arşiv = geçmiş aylar)
        const now=new Date();
        const thisKey=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`;
        const archiveKeys=Object.keys(monthGroups).filter(k=>k!==thisKey).sort().reverse();
        const monthName=(m,lang)=>{
          const names={
            tr:["","Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"],
            en:["","January","February","March","April","May","June","July","August","September","October","November","December"],
            ru:["","Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"],
            es:["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],
            de:["","Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"],
            fr:["","Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"],
            zh:["",...Array.from({length:12},(_,i)=>(i+1)+"月")],
            ar:["","يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
          };
          return(names[t.lang]||names.tr)[m];
        };
        const exportMonthPDF=(key)=>{
          const g=monthGroups[key];
          const L=t.L;
          const dept=(profile&&profile.department)||L.fr06DefaultDept;
          const workplace=(profile&&profile.workplace)||"";
          const w=window.open("","","width=900,height=700");
          const allRecs=[...g.prods,...g.samples,...g.fires].sort((a,b)=>new Date(a.producedAt)-new Date(b.producedAt));
          w.document.write(`<html><head><meta charset="utf-8"><style>
@page{size:A4;margin:12mm}*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
body{font-size:9px;color:#000;padding:8px}
.hdr{display:flex;justify-content:space-between;align-items:center;border:1px solid #000;padding:5px 8px;margin-bottom:8px}
h1{font-size:14px;font-weight:700;text-align:center}
table{width:100%;border-collapse:collapse;margin-bottom:10px;font-size:8px}
th,td{border:1px solid #000;padding:3px 4px;text-align:left}th{background:#f0f0f0;font-weight:700}
.section{font-weight:700;font-size:10px;margin:6px 0 3px;border-left:3px solid #333;padding-left:6px}
@media print{body{padding:5mm}}
</style></head><body>
<div class="hdr"><div>${dept}${workplace?" — "+workplace:""}</div><h1>📊 ${monthName(g.month,t.lang)} ${g.year} — ${L.archiveTitle||"Aylık Rapor Arşivi"}</h1><div>${L.fr06DateLbl}: ${String(g.month).padStart(2,"0")}/${g.year}</div></div>
<div style="display:flex;gap:12px;margin-bottom:8px;font-size:10px">
  <span>✅ ${L.productionsCap}: <strong>${g.prods.length}</strong></span>
  <span>🧪 ${L.samplesCap}: <strong>${g.samples.length}</strong></span>
  <span>🗑 ${L.firesCap}: <strong>${g.fires.length}</strong></span>
  <span>📋 <strong>${allRecs.length}</strong></span>
</div>
${g.prods.length>0?`<div class="section">✅ ${L.productionsCap}</div>
<table><thead><tr><th>${L.date}</th><th>${L.recipeName||"Recipe"}</th><th>${L.lot||"Lot"}</th><th>${L.portions||"Portions"}</th><th>${L.storage||"Storage"}</th><th>${L.fr06Controller}</th><th>${L.note||"Note"}</th></tr></thead><tbody>
${g.prods.map(p=>{const d=new Date(p.producedAt);return`<tr><td>${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}</td><td>${p.recipeName}</td><td>${p.lotNumber||""}</td><td>${p.portions}</td><td>${p.storageName||""}</td><td>${p.producedBy||""}</td><td>${p.note||""}</td></tr>`}).join("")}
</tbody></table>`:""}
${g.samples.length>0?`<div class="section">🧪 ${L.samplesCap}</div>
<table><thead><tr><th>${L.date}</th><th>${L.recipeName||"Recipe"}</th><th>${L.sampleOrgLabel||"Org"}</th><th>${L.samplePax||"Pax"}</th><th>${L.sampleTaker||"Taker"}</th></tr></thead><tbody>
${g.samples.map(p=>{const d=new Date(p.producedAt);return`<tr><td>${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}</td><td>${p.recipeName}</td><td>${p.sampleOrganization||p.sampleLocation||""}</td><td>${p.samplePax||1}</td><td>${p.takerName||""}</td></tr>`}).join("")}
</tbody></table>`:""}
${g.fires.length>0?`<div class="section">🗑 ${L.firesCap}</div>
<table><thead><tr><th>${L.date}</th><th>${L.recipeName||"Recipe"}</th><th>${L.portions||"Portions"}</th></tr></thead><tbody>
${g.fires.map(p=>{const d=new Date(p.producedAt);return`<tr><td>${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}</td><td>${p.recipeName}</td><td>${p.fireAmount||p.portions}</td></tr>`}).join("")}
</tbody></table>`:""}
</body></html>`);
          w.document.close();setTimeout(()=>w.print(),500);
        };
        return <div>
          <h3 style={{fontSize:18,color:t.text,marginBottom:4}}>{t.L.archiveTitle||"Aylık Rapor Arşivi"}</h3>
          <div style={{fontSize:12,color:t.tm,marginBottom:14,lineHeight:1.5}}>{t.L.archiveHint||"💡 Her ay sonunda raporlar otomatik olarak burada arşivlenir."}</div>
          {archiveKeys.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
            <div style={{fontSize:40,opacity:0.4,marginBottom:10}}>📅</div>
            <div>{t.L.archiveEmpty||"Henüz arşivlenmiş rapor yok"}</div>
            <div style={{fontSize:11,marginTop:4,opacity:0.6}}>{t.lang==="tr"?"Bu ay bittikten sonra kayıtlar burada görünecek":t.lang==="en"?"Records appear here after month ends":""}</div>
          </div>:archiveKeys.map(key=>{
            const g=monthGroups[key];
            const total=g.prods.length+g.samples.length+g.fires.length;
            return <div key={key} style={{...cSt(t),padding:"14px 16px",marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:16,fontWeight:700,color:t.text}}>📅 {monthName(g.month,t.lang)} {g.year}</div>
                  <div style={{fontSize:12,color:t.tm,marginTop:4}}>
                    ✅ {g.prods.length} · 🧪 {g.samples.length} · 🗑 {g.fires.length} · {total} {t.L.archiveTotal||"kayıt"}
                  </div>
                </div>
                <button onClick={()=>exportMonthPDF(key)} style={{...bSt("p",t),fontSize:11,padding:"6px 12px"}}>{t.L.archiveExport||"📥 PDF"}</button>
              </div>
            </div>;
          })}
        </div>;
      })()}

      {reportView==="storage"&&<div>
        <div style={{fontSize:12,color:t.tm,marginBottom:10,lineHeight:1.5}}>{t.L.storageCheckHint}</div>
        {(storageAreas||[]).map(s=>{
          const typ=STORAGE_TYPES.find(x=>x.id===s.type);
          // Bugünün kontrolleri
          const today=(()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`})();
          const todayChecks=storageChecks.filter(c=>c.storageId===s.id&&c.date===today);
          // Son 7 gün grupla
          const byDate={};
          storageChecks.filter(c=>c.storageId===s.id).forEach(c=>{
            if(!byDate[c.date])byDate[c.date]=[];
            byDate[c.date].push(c);
          });
          const dates=Object.keys(byDate).sort().reverse().slice(0,7);
          const scrollTo=LS.get("tk_reportscrollstorage",null);
          const highlight=scrollTo===s.id;
          return <div key={s.id} ref={el=>{if(el&&highlight){setTimeout(()=>{el.scrollIntoView({behavior:"smooth",block:"center"});LS.set("tk_reportscrollstorage",null)},100)}}} style={{...cSt(t),padding:"14px 16px",marginBottom:10,border:highlight?`2px solid ${t.accent}`:undefined,boxShadow:highlight?`0 0 0 4px ${t.acB}`:undefined,transition:"all 0.3s"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:t.text}}>{typ?.icon} {s.name}</div>
                <div style={{fontSize:11,color:t.tm,marginTop:2}}>{storageTypeL(typ,t.lang)} · {s.temp}°C · {t.lang==="tr"?"Bugün:":t.lang==="en"?"Today:":t.lang==="ru"?"Сегодня:":t.lang==="de"?"Heute:":t.lang==="fr"?"Auj.":"Today:"} {todayChecks.length}/3</div>
              </div>
              <button onClick={()=>setShowStorageCheckModal({storageId:s.id,date:today})} style={{...bSt("p",t),fontSize:11,padding:"6px 10px"}}>{t.L.addCheck||"+ Add Check"}</button>
            </div>
            {dates.length===0?<div style={{fontSize:11,color:t.tm,fontStyle:"italic"}}>Henüz kontrol girilmedi.</div>:dates.map(d=>{
              const dChecks=byDate[d];
              const timeSet=new Set(dChecks.map(c=>c.time));
              const complete=timeSet.has("09.00")&&timeSet.has("15.00")&&timeSet.has("21.00");
              return <div key={d} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,padding:"6px 0",borderTop:`1px solid ${t.border}`,color:t.ts}}>
                <div>
                  <span>{fmtDate(new Date(d).toISOString())}</span>
                  <span style={{color:complete?t.success:t.warn,marginLeft:6,fontSize:10,fontWeight:600}}>{complete?"✓ Tam":`${dChecks.length}/3`}</span>
                </div>
                <button onClick={()=>generateFR05(d,s.id)} style={{...bSt("s",t),fontSize:10,padding:"4px 8px"}}>📄 FR.05</button>
              </div>;
            })}
          </div>;
        })}
      </div>}

      {reportView==="production"&&<>
      {sortedDays.length===0&&<div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
        <div style={{fontSize:48,opacity:0.4,marginBottom:12}}>📊</div>
        <div style={{fontSize:14}}>{t.L.noReports}</div>
      </div>}
      {sortedDays.map(dayKey=>{
        const day=byDay[dayKey];
        const dateStr=fmtDate(day.date.toISOString());
        const total=day.productions.length+day.samples.length+day.fires.length;
        const scrollTo=LS.get("tk_reportscrollto",null);
        const highlight=scrollTo===dayKey;
        return <div key={dayKey} id={`day-${dayKey}`} ref={el=>{if(el&&highlight){setTimeout(()=>{el.scrollIntoView({behavior:"smooth",block:"center"});LS.set("tk_reportscrollto",null)},100)}}} style={{...cSt(t),padding:"14px 16px",marginBottom:10,border:highlight?`2px solid ${t.accent}`:undefined,boxShadow:highlight?`0 0 0 4px ${t.acB}`:undefined,transition:"all 0.3s"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div>
              <div style={{fontSize:15,fontWeight:700,color:t.text}}>📅 {dateStr}</div>
              <div style={{fontSize:11,color:t.tm,marginTop:2}}>{day.productions.length} {t.L.production.toLowerCase()} · {day.samples.length} {t.L.sampleName.toLowerCase()} · {day.fires.length} {t.L.fire.toLowerCase()}</div>
            </div>
            <div style={{display:"flex",gap:6}}>
              {day.productions.length>0&&<button onClick={()=>generateFR06(dayKey)} style={{...bSt("p",t),fontSize:11,padding:"6px 10px"}}>📄 FR.06</button>}
              {day.samples.length>0&&<button onClick={()=>generateFR12(dayKey)} style={{...bSt("s",t),fontSize:11,padding:"6px 10px"}}>🧪 FR.12</button>}
            </div>
          </div>
          {day.productions.length>0&&<div style={{marginBottom:6}}>
            <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:4}}>{t.L.productionsCap}</div>
            {day.productions.map(p=><div key={p.id} style={{fontSize:12,color:t.ts,padding:"3px 0",display:"flex",justifyContent:"space-between"}}>
              <span>• {p.recipeName} <span style={{color:t.tm}}>({p.portions} {t.L.portions} · {t.L.lot} {p.lotNumber})</span></span>
              <span style={{color:p.status==="consumed"?t.success:t.tm,fontSize:10}}>{p.status==="consumed"?"✓ "+t.L.consumed:p.status==="active"?t.L.active:p.status}</span>
            </div>)}
          </div>}
          {day.samples.length>0&&<div style={{marginBottom:6}}>
            <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:4}}>{t.L.samplesCap}</div>
            {day.samples.map(p=><div key={p.id} style={{fontSize:12,color:t.ts,padding:"3px 0"}}>
              🧪 {p.recipeName} <span style={{color:t.tm}}>{p.sampleOrganization||p.sampleLocation||""} {p.samplePax>1?"("+p.samplePax+" pax)":""}</span>
            </div>)}
          </div>}
          {day.fires.length>0&&<div>
            <div style={{fontSize:11,fontWeight:700,color:t.danger,marginBottom:4}}>{t.L.firesCap}</div>
            {day.fires.map(p=><div key={p.id} style={{fontSize:12,color:t.danger,padding:"3px 0"}}>
              × {p.recipeName} <span style={{color:t.tm}}>({p.fireAmount||p.portions} porsiyon)</span>
            </div>)}
          </div>}
        </div>;
      })}
      </>}

      {/* Depo Kontrol Modal */}
      {showStorageCheckModal&&(()=>{
        const sid=showStorageCheckModal.storageId;
        const s=(storageAreas||[]).find(x=>x.id===sid);
        const typ=STORAGE_TYPES.find(x=>x.id===s?.type);
        return <StorageCheckModal storage={s} storageType={typ} date={showStorageCheckModal.date} profile={profile} onSave={(check)=>{setStorageChecks(prev=>[...prev,check]);setShowStorageCheckModal(null)}} onClose={()=>setShowStorageCheckModal(null)} t={t}/>;
      })()}
    </div>;
  }

  // Detay modalı
  if(activeProd){
    const dl=daysLeft(activeProd.expiresAt);
    return <div style={mOv(t)} onClick={()=>setActiveProd(null)}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),padding:0}}>
      <div style={{padding:"22px 26px 18px",borderBottom:`1px solid ${t.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,color:t.tm,marginBottom:4}}>Lot: {activeProd.lotNumber}</div>
            <h2 style={{fontSize:22,color:t.text,marginBottom:8}}>{activeProd.recipeName}</h2>
            <div style={{fontSize:14,color:t.ts}}>{activeProd.portions} porsiyon · {activeProd.storageIcon} {activeProd.storageName}</div>
          </div>
          <button onClick={()=>setActiveProd(null)} style={{...bSt("s",t),padding:"6px 12px",fontSize:18}}>×</button>
        </div>
      </div>
      <div style={{padding:"18px 26px"}}>
        <div style={{background:t.inBg,borderRadius:12,padding:"12px 14px",marginBottom:14,fontSize:13,color:t.ts,lineHeight:1.8}}>
          <div><strong>{t.L.produced}:</strong> {fmtDate(activeProd.producedAt)} {fmtTime(activeProd.producedAt)}</div>
          <div><strong>{t.L.expires}:</strong> {fmtDate(activeProd.expiresAt)} <span style={{color:dl<=1?t.danger:dl<=3?t.warn:t.success,fontWeight:600}}>({dl<0?"⚠":dl+" "+t.L.daysLater})</span></div>
          <div><strong>{t.L.storage}:</strong> {activeProd.storageName} ({activeProd.storageTemp}°C)</div>
          {activeProd.note&&<div><strong>{t.L.note}:</strong> {activeProd.note}</div>}
        </div>
        {activeProd.allergens&&activeProd.allergens.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:14}}>
          {activeProd.allergens.map(a=><ABadge key={a} a={a} t={t}/>)}
        </div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <button onClick={()=>consumeAll(activeProd)} style={{...bSt("p",t),fontSize:13}}>{t.L.consumeAll}</button>
          <button onClick={()=>printLabel(activeProd)} style={{...bSt("s",t),fontSize:13}}>{t.L.printLabel}</button>
        </div>
        <button onClick={()=>{
          const recipe=(recipes||[]).find(r=>r.id===activeProd.recipeId)||null;
          printCustomerQR(activeProd,recipe);
        }} style={{...bSt("s",t),fontSize:12,width:"100%",marginBottom:10}}>{t.L.customerQRBtn||"📱 Müşteri QR Bas"}</button>
        {traceability&&setTab&&<button onClick={()=>{
          // Rapora git: üretim gününü localStorage'a yaz ki reports sekmesi o günü vurgulasın
          const d=new Date(activeProd.producedAt);
          const dayKey=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
          LS.set("tk_reportscrollto",dayKey);
          setActiveProd(null);
          setTab("reports");
        }} style={{...bSt("s",t),width:"100%",marginBottom:10,fontSize:13}}>{t.L.gotoReport}</button>}
        <div style={{background:t.inBg,borderRadius:10,padding:"10px 12px",marginBottom:10}}>
          <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:6}}>{t.L.partialConsume}</div>
          <div style={{display:"flex",gap:6}}>
            <input type="number" min="0" max={activeProd.portions} value={partialAmount||""} onChange={e=>setPartialAmount(parseInt(e.target.value,10)||0)} placeholder={t.L.portions} style={{...iSt(t),flex:1}}/>
            <button onClick={()=>consumePartial(activeProd,partialAmount)} style={{...bSt("p",t),fontSize:13}}>{t.L.decrease}</button>
          </div>
        </div>
        <button onClick={()=>markFire(activeProd)} style={{...bSt("d",t),width:"100%",fontSize:13}}>{t.L.markFire}</button>
      </div>
    </div></div>;
  }

  // Ana liste
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
      <h3 style={{fontSize:22,color:t.text,margin:0}}>🍱 {t.L.production} <span style={{fontSize:14,color:t.tm,fontWeight:400}}>({activeList.length})</span></h3>
      {!traceability&&<button onClick={()=>setShowReports(true)} style={{...bSt("s",t),fontSize:12}}>📊 {t.L.reports}</button>}
    </div>
    {/* Takip No Arama */}
    <div style={{position:"relative",marginBottom:10}}>
      <input
        style={{...iSt(t),paddingLeft:36,fontSize:13,letterSpacing:"0.04em"}}
        placeholder={t.lang==="tr"?"№ Takip No ara (örn: 20260425-0002)":t.lang==="en"?"№ Search Track No (e.g. 20260425-0002)":"№ Track No..."}
        value={trackSearch}
        onChange={e=>setTrackSearch(e.target.value.replace(/[^0-9\-]/g,""))}
      />
      <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",fontSize:14,color:t.tm,pointerEvents:"none"}}>🔍</span>
      {trackSearch&&<button onClick={()=>setTrackSearch("")} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:t.tm,fontSize:16}}>✕</button>}
    </div>
    {trackSearch.trim()&&<div style={{fontSize:12,color:filtered.length>0?t.success:t.danger,marginBottom:8,fontWeight:600}}>
      {filtered.length>0?`✓ ${filtered.length} ${t.lang==="tr"?"sonuç":"result"}`:`✕ ${t.lang==="tr"?"Bulunamadı":"Not found"}`}
    </div>}
    {/* Depo filtreleri */}
    <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap",overflowX:"auto",alignItems:"center"}}>
      <button onClick={()=>setSelectedStorage("all")} style={{padding:"8px 12px",borderRadius:10,fontSize:12,fontWeight:600,border:`1px solid ${selectedStorage==="all"?t.accent:t.inBo}`,background:selectedStorage==="all"?t.acB:"transparent",color:selectedStorage==="all"?t.accent:t.tm,cursor:"pointer",whiteSpace:"nowrap"}}>{t.L.all} ({activeList.length})</button>
      {storageAreas.map(s=>{
        const cnt=activeList.filter(p=>p.storageId===s.id).length;
        const typ=STORAGE_TYPES.find(x=>x.id===s.type);
        // Bugünün kontrol sayısı
        const todayKey=(()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`})();
        const todayCheckCnt=(storageChecks||[]).filter(c=>c.storageId===s.id&&c.date===todayKey).length;
        return <div key={s.id} style={{display:"inline-flex",gap:0,alignItems:"stretch"}}>
          <button onClick={()=>setSelectedStorage(s.id)} style={{padding:"8px 10px",borderRadius:traceability?"10px 0 0 10px":"10px",fontSize:12,fontWeight:600,border:`1px solid ${selectedStorage===s.id?t.accent:t.inBo}`,background:selectedStorage===s.id?t.acB:"transparent",color:selectedStorage===s.id?t.accent:t.tm,cursor:"pointer",whiteSpace:"nowrap",borderRight:traceability?"none":undefined}}>{typ?.icon} {s.name} ({cnt})</button>
          {traceability&&<button title="Depo Kontrol Formu" onClick={()=>{
            LS.set("tk_reportscrollstorage",s.id);
            if(setTab)setTab("reports");
          }} style={{padding:"8px 8px",borderRadius:"0 10px 10px 0",fontSize:11,fontWeight:600,border:`1px solid ${t.inBo}`,borderLeft:"none",background:todayCheckCnt>=3?t.success+"22":todayCheckCnt>0?t.warn+"22":t.danger+"22",color:todayCheckCnt>=3?t.success:todayCheckCnt>0?t.warn:t.danger,cursor:"pointer",whiteSpace:"nowrap"}}>🗄 {todayCheckCnt}/3</button>}
        </div>;
      })}
    </div>
    {/* Üretim kartları */}
    {filtered.length===0&&<div style={{textAlign:"center",padding:"50px 20px",color:t.tm}}>
      <div style={{fontSize:40,opacity:0.4,marginBottom:10}}>🍱</div>
      <div style={{fontSize:14}}>{t.L.noProduction}</div>
      <div style={{fontSize:11,marginTop:4}}>{t.L.noProductionHint}</div>
    </div>}
    {filtered.map(p=>{
      const dl=daysLeft(p.expiresAt);
      const dlColor=dl<=1?t.danger:dl<=3?t.warn:t.success;
      return <div key={p.id} onClick={()=>{setActiveProd(p);setPartialAmount(0)}} style={{...cSt(t),padding:"14px 16px",marginBottom:8,cursor:"pointer"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:16,fontWeight:700,color:t.text}}>{p.recipeName}</div>
            <div style={{fontSize:12,color:t.tm,marginTop:2}}>{p.storageIcon} {p.storageName} · {p.portions} {t.L.portions} · {t.L.lot} {p.lotNumber}</div>
            <div style={{fontSize:11,color:t.tm,marginTop:4}}>{t.L.produced}: {fmtDate(p.producedAt)} · <span style={{color:dlColor,fontWeight:600}}>{t.L.expires}: {fmtDate(p.expiresAt)} ({dl<0?"⚠":dl+"g"})</span></div>
          </div>
          {p.allergens&&p.allergens.length>0&&<div style={{fontSize:14,opacity:0.7}}>{p.allergens.slice(0,3).map(a=>ALLERGENS.find(x=>x.id===a)?.icon).filter(Boolean).join(" ")}</div>}
        </div>
      </div>;
    })}
  </div>;
};

// ═══ MENÜ ═══
// ═══ IMAGE CROP MODAL ═══
const ImageCropModal=({image,targetRatio,onClose,onCrop,t})=>{
  const [scale,setScale]=useState(1);
  const [offsetX,setOffsetX]=useState(0);
  const [offsetY,setOffsetY]=useState(0);
  const [dragging,setDragging]=useState(false);
  const [dragStart,setDragStart]=useState({x:0,y:0,offsetX:0,offsetY:0});
  const [imgSize,setImgSize]=useState({w:0,h:0});
  const frameRef=React.useRef(null);
  const imgRef=React.useRef(null);

  // Frame boyutu (ekranda gösterilecek)
  const frameWidth=300;
  const frameHeight=frameWidth/targetRatio;

  const onImgLoad=(e)=>{
    const img=e.target;
    setImgSize({w:img.naturalWidth,h:img.naturalHeight});
    // Frame'i tam dolduracak min scale hesapla
    const imgRatio=img.naturalWidth/img.naturalHeight;
    const frameRatio=targetRatio;
    let initialScale=1;
    if(imgRatio>frameRatio){
      // Resim daha geniş — yükseklik sığdır
      initialScale=frameHeight/img.naturalHeight;
    }else{
      // Resim daha dar — genişlik sığdır
      initialScale=frameWidth/img.naturalWidth;
    }
    setScale(initialScale);
    setOffsetX(0);
    setOffsetY(0);
  };

  const startDrag=(e)=>{
    const pt=e.touches?e.touches[0]:e;
    setDragging(true);
    setDragStart({x:pt.clientX,y:pt.clientY,offsetX,offsetY});
  };
  const onDrag=(e)=>{
    if(!dragging)return;
    e.preventDefault();
    const pt=e.touches?e.touches[0]:e;
    setOffsetX(dragStart.offsetX+(pt.clientX-dragStart.x));
    setOffsetY(dragStart.offsetY+(pt.clientY-dragStart.y));
  };
  const endDrag=()=>setDragging(false);

  const doCrop=()=>{
    // Canvas oluştur, kırpılmış alanı oraya çiz
    const canvas=document.createElement("canvas");
    const outW=1200; // çıktı kalitesi
    const outH=Math.round(outW/targetRatio);
    canvas.width=outW;
    canvas.height=outH;
    const ctx=canvas.getContext("2d");

    const img=imgRef.current;
    if(!img){onClose();return}

    // Frame içinde görünen bölge: frameWidth x frameHeight
    // Resim offset+scale ile gösteriliyor
    // Görünen alanın resim koordinatlarındaki karşılığı:
    const renderedW=imgSize.w*scale;
    const renderedH=imgSize.h*scale;
    // Frame merkezi resmi ne kadar kaplıyor? frame/2 orijinden itibaren +/- offset
    // Resim sol-üst köşesi: (frameWidth/2 - renderedW/2 + offsetX, frameHeight/2 - renderedH/2 + offsetY)
    const imgLeft=frameWidth/2-renderedW/2+offsetX;
    const imgTop=frameHeight/2-renderedH/2+offsetY;
    // Frame'in sol-üst (0,0) koordinatının resimdeki karşılığı:
    const sx=(-imgLeft)/scale;
    const sy=(-imgTop)/scale;
    const sw=frameWidth/scale;
    const sh=frameHeight/scale;

    ctx.fillStyle="#fff";
    ctx.fillRect(0,0,outW,outH);
    try{
      ctx.drawImage(img,sx,sy,sw,sh,0,0,outW,outH);
      const dataUrl=canvas.toDataURL("image/jpeg",0.85);
      onCrop(dataUrl);
    }catch(e){
      window.toast.error("Kırpma hatası: "+e.message);
    }
  };

  return <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.8)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,zIndex:10000}}>
    <div style={{background:t.card,borderRadius:18,padding:22,maxWidth:420,width:"100%",maxHeight:"95vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.6)"}}>
      <h3 style={{fontSize:18,color:t.text,marginBottom:6}}>{t.L.cropTitle}</h3>
      <div style={{fontSize:11,color:t.tm,marginBottom:14}}>{t.L.cropDesc}</div>

      <div style={{display:"flex",justifyContent:"center",marginBottom:12}}>
        <div ref={frameRef} style={{position:"relative",width:frameWidth,height:frameHeight,border:`3px solid ${t.accent}`,overflow:"hidden",background:"#000",cursor:dragging?"grabbing":"grab",touchAction:"none"}}
          onMouseDown={startDrag} onMouseMove={onDrag} onMouseUp={endDrag} onMouseLeave={endDrag}
          onTouchStart={startDrag} onTouchMove={onDrag} onTouchEnd={endDrag}>
          <img ref={imgRef} src={image} onLoad={onImgLoad} alt="" style={{position:"absolute",left:"50%",top:"50%",transform:`translate(calc(-50% + ${offsetX}px), calc(-50% + ${offsetY}px)) scale(${scale})`,transformOrigin:"center",pointerEvents:"none",userSelect:"none",maxWidth:"none"}}/>
        </div>
      </div>

      <div style={{marginBottom:12}}>
        <label style={{fontSize:11,fontWeight:600,color:t.tm,display:"block",marginBottom:4}}>{t.L.cropZoom}</label>
        <input type="range" min="0.1" max="5" step="0.05" value={scale} onChange={e=>setScale(parseFloat(e.target.value))} style={{width:"100%"}}/>
      </div>

      <div style={{display:"flex",gap:8}}>
        <button onClick={onClose} style={{...bSt("s",t),flex:1}}>{t.L.cancel}</button>
        <button onClick={doCrop} style={{...bSt("p",t),flex:2}}>{t.L.cropApply}</button>
      </div>
    </div>
  </div>;
};

const MenuTab=({menus,setMenus,recipes,menuTemplates,setMenuTemplates,t})=>{
  const[showCreate,setSC]=useState(false);
  const[editMenu,setEM]=useState(null);
  const[previewMenu,setPM]=useState(null);
  const[cropImage,setCropImage]=useState(null);

  // Yeni menü oluştur
  const createMenu=(type)=>{
    const mt=MENU_TYPES.find(m=>m.id===type);
    const newMenu={
      id:Date.now(),
      name:(t.L.newMenuName||"Yeni {type} Menüsü").replace("{type}",menuTypeL(mt,t.lang)),
      type:type,
      font:"Fraunces",
      theme:"warm",
      showPrice:type!=="banquet",
      showCalorie:true,
      showAllergen:true,
      sections:menuSectionsL(mt,t.lang).map(s=>({title:s,items:[]})),
      date:new Date().toISOString().split("T")[0]
    };
    setMenus(p=>[newMenu,...p]);
    setEM(newMenu);
    setSC(false);
  };

  // Menü düzenleme
  const updateMenu=(updated)=>{
    setMenus(p=>p.map(m=>m.id===updated.id?updated:m));
    setEM(updated);
  };

  // Bölüme reçete ekle
  const addToSection=(menu,secIdx,recipe)=>{
    const item={
      recipeId:recipe.id,
      title:recipe.name,
      description:"",
      price:"",
      allergens:recipe.allergens||[],
      calories:recipe.calories||null
    };
    const updated={...menu,sections:menu.sections.map((s,i)=>i===secIdx?{...s,items:[...s.items,item]}:s)};
    updateMenu(updated);
  };

  // Önizleme render
  const renderPreview=(menu)=>{
    const th=MENU_THEMES.find(x=>x.id===menu.theme)||MENU_THEMES[1];
    const bgOp=(menu.bgOpacity||25)/100;
    // Kağıt boyutu — en-boy oranı
    const paperSizes={a4:[210,297],a5:[148,210],a6:[105,148],letter:[216,279]};
    let pW=210,pH=297;
    if(menu.paperSize==="custom"&&menu.customWidth&&menu.customHeight){
      pW=menu.customWidth;pH=menu.customHeight;
    }else if(paperSizes[menu.paperSize]){
      [pW,pH]=paperSizes[menu.paperSize];
    }
    const ratio=pW/pH;
    // Ekranda görünen genişlik
    const previewW=500;
    const previewH=previewW/ratio;
    return <div style={{display:"flex",justifyContent:"center"}}>
      <div style={{position:"relative",fontFamily:menu.font+",serif",background:th.bg,color:th.text,padding:"40px 30px",borderRadius:8,width:previewW,minHeight:previewH,overflow:"hidden",boxShadow:"0 4px 20px rgba(0,0,0,0.15)"}}>
      {menu.bgImage&&<div style={{position:"absolute",inset:0,backgroundImage:`url(${menu.bgImage})`,backgroundSize:"cover",backgroundPosition:"center",opacity:bgOp,zIndex:0,pointerEvents:"none"}}/>}
      <div style={{position:"relative",zIndex:1}}>
      <div style={{textAlign:"center",marginBottom:30}}>
        <h1 style={{fontSize:28,fontWeight:700,color:th.accent,margin:0,letterSpacing:"0.05em"}}>{menu.name}</h1>
        <div style={{fontSize:12,color:th.text,opacity:0.5,marginTop:6}}>{menu.date}</div>
      </div>
      {menu.sections.map((sec,si)=>{
        if(!sec.items.length)return null;
        return <div key={si} style={{marginBottom:24}}>
          <div style={{fontSize:14,fontWeight:700,color:th.accent,letterSpacing:"0.15em",textTransform:"uppercase",borderBottom:`1px solid ${th.border}`,paddingBottom:6,marginBottom:12}}>{sec.title}</div>
          {sec.items.map((item,ii)=><div key={ii} style={{marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline"}}>
              <span style={{fontSize:17,fontWeight:600}}>{item.title}</span>
              {menu.showPrice&&item.price&&<span style={{fontSize:15,fontWeight:600,color:th.accent}}>₺{item.price}</span>}
            </div>
            {item.description&&<div style={{fontSize:13,opacity:0.7,fontStyle:"italic",marginTop:2}}>{item.description}</div>}
            <div style={{display:"flex",gap:8,marginTop:4,flexWrap:"wrap"}}>
              {menu.showAllergen&&item.allergens.map(a=>{const al=ALLERGENS.find(x=>x.id===a);return al?<span key={a} style={{fontSize:10,opacity:0.6}}>{al.icon}</span>:null})}
              {menu.showCalorie&&item.calories&&<span style={{fontSize:10,opacity:0.5}}>🔥 {item.calories} kcal/100g</span>}
            </div>
          </div>)}
        </div>;
      })}
      </div>
    </div>
    </div>;
  };

  // PDF oluştur
  const genPDF=async(menu,paper="a4",mode="auto",action="download")=>{
    const th=MENU_THEMES.find(x=>x.id===menu.theme)||MENU_THEMES[1];
    const el=document.getElementById("menu-preview-render");
    if(!el)return;
    try{
      const canvas=await window.html2canvas(el,{scale:2,backgroundColor:th.bg,useCORS:true});
      const imgData=canvas.toDataURL("image/png");
      const{jsPDF}=window.jspdf;
      // Menü kayıtlı özel boyut varsa öncelik ver
      const sizes={a4:[210,297],a5:[148,210],a6:[105,148],letter:[216,279]};
      let w,h,format;
      const menuPaper=menu.paperSize||paper;
      if(menuPaper==="custom"&&menu.customWidth&&menu.customHeight){
        w=menu.customWidth;
        h=menu.customHeight;
        format=[w,h];
      }else if(sizes[menuPaper]){
        [w,h]=sizes[menuPaper];
        format=menuPaper;
      }else{
        [w,h]=sizes[paper]||sizes.a4;
        format=paper;
      }
      const pdf=new jsPDF({unit:"mm",format:format,orientation:w>h?"landscape":"portrait"});
      const imgW=w, imgH=(canvas.height*imgW)/canvas.width;
      if(mode==="single"||(mode==="auto"&&imgH<=h)){
        pdf.addImage(imgData,"PNG",0,0,imgW,Math.min(imgH,h));
      }else{
        let y=0;
        while(y<imgH){
          pdf.addImage(imgData,"PNG",0,-y,imgW,imgH);
          y+=h;
          if(y<imgH)pdf.addPage();
        }
      }
      if(action==="download"){
        pdf.save(`${menu.name}.pdf`);
      }else if(action==="share"){
        const blob=pdf.output("blob");
        const file=new File([blob],`${menu.name}.pdf`,{type:"application/pdf"});
        if(navigator.canShare&&navigator.canShare({files:[file]})){
          await navigator.share({files:[file],title:menu.name});
        }else{
          pdf.save(`${menu.name}.pdf`);
        }
      }
    }catch(e){window.toast.error("PDF hatası: "+e.message);}
  };

  // Menü kopyala
  const copyMenu=(menu)=>{
    const copy={...JSON.parse(JSON.stringify(menu)),id:Date.now(),name:menu.name+" (kopya)",date:new Date().toISOString().split("T")[0]};
    setMenus(p=>[copy,...p]);
  };

  // Yazdır
  const printMenu=(menu)=>{
    const w=window.open("","","width=800,height=1000");
    const th=MENU_THEMES.find(x=>x.id===menu.theme)||MENU_THEMES[1];
    const bgOp=(menu.bgOpacity||25)/100;
    let html="<html><head><style>@import url('https://fonts.googleapis.com/css2?family="+encodeURIComponent(menu.font)+":wght@400;600;700&display=swap');*{margin:0;box-sizing:border-box}body{font-family:'"+menu.font+"',serif;background:"+th.bg+";color:"+th.text+";padding:50px 40px;position:relative;min-height:100vh}"+
      (menu.bgImage?`.bg{position:fixed;inset:0;background-image:url('${menu.bgImage}');background-size:cover;background-position:center;opacity:${bgOp};z-index:0}`:"")+
      ".content{position:relative;z-index:1}</style></head><body>";
    if(menu.bgImage)html+="<div class='bg'></div>";
    html+="<div class='content'>";
    html+="<div style='text-align:center;margin-bottom:36px'><h1 style='font-size:32px;color:"+th.accent+";letter-spacing:0.05em'>"+menu.name+"</h1><div style='font-size:12px;opacity:0.5;margin-top:6px'>"+menu.date+"</div></div>";
    menu.sections.forEach(sec=>{
      if(!sec.items.length)return;
      html+="<div style='margin-bottom:28px'><div style='font-size:14px;font-weight:700;color:"+th.accent+";letter-spacing:0.15em;text-transform:uppercase;border-bottom:1px solid "+th.border+";padding-bottom:6px;margin-bottom:14px'>"+sec.title+"</div>";
      sec.items.forEach(item=>{
        html+="<div style='margin-bottom:14px'><div style='display:flex;justify-content:space-between;align-items:baseline'><span style='font-size:18px;font-weight:600'>"+item.title+"</span>";
        if(menu.showPrice&&item.price)html+="<span style='font-size:16px;font-weight:600;color:"+th.accent+"'>"+item.price+" TL</span>";
        html+="</div>";
        if(item.description)html+="<div style='font-size:13px;opacity:0.7;font-style:italic;margin-top:3px'>"+item.description+"</div>";
        let badges="";
        if(menu.showAllergen&&item.allergens.length){item.allergens.forEach(a=>{const al=ALLERGENS.find(x=>x.id===a);if(al)badges+=al.icon+" "});}
        if(menu.showCalorie&&item.calories)badges+="🔥 "+item.calories+" kcal/100g";
        if(badges)html+="<div style='font-size:10px;opacity:0.5;margin-top:4px'>"+badges+"</div>";
        html+="</div>";
      });
      html+="</div>";
    });
    html+="</div></body></html>";
    w.document.write(html);w.document.close();
    setTimeout(()=>{w.print();},500);
  };

  // Önizleme state'leri
  const[paperSize,setPaperSize]=useState("a4");
  const[pageMode,setPageMode]=useState("auto");

  // Ana render
  // Crop modal (tüm ekranlarda aktif olacak)
  const cropModalJSX=cropImage?<ImageCropModal image={cropImage} targetRatio={(editMenu?.customWidth||210)/(editMenu?.customHeight||297)} onClose={()=>setCropImage(null)} onCrop={(croppedDataUrl)=>{
    if(editMenu)updateMenu({...editMenu,bgImage:croppedDataUrl});
    setCropImage(null);
  }} t={t}/>:null;

  if(previewMenu)return <div>{cropModalJSX}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
      <h3 style={{fontSize:22,color:t.text,margin:0}}>Önizleme</h3>
      <button onClick={()=>setPM(null)} style={{...bSt("s",t),fontSize:13}}>← Geri</button>
    </div>
    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14,padding:"10px 12px",background:t.inBg,borderRadius:12}}>
      <div style={{display:"flex",gap:4}}>
        {[["a4","A4"],["a5","A5"],["letter","Letter"]].map(([id,l])=><button key={id} onClick={()=>setPaperSize(id)} style={{padding:"6px 10px",borderRadius:8,fontSize:11,fontWeight:600,border:`1px solid ${paperSize===id?t.accent:t.inBo}`,background:paperSize===id?t.acB:"transparent",color:paperSize===id?t.accent:t.tm,cursor:"pointer"}}>{l}</button>)}
      </div>
      <div style={{display:"flex",gap:4}}>
        {[["auto","Otomatik"],["single","Tek sayfa"],["multi","Çok sayfa"]].map(([id,l])=><button key={id} onClick={()=>setPageMode(id)} style={{padding:"6px 10px",borderRadius:8,fontSize:11,fontWeight:600,border:`1px solid ${pageMode===id?t.accent:t.inBo}`,background:pageMode===id?t.acB:"transparent",color:pageMode===id?t.accent:t.tm,cursor:"pointer"}}>{l}</button>)}
      </div>
      <div style={{flex:1,minWidth:10}}/>
      <button onClick={()=>genPDF(previewMenu,paperSize,pageMode,"download")} style={{...bSt("p",t),padding:"6px 12px",fontSize:12}}>📥 PDF</button>
      <button onClick={()=>genPDF(previewMenu,paperSize,pageMode,"share")} style={{...bSt("s",t),padding:"6px 12px",fontSize:12}}>📤 Paylaş</button>
      <button onClick={()=>printMenu(previewMenu)} style={{...bSt("s",t),padding:"6px 12px",fontSize:12}}>🖨 Yazdır</button>
    </div>
    <div id="menu-preview-render">
      {renderPreview(previewMenu)}
    </div>
  </div>;

  if(editMenu)return <div>{cropModalJSX}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
      <h3 style={{fontSize:22,color:t.text}}>{t.L.menuEdit}</h3>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>{setPM(editMenu);setEM(null)}} style={{...bSt("p",t),fontSize:13}}>{t.L.menuPreviewBtn}</button>
        <button onClick={()=>setEM(null)} style={{...bSt("s",t),fontSize:13}}>← {t.L.back}</button>
      </div>
    </div>
    {/* Menü ayarları */}
    <div style={{...cSt(t),padding:"14px 16px",marginBottom:14}}>
      <div><label style={lSt(t)}>{t.L.menuName}</label><input style={iSt(t)} value={editMenu.name} onChange={e=>updateMenu({...editMenu,name:e.target.value})}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:10}}>
        <div><label style={lSt(t)}>{t.L.font}</label><select style={{...iSt(t),fontFamily:editMenu.font}} value={editMenu.font} onChange={e=>updateMenu({...editMenu,font:e.target.value})}>{MENU_FONTS.map(f=><option key={f} value={f} style={{fontFamily:f}}>{f}</option>)}</select></div>
        <div><label style={lSt(t)}>{t.L.theme}</label><select style={iSt(t)} value={editMenu.theme} onChange={e=>updateMenu({...editMenu,theme:e.target.value})}>{MENU_THEMES.map(th=><option key={th.id} value={th.id}>{themeL(th,t.lang)}</option>)}</select></div>
      </div>
      <div style={{marginTop:10,padding:"10px 12px",background:t.inBg,borderRadius:10}}>
        <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:6}}>{t.L.customSizes}</div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:4}}>
            {[["a4",210,297,"A4"],["a5",148,210,"A5"],["a6",105,148,"A6"],["letter",216,279,"Letter"]].map(([id,w,h,l])=><button key={id} onClick={()=>updateMenu({...editMenu,paperSize:id,customWidth:w,customHeight:h})} style={{padding:"5px 8px",borderRadius:6,fontSize:10,fontWeight:600,border:`1px solid ${editMenu.paperSize===id?t.accent:t.inBo}`,background:editMenu.paperSize===id?t.acB:"transparent",color:editMenu.paperSize===id?t.accent:t.tm,cursor:"pointer"}}>{l}</button>)}
          </div>
          <span style={{fontSize:10,color:t.tm}}>veya</span>
          <input type="number" min="50" max="500" value={editMenu.customWidth||210} onChange={e=>updateMenu({...editMenu,paperSize:"custom",customWidth:parseInt(e.target.value,10)||210})} style={{...iSt(t),width:60,textAlign:"center",fontSize:12,padding:"4px 6px"}}/>
          <span style={{fontSize:11,color:t.tm}}>×</span>
          <input type="number" min="50" max="500" value={editMenu.customHeight||297} onChange={e=>updateMenu({...editMenu,paperSize:"custom",customHeight:parseInt(e.target.value,10)||297})} style={{...iSt(t),width:60,textAlign:"center",fontSize:12,padding:"4px 6px"}}/>
          <span style={{fontSize:10,color:t.tm}}>mm</span>
        </div>
      </div>
      <div style={{display:"flex",gap:10,marginTop:10,flexWrap:"wrap"}}>
        <label style={{fontSize:13,color:t.ts,display:"flex",alignItems:"center",gap:4,cursor:"pointer"}}><input type="checkbox" checked={editMenu.showPrice} onChange={e=>updateMenu({...editMenu,showPrice:e.target.checked})}/> {t.L.showPrice}</label>
        <label style={{fontSize:13,color:t.ts,display:"flex",alignItems:"center",gap:4,cursor:"pointer"}}><input type="checkbox" checked={editMenu.showCalorie} onChange={e=>updateMenu({...editMenu,showCalorie:e.target.checked})}/> {t.L.showCalorie}</label>
        <label style={{fontSize:13,color:t.ts,display:"flex",alignItems:"center",gap:4,cursor:"pointer"}}><input type="checkbox" checked={editMenu.showAllergen} onChange={e=>updateMenu({...editMenu,showAllergen:e.target.checked})}/> {t.L.showAllergen}</label>
      </div>
      <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${t.border}`}}>
        <label style={lSt(t)}>{t.L.bgImage}</label>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          {editMenu.bgImage&&<img src={editMenu.bgImage} style={{width:60,height:60,objectFit:"cover",borderRadius:8,border:`1px solid ${t.border}`}}/>}
          <input type="file" accept="image/*" id="menuBgInput" style={{display:"none"}} onChange={e=>{
            const f=e.target.files[0];if(!f)return;
            const rd=new FileReader();
            rd.onload=ev=>{
              setCropImage(ev.target.result);
            };
            rd.readAsDataURL(f);
            e.target.value="";
          }}/>
          <button onClick={()=>document.getElementById("menuBgInput").click()} style={{...bSt("s",t),fontSize:12}}>{editMenu.bgImage?t.L.changeBg:t.L.chooseBg}</button>
          {editMenu.bgImage&&<button onClick={()=>setCropImage(editMenu.bgImage)} style={{...bSt("s",t),fontSize:12}}>{t.L.cropBg}</button>}
          {editMenu.bgImage&&<button onClick={()=>updateMenu({...editMenu,bgImage:null})} style={{...bSt("d",t),fontSize:12}}>× {t.L.remove}</button>}
          <div style={{fontSize:11,color:t.tm,flex:1,minWidth:120}}>
            {t.L.bgOpacity}: 
            <input type="range" min="0" max="100" value={editMenu.bgOpacity||25} onChange={e=>updateMenu({...editMenu,bgOpacity:parseInt(e.target.value,10)})} style={{width:80,marginLeft:4,verticalAlign:"middle"}}/>
            <span style={{marginLeft:4}}>%{editMenu.bgOpacity||25}</span>
          </div>
        </div>
        <div style={{fontSize:10,color:t.tm,marginTop:4}}>{t.L.bgHint}</div>
      </div>
      <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${t.border}`}}>
        <div style={{fontSize:11,fontWeight:700,color:t.tm,marginBottom:6}}>{t.L.menuTemplates}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
          <select style={{...iSt(t),flex:1,minWidth:140,fontSize:12}} onChange={e=>{
            if(!e.target.value)return;
            const tpl=(menuTemplates||[]).find(x=>x.id===e.target.value);
            if(!tpl)return;
            if(!window.confirm(tr_(t.lang,"templateSaveConfirm",{name:tpl.name}))){e.target.value="";return}
            const{id,name,...tplDesign}=tpl;
            updateMenu({...editMenu,...tplDesign});
            e.target.value="";
          }}>
            <option value="">{t.L.loadTemplate}</option>
            {(menuTemplates||[]).map(tpl=><option key={tpl.id} value={tpl.id}>{tpl.name}</option>)}
          </select>
          <button onClick={()=>{
            const name=window.prompt(t.L.templateName,tr_(t.lang,"defaultTemplateName",{menuName:editMenu.name}));
            if(!name)return;
            const newTpl={
              id:"tpl_"+Date.now(),
              name,
              font:editMenu.font,
              theme:editMenu.theme,
              paperSize:editMenu.paperSize,
              customWidth:editMenu.customWidth,
              customHeight:editMenu.customHeight,
              bgImage:editMenu.bgImage,
              bgOpacity:editMenu.bgOpacity,
              showPrice:editMenu.showPrice,
              showCalorie:editMenu.showCalorie,
              showAllergen:editMenu.showAllergen
            };
            setMenuTemplates([...(menuTemplates||[]),newTpl]);
            window.toast.info(tr_(t.lang,"templateSaved",{name}));
          }} style={{...bSt("p",t),fontSize:11}}>{t.L.saveTemplate}</button>
          {(menuTemplates||[]).length>0&&<button onClick={()=>{
            const name=window.prompt("Silinecek şablon ismi?");
            if(!name)return;
            const tpl=(menuTemplates||[]).find(x=>x.name===name);
            if(!tpl){window.toast.info(t.L.templateNotFound);return}
            setMenuTemplates(menuTemplates.filter(x=>x.id!==tpl.id));
          }} style={{...bSt("d",t),fontSize:11}}>× {t.L.delete}</button>}
        </div>
        <div style={{fontSize:10,color:t.tm,marginTop:4}}>{t.L.templateHint}</div>
      </div>
    </div>
    {/* Bölümler */}
    {editMenu.sections.map((sec,si)=><div key={si} style={{...cSt(t),padding:"14px 16px",marginBottom:10}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:10}}>
        <input style={{...iSt(t),fontWeight:700,fontSize:15,border:"none",background:"transparent",padding:0,flex:1}} value={sec.title} onChange={e=>{const s=[...editMenu.sections];s[si]={...s[si],title:e.target.value};updateMenu({...editMenu,sections:s});}}/>
        <span style={{fontSize:11,color:t.tm}}>{sec.items.length}</span>
        <button onClick={()=>{
          if(!window.confirm(`"${sec.title}" ${t.lang==="tr"?"bölümü silinsin mi?":"delete this section?"}`))return;
          updateMenu({...editMenu,sections:editMenu.sections.filter((_,j)=>j!==si)});
        }} style={{...bSt("d",t),padding:"4px 8px",fontSize:11}} title={t.L.delete}>×</button>
      </div>
      {sec.items.map((item,ii)=><div key={ii} style={{background:t.inBg,borderRadius:10,padding:"10px 12px",marginBottom:6}}>
        <div style={{display:"flex",gap:8,marginBottom:6}}>
          <input style={{...iSt(t),flex:1,fontWeight:600,fontSize:14}} value={item.title} placeholder={t.L.menuItemTitle} onChange={e=>{const s=[...editMenu.sections];s[si].items[ii]={...item,title:e.target.value};updateMenu({...editMenu,sections:s});}}/>
          {editMenu.showPrice&&<input style={{...iSt(t),width:80,textAlign:"right"}} value={item.price} placeholder="₺" onChange={e=>{const s=[...editMenu.sections];s[si].items[ii]={...item,price:e.target.value};updateMenu({...editMenu,sections:s});}}/>}
          <button onClick={()=>{const s=[...editMenu.sections];s[si]={...s[si],items:s[si].items.filter((_,j)=>j!==ii)};updateMenu({...editMenu,sections:s});}} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
        </div>
        <input style={{...iSt(t),fontSize:12}} value={item.description} placeholder={t.L.menuItemDesc} onChange={e=>{const s=[...editMenu.sections];s[si].items[ii]={...item,description:e.target.value};updateMenu({...editMenu,sections:s});}}/>
        {item.allergens.length>0&&<div style={{display:"flex",gap:3,marginTop:6,flexWrap:"wrap"}}>{item.allergens.map(a=><ABadge key={a} a={a} t={t}/>)}</div>}
      </div>)}
      {/* Reçete ekle dropdown */}
      <select style={{...iSt(t),fontSize:13,marginTop:4}} value="" onChange={e=>{if(!e.target.value)return;const recipe=recipes.find(r=>r.id===parseInt(e.target.value,10));if(recipe)addToSection(editMenu,si,recipe);e.target.value="";}}>
        <option value="">{t.L.menuAutoLoad}...</option>
        {recipes.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}
      </select>
    </div>)}
    {/* Yeni bölüm ekle */}
    <button onClick={()=>{
      const newTitle=window.prompt(t.lang==="tr"?"Yeni bölüm başlığı:":t.lang==="en"?"New section title:":"New:",t.lang==="tr"?"Yeni Bölüm":"New Section");
      if(!newTitle)return;
      updateMenu({...editMenu,sections:[...editMenu.sections,{title:newTitle,items:[]}]});
    }} style={{...bSt("s",t),width:"100%",marginTop:8,fontSize:13}}>+ {t.lang==="tr"?"Yeni Bölüm":t.lang==="en"?"New Section":"+"}</button>
  </div>;

  // Liste
  return <div>{cropModalJSX}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
      <h3 style={{fontSize:22,color:t.text}}>{t.L.menus} <span style={{fontSize:14,color:t.tm,fontWeight:400}}>({menus.length})</span></h3>
      <button onClick={()=>setSC(true)} style={{...bSt("p",t),fontSize:13}}>{t.L.newMenu}</button>
    </div>
    {/* Oluşturma modal */}
    {showCreate&&<div style={{...cSt(t),padding:"18px 20px",marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:700,color:t.text,marginBottom:12}}>{t.L.chooseMenuType}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        {MENU_TYPES.map(mt=><button key={mt.id} onClick={()=>createMenu(mt.id)} style={{...cSt(t),padding:"16px 14px",textAlign:"center",cursor:"pointer",border:`1px solid ${t.inBo}`}}>
          <div style={{fontSize:28,marginBottom:6}}>{mt.icon}</div>
          <div style={{fontSize:14,fontWeight:600,color:t.text}}>{menuTypeL(mt,t.lang)}</div>
          <div style={{fontSize:10,color:t.tm,marginTop:4}}>{menuSectionsL(mt,t.lang).join(" · ")}</div>
        </button>)}
      </div>
      <button onClick={()=>setSC(false)} style={{...bSt("s",t),width:"100%",marginTop:10}}>{t.L.cancel}</button>
    </div>}
    {menus.length===0&&!showCreate&&<div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
      <div style={{fontSize:48,opacity:0.4,marginBottom:12}}>📋</div>
      <div style={{fontSize:16,fontFamily:"'Fraunces',serif"}}>{t.L.menus}</div>
      <div style={{fontSize:13,marginTop:6}}>{t.L.emptyStateMenus}</div>
    </div>}
    {menus.map(menu=>{
      const mt=MENU_TYPES.find(x=>x.id===menu.type);
      const itemCount=menu.sections.reduce((a,s)=>a+s.items.length,0);
      return <div key={menu.id} style={{...cSt(t),padding:"16px 18px",marginBottom:10,cursor:"pointer"}} onClick={()=>setEM(menu)}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <div style={{fontSize:11,color:t.tm}}>{mt?.icon} {menuTypeL(mt,t.lang)} · {menu.date}</div>
            <div style={{fontSize:17,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif",marginTop:4}}>{menu.name}</div>
            <div style={{fontSize:12,color:t.tm,marginTop:4}}>{itemCount} · {menu.font} · {themeL(MENU_THEMES.find(x=>x.id===menu.theme),t.lang)}</div>
          </div>
          <div style={{display:"flex",gap:6}}>
            <button onClick={e=>{e.stopPropagation();setPM(menu)}} style={{...bSt("s",t),padding:"6px 10px",fontSize:11}}>👁</button>
            <button onClick={e=>{e.stopPropagation();copyMenu(menu)}} style={{...bSt("s",t),padding:"6px 10px",fontSize:11}}>📋</button>
            <button onClick={e=>{e.stopPropagation();printMenu(menu)}} style={{...bSt("s",t),padding:"6px 10px",fontSize:11}}>🖨</button>
            <button onClick={e=>{e.stopPropagation();setMenus(p=>p.filter(m=>m.id!==menu.id))}} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
          </div>
        </div>
      </div>;
    })}
  </div>;
};

// ═══ ASSISTANT (Çoklu oturum + Bot bildirimleri) ═══
const AssistantTab=({recipes,stock,apiKey,conversations,setConversations,activeConvId,setActiveConvId,botMessages,productions,storageChecks,t})=>{
  const[input,setIn]=useState("");
  const[loading,setL]=useState(false);
  const[view,setView]=useState("chat");  // "list" | "chat" - mobil için
  const[editingName,setEditingName]=useState(null);
  const[showArchived,setShowArchived]=useState(false);
  const endRef=useRef();

  // Aktif sohbet - yoksa ilk aktifi seç ya da yeni oluştur
  const active=conversations.find(c=>c.id===activeConvId);
  const activeConversations=conversations.filter(c=>!c.archived);
  const archivedConversations=conversations.filter(c=>c.archived);

  useEffect(()=>{
    // Aktif sohbet yok ve aktif liste boş → otomatik yeni sohbet oluştur
    if(!active){
      if(activeConversations.length>0){
        setActiveConvId(activeConversations[0].id);
        setView("chat");
      }else{
        // Hiç aktif sohbet yok, yeni oluştur
        const newConv={id:"conv_"+Date.now(),name:t.L.newChat||"Yeni Sohbet",createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),archived:false,messages:[]};
        setConversations(p=>[newConv,...p.filter(c=>c.archived)]);
        setActiveConvId(newConv.id);
        setView("chat");
      }
    }
  // eslint-disable-next-line
  },[active,activeConversations.length]);

  useEffect(()=>{endRef.current?.scrollIntoView({behavior:"smooth"})},[active?.messages?.length]);

  const newChat=()=>{
    const newConv={id:"conv_"+Date.now(),name:t.L.newChat||"Yeni Sohbet",createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),archived:false,messages:[]};
    setConversations(p=>[newConv,...p]);
    setActiveConvId(newConv.id);
    setView("chat");
  };

  const updateConv=(id,updates)=>{
    setConversations(p=>p.map(c=>c.id===id?{...c,...updates,updatedAt:new Date().toISOString()}:c));
  };

  const deleteConv=(id)=>{
    if(!window.confirm(t.L.deleteConvConfirm||"Bu sohbet silinecek. Onaylıyor musun?"))return;
    const remaining=conversations.filter(c=>c.id!==id);
    setConversations(remaining);
    if(activeConvId===id){
      const nextActive=remaining.find(c=>!c.archived);
      setActiveConvId(nextActive?nextActive.id:null);
      // Boşsa useEffect yeni sohbet oluşturacak; ama view'i de chat'e çek
      setView("chat");
    }
  };

  const archiveConv=(id)=>{
    setConversations(p=>p.map(c=>c.id===id?{...c,archived:true,updatedAt:new Date().toISOString()}:c));
    if(activeConvId===id){
      const nextActive=conversations.find(c=>c.id!==id&&!c.archived);
      setActiveConvId(nextActive?nextActive.id:null);
      setView("chat");
    }
  };

  const unarchiveConv=(id)=>updateConv(id,{archived:false});

  const renameConv=(id,newName)=>{
    updateConv(id,{name:newName.trim()||t.L.newChat||"Yeni Sohbet"});
    setEditingName(null);
  };

  const send=async()=>{
    if(!input.trim()||loading||!active)return;
    const q=input;setIn("");setL(true);
    const userMsg={role:"user",text:q,ts:new Date().toISOString()};
    // Otomatik isim: boş sohbetin ilk mesajından
    const autoName=active.messages.length===0?q.slice(0,30):active.name;
    updateConv(active.id,{messages:[...active.messages,userMsg],name:autoName});
    try{
      const rl=recipes.map(r=>`${r.id}:${r.name}(${r.cuisine})`).join(", ");
      const sl=stock.map(s=>`${s.name}:${s.qty}${s.unit}`).join(", ");
      const langName={tr:"Turkish",en:"English",ru:"Russian",es:"Spanish",de:"German",fr:"French",zh:"Chinese",ar:"Arabic"}[t.lang]||"Turkish";
      // Raporları hazırla
      const now=new Date();
      const recentProds=productions?productions.filter(p=>{
        const d=new Date(p.producedAt||p.createdAt||"");
        return (now-d)<30*24*60*60*1000; // son 30 gün
      }).slice(0,20):[];
      const recentChecks=storageChecks?storageChecks.slice(-20):[];
      const sys=`You are Kitchen Manager AI assistant. Respond in ${langName} ONLY. Be concise and direct.

CURRENT DATA ACCESS:
Recipes (${(recipes||[]).length}): ${rl}
Stock (${(stock||[]).length} items): ${sl}
Recent Productions (last 30 days, ${recentProds.length}): ${JSON.stringify(recentProds.map(p=>({id:p.id,name:p.recipeName,date:p.producedAt,labelSeq:p.labelSeq,status:p.status,storage:p.storageId})))}
Storage Checks (${recentChecks.length}): ${JSON.stringify(recentChecks.map(c=>({date:c.date,storage:c.storageId,temp:c.temp,status:c.status})))}

REPORT CAPABILITIES:
- If user asks for reports (rapor, FR.06, FR.12, FR.05), ask: which date? which type?
- If date given, filter data and list matching records
- If user says "all reports" or "tüm raporlar", list all available
- For production reports: show producedAt, recipeName, labelSeq, status
- For storage check reports: show date, storage, temperature
- Format dates clearly, use emojis for readability

EXAMPLES:
User: "23 Nisan raporları" → Ask: "Hangi rapor? 1.Üretim 2.Depo 3.Hepsi"
User: "Hepsi" → List all records from April 23
User: "FR.06 getir" → List production records, ask date if needed`;

      // Context: son 10 mesaj
      const history=[...active.messages,userMsg].slice(-10).map(m=>m.role==="user"?`K: ${m.text}`:`A: ${m.text}`).join("\n");
      const raw=await callAI(apiKey,sys,history+"\n\nSon soruma yanıt ver.","haiku");
      const aiMsg={role:"assistant",text:raw,ts:new Date().toISOString()};
      // Güncel active'den al (bu arada bot mesajı eklenmiş olabilir)
      setConversations(p=>p.map(c=>c.id===active.id?{...c,messages:[...c.messages,aiMsg],updatedAt:new Date().toISOString()}:c));
    }catch(e){
      const errMsg={role:"assistant",text:"Hata: "+e.message,ts:new Date().toISOString()};
      setConversations(p=>p.map(c=>c.id===active.id?{...c,messages:[...c.messages,errMsg],updatedAt:new Date().toISOString()}:c));
    }
    setL(false);
  };

  // Aktif sohbete bot mesajları entegre et (yalnızca render için)
  const renderMessages=[...(active?.messages||[])];
  // Bot mesajları aktif sohbetin mesajlarıyla zaman sırasına göre birleştirilmiş olarak render edilecek
  const allMsgs=[...renderMessages,...(botMessages||[]).map(b=>({role:"bot",text:b.text,ts:b.ts,icon:b.icon||"🤖"}))].sort((a,b)=>new Date(a.ts||0)-new Date(b.ts||0));

  // Sohbet listesi panel
  const ListPanel=()=>(
    <div style={{borderRight:`1px solid ${t.border}`,height:"100%",overflowY:"auto",padding:"10px 8px"}}>
      <button onClick={newChat} style={{...bSt("p",t),width:"100%",marginBottom:10,fontSize:13}}>+ {t.L.newChat||"Yeni Sohbet"}</button>
      {activeConversations.map(c=>{
        const msgCount=c.messages.length;
        const preview=c.messages[c.messages.length-1]?.text?.slice(0,40)||"";
        return <div key={c.id} onClick={()=>{setActiveConvId(c.id);setView("chat")}} style={{padding:"10px 12px",marginBottom:6,borderRadius:10,cursor:"pointer",background:activeConvId===c.id?t.acB:"transparent",border:`1px solid ${activeConvId===c.id?t.acBo:"transparent"}`}}>
          <div style={{fontSize:13,fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
          {preview&&<div style={{fontSize:11,color:t.tm,marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{preview}</div>}
          <div style={{fontSize:10,color:t.tm,marginTop:2}}>{msgCount} {t.L.msgCount||"mesaj"}</div>
        </div>;
      })}
      {archivedConversations.length>0&&<div style={{marginTop:12,paddingTop:10,borderTop:`1px solid ${t.border}`}}>
        <button onClick={()=>setShowArchived(!showArchived)} style={{...bSt("s",t),width:"100%",fontSize:11,padding:"6px 10px"}}>
          🗂 {t.L.archived||"Arşivlenenler"} ({archivedConversations.length}) {showArchived?"▲":"▼"}
        </button>
        {showArchived&&archivedConversations.map(c=>(
          <div key={c.id} style={{padding:"8px 10px",marginTop:6,borderRadius:8,background:t.inBg,fontSize:12}}>
            <div style={{color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
            <div style={{display:"flex",gap:4,marginTop:4}}>
              <button onClick={()=>unarchiveConv(c.id)} style={{...bSt("s",t),fontSize:10,padding:"3px 8px",flex:1}}>↩ {t.L.unarchive||"Geri Al"}</button>
              <button onClick={()=>deleteConv(c.id)} style={{...bSt("d",t),fontSize:10,padding:"3px 8px"}}>×</button>
            </div>
          </div>
        ))}
      </div>}
    </div>
  );

  return <div style={{display:"flex",height:"calc(100dvh - 200px)",minHeight:400,gap:0,paddingBottom:"env(safe-area-inset-bottom)"}}>
    {/* Desktop: iki sütun, Mobil: tek */}
    <div style={{width:260,flexShrink:0,display:view==="list"?"block":"none"}} className="chat-list-panel-mobile"><ListPanel/></div>
    <div style={{width:260,flexShrink:0,display:"none"}} className="chat-list-panel-desktop"><ListPanel/></div>

    <div style={{flex:1,display:view==="chat"||window.innerWidth>700?"flex":"none",flexDirection:"column",padding:"0 8px"}}>
      {active&&<>
        <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderBottom:`1px solid ${t.border}`,marginBottom:10}}>
          <button onClick={()=>setView("list")} style={{...bSt("s",t),fontSize:13,padding:"6px 10px"}} className="chat-back-btn">☰</button>
          {editingName===active.id?
            <input autoFocus defaultValue={active.name} onBlur={e=>renameConv(active.id,e.target.value)} onKeyDown={e=>{if(e.key==="Enter")renameConv(active.id,e.target.value)}} style={{...iSt(t),flex:1,fontSize:14,padding:"6px 10px"}}/>
            :<div onClick={()=>setEditingName(active.id)} style={{flex:1,fontSize:15,fontWeight:600,color:t.text,cursor:"pointer",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{active.name}</div>
          }
          <button onClick={()=>archiveConv(active.id)} title={t.L.archive||"Arşivle"} style={{...bSt("s",t),fontSize:13,padding:"6px 10px"}}>📦</button>
          <button onClick={()=>deleteConv(active.id)} title={t.L.delete||"Sil"} style={{...bSt("d",t),fontSize:13,padding:"6px 10px"}}>🗑</button>
        </div>

        <div style={{flex:1,overflowY:"auto",padding:"0 0 16px",display:"flex",flexDirection:"column",gap:10}}>
          {allMsgs.length===0&&<div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
            <div style={{fontSize:40,marginBottom:12,opacity:0.5}}>💬</div>
            <div style={{fontSize:14}}>{t.L.chatEmptyHint||"Menü planla, maliyet sor, alışveriş listesi çıkar..."}</div>
          </div>}
          {allMsgs.map((m,i)=>{
            const isUser=m.role==="user";
            const isBot=m.role==="bot";
            const bg=isUser?t.accent:(isBot?t.waBg:t.card);
            const color=isUser?(t.bg==="#1a1612"?"#1a1612":"#fff"):(isBot?t.warn:t.text);
            const border=isBot?`1px solid ${t.waBo}`:(isUser?"none":`1px solid ${t.cardB}`);
            return <div key={i} style={{alignSelf:isUser?"flex-end":"flex-start",maxWidth:"85%",background:bg,color,borderRadius:isUser?"16px 16px 4px 16px":"16px 16px 16px 4px",padding:"12px 16px",fontSize:14,lineHeight:1.6,border,whiteSpace:"pre-wrap"}}>
              {isBot?<span>{m.icon||"🤖"} </span>:null}{m.text}
            </div>;
          })}
          {loading&&<div style={{alignSelf:"flex-start",background:t.card,border:`1px solid ${t.cardB}`,borderRadius:16,padding:"12px 16px",fontSize:14,color:t.tm}}>⏳ ...</div>}
          <div ref={endRef}/>
        </div>
        <div style={{display:"flex",gap:8}}>
          <input value={input} onChange={e=>setIn(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")send()}} placeholder={t.L.askQuestion||"Soru sor..."} style={{...iSt(t),flex:1}}/>
          <button onClick={send} disabled={loading} style={{...bSt("p",t),padding:"11px 18px"}}>↑</button>
        </div>
      </>}
    </div>
  </div>;
};

// ═══ SETTINGS ═══
const SettingsTab=({apiKey,setApiKey,dark,setDark,lang,setLang,recipes,stock,invoices,setRecipes,setStock,setInvoices,expenses,setExpenses,storageAreas,setStorageAreas,profile,setProfile,traceability,setTraceability,trackedIngs,setTrackedIngs,resetHour,setResetHour,organizations,setOrganizations,notifSettings,setNotifSettings,printers,setPrinters,setBotMessages,calorieDB,setCalorieDB,user,setUser,authRequired,setAuthRequired,setShowAuth,handleLogout,externalSection,setExternalSection,team,setTeam,teamMembers,setTeamMembers,wallpaper,setWallpaper,customWP,setCustomWP,t})=>{
  const[msg,setMsg]=useState("");
  const[section,setSection_]=useState(externalSection||null);
  const setSection=(s)=>{setSection_(s);if(setExternalSection)setExternalSection(s);};
  const fileRef=useRef();
  const flash=m=>{setMsg(m);setTimeout(()=>setMsg(""),3000)};

  const backup=()=>{
    const d=JSON.stringify({recipes,stock,invoices,profile,expenses,storageAreas,version:"2.0"},null,2);
    const b=new Blob([d],{type:"application/json"});
    const u=URL.createObjectURL(b);const a=document.createElement("a");
    a.href=u;a.download=`kitchenmanager-${new Date().toISOString().split("T")[0]}.json`;a.click();
    URL.revokeObjectURL(u);flash("✓ "+(t.L.backup||"Yedek"));
  };
  const restore=f=>{const r=new FileReader();r.onload=e=>{try{const d=JSON.parse(e.target.result);if(!window.confirm(t.L.settingsRestoreConfirm))return;if(d.recipes)setRecipes(d.recipes);if(d.stock)setStock(d.stock);if(d.invoices)setInvoices(d.invoices);if(d.profile)setProfile(d.profile);if(d.expenses)setExpenses(d.expenses);if(d.storageAreas)setStorageAreas(d.storageAreas);flash("✓ "+t.L.settingsRestoreSuccess)}catch(e){flash("Hata: "+e.message)}};r.readAsText(f)};
  const exportLogs=()=>{
    const txt=JSON.stringify(LOGS,null,2);
    if(navigator.clipboard)navigator.clipboard.writeText(txt);
    flash(`✓ ${LOGS.length} log`);
  };

  // Test bot mesajı
  const sendTestNotif=()=>{
    setBotMessages(p=>[...p,{
      id:"bot_test_"+Date.now(),
      key:"test_"+Date.now(),
      text:"✅ "+(lang==="tr"?"Test bildirimi — sistem çalışıyor":lang==="en"?"Test notification — system working":"Test"),
      icon:"✅",
      ts:new Date().toISOString()
    }]);
    flash("✓ "+(t.L.testNotif||"Test"));
  };

  const sections=[
    {id:"appearance",label:lang==="tr"?"Görünüm":lang==="en"?"Appearance":"Appearance",icon:"🎨",sub:lang==="tr"?"Arka plan, tema":"Wallpaper, theme"},
    {id:"user",label:lang==="tr"?"Hesap & Profil":lang==="en"?"Account & Profile":lang==="ru"?"Аккаунт":lang==="es"?"Cuenta":lang==="de"?"Konto":lang==="fr"?"Compte":lang==="zh"?"账户":"الحساب",icon:"👤",sub:user?user.email:(lang==="tr"?"Giriş yapılmadı":"Not signed in")},
    {id:"language",label:lang==="tr"?"Dil":lang==="en"?"Language":lang==="ru"?"Язык":lang==="es"?"Idioma":lang==="de"?"Sprache":lang==="fr"?"Langue":lang==="zh"?"语言":"اللغة",icon:"🌐",sub:LANGS.find(l=>l.code===lang)?.fullName||lang.toUpperCase()},
    {id:"storage",label:lang==="tr"?"Depolama Alanları":lang==="en"?"Storage Areas":lang==="ru"?"Хранилище":lang==="es"?"Almacenes":lang==="de"?"Lagerbereiche":lang==="fr"?"Stockage":lang==="zh"?"存储":"التخزين",icon:"🗄",sub:`${storageAreas.length} ${lang==="tr"?"alan":lang==="en"?"areas":""}`},
    {id:"iso",label:"ISO 22000",icon:"🔍",sub:traceability?(lang==="tr"?"Aktif":"Active"):(lang==="tr"?"Pasif":"Inactive")},
    {id:"expenses",label:lang==="tr"?"İşletme Giderleri":lang==="en"?"Expenses":lang==="ru"?"Расходы":lang==="es"?"Gastos":lang==="de"?"Kosten":lang==="fr"?"Dépenses":lang==="zh"?"开支":"المصاريف",icon:"💼",sub:`${expenses.fixed.length+expenses.personnel.length} ${lang==="tr"?"kalem":"items"}`},
    {id:"notif",label:lang==="tr"?"Bildirimler":lang==="en"?"Notifications":lang==="ru"?"Уведомления":lang==="es"?"Notificaciones":lang==="de"?"Benachrichtigungen":lang==="fr"?"Notifications":lang==="zh"?"通知":"الإشعارات",icon:"🔔",sub:notifSettings.enabled?(lang==="tr"?"Açık":"On"):(lang==="tr"?"Kapalı":"Off")},
    {id:"printers",label:lang==="tr"?"Yazıcılar":lang==="en"?"Printers":lang==="ru"?"Принтеры":lang==="es"?"Impresoras":lang==="de"?"Drucker":lang==="fr"?"Imprimantes":lang==="zh"?"打印机":"الطابعات",icon:"🖨",sub:`${printers.length} ${lang==="tr"?"yazıcı":"printers"}`},
    {id:"teamJoin",label:lang==="tr"?"Ekip":lang==="en"?"Team":lang==="ru"?"Команда":lang==="es"?"Equipo":lang==="de"?"Team":lang==="fr"?"Équipe":lang==="zh"?"团队":"الفريق",icon:"👥",sub:team?team.name:(lang==="tr"?"Ekip yok":lang==="en"?"No team":"—")}
  ];

  // Section başlığı (geri butonu ile)
  const SectionHeader=({title})=><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:18,paddingBottom:14,borderBottom:`1px solid ${t.border}`}}>
    <button onClick={()=>setSection(null)} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 10px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {lang==="tr"?"Ayarlar":lang==="en"?"Settings":"‹"}</button>
    <h3 style={{fontSize:18,color:t.text,margin:0,fontWeight:600}}>{title}</h3>
  </div>;

  // ANA MENÜ — section seçili değilse göster
  if(!section){
    return <div style={{maxWidth:520,margin:"0 auto"}}>
      {msg&&<div style={{background:t.sucBg,border:`1px solid ${t.sucBo}`,borderRadius:12,padding:"10px 14px",marginBottom:16,fontSize:13,color:t.success}}>{msg}</div>}

      <h2 style={{fontSize:24,color:t.text,marginBottom:6,fontFamily:"'Fraunces',serif",fontWeight:600}}>{lang==="tr"?"Ayarlar":lang==="en"?"Settings":lang==="ru"?"Настройки":lang==="es"?"Ajustes":lang==="de"?"Einstellungen":lang==="fr"?"Paramètres":lang==="zh"?"设置":"الإعدادات"}</h2>
      <div style={{fontSize:12,color:t.tm,marginBottom:18}}>{user?user.email:(lang==="tr"?"Giriş yapılmadı":"Not signed in")}</div>

      {/* iOS tarzı liste - gruplandırılmış */}
      <div style={{...cSt(t),padding:0,overflow:"hidden",marginBottom:14}}>
        {sections.map((s,i)=><button key={s.id} onClick={()=>setSection(s.id)} style={{
          width:"100%",
          display:"flex",
          alignItems:"center",
          gap:14,
          padding:"14px 16px",
          background:"transparent",
          border:"none",
          borderTop:i>0?`1px solid ${t.border}`:"none",
          cursor:"pointer",
          textAlign:"left",
          transition:"background 0.15s"
        }} onMouseOver={e=>e.currentTarget.style.background=t.inBg} onMouseOut={e=>e.currentTarget.style.background="transparent"}>
          <span style={{fontSize:22,width:32,textAlign:"center",flexShrink:0}}>{s.icon}</span>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:15,fontWeight:500,color:t.text,marginBottom:2}}>{s.label}</div>
            {s.sub&&<div style={{fontSize:11,color:t.tm,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.sub}</div>}
          </div>
          <span style={{color:t.tm,fontSize:18,opacity:0.5,flexShrink:0}}>›</span>
        </button>)}
      </div>

      <div style={{textAlign:"center",fontSize:12,color:t.tm,marginTop:24,padding:"16px 0"}}>
        <div style={{fontWeight:600,color:t.ts,fontFamily:"'Fraunces',serif",fontSize:14}}>Kitchen Manager Pro</div>
        <div style={{fontSize:10,marginTop:4,opacity:0.7}}>v1.0.0 · by Tulpar Kitchen Software</div>
      </div>
    </div>;
  }

  return <div style={{maxWidth:520,margin:"0 auto"}}>
    {msg&&<div style={{background:t.sucBg,border:`1px solid ${t.sucBo}`,borderRadius:12,padding:"10px 14px",marginBottom:16,fontSize:13,color:t.success}}>{msg}</div>}

    {/* DİL SECILMIŞ */}
    {section==="appearance"&&<div>
      <SectionHeader title={lang==="tr"?"Görünüm":"Appearance"}/>
      <div style={{...cSt(t),padding:16,marginBottom:12}}>
        <div style={{fontSize:13,fontWeight:600,color:t.text,marginBottom:12}}>{lang==="tr"?"Arka Plan":"Wallpaper"}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8,marginBottom:6}}>
          {WALLPAPERS.filter(w=>w.id!=="custom").map(wp=><button key={wp.id} onClick={()=>{
            setWallpaper(wp.id);localStorage.setItem("kmp_wallpaper",wp.id);
          }} style={{aspectRatio:"1",borderRadius:10,border:wallpaper===wp.id?`2px solid ${t.accent}`:`2px solid ${t.border}`,cursor:"pointer",overflow:"hidden",position:"relative",...(wp.id==="default"?{background:t.bg}:wp.style),minHeight:40}}>
            {wallpaper===wp.id&&<span style={{position:"absolute",bottom:2,right:2,fontSize:7,background:t.accent,color:"#fff",borderRadius:"50%",width:14,height:14,display:"flex",alignItems:"center",justifyContent:"center"}}>✓</span>}
          </button>)}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8,marginBottom:14}}>
          {WALLPAPERS.filter(w=>w.id!=="custom").map(wp=><div key={wp.id} style={{fontSize:7,color:t.tm,textAlign:"center",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{wp.label[lang]||wp.label.en}</div>)}
        </div>
        <label style={{...bSt("s",t),display:"flex",alignItems:"center",justifyContent:"center",gap:6,cursor:"pointer",width:"100%",padding:"10px 0"}}>
          📷 {lang==="tr"?"Fotoğraf Yükle":"Upload Photo"}
          <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
            const file=e.target.files?.[0];if(!file)return;
            const reader=new FileReader();
            reader.onload=ev=>{
              const img=new Image();
              img.onload=()=>{
                // Resize: max kenar 1280px, quality 0.75 → ~200-400KB
                const MAX=1280;
                let w=img.width,h=img.height;
                if(w>MAX||h>MAX){
                  if(w>h){h=Math.round(h*MAX/w);w=MAX;}
                  else{w=Math.round(w*MAX/h);h=MAX;}
                }
                const canvas=document.createElement("canvas");
                canvas.width=w;canvas.height=h;
                canvas.getContext("2d").drawImage(img,0,0,w,h);
                const url=canvas.toDataURL("image/jpeg",0.75);
                try{
                  localStorage.setItem("kmp_customwp",url);
                  localStorage.setItem("kmp_wallpaper","custom");
                  setCustomWP(url);setWallpaper("custom");
                }catch(err){
                  window.toast.info(lang==="tr"?"Fotoğraf çok büyük, daha küçük seçin":"Photo too large, try a smaller one");
                }
              };
              img.onerror=()=>window.toast.info(lang==="tr"?"Fotoğraf okunamadı":"Could not read photo");
              img.src=ev.target.result;
            };
            reader.readAsDataURL(file);e.target.value="";
          }}/>
        </label>
        {wallpaper==="custom"&&customWP&&<div style={{marginTop:8,position:"relative"}}>
          <img src={customWP} style={{width:"100%",height:80,objectFit:"cover",borderRadius:8}} alt="wallpaper"/>
          <button onClick={()=>{setWallpaper("default");setCustomWP("");localStorage.removeItem("kmp_wallpaper");localStorage.removeItem("kmp_customwp");}} style={{position:"absolute",top:4,right:4,background:"rgba(0,0,0,0.6)",border:"none",borderRadius:"50%",width:20,height:20,color:"#fff",cursor:"pointer",fontSize:11}}>✕</button>
        </div>}
      </div>
    </div>}
    {section==="language"&&<div>
      <SectionHeader title={lang==="tr"?"Dil":lang==="en"?"Language":"Language"}/>
      <div style={{...cSt(t),padding:0,overflow:"hidden"}}>
        {LANGS.map((ln,i)=><button key={ln.code} onClick={()=>setLang(ln.code)} style={{
          width:"100%",display:"flex",alignItems:"center",gap:14,padding:"14px 16px",
          background:lang===ln.code?t.acB:"transparent",border:"none",
          borderTop:i>0?`1px solid ${t.border}`:"none",cursor:"pointer",textAlign:"left"
        }}>
          <span style={{fontSize:24,flexShrink:0}}>{ln.flag}</span>
          <div style={{flex:1}}>
            <div style={{fontSize:15,fontWeight:500,color:t.text}}>{ln.fullName||ln.label}</div>
            <div style={{fontSize:11,color:t.tm}}>{ln.label}</div>
          </div>
          {lang===ln.code&&<span style={{color:t.accent,fontSize:18,fontWeight:700}}>✓</span>}
        </button>)}
      </div>
    </div>}

    {/* GÖRÜNÜM (TEMA) */}
    {/* ═══ KULLANICI / HESAP ═══ */}
    {section==="user"&&<div>
      <SectionHeader title={lang==="tr"?"Hesap & Profil":"Account & Profile"}/>
      {/* GİRİŞ SİSTEMİ */}
      <div style={{...cSt(t),padding:"14px 16px",marginBottom:20}}>
        {user?<>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
            <div style={{width:42,height:42,borderRadius:"50%",background:t.acB,color:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:700}}>{(user.name||user.email||"U")[0].toUpperCase()}</div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:14,fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.name||user.email}</div>
              <div style={{fontSize:11,color:t.tm}}>{user.email} {user.verified?"✓":"⚠️ "+(lang==="tr"?"doğrulanmamış":"unverified")}</div>
            </div>
            <button onClick={()=>{
              if(window.confirm(lang==="tr"?"Çıkış yapılsın mı?":"Logout?")){handleLogout()}
            }} style={{...bSt("d",t),fontSize:11,padding:"6px 10px"}}>{lang==="tr"?"Çıkış":"Logout"}</button>
          </div>
        </>:<>
          <div style={{fontSize:13,color:t.tm,marginBottom:10}}>{lang==="tr"?"Hesabınızla giriş yaparak verilerinize her cihazdan ulaşın":"Login to access your data from any device"}</div>
          <button onClick={()=>setShowAuth(true)} style={{...bSt("p",t),width:"100%"}}>{lang==="tr"?"Giriş / Kayıt":"Login / Sign Up"}</button>
        </>}
      </div>

      <div style={{fontSize:13,fontWeight:600,color:t.tm,marginBottom:10,letterSpacing:"0.05em"}}>{t.L.userProfile}</div>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr",gap:10}}>
          <div><label style={lSt(t)}>{t.L.fullName}</label><input style={iSt(t)} placeholder={t.L.placeholderName||t.L.settingsProfileNamePh} value={profile.fullName} onChange={e=>setProfile({...profile,fullName:e.target.value})}/></div>
          <div><label style={lSt(t)}>{t.L.workplace}</label><input style={iSt(t)} placeholder={t.L.placeholderWorkplace||t.L.settingsProfileWorkplacePh} value={profile.workplace} onChange={e=>setProfile({...profile,workplace:e.target.value})}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label style={lSt(t)}>{lang==="tr"?"Ülke":"Country"}</label>
              <select style={iSt(t)} value={profile.country||""} onChange={e=>setProfile({...profile,country:e.target.value})}>
                <option value="">{lang==="tr"?"-- Seçin --":"-- Select --"}</option>
                {COUNTRIES.map(c=><option key={c.code} value={c.code}>{c.flag} {c.name}</option>)}
              </select>
            </div>
            <div><label style={lSt(t)}>{lang==="tr"?"Sektör":"Sector"}</label>
              <select style={iSt(t)} value={profile.sector||""} onChange={e=>setProfile({...profile,sector:e.target.value})}>
                <option value="">{lang==="tr"?"-- Seçin --":"-- Select --"}</option>
                {SECTORS.map(s=><option key={s.id} value={s.id}>{s.emoji} {s[lang]||s.en}</option>)}
              </select>
            </div>
          </div>
          <div><label style={lSt(t)}>{lang==="tr"?"Ölçü Birimi Sistemi":"Measurement System"}</label>
            <div style={{display:"flex",gap:6}}>
              {[
                {id:"metric",label:lang==="tr"?"📏 Metrik (g, ml, °C)":"📏 Metric (g, ml, °C)"},
                {id:"imperial",label:lang==="tr"?"📐 İmperyal (oz, fl oz, °F)":"📐 Imperial (oz, fl oz, °F)"}
              ].map(o=>{const on=(profile.unitSystem||"metric")===o.id;return <button key={o.id} onClick={()=>setProfile({...profile,unitSystem:o.id})} style={{flex:1,padding:"10px 8px",borderRadius:8,fontSize:12,fontWeight:600,border:`1px solid ${on?t.accent:t.border}`,background:on?t.accent+"22":t.inBg,color:on?t.accent:t.text,cursor:"pointer"}}>{o.label}</button>;})}
            </div>
          </div>
          <div><label style={lSt(t)}>{t.L.department}</label><input style={iSt(t)} placeholder={t.L.placeholderDept||t.L.settingsProfileDeptPh} value={profile.department} onChange={e=>setProfile({...profile,department:e.target.value})}/></div>
          <div><label style={lSt(t)}>{t.L.role} <span style={{color:t.accent,fontSize:10,fontWeight:700}}>(PRO)</span></label>
            <select style={iSt(t)} value={ROLE_HIERARCHY[profile.role]?profile.role:""}
              onChange={e=>setProfile({...profile,role:e.target.value})}>
              <option value="">{lang==="tr"?"-- Görev Seçin --":"-- Select Role --"}</option>
              <optgroup label={lang==="tr"?"👑 Üst Yönetim (Pro)":"👑 Top Management (Pro)"}>
                {PRO_ROLES.map(r=><option key={r} value={r}>{ROLE_HIERARCHY[r].icon} {ROLE_HIERARCHY[r].label[lang]||ROLE_HIERARCHY[r].label.en}</option>)}
              </optgroup>
            </select>
            <div style={{fontSize:10,color:t.tm,marginTop:6,lineHeight:1.5}}>
              💡 {lang==="tr"?"Pro yöneticisi olarak alt kademedeki herkese görev atayabilir, ekipler kurabilirsiniz.":"As Pro manager, you can assign tasks to all subordinates and create teams."}
            </div>
          </div>
        </div>
        <div style={{fontSize:10,color:t.tm,marginTop:8,lineHeight:1.5}}>{t.L.profileHint}</div>
        {/* Supabase'deki adı güncelle */}
        {user&&<button onClick={async()=>{
          if(!profile.fullName?.trim()){window.toast.info(lang==="tr"?"Ad girin":"Enter name");return;}
          const sb=initSupabase();
          if(sb){
            await sb.auth.updateUser({data:{name:profile?.fullName?.trim(),full_name:profile?.fullName?.trim()}});
            // team_members tablosundaki adı da güncelle
            const{data:{session}}=await sb.auth.getSession();
            if(session?.user?.id&&team?.id){
              await sb.from("team_members").update({name:profile?.fullName?.trim()})
                .eq("user_id",session.user.id).eq("team_id",team.id);
            }
          }
          flash(lang==="tr"?"✓ Ad güncellendi":"✓ Name updated");
        }} style={{...bSt("p",t),width:"100%",marginTop:10,fontSize:13}}>
          {lang==="tr"?"✓ Adı Kaydet":"✓ Save Name"}
        </button>}
      </div>
    </div>}

    {/* ═══ DEPOLAMA ALANLARI ═══ */}
    {section==="storage"&&<div>
      <SectionHeader title={t.L.storageAreas}/>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        {storageAreas.map((s,i)=><div key={s.id} style={{display:"flex",gap:6,marginBottom:8,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:20}}>{STORAGE_TYPES.find(x=>x.id===s.type)?.icon||"📦"}</span>
          <input style={{...iSt(t),flex:"2 1 120px",minWidth:100}} placeholder={t.L.settingsStorageName} value={s.name} onChange={e=>{const a=[...storageAreas];a[i]={...s,name:e.target.value};setStorageAreas(a)}}/>
          <select style={{...iSt(t),flex:"2 1 120px",minWidth:110}} value={s.type} onChange={e=>{const a=[...storageAreas];const typ=STORAGE_TYPES.find(x=>x.id===e.target.value);a[i]={...s,type:e.target.value,temp:typ?typ.defaultTemp:s.temp};setStorageAreas(a)}}>
            {STORAGE_TYPES.map(typ=><option key={typ.id} value={typ.id}>{typ.icon} {storageTypeL(typ,lang)}</option>)}
          </select>
          <input style={{...iSt(t),width:70,textAlign:"center"}} type="number" placeholder="°C" value={s.temp} onChange={e=>{const a=[...storageAreas];a[i]={...s,temp:parseFloat(e.target.value)||0};setStorageAreas(a)}}/>
          <button onClick={()=>setStorageAreas(storageAreas.filter((_,j)=>j!==i))} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
        </div>)}
        <button onClick={()=>setStorageAreas([...storageAreas,{id:"st"+Date.now(),name:t.L.settingsStorageName,type:"fridge",temp:4,capacity:null}])} style={{...bSt("s",t),fontSize:12,width:"100%",marginTop:4}}>{t.L.addStorage}</button>
        <div style={{fontSize:10,color:t.tm,marginTop:8,lineHeight:1.5}}>{t.L.storageHint}</div>
      </div>
    </div>}

    {/* ═══ ISO 22000 (izlenebilirlik + organizasyon + numune) ═══ */}
    {section==="iso"&&<div>
      <SectionHeader title="ISO 22000"/>
      <div style={{...cSt(t),padding:"14px 16px",marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
          <div style={{flex:1}}>
            <div style={{fontSize:14,fontWeight:700,color:t.text}}>{t.L.traceabilityMode}</div>
            <div style={{fontSize:11,color:t.tm,marginTop:2}}>{t.L.traceabilityDesc}</div>
          </div>
          <label style={{position:"relative",display:"inline-block",width:48,height:26,marginLeft:12}}>
            <input type="checkbox" checked={traceability} onChange={e=>setTraceability(e.target.checked)} style={{opacity:0,width:0,height:0}}/>
            <span style={{position:"absolute",cursor:"pointer",top:0,left:0,right:0,bottom:0,background:traceability?t.accent:"#ccc",borderRadius:13,transition:"0.3s"}}>
              <span style={{position:"absolute",height:20,width:20,left:traceability?25:3,bottom:3,background:"white",borderRadius:"50%",transition:"0.3s"}}/>
            </span>
          </label>
        </div>
        {traceability&&<>
          <div style={{borderTop:`1px solid ${t.border}`,paddingTop:12,marginTop:6}}>
            <label style={lSt(t)}>{t.L.resetHourLabel}</label>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <input type="number" min="0" max="23" value={resetHour} onChange={e=>setResetHour(Math.min(23,Math.max(0,parseInt(e.target.value,10)||23)))} style={{...iSt(t),width:80,textAlign:"center",fontWeight:700}}/>
              <span style={{fontSize:12,color:t.tm}}>:00 — {t.L.resetHourDesc}</span>
            </div>
          </div>
          <div style={{marginTop:14}}>
            <label style={lSt(t)}>{t.L.trackedIngsLabel}</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
              {trackedIngs.map((ing,i)=><span key={i} style={{background:t.acB,color:t.accent,padding:"4px 10px",borderRadius:12,fontSize:12,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6}}>
                {ing}
                <span onClick={()=>setTrackedIngs(trackedIngs.filter((_,j)=>j!==i))} style={{cursor:"pointer",opacity:0.6}}>×</span>
              </span>)}
            </div>
            <div style={{display:"flex",gap:6}}>
              <input id="newTrackedIng" style={{...iSt(t),flex:1}} placeholder={t.L.trackedIngsPh} onKeyDown={e=>{if(e.key==="Enter"&&e.target.value.trim()){setTrackedIngs([...trackedIngs,e.target.value.trim()]);e.target.value=""}}}/>
              <button onClick={()=>{const inp=document.getElementById("newTrackedIng");if(inp.value.trim()){setTrackedIngs([...trackedIngs,inp.value.trim()]);inp.value=""}}} style={{...bSt("s",t),fontSize:12}}>{t.L.stockAddBtn}</button>
            </div>
            <div style={{fontSize:10,color:t.tm,marginTop:6,lineHeight:1.5}}>{t.L.trackedIngsHint}</div>
          </div>
        </>}
      </div>

      {/* Organizasyonlar (Numune için) */}
      <h4 style={{fontSize:16,marginBottom:10,color:t.text}}>{t.L.organizationsTitle}</h4>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
          {organizations.map((org,i)=><span key={i} style={{background:t.acB,color:t.accent,padding:"4px 10px",borderRadius:12,fontSize:12,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6}}>
            {org}
            <span onClick={()=>setOrganizations(organizations.filter((_,j)=>j!==i))} style={{cursor:"pointer",opacity:0.6}}>×</span>
          </span>)}
        </div>
        <div style={{display:"flex",gap:6}}>
          <input id="newOrg" style={{...iSt(t),flex:1}} placeholder={t.L.organizationsPh} onKeyDown={e=>{if(e.key==="Enter"&&e.target.value.trim()){setOrganizations([...organizations,e.target.value.trim()]);e.target.value=""}}}/>
          <button onClick={()=>{const inp=document.getElementById("newOrg");if(inp.value.trim()){setOrganizations([...organizations,inp.value.trim()]);inp.value=""}}} style={{...bSt("s",t),fontSize:12}}>{t.L.stockAddBtn}</button>
        </div>
        <div style={{fontSize:10,color:t.tm,marginTop:6,lineHeight:1.5}}>{t.L.organizationsHint}</div>
      </div>
    </div>}

    {/* ═══ İŞLETME GİDERLERİ ═══ */}
    {section==="expenses"&&<div>
      <SectionHeader title={t.L.expensesTitle}/>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        <div style={{fontSize:12,fontWeight:700,color:t.tm,marginBottom:8}}>{t.L.fixedExpenses}</div>
        {expenses.fixed.map((it,i)=><div key={i} style={{display:"flex",gap:6,marginBottom:6}}>
          <input style={{...iSt(t),flex:2}} placeholder={t.L.settingsExpenseName} value={it.name} onChange={e=>{const a=[...expenses.fixed];a[i]={...it,name:e.target.value};setExpenses({...expenses,fixed:a})}}/>
          <input style={{...iSt(t),flex:1,textAlign:"right"}} type="number" placeholder="₺" value={it.amount||""} onChange={e=>{const a=[...expenses.fixed];a[i]={...it,amount:parseFloat(e.target.value)||0};setExpenses({...expenses,fixed:a})}}/>
          <button onClick={()=>setExpenses({...expenses,fixed:expenses.fixed.filter((_,j)=>j!==i)})} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
        </div>)}
        <button onClick={()=>setExpenses({...expenses,fixed:[...expenses.fixed,{name:"",amount:0}]})} style={{...bSt("s",t),fontSize:12,width:"100%",marginTop:4}}>{t.L.addFixed}</button>

        <div style={{fontSize:12,fontWeight:700,color:t.tm,marginTop:16,marginBottom:8}}>{t.L.personnel}</div>
        {expenses.personnel.map((p,i)=><div key={i} style={{display:"flex",gap:6,marginBottom:6}}>
          <input style={{...iSt(t),flex:2}} placeholder={t.L.settingsPersonnelName} value={p.name} onChange={e=>{const a=[...expenses.personnel];a[i]={...p,name:e.target.value};setExpenses({...expenses,personnel:a})}}/>
          <input style={{...iSt(t),flex:2}} placeholder={t.L.role} value={p.role||""} onChange={e=>{const a=[...expenses.personnel];a[i]={...p,role:e.target.value};setExpenses({...expenses,personnel:a})}}/>
          <input style={{...iSt(t),flex:1,textAlign:"right"}} type="number" placeholder="₺" value={p.salary||""} onChange={e=>{const a=[...expenses.personnel];a[i]={...p,salary:parseFloat(e.target.value)||0};setExpenses({...expenses,personnel:a})}}/>
          <button onClick={()=>setExpenses({...expenses,personnel:expenses.personnel.filter((_,j)=>j!==i)})} style={{...bSt("d",t),padding:"6px 10px",fontSize:11}}>×</button>
        </div>)}
        <button onClick={()=>setExpenses({...expenses,personnel:[...expenses.personnel,{name:"",role:"",salary:0}]})} style={{...bSt("s",t),fontSize:12,width:"100%",marginTop:4}}>{t.L.addPersonnel}</button>

        <div style={{marginTop:16,padding:"10px 12px",background:t.acB,borderRadius:10}}>
          <label style={{fontSize:11,fontWeight:700,color:t.accent}}>{t.L.monthlyPortions}</label>
          <input type="number" min="1" value={expenses.monthlyPortions||1000} onChange={e=>setExpenses({...expenses,monthlyPortions:Math.max(1,parseInt(e.target.value,10)||1000)})} style={{...iSt(t),marginTop:4,textAlign:"center",fontWeight:700}}/>
          <div style={{fontSize:10,color:t.tm,marginTop:4}}>{t.L.portionHint}</div>
        </div>

        {(()=>{
          const totalFixed=expenses.fixed.reduce((a,x)=>a+(x.amount||0),0);
          const totalPersonnel=expenses.personnel.reduce((a,x)=>a+(x.salary||0),0);
          const total=totalFixed+totalPersonnel;
          const mp=expenses.monthlyPortions||1000;
          const perPortion=total/mp;
          return <div style={{marginTop:12,padding:"10px 12px",background:t.waBg,border:`1px solid ${t.waBo}`,borderRadius:10}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:t.warn}}><span>{t.L.fixedTotal}</span><span style={{fontWeight:700}}>₺{totalFixed.toLocaleString("tr-TR")}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:t.warn}}><span>{t.L.personnelTotal}</span><span style={{fontWeight:700}}>₺{totalPersonnel.toLocaleString("tr-TR")}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:14,color:t.warn,fontWeight:700,borderTop:`1px solid ${t.waBo}`,marginTop:6,paddingTop:6}}><span>{t.L.monthlyTotal}</span><span>₺{total.toLocaleString("tr-TR")}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:13,color:t.warn,marginTop:4}}><span>{t.L.perPortionShare}</span><span style={{fontWeight:700}}>₺{perPortion.toFixed(2)}</span></div>
          </div>;
        })()}
      </div>
    </div>}

    {/* ═══ BİLDİRİMLER ═══ */}
    {section==="notif"&&<div>
      <SectionHeader title={t.L.notificationsTitle||(lang==="tr"?"Bildirimler":"Notifications")}/>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,paddingBottom:14,borderBottom:`1px solid ${t.border}`}}>
          <div style={{flex:1}}>
            <div style={{fontSize:14,fontWeight:700,color:t.text}}>{t.L.notifEnabled||"Bot bildirimleri aktif"}</div>
            <div style={{fontSize:11,color:t.tm,marginTop:2}}>{t.L.notifDesc||"Bildirimler sohbet ekranına düşer, pop-up olmaz"}</div>
          </div>
          <label style={{position:"relative",display:"inline-block",width:48,height:26,marginLeft:12}}>
            <input type="checkbox" checked={notifSettings.enabled} onChange={e=>setNotifSettings({...notifSettings,enabled:e.target.checked})} style={{opacity:0,width:0,height:0}}/>
            <span style={{position:"absolute",cursor:"pointer",top:0,left:0,right:0,bottom:0,background:notifSettings.enabled?t.accent:"#ccc",borderRadius:13,transition:"0.3s"}}>
              <span style={{position:"absolute",height:20,width:20,left:notifSettings.enabled?25:3,bottom:3,background:"white",borderRadius:"50%",transition:"0.3s"}}/>
            </span>
          </label>
        </div>
        {[
          {key:"storageCheck",label:t.L.notifStorage||"Depo kontrolü hatırlatma"},
          {key:"expiredSKT",label:t.L.notifExpired||"SKT geçen ürün uyarısı"},
          {key:"lowStock",label:t.L.notifLow||"Düşük stok uyarısı"}
        ].map(opt=><div key={opt.key} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${t.border}`,opacity:notifSettings.enabled?1:0.4}}>
          <span style={{fontSize:13,color:t.text}}>{opt.label}</span>
          <label style={{position:"relative",display:"inline-block",width:40,height:22}}>
            <input type="checkbox" checked={notifSettings[opt.key]} disabled={!notifSettings.enabled} onChange={e=>setNotifSettings({...notifSettings,[opt.key]:e.target.checked})} style={{opacity:0,width:0,height:0}}/>
            <span style={{position:"absolute",cursor:notifSettings.enabled?"pointer":"not-allowed",top:0,left:0,right:0,bottom:0,background:notifSettings[opt.key]&&notifSettings.enabled?t.accent:"#ccc",borderRadius:11,transition:"0.3s"}}>
              <span style={{position:"absolute",height:16,width:16,left:notifSettings[opt.key]?22:3,bottom:3,background:"white",borderRadius:"50%",transition:"0.3s"}}/>
            </span>
          </label>
        </div>)}
        <button onClick={sendTestNotif} style={{...bSt("s",t),width:"100%",marginTop:14}}>{t.L.testNotif||"Test mesajı gönder"}</button>
      </div>
    </div>}

    {/* ═══ YAZICILAR ═══ */}
    {section==="printers"&&<div>
      <SectionHeader title={t.L.printersTitle||(lang==="tr"?"Etiket Yazıcıları":"Label Printers")}/>
      <div style={{...cSt(t),padding:"14px 16px"}}>
        {printers.length===0&&<div style={{fontSize:12,color:t.tm,textAlign:"center",padding:"20px 0"}}>
          {lang==="tr"?"Henüz yazıcı eklenmedi":"No printer added yet"}
        </div>}
        {printers.map((pr,i)=><div key={pr.id} style={{background:t.inBg,borderRadius:10,padding:"10px 12px",marginBottom:8}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
            <input style={iSt(t)} placeholder={t.L.printerName||"Yazıcı adı"} value={pr.name} onChange={e=>{const a=[...printers];a[i]={...pr,name:e.target.value};setPrinters(a)}}/>
            <select style={iSt(t)} value={pr.protocol} onChange={e=>{const a=[...printers];a[i]={...pr,protocol:e.target.value};setPrinters(a)}}>
              <option value="wifi">{t.L.printerWifi||"WiFi"}</option>
              <option value="bluetooth">{t.L.printerBt||"Bluetooth"}</option>
            </select>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
            <input style={iSt(t)} placeholder={t.L.printerAddress||"IP / MAC"} value={pr.address} onChange={e=>{const a=[...printers];a[i]={...pr,address:e.target.value};setPrinters(a)}}/>
            <select style={iSt(t)} value={pr.labelColor} onChange={e=>{const a=[...printers];a[i]={...pr,labelColor:e.target.value};setPrinters(a)}}>
              <option value="white">{t.L.labelWhite||"Beyaz"}</option>
              <option value="blue">{t.L.labelBlue||"Mavi"}</option>
              <option value="customer">{t.L.labelCustomer||"Müşteri QR"}</option>
            </select>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:6}}>
            <select style={iSt(t)} value={pr.area||""} onChange={e=>{const a=[...printers];a[i]={...pr,area:e.target.value};setPrinters(a)}}>
              <option value="">— {t.L.printerArea||"Üretim alanı"} —</option>
              {storageAreas.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <button onClick={()=>setPrinters(printers.filter((_,j)=>j!==i))} style={{...bSt("d",t),fontSize:11}}>× {t.L.delete}</button>
          </div>
        </div>)}
        <button onClick={()=>setPrinters([...printers,{id:"pr_"+Date.now(),name:"",protocol:"wifi",address:"",labelColor:"white",area:""}])} style={{...bSt("s",t),width:"100%",marginTop:8}}>{t.L.addPrinter||"+ Yazıcı Ekle"}</button>
        <div style={{fontSize:10,color:t.tm,marginTop:8,lineHeight:1.5}}>
          {lang==="tr"?"💡 WiFi/Bluetooth yazıcı desteği v2.1'de aktifleşecek. Şimdilik yapılandırma kaydedilir.":"💡 WiFi/Bluetooth printer support activates in v2.1. For now configuration is saved."}
        </div>
      </div>
    </div>}

    {/* ═══ EKİP YÖNETİMİ ═══ */}
    {section==="teamJoin"&&<div>
      <SectionHeader title={lang==="tr"?"Ekip":lang==="en"?"Team":"Team"}/>
      {!team?<div style={{...cSt(t),padding:16,marginBottom:12}}>
          <div style={{fontSize:14,fontWeight:700,color:t.text,marginBottom:10}}>👑 {lang==="tr"?"Ekip Oluştur":"Create Team"}</div>
          <div style={{fontSize:12,color:t.tm,marginBottom:10,lineHeight:1.5}}>
            {lang==="tr"?"Yeni bir ekip oluşturun. Davet koduyla çalışanlarınızı ekleyebilirsiniz.":"Create a new team and invite your staff with the invite code."}
          </div>
          <input style={iSt(t)} placeholder={lang==="tr"?"Ekip adı (örn: Pastane Ekibi)":"Team name (e.g. Pastry Team)"} id="chefTeamNameInput"/>
          <button onClick={async()=>{
            const name=document.getElementById("chefTeamNameInput")?.value?.trim();
            if(!name){window.toast.info(lang==="tr"?"Ekip adı girin":"Enter team name");return;}
            if(!user?.userId){window.toast.info(lang==="tr"?"Önce giriş yapın":"Login first");return;}
            try{
              const sb2=initSupabase();
              let realName=user.name||user.email;
              if(sb2){
                const{data:{session}}=await sb2.auth.getSession();
                if(session?.user)realName=session.user.user_metadata?.name||session.user.user_metadata?.full_name||session.user.email.split("@")[0];
              }
              const newTeam=await createTeam(name,user.userId,realName);
              const newTeam2={...newTeam,role:"chef",inviteCode:newTeam.invite_code};
              setTeam(newTeam2);
              LS.set("kmp_team",newTeam2);
              const chefMember=[{userId:user.userId,name:realName,role:"chef"}];
              setTeamMembers(chefMember);
              LS.set("kmp_team_members",chefMember);
              // Davet kodunu göster ve kopyala
              const code=newTeam.invite_code;
              const shareText=`${lang==="tr"?"Kitchen Manager'a katıl!":"Join Kitchen Manager!"}\n${lang==="tr"?"Ekip":"Team"}: ${name}\n${lang==="tr"?"Davet Kodu":"Invite Code"}: ${code}\n${window.location.origin}${window.location.pathname}`;
              if(navigator.share){navigator.share({title:"Kitchen Manager",text:shareText}).catch(()=>{});}
              else{const ta=document.createElement("textarea");ta.value=shareText;ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.focus();ta.select();try{document.execCommand("copy");}catch{}document.body.removeChild(ta);}
              flash(lang==="tr"?`✓ Ekip oluşturuldu! Davet kodu: ${code} (kopyalandı)`:`✓ Team created! Invite code: ${code} (copied)`);
            }catch(e){window.toast.info(e.message);}
          }} style={{...bSt("p",t),width:"100%",marginTop:10,padding:12}}>
            ✦ {lang==="tr"?"Oluştur":"Create"}
          </button>
      </div>:<div style={{...cSt(t),padding:"16px"}}>
        <div style={{fontSize:16,fontWeight:700,color:t.text,marginBottom:4}}>{team.name}</div>
        <div style={{fontSize:12,color:t.tm,marginBottom:12}}>{team.role==="chef"?"👑 Şef":"👤 Üye"}</div>
        {/* Davet kodu — sadece şef görür */}
        {(team.inviteCode||team.invite_code)&&<div style={{background:t.acB,border:`1px solid ${t.accent}`,borderRadius:12,padding:"12px 14px",marginBottom:12}}>
          <div style={{fontSize:11,color:t.accent,fontWeight:700,marginBottom:6,letterSpacing:"0.05em"}}>{lang==="tr"?"📨 DAVETİYE KODU":"📨 INVITE CODE"}</div>
          <div style={{fontSize:28,fontWeight:900,color:t.accent,letterSpacing:"0.3em",textAlign:"center",marginBottom:8}}>{team.inviteCode||team.invite_code}</div>
          <button onClick={async()=>{
            const code=team.inviteCode||team.invite_code;
            const shareText=`${lang==="tr"?"Kitchen Manager'a katıl!":"Join Kitchen Manager!"}\n${lang==="tr"?"Ekip":"Team"}: ${team.name}\n${lang==="tr"?"Davet Kodu":"Invite Code"}: ${code}\n${window.location.origin}${window.location.pathname.replace("pro.html","index.html")}`;
            if(navigator.share){try{await navigator.share({title:"Kitchen Manager",text:shareText});}catch{}}
            else{const ta=document.createElement("textarea");ta.value=shareText;ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.focus();ta.select();try{document.execCommand("copy");}catch{}document.body.removeChild(ta);flash(lang==="tr"?"✓ Kopyalandı":"✓ Copied");}
          }} style={{...bSt("p",t),width:"100%",fontSize:13}}>
            📤 {lang==="tr"?"Davet Linkini Paylaş":"Share Invite Link"}
          </button>
        </div>}
        {/* Ülke / Bölge ayarı */}
        {team.role==="chef"&&<div style={{...cSt(t),padding:"12px 14px",marginBottom:12}}>
          <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em"}}>🌍 {lang==="tr"?"ÜLKE / BÖLGE":"COUNTRY / REGION"}</div>
          <div style={{fontSize:11,color:t.tm,marginBottom:8,lineHeight:1.4}}>
            {lang==="tr"?"Resmi tatiller, hafta sonu günleri ve varsayılan yıllık izin günleri bu seçime göre belirlenir.":"Holidays, weekend days and default annual leave are determined by this selection."}
          </div>
          <select value={team.country||"TR"} onChange={async e=>{
            const newCountry=e.target.value;
            const newWeekend=getWeekendDays(newCountry);
            const sb=initSupabase();if(!sb)return;
            const{error}=await sb.from("teams").update({country:newCountry,weekend_days:newWeekend}).eq("id",team.id);
            if(error){window.toast.error(error.message);return;}
            const updated={...team,country:newCountry,weekend_days:newWeekend};
            setTeam(updated);
            LS.set("kmp_team",updated);
            window.toast.success(lang==="tr"?"✓ Ülke güncellendi":"✓ Country updated");
          }} style={{...iSt(t),width:"100%",fontSize:13}}>
            {COMMON_COUNTRIES.map(c=><option key={c.code} value={c.code}>{c.flag} {c.name}</option>)}
          </select>
          <div style={{fontSize:10,color:t.tm,marginTop:6,lineHeight:1.4}}>
            {lang==="tr"?`Varsayılan yıllık izin: ${getDefaultAnnualLeave(team.country||"TR")} gün • Hafta sonu: ${(team.weekend_days||[0,6]).map(d=>["Paz","Pzt","Sal","Çar","Per","Cum","Cmt"][d]).join(", ")}`:`Default leave: ${getDefaultAnnualLeave(team.country||"TR")} days • Weekend: ${(team.weekend_days||[0,6]).map(d=>["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][d]).join(", ")}`}
          </div>
        </div>}
        {/* Hızlı Vardiya Şablonları */}
        {team.role==="chef"&&<ShiftPresetsCard team={team} setTeam={setTeam} t={t} lang={lang}/>}
        {/* Üst Ekibe Bağlan */}
        {<div style={{...cSt(t),padding:"12px 14px",marginBottom:12}}>
          <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em"}}>🔗 {lang==="tr"?"ÜST EKİP":"PARENT TEAM"}</div>
          {team.parent_team_id?<div>
            <div style={{fontSize:13,color:t.text,marginBottom:6}}>{lang==="tr"?"Bu ekip ":"This team is linked to "}<strong>{team.parent_team_name||(lang==="tr"?"bir üst ekibe":"a parent team")}</strong>{lang==="tr"?" ekibine bağlı":""}</div>
            <button onClick={async()=>{
              if(!confirm(lang==="tr"?"Üst ekip bağlantısını kaldır?":"Remove parent team link?"))return;
              try{
                const sb=initSupabase();if(!sb)return;
                await sb.from("teams").update({parent_team_id:null}).eq("id",team.id);
                setTeam({...team,parent_team_id:null,parent_team_name:null});
                window.toast.success(lang==="tr"?"✓ Bağlantı kaldırıldı":"✓ Link removed");
              }catch(e){window.toast.error(e.message);}
            }} style={{...bSt("s",t),fontSize:12,padding:"6px 12px"}}>{lang==="tr"?"Bağlantıyı Kaldır":"Unlink"}</button>
          </div>:<div>
            <div style={{fontSize:12,color:t.tm,marginBottom:8,lineHeight:1.5}}>{lang==="tr"?"Daha üst bir ekibe (Pro / Yönetim) bağlanmak için davet kodunu girin.":"Enter invite code to link to a higher team (Pro / Management)."}</div>
            <div style={{display:"flex",gap:6}}>
              <input style={{...iSt(t),flex:1,textTransform:"uppercase",letterSpacing:"0.15em",fontFamily:"monospace"}} placeholder="ABC123" id="parentCodeInput" maxLength={6}/>
              <button onClick={async()=>{
                const code=document.getElementById("parentCodeInput")?.value?.trim();
                if(!code){window.toast.info(lang==="tr"?"Kod girin":"Enter code");return;}
                try{
                  const parent=await linkParentTeam(team.id,code);
                  setTeam({...team,parent_team_id:parent.id,parent_team_name:parent.name});
                  window.toast.success(lang==="tr"?`✓ ${parent.name} ekibine bağlandı`:`✓ Linked to ${parent.name}`);
                  document.getElementById("parentCodeInput").value="";
                }catch(e){window.toast.error(e.message);}
              }} style={{...bSt("p",t),fontSize:12,padding:"8px 14px"}}>↑</button>
            </div>
          </div>}
        </div>}

        {/* Alt Ekipler */}
        {<ChildTeamsSection teamId={team.id} t={t} lang={lang}/>}

        {/* Ekip Üyeleri Listesi (Gerçek + Phantom) */}
        <PhantomMembersSection team={team} teamMembers={teamMembers} phantomMembers={phantomMembers} setPhantomMembers={setPhantomMembers} user={user} t={t} lang={lang}/>
        <button onClick={async()=>{
          if(!team?.id||!user?.userId)return;
          try{
            const[syncStock,syncProd,syncRecipes,syncTodos]=await Promise.all([
              syncFromTeam(team.id,"team_stock"),
              syncFromTeam(team.id,"team_productions"),
              syncFromTeam(team.id,"team_recipes"),
              syncFromTeam(team.id,"team_todos")
            ]);
            if(syncStock&&syncStock.length>0)setStock(syncStock);
            if(syncProd&&syncProd.length>0)setProductions(syncProd);
            if(syncRecipes&&syncRecipes.length>0)setRecipes(syncRecipes);
            if(syncTodos&&syncTodos.length>0)setTodos(syncTodos);
            LS.set("kmp_last_sync",new Date().toISOString());
            flash(lang==="tr"?"✓ Senkronize edildi":"✓ Synced");
          }catch(e){window.toast.info(e.message);}
        }} style={{...bSt("s",t),width:"100%",marginBottom:10}}>
          🔄 {lang==="tr"?"Şimdi Senkronize Et":"Sync Now"}
        </button>
        {LS.get("kmp_last_sync",null)&&<div style={{fontSize:10,color:t.tm,textAlign:"center",marginBottom:12}}>
          {lang==="tr"?"Son sync:":"Last sync:"} {new Date(LS.get("kmp_last_sync","")).toLocaleString()}
        </div>}
        <button onClick={()=>{
          if(window.confirm(lang==="tr"?"Ekipten ayrılmak istediğinize emin misiniz?":"Leave team?")){
            setTeam(null);setTeamMembers([]);
            flash(lang==="tr"?"Ekipten ayrıldınız":"Left team");
          }
        }} style={{...bSt("d",t),width:"100%",fontSize:12}}>
          {lang==="tr"?"Ekipten Ayrıl":"Leave Team"}
        </button>
      </div>}
    </div>}

  </div>;
};

// ═══ TODO TAB ═══
const TodoTab=({todos,setTodos,t})=>{
  const[newText,setNewText]=useState("");
  const[newDate,setNewDate]=useState("");
  const[newTime,setNewTime]=useState("");
  const[newPrio,setNewPrio]=useState("medium");
  const[newType,setNewType]=useState("task"); // task | meeting | recurring
  const[newRecurring,setNewRecurring]=useState("daily"); // daily | weekly | monthly
  const[filter,setFilter]=useState("all");
  const lang=t.lang;

  const addTodo=()=>{
    if(!newText.trim())return;
    const todo={
      id:Date.now(),
      text:newText.trim(),
      done:false,
      priority:newPrio,
      type:newType,
      dueDate:newDate||null,
      dueTime:newTime||null,
      recurring:newType==="recurring"?newRecurring:null,
      createdAt:new Date().toISOString()
    };
    setTodos(p=>[todo,...p]);
    setNewText("");setNewDate("");setNewTime("");setNewPrio("medium");setNewType("task");
  };

  const toggleDone=(id)=>setTodos(p=>p.map(t=>t.id===id?{...t,done:!t.done,doneAt:!t.done?new Date().toISOString():null}:t));
  const deleteTodo=(id)=>setTodos(p=>p.filter(t=>t.id!==id));

  const prioColors={high:"#dc2626",medium:"#d97706",low:"#16a34a"};
  const prioLabels={
    high:{tr:"Yüksek",en:"High",ru:"Высокий",es:"Alta",de:"Hoch",fr:"Haute",zh:"高",ar:"عالية"},
    medium:{tr:"Orta",en:"Medium",ru:"Средний",es:"Media",de:"Mittel",fr:"Moyenne",zh:"中",ar:"متوسطة"},
    low:{tr:"Düşük",en:"Low",ru:"Низкий",es:"Baja",de:"Niedrig",fr:"Basse",zh:"低",ar:"منخفضة"}
  };
  const pL=(p)=>prioLabels[p]?.[lang]||prioLabels[p]?.en||p;

  const L={
    title:{tr:"Görev Listesi",en:"To-Do List",ru:"Список задач",es:"Lista de tareas",de:"Aufgabenliste",fr:"Liste de tâches",zh:"待办事项",ar:"قائمة المهام"}[lang]||"To-Do",
    addPh:{tr:"Yeni görev ekle...",en:"Add new task...",ru:"Добавить задачу...",es:"Añadir tarea...",de:"Aufgabe hinzufügen...",fr:"Ajouter une tâche...",zh:"添加任务...",ar:"إضافة مهمة..."}[lang]||"Add task...",
    add:{tr:"Ekle",en:"Add",ru:"Добавить",es:"Añadir",de:"Hinzufügen",fr:"Ajouter",zh:"添加",ar:"إضافة"}[lang]||"Add",
    all:{tr:"Tümü",en:"All",ru:"Все",es:"Todos",de:"Alle",fr:"Tout",zh:"全部",ar:"الكل"}[lang]||"All",
    active:{tr:"Bekleyen",en:"Active",ru:"Активные",es:"Activos",de:"Aktiv",fr:"Actifs",zh:"进行中",ar:"نشط"}[lang]||"Active",
    done:{tr:"Tamamlanan",en:"Done",ru:"Выполненные",es:"Completados",de:"Erledigt",fr:"Terminés",zh:"已完成",ar:"مكتمل"}[lang]||"Done",
    empty:{tr:"Henüz görev yok",en:"No tasks yet",ru:"Задач пока нет",es:"Sin tareas aún",de:"Noch keine Aufgaben",fr:"Aucune tâche encore",zh:"暂无任务",ar:"لا توجد مهام بعد"}[lang]||"No tasks",
    due:{tr:"Son:",en:"Due:",ru:"Срок:",es:"Vence:",de:"Fällig:",fr:"Échéance:",zh:"截止:",ar:"موعد:"}[lang]||"Due:",
    clearDone:{tr:"Tamamlananları Sil",en:"Clear Done",ru:"Удалить выполненные",es:"Borrar completados",de:"Erledigte löschen",fr:"Supprimer terminés",zh:"清除已完成",ar:"مسح المنجزة"}[lang]||"Clear Done"
  };

  const filtered=todos.filter(td=>
    filter==="all"?true:filter==="done"?td.done:!td.done
  );
  const doneCount=todos.filter(t=>t.done).length;
  const activeCount=todos.length-doneCount;

  return <div style={{maxWidth:520,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
      <h2 style={{fontSize:22,color:t.text,fontFamily:"'Fraunces',serif",fontWeight:600,margin:0}}>{L.title}</h2>
      <div style={{fontSize:12,color:t.tm}}>{activeCount} {L.active}</div>
    </div>

    {/* Yeni görev ekle */}
    <div style={{...cSt(t),padding:"14px 16px",marginBottom:16}}>
      {/* Tip seçici */}
      <div style={{display:"flex",gap:4,marginBottom:10,background:t.inBg,padding:3,borderRadius:10}}>
        {[["task","✅ "+(lang==="tr"?"Görev":"Task")],["meeting","📅 "+(lang==="tr"?"Toplantı":"Meeting")],["recurring","🔁 "+(lang==="tr"?"Tekrarlayan":"Recurring")]].map(([tp,lbl])=>
          <button key={tp} onClick={()=>setNewType(tp)} style={{flex:1,padding:"6px 4px",borderRadius:8,fontSize:11,fontWeight:600,border:"none",cursor:"pointer",background:newType===tp?t.card:"transparent",color:newType===tp?t.text:t.tm}}>{lbl}</button>
        )}
      </div>
      <textarea style={{...iSt(t),minHeight:56,resize:"vertical",marginBottom:10,fontSize:14}}
        placeholder={newType==="meeting"?(lang==="tr"?"Toplantı başlığı...":"Meeting title..."):newType==="recurring"?(lang==="tr"?"Tekrarlayan görev...":"Recurring task..."):L.addPh}
        value={newText} onChange={e=>setNewText(e.target.value)}
        onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();addTodo()}}}/>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <input type="date" style={{...iSt(t),flex:"1 1 120px",fontSize:13}} value={newDate} onChange={e=>setNewDate(e.target.value)}/>
        {newType==="meeting"&&<input type="time" style={{...iSt(t),flex:"1 1 90px",fontSize:13}} value={newTime} onChange={e=>setNewTime(e.target.value)}/>}
        {newType==="recurring"&&<select style={{...iSt(t),flex:"1 1 100px",fontSize:13}} value={newRecurring} onChange={e=>setNewRecurring(e.target.value)}>
          <option value="daily">{lang==="tr"?"Her gün":"Daily"}</option>
          <option value="weekly">{lang==="tr"?"Her hafta":"Weekly"}</option>
          <option value="monthly">{lang==="tr"?"Her ay":"Monthly"}</option>
        </select>}
        <select style={{...iSt(t),flex:"1 1 90px",fontSize:13}} value={newPrio} onChange={e=>setNewPrio(e.target.value)}>
          {Object.keys(prioLabels).map(p=><option key={p} value={p}>{pL(p)}</option>)}
        </select>
        <button onClick={addTodo} disabled={!newText.trim()} style={{...bSt("p",t),flex:"1 1 70px",opacity:newText.trim()?1:0.5}}>{L.add}</button>
      </div>
    </div>

    {/* Filtre */}
    <div style={{display:"flex",gap:6,marginBottom:14,background:t.inBg,padding:3,borderRadius:12}}>
      {[["all",L.all],[" active",L.active],["done",L.done]].map(([f2,lbl])=><button key={f2} onClick={()=>setFilter(f2.trim())} style={{flex:1,padding:"7px",borderRadius:9,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",background:filter===f2.trim()?t.card:"transparent",color:filter===f2.trim()?t.text:t.tm,transition:"all 0.15s"}}>{lbl}</button>)}
    </div>

    {/* Liste */}
    {filtered.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}>
      <div style={{fontSize:36,marginBottom:10}}>✅</div>
      <div style={{fontSize:14}}>{L.empty}</div>
    </div>}

    {filtered.map(td=>{
      const overdue=td.dueDate&&!td.done&&new Date(td.dueDate)<new Date();
      const typeIcon=td.type==="meeting"?"📅":td.type==="recurring"?"🔁":"✅";
      return <div key={td.id} style={{...cSt(t),padding:"12px 14px",marginBottom:8,display:"flex",gap:12,alignItems:"flex-start",opacity:td.done?0.6:1,borderLeft:`3px solid ${prioColors[td.priority]||t.accent}`}}>
        <button onClick={()=>toggleDone(td.id)} style={{
          width:22,height:22,borderRadius:"50%",border:`2px solid ${td.done?t.success:t.inBo}`,
          background:td.done?t.success:"transparent",cursor:"pointer",flexShrink:0,marginTop:2,
          display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:12,fontWeight:700
        }}>{td.done?"✓":""}</button>
        <div style={{flex:1,minWidth:0}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
            <span style={{fontSize:12}}>{typeIcon}</span>
            <div style={{fontSize:14,color:t.text,textDecoration:td.done?"line-through":"none",lineHeight:1.4,wordBreak:"break-word"}}>{td.text}</div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:4,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:10,fontWeight:700,color:prioColors[td.priority],background:prioColors[td.priority]+"15",padding:"2px 7px",borderRadius:6}}>{pL(td.priority)}</span>
            {td.dueDate&&<span style={{fontSize:10,color:overdue?t.danger:t.tm}}>📅 {td.dueDate}{td.dueTime&&" "+td.dueTime}</span>}
            {td.recurring&&<span style={{fontSize:10,color:t.accent}}>🔁 {td.recurring==="daily"?(lang==="tr"?"Her gün":"Daily"):td.recurring==="weekly"?(lang==="tr"?"Her hafta":"Weekly"):(lang==="tr"?"Her ay":"Monthly")}</span>}
          </div>
        </div>
        <button onClick={()=>deleteTodo(td.id)} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:16,padding:"0 4px",flexShrink:0}}>✕</button>
      </div>;
    })}

    {/* Tamamlananları sil */}
    {doneCount>0&&<button onClick={()=>{if(window.confirm(`${doneCount} ${lang==="tr"?"tamamlanan silinecek":"done tasks will be deleted"}. OK?`))setTodos(p=>p.filter(td=>!td.done))}} style={{...bSt("d",t),width:"100%",marginTop:8,fontSize:12}}>{L.clearDone} ({doneCount})</button>}
  </div>;
};

// ═══ KM: EKİP SOHBETİ ═══
const TeamChatTab=({team,user,t})=>{
  const[messages,setMessages]=useState([]);
  const[newMsg,setNewMsg]=useState("");
  const[loading,setLoading]=useState(true);
  const[sending,setSending]=useState(false);
  const endRef=useRef(null);
  const lang=t.lang;

  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    sb.from("team_messages").select("*").eq("team_id",team.id).is("private_to",null)
      .order("created_at",{ascending:true}).limit(100)
      .then(({data})=>{if(data)setMessages(data);setLoading(false);
        setTimeout(()=>endRef.current?.scrollIntoView({behavior:"smooth"}),100);});
    const channel=sb.channel(`km-chat-${team.id}`)
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"team_messages",filter:`team_id=eq.${team.id}`},
        (payload)=>{if(payload.new.private_to)return;setMessages(p=>p.find(m=>m.id===payload.new.id)?p:[...p,payload.new]);
          setTimeout(()=>endRef.current?.scrollIntoView({behavior:"smooth"}),50);})
      .subscribe();
    return()=>{sb.removeChannel(channel);};
  },[user?.userId,user?.id]);

  const sendMessage=async()=>{
    if(!newMsg.trim()||!team?.id||!user?.userId)return;
    setSending(true);
    const sb=initSupabase();
    if(sb)await sb.from("team_messages").insert({
      team_id:team.id,user_id:user.userId,
      user_name:user.name||user.email||"?",text:newMsg.trim()
    });
    setNewMsg("");setSending(false);
  };

  const fmtTime=(iso)=>{
    const d=new Date(iso);const now=new Date();
    const isToday=d.toDateString()===now.toDateString();
    const time=`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
    return isToday?time:`${d.getDate()}/${d.getMonth()+1} ${time}`;
  };

  if(!team)return <div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
    <div style={{fontSize:40,marginBottom:12}}>👥</div>
    <div style={{fontSize:14,color:t.tm}}>{lang==="tr"?"Ekip sohbeti için ayarlardan bir ekibe katılın":lang==="en"?"Join a team in settings to use team chat":"Join a team first"}</div>
    <div style={{fontSize:12,color:t.accent,marginTop:8}}>⚙ → {lang==="tr"?"Ekip":"Team"}</div>
  </div>;

  return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)",maxWidth:520,margin:"0 auto"}}>
    <div style={{fontSize:16,fontWeight:700,color:t.text,marginBottom:8,fontFamily:"'Fraunces',serif"}}>
      🏠 {team.name}
    </div>
    <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:8,paddingBottom:8}}>
      {loading&&<div style={{textAlign:"center",padding:20,color:t.tm,fontSize:13}}>⏳</div>}
      {messages.map((msg,i)=>{
        const isMe=msg.user_id===user?.userId;
        const prev=messages[i-1];
        const showName=!isMe&&(!prev||prev.user_id!==msg.user_id);
        return <div key={msg.id} style={{display:"flex",flexDirection:"column",alignItems:isMe?"flex-end":"flex-start"}}>
          {showName&&<div style={{fontSize:10,color:t.tm,marginBottom:2,marginLeft:12}}>{msg.user_name}</div>}
          <div style={{maxWidth:"80%",background:isMe?t.accent:t.card,color:isMe?"#fff":t.text,
            borderRadius:isMe?"18px 18px 4px 18px":"18px 18px 18px 4px",
            padding:"9px 13px",fontSize:14,lineHeight:1.4,
            border:isMe?"none":`1px solid ${t.border}`,wordBreak:"break-word"}}>
            {msg.text}
          </div>
          <div style={{fontSize:9,color:t.tm,marginTop:2,marginLeft:isMe?0:12,marginRight:isMe?12:0}}>
            {fmtTime(msg.created_at)}
          </div>
        </div>;
      })}
      <div ref={endRef}/>
    </div>
    <div style={{display:"flex",gap:8,paddingTop:10,borderTop:`1px solid ${t.border}`}}>
      <input style={{...iSt(t),flex:1,borderRadius:20,padding:"10px 16px"}}
        placeholder={lang==="tr"?"Mesaj yaz...":lang==="en"?"Write a message...":"..."}
        value={newMsg} onChange={e=>setNewMsg(e.target.value)}
        onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMessage();}}}/>
      <button onClick={sendMessage} disabled={!newMsg.trim()||sending}
        style={{...bSt("p",t),borderRadius:"50%",width:42,height:42,padding:0,flexShrink:0,
          opacity:!newMsg.trim()||sending?0.5:1,fontSize:18}}>↑</button>
    </div>
  </div>;
};

// ═══ AUTH MODAL (Supabase hazırlığı - şu an local mod) ═══
const AuthModal=({onClose,onLogin,t})=>{
  const urlMode=new URLSearchParams(window.location.search).get("mode");
  // Supabase hash fragment'ten recovery token'ı oku
  const hashParams=new URLSearchParams(window.location.hash.replace("#",""));
  const isRecovery=hashParams.get("type")==="recovery"||urlMode==="reset"||localStorage.getItem("km_password_recovery")==="true";
  // PASSWORD_RECOVERY eventi varsa reset modunu aç
  const initialMode=isRecovery?"reset":"login";
  const[mode,setMode]=useState(initialMode);
  const[email,setEmail]=useState("");
  const[name,setName]=useState("");
  const[password,setPassword]=useState("");
  const[password2,setPassword2]=useState("");
  const[loading,setLoading]=useState(false);
  const[err,setErr]=useState("");
  const[info,setInfo]=useState("");
  const lang=t.lang;

  const handleLogin=async()=>{
    if(!email||!password){setErr(lang==="tr"?"Email ve şifre gerekli":"Email and password required");return}
    setLoading(true);setErr("");
    try{
      const sb=initSupabase();
      if(!sb){setErr(lang==="tr"?"Supabase yüklenemedi":"Supabase failed");setLoading(false);return}
      const{data,error}=await sb.auth.signInWithPassword({email,password});
      if(error)throw error;
      // Kullanıcının gerçek adını al
      const displayName=data.user.user_metadata?.name||data.user.user_metadata?.full_name||data.user.email.split("@")[0];
      onLogin({
        email:data.user.email,
        name:displayName,
        verified:!!data.user.email_confirmed_at,
        userId:data.user.id,
        accessToken:data.session.access_token
      });
    }catch(e){
      setErr(e.message||(lang==="tr"?"Giriş başarısız":"Login failed"));
    }
    setLoading(false);
  };

  const handleSignup=async()=>{
    if(!email||!password){setErr(lang==="tr"?"Email ve şifre gerekli":"Email and password required");return}
    if(!name.trim()){setErr(lang==="tr"?"Ad Soyad gerekli":"Name is required");return}
    if(password.length<6){setErr(lang==="tr"?"Şifre en az 6 karakter":"Password min 6 chars");return}
    if(password!==password2){setErr(lang==="tr"?"Şifreler eşleşmiyor":"Passwords don't match");return}
    setLoading(true);setErr("");
    try{
      const sb=initSupabase();
      if(!sb){setErr(lang==="tr"?"Supabase yüklenemedi":"Supabase failed");setLoading(false);return}
      const{data,error}=await sb.auth.signUp({
        email,password,
        options:{data:{name:name.trim(),full_name:name.trim()}}
      });
      if(error)throw error;
      if(data.user&&!data.user.email_confirmed_at){
        setMode("verify");
        setInfo(lang==="tr"
          ?"Email'inize doğrulama linki gönderildi. Linke tıkladıktan sonra giriş yapabilirsiniz."
          :"Verification link sent to your email. Click the link to login.");
      }else{
        onLogin({
          email:data.user.email,
          name:name.trim(),
          verified:true,
          userId:data.user.id,
          accessToken:data.session?.access_token
        });
      }
    }catch(e){
      setErr(e.message||(lang==="tr"?"Kayıt başarısız":"Signup failed"));
    }
    setLoading(false);
  };

  const handleVerify=()=>{
    setMode("login");
    setInfo(lang==="tr"?"Email'inizdeki linke tıkladıktan sonra giriş yapabilirsiniz.":"After clicking link in email, you can login.");
  };

  const handleForgot=async()=>{
    if(!email){setErr(lang==="tr"?"Email girin":"Enter email");return}
    setLoading(true);setErr("");
    try{
      const sb=initSupabase();
      if(!sb){setErr(lang==="tr"?"Supabase yüklenemedi":"Supabase failed");setLoading(false);return}
      const{error}=await sb.auth.resetPasswordForEmail(email,{
        redirectTo:`${window.location.origin}${window.location.pathname}?mode=reset`
      });
      if(error)throw error;
      setInfo(lang==="tr"
        ?"Şifre sıfırlama linki email'inize gönderildi."
        :"Password reset link sent to your email.");
    }catch(e){
      setErr(e.message||(lang==="tr"?"İşlem başarısız":"Failed"));
    }
    setLoading(false);
  };

  const handleResetPassword=async()=>{
    if(!password||password.length<6){setErr(lang==="tr"?"Şifre en az 6 karakter":"Min 6 chars");return}
    if(password!==password2){setErr(lang==="tr"?"Şifreler eşleşmiyor":"Passwords don't match");return}
    setLoading(true);setErr("");
    try{
      const sb=initSupabase();
      if(!sb){setErr("Supabase failed");setLoading(false);return}
      // URL'deki token ile session kur (eğer henüz yoksa)
      const params=new URLSearchParams(window.location.search);
      const token=params.get("token");
      const{data:sessionData}=await sb.auth.getSession();
      if(!sessionData?.session&&token){
        const{error:vErr}=await sb.auth.verifyOtp({token_hash:token,type:"recovery"});
        if(vErr)throw vErr;
      }
      const{error}=await sb.auth.updateUser({password});
      if(error)throw error;
      // Flag temizle
      if(typeof window!=="undefined")delete window.__kmPasswordRecovery;
      localStorage.removeItem("km_password_recovery");
      setInfo(lang==="tr"?"✓ Şifreniz güncellendi! Giriş yapabilirsiniz.":"✓ Password updated! You can login now.");
      setTimeout(()=>setMode("login"),2000);
    }catch(e){
      setErr(e.message||(lang==="tr"?"İşlem başarısız":"Failed"));
    }
    setLoading(false);
  };

  return <div style={mOv(t)}><div onClick={e=>e.stopPropagation()} style={{...mPn(t),maxWidth:400}}>
    <div style={{textAlign:"center",marginBottom:18}}>
      <div style={{fontSize:28,fontFamily:"'Fraunces',serif",fontWeight:700,color:t.text}}>Kitchen Manager <span style={{fontSize:14,color:"#fff",background:`linear-gradient(135deg,${t.accent} 0%,#8b6332 100%)`,padding:"3px 8px",borderRadius:5,letterSpacing:"0.1em",fontWeight:800,marginLeft:4,verticalAlign:"middle"}}>PRO</span></div>
      <div style={{fontSize:11,color:t.tm,marginTop:4}}>by Tulpar Kitchen Software</div>
    </div>
    <div style={{display:"flex",gap:4,marginBottom:16,background:t.inBg,padding:3,borderRadius:10}}>
      <button onClick={()=>{setMode("login");setErr("");setInfo("")}} style={{flex:1,padding:"8px",borderRadius:8,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:mode==="login"?t.card:"transparent",color:mode==="login"?t.text:t.tm}}>{lang==="tr"?"Giriş":"Login"}</button>
      <button onClick={()=>{setMode("signup");setErr("");setInfo("")}} style={{flex:1,padding:"8px",borderRadius:8,fontSize:13,fontWeight:600,border:"none",cursor:"pointer",background:mode==="signup"?t.card:"transparent",color:mode==="signup"?t.text:t.tm}}>{lang==="tr"?"Kayıt":"Signup"}</button>
    </div>

    {(mode==="login"||mode==="signup"||mode==="forgot")&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
      {mode==="signup"&&<input style={iSt(t)} type="text" placeholder={lang==="tr"?"Ad Soyad *":"Full Name *"} value={name} onChange={e=>setName(e.target.value)}/>}
      <input style={iSt(t)} type="email" placeholder="email@example.com" value={email} onChange={e=>setEmail(e.target.value)}/>
      {mode!=="forgot"&&<input style={iSt(t)} type="password" placeholder={lang==="tr"?"Şifre":"Password"} value={password} onChange={e=>setPassword(e.target.value)}/>}
      {mode==="signup"&&<input style={iSt(t)} type="password" placeholder={lang==="tr"?"Şifre (tekrar)":"Confirm password"} value={password2} onChange={e=>setPassword2(e.target.value)}/>}
    </div>}

    {mode==="verify"&&<div>
      <div style={{fontSize:13,color:t.ts,marginBottom:10,textAlign:"center",lineHeight:1.5}}>{lang==="tr"?"📧 Email'inize doğrulama linki gönderildi. Linke tıkladıktan sonra giriş yapabilirsiniz.":"📧 Verification link sent to your email. After clicking, you can login."}</div>
    </div>}

    {mode==="reset"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div style={{fontSize:13,color:t.ts,marginBottom:4,textAlign:"center"}}>🔑 {lang==="tr"?"Yeni şifrenizi girin":"Enter your new password"}</div>
      <input style={iSt(t)} type="password" placeholder={lang==="tr"?"Yeni şifre":"New password"} value={password} onChange={e=>setPassword(e.target.value)}/>
      <input style={iSt(t)} type="password" placeholder={lang==="tr"?"Şifreyi tekrarla":"Confirm password"} value={password2} onChange={e=>setPassword2(e.target.value)}/>
    </div>}

    {err&&<div style={{color:t.danger,fontSize:13,marginTop:10,background:t.danBg,borderRadius:10,padding:"8px 12px"}}>{err}</div>}
    {info&&<div style={{color:t.accent,fontSize:12,marginTop:10,background:t.acB,borderRadius:10,padding:"8px 12px",lineHeight:1.4}}>ℹ️ {info}</div>}

    <div style={{marginTop:16}}>
      {mode==="login"&&<button onClick={handleLogin} disabled={loading} style={{...bSt("p",t),width:"100%",padding:12,opacity:loading?0.6:1}}>{loading?"⏳":(lang==="tr"?"Giriş Yap":"Login")}</button>}
      {mode==="signup"&&<button onClick={handleSignup} disabled={loading} style={{...bSt("p",t),width:"100%",padding:12,opacity:loading?0.6:1}}>{loading?"⏳":(lang==="tr"?"Kayıt Ol":"Sign Up")}</button>}
      {mode==="forgot"&&<button onClick={handleForgot} disabled={loading} style={{...bSt("p",t),width:"100%",padding:12,opacity:loading?0.6:1}}>{loading?"⏳":(lang==="tr"?"Link Gönder":"Send Link")}</button>}
      {mode==="verify"&&<button onClick={handleVerify} style={{...bSt("p",t),width:"100%",padding:12}}>{lang==="tr"?"Giriş Ekranına Dön":"Back to Login"}</button>}
      {mode==="reset"&&<button onClick={handleResetPassword} disabled={loading} style={{...bSt("p",t),width:"100%",padding:12,opacity:loading?0.6:1}}>{loading?"⏳":(lang==="tr"?"Şifreyi Güncelle":"Update Password")}</button>}
    </div>

    {mode==="login"&&<button onClick={()=>{setMode("forgot");setErr("");setInfo("")}} style={{background:"none",border:"none",color:t.accent,fontSize:12,cursor:"pointer",width:"100%",marginTop:10,padding:6}}>{lang==="tr"?"Şifremi unuttum":"Forgot password?"}</button>}
    {onClose&&mode!=="reset"&&<button onClick={onClose} style={{...bSt("s",t),width:"100%",marginTop:8}}>{lang==="tr"?"Kapat":"Close"}</button>}
  </div></div>;
};

// ═══ SOHBET SİSTEMİ ═══
const getOrCreateGroupConv=async(teamId,teamName,userId,userName,userRole)=>{
  const sb=initSupabase();if(!sb)return null;
  // Ekip grubu var mı?
  const{data:convs}=await sb.from("conversations")
    .select("*,conversation_members(*)")
    .eq("team_id",teamId).eq("type","group");
  if(convs&&convs.length>0){
    // Üye değilsem ekle
    const conv=convs[0];
    const isMember=conv.conversation_members?.find(m=>m.user_id===userId);
    if(!isMember){
      await sb.from("conversation_members").insert({
        conversation_id:conv.id,user_id:userId,user_name:userName,user_role:userRole
      });
    }
    return conv;
  }
  // Yeni grup oluştur
  const{data:newConv,error}=await sb.from("conversations").insert({
    team_id:teamId,type:"group",name:teamName,created_by:userId
  }).select().single();
  if(error)throw error;
  await sb.from("conversation_members").insert({
    conversation_id:newConv.id,user_id:userId,user_name:userName,user_role:userRole
  });
  return newConv;
};

const getOrCreatePrivateConv=async(teamId,myId,myName,myRole,otherId,otherName,otherRole)=>{
  const sb=initSupabase();if(!sb)return null;
  // Mevcut özel konuşma var mı?
  const{data:myConvs}=await sb.from("conversation_members")
    .select("conversation_id").eq("user_id",myId);
  if(myConvs&&myConvs.length>0){
    const myConvIds=myConvs.map(c=>c.conversation_id);
    const{data:shared}=await sb.from("conversation_members")
      .select("conversation_id").eq("user_id",otherId).in("conversation_id",myConvIds);
    if(shared&&shared.length>0){
      const{data:privConv}=await sb.from("conversations")
        .select("*").eq("id",shared[0].conversation_id).eq("type","private").single();
      if(privConv)return privConv;
    }
  }
  // Yeni özel konuşma
  const{data:newConv,error}=await sb.from("conversations").insert({
    team_id:teamId,type:"private",name:null,created_by:myId
  }).select().single();
  if(error)throw error;
  await sb.from("conversation_members").insert([
    {conversation_id:newConv.id,user_id:myId,user_name:myName,user_role:myRole},
    {conversation_id:newConv.id,user_id:otherId,user_name:otherName,user_role:otherRole}
  ]);
  return newConv;
};

// ═══ CHAT TAB COMPONENT ═══
const ChatTab=({team,teamMembers,user,recipes,menus,stock,productions,t})=>{
  const[messages,setMessages]=useState([]);
  const[newMsg,setNewMsg]=useState("");
  const[loading,setLoading]=useState(false);
  const[sending,setSending]=useState(false);
  const[showShare,setShowShare]=useState(false);
  const endRef=useRef(null);
  const lang=t.lang;

  const myName=user?.name||user?.email||"?";
  const myRole=(()=>{try{return JSON.parse(localStorage.getItem("kmp_profile")||"{}")?.role||"";}catch(e){return "";}})();
  const[chatError,setChatError]=useState("");

  const L={
    chats:{tr:"Sohbetler",en:"Chats",ru:"Чаты",es:"Chats",de:"Chats",fr:"Discussions",zh:"聊天",ar:"المحادثات"}[lang]||"Chats",
    group:{tr:"Ekip Grubu",en:"Team Group",ru:"Группа",es:"Grupo",de:"Gruppe",fr:"Groupe",zh:"群组",ar:"المجموعة"}[lang]||"Team Group",
    private:{tr:"Özel",en:"Private",ru:"Личное",es:"Privado",de:"Privat",fr:"Privé",zh:"私聊",ar:"خاص"}[lang]||"Private",
    newPrivate:{tr:"Yeni Özel Sohbet",en:"New Private Chat",ru:"Новый чат",es:"Nueva conversación",de:"Neuer Chat",fr:"Nouvelle discussion",zh:"新私聊",ar:"محادثة جديدة"}[lang]||"New Chat",
    send:{tr:"Gönder",en:"Send",ru:"Отправить",es:"Enviar",de:"Senden",fr:"Envoyer",zh:"发送",ar:"İrsal"}[lang]||"Send",
    share:{tr:"Paylaş",en:"Share",ru:"Поделиться",es:"Compartir",de:"Teilen",fr:"Partager",zh:"分享",ar:"مشاركة"}[lang]||"Share",
    noTeam:{tr:"Ekibe katılın",en:"Join a team first",ru:"Присоединитесь к команде",es:"Únete a un equipo",de:"Team beitreten",fr:"Rejoignez l'équipe",zh:"先加入团队",ar:"انضم لفريق"}[lang]||"Join a team"
  };

  useEffect(()=>{
    if(!team?.id||!user?.userId)return;
    setLoading(true);
    setChatError("");
    const sb=initSupabase();
    if(!sb){setChatError("Supabase bağlanamadı");setLoading(false);return;}
    // team_messages tablosundan yükle
    sb.from("team_messages").select("*")
      .eq("team_id",team.id)
      .is("private_to",null)
      .order("created_at",{ascending:true})
      .limit(100)
      .then(({data,error})=>{
        if(error){setChatError("Hata: "+error.message);setLoading(false);return;}
        if(data)setMessages(data);
        setLoading(false);
        setTimeout(()=>endRef.current?.scrollIntoView({behavior:"smooth"}),100);
      });
    // Realtime
    const channel=sb.channel(`chat-${team.id}`)
      .on("postgres_changes",{
        event:"INSERT",schema:"public",
        table:"team_messages",
        filter:`team_id=eq.${team.id}`
      },(payload)=>{
        if(payload.new.private_to)return;
        setMessages(p=>p.find(m=>m.id===payload.new.id)?p:[...p,payload.new]);
        setTimeout(()=>endRef.current?.scrollIntoView({behavior:"smooth"}),50);
      }).subscribe();
    return()=>{sb.removeChannel(channel);};
  },[team?.id,user?.userId]);

  const sendMsg=async(type="text",text=newMsg,attachment=null)=>{
    if(!text.trim()&&!attachment)return;
    if(!team?.id)return;
    setSending(true);
    const sb=initSupabase();
    if(sb){
      // Supabase session'dan user_id al — en güvenilir kaynak
      const{data:{session}}=await sb.auth.getSession();
      const uid=session?.user?.id||user?.userId||user?.id;
      const uname=session?.user?.user_metadata?.name||session?.user?.user_metadata?.full_name||session?.user?.email?.split("@")[0]||myName;
      if(!uid){setSending(false);return;}
      const profile=JSON.parse(localStorage.getItem("kmp_profile")||"{}");
      await sb.from("team_messages").insert({
        team_id:team.id,
        user_id:uid,
        user_name:uname,
        user_role:profile.role||"",
        type,text:text.trim(),attachment
      });
    }
    if(type==="text")setNewMsg("");
    setSending(false);
    setShowShare(false);
  };

  const fmtTime=(iso)=>{
    const d=new Date(iso);const now=new Date();
    const isToday=d.toDateString()===now.toDateString();
    return isToday?`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`:
      `${d.getDate()}/${d.getMonth()+1}`;
  };

  const getConvName=(conv)=>{
    if(conv.type==="group")return conv.name||L.group;
    const other=conv.conversation_members?.find(m=>m.user_id!==user?.userId);
    return other?`${other.user_name}${other.user_role?" · "+other.user_role:""}`:L.private;
  };

  const getConvIcon=(conv)=>conv.type==="group"?"👥":"👤";

  const myUserId=user?.userId||user?.id||"";
  const myName2=user?.name||user?.email||"";
  const[sessionUid,setSessionUid]=useState("");
  useEffect(()=>{
    const sb=initSupabase();if(!sb)return;
    sb.auth.getSession().then(({data:{session}})=>{
      if(session?.user?.id)setSessionUid(session.user.id);
    });
  },[]);
  const isMe=(uid,uname)=>{
    const effectiveUid=sessionUid||myUserId;
    if(uid&&effectiveUid&&uid===effectiveUid)return true;
    if(!effectiveUid&&uname&&myName2&&uname===myName2)return true;
    return false;
  };

  if(!team||!user?.userId)return <div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
    <div style={{fontSize:40,marginBottom:12}}>💬</div>
    <div style={{fontSize:14}}>{L.noTeam}</div>
    <div style={{fontSize:12,color:t.accent,marginTop:8}}>⚙ → 👥 {lang==="tr"?"Ekip":"Team"}</div>
  </div>;

  const otherMembers=(teamMembers||[]).filter(m=>m.userId!==myUserId&&m.user_id!==myUserId);

  return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)",maxWidth:520,margin:"0 auto"}}>
    {/* Header */}
    <div style={{marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${t.border}`}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:otherMembers.length>0?8:0}}>
        <div style={{width:36,height:36,borderRadius:"50%",background:t.acB,color:t.accent,
          display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>👥</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:15,fontWeight:700,color:t.text}}>{team.name}</div>
          <div style={{fontSize:10,color:t.tm}}>{lang==="tr"?"Ekip Grubu":"Team Group"} · {(teamMembers||[]).length} {lang==="tr"?"kişi":"members"}</div>
        </div>
        {chatError&&<div style={{fontSize:10,color:t.danger}}>⚠️</div>}
        {loading&&<div style={{fontSize:12,color:t.tm}}>⏳</div>}
      </div>
      {/* Ekip üyeleri - özel mesaj */}
      {otherMembers.length>0&&<div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:4}}>
        {otherMembers.map((m,i)=><button key={m.userId||m.user_id||i} onClick={async()=>{
          // Özel konuşma başlat — team_messages'da type:private olarak gönder
          const otherUid=m.userId||m.user_id;
          const otherName=m.name||"?";
          window.toast.info(lang==="tr"
            ?`💬 ${otherName} ile özel mesajlaşma yakında gelecek!\n\nŞimdilik grup sohbeti üzerinden iletişim kurabilirsiniz.`
            :`💬 Private chat with ${otherName} coming soon!\n\nFor now, use the group chat.`);
        }} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3,padding:"6px 10px",borderRadius:12,border:`1px solid ${t.border}`,background:t.inBg,cursor:"pointer",flexShrink:0,minWidth:60}}>
          <div style={{width:28,height:28,borderRadius:"50%",background:t.acB,color:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700}}>
            {(m.name||"?")[0].toUpperCase()}
          </div>
          <span style={{fontSize:9,color:t.tm,maxWidth:55,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.name||"?"}</span>
          {m.role&&<span style={{fontSize:8,color:t.accent,fontWeight:600}}>{m.role==="chef"?"👑":""}{m.role}</span>}
        </button>)}
      </div>}
    </div>

    {/* Mesajlar */}
    <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:6,paddingBottom:8}}>
      {messages.map((msg,i)=>{
        const me=isMe(msg.user_id,msg.user_name);
        const prev=messages[i-1];
        const showMeta=!prev||prev.user_id!==msg.user_id;
        const isAttachment=msg.type!=="text"&&msg.attachment;
        const avatarColor=me?t.accent:"#"+Math.abs(msg.user_id?.split("").reduce((a,c)=>a+c.charCodeAt(0),0)||0).toString(16).slice(0,6).padEnd(6,"a");
        return <div key={msg.id} style={{display:"flex",flexDirection:"column",alignItems:me?"flex-end":"flex-start"}}>
          {showMeta&&<div style={{display:"flex",alignItems:"center",gap:6,marginBottom:3,flexDirection:me?"row-reverse":"row",paddingLeft:me?0:4,paddingRight:me?4:0}}>
            <div style={{width:24,height:24,borderRadius:"50%",background:me?t.acB:t.inBg,
              color:me?t.accent:t.ts,display:"flex",alignItems:"center",justifyContent:"center",
              fontSize:11,fontWeight:700,flexShrink:0,border:`1px solid ${me?t.accent:t.border}`}}>
              {(msg.user_name||"?")[0].toUpperCase()}
            </div>
            <div style={{fontSize:11,fontWeight:700,color:me?t.accent:t.text}}>
              {me?(lang==="tr"?"Sen":"You"):msg.user_name}
              {msg.user_role&&<span style={{fontSize:10,color:t.tm,fontWeight:400}}> · {msg.user_role}</span>}
            </div>
          </div>}
          <div style={{maxWidth:"82%",background:me?t.accent:t.card,color:me?"#fff":t.text,
            borderRadius:me?"18px 18px 4px 18px":"4px 18px 18px 18px",
            padding:isAttachment?"10px 13px":"9px 13px",fontSize:14,lineHeight:1.4,
            border:me?"none":`1px solid ${t.border}`,wordBreak:"break-word",
            marginLeft:me?0:30,marginRight:me?30:0,
            cursor:isAttachment?"pointer":"default"}}
            onClick={()=>{
              if(!isAttachment)return;
              if(msg.type==="file"&&msg.attachment?.data){
                // Direkt indir
                const a=document.createElement("a");
                a.href=msg.attachment.data;
                a.download=msg.attachment.name||"download";
                document.body.appendChild(a);a.click();a.remove();
                return;
              }
              // Custom event ile App'e ilet
              window.dispatchEvent(new CustomEvent("km-open-attachment",{detail:{attachment:msg.attachment,type:msg.type}}));
            }}>
            {isAttachment&&<div>
              <div style={{fontSize:10,fontWeight:700,opacity:0.7,marginBottom:4,letterSpacing:"0.05em"}}>
                {msg.type==="recipe"?"🍽 REÇETE":msg.type==="menu"?"📋 MENÜ":msg.type==="report"?"📊 RAPOR":msg.type==="image"?"📸":msg.type==="file"?"📄 "+(lang==="tr"?"DOSYA":"FILE"):"📎"}
                {msg.type!=="image"&&<span style={{float:"right",fontSize:9,opacity:0.6}}>{msg.type==="file"?(lang==="tr"?"İndir →":"Download →"):(lang==="tr"?"Aç →":"Open →")}</span>}
              </div>
              {msg.type==="image"&&msg.attachment?.data?<img src={msg.attachment.data} alt={msg.attachment.name||"photo"} style={{width:"100%",maxWidth:220,borderRadius:8,display:"block",marginTop:4}} onClick={(e)=>{e.stopPropagation();window.open(msg.attachment.data,"_blank");}}/>:
              msg.type==="file"?<><div style={{fontWeight:700,fontSize:13}}>{msg.attachment.name}</div>
              <div style={{fontSize:10,opacity:0.7,marginTop:2}}>{msg.attachment.size?(msg.attachment.size/1024).toFixed(0)+" KB":""}</div></>:
              <><div style={{fontWeight:700,fontSize:13}}>{msg.attachment.name}</div>
              {msg.attachment.desc&&<div style={{fontSize:11,opacity:0.8,marginTop:2}}>{msg.attachment.desc}</div>}</>}
            </div>}
            {msg.text&&<div style={{marginTop:isAttachment?6:0}}>{msg.text}</div>}
          </div>
          <div style={{fontSize:9,color:t.tm,marginTop:2,marginLeft:me?0:34,marginRight:me?34:0}}>
            {fmtTime(msg.created_at)}
          </div>
        </div>;
      })}
      <div ref={endRef}/>
    </div>

    {/* Paylaş paneli */}
    {showShare&&<div style={{...cSt(t),padding:"12px",marginBottom:8,borderRadius:12,maxHeight:200,overflowY:"auto"}}>
      <div style={{fontSize:11,color:t.tm,marginBottom:8,fontWeight:700}}>{lang==="tr"?"Ne paylaşmak istiyorsunuz?":"What would you like to share?"}</div>
      {/* Reçeteler */}
      {(recipes||[]).length>0&&<div style={{marginBottom:8}}>
        <div style={{fontSize:10,color:t.tm,marginBottom:4,fontWeight:600}}>🍽 {lang==="tr"?"Reçeteler":"Recipes"}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {(recipes||[]).map(r=><button key={r.id} onClick={()=>sendMsg("recipe",`🍽 ${r.name}`,{
            type:"recipe",id:r.id,name:r.name,
            desc:`${(r.ingredients||[]).length} ${lang==="tr"?"malzeme":"ingredients"} · ${r.calories||"?"}kcal`,
            data:{name:r.name,ingredients:r.ingredients,steps:r.steps,allergens:r.allergens,calories:r.calories,difficulty:r.difficulty}
          })} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>
            {r.name}
          </button>)}
        </div>
      </div>}
      {/* Menüler */}
      {(menus||[]).length>0&&<div style={{marginBottom:8}}>
        <div style={{fontSize:10,color:t.tm,marginBottom:4,fontWeight:600}}>📋 {lang==="tr"?"Menüler":"Menus"}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {(menus||[]).map(m=><button key={m.id} onClick={()=>sendMsg("menu",`📋 ${m.name}`,{
            type:"menu",id:m.id,name:m.name,desc:m.date||""
          })} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>{m.name}</button>)}
        </div>
      </div>}
      {/* Raporlar */}
      <div style={{marginBottom:8}}>
        <div style={{fontSize:10,color:t.tm,marginBottom:4,fontWeight:600}}>📊 {lang==="tr"?"Raporlar":"Reports"}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          <button onClick={()=>{
            const today=new Date().toISOString().slice(0,10);
            sendMsg("report",`📊 FR.06 · ${today}`,{
              type:"report",name:`FR.06 · ${today}`,
              desc:lang==="tr"?"Günlük üretim raporu":"Daily production report",
              report_type:"production",date:today
            });setShowShare(false);
          }} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>📊 FR.06</button>
          <button onClick={()=>{
            const today=new Date().toISOString().slice(0,10);
            sendMsg("report",`🗄 FR.12 · ${today}`,{
              type:"report",name:`FR.12 · ${today}`,
              desc:lang==="tr"?"Günlük depo kontrol":"Daily storage check",
              report_type:"storage",date:today
            });setShowShare(false);
          }} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>🗄 FR.12</button>
          <button onClick={()=>{
            const today=new Date().toISOString().slice(0,10);
            sendMsg("report",`🏷 FR.05 · ${today}`,{
              type:"report",name:`FR.05 · ${today}`,
              desc:lang==="tr"?"Numune/Şahit etiketi raporu":"Sample/Witness label report",
              report_type:"sample",date:today
            });setShowShare(false);
          }} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>🏷 FR.05</button>
          <button onClick={()=>{
            const today=new Date().toISOString().slice(0,10);
            const lowItems=(stock||[]).filter(s=>s.qty<=(s.low||100));
            const desc=lowItems.length>0
              ?`⚠️ ${lowItems.length} ${lang==="tr"?"ürün kritik seviyede":"items at critical level"}: ${lowItems.slice(0,3).map(s=>s.name).join(", ")}${lowItems.length>3?"...":""}`
              :lang==="tr"?"Tüm stoklar yeterli seviyede":"All stocks at sufficient levels";
            sendMsg("report",`📦 ${lang==="tr"?"Stok Durumu":"Stock Status"} · ${today}`,{
              type:"report",name:`${lang==="tr"?"Stok":"Stock"} · ${today}`,
              desc,report_type:"stock",date:today
            });setShowShare(false);
          }} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>📦 {lang==="tr"?"Stok":"Stock"}</button>
          {(productions||[]).filter(p=>p.date===new Date().toISOString().slice(0,10)).length>0&&
          <button onClick={()=>{
            const today=new Date().toISOString().slice(0,10);
            const todayProds=(productions||[]).filter(p=>p.date===today);
            sendMsg("report",`🍱 ${lang==="tr"?"Bugünkü Üretimler":"Today's Productions"} · ${today}`,{
              type:"report",name:`${lang==="tr"?"Üretim":"Production"} · ${today}`,
              desc:`${todayProds.length} ${lang==="tr"?"üretim kaydı":"production records"}: ${todayProds.slice(0,3).map(p=>p.name).join(", ")}`,
              report_type:"production_list",date:today
            });setShowShare(false);
          }} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>🍱 {lang==="tr"?"Bugün":"Today"}</button>}
        </div>
      </div>
      <button onClick={()=>setShowShare(false)} style={{...bSt("d",t),width:"100%",marginTop:4,fontSize:11,padding:"6px"}}>{lang==="tr"?"İptal":"Cancel"}</button>
    </div>}

    {/* Input */}
    <div style={{display:"flex",gap:6,paddingTop:8,borderTop:`1px solid ${t.border}`,alignItems:"flex-end"}}>
      <button onClick={()=>setShowShare(s=>!s)} style={{...bSt("s",t),padding:"10px 12px",flexShrink:0,borderRadius:12,fontSize:16}}>📎</button>
      {/* Fotoğraf butonu */}
      <label style={{...bSt("s",t),padding:"10px 12px",flexShrink:0,borderRadius:12,fontSize:16,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
        📷
        <input type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={async e=>{
          const file=e.target.files?.[0];
          if(!file)return;
          const reader=new FileReader();
          reader.onload=async(ev)=>{
            const base64=ev.target.result;
            await sendMsg("image","📸 "+file.name,{type:"image",name:file.name,data:base64,size:file.size});
          };
          reader.readAsDataURL(file);
          e.target.value="";
        }}/>
      </label>
      {/* Dosya butonu (PDF, doc vs) */}
      <label style={{...bSt("s",t),padding:"10px 12px",flexShrink:0,borderRadius:12,fontSize:16,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
        📄
        <input type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt" style={{display:"none"}} onChange={async e=>{
          const file=e.target.files?.[0];
          if(!file)return;
          // Limit 4MB (Supabase row limit ~8MB ama base64 ~33% şişer)
          if(file.size>4*1024*1024){
            window.toast.info(lang==="tr"?"Dosya 4MB'dan büyük olamaz":"File must be under 4MB");
            e.target.value="";return;
          }
          const reader=new FileReader();
          reader.onload=async(ev)=>{
            const base64=ev.target.result;
            await sendMsg("file","📄 "+file.name,{type:"file",name:file.name,data:base64,size:file.size,mime:file.type});
          };
          reader.readAsDataURL(file);
          e.target.value="";
        }}/>
      </label>
      <input style={{...iSt(t),flex:1,borderRadius:20,padding:"10px 14px"}}
        placeholder={lang==="tr"?"Mesaj yaz...":"Write a message..."}
        value={newMsg} onChange={e=>setNewMsg(e.target.value)}
        onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMsg();}}}/>
      <button onClick={()=>sendMsg()} disabled={!newMsg.trim()||sending}
        style={{...bSt("p",t),borderRadius:"50%",width:42,height:42,padding:0,flexShrink:0,
          opacity:!newMsg.trim()||sending?0.5:1,fontSize:18}}>↑</button>
    </div>
  </div>;
};

// ═══ GLOBAL MALZEME DB ═══
// Supabase'deki ingredient_db'den kalori ve birim verisi çeker
// AI'a gitmeden önce her zaman buraya bakılır

const ingredientCache={}; // Session cache

const lookupIngredient=async(name)=>{
  if(!name)return null;
  const key=name.toLowerCase().trim();
  if(ingredientCache[key])return ingredientCache[key];
  const sb=initSupabase();if(!sb)return null;
  try{
    // Önce tam eşleşme
    let{data}=await sb.from("ingredient_db").select("*")
      .or(`name_tr.ilike.${key},name_en.ilike.${key}`)
      .limit(1).single();
    if(!data){
      // aliases içinde ara
      const{data:aliasData}=await sb.from("ingredient_db").select("*")
        .contains("aliases",`{${key}}`).limit(1).single();
      data=aliasData;
    }
    if(data){ingredientCache[key]=data;return data;}
    return null;
  }catch(e){return null;}
};

const getKcalFromDB=async(ingredientName)=>{
  const data=await lookupIngredient(ingredientName);
  return data?data.kcal_per_100g:null;
};

const convertUnitToGrams=async(amount,unit,ingredientName=null)=>{
  if(!unit||!amount)return parseFloat(amount)||0;
  const u=unit.toLowerCase().trim();
  const num=parseFloat(String(amount).replace(",","."))||0;
  // Önce UNIT_DB (inline) kontrol
  const inlineResult=parseAmountToGram(`${num} ${u}`);
  if(inlineResult>0)return inlineResult;
  // Supabase unit_conversions'a bak
  const sb=initSupabase();if(!sb)return num;
  try{
    const{data}=await sb.from("unit_conversions").select("*")
      .eq("from_unit",u).eq("to_unit","ml").single();
    if(data){
      const ml=num*data.factor;
      // Sıvı ise density ile gramaj hesapla
      if(ingredientName){
        const ing=await lookupIngredient(ingredientName);
        if(ing?.density_g_per_ml)return ml*ing.density_g_per_ml;
      }
      return ml; // 1ml ≈ 1g yaklaşımı
    }
  }catch(e){}
  return num;
};

// Reçete kaydedilirken global DB'ye kalori katkısı yap
const contributeToGlobalDB=async(ingredientName,kcalPer100g)=>{
  if(!ingredientName||!kcalPer100g)return;
  const sb=initSupabase();if(!sb)return;
  const key=ingredientName.toLowerCase().trim();
  try{
    // Mevcut kayıt var mı?
    const{data:existing}=await sb.from("ingredient_db").select("id,contributed_count")
      .or(`name_tr.ilike.${key},name_en.ilike.${key}`).limit(1).single();
    if(existing){
      await sb.from("ingredient_db").update({contributed_count:(existing.contributed_count||1)+1})
        .eq("id",existing.id);
    }else{
      // Yeni malzeme — anonim katkı
      await sb.from("ingredient_db").insert({
        name_tr:ingredientName,name_en:ingredientName,
        aliases:[key],category:"other",
        kcal_per_100g:kcalPer100g,verified:false
      });
    }
    // Local cache güncelle
    ingredientCache[key]={...ingredientCache[key],kcal_per_100g:kcalPer100g};
  }catch(e){}
};

// ═══ BİRLEŞİK SOHBET (WhatsApp tarzı) ═══
const UnifiedChatTab=({team,teamMembers,user,recipes,menus,stock,productions,apiKey,conversations,setConversations,activeConvId,setActiveConvId,botMessages,storageChecks,t})=>{
  const[activeChat,setActiveChat]=useState(null);
  const[dmMessages,setDmMessages]=useState([]);
  const[dmNew,setDmNew]=useState("");
  const[dmLoading,setDmLoading]=useState(false);
  const dmEndRef=useRef(null);
  const lang=t.lang;
  const myUid=user?.userId||user?.id||"";

  const L={
    title:{tr:"Sohbet",en:"Chats"}[lang]||"Chats",
    ai:{tr:"AI Asistan",en:"AI Assistant"}[lang]||"AI Assistant",
    aiDesc:{tr:"Reçete, rapor ve mutfak soruları",en:"Recipes, reports & kitchen questions"}[lang]||"Kitchen questions",
    teamChat:{tr:"Ekip Grubu",en:"Team Group"}[lang]||"Team Group",
    back:{tr:"Geri",en:"Back"}[lang]||"Back",
    noTeam:{tr:"Ekibe katılın",en:"Join a team"}[lang]||"Join a team"
  };

  const[groupName,setGroupName]=useState("");
  const[groupMembers,setGroupMembers]=useState([]);

  const fmtTime=(iso)=>{
    const d=new Date(iso);const now=new Date();
    const isToday=d.toDateString()===now.toDateString();
    return isToday?`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`:
      `${d.getDate()}/${d.getMonth()+1}`;
  };

  // Grup oluşturma ekranı
  if(activeChat?.newGroup){
    return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,paddingBottom:10,borderBottom:`1px solid ${t.border}`}}>
        <button onClick={()=>setActiveChat(null)} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 12px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {L.back}</button>
        <div style={{flex:1,fontSize:15,fontWeight:700,color:t.text}}>+ {lang==="tr"?"Yeni Grup":"New Group"}</div>
      </div>
      <input style={{...iSt(t),marginBottom:14}} placeholder={lang==="tr"?"Grup adı...":"Group name..."} value={groupName} onChange={e=>setGroupName(e.target.value)}/>
      <div style={{fontSize:12,fontWeight:700,color:t.tm,marginBottom:8}}>{lang==="tr"?"Üyeleri Seç":"Select Members"}</div>
      {(teamMembers||[]).filter(m=>(m.userId||m.user_id)!==myUid).map((m,i)=>{
        const uid=m.userId||m.user_id;
        const sel=groupMembers.includes(uid);
        return <button key={i} onClick={()=>setGroupMembers(p=>sel?p.filter(x=>x!==uid):[...p,uid])} style={{width:"100%",display:"flex",alignItems:"center",gap:12,padding:"10px 14px",background:sel?t.acB:t.card,border:`1px solid ${sel?t.accent:t.border}`,borderRadius:12,cursor:"pointer",marginBottom:8,textAlign:"left"}}>
          <div style={{width:36,height:36,borderRadius:"50%",background:sel?t.accent:t.inBg,color:sel?"#fff":t.ts,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:700,flexShrink:0}}>{sel?"✓":(m.name||"?")[0].toUpperCase()}</div>
          <div style={{flex:1}}><div style={{fontSize:14,fontWeight:600,color:t.text}}>{m.name}</div>{m.role&&<div style={{fontSize:11,color:t.tm}}>{m.role}</div>}</div>
        </button>;
      })}
      <button onClick={()=>{
        if(!groupName.trim()||groupMembers.length===0){window.toast.info(lang==="tr"?"Grup adı ve en az 1 üye":"Group name + 1 member");return;}
        setActiveChat({customGroup:{id:`grp_${Date.now()}`,name:groupName.trim(),members:[myUid,...groupMembers]}});
        setGroupName("");setGroupMembers([]);
      }} disabled={!groupName.trim()||groupMembers.length===0} style={{...bSt("p",t),marginTop:12,opacity:groupName.trim()&&groupMembers.length>0?1:0.5}}>
        {lang==="tr"?"Grubu Oluştur":"Create Group"}
      </button>
    </div>;
  }

  // DM yükle
  useEffect(()=>{
    if(!activeChat?.dm||!team?.id||!myUid)return;
    const otherUid=activeChat.dm.userId||activeChat.dm.user_id;
    setDmLoading(true);setDmMessages([]);
    const sb=initSupabase();if(!sb)return;
    sb.from("team_messages").select("*")
      .eq("team_id",team.id)
      .or(`and(user_id.eq.${myUid},private_to.eq.${otherUid}),and(user_id.eq.${otherUid},private_to.eq.${myUid})`)
      .order("created_at",{ascending:true}).limit(100)
      .then(({data})=>{if(data)setDmMessages(data);setDmLoading(false);
        setTimeout(()=>dmEndRef.current?.scrollIntoView({behavior:"smooth"}),100);});
    const channel=sb.channel(`dm-${[myUid,otherUid].sort().join("-")}`)
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"team_messages",filter:`team_id=eq.${team.id}`},
        (payload)=>{
          const m=payload.new;
          if((m.user_id===myUid&&m.private_to===otherUid)||(m.user_id===otherUid&&m.private_to===myUid)){
            setDmMessages(p=>p.find(x=>x.id===m.id)?p:[...p,m]);
            setTimeout(()=>dmEndRef.current?.scrollIntoView({behavior:"smooth"}),50);
          }
        }).subscribe();
    return()=>{sb.removeChannel(channel);};
  },[activeChat?.dm?.userId,activeChat?.dm?.user_id,team?.id,myUid]);

  const sendDm=async()=>{
    if(!dmNew.trim()||!activeChat?.dm||!team?.id||!myUid)return;
    const otherUid=activeChat.dm.userId||activeChat.dm.user_id;
    const sb=initSupabase();if(!sb)return;
    const{data:{session}}=await sb.auth.getSession();
    const uid=session?.user?.id||myUid;
    const uname=session?.user?.user_metadata?.name||session?.user?.user_metadata?.full_name||session?.user?.email?.split("@")[0]||user?.name||"?";
    const profile=JSON.parse(localStorage.getItem("kmp_profile")||"{}");
    await sb.from("team_messages").insert({
      team_id:team.id,user_id:uid,user_name:uname,
      user_role:profile.role||"",private_to:otherUid,
      type:"text",text:dmNew.trim()
    });
    setDmNew("");
  };

  // DM ekranı
  if(activeChat?.dm){
    const other=activeChat.dm;
    return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${t.border}`}}>
        <button onClick={()=>{setActiveChat(null);setDmMessages([]);}} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 12px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {L.back}</button>
        <div style={{width:36,height:36,borderRadius:"50%",background:t.inBg,color:t.ts,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,flexShrink:0}}>
          {(other.name||"?")[0].toUpperCase()}
        </div>
        <div style={{flex:1}}>
          <div style={{fontSize:15,fontWeight:700,color:t.text}}>{other.name}</div>
          {other.role&&<div style={{fontSize:10,color:other.role==="chef"?t.accent:t.tm}}>{other.role==="chef"?"👑 ":""}{other.role}</div>}
        </div>
      </div>
      <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:8,paddingBottom:8}}>
        {dmLoading&&<div style={{textAlign:"center",padding:20,color:t.tm}}>⏳</div>}
        {!dmLoading&&dmMessages.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:t.tm,fontSize:13}}>
          <div style={{fontSize:32,marginBottom:8}}>💬</div>
          {lang==="tr"?"Henüz mesaj yok":"No messages yet"}
        </div>}
        {dmMessages.map((msg,i)=>{
          const me=msg.user_id===myUid;
          const prev=dmMessages[i-1];
          const showMeta=!prev||prev.user_id!==msg.user_id;
          return <div key={msg.id||i} style={{display:"flex",flexDirection:"column",alignItems:me?"flex-end":"flex-start"}}>
            {showMeta&&<div style={{fontSize:10,color:me?t.accent:t.tm,marginBottom:2,marginLeft:me?0:8,marginRight:me?8:0,fontWeight:600}}>
              {me?(lang==="tr"?"Sen":"You"):msg.user_name}
            </div>}
            <div style={{maxWidth:"80%",background:me?t.accent:t.card,color:me?"#fff":t.text,
              borderRadius:me?"18px 18px 4px 18px":"4px 18px 18px 18px",
              padding:"9px 13px",fontSize:14,lineHeight:1.4,
              border:me?"none":`1px solid ${t.border}`,wordBreak:"break-word"}}>
              {msg.text}
            </div>
            <div style={{fontSize:9,color:t.tm,marginTop:2,marginLeft:me?0:8,marginRight:me?8:0}}>{fmtTime(msg.created_at)}</div>
          </div>;
        })}
        <div ref={dmEndRef}/>
      </div>
      <div style={{display:"flex",gap:6,paddingTop:8,borderTop:`1px solid ${t.border}`}}>
        <input style={{...iSt(t),flex:1,borderRadius:20,padding:"10px 14px"}}
          placeholder={lang==="tr"?"Mesaj yaz...":"Write a message..."}
          value={dmNew} onChange={e=>setDmNew(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendDm();}}}/>
        <button onClick={sendDm} disabled={!dmNew.trim()} style={{...bSt("p",t),borderRadius:"50%",width:42,height:42,padding:0,flexShrink:0,opacity:dmNew.trim()?1:0.5,fontSize:18}}>↑</button>
      </div>
    </div>;
  }

  // AI ekranı
  if(activeChat==="ai"){
    return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${t.border}`}}>
        <button onClick={()=>setActiveChat(null)} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 12px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {L.back}</button>
        <span style={{fontSize:18}}>🤖</span>
        <div style={{flex:1}}><div style={{fontSize:15,fontWeight:700,color:t.text}}>{L.ai}</div></div>
      </div>
      <div style={{flex:1,overflow:"hidden"}}>
        <AssistantTab recipes={recipes} stock={[]} apiKey={apiKey} conversations={conversations} setConversations={setConversations} activeConvId={activeConvId} setActiveConvId={setActiveConvId} botMessages={botMessages} productions={productions} storageChecks={storageChecks} t={t}/>
      </div>
    </div>;
  }

  // Ekip grubu ekranı
  if(activeChat==="team"){
    return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 160px)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${t.border}`}}>
        <button onClick={()=>setActiveChat(null)} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 12px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {L.back}</button>
        <div style={{width:36,height:36,borderRadius:"50%",background:t.acB,color:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>👥</div>
        <div style={{flex:1}}>
          <div style={{fontSize:15,fontWeight:700,color:t.text}}>{team?.name||L.teamChat}</div>
          <div style={{fontSize:10,color:t.tm}}>{(teamMembers||[]).length} {lang==="tr"?"üye":"members"}</div>
        </div>
      </div>
      <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
        <ChatTab team={team} teamMembers={teamMembers} user={user} recipes={recipes} menus={menus} stock={stock} productions={productions} t={t}/>
      </div>
    </div>;
  }

  // Liste
  const otherMembers=(teamMembers||[]).filter(m=>(m.userId||m.user_id)!==myUid);
  return <div style={{maxWidth:520,margin:"0 auto"}}>
    <h2 style={{fontSize:22,color:t.text,fontFamily:"'Fraunces',serif",marginBottom:16}}>{L.title}</h2>
    <button onClick={()=>setActiveChat("ai")} style={{width:"100%",display:"flex",alignItems:"center",gap:14,padding:"14px 16px",background:t.card,border:`1px solid ${t.border}`,borderRadius:16,cursor:"pointer",marginBottom:10,textAlign:"left"}}>
      <div style={{width:48,height:48,borderRadius:"50%",background:`linear-gradient(135deg,${t.accent},${t.accent}99)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>🤖</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:15,fontWeight:700,color:t.text}}>{L.ai}</div>
        <div style={{fontSize:12,color:t.tm,marginTop:2}}>{L.aiDesc}</div>
      </div>
      <span style={{color:t.tm,fontSize:20,opacity:0.4}}>›</span>
    </button>
    {team?<button onClick={()=>setActiveChat("team")} style={{width:"100%",display:"flex",alignItems:"center",gap:14,padding:"14px 16px",background:t.card,border:`1px solid ${t.border}`,borderRadius:16,cursor:"pointer",marginBottom:10,textAlign:"left"}}>
      <div style={{width:48,height:48,borderRadius:"50%",background:t.acB,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>👥</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:15,fontWeight:700,color:t.text}}>{team.name}</div>
        <div style={{fontSize:12,color:t.tm,marginTop:2}}>{L.teamChat} · {(teamMembers||[]).length} {lang==="tr"?"üye":"members"}</div>
      </div>
      <span style={{color:t.tm,fontSize:20,opacity:0.4}}>›</span>
    </button>:<div style={{...cSt(t),padding:"14px 16px",marginBottom:10,display:"flex",alignItems:"center",gap:14,opacity:0.5}}>
      <div style={{width:48,height:48,borderRadius:"50%",background:t.inBg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>👥</div>
      <div style={{flex:1}}><div style={{fontSize:15,fontWeight:600,color:t.text}}>{L.teamChat}</div>
        <div style={{fontSize:12,color:t.accent,marginTop:2}}>⚙ → {lang==="tr"?"Ekip":"Team"} → {L.noTeam}</div>
      </div>
    </div>}
    {team&&otherMembers.length>0&&<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <div style={{fontSize:11,fontWeight:700,color:t.tm,letterSpacing:"0.08em"}}>{lang==="tr"?"ÖZEL MESAJLAR":"DIRECT MESSAGES"}</div>
        <button onClick={()=>setActiveChat({newGroup:true})} style={{fontSize:11,color:t.accent,background:"none",border:`1px solid ${t.accent}`,borderRadius:8,padding:"3px 8px",cursor:"pointer"}}>
          + {lang==="tr"?"Grup Oluştur":"New Group"}
        </button>
      </div>
      {otherMembers.map((m,i)=>{
        const uid=m.userId||m.user_id;
        const initial=(m.name||"?")[0].toUpperCase();
        return <button key={i} onClick={()=>setActiveChat({dm:m})} style={{width:"100%",display:"flex",alignItems:"center",gap:12,padding:"12px 14px",background:t.card,border:`1px solid ${t.border}`,borderRadius:14,cursor:"pointer",marginBottom:8,textAlign:"left"}}>
          <div style={{width:40,height:40,borderRadius:"50%",background:t.acB,color:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,flexShrink:0}}>
            {initial}
          </div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:14,fontWeight:600,color:t.text}}>{m.name||"?"}</div>
            {m.role&&<div style={{fontSize:11,color:m.role==="chef"?t.accent:t.tm}}>{m.role==="chef"?"👑 ":""}{m.role}</div>}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:t.success}}/>
            <span style={{color:t.tm,fontSize:20,opacity:0.4}}>›</span>
          </div>
        </button>;
      })}
    </>}
  </div>;
};

// ═══ WHATSAPP SOHBET SİSTEMİ ═══
const WAChatTab=({team,teamMembers,user,apiKey,t,tier})=>{
  const lang=t.lang;
  const myUid=user?.userId||user?.id||"";
  const myName=user?.name||user?.email||"";
  const sb=initSupabase();

  const[convList,setConvList]=useState([]);
  const[activeConv,setActiveConv]=useState(null);
  const[messages,setMessages]=useState([]);
  const[newMsg,setNewMsg]=useState("");
  const[sending,setSending]=useState(false);
  const[loading,setLoading]=useState(true);
  const[showNewGroup,setShowNewGroup]=useState(false);
  const[groupName,setGroupName]=useState("");
  const[groupMembers,setGroupMembers]=useState([]);
  const[uploading,setUploading]=useState(false);
  const[msgMenu,setMsgMenu]=useState(null); // {id, x, y}
  const[convMenu,setConvMenu]=useState(null); // {conv, x, y}
  const fileRef=useRef(null);
  const msgEndRef=useRef(null);

  const L={
    chats:{tr:"Sohbetler",en:"Chats"}[lang]||"Chats",
    team:{tr:"Ekip",en:"Team"}[lang]||"Team",
    newGroup:{tr:"Yeni Grup",en:"New Group"}[lang]||"New Group",
    dm:{tr:"Direkt Mesaj",en:"Direct Message"}[lang]||"DM",
    noTeam:{tr:"Önce bir ekibe katılın",en:"Join a team first"}[lang]||"Join a team first",
    groupNamePh:{tr:"Grup adı...",en:"Group name..."}[lang]||"Group name...",
    create:{tr:"Oluştur",en:"Create"}[lang]||"Create",
    members:{tr:"Üyeler",en:"Members"}[lang]||"Members",
    typeMsg:{tr:"Mesaj yaz...",en:"Type a message..."}[lang]||"Type a message...",
    back:{tr:"Geri",en:"Back"}[lang]||"Back",
    noMessages:{tr:"Henüz mesaj yok",en:"No messages yet"}[lang]||"No messages yet",
    deleteMsg:{tr:"Mesajı Sil",en:"Delete Message"}[lang]||"Delete Message",
    deleteConv:{tr:"Sohbeti Sil",en:"Delete Chat"}[lang]||"Delete Chat",
    leaveGroup:{tr:"Gruptan Ayrıl",en:"Leave Group"}[lang]||"Leave Group",
    confirmDelete:{tr:"Emin misiniz?",en:"Are you sure?"}[lang]||"Are you sure?",
  };

  const loadConversations=async()=>{
    if(!sb||!team?.id||!myUid)return;
    setLoading(true);
    try{
      await ensureTeamConversations();
      const{data:memberOf}=await sb.from("conversation_members").select("conversation_id").eq("user_id",myUid);
      const convIds=(memberOf||[]).map(m=>m.conversation_id);
      if(convIds.length>0){
        const{data:convs}=await sb.from("conversations").select("*").in("id",convIds).order("created_at",{ascending:false});
        const filtered=(convs||[]).filter(c=>{
          if(c.type!=="team")return true;
          if(tier==="pro")return c.tier==="pro";
          if(tier==="manager")return c.tier==="manager"||c.tier==="pro";
          if(tier==="chef")return c.tier==="chef"||c.tier==="manager";
          return true;
        });
        setConvList(filtered);
      }
    }catch(e){console.warn("Conv load:",e);}
    setLoading(false);
  };

  const ensureTeamConversations=async()=>{
    if(!sb||!team?.id||!myUid)return;
    let convId=null;
    const{data:existing}=await sb.from("conversations").select("id").eq("type","team").eq("team_id",team.id).maybeSingle();
    if(!existing){
      const{data:conv,error:ce}=await sb.from("conversations").insert({type:"team",name:team.name,team_id:team.id,tier:tier||"chef",created_by:myUid}).select().single();
      if(ce){console.warn("Conv create:",ce.message);return;}
      convId=conv.id;
      const uids=[...new Set([myUid,...(teamMembers||[]).map(m=>m.userId||m.user_id)])];
      for(const uid of uids){
        await sb.from("conversation_members").insert({conversation_id:convId,user_id:uid}).then(r=>{if(r.error&&!r.error.message.includes("duplicate"))console.warn(r.error.message);});
      }
    }else{
      convId=existing.id;
      const uids=[...new Set([myUid,...(teamMembers||[]).map(m=>m.userId||m.user_id)])];
      const{data:currentMembers}=await sb.from("conversation_members").select("user_id").eq("conversation_id",convId);
      const existing_uids=new Set((currentMembers||[]).map(m=>m.user_id));
      for(const uid of uids){
        if(!existing_uids.has(uid)){
          await sb.from("conversation_members").insert({conversation_id:convId,user_id:uid}).then(r=>{if(r.error&&!r.error.message.includes("duplicate"))console.warn(r.error.message);});
        }
      }
    }
    if(tier!=="pro"&&team.parent_team_id){
      const{data:parentConv}=await sb.from("conversations").select("id").eq("type","team").eq("team_id",team.parent_team_id).maybeSingle();
      if(parentConv){
        const{data:isMember}=await sb.from("conversation_members").select("id").eq("conversation_id",parentConv.id).eq("user_id",myUid).maybeSingle();
        if(!isMember){
          await sb.from("conversation_members").insert({conversation_id:parentConv.id,user_id:myUid}).then(r=>{if(r.error&&!r.error.message.includes("duplicate"))console.warn(r.error.message);});
        }
      }
    }
  };

  const loadMessages=async(convId)=>{
    if(!sb||!convId)return;
    const{data}=await sb.from("messages").select("*").eq("conversation_id",convId).order("created_at",{ascending:true}).limit(100);
    setMessages(data||[]);
    setTimeout(()=>msgEndRef.current?.scrollIntoView({behavior:"smooth"}),100);
  };

  useEffect(()=>{
    if(!activeConv?.id||!sb)return;
    const ch=sb.channel(`conv-${activeConv.id}`)
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"messages",filter:`conversation_id=eq.${activeConv.id}`},(payload)=>{
        setMessages(p=>[...p,payload.new]);
        setTimeout(()=>msgEndRef.current?.scrollIntoView({behavior:"smooth"}),50);
      })
      .on("postgres_changes",{event:"DELETE",schema:"public",table:"messages",filter:`conversation_id=eq.${activeConv.id}`},(payload)=>{
        setMessages(p=>p.filter(m=>m.id!==payload.old.id));
      })
      .subscribe();
    return()=>sb.removeChannel(ch);
  },[activeConv?.id]);

  useEffect(()=>{loadConversations();},[team?.id,myUid]);
  useEffect(()=>{if(activeConv?.id)loadMessages(activeConv.id);},[activeConv?.id]);

  const sendMessage=async(text,attachment=null)=>{
    if(!sb||!activeConv?.id||!myUid)return;
    if(!text?.trim()&&!attachment)return;
    setSending(true);
    try{
      await sb.from("messages").insert({conversation_id:activeConv.id,user_id:myUid,user_name:myName,text:text?.trim()||"",attachment});
      setNewMsg("");
    }catch(e){window.toast?.error(e.message);}
    setSending(false);
  };

  const sendFile=async(file)=>{
    if(!file||!team?.id)return;
    setUploading(true);
    try{
      const uploaded=await uploadFile(file,team.id,"chat");
      await sendMessage("",{url:uploaded.url,path:uploaded.path,name:uploaded.name,type:uploaded.type,ext:uploaded.ext});
    }catch(e){window.toast?.error(e.message);}
    setUploading(false);
  };

  // Mesaj sil
  const deleteMessage=async(msgId)=>{
    if(!sb||!msgId)return;
    await sb.from("messages").delete().eq("id",msgId).eq("user_id",myUid);
    setMessages(p=>p.filter(m=>m.id!==msgId));
    setMsgMenu(null);
  };

  // Sohbeti sil (DM/grup)
  const deleteConversation=async(conv)=>{
    if(!sb||!conv?.id)return;
    if(!window.confirm(L.confirmDelete))return;
    await sb.from("conversation_members").delete().eq("conversation_id",conv.id);
    await sb.from("messages").delete().eq("conversation_id",conv.id);
    await sb.from("conversations").delete().eq("id",conv.id);
    setConvList(p=>p.filter(c=>c.id!==conv.id));
    if(activeConv?.id===conv.id)setActiveConv(null);
    setConvMenu(null);
  };

  // Gruptan ayrıl
  const leaveGroup=async(conv)=>{
    if(!sb||!conv?.id||!myUid)return;
    if(!window.confirm(L.confirmDelete))return;
    await sb.from("conversation_members").delete().eq("conversation_id",conv.id).eq("user_id",myUid);
    setConvList(p=>p.filter(c=>c.id!==conv.id));
    if(activeConv?.id===conv.id)setActiveConv(null);
    setConvMenu(null);
  };

  const createGroup=async()=>{
    if(!groupName.trim()||!sb||!myUid)return;
    try{
      const grpTier=tier||"chef";
      const{data:conv,error:ce}=await sb.from("conversations").insert({type:"group",name:groupName.trim(),team_id:team?.id,tier:grpTier,created_by:myUid}).select().single();
      if(ce)throw ce;
      if(conv){
        const uids=[...new Set([myUid,...groupMembers])];
        for(const uid of uids){
          await sb.from("conversation_members").insert({conversation_id:conv.id,user_id:uid}).then(r=>{if(r.error&&!r.error.message.includes("duplicate"))console.warn(r.error.message);});
        }
        setConvList(p=>[conv,...p]);
        setActiveConv(conv);
        setShowNewGroup(false);
        setGroupName("");setGroupMembers([]);
        window.toast?.success(lang==="tr"?"Grup oluşturuldu":"Group created");
      }
    }catch(e){window.toast?.error(e.message);}
  };

  const openDM=async(otherUid,otherName)=>{
    if(!sb||!myUid)return;
    const{data:myConvs}=await sb.from("conversation_members").select("conversation_id").eq("user_id",myUid);
    const myIds=(myConvs||[]).map(m=>m.conversation_id);
    if(myIds.length>0){
      const{data:otherConvs}=await sb.from("conversation_members").select("conversation_id").eq("user_id",otherUid).in("conversation_id",myIds);
      if(otherConvs?.length>0){
        const{data:dmConv}=await sb.from("conversations").select("*").eq("type","dm").in("id",otherConvs.map(c=>c.conversation_id)).maybeSingle();
        if(dmConv){setActiveConv(dmConv);return;}
      }
    }
    const dmTier=tier||"chef";
    const{data:conv}=await sb.from("conversations").insert({type:"dm",name:otherName,team_id:team?.id,tier:dmTier,created_by:myUid}).select().single();
    if(conv){
      await sb.from("conversation_members").insert([{conversation_id:conv.id,user_id:myUid},{conversation_id:conv.id,user_id:otherUid}]);
      setConvList(p=>p.find(c=>c.id===conv.id)?p:[conv,...p]);
      setActiveConv(conv);
    }
  };

  const getMemberName=(uid)=>{
    const m=(teamMembers||[]).find(m=>(m.userId||m.user_id)===uid);
    return m?.name||uid?.slice(0,8)||"?";
  };

  const convIcon=(type)=>type==="team"?"🏢":type==="group"?"👥":"💬";

  if(!team?.id)return <div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>{L.noTeam}</div>;

  // Mesaj görünümü
  if(activeConv){
    return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 180px)",maxHeight:700}} onClick={()=>{setMsgMenu(null);}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 0",marginBottom:8,borderBottom:`1px solid ${t.border}`}}>
        <button onClick={()=>setActiveConv(null)} style={{background:"none",border:"none",color:t.accent,cursor:"pointer",fontSize:14,fontWeight:600}}>← {L.back}</button>
        <span style={{fontSize:18}}>{convIcon(activeConv.type)}</span>
        <div style={{flex:1}}>
          <div style={{fontSize:15,fontWeight:700,color:t.text}}>{activeConv.name}</div>
          <div style={{fontSize:10,color:t.tm}}>{activeConv.type==="team"?L.team:activeConv.type==="group"?L.newGroup:L.dm}</div>
        </div>
        {/* Grup/DM menü */}
        {activeConv.type!=="team"&&<button onClick={e=>{e.stopPropagation();setConvMenu(convMenu?null:{conv:activeConv,x:e.clientX,y:e.clientY});}} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:20,padding:"0 4px"}}>⋮</button>}
      </div>
      {/* Sohbet menü */}
      {convMenu&&<div style={{position:"fixed",top:convMenu.y,right:16,background:t.card,border:`1px solid ${t.border}`,borderRadius:10,boxShadow:"0 4px 20px rgba(0,0,0,0.15)",zIndex:500,minWidth:160}} onClick={e=>e.stopPropagation()}>
        {convMenu.conv.type==="group"&&<button onClick={()=>leaveGroup(convMenu.conv)} style={{display:"block",width:"100%",padding:"10px 16px",background:"none",border:"none",color:"#f97316",cursor:"pointer",textAlign:"left",fontSize:14}}>🚪 {L.leaveGroup}</button>}
        {(convMenu.conv.type==="dm"||(convMenu.conv.type==="group"&&convMenu.conv.created_by===myUid))&&<button onClick={()=>deleteConversation(convMenu.conv)} style={{display:"block",width:"100%",padding:"10px 16px",background:"none",border:"none",color:"#ef4444",cursor:"pointer",textAlign:"left",fontSize:14}}>🗑 {L.deleteConv}</button>}
      </div>}
      {/* Mesajlar */}
      <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:6,padding:"4px 0"}}>
        {messages.length===0&&<div style={{textAlign:"center",color:t.tm,padding:40,fontSize:13}}>{L.noMessages}</div>}
        {messages.map((msg,i)=>{
          const isMe=msg.user_id===myUid;
          const att=msg.attachment;
          return <div key={msg.id||i} style={{display:"flex",flexDirection:"column",alignItems:isMe?"flex-end":"flex-start",padding:"0 4px"}}>
            {!isMe&&<div style={{fontSize:10,color:t.accent,fontWeight:600,marginBottom:2,marginLeft:2}}>{msg.user_name||getMemberName(msg.user_id)}</div>}
            <div
              onContextMenu={isMe?e=>{e.preventDefault();setMsgMenu({id:msg.id,x:e.clientX,y:e.clientY});}:undefined}
              style={{maxWidth:"75%",background:isMe?t.accent:"#fff",color:isMe?"#fff":t.text,borderRadius:isMe?"16px 16px 4px 16px":"16px 16px 16px 4px",padding:"8px 12px",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",cursor:isMe?"context-menu":"default"}}>
              {att&&<div style={{marginBottom:msg.text?6:0}}>
                {isImage(att.ext)?<img src={att.url} style={{maxWidth:"100%",maxHeight:180,borderRadius:8,display:"block",cursor:"pointer"}} onClick={()=>window.open(att.url,"_blank")} alt={att.name}/>:
                isPDF(att.ext)?<a href={att.url} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:6,color:isMe?"#fff":t.accent,fontSize:12}}>📄 {att.name}</a>:
                <a href={att.url} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:6,color:isMe?"#fff":t.accent,fontSize:12}}>📎 {att.name}</a>}
              </div>}
              {msg.text&&<div style={{fontSize:14,lineHeight:1.5,wordBreak:"break-word"}}>{msg.text}</div>}
              <div style={{fontSize:9,opacity:0.6,marginTop:3,textAlign:"right"}}>{new Date(msg.created_at).toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"})}</div>
            </div>
          </div>;
        })}
        <div ref={msgEndRef}/>
      </div>
      {/* Mesaj sil menü */}
      {msgMenu&&<div style={{position:"fixed",top:msgMenu.y,left:msgMenu.x,background:t.card,border:`1px solid ${t.border}`,borderRadius:10,boxShadow:"0 4px 20px rgba(0,0,0,0.15)",zIndex:500}} onClick={e=>e.stopPropagation()}>
        <button onClick={()=>deleteMessage(msgMenu.id)} style={{display:"block",padding:"10px 16px",background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:14}}>🗑 {L.deleteMsg}</button>
      </div>}
      {/* Input */}
      <div style={{borderTop:`1px solid ${t.border}`,paddingTop:10,marginTop:6}}>
        <input ref={fileRef} type="file" accept="image/*,video/*,.pdf,.xlsx,.docx,.txt" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)sendFile(f);e.target.value="";}}/>
        <div style={{display:"flex",gap:6,alignItems:"flex-end"}}>
          <button onClick={()=>fileRef.current?.click()} disabled={uploading} style={{...bSt("s",t),padding:"10px 12px",flexShrink:0,fontSize:16}}>{uploading?"⏳":"📎"}</button>
          <input
            style={{...iSt(t),flex:1,padding:"10px 14px",borderRadius:20,fontSize:14}}
            placeholder={L.typeMsg}
            value={newMsg}
            onChange={e=>setNewMsg(e.target.value)}
            onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMessage(newMsg);}}}
          />
          <button onClick={()=>sendMessage(newMsg)} disabled={sending||!newMsg.trim()} style={{...bSt("p",t),padding:"10px 16px",flexShrink:0,borderRadius:20,opacity:newMsg.trim()?1:0.5}}>{sending?"⏳":"↑"}</button>
        </div>
      </div>
    </div>;
  }

  // Yeni grup modalı
  if(showNewGroup){
    return <div style={{padding:"0 4px"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
        <button onClick={()=>setShowNewGroup(false)} style={{background:"none",border:"none",color:t.accent,cursor:"pointer",fontSize:14}}>← {L.back}</button>
        <strong style={{fontSize:16,color:t.text}}>{L.newGroup}</strong>
      </div>
      <input style={{...iSt(t),marginBottom:12}} placeholder={L.groupNamePh} value={groupName} onChange={e=>setGroupName(e.target.value)}/>
      <div style={{fontSize:12,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em"}}>{L.members.toUpperCase()}</div>
      <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:16,maxHeight:300,overflowY:"auto"}}>
        {(teamMembers||[]).filter(m=>(m.userId||m.user_id)!==myUid).map(m=>{
          const uid=m.userId||m.user_id;
          const sel=groupMembers.includes(uid);
          return <button key={uid} onClick={()=>setGroupMembers(p=>sel?p.filter(id=>id!==uid):[...p,uid])} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:sel?t.accent+"22":t.inBg,border:`1px solid ${sel?t.accent:t.border}`,borderRadius:10,cursor:"pointer",textAlign:"left"}}>
            <div style={{width:32,height:32,borderRadius:"50%",background:sel?t.accent:t.tm,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13,flexShrink:0}}>{(m.name||"?")[0].toUpperCase()}</div>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:600,color:t.text}}>{m.name}</div>
              <div style={{fontSize:10,color:t.tm}}>{m.role}</div>
            </div>
            {sel&&<span style={{color:t.accent,fontSize:16}}>✓</span>}
          </button>;
        })}
      </div>
      <button onClick={createGroup} disabled={!groupName.trim()||groupMembers.length===0} style={{...bSt("p",t),width:"100%",opacity:groupName.trim()&&groupMembers.length>0?1:0.5}}>
        {L.create} {groupMembers.length>0&&`(${groupMembers.length+1})`}
      </button>
    </div>;
  }

  // Ana liste
  return <div onClick={()=>{setConvMenu(null);}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
      <strong style={{fontSize:18,color:t.text}}>{L.chats}</strong>
      <button onClick={()=>setShowNewGroup(true)} style={{...bSt("p",t),fontSize:12,padding:"6px 12px"}}>+ {L.newGroup}</button>
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:16}}>
      {loading&&<div style={{textAlign:"center",padding:30,color:t.tm}}>⏳</div>}
      {!loading&&convList.map(conv=><div key={conv.id} style={{display:"flex",alignItems:"center",gap:0,background:t.card,border:`1px solid ${t.border}`,borderRadius:12}}>
        <button onClick={()=>setActiveConv(conv)} style={{display:"flex",alignItems:"center",gap:10,padding:"12px 14px",background:"none",border:"none",cursor:"pointer",textAlign:"left",flex:1}}>
          <div style={{width:40,height:40,borderRadius:"50%",background:conv.type==="team"?t.accent:conv.type==="group"?"#8b5cf6":"#10b981",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{convIcon(conv.type)}</div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:14,fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{conv.name}</div>
            <div style={{fontSize:11,color:t.tm}}>{conv.type==="team"?L.team:conv.type==="group"?L.newGroup:L.dm}</div>
          </div>
        </button>
        {conv.type!=="team"&&<button onClick={e=>{e.stopPropagation();setConvMenu(convMenu?.conv?.id===conv.id?null:{conv,x:e.clientX,y:e.clientY});}} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:20,padding:"0 12px",alignSelf:"stretch",display:"flex",alignItems:"center"}}>⋮</button>}
      </div>)}
    </div>
    {/* Sohbet menü */}
    {convMenu&&<div style={{position:"fixed",top:convMenu.y,right:16,background:t.card,border:`1px solid ${t.border}`,borderRadius:10,boxShadow:"0 4px 20px rgba(0,0,0,0.15)",zIndex:500,minWidth:160}} onClick={e=>e.stopPropagation()}>
      {convMenu.conv.type==="group"&&<button onClick={()=>leaveGroup(convMenu.conv)} style={{display:"block",width:"100%",padding:"10px 16px",background:"none",border:"none",color:"#f97316",cursor:"pointer",textAlign:"left",fontSize:14}}>🚪 {L.leaveGroup}</button>}
      {(convMenu.conv.type==="dm"||(convMenu.conv.type==="group"&&convMenu.conv.created_by===myUid))&&<button onClick={()=>deleteConversation(convMenu.conv)} style={{display:"block",width:"100%",padding:"10px 16px",background:"none",border:"none",color:"#ef4444",cursor:"pointer",textAlign:"left",fontSize:14}}>🗑 {L.deleteConv}</button>}
    </div>}
    {/* DM */}
    {(teamMembers||[]).filter(m=>(m.userId||m.user_id)!==myUid).length>0&&<div>
      <div style={{fontSize:11,color:t.tm,fontWeight:700,letterSpacing:"0.05em",marginBottom:8}}>💬 {L.dm.toUpperCase()}</div>
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {(teamMembers||[]).filter(m=>(m.userId||m.user_id)!==myUid).map(m=>{
          const uid=m.userId||m.user_id;
          return <button key={uid} onClick={()=>openDM(uid,m.name||"?")} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:t.card,border:`1px solid ${t.border}`,borderRadius:10,cursor:"pointer",textAlign:"left",width:"100%"}}>
            <div style={{width:36,height:36,borderRadius:"50%",background:t.acB,color:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:14,flexShrink:0}}>{(m.name||"?")[0].toUpperCase()}</div>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:600,color:t.text}}>{m.name||"?"}</div>
              <div style={{fontSize:10,color:t.tm}}>{m.role}</div>
            </div>
            <span style={{color:t.tm,fontSize:14}}>›</span>
          </button>;
        })}
      </div>
    </div>}
  </div>;
};


// ═══ KM: EKİP KATILMA SİSTEMİ ═══
const generateInviteCode=()=>Math.random().toString(36).substring(2,8).toUpperCase();
const createTeam=async(teamName,userId,userName,parentTeamId=null)=>{
  const sb=initSupabase();if(!sb)throw new Error("Supabase yüklenemedi");
  const code=generateInviteCode();
  const insertData={name:teamName,invite_code:code.toUpperCase(),owner_id:userId,app_type:"pro"};
  if(parentTeamId)insertData.parent_team_id=parentTeamId;
  const{data:team,error:te}=await sb.from("teams").insert(insertData).select().single();
  if(te)throw te;
  const{error:me}=await sb.from("team_members").insert({team_id:team.id,user_id:userId,role:"worker",position:userName});
  if(me)throw me;
  return team;
};

const linkParentTeam=async(currentTeamId,inviteCode)=>{
  const sb=initSupabase();if(!sb)throw new Error("Supabase yüklenemedi");
  const{data:parentTeam,error:te}=await sb.from("teams").select("id,name").eq("invite_code",inviteCode.toUpperCase()).single();
  if(te||!parentTeam)throw new Error("Geçersiz davet kodu");
  if(parentTeam.id===currentTeamId)throw new Error("Kendi ekibinize bağlanamazsınız");
  const{error:ue}=await sb.from("teams").update({parent_team_id:parentTeam.id}).eq("id",currentTeamId);
  if(ue)throw ue;
  return parentTeam;
};

const getChildTeams=async(parentId)=>{
  const sb=initSupabase();if(!sb)return[];
  const{data}=await sb.from("teams").select("*").eq("parent_team_id",parentId);
  return data||[];
};

const joinTeam=async(inviteCode,userId,userName)=>{
  const sb=initSupabase();if(!sb)throw new Error("Supabase yüklenemedi");
  const{data:team,error:te}=await sb.from("teams").select("*").eq("invite_code",inviteCode.toUpperCase()).eq("app_type","pro").single();
  if(te||!team)throw new Error("Geçersiz davet kodu");
  const{data:existing}=await sb.from("team_members").select("id").eq("team_id",team.id).eq("user_id",userId).single();
  if(existing)return team;
  const{error:me}=await sb.from("team_members").insert({
    team_id:team.id,user_id:userId,role:"worker",position:userName
  });
  if(me)throw me;
  return team;
};

const syncFromTeam=async(teamId,table)=>{
  const sb=initSupabase();if(!sb)return null;
  const{data,error}=await sb.from(table).select("*").eq("team_id",teamId).order("updated_at",{ascending:false}).limit(1).single();
  if(error)return null;
  return data?.data||null;
};

const setTeamData=async(teamId,table,jsonData,userId)=>{
  const sb=initSupabase();if(!sb)return;
  const{data:existing}=await sb.from(table).select("id").eq("team_id",teamId).single();
  if(existing){
    await sb.from(table).update({data:jsonData,updated_at:new Date().toISOString(),updated_by:userId}).eq("id",existing.id);
  }else{
    await sb.from(table).insert({team_id:teamId,data:jsonData,updated_by:userId});
  }
};


const EventsTab=({team,teamMembers,user,apiKey,t})=>{
  const lang=t.lang;
  const[events,setEvents]=useState([]);
  const[loading,setLoading]=useState(true);
  const[showNew,setShowNew]=useState(false);
  const[parsing,setParsing]=useState(false);
  const[parseProgress,setParseProgress]=useState("");
  const[selectedEvent,setSelectedEvent]=useState(null);
  const[error,setError]=useState("");
  const[showManual,setShowManual]=useState(false);
  const[manualForm,setManualForm]=useState({name:"",event_date:"",start_time:"",pax:"",location:"",notes:"",items:"",photos:[]});
  const[manualBusy,setManualBusy]=useState(false);
  const[manualPreview,setManualPreview]=useState(null);

  // Departman tanımları
  const DEPARTMENTS=[
    {id:"kitchen",icon:"🍳",tr:"Sıcak Mutfak",en:"Hot Kitchen",color:"#dc2626"},
    {id:"cold",icon:"🥗",tr:"Soğuk Mutfak",en:"Cold Kitchen",color:"#0891b2"},
    {id:"pastry",icon:"🥐",tr:"Pastane",en:"Pastry",color:"#c8965a"},
    {id:"bakery",icon:"🍞",tr:"Fırın",en:"Bakery",color:"#92400e"},
    {id:"butcher",icon:"🥩",tr:"Kasap",en:"Butchery",color:"#7f1d1d"},
    {id:"service",icon:"🍽",tr:"Servis",en:"Banquet Service",color:"#7c3aed"},
    {id:"bar",icon:"🍷",tr:"Bar",en:"Bar",color:"#059669"},
    {id:"setup",icon:"🪑",tr:"Kurulum",en:"Setup",color:"#525252"},
    {id:"accounting",icon:"💰",tr:"Muhasebe",en:"Accounting",color:"#1e40af"},
    {id:"general",icon:"📋",tr:"Genel",en:"General",color:"#6b7280"}
  ];

  // Yükle
  useEffect(()=>{
    if(!team?.id){setLoading(false);return;}
    const sb=initSupabase();if(!sb){setLoading(false);return;}
    sb.from("events").select("*").eq("team_id",team.id).order("event_date",{ascending:false}).limit(50)
      .then(({data,error})=>{
        if(error){console.warn("[events] load error:",error.message);setEvents([]);setLoading(false);return;}
        setEvents(data||[]);setLoading(false);
      });
  },[team?.id]);

  // PDF metnini çıkar (PDF.js)
  const extractPDFText=async(file)=>{
    // PDF.js yükle
    if(!window.pdfjsLib){
      await new Promise((res,rej)=>{
        const s=document.createElement("script");
        s.src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
        s.onload=res;s.onerror=rej;document.head.appendChild(s);
      });
      window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    }
    const buf=await file.arrayBuffer();
    const pdf=await window.pdfjsLib.getDocument({data:buf}).promise;
    const pages=[];
    let totalLen=0;
    for(let i=1;i<=pdf.numPages;i++){
      const page=await pdf.getPage(i);
      const content=await page.getTextContent();
      const txt=content.items.map(it=>it.str).join(" ");
      pages.push(txt);
      totalLen+=txt.replace(/\s/g,"").length;
    }
    return{text:pages.join("\n\n--- SAYFA ---\n\n"),isImageBased:totalLen<200,pageCount:pdf.numPages};
  };

  // PDF'i base64'e çevir (vision için)
  const pdfToBase64=(file)=>new Promise((res,rej)=>{
    const r=new FileReader();
    r.onload=()=>res(r.result.split(",")[1]);
    r.onerror=()=>rej(new Error("Dosya okunamadı"));
    r.readAsDataURL(file);
  });

  // AI ile parse et
  const parseWithAI=async(file)=>{
    setParsing(true);setError("");
    try{
      setParseProgress(lang==="tr"?"PDF okunuyor...":"Reading PDF...");
      const{text,isImageBased,pageCount}=await extractPDFText(file);

      const sysPrompt=`You are a professional kitchen operations assistant analyzing a Banquet Event Order (BEO) document. Extract structured event data with multi-day support and sub-events (e.g. AM Coffee Break, Lunch, PM Coffee Break).

Output VALID JSON ONLY (no markdown, no explanation), matching this schema:
{
  "name": "Event/booking name (e.g. Live Consulting Meeting, Wedding Tasting)",
  "contractNo": "Contract/booking number if shown (e.g. 714268)",
  "accountName": "Account/company name (e.g. Asemble Organizasyon)",
  "confManager": "Conf/Cat Manager name if shown",
  "contactName": "Primary contact person name (e.g. Emel Duman)",
  "contactPhone": "Contact phone number if shown",
  "startDate": "YYYY-MM-DD — earliest date in BEO",
  "endDate": "YYYY-MM-DD — latest date in BEO (same as startDate if single-day)",
  "currency": "EUR/USD/TRY/GBP",
  "totalAmount": number or null,
  "pricePerPerson": number or null,
  "vatRate": number (default 20),
  "subEvents": [
    {
      "date": "YYYY-MM-DD",
      "timeStart": "HH:MM",
      "timeEnd": "HH:MM",
      "name": "Sub-event name (Meeting, AM Coffee Break, Lunch, PM Coffee Break, Tea & Coffee)",
      "room": "Room/venue (Kaftan, Tugra Lobby, Tugra Restaurant)",
      "setUp": "Set-up type (Lounge, Theater, Coffee Break, Existing Setup)",
      "pax": number,
      "items": [
        {
          "name": "Menu item exact name (e.g. 'Mekik çeşitleri', 'Grilled lamb loin')",
          "departments": ["pastry","bakery"]
        }
      ],
      "notesTr": "Turkish notes/instructions from ATT TO X sections (if any)",
      "notesEn": "English notes/instructions from ATT TO X sections (if any)"
    }
  ],
  "summary": "2-3 sentence summary in ${lang==="tr"?"Turkish":"English"}"
}

Department codes (each item can have MULTIPLE departments):
- "kitchen" = Hot kitchen: main courses, hot starters, hot canapes, hot soups
- "cold" = Cold kitchen: cold starters, salads, cold canapes
- "pastry" = Desserts, cakes, baklava, panna cotta, profiterole, tarts, tiramisu
- "bakery" = Bread, viennoiserie, simit, brioche, focaccia, poğaça, çörek, muffin
- "butcher" = Meat prep: lamb cuts, beef tenderloin, marinades
- "service" = Banquet service: table arrangements, plating, lounge setup
- "bar" = Beverages: cocktails, wine, soft drinks (note: Tea & Coffee is service, not bar)
- "setup" = Furniture/equipment: podium, AV, signage, skirt, chairs
- "accounting" = Pricing, billing notes, VAT
- "general" = Other notes

Rules:
- ALWAYS extract every sub-event separately (Meeting, AM Break, Lunch, PM Break, Tea Service).
- Each menu item is one entry with departments array (can be multi: "Tahini buns" = ["pastry","bakery"]).
- Notes (ATT TO ...) belong to the sub-event they're under. Split TR and EN if both languages present.
- Setup instructions (podium, chair, skirt) → "setup" department + put in notes.
- If a sub-event has no menu items (e.g. pure meeting), items can be empty array.
- For meat dishes, also add "butcher" if prep cuts are needed.
- Bakery vs Pastry: Bread/savoury baked = bakery; Sweet desserts = pastry. Hybrid items (tahini buns) = both.
- Convert dates: "27. February 2026" → "2026-02-27".
- Extract exact pax from "Exp/Gtd: 12 / 12" or "for 40 pax".
- Keep menu items in ORIGINAL language exactly as in BEO.`;

      let userMessages;
      if(isImageBased){
        setParseProgress(lang==="tr"?"Resim PDF tespit edildi, vision ile analiz ediliyor...":"Image-based PDF detected, analyzing with vision...");
        // Sayfa sayfa base64 ve vision
        const base64=await pdfToBase64(file);
        userMessages=[{
          role:"user",
          content:[
            {type:"document",source:{type:"base64",media_type:"application/pdf",data:base64}},
            {type:"text",text:"Extract event details and assign items to departments per the schema."}
          ]
        }];
      }else{
        setParseProgress(lang==="tr"?"AI analiz ediyor...":"AI analyzing..."+ ` (${pageCount} pages)`);
        userMessages=[{
          role:"user",
          content:`Analyze this BEO and output JSON:\n\n${text.slice(0,12000)}`
        }];
      }

      const model=isImageBased?"claude-sonnet-4-5":"claude-haiku-4-5";
      const resp=await fetch("https://kitchen-manager-ai.aligny0.workers.dev",{
        method:"POST",
        headers:{"Content-Type":"application/json","X-Auth-Token":WORKER_AUTH_TOKEN},
        body:JSON.stringify({model,max_tokens:3000,system:sysPrompt,messages:userMessages})
      });
      if(!resp.ok){
        const errText=await resp.text().catch(()=>"");
        let errMsg="HTTP "+resp.status;
        try{const j=JSON.parse(errText);if(j.error?.message)errMsg+=" — "+j.error.message;else if(j.error)errMsg+=" — "+JSON.stringify(j.error);}catch{errMsg+=" — "+errText.slice(0,200);}
        // 503/529 → bir kez retry
        if(resp.status===503||resp.status===529){
          setParseProgress(lang==="tr"?"AI yoğun, tekrar deniyor...":"AI busy, retrying...");
          await new Promise(r=>setTimeout(r,2500));
          const resp2=await fetch("https://kitchen-manager-ai.aligny0.workers.dev/",{
            method:"POST",
            headers:{"Content-Type":"application/json","X-Auth-Token":WORKER_AUTH_TOKEN},
            body:JSON.stringify({model,max_tokens:3000,system:sysPrompt,messages:userMessages})
          });
          if(!resp2.ok){
            const t2=await resp2.text().catch(()=>"");
            throw new Error("AI hatası ("+resp2.status+"): "+(t2.slice(0,200)||errMsg));
          }
          const data2=await resp2.json();
          const aiText2=data2?.content?.[0]?.text||"";
          let jsonStr2=aiText2.trim().replace(/^```(?:json|JSON)?\s*\n?/,"").replace(/\n?```\s*$/,"").trim();
          const fb=jsonStr2.search(/[\{\[]/);const lb=Math.max(jsonStr2.lastIndexOf("}"),jsonStr2.lastIndexOf("]"));
          if(fb>=0&&lb>fb)jsonStr2=jsonStr2.substring(fb,lb+1);
          const parsed2=JSON.parse(jsonStr2);
          setParseProgress(lang==="tr"?"Tamamlandı":"Done");
          return{parsed:parsed2,rawText:text.slice(0,5000),isImageBased};
        }
        throw new Error("AI hatası: "+errMsg);
      }
      const data=await resp.json();
      const aiText=data?.content?.[0]?.text||"";

      // JSON parse — markdown fence varsa temizle (daha sağlam)
      let jsonStr=aiText.trim();
      // 1) ```json ... ``` veya ``` ... ``` fence'ini her yerden kaldır
      jsonStr=jsonStr.replace(/^```(?:json|JSON)?\s*\n?/,"").replace(/\n?```\s*$/,"").trim();
      // 2) İlk { veya [ ile son } veya ] arasını al (içinde başka şey varsa kırp)
      const firstBrace=jsonStr.search(/[\{\[]/);
      const lastBrace=Math.max(jsonStr.lastIndexOf("}"),jsonStr.lastIndexOf("]"));
      if(firstBrace>=0&&lastBrace>firstBrace){jsonStr=jsonStr.substring(firstBrace,lastBrace+1);}
      let parsed;
      try{parsed=JSON.parse(jsonStr);}
      catch(e){
        // Son çare: control char'ları temizleyip dene
        try{
          const cleaned=jsonStr.replace(/[\u0000-\u001F]+/g,(m)=>m.replace(/\n/g," ").replace(/\t/g," ").replace(/[^\s]/g,""));
          parsed=JSON.parse(cleaned);
        }catch(e2){
          console.warn("AI JSON parse hatası:",e.message,"\nİlk 500 char:",jsonStr.slice(0,500));
          throw new Error("AI JSON parse edilemedi: "+e.message);
        }
      }

      setParseProgress(lang==="tr"?"Tamamlandı":"Done");
      return{parsed,rawText:text.slice(0,5000),isImageBased};
    }finally{
      setParsing(false);
    }
  };

  // Yeni event yükle
  const handleUpload=async(file)=>{
    if(!file||!team?.id)return;
    if(file.size>10*1024*1024){setError(lang==="tr"?"PDF 10MB'dan büyük olamaz":"PDF must be under 10MB");return;}
    setError("");
    try{
      const{parsed,rawText,isImageBased}=await parseWithAI(file);
      
      // PDF'i Supabase Storage'a yükle
      let pdfPath=null;
      try{
        const sb=initSupabase();
        if(sb){
          const ts=Date.now();
          const safeName=file.name.replace(/[^a-zA-Z0-9._-]/g,"_");
          pdfPath=`${team.id}/${ts}_${safeName}`;
          const{error:upErr}=await sb.storage.from("event-pdfs").upload(pdfPath,file,{contentType:"application/pdf",upsert:false});
          if(upErr){
            console.warn("PDF Storage upload hatası:",upErr.message);
            pdfPath=null;
          }
        }
      }catch(e){console.warn("PDF upload exception:",e.message);}

      // Departments aggregate: tüm subEvent'lerden tek bir map (eski uyumluluk için)
      const aggDepts={};
      (parsed.subEvents||[]).forEach(se=>{
        (se.items||[]).forEach(it=>{
          (it.departments||[]).forEach(d=>{
            if(!aggDepts[d])aggDepts[d]=[];
            if(!aggDepts[d].includes(it.name))aggDepts[d].push(it.name);
          });
        });
      });

      setSelectedEvent({
        id:null,
        team_id:team.id,
        name:parsed.name||file.name.replace(/\.pdf$/i,""),
        contract_no:parsed.contractNo||"",
        account_name:parsed.accountName||"",
        conf_manager:parsed.confManager||"",
        contact_name:parsed.contactName||"",
        contact_phone:parsed.contactPhone||"",
        event_date:parsed.startDate||parsed.date||"",
        end_date:parsed.endDate||parsed.startDate||parsed.date||"",
        start_time:(parsed.subEvents?.[0]?.timeStart)||parsed.startTime||"",
        end_time:(parsed.subEvents?.[parsed.subEvents.length-1]?.timeEnd)||parsed.endTime||"",
        pax:parsed.subEvents?.[0]?.pax||parsed.pax||0,
        location:parsed.subEvents?.[0]?.room||parsed.location||"",
        departments:Object.keys(aggDepts).length?aggDepts:(parsed.departments||{}),
        sub_events:parsed.subEvents||[],
        currency:parsed.currency||"EUR",
        total_amount:parsed.totalAmount||null,
        price_per_person:parsed.pricePerPerson||null,
        vat_rate:parsed.vatRate||20,
        original_pdf_path:pdfPath,
        original_pdf_size:file.size,
        ai_summary:parsed.summary||"",
        raw_text:rawText,
        pdf_name:file.name,
        status:"draft",
        notes:"",
        _isNew:true,
        _isImageBased:isImageBased
      });
      setShowNew(false);
    }catch(e){
      setError((lang==="tr"?"Hata: ":"Error: ")+e.message);
    }
  };

  // Kaydet
  const saveEvent=async()=>{
    if(!selectedEvent||!selectedEvent.name?.trim())return;
    const sb=initSupabase();if(!sb)return;
    const payload={
      team_id:team.id,
      name:selectedEvent.name.trim(),
      contract_no:selectedEvent.contract_no||null,
      account_name:selectedEvent.account_name||null,
      conf_manager:selectedEvent.conf_manager||null,
      contact_name:selectedEvent.contact_name||null,
      contact_phone:selectedEvent.contact_phone||null,
      event_date:selectedEvent.event_date||null,
      end_date:selectedEvent.end_date||null,
      start_time:selectedEvent.start_time||null,
      end_time:selectedEvent.end_time||null,
      pax:selectedEvent.pax||null,
      location:selectedEvent.location||null,
      departments:selectedEvent.departments||{},
      sub_events:selectedEvent.sub_events||[],
      currency:selectedEvent.currency||"EUR",
      total_amount:selectedEvent.total_amount||null,
      price_per_person:selectedEvent.price_per_person||null,
      vat_rate:selectedEvent.vat_rate||20,
      original_pdf_path:selectedEvent.original_pdf_path||null,
      original_pdf_size:selectedEvent.original_pdf_size||null,
      ai_summary:selectedEvent.ai_summary||null,
      raw_text:selectedEvent.raw_text||null,
      pdf_name:selectedEvent.pdf_name||null,
      status:selectedEvent.status||"draft",
      notes:selectedEvent.notes||null,
      created_by:user?.userId||null
    };
    let res;
    if(selectedEvent.id){
      res=await sb.from("events").update(payload).eq("id",selectedEvent.id).select().single();
    }else{
      res=await sb.from("events").insert(payload).select().single();
    }
    if(res.error){window.toast.error((lang==="tr"?"Kayıt başarısız: ":"Save failed: ")+res.error.message);return;}
    if(res.data){
      if(selectedEvent.id){
        setEvents(p=>p.map(e=>e.id===res.data.id?res.data:e));
      }else{
        setEvents(p=>[res.data,...p]);
      }
      setSelectedEvent(null);
    }
  };

  // Departmanlara dağıt — sohbete mesaj at

  // Manuel etkinlik AI ile departmanlara dağıt
  const aiDistributeManual=async()=>{
    if(!manualForm.name?.trim()){window.toast.error(lang==="tr"?"Etkinlik adı gerekli":"Event name required");return;}
    if(!manualForm.items?.trim()){window.toast.error(lang==="tr"?"En az 1 menü kalemi gerekli":"At least 1 menu item required");return;}
    setManualBusy(true);setError("");
    try{
      const items=manualForm.items.split(/[\n,;]+/).map(x=>x.trim()).filter(Boolean);
      if(!items.length){throw new Error(lang==="tr"?"Geçerli kalem yok":"No valid items");}
      
      // AI prompt
      const prompt=`You are a banquet kitchen dispatcher. Categorize each menu/event item into one of these departments:
- kitchen (Hot Kitchen): hot mains, hot starters, soups, hot canapes, grilled, fried items
- cold (Cold Kitchen): cold starters, salads, cold canapes, mezes, hummus
- pastry: desserts, baklava, cakes, ice cream, sweets
- bakery: bread, simit, brioche, viennoiserie
- butcher: meat preparation requirements, lamb, beef tenderloin
- service (Banquet Service): table setup, service notes, decorations
- bar: drinks, cocktails, wine, beverages, coffee, tea
- setup: equipment, av, signage, podium
- accounting: pricing, billing
- general: anything that doesn't fit

Items to categorize:
${items.map((it,i)=>`${i+1}. ${it}`).join("\n")}

Respond with ONLY a JSON object, no other text:
{
  "kitchen": ["item1","item2"],
  "cold": [],
  "pastry": ["item3"],
  ...all 10 departments, empty arrays if none...
}

Use the EXACT item text as input. Each item should appear in exactly one department.`;
      
      const proxyUrl="https://kitchen-manager-ai.aligny0.workers.dev";
      const res=await fetch(proxyUrl,{
        method:"POST",
        headers:{"Content-Type":"application/json","X-Auth-Token":WORKER_AUTH_TOKEN},
        body:JSON.stringify({model:"claude-haiku-4-5-20251001",max_tokens:1024,messages:[{role:"user",content:prompt}]})
      });
      if(!res.ok){const txt=await res.text();throw new Error(`API ${res.status}: ${txt.slice(0,100)}`);}
      const data=await res.json();
      const text=data.content?.[0]?.text||"";
      // JSON çıkar
      const m=text.match(/\{[\s\S]*\}/);
      if(!m)throw new Error(lang==="tr"?"AI yanıtı anlaşılamadı":"AI response unclear");
      const departments=JSON.parse(m[0]);
      // Boş departmanları temizle
      Object.keys(departments).forEach(k=>{if(!Array.isArray(departments[k])||!departments[k].length)delete departments[k];});
      
      // Önizlemeyi göster
      setManualPreview({...manualForm,departments,_isNew:true,team_id:team.id});
    }catch(e){
      console.warn("[manual ai]",e);
      // AI hata verirse → kullanıcıya tüm kalemleri "general" olarak ata, manuel düzeltsin
      const items=manualForm.items.split(/[\n,;]+/).map(x=>x.trim()).filter(Boolean);
      setManualPreview({...manualForm,departments:{general:items},_isNew:true,team_id:team.id});
      window.toast.info(lang==="tr"?"AI ulaşılamadı, manuel atayın":"AI unavailable, assign manually");
    }
    setManualBusy(false);
  };

  // Manuel formu kaydet (önizlemeden)
  const saveManualEvent=async()=>{
    if(!manualPreview)return;
    const sb=initSupabase();if(!sb)return;
    const payload={
      team_id:team.id,
      name:manualPreview.name,
      event_date:manualPreview.event_date||null,
      start_time:manualPreview.start_time||null,
      pax:manualPreview.pax?parseInt(manualPreview.pax,10):null,
      location:manualPreview.location||null,
      summary:manualPreview.notes||null,
      departments:manualPreview.departments||{},
      photos:manualPreview.photos||[],
      source:"manual",
      created_by:user?.userId||null
    };
    try{
      const{data,error}=await sb.from("events").insert(payload).select().single();
      if(error)throw error;
      setEvents(p=>[data,...p]);
      window.toast.success(lang==="tr"?"✓ Etkinlik kaydedildi":"✓ Event saved");
      setShowManual(false);
      setManualPreview(null);
      setManualForm({name:"",event_date:"",start_time:"",pax:"",location:"",notes:"",items:"",photos:[]});
    }catch(e){window.toast.error(e.message);}
  };

  // Foto yükleme (base64'e çevir, küçült)
  const handleManualPhoto=async(file)=>{
    if(!file)return;
    if(file.size>5*1024*1024){window.toast.error(lang==="tr"?"Foto 5MB'dan büyük olamaz":"Photo > 5MB");return;}
    const reader=new FileReader();
    reader.onload=ev=>{
      const img=new Image();
      img.onload=()=>{
        const MAX=800;
        const sc=Math.min(MAX/img.width,MAX/img.height,1);
        const cv=document.createElement("canvas");
        cv.width=img.width*sc;cv.height=img.height*sc;
        cv.getContext("2d").drawImage(img,0,0,cv.width,cv.height);
        const dataUrl=cv.toDataURL("image/jpeg",0.75);
        setManualForm(f=>({...f,photos:[...(f.photos||[]),dataUrl].slice(0,5)}));
      };
      img.src=ev.target.result;
    };
    reader.readAsDataURL(file);
  };

  const distributeToDepartments=async(ev)=>{
    if(!ev?.departments||!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    if(!window.confirm(lang==="tr"?"Departmanlara görevler ekip sohbetine gönderilecek. Onaylıyor musun?":"Tasks will be sent to team chat. Confirm?"))return;
    let count=0;
    for(const[deptId,items] of Object.entries(ev.departments)){
      if(!items||!items.length)continue;
      const dept=DEPARTMENTS.find(d=>d.id===deptId);
      const deptLabel=dept?(lang==="tr"?dept.tr:dept.en):deptId;
      const deptIcon=dept?.icon||"📋";
      const dateStr=ev.event_date?new Date(ev.event_date+"T12:00:00").toLocaleDateString(lang==="tr"?"tr-TR":"en-US"):"";
      const msg=`${deptIcon} **${deptLabel.toUpperCase()}** — ${ev.name}\n📅 ${dateStr}${ev.start_time?` ${ev.start_time}`:""} ${ev.pax?`· ${ev.pax} pax`:""}${ev.location?`\n📍 ${ev.location}`:""}\n\n${items.map((it,i)=>`${i+1}. ${it}`).join("\n")}`;
      await sb.from("team_messages").insert({
        team_id:team.id,
        user_id:user?.userId||"event",
        user_name:`📅 Event: ${ev.name}`,
        user_role:"event",
        type:"text",
        text:msg
      });
      // Fotolar varsa her departmana ayrı mesaj olarak gönder
      if(Array.isArray(ev.photos)&&ev.photos.length){
        for(const photo of ev.photos.slice(0,5)){
          await sb.from("team_messages").insert({
            team_id:team.id,
            user_id:user?.userId||"event",
            user_name:`📅 Event: ${ev.name}`,
            user_role:"event",
            type:"image",
            text:`📷 ${ev.name} — ${deptLabel}`,
            attachment:photo
          }).then(()=>{}).catch(()=>{});
        }
      }
      count++;
    }
    window.toast.success(lang==="tr"?`✓ ${count} departmana dağıtıldı`:`✓ Distributed to ${count} departments`);
  };

  const deleteEvent=async(id)=>{
    if(!window.confirm(lang==="tr"?"Etkinlik silinsin mi?":"Delete event?"))return;
    const sb=initSupabase();if(!sb)return;
    await sb.from("events").delete().eq("id",id);
    setEvents(p=>p.filter(e=>e.id!==id));
    if(selectedEvent?.id===id)setSelectedEvent(null);
  };

  // Form: department editor
  const updateDept=(deptId,items)=>{
    setSelectedEvent(s=>({...s,departments:{...s.departments,[deptId]:items}}));
  };

  if(!team){
    return <div style={{padding:20,textAlign:"center"}}>
      <div style={{fontSize:48,marginBottom:12,opacity:0.3}}>🎉</div>
      <div style={{fontSize:14,color:t.tm}}>{lang==="tr"?"Ekip kurun veya katılın — Etkinlikler özelliği için ekip gerekli.":"Create or join a team — Events feature requires a team."}</div>
    </div>;
  }

  // Detail/edit form
  if(selectedEvent){
    return <div style={{padding:"12px 14px",paddingBottom:60}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <button onClick={()=>setSelectedEvent(null)} style={{...bSt("g",t),padding:"6px 12px",fontSize:12}}>← {lang==="tr"?"Geri":"Back"}</button>
        <div style={{fontSize:12,color:t.tm}}>{selectedEvent._isNew?(lang==="tr"?"Yeni Etkinlik":"New Event"):(lang==="tr"?"Etkinlik Düzenle":"Edit Event")}</div>
      </div>

      {selectedEvent.ai_summary&&<div style={{...cSt(t),padding:"10px 12px",marginBottom:12,background:t.accent+"15",border:`1px solid ${t.accent}40`}}>
        <div style={{fontSize:9,fontWeight:700,color:t.accent,letterSpacing:"0.1em",marginBottom:4}}>🤖 AI ÖZET</div>
        <div style={{fontSize:13,color:t.text,lineHeight:1.5}}>{selectedEvent.ai_summary}</div>
        {selectedEvent._isImageBased&&<div style={{fontSize:9,color:t.tm,marginTop:6}}>📷 {lang==="tr"?"Resim PDF — Vision ile analiz edildi":"Image PDF — Analyzed with Vision"}</div>}
      </div>}

      <div style={{...cSt(t),padding:"12px 14px",marginBottom:12}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr",gap:10}}>
          <div><label style={lSt(t)}>{lang==="tr"?"Etkinlik Adı":"Event Name"} *</label>
            <input style={iSt(t)} value={selectedEvent.name||""} onChange={e=>setSelectedEvent(s=>({...s,name:e.target.value}))}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div><label style={lSt(t)}>{lang==="tr"?"Tarih":"Date"}</label>
              <input type="date" style={iSt(t)} value={selectedEvent.event_date||""} onChange={e=>setSelectedEvent(s=>({...s,event_date:e.target.value}))}/>
            </div>
            <div><label style={lSt(t)}>{lang==="tr"?"Misafir":"Pax"}</label>
              <input type="number" style={iSt(t)} value={selectedEvent.pax||""} onChange={e=>setSelectedEvent(s=>({...s,pax:parseInt(e.target.value,10)||0}))}/>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div><label style={lSt(t)}>{lang==="tr"?"Başlangıç":"Start"}</label>
              <input type="time" style={iSt(t)} value={selectedEvent.start_time||""} onChange={e=>setSelectedEvent(s=>({...s,start_time:e.target.value}))}/>
            </div>
            <div><label style={lSt(t)}>{lang==="tr"?"Bitiş":"End"}</label>
              <input type="time" style={iSt(t)} value={selectedEvent.end_time||""} onChange={e=>setSelectedEvent(s=>({...s,end_time:e.target.value}))}/>
            </div>
          </div>
          <div><label style={lSt(t)}>{lang==="tr"?"Konum":"Location"}</label>
            <input style={iSt(t)} value={selectedEvent.location||""} placeholder={lang==="tr"?"Salon adı...":"Venue name..."} onChange={e=>setSelectedEvent(s=>({...s,location:e.target.value}))}/>
          </div>
          <div><label style={lSt(t)}>{lang==="tr"?"Kontrat No":"Contract No"}</label>
            <input style={iSt(t)} value={selectedEvent.contract_no||""} onChange={e=>setSelectedEvent(s=>({...s,contract_no:e.target.value}))}/>
          </div>
        </div>
      </div>

      <div style={{fontSize:11,fontWeight:700,color:t.tm,letterSpacing:"0.05em",marginBottom:8,marginTop:16}}>
        🏢 {lang==="tr"?"DEPARTMAN GÖREVLERİ":"DEPARTMENT TASKS"}
      </div>
      {DEPARTMENTS.map(d=>{
        const items=(selectedEvent.departments?.[d.id])||[];
        return <div key={d.id} style={{...cSt(t),padding:"10px 12px",marginBottom:8,borderLeft:`3px solid ${d.color}`}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:items.length?8:0}}>
            <div style={{fontSize:12,fontWeight:700,color:t.text}}>{d.icon} {lang==="tr"?d.tr:d.en}</div>
            <button onClick={()=>updateDept(d.id,[...items,""])} style={{...bSt("g",t),padding:"3px 8px",fontSize:11}}>+ {lang==="tr"?"Ekle":"Add"}</button>
          </div>
          {items.map((item,i)=><div key={i} style={{display:"flex",gap:6,marginBottom:4}}>
            <input style={{...iSt(t),flex:1,fontSize:12,padding:"6px 8px"}} value={item} onChange={e=>{
              const newItems=[...items];newItems[i]=e.target.value;updateDept(d.id,newItems);
            }}/>
            <button onClick={()=>updateDept(d.id,items.filter((_,x)=>x!==i))} style={{...bSt("d",t),padding:"4px 8px",fontSize:11}}>✕</button>
          </div>)}
        </div>;
      })}

      <div><label style={lSt(t)}>{lang==="tr"?"Notlar":"Notes"}</label>
        <textarea style={{...iSt(t),minHeight:60,resize:"vertical"}} value={selectedEvent.notes||""} onChange={e=>setSelectedEvent(s=>({...s,notes:e.target.value}))}/>
      </div>

      <div style={{display:"flex",gap:8,marginTop:16,position:"sticky",bottom:60,background:t.bg+"e0",backdropFilter:"blur(10px)",padding:"8px 0"}}>
        <button onClick={()=>setSelectedEvent(null)} style={{...bSt("g",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
        <button onClick={saveEvent} disabled={!selectedEvent.name?.trim()} style={{...bSt("p",t),flex:2,opacity:selectedEvent.name?.trim()?1:0.5}}>
          ✓ {lang==="tr"?"Kaydet":"Save"}
        </button>
      </div>

      {!selectedEvent._isNew&&<button onClick={()=>distributeToDepartments(selectedEvent)} style={{...bSt("s",t),width:"100%",marginTop:10,fontSize:13,fontWeight:700}}>
        📤 {lang==="tr"?"Departmanlara Dağıt":"Distribute to Departments"}
      </button>}
    </div>;
  }

  // List view
  return <div style={{padding:"12px 14px",paddingBottom:60}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
      <div>
        <div style={{fontSize:18,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif"}}>🎉 {lang==="tr"?"Etkinlikler":"Events"}</div>
        <div style={{fontSize:11,color:t.tm,marginTop:2}}>{lang==="tr"?"BEO yükle, AI departmanlara dağıtsın":"Upload BEO, AI distributes to departments"}</div>
      </div>
      <div style={{display:"flex",gap:6}}>
        <button onClick={()=>{setShowManual(true);setManualPreview(null);}} style={{...bSt("s",t),padding:"8px 12px",fontSize:13,display:"flex",alignItems:"center",gap:4}}>
          ✍️ {lang==="tr"?"Manuel":"Manual"}
        </button>
        <label style={{...bSt("p",t),padding:"8px 14px",fontSize:13,cursor:parsing?"wait":"pointer",opacity:parsing?0.6:1,display:"flex",alignItems:"center",gap:6}}>
          {parsing?"⏳":"📄+"} {parsing?(parseProgress||"..."):(lang==="tr"?"PDF":"PDF")}
          <input type="file" accept=".pdf" disabled={parsing} style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleUpload(f);e.target.value="";}}/>
        </label>
      </div>
    </div>

    {error&&<div style={{...cSt(t),padding:"10px 12px",marginBottom:12,background:"#fee",border:"1px solid #fbb",color:"#900"}}>
      ⚠️ {error}
    </div>}

    {parsing&&<div style={{...cSt(t),padding:"14px",marginBottom:12,background:t.accent+"15",border:`1px dashed ${t.accent}`}}>
      <div style={{fontSize:13,fontWeight:600,color:t.accent,marginBottom:4}}>🤖 {lang==="tr"?"AI Analiz Ediyor":"AI Analyzing"}</div>
      <div style={{fontSize:11,color:t.tm}}>{parseProgress}</div>
    </div>}

    {loading?<div style={{padding:30,textAlign:"center",color:t.tm}}>{lang==="tr"?"Yükleniyor...":"Loading..."}</div>:
     events.length===0?<div style={{padding:40,textAlign:"center"}}>
      <div style={{fontSize:48,marginBottom:12,opacity:0.3}}>📭</div>
      <div style={{fontSize:13,color:t.tm}}>{lang==="tr"?"Henüz etkinlik yok. PDF yükle veya manuel ekle.":"No events yet. Upload a PDF or add manually."}</div>
    </div>:events.map(ev=>{
      const dateStr=ev.event_date?new Date(ev.event_date+"T12:00:00").toLocaleDateString(lang==="tr"?"tr-TR":"en-US",{day:"numeric",month:"short",year:"numeric"}):"";
      const deptCount=Object.keys(ev.departments||{}).filter(k=>ev.departments[k]?.length).length;
      const totalItems=Object.values(ev.departments||{}).reduce((sum,arr)=>sum+(arr?.length||0),0);
      const isPast=ev.event_date&&new Date(ev.event_date)<new Date(new Date().toDateString());
      return <div key={ev.id} style={{...cSt(t),padding:"12px 14px",marginBottom:8,opacity:isPast?0.6:1,cursor:"pointer"}}
        onClick={()=>setSelectedEvent(ev)}>
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:8}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:14,fontWeight:700,color:t.text,marginBottom:3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{ev.name}</div>
            <div style={{fontSize:10,color:t.tm,display:"flex",gap:8,flexWrap:"wrap"}}>
              {dateStr&&<span>📅 {dateStr}</span>}
              {ev.start_time&&<span>🕐 {ev.start_time}</span>}
              {ev.pax&&<span>👥 {ev.pax} pax</span>}
              {ev.location&&<span style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:140}}>📍 {ev.location}</span>}
            </div>
            <div style={{fontSize:10,color:t.accent,marginTop:6,fontWeight:600}}>
              🏢 {deptCount} {lang==="tr"?"departman":"departments"} · {totalItems} {lang==="tr"?"görev":"tasks"}
            </div>
          </div>
          <button onClick={e=>{e.stopPropagation();deleteEvent(ev.id);}} style={{...bSt("d",t),padding:"4px 8px",fontSize:11}}>✕</button>
        </div>
      </div>;
    })}
    {/* Manuel Etkinlik Modal */}
    {showManual&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:999,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)setShowManual(false);}}>
      <div style={{background:t.bg,borderRadius:"20px 20px 0 0",width:"100%",maxWidth:600,maxHeight:"92vh",overflowY:"auto",padding:"16px 14px calc(20px + env(safe-area-inset-bottom))"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:16,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif"}}>{manualPreview?(lang==="tr"?"Önizleme — Düzenle":"Preview — Edit"):(lang==="tr"?"Manuel Etkinlik":"Manual Event")}</div>
          <button onClick={()=>{setShowManual(false);setManualPreview(null);}} style={{background:"none",border:"none",fontSize:22,color:t.tm,cursor:"pointer",padding:"0 6px"}}>✕</button>
        </div>

        {!manualPreview?<>
          {/* Form */}
          <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:12}}>
            <input style={iSt(t)} placeholder={lang==="tr"?"Etkinlik adı *":"Event name *"} value={manualForm.name} onChange={e=>setManualForm(f=>({...f,name:e.target.value}))}/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
              <input style={iSt(t)} type="date" value={manualForm.event_date} onChange={e=>setManualForm(f=>({...f,event_date:e.target.value}))}/>
              <input style={iSt(t)} type="time" value={manualForm.start_time} onChange={e=>setManualForm(f=>({...f,start_time:e.target.value}))}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 2fr",gap:6}}>
              <input style={iSt(t)} type="number" placeholder={lang==="tr"?"Kişi (pax)":"Pax"} value={manualForm.pax} onChange={e=>setManualForm(f=>({...f,pax:e.target.value}))}/>
              <input style={iSt(t)} placeholder={lang==="tr"?"Lokasyon":"Location"} value={manualForm.location} onChange={e=>setManualForm(f=>({...f,location:e.target.value}))}/>
            </div>
            <textarea style={{...iSt(t),minHeight:60,resize:"vertical"}} placeholder={lang==="tr"?"Notlar (opsiyonel)":"Notes (optional)"} value={manualForm.notes} onChange={e=>setManualForm(f=>({...f,notes:e.target.value}))}/>
            <div>
              <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:4,letterSpacing:"0.05em"}}>🍽 {lang==="tr"?"MENÜ KALEMLERI (her satıra bir kalem)":"MENU ITEMS (one per line)"}</div>
              <textarea style={{...iSt(t),minHeight:120,resize:"vertical",fontFamily:"monospace",fontSize:13}} placeholder={lang==="tr"?"Domates çorbası\nLevrek ızgara\nCrème brûlée\nKokteyl seçimi":"Tomato soup\nGrilled sea bass\nCrème brûlée\nCocktail selection"} value={manualForm.items} onChange={e=>setManualForm(f=>({...f,items:e.target.value}))}/>
            </div>
            <div>
              <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:4,letterSpacing:"0.05em"}}>📷 {lang==="tr"?"FOTOĞRAFLAR (opsiyonel, max 5)":"PHOTOS (optional, max 5)"}</div>
              {manualForm.photos.length>0&&<div style={{display:"flex",gap:6,marginBottom:6,flexWrap:"wrap"}}>
                {manualForm.photos.map((p,i)=><div key={i} style={{position:"relative",width:60,height:60}}>
                  <img src={p} style={{width:"100%",height:"100%",objectFit:"cover",borderRadius:8,border:`1px solid ${t.border}`}}/>
                  <button onClick={()=>setManualForm(f=>({...f,photos:f.photos.filter((_,j)=>j!==i)}))} style={{position:"absolute",top:-6,right:-6,width:20,height:20,borderRadius:10,background:t.danger,color:"#fff",border:"none",fontSize:11,cursor:"pointer"}}>✕</button>
                </div>)}
              </div>}
              {manualForm.photos.length<5&&<label style={{...bSt("g",t),fontSize:12,padding:"8px 12px",cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}>
                + {lang==="tr"?"Foto Ekle":"Add Photo"}
                <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleManualPhoto(f);e.target.value="";}}/>
              </label>}
            </div>
          </div>

          <button onClick={aiDistributeManual} disabled={manualBusy} style={{...bSt("p",t),width:"100%",padding:"12px",fontSize:14,fontWeight:700,opacity:manualBusy?0.6:1}}>
            {manualBusy?"🤖 "+(lang==="tr"?"AI Dağıtıyor...":"AI Distributing..."):"🤖 "+(lang==="tr"?"AI ile Departmanlara Ata":"AI Distribute to Departments")}
          </button>
          <div style={{fontSize:11,color:t.tm,textAlign:"center",marginTop:6,lineHeight:1.4}}>{lang==="tr"?"AI menü kalemlerini analiz edip uygun departmanlara atayacak. Sonra düzenleyebilirsin.":"AI will analyze items and assign departments. You can edit after."}</div>
        </>:<>
          {/* Önizleme + düzenle */}
          <div style={{...cSt(t),padding:"10px 12px",marginBottom:12,background:t.acB,borderColor:t.accent}}>
            <div style={{fontSize:14,fontWeight:700,color:t.text}}>{manualPreview.name}</div>
            <div style={{fontSize:11,color:t.tm,marginTop:2}}>
              {manualPreview.event_date&&<>📅 {manualPreview.event_date} </>}
              {manualPreview.start_time&&<>🕐 {manualPreview.start_time} </>}
              {manualPreview.pax&&<>👥 {manualPreview.pax} pax </>}
              {manualPreview.location&&<>📍 {manualPreview.location}</>}
            </div>
          </div>

          <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em"}}>🏢 {lang==="tr"?"DEPARTMAN ATAMASI (yanlışsa düzenle)":"DEPARTMENT ASSIGNMENT (edit if wrong)"}</div>
          <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>
            {Object.entries(manualPreview.departments).map(([deptId,items])=>{
              const dept=DEPARTMENTS.find(d=>d.id===deptId);
              return (items||[]).map((item,idx)=><div key={`${deptId}-${idx}`} style={{display:"flex",gap:6,alignItems:"center",background:t.inBg,padding:"6px 10px",borderRadius:8}}>
                <span style={{flex:1,fontSize:12,color:t.text}}>{item}</span>
                <select value={deptId} onChange={e=>{
                  const newDept=e.target.value;
                  if(newDept===deptId)return;
                  setManualPreview(p=>{
                    const newDepts={...p.departments};
                    newDepts[deptId]=(newDepts[deptId]||[]).filter((_,i)=>i!==idx);
                    if(!newDepts[deptId].length)delete newDepts[deptId];
                    newDepts[newDept]=[...(newDepts[newDept]||[]),item];
                    return{...p,departments:newDepts};
                  });
                }} style={{...iSt(t),fontSize:11,padding:"4px 8px",width:130}}>
                  {DEPARTMENTS.map(d=><option key={d.id} value={d.id}>{d.icon} {lang==="tr"?d.tr:d.en}</option>)}
                </select>
                <button onClick={()=>setManualPreview(p=>{
                  const newDepts={...p.departments};
                  newDepts[deptId]=(newDepts[deptId]||[]).filter((_,i)=>i!==idx);
                  if(!newDepts[deptId].length)delete newDepts[deptId];
                  return{...p,departments:newDepts};
                })} style={{background:"none",border:"none",color:t.danger,cursor:"pointer",fontSize:13,padding:"0 4px"}}>✕</button>
              </div>);
            })}
          </div>

          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>setManualPreview(null)} style={{...bSt("g",t),flex:1,fontSize:13}}>← {lang==="tr"?"Geri":"Back"}</button>
            <button onClick={saveManualEvent} style={{...bSt("p",t),flex:2,fontSize:13,fontWeight:700}}>✓ {lang==="tr"?"Kaydet":"Save"}</button>
          </div>
        </>}
      </div>
    </div>}

  </div>;
};


// ═══ SHIFT TAB ═══
// ═══ TATİL HELPERS (date-holidays paketi) ═══
const _hdCache={};
const getHolidaysForCountry=(country="TR",year)=>{
  const key=`${country}-${year}`;
  if(_hdCache[key])return _hdCache[key];
  try{
    const hd=new Holidays(country);
    const list=hd.getHolidays(year)||[];
    const result={};
    list.forEach(h=>{
      if(h.type==="public"){
        // date-holidays "YYYY-MM-DD HH:MM:SS" formatında veriyor, ilk 10 karakteri al (saat dilimi dönüşümü olmasın)
        const dStr=String(h.date).slice(0,10);
        result[dStr]=h.name||"Holiday";
        // Ramazan Bayramı sadece 1. günü veriyor, +2 gün ekleyelim (toplam 3 gün)
        const nm=(h.name||"").toLowerCase();
        if(nm.includes("ramazan")||nm.includes("eid al-fitr")||nm.includes("ramadan")){
          for(let i=1;i<=2;i++){
            const d=new Date(dStr+"T12:00:00");
            d.setDate(d.getDate()+i);
            const next=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
            result[next]=(h.name||"Eid")+` (${i+1}. gün)`;
          }
        }
        // Kurban Bayramı sadece 1. günü veriyor, +3 gün ekleyelim (toplam 4 gün)
        if(nm.includes("kurban")||nm.includes("eid al-adha")||nm.includes("eid al adha")){
          for(let i=1;i<=3;i++){
            const d=new Date(dStr+"T12:00:00");
            d.setDate(d.getDate()+i);
            const next=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
            result[next]=(h.name||"Eid")+` (${i+1}. gün)`;
          }
        }
      }
    });
    _hdCache[key]=result;
    return result;
  }catch(e){
    console.warn("Tatil verisi alınamadı:",country,e.message);
    return {};
  }
};
// Geri uyumluluk için (eski kod hala kullanıyor olabilir)
const getTurkishHolidays=(year)=>getHolidaysForCountry("TR",year);

// Ülkeye göre hafta sonu günleri (0=Paz, 6=Cmt)
const getWeekendDays=(country="TR")=>{
  const fridayCountries=["SA","AE","BH","KW","OM","QA","YE","DZ","EG","IQ","JO","LY","SY","SD"];
  const israelLike=["IL"];
  if(israelLike.includes(country))return[5,6]; // Cuma, Cumartesi
  if(fridayCountries.includes(country))return[5,6]; // Cuma, Cumartesi
  return[0,6]; // Pazar, Cumartesi (varsayılan)
};

// Ülkeye göre default yıllık izin
const getDefaultAnnualLeave=(country="TR")=>{
  const map={TR:14,DE:20,FR:25,GB:28,US:10,IT:20,ES:22,NL:20,BE:20,AT:25,CH:20,PL:20,AU:20,JP:10,CA:10,SE:25,NO:25,DK:25,FI:25};
  return map[country]||14;
};

// Yaygın ülke listesi (dropdown için)
const COMMON_COUNTRIES=[
  {code:"TR",name:"Türkiye",flag:"🇹🇷"},
  {code:"DE",name:"Deutschland",flag:"🇩🇪"},
  {code:"FR",name:"France",flag:"🇫🇷"},
  {code:"GB",name:"United Kingdom",flag:"🇬🇧"},
  {code:"US",name:"United States",flag:"🇺🇸"},
  {code:"IT",name:"Italia",flag:"🇮🇹"},
  {code:"ES",name:"España",flag:"🇪🇸"},
  {code:"NL",name:"Nederland",flag:"🇳🇱"},
  {code:"BE",name:"België",flag:"🇧🇪"},
  {code:"AT",name:"Österreich",flag:"🇦🇹"},
  {code:"CH",name:"Schweiz",flag:"🇨🇭"},
  {code:"PL",name:"Polska",flag:"🇵🇱"},
  {code:"SE",name:"Sverige",flag:"🇸🇪"},
  {code:"NO",name:"Norge",flag:"🇳🇴"},
  {code:"DK",name:"Danmark",flag:"🇩🇰"},
  {code:"FI",name:"Suomi",flag:"🇫🇮"},
  {code:"GR",name:"Ελλάδα",flag:"🇬🇷"},
  {code:"PT",name:"Portugal",flag:"🇵🇹"},
  {code:"IE",name:"Ireland",flag:"🇮🇪"},
  {code:"CZ",name:"Česko",flag:"🇨🇿"},
  {code:"HU",name:"Magyarország",flag:"🇭🇺"},
  {code:"RO",name:"România",flag:"🇷🇴"},
  {code:"RU",name:"Россия",flag:"🇷🇺"},
  {code:"UA",name:"Україна",flag:"🇺🇦"},
  {code:"US",name:"United States",flag:"🇺🇸"},
  {code:"CA",name:"Canada",flag:"🇨🇦"},
  {code:"MX",name:"México",flag:"🇲🇽"},
  {code:"BR",name:"Brasil",flag:"🇧🇷"},
  {code:"AR",name:"Argentina",flag:"🇦🇷"},
  {code:"AU",name:"Australia",flag:"🇦🇺"},
  {code:"NZ",name:"New Zealand",flag:"🇳🇿"},
  {code:"JP",name:"日本",flag:"🇯🇵"},
  {code:"KR",name:"한국",flag:"🇰🇷"},
  {code:"CN",name:"中国",flag:"🇨🇳"},
  {code:"IN",name:"भारत",flag:"🇮🇳"},
  {code:"AE",name:"الإمارات",flag:"🇦🇪"},
  {code:"SA",name:"السعودية",flag:"🇸🇦"},
  {code:"IL",name:"ישראל",flag:"🇮🇱"},
  {code:"EG",name:"مصر",flag:"🇪🇬"},
  {code:"ZA",name:"South Africa",flag:"🇿🇦"},
  {code:"NG",name:"Nigeria",flag:"🇳🇬"}
];

// ═══ SHIFT TAB v2 (Sıfırdan, AI tabanlı) ═══
const ShiftTab=({team,teamMembers,phantomMembers=[],setPhantomMembers,user,t})=>{
  const lang=t.lang;
  const[shifts,setShifts]=useState([]);
  const[loading,setLoading]=useState(true);
  const fileInputRef=useRef(null);
  const[importDate,setImportDate]=useState(new Date().toISOString().slice(0,10));
  const[aiLoading,setAiLoading]=useState(false);
  const[previewShifts,setPreviewShifts]=useState(null);
  const[holidays,setHolidays]=useState({});
  const[showHoliday,setShowHoliday]=useState(false);
  const[exportMenuOpen,setExportMenuOpen]=useState(false);
  // Hafta navigasyonu (Pazartesi başlangıçlı hafta)
  // Verilen tarihin haftasının Pazartesi'sini döndür (lokal saat dilimi)
  const getMondayOf=(date)=>{
    const d=new Date(date);
    d.setHours(12,0,0,0); // gün geçişi sorunlarını önlemek için öğlene sabit
    const day=d.getDay(); // 0=Paz, 1=Pzt, ..., 6=Cmt
    const diff=day===0?-6:(1-day); // Paz ise 6 gün geri, değilse (1-day) gün
    d.setDate(d.getDate()+diff);
    // Yerel YYYY-MM-DD formatı (toISOString UTC verir, bizde kayma olabilir)
    const y=d.getFullYear();
    const m=String(d.getMonth()+1).padStart(2,"0");
    const dd=String(d.getDate()).padStart(2,"0");
    return `${y}-${m}-${dd}`;
  };
  const[weekStart,setWeekStart]=useState(()=>getMondayOf(new Date()));
  // Hücre düzenleme modal
  const[cellEdit,setCellEdit]=useState(null); // {memberId, memberType, date, existing, name}
  const[yearSummary,setYearSummary]=useState(null);
  // Hızlı vardiya şablonları (team'den yüklenir)
  const[presets,setPresets]=useState(team?.shift_presets||[
    {name:lang==="tr"?"Sabah":"Morning",start:"07:00",end:"15:00"},
    {name:lang==="tr"?"Akşam":"Evening",start:"15:00",end:"23:00"},
    {name:lang==="tr"?"Gece":"Night",start:"23:00",end:"07:00"}
  ]);
  useEffect(()=>{
    if(team?.shift_presets&&Array.isArray(team.shift_presets))setPresets(team.shift_presets);
  },[team?.id]);

  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    setLoading(true);
    sb.from("shifts").select("*").eq("team_id",team.id).order("date",{ascending:true}).limit(2000).then(({data,error})=>{
      if(!error&&data)setShifts(data);
      setLoading(false);
    });
    const year=new Date().getFullYear();
    const country=team?.country||"TR";
    setHolidays({...getHolidaysForCountry(country,year),...getHolidaysForCountry(country,year+1)});
  },[team?.id]);

  // Birleşik üye listesi (gerçek + phantom, linked olmayan)
  const allMembers=useMemo(()=>{
    const list=[
      ...teamMembers.map(m=>({id:m.userId||m.user_id,type:"real",name:m.name||"",position:m.position||m.role||"",annual_leave_total:m.annual_leave_total||14,_obj:m})),
      ...phantomMembers.filter(p=>!p.linked_user_id).map(p=>({id:p.id,type:"phantom",name:p.name||"",position:p.position||"",annual_leave_total:p.annual_leave_total||14,_obj:p}))
    ];
    return list;
  },[teamMembers,phantomMembers]);

  // Haftanın 7 günü
  const weekDays=useMemo(()=>{
    const days=[];
    const start=new Date(weekStart+"T12:00:00");
    for(let i=0;i<7;i++){
      const d=new Date(start);
      d.setDate(d.getDate()+i);
      const y=d.getFullYear();
      const m=String(d.getMonth()+1).padStart(2,"0");
      const dd=String(d.getDate()).padStart(2,"0");
      days.push({
        date:`${y}-${m}-${dd}`,
        dayNum:d.getDate(),
        weekday:d.toLocaleDateString(lang==="tr"?"tr-TR":"en-US",{weekday:"short"}),
        isWeekend:((team?.country?getWeekendDays(team.country):[0,6])).includes(d.getDay())
      });
    }
    return days;
  },[weekStart,lang]);

  // Bu haftaki vardiyalar (member×date map)
  const shiftMap=useMemo(()=>{
    const m={};
    const weekDates=new Set(weekDays.map(d=>d.date));
    shifts.forEach(s=>{
      if(!weekDates.has(s.date))return;
      const key=(s.phantom_member_id||s.created_by)+"|"+s.date;
      m[key]=s;
    });
    return m;
  },[shifts,weekDays]);

  // Üyenin haftalık toplam saati hesapla
  const calcHours=(memberId)=>{
    let total=0;
    weekDays.forEach(d=>{
      const s=shiftMap[memberId+"|"+d.date];
      if(s&&s.type!=="leave"&&s.type!=="off"&&s.start_time&&s.end_time){
        const[sh,sm]=s.start_time.split(":").map(Number);
        const[eh,em]=s.end_time.split(":").map(Number);
        let diff=(eh*60+em)-(sh*60+sm);
        if(diff<0)diff+=24*60; // Gece vardiyası
        total+=diff/60;
      }
    });
    return total;
  };

  // Kullanılan yıllık izin
  const calcUsedLeave=(memberId)=>{
    return shifts.filter(s=>(s.phantom_member_id===memberId||s.created_by===memberId)&&s.type==="leave").length;
  };
  // Tüm izin tiplerini say
  const calcLeaveByType=(memberId,type)=>{
    return shifts.filter(s=>(s.phantom_member_id===memberId||s.created_by===memberId)&&s.type===type).length;
  };
  // Yıllık özet hesabı
  const getYearStats=(memberId,year)=>{
    const yearStr=String(year);
    const memberShifts=shifts.filter(s=>(s.phantom_member_id===memberId||s.created_by===memberId)&&s.date?.startsWith(yearStr));
    let totalHours=0,shiftCount=0;
    const byType={leave:0,sick:0,parental:0,training:0,unpaid:0,off:0};
    memberShifts.forEach(s=>{
      if(s.type==="shift"||!s.type){
        shiftCount++;
        if(s.start_time&&s.end_time){
          const[sh,sm]=s.start_time.split(":").map(Number);
          const[eh,em]=s.end_time.split(":").map(Number);
          let diff=(eh*60+em)-(sh*60+sm);
          if(diff<0)diff+=24*60;
          totalHours+=diff/60;
        }
      }else if(byType[s.type]!==undefined){
        byType[s.type]++;
      }
    });
    return{shiftCount,totalHours,...byType};
  };

  // Hafta navigasyonu
  const shiftWeek=(delta)=>{
    const d=new Date(weekStart+"T12:00:00");
    d.setDate(d.getDate()+delta*7);
    const y=d.getFullYear();
    const m=String(d.getMonth()+1).padStart(2,"0");
    const dd=String(d.getDate()).padStart(2,"0");
    setWeekStart(`${y}-${m}-${dd}`);
  };

  // Hücre kaydet
  const saveCell=async({memberId,memberType,date,start,end,type,position,name})=>{
    const sb=initSupabase();if(!sb)return;
    // Önce mevcut kaydı sil
    let delQ=sb.from("shifts").delete().eq("team_id",team.id).eq("date",date);
    if(memberType==="real")delQ=delQ.eq("created_by",memberId);
    else delQ=delQ.eq("phantom_member_id",memberId);
    await delQ;
    // type "delete" ise sadece sil, kayıt ekleme (hücre tamamen boş kalır)
    if(type==="delete"){
      const{data}=await sb.from("shifts").select("*").eq("team_id",team.id).order("date",{ascending:true}).limit(2000);
      if(data)setShifts(data);
      return true;
    }
    // OFF dahil tüm diğer tipler kayıt olur
    const payload={
      team_id:team.id,
      name:position||"Vardiya",
      member_name:name,
      start_time:(type==="leave"||type==="off"||type==="sick"||type==="parental"||type==="training"||type==="unpaid")?"00:00:00":(start.length===5?start+":00":start),
      end_time:(type==="leave"||type==="off"||type==="sick"||type==="parental"||type==="training"||type==="unpaid")?"00:00:00":(end.length===5?end+":00":end),
      date,
      tasks:[],
      type:type||"shift",
      ...(memberType==="real"?{created_by:memberId}:{phantom_member_id:memberId})
    };
    const{error}=await sb.from("shifts").insert(payload);
    if(error){window.toast.error(error.message);return false;}
    // Refresh
    const{data}=await sb.from("shifts").select("*").eq("team_id",team.id).order("date",{ascending:true}).limit(2000);
    if(data)setShifts(data);
    return true;
  };

  // Hücreye tıkla
  const openCellEdit=(member,date)=>{
    const existing=shiftMap[member.id+"|"+date];
    setCellEdit({
      memberId:member.id,
      memberType:member.type,
      memberName:member.name,
      position:member.position,
      date,
      existing,
      start:existing?.start_time?.slice(0,5)||"09:00",
      end:existing?.end_time?.slice(0,5)||"18:00",
      type:existing?.type||"shift"
    });
  };

  // Excel İçe (AI parser)
  const handleFileChange=async(e)=>{
    const file=e.target.files?.[0];
    e.target.value="";
    if(!file)return;
    if(!team?.id){window.toast.error(lang==="tr"?"Önce ekip oluşturun":"Create a team first");return;}
    let rows=[];
    try{
      const isExcel=/\.(xlsx|xls)$/i.test(file.name);
      if(isExcel){
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}).filter(r=>r.some(c=>String(c).trim()));
      }else{
        const text=await file.text();
        rows=text.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(/\t|,/));
      }
    }catch(err){window.toast.error((lang==="tr"?"Dosya okunamadı: ":"File read error: ")+err.message);return;}
    if(rows.length<2){window.toast.error(lang==="tr"?"Geçersiz dosya":"Invalid file");return;}
    setAiLoading(true);
    try{
      const resp=await fetch("https://kitchen-manager-ai.aligny0.workers.dev/parse-shift-excel",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({rows:rows.map(r=>r.map(c=>String(c))),startDate:importDate,lang})
      });
      if(!resp.ok)throw new Error("HTTP "+resp.status);
      const data=await resp.json();
      if(data.error)throw new Error(data.error);
      const flat=[];
      for(const s of (data.shifts||[])){
        for(const e of (s.entries||[])){
          if(!e.date||!e.start||!e.end)continue;
          flat.push({name:s.name||"",position:s.position||"Vardiya",date:e.date,start:e.start.slice(0,5),end:e.end.slice(0,5)});
        }
      }
      if(flat.length===0)window.toast.error(lang==="tr"?"AI hiç vardiya bulamadı":"AI found no shifts");
      else setPreviewShifts(flat);
    }catch(err){window.toast.error((lang==="tr"?"AI hatası: ":"AI error: ")+err.message);}
    finally{setAiLoading(false);}
  };

  const updatePreview=(i,field,value)=>setPreviewShifts(prev=>{const c=[...prev];c[i]={...c[i],[field]:value};return c;});
  const removePreview=(i)=>setPreviewShifts(prev=>prev.filter((_,idx)=>idx!==i));

  const importAll=async()=>{
    if(!previewShifts||!previewShifts.length)return;
    const sb=initSupabase();if(!sb)return;
    let success=0,failed=0,newPhantoms=[];
    for(const s of previewShifts){
      const sName=(s.name||"").trim();
      const sNameLower=sName.toLowerCase();
      const realMember=teamMembers.find(m=>{const mn=(m.name||"").toLowerCase().trim();return mn===sNameLower||mn.includes(sNameLower)||sNameLower.includes(mn);});
      let phantomId=null,creator=null;
      if(realMember)creator=realMember.userId||realMember.user_id;
      else{
        const allP=[...phantomMembers,...newPhantoms];
        let phantom=allP.find(p=>{const pn=(p.name||"").toLowerCase().trim();return pn===sNameLower||pn.includes(sNameLower)||sNameLower.includes(pn);});
        if(!phantom){
          const{data:created,error:perr}=await sb.from("team_phantom_members").insert({team_id:team.id,name:sName,position:s.position||null,created_by:user.userId}).select().single();
          if(perr){failed++;continue;}
          phantom=created;newPhantoms.push(created);
        }
        phantomId=phantom.id;
      }
      try{
        let delQ=sb.from("shifts").delete().eq("team_id",team.id).eq("date",s.date);
        if(creator)delQ=delQ.eq("created_by",creator);
        else if(phantomId)delQ=delQ.eq("phantom_member_id",phantomId);
        await delQ;
        const{error}=await sb.from("shifts").insert({
          team_id:team.id,name:s.position||"Vardiya",member_name:sName,
          start_time:s.start+":00",end_time:s.end+":00",date:s.date,tasks:[],type:"shift",
          created_by:creator,phantom_member_id:phantomId
        });
        if(error){failed++;continue;}
        success++;
      }catch(e){failed++;}
    }
    if(newPhantoms.length>0&&setPhantomMembers){
      setPhantomMembers(prev=>[...prev,...newPhantoms]);
      LS.set("kmp_phantom_members",[...phantomMembers,...newPhantoms]);
    }
    const{data}=await sb.from("shifts").select("*").eq("team_id",team.id).order("date",{ascending:true}).limit(2000);
    if(data)setShifts(data);
    window.toast.success(lang==="tr"?`✓ ${success} vardiya${newPhantoms.length?` (${newPhantoms.length} yeni üye)`:""}`:`✓ ${success} shifts`);
    setPreviewShifts(null);
  };

  // ═══ DIŞA AKTAR — Excel ═══
  const exportExcel=()=>{
    if(typeof XLSX==="undefined"){window.toast.error("XLSX yüklenemedi");return;}
    const headerRow=[lang==="tr"?"Pozisyon":"Position",lang==="tr"?"İsim":"Name",lang==="tr"?"Toplam Saat":"Total Hours",lang==="tr"?"Kalan İzin":"Leave Left"];
    weekDays.forEach(d=>headerRow.push(`${d.dayNum} ${d.weekday}`));
    const data=[headerRow];
    allMembers.forEach(m=>{
      const row=[m.position||"",m.name,calcHours(m.id),m.annual_leave_total-calcUsedLeave(m.id)];
      weekDays.forEach(d=>{
        const s=shiftMap[m.id+"|"+d.date];
        if(!s)row.push("");
        else if(s.type==="leave")row.push(lang==="tr"?"İZİN":"LEAVE");
        else if(s.type==="off")row.push("OFF");
        else row.push(`${s.start_time?.slice(0,5)}-${s.end_time?.slice(0,5)}`);
      });
      data.push(row);
    });
    const ws=XLSX.utils.aoa_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Vardiya");
    XLSX.writeFile(wb,`vardiya_${weekStart}.xlsx`);
    setExportMenuOpen(false);
  };

  // ═══ DIŞA AKTAR — PDF (yazdır penceresi PDF olarak kaydedebilir) ═══
  const exportPDF=()=>{printTable("pdf");setExportMenuOpen(false);};
  const exportPrint=()=>{printTable("print");setExportMenuOpen(false);};

  const printTable=(mode)=>{
    const w=window.open("","_blank");
    if(!w){window.toast.error(lang==="tr"?"Popup engellendi":"Popup blocked");return;}
    const weekLabel=`${new Date(weekDays[0].date).toLocaleDateString("tr-TR")} - ${new Date(weekDays[6].date).toLocaleDateString("tr-TR")}`;
    let html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Vardiya ${weekLabel}</title>
      <style>
        body{font-family:Arial,sans-serif;font-size:11px;margin:20px}
        h1{font-size:16px;margin-bottom:6px}
        .meta{color:#666;font-size:10px;margin-bottom:12px}
        table{border-collapse:collapse;width:100%}
        th,td{border:1px solid #888;padding:4px 6px;text-align:center}
        th{background:#f3f4f6;font-weight:bold}
        .pos{text-align:left;background:#fbbf24;font-weight:bold}
        .name{text-align:left;font-weight:600}
        .weekend{background:#fef3c7}
        .holiday{background:#fee2e2}
        .leave{background:#dcfce7;font-weight:bold}
        .off{background:#f3f4f6;color:#999}
        @media print{body{margin:0}}
      </style></head><body>
      <h1>📅 ${team.name} — ${lang==="tr"?"Haftalık Vardiya":"Weekly Shifts"}</h1>
      <div class="meta">${weekLabel}</div>
      <table><thead><tr>
        <th>${lang==="tr"?"Pozisyon":"Position"}</th>
        <th>${lang==="tr"?"İsim":"Name"}</th>
        <th>${lang==="tr"?"Toplam":"Total"}</th>
        <th>${lang==="tr"?"İzin":"Leave"}</th>`;
    weekDays.forEach(d=>{
      const isH=!!holidays[d.date];
      html+=`<th class="${isH?"holiday":(d.isWeekend?"weekend":"")}">${d.dayNum} ${d.weekday}${isH?" 🎌":""}</th>`;
    });
    html+=`</tr></thead><tbody>`;
    allMembers.forEach(m=>{
      html+=`<tr><td class="pos">${m.position||""}</td><td class="name">${m.name}</td><td>${calcHours(m.id)}</td><td>${m.annual_leave_total-calcUsedLeave(m.id)}</td>`;
      weekDays.forEach(d=>{
        const s=shiftMap[m.id+"|"+d.date];
        const isH=!!holidays[d.date];
        const cls=s?.type==="leave"?"leave":(s?.type==="off"?"off":(isH?"holiday":(d.isWeekend?"weekend":"")));
        let txt="";
        if(s){
          if(s.type==="leave")txt=lang==="tr"?"İZİN":"LEAVE";
          else if(s.type==="off")txt="OFF";
          else txt=`${s.start_time?.slice(0,5)}-${s.end_time?.slice(0,5)}`;
        }
        html+=`<td class="${cls}">${txt}</td>`;
      });
      html+=`</tr>`;
    });
    html+=`</tbody></table>
      <script>window.onload=()=>{${mode==="print"?"window.print();":""}};</script>
      </body></html>`;
    w.document.write(html);
    w.document.close();
  };

  if(!team){
    return <div style={{padding:24,textAlign:"center",color:t.tm}}>{lang==="tr"?"Vardiya için önce ekip oluşturun.":"Create a team first."}</div>;
  }

  const weekLabel=`${new Date(weekDays[0].date).toLocaleDateString(lang==="tr"?"tr-TR":"en-US",{day:"numeric",month:"short"})} - ${new Date(weekDays[6].date).toLocaleDateString(lang==="tr"?"tr-TR":"en-US",{day:"numeric",month:"short",year:"numeric"})}`;

  return <div style={{padding:"10px 12px"}}>
    {/* Üst bar */}
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:6}}>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <button onClick={()=>shiftWeek(-1)} style={{...bSt("s",t),padding:"6px 10px",fontSize:13}}>◀</button>
        <div style={{fontSize:13,fontWeight:700,color:t.text,minWidth:140,textAlign:"center"}}>{weekLabel}</div>
        <button onClick={()=>shiftWeek(1)} style={{...bSt("s",t),padding:"6px 10px",fontSize:13}}>▶</button>
        <button onClick={()=>setWeekStart(getMondayOf(new Date()))} style={{...bSt("s",t),fontSize:10,padding:"4px 8px"}}>{lang==="tr"?"Bu Hafta":"This Week"}</button>
        <button onClick={async()=>{
          if(!confirm(lang==="tr"?"Bu haftanın tüm vardiyaları sonraki haftaya kopyalansın mı?":"Copy all this week's shifts to next week?"))return;
          const sb=initSupabase();if(!sb)return;
          const weekShifts=shifts.filter(s=>weekDays.some(d=>d.date===s.date));
          if(weekShifts.length===0){window.toast.info(lang==="tr"?"Bu hafta vardiya yok":"No shifts this week");return;}
          const nextWeekStart=new Date(weekStart+"T12:00:00");
          nextWeekStart.setDate(nextWeekStart.getDate()+7);
          let success=0;
          for(const s of weekShifts){
            const oldDate=new Date(s.date+"T12:00:00");
            const newDate=new Date(oldDate);
            newDate.setDate(newDate.getDate()+7);
            const newDateStr=`${newDate.getFullYear()}-${String(newDate.getMonth()+1).padStart(2,"0")}-${String(newDate.getDate()).padStart(2,"0")}`;
            try{
              let delQ=sb.from("shifts").delete().eq("team_id",team.id).eq("date",newDateStr);
              if(s.created_by)delQ=delQ.eq("created_by",s.created_by);
              else if(s.phantom_member_id)delQ=delQ.eq("phantom_member_id",s.phantom_member_id);
              await delQ;
              const{error}=await sb.from("shifts").insert({
                team_id:team.id,name:s.name,member_name:s.member_name,
                start_time:s.start_time,end_time:s.end_time,
                date:newDateStr,tasks:[],type:s.type||"shift",
                created_by:s.created_by,phantom_member_id:s.phantom_member_id
              });
              if(!error)success++;
            }catch(e){}
          }
          const{data}=await sb.from("shifts").select("*").eq("team_id",team.id).order("date",{ascending:true}).limit(2000);
          if(data)setShifts(data);
          window.toast.success(lang==="tr"?`✓ ${success} vardiya kopyalandı`:`✓ ${success} copied`);
          // Sonraki haftaya geç
          const y=nextWeekStart.getFullYear();
          const m=String(nextWeekStart.getMonth()+1).padStart(2,"0");
          const dd=String(nextWeekStart.getDate()).padStart(2,"0");
          setWeekStart(`${y}-${m}-${dd}`);
        }} title={lang==="tr"?"Bu haftayı sonrakine kopyala":"Copy to next week"} style={{...bSt("s",t),fontSize:10,padding:"4px 8px"}}>📋➡</button>
      </div>
      <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",position:"relative"}}>
        <input type="date" value={importDate} onChange={e=>setImportDate(e.target.value)} title={lang==="tr"?"Excel başlangıç tarihi":"Excel start"} style={{...iSt(t),fontSize:10,padding:"4px 6px",width:120}}/>
        <button onClick={()=>fileInputRef.current?.click()} style={{...bSt("s",t),fontSize:11}}>📥 {lang==="tr"?"İçe":"Import"}</button>
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={handleFileChange}/>
        <button onClick={()=>setExportMenuOpen(!exportMenuOpen)} style={{...bSt("s",t),fontSize:11}}>📤 {lang==="tr"?"Dışa":"Export"}</button>
        {exportMenuOpen&&<div style={{position:"absolute",top:34,right:0,background:t.cardBg||t.bg,border:`1px solid ${t.border}`,borderRadius:8,boxShadow:"0 4px 12px rgba(0,0,0,0.15)",zIndex:100,minWidth:160}}>
          <button onClick={exportExcel} style={{display:"block",width:"100%",textAlign:"left",padding:"8px 12px",background:"transparent",border:"none",cursor:"pointer",color:t.text,fontSize:12}}>📊 Excel</button>
          <button onClick={exportPDF} style={{display:"block",width:"100%",textAlign:"left",padding:"8px 12px",background:"transparent",border:"none",cursor:"pointer",color:t.text,fontSize:12}}>📄 PDF</button>
          <button onClick={exportPrint} style={{display:"block",width:"100%",textAlign:"left",padding:"8px 12px",background:"transparent",border:"none",cursor:"pointer",color:t.text,fontSize:12}}>🖨️ {lang==="tr"?"Yazdır":"Print"}</button>
        </div>}
        <button onClick={()=>setShowHoliday(!showHoliday)} style={{...bSt("s",t),fontSize:11}}>🎌</button>
      </div>
    </div>

    {/* Tablo */}
    {allMembers.length===0?<div style={{textAlign:"center",padding:40,color:t.tm}}>
      <div style={{fontSize:32,marginBottom:8}}>👥</div>
      <div style={{fontSize:13}}>{lang==="tr"?"Henüz üye yok":"No members"}</div>
      <div style={{fontSize:11,marginTop:4}}>{lang==="tr"?"Ayarlar → Ekip'ten üye ekleyin":"Add members from Settings → Team"}</div>
    </div>:<div style={{overflowX:"auto",border:`1px solid ${t.border}`,borderRadius:8,background:t.cardBg||t.bg}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:800}}>
        <thead>
          <tr style={{background:t.inBg||"#f9fafb"}}>
            <th style={{padding:"6px 8px",textAlign:"left",borderBottom:`1px solid ${t.border}`,fontWeight:700,fontSize:10,color:t.tm,whiteSpace:"nowrap",position:"sticky",left:0,background:t.inBg||"#f9fafb",zIndex:2}}>{lang==="tr"?"Pozisyon":"Position"}</th>
            <th style={{padding:"6px 8px",textAlign:"left",borderBottom:`1px solid ${t.border}`,fontWeight:700,fontSize:10,color:t.tm,whiteSpace:"nowrap"}}>{lang==="tr"?"İsim":"Name"}</th>
            <th style={{padding:"6px 6px",textAlign:"center",borderBottom:`1px solid ${t.border}`,fontWeight:700,fontSize:10,color:t.tm,whiteSpace:"nowrap"}}>{lang==="tr"?"Saat":"Hrs"}</th>
            <th style={{padding:"6px 6px",textAlign:"center",borderBottom:`1px solid ${t.border}`,fontWeight:700,fontSize:10,color:t.tm,whiteSpace:"nowrap"}}>{lang==="tr"?"İzin":"Lv"}</th>
            {weekDays.map(d=>{
              const isH=!!holidays[d.date];
              const bg=isH?"#fee2e2":(d.isWeekend?"#fef3c7":(t.inBg||"#f9fafb"));
              return <th key={d.date} style={{padding:"6px 4px",textAlign:"center",borderBottom:`1px solid ${t.border}`,fontWeight:700,fontSize:10,color:t.text,whiteSpace:"nowrap",background:bg,minWidth:80}}>
                <div>{d.dayNum} {d.weekday}</div>
                {isH&&<div style={{fontSize:9,color:"#dc2626"}}>🎌</div>}
              </th>;
            })}
          </tr>
        </thead>
        <tbody>
          {allMembers.map((m,idx)=>{
            const hours=calcHours(m.id);
            const leaveLeft=m.annual_leave_total-calcUsedLeave(m.id);
            return <tr key={m.id}>
              <td style={{padding:"5px 8px",borderBottom:`1px solid ${t.border}`,fontWeight:600,color:t.text,whiteSpace:"nowrap",position:"sticky",left:0,background:idx%2?t.bg:(t.cardBg||t.bg),zIndex:1,fontSize:11}}>{m.position||"—"}</td>
              <td onClick={()=>setYearSummary({memberId:m.id,memberName:m.name,memberType:m.type})} style={{padding:"5px 8px",borderBottom:`1px solid ${t.border}`,color:t.text,whiteSpace:"nowrap",fontSize:11,cursor:"pointer",textDecoration:"underline",textDecorationColor:t.tm,textDecorationStyle:"dotted",textUnderlineOffset:3}}>{m.name}{m.type==="phantom"&&<span style={{fontSize:9,color:t.tm,marginLeft:4}}>👤❓</span>}</td>
              <td style={{padding:"5px 6px",borderBottom:`1px solid ${t.border}`,textAlign:"center",color:t.text,fontSize:11,fontWeight:600}}>{hours}</td>
              <td style={{padding:"5px 6px",borderBottom:`1px solid ${t.border}`,textAlign:"center",color:leaveLeft<3?t.danger:t.tm,fontSize:11}}>{leaveLeft}</td>
              {weekDays.map(d=>{
                const s=shiftMap[m.id+"|"+d.date];
                const isH=!!holidays[d.date];
                let bg="transparent",txt="",fg=t.text,fw=400;
                if(s){
                  if(s.type==="leave"){bg="#dcfce7";txt=lang==="tr"?"İZİN":"LEAVE";fg="#15803d";fw=700;}
                  else if(s.type==="sick"){bg="#ffedd5";txt=lang==="tr"?"HSP":"SICK";fg="#c2410c";fw=700;}
                  else if(s.type==="parental"){bg="#fce7f3";txt=lang==="tr"?"DOĞ":"PAR";fg="#9d174d";fw=700;}
                  else if(s.type==="training"){bg="#dbeafe";txt=lang==="tr"?"EĞT":"TRN";fg="#1e40af";fw=700;}
                  else if(s.type==="unpaid"){bg="#e5e7eb";txt=lang==="tr"?"ÜCR":"UNP";fg="#374151";fw=700;}
                  else if(s.type==="off"){bg="#f3f4f6";txt="OFF";fg="#9ca3af";}
                  else{txt=`${s.start_time?.slice(0,5)}-${s.end_time?.slice(0,5)}`;}
                }
                if(!s&&isH)bg="#fee2e2";
                else if(!s&&d.isWeekend)bg="#fef9c3";
                return <td key={d.date} onClick={()=>openCellEdit(m,d.date)} style={{padding:"5px 4px",borderBottom:`1px solid ${t.border}`,borderLeft:`1px solid ${t.border}`,textAlign:"center",cursor:"pointer",background:bg,color:fg,fontSize:10,fontWeight:fw,whiteSpace:"nowrap"}}>{txt||(isH?"🎌":"")}</td>;
              })}
            </tr>;
          })}
        </tbody>
      </table>
    </div>}

    {/* Renk açıklamaları */}
    <div style={{display:"flex",gap:12,flexWrap:"wrap",marginTop:8,fontSize:10,color:t.tm}}>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#fef9c3",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Hafta Sonu":"Weekend"}</span>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#fee2e2",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Resmi Tatil":"Holiday"}</span>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#dcfce7",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Yıllık İzin":"Annual Leave"}</span>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#ffedd5",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Hastalık":"Sick"}</span>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#fce7f3",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Doğum":"Parental"}</span>
      <span><span style={{display:"inline-block",width:10,height:10,background:"#dbeafe",border:`1px solid ${t.border}`,verticalAlign:"middle",marginRight:4}}/>{lang==="tr"?"Eğitim":"Training"}</span>
      <span>👤❓ {lang==="tr"?"Kayıtsız":"Unregistered"}</span>
    </div>

    {/* AI Loading */}
    {aiLoading&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:28,textAlign:"center",maxWidth:300}}>
        <div style={{fontSize:36,marginBottom:10}}>🤖</div>
        <div style={{fontSize:14,fontWeight:600,color:t.text}}>{lang==="tr"?"AI tabloyu okuyor...":"AI is reading..."}</div>
      </div>
    </div>}

    {/* Önizleme modal (Excel import sonrası) */}
    {previewShifts&&<div onClick={e=>{if(e.target===e.currentTarget)setPreviewShifts(null);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:18,maxWidth:700,width:"100%",maxHeight:"90vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <div style={{fontSize:16,fontWeight:700,color:t.text}}>🤖 {lang==="tr"?"Önizleme":"Preview"}</div>
          <div style={{fontSize:11,color:t.tm}}>{previewShifts.length} {lang==="tr"?"vardiya":"shifts"}</div>
        </div>
        <div style={{flex:1,overflow:"auto",border:`1px solid ${t.border}`,borderRadius:8,padding:8}}>
          {previewShifts.map((s,i)=><div key={i} style={{display:"flex",gap:6,alignItems:"center",padding:"4px 0",borderBottom:`1px solid ${t.border}`,flexWrap:"wrap"}}>
            <input value={s.name} onChange={e=>updatePreview(i,"name",e.target.value)} style={{...iSt(t),fontSize:11,padding:"4px 6px",flex:"1 1 100px"}}/>
            <input value={s.position} onChange={e=>updatePreview(i,"position",e.target.value)} style={{...iSt(t),fontSize:11,padding:"4px 6px",flex:"1 1 80px"}}/>
            <input type="date" value={s.date} onChange={e=>updatePreview(i,"date",e.target.value)} style={{...iSt(t),fontSize:11,padding:"4px 6px",width:130}}/>
            <input type="time" value={s.start} onChange={e=>updatePreview(i,"start",e.target.value)} style={{...iSt(t),fontSize:11,padding:"4px 6px",width:80}}/>
            <input type="time" value={s.end} onChange={e=>updatePreview(i,"end",e.target.value)} style={{...iSt(t),fontSize:11,padding:"4px 6px",width:80}}/>
            <button onClick={()=>removePreview(i)} style={{background:"transparent",border:"none",color:t.danger,cursor:"pointer",fontSize:14}}>✕</button>
          </div>)}
        </div>
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <button onClick={()=>setPreviewShifts(null)} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
          <button onClick={importAll} style={{...bSt("p",t),flex:2}}>{lang==="tr"?`✓ ${previewShifts.length} Aktar`:`✓ Import ${previewShifts.length}`}</button>
        </div>
      </div>
    </div>}

    {/* Yıllık Özet Modal */}
    {yearSummary&&(()=>{
      const year=new Date(weekStart).getFullYear();
      const stats=getYearStats(yearSummary.memberId,year);
      const m=allMembers.find(x=>x.id===yearSummary.memberId);
      const leaveLeft=(m?.annual_leave_total||14)-stats.leave;
      return <div onClick={e=>{if(e.target===e.currentTarget)setYearSummary(null);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
        <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:420,width:"100%"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
            <div>
              <div style={{fontSize:18,fontWeight:700,color:t.text}}>{yearSummary.memberName}</div>
              <div style={{fontSize:12,color:t.tm}}>{year} {lang==="tr"?"yıllık özeti":"summary"}</div>
            </div>
            <button onClick={()=>setYearSummary(null)} style={{background:"transparent",border:"none",fontSize:18,cursor:"pointer",color:t.tm}}>✕</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div style={{padding:10,background:t.inBg||"#f9fafb",borderRadius:8}}>
              <div style={{fontSize:10,color:t.tm,fontWeight:600}}>{lang==="tr"?"Toplam Vardiya":"Total Shifts"}</div>
              <div style={{fontSize:20,fontWeight:700,color:t.text}}>{stats.shiftCount}</div>
            </div>
            <div style={{padding:10,background:t.inBg||"#f9fafb",borderRadius:8}}>
              <div style={{fontSize:10,color:t.tm,fontWeight:600}}>{lang==="tr"?"Toplam Saat":"Total Hours"}</div>
              <div style={{fontSize:20,fontWeight:700,color:t.text}}>{Math.round(stats.totalHours)}</div>
            </div>
            <div style={{padding:10,background:"#dcfce7",borderRadius:8}}>
              <div style={{fontSize:10,color:"#15803d",fontWeight:600}}>🌴 {lang==="tr"?"Yıllık İzin":"Leave"}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#15803d"}}>{stats.leave} / {m?.annual_leave_total||14}</div>
              <div style={{fontSize:9,color:"#15803d",marginTop:2}}>{lang==="tr"?`${leaveLeft} gün kaldı`:`${leaveLeft} days left`}</div>
            </div>
            <div style={{padding:10,background:"#ffedd5",borderRadius:8}}>
              <div style={{fontSize:10,color:"#c2410c",fontWeight:600}}>🤒 {lang==="tr"?"Hastalık":"Sick"}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#c2410c"}}>{stats.sick}</div>
            </div>
            <div style={{padding:10,background:"#fce7f3",borderRadius:8}}>
              <div style={{fontSize:10,color:"#9d174d",fontWeight:600}}>👶 {lang==="tr"?"Doğum/Babalık":"Parental"}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#9d174d"}}>{stats.parental}</div>
            </div>
            <div style={{padding:10,background:"#dbeafe",borderRadius:8}}>
              <div style={{fontSize:10,color:"#1e40af",fontWeight:600}}>📚 {lang==="tr"?"Eğitim":"Training"}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#1e40af"}}>{stats.training}</div>
            </div>
            <div style={{padding:10,background:"#e5e7eb",borderRadius:8,gridColumn:"span 2"}}>
              <div style={{fontSize:10,color:"#374151",fontWeight:600}}>💸 {lang==="tr"?"Ücretsiz İzin":"Unpaid Leave"}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#374151"}}>{stats.unpaid}</div>
            </div>
          </div>
        </div>
      </div>;
    })()}

    {/* Hücre düzenleme modal */}
    {cellEdit&&<div onClick={e=>{if(e.target===e.currentTarget)setCellEdit(null);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:380,width:"100%"}}>
        <div style={{fontSize:15,fontWeight:700,color:t.text,marginBottom:4}}>{cellEdit.memberName}</div>
        <div style={{fontSize:11,color:t.tm,marginBottom:14}}>{new Date(cellEdit.date).toLocaleDateString(lang==="tr"?"tr-TR":"en-US",{weekday:"long",day:"numeric",month:"long"})}</div>
        
        {/* ÜST SIRA: OFF + YILLIK İZİN (en sık kullanılanlar) */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
          <button onClick={async()=>{
            const ok=await saveCell({...cellEdit,type:"off",name:cellEdit.memberName});
            if(ok){window.toast.success(lang==="tr"?"✓ OFF":"✓ OFF");setCellEdit(null);}
          }} style={{...bSt("s",t),background:"#f3f4f6",color:"#374151",fontSize:13,padding:"12px 8px",fontWeight:700}}>⊘ OFF</button>
          <button onClick={async()=>{
            const ok=await saveCell({...cellEdit,type:"leave",name:cellEdit.memberName});
            if(ok){window.toast.success(lang==="tr"?"✓ Yıllık izin":"✓ Annual leave");setCellEdit(null);}
          }} style={{...bSt("s",t),background:"#dcfce7",color:"#15803d",fontSize:13,padding:"12px 8px",fontWeight:700}}>🌴 {lang==="tr"?"YILLIK İZİN":"LEAVE"}</button>
        </div>

        {/* Vardiya saati (manuel) */}
        <div style={{padding:12,background:t.inBg||"#f9fafb",borderRadius:8,marginBottom:10}}>
          <div style={{fontSize:11,fontWeight:600,color:t.tm,marginBottom:6}}>{lang==="tr"?"VARDİYA SAATİ":"SHIFT TIME"}</div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <input type="time" value={cellEdit.start} onChange={e=>setCellEdit({...cellEdit,start:e.target.value,type:"shift"})} style={{...iSt(t),flex:1}}/>
            <span style={{color:t.tm}}>→</span>
            <input type="time" value={cellEdit.end} onChange={e=>setCellEdit({...cellEdit,end:e.target.value,type:"shift"})} style={{...iSt(t),flex:1}}/>
          </div>
          <button onClick={async()=>{
            const ok=await saveCell({...cellEdit,type:"shift",name:cellEdit.memberName});
            if(ok){window.toast.success(lang==="tr"?"✓ Kaydedildi":"✓ Saved");setCellEdit(null);}
          }} style={{...bSt("p",t),width:"100%",marginTop:8,fontSize:12}}>{lang==="tr"?"✓ Vardiyayı Kaydet":"✓ Save Shift"}</button>
        </div>

        {/* + Diğer izin tipleri */}
        {cellEdit.showMore?<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
            <button onClick={async()=>{
              const ok=await saveCell({...cellEdit,type:"sick",name:cellEdit.memberName});
              if(ok){window.toast.success(lang==="tr"?"✓ Hastalık raporu":"✓ Sick");setCellEdit(null);}
            }} style={{...bSt("s",t),background:"#ffedd5",color:"#c2410c",fontSize:11,padding:"8px 4px"}}>🤒 {lang==="tr"?"Hastalık":"Sick"}</button>
            <button onClick={async()=>{
              const ok=await saveCell({...cellEdit,type:"parental",name:cellEdit.memberName});
              if(ok){window.toast.success(lang==="tr"?"✓ Doğum/Babalık":"✓ Parental");setCellEdit(null);}
            }} style={{...bSt("s",t),background:"#fce7f3",color:"#9d174d",fontSize:11,padding:"8px 4px"}}>👶 {lang==="tr"?"Doğum/Babalık":"Parental"}</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:8}}>
            <button onClick={async()=>{
              const ok=await saveCell({...cellEdit,type:"training",name:cellEdit.memberName});
              if(ok){window.toast.success(lang==="tr"?"✓ Eğitim":"✓ Training");setCellEdit(null);}
            }} style={{...bSt("s",t),background:"#dbeafe",color:"#1e40af",fontSize:11,padding:"8px 4px"}}>📚 {lang==="tr"?"Eğitim":"Training"}</button>
            <button onClick={async()=>{
              const ok=await saveCell({...cellEdit,type:"unpaid",name:cellEdit.memberName});
              if(ok){window.toast.success(lang==="tr"?"✓ Ücretsiz izin":"✓ Unpaid");setCellEdit(null);}
            }} style={{...bSt("s",t),background:"#e5e7eb",color:"#374151",fontSize:11,padding:"8px 4px"}}>💸 {lang==="tr"?"Ücretsiz":"Unpaid"}</button>
          </div>
        </>:<button onClick={()=>setCellEdit({...cellEdit,showMore:true})} style={{...bSt("s",t),width:"100%",fontSize:11,padding:"6px",marginBottom:8,color:t.tm}}>+ {lang==="tr"?"Diğer izin tipleri":"More leave types"}</button>}
        
        {/* En altta: Sil (sadece mevcut vardiya varsa) ve İptal */}
        <div style={{display:"flex",gap:8}}>
          {cellEdit.existing&&<button onClick={async()=>{
            const ok=await saveCell({...cellEdit,type:"delete",name:cellEdit.memberName});
            if(ok){window.toast.success(lang==="tr"?"✓ Silindi":"✓ Deleted");setCellEdit(null);}
          }} style={{...bSt("s",t),flex:1,fontSize:12,color:t.danger}}>🗑️ {lang==="tr"?"Sil":"Delete"}</button>}
          <button onClick={()=>setCellEdit(null)} style={{...bSt("s",t),flex:1,fontSize:12}}>{lang==="tr"?"İptal":"Cancel"}</button>
        </div>
      </div>
    </div>}

    {/* Tatiller modal */}
    {showHoliday&&<div onClick={e=>{if(e.target===e.currentTarget)setShowHoliday(false);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:400,width:"100%"}}>
        <div style={{fontSize:16,fontWeight:700,color:t.text,marginBottom:10}}>🎌 {lang==="tr"?"Resmi Tatiller":"Holidays"}</div>
        <div style={{maxHeight:300,overflow:"auto"}}>
          {Object.entries(holidays).sort().map(([d,name])=><div key={d} style={{padding:6,fontSize:12,borderBottom:`1px solid ${t.border}`}}>{new Date(d).toLocaleDateString(lang==="tr"?"tr-TR":"en-US")} — {name}</div>)}
        </div>
        <button onClick={()=>setShowHoliday(false)} style={{...bSt("p",t),width:"100%",marginTop:12}}>{lang==="tr"?"Kapat":"Close"}</button>
      </div>
    </div>}
  </div>;
};


// ═══ KANBAN TAB ═══
// ═══ HIZLI VARDİYA ŞABLONLARI ═══
const ShiftPresetsCard=({team,setTeam,t,lang})=>{
  const defaults=[
    {name:lang==="tr"?"Sabah":"Morning",start:"07:00",end:"15:00"},
    {name:lang==="tr"?"Akşam":"Evening",start:"15:00",end:"23:00"},
    {name:lang==="tr"?"Gece":"Night",start:"23:00",end:"07:00"}
  ];
  const[presets,setPresets]=useState(team?.shift_presets&&Array.isArray(team.shift_presets)?team.shift_presets:defaults);
  const[showAdd,setShowAdd]=useState(false);
  const[newP,setNewP]=useState({name:"",start:"09:00",end:"18:00"});

  const savePresets=async(list)=>{
    const sb=initSupabase();if(!sb)return;
    const{error}=await sb.from("teams").update({shift_presets:list}).eq("id",team.id);
    if(error){window.toast.error(error.message);return;}
    setPresets(list);
    setTeam({...team,shift_presets:list});
    LS.set("kmp_team",{...team,shift_presets:list});
  };

  return <div style={{...cSt(t),padding:"12px 14px",marginBottom:12}}>
    <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span>⚡ {lang==="tr"?"HIZLI VARDİYA ŞABLONLARI":"QUICK SHIFT TEMPLATES"}</span>
      <button onClick={()=>setShowAdd(true)} style={{...bSt("s",t),fontSize:10,padding:"4px 10px"}}>+ {lang==="tr"?"Ekle":"Add"}</button>
    </div>
    <div style={{fontSize:11,color:t.tm,marginBottom:8,lineHeight:1.4}}>
      {lang==="tr"?"Vardiya tablosunda bir hücreye tıkladığında bu şablonlar tek tıkla atanabilir.":"These templates appear as one-click options when editing shift cells."}
    </div>
    {presets.length===0&&<div style={{fontSize:11,color:t.tm,padding:8,textAlign:"center"}}>{lang==="tr"?"Henüz şablon yok":"No templates"}</div>}
    <div style={{display:"flex",flexDirection:"column",gap:6}}>
      {presets.map((p,i)=><div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 10px",background:t.bg,border:`1px solid ${t.border}`,borderRadius:8}}>
        <div style={{display:"flex",alignItems:"center",gap:10,flex:1}}>
          <div style={{fontSize:13,fontWeight:600,color:t.text}}>{p.name}</div>
          <div style={{fontSize:11,color:t.tm}}>{p.start} → {p.end}</div>
        </div>
        <button onClick={()=>{if(confirm(lang==="tr"?"Silinsin mi?":"Delete?"))savePresets(presets.filter((_,idx)=>idx!==i));}} style={{background:"transparent",border:"none",color:t.danger,cursor:"pointer",fontSize:14}}>✕</button>
      </div>)}
    </div>
    {showAdd&&<div onClick={e=>{if(e.target===e.currentTarget)setShowAdd(false);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:380,width:"100%"}}>
        <div style={{fontSize:15,fontWeight:700,color:t.text,marginBottom:10}}>+ {lang==="tr"?"Yeni Şablon":"New Template"}</div>
        <div style={{display:"grid",gap:8}}>
          <input value={newP.name} onChange={e=>setNewP({...newP,name:e.target.value})} placeholder={lang==="tr"?"İsim (Sabah, Pastane vb.)":"Name"} style={{...iSt(t)}}/>
          <div style={{display:"flex",gap:8}}>
            <input type="time" value={newP.start} onChange={e=>setNewP({...newP,start:e.target.value})} style={{...iSt(t),flex:1}}/>
            <input type="time" value={newP.end} onChange={e=>setNewP({...newP,end:e.target.value})} style={{...iSt(t),flex:1}}/>
          </div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <button onClick={()=>setShowAdd(false)} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
          <button onClick={async()=>{
            if(!newP.name.trim()){window.toast.error(lang==="tr"?"İsim gerekli":"Name required");return;}
            await savePresets([...presets,newP]);
            setNewP({name:"",start:"09:00",end:"18:00"});
            setShowAdd(false);
          }} style={{...bSt("p",t),flex:1}}>{lang==="tr"?"Ekle":"Add"}</button>
        </div>
      </div>
    </div>}
  </div>;
};

// ═══ PHANTOM MEMBERS SECTION — Ekip üyeleri (gerçek + phantom) ═══
const PhantomMembersSection=({team,teamMembers,phantomMembers,setPhantomMembers,user,t,lang})=>{
  const[showAdd,setShowAdd]=useState(false);
  const[newPhantom,setNewPhantom]=useState({name:"",position:"",department:"",email:"",phone:""});
  const[inviteFor,setInviteFor]=useState(null); // phantom obj — davet linki gösterme

  const addPhantom=async()=>{
    if(!newPhantom.name.trim()){window.toast.error(lang==="tr"?"İsim gerekli":"Name required");return;}
    const sb=initSupabase();if(!sb)return;
    const{data,error}=await sb.from("team_phantom_members").insert({
      team_id:team.id,
      name:newPhantom.name.trim(),
      position:newPhantom.position.trim()||null,
      department:newPhantom.department.trim()||null,
      email:newPhantom.email.trim()||null,
      phone:newPhantom.phone.trim()||null,
      created_by:user.userId
    }).select().single();
    if(error){window.toast.error(error.message);return;}
    setPhantomMembers(prev=>[...prev,data]);
    LS.set("kmp_phantom_members",[...phantomMembers,data]);
    setNewPhantom({name:"",position:"",department:"",email:"",phone:""});
    setShowAdd(false);
    window.toast.success(lang==="tr"?"✓ Üye eklendi":"✓ Member added");
  };

  const removePhantom=async(p)=>{
    if(!confirm(lang==="tr"?`${p.name} silinsin mi? Bu kişinin tüm vardiyaları da silinecek.`:`Delete ${p.name}? All their shifts will be removed.`))return;
    const sb=initSupabase();if(!sb)return;
    const{error}=await sb.from("team_phantom_members").delete().eq("id",p.id);
    if(error){window.toast.error(error.message);return;}
    setPhantomMembers(prev=>prev.filter(x=>x.id!==p.id));
    LS.set("kmp_phantom_members",phantomMembers.filter(x=>x.id!==p.id));
    window.toast.success(lang==="tr"?"Silindi":"Deleted");
  };

  // Yıllık izin günü güncelle (gerçek veya phantom üye için)
  const updateLeaveTotal=async(m,newTotal)=>{
    const sb=initSupabase();if(!sb)return;
    const val=parseInt(newTotal)||0;
    if(m.type==="phantom"){
      const{error}=await sb.from("team_phantom_members").update({annual_leave_total:val}).eq("id",m._id);
      if(error){window.toast.error(error.message);return;}
      setPhantomMembers(prev=>prev.map(x=>x.id===m._id?{...x,annual_leave_total:val}:x));
      LS.set("kmp_phantom_members",phantomMembers.map(x=>x.id===m._id?{...x,annual_leave_total:val}:x));
    }else{
      const{error}=await sb.from("team_members").update({annual_leave_total:val}).eq("team_id",team.id).eq("user_id",m._id);
      if(error){window.toast.error(error.message);return;}
    }
  };

  const generateInviteToken=async(p)=>{
    if(p.invite_token){
      setInviteFor(p);
      return;
    }
    const sb=initSupabase();if(!sb)return;
    // 24 karakter rastgele token
    const token=Array.from(crypto.getRandomValues(new Uint8Array(18))).map(b=>b.toString(36)).join("").substring(0,20);
    const{data,error}=await sb.from("team_phantom_members").update({invite_token:token,invite_created_at:new Date().toISOString()}).eq("id",p.id).select().single();
    if(error){window.toast.error(error.message);return;}
    setPhantomMembers(prev=>prev.map(x=>x.id===p.id?data:x));
    LS.set("kmp_phantom_members",phantomMembers.map(x=>x.id===p.id?data:x));
    setInviteFor(data);
  };

  const copyInviteLink=async(p)=>{
    const link=`${window.location.origin}${window.location.pathname}?invite=${p.invite_token}&team=${team.id}`;
    try{
      await navigator.clipboard.writeText(link);
      window.toast.success(lang==="tr"?"✓ Link kopyalandı":"✓ Link copied");
    }catch{
      const ta=document.createElement("textarea");ta.value=link;document.body.appendChild(ta);ta.select();document.execCommand("copy");document.body.removeChild(ta);
      window.toast.success(lang==="tr"?"✓ Link kopyalandı":"✓ Link copied");
    }
  };

  const shareWhatsApp=(p)=>{
    const link=`${window.location.origin}${window.location.pathname}?invite=${p.invite_token}&team=${team.id}`;
    const msg=lang==="tr"
      ?`Merhaba ${p.name}, ${team.name} ekibine katılmak için bu linke tıkla:\n${link}`
      :`Hi ${p.name}, click this link to join ${team.name}:\n${link}`;
    window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`,"_blank");
  };

  // Birleşik liste — gerçek + phantom (linked olmayan)
  const allMembers=[
    ...teamMembers.map(m=>({...m,type:"real",displayName:m.name,_id:m.userId||m.user_id})),
    ...phantomMembers.filter(p=>!p.linked_user_id).map(p=>({...p,type:"phantom",displayName:p.name,_id:p.id}))
  ];

  return <div style={{...cSt(t),padding:"12px 14px",marginBottom:12,background:t.inBg}}>
    <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:8,letterSpacing:"0.05em",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
      <span>👥 {lang==="tr"?"EKİP ÜYELERİ":"TEAM MEMBERS"} ({allMembers.length})</span>
      <button onClick={()=>setShowAdd(true)} style={{...bSt("s",t),fontSize:10,padding:"4px 10px"}}>+ {lang==="tr"?"Üye":"Member"}</button>
    </div>
    {allMembers.length===0&&<div style={{fontSize:11,color:t.tm,padding:8,textAlign:"center"}}>{lang==="tr"?"Henüz ekip üyesi yok":"No members yet"}</div>}
    <div style={{display:"flex",flexDirection:"column",gap:6}}>
      {allMembers.map((m,i)=>{
        const isMe=m.type==="real"&&m._id===user?.userId;
        const isChef=m.type==="real"&&(m.role==="chef"||m.role==="head_chef"||m.role==="executive_chef");
        const isPhantom=m.type==="phantom";
        return <div key={m._id||i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 10px",background:t.bg,border:`1px solid ${isPhantom?t.tm:t.border}`,borderRadius:8,opacity:isPhantom?0.85:1}}>
          <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0,flex:1}}>
            <div style={{width:32,height:32,borderRadius:"50%",background:isChef?t.accent:(isPhantom?"transparent":t.tm),color:isPhantom?t.tm:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:12,flexShrink:0,border:isPhantom?`2px dashed ${t.tm}`:"none"}}>
              {(m.displayName||"?").charAt(0).toUpperCase()}
            </div>
            <div style={{minWidth:0,flex:1}}>
              <div style={{fontSize:13,fontWeight:600,color:t.text,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",display:"flex",alignItems:"center",gap:6}}>
                {m.displayName||"-"}
                {isMe&&<span style={{fontSize:10,color:t.tm,fontWeight:400}}>({lang==="tr"?"Sen":"You"})</span>}
                {isPhantom&&<span style={{fontSize:9,color:t.tm,fontWeight:400,background:t.bg2||"transparent",padding:"1px 6px",borderRadius:4,border:`1px solid ${t.tm}`}}>{lang==="tr"?"Kayıtsız":"Pending"}</span>}
              </div>
              <div style={{fontSize:10,color:t.tm,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                {isChef?"👑 ":""}
                {isPhantom?(m.position||m.department||(lang==="tr"?"Pozisyon belirsiz":"No position")):
                  (m.role==="chef"?(lang==="tr"?"Şef":"Chef"):
                   m.role==="head_chef"?(lang==="tr"?"Baş Şef":"Head Chef"):
                   m.role==="executive_chef"?(lang==="tr"?"Executive Şef":"Executive Chef"):
                   m.role==="sous_chef"?"Sous Chef":
                   m.role==="member"?(lang==="tr"?"Üye":"Member"):(m.role||""))}
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:6,alignItems:"center"}}>
            <div style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:t.tm}}>
              <span title={lang==="tr"?"Yıllık izin günü":"Annual leave days"}>🌴</span>
              <input type="number" min="0" max="60" defaultValue={m.annual_leave_total||getDefaultAnnualLeave(team?.country||"TR")} onBlur={e=>updateLeaveTotal(m,e.target.value)} style={{...iSt(t),width:48,fontSize:11,padding:"3px 4px",textAlign:"center"}}/>
            </div>
            {isPhantom&&<>
              <button onClick={()=>generateInviteToken(m)} title={lang==="tr"?"Davet linki":"Invite link"} style={{...bSt("s",t),fontSize:11,padding:"4px 8px"}}>🔗</button>
              <button onClick={()=>removePhantom(m)} title={lang==="tr"?"Sil":"Delete"} style={{background:"transparent",border:"none",color:t.danger,cursor:"pointer",fontSize:14,padding:"0 4px"}}>✕</button>
            </>}
          </div>
        </div>;
      })}
    </div>

    {/* Yeni Phantom Ekle Modal */}
    {showAdd&&<div onClick={e=>{if(e.target===e.currentTarget)setShowAdd(false);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:400,width:"100%"}}>
        <div style={{fontSize:16,fontWeight:700,color:t.text,marginBottom:6}}>+ {lang==="tr"?"Yeni Üye":"New Member"}</div>
        <div style={{fontSize:11,color:t.tm,marginBottom:12,lineHeight:1.4}}>{lang==="tr"?"Uygulamaya kayıtlı olmayan bir ekip üyesi ekle. Daha sonra davet linki gönderebilirsin.":"Add a member who isn't registered in the app yet. You can send an invite link later."}</div>
        <div style={{display:"grid",gap:8}}>
          <input value={newPhantom.name} onChange={e=>setNewPhantom({...newPhantom,name:e.target.value})} placeholder={lang==="tr"?"İsim Soyisim *":"Full name *"} style={{...iSt(t)}}/>
          <input value={newPhantom.position} onChange={e=>setNewPhantom({...newPhantom,position:e.target.value})} placeholder={lang==="tr"?"Pozisyon (Pastry Chef)":"Position"} style={{...iSt(t)}}/>
          <input value={newPhantom.department} onChange={e=>setNewPhantom({...newPhantom,department:e.target.value})} placeholder={lang==="tr"?"Departman (Pastane)":"Department"} style={{...iSt(t)}}/>
          <input value={newPhantom.email} onChange={e=>setNewPhantom({...newPhantom,email:e.target.value})} placeholder="email@example.com" style={{...iSt(t)}}/>
        </div>
        <div style={{display:"flex",gap:8,marginTop:14}}>
          <button onClick={()=>setShowAdd(false)} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
          <button onClick={addPhantom} style={{...bSt("p",t),flex:1}}>{lang==="tr"?"Ekle":"Add"}</button>
        </div>
      </div>
    </div>}

    {/* Davet Linki Modal */}
    {inviteFor&&<div onClick={e=>{if(e.target===e.currentTarget)setInviteFor(null);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:t.cardBg||t.bg,borderRadius:14,padding:20,maxWidth:420,width:"100%"}}>
        <div style={{fontSize:16,fontWeight:700,color:t.text,marginBottom:4}}>🔗 {lang==="tr"?"Davet Linki":"Invite Link"}</div>
        <div style={{fontSize:12,color:t.tm,marginBottom:14,lineHeight:1.4}}>
          <strong>{inviteFor.name}</strong> {lang==="tr"?"için davet linki hazır. Linke tıklayan kişi kayıt olduğunda otomatik olarak bu phantom hesabıyla eşleşir.":"link is ready. When clicked and registered, the user is auto-linked to this phantom."}
        </div>
        <div style={{background:t.bg2||t.bg,border:`1px solid ${t.border}`,borderRadius:8,padding:10,marginBottom:12,fontSize:11,wordBreak:"break-all",fontFamily:"monospace",color:t.text}}>
          {window.location.origin}{window.location.pathname}?invite={inviteFor.invite_token}&team={team.id}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>copyInviteLink(inviteFor)} style={{...bSt("p",t),flex:"1 1 100px"}}>📋 {lang==="tr"?"Kopyala":"Copy"}</button>
          <button onClick={()=>shareWhatsApp(inviteFor)} style={{...bSt("s",t),flex:"1 1 100px"}}>💬 WhatsApp</button>
          <button onClick={()=>setInviteFor(null)} style={{...bSt("s",t),flex:"1 1 100px"}}>{lang==="tr"?"Kapat":"Close"}</button>
        </div>
      </div>
    </div>}
  </div>;
};

const ChildTeamsSection=({teamId,t,lang})=>{
  const[children,setChildren]=useState([]);
  const[loading,setLoading]=useState(true);
  const[stats,setStats]=useState({});

  const loadChildren=async()=>{
    setLoading(true);
    try{
      const sb=initSupabase();if(!sb){setLoading(false);return;}
      const{data}=await sb.from("teams").select("id,name,invite_code").eq("parent_team_id",teamId);
      const list=data||[];
      setChildren(list);
      // Her alt ekip için stat çek
      const newStats={};
      await Promise.all(list.map(async(child)=>{
        const[stockRes,prodRes,membersRes]=await Promise.all([
          sb.from("stock").select("*").eq("team_id",child.id),
          sb.from("productions").select("*").eq("team_id",child.id),
          sb.from("team_members").select("id").eq("team_id",child.id),
        ]);
        const stock=Array.isArray(stockRes?.data?.data)?stockRes.data.data:[];
        const prod=Array.isArray(prodRes?.data?.data)?prodRes.data.data:[];
        const today=new Date().toDateString();
        const todayProd=prod.filter(p=>{const d=p.producedAt||p.created_at;return d&&new Date(d).toDateString()===today;});
        newStats[child.id]={
          stockCount:stock.length,
          totalProd:prod.length,
          todayProd:todayProd.length,
          memberCount:(membersRes?.data||[]).length,
        };
      }));
      setStats(newStats);
    }catch(e){console.warn("Alt ekipler yüklenemedi:",e);}
    setLoading(false);
  };

  useEffect(()=>{loadChildren();},[teamId]);

  if(loading)return <div style={{...cSt(t),padding:14,marginBottom:12}}>
    <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:6,letterSpacing:"0.05em"}}>🏢 {lang==="tr"?"ALT EKİPLER":"CHILD TEAMS"}</div>
    <div style={{fontSize:12,color:t.tm}}>⏳</div>
  </div>;

  if(!children.length)return <div style={{...cSt(t),padding:14,marginBottom:12,background:t.inBg}}>
    <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:6,letterSpacing:"0.05em"}}>🏢 {lang==="tr"?"ALT EKİPLER":"CHILD TEAMS"}</div>
    <div style={{fontSize:12,color:t.tm,lineHeight:1.5}}>{lang==="tr"?"Henüz alt ekip yok. Davet kodunuzu paylaşarak Manager/Pro kullanıcılarının size bağlanmasını sağlayın.":"No child teams yet. Share your invite code so Manager/Pro users can link to you."}</div>
  </div>;

  return <div style={{...cSt(t),padding:14,marginBottom:12}}>
    <div style={{fontSize:11,color:t.tm,fontWeight:700,marginBottom:10,letterSpacing:"0.05em",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span>🏢 {lang==="tr"?"ALT EKİPLER":"CHILD TEAMS"} ({children.length})</span>
      <button onClick={loadChildren} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:13}}>↻</button>
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      {children.map(child=>{
        const s=stats[child.id]||{};
        return <div key={child.id} style={{background:t.inBg,border:`1px solid ${t.border}`,borderRadius:10,padding:"10px 12px"}}>
          <div style={{fontSize:13,fontWeight:700,color:t.text,marginBottom:6}}>🍳 {child.name}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:11,color:t.tm}}>
            <div>👥 {s.memberCount||0} {lang==="tr"?"üye":"members"}</div>
            <div>📦 {s.stockCount||0} {lang==="tr"?"stok":"stock"}</div>
            <div>🍳 {s.totalProd||0} {lang==="tr"?"üretim":"productions"}</div>
            <div style={{color:s.todayProd>0?t.success:t.tm}}>📅 {s.todayProd||0} {lang==="tr"?"bugün":"today"}</div>
          </div>
        </div>;
      })}
    </div>
  </div>;
};


const HubTab=({team,user,t})=>{
  const lang=t.lang;
  const[children,setChildren]=useState([]);
  const[loading,setLoading]=useState(true);
  const[expandedId,setExpandedId]=useState(null);
  const[period,setPeriod]=useState("today"); // today, week, month
  const[allData,setAllData]=useState({}); // {teamId: {stock, productions, members}}

  const periodFilter=(items,dateField="producedAt")=>{
    const now=new Date();
    return items.filter(item=>{
      const d=item[dateField]||item.created_at;
      if(!d)return false;
      const dt=new Date(d);
      if(period==="today")return dt.toDateString()===now.toDateString();
      if(period==="week"){const w=new Date();w.setDate(w.getDate()-7);return dt>=w;}
      if(period==="month"){const m=new Date();m.setMonth(m.getMonth()-1);return dt>=m;}
      return true;
    });
  };

  const loadAll=async()=>{
    setLoading(true);
    try{
      const sb=initSupabase();if(!sb){setLoading(false);return;}
      const{data:list}=await sb.from("teams").select("id,name,invite_code").eq("parent_team_id",team.id);
      const teams=list||[];
      setChildren(teams);
      const data={};
      await Promise.all(teams.map(async(child)=>{
        const[stockRes,prodRes,membersRes]=await Promise.all([
          sb.from("stock").select("*").eq("team_id",child.id),
          sb.from("productions").select("*").eq("team_id",child.id),
          sb.from("team_members").select("user_id,position,role").eq("team_id",child.id),
        ]);
        const memberData=membersRes?.data||[];
        const uids=memberData.map(m=>m.user_id);
        let profileMap={};
        if(uids.length>0){
          const{data:profs}=await sb.from("profiles").select("id,full_name,email").in("id",uids);
          (profs||[]).forEach(p=>{profileMap[p.id]=p.full_name||p.email?.split("@")[0]||p.id;});
        }
        data[child.id]={
          stock:stockRes?.data||[],
          productions:prodRes?.data||[],
          members:memberData.map(m=>({...m,name:profileMap[m.user_id]||m.position||m.user_id})),
        };
      }));
      setAllData(data);
    }catch(e){console.warn("Hub load:",e);}
    setLoading(false);
  };

  useEffect(()=>{loadAll();},[team?.id]);

  // Toplam istatistikler
  const totals=useMemo(()=>{
    let totalStock=0,totalProd=0,totalFire=0,totalConsumed=0,totalActive=0,totalMembers=0;
    let totalCost=0,totalFireCost=0;
    Object.values(allData).forEach(d=>{
      totalStock+=d.stock.length;
      totalMembers+=d.members.length;
      const filtered=periodFilter(d.productions);
      filtered.forEach(p=>{
        totalProd++;
        if(p.status==="fire"||p.fired)totalFire++;
        else if(p.status==="consumed"||p.consumed===p.portions)totalConsumed++;
        else if(p.status==="active")totalActive++;
        if(p.cost)totalCost+=p.cost;
        if((p.status==="fire"||p.fired)&&p.cost)totalFireCost+=p.cost;
      });
    });
    return{totalStock,totalProd,totalFire,totalConsumed,totalActive,totalMembers,totalCost,totalFireCost};
  },[allData,period]);

  // Düşük stok uyarıları
  const lowStockAlerts=useMemo(()=>{
    const alerts=[];
    Object.entries(allData).forEach(([teamId,d])=>{
      const child=children.find(c=>c.id===teamId);
      d.stock.forEach(s=>{
        if(s.qty<=(s.low||100)){
          alerts.push({teamName:child?.name||"?",item:s.name,qty:s.qty,unit:s.unit||"g"});
        }
      });
    });
    return alerts.slice(0,10);
  },[allData,children]);

  // Yakın SKT
  const expiringSoon=useMemo(()=>{
    const items=[];
    const now=new Date();
    const week=new Date();week.setDate(week.getDate()+7);
    Object.entries(allData).forEach(([teamId,d])=>{
      const child=children.find(c=>c.id===teamId);
      d.stock.forEach(s=>{
        if(!s.skt)return;
        const exp=new Date(s.skt);
        if(exp<=week){
          items.push({teamName:child?.name||"?",item:s.name,skt:s.skt,daysLeft:Math.ceil((exp-now)/(24*3600*1000))});
        }
      });
    });
    return items.sort((a,b)=>a.daysLeft-b.daysLeft).slice(0,10);
  },[allData,children]);

  if(loading)return <div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>⏳ {lang==="tr"?"Yükleniyor...":"Loading..."}</div>;

  if(!team)return <div style={{...cSt(t),padding:24,textAlign:"center",color:t.tm}}>
    <div style={{fontSize:32,marginBottom:8}}>🏢</div>
    <div>{lang==="tr"?"Önce ekip oluşturun.":"Create a team first."}</div>
  </div>;

  if(!children.length)return <div style={{...cSt(t),padding:24,textAlign:"center"}}>
    <div style={{fontSize:32,marginBottom:8}}>🏢</div>
    <div style={{fontSize:14,color:t.text,fontWeight:700,marginBottom:6}}>{lang==="tr"?"Henüz alt ekip yok":"No child teams yet"}</div>
    <div style={{fontSize:12,color:t.tm,lineHeight:1.5}}>{lang==="tr"?"Ekibinizin davet kodunu Manager kullanıcılarla paylaşın. Onlar kendi ekiplerini sizin ekibinize bağladığında burada görünecekler.":"Share your team invite code with Manager users. They will appear here when they link their teams to yours."}</div>
    <div style={{marginTop:12,padding:"10px 14px",background:t.acB,borderRadius:10,display:"inline-block"}}>
      <div style={{fontSize:11,color:t.accent,fontWeight:700}}>{lang==="tr"?"DAVET KODU":"INVITE CODE"}</div>
      <div style={{fontSize:24,fontWeight:900,color:t.accent,letterSpacing:"0.2em"}}>{team.inviteCode||team.invite_code}</div>
    </div>
  </div>;

  return <div>
    {/* Başlık */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
      <h2 style={{fontSize:20,color:t.text,fontFamily:"'Fraunces',serif",margin:0}}>📊 {lang==="tr"?"Departman Raporu":"Department Report"}</h2>
      <button onClick={loadAll} style={{...bSt("s",t),fontSize:11,padding:"5px 10px"}}>↻</button>
    </div>

    {/* Periyot Filtre */}
    <div style={{display:"flex",gap:4,marginBottom:14,background:t.inBg,padding:3,borderRadius:10}}>
      {[["today",lang==="tr"?"Bugün":"Today"],["week",lang==="tr"?"Hafta":"Week"],["month",lang==="tr"?"Ay":"Month"]].map(([p,l])=>
        <button key={p} onClick={()=>setPeriod(p)} style={{flex:1,padding:"7px",borderRadius:8,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",background:period===p?t.card:"transparent",color:period===p?t.text:t.tm}}>{l}</button>
      )}
    </div>

    {/* Genel Özet Kartları */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
      <div style={{...cSt(t),padding:"12px 14px"}}>
        <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em"}}>{lang==="tr"?"DEPARTMAN":"DEPARTMENTS"}</div>
        <div style={{fontSize:24,fontWeight:800,color:t.accent}}>{children.length}</div>
        <div style={{fontSize:10,color:t.tm}}>👥 {totals.totalMembers} {lang==="tr"?"toplam üye":"total members"}</div>
      </div>
      <div style={{...cSt(t),padding:"12px 14px"}}>
        <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em"}}>{lang==="tr"?"ÜRETİM":"PRODUCTIONS"}</div>
        <div style={{fontSize:24,fontWeight:800,color:t.success}}>{totals.totalProd}</div>
        <div style={{fontSize:10,color:t.tm}}>✓ {totals.totalConsumed} · ⚡ {totals.totalActive}</div>
      </div>
      <div style={{...cSt(t),padding:"12px 14px"}}>
        <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em"}}>{lang==="tr"?"STOK":"STOCK"}</div>
        <div style={{fontSize:24,fontWeight:800,color:t.text}}>{totals.totalStock}</div>
        <div style={{fontSize:10,color:t.danger}}>⚠️ {lowStockAlerts.length} {lang==="tr"?"düşük":"low"}</div>
      </div>
      <div style={{...cSt(t),padding:"12px 14px"}}>
        <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em"}}>{lang==="tr"?"FİRE":"WASTE"}</div>
        <div style={{fontSize:24,fontWeight:800,color:t.danger}}>{totals.totalFire}</div>
        <div style={{fontSize:10,color:t.tm}}>{totals.totalProd>0?Math.round(totals.totalFire/totals.totalProd*100):0}% {lang==="tr"?"fire oranı":"waste rate"}</div>
      </div>
    </div>

    {/* Departmanlar */}
    <div style={{fontSize:11,color:t.tm,fontWeight:700,letterSpacing:"0.05em",marginBottom:8}}>🏢 {lang==="tr"?"DEPARTMANLAR":"DEPARTMENTS"}</div>
    <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:14}}>
      {children.map(child=>{
        const d=allData[child.id]||{stock:[],productions:[],members:[]};
        const filteredProd=periodFilter(d.productions);
        const fireCount=filteredProd.filter(p=>p.status==="fire"||p.fired).length;
        const lowStock=d.stock.filter(s=>s.qty<=(s.low||100)).length;
        const isExp=expandedId===child.id;
        return <div key={child.id} style={{...cSt(t)}}>
          <button onClick={()=>setExpandedId(isExp?null:child.id)} style={{width:"100%",padding:"12px 14px",background:"none",border:"none",cursor:"pointer",textAlign:"left",display:"flex",alignItems:"center",gap:10}}>
            <div style={{flex:1}}>
              <div style={{fontSize:14,fontWeight:700,color:t.text,marginBottom:4}}>🍳 {child.name}</div>
              <div style={{display:"flex",gap:10,fontSize:11,color:t.tm,flexWrap:"wrap"}}>
                <span>👥 {d.members.length}</span>
                <span>📦 {d.stock.length}</span>
                <span style={{color:t.success}}>🍳 {filteredProd.length}</span>
                {fireCount>0&&<span style={{color:t.danger}}>🗑 {fireCount}</span>}
                {lowStock>0&&<span style={{color:t.danger}}>⚠️ {lowStock}</span>}
              </div>
            </div>
            <span style={{color:t.tm,fontSize:14}}>{isExp?"▾":"▸"}</span>
          </button>
          {isExp&&<div style={{padding:"0 14px 14px",borderTop:`1px solid ${t.border}`}}>
            {/* Üyeler */}
            {d.members.length>0&&<div style={{marginTop:12}}>
              <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em",marginBottom:6}}>👥 {lang==="tr"?"ÜYELER":"MEMBERS"}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                {d.members.map(m=><span key={m.user_id} style={{fontSize:11,background:t.inBg,padding:"3px 8px",borderRadius:6,color:t.text}}>{m.role==="chef"?"👑 ":""}{m.name}</span>)}
              </div>
            </div>}
            {/* Son Üretimler */}
            {filteredProd.length>0&&<div style={{marginTop:12}}>
              <div style={{fontSize:10,color:t.tm,fontWeight:700,letterSpacing:"0.05em",marginBottom:6}}>🍱 {lang==="tr"?"ÜRETİMLER":"PRODUCTIONS"} ({filteredProd.length})</div>
              <div style={{display:"flex",flexDirection:"column",gap:4,maxHeight:200,overflowY:"auto"}}>
                {filteredProd.slice(0,8).map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,padding:"4px 6px",background:t.inBg,borderRadius:6}}>
                  <span style={{color:t.text,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.recipeName||"?"}</span>
                  <span style={{color:t.tm,fontSize:10,marginLeft:8}}>{p.portions||0}p</span>
                  <span style={{color:p.status==="fire"||p.fired?t.danger:p.status==="consumed"?t.success:t.tm,fontSize:10,marginLeft:8}}>
                    {p.status==="fire"||p.fired?"🗑":p.status==="consumed"?"✓":"⚡"}
                  </span>
                </div>)}
              </div>
            </div>}
            {/* Düşük Stok */}
            {lowStock>0&&<div style={{marginTop:12}}>
              <div style={{fontSize:10,color:t.danger,fontWeight:700,letterSpacing:"0.05em",marginBottom:6}}>⚠️ {lang==="tr"?"DÜŞÜK STOK":"LOW STOCK"}</div>
              <div style={{fontSize:11,color:t.tm}}>{d.stock.filter(s=>s.qty<=(s.low||100)).slice(0,5).map(s=>s.name).join(" · ")}</div>
            </div>}
          </div>}
        </div>;
      })}
    </div>

    {/* Cross-departman uyarılar */}
    {(lowStockAlerts.length>0||expiringSoon.length>0)&&<div style={{marginTop:14}}>
      <div style={{fontSize:11,color:t.tm,fontWeight:700,letterSpacing:"0.05em",marginBottom:8}}>🚨 {lang==="tr"?"GENEL UYARILAR":"GLOBAL ALERTS"}</div>
      {lowStockAlerts.length>0&&<div style={{...cSt(t),padding:"10px 14px",marginBottom:8,borderLeft:`3px solid ${t.danger}`}}>
        <div style={{fontSize:12,fontWeight:700,color:t.danger,marginBottom:4}}>⚠️ {lang==="tr"?"Düşük Stok":"Low Stock"} ({lowStockAlerts.length})</div>
        {lowStockAlerts.slice(0,5).map((a,i)=><div key={i} style={{fontSize:11,color:t.text,padding:"2px 0"}}>
          <span style={{color:t.tm}}>{a.teamName}</span> · {a.item} ({a.qty}{a.unit})
        </div>)}
      </div>}
      {expiringSoon.length>0&&<div style={{...cSt(t),padding:"10px 14px",borderLeft:`3px solid #f59e0b`}}>
        <div style={{fontSize:12,fontWeight:700,color:"#f59e0b",marginBottom:4}}>📅 {lang==="tr"?"Yakında SKT":"Expiring Soon"} ({expiringSoon.length})</div>
        {expiringSoon.slice(0,5).map((a,i)=><div key={i} style={{fontSize:11,color:t.text,padding:"2px 0"}}>
          <span style={{color:t.tm}}>{a.teamName}</span> · {a.item} ({a.daysLeft<0?(lang==="tr"?"süresi geçti":"expired"):a.daysLeft+(lang==="tr"?" gün":" days")})
        </div>)}
      </div>}
    </div>}
  </div>;
};


const KanbanTab=({team,teamMembers,user,t,profile,isManager=false,isPro=false})=>{
  const[cards,setCards]=useState([]);
  const[loading,setLoading]=useState(true);
  const[showNew,setShowNew]=useState(false);
  const[filter,setFilter]=useState("all");
  const[view,setView]=useState("kanban");
  const[dragCard,setDragCard]=useState(null);
  const[editCard,setEditCard]=useState(null);
  const[detailCard,setDetailCard]=useState(null);
  const[showProgress,setShowProgress]=useState(false);
  const[progressCard,setProgressCard]=useState(null);
  const[progressNote,setProgressNote]=useState("");
  const[progressUploading,setProgressUploading]=useState(false);
  const[newCard,setNewCard]=useState({text:"",col:"todo",assignedTo:"",priority:"medium",dueDate:"",visibility:"team",colorLabel:"normal",cardType:"task",recurrence:"",meetingAt:"",meetingAttendees:[]});
  const[newComment,setNewComment]=useState("");
  const lang=t.lang;
  const myUid=user?.userId||"";
  const myName=(profile&&profile.fullName)||user?.name||"";
  const canManage=isManager||isPro;
  const COLS=[
    {id:"todo",label:{tr:"Bekleyen",en:"To Do"},icon:"📋",color:"#6366f1"},
    {id:"doing",label:{tr:"Devam",en:"In Progress"},icon:"⚡",color:"#f59e0b"},
    {id:"done",label:{tr:"Tamamlandı",en:"Done"},icon:"✅",color:"#10b981"}
  ];
  const PRIOS={
    high:{color:"#dc2626",label:{tr:"Yüksek",en:"High"},icon:"🔴"},
    medium:{color:"#f59e0b",label:{tr:"Orta",en:"Medium"},icon:"🟡"},
    low:{color:"#16a34a",label:{tr:"Düşük",en:"Low"},icon:"🟢"}
  };
  const COLOR_LABELS={
    normal:{color:null,label:{tr:"Normal",en:"Normal"}},
    urgent:{color:"#dc2626",label:{tr:"Acil",en:"Urgent"}},
    today:{color:"#f59e0b",label:{tr:"Bugün",en:"Today"}},
    thisweek:{color:"#3b82f6",label:{tr:"Bu Hafta",en:"This Week"}}
  };
  const RECURRENCE=[
    {id:"",label:{tr:"Tekrarsız",en:"No Repeat"}},
    {id:"daily",label:{tr:"Her Gün",en:"Daily"}},
    {id:"weekly",label:{tr:"Her Hafta",en:"Weekly"}},
    {id:"shift",label:{tr:"Her Vardiya",en:"Every Shift"}}
  ];
  const colLabel=c=>c.label[lang]||c.label.en;

  const loadCards=async()=>{
    setLoading(true);
    try{
      const sb=initSupabase();
      if(sb&&team?.id){
        const{data}=await sb.from("kanban_cards").select("*").eq("team_id",team.id).order("created_at",{ascending:false});
        if(data)setCards(data);
      }
    }catch(e){console.warn("Kanban load:",e);}
    setLoading(false);
  };

  useEffect(()=>{loadCards();},[team?.id]);

  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    const ch=sb.channel(`kanban-${team.id}`)
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"kanban_cards",filter:`team_id=eq.${team.id}`},(payload)=>{
        setCards(p=>p.find(c=>c.id===payload.new.id)?p:[payload.new,...p]);
      })
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"kanban_cards",filter:`team_id=eq.${team.id}`},(payload)=>{
        setCards(p=>p.map(c=>c.id===payload.new.id?payload.new:c));
      })
      .on("postgres_changes",{event:"DELETE",schema:"public",table:"kanban_cards",filter:`team_id=eq.${team.id}`},(payload)=>{
        setCards(p=>p.filter(c=>c.id!==payload.old.id));
      })
      .subscribe();
    return()=>{sb.removeChannel(ch);};
  },[team?.id]);

  // Tekrarlayan görev kontrolü
  useEffect(()=>{
    if(!cards.length)return;
    const today=new Date().toISOString().slice(0,10);
    cards.forEach(async card=>{
      if(!card.recurrence||card.col!=="done")return;
      const lastDone=new Date(card.updated_at||card.created_at);
      const now=new Date();
      let shouldReset=false;
      if(card.recurrence==="daily"&&lastDone.toDateString()!==now.toDateString())shouldReset=true;
      if(card.recurrence==="weekly"&&(now-lastDone)>7*24*3600*1000)shouldReset=true;
      if(shouldReset){
        const sb=initSupabase();if(!sb)return;
        await sb.from("kanban_cards").update({col:"todo"}).eq("id",card.id);
      }
    });
  },[cards]);

  const addCard=async()=>{
    if(!newCard.text.trim())return;
    const sb=initSupabase();if(!sb)return;
    const attendees=newCard.cardType==="meeting"?newCard.meetingAttendees:[];
    await sb.from("kanban_cards").insert({
      team_id:team.id,
      text:newCard.text.trim(),
      col:newCard.col,
      assigned_to:newCard.assignedTo||null,
      created_by:myUid,
      priority:newCard.priority,
      due_date:newCard.dueDate||null,
      visibility:newCard.visibility,
      color_label:newCard.colorLabel||"normal",
      card_type:newCard.cardType||"task",
      recurrence:newCard.recurrence||null,
      meeting_at:newCard.meetingAt||null,
      meeting_attendees:attendees,
      checklist:[],
      comments:[]
    });
    setNewCard({text:"",col:"todo",assignedTo:"",priority:"medium",dueDate:"",visibility:"team",colorLabel:"normal",cardType:"task",recurrence:"",meetingAt:"",meetingAttendees:[]});
    setShowNew(false);
  };

  const moveCard=async(id,toCol)=>{
    const sb=initSupabase();if(!sb)return;
    setCards(p=>p.map(c=>c.id===id?{...c,col:toCol}:c));
    await sb.from("kanban_cards").update({col:toCol,updated_at:new Date().toISOString()}).eq("id",id);
  };

  const deleteCard=async(id,card)=>{
    if(!canManage&&card.created_by!==myUid&&card.assigned_to!==myUid)return;
    const sb=initSupabase();if(!sb)return;
    setCards(p=>p.filter(c=>c.id!==id));
    await sb.from("kanban_cards").delete().eq("id",id);
    if(detailCard?.id===id)setDetailCard(null);
    if(editCard?.id===id)setEditCard(null);
  };

  const updateCard=async(id,updates)=>{
    const sb=initSupabase();if(!sb)return;
    setCards(p=>p.map(c=>c.id===id?{...c,...updates}:c));
    if(detailCard?.id===id)setDetailCard(d=>({...d,...updates}));
    await sb.from("kanban_cards").update({...updates,updated_at:new Date().toISOString()}).eq("id",id);
    setEditCard(null);
  };

  // Checklist
  const toggleCheckItem=async(card,idx)=>{
    const newList=(card.checklist||[]).map((item,i)=>i===idx?{...item,done:!item.done}:item);
    await updateCard(card.id,{checklist:newList});
  };
  const addCheckItem=async(card,text)=>{
    if(!text.trim())return;
    const newList=[...(card.checklist||[]),{id:Date.now(),text:text.trim(),done:false}];
    await updateCard(card.id,{checklist:newList});
  };
  const deleteCheckItem=async(card,idx)=>{
    const newList=(card.checklist||[]).filter((_,i)=>i!==idx);
    await updateCard(card.id,{checklist:newList});
  };

  // Yorum
  const addComment=async(card)=>{
    if(!newComment.trim())return;
    const newComments=[...(card.comments||[]),{id:Date.now(),text:newComment.trim(),by:myUid,name:myName,at:new Date().toISOString()}];
    await updateCard(card.id,{comments:newComments});
    setNewComment("");
  };
  const deleteComment=async(card,idx)=>{
    const comment=(card.comments||[])[idx];
    if(comment.by!==myUid&&!canManage)return;
    const newComments=(card.comments||[]).filter((_,i)=>i!==idx);
    await updateCard(card.id,{comments:newComments});
  };

  const getMemberName=uid=>{const m=(teamMembers||[]).find(m=>(m.userId||m.user_id)===uid);return m?.name||"?";};

  const filteredCards=cards.filter(c=>{
    if(filter==="mine")return c.created_by===myUid||c.assigned_to===myUid;
    if(filter==="team")return c.visibility==="team";
    if(filter==="personal")return c.visibility==="personal"&&c.created_by===myUid;
    if(filter==="meeting")return c.card_type==="meeting";
    return true;
  });

  const printKanban=()=>{
    const ch=COLS.map(col=>{const cc=filteredCards.filter(c=>c.col===col.id);return `<div class="col"><div class="ch" style="color:${col.color}">${col.icon} ${colLabel(col)} (${cc.length})</div>${cc.map(c=>`<div class="card">${c.card_type==="meeting"?"📅 ":""}${c.text}${c.assigned_to?`<div class="meta">👤 ${getMemberName(c.assigned_to)}</div>`:""}</div>`).join("")}</div>`;}).join("");
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Arial;margin:20px;font-size:12px}.board{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}.col{background:#f8f8f8;border-radius:8px;padding:10px}.ch{font-weight:700;font-size:13px;margin-bottom:8px;padding-bottom:4px;border-bottom:1px solid #ddd}.card{background:#fff;border-radius:6px;padding:8px;margin-bottom:6px;border:1px solid #e5e7eb}.meta{font-size:10px;color:#888}</style></head><body><h3>📋 ${lang==="tr"?"Görevler":"Tasks"} — ${team?.name||""}</h3><div class="board">${ch}</div><script>window.onload=()=>window.print()<\/script></body></html>`;
    const w=window.open("","_blank");if(w){w.document.write(html);w.document.close();}
  };

  const shareKanban=()=>{
    const text=COLS.map(col=>{const cc=filteredCards.filter(c=>c.col===col.id);if(!cc.length)return null;return `${col.icon} ${colLabel(col)}:\n${cc.map(c=>`• ${c.text}${c.assigned_to?" ("+getMemberName(c.assigned_to)+")":""}`).join("\n")}`;}).filter(Boolean).join("\n\n");
    if(navigator.share){navigator.share({title:team?.name||"Görevler",text});}
    else{const ta=document.createElement("textarea");ta.value=text;ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.focus();ta.select();try{document.execCommand("copy");}catch{}document.body.removeChild(ta);}
  };

  if(loading)return <div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}>⏳</div>;

  // ── Detail Modal ──
  const DetailModal=({card,onClose})=>{
    const[checkInput,setCheckInput]=useState("");
    const prio=PRIOS[card.priority||"medium"];
    const col=COLS.find(c=>c.id===card.col)||COLS[0];
    const overdue=card.due_date&&card.col!=="done"&&new Date(card.due_date)<new Date();
    const checkDone=(card.checklist||[]).filter(i=>i.done).length;
    const checkTotal=(card.checklist||[]).length;
    const canEdit=canManage||card.created_by===myUid;
    const isAssigned=card.assigned_to===myUid;
    return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{...cSt(t),width:"100%",maxWidth:520,borderRadius:"20px 20px 0 0",padding:"20px 18px",maxHeight:"90vh",overflowY:"auto"}}>
        {/* Başlık */}
        <div style={{display:"flex",gap:8,alignItems:"flex-start",marginBottom:14}}>
          {card.card_type==="meeting"&&<span style={{fontSize:24}}>📅</span>}
          <div style={{flex:1}}>
            <div style={{fontSize:16,fontWeight:700,color:t.text,lineHeight:1.4}}>{card.text}</div>
            <div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>
              <span style={{fontSize:10,background:col.color+"20",color:col.color,padding:"2px 7px",borderRadius:6}}>{col.icon} {colLabel(col)}</span>
              <span style={{fontSize:10,background:prio.color+"20",color:prio.color,padding:"2px 7px",borderRadius:6}}>{prio.icon} {prio.label[lang]}</span>
              {card.color_label&&card.color_label!=="normal"&&<span style={{fontSize:10,background:COLOR_LABELS[card.color_label]?.color+"20",color:COLOR_LABELS[card.color_label]?.color,padding:"2px 7px",borderRadius:6}}>{COLOR_LABELS[card.color_label]?.label[lang]}</span>}
              {card.recurrence&&<span style={{fontSize:10,color:t.tm,background:t.inBg,padding:"2px 7px",borderRadius:6}}>🔄 {RECURRENCE.find(r=>r.id===card.recurrence)?.label[lang]}</span>}
            </div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:22,color:t.tm,cursor:"pointer",padding:"0 4px"}}>×</button>
        </div>

        {/* Meta */}
        <div style={{...cSt(t),padding:"10px 12px",marginBottom:14,background:t.inBg,borderRadius:10}}>
          {card.assigned_to&&<div style={{fontSize:13,color:t.text,marginBottom:4}}>👤 {getMemberName(card.assigned_to)}</div>}
          {card.due_date&&<div style={{fontSize:13,color:overdue?t.danger:t.tm,marginBottom:4}}>📅 {card.due_date}{overdue?" ⚠️":""}</div>}
          {card.card_type==="meeting"&&card.meeting_at&&<div style={{fontSize:13,color:t.accent,marginBottom:4}}>🕐 {new Date(card.meeting_at).toLocaleString()}</div>}
          {card.card_type==="meeting"&&(card.meeting_attendees||[]).length>0&&<div style={{fontSize:12,color:t.tm}}>👥 {(card.meeting_attendees||[]).map(uid=>getMemberName(uid)).join(", ")}</div>}
          <div style={{fontSize:11,color:t.tm,marginTop:4}}>✍️ {getMemberName(card.created_by)} · {new Date(card.created_at).toLocaleDateString()}</div>
        </div>

        {/* Checklist */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,fontWeight:700,color:t.tm,marginBottom:8,display:"flex",justifyContent:"space-between"}}>
            <span>☑️ {lang==="tr"?"Kontrol Listesi":"Checklist"}</span>
            {checkTotal>0&&<span style={{color:t.accent}}>{checkDone}/{checkTotal}</span>}
          </div>
          {checkTotal>0&&<div style={{height:4,background:t.inBg,borderRadius:4,marginBottom:10}}>
            <div style={{height:4,background:t.accent,borderRadius:4,width:`${checkDone/checkTotal*100}%`,transition:"width 0.3s"}}/>
          </div>}
          {(card.checklist||[]).map((item,i)=><div key={item.id||i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:`1px solid ${t.border}`}}>
            <button onClick={()=>toggleCheckItem(card,i)} style={{width:20,height:20,borderRadius:5,border:`2px solid ${item.done?t.accent:t.inBo}`,background:item.done?t.accent:"transparent",cursor:"pointer",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:11}}>{item.done?"✓":""}</button>
            <span style={{flex:1,fontSize:13,color:t.text,textDecoration:item.done?"line-through":"none",opacity:item.done?0.6:1}}>{item.text}</span>
            {canEdit&&<button onClick={()=>deleteCheckItem(card,i)} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:14}}>×</button>}
            {isAssigned&&!item.done&&<button onClick={()=>toggleCheckItem(card,i)} style={{fontSize:10,background:t.accent+"22",color:t.accent,border:"none",borderRadius:4,padding:"2px 6px",cursor:"pointer"}}>✓</button>}
          </div>)}
          {canEdit&&<div style={{display:"flex",gap:6,marginTop:8}}>
            <input style={{...iSt(t),flex:1,fontSize:13}} placeholder={lang==="tr"?"Madde ekle...":"Add item..."} value={checkInput} onChange={e=>setCheckInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){addCheckItem(card,checkInput);setCheckInput("");}}}/>
            <button onClick={()=>{addCheckItem(card,checkInput);setCheckInput("");}} style={{...bSt("p",t),padding:"8px 12px",fontSize:12}}>+</button>
          </div>}
        </div>

        {/* Yorumlar */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,fontWeight:700,color:t.tm,marginBottom:8}}>💬 {lang==="tr"?"Yorumlar":"Comments"} {(card.comments||[]).length>0&&`(${card.comments.length})`}</div>
          {(card.comments||[]).map((c,i)=><div key={c.id||i} style={{...cSt(t),padding:"8px 10px",marginBottom:6,background:t.inBg,borderRadius:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
              <span style={{fontSize:11,fontWeight:600,color:t.accent}}>{c.name||getMemberName(c.by)}</span>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <span style={{fontSize:10,color:t.tm}}>{new Date(c.at).toLocaleString()}</span>
                {(c.by===myUid||canManage)&&<button onClick={()=>deleteComment(card,i)} style={{background:"none",border:"none",color:t.tm,cursor:"pointer",fontSize:12}}>×</button>}
              </div>
            </div>
            {c.isProgress&&<span style={{fontSize:10,background:t.accent+"22",color:t.accent,padding:"2px 6px",borderRadius:4,marginBottom:4,display:"inline-block"}}>📝 {lang==="tr"?"İlerleme":"Progress"}</span>}
            {c.text&&<div style={{fontSize:13,color:t.text,lineHeight:1.5}}>{c.text}</div>}
            {c.attachment&&<div style={{marginTop:6}}>
              {isImage(c.attachment.ext)?
                <img src={c.attachment.url} style={{maxWidth:"100%",maxHeight:200,borderRadius:8,display:"block"}} alt={c.attachment.name} onClick={()=>window.open(c.attachment.url,"_blank")}/>:
                <a href={c.attachment.url} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:6,padding:"6px 10px",background:t.bg,borderRadius:6,fontSize:12,color:t.accent,textDecoration:"none"}}>
                  {isPDF(c.attachment.ext)?"📄":"📎"} {c.attachment.name}
                </a>
              }
            </div>}
          </div>)}
          <div style={{display:"flex",gap:6,marginTop:8}}>
            <input style={{...iSt(t),flex:1,fontSize:13}} placeholder={lang==="tr"?"Yorum yaz...":"Add comment..."} value={newComment} onChange={e=>setNewComment(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){addComment(card);}}}/>
            <button onClick={()=>addComment(card)} style={{...bSt("p",t),padding:"8px 12px",fontSize:12}}>↑</button>
          </div>
        </div>

        {/* Aksiyonlar */}
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {canEdit&&<button onClick={()=>{setEditCard(card);onClose();}} style={{...bSt("s",t),flex:1,fontSize:13}}>✏️ {lang==="tr"?"Düzenle":"Edit"}</button>}
          {isAssigned&&!canEdit&&<button onClick={()=>{setProgressCard(card);setShowProgress(true);onClose();}} style={{...bSt("p",t),flex:1,fontSize:13}}>📝 {lang==="tr"?"İlerleme Ekle":"Add Progress"}</button>}
          {COLS.filter(c=>c.id!==card.col).map(c=><button key={c.id} onClick={()=>{moveCard(card.id,c.id);onClose();}} style={{...bSt("s",t),flex:1,fontSize:12}}>{c.icon} {colLabel(c)}</button>)}
          {(canManage||card.created_by===myUid)&&<button onClick={()=>deleteCard(card.id,card)} style={{...bSt("s",t),flex:1,fontSize:13,color:t.danger}}>🗑</button>}
        </div>
      </div>
    </div>;
  };

  // İlerleme modalı
  const ProgressModal=()=>{
    if(!showProgress||!progressCard)return null;
    const fileRef=React.useRef(null);
    const addProgress=async()=>{
      if(!progressNote.trim()&&!progressCard._pendingFile)return;
      setProgressUploading(true);
      try{
        let attachment=null;
        if(progressCard._pendingFile){
          const f=progressCard._pendingFile;
          const uploaded=await uploadFile(f,team.id,"kanban");
          attachment={url:uploaded.url,path:uploaded.path,name:uploaded.name,type:uploaded.type,ext:uploaded.ext};
        }
        const comment={
          id:Date.now().toString(),
          by:myUid,
          name:myName,
          text:progressNote.trim()||"",
          at:new Date().toISOString(),
          isProgress:true,
          attachment
        };
        const updated={...progressCard,comments:[...(progressCard.comments||[]),comment]};
        await updateCard(progressCard.id,{comments:updated.comments});
        setCards(p=>p.map(c=>c.id===progressCard.id?updated:c));
        setProgressNote("");
        setProgressCard(c=>({...c,_pendingFile:null}));
        setShowProgress(false);
        window.toast.success(lang==="tr"?"İlerleme eklendi":"Progress added");
      }catch(e){window.toast.error(e.message);}
      setProgressUploading(false);
    };
    return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",zIndex:400,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget){setShowProgress(false);}}}>
      <div style={{...cSt(t),width:"100%",maxWidth:520,borderRadius:"20px 20px 0 0",padding:"20px 18px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <strong style={{fontSize:16,color:t.text}}>📝 {lang==="tr"?"İlerleme Ekle":"Add Progress"}</strong>
          <button onClick={()=>setShowProgress(false)} style={{background:"none",border:"none",fontSize:22,color:t.tm,cursor:"pointer"}}>×</button>
        </div>
        <div style={{fontSize:13,color:t.tm,marginBottom:12,padding:"8px 12px",background:t.inBg,borderRadius:8}}>
          📋 {progressCard.text}
        </div>
        <textarea
          style={{...iSt(t),minHeight:80,resize:"none",marginBottom:10}}
          placeholder={lang==="tr"?"Not yaz... (fotoğraf veya dosya da ekleyebilirsin)":"Add note... (you can also attach photo or file)"}
          value={progressNote}
          onChange={e=>setProgressNote(e.target.value)}
        />
        {progressCard._pendingFile&&<div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:t.inBg,borderRadius:8,marginBottom:10}}>
          {isImage(progressCard._pendingFile.name.split(".").pop())?
            <img src={URL.createObjectURL(progressCard._pendingFile)} style={{width:48,height:48,objectFit:"cover",borderRadius:6}} alt=""/>:
            <span style={{fontSize:24}}>{isPDF(progressCard._pendingFile.name.split(".").pop())?"📄":"📎"}</span>
          }
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:12,fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{progressCard._pendingFile.name}</div>
            <div style={{fontSize:10,color:t.tm}}>{(progressCard._pendingFile.size/1024).toFixed(0)}KB</div>
          </div>
          <button onClick={()=>setProgressCard(c=>({...c,_pendingFile:null}))} style={{background:"none",border:"none",color:t.danger,cursor:"pointer",fontSize:16}}>×</button>
        </div>}
        <input ref={fileRef} type="file" accept="image/*,video/*,.pdf,.xlsx,.xls,.docx,.txt" style={{display:"none"}} onChange={e=>{
          const f=e.target.files?.[0];
          if(f){
            if(f.size>50*1024*1024){window.toast.error(lang==="tr"?"Dosya 50MB'dan büyük":"File too large (50MB max)");return;}
            setProgressCard(c=>({...c,_pendingFile:f}));
          }
          e.target.value="";
        }}/>
        <div style={{display:"flex",gap:8,marginBottom:12}}>
          <button onClick={()=>fileRef.current?.click()} style={{...bSt("s",t),flex:1,fontSize:13}}>📷 {lang==="tr"?"Fotoğraf/Dosya":"Photo/File"}</button>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setShowProgress(false)} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
          <button onClick={addProgress} disabled={progressUploading||(!progressNote.trim()&&!progressCard._pendingFile)} style={{...bSt("p",t),flex:2,opacity:(progressUploading||(!progressNote.trim()&&!progressCard._pendingFile))?0.5:1}}>
            {progressUploading?"⏳ ":""}
            {lang==="tr"?"Kaydet":"Save"}
          </button>
        </div>
      </div>
    </div>;
  };

  return <div style={{maxWidth:"100%"}}>
    {showProgress&&<ProgressModal/>}
    {/* Başlık */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <h2 style={{fontSize:20,color:t.text,fontFamily:"'Fraunces',serif",margin:0}}>📋 {lang==="tr"?"Görevler":"Tasks"}</h2>
      <div style={{display:"flex",gap:6}}>
        <button onClick={shareKanban} style={{...bSt("s",t),fontSize:11,padding:"5px 8px"}}>📤</button>
        <button onClick={printKanban} style={{...bSt("s",t),fontSize:11,padding:"5px 8px"}}>🖨</button>
        <button onClick={()=>setView(v=>v==="kanban"?"list":"kanban")} style={{...bSt("s",t),fontSize:11,padding:"5px 8px"}}>{view==="kanban"?"☰":"▦"}</button>
        <button onClick={()=>setShowNew(s=>!s)} style={{...bSt("p",t),fontSize:12}}>+</button>
      </div>
    </div>

    {/* Filtre */}
    <div style={{display:"flex",gap:4,marginBottom:12,background:t.inBg,padding:3,borderRadius:10,overflowX:"auto"}}>
      {[["all",lang==="tr"?"Tümü":"All"],["mine",lang==="tr"?"Benimkiler":"Mine"],["team",lang==="tr"?"Ekip":"Team"],["personal",lang==="tr"?"Kişisel":"Personal"],["meeting","📅"]].map(([f,l])=>
        <button key={f} onClick={()=>setFilter(f)} style={{flex:"0 0 auto",padding:"6px 10px",borderRadius:8,fontSize:11,fontWeight:600,border:"none",cursor:"pointer",background:filter===f?t.card:"transparent",color:filter===f?t.text:t.tm}}>{l}</button>
      )}
    </div>

    {/* Yeni Kart Formu */}
    {showNew&&<div style={{...cSt(t),padding:14,marginBottom:12}}>
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["task",lang==="tr"?"📋 Görev":"📋 Task"],["meeting","📅 "+(lang==="tr"?"Toplantı":"Meeting")]].map(([type,label])=>
          <button key={type} onClick={()=>setNewCard(c=>({...c,cardType:type}))} style={{flex:1,padding:"7px",borderRadius:8,fontSize:12,fontWeight:600,border:`1px solid ${newCard.cardType===type?t.accent:t.border}`,background:newCard.cardType===type?t.accent+"22":t.inBg,color:newCard.cardType===type?t.accent:t.text,cursor:"pointer"}}>{label}</button>
        )}
      </div>
      <textarea style={{...iSt(t),minHeight:56,resize:"none",marginBottom:10,fontSize:14}}
        placeholder={newCard.cardType==="meeting"?(lang==="tr"?"Toplantı konusu...":"Meeting subject..."):(lang==="tr"?"Görev açıklaması...":"Task description...")}
        value={newCard.text} onChange={e=>setNewCard(c=>({...c,text:e.target.value}))}
        onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();addCard();}}}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <select style={iSt(t)} value={newCard.col} onChange={e=>setNewCard(c=>({...c,col:e.target.value}))}>
          {COLS.map(c=><option key={c.id} value={c.id}>{c.icon} {colLabel(c)}</option>)}
        </select>
        <select style={iSt(t)} value={newCard.priority} onChange={e=>setNewCard(c=>({...c,priority:e.target.value}))}>
          {Object.entries(PRIOS).map(([k,v])=><option key={k} value={k}>{v.icon} {v.label[lang]||v.label.en}</option>)}
        </select>
        <select style={iSt(t)} value={newCard.colorLabel} onChange={e=>setNewCard(c=>({...c,colorLabel:e.target.value}))}>
          {Object.entries(COLOR_LABELS).map(([k,v])=><option key={k} value={k}>{v.label[lang]||v.label.en}</option>)}
        </select>
        <select style={iSt(t)} value={newCard.recurrence} onChange={e=>setNewCard(c=>({...c,recurrence:e.target.value}))}>
          {RECURRENCE.map(r=><option key={r.id} value={r.id}>{r.label[lang]||r.label.en}</option>)}
        </select>
        <select style={iSt(t)} value={newCard.visibility} onChange={e=>setNewCard(c=>({...c,visibility:e.target.value}))}>
          <option value="team">👥 {lang==="tr"?"Ekip":"Team"}</option>
          <option value="personal">🔒 {lang==="tr"?"Kişisel":"Personal"}</option>
        </select>
        {canManage&&<select style={iSt(t)} value={newCard.assignedTo} onChange={e=>setNewCard(c=>({...c,assignedTo:e.target.value}))}>
          <option value="">{lang==="tr"?"Kişiye Ata":"Assign to..."}</option>
          {(teamMembers||[]).map(m=><option key={m.userId||m.user_id} value={m.userId||m.user_id}>{m.name}</option>)}
        </select>}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:8}}>
        <input type="date" style={{...iSt(t),flex:1,fontSize:12}} value={newCard.dueDate} onChange={e=>setNewCard(c=>({...c,dueDate:e.target.value}))}/>
        {newCard.cardType==="meeting"&&<input type="datetime-local" style={{...iSt(t),flex:1,fontSize:12}} value={newCard.meetingAt} onChange={e=>setNewCard(c=>({...c,meetingAt:e.target.value}))}/>}
      </div>
      {newCard.cardType==="meeting"&&canManage&&<div style={{marginBottom:8}}>
        <div style={{fontSize:11,color:t.tm,marginBottom:4}}>{lang==="tr"?"Katılımcılar":"Attendees"}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
          {(teamMembers||[]).map(m=>{
            const uid=m.userId||m.user_id;
            const sel=newCard.meetingAttendees.includes(uid);
            return <button key={uid} onClick={()=>setNewCard(c=>({...c,meetingAttendees:sel?c.meetingAttendees.filter(id=>id!==uid):[...c.meetingAttendees,uid]}))} style={{padding:"4px 8px",borderRadius:6,fontSize:11,border:`1px solid ${sel?t.accent:t.border}`,background:sel?t.accent+"22":t.inBg,color:sel?t.accent:t.text,cursor:"pointer"}}>{m.name}</button>;
          })}
        </div>
      </div>}
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>{setShowNew(false);setNewCard({text:"",col:"todo",assignedTo:"",priority:"medium",dueDate:"",visibility:"team",colorLabel:"normal",cardType:"task",recurrence:"",meetingAt:"",meetingAttendees:[]});}} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
        <button onClick={addCard} disabled={!newCard.text.trim()} style={{...bSt("p",t),flex:2,opacity:newCard.text.trim()?1:0.5}}>{lang==="tr"?"Ekle":"Add"}</button>
      </div>
    </div>}

    {/* Edit Modal */}
    {editCard&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{...cSt(t),padding:20,width:"100%",maxWidth:420,borderRadius:16}}>
        <h3 style={{margin:"0 0 14px",color:t.text,fontSize:16}}>{lang==="tr"?"Kartı Düzenle":"Edit Card"}</h3>
        <textarea style={{...iSt(t),minHeight:70,resize:"none",marginBottom:10}} value={editCard.text} onChange={e=>setEditCard(c=>({...c,text:e.target.value}))}/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
          <select style={iSt(t)} value={editCard.col} onChange={e=>setEditCard(c=>({...c,col:e.target.value}))}>{COLS.map(c=><option key={c.id} value={c.id}>{c.icon} {colLabel(c)}</option>)}</select>
          <select style={iSt(t)} value={editCard.priority} onChange={e=>setEditCard(c=>({...c,priority:e.target.value}))}>{Object.entries(PRIOS).map(([k,v])=><option key={k} value={k}>{v.icon} {v.label[lang]}</option>)}</select>
          <select style={iSt(t)} value={editCard.color_label||"normal"} onChange={e=>setEditCard(c=>({...c,color_label:e.target.value}))}>{Object.entries(COLOR_LABELS).map(([k,v])=><option key={k} value={k}>{v.label[lang]}</option>)}</select>
          <select style={iSt(t)} value={editCard.recurrence||""} onChange={e=>setEditCard(c=>({...c,recurrence:e.target.value}))}>{RECURRENCE.map(r=><option key={r.id} value={r.id}>{r.label[lang]}</option>)}</select>
          {canManage&&<select style={iSt(t)} value={editCard.assigned_to||""} onChange={e=>setEditCard(c=>({...c,assigned_to:e.target.value||null}))}>
            <option value="">{lang==="tr"?"Atama Yok":"Unassigned"}</option>
            {(teamMembers||[]).map(m=><option key={m.userId||m.user_id} value={m.userId||m.user_id}>{m.name}</option>)}
          </select>}
          <input type="date" style={{...iSt(t),fontSize:12}} value={editCard.due_date||""} onChange={e=>setEditCard(c=>({...c,due_date:e.target.value||null}))}/>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setEditCard(null)} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
          <button onClick={()=>updateCard(editCard.id,{text:editCard.text,col:editCard.col,priority:editCard.priority,assigned_to:editCard.assigned_to,due_date:editCard.due_date,color_label:editCard.color_label,recurrence:editCard.recurrence})} style={{...bSt("p",t),flex:2}}>{lang==="tr"?"Kaydet":"Save"}</button>
        </div>
      </div>
    </div>}

    {/* Detail Modal */}
    {detailCard&&<DetailModal card={cards.find(c=>c.id===detailCard.id)||detailCard} onClose={()=>setDetailCard(null)}/>}

    {/* Kanban View */}
    {view==="kanban"&&<div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:16}}>
      {COLS.map(col=>{
        const colCards=filteredCards.filter(c=>c.col===col.id);
        return <div key={col.id} style={{minWidth:200,flex:"0 0 200px",background:t.inBg,borderRadius:14,padding:10}}
          onDragOver={e=>e.preventDefault()}
          onDrop={e=>{e.preventDefault();if(dragCard)moveCard(dragCard,col.id);setDragCard(null);}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10}}>
            <span style={{fontSize:16}}>{col.icon}</span>
            <span style={{fontSize:12,fontWeight:700,color:col.color,flex:1}}>{colLabel(col)}</span>
            <span style={{fontSize:10,background:t.card,color:t.tm,padding:"2px 6px",borderRadius:8}}>{colCards.length}</span>
          </div>
          {colCards.map(card=>{
            const prio=PRIOS[card.priority||"medium"];
            const clr=COLOR_LABELS[card.color_label||"normal"];
            const overdue=card.due_date&&card.col!=="done"&&new Date(card.due_date)<new Date();
            const checkDone=(card.checklist||[]).filter(i=>i.done).length;
            const checkTotal=(card.checklist||[]).length;
            const borderColor=clr?.color||prio.color;
            return <div key={card.id} draggable onDragStart={()=>setDragCard(card.id)}
              onClick={()=>setDetailCard(card)}
              style={{background:t.card,borderRadius:10,padding:"10px 11px",marginBottom:8,cursor:"pointer",
                opacity:dragCard===card.id?0.4:1,border:`1px solid ${t.border}`,borderLeft:`3px solid ${borderColor}`}}>
              {card.card_type==="meeting"&&<div style={{fontSize:10,color:t.accent,fontWeight:700,marginBottom:3}}>📅 {lang==="tr"?"TOPLANTI":"MEETING"}</div>}
              <div style={{fontSize:13,color:t.text,lineHeight:1.4,marginBottom:6,wordBreak:"break-word"}}>{card.text}</div>
              {checkTotal>0&&<div style={{height:3,background:t.inBg,borderRadius:3,marginBottom:6}}>
                <div style={{height:3,background:t.accent,borderRadius:3,width:`${checkDone/checkTotal*100}%`}}/>
              </div>}
              <div style={{display:"flex",gap:4,flexWrap:"wrap",alignItems:"center"}}>
                {card.assigned_to&&<span style={{fontSize:10,color:t.accent,background:t.acB,padding:"2px 6px",borderRadius:6}}>👤 {getMemberName(card.assigned_to)}</span>}
                {card.visibility==="personal"&&<span style={{fontSize:9,color:t.tm,background:t.inBg,padding:"1px 5px",borderRadius:4}}>🔒</span>}
                {card.recurrence&&<span style={{fontSize:9,color:t.tm}}>🔄</span>}
                {card.due_date&&<span style={{fontSize:10,color:overdue?t.danger:t.tm}}>📅 {card.due_date}{overdue?" ⚠️":""}</span>}
                {checkTotal>0&&<span style={{fontSize:9,color:t.tm,marginLeft:"auto"}}>☑️ {checkDone}/{checkTotal}</span>}
                {(card.comments||[]).length>0&&<span style={{fontSize:9,color:t.tm}}>💬 {card.comments.length}</span>}
              </div>
            </div>;
          })}
          {!colCards.length&&<div style={{textAlign:"center",padding:"16px 8px",color:t.tm,fontSize:11,opacity:0.5}}>{lang==="tr"?"Boş":"Empty"}</div>}
        </div>;
      })}
    </div>}

    {/* Liste View */}
    {view==="list"&&<div>
      {!filteredCards.length&&<div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}><div style={{fontSize:32}}>📋</div><div style={{marginTop:8,fontSize:13}}>{lang==="tr"?"Görev yok":"No tasks"}</div></div>}
      {filteredCards.map(card=>{
        const prio=PRIOS[card.priority||"medium"];
        const col=COLS.find(c=>c.id===card.col)||COLS[0];
        const clr=COLOR_LABELS[card.color_label||"normal"];
        const overdue=card.due_date&&card.col!=="done"&&new Date(card.due_date)<new Date();
        const borderColor=clr?.color||prio.color;
        return <div key={card.id} onClick={()=>setDetailCard(card)} style={{...cSt(t),padding:"11px 14px",marginBottom:8,display:"flex",gap:10,alignItems:"flex-start",borderLeft:`3px solid ${borderColor}`,opacity:card.col==="done"?0.6:1,cursor:"pointer"}}>
          <button onClick={e=>{e.stopPropagation();moveCard(card.id,card.col==="done"?"todo":"done");}}
            style={{width:20,height:20,borderRadius:"50%",border:`2px solid ${card.col==="done"?t.success:t.inBo}`,
              background:card.col==="done"?t.success:"transparent",cursor:"pointer",flexShrink:0,marginTop:2,
              display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:11}}>
            {card.col==="done"?"✓":""}
          </button>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:14,color:t.text,textDecoration:card.col==="done"?"line-through":"none",lineHeight:1.4}}>
              {card.card_type==="meeting"&&"📅 "}{card.text}
            </div>
            <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>
              <span style={{fontSize:10,background:col.color+"20",color:col.color,padding:"2px 6px",borderRadius:6}}>{col.icon} {colLabel(col)}</span>
              {card.assigned_to&&<span style={{fontSize:10,color:t.accent}}>👤 {getMemberName(card.assigned_to)}</span>}
              {card.due_date&&<span style={{fontSize:10,color:overdue?t.danger:t.tm}}>📅 {card.due_date}{overdue?" ⚠️":""}</span>}
              {card.recurrence&&<span style={{fontSize:10,color:t.tm}}>🔄</span>}
              {(card.checklist||[]).length>0&&<span style={{fontSize:10,color:t.tm}}>☑️ {(card.checklist||[]).filter(i=>i.done).length}/{(card.checklist||[]).length}</span>}
              {(card.comments||[]).length>0&&<span style={{fontSize:10,color:t.tm}}>💬 {(card.comments||[]).length}</span>}
            </div>
          </div>
        </div>;
      })}
    </div>}
  </div>;
};

const BotRulesTab=({team,teamMembers,user,stock,setBotMessages,t})=>{
  const[rules,setRules]=useState([]);const[loading,setLoading]=useState(true);const[showNew,setShowNew]=useState(false);
  const EMPTY={trigger_type:"time",message:"",trigger_time:"09:00",days:["MO","TU","WE","TH","FR","SA","SU"],stock_item:"",stock_threshold:100,skt_days:3,shift_minutes:30,action_type:"team_chat",assign_to:"",active:true};
  const[form,setForm]=useState(EMPTY);
  const lang=t.lang;
  const DAYS=["MO","TU","WE","TH","FR","SA","SU"];
  const DL={tr:["Pzt","Sal","Çar","Per","Cum","Cmt","Paz"],en:["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]};
  const dl=DL[lang]||DL.en;
  const TT=[{id:"time",icon:"⏰",tr:"Zaman",en:"Time"},{id:"stock",icon:"📦",tr:"Stok",en:"Stock"},{id:"skt",icon:"📅",tr:"SKT",en:"Expiry"},{id:"shift",icon:"🕐",tr:"Vardiya",en:"Shift"},{id:"task",icon:"📋",tr:"Görev",en:"Task"}];
  const AT=[{id:"team_chat",icon:"👥",tr:"Ekip Sohbeti",en:"Team Chat"},{id:"dm",icon:"👤",tr:"Özel Mesaj",en:"DM"},{id:"create_task",icon:"📌",tr:"Görev Oluştur",en:"Create Task"}];
  useEffect(()=>{if(!team?.id)return;const sb=initSupabase();if(!sb)return;sb.from("bot_rules").select("*").eq("team_id",team.id).then(({data})=>{if(data)setRules(data);setLoading(false);});},[team?.id]);
  const saveRule=async()=>{
    if(!form.message.trim()||!team?.id)return;
    const sb=initSupabase();if(!sb){window.toast.error("Supabase yok");return;}
    // Empty UUID/string'leri NULL yap (Postgres UUID kolonu boş string kabul etmez)
    const payload={
      team_id:team.id,
      trigger_type:form.trigger_type,
      message:form.message.trim(),
      trigger_time:form.trigger_time||null,
      days:form.days&&form.days.length?form.days:null,
      stock_item:form.stock_item||null,
      stock_threshold:form.stock_threshold||null,
      skt_days:form.skt_days||null,
      shift_minutes:form.shift_minutes||null,
      action_type:form.action_type,
      assign_to:form.assign_to||null,
      active:form.active,
      created_by:user?.userId||null
    };
    const{data,error}=await sb.from("bot_rules").insert(payload).select().single();
    if(error){
      window.toast.error((lang==="tr"?"Kayıt başarısız: ":"Save failed: ")+error.message);
      return;
    }
    if(data){setRules(p=>[...p,data]);setShowNew(false);setForm(EMPTY);}
  };
  const toggleRule=async(rule)=>{const sb=initSupabase();if(!sb)return;await sb.from("bot_rules").update({active:!rule.active}).eq("id",rule.id);setRules(p=>p.map(r=>r.id===rule.id?{...r,active:!r.active}:r));};
  const deleteRule=async(id)=>{const sb=initSupabase();if(!sb)return;await sb.from("bot_rules").delete().eq("id",id);setRules(p=>p.filter(r=>r.id!==id));};
  const testRule=(rule)=>{
    if(setBotMessages)setBotMessages(p=>[...p,{id:Date.now(),text:`🤖 [TEST] ${rule.message.slice(0,60)}`,time:new Date().toLocaleTimeString(),type:"bot"}]);
    const sb=initSupabase();if(sb&&rule.action_type==="team_chat"&&team?.id&&user?.userId)sb.from("team_messages").insert({team_id:team.id,user_id:user.userId,user_name:"🤖 Bot",user_role:"bot",type:"text",text:`🤖 ${rule.message}`});
  };
  if(!team)return <div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}><div style={{fontSize:40}}>🤖</div><div style={{fontSize:13,marginTop:8,color:t.tm}}>{lang==="tr"?"Ekibe katılın":"Join a team"}</div></div>;
  return <div style={{maxWidth:520,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
      <h2 style={{fontSize:22,color:t.text,fontFamily:"'Fraunces',serif",margin:0}}>⚙️ {lang==="tr"?"Otomasyon":"Automation"}</h2>
      <button onClick={()=>setShowNew(s=>!s)} style={{...bSt("p",t),fontSize:12}}>+ {lang==="tr"?"Yeni":"New"}</button>
    </div>
    {showNew&&<div style={{...cSt(t),padding:16,marginBottom:16}}>
      <div style={{fontSize:12,fontWeight:700,color:t.text,marginBottom:8}}>{lang==="tr"?"1. Tetikleyici:":"1. Trigger:"}</div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
        {TT.map(tt=><button key={tt.id} onClick={()=>setForm(f=>({...f,trigger_type:tt.id}))} style={{padding:"6px 10px",borderRadius:10,fontSize:11,fontWeight:600,border:"none",cursor:"pointer",background:form.trigger_type===tt.id?t.accent:t.inBg,color:form.trigger_type===tt.id?"#fff":t.tm}}>{tt.icon} {tt[lang]||tt.en}</button>)}
      </div>
      {form.trigger_type==="time"&&<><input type="time" style={{...iSt(t),marginBottom:8}} value={form.trigger_time} onChange={e=>setForm(f=>({...f,trigger_time:e.target.value}))}/>
        <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:12}}>{DAYS.map((d,i)=><button key={d} onClick={()=>setForm(f=>({...f,days:f.days.includes(d)?f.days.filter(x=>x!==d):[...f.days,d]}))} style={{padding:"5px 8px",borderRadius:8,fontSize:11,fontWeight:600,border:"none",cursor:"pointer",background:form.days.includes(d)?t.accent:t.inBg,color:form.days.includes(d)?"#fff":t.tm}}>{dl[i]}</button>)}</div></>}
      {form.trigger_type==="stock"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}><input style={iSt(t)} placeholder="Ürün" value={form.stock_item} onChange={e=>setForm(f=>({...f,stock_item:e.target.value}))}/><input type="number" style={iSt(t)} value={form.stock_threshold} onChange={e=>setForm(f=>({...f,stock_threshold:parseInt(e.target.value,10)||0}))}/></div>}
      {form.trigger_type==="skt"&&<div style={{display:"flex",gap:6,marginBottom:12}}>{[1,3,7,14].map(d=><button key={d} onClick={()=>setForm(f=>({...f,skt_days:d}))} style={{flex:1,padding:"8px",borderRadius:10,fontSize:13,fontWeight:700,border:"none",cursor:"pointer",background:form.skt_days===d?t.accent:t.inBg,color:form.skt_days===d?"#fff":t.tm}}>{d}</button>)}</div>}
      {form.trigger_type==="shift"&&<div style={{display:"flex",gap:6,marginBottom:12}}>{[15,30,60].map(m=><button key={m} onClick={()=>setForm(f=>({...f,shift_minutes:m}))} style={{flex:1,padding:"8px",borderRadius:10,fontSize:13,fontWeight:700,border:"none",cursor:"pointer",background:form.shift_minutes===m?t.accent:t.inBg,color:form.shift_minutes===m?"#fff":t.tm}}>{m}dk</button>)}</div>}
      <div style={{fontSize:12,fontWeight:700,color:t.text,marginBottom:8}}>{lang==="tr"?"2. Mesaj:":"2. Message:"}</div>
      <textarea style={{...iSt(t),minHeight:70,resize:"vertical",marginBottom:12,fontSize:13}} placeholder={lang==="tr"?"Günaydın ekip!":"Good morning team!"} value={form.message} onChange={e=>setForm(f=>({...f,message:e.target.value}))}/>
      <div style={{fontSize:12,fontWeight:700,color:t.text,marginBottom:8}}>{lang==="tr"?"3. Aksiyon:":"3. Action:"}</div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>{AT.map(at=><button key={at.id} onClick={()=>setForm(f=>({...f,action_type:at.id}))} style={{padding:"6px 10px",borderRadius:10,fontSize:11,fontWeight:600,border:"none",cursor:"pointer",background:form.action_type===at.id?t.accent:t.inBg,color:form.action_type===at.id?"#fff":t.tm}}>{at.icon} {at[lang]||at.en}</button>)}</div>
      {(form.action_type==="dm"||form.action_type==="create_task")&&<select style={{...iSt(t),marginBottom:10}} value={form.assign_to} onChange={e=>setForm(f=>({...f,assign_to:e.target.value}))}>
        <option value="">{lang==="tr"?"Tüm ekip":"Everyone"}</option>
        {(teamMembers||[]).map(m=><option key={m.userId||m.user_id} value={m.userId||m.user_id}>{m.name}</option>)}
      </select>}
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>{setShowNew(false);setForm(EMPTY);}} style={{...bSt("s",t),flex:1}}>{lang==="tr"?"İptal":"Cancel"}</button>
        <button onClick={saveRule} disabled={!form.message.trim()} style={{...bSt("p",t),flex:2,opacity:form.message.trim()?1:0.5}}>✓ {lang==="tr"?"Kaydet":"Save"}</button>
      </div>
    </div>}
    {loading&&<div style={{textAlign:"center",padding:20,color:t.tm}}>⏳</div>}
    {!loading&&!rules.length&&<div style={{textAlign:"center",padding:"40px 20px",color:t.tm}}><div style={{fontSize:36}}>🤖</div><div style={{fontSize:13,marginTop:8}}>{lang==="tr"?"Kural yok":"No rules"}</div></div>}
    {rules.map(rule=>{const tt=TT.find(x=>x.id===rule.trigger_type)||TT[0];return <div key={rule.id} style={{...cSt(t),padding:"14px 16px",marginBottom:10,opacity:rule.active?1:0.5}}>
      <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
        <div style={{width:38,height:38,borderRadius:10,background:rule.active?t.acB:t.inBg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{tt.icon}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:12,color:t.tm,marginBottom:3}}>{tt[lang]||tt.en}{rule.trigger_type==="time"?` · ${rule.trigger_time}`:rule.trigger_type==="stock"?` · ${rule.stock_item}`:""}</div>
          <div style={{fontSize:13,color:t.text,lineHeight:1.4}}>{rule.message.slice(0,80)}{rule.message.length>80?"...":""}</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:5,alignItems:"center"}}>
          <label style={{position:"relative",display:"inline-block",width:36,height:20,cursor:"pointer"}}>
            <input type="checkbox" checked={!!rule.active} onChange={()=>toggleRule(rule)} style={{opacity:0,width:0,height:0}}/>
            <span style={{position:"absolute",inset:0,background:rule.active?t.accent:"#ccc",borderRadius:10,transition:"0.2s"}}><span style={{position:"absolute",width:14,height:14,left:rule.active?19:3,top:3,background:"#fff",borderRadius:"50%",transition:"0.2s"}}/></span>
          </label>
          <button onClick={()=>testRule(rule)} style={{...bSt("s",t),padding:"2px 7px",fontSize:10}}>▶</button>
          <button onClick={()=>deleteRule(rule.id)} style={{background:"none",border:"none",color:t.danger,cursor:"pointer",fontSize:16}}>×</button>
        </div>
      </div>
    </div>;})}
  </div>;
};


// ═══ SVG İKON SİSTEMİ ═══
const ICONS = {
  recipes: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="8" y="4" width="28" height="36" rx="2" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <circle cx="8" cy="12" r="3" stroke="${c}" stroke-width="2.2"/>
    <circle cx="8" cy="22" r="3" stroke="${c}" stroke-width="2.2"/>
    <circle cx="8" cy="32" r="3" stroke="${c}" stroke-width="2.2"/>
    <line x1="16" y1="12" x2="32" y2="12" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="16" y1="22" x2="28" y2="22" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="16" y1="32" x2="30" y2="32" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
  </svg>`,
  stock: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <line x1="6" y1="6" x2="6" y2="40" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="38" y1="6" x2="38" y2="40" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="6" y1="16" x2="38" y2="16" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="6" y1="28" x2="38" y2="28" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <rect x="9" y="7" width="10" height="9" rx="1" stroke="${c}" stroke-width="1.8"/>
    <rect x="22" y="7" width="14" height="9" rx="1" stroke="${c}" stroke-width="1.8"/>
    <rect x="9" y="19" width="16" height="9" rx="1" stroke="${c}" stroke-width="1.8"/>
    <rect x="28" y="19" width="8" height="9" rx="1" stroke="${c}" stroke-width="1.8"/>
  </svg>`,
  production: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="22" cy="28" r="14" stroke="${c}" stroke-width="2.2"/>
    <circle cx="22" cy="28" r="8" stroke="${c}" stroke-width="2.2"/>
    <rect x="14" y="18" width="16" height="8" rx="2" stroke="${c}" stroke-width="2.2"/>
    <line x1="19" y1="8" x2="19" y2="18" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <path d="M12 6 Q14 2 16 6" stroke="${c}" stroke-width="2" fill="none" stroke-linecap="round"/>
    <path d="M19 4 Q21 0 23 4" stroke="${c}" stroke-width="2" fill="none" stroke-linecap="round"/>
    <path d="M26 6 Q28 2 30 6" stroke="${c}" stroke-width="2" fill="none" stroke-linecap="round"/>
  </svg>`,
  menus: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="4" y="4" width="12" height="36" rx="1.5" stroke="${c}" stroke-width="2.2"/>
    <rect x="18" y="4" width="12" height="36" rx="1.5" stroke="${c}" stroke-width="2.2"/>
    <rect x="32" y="4" width="8" height="36" rx="1.5" stroke="${c}" stroke-width="2.2"/>
    <line x1="7" y1="12" x2="13" y2="12" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="7" y1="18" x2="13" y2="18" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="7" y1="24" x2="13" y2="24" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="21" y1="12" x2="27" y2="12" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="21" y1="18" x2="27" y2="18" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="35" y1="12" x2="38" y2="12" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
  </svg>`,
  kanban: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="4" y="4" width="36" height="36" rx="3" stroke="${c}" stroke-width="2.2"/>
    <line x1="17" y1="4" x2="17" y2="40" stroke="${c}" stroke-width="2.2"/>
    <line x1="30" y1="4" x2="30" y2="40" stroke="${c}" stroke-width="2.2"/>
    <rect x="7" y="8" width="7" height="12" rx="1.5" stroke="${c}" stroke-width="1.8"/>
    <rect x="7" y="24" width="7" height="8" rx="1.5" stroke="${c}" stroke-width="1.8"/>
    <rect x="20" y="8" width="7" height="18" rx="1.5" stroke="${c}" stroke-width="1.8"/>
    <rect x="33" y="8" width="4" height="8" rx="1.5" stroke="${c}" stroke-width="1.8"/>
    <rect x="33" y="20" width="4" height="14" rx="1.5" stroke="${c}" stroke-width="1.8"/>
  </svg>`,
  chat: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M4 8 L4 30 L14 30 L14 40 L26 30 L40 30 L40 8 Z" stroke="${c}" stroke-width="2.2" stroke-linejoin="round"/>
    <line x1="12" y1="17" x2="32" y2="17" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="12" y1="24" x2="24" y2="24" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <circle cx="36" cy="8" r="5" stroke="${c}" stroke-width="2"/>
    <circle cx="36" cy="8" r="2" fill="${c}"/>
  </svg>`,
  events: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M10 36 L8 14 L26 32 L10 36 Z" stroke="${c}" stroke-width="2.2" stroke-linejoin="round"/>
    <line x1="14" y1="20" x2="18" y2="24" stroke="${c}" stroke-width="2" stroke-linecap="round"/>
    <line x1="18" y1="16" x2="22" y2="20" stroke="${c}" stroke-width="2" stroke-linecap="round"/>
    <line x1="32" y1="6" x2="34" y2="8" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="38" y1="10" x2="40" y2="12" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <line x1="36" y1="18" x2="38" y2="20" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <circle cx="30" cy="14" r="1.5" fill="${c}"/>
    <circle cx="40" cy="22" r="1.5" fill="${c}"/>
    <circle cx="34" cy="26" r="1.5" fill="${c}"/>
  </svg>`,
  hub: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="6" y="6" width="32" height="34" rx="2" stroke="${c}" stroke-width="2.2"/>
    <line x1="22" y1="6" x2="22" y2="40" stroke="${c}" stroke-width="2.2"/>
    <rect x="10" y="11" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
    <rect x="10" y="20" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
    <rect x="10" y="29" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
    <rect x="28" y="11" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
    <rect x="28" y="20" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
    <rect x="28" y="29" width="6" height="5" rx="0.5" stroke="${c}" stroke-width="1.6"/>
  </svg>`,
  shift: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="22" cy="24" r="16" stroke="${c}" stroke-width="2.2"/>
    <line x1="22" y1="24" x2="22" y2="12" stroke="${c}" stroke-width="2.4" stroke-linecap="round"/>
    <line x1="22" y1="24" x2="30" y2="29" stroke="${c}" stroke-width="2.2" stroke-linecap="round"/>
    <circle cx="22" cy="24" r="2.5" fill="${c}"/>
    <line x1="16" y1="4" x2="16" y2="8" stroke="${c}" stroke-width="2" stroke-linecap="round"/>
    <line x1="28" y1="4" x2="28" y2="8" stroke="${c}" stroke-width="2" stroke-linecap="round"/>
    <line x1="22" y1="8" x2="22" y2="10" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="22" y1="38" x2="22" y2="40" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="6" y1="24" x2="8" y2="24" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
    <line x1="36" y1="24" x2="38" y2="24" stroke="${c}" stroke-width="1.6" stroke-linecap="round"/>
  </svg>`,
  automation: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="22" cy="22" r="10" stroke="${c}" stroke-width="2.2"/>
    <circle cx="22" cy="22" r="4" stroke="${c}" stroke-width="2.2"/>
    <rect x="20" y="4" width="4" height="6" rx="2" fill="${c}"/>
    <rect x="20" y="34" width="4" height="6" rx="2" fill="${c}"/>
    <rect x="4" y="20" width="6" height="4" rx="2" fill="${c}"/>
    <rect x="34" y="20" width="6" height="4" rx="2" fill="${c}"/>
    <rect x="9" y="9" width="4" height="4" rx="2" fill="${c}" transform="rotate(45,11,11)"/>
    <rect x="31" y="9" width="4" height="4" rx="2" fill="${c}" transform="rotate(45,33,11)"/>
    <rect x="9" y="31" width="4" height="4" rx="2" fill="${c}" transform="rotate(45,11,33)"/>
    <rect x="31" y="31" width="4" height="4" rx="2" fill="${c}" transform="rotate(45,33,33)"/>
  </svg>`,
  settings: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="22" cy="22" r="6" stroke="${c}" stroke-width="2.2"/>
    <path d="M22 4 L25 10 L32 8 L32 16 L38 19 L34 25 L38 31 L32 34 L32 42 L25 40 L22 44 L19 40 L12 42 L12 34 L6 31 L10 25 L6 19 L12 16 L12 8 L19 10 Z" stroke="${c}" stroke-width="2.2" stroke-linejoin="round" fill="none"/>
  </svg>`,
  reports: (c='currentColor') => `<svg width="22" height="22" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="4" y="4" width="36" height="36" rx="3" stroke="${c}" stroke-width="2.2"/>
    <line x1="12" y1="32" x2="12" y2="24" stroke="${c}" stroke-width="3" stroke-linecap="round"/>
    <line x1="20" y1="32" x2="20" y2="14" stroke="${c}" stroke-width="3" stroke-linecap="round"/>
    <line x1="28" y1="32" x2="28" y2="20" stroke="${c}" stroke-width="3" stroke-linecap="round"/>
    <line x1="36" y1="32" x2="36" y2="10" stroke="${c}" stroke-width="3" stroke-linecap="round"/>
    <polyline points="12,24 20,14 28,20 36,10" stroke="${c}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  </svg>`,
};

// Tab için SVG ikon render
const TabIcon = ({name, active, size=22}) => {
  const color = active ? '#C8965A' : 'currentColor';
  const fn = ICONS[name];
  if(!fn) return null;
  return <span style={{display:'flex',alignItems:'center',justifyContent:'center'}}
    dangerouslySetInnerHTML={{__html: fn(color)}}/>;
};

// ═══ ARKA PLAN SİSTEMİ ═══
const WALLPAPERS = [
  {id:'default', label:{tr:'Varsayılan',en:'Default'}, preview:'#F5F0E8', style:{background:'#F5F0E8'}},
  {id:'white', label:{tr:'Beyaz',en:'White'}, preview:'#FFFFFF', style:{background:'#FFFFFF'}},
  {id:'dark', label:{tr:'Koyu',en:'Dark'}, preview:'#1A1A16', style:{background:'#1A1A16'}},
  {id:'linen', label:{tr:'Keten',en:'Linen'}, preview:'#F4EDE4',
    style:{background:"url('data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%224%22 height=%224%22%3E%3Crect width=%224%22 height=%224%22 fill=%22%23F4EDE4%22/%3E%3Crect x=%220%22 y=%220%22 width=%221%22 height=%221%22 fill=%22%23E8DDD2%22 opacity=%220.4%22/%3E%3Crect x=%222%22 y=%222%22 width=%221%22 height=%221%22 fill=%22%23E8DDD2%22 opacity=%220.4%22/%3E%3C/svg%3E') repeat"}},
  {id:'slate', label:{tr:'Taş',en:'Slate'}, preview:'#8A9BA8', style:{background:'linear-gradient(135deg,#6B7F8A 0%,#8A9BA8 50%,#7A8E9A 100%)'}},
  {id:'forest', label:{tr:'Orman',en:'Forest'}, preview:'#2D4A3E', style:{background:'linear-gradient(160deg,#1E3A2F 0%,#2D4A3E 60%,#1A3028 100%)'}},
  {id:'dusk', label:{tr:'Alacakaranlık',en:'Dusk'}, preview:'#3D2B4A', style:{background:'linear-gradient(160deg,#2A1A38 0%,#3D2B4A 50%,#1E1428 100%)'}},
  {id:'coffee', label:{tr:'Kahve',en:'Coffee'}, preview:'#4A3020', style:{background:'linear-gradient(160deg,#2E1C0E 0%,#4A3020 60%,#3A2414 100%)'}},
  {id:'cream', label:{tr:'Krem',en:'Cream'}, preview:'#FDF6E8', style:{background:'linear-gradient(135deg,#FDF6E8 0%,#F5EDD8 100%)'}},
  {id:'custom', label:{tr:'Fotoğraf',en:'Photo'}, preview:'gradient', style:{}},
];

const useSwipeNav=(tabs,tab,setTab)=>{
  const startX=useRef(0);
  const startY=useRef(0);
  const startT=useRef(0);
  const onTouchStart=(e)=>{
    const t=e.touches[0];
    startX.current=t.clientX;
    startY.current=t.clientY;
    startT.current=Date.now();
  };
  const onTouchEnd=(e)=>{
    const t=e.changedTouches[0];
    const dx=t.clientX-startX.current;
    const dy=t.clientY-startY.current;
    const dt=Date.now()-startT.current;
    // Yatay swipe: x > 60px, y < 50px (dikey scroll'la karışmasın), süre < 600ms
    if(Math.abs(dx)<60||Math.abs(dy)>50||dt>600)return;
    // Form/input/textarea/scroll içindeyse iptal et
    const target=e.target;
    if(target.closest("input,textarea,select,[data-no-swipe]"))return;
    // Yatay scroll edilebilir bir element üzerindeyse iptal
    let p=target;
    while(p&&p!==document.body){
      const cs=window.getComputedStyle(p);
      if((cs.overflowX==="auto"||cs.overflowX==="scroll")&&p.scrollWidth>p.clientWidth)return;
      p=p.parentElement;
    }
    const idx=tabs.findIndex(x=>x.id===tab);
    if(idx===-1)return;
    if(dx<0&&idx<tabs.length-1)setTab(tabs[idx+1].id); // sola swipe → sonraki
    else if(dx>0&&idx>0)setTab(tabs[idx-1].id); // sağa swipe → önceki
  };
  return{onTouchStart,onTouchEnd};
};


const DockTabBar=({tabs,tab,setTab,t,ICONS,lowCount})=>{
  const scrollerRef=useRef(null);
  const tabRefs=useRef({});
  const isJumping=useRef(false);

  // 3x kopyalanmış tab listesi — sonsuz döngü için
  const triple=useMemo(()=>[
    ...tabs.map((tb,i)=>({...tb,_idx:i,_copy:0,_uid:`a-${tb.id}`})),
    ...tabs.map((tb,i)=>({...tb,_idx:i,_copy:1,_uid:`b-${tb.id}`})),
    ...tabs.map((tb,i)=>({...tb,_idx:i,_copy:2,_uid:`c-${tb.id}`})),
  ],[tabs]);

  // Aktif tab değişince ORTADAKİ kopyayı merkeze kaydır
  useEffect(()=>{
    const sc=scrollerRef.current;
    if(!sc)return;
    const el=tabRefs.current[`b-${tab}`];
    if(!el)return;
    const elRect=el.getBoundingClientRect();
    const scRect=sc.getBoundingClientRect();
    const targetScroll=sc.scrollLeft+elRect.left-scRect.left-(scRect.width/2)+(elRect.width/2);
    isJumping.current=true;
    sc.scrollTo({left:targetScroll,behavior:"smooth"});
    setTimeout(()=>{isJumping.current=false;},400);
  },[tab,tabs.length]);

  // Scroll teleport — kullanıcı a veya c kopyasındaysa b'ye teleport et
  const onScroll=useCallback(()=>{
    if(isJumping.current)return;
    const sc=scrollerRef.current;
    if(!sc)return;
    const sw=sc.scrollWidth/3; // bir kopyanın genişliği
    if(sc.scrollLeft<sw*0.5){
      // a kopyasındayız, b'ye teleport
      isJumping.current=true;
      sc.scrollLeft=sc.scrollLeft+sw;
      setTimeout(()=>{isJumping.current=false;},50);
    }else if(sc.scrollLeft>sw*2.0){
      // c kopyasındayız, b'ye teleport
      isJumping.current=true;
      sc.scrollLeft=sc.scrollLeft-sw;
      setTimeout(()=>{isJumping.current=false;},50);
    }
  },[]);

  // İlk yüklemede ortadaki kopyayı (b) ortala
  useEffect(()=>{
    const sc=scrollerRef.current;
    if(!sc)return;
    const t0=setTimeout(()=>{
      const sw=sc.scrollWidth/3;
      const initialEl=tabRefs.current[`b-${tab}`];
      if(initialEl){
        const elRect=initialEl.getBoundingClientRect();
        const scRect=sc.getBoundingClientRect();
        sc.scrollLeft=sc.scrollLeft+elRect.left-scRect.left-(scRect.width/2)+(elRect.width/2);
      }else{
        sc.scrollLeft=sw;
      }
    },50);
    return()=>clearTimeout(t0);
  },[]);

  return <div style={{position:"fixed",bottom:0,left:0,right:0,background:t.topBar,backdropFilter:"blur(20px)",WebkitBackdropFilter:"blur(20px)",borderTop:`1px solid ${t.border}`,zIndex:100,padding:"6px 0 calc(env(safe-area-inset-bottom) + 6px)"}}>
    <div ref={scrollerRef} onScroll={onScroll} style={{display:"flex",overflowX:"auto",scrollbarWidth:"none",msOverflowStyle:"none",WebkitOverflowScrolling:"touch",gap:2}}>
      <style dangerouslySetInnerHTML={{__html:`.km-dock::-webkit-scrollbar{display:none}.km-dock-btn{transition:transform 0.25s cubic-bezier(0.34,1.56,0.64,1),opacity 0.2s,padding 0.2s}.km-dock-btn:active{transform:scale(0.92)}`}}/>
      {triple.map(tb=>{
        const isActive=tab===tb.id;
        return <button
          key={tb._uid}
          ref={el=>{if(el)tabRefs.current[tb._uid]=el;}}
          onClick={()=>setTab(tb.id)}
          className="km-dock-btn"
          style={{
            position:"relative",
            background:isActive?t.acB:"transparent",
            border:"none",
            cursor:"pointer",
            display:"flex",
            flexDirection:"column",
            alignItems:"center",
            justifyContent:"center",
            gap:2,
            padding:isActive?"7px 14px":"7px 10px",
            margin:"0 1px",
            borderRadius:14,
            color:isActive?t.accent:t.tm,
            flexShrink:0,
            transform:isActive?"scale(1.08)":"scale(1)",
            opacity:isActive?1:0.55,
          }}>
          <span style={{width:24,height:24,display:"flex",alignItems:"center",justifyContent:"center"}}>
            {ICONS&&ICONS[tb.icon]?<span dangerouslySetInnerHTML={{__html:ICONS[tb.icon](isActive?(t.accent||"#C8965A"):"currentColor")}}/>:<span style={{fontSize:20}}>{tb.i}</span>}
          </span>
          <span style={{fontSize:isActive?10:9,fontWeight:isActive?700:500,whiteSpace:"nowrap",maxWidth:64,overflow:"hidden",textOverflow:"ellipsis",lineHeight:1.2}}>{tb.l}</span>
          {tb.id==="stock"&&lowCount>0&&<div style={{position:"absolute",top:4,right:6,width:7,height:7,borderRadius:4,background:t.danger,boxShadow:"0 0 0 2px "+t.topBar}}/>}
        </button>;
      })}
    </div>
  </div>;
};



export default function App(){
  const[apiKey,setApiKey]=useState("worker-proxy");
  const[dark,setDark]=useState(LS.get(SK.dark,false));
  const[lang,setLang]=useState(LS.get(SK.lang,"tr"));
  const[tab,setTab]=useState("recipes");
  const[recipes,setRecipes]=useState(()=>{
    const raw=LS.get(SK.recipes,DEF_RECIPES);
    return raw.map(r=>({...r,totalWeight:r.totalWeight||0}));
  });
  const[stock,setStock]=useState(()=>{
    const raw=LS.get(SK.stock,DEF_STOCK);
    // Migration: eski kategori → yeni kategori
    return raw.map(migrateStockCat);
  });
  const[invoices,setInvoices]=useState(LS.get(SK.invoices,[]));
  const[menus,setMenus]=useState(LS.get(SK.menus,[]));
  const[expenses,setExpenses]=useState(LS.get(SK.expenses,DEF_EXPENSES));
  const[storageAreas,setStorageAreas]=useState(LS.get(SK.storage,DEF_STORAGE));
  const[productions,setProductions]=useState(LS.get(SK.productions,[]));
  const[reportCats,setReportCats]=useState(LS.get(SK.reportCats,DEF_REPORT_CATS));
  const[profile,setProfile]=useState(LS.get(SK.profile,DEF_PROFILE));
  const[traceability,setTraceability]=useState(LS.get(SK.traceability,false));
  const[lots,setLots]=useState(LS.get(SK.lots,{}));
  const[trackedIngs,setTrackedIngs]=useState(LS.get(SK.trackedIngs,[]));
  const[resetHour,setResetHour]=useState(LS.get(SK.resetHour,23));
  const[organizations,setOrganizations]=useState(LS.get(SK.organizations,DEF_ORGANIZATIONS));
  const[storageChecks,setStorageChecks]=useState(LS.get(SK.storageChecks,[]));
  const[menuTemplates,setMenuTemplates]=useState(LS.get(SK.menuTemplates,[]));
  // Sohbet oturumları + bot sistemi
  const[conversations,setConversations]=useState(LS.get(SK.conversations,[]));
  const[activeConvId,setActiveConvId]=useState(LS.get(SK.activeConvId,null));
  const[botMessages,setBotMessages]=useState(LS.get(SK.botMessages,[]));
  const[notifSettings,setNotifSettings]=useState(LS.get(SK.notifSettings,{enabled:true,storageCheck:true,expiredSKT:true,lowStock:true,lotReminder:true}));
  const[calorieDB,setCalorieDB]=useState(LS.get(SK.calorieDB,{}));
  const[printers,setPrinters]=useState(LS.get(SK.printers,[]));
  const[todos,setTodos]=useState(LS.get("kmp_todos",[]));
  const[team,setTeam]=useState(LS.get("kmp_team",null));
  const[teamMembers,setTeamMembers]=useState(()=>{
    // Eski cache'te UUID varsa temizle, Supabase'den yeniden çekecek
    const cached=LS.get("kmp_team_members",[]);
    if(cached.some(m=>m.name&&m.name.length===36&&m.name.includes("-")))return[];
    return cached;
  });
  // Phantom üyeler — uygulamaya kayıtlı olmayan ekip üyeleri
  const[phantomMembers,setPhantomMembers]=useState(LS.get("kmp_phantom_members",[]));
  // Phantom üyeleri Supabase'den yükle
  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    sb.from("team_phantom_members").select("*").eq("team_id",team.id).order("created_at",{ascending:true}).then(({data,error})=>{
      if(!error&&data){
        setPhantomMembers(data);
        LS.set("kmp_phantom_members",data);
      }
    });
  },[team?.id]);


  // Parent team bilgisini yükle
  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    sb.from("teams").select("parent_team_id").eq("id",team.id).single().then(async({data})=>{
      if(!data?.parent_team_id){
        if(team.parent_team_id)setTeam(t=>({...t,parent_team_id:null,parent_team_name:null}));
        return;
      }
      const{data:parent}=await sb.from("teams").select("name").eq("id",data.parent_team_id).single();
      setTeam(t=>({...t,parent_team_id:data.parent_team_id,parent_team_name:parent?.name||null}));
    });
  },[team?.id]);

  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    // Tüm üyeleri yükle: kendi ekibi + alt ekipler
    (async()=>{
      try{
        // 1. Kendi ekibinin üyeleri
        const{data:ownMembers}=await sb.from("team_members").select("user_id,role,position").eq("team_id",team.id);
        // 2. Alt ekiplerin üyeleri
        const{data:childTeams}=await sb.from("teams").select("id,name").eq("parent_team_id",team.id);
        let childMembers=[];
        if(childTeams&&childTeams.length>0){
          const childIds=childTeams.map(t=>t.id);
          const{data:cm}=await sb.from("team_members").select("user_id,role,position,team_id").in("team_id",childIds);
          if(cm)childMembers=cm.map(m=>({...m,teamName:(childTeams.find(t=>t.id===m.team_id)||{}).name||""}));
        }
        // 3. Tüm user_id'leri topla, profiles'dan isim çek
        const allRaw=[...(ownMembers||[]),...childMembers];
        const allUids=[...new Set(allRaw.map(m=>m.user_id))];
        const{data:profiles}=await sb.from("profiles").select("id,full_name,email").in("id",allUids);
        const getName=(uid)=>{const p=(profiles||[]).find(p=>p.id===uid);return p?.full_name||p?.email?.split("@")[0]||uid;};
        const mapped=allRaw.map(m=>({userId:m.user_id,name:getName(m.user_id),role:m.role,position:m.position,teamName:m.teamName||""}));
        setTeamMembers(mapped);
        LS.set("kmp_team_members",mapped);
      }catch(e){console.warn("Üye yüklenemedi:",e.message);}
    })();
    // Team'in kendisini de doğrula
    sb.from("teams").select("*").eq("id",team.id).single().then(({data,error})=>{
      if(error){console.warn("team fetch error:",error.message);return;}
      if(data){
        const updated={...data,role:team.role,inviteCode:data.invite_code};
        if(JSON.stringify(updated)!==JSON.stringify(team)){setTeam(updated);}
      }
    });
  },[team?.id]);
  useEffect(()=>{
    if(team)LS.set("kmp_team",team);
    else localStorage.removeItem("kmp_team");
  },[team]);
  useEffect(()=>{
    if(teamMembers?.length>0)LS.set("kmp_team_members",teamMembers);
  },[teamMembers]);
  useEffect(()=>{LS.set("kmp_todos",todos)},[todos]);
  // Etiket takip numarası: YYYYMMDD-XXXX formatı, günlük sıfırlanır
  const getLabelSeq=()=>{
    const today=new Date().toISOString().slice(0,10).replace(/-/g,"");
    const stored=LS.get("kmp_labelseq",{date:"",seq:0});
    let seq=stored.date===today?stored.seq+1:1;
    LS.set("kmp_labelseq",{date:today,seq});
    return `${today}-${String(seq).padStart(4,"0")}`;
  };
  // AUTH
  const[user,setUser]=useState(null);

  // ═══ OTOMATİK BOT KONTROLÜ — App açılınca çalışır ═══
  const stockRef=useRef(stock);
  useEffect(()=>{stockRef.current=stock;},[stock]);
  const userRef=useRef(user);
  useEffect(()=>{userRef.current=user;},[user]);
  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    const checkBots=async()=>{
      const{data:rules}=await sb.from("bot_rules").select("*").eq("team_id",team.id).eq("active",true);
      if(!rules?.length)return;
      const now=new Date();
      const today=now.toISOString().slice(0,10);
      const currentMinutes=now.getHours()*60+now.getMinutes();
      const DAY_MAP={0:"SU",1:"MO",2:"TU",3:"WE",4:"TH",5:"FR",6:"SA"};
      const todayKey=DAY_MAP[now.getDay()];
      const lastRunKey=`kmp_bot_lastrun_${team.id}`;
      const lastRun=JSON.parse(localStorage.getItem(lastRunKey)||"{}");
      const currentStock=stockRef.current||[];
      const currentUser=userRef.current;

      for(const rule of rules){
        const ruleKey=`${rule.id}_${today}`;
        if(lastRun[ruleKey])continue;
        let shouldFire=false;
        if(rule.trigger_type==="time"){
          const ruleTime=rule.trigger_time||"09:00";
          const ruleDays=rule.days||["MO","TU","WE","TH","FR","SA","SU"];
          const [rh,rm]=ruleTime.split(":").map(Number);
          const ruleMinutes=rh*60+rm;
          if(ruleDays.includes(todayKey)&&currentMinutes>=ruleMinutes&&currentMinutes<ruleMinutes+30){
            shouldFire=true;
          }
        } else if(rule.trigger_type==="stock"){
          const item=currentStock.find(s=>s.name?.toLowerCase().includes((rule.stock_item||"").toLowerCase()));
          if(item&&item.qty<=(rule.stock_threshold||100))shouldFire=true;
        } else if(rule.trigger_type==="skt"){
          const days=rule.skt_days||3;
          const threshold=new Date(Date.now()+days*86400000).toISOString().slice(0,10);
          const expiring=currentStock.filter(s=>s.skt&&s.skt<=threshold&&s.skt>=today);
          if(expiring.length>0)shouldFire=true;
        }
        if(shouldFire){
          const msg=rule.message;
          if(rule.action_type==="team_chat"){
            await sb.from("team_messages").insert({team_id:team.id,user_id:currentUser?.userId||"bot",user_name:"⚙️ Otomasyon",user_role:"bot",type:"text",text:`⚙️ ${msg}`});
          } else if(rule.action_type==="dm"&&rule.assign_to){
            await sb.from("team_messages").insert({team_id:team.id,user_id:currentUser?.userId||"bot",user_name:"⚙️ Otomasyon",user_role:"bot",type:"text",text:`⚙️ ${msg}`,private_to:rule.assign_to});
          }
          setBotMessages(p=>[...p,{id:Date.now(),text:msg,ts:new Date().toISOString(),type:"bot",icon:"⚙️"}]);
          lastRun[ruleKey]=now.toISOString();
          localStorage.setItem(lastRunKey,JSON.stringify(lastRun));
        }
      }
    };
    checkBots();
    const interval=setInterval(checkBots,2*60*1000);
    return()=>clearInterval(interval);
  },[team?.id]);

  const[wallpaper,setWallpaper]=useState(localStorage.getItem("kmp_wallpaper")||"default");
  const[customWP,setCustomWP]=useState(localStorage.getItem("kmp_customwp")||"");
  // Kullanıcı adı güncelleme — session'dan gerçek isim al
  useEffect(()=>{
    if(!user?.userId)return;
    const sb=initSupabase();if(!sb)return;
    sb.auth.getSession().then(({data:{session}})=>{
      if(!session?.user)return;
      const realName=session.user.user_metadata?.name||session.user.user_metadata?.full_name||session.user.email?.split("@")[0];
      if(realName&&realName!==user.name){
        setUser(u=>u?{...u,name:realName}:u);
        LS.set("kmp_user",{...user,name:realName});
      }
    });
  },[user?.userId]);
  const[authChecked,setAuthChecked]=useState(false);
  const[showAuth,setShowAuth]=useState(()=>{
    const p=new URLSearchParams(window.location.search);
    return p.get("mode")==="reset";
  });
  // mode=reset gelince user temizle
  useEffect(()=>{
    const p=new URLSearchParams(window.location.search);
    if(p.get("mode")==="reset"){
      setUser(null);
      setShowAuth(true);
    }
  },[]);
  const[authRequired,setAuthRequired]=useState(LS.get("kmp_authrequired",true));
  useEffect(()=>{LS.set("kmp_authrequired",authRequired)},[authRequired]);

  // ═══ DAVET TOKEN İŞLEME (?invite=xxx&team=yyy) ═══
  useEffect(()=>{
    if(!user?.userId)return; // Önce giriş yapması lazım
    const p=new URLSearchParams(window.location.search);
    const token=p.get("invite");
    const inviteTeamId=p.get("team");
    if(!token||!inviteTeamId)return;
    
    (async()=>{
      const sb=initSupabase();if(!sb)return;
      try{
        // Token'la phantom bul
        const{data:phantom,error}=await sb.from("team_phantom_members").select("*").eq("invite_token",token).eq("team_id",inviteTeamId).maybeSingle();
        if(error||!phantom){
          window.toast.error(lang==="tr"?"Geçersiz davet linki":"Invalid invite link");
          return;
        }
        if(phantom.linked_user_id){
          window.toast.info(lang==="tr"?"Bu davet zaten kullanılmış":"Invite already used");
          return;
        }
        // Phantom'u user'a bağla
        const{error:upErr}=await sb.from("team_phantom_members").update({
          linked_user_id:user.userId,
          linked_at:new Date().toISOString()
        }).eq("id",phantom.id);
        if(upErr){window.toast.error(upErr.message);return;}
        
        // team_members'a ekle
        const{data:teamData}=await sb.from("teams").select("*").eq("id",inviteTeamId).maybeSingle();
        if(teamData){
          await sb.from("team_members").upsert({
            team_id:inviteTeamId,
            user_id:user.userId,
            role:"member",
            position:phantom.position||null
          },{onConflict:"team_id,user_id"});
          setTeam({...teamData,role:"member",inviteCode:teamData.invite_code});
          LS.set("kmp_team",{...teamData,role:"member",inviteCode:teamData.invite_code});
        }
        // Phantom'un yerini gerçek üye alıyor → vardiyalarını da güncelle
        await sb.from("shifts").update({created_by:user.userId,phantom_member_id:null}).eq("phantom_member_id",phantom.id);
        
        // URL'den token'ı temizle
        const newUrl=window.location.pathname;
        window.history.replaceState({},"",newUrl);
        window.toast.success(lang==="tr"?`✓ ${teamData?.name||"Ekibe"} katıldın!`:`✓ Joined ${teamData?.name||"team"}!`);
      }catch(e){
        console.error("Davet işleme hatası:",e);
        window.toast.error(e.message);
      }
    })();
  },[user?.userId]);

  // Supabase oturum kontrolü - sayfa açıldığında mevcut oturumu yükle
  useEffect(()=>{
    // Fallback: 5 saniye geçtiyse zorla devam et (Supabase erişilemez ise)
    const _authTimeout=setTimeout(()=>setAuthChecked(true),5000);
    try{
      const sb=initSupabase();
      if(!sb){clearTimeout(_authTimeout);setAuthChecked(true);return;}
      sb.auth.getSession().then(({data})=>{
        clearTimeout(_authTimeout);
        const isReset=new URLSearchParams(window.location.search).get("mode")==="reset";
        if(!isReset&&data.session&&data.session.user){
          const u={
            email:data.session.user.email,
            name:data.session.user.user_metadata?.name||data.session.user.user_metadata?.full_name||data.session.user.email.split("@")[0],
            verified:!!data.session.user.email_confirmed_at,
            userId:data.session.user.id,
            accessToken:data.session.access_token
          };
          setUser(u);
          LS.set("kmp_user",u);
          // localStorage'da team yoksa Supabase'den otomatik yükle
          if(!LS.get("kmp_team",null)){
            (async()=>{try{
              const uid=data.session.user.id;
              let teamData=null;let memberRole="pro";
              const{data:ownedTeam}=await sb.from("teams").select("*").eq("owner_id",uid).eq("app_type","pro").order("created_at",{ascending:false}).limit(1).single();
              if(ownedTeam){teamData=ownedTeam;memberRole="pro";}
              else{
                const{data:members_raw}=await sb.from("team_members").select("team_id,role,position").eq("user_id",uid).order("joined_at",{ascending:false}).limit(1);
                const membership=members_raw?.[0]||null;
                if(membership?.team_id){
                  const{data:td}=await sb.from("teams").select("*").eq("id",membership.team_id).eq("app_type","pro").single();
                  if(td){teamData=td;memberRole=membership.role||"pro";}
                }
              }
              if(teamData){
                const loadedTeam={...teamData,role:memberRole,inviteCode:teamData.invite_code};
                setTeam(loadedTeam);LS.set("kmp_team",loadedTeam);
                const{data:members}=await sb.from("team_members").select("*").eq("team_id",teamData.id);
                if(members){
                  const uids=members.map(m=>m.user_id);
                  const{data:profs}=await sb.from("profiles").select("id,full_name,email").in("id",uids);
                  const getName=(uid2)=>{const p=(profs||[]).find(p=>p.id===uid2);return p?.full_name||p?.email?.split("@")[0]||uid2;};
                  setTeamMembers(members.map(m=>({userId:m.user_id,name:getName(m.user_id),role:m.role,position:m.position})));
                }
              }
            }catch(e){console.warn("Ekip yüklenemedi (session):",e.message);}})();
          }
        }
        setAuthChecked(true);
      }).catch(()=>{clearTimeout(_authTimeout);setAuthChecked(true);});
      const{data:listener}=sb.auth.onAuthStateChange((event,session)=>{
        if(event==="SIGNED_OUT"){setUser(null);LS.set("kmp_user",null);}
        else if(event==="PASSWORD_RECOVERY"){
          // Şifre sıfırlama linki tıklandı — yeni şifre ekranını göster, giriş yapma
          localStorage.setItem("km_password_recovery","true");
          window.__kmPasswordRecovery=true;
          setShowAuth(true);
        }
        else if((event==="SIGNED_IN"||event==="TOKEN_REFRESHED")&&session?.user){
          // PASSWORD_RECOVERY sonrası SIGNED_IN gelirse atla
          if(window.__kmPasswordRecovery)return;
          const u={
            email:session.user.email,
            name:session.user.user_metadata?.name||session.user.user_metadata?.full_name||session.user.email.split("@")[0],
            verified:!!session.user.email_confirmed_at,
            userId:session.user.id,
            accessToken:session.access_token
          };
          setUser(u);LS.set("kmp_user",u);
          // Ekibi Supabase'den yükle (owner VEYA member)
          (async()=>{try{
            const uid=session.user.id;
            // Önce kendi oluşturduğu team'e bak (owner)
            let teamData=null;let memberRole="pro";
            const{data:ownedTeam}=await sb.from("teams").select("*").eq("owner_id",uid).eq("app_type","pro").order("created_at",{ascending:false}).limit(1).single();
            if(ownedTeam){teamData=ownedTeam;memberRole="pro";}
            else{
              // Üye olduğu team'e bak (herhangi role)
              const{data:members_raw}=await sb.from("team_members").select("team_id,role,position").eq("user_id",uid).order("joined_at",{ascending:false}).limit(1);
              const membership=members_raw?.[0]||null;
              if(membership?.team_id){
                const{data:td}=await sb.from("teams").select("*").eq("id",membership.team_id).eq("app_type","pro").single();
                if(td){teamData=td;memberRole=membership.role||"pro";}
              }
            }
            if(teamData){
              const loadedTeam={...teamData,role:memberRole,inviteCode:teamData.invite_code};
              setTeam(loadedTeam);LS.set("kmp_team",loadedTeam);
              const{data:members}=await sb.from("team_members").select("*").eq("team_id",teamData.id);
              if(members){
                const uids=members.map(m=>m.user_id);
                const{data:profs}=await sb.from("profiles").select("id,full_name,email").in("id",uids);
                const getName=(uid2)=>{const p=(profs||[]).find(p=>p.id===uid2);return p?.full_name||p?.email?.split("@")[0]||uid2;};
                setTeamMembers(members.map(m=>({userId:m.user_id,name:getName(m.user_id),role:m.role,position:m.position})));
              }
            }
          }catch(e){console.warn("Ekip yüklenemedi:",e.message);}})();
        }
      });
      return()=>{listener?.subscription?.unsubscribe()};
    }catch(e){setAuthChecked(true);}
  },[]);

  const handleLogout=async()=>{
    try{
      const sb=initSupabase();
      if(sb)await sb.auth.signOut();
    }catch(e){console.warn("Logout hata:",e)}
    setUser(null);
  };
  const[mainCat,setMC]=useState("all");const[subCat,setSC]=useState("all");
  // Akıllı arama: metin + diyet filtresi + alerjen hariç tutma (madde 3)
  const[search,setSearch]=useState("");
  const[dietFilter,setDF]=useState([]);       // OR mantığı: bunlardan EN AZ biri olsun
  const[showFilters,setShowFilters]=useState(false);
  const[allergenExclude,setAE]=useState([]);  // bunlar OLMAYAN reçeteler
  const[showAdd,setSAdd]=useState(false);
  const[showSettingsDrawer,setShowSettingsDrawer]=useState(false);
  const[chatAttachment,setChatAttachment]=useState(null);
  useEffect(()=>{
    const h=(e)=>setChatAttachment({...e.detail.attachment,_type:e.detail.type});
    window.addEventListener("km-open-attachment",h);
    return()=>window.removeEventListener("km-open-attachment",h);
  },[]);
  const[drawerSection,setDrawerSection]=useState(null);
  const[showAllRecipes,setShowAllRecipes]=useState(false);
  const[activeR,setAR]=useState(null);const[editR,setER]=useState(null);

  // Reçete kaydetme sarmalayıcı: kalori boşsa DB'den tahmin et, sonra uygula
  const saveRecipeWithCalorie=async(recipe)=>{
    let finalRecipe={...recipe};
    // Fotoğraf base64 ise Storage'a yükle
    if(finalRecipe.photo&&finalRecipe.photo.startsWith("data:")){
      try{
        const res=await fetch(finalRecipe.photo);
        const blob=await res.blob();
        const file=new File([blob],`recipe_${finalRecipe.id||Date.now()}.jpg`,{type:"image/jpeg"});
        const uid=user?.userId||user?.id;
        if(uid){
          const uploaded=await uploadFile(file,null,"recipes",null,uid);
          finalRecipe={...finalRecipe,photo:uploaded.url};
        }
      }catch(e){console.warn("Foto upload hatası:",e.message);}
    }
    if(!finalRecipe.calories||finalRecipe.calories===0){
      // 1. Local correction DB
      const est=estimateRecipeCalories(finalRecipe,calorieDB);
      if(est.kcal){
        finalRecipe.calories=est.kcal;
      }else{
        // 2. Global Supabase DB
        let totalKcal=0,totalWeight=0;
        const missing=[];
        for(const ing of (finalRecipe.ingredients||[])){
          const weight=parseAmountToGram(ing.amount);
          if(weight<=0){missing.push(ing);continue;}
          const kcal=await getKcalFromDB(ing.name);
          if(kcal!==null){
            totalKcal+=(kcal*weight)/100;
            totalWeight+=weight;
            saveCalorieToDB(setCalorieDB,ing.name,kcal);
          }else{
            missing.push(ing);
          }
        }
        if(totalWeight>0&&missing.length===0){
          finalRecipe.calories=Math.round((totalKcal/totalWeight)*100);
        }else if(missing.length>0&&apiKey){
          // Sadece eksik malzemeler için AI
          try{
            const ingList=missing.map(i=>`${i.name}: ${i.amount}`).join("\n");
            const raw=await callAI(apiKey,`Nutrition expert. For each ingredient estimate kcal per 100g.
Return ONLY JSON: {"ingredients":[{"name":"...","kcal_per_100g":XXX}]}
Ingredients:\n${ingList}`,"Return JSON only.","haiku");
            const result=parseJSON(raw);
            if(result.ingredients){
              for(const item of result.ingredients){
                if(item.name&&item.kcal_per_100g){
                  saveCalorieToDB(setCalorieDB,item.name,item.kcal_per_100g);
                  contributeToGlobalDB(item.name,item.kcal_per_100g);
                  const w=parseAmountToGram((missing.find(m=>normalizeName(m.name).toLowerCase()===normalizeName(item.name).toLowerCase())||{}).amount||"100g");
                  totalKcal+=(item.kcal_per_100g*w)/100;
                  totalWeight+=w;
                }
              }
              if(totalWeight>0)finalRecipe.calories=Math.round((totalKcal/totalWeight)*100);
            }
          }catch(e){log("calorie-ai",e.message,"")}
        }
      }
    }
    return finalRecipe;
  };

  const t=useMemo(()=>{
    const baseTheme=dark?THEMES.dark:THEMES.light;
    return {...baseTheme,L:I18N[lang]||I18N.tr,lang,rtl:lang==="ar"};
  },[dark,lang]);

  useEffect(()=>{LS.set(SK.key,apiKey)},[apiKey]);
  useEffect(()=>{LS.set(SK.dark,dark);document.body.className=dark?"dark":""},[dark]);
  useEffect(()=>{LS.set(SK.lang,lang);document.documentElement.dir=lang==="ar"?"rtl":"ltr";document.documentElement.lang=lang},[lang]);
  useEffect(()=>{LS.set(SK.recipes,recipes)},[recipes]);
  // Mevcut base64 fotoğrafları Storage'a migrate et
  useEffect(()=>{
    const uid=user?.userId||user?.id;
    if(!uid||!recipes?.length)return;
    const needsMigrate=recipes.filter(r=>r.photo&&r.photo.startsWith("data:"));
    if(needsMigrate.length===0)return;
    (async()=>{
      let updated=[...recipes];
      for(const r of needsMigrate){
        try{
          const res=await fetch(r.photo);
          const blob=await res.blob();
          const file=new File([blob],`recipe_${r.id||Date.now()}.jpg`,{type:"image/jpeg"});
          const uid=user?.userId||user?.id;
          if(!uid)continue;
          const uploaded=await uploadFile(file,null,"recipes",null,uid);
          updated=updated.map(x=>x.id===r.id?{...x,photo:uploaded.url}:x);
        }catch(e){console.warn("Migrate foto:",r.id,e.message);}
      }
      setRecipes(updated);
    })();
  },[team?.id]);
  useEffect(()=>{LS.set(SK.stock,stock)},[stock]);
  useEffect(()=>{LS.set(SK.invoices,invoices)},[invoices]);
  useEffect(()=>{LS.set(SK.menus,menus)},[menus]);
  useEffect(()=>{LS.set(SK.expenses,expenses)},[expenses]);
  useEffect(()=>{LS.set(SK.storage,storageAreas)},[storageAreas]);
  useEffect(()=>{LS.set(SK.productions,productions)},[productions]);
  useEffect(()=>{LS.set(SK.reportCats,reportCats)},[reportCats]);
  useEffect(()=>{LS.set(SK.profile,profile)},[profile]);
  useEffect(()=>{LS.set(SK.traceability,traceability)},[traceability]);
  useEffect(()=>{LS.set(SK.lots,lots)},[lots]);
  useEffect(()=>{LS.set(SK.trackedIngs,trackedIngs)},[trackedIngs]);
  useEffect(()=>{LS.set(SK.resetHour,resetHour)},[resetHour]);
  useEffect(()=>{LS.set(SK.organizations,organizations)},[organizations]);
  useEffect(()=>{LS.set(SK.storageChecks,storageChecks)},[storageChecks]);
  useEffect(()=>{LS.set(SK.menuTemplates,menuTemplates)},[menuTemplates]);
  useEffect(()=>{LS.set(SK.conversations,conversations)},[conversations]);
  useEffect(()=>{LS.set(SK.activeConvId,activeConvId)},[activeConvId]);
  useEffect(()=>{LS.set(SK.botMessages,botMessages)},[botMessages]);
  useEffect(()=>{LS.set(SK.notifSettings,notifSettings)},[notifSettings]);
  useEffect(()=>{LS.set(SK.calorieDB,calorieDB)},[calorieDB]);
  useEffect(()=>{LS.set(SK.printers,printers)},[printers]);

  // ═══ REALTIME SYNC ═══
  const _sbStockTimer=useRef(null);
  const _sbProdTimer=useRef(null);
  const _sbRecipeTimer=useRef(null);
  const _sbTodosTimer=useRef(null);

  useEffect(()=>{
    if(!team?.id||!user?.userId)return;
    if(_sbStockTimer.current)clearTimeout(_sbStockTimer.current);
    _sbStockTimer.current=setTimeout(async()=>{
      await setTeamData(team.id,"team_stock",stock,user.userId);
    },2000);
    return()=>{if(_sbStockTimer.current)clearTimeout(_sbStockTimer.current);};
  },[stock,team?.id,user?.userId]);

  useEffect(()=>{
    if(!team?.id||!user?.userId)return;
    if(_sbProdTimer.current)clearTimeout(_sbProdTimer.current);
    _sbProdTimer.current=setTimeout(async()=>{
      await setTeamData(team.id,"team_productions",productions,user.userId);
    },2000);
    return()=>{if(_sbProdTimer.current)clearTimeout(_sbProdTimer.current);};
  },[productions,team?.id,user?.userId]);

  useEffect(()=>{
    if(!team?.id||!user?.userId)return;
    if(_sbRecipeTimer.current)clearTimeout(_sbRecipeTimer.current);
    _sbRecipeTimer.current=setTimeout(async()=>{
      await setTeamData(team.id,"team_recipes",recipes,user.userId);
    },3000);
    return()=>{if(_sbRecipeTimer.current)clearTimeout(_sbRecipeTimer.current);};
  },[recipes,team?.id,user?.userId]);

  useEffect(()=>{
    if(!team?.id||!user?.userId)return;
    if(_sbTodosTimer.current)clearTimeout(_sbTodosTimer.current);
    _sbTodosTimer.current=setTimeout(async()=>{
      await setTeamData(team.id,"team_todos",todos,user.userId);
    },2000);
    return()=>{if(_sbTodosTimer.current)clearTimeout(_sbTodosTimer.current);};
  },[todos,team?.id,user?.userId]);

  useEffect(()=>{
    if(!team?.id)return;
    const sb=initSupabase();if(!sb)return;
    const channel=sb.channel(`kmp-sync-${team.id}`)
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"team_stock",filter:`team_id=eq.${team.id}`},
        (payload)=>{
          if(payload.new?.updated_by===user?.userId)return;
          if(payload.new?.data)setStock(payload.new.data);
        })
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"team_productions",filter:`team_id=eq.${team.id}`},
        (payload)=>{
          if(payload.new?.updated_by===user?.userId)return;
          if(payload.new?.data)setProductions(payload.new.data);
        })
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"team_recipes",filter:`team_id=eq.${team.id}`},
        (payload)=>{
          if(payload.new?.updated_by===user?.userId)return;
          if(payload.new?.data)setRecipes(payload.new.data);
        })
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"team_todos",filter:`team_id=eq.${team.id}`},
        (payload)=>{
          if(payload.new?.updated_by===user?.userId)return;
          if(payload.new?.data)setTodos(payload.new.data);
        })
      .subscribe();
    return()=>{sb.removeChannel(channel);};
  },[team?.id,user?.userId]);

  // ═══ REALTIME SYNC SONU ═══

  // Bot kontrolü - her dakika çalışır, kontrol edilen saatlerde bot mesajı ekler
  useEffect(()=>{
    if(!notifSettings.enabled)return;
    const check=()=>{
      const now=new Date();
      const today=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`;
      const hour=now.getHours();
      const minute=now.getMinutes();
      const minutesSince=hour*60+minute;
      const addBot=(key,text,icon)=>{
        // Aynı gün + aynı key ile mesaj varsa ekleme (duplicate koruma)
        const exists=botMessages.some(b=>b.key===key+"_"+today);
        if(!exists){
          setBotMessages(p=>[...p,{
            id:"bot_"+Date.now()+Math.random(),
            key:key+"_"+today,
            text,
            icon:icon||"🤖",
            ts:new Date().toISOString()
          }].slice(-50)); // son 50 mesajla sınırla
        }
      };
      // 09:05 / 15:05 / 21:05 depo kontrolü
      if(notifSettings.storageCheck&&traceability&&(storageAreas||[]).length>0){
        const checks=[
          {time:"09.00",trigger:9*60+5,label:"09:00"},
          {time:"15.00",trigger:15*60+5,label:"15:00"},
          {time:"21.00",trigger:21*60+5,label:"21:00"}
        ];
        checks.forEach(c=>{
          if(minutesSince>=c.trigger&&minutesSince<c.trigger+60){
            const missing=(storageAreas||[]).filter(s=>!storageChecks.some(sc=>sc.storageId===s.id&&sc.date===today&&sc.time===c.time));
            if(missing.length>0){
              addBot(`storagecheck_${c.time}`,`Saat ${c.label} — ${missing.length} depo kontrolü henüz yapılmadı: ${missing.map(m=>m.name).join(", ")}`,"⏰");
            }
          }
        });
      }
      // SKT geçen ürün
      if(notifSettings.expiredSKT){
        const expired=productions.filter(p=>p.status==="active"&&new Date(p.expiresAt)<now);
        if(expired.length>0){
          addBot("expired",`${expired.length} ürünün son kullanma tarihi geçti: ${expired.slice(0,3).map(p=>p.recipeName).join(", ")}${expired.length>3?"...":""}`,"⚠️");
        }
      }
      // Düşük stok
      if(notifSettings.lowStock){
        const low=stock.filter(s=>s.qty<=(s.low||100));
        if(low.length>0){
          addBot("lowstock",`${low.length} ürün kritik seviyede: ${low.slice(0,3).map(s=>s.name).join(", ")}${low.length>3?"...":""}`,"📉");
        }
      }
      // Parti no hatırlatma - 22:55
      if(notifSettings.lotReminder&&traceability&&trackedIngs.length>0){
        if(minutesSince>=22*60+55&&minutesSince<23*60){
          const missing=trackedIngs.filter(ing=>!lots[ing]);
          if(missing.length>0){
            addBot("lotreminder",`Gün bitmeden parti numaraları kontrol et! ${missing.length} hammaddede parti no yok.`,"🌙");
          }
        }
      }
    };
    check();
    const id=setInterval(check,60000);
    return()=>clearInterval(id);
  },[notifSettings,traceability,storageAreas,storageChecks,productions,stock,trackedIngs,lots]);

  // Gün sonu otomatik parti no sıfırlama
  useEffect(()=>{
    const checkReset=()=>{
      const now=new Date();
      const lastReset=LS.get("kmp_lastreset",null);
      const today=`${now.getFullYear()}-${now.getMonth()}-${now.getDate()}`;
      if(now.getHours()>=resetHour&&lastReset!==today){
        // Tüm parti no'ları sıfırla
        setLots({});
        LS.set("kmp_lastreset",today);
      }
    };
    checkReset();
    const id=setInterval(checkReset,60000); // her dakika kontrol
    return()=>clearInterval(id);
  },[resetHour]);
  useEffect(()=>{document.body.style.background=t.bg;document.body.style.color=t.text},[t]);

  // Deep link + şifre sıfırlama + paylaşım URL parametreleri
  const[sharedRecipeModal,setSharedRecipeModal]=useState(null);
  useEffect(()=>{
    const params=new URLSearchParams(window.location.search);
    const prodId=params.get("prod");
    const mode=params.get("mode");
    const shareType=params.get("share");
    const shareId=params.get("id");
    const shareName=params.get("name");
    // Şifre sıfırlama
    if(mode==="reset"){setShowAuth(true);}
    // Paylaşılan reçete linki
    if(shareType==="recipe"){
      setSharedRecipeModal({id:shareId,name:decodeURIComponent(shareName||"")});
      window.history.replaceState({},"",window.location.pathname);
    }
    if(prodId&&productions.length>0){
      const found=productions.find(p=>String(p.id)===prodId||p.labelSeq===prodId);
      if(found){
        setTab("production");
        setTimeout(()=>{window.dispatchEvent(new CustomEvent("km-open-prod",{detail:{prod:found}}));},300);
        window.history.replaceState({},"",window.location.pathname);
      }
    }
  },[productions]);

  const filtered=useMemo(()=>recipes.slice().sort((a,b)=>(b.id||0)-(a.id||0)).filter(r=>{
    if(mainCat!=="all"&&r.mainCat!==mainCat)return false;
    if(subCat!=="all"&&r.subCat!==subCat)return false;
    if(search.trim()){
      const q=search.toLowerCase();
      const nameMatch=r.name.toLowerCase().includes(q);
      const ingMatch=r.ingredients.some(i=>i.name.toLowerCase().includes(q));
      const noteMatch=(r.notes||"").toLowerCase().includes(q);
      const dietMatch=(r.diets||[]).some(d=>{const di=DIETS.find(x=>x.id===d);return di&&di.l.toLowerCase().includes(q);});
      const allergenMatch=(r.allergens||[]).some(a=>{const al=ALLERGENS.find(x=>x.id===a);return al&&al.l.toLowerCase().includes(q);});
      const cuisineMatch=(CUISINES.find(c=>c.id===r.cuisine)?.l||"").toLowerCase().includes(q);
      if(!nameMatch&&!ingMatch&&!noteMatch&&!dietMatch&&!allergenMatch&&!cuisineMatch)return false;
    }
    // Diyet filtresi: seçili diyetlerin HEPSI reçetede olmalı
    if(dietFilter.length>0&&!dietFilter.every(d=>(r.diets||[]).includes(d)))return false;
    // Alerjen hariç tutma: seçili alerjenlerden HERHANGİ BİRİ varsa gösterme
    if(allergenExclude.length>0&&allergenExclude.some(a=>(r.allergens||[]).includes(a)))return false;
    return true;
  }),[recipes,mainCat,subCat,search,dietFilter,allergenExclude]);

  // STOKTAN DÜŞME BUG FIX (madde 8):
  // Orijinal bug: parseFloat ile miktar çıkarmak TR sayı formatını dikkate almıyordu
  // ve stok match çok kısıtlıydı — ilk kelime eşleşmesi yeterliydi
  // Fix: hem TR hem EN format, daha iyi ad eşleşmesi, negatife düşmeme
  const deductStock=useCallback((r,m=1)=>{
    setStock(prev=>{
      let next=[...prev];
      for(const ing of r.ingredients){
        const nmMatch=ing.amount.match(/[\d.,]+/);
        if(!nmMatch)continue;
        const amStr=nmMatch[0];
        // normalize: TR virgüllü sayı desteği
        const am=parseFloat(amStr.replace(/\./g,"").replace(",","."))*m;
        if(isNaN(am)||am<=0)continue;
        const ingNameLower=ing.name.toLowerCase();
        const idx=next.findIndex(s=>{
          const sLower=s.name.toLowerCase();
          return ingNameLower.includes(sLower)||sLower.includes(ingNameLower)||
            ingNameLower.split(" ").some(w=>w.length>2&&sLower.includes(w));
        });
        if(idx>=0){
          next=next.map((s,i)=>i===idx?{...s,qty:Math.max(0,Math.round((s.qty-am)*1000)/1000)}:s);
        }
      }
      return next;
    });
  },[]);

  // API key yoksa Ayarlar'a yönlendir, ama uygulamayı engelleme
  // (Ayarlar > Geliştirici sekmesinden manuel girilebilir veya Cloudflare proxy ile sağlanacak)

  const subs=mainCat!=="all"?(SUB_CATS[mainCat]||[]):[];
  const baseTabs=[{id:"recipes",l:t.L.tabRecipes,i:"🍽",icon:"recipes"},{id:"stock",l:t.L.tabStock,i:"📦",icon:"stock"},{id:"production",l:t.L.tabProduction,i:"🍱",icon:"production"}];
  const traceTabs=traceability?[{id:"reports",l:t.L.tabReports,i:"📊",icon:"reports"}]:[];
  const endTabs=[
    {id:"menus",l:t.L.tabMenus,i:"📋",icon:"menus"},
    {id:"kanban",l:"Kanban",i:"📋",icon:"kanban"},
    {id:"chat",l:lang==="tr"?"Sohbet":"Chats",i:"💬",icon:"chat"},
    ...(team?[
      {id:"hub",l:lang==="tr"?"Departmanlar":"Departments",i:"🏢",icon:"hub"},
      {id:"events",l:lang==="tr"?"Etkinlikler":"Events",i:"🎉",icon:"events"},
      {id:"shift",l:lang==="tr"?"Vardiya":"Shifts",i:"🕐",icon:"shift"},
      {id:"botrules",l:lang==="tr"?"Otomasyon":"Automation",i:"⚙️",icon:"automation"},
    ]:[])
  ];
  const tabs=[...baseTabs,...traceTabs,...endTabs];
  const swipeNav=useSwipeNav(tabs,tab,setTab);
  const lowCount=stock.filter(s=>s.qty<=(s.low||100)).length;

  // ═══ SPLASH SCREEN ═══
  if(!authChecked)return <div style={{position:"fixed",inset:0,background:t.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",zIndex:9999,gap:20}}>
    <div style={{width:80,height:80,borderRadius:20,background:t.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:40,boxShadow:"0 8px 24px rgba(200,150,90,0.3)",animation:"kmPulse 1.5s ease-in-out infinite"}}>🍳</div>
    <div style={{fontSize:22,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif"}}>Kitchen Manager</div>
    <div style={{display:"flex",gap:6}}>
      {[0,1,2].map(i=><div key={i} style={{width:8,height:8,borderRadius:"50%",background:t.accent,animation:`kmDot 1s ease-in-out ${i*0.15}s infinite`}}/>)}
    </div>
    <style dangerouslySetInnerHTML={{__html:`@keyframes kmPulse{0%,100%{transform:scale(1);opacity:1}50%{transform:scale(0.95);opacity:0.85}}@keyframes kmDot{0%,100%{transform:translateY(0);opacity:0.3}50%{transform:translateY(-8px);opacity:1}}`}}/>
  </div>;

  return <div {...swipeNav} style={{minHeight:"100vh",paddingBottom:80,
    ...(wallpaper==="custom"&&customWP?{backgroundImage:`url(${customWP})`,backgroundSize:"cover",backgroundPosition:"center",backgroundAttachment:"fixed"}:
    WALLPAPERS.find(w=>w.id===wallpaper)?.id==="default"?{background:t.bg}:
    WALLPAPERS.find(w=>w.id===wallpaper)?.style||{background:t.bg})}}>
    <div style={{background:t.topBar,backdropFilter:"blur(16px)",WebkitBackdropFilter:"blur(16px)",borderBottom:`1px solid ${t.border}`,padding:"0 16px",position:"sticky",top:0,zIndex:100}}>
      <div className="app" style={{display:"flex",alignItems:"center",justifyContent:"space-between",height:54}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <Logo size={32} c={t.accent}/>
          <div>
            <div style={{fontSize:18,fontFamily:"'Fraunces',serif",color:t.text,fontWeight:700,lineHeight:1}}>Kitchen</div>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              <div style={{fontSize:9,color:t.accent,letterSpacing:"0.15em",fontWeight:600}}>MANAGER</div>
              <div style={{fontSize:8,color:"#fff",background:`linear-gradient(135deg,${t.accent} 0%,#8b6332 100%)`,padding:"1px 5px",borderRadius:3,letterSpacing:"0.1em",fontWeight:800}}>PRO</div>
            </div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {tab==="recipes"&&<button onClick={()=>setSAdd(true)} style={{...bSt("p",t),padding:"7px 14px",fontSize:13}}>{t.L.stockAddBtn}</button>}
          <button onClick={()=>setDark(d=>!d)} style={{
            background:t.inBg,border:`1px solid ${t.inBo}`,
            borderRadius:20,padding:"5px 10px",cursor:"pointer",
            fontSize:16,color:t.ts,display:"flex",alignItems:"center",gap:4,
            transition:"all 0.2s"
          }} title={dark?"Light mode":"Dark mode"}>
            {dark?"☀️":"🌙"}
          </button>
          <button onClick={()=>{setShowSettingsDrawer(true);setDrawerSection(null);}} style={{
            background:showSettingsDrawer?t.acB:"transparent",
            border:`1px solid ${showSettingsDrawer?t.accent:t.inBo}`,
            borderRadius:10,padding:"7px 10px",cursor:"pointer",
            fontSize:18,color:showSettingsDrawer?t.accent:t.ts,
            display:"flex",alignItems:"center",justifyContent:"center"
          }} title={t.L.tabSettings}>⚙</button>
        </div>
      </div>
    </div>

    <div className="app" style={{padding:"20px 16px"}}>
      {tab==="recipes"&&<>
        {/* Arama + filtre toggle */}
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={"🔍 "+t.L.search+"..."} style={{...iSt(t),flex:1}}/>
          <button onClick={()=>setShowFilters(!showFilters)} style={{...bSt(showFilters?"p":"s",t),padding:"11px 14px",fontSize:13,whiteSpace:"nowrap"}}>
            🔽 {t.L.filter}{(dietFilter.length+allergenExclude.length)>0?` (${dietFilter.length+allergenExclude.length})`:""}
          </button>
        </div>
        {showFilters&&<div style={{...cSt(t),padding:"12px 14px",marginBottom:10}}>
          <div style={{fontSize:10,color:t.tm,fontWeight:700,marginBottom:6,letterSpacing:"0.1em"}}>DİYET (DAHIL ET)</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:12}}>
            {DIETS.map(d=>{const on=dietFilter.includes(d.id);return <button key={d.id} onClick={()=>setDF(on?dietFilter.filter(x=>x!==d.id):[...dietFilter,d.id])} style={{padding:"5px 11px",borderRadius:14,fontSize:11,fontWeight:600,border:`1px solid ${on?t.success:t.inBo}`,background:on?t.sucBg:"transparent",color:on?t.success:t.tm,cursor:"pointer"}}>{d.icon} {dietL(d,t?.lang||"tr")}</button>})}
          </div>
          <div style={{fontSize:10,color:t.tm,fontWeight:700,marginBottom:6,letterSpacing:"0.1em"}}>ALERJEN (HARİÇ TUT)</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {ALLERGENS.map(a=>{const on=allergenExclude.includes(a.id);return <button key={a.id} onClick={()=>setAE(on?allergenExclude.filter(x=>x!==a.id):[...allergenExclude,a.id])} style={{padding:"5px 11px",borderRadius:14,fontSize:11,fontWeight:600,border:`1px solid ${on?a.c:t.inBo}`,background:on?a.c+"15":"transparent",color:on?a.c:t.tm,cursor:"pointer"}}>{a.icon} {allergenL(a,t?.lang||"tr")}</button>})}
          </div>
        </div>}
        <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:6,marginBottom:8}}>
          {MAIN_CATS.map(c=><button key={c.id} onClick={()=>{setMC(c.id);setSC("all")}} style={{whiteSpace:"nowrap",padding:"7px 14px",borderRadius:20,fontSize:13,fontWeight:500,border:"1px solid",background:mainCat===c.id?t.pA:"transparent",color:mainCat===c.id?t.pAT:t.ts,borderColor:mainCat===c.id?t.pA:t.inBo,cursor:"pointer"}}>{c.icon} {mainCatL(c,t?.lang||"tr")}</button>)}
        </div>
        {subs.length>0&&<div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:6,marginBottom:14}}>
          <button onClick={()=>setSC("all")} style={{whiteSpace:"nowrap",padding:"5px 12px",borderRadius:16,fontSize:12,fontWeight:500,border:`1px solid ${subCat==="all"?t.acBo:t.inBo}`,background:subCat==="all"?t.acB:"transparent",color:subCat==="all"?t.accent:t.tm,cursor:"pointer"}}>Tümü</button>
          {subs.map(s=><button key={s.id} onClick={()=>setSC(s.id)} style={{whiteSpace:"nowrap",padding:"5px 12px",borderRadius:16,fontSize:12,fontWeight:500,border:`1px solid ${subCat===s.id?t.acBo:t.inBo}`,background:subCat===s.id?t.acB:"transparent",color:subCat===s.id?t.accent:t.tm,cursor:"pointer"}}>{s.label}</button>)}
        </div>}
        <div style={{fontSize:12,color:t.tm,marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span>
            {(()=>{
              const isEmpty=!search.trim()&&mainCat==="all"&&subCat==="all"&&dietFilter.length===0&&allergenExclude.length===0;
              if(isEmpty&&!showAllRecipes&&filtered.length>5){
                return `${lang==="tr"?"Son 5 reçete":lang==="en"?"Last 5 recipes":"Son 5"} · ${filtered.length} ${lang==="tr"?"toplam":"total"}`;
              }
              return `${filtered.length} ${lang==="tr"?"reçete":"recipes"}${(dietFilter.length>0||allergenExclude.length>0)?" · "+(lang==="tr"?"filtreli":"filtered"):""}`;
            })()}
          </span>
          {(()=>{
            const isEmpty=!search.trim()&&mainCat==="all"&&subCat==="all"&&dietFilter.length===0&&allergenExclude.length===0;
            if(isEmpty&&filtered.length>5){
              return <button onClick={()=>setShowAllRecipes(!showAllRecipes)} style={{...bSt("s",t),fontSize:11,padding:"4px 10px"}}>
                {showAllRecipes?(lang==="tr"?"← Son 5":lang==="en"?"← Last 5":"← 5"):(lang==="tr"?"Tümünü Göster":lang==="en"?"Show All":"All")}
              </button>;
            }
            return null;
          })()}
        </div>
        {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:t.tm}}>
          <div style={{fontSize:48,opacity:0.4,marginBottom:12}}>🍽</div>
          <div style={{fontSize:16,fontFamily:"'Fraunces',serif"}}>{lang==="tr"?"Reçete bulunamadı":lang==="en"?"No recipes found":"—"}</div>
        </div>:<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
          {(()=>{
            const isEmpty=!search.trim()&&mainCat==="all"&&subCat==="all"&&dietFilter.length===0&&allergenExclude.length===0;
            const toShow=isEmpty&&!showAllRecipes?filtered.slice(0,5):filtered;
            return toShow.map(r=><RCard key={r.id} r={r} onClick={setAR} t={t}/>);
          })()}
        </div>}
      </>}
      {tab==="stock"&&<StockTab stock={stock} setStock={setStock} invoices={invoices} setInvoices={setInvoices} apiKey={apiKey} traceability={traceability} lots={lots} setLots={setLots} trackedIngs={trackedIngs} profile={profile} calorieDB={calorieDB} setCalorieDB={setCalorieDB} t={t}/>}
      {tab==="production"&&<ProductionTab productions={productions} setProductions={setProductions} storageAreas={storageAreas} reportCats={reportCats} setReportCats={setReportCats} profile={profile} traceability={traceability} setTab={setTab} storageChecks={storageChecks} setStorageChecks={setStorageChecks} recipes={recipes} getLabelSeq={getLabelSeq} t={t}/>}
      {tab==="reports"&&<ProductionTab productions={productions} setProductions={setProductions} storageAreas={storageAreas} reportCats={reportCats} setReportCats={setReportCats} profile={profile} traceability={traceability} setTab={setTab} storageChecks={storageChecks} setStorageChecks={setStorageChecks} recipes={recipes} getLabelSeq={getLabelSeq} initialShowReports={true} t={t}/>}
      {tab==="menus"&&<MenuTab menus={menus} setMenus={setMenus} recipes={recipes} menuTemplates={menuTemplates} setMenuTemplates={setMenuTemplates} t={t}/>}
      {tab==="todo"&&<TodoTab todos={todos} setTodos={setTodos} t={t}/>}
      {tab==="hub"&&<HubTab team={team} user={user} t={t}/>}
      {tab==="kanban"&&<KanbanTab team={team} teamMembers={teamMembers} user={user} t={t} profile={profile} isPro={true}/>}
      {tab==="events"&&<EventsTab team={team} teamMembers={teamMembers} user={user} apiKey={apiKey} t={t}/>}
      {tab==="shift"&&<ShiftTab team={team} teamMembers={teamMembers} phantomMembers={phantomMembers} setPhantomMembers={setPhantomMembers} user={user} t={t}/>}
      {tab==="botrules"&&<BotRulesTab team={team} teamMembers={teamMembers} user={user} stock={stock} setBotMessages={setBotMessages} t={t}/>}
      {tab==="chat"&&<WAChatTab team={team} teamMembers={teamMembers} user={user} apiKey={apiKey} t={t} tier="pro"/>}
      {tab==="settings"&&<SettingsTab apiKey={apiKey} setApiKey={setApiKey} dark={dark} setDark={setDark} lang={lang} setLang={setLang} recipes={recipes} stock={stock} invoices={invoices} setRecipes={setRecipes} setStock={setStock} setInvoices={setInvoices} expenses={expenses} setExpenses={setExpenses} storageAreas={storageAreas} setStorageAreas={setStorageAreas} profile={profile} setProfile={setProfile} traceability={traceability} setTraceability={setTraceability} trackedIngs={trackedIngs} setTrackedIngs={setTrackedIngs} resetHour={resetHour} setResetHour={setResetHour} organizations={organizations} setOrganizations={setOrganizations} notifSettings={notifSettings} setNotifSettings={setNotifSettings} printers={printers} setPrinters={setPrinters} setBotMessages={setBotMessages} calorieDB={calorieDB} setCalorieDB={setCalorieDB} user={user} setUser={setUser} authRequired={authRequired} setAuthRequired={setAuthRequired} setShowAuth={setShowAuth} handleLogout={handleLogout} team={team} setTeam={setTeam} teamMembers={teamMembers} setTeamMembers={setTeamMembers} wallpaper={wallpaper} setWallpaper={setWallpaper} customWP={customWP} setCustomWP={setCustomWP} t={t}/>}
    </div>

    <DockTabBar tabs={tabs} tab={tab} setTab={setTab} t={t} ICONS={ICONS} lowCount={lowCount}/>

    {/* SETTINGS DRAWER */}

    {showSettingsDrawer&&<>
      {/* Overlay - tıklayınca kapanır */}
      <div onClick={()=>{setShowSettingsDrawer(false);setDrawerSection(null);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:200,backdropFilter:"blur(2px)"}}/>
      {/* Drawer panel - sağdan kayan */}
      <div style={{
        position:"fixed",top:0,right:0,bottom:0,
        width:Math.min(360,window.innerWidth*0.88)+"px",
        background:t.card,zIndex:201,
        boxShadow:"-4px 0 24px rgba(0,0,0,0.18)",
        overflowY:"auto",
        display:"flex",flexDirection:"column"
      }}>
        {/* Drawer header */}
        <div style={{padding:"16px 18px 12px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,background:t.card,zIndex:1}}>
          {drawerSection?<button onClick={()=>setDrawerSection(null)} style={{background:"none",border:`1px solid ${t.inBo}`,borderRadius:10,padding:"6px 12px",fontSize:13,color:t.accent,cursor:"pointer",fontWeight:600}}>‹ {t.lang==="tr"?"Ayarlar":"Settings"}</button>
          :<span style={{fontSize:18,fontWeight:700,color:t.text,fontFamily:"'Fraunces',serif"}}>⚙ {t.L.tabSettings}</span>}
          <button onClick={()=>{setShowSettingsDrawer(false);setDrawerSection(null);}} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:t.tm,padding:"0 4px"}}>✕</button>
        </div>
        {/* Drawer içerik */}
        <div style={{flex:1,padding:"12px 16px 24px"}}>
          <SettingsTab
            apiKey={apiKey} setApiKey={setApiKey} dark={dark} setDark={setDark}
            lang={lang} setLang={setLang} recipes={recipes} stock={stock}
            invoices={invoices} setRecipes={setRecipes} setStock={setStock}
            setInvoices={setInvoices} expenses={expenses} setExpenses={setExpenses}
            storageAreas={storageAreas} setStorageAreas={setStorageAreas}
            profile={profile} setProfile={setProfile} traceability={traceability}
            setTraceability={setTraceability} trackedIngs={trackedIngs}
            setTrackedIngs={setTrackedIngs} resetHour={resetHour}
            setResetHour={setResetHour} organizations={organizations}
            setOrganizations={setOrganizations} notifSettings={notifSettings}
            setNotifSettings={setNotifSettings} printers={printers}
            setPrinters={setPrinters} setBotMessages={setBotMessages}
            calorieDB={calorieDB} setCalorieDB={setCalorieDB}
            user={user} setUser={setUser} authRequired={authRequired}
            setAuthRequired={setAuthRequired} setShowAuth={setShowAuth}
            handleLogout={handleLogout}
            externalSection={drawerSection} setExternalSection={setDrawerSection}
            team={team} setTeam={setTeam} teamMembers={teamMembers} setTeamMembers={setTeamMembers}
           
            wallpaper={wallpaper} setWallpaper={setWallpaper} customWP={customWP} setCustomWP={setCustomWP}
            t={t}
          />
        </div>
      </div>
    </>}

    {showAuth&&<div style={{position:"fixed",inset:0,zIndex:9999}}><AuthModal onClose={()=>setShowAuth(false)} onLogin={u=>{setUser(u);setShowAuth(false);}} t={t}/></div>}

    {/* Sohbetten açılan içerik modal */}
    {chatAttachment&&<div onClick={()=>setChatAttachment(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:"20px 20px 0 0",padding:20,width:"100%",maxWidth:520,maxHeight:"80vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontSize:16,fontWeight:700,color:t.text}}>
            {chatAttachment._type==="recipe"?"🍽":chatAttachment._type==="menu"?"📋":"📊"} {chatAttachment.name}
          </div>
          <button onClick={()=>setChatAttachment(null)} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:t.tm}}>✕</button>
        </div>
        {chatAttachment._type==="recipe"&&chatAttachment.data&&<>
          {chatAttachment.data.calories&&<div style={{fontSize:12,color:t.tm,marginBottom:10}}>🔥 {chatAttachment.data.calories} kcal</div>}
          {(chatAttachment.data.ingredients||[]).length>0&&<>
            <div style={{fontSize:13,fontWeight:700,color:t.text,marginBottom:6}}>{lang==="tr"?"Malzemeler":"Ingredients"}</div>
            {chatAttachment.data.ingredients.map((ing,i)=><div key={i} style={{fontSize:13,color:t.ts,padding:"4px 0",borderBottom:`1px solid ${t.border}`,display:"flex",justifyContent:"space-between"}}>
              <span>{ing.name}</span><span style={{color:t.tm}}>{ing.amount}</span>
            </div>)}
          </>}
          {(chatAttachment.data.steps||[]).length>0&&<>
            <div style={{fontSize:13,fontWeight:700,color:t.text,marginTop:12,marginBottom:6}}>{lang==="tr"?"Hazırlanışı":"Instructions"}</div>
            {chatAttachment.data.steps.map((step,i)=><div key={i} style={{fontSize:13,color:t.ts,padding:"6px 0",borderBottom:`1px solid ${t.border}`}}>
              <span style={{fontWeight:700,color:t.accent}}>{i+1}.</span> {step}
            </div>)}
          </>}
        </>}
        {chatAttachment._type==="menu"&&<div style={{fontSize:13,color:t.ts}}>{chatAttachment.desc}</div>}
        {chatAttachment._type==="report"&&<div>
          <div style={{fontSize:13,color:t.ts,marginBottom:10}}>{chatAttachment.desc}</div>
          <div style={{fontSize:12,color:t.tm}}>{lang==="tr"?"Raporu görüntülemek için Raporlar sekmesine gidin.":"Go to Reports tab to view the full report."}</div>
          <button onClick={()=>{setTab("reports");setChatAttachment(null);}} style={{...bSt("p",t),width:"100%",marginTop:12}}>
            {lang==="tr"?"Raporlara Git":"Go to Reports"}
          </button>
        </div>}
      </div>
    </div>}
    {authChecked&&authRequired&&!user&&<AuthModal onLogin={u=>{setUser(u);}} t={t}/>}

    {/* Paylaşılan reçete linki - giriş yapılmamışsa kayıt ol */}
    {sharedRecipeModal&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:t.card,borderRadius:20,padding:24,width:"100%",maxWidth:400,textAlign:"center"}}>
        <div style={{fontSize:36,marginBottom:12}}>🍽</div>
        <div style={{fontSize:18,fontWeight:700,color:t.text,marginBottom:8}}>{sharedRecipeModal.name}</div>
        <div style={{fontSize:13,color:t.tm,marginBottom:20,lineHeight:1.5}}>
          {lang==="tr"
            ?"Bu reçeteyi görüntülemek için Kitchen Manager'a giriş yapın veya kayıt olun."
            :"Sign in or create an account to view this recipe on Kitchen Manager."}
        </div>
        {user?<>
          <button onClick={()=>{
            const r=recipes.find(x=>String(x.id)===sharedRecipeModal.id);
            if(r)setAR(r);
            setSharedRecipeModal(null);
          }} style={{...bSt("p",t),width:"100%",marginBottom:8}}>
            {lang==="tr"?"Reçeteyi Aç":"Open Recipe"}
          </button>
          <button onClick={()=>setSharedRecipeModal(null)} style={{...bSt("s",t),width:"100%"}}>{lang==="tr"?"Kapat":"Close"}</button>
        </>:<>
          <button onClick={()=>{setShowAuth(true);}} style={{...bSt("p",t),width:"100%",marginBottom:8}}>
            {lang==="tr"?"Giriş Yap / Kayıt Ol":"Sign In / Sign Up"}
          </button>
          <button onClick={()=>setSharedRecipeModal(null)} style={{...bSt("s",t),width:"100%"}}>{lang==="tr"?"Kapat":"Close"}</button>
        </>}
      </div>
    </div>}

    {showAdd&&<AddModal onClose={()=>setSAdd(false)} onAdd={async r=>{const withCal=await saveRecipeWithCalorie(r);setRecipes(p=>[withCal,...p])}} apiKey={apiKey} t={t}/>}
    {activeR&&!editR&&<Detail r={activeR} onClose={()=>setAR(null)} onDel={id=>{setRecipes(p=>p.filter(r=>r.id!==id));setAR(null)}} onEdit={async r=>{if(r._photoOnly){const clean={...r};delete clean._photoOnly;if(clean.photo&&clean.photo.startsWith("data:")){const uid=user?.userId||user?.id;if(uid){try{const res=await fetch(clean.photo);const blob=await res.blob();const file=new File([blob],`recipe_${clean.id||Date.now()}.jpg`,{type:"image/jpeg"});const uploaded=await uploadFile(file,null,"recipes",null,uid);clean.photo=uploaded.url;}catch(e){console.warn("Foto upload:",e.message);}}}setRecipes(p=>p.map(x=>x.id===clean.id?clean:x));setAR(clean);}else{setER(r);setAR(null);}}} onDeduct={deductStock} stock={stock} expenses={expenses} storageAreas={storageAreas} productions={productions} setProductions={setProductions} profile={profile} organizations={organizations} lots={lots} trackedIngs={trackedIngs} traceability={traceability}
      team={team}
      onShareToChat={(att)=>{
        // Sohbet sekmesine geç ve mesaj gönder
        const sb=initSupabase();
        if(sb&&team?.id){
          sb.auth.getSession().then(({data:{session}})=>{
            const uid=session?.user?.id;
            const uname=session?.user?.user_metadata?.name||session?.user?.email?.split("@")[0]||"?";
            const profile2=JSON.parse(localStorage.getItem("kmp_profile")||"{}");
            sb.from("team_messages").insert({
              team_id:team.id,user_id:uid,user_name:uname,
              user_role:profile2.role||"",
              type:"recipe",text:`🍽 ${att.name}`,attachment:att
            });
          });
        }
        setTab("chat");
        setAR(null);
      }}
      t={t}/>}
    {editR&&<div style={mOv(t)} onClick={()=>setER(null)}><div onClick={e=>e.stopPropagation()} style={mPn(t)}>
      <h3 style={{fontSize:22,marginBottom:18,color:t.text}}>Düzenle</h3>
      <EditForm init={editR} onSave={async f=>{const withCal=await saveRecipeWithCalorie(f);setRecipes(p=>p.map(r=>r.id===withCal.id?withCal:r));setER(null);setAR(null)}} onCancel={()=>setER(null)} t={t}/>
    </div></div>}
  </div>;
}

// Hata yakalayıcı - beyaz ekran olursa hata mesajı göster
class ErrorBoundary extends React.Component{
  constructor(props){super(props);this.state={hasError:false,error:null,info:null}}
  static getDerivedStateFromError(error){return{hasError:true,error}}
  componentDidCatch(error,info){
    this.setState({info});
    console.error("App hatası:",error,info);
  }
  render(){
    if(this.state.hasError){
      return React.createElement("div",{style:{padding:"30px 20px",fontFamily:"system-ui,Arial",maxWidth:600,margin:"40px auto",background:"#fee",border:"1px solid #c00",borderRadius:12,color:"#000"}},
        React.createElement("h2",{style:{color:"#c00",fontSize:18,marginBottom:12}},"⚠ Uygulama Hatası"),
        React.createElement("p",{style:{fontSize:14,marginBottom:14,lineHeight:1.5}},"Bir hata oluştu. Çözüm yolları:"),
        React.createElement("ol",{style:{fontSize:13,paddingLeft:20,lineHeight:1.7,marginBottom:16}},
          React.createElement("li",null,"Sayfayı yenile (Cmd+R / Ctrl+R)"),
          React.createElement("li",null,"Internet bağlantını kontrol et"),
          React.createElement("li",null,"Tarayıcı önbelleğini temizle")
        ),
        React.createElement("details",{style:{marginTop:14}},
          React.createElement("summary",{style:{cursor:"pointer",color:"#c00",fontSize:12}},"Teknik detay"),
          React.createElement("pre",{style:{fontSize:11,background:"#fff",padding:10,marginTop:8,borderRadius:6,overflow:"auto",maxHeight:200}},String(this.state.error?.message||this.state.error||"Bilinmeyen hata")+"\n\n"+(this.state.info?.componentStack||""))
        ),
        React.createElement("button",{
          onClick:()=>{LS.set("kmp_user",null);location.reload()},
          style:{marginTop:14,padding:"10px 16px",background:"#c00",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}
        },"Yenile (oturumu temizle)")
      );
    }
    return this.props.children;
  }
}

