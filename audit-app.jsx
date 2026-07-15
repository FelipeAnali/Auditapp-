import React, { useState, useCallback, useRef, useMemo, Component } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, Cell, ComposedChart, Area, Line, LineChart, PieChart, Pie } from "recharts";

/* ═══ CONSTANTS ═══ */
const CL = { pri: "#1a6b3c", priDk: "#0d4a28", priLt: "#e6f4ec", dan: "#e74c3c", warn: "#f39c12", mut: "#95a5a6", bg: "#f0f4f1", card: "#fff", txt: "#1a1a2e", txtL: "#5a6672" };
const RC_TOP=["#FFD700","#C0C0C0","#CD7F32"];
function rColor(rank,total){if(rank<=3)return RC_TOP[rank-1];const pct=(rank-4)/(Math.max(total-4,1));if(pct<.25)return"#2ecc71";if(pct<.5)return"#3498db";if(pct<.75)return"#f39c12";return"#e74c3c";}
const DIAS = ["Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado"];
const DC = ["Dom","Lun","Mar","Mie","Jue","Vie","Sab"];
const EXCEL_EPOCH=25569; // Days between Excel epoch (1900-01-01) and Unix epoch (1970-01-01)
const MS_PER_DAY=86400000;
const MS_PER_HOUR=3600000;
const MIN_HOURS_PER_DAY=1; // Minimum hours credited per working day
const OUTLIER_THRESHOLD=0.4; // Day F/H below 40% of personal avg = outlier
const KPI_CUMPLE_OFFSET=2; // Default: Cumple = avg + 2
function rankCajeros(cajeros,avg,kpiCfg){cajeros.sort((a,b)=>b.pfH-a.pfH);cajeros.forEach((c,i)=>{c.rank=i+1;c.frase=getPhrase(c.rank,cajeros.length);c.kpi=getKPI(c.pfH,avg,kpiCfg);});return cajeros;}

function getPhrase(r,t){if(r===1)return{em:"🏆",tx:"Mejor cajero/a de la sede.",cl:"#FFD700"};if(r<=3)return{em:"👏",tx:"Excelente!",cl:"#2ecc71"};if(r<=6)return{em:"💪",tx:"Muy buen trabajo!",cl:"#27ae60"};if(r<=Math.ceil(t*.55))return{em:"👍",tx:"Cumple expectativas basicas.",cl:"#3498db"};if(r<=Math.ceil(t*.77))return{em:"⚠️",tx:"Por debajo. Requiere seguimiento.",cl:"#f39c12"};if(r<t)return{em:"🔻",tx:"Muy por debajo. Plan de mejora.",cl:"#e67e22"};return{em:"🔻",tx:"Ultimo. Intervencion inmediata.",cl:"#e74c3c"};}
function getKPI(fh,avg,custom){if(custom){if(fh>=custom.cumple)return{lab:"Cumple",ic:"✅",cl:"#2ecc71"};if(fh>=custom.enProm)return{lab:"En promedio",ic:"😐",cl:"#f39c12"};return{lab:"No cumple",ic:"❌",cl:"#e74c3c"};}if(fh>=avg+2)return{lab:"Cumple",ic:"✅",cl:"#2ecc71"};if(fh>=avg)return{lab:"En promedio",ic:"😐",cl:"#f39c12"};return{lab:"No cumple",ic:"❌",cl:"#e74c3c"};}
function recalcKPIs(data,kpiCfg){const c=kpiCfg&&kpiCfg.active?kpiCfg:null;return{...data,kpiCfg:c,cajeros:data.cajeros.map(x=>({...x,kpi:getKPI(x.pfH,data.avg,c)}))};}
function kpiT(data){const c=data.kpiCfg;if(c)return{cumple:c.cumple,enProm:c.enProm,custom:true};return{cumple:data.avg+KPI_CUMPLE_OFFSET,enProm:data.avg,custom:false};}

function parseDate(v){if(!v)return null;if(v instanceof Date)return isNaN(v.getTime())?null:v;if(typeof v==="number"){const d=new Date((v-EXCEL_EPOCH)*MS_PER_DAY);return isNaN(d.getTime())?null:d;}
  let s=String(v).trim();
  // Handle Spanish AM/PM: "a. m." → "AM", "p. m." → "PM"  
  s=s.replace(/a\.\s*m\./gi,"AM").replace(/p\.\s*m\./gi,"PM");
  // Try DD/MM/YYYY [H:MM[:SS] [AM/PM]]
  const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i);
  if(m){let hr=parseInt(m[4]||"0",10);const mn=parseInt(m[5]||"0",10),sc=parseInt(m[6]||"0",10);
    if(m[7]){const pm=m[7].toUpperCase()==="PM";if(pm&&hr<12)hr+=12;if(!pm&&hr===12)hr=0;}
    return new Date(+m[3],+m[2]-1,+m[1],hr,mn,sc);}
  // Try ISO-like: YYYY-MM-DD[T| ]HH:MM:SS
  s=s.replace(/^(\d{4}-\d{2}-\d{2})\s+(\d)/,"$1T$2");
  const d=new Date(s);return isNaN(d.getTime())?null:d;}
const fD=d=>d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}`:"";
const fDF=d=>d?`${fD(d)}/${d.getFullYear()}`:"";
const fN=(n,dc=2)=>typeof n==="number"?n.toFixed(dc):"---";
const fH=h=>`${Math.floor(h)}h ${Math.round((h-Math.floor(h))*60)}m`;
const sN=f=>{if(!f)return"";const p=f.trim().split(/\s+/);if(p.length>=3)return`${p[2]} ${p[0]}`;return f;};
const esc=s=>String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

/* Consistency: coefficient of variation */
function calcCV(arr){if(arr.length<2)return{cv:0,label:"N/A",cl:"#95a5a6"};const mean=arr.reduce((a,b)=>a+b,0)/arr.length;const sd=Math.sqrt(arr.reduce((s,v)=>s+Math.pow(v-mean,2),0)/arr.length);const cv=mean>0?(sd/mean)*100:0;if(cv<=15)return{cv,label:"Muy estable",cl:"#2ecc71"};if(cv<=30)return{cv,label:"Estable",cl:"#3498db"};if(cv<=50)return{cv,label:"Variable",cl:"#f39c12"};return{cv,label:"Muy variable",cl:"#e74c3c"};}

/* ═══ PROCESS SCHEDULE ═══ */
const BREAK_THRESHOLD=20; // Minutes: breaks > this are deducted, <= are working time
function processSchedule(jr){
  const hd=jr[0].map(h=>String(h||"").trim().toUpperCase());
  // Smart column finder: exact match first, then startsWith, then includes
  const fc=ks=>{for(const k of ks){const i=hd.findIndex(h=>h===k);if(i>=0)return i;}for(const k of ks){const i=hd.findIndex(h=>h.startsWith(k));if(i>=0)return i;}for(const k of ks){const i=hd.findIndex(h=>h.includes(k));if(i>=0)return i;}return-1;};
  const cFunc=fc(["FUNCION"]),cHora=fc(["HORA"]),cId=fc(["IDENTIFICACION","CEDULA","DOCUMENTO"]),cFecha=fc(["FECHA"]),cEmp=fc(["EMPLEADO","NOMBRE"]);
  if(cFunc>=0&&cHora>=0&&cId>=0){
    const byEmp={};
    for(let i=1;i<jr.length;i++){const r=jr[i];if(!r)continue;
      const ced=String(r[cId]||"").trim(),func=String(r[cFunc]||"").trim().toUpperCase(),nombre=cEmp>=0?String(r[cEmp]||"").trim():"";
      if(!ced||!func)continue;
      let hora=r[cHora];if(!hora)continue;
      let mins=0;
      if(typeof hora==="object"&&hora.getHours){mins=hora.getHours()*60+hora.getMinutes();}
      else if(typeof hora==="number"){const frac=hora>1?hora-Math.floor(hora):hora;mins=Math.round(frac*24*60);}
      else{const s=String(hora).trim();const p=s.split(":");if(p.length>=2)mins=(parseInt(p[0],10)||0)*60+(parseInt(p[1],10)||0);else mins=0;}
      let fechaRaw=cFecha>=0?String(r[cFecha]||"").trim():"";if(!fechaRaw)continue;
      // Normalize to YYYY-MM-DD: handles "2.02.2026", "2/02/2026", "2026-02-02", "2026.02.02"
      let fecha;const fm=fechaRaw.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{4})/);
      if(fm){fecha=`${fm[3]}-${fm[2].padStart(2,"0")}-${fm[1].padStart(2,"0")}`;}
      else{const fm2=fechaRaw.match(/^(\d{4})[.\/\-](\d{1,2})[.\/\-](\d{1,2})/);fecha=fm2?`${fm2[1]}-${fm2[2].padStart(2,"0")}-${fm2[3].padStart(2,"0")}`:fechaRaw;}
      const isFallido=func.includes("FALLID");
      if(!byEmp[ced])byEmp[ced]={nombre,dias:{}};
      if(!byEmp[ced].dias[fecha])byEmp[ced].dias[fecha]=[];
      byEmp[ced].dias[fecha].push({func,mins,fallido:isFallido});
      if(nombre&&!byEmp[ced].nombre)byEmp[ced].nombre=nombre;
    }
    const map={};
    for(const[ced,emp]of Object.entries(byEmp)){
      let totalHrs=0,totalDias=0,diasFallidos=0;const byDay={};
      for(const[fecha,events]of Object.entries(emp.dias)){
        events.sort((a,b)=>a.mins-b.mins);
        const hasFallido=events.some(e=>e.fallido);
        if(hasFallido)diasFallidos++;
        const valid=events.filter(e=>!e.fallido);
        if(valid.length<2){byDay[fecha]={hrsMar:0,hasFallido,bruto:0,breakMin:0};continue;}
        const entrada=valid.find(e=>e.func==="ENTRADA");
        const salida=[...valid].reverse().find(e=>e.func==="SALIDA");
        if(!entrada||!salida||salida.mins<=entrada.mins){byDay[fecha]={hrsMar:0,hasFallido,bruto:0,breakMin:0};continue;}
        const bruto=salida.mins-entrada.mins;
        // New break rule: only deduct breaks > BREAK_THRESHOLD minutes
        let breakMins=0;
        for(let k=0;k<valid.length;k++){
          if(valid[k].func.startsWith("SALIDA A")){
            const llegada=valid.slice(k+1).find(e=>e.func.startsWith("LLEGADA"));
            if(llegada&&llegada.mins>valid[k].mins){
              const bk=llegada.mins-valid[k].mins;
              if(bk>BREAK_THRESHOLD)breakMins+=bk; // Only deduct long breaks
            }
          }
        }
        const neto=Math.max(0,bruto-breakMins);
        const hrsDay=neto/60;
        byDay[fecha]={hrsMar:Math.round(hrsDay*100)/100,hasFallido,bruto:Math.round(bruto/60*100)/100,breakMin:breakMins};
        if(!hasFallido&&neto>0){totalHrs+=hrsDay;totalDias++;}
      }
      map[ced]={hrs:Math.round(totalHrs*100)/100,dias:totalDias,diasFallidos,nombre:emp.nombre,byDay};
    }
    return{byCed:true,data:map};
  }
  const cn=fc(["NOMBRE","CAJERO","COLABORADOR","EMPLEADO"]),ch=fc(["HORA","HORAS","HOURS","HRS"]),cd=fc(["DIA","DIAS","DAYS"]);
  if(cn<0||ch<0)return null;
  const map={};for(let i=1;i<jr.length;i++){const r=jr[i];if(!r||!r[cn])continue;const nm=String(r[cn]).trim().toUpperCase(),hrs=parseFloat(r[ch]),dias=cd>=0?parseInt(r[cd]):null;if(nm&&!isNaN(hrs)){if(!map[nm])map[nm]={hrs:0,dias:null,diasFallidos:0,byDay:{}};map[nm].hrs+=hrs;if(dias&&!isNaN(dias))map[nm].dias=(map[nm].dias||0)+dias;}}
  return{byCed:false,data:map};
}

/* ═══ PROCESS DATA ═══ */
function processData(jr,schedMap,bdP){
  const hd=jr[0].map(h=>String(h||"").trim().toUpperCase());
  const fc=ks=>{for(const k of ks){const idx=hd.findIndex(h=>h.includes(k));if(idx>=0)return idx;}return-1;};
  const cols={sede:fc(["DESC. C.O","DESC C.O","SEDE"]),nro:fc(["NRO DOCUMENTO","NRO DOC","FACTURA","DOCUMENTO"]),vend:fc(["VENDEDOR","CEDULA"]),caj:fc(["NOMBRE CAJERO","NOMBRE CAJ","NOM CAJERO","CAJERO"]),fec:fc(["FECHA CREACI","FECHA CREAC"]),tpv:fc(["T.P.V","TPV","CAJA","PUNTO DE VENTA"])};
  if(cols.caj<0||cols.fec<0)throw new Error("Columnas 'Nombre cajero' y 'Fecha creacion' no encontradas.");
  const map={};
  function extractHour(raw,fec){
    if(raw!=null){let s=String(raw);s=s.replace(/a\.\s*m\./gi,"AM").replace(/p\.\s*m\./gi,"PM");
      const m=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
      if(m){let h=parseInt(m[1],10);if(m[4]){const pm=m[4].toUpperCase()==="PM";if(pm&&h<12)h+=12;if(!pm&&h===12)h=0;}if(h>=0&&h<24)return h;}}
    if(typeof raw==="number"&&raw>0){const frac=raw-Math.floor(raw);if(frac>0)return Math.floor(frac*24);}
    if(fec&&!isNaN(fec.getTime()))return fec.getHours();
    return 0;
  }
  const tpvMap={};
  for(let i=1;i<jr.length;i++){const r=jr[i];if(!r||!r[cols.caj])continue;const nm=String(r[cols.caj]).trim();if(!nm||nm==="#N/D")continue;
    const ced=cols.vend>=0?String(r[cols.vend]||"").trim():"",sede=cols.sede>=0?String(r[cols.sede]||"").trim():"",doc=cols.nro>=0?String(r[cols.nro]||"").trim():`d${i}`,raw=r[cols.fec],fec=parseDate(raw);if(!fec)continue;
    const dk=`${fec.getFullYear()}-${String(fec.getMonth()+1).padStart(2,"0")}-${String(fec.getDate()).padStart(2,"0")}`;
    const hora=extractHour(raw,fec);const tpv=cols.tpv>=0?String(r[cols.tpv]||"").trim():"";
    if(!map[nm])map[nm]={nombre:nm,cedula:ced,sede,regs:[]};map[nm].regs.push({doc,fec,dk,hora,ds:fec.getDay(),tpv});
    if(tpv){if(!tpvMap[tpv])tpvMap[tpv]={regs:0,facs:new Set(),cajeros:new Set(),dias:new Set(),horas:{}};tpvMap[tpv].regs++;tpvMap[tpv].facs.add(doc);tpvMap[tpv].cajeros.add(nm);tpvMap[tpv].dias.add(dk);if(!tpvMap[tpv].horas[hora])tpvMap[tpv].horas[hora]=0;tpvMap[tpv].horas[hora]++;}}
  const cajeros=Object.values(map).map(c=>{
    const byDay={};c.regs.forEach(r=>{if(!byDay[r.dk])byDay[r.dk]={facs:new Set(),regs:0,ts:[],byH:{}};byDay[r.dk].facs.add(r.doc);byDay[r.dk].regs++;byDay[r.dk].ts.push(r.fec.getTime());if(!byDay[r.dk].byH[r.hora])byDay[r.dk].byH[r.hora]={regs:0};byDay[r.dk].byH[r.hora].regs++;});
    let tH=0;const pH={};for(let h=0;h<24;h++)pH[h]={regs:0,dias:0};const pS={};for(let d=0;d<7;d++)pS[d]={regs:0,facs:new Set(),dias:0,horas:0};
    const dias=Object.entries(byDay).sort(([a],[b])=>a.localeCompare(b)).map(([dk,dy])=>{const mn=Math.min(...dy.ts),mx=Math.max(...dy.ts);let hrs=(mx-mn)/MS_PER_HOUR;const minH=hrs<MIN_HOURS_PER_DAY;if(minH)hrs=MIN_HOURS_PER_DAY;tH+=hrs;const facs=dy.facs.size,regs=dy.regs,dO=new Date(dk+"T12:00:00"),ds=dO.getDay();
      Object.entries(dy.byH).forEach(([h,v])=>{pH[h].regs+=v.regs;pH[h].dias++;});pS[ds].regs+=regs;dy.facs.forEach(f=>pS[ds].facs.add(f));pS[ds].dias++;pS[ds].horas+=hrs;
      return{fecha:dk,fechaD:fD(dO),fechaF:fDF(dO),diaSem:DC[ds],diaSemL:DIAS[ds],facs,regs,hrs,fH:facs/hrs,rH:regs/hrs,rF:facs>0?regs/facs:0,minH,hi:new Date(mn).toLocaleTimeString("es-CO",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),hf:new Date(mx).toLocaleTimeString("es-CO",{hour:"2-digit",minute:"2-digit",second:"2-digit"})};});
    const tF=dias.reduce((s,d)=>s+d.facs,0),tR=c.regs.length;
    const mD=dias.length>0?dias.reduce((a,b)=>a.fH>b.fH?a:b):null,pDy=dias.length>0?dias.reduce((a,b)=>a.fH<b.fH?a:b):null;
    const hA=Object.entries(pH).filter(([,v])=>v.regs>0).map(([h,v])=>({hora:+h,horaStr:`${h}:00`,regs:v.regs,dias:v.dias,prom:v.dias>0?v.regs/v.dias:0})).sort((a,b)=>b.regs-a.regs);
    const dA=Object.entries(pS).filter(([,v])=>v.dias>0).map(([d,v])=>({dia:DC[d],regs:v.regs,facs:v.facs.size,dias:v.dias,horas:v.horas,fH:v.horas>0?v.facs.size/v.horas:0}));
    const consist=calcCV(dias.map(d=>d.fH));
    let hrsHor=null,diasProg=null,diasFallidos=null,schedByDay=null;
    if(schedMap){
      let found=null;
      if(schedMap.byCed){
        const ced=c.cedula.replace(/\D/g,"");
        // Try cedula match first
        if(ced){const match=Object.entries(schedMap.data).find(([k])=>k.replace(/\D/g,"")===ced);if(match)found=match[1];}
        // Fallback: name match when no cedula or no cedula match
        if(!found){const nu=c.nombre.toUpperCase();found=Object.values(schedMap.data).find(v=>v.nombre&&v.nombre.toUpperCase()===nu);
          if(!found){const pts=nu.split(/\s+/);found=Object.values(schedMap.data).find(v=>{const vn=v.nombre?v.nombre.toUpperCase():"";return pts.filter(p=>p.length>3).every(p=>vn.includes(p));});}
        }
      }else{
        const sd=schedMap.data;const nu=c.nombre.toUpperCase();
        found=sd[nu];if(!found){const ks=Object.keys(sd),pts=nu.split(/\s+/),mt=ks.find(k=>pts.some(p=>p.length>3&&k.includes(p)));if(mt)found=sd[mt];}
      }
      if(found){hrsHor=found.hrs;diasProg=found.dias;diasFallidos=found.diasFallidos||0;schedByDay=found.byDay||null;}
    }
    // Merge marcación data per day
    let totalHrsMar=0,diasMarValid=0;
    dias.forEach(d=>{
      if(schedByDay){
        // Try matching by fecha key (normalize both sides)
        const mk=Object.keys(schedByDay).find(k=>k===d.fecha||k.includes(d.fecha)||d.fecha.includes(k));
        if(mk){const md=schedByDay[mk];d.hrsMar=md.hrsMar;d.hasFallido=md.hasFallido;d.breakMin=md.breakMin;
          if(!md.hasFallido&&md.hrsMar>0){totalHrsMar+=md.hrsMar;diasMarValid++;}
        }else{d.hrsMar=null;d.hasFallido=false;d.breakMin=0;}
      }else{d.hrsMar=null;d.hasFallido=false;d.breakMin=0;}
      d.rHMar=d.hrsMar&&d.hrsMar>0?d.regs/d.hrsMar:null;
      d.fHMar=d.hrsMar&&d.hrsMar>0?d.facs/d.hrsMar:null;
    });
    const diasNoLab=diasProg?Math.max(0,diasProg-dias.length):null;
    const bdInfo=bdP?bdP[c.nombre.toUpperCase()]||null:null;
    const activo=bdInfo?bdInfo.activo:null;const cargo=bdInfo?bdInfo.cargo:"";const ccosto=bdInfo?bdInfo.ccosto:"";
    const rHMar=totalHrsMar>0?tR/totalHrsMar:0; // Registros per marcación hour (RANKING METRIC)
    const fHMar=totalHrsMar>0?tF/totalHrsMar:0;
    return{nombre:c.nombre,cedula:c.cedula,sede:c.sede,tF,tR,numDias:dias.length,hrsEfec:tH,hrsHor,totalHrsMar,diasMarValid,diasProg,diasNoLab,diasFallidos,activo,cargo,ccosto,pfH:tH>0?tF/tH:0,prH:tH>0?tR/tH:0,prF:tF>0?tR/tF:0,rHMar,fHMar,dias,mD,pDy,hA,dA,consist,rank:0,frase:null,kpi:null};});
  // RANKING: if marcaciones loaded, rank by R/HMar; else by F/HEfec
  const hasSched=cajeros.some(c=>c.hrsHor!==null);
  if(hasSched){cajeros.sort((a,b)=>b.rHMar-a.rHMar);}else{cajeros.sort((a,b)=>b.pfH-a.pfH);}
  const avg=cajeros.length>0?cajeros.reduce((s,c)=>s+c.pfH,0)/cajeros.length:0;
  const avgR=cajeros.length>0?cajeros.reduce((s,c)=>s+c.prH,0)/cajeros.length:0;
  const avgRHMar=hasSched&&cajeros.length>0?cajeros.reduce((s,c)=>s+c.rHMar,0)/cajeros.length:0;
  rankCajeros(cajeros,avg,null);
  const allDates=[...new Set(cajeros.flatMap(c=>c.dias.map(d=>d.fecha)))].sort();
  const dS=allDates.map(dk=>{let f=0,r=0,h=0,a=0;cajeros.forEach(c=>{const d=c.dias.find(x=>x.fecha===dk);if(d){f+=d.facs;r+=d.regs;h+=d.hrs;a++;}});return{fecha:dk,fechaD:fD(new Date(dk+"T12:00:00")),facs:f,regs:r,hrs:h,activos:a,fH:h>0?f/h:0};});
  const tpvStats=Object.entries(tpvMap).map(([tpv,v])=>{const peakH=Object.entries(v.horas).sort((a,b)=>b[1]-a[1])[0];return{tpv,regs:v.regs,facs:v.facs.size,cajeros:v.cajeros.size,dias:v.dias.size,promRegDia:Math.round(v.regs/v.dias.size),promFacDia:Math.round(v.facs.size/v.dias.size),peakHora:peakH?+peakH[0]:null};}).sort((a,b)=>b.regs-a.regs);
  return{cajeros,sede:cajeros[0]?.sede||"Sede",avg,avgR,avgRHMar,hasSched,tR:cajeros.reduce((s,c)=>s+c.tR,0),tF:cajeros.reduce((s,c)=>s+c.tF,0),periodo:{desde:allDates[0],hasta:allDates[allDates.length-1]},dS,allDates,kpiCfg:null,tpvStats};}

/* ═══ FILTER BY DATE RANGE ═══ */
function filterByDates(data,from,to){
  if(!from&&!to)return{...data,filtered:false};
  const cajeros=data.cajeros.map(c=>{const dias=c.dias.filter(d=>{if(from&&d.fecha<from)return false;if(to&&d.fecha>to)return false;return true;});if(dias.length===0)return null;
    const tF=dias.reduce((s,d)=>s+d.facs,0),tR=dias.reduce((s,d)=>s+d.regs,0),tH=dias.reduce((s,d)=>s+d.hrs,0);const consist=calcCV(dias.map(d=>d.fH));
    return{...c,dias,tF,tR,numDias:dias.length,hrsEfec:tH,hrsHor:null,diasProg:null,diasNoLab:null,pfH:tH>0?tF/tH:0,prH:tH>0?tR/tH:0,prF:tF>0?tR/tF:0,mD:dias.reduce((a,b)=>a.fH>b.fH?a:b),pDy:dias.reduce((a,b)=>a.fH<b.fH?a:b),consist};}).filter(Boolean);
  if(cajeros.length===0)return{...data,filtered:true,cajeros:[],avg:0,avgR:0,tR:0,tF:0,periodo:{desde:from||data.periodo.desde,hasta:to||data.periodo.hasta},dS:[],allDates:[]};
  cajeros.sort((a,b)=>b.pfH-a.pfH);const avg=cajeros.reduce((s,c)=>s+c.pfH,0)/cajeros.length;const avgR=cajeros.reduce((s,c)=>s+c.prH,0)/cajeros.length;
  rankCajeros(cajeros,avg,data.kpiCfg);
  const allDates=[...new Set(cajeros.flatMap(c=>c.dias.map(d=>d.fecha)))].sort();
  const dS=allDates.map(dk=>{let f=0,r=0,h=0,a=0;cajeros.forEach(c=>{const d=c.dias.find(x=>x.fecha===dk);if(d){f+=d.facs;r+=d.regs;h+=d.hrs;a++;}});return{fecha:dk,fechaD:fD(new Date(dk+"T12:00:00")),facs:f,regs:r,hrs:h,activos:a,fH:h>0?f/h:0};});
  return{...data,filtered:true,hasSched:false,cajeros,avg,avgR,tR:cajeros.reduce((s,c)=>s+c.tR,0),tF:cajeros.reduce((s,c)=>s+c.tF,0),periodo:{desde:allDates[0],hasta:allDates[allDates.length-1]},dS,allDates};}

/* ═══ EXCEL EXPORTS ═══ */
function aS(wb,nm,hd,rw,cw){const ws=XLSX.utils.aoa_to_sheet([hd,...rw]);if(cw)ws["!cols"]=cw.map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws,nm.substring(0,31));}

function exAll(data){toast("⏳ Generando Excel...");setTimeout(()=>{const wb=XLSX.utils.book_new();const hh=data.hasSched;const kt=kpiT(data);
  aS(wb,"Ranking",["#","CED","CAJERO","SEDE","FACT","REG","DIAS",...(hh?["H.HOR","H.EFEC"]:["H.EFEC"]),"F/H","R/H","R/F","CONSIST","KPI","ESTADO"],
    data.cajeros.map(c=>[c.rank,c.cedula,c.nombre,c.sede,c.tF,c.tR,c.numDias,...(hh?[c.hrsHor!==null?c.hrsHor:"-",+c.hrsEfec.toFixed(2)]:[+c.hrsEfec.toFixed(2)]),+c.pfH.toFixed(2),+c.prH.toFixed(2),+c.prF.toFixed(2),`${c.consist.label} (${c.consist.cv.toFixed(0)}%)`,`${c.kpi.ic} ${c.kpi.lab}`,`${c.frase.em} ${c.frase.tx}`]),
    [6,14,38,20,10,10,6,...(hh?[10,10]:[10]),10,10,10,18,16,36]);
  aS(wb,"Promedios",["METRICA","VALOR","DETALLE"],
    [["Cajeros",data.cajeros.length,""],["Facturas",data.tF,""],["Registros",data.tR,""],["Prom F/H",+data.avg.toFixed(4),""],["Prom R/H",+data.avgR.toFixed(4),""],["Cumple >=",+kt.cumple.toFixed(2),kt.custom?"Personalizado":"Prom+2"],["En Prom >=",+kt.enProm.toFixed(2),kt.custom?"Personalizado":"= Prom"],["Tipo KPIs",kt.custom?"PERSONALIZADOS":"AUTOMATICOS",""],["Desde",data.periodo.desde,""],["Hasta",data.periodo.hasta,""]],[26,16,30]);
  aS(wb,"Detalle Diario",["CAJERO","CED","FECHA","DIA","FACT","REG","INI","FIN","HRS","F/H","R/H","R/F"],
    data.cajeros.flatMap(c=>c.dias.map(d=>[c.nombre,c.cedula,d.fecha,d.diaSem,d.facs,d.regs,d.hi,d.hf,+d.hrs.toFixed(2),+d.fH.toFixed(2),+d.rH.toFixed(2),+d.rF.toFixed(2)])),[38,14,12,6,8,8,10,10,8,10,10,10]);
  aS(wb,"Por Hora",["CAJERO","HORA","REG","DIAS","PROM"],data.cajeros.flatMap(c=>c.hA.map(h=>[c.nombre,h.horaStr,h.regs,h.dias,+h.prom.toFixed(2)])),[38,8,10,8,12]);
  aS(wb,"Por Dia Semana",["CAJERO","DIA","REG","FACT","DIAS","HRS","F/H"],data.cajeros.flatMap(c=>c.dA.map(d=>[c.nombre,d.dia,d.regs,d.facs,d.dias,+d.horas.toFixed(2),+d.fH.toFixed(2)])),[38,8,10,10,8,10,10]);
  aS(wb,"Consistencia",["#","CAJERO","F/H","CV%","CALIFICACION","MEJOR F/H","PEOR F/H","DIF"],
    data.cajeros.map(c=>[c.rank,c.nombre,+c.pfH.toFixed(2),+c.consist.cv.toFixed(1),c.consist.label,c.mD?+c.mD.fH.toFixed(2):"",c.pDy?+c.pDy.fH.toFixed(2):"",c.mD&&c.pDy?+(c.mD.fH-c.pDy.fH).toFixed(2):""]),[6,38,10,8,16,10,10,10]);
  if(data.cajeros.some(c=>c.diasNoLab!==null))aS(wb,"Asistencia",["CAJERO","DIAS PROG","DIAS TRAB","NO LABORADOS","MARC FALLIDAS","% ASIST"],
    data.cajeros.filter(c=>c.diasProg).map(c=>[c.nombre,c.diasProg,c.numDias,c.diasNoLab,c.diasFallidos||0,((c.numDias/c.diasProg)*100).toFixed(1)+"%"]),[38,12,12,14,14,10]);
  aS(wb,"Sede Diario",["FECHA","FACT","REG","ACTIVOS","F/H"],data.dS.map(d=>[d.fecha,d.facs,d.regs,d.activos,+d.fH.toFixed(2)]),[12,10,10,10,10]);
  const fR=data.cajeros.map((c,i)=>{const r=i+2;return[c.nombre,c.tF,c.tR,+c.hrsEfec.toFixed(4),{f:`B${r}/D${r}`},{f:`C${r}/D${r}`}];});
  aS(wb,"Formulas",["CAJERO","FACT","REG","HRS","F/H","R/H"],fR,[38,10,10,12,12,12]);
  aS(wb,"Metodologia",["CONCEPTO","EXPLICACION"],[["REGISTROS","Cada fila = 1 producto"],["FACTURAS","Nro documento unicos"],["HRS EFECTIVAS","Ultimo-Primer registro/dia. Min 1h"],["HRS HORARIO","Del archivo de horarios (opcional)"],["CONSISTENCIA","CV% de Fact/Hora diaria. Menor = mas estable"],["F/H","Facturas/Horas. PRINCIPAL"],["CUMPLE",`>= ${fN(kt.cumple)}${kt.custom?" (personalizado)":" (Prom+2)"}`],["EN PROM",`>= ${fN(kt.enProm)} y < ${fN(kt.cumple)}`],["NO CUMPLE",`< ${fN(kt.enProm)}`]],[24,70]);
  XLSX.writeFile(wb,`Auditoria_${data.sede.replace(/\s/g,"_")}.xlsx`);toast("✅ Excel listo");},50);}

function exInd(c,data){toast("⏳ Generando Excel...");setTimeout(()=>{const wb=XLSX.utils.book_new();const kt=kpiT(data);
  const rows=[["Nombre",c.nombre,""],["Cedula",c.cedula,""],["Sede",c.sede,""],["Ranking",`#${c.rank} de ${data.cajeros.length}`,""],["Facturas",c.tF,""],["Registros",c.tR,""],["Dias",c.numDias,""]];
  if(c.hrsHor!==null)rows.push(["Hrs Horario",c.hrsHor,""]);rows.push(["Hrs Efectivas",+c.hrsEfec.toFixed(2),""]);
  if(c.hrsHor!==null&&c.hrsHor>0)rows.push(["Aprovechamiento",((c.hrsEfec/c.hrsHor)*100).toFixed(1)+"%",""]);
  if(c.diasProg)rows.push(["Dias Programados",c.diasProg,""],["Dias No Laborados",c.diasNoLab,""]);
  if(c.diasFallidos>0)rows.push(["Marc. Fallidas",c.diasFallidos+" dias","Dias con marcacion fallida"]);
  rows.push(["F/H",+c.pfH.toFixed(4),""],["Consistencia",`${c.consist.label} (CV ${c.consist.cv.toFixed(0)}%)`,""],["KPI",`${c.kpi.ic} ${c.kpi.lab}`,`Umbral: ${fN(kt.cumple)}`]);
  aS(wb,"Resumen",["DATO","VALOR","DETALLE"],rows,[20,22,30]);
  aS(wb,"Diario",["FECHA","DIA","FACT","REG","INI","FIN","HRS","F/H","R/H"],c.dias.map(d=>[d.fecha,d.diaSem,d.facs,d.regs,d.hi,d.hf,+d.hrs.toFixed(2),+d.fH.toFixed(2),+d.rH.toFixed(2)]),[12,6,8,8,10,10,8,10,10]);
  aS(wb,"Por Hora",["HORA","REG","DIAS","PROM"],c.hA.map(h=>[h.horaStr,h.regs,h.dias,+h.prom.toFixed(2)]),[8,10,8,12]);
  XLSX.writeFile(wb,`Informe_${c.nombre.split(" ").slice(0,2).join("_")}.xlsx`);toast("✅ Excel listo");},50);}

function exCmp(a,b){const wb=XLSX.utils.book_new();aS(wb,"Comparacion",["METRICA",a.nombre,b.nombre,"MEJOR"],[["Ranking",a.rank,b.rank,a.rank<b.rank?a.nombre:b.nombre],["F/H",+a.pfH.toFixed(2),+b.pfH.toFixed(2),a.pfH>b.pfH?a.nombre:b.nombre],["Consistencia",a.consist.label,b.consist.label,a.consist.cv<b.consist.cv?a.nombre:b.nombre],["KPI",`${a.kpi.ic} ${a.kpi.lab}`,`${b.kpi.ic} ${b.kpi.lab}`,""]],[14,24,24,28]);XLSX.writeFile(wb,`Comp_${sN(a.nombre)}_vs_${sN(b.nombre)}.xlsx`);toast("✅ Excel listo");}

function exAlt(data){toast("⏳ Generando Excel...");setTimeout(()=>{const wb=XLSX.utils.book_new();const kt=kpiT(data);
  aS(wb,"No Cumplen",["#","CAJERO","F/H","UMBRAL","DIF"],data.cajeros.filter(c=>c.kpi.lab==="No cumple").map(c=>[c.rank,c.nombre,+c.pfH.toFixed(2),+kt.enProm.toFixed(2),+(c.pfH-kt.enProm).toFixed(2)]),[6,38,10,10,10]);
  aS(wb,"En Promedio",["#","CAJERO","F/H","FALTA"],data.cajeros.filter(c=>c.kpi.lab==="En promedio").map(c=>[c.rank,c.nombre,+c.pfH.toFixed(2),+(kt.cumple-c.pfH).toFixed(2)]),[6,38,10,10]);
  aS(wb,"Config",["DATO","VALOR"],[["Tipo",kt.custom?"PERSONALIZADOS":"AUTOMATICOS"],["Cumple >=",+kt.cumple.toFixed(2)],["En Prom >=",+kt.enProm.toFixed(2)]],[20,16]);
  XLSX.writeFile(wb,`Alertas_${data.sede.replace(/\s/g,"_")}.xlsx`);toast("✅ Excel listo");},50);}

/* ═══ NARRATIVE PDF ═══ */
const PCSS=`*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',Arial,sans-serif;padding:30px 40px;color:#1a1a2e;font-size:13px;line-height:1.7;max-width:800px;margin:0 auto}h1{font-size:20px;font-weight:800;color:#0d4a28;margin-bottom:4px}h2{font-size:15px;font-weight:700;color:#1a6b3c;margin:22px 0 8px;padding-bottom:4px;border-bottom:2px solid #1a6b3c}h3{font-size:13px;font-weight:700;color:#333;margin:14px 0 6px}p{margin:6px 0;text-align:justify}.sub{font-size:12px;color:#5a6672;margin-bottom:16px}.badge{display:inline-block;padding:3px 12px;border-radius:16px;font-size:12px;font-weight:600}.badge-g{background:#e6f9ee;color:#1a6b3c}.badge-y{background:#fff9e6;color:#e67e22}.badge-r{background:#fde8e8;color:#e74c3c}table{width:100%;border-collapse:collapse;margin:10px 0;font-size:12px}th{background:#1a6b3c;color:#fff;padding:7px 10px;text-align:left;font-size:11px;-webkit-print-color-adjust:exact;print-color-adjust:exact}td{padding:6px 10px;border-bottom:1px solid #e0e0e0}tr:nth-child(even){background:#f8faf8}.mb{display:inline-block;background:#f0f4f1;border-radius:8px;padding:8px 14px;margin:4px;text-align:center;min-width:100px}.mb .v{font-size:18px;font-weight:800;color:#1a6b3c}.mb .l{font-size:10px;color:#666}.hl{background:#e6f4ec;padding:12px 16px;border-radius:10px;border-left:4px solid #1a6b3c;margin:10px 0}.wb{background:#fde8e8;padding:12px 16px;border-radius:10px;border-left:4px solid #e74c3c;margin:10px 0}.ft{margin-top:28px;padding-top:10px;border-top:1px solid #ccc;font-size:10px;color:#888;text-align:center}@media print{body{padding:15px 20px}@page{margin:1.5cm}.pb{page-break-before:always}}`;

function narInd(c,data){const kt=kpiT(data);const diff=Math.abs(c.pfH-data.avg).toFixed(2);const nm=esc(c.nombre),sd=esc(c.sede),cd=esc(c.cedula);
  const hrsL=c.hrsHor!==null&&c.hrsHor>0?`${fH(c.hrsEfec)} efectivas de ${c.hrsHor}h programadas (${((c.hrsEfec/c.hrsHor)*100).toFixed(1)}% aprovechamiento)`:`${fH(c.hrsEfec)} efectivas`;
  let ev,eb;if(c.kpi.lab==="Cumple"){ev=`${nm} ha superado las expectativas. Su productividad de ${fN(c.pfH)} f/h esta ${diff} puntos por encima del promedio (${fN(data.avg)}), posicion ${c.rank} de ${data.cajeros.length}. Demuestra eficiencia y compromiso.`;eb="badge-g";}else if(c.kpi.lab==="En promedio"){ev=`${nm} cumple con expectativas basicas. Su productividad de ${fN(c.pfH)} f/h es congruente con el promedio (${fN(data.avg)}), posicion ${c.rank} de ${data.cajeros.length}. Se recomiendan oportunidades de mejora.`;eb="badge-y";}else{ev=`${nm} presenta rendimiento por debajo. Su productividad de ${fN(c.pfH)} f/h esta ${diff} puntos bajo el promedio (${fN(data.avg)}), posicion ${c.rank} de ${data.cajeros.length}. Requiere plan de accion inmediato.`;eb="badge-r";}
  const rows=c.dias.map((d,i)=>`<tr${i%2===0?" style='background:#f8faf8'":""}><td>${d.fechaF}</td><td>${d.diaSemL}</td><td>${d.facs}</td><td>${d.regs}</td><td>${d.hi}</td><td>${d.hf}</td><td>${fH(d.hrs)}${d.minH?" ⚠️":""}</td><td style="font-weight:700;color:${d.fH>=data.avg?"#1a6b3c":"#e74c3c"}">${fN(d.fH)}</td><td>${fN(d.rH)}</td></tr>`).join("");
  const mbs=[["Facturas",c.tF.toLocaleString()],["Registros",c.tR.toLocaleString()],["Dias",c.numDias],["Hrs Efec",fH(c.hrsEfec)],...(c.hrsHor!==null?[["Hrs Horario",c.hrsHor+"h"]]:[]),...(c.diasNoLab!==null?[["No Laborados",c.diasNoLab+" dias"]]:[]),...(c.diasFallidos>0?[["Marc. Fallidas",c.diasFallidos+" dias"]]:[]),["F/H",fN(c.pfH)],["R/H",fN(c.prH)],["Consistencia",`${c.consist.label}`]].map(([l,v])=>`<div class="mb"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");
  return`<h1>Informe de Productividad - ${nm}</h1><p class="sub">Periodo: ${data.periodo.desde} a ${data.periodo.hasta} | ${sd} | Cedula: ${cd}</p>
<h2>Resumen Ejecutivo</h2><p>${ev}</p>
<h2>Metricas</h2><div style="text-align:center">${mbs}</div>
<h2>Analisis de Resultados</h2><p>Proceso <b>${c.tF.toLocaleString()} facturas</b> y <b>${c.tR.toLocaleString()} productos</b> en <b>${hrsL}</b> a lo largo de <b>${c.numDias} dias</b>. Productividad: <b>${fN(c.pfH)} f/h</b>, posicion <b>${c.rank} de ${data.cajeros.length}</b>.</p>
<p>Su nivel de consistencia es <b>${c.consist.label}</b> (CV: ${c.consist.cv.toFixed(1)}%), lo que indica ${c.consist.cv<=30?"un rendimiento predecible y confiable dia a dia":"variabilidad significativa entre dias, lo que sugiere inconsistencia operativa"}.</p>
${c.mD?`<div class="hl"><b>Mejor dia:</b> ${c.mD.fechaF} (${c.mD.diaSemL}) con ${fN(c.mD.fH)} f/h y ${c.mD.facs} facturas.</div>`:""}
${c.pDy?`<div class="${c.pDy.fH<data.avg?"wb":"hl"}"><b>Dia mas bajo:</b> ${c.pDy.fechaF} (${c.pDy.diaSemL}) con ${fN(c.pDy.fH)} f/h.</div>`:""}
${c.diasNoLab!==null?`<p>De ${c.diasProg} dias programados, trabajo <b>${c.numDias} dias</b> con <b>${c.diasNoLab} dia(s) no laborados</b> (${((c.numDias/c.diasProg)*100).toFixed(1)}% asistencia).${c.diasFallidos>0?` <span style="color:#e67e22"><b>${c.diasFallidos} dia(s) con marcacion fallida</b> (la hora se uso pero la marcacion no fue valida).</span>`:""}</p>`:""}
<h2>Detalle Diario</h2><table><thead><tr><th>Fecha</th><th>Dia</th><th>Fact</th><th>Reg</th><th>Ini</th><th>Fin</th><th>Hrs</th><th>F/H</th><th>R/H</th></tr></thead><tbody>${rows}</tbody></table>
<h2>Conclusion</h2><p><span class="badge ${eb}">${c.kpi.ic} ${c.kpi.lab}</span></p>
<p>${c.kpi.lab==="Cumple"?`Desempeno sobresaliente. Supera el promedio en ${diff} f/h.`:c.kpi.lab==="En promedio"?`Desempeno aceptable. Se recomienda optimizacion para el proximo ciclo.`:`Requiere atencion prioritaria. Plan de accion con seguimiento semanal.`}</p>
${kt.custom?`<div class="hl"><b>KPIs utilizados:</b> Cumple >= ${fN(kt.cumple)} | En prom >= ${fN(kt.enProm)} (personalizados)</div>`:""}
<div class="ft">Sistema de Auditoria | ${new Date().toLocaleDateString("es-CO")} | Prom sede: ${fN(data.avg)} f/h</div>`;}

function narGen(data){const cu=data.cajeros.filter(c=>c.kpi.lab==="Cumple"),ep=data.cajeros.filter(c=>c.kpi.lab==="En promedio"),nc=data.cajeros.filter(c=>c.kpi.lab==="No cumple");const kt=kpiT(data);const sd=esc(data.sede);
  const t3=data.cajeros.slice(0,3),pCu=((cu.length/data.cajeros.length)*100).toFixed(0),pNc=((nc.length/data.cajeros.length)*100).toFixed(0);
  const bD=data.dS.reduce((a,b)=>a.fH>b.fH?a:b),wD=data.dS.reduce((a,b)=>a.fH<b.fH?a:b);
  const mostConsist=data.cajeros.slice().sort((a,b)=>a.consist.cv-b.consist.cv).slice(0,3);
  const leastConsist=data.cajeros.slice().sort((a,b)=>b.consist.cv-a.consist.cv).slice(0,3);
  const rows=data.cajeros.map((c,i)=>{const bg=c.kpi.lab==="Cumple"?"#e6f9ee":c.kpi.lab==="No cumple"?"#fde8e8":"#fff9e6";return`<tr style="background:${bg}"><td style="font-weight:800;color:${rColor(i+1,data.cajeros.length)}">${c.rank}</td><td>${esc(c.nombre)}</td><td>${c.tF.toLocaleString()}</td><td>${fH(c.hrsEfec)}</td><td style="font-weight:700;color:${c.pfH>=data.avg?"#1a6b3c":"#e74c3c"}">${fN(c.pfH)}</td><td>${c.consist.label}</td><td>${c.kpi.ic} ${c.kpi.lab}</td></tr>`;}).join("");
  return`<h1>Informe General de Productividad</h1><p class="sub">${sd} | ${data.periodo.desde} a ${data.periodo.hasta}</p>
<h2>Resumen Ejecutivo</h2><p>La sede <b>${sd}</b> cuenta con <b>${data.cajeros.length} cajeros</b>. Se procesaron <b>${data.tF.toLocaleString()} facturas</b> y <b>${data.tR.toLocaleString()} registros</b>.</p>
<p>${kt.custom?`<b>KPIs personalizados:</b> Cumple >= ${fN(kt.cumple)} f/h, En promedio >= ${fN(kt.enProm)} f/h.`:`Umbrales automaticos: Cumple >= ${fN(kt.cumple)} f/h, En promedio >= ${fN(kt.enProm)} f/h.`} Promedio sede: <b>${fN(data.avg)} f/h</b>. De ${data.cajeros.length} colaboradores: <b>${cu.length} (${pCu}%) cumplen</b>, <b>${ep.length} en promedio</b>, <b>${nc.length} (${pNc}%) no cumplen</b>.</p>
${nc.length>0?`<div class="wb"><b>Atencion:</b> ${nc.length} colaboradores criticos: ${nc.map(c=>`${esc(c.nombre)} (${fN(c.pfH)} f/h)`).join(", ")}.</div>`:`<div class="hl"><b>Excelente:</b> Todos los colaboradores dentro o por encima del promedio.</div>`}
<h2>Top 3</h2><p>${t3.map((c,i)=>`<b>${i+1}. ${esc(c.nombre)}</b> con ${fN(c.pfH)} f/h (${c.consist.label})`).join(". ")}.</p>
<h2>Consistencia del Equipo</h2><p>Los mas estables: ${mostConsist.map(c=>`<b>${esc(sN(c.nombre))}</b> (CV ${c.consist.cv.toFixed(0)}%)`).join(", ")}. Los mas variables: ${leastConsist.map(c=>`<b>${esc(sN(c.nombre))}</b> (CV ${c.consist.cv.toFixed(0)}%)`).join(", ")}.</p>
${nc.length>0?`<h2>Requieren Atencion</h2><p>${nc.map(c=>`<b>${esc(c.nombre)}</b> (#${c.rank}) con ${fN(c.pfH)} f/h, ${fN(Math.abs(c.pfH-data.avg))} bajo promedio`).join(". ")}.</p>`:""}
<h2>Ranking</h2><table><thead><tr><th>#</th><th>Cajero</th><th>Fact</th><th>Hrs</th><th>F/H</th><th>Consist</th><th>KPI</th></tr></thead><tbody>${rows}</tbody></table>
<h2>Conclusiones</h2><p>${cu.length>0?`Reconocer a los ${cu.length} que cumplen, especialmente ${esc(t3[0].nombre)}. `:""}${nc.length>0?`Para los ${nc.length} que no cumplen: (1) retroalimentacion individual, (2) plan de mejora semanal, (3) seguimiento quincenal.`:"Buen nivel general. Buscar mejora continua."}</p>
<div class="ft">${new Date().toLocaleDateString("es-CO")}</div>`;}

function narAllInd(data){return data.cajeros.map((c,i)=>`${i>0?'<div class="pb"></div>':""}${narInd(c,data)}`).join("");}

function dlReport(html,title){const full=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title><style>${PCSS}\n.print-btn{position:fixed;top:10px;right:10px;padding:12px 28px;background:#1a6b3c;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;box-shadow:0 4px 15px rgba(0,0,0,.3);z-index:9999}.print-btn:hover{background:#0d4a28}@media print{.print-btn{display:none!important}}</style></head><body><button class="print-btn" onclick="window.print()">Guardar como PDF</button>${html}<script>window.onload=function(){setTimeout(function(){window.print()},600)}<\/script></body></html>`;
  const blob=new Blob([full],{type:"text/html;charset=utf-8"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=`${title.replace(/[^a-zA-Z0-9_\- ]/g,"").replace(/\s+/g,"_")}.html`;document.body.appendChild(a);a.click();setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},300);toast("✅ Informe descargado");}

/* ═══ UI ═══ */
const Btn=({onClick,children,green,small})=><button onClick={onClick} style={{padding:small?"5px 10px":"8px 16px",borderRadius:10,border:"none",fontSize:small?11:13,fontWeight:600,cursor:"pointer",background:green?"#27ae60":CL.pri,color:"#fff"}}>{children}</button>;
function toast(msg,ms=2000){const d=document.createElement("div");d.textContent=msg;d.style.cssText="position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#1a6b3c;color:#fff;padding:10px 24px;border-radius:10px;font-size:13px;font-weight:700;z-index:99999;box-shadow:0 4px 15px rgba(0,0,0,.3);transition:opacity .3s";document.body.appendChild(d);setTimeout(()=>{d.style.opacity="0";setTimeout(()=>d.remove(),300);},ms);}
const KC=({label,value,color,big})=><div style={{background:`${color}12`,borderRadius:12,padding:"12px",border:`1px solid ${color}30`,flex:big?"1.4 1 170px":"1 1 130px",minWidth:115}}><div style={{fontSize:big?20:17,fontWeight:800,color}}>{value}</div><div style={{fontSize:11,color:CL.txtL}}>{label}</div></div>;
const Bg=({color,children})=><span style={{display:"inline-flex",alignItems:"center",gap:3,padding:"3px 9px",borderRadius:18,fontSize:11,fontWeight:600,background:`${color}18`,color,border:`1px solid ${color}30`}}>{children}</span>;
const rBadge=c=>c.activo===false?" ⛔":"";
const crd={background:CL.card,borderRadius:14,padding:"18px",boxShadow:"0 1px 6px rgba(0,0,0,.06)",border:"1px solid #e8ece9"};
const TH={padding:"7px 9px",textAlign:"left",fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:".4px",color:CL.txtL,borderBottom:"2px solid #e8ece9",whiteSpace:"nowrap"};
const TD={padding:"6px 9px",fontSize:12,borderBottom:"1px solid #f0f2f1"};
const SL={padding:"10px 12px",borderRadius:10,border:`2px solid ${CL.pri}40`,fontSize:14,background:"#fff",color:CL.txt,outline:"none",cursor:"pointer",minWidth:220};

const KPIBanner=({data})=>{const k=kpiT(data);return <div style={{...crd,padding:"8px 14px",display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",background:k.custom?"#fff9e6":"#f8faf8",border:`1px solid ${k.custom?"#f39c1240":"#e8ece9"}`}}><span style={{fontSize:16}}>{k.custom?"⚙️":"📊"}</span><div style={{flex:1,fontSize:12}}><span style={{fontWeight:700,color:k.custom?CL.warn:CL.pri}}>{k.custom?"KPIs Personalizados":"KPIs Auto"}</span><span style={{color:CL.txtL,marginLeft:6}}>Cumple ≥ <b style={{color:"#2ecc71"}}>{fN(k.cumple)}</b> | En prom ≥ <b style={{color:"#f39c12"}}>{fN(k.enProm)}</b></span></div></div>;};

/* ═══ UPLOAD ═══ */
function parsePersonnel(jr){
  const hd=jr[0].map(h=>String(h||"").trim().toUpperCase());
  const fc=ks=>{for(const k of ks){const idx=hd.findIndex(h=>h.includes(k));if(idx>=0)return idx;}return-1;};
  const cNom=fc(["NOMBRE DEL EMPLEADO","NOMBRE EMPLEADO","NOMBRE"]);
  const cSede=fc(["DESCRIPCION C.O","DESC C.O","SEDE"]);
  const cRetiro=fc(["FECHA RETIRO","RETIRO"]);
  const cCargo=fc(["DESCRIPCION DEL CARGO","DESC CARGO","CARGO"]);
  const cCcosto=fc(["DESCRIPCION CCOSTO","DESC CCOSTO","CCOSTO","CENTRO COSTO","SECCION"]);
  if(cNom<0)return null;
  const map={};
  for(let i=1;i<jr.length;i++){const r=jr[i];if(!r||!r[cNom])continue;
    const nm=String(r[cNom]).trim().toUpperCase();
    const sede=cSede>=0?String(r[cSede]||"").trim():"";
    const cargo=cCargo>=0?String(r[cCargo]||"").trim():"";
    const ccosto=cCcosto>=0?String(r[cCcosto]||"").trim():"";
    let retiro=null;if(cRetiro>=0&&r[cRetiro]){const v=r[cRetiro];if(v instanceof Date)retiro=v;else if(typeof v==="string"&&v.trim())retiro=parseDate(v);else if(typeof v==="number")retiro=parseDate(v);}
    map[nm]={sede,cargo,ccosto,retiro,activo:!retiro};
  }
  return map;
}

function Upload({onData}){const r1=useRef(),r2=useRef(),r3=useRef(),r4=useRef();const[mD,setMD]=useState(null);const[sD,setSD]=useState(null);const[bdP,setBdP]=useState(null);const[st,setSt]=useState("");const[er,setEr]=useState(null);
  const rf=useCallback(async f=>{const buf=await f.arrayBuffer();const nm=f.name.toLowerCase();
    if(nm.endsWith(".csv")){
      let txt;try{txt=new TextDecoder("utf-8",{fatal:true}).decode(buf);}catch(e){txt=new TextDecoder("latin1").decode(buf);}
      const rows=[];let cur="",inQ=false;const pushRow=()=>{rows.push(cur.split(",").map(c=>{c=c.trim();if(c.startsWith('"')&&c.endsWith('"'))c=c.slice(1,-1).replace(/""/g,'"');return c;}));cur="";};
      for(const ch of txt){if(ch==='"')inQ=!inQ;else if(ch==="\n"&&!inQ){pushRow();}else{cur+=ch;}}if(cur.trim())pushRow();
      return rows;
    }
    const wb=XLSX.read(buf,{type:"array",cellDates:false,raw:false});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:""});
  },[]);
  const lm=useCallback(async f=>{setEr(null);setSt("Leyendo...");try{const j=await rf(f);setSt(`${j.length-1} registros.`);setMD(j);}catch(e){setEr(e.message);}},[rf]);
  const ls=useCallback(async f=>{setEr(null);setSt("Leyendo marcaciones...");try{const j=await rf(f);
    if(!j||j.length<2){setEr("Archivo vacio: "+(j?j.length:0)+" filas");return;}
    toast("Leido: "+j.length+" filas | "+String(j[0]).substring(0,60),4000);
    const m=processSchedule(j);if(!m){setEr("No encontro columnas. Headers: "+j[0].slice(0,8).join(", "));return;}
    const cnt=Object.keys(m.data).length;const s1=Object.values(m.data)[0];
    const byDayN=s1&&s1.byDay?Object.keys(s1.byDay).length:0;const byDaySample=s1&&s1.byDay?Object.keys(s1.byDay)[0]:"none";
    toast(cnt+" emp | byCed:"+m.byCed+" | ej:"+((s1&&s1.nombre)?s1.nombre.substring(0,20):"?")+" "+((s1)?s1.hrs:0)+"h byDay:"+byDayN+" sample:"+byDaySample,6000);
    setSD(m);setSt("Marcaciones: "+cnt+" empleados");}catch(e){setEr("Error: "+e.message);toast("ERROR: "+e.message,5000);}},[rf]);
  const lbd=useCallback(async f=>{setEr(null);setSt("Leyendo base de datos...");try{const j=await rf(f);const m=parsePersonnel(j);if(!m){setEr("No se encontro columna 'Nombre del empleado'.");return;}
    const total=Object.keys(m).length;const activos=Object.values(m).filter(v=>v.activo).length;
    setBdP(m);setSt(`BD Personal: ${total} empleados (${activos} activos, ${total-activos} retirados).`);}catch(e){setEr(e.message);}},[rf]);
  const lj=useCallback(async f=>{setEr(null);setSt("Cargando memoria...");try{const txt=await f.text();const json=JSON.parse(txt);const d=loadSnapshot(json);
    if(!d){setEr("Archivo no valido.");return;}
    const c1=d.cajeros[0];
    toast(d.cajeros.length+"caj | tpv:"+((d.tpvStats||[]).length)+" | hA:"+((c1&&c1.hA)?c1.hA.length:0)+" | dias:"+((c1)?c1.dias.length:0)+" | dS:"+d.dS.length,6000);
    setSt("Memoria: "+d.sede+" | "+d.cajeros.length+" cajeros");
    setTimeout(()=>onData(d),600);}catch(e){setEr("Error: "+e.message);}},[onData]);
  const go=useCallback(()=>{if(!mD)return;setSt("Procesando...");setEr(null);setTimeout(()=>{try{const d=processData(mD,sD,bdP);
    const matched=d.cajeros.filter(c=>c.hrsHor!==null).length;
    const c1=d.cajeros[0];const c1info=c1?c1.nombre.substring(0,20)+"|hrsHor:"+c1.hrsHor+"|hrsMar:"+c1.totalHrsMar:"none";
    const sdSample=sD?Object.values(sD.data)[0]:null;const sdName=sdSample&&sdSample.nombre?sdSample.nombre.substring(0,25):"NO_NAME";
    const msg=d.hasSched?matched+"/"+d.cajeros.length+" match | TRUE | "+c1info:d.cajeros.length+"caj | FALSE | sD:"+(sD?"byCed:"+sD.byCed:"NULL")+" | sdName:"+sdName+" | caj1:"+c1info;
    toast(msg,8000);setTimeout(()=>onData(d),900);}catch(e){setEr(e.message);toast("ERROR: "+e.message,5000);}},200);},[mD,sD,bdP,onData]);
  return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:`linear-gradient(135deg,${CL.priDk},${CL.pri},#1a8c4e)`,padding:20}}>
    <div style={{textAlign:"center",maxWidth:560,width:"100%"}}><div style={{fontSize:42,marginBottom:8}}>📊</div><h1 style={{color:"#fff",fontSize:26,fontWeight:800,marginBottom:4}}>Auditoria de Cajeros</h1><p style={{color:"rgba(255,255,255,.7)",fontSize:14,marginBottom:24}}>Sistema de productividad</p>
      <div onClick={()=>r1.current?.click()} style={{background:mD?"rgba(46,204,113,.2)":"rgba(255,255,255,.1)",border:`2px dashed ${mD?"rgba(46,204,113,.6)":"rgba(255,255,255,.35)"}`,borderRadius:16,padding:"28px 20px",cursor:"pointer",marginBottom:12}}>
        <input ref={r1} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{display:"none"}} onChange={e=>e.target.files[0]&&lm(e.target.files[0])}/><div style={{fontSize:28,marginBottom:4}}>{mD?"✅":"📁"}</div><p style={{color:"#fff",fontSize:14,fontWeight:600}}>{mD?`Registros cargados (${mD.length-1} filas)`:"1. Sube el archivo de REGISTROS"}</p></div>
      <div onClick={()=>r2.current?.click()} style={{background:sD?"rgba(46,204,113,.2)":"rgba(255,255,255,.06)",border:`2px dashed ${sD?"rgba(46,204,113,.6)":"rgba(255,255,255,.2)"}`,borderRadius:16,padding:"18px",cursor:"pointer",marginBottom:16}}>
        <input ref={r2} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{display:"none"}} onChange={e=>e.target.files[0]&&ls(e.target.files[0])}/><div style={{fontSize:22,marginBottom:2}}>{sD?"✅":"📊"}</div><p style={{color:"#fff",fontSize:13,fontWeight:600}}>{sD?`Marcaciones cargadas`:"2. Sube MARCACIONES (entrada/salida)"}</p><p style={{color:"rgba(255,255,255,.4)",fontSize:11}}>Para el ranking R/H y horas laboradas</p></div>
      <div onClick={()=>r4.current?.click()} style={{background:bdP?"rgba(46,204,113,.2)":"rgba(255,255,255,.06)",border:`2px dashed ${bdP?"rgba(46,204,113,.6)":"rgba(255,255,255,.2)"}`,borderRadius:16,padding:"18px",cursor:"pointer",marginBottom:16}}>
        <input ref={r4} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{display:"none"}} onChange={e=>e.target.files[0]&&lbd(e.target.files[0])}/><div style={{fontSize:22,marginBottom:2}}>{bdP?"✅":"👥"}</div><p style={{color:"#fff",fontSize:13,fontWeight:600}}>{bdP?`BD Personal cargada (${Object.keys(bdP).length})`:"3. (Opcional) Sube MAESTRO DE PERSONAL"}</p><p style={{color:"rgba(255,255,255,.4)",fontSize:11}}>Para identificar activos vs retirados</p></div>
      {mD&&sD&&<button onClick={go} style={{padding:"14px 40px",borderRadius:14,border:"none",background:"#fff",color:CL.pri,fontSize:16,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 20px rgba(0,0,0,.2)"}}>🚀 Procesar</button>}
      {mD&&!sD&&<p style={{color:"rgba(255,255,255,.6)",fontSize:12,textAlign:"center"}}>⬆️ Sube las marcaciones para continuar</p>}
      <div style={{marginTop:24,paddingTop:20,borderTop:"1px solid rgba(255,255,255,.15)"}}>
        <p style={{color:"rgba(255,255,255,.5)",fontSize:12,marginBottom:10}}>¿Ya tienes un analisis guardado?</p>
        <button onClick={()=>r3.current?.click()} style={{padding:"12px 28px",borderRadius:12,border:"2px solid rgba(255,255,255,.3)",background:"rgba(255,255,255,.08)",color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>💾 Cargar Memoria (.json)</button>
        <input ref={r3} type="file" accept=".json" style={{display:"none"}} onChange={e=>e.target.files[0]&&lj(e.target.files[0])}/>
        <p style={{color:"rgba(255,255,255,.35)",fontSize:11,marginTop:6}}>Carga un .json guardado previamente para ver el analisis sin necesidad de los archivos originales</p></div>
      {st&&<p style={{color:"rgba(255,255,255,.8)",fontSize:13,marginTop:10,fontWeight:600}}>{st}</p>}
      {er&&<div style={{marginTop:10,padding:"8px 14px",background:"rgba(231,76,60,.2)",borderRadius:12,color:"#fff",fontSize:13}}>⚠️ {er}</div>}</div></div>);}

/* ═══ FILTER BAR ═══ */
function FilterBar({data,from,to,setFrom,setTo,sedes,sedeFilter,setSedeFilter,hasBD,showRetirados,setShowRetirados,ccostos,ccostoFilter,setCcostoFilter,excSuperv,setExcSuperv}){const mn=data.allDates?.[0]||"",mx=data.allDates?.[data.allDates.length-1]||"";const active=from||to;const invalid=from&&to&&from>to;
  const[ccOpen,setCcOpen]=useState(false);
  const toggleCc=v=>{setCcostoFilter(prev=>prev.includes(v)?prev.filter(x=>x!==v):[...prev,v]);};
  return <div style={{...crd,padding:"10px 16px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",background:invalid?"#fde8e8":active||sedeFilter||!showRetirados||ccostoFilter.length?"#fff9e6":"#f8faf8"}}>
    {sedes.length>1&&<><span style={{fontSize:14}}>🏢</span><select value={sedeFilter} onChange={e=>setSedeFilter(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:"2px solid #9b59b6",fontSize:12,fontWeight:600,maxWidth:200}}>
      <option value="">Todas las sedes ({sedes.length})</option>{sedes.map(s=><option key={s} value={s}>{s}</option>)}</select></>}
    {ccostos.length>1&&<div style={{position:"relative"}}><button onClick={()=>setCcOpen(!ccOpen)} style={{padding:"6px 10px",borderRadius:8,border:"2px solid #e67e22",fontSize:12,fontWeight:600,background:ccostoFilter.length?"#fff3e0":"#fff",color:"#333",cursor:"pointer"}}>
      🏷️ {ccostoFilter.length?ccostoFilter.length+" secciones":"Secciones ("+ccostos.length+")"} ▾</button>
      {ccOpen&&<div style={{position:"absolute",top:"100%",left:0,zIndex:999,background:"#fff",border:"1px solid #e67e22",borderRadius:8,padding:6,maxHeight:200,overflowY:"auto",minWidth:220,boxShadow:"0 4px 12px rgba(0,0,0,.15)"}}>
        {ccostos.map(s=><label key={s} style={{display:"flex",alignItems:"center",gap:4,padding:"3px 4px",fontSize:11,cursor:"pointer",borderRadius:4,background:ccostoFilter.includes(s)?"#fff3e0":"transparent"}}>
          <input type="checkbox" checked={ccostoFilter.includes(s)} onChange={()=>toggleCc(s)}/>{s}</label>)}
        {ccostoFilter.length>0&&<button onClick={()=>setCcostoFilter([])} style={{width:"100%",padding:4,fontSize:10,border:"none",background:"#fde8e8",borderRadius:4,cursor:"pointer",marginTop:4}}>✕ Limpiar secciones</button>}
      </div>}</div>}
    {hasBD&&<label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,fontWeight:600,color:showRetirados?"#e74c3c":"#2ecc71",cursor:"pointer"}}>
      <input type="checkbox" checked={!showRetirados} onChange={e=>setShowRetirados(!e.target.checked)}/>{showRetirados?"👥 Con retirados":"✅ Solo activos"}</label>}
    {hasBD&&<label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,fontWeight:600,color:excSuperv?"#2ecc71":"#9b59b6",cursor:"pointer"}}>
      <input type="checkbox" checked={excSuperv} onChange={e=>setExcSuperv(e.target.checked)}/>{excSuperv?"🚫 Sin supervisores":"👔 Con supervisores"}</label>}
    <span style={{fontSize:14}}>📅</span><span style={{fontSize:12,fontWeight:700,color:CL.txtL}}>Fechas:</span>
    <input type="date" value={from} min={mn} max={to||mx} onChange={e=>setFrom(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${invalid?"#e74c3c":"#ddd"}`,fontSize:12}}/>
    <span style={{fontSize:12,color:CL.txtL}}>a</span>
    <input type="date" value={to} min={from||mn} max={mx} onChange={e=>setTo(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${invalid?"#e74c3c":"#ddd"}`,fontSize:12}}/>
    {(active||sedeFilter||ccostoFilter.length)&&<button onClick={()=>{setFrom("");setTo("");setSedeFilter("");setCcostoFilter([]);}} style={{padding:"4px 10px",borderRadius:8,border:"1px solid #e74c3c",background:"#fff",color:"#e74c3c",fontSize:11,cursor:"pointer",fontWeight:600}}>✕ Limpiar</button>}
    {invalid&&<span style={{fontSize:11,color:CL.dan,fontWeight:600}}>❌ Desde no puede ser mayor que Hasta</span>}
    {(active||sedeFilter||ccostoFilter.length)&&!invalid&&<span style={{fontSize:11,color:CL.warn,fontWeight:600}}>⚠️ Filtro{sedeFilter?` | ${sedeFilter}`:""}{ccostoFilter.length?` | ${ccostoFilter.join(", ")}`:""}</span>}
  </div>;}

/* ═══ DASHBOARD ═══ */
function Dash({data,showR}){
  if(!data.cajeros.length)return <div style={{...crd,textAlign:"center",padding:40,color:CL.txtL}}><div style={{fontSize:40}}>📭</div><p style={{fontSize:14,fontWeight:600}}>Sin datos en este rango de fechas</p><p style={{fontSize:12}}>Ajusta el filtro de fechas para ver resultados.</p></div>;
  const t5=data.cajeros.slice(0,5),b3=data.cajeros.slice(-3).reverse();const ch=data.cajeros.map(c=>({n:sN(c.nombre),fH:+c.pfH.toFixed(2)}));const sc=data.dS.map(d=>({f:d.fechaD,fH:+d.fH.toFixed(2)}));
  const sedeTH=data.cajeros.reduce((s,c)=>s+c.hrsEfec,0),sedeTD=data.cajeros.reduce((s,c)=>s+c.numDias,0);
  const pp=[{name:"Cumplen",value:data.cajeros.filter(c=>c.kpi.lab==="Cumple").length,fill:"#2ecc71"},{name:"En Prom",value:data.cajeros.filter(c=>c.kpi.lab==="En promedio").length,fill:"#f39c12"},{name:"No Cumplen",value:data.cajeros.filter(c=>c.kpi.lab==="No cumple").length,fill:"#e74c3c"}];
  // Weekly trend
  const weeks=useMemo(()=>{const w={};const meses=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];data.dS.forEach(d=>{const dt=new Date(d.fecha+"T12:00:00");const wk=`${meses[dt.getMonth()]}-S${Math.ceil(dt.getDate()/7)}`;if(!w[wk])w[wk]={fH:[],f:0,r:0};w[wk].fH.push(d.fH);w[wk].f+=d.facs;w[wk].r+=d.regs;});return Object.entries(w).map(([k,v])=>({sem:k,fH:+(v.fH.reduce((a,b)=>a+b,0)/v.fH.length).toFixed(2),fact:v.f}));},[data]);
  // Dead hours (sede-wide hourly)
  const hSede=useMemo(()=>{const h={};for(let i=0;i<24;i++)h[i]={regs:0,dias:0};data.cajeros.forEach(c=>c.hA.forEach(ha=>{h[ha.hora].regs+=ha.regs;h[ha.hora].dias=Math.max(h[ha.hora].dias,ha.dias);}));
    return Object.entries(h).filter(([,v])=>v.regs>0).map(([k,v])=>({hora:`${k}:00`,regs:v.regs,h:+k})).sort((a,b)=>a.h-b.h);},[data]);
  const peakH=hSede.length>0?hSede.reduce((a,b)=>a.regs>b.regs?a:b):null;
  const deadH=hSede.length>2?hSede.filter(h=>h.regs>0).sort((a,b)=>a.regs-b.regs).slice(0,3):[];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <KPIBanner data={data}/>
    <div style={{...crd,background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,color:"#fff",padding:"12px 18px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
      <div><div style={{fontSize:12,opacity:.8}}>Sede</div><div style={{fontSize:18,fontWeight:700}}>{data.sede}</div></div>
      <div style={{display:"flex",gap:6}}><Btn onClick={()=>showR({html:narGen(data),title:`General - ${data.sede}`})}>📄 Informe</Btn><Btn green onClick={()=>exAll(data)}>📥 Excel</Btn></div>
      <div style={{textAlign:"right"}}><div style={{fontSize:12,opacity:.8}}>Periodo</div><div style={{fontSize:14,fontWeight:600}}>{fDF(new Date(data.periodo.desde+"T12:00:00"))} a {fDF(new Date(data.periodo.hasta+"T12:00:00"))}</div></div></div>
    <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><KC label="Cajeros" value={data.cajeros.length} color={CL.pri}/><KC label="Facturas" value={data.tF.toLocaleString()} color="#3498db"/><KC label="Registros" value={data.tR.toLocaleString()} color="#9b59b6"/><KC label="Prom F/H" value={fN(data.avg)} color="#e67e22"/><KC label="Prom R/H" value={fN(data.avgR)} color="#2ecc71"/>
      <KC label="Prom Hrs/Dia" value={sedeTD>0?(sedeTH/sedeTD).toFixed(1):"—"} color="#8e44ad"/><KC label="Total Hrs" value={Math.round(sedeTH)+"h"} color="#2c3e50"/></div>
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>🏆 Top 5</h3>{t5.map((c,i)=><div key={c.nombre} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:i<4?"1px solid #f0f2f1":"none"}}><div style={{width:26,height:26,borderRadius:"50%",background:rColor(i+1,data.cajeros.length),color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:11,flexShrink:0}}>{i+1}</div><div style={{flex:1,fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nombre}{rBadge(c)}</div><div style={{fontSize:13,fontWeight:800,color:CL.pri}}>{fN(c.pfH)}</div><Bg color={c.consist.cl}>{c.consist.label}</Bg></div>)}</div>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📊 KPI</h3><ResponsiveContainer width="100%" height={180}><PieChart><Pie data={pp} cx="50%" cy="50%" outerRadius={55} dataKey="value" label={e=>`${e.name}:${e.value}`}>{pp.map((e,i)=><Cell key={i} fill={e.fill}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer>
        <h4 style={{fontSize:12,fontWeight:700,marginTop:4}}>⚠️ Atencion</h4>{b3.map(c=><div key={c.nombre} style={{display:"flex",gap:6,padding:"2px 0",fontSize:11}}><span style={{fontWeight:800,color:CL.dan}}>#{c.rank}</span><span style={{flex:1,fontWeight:600}}>{sN(c.nombre)}{rBadge(c)}</span><span style={{fontWeight:800,color:CL.dan}}>{fN(c.pfH)}</span></div>)}</div></div>
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📈 Tendencia Semanal</h3>
        <ResponsiveContainer width="100%" height={160}><ComposedChart data={weeks}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="sem" tick={{fontSize:11}}/><YAxis tick={{fontSize:10}}/><Tooltip/><Bar dataKey="fH" fill={`${CL.pri}40`} radius={[4,4,0,0]}/><Line type="monotone" dataKey="fH" stroke={CL.pri} strokeWidth={2.5} dot={{r:4}} name="Prom F/H"/></ComposedChart></ResponsiveContainer>
        <p style={{fontSize:11,color:CL.txtL,marginTop:4}}>{weeks.length>=2?(weeks[weeks.length-1].fH>weeks[0].fH?`📈 Tendencia al alza: +${fN(weeks[weeks.length-1].fH-weeks[0].fH)} f/h`:`📉 Tendencia a la baja: ${fN(weeks[weeks.length-1].fH-weeks[0].fH)} f/h`):"Solo 1 semana"}</p></div>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>⏰ Actividad por Hora</h3>
        {hSede.length>0?<><ResponsiveContainer width="100%" height={160}><BarChart data={hSede}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="hora" tick={{fontSize:9}}/><YAxis tick={{fontSize:9}}/><Tooltip/><Bar dataKey="regs" name="Registros" radius={[3,3,0,0]}>{hSede.map(h=><Cell key={h.hora} fill={deadH.some(d=>d.hora===h.hora)?"#e74c3c40":CL.pri}/>)}</Bar></BarChart></ResponsiveContainer>
        <div style={{fontSize:11,marginTop:4}}>{peakH&&<span style={{color:CL.pri,fontWeight:700}}>🔥 Pico: {peakH.hora}</span>}{deadH.length>0&&<span style={{color:CL.dan,fontWeight:600,marginLeft:8}}>💤 Muertas: {deadH.map(d=>d.hora).join(", ")}</span>}</div></>
        :<p style={{fontSize:12,color:CL.txtL,padding:20,textAlign:"center"}}>Sin datos horarios disponibles</p>}</div></div>
    <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📊 Ranking F/H + Consistencia</h3>
      <ResponsiveContainer width="100%" height={280}><BarChart data={ch} margin={{left:8,right:8,bottom:50}}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="n" tick={{fontSize:9,fill:CL.txtL}} angle={-40} textAnchor="end" interval={0}/><YAxis tick={{fontSize:10}}/><Tooltip/><Bar dataKey="fH" name="F/H" radius={[5,5,0,0]}>{ch.map((_,i)=><Cell key={i} fill={rColor(i+1,data.cajeros.length)}/>)}</Bar></BarChart></ResponsiveContainer></div>
    <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📅 Sede Diaria</h3>
      <ResponsiveContainer width="100%" height={180}><ComposedChart data={sc}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="f" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/><Area type="monotone" dataKey="fH" fill={`${CL.pri}20`} stroke="none"/><Line type="monotone" dataKey="fH" stroke={CL.pri} strokeWidth={2.5} dot={{r:3}} name="F/H"/></ComposedChart></ResponsiveContainer></div>
    <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📋 Ranking {data.hasSched?"(R/H Marcación)":"(F/H Efectiva)"}</h3><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["#","Cajero","Reg","Fact","Dias Reg",...(data.hasSched?["Dias Mar","H.Mar","R/HMar"]:[]),"H.Efec","F/HEfec","Min/Fact","Consist","KPI"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>
      {data.cajeros.map((c,idx)=>{const mf=c.tF>0?((c.hrsEfec*60)/c.tF).toFixed(1):"—";const rc=rColor(c.rank,data.cajeros.length);const bg=c.kpi.lab==="Cumple"?"#f0faf4":c.kpi.lab==="No cumple"?"#fef5f5":"#fffcf0";return <tr key={c.nombre} style={{background:bg}}><td style={{...TD,fontWeight:800,color:rc}}>{c.rank}</td><td style={{...TD,fontWeight:600,fontSize:11}}>{c.nombre}{rBadge(c)}{c.diasFallidos>0&&<span style={{color:"#e67e22",fontSize:9,marginLeft:2}}>⚠{c.diasFallidos}f</span>}</td><td style={TD}>{c.tR.toLocaleString()}</td><td style={TD}>{c.tF.toLocaleString()}</td><td style={TD}>{c.numDias}</td>{data.hasSched&&<><td style={{...TD,color:"#8e44ad"}}>{c.diasMarValid||"-"}</td><td style={{...TD,fontSize:11,color:"#8e44ad"}}>{c.totalHrsMar?fH(c.totalHrsMar):"-"}</td><td style={{...TD,fontWeight:800,color:CL.pri,fontSize:13}}>{c.totalHrsMar>0?fN(c.rHMar):"-"}</td></>}<td style={{...TD,fontSize:11}}>{fH(c.hrsEfec)}</td><td style={{...TD,fontWeight:data.hasSched?400:700,color:data.hasSched?CL.txtL:CL.pri}}>{fN(c.pfH)}</td><td style={{...TD,fontSize:11,color:"#3498db"}}>{mf}</td><td style={TD}><Bg color={c.consist.cl}>{c.consist.label}</Bg></td><td style={TD}><Bg color={c.kpi.cl}>{c.kpi.ic} {c.kpi.lab}</Bg></td></tr>})}</tbody></table></div></div></div>;}

/* ═══ INDIVIDUAL ═══ */
function Indiv({data,sel,setSel,showR}){const c=data.cajeros.find(x=>x.nombre===sel);const[busq,setBusq]=useState("");
  const filtrados=busq?data.cajeros.filter(x=>x.nombre.toLowerCase().includes(busq.toLowerCase())):data.cajeros;
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
      <input type="text" placeholder="🔍 Buscar cajero..." value={busq} onChange={e=>setBusq(e.target.value)} style={{padding:"7px 12px",borderRadius:8,border:"2px solid #ddd",fontSize:12,width:160,outline:"none"}}/>
      <select style={SL} value={sel||""} onChange={e=>{setSel(e.target.value);setBusq("");}}><option value="">-- Seleccionar ({filtrados.length}) --</option>{filtrados.map(c=><option key={c.nombre} value={c.nombre}>#{c.rank} {c.nombre}{rBadge(c)}</option>)}</select>
      {c&&<Btn onClick={()=>showR({html:narInd(c,data),title:`Informe - ${c.nombre}`})}>📄 Informe</Btn>}{c&&<Btn green onClick={()=>exInd(c,data)}>📥 Excel</Btn>}</div>
    {c?<><div style={{...crd,background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,color:"#fff"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12}}><div><div style={{fontSize:10,opacity:.7}}>FICHA</div><h2 style={{margin:"4px 0",fontSize:20,fontWeight:800}}>{c.nombre}{rBadge(c)}</h2><div style={{fontSize:12,opacity:.85}}>{c.sede} - {c.cedula}</div></div><div style={{textAlign:"center",background:"rgba(255,255,255,.15)",borderRadius:12,padding:"10px 20px"}}><div style={{fontSize:30,fontWeight:900}}>#{c.rank}</div><div style={{fontSize:10,opacity:.8}}>de {data.cajeros.length}</div></div></div><div style={{marginTop:10,padding:"7px 12px",background:`${c.frase.cl}30`,borderRadius:8,fontSize:13}}>{c.frase.em} {c.frase.tx} - {c.kpi.ic} {c.kpi.lab}</div></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><KC label="Registros" value={c.tR.toLocaleString()} color="#9b59b6"/><KC label="Facturas" value={c.tF.toLocaleString()} color="#3498db"/><KC label="Dias Reg" value={c.numDias} color="#e67e22"/>
        {c.diasMarValid>0&&<KC label="Dias Mar" value={c.diasMarValid} color="#8e44ad"/>}
        <KC label="Hrs Efec" value={fH(c.hrsEfec)} color="#2ecc71"/>
        {c.totalHrsMar>0&&<KC label="Hrs Mar" value={fH(c.totalHrsMar)} color="#8e44ad"/>}
        {c.totalHrsMar>0&&<KC label="R/H Mar" value={fN(c.rHMar)} color={CL.pri} big/>}
        <KC label="F/H Efec" value={fN(c.pfH)} color={c.totalHrsMar>0?CL.txtL:CL.pri} big={!c.totalHrsMar}/>
        {c.hrsHor!==null&&c.hrsHor>0&&<KC label="Aprovech." value={((c.hrsEfec/c.hrsHor)*100).toFixed(1)+"%"} color={c.hrsEfec/c.hrsHor>=.9?"#2ecc71":"#e74c3c"}/>}
        {c.diasNoLab!==null&&<KC label="No Laborados" value={c.diasNoLab+" dias"} color={c.diasNoLab>3?"#e74c3c":"#f39c12"}/>}
        {c.diasFallidos>0&&<KC label="Marc. Fallidas" value={c.diasFallidos+" dias"} color="#e67e22"/>}
        <KC label="Consistencia" value={c.consist.label} color={c.consist.cl}/></div>
      {/* Audit efficiency panel */}
      <div style={{...crd,borderLeft:`4px solid ${CL.pri}`,background:CL.priLt}}>
        <div style={{fontSize:12,fontWeight:700,color:CL.priDk,marginBottom:6}}>📋 Eficiencia Detallada</div>
        <div style={{background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,color:"#fff",borderRadius:10,padding:"12px 16px",marginBottom:10,fontSize:13,lineHeight:1.8}}>
          {c.totalHrsMar>0?<>🏷️ <b style={{fontSize:18}}>{fN(c.rHMar)}</b> reg/h marcada | <b>{fN(c.fHMar)}</b> fact/h marcada | <b>{c.diasMarValid>0?(c.totalHrsMar/c.diasMarValid).toFixed(1):0}h</b> marc/dia
          <br/>⚡ Registra <b style={{fontSize:16}}>{c.totalHrsMar>0?Math.round(c.tR/(c.totalHrsMar*60)):0}</b> productos/min (hrs marcadas) | <b>{c.hrsEfec>0?Math.round(c.tR/(c.hrsEfec*60)):0}</b> productos/min (hrs efectivas)
          <br/>📦 <b>{c.numDias>0?Math.round(c.tR/c.numDias):0}</b> reg/dia en <b>{c.numDias>0?Math.round(c.tF/c.numDias):0}</b> fact | Dias reg: <b>{c.numDias}</b> | Dias mar: <b>{c.diasMarValid}</b></>
          :<>⚡ Registra <b style={{fontSize:18}}>{c.hrsEfec>0?Math.round(c.tR/(c.hrsEfec*60)):0}</b> productos/min
          <br/>📦 <b>{c.numDias>0?Math.round(c.tR/c.numDias):0}</b> reg/dia en <b>{c.numDias>0?Math.round(c.tF/c.numDias):0}</b> fact | Trabaja <b>{c.numDias>0?(c.hrsEfec/c.numDias).toFixed(1):0}h</b>/dia</>}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:10,padding:"10px 12px"}}>
            <div style={{fontSize:10,color:CL.txtL,fontWeight:700,marginBottom:6}}>⏱️ JORNADA</div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:CL.pri}}>{c.numDias>0?(c.hrsEfec/c.numDias).toFixed(1):0}h</div><div style={{color:CL.txtL,fontSize:9}}>Prom Hrs/Dia</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#2ecc71"}}>{fH(c.hrsEfec)}</div><div style={{color:CL.txtL,fontSize:9}}>Total Mes</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#e67e22"}}>{(c.hrsEfec/8).toFixed(1)}</div><div style={{color:CL.txtL,fontSize:9}}>Equiv Dias 8h</div></div></div></div>
          <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:10,padding:"10px 12px"}}>
            <div style={{fontSize:10,color:CL.txtL,fontWeight:700,marginBottom:6}}>⚡ VELOCIDAD</div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#9b59b6"}}>{c.hrsEfec>0?Math.round(c.tR/(c.hrsEfec*60)):0}</div><div style={{color:CL.txtL,fontSize:9}}>Reg/Min</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#3498db"}}>{c.hrsEfec>0?(c.tF/(c.hrsEfec*60)).toFixed(1):0}</div><div style={{color:CL.txtL,fontSize:9}}>Fact/Min</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:CL.dan}}>{c.numDias>0?Math.round(c.tR/c.numDias):0}</div><div style={{color:CL.txtL,fontSize:9}}>Reg/Dia</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:CL.pri}}>{c.numDias>0?Math.round(c.tF/c.numDias):0}</div><div style={{color:CL.txtL,fontSize:9}}>Fact/Dia</div></div></div></div>
          <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:10,padding:"10px 12px"}}>
            <div style={{fontSize:10,color:CL.txtL,fontWeight:700,marginBottom:6}}>🕐 TIEMPO ENTRE</div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#9b59b6"}}>{c.tR>0?((c.hrsEfec*60)/c.tR).toFixed(1):0} min</div><div style={{color:CL.txtL,fontSize:9}}>por Registro</div></div>
              <div style={{textAlign:"center",minWidth:60}}><div style={{fontSize:17,fontWeight:900,color:"#3498db"}}>{c.tF>0?((c.hrsEfec*60)/c.tF).toFixed(1):0} min</div><div style={{color:CL.txtL,fontSize:9}}>por Factura</div></div></div></div></div></div>
      {c.mD&&<div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div style={{...crd,borderLeft:"4px solid #2ecc71"}}><div style={{fontSize:11,color:CL.txtL,fontWeight:600}}>🏆 Mejor</div><div style={{fontSize:16,fontWeight:800,color:CL.pri}}>{c.mD.fechaF} ({c.mD.diaSemL})</div><div style={{fontSize:12}}>{fN(c.mD.fH)} f/h - {c.mD.facs} fact</div></div>
        <div style={{...crd,borderLeft:"4px solid #e74c3c"}}><div style={{fontSize:11,color:CL.txtL,fontWeight:600}}>📉 Peor</div><div style={{fontSize:16,fontWeight:800,color:CL.dan}}>{c.pDy.fechaF} ({c.pDy.diaSemL})</div><div style={{fontSize:12}}>{fN(c.pDy.fH)} f/h - {c.pDy.facs} fact</div></div></div>}
      <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>📈 F/H</h3><ResponsiveContainer width="100%" height={170}><ComposedChart data={c.dias}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="fechaD" tick={{fontSize:9}}/><YAxis tick={{fontSize:9}}/><Tooltip/><Area type="monotone" dataKey="fH" fill={`${CL.pri}20`} stroke="none"/><Line type="monotone" dataKey="fH" stroke={CL.pri} strokeWidth={2.5} dot={{r:3,fill:CL.pri}} name="F/H"/></ComposedChart></ResponsiveContainer></div>
        <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>📦 R/H</h3><ResponsiveContainer width="100%" height={170}><ComposedChart data={c.dias}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="fechaD" tick={{fontSize:9}}/><YAxis tick={{fontSize:9}}/><Tooltip/><Area type="monotone" dataKey="rH" fill="#3498db20" stroke="none"/><Line type="monotone" dataKey="rH" stroke="#3498db" strokeWidth={2.5} dot={{r:3}} name="R/H"/></ComposedChart></ResponsiveContainer></div></div>
      <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>⏰ Por Hora</h3><ResponsiveContainer width="100%" height={150}><BarChart data={c.hA.slice().sort((a,b)=>a.hora-b.hora)}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="horaStr" tick={{fontSize:9}}/><YAxis tick={{fontSize:9}}/><Tooltip/><Bar dataKey="prom" name="Prom Reg/Dia" fill={CL.pri} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer></div>
      {/* Gap / Tiempo muerto analysis */}
      {c.hA.length>0&&(()=>{const sorted=c.hA.slice().sort((a,b)=>a.hora-b.hora);const minH=sorted[0].hora,maxH=sorted[sorted.length-1].hora;
        const activeHrs=new Set(sorted.map(h=>h.hora));const gapHrs=[];for(let h=minH+1;h<maxH;h++){if(!activeHrs.has(h))gapHrs.push(h);}
        const pctActivo=maxH>minH?((activeHrs.size/(maxH-minH+1))*100).toFixed(0):100;
        const totalRegMin=c.hrsEfec>0?(c.tR/(c.hrsEfec*60)).toFixed(1):0;
        const difMarcEfec=c.totalHrsMar>0?(c.totalHrsMar-c.hrsEfec).toFixed(1):null;
        return <div style={{...crd,borderLeft:"4px solid #e67e22",background:"#fffcf0"}}>
          <h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700,color:"#e67e22"}}>🔍 Analisis de Actividad</h3>
          <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:12,marginBottom:8}}>
            <div style={{textAlign:"center",minWidth:80}}><div style={{fontSize:20,fontWeight:900,color:+pctActivo>=80?"#2ecc71":+pctActivo>=60?"#f39c12":"#e74c3c"}}>{pctActivo}%</div><div style={{color:CL.txtL,fontSize:10}}>Hrs con actividad</div></div>
            <div style={{textAlign:"center",minWidth:80}}><div style={{fontSize:20,fontWeight:900,color:"#3498db"}}>{activeHrs.size}h</div><div style={{color:CL.txtL,fontSize:10}}>Horas activas</div></div>
            <div style={{textAlign:"center",minWidth:80}}><div style={{fontSize:20,fontWeight:900,color:gapHrs.length>0?"#e74c3c":"#2ecc71"}}>{gapHrs.length}h</div><div style={{color:CL.txtL,fontSize:10}}>Horas sin registro</div></div>
            {difMarcEfec&&<div style={{textAlign:"center",minWidth:80}}><div style={{fontSize:20,fontWeight:900,color:+difMarcEfec>2?"#e74c3c":"#f39c12"}}>{difMarcEfec}h</div><div style={{color:CL.txtL,fontSize:10}}>Dif Marc vs Efec</div></div>}
          </div>
          {gapHrs.length>0&&<div style={{fontSize:11,color:"#e67e22",padding:"6px 10px",background:"#fff",borderRadius:6}}>
            ⚠️ Horas sin registros: {gapHrs.map(h=><span key={h} style={{display:"inline-block",padding:"1px 6px",margin:"1px",borderRadius:4,background:"#fde8e8",fontWeight:600}}>{h}:00</span>)}
            <span style={{color:CL.txtL,marginLeft:6}}>(entre primera y ultima actividad del periodo)</span></div>}
          {difMarcEfec&&+difMarcEfec>1&&<div style={{fontSize:11,color:"#e74c3c",marginTop:6,padding:"6px 10px",background:"#fff",borderRadius:6}}>
            🔴 Marcacion indica <b>{fH(c.totalHrsMar)}</b> en tienda pero solo registra <b>{fH(c.hrsEfec)}</b> efectivas. Diferencia: <b>{difMarcEfec}h</b> sin actividad de registro.</div>}
        </div>;})()}
      <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>📅 Detalle Diario</h3><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Fecha","Dia","Reg","Fact","Ini","Fin","HrsEfec",...(data.hasSched?["HrsMar","Dif","R/HMar"]:[]),"F/HEfec","Min/Fact","Min/Reg"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>
        {c.dias.map(d=>{const mf=d.facs>0?((d.hrs*60)/d.facs).toFixed(1):"—";const mr=d.regs>0?((d.hrs*60)/d.regs).toFixed(1):"—";const dif=d.hrsMar!=null?(d.hrsMar-d.hrs).toFixed(1):null;return <tr key={d.fecha} style={{background:d.hasFallido?"#fff3e0":"transparent"}}><td style={{...TD,fontWeight:600}}>{d.fechaD}</td><td style={{...TD,fontSize:11}}>{d.diaSem}</td><td style={TD}>{d.regs}</td><td style={TD}>{d.facs}</td><td style={{...TD,fontSize:11}}>{d.hi}</td><td style={{...TD,fontSize:11}}>{d.hf}</td><td style={TD}>{fH(d.hrs)}{d.minH?" ⚠️":""}</td>
          {data.hasSched&&<><td style={{...TD,fontSize:11,color:"#8e44ad",fontWeight:600}}>{d.hrsMar!=null?fH(d.hrsMar):"—"}{d.hasFallido&&<span style={{color:"#e67e22",fontSize:9}}> ⚠F</span>}</td>
          <td style={{...TD,fontSize:11,color:dif>0?"#2ecc71":"#e74c3c"}}>{dif!=null?(dif>0?"+":"")+dif+"h":"—"}</td>
          <td style={{...TD,fontWeight:700,color:CL.pri}}>{d.rHMar!=null?fN(d.rHMar):"—"}</td></>}
          <td style={{...TD,color:d.fH>=data.avg?CL.pri:CL.dan}}>{fN(d.fH)}</td><td style={{...TD,fontSize:11,color:"#3498db"}}>{mf}</td><td style={{...TD,fontSize:11,color:"#9b59b6"}}>{mr}</td></tr>})}</tbody></table></div></div>
    </>:<div style={{...crd,textAlign:"center",padding:46,color:CL.txtL}}><div style={{fontSize:42,marginBottom:8}}>👤</div><p>Selecciona un cajero</p></div>}</div>;}

/* ═══ COMPARE ═══ */
function Cmp({data}){const[s1,sS1]=useState(""),[s2,sS2]=useState("");const c1=data.cajeros.find(x=>x.nombre===s1),c2=data.cajeros.find(x=>x.nombre===s2);
  const mt=c1&&c2?[{l:"Ranking",a:`#${c1.rank}`,b:`#${c2.rank}`,w:c1.rank<c2.rank?1:c2.rank<c1.rank?2:0},{l:"F/H",a:fN(c1.pfH),b:fN(c2.pfH),w:c1.pfH>c2.pfH?1:2,key:1},{l:"R/H",a:fN(c1.prH),b:fN(c2.prH),w:c1.prH>c2.prH?1:2},{l:"Consist.",a:c1.consist.label,b:c2.consist.label,w:c1.consist.cv<c2.consist.cv?1:2},{l:"KPI",a:`${c1.kpi.ic} ${c1.kpi.lab}`,b:`${c2.kpi.ic} ${c2.kpi.lab}`,w:0}]:[];
  const dts=c1&&c2?[...new Set([...c1.dias.map(x=>x.fecha),...c2.dias.map(x=>x.fecha)])].sort():[];
  const cd=dts.map(f=>{const d1=c1.dias.find(x=>x.fecha===f),d2=c2.dias.find(x=>x.fecha===f);const o={f:fD(new Date(f+"T12:00:00"))};o[sN(c1.nombre)]=d1?+d1.fH.toFixed(2):null;o[sN(c2.nombre)]=d2?+d2.fH.toFixed(2):null;return o;});
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{...crd,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}><select style={SL} value={s1} onChange={e=>sS1(e.target.value)}><option value="">-- Cajero 1 --</option>{data.cajeros.map(c=><option key={c.nombre} value={c.nombre}>#{c.rank} {c.nombre}{rBadge(c)}</option>)}</select><span style={{fontSize:20,fontWeight:800,color:CL.pri}}>VS</span>
      <select style={SL} value={s2} onChange={e=>sS2(e.target.value)}><option value="">-- Cajero 2 --</option>{data.cajeros.map(c=><option key={c.nombre} value={c.nombre}>#{c.rank} {c.nombre}{rBadge(c)}</option>)}</select>{c1&&c2&&<Btn green onClick={()=>exCmp(c1,c2)}>📥 Excel</Btn>}</div>
    {c1&&c2?<><div style={crd}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><th style={{...TH,textAlign:"right",width:"35%"}}>{sN(c1.nombre)}</th><th style={{...TH,textAlign:"center",width:"30%"}}>Metrica</th><th style={{...TH,width:"35%"}}>{sN(c2.nombre)}</th></tr></thead>
      <tbody>{mt.map(m=><tr key={m.l} style={{background:m.key?CL.priLt:"transparent"}}><td style={{...TD,textAlign:"right",fontWeight:m.key?800:600,color:m.w===1?CL.pri:m.w===2?CL.dan:CL.txt,fontSize:m.key?16:13}}>{m.w===1?"✓ ":""}{m.a}</td><td style={{...TD,textAlign:"center",fontSize:11,color:CL.txtL,fontWeight:700}}>{m.l}</td><td style={{...TD,fontWeight:m.key?800:600,color:m.w===2?CL.pri:m.w===1?CL.dan:CL.txt,fontSize:m.key?16:13}}>{m.b}{m.w===2?" ✓":""}</td></tr>)}</tbody></table></div>
      <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:14,fontWeight:700}}>📈 F/H</h3><ResponsiveContainer width="100%" height={200}><LineChart data={cd}><CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis dataKey="f" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/><Legend/><Line type="monotone" dataKey={sN(c1.nombre)} stroke={CL.pri} strokeWidth={2.5} dot={{r:3}} connectNulls/><Line type="monotone" dataKey={sN(c2.nombre)} stroke="#e74c3c" strokeWidth={2.5} dot={{r:3}} connectNulls/></LineChart></ResponsiveContainer></div>
    </>:<div style={{...crd,textAlign:"center",padding:46,color:CL.txtL}}><div style={{fontSize:42}}>⚔️</div><p>Selecciona dos cajeros</p></div>}</div>;}

/* ═══ ALERTS ═══ */
function Alt({data}){const no=data.cajeros.filter(c=>c.kpi.lab==="No cumple"),at=data.cajeros.filter(c=>c.kpi.lab==="En promedio");const kt=kpiT(data);
  // Smart alerts: detect patterns
  const patterns=useMemo(()=>{const al=[];const dias=["Dom","Lun","Mar","Mie","Jue","Vie","Sab"];
    data.cajeros.forEach(c=>{if(c.dias.length<3)return;
      // Day-of-week pattern: best and worst day
      const byDay={};c.dias.forEach(d=>{const dw=new Date(d.fecha+"T12:00:00").getDay();if(!byDay[dw])byDay[dw]=[];byDay[dw].push(d.fH);});
      const dayAvgs=Object.entries(byDay).map(([d,v])=>({d:+d,avg:v.reduce((a,b)=>a+b,0)/v.length,n:v.length})).filter(x=>x.n>=2);
      if(dayAvgs.length>=2){const best=dayAvgs.reduce((a,b)=>a.avg>b.avg?a:b),worst=dayAvgs.reduce((a,b)=>a.avg<b.avg?a:b);
        if(best.avg-worst.avg>3)al.push({cajero:c.nombre,tipo:"📅",msg:`Rinde ${fN(best.avg)} f/h los ${dias[best.d]} pero baja a ${fN(worst.avg)} los ${dias[worst.d]}`,sev:2});}
      // Declining trend: last 3 days dropping
      if(c.dias.length>=4){const last4=c.dias.slice(-4);if(last4[3].fH<last4[2].fH&&last4[2].fH<last4[1].fH&&last4[1].fH<last4[0].fH)
        al.push({cajero:c.nombre,tipo:"📉",msg:`En caida: ${last4.map(d=>fN(d.fH)).join(" → ")} (ultimos 4 dias)`,sev:3});}
      // Outlier day: one day way below their own average
      const myAvg=c.pfH;c.dias.forEach(d=>{if(d.fH<myAvg*OUTLIER_THRESHOLD&&myAvg>5)al.push({cajero:c.nombre,tipo:"⚡",msg:`Dia atipico ${d.fechaF}: solo ${fN(d.fH)} f/h (su prom es ${fN(myAvg)})`,sev:1});});
      // High variability warning
      if(c.consist.cv>50)al.push({cajero:c.nombre,tipo:"🎲",msg:`Muy variable (CV ${c.consist.cv.toFixed(0)}%): rendimiento impredecible`,sev:2});
    });return al.sort((a,b)=>b.sev-a.sev);},[data]);
  return <div style={{display:"flex",flexDirection:"column",gap:14}}><KPIBanner data={data}/>
    <div style={{...crd,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,background:no.length>0?"#fde8e8":"#e6f9ee"}}>
      <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:24}}>{no.length>0?"🚨":"✅"}</span><div><div style={{fontSize:16,fontWeight:700}}>{no.length>0?`${no.length} NO cumplen`:"Todos OK!"}</div><div style={{fontSize:12,color:CL.txtL}}>Cumple ≥ {fN(kt.cumple)} | En prom ≥ {fN(kt.enProm)}</div></div></div>
      <Btn green onClick={()=>exAlt(data)}>📥 Excel</Btn></div>
    {patterns.length>0&&<div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700,color:"#8e44ad"}}>🔔 Alertas Inteligentes</h3>
      <p style={{fontSize:11,color:CL.txtL,marginBottom:8}}>Patrones detectados automaticamente en los datos</p>
      {patterns.map((p,i)=><div key={i} style={{display:"flex",gap:10,padding:"8px 12px",marginBottom:4,background:p.sev>=3?"#fde8e8":p.sev>=2?"#fff9e6":"#f0f6ff",borderRadius:8,borderLeft:`3px solid ${p.sev>=3?"#e74c3c":p.sev>=2?"#f39c12":"#3498db"}`,alignItems:"center"}}>
        <span style={{fontSize:18}}>{p.tipo}</span>
        <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700}}>{sN(p.cajero)}</div><div style={{fontSize:11,color:CL.txtL}}>{p.msg}</div></div>
        <Bg color={p.sev>=3?"#e74c3c":p.sev>=2?"#f39c12":"#3498db"}>{p.sev>=3?"Critico":p.sev>=2?"Atencion":"Info"}</Bg></div>)}
      {patterns.length===0&&<p style={{fontSize:12,color:CL.txtL}}>No se detectaron patrones inusuales.</p>}</div>}
    {no.length>0&&<div style={crd}><h3 style={{margin:"0 0 8px",fontSize:14,fontWeight:700,color:CL.dan}}>❌ No Cumplen</h3><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["#","Cajero","F/H","Consist","vs Prom"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>{no.map(c=><tr key={c.nombre} style={{background:"#fde8e8"}}><td style={{...TD,fontWeight:800,color:CL.dan}}>{c.rank}</td><td style={{...TD,fontWeight:600}}>{c.nombre}{rBadge(c)}</td><td style={{...TD,fontWeight:700,color:CL.dan}}>{fN(c.pfH)}</td><td style={TD}><Bg color={c.consist.cl}>{c.consist.label}</Bg></td><td style={{...TD,color:CL.dan}}>{fN(c.pfH-data.avg)}</td></tr>)}</tbody></table></div></div>}
    {at.length>0&&<div style={crd}><h3 style={{margin:"0 0 8px",fontSize:14,fontWeight:700,color:CL.warn}}>😐 En Promedio</h3><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["#","Cajero","F/H","Falta"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>{at.map(c=><tr key={c.nombre} style={{background:"#fff9e6"}}><td style={{...TD,fontWeight:800,color:CL.warn}}>{c.rank}</td><td style={{...TD,fontWeight:600}}>{c.nombre}{rBadge(c)}</td><td style={{...TD,fontWeight:700}}>{fN(c.pfH)}</td><td style={{...TD,color:CL.warn}}>+{fN(kt.cumple-c.pfH)}</td></tr>)}</tbody></table></div></div>}</div>;}

/* ═══ REPORTS ═══ */
function Rpt({data,showR}){const[s,sS]=useState("");const c=data.cajeros.find(x=>x.nombre===s);const kt=kpiT(data);
  const met=[{t:"📦 Registros",x:"Cada fila del archivo = 1 producto escaneado por cajero."},{t:"🧾 Facturas",x:"Valores UNICOS de Nro documento. 1 factura puede tener muchos productos."},{t:"📅 Dias",x:"Dias distintos con al menos 1 registro."},{t:"⏰ Horas Efectivas",x:"Ultimo registro - Primer registro por dia. Minimo 1h. Se suman todos los dias."},{t:"📅 Horas Horario",x:"Del archivo de horarios subido (opcional). Horas programadas/asignadas."},{t:"📊 Aprovechamiento",x:"= Horas Efectivas / Horas Horario x 100. Solo si se sube archivo de horarios."},{t:"⚡ Fact/Hora (PRINCIPAL)",x:"= Total Facturas / Total Horas Efectivas. Es el indicador clave."},{t:"📦 Reg/Hora",x:"= Total Registros / Total Horas Efectivas."},{t:"📦 Reg/Factura",x:"= Total Registros / Total Facturas. Productos promedio por factura."},{t:"🏆 Ranking",x:"Ordenamiento de mayor a menor por Fact/Hora."},{t:"🎯 Consistencia (CV%)",x:"Coeficiente de variacion de F/H diaria. Mide que tan estable es el cajero dia a dia. Muy estable (<=15%), Estable (<=30%), Variable (<=50%), Muy variable (>50%)."},{t:`📏 Prom Sede: ${fN(data.avg)}`,x:"= Suma de F/H de cada cajero / Total cajeros."},{t:`✅ Cumple: >= ${fN(kt.cumple)}`,x:kt.custom?"KPI personalizado definido por el evaluador.":"F/H >= Promedio + 2 (calculado automaticamente)."},{t:`😐 En promedio: >= ${fN(kt.enProm)}`,x:kt.custom?"KPI personalizado definido por el evaluador.":`F/H >= ${fN(kt.enProm)} pero menor que ${fN(kt.cumple)}.`},{t:`❌ No Cumple: < ${fN(kt.enProm)}`,x:kt.custom?"KPI personalizado definido por el evaluador.":`F/H por debajo de ${fN(kt.enProm)}.`},{t:"💾 Archivo de Memoria (.json)",x:"Guarda una 'foto' del analisis actual. Se puede cargar despues en Periodos para comparar con un nuevo analisis o comparar entre sedes."},{t:"⚠️ Marcacion Fallida",x:"Dias donde el empleado marco el reloj pero la marcacion salio fallida. La app usa la hora igualmente y deduce si fue entrada, salida o break por su posicion en el dia."}];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:15,fontWeight:700}}>📊 Informe General</h3><div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn onClick={()=>showR({html:narGen(data),title:`General - ${data.sede}`})}>📄 Ver</Btn><Btn green onClick={()=>exAll(data)}>📥 Excel</Btn></div></div>
    <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:15,fontWeight:700}}>👤 Individual</h3><div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}><select style={SL} value={s} onChange={e=>sS(e.target.value)}><option value="">-- Seleccionar --</option>{data.cajeros.map(c=><option key={c.nombre} value={c.nombre}>#{c.rank} {c.nombre}{rBadge(c)}</option>)}</select>
      {c&&<Btn onClick={()=>showR({html:narInd(c,data),title:`Informe - ${c.nombre}`})}>📄 Ver</Btn>}{c&&<Btn green onClick={()=>exInd(c,data)}>📥 Excel</Btn>}</div></div>
    <div style={{...crd,background:"#f0faf4",border:`1px solid ${CL.pri}40`}}><h3 style={{margin:"0 0 8px",fontSize:15,fontWeight:700,color:CL.pri}}>📦 Descargar TODOS los Informes</h3>
      <p style={{fontSize:13,color:CL.txtL,marginBottom:10}}>Genera un archivo con los {data.cajeros.length} informes narrativos.</p>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn onClick={()=>dlReport(narAllInd(data),`Todos_Informes_${data.sede}`)}>📥 Todos ({data.cajeros.length})</Btn><Btn green onClick={()=>dlReport(narGen(data),`General_${data.sede}`)}>📥 General</Btn></div></div>
    <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:15,fontWeight:700}}>🚨 Alertas</h3><Btn green onClick={()=>exAlt(data)}>📥 Excel Alertas</Btn></div>
    <div style={{...crd,border:"2px solid #3498db40",background:"#f0f6ff"}}><h3 style={{margin:"0 0 12px",fontSize:16,fontWeight:800,color:"#2c3e50"}}>📖 Manual de Metodologia</h3>
      <p style={{fontSize:13,color:CL.txtL,margin:"0 0 14px"}}>Asi se calculan todos los indicadores de este sistema. Importante para que cualquier persona entienda las reglas del juego.</p>
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {met.map((m,i)=><div key={i} style={{display:"flex",gap:10,padding:"8px 12px",background:i%2===0?"#fff":"#f8fafc",borderRadius:8,border:"1px solid #e8ecf0"}}>
          <div style={{fontWeight:700,fontSize:13,minWidth:200,color:"#2c3e50"}}>{m.t}</div>
          <div style={{fontSize:12,color:CL.txtL,flex:1}}>{m.x}</div></div>)}
      </div>
      <div style={{marginTop:14,padding:"10px 14px",background:"#e8f4fd",borderRadius:8,fontSize:12,color:"#2c3e50",border:"1px solid #b8daff"}}>
        💡 <b>Nota:</b> Los KPIs pueden ser automaticos (basados en promedio) o personalizados (definidos en ⚙️). El archivo de horarios es opcional y agrega metricas de aprovechamiento y asistencia. El archivo de memoria (.json) permite comparar periodos y sedes sin perder datos.
      </div></div></div>;}

/* ═══ SNAPSHOT SAVE/LOAD (MEMORY) ═══ */
function saveSnapshot(data){
  const snap={_type:"audit_snapshot",_v:5,sede:data.sede,periodo:data.periodo,avg:data.avg,avgR:data.avgR,avgRHMar:data.avgRHMar||0,kpiCfg:data.kpiCfg,tpvStats:data.tpvStats||[],
    cajeros:data.cajeros.map(c=>({nombre:c.nombre,cedula:c.cedula,sede:c.sede,tF:c.tF,tR:c.tR,numDias:c.numDias,hrsEfec:c.hrsEfec,hrsHor:c.hrsHor,totalHrsMar:c.totalHrsMar||0,diasMarValid:c.diasMarValid||0,diasProg:c.diasProg,diasNoLab:c.diasNoLab,diasFallidos:c.diasFallidos||0,activo:c.activo,cargo:c.cargo||"",ccosto:c.ccosto||"",pfH:c.pfH,prH:c.prH,prF:c.prF,rHMar:c.rHMar||0,fHMar:c.fHMar||0,rank:c.rank,consist:{cv:c.consist.cv,label:c.consist.label},kpiLab:c.kpi.lab,
      hA:c.hA||[],dA:c.dA||[],
      dias:c.dias.map(d=>({fecha:d.fecha,fechaD:d.fechaD,fechaF:d.fechaF,diaSem:d.diaSem,diaSemL:d.diaSemL,facs:d.facs,regs:d.regs,hrs:d.hrs,fH:d.fH,rH:d.rH,rF:d.rF,minH:d.minH||false,hi:d.hi,hf:d.hf,hrsMar:d.hrsMar!=null?d.hrsMar:null,hasFallido:d.hasFallido||false,breakMin:d.breakMin||0,rHMar:d.rHMar||null,fHMar:d.fHMar||null}))}))};
  const blob=new Blob([JSON.stringify(snap)],{type:"application/json"});const url=URL.createObjectURL(blob);const a=document.createElement("a");
  a.href=url;a.download=`Auditoria_${data.sede.replace(/\s/g,"_")}_${data.periodo.desde}_a_${data.periodo.hasta}.json`;
  document.body.appendChild(a);a.click();setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},300);toast("✅ Memoria guardada");}

const SNAP_VERSION=5;
function migrateSnapshot(s){
  if(!s._v||s._v<2){s.cajeros.forEach(c=>{c.diasFallidos=c.diasFallidos||0;c.consist=c.consist||{cv:0,label:"N/A"};});}
  if(!s._v||s._v<3){s.avgR=s.avgR||0;s.cajeros.forEach(c=>{c.prF=c.prF||0;});}
  if(!s._v||s._v<4){s.tpvStats=s.tpvStats||[];s.cajeros.forEach(c=>{c.activo=c.activo!=null?c.activo:null;c.cargo=c.cargo||"";c.hA=c.hA||[];c.dA=c.dA||[];c.dias.forEach(d=>{d.minH=d.minH||false;});});}
  if(!s._v||s._v<5){s.avgRHMar=s.avgRHMar||0;s.cajeros.forEach(c=>{c.totalHrsMar=c.totalHrsMar||0;c.diasMarValid=c.diasMarValid||0;c.rHMar=c.rHMar||0;c.fHMar=c.fHMar||0;c.dias.forEach(d=>{d.hrsMar=d.hrsMar!=null?d.hrsMar:null;d.hasFallido=d.hasFallido||false;d.breakMin=d.breakMin||0;d.rHMar=d.rHMar||null;d.fHMar=d.fHMar||null;});});}
  s._v=SNAP_VERSION;return s;
}

function loadSnapshot(json){
  const s=json;if(!s._type||s._type!=="audit_snapshot")return null;
  migrateSnapshot(s);
  const cajeros=s.cajeros.map(c=>({...c,consist:c.consist||{cv:0,label:"N/A",cl:"#95a5a6"},frase:getPhrase(c.rank,s.cajeros.length),kpi:getKPI(c.pfH,s.avg,s.kpiCfg),
    mD:c.dias.length>0?c.dias.reduce((a,b)=>a.fH>b.fH?a:b):null,pDy:c.dias.length>0?c.dias.reduce((a,b)=>a.fH<b.fH?a:b):null,hA:c.hA||[],dA:c.dA||[]}));
  cajeros.forEach(c=>{c.consist.cl=c.consist.cv<=15?"#2ecc71":c.consist.cv<=30?"#3498db":c.consist.cv<=50?"#f39c12":"#e74c3c";});
  const allDates=[...new Set(cajeros.flatMap(c=>c.dias.map(d=>d.fecha)))].sort();
  const dS=allDates.map(dk=>{let f=0,r=0,h=0,a=0;cajeros.forEach(c=>{const d=c.dias.find(x=>x.fecha===dk);if(d){f+=d.facs;r+=d.regs;h+=d.hrs;a++;}});return{fecha:dk,fechaD:fD(new Date(dk+"T12:00:00")),facs:f,regs:r,hrs:h,activos:a,fH:h>0?f/h:0};});
  return{cajeros,sede:s.sede,avg:s.avg,avgR:s.avgR,avgRHMar:s.avgRHMar||0,hasSched:cajeros.some(c=>c.hrsHor!==null||c.totalHrsMar>0),tR:cajeros.reduce((a,c)=>a+c.tR,0),tF:cajeros.reduce((a,c)=>a+c.tF,0),periodo:s.periodo,dS,allDates,kpiCfg:s.kpiCfg||null,tpvStats:s.tpvStats||[]};}

/* ═══ PERIOD / SEDE COMPARE ═══ */
function PComp({data}){const r2=useRef();const[d2,setD2]=useState(null);const[st,setSt]=useState("");const[mode,setMode]=useState(null);
  const load2=useCallback(async f=>{setSt("Cargando...");try{
    const buf=await f.arrayBuffer();const nm=f.name.toLowerCase();let d;
    if(nm.endsWith(".json")){const txt=new TextDecoder().decode(buf);const json=JSON.parse(txt);d=loadSnapshot(json);if(!d){setSt("Error: no es un archivo de memoria valido.");return;}}
    else{let j;if(nm.endsWith(".csv")){let txt;try{txt=new TextDecoder("utf-8",{fatal:true}).decode(buf);}catch(e){txt=new TextDecoder("latin1").decode(buf);}
      j=[];let cur="",inQ=false;const pushRow=()=>{j.push(cur.split(",").map(c=>{c=c.trim();if(c.startsWith('"')&&c.endsWith('"'))c=c.slice(1,-1).replace(/""/g,'"');return c;}));cur="";};
      for(const ch of txt){if(ch==='"')inQ=!inQ;else if(ch==="\n"&&!inQ){pushRow();}else{cur+=ch;}}if(cur.trim())pushRow();
    }else{const wb=XLSX.read(buf,{type:"array",cellDates:false,raw:false});j=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:""});}d=processData(j,null);}
    if(data.kpiCfg){d.kpiCfg=data.kpiCfg;d.cajeros.forEach(c=>{c.kpi=getKPI(c.pfH,d.avg,d.kpiCfg);});}
    const same=d.sede.toUpperCase().trim()===data.sede.toUpperCase().trim();setMode(same?"periodo":"sede");setD2(d);
    setSt(same?`📅 Misma sede detectada → Comparacion de periodos | ${d.cajeros.length} cajeros | ${d.periodo.desde} a ${d.periodo.hasta}`:`🏢 Sede diferente detectada: ${d.sede} → Comparacion entre sedes | ${d.cajeros.length} cajeros`);}catch(e){setSt("Error: "+e.message);}}, [data]);

  const comp=useMemo(()=>{if(!d2||mode!=="periodo")return null;
    return data.cajeros.map(c=>{const c2=d2.cajeros.find(x=>x.nombre===c.nombre);
      if(!c2)return{...c,prev:null,dFH:null,dRk:null,dReg:null,dDias:null,dRHM:null};
      return{...c,prev:c2,dFH:c.pfH-c2.pfH,dRk:c2.rank-c.rank,dReg:c.tR-c2.tR,dDias:c.numDias-c2.numDias,dRHM:c.rHMar&&c2.rHMar?c.rHMar-c2.rHMar:null};}).sort((a,b)=>(b.dFH||0)-(a.dFH||0));},[data,d2,mode]);

  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={crd}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8,marginBottom:10}}>
        <div><h3 style={{margin:"0 0 4px",fontSize:15,fontWeight:700}}>📈 Comparar Periodos</h3>
          <p style={{fontSize:12,color:CL.txtL,margin:0}}>Sube un <b>.json</b> (guardado con 💾) o <b>.xlsx</b> de otro periodo.</p></div>
        <Btn green onClick={()=>saveSnapshot(raw)}>💾 Guardar Memoria</Btn></div>
      <p style={{fontSize:12,color:CL.txtL,marginBottom:8}}>📍 Actual: <b>{data.sede}</b> | {data.periodo.desde} a {data.periodo.hasta} | {data.cajeros.length} cajeros</p>
      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <button onClick={()=>r2.current?.click()} style={{padding:"10px 20px",borderRadius:10,border:`2px dashed ${d2?CL.pri:CL.pri}`,background:d2?CL.priLt:"#fff",color:CL.pri,fontSize:13,fontWeight:700,cursor:"pointer"}}>{d2?`✅ ${d2.sede} (${d2.cajeros.length} caj)`:"📁 Subir .json o .xlsx"}</button>
        <input ref={r2} type="file" accept=".xlsx,.xls,.xlsm,.csv,.json" style={{display:"none"}} onChange={e=>{if(e.target.files[0])load2(e.target.files[0]);e.target.value="";}}/>
        {d2&&<button onClick={()=>{setD2(null);setMode(null);setSt("");}} style={{padding:"6px 12px",borderRadius:8,border:"1px solid #e74c3c",background:"#fff",color:"#e74c3c",fontSize:11,cursor:"pointer",fontWeight:600}}>✕ Quitar</button>}
      </div>
      {st&&<p style={{fontSize:12,color:CL.pri,marginTop:8,fontWeight:600}}>{st}</p>}</div>

    {comp&&mode==="periodo"&&<>
      {/* Resumen comparativo */}
      <div style={{...crd,background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,color:"#fff",padding:"16px 20px"}}>
        <h3 style={{margin:"0 0 10px",fontSize:16,fontWeight:800}}>📊 {d2.periodo.desde} → {data.periodo.desde}</h3>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10}}>
          {[{l:"Facturas",a:d2.tF,b:data.tF},{l:"Registros",a:d2.tR,b:data.tR},{l:"Cajeros",a:d2.cajeros.length,b:data.cajeros.length},{l:"Prom F/H",a:d2.avg,b:data.avg,dec:true},{l:"Dias",a:d2.dS.length,b:data.dS.length}].map(m=>{const d=m.dec?(m.b-m.a).toFixed(1):m.b-m.a;const up=d>0;return <div key={m.l} style={{background:"rgba(255,255,255,.12)",borderRadius:8,padding:"8px 10px",textAlign:"center"}}>
            <div style={{fontSize:10,opacity:.7}}>{m.l}</div>
            <div style={{fontSize:11}}>{m.dec?fN(m.a):m.a.toLocaleString()} → <b>{m.dec?fN(m.b):m.b.toLocaleString()}</b></div>
            <div style={{fontSize:13,fontWeight:800,color:up?"#7dff7d":"#ff9999"}}>{up?"+":""}{m.dec?d:d.toLocaleString()}</div></div>})}</div></div>

      {/* Tabla comparativa completa */}
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📋 Comparativa por Cajero</h3>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["#","Cajero","Reg Ant","Reg Act","ΔReg","Dias Ant","Dias Act","F/H Ant","F/H Act","ΔF/H","Rank Ant","Rank Act","Mov"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
          <tbody>{comp.map(c=>{const up=c.dFH>0,dn=c.dFH<0;return <tr key={c.nombre} style={{background:c.prev?up?"#f0faf4":dn?"#fef5f5":"#fffcf0":"#f8f8f8"}}>
            <td style={{...TD,fontWeight:800}}>{c.rank}</td><td style={{...TD,fontWeight:600,fontSize:11}}>{c.nombre}{rBadge(c)}</td>
            <td style={{...TD,fontSize:11}}>{c.prev?c.prev.tR.toLocaleString():"-"}</td><td style={TD}>{c.tR.toLocaleString()}</td>
            <td style={{...TD,fontSize:11,fontWeight:600,color:c.dReg>0?"#2ecc71":c.dReg<0?"#e74c3c":"#666"}}>{c.dReg!==null?(c.dReg>0?"+":"")+c.dReg.toLocaleString():"-"}</td>
            <td style={{...TD,fontSize:11}}>{c.prev?c.prev.numDias:"-"}</td><td style={TD}>{c.numDias}</td>
            <td style={{...TD,fontSize:11}}>{c.prev?fN(c.prev.pfH):"-"}</td><td style={{...TD,fontWeight:700}}>{fN(c.pfH)}</td>
            <td style={{...TD,fontWeight:800,color:up?"#2ecc71":dn?"#e74c3c":"#666",fontSize:13}}>{c.dFH!==null?`${up?"+":""}${fN(c.dFH)}`:"-"}</td>
            <td style={{...TD,fontSize:11}}>{c.prev?`#${c.prev.rank}`:"-"}</td><td style={TD}>#{c.rank}</td>
            <td style={{...TD,fontWeight:700,fontSize:14,color:c.dRk>0?"#2ecc71":c.dRk<0?"#e74c3c":"#666"}}>{c.dRk!==null?c.dRk>0?`↑${c.dRk}`:c.dRk<0?`↓${Math.abs(c.dRk)}`:"=":"-"}</td></tr>})}</tbody></table></div></div>

      {/* Mejoraron / Bajaron */}
      <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <div style={{...crd,borderLeft:"4px solid #2ecc71"}}><h3 style={{margin:"0 0 8px",fontSize:14,fontWeight:700,color:"#2ecc71"}}>📈 Mas Mejoraron</h3>
          {comp.filter(c=>c.dFH>0).slice(0,5).map(c=><div key={c.nombre} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",fontSize:12,borderBottom:"1px solid #f0f2f1"}}><span style={{fontWeight:600}}>{sN(c.nombre)}{rBadge(c)}</span><span style={{fontWeight:800,color:"#2ecc71"}}>+{fN(c.dFH)} f/h | +{(c.dReg||0).toLocaleString()} reg</span></div>)}
          {comp.filter(c=>c.dFH>0).length===0&&<p style={{fontSize:12,color:CL.txtL}}>Ninguno mejoro</p>}</div>
        <div style={{...crd,borderLeft:"4px solid #e74c3c"}}><h3 style={{margin:"0 0 8px",fontSize:14,fontWeight:700,color:"#e74c3c"}}>📉 Mas Bajaron</h3>
          {comp.filter(c=>c.dFH!==null&&c.dFH<0).sort((a,b)=>a.dFH-b.dFH).slice(0,5).map(c=><div key={c.nombre} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",fontSize:12,borderBottom:"1px solid #f0f2f1"}}><span style={{fontWeight:600}}>{sN(c.nombre)}{rBadge(c)}</span><span style={{fontWeight:800,color:"#e74c3c"}}>{fN(c.dFH)} f/h | {(c.dReg||0).toLocaleString()} reg</span></div>)}
          {comp.filter(c=>c.dFH!==null&&c.dFH<0).length===0&&<p style={{fontSize:12,color:CL.txtL}}>Ninguno bajo</p>}</div></div>

      {/* Nuevos y salieron */}
      {comp.filter(c=>!c.prev).length>0&&<div style={{...crd,borderLeft:"4px solid #f39c12"}}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700,color:CL.warn}}>🆕 Nuevos (no estaban antes)</h3>
        {comp.filter(c=>!c.prev).map(c=><span key={c.nombre} style={{display:"inline-block",padding:"3px 10px",margin:2,borderRadius:12,background:"#fff9e6",fontSize:11,fontWeight:600}}>{c.nombre} ({fN(c.pfH)} f/h)</span>)}</div>}
      {d2.cajeros.filter(c2=>!data.cajeros.find(c=>c.nombre===c2.nombre)).length>0&&<div style={{...crd,borderLeft:"4px solid #95a5a6"}}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700,color:"#95a5a6"}}>👋 Salieron (estaban antes, ya no)</h3>
        {d2.cajeros.filter(c2=>!data.cajeros.find(c=>c.nombre===c2.nombre)).map(c2=><span key={c2.nombre} style={{display:"inline-block",padding:"3px 10px",margin:2,borderRadius:12,background:"#f0f0f0",fontSize:11,fontWeight:600}}>{c2.nombre} ({fN(c2.pfH)} f/h)</span>)}</div>}
    </>}

    {/* SEDE COMPARISON */}
    {d2&&mode==="sede"&&<SedeComp data={data} d2={d2}/>}

    {!d2&&<div style={{...crd,textAlign:"center",padding:40,color:CL.txtL,background:"#f8faf8"}}>
      <div style={{fontSize:40,marginBottom:8}}>💾</div>
      <p style={{fontSize:14,fontWeight:600,marginBottom:4}}>Como funciona la memoria</p>
      <p style={{fontSize:12,maxWidth:400,margin:"0 auto"}}>1. Haz tu analisis con los datos actuales. 2. Click en <b>💾 Guardar Memoria</b> para descargar un .json. 3. El proximo mes, sube nuevos datos y luego carga el .json aqui. 4. La app detecta automaticamente si es la misma sede (compara periodos) o diferente (compara sedes).</p></div>}
  </div>;}

function SedeComp({data,d2}){
  const s1={nm:data.sede,n:data.cajeros.length,avg:data.avg,avgR:data.avgR,tF:data.tF,tR:data.tR,per:data.periodo,cu:data.cajeros.filter(c=>c.kpi.lab==="Cumple").length,nc:data.cajeros.filter(c=>c.kpi.lab==="No cumple").length,top:data.cajeros[0],consist:data.cajeros.reduce((s,c)=>s+c.consist.cv,0)/data.cajeros.length};
  const s2={nm:d2.sede,n:d2.cajeros.length,avg:d2.avg,avgR:d2.avgR,tF:d2.tF,tR:d2.tR,per:d2.periodo,cu:d2.cajeros.filter(c=>c.kpi.lab==="Cumple").length,nc:d2.cajeros.filter(c=>c.kpi.lab==="No cumple").length,top:d2.cajeros[0],consist:d2.cajeros.reduce((s,c)=>s+c.consist.cv,0)/d2.cajeros.length};
  const mt=[{l:"Cajeros",a:s1.n,b:s2.n,w:0},{l:"Prom F/H",a:fN(s1.avg),b:fN(s2.avg),w:s1.avg>s2.avg?1:2,key:1},{l:"Prom R/H",a:fN(s1.avgR),b:fN(s2.avgR),w:s1.avgR>s2.avgR?1:2},{l:"Total Fact",a:s1.tF.toLocaleString(),b:s2.tF.toLocaleString(),w:0},{l:"Cumplen",a:`${s1.cu} (${((s1.cu/s1.n)*100).toFixed(0)}%)`,b:`${s2.cu} (${((s2.cu/s2.n)*100).toFixed(0)}%)`,w:s1.cu/s1.n>s2.cu/s2.n?1:2},{l:"No Cumplen",a:`${s1.nc} (${((s1.nc/s1.n)*100).toFixed(0)}%)`,b:`${s2.nc} (${((s2.nc/s2.n)*100).toFixed(0)}%)`,w:s1.nc/s1.n<s2.nc/s2.n?1:2},{l:"Consist Prom",a:`CV ${s1.consist.toFixed(0)}%`,b:`CV ${s2.consist.toFixed(0)}%`,w:s1.consist<s2.consist?1:2},{l:"Mejor Cajero",a:s1.top?`${sN(s1.top.nombre)} (${fN(s1.top.pfH)})`:"",b:s2.top?`${sN(s2.top.nombre)} (${fN(s2.top.pfH)})`:"",w:0},{l:"Periodo",a:`${s1.per.desde} a ${s1.per.hasta}`,b:`${s2.per.desde} a ${s2.per.hasta}`,w:0}];
  return <><div style={crd}><h3 style={{margin:"0 0 10px",fontSize:15,fontWeight:700,color:"#9b59b6"}}>🏢 Comparacion entre Sedes</h3>
    <table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><th style={{...TH,textAlign:"right",width:"35%",color:"#9b59b6"}}>{s1.nm}</th><th style={{...TH,textAlign:"center",width:"30%"}}>Metrica</th><th style={{...TH,width:"35%",color:CL.pri}}>{s2.nm}</th></tr></thead>
      <tbody>{mt.map(m=><tr key={m.l} style={{background:m.key?CL.priLt:"transparent"}}><td style={{...TD,textAlign:"right",fontWeight:m.key?800:600,color:m.w===1?"#2ecc71":m.w===2?"#e74c3c":CL.txt,fontSize:m.key?16:13}}>{m.w===1?"✓ ":""}{m.a}</td><td style={{...TD,textAlign:"center",fontSize:11,color:CL.txtL,fontWeight:700}}>{m.l}</td><td style={{...TD,fontWeight:m.key?800:600,color:m.w===2?"#2ecc71":m.w===1?"#e74c3c":CL.txt,fontSize:m.key?16:13}}>{m.b}{m.w===2?" ✓":""}</td></tr>)}</tbody></table></div>
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={{...crd,borderTop:"3px solid #9b59b6"}}><h4 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>🏆 Top 5 - {s1.nm}</h4>
        {data.cajeros.slice(0,5).map((c,i)=><div key={c.nombre} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",fontSize:12,borderBottom:"1px solid #f0f2f1"}}><span style={{fontWeight:600}}>{i+1}. {sN(c.nombre)}</span><span style={{fontWeight:800,color:CL.pri}}>{fN(c.pfH)}</span></div>)}</div>
      <div style={{...crd,borderTop:`3px solid ${CL.pri}`}}><h4 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>🏆 Top 5 - {s2.nm}</h4>
        {d2.cajeros.slice(0,5).map((c,i)=><div key={c.nombre} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",fontSize:12,borderBottom:"1px solid #f0f2f1"}}><span style={{fontWeight:600}}>{i+1}. {sN(c.nombre)}</span><span style={{fontWeight:800,color:CL.pri}}>{fN(c.pfH)}</span></div>)}</div></div></>;}


/* ═══ SIMULATOR ═══ */
function Sim({data}){const[excl,setExcl]=useState([]);
  const sim=useMemo(()=>{const active=data.cajeros.filter(c=>!excl.includes(c.nombre));if(active.length===0)return null;
    const avg=active.reduce((s,c)=>s+c.pfH,0)/active.length;const avgR=active.reduce((s,c)=>s+c.prH,0)/active.length;
    const tF=active.reduce((s,c)=>s+c.tF,0),tR=active.reduce((s,c)=>s+c.tR,0);
    const cu=active.filter(c=>c.pfH>=avg+2).length,ep=active.filter(c=>c.pfH>=avg&&c.pfH<avg+2).length,nc=active.filter(c=>c.pfH<avg).length;
    return{n:active.length,avg,avgR,tF,tR,cu,ep,nc};},[data,excl]);
  const real={n:data.cajeros.length,avg:data.avg,avgR:data.avgR,tF:data.tF,tR:data.tR,cu:data.cajeros.filter(c=>c.kpi.lab==="Cumple").length,nc:data.cajeros.filter(c=>c.kpi.lab==="No cumple").length};
  const diff=sim?sim.avg-real.avg:0;
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={crd}><h3 style={{margin:"0 0 8px",fontSize:15,fontWeight:700}}>🧮 Simulador: ¿Que pasa si...?</h3>
      <p style={{fontSize:12,color:CL.txtL,marginBottom:10}}>Excluye cajeros para ver como cambiaria el promedio de la sede. Util para simular rotaciones o ausencias.</p>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {data.cajeros.map(c=>{const ex=excl.includes(c.nombre);return <button key={c.nombre} onClick={()=>setExcl(ex?excl.filter(x=>x!==c.nombre):[...excl,c.nombre])}
          style={{padding:"5px 10px",borderRadius:8,border:`1.5px solid ${ex?"#e74c3c":CL.pri+"60"}`,background:ex?"#fde8e8":"#fff",color:ex?"#e74c3c":CL.txt,fontSize:11,fontWeight:600,cursor:"pointer",opacity:ex?.6:1}}>
          {ex?"✕ ":""}{sN(c.nombre)}{rBadge(c)} ({fN(c.pfH)})</button>})}</div>
      {excl.length>0&&<p style={{fontSize:11,color:CL.dan,marginTop:6,fontWeight:600}}>Excluidos: {excl.length} cajero(s)</p>}</div>
    {sim&&<div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr",gap:14,alignItems:"start"}}>
      <div style={{...crd,borderTop:`3px solid ${CL.txtL}`}}><h4 style={{margin:"0 0 8px",fontSize:13,fontWeight:700,color:CL.txtL}}>📊 Real</h4>
        <div style={{fontSize:11}}><p>Cajeros: <b>{real.n}</b></p><p>Prom F/H: <b>{fN(real.avg)}</b></p><p>Prom R/H: <b>{fN(real.avgR)}</b></p><p>Facturas: <b>{real.tF.toLocaleString()}</b></p></div></div>
      <div style={{textAlign:"center",paddingTop:30}}><div style={{fontSize:28,fontWeight:900,color:diff>0?"#2ecc71":diff<0?"#e74c3c":"#666"}}>{diff>0?"+":""}{fN(diff)}</div><div style={{fontSize:11,color:CL.txtL}}>f/h</div></div>
      <div style={{...crd,borderTop:`3px solid ${diff>0?"#2ecc71":"#e74c3c"}`}}><h4 style={{margin:"0 0 8px",fontSize:13,fontWeight:700,color:diff>0?"#2ecc71":"#e74c3c"}}>🧮 Simulado</h4>
        <div style={{fontSize:11}}><p>Cajeros: <b>{sim.n}</b></p><p>Prom F/H: <b style={{color:diff>0?"#2ecc71":"#e74c3c"}}>{fN(sim.avg)}</b></p><p>Prom R/H: <b>{fN(sim.avgR)}</b></p><p>Facturas: <b>{sim.tF.toLocaleString()}</b></p></div></div></div>}
    {excl.length>0&&<div style={{...crd,background:"#f0f6ff",borderLeft:"3px solid #3498db"}}><p style={{fontSize:12,margin:0}}>💡 Sin {excl.map(x=>sN(x)).join(", ")} el promedio {diff>0?"sube":"baja"} <b>{Math.abs(diff).toFixed(2)} f/h</b> ({diff>0?"mejor":"peor"}). {diff>0?`Se perderian ${(real.tF-sim.tF).toLocaleString()} facturas pero la eficiencia general mejora.`:`La eficiencia general empeora.`}</p></div>}
    {excl.length===0&&<div style={{...crd,textAlign:"center",padding:40,color:CL.txtL}}><div style={{fontSize:40}}>🧮</div><p>Haz click en los cajeros arriba para excluirlos de la simulacion</p></div>}</div>;}

/* ═══ SCORECARD ═══ */
function Score({data,showR}){const kt=kpiT(data);const cu=data.cajeros.filter(c=>c.kpi.lab==="Cumple"),nc=data.cajeros.filter(c=>c.kpi.lab==="No cumple"),ep=data.cajeros.filter(c=>c.kpi.lab==="En promedio");
  const t3=data.cajeros.slice(0,3),b3=data.cajeros.slice(-3).reverse();
  const avgConsist=data.cajeros.length>0?data.cajeros.reduce((s,c)=>s+c.consist.cv,0)/data.cajeros.length:0;
  const health=nc.length===0?"🟢 Excelente":nc.length<=2?"🟡 Bueno":nc.length<=Math.ceil(data.cajeros.length/3)?"🟠 Regular":"🔴 Critico";
  const sedeTH=data.cajeros.reduce((s,c)=>s+c.hrsEfec,0),sedeTD=data.cajeros.reduce((s,c)=>s+c.numDias,0);
  // Day analysis
  const bestDay=data.dS.length>0?data.dS.reduce((a,b)=>a.facs>b.facs?a:b):null;
  const worstDay=data.dS.length>0?data.dS.reduce((a,b)=>a.facs<b.facs?a:b):null;
  // Day of week aggregation
  const dowMap={};data.dS.forEach(d=>{const dt=new Date(d.fecha+"T12:00:00");const dw=dt.getDay();if(!dowMap[dw])dowMap[dw]={regs:0,facs:0,dias:0};dowMap[dw].regs+=d.regs;dowMap[dw].facs+=d.facs;dowMap[dw].dias++;});
  const dowArr=Object.entries(dowMap).map(([d,v])=>({dia:DC[+d],diaL:DIAS[+d],regs:v.regs,facs:v.facs,dias:v.dias,promFac:Math.round(v.facs/v.dias),promReg:Math.round(v.regs/v.dias)})).sort((a,b)=>b.promReg-a.promReg);
  const bestDow=dowArr[0],worstDow=dowArr[dowArr.length-1];
  // Hour analysis
  const hMap={};data.cajeros.forEach(c=>c.hA.forEach(h=>{if(!hMap[h.hora])hMap[h.hora]=0;hMap[h.hora]+=h.regs;}));
  const hArr=Object.entries(hMap).map(([h,r])=>({hora:+h,regs:r})).sort((a,b)=>b.regs-a.regs);
  const peakHrs=hArr.slice(0,3),deadHrs=hArr.slice(-3).reverse();
  // Most/least registros
  const mostReg=data.cajeros.slice().sort((a,b)=>b.tR-a.tR)[0];
  const leastReg=data.cajeros.slice().sort((a,b)=>a.tR-b.tR)[0];
  const mostFac=data.cajeros.slice().sort((a,b)=>b.tF-a.tF)[0];
  const leastFac=data.cajeros.slice().sort((a,b)=>a.tF-b.tF)[0];
  // TPV stats
  const tpv=data.tpvStats||[];
  const tpvBest=tpv[0],tpvWorst=tpv.length>1?tpv[tpv.length-1]:null;

  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{...crd,background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,color:"#fff",padding:"16px 20px"}}>
      <h2 style={{margin:"0 0 4px",fontSize:20,fontWeight:800}}>🏅 Scorecard Profundo — {esc(data.sede)}</h2>
      <p style={{margin:0,fontSize:12,opacity:.8}}>{data.periodo.desde} a {data.periodo.hasta} | {data.cajeros.length} cajeros | {data.dS.length} dias</p>
      <div style={{marginTop:10,padding:"8px 14px",borderRadius:8,background:"rgba(255,255,255,.15)",fontSize:14,fontWeight:700,textAlign:"center"}}>{health}</div></div>

    {/* KPIs generales */}
    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><KC label="Facturas" value={data.tF.toLocaleString()} color="#3498db"/><KC label="Registros" value={data.tR.toLocaleString()} color="#9b59b6"/>
      <KC label="Prom F/H" value={fN(data.avg)} color="#e67e22"/><KC label="Total Hrs" value={Math.round(sedeTH)+"h"} color="#2ecc71"/>
      <KC label="Prom Hrs/Dia" value={sedeTD>0?(sedeTH/sedeTD).toFixed(1)+"h":"—"} color="#8e44ad"/>
      <KC label="Consistencia" value={`CV ${avgConsist.toFixed(0)}%`} color={avgConsist<30?"#2ecc71":avgConsist<50?"#f39c12":"#e74c3c"}/></div>

    {/* KPI Distribution + Top/Bottom */}
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📊 Distribucion KPI</h3>
        <div style={{display:"flex",gap:8,marginBottom:8}}><div style={{flex:cu.length,background:"#2ecc71",borderRadius:6,padding:6,textAlign:"center",color:"#fff",fontSize:11,fontWeight:700,minWidth:30}}>{cu.length}</div>
          <div style={{flex:ep.length||.5,background:"#f39c12",borderRadius:6,padding:6,textAlign:"center",color:"#fff",fontSize:11,fontWeight:700,minWidth:30}}>{ep.length}</div>
          <div style={{flex:nc.length||.5,background:"#e74c3c",borderRadius:6,padding:6,textAlign:"center",color:"#fff",fontSize:11,fontWeight:700,minWidth:30}}>{nc.length}</div></div>
        <div style={{fontSize:11,color:CL.txtL}}>✅ Cumplen: <b>{cu.length}</b> ({data.cajeros.length>0?((cu.length/data.cajeros.length)*100).toFixed(0):0}%) | 😐 En promedio: <b>{ep.length}</b> ({data.cajeros.length>0?((ep.length/data.cajeros.length)*100).toFixed(0):0}%) | ❌ No cumplen: <b>{nc.length}</b> ({data.cajeros.length>0?((nc.length/data.cajeros.length)*100).toFixed(0):0}%)</div></div>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>🏆 Top 3 vs ⚠️ Bottom 3</h3>
        <div style={{display:"flex",gap:10}}>
          <div style={{flex:1}}>{t3.map((c,i)=><div key={c.nombre} style={{fontSize:11,padding:"3px 0",color:CL.pri}}><b>{i+1}.</b> {sN(c.nombre)}{rBadge(c)} — <b>{fN(c.pfH)}</b></div>)}</div>
          <div style={{flex:1}}>{b3.map(c=><div key={c.nombre} style={{fontSize:11,padding:"3px 0",color:CL.dan}}>#{c.rank} {sN(c.nombre)}{rBadge(c)} — <b>{fN(c.pfH)}</b></div>)}</div></div></div></div>

    {/* Quien registra mas/menos */}
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={{...crd,borderLeft:"4px solid #2ecc71"}}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>📦 Mas registros</h3>
        {mostReg&&<div style={{fontSize:12}}><b>{sN(mostReg.nombre)}{rBadge(mostReg)}</b> — {mostReg.tR.toLocaleString()} reg ({mostReg.tF.toLocaleString()} fact)</div>}
        {mostFac&&mostFac.nombre!==mostReg?.nombre&&<div style={{fontSize:11,color:CL.txtL,marginTop:4}}>Mas facturas: <b>{sN(mostFac.nombre)}</b> — {mostFac.tF.toLocaleString()}</div>}</div>
      <div style={{...crd,borderLeft:"4px solid #e74c3c"}}><h3 style={{margin:"0 0 8px",fontSize:13,fontWeight:700}}>📉 Menos registros</h3>
        {leastReg&&<div style={{fontSize:12}}><b>{sN(leastReg.nombre)}{rBadge(leastReg)}</b> — {leastReg.tR.toLocaleString()} reg ({leastReg.tF.toLocaleString()} fact)</div>}
        {leastFac&&leastFac.nombre!==leastReg?.nombre&&<div style={{fontSize:11,color:CL.txtL,marginTop:4}}>Menos facturas: <b>{sN(leastFac.nombre)}</b> — {leastFac.tF.toLocaleString()}</div>}</div></div>

    {/* Dias y Horas */}
    <div className="desk-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>📅 Analisis por Dia</h3>
        {bestDay&&<div style={{fontSize:12,marginBottom:6}}>📈 Mejor dia: <b>{bestDay.fechaD}</b> — {bestDay.facs.toLocaleString()} fact, {bestDay.regs.toLocaleString()} reg</div>}
        {worstDay&&<div style={{fontSize:12,marginBottom:10}}>📉 Peor dia: <b>{worstDay.fechaD}</b> — {worstDay.facs.toLocaleString()} fact, {worstDay.regs.toLocaleString()} reg</div>}
        <div style={{fontSize:11,fontWeight:700,color:CL.txtL,marginBottom:4}}>Promedio por dia de semana:</div>
        {dowArr.map(d=><div key={d.dia} style={{display:"flex",justifyContent:"space-between",padding:"2px 0",fontSize:11,borderBottom:"1px solid #f0f2f1"}}>
          <span style={{fontWeight:600,color:d===bestDow?"#2ecc71":d===worstDow?"#e74c3c":CL.txt}}>{d===bestDow?"🔥 ":d===worstDow?"💤 ":""}{d.diaL}</span>
          <span>{d.promFac} fact | {d.promReg} reg ({d.dias} dias)</span></div>)}</div>
      <div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>⏰ Analisis por Hora</h3>
        {peakHrs.length>0&&<div style={{fontSize:12,marginBottom:6}}>🔥 Horas pico: {peakHrs.map(h=><span key={h.hora} style={{display:"inline-block",padding:"2px 8px",margin:"1px",borderRadius:6,background:"#e6f9ee",fontWeight:700,fontSize:11}}>{h.hora}:00 ({h.regs.toLocaleString()})</span>)}</div>}
        {deadHrs.length>0&&<div style={{fontSize:12,marginBottom:10}}>💤 Horas muertas: {deadHrs.map(h=><span key={h.hora} style={{display:"inline-block",padding:"2px 8px",margin:"1px",borderRadius:6,background:"#fde8e8",fontWeight:700,fontSize:11}}>{h.hora}:00 ({h.regs.toLocaleString()})</span>)}</div>}
        <div style={{fontSize:11,fontWeight:700,color:CL.txtL,marginBottom:4}}>Desglose por hora:</div>
        <div style={{maxHeight:200,overflowY:"auto"}}>{hArr.map(h=><div key={h.hora} style={{display:"flex",justifyContent:"space-between",padding:"2px 0",fontSize:11,borderBottom:"1px solid #f0f2f1"}}>
          <span style={{fontWeight:600}}>{h.hora}:00</span>
          <div style={{flex:1,margin:"0 8px"}}><div style={{height:6,borderRadius:3,background:`${CL.pri}20`}}><div style={{height:6,borderRadius:3,background:CL.pri,width:`${hArr[0]?.regs>0?(h.regs/hArr[0].regs*100):0}%`}}/></div></div>
          <span style={{minWidth:50,textAlign:"right"}}>{h.regs.toLocaleString()}</span></div>)}</div></div></div>

    {/* TPV Analysis */}
    {tpv.length>0&&<div style={crd}><h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700}}>🖥️ Analisis por Caja (T.P.V.)</h3>
      <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
        <KC label="Cajas activas" value={tpv.length} color="#3498db"/>
        {tpvBest&&<KC label="Mas registros" value={`${tpvBest.tpv} (${tpvBest.regs.toLocaleString()})`} color="#2ecc71"/>}
        {tpvWorst&&<KC label="Menos registros" value={`${tpvWorst.tpv} (${tpvWorst.regs.toLocaleString()})`} color="#e74c3c"/>}</div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Caja","Reg","Fact","Cajeros","Dias","Reg/Dia","Fact/Dia","Hora Pico"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>
        {tpv.map((t,i)=><tr key={t.tpv} style={{background:i===0?"#e6f9ee":i===tpv.length-1?"#fde8e8":"transparent"}}><td style={{...TD,fontWeight:700,fontSize:11}}>{t.tpv}</td><td style={TD}>{t.regs.toLocaleString()}</td><td style={TD}>{t.facs.toLocaleString()}</td><td style={TD}>{t.cajeros}</td><td style={TD}>{t.dias}</td><td style={{...TD,fontWeight:700,color:CL.pri}}>{t.promRegDia}</td><td style={TD}>{t.promFacDia}</td><td style={TD}>{t.peakHora!==null?t.peakHora+":00":"—"}</td></tr>)}</tbody></table></div></div>}

    {/* Download */}
    <div style={{...crd,textAlign:"center",padding:12}}>
      <Btn green onClick={()=>{toast("⏳ Generando...");setTimeout(()=>{const h=buildScoreHTML(data,kt,cu,nc,ep,t3,b3,health,avgConsist,tpv,dowArr,hArr,bestDay,worstDay,sedeTH,sedeTD);dlReport(h,`Scorecard_${data.sede}`);},50);}}>📥 Descargar Scorecard PDF</Btn></div></div>;}

function buildScoreHTML(data,kt,cu,nc,ep,t3,b3,health,avgConsist,tpv,dowArr,hArr,bestDay,worstDay,sedeTH,sedeTD){
  return`<div style="font-family:'Segoe UI',sans-serif;max-width:800px;margin:0 auto;padding:30px">
<h1 style="text-align:center;color:#0d4a28">📊 Scorecard Profundo — ${esc(data.sede)}</h1>
<p style="text-align:center;color:#666">${data.periodo.desde} a ${data.periodo.hasta} | ${data.cajeros.length} cajeros | ${data.dS.length} dias</p>
<div style="text-align:center;padding:12px;background:${nc.length===0?"#e6f9ee":nc.length<=2?"#fff9e6":"#fde8e8"};border-radius:10px;font-size:18px;font-weight:700;margin:20px 0">${health}</div>
<table style="width:100%;border-collapse:collapse;margin:14px 0;font-size:13px"><tr><td style="padding:8px;border:1px solid #eee"><b>Facturas</b></td><td style="padding:8px;border:1px solid #eee">${data.tF.toLocaleString()}</td><td style="padding:8px;border:1px solid #eee"><b>Registros</b></td><td style="padding:8px;border:1px solid #eee">${data.tR.toLocaleString()}</td></tr>
<tr><td style="padding:8px;border:1px solid #eee"><b>Prom F/H</b></td><td style="padding:8px;border:1px solid #eee;font-weight:700;color:#1a6b3c">${fN(data.avg)}</td><td style="padding:8px;border:1px solid #eee"><b>Total Hrs</b></td><td style="padding:8px;border:1px solid #eee">${Math.round(sedeTH)}h</td></tr>
<tr><td style="padding:8px;border:1px solid #eee"><b>Cumplen</b></td><td style="padding:8px;border:1px solid #eee;color:#2ecc71">${cu.length} (${data.cajeros.length>0?((cu.length/data.cajeros.length)*100).toFixed(0):0}%)</td><td style="padding:8px;border:1px solid #eee"><b>No cumplen</b></td><td style="padding:8px;border:1px solid #eee;color:#e74c3c">${nc.length}</td></tr></table>
<h2>🏆 Top 3 / ⚠️ Bottom 3</h2>
<div style="display:flex;gap:14px"><div style="flex:1;background:#e6f9ee;padding:12px;border-radius:10px">${t3.map((c,i)=>`<div style="font-size:12px;padding:2px 0"><b>${i+1}.</b> ${esc(sN(c.nombre))} — <b>${fN(c.pfH)}</b> f/h</div>`).join("")}</div>
<div style="flex:1;background:#fde8e8;padding:12px;border-radius:10px">${b3.map(c=>`<div style="font-size:12px;padding:2px 0">#${c.rank} ${esc(sN(c.nombre))} — <b>${fN(c.pfH)}</b> f/h</div>`).join("")}</div></div>
${bestDay?`<h2>📅 Dias</h2><p>📈 Mejor: <b>${bestDay.fechaD}</b> (${bestDay.facs} fact) | 📉 Peor: <b>${worstDay.fechaD}</b> (${worstDay.facs} fact)</p>`:""}
${dowArr.length>0?`<table style="width:100%;border-collapse:collapse;font-size:12px;margin:10px 0"><thead><tr><th style="background:#1a6b3c;color:#fff;padding:6px">Dia</th><th style="background:#1a6b3c;color:#fff;padding:6px">Prom Fact</th><th style="background:#1a6b3c;color:#fff;padding:6px">Prom Reg</th></tr></thead><tbody>${dowArr.map(d=>`<tr><td style="padding:5px;border:1px solid #eee">${d.diaL}</td><td style="padding:5px;border:1px solid #eee">${d.promFac}</td><td style="padding:5px;border:1px solid #eee">${d.promReg}</td></tr>`).join("")}</tbody></table>`:""}
${hArr.length>0?`<h2>⏰ Horas</h2><p>🔥 Pico: ${hArr.slice(0,3).map(h=>`${h.hora}:00`).join(", ")} | 💤 Muertas: ${hArr.slice(-3).map(h=>`${h.hora}:00`).join(", ")}</p>`:""}
${tpv.length>0?`<h2>🖥️ Cajas (T.P.V.)</h2><table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr><th style="background:#1a6b3c;color:#fff;padding:6px">Caja</th><th style="background:#1a6b3c;color:#fff;padding:6px">Reg</th><th style="background:#1a6b3c;color:#fff;padding:6px">Fact</th><th style="background:#1a6b3c;color:#fff;padding:6px">Reg/Dia</th></tr></thead><tbody>${tpv.map(t=>`<tr><td style="padding:5px;border:1px solid #eee">${t.tpv}</td><td style="padding:5px;border:1px solid #eee">${t.regs.toLocaleString()}</td><td style="padding:5px;border:1px solid #eee">${t.facs.toLocaleString()}</td><td style="padding:5px;border:1px solid #eee;font-weight:700">${t.promRegDia}</td></tr>`).join("")}</tbody></table>`:""}
<div style="text-align:center;font-size:10px;color:#999;margin-top:20px;border-top:1px solid #eee;padding-top:10px">${new Date().toLocaleDateString("es-CO")} | Generado automaticamente</div></div>`;}

/* ═══ METAS ═══ */
function Metas({data}){const[meta,setMeta]=useState("");const metaV=parseFloat(meta);
  const valid=!isNaN(metaV)&&metaV>0;
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{...crd,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}><h3 style={{margin:0,fontSize:15,fontWeight:700}}>🎯 Metas</h3>
      <div style={{display:"flex",gap:8,alignItems:"center"}}><span style={{fontSize:12,fontWeight:600}}>Meta F/H:</span>
        <input type="number" step="0.5" placeholder={fN(data.avg+KPI_CUMPLE_OFFSET)} value={meta} onChange={e=>setMeta(e.target.value)} style={{padding:"8px 12px",borderRadius:8,border:"2px solid #2ecc71",fontSize:15,fontWeight:700,width:100,textAlign:"center",outline:"none"}}/></div>
      <p style={{fontSize:11,color:CL.txtL,margin:0}}>Define una meta y ve quien la cumple y cuanto les falta</p></div>
    {valid&&<div style={crd}><div style={{display:"flex",gap:14,marginBottom:12,flexWrap:"wrap"}}>
      <KC label="Meta" value={fN(metaV)} color="#2ecc71"/>
      <KC label="Cumplen meta" value={data.cajeros.filter(c=>c.pfH>=metaV).length+"/"+data.cajeros.length} color={data.cajeros.filter(c=>c.pfH>=metaV).length>=data.cajeros.length/2?"#2ecc71":"#e74c3c"}/>
      <KC label="Prom actual" value={fN(data.avg)} color="#e67e22"/></div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["#","Cajero","F/H","Meta","Dif","%","Estado"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>
        {data.cajeros.map(c=>{const d=c.pfH-metaV,pct=(c.pfH/metaV*100);const ok=d>=0;
          return <tr key={c.nombre} style={{background:ok?"#e6f9ee":"#fde8e8"}}><td style={{...TD,fontWeight:800}}>{c.rank}</td><td style={{...TD,fontWeight:600,fontSize:11}}>{c.nombre}{rBadge(c)}</td>
            <td style={{...TD,fontWeight:700}}>{fN(c.pfH)}</td><td style={TD}>{fN(metaV)}</td>
            <td style={{...TD,fontWeight:800,color:ok?"#2ecc71":"#e74c3c"}}>{ok?"+":""}{fN(d)}</td>
            <td style={TD}>{pct.toFixed(0)}%</td>
            <td style={TD}><Bg color={ok?"#2ecc71":pct>=80?"#f39c12":"#e74c3c"}>{ok?"✅ Cumple":pct>=80?"🔶 Cerca":"❌ Lejos"}</Bg></td></tr>})}</tbody></table></div></div>}
    {!valid&&<div style={{...crd,textAlign:"center",padding:40,color:CL.txtL}}><div style={{fontSize:40}}>🎯</div><p>Ingresa una meta de Fact/Hora arriba para ver el progreso de cada cajero</p></div>}</div>;}


/* ═══ ERROR BOUNDARY ═══ */
class ErrBound extends Component{
  constructor(p){super(p);this.state={hasErr:false,err:null};}
  static getDerivedStateFromError(e){return{hasErr:true,err:e};}
  componentDidCatch(e,info){console.error("AuditApp error:",e,info);}
  render(){if(this.state.hasErr)return <div style={{padding:40,textAlign:"center",background:"#fde8e8",borderRadius:16,margin:16}}>
    <div style={{fontSize:48,marginBottom:12}}>⚠️</div>
    <h3 style={{fontSize:16,fontWeight:700,color:"#e74c3c",marginBottom:8}}>Algo salio mal</h3>
    <p style={{fontSize:13,color:"#666",marginBottom:16}}>{this.state.err?.message||"Error desconocido"}</p>
    <button onClick={()=>this.setState({hasErr:false,err:null})} style={{padding:"10px 24px",borderRadius:10,border:"none",background:"#1a6b3c",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>🔄 Reintentar</button></div>;
  return this.props.children;}
}

/* ═══ APP ═══ */
export default function App(){
  const[raw,setRaw]=useState(null);const[view,setView]=useState("da");const[sel,setSel]=useState("");const[report,setReport]=useState(null);
  const[showKPI,setShowKPI]=useState(false);const[kC,setKC]=useState("");const[kE,setKE]=useState("");
  const[dateFrom,setDateFrom]=useState("");const[dateTo,setDateTo]=useState("");
  const[sedeFilter,setSedeFilter]=useState("");
  const[showRetirados,setShowRetirados]=useState(true);
  const[ccostoFilter,setCcostoFilter]=useState([]);
  const[excSuperv,setExcSuperv]=useState(true);

  const sedes=useMemo(()=>{if(!raw)return[];const s=[...new Set(raw.cajeros.map(c=>c.sede).filter(Boolean))].sort();return s;},[raw]);
  const hasBD=raw?raw.cajeros.some(c=>c.activo!==null):false;
  const ccostos=useMemo(()=>{if(!raw||!hasBD)return[];return[...new Set(raw.cajeros.map(c=>c.ccosto).filter(Boolean))].sort();},[raw,hasBD]);
  const sedeData=useMemo(()=>{if(!raw)return null;
    let cajeros=raw.cajeros;
    if(hasBD&&!showRetirados)cajeros=cajeros.filter(c=>c.activo!==false);
    if(hasBD&&excSuperv)cajeros=cajeros.filter(c=>!c.cargo||!c.cargo.toUpperCase().includes("SUPERVISOR"));
    if(ccostoFilter.length>0)cajeros=cajeros.filter(c=>ccostoFilter.includes(c.ccosto));
    if(sedeFilter&&sedes.length>1){cajeros=cajeros.filter(c=>c.sede===sedeFilter);}
    if(cajeros.length===0)return{...raw,cajeros:[],avg:0,avgR:0,tR:0,tF:0,dS:[],allDates:[],hasBD};
    if(cajeros===raw.cajeros&&!sedeFilter)return{...raw,hasBD};
    const avg=cajeros.reduce((s,c)=>s+c.pfH,0)/cajeros.length;const avgR=cajeros.reduce((s,c)=>s+c.prH,0)/cajeros.length;
    rankCajeros(cajeros,avg,raw.kpiCfg);
    const allDates=[...new Set(cajeros.flatMap(c=>c.dias.map(d=>d.fecha)))].sort();
    const dS=allDates.map(dk=>{let f=0,r=0,h=0,a=0;cajeros.forEach(c=>{const d=c.dias.find(x=>x.fecha===dk);if(d){f+=d.facs;r+=d.regs;h+=d.hrs;a++;}});return{fecha:dk,fechaD:fD(new Date(dk+"T12:00:00")),facs:f,regs:r,hrs:h,activos:a,fH:h>0?f/h:0};});
    const sede=sedeFilter||raw.sede;
    return{...raw,cajeros,avg,avgR,sede,tR:cajeros.reduce((s,c)=>s+c.tR,0),tF:cajeros.reduce((s,c)=>s+c.tF,0),periodo:{desde:allDates[0]||raw.periodo.desde,hasta:allDates[allDates.length-1]||raw.periodo.hasta},dS,allDates,hasSched:cajeros.some(c=>c.hrsHor!==null),hasBD};
  },[raw,sedeFilter,sedes,showRetirados,hasBD,ccostoFilter,excSuperv]);
  const data=useMemo(()=>sedeData?filterByDates(sedeData,dateFrom,dateTo):null,[sedeData,dateFrom,dateTo]);

  const applyKPI=()=>{const c=parseFloat(kC),e=parseFloat(kE);if(!isNaN(c)&&!isNaN(e)&&c>e){setRaw(recalcKPIs(raw,{active:true,cumple:c,enProm:e}));setShowKPI(false);}};
  const clearKPI=()=>{setKC("");setKE("");setRaw(recalcKPIs(raw,null));setShowKPI(false);};

  if(!data)return <Upload onData={setRaw}/>;

  if(report){return <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#fff",minHeight:"100vh"}}>
    <style>{`@media print{.np{display:none!important}.pa{padding:0!important}}`}</style>
    <div className="np" style={{position:"sticky",top:0,zIndex:100,background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,padding:"10px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
      <button onClick={()=>setReport(null)} style={{padding:"8px 16px",borderRadius:10,border:"1px solid rgba(255,255,255,.3)",background:"rgba(255,255,255,.15)",color:"#fff",fontSize:13,fontWeight:600,cursor:"pointer"}}>← Volver</button>
      <div style={{color:"#fff",fontSize:15,fontWeight:700}}>{report.title}</div>
      <button onClick={()=>dlReport(report.html,report.title)} style={{padding:"10px 24px",borderRadius:10,border:"none",background:"#fff",color:CL.pri,fontSize:14,fontWeight:700,cursor:"pointer",boxShadow:"0 2px 8px rgba(0,0,0,.15)"}}>📥 Descargar</button></div>
    <div className="pa" style={{maxWidth:820,margin:"0 auto",padding:"20px"}} dangerouslySetInnerHTML={{__html:`<style>${PCSS}</style>${report.html}`}}/></div>;}

  const hC=data.kpiCfg&&data.kpiCfg.active;
  const vs=[{id:"da",l:"📊"},{id:"in",l:"👤"},{id:"cm",l:"⚔️"},{id:"al",l:"🚨"},{id:"mt",l:"🎯"},{id:"si",l:"🧮"},{id:"sc",l:"🏅"},{id:"rp",l:"📄"},{id:"pc",l:"📈"}];
  const vLabels={da:"Dashboard",in:"Individual",cm:"Comparar",al:"Alertas",mt:"Metas",si:"Simulador",sc:"Scorecard",rp:"Informes",pc:"Periodos"};
  const inpSt={padding:"10px 12px",borderRadius:10,border:"2px solid #ddd",fontSize:16,fontWeight:700,width:120,textAlign:"center",outline:"none"};

  return <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:CL.bg,minHeight:"100vh",color:CL.txt}}>
    <style>{`@media(max-width:768px){.nav-lbl{display:none}.desk-grid{grid-template-columns:1fr!important}.aud-content *{max-width:100%!important;overflow-x:auto}.aud-content div[style]{min-width:0!important}}`}</style>
    {showKPI&&<div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={()=>setShowKPI(false)}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:20,padding:"28px 32px",maxWidth:460,width:"100%",boxShadow:"0 20px 60px rgba(0,0,0,.25)"}}>
        <h2 style={{fontSize:18,fontWeight:800,marginBottom:4}}>⚙️ Configurar KPIs</h2>
        <p style={{fontSize:13,color:CL.txtL,marginBottom:16}}>Define umbrales de Fact/Hora. Sin definir = calculo automatico.</p>
        {hC&&<div style={{background:CL.priLt,padding:"10px 14px",borderRadius:10,marginBottom:14,fontSize:13}}>✅ <b>Activos:</b> Cumple ≥ {data.kpiCfg.cumple} | En prom ≥ {data.kpiCfg.enProm}</div>}
        {!hC&&<div style={{background:"#f8f9fa",padding:"10px 14px",borderRadius:10,marginBottom:14,fontSize:13}}>📊 <b>Auto:</b> Cumple ≥ {fN(data.avg+KPI_CUMPLE_OFFSET)} | En prom ≥ {fN(data.avg)}</div>}
        <div style={{display:"flex",gap:20,marginBottom:20,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:140}}><label style={{display:"block",fontSize:12,fontWeight:700,color:CL.txtL,marginBottom:6}}>✅ Cumple (F/H ≥)</label><input style={{...inpSt,borderColor:"#2ecc71"}} type="number" step="0.1" placeholder={fN(data.avg+KPI_CUMPLE_OFFSET)} value={kC} onChange={e=>setKC(e.target.value)}/></div>
          <div style={{flex:1,minWidth:140}}><label style={{display:"block",fontSize:12,fontWeight:700,color:CL.txtL,marginBottom:6}}>😐 En prom (F/H ≥)</label><input style={{...inpSt,borderColor:"#f39c12"}} type="number" step="0.1" placeholder={fN(data.avg)} value={kE} onChange={e=>setKE(e.target.value)}/></div></div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
          {hC&&<button onClick={clearKPI} style={{padding:"10px 20px",borderRadius:10,border:`2px solid ${CL.dan}`,background:"#fff",color:CL.dan,fontSize:13,fontWeight:700,cursor:"pointer"}}>🗑️ Quitar</button>}
          <button onClick={()=>setShowKPI(false)} style={{padding:"10px 20px",borderRadius:10,border:"2px solid #ddd",background:"#fff",color:CL.txtL,fontSize:13,fontWeight:600,cursor:"pointer"}}>Cancelar</button>
          <button onClick={applyKPI} disabled={!kC||!kE||parseFloat(kC)<=parseFloat(kE)} style={{padding:"10px 24px",borderRadius:10,border:"none",background:(!kC||!kE||parseFloat(kC)<=parseFloat(kE))?"#ccc":CL.pri,color:"#fff",fontSize:13,fontWeight:700,cursor:(!kC||!kE||parseFloat(kC)<=parseFloat(kE))?"not-allowed":"pointer"}}>✅ Aplicar</button></div></div></div>}

    <nav style={{background:`linear-gradient(135deg,${CL.priDk},${CL.pri})`,position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 16px rgba(0,0,0,.15)"}}>
      <div style={{maxWidth:1280,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 18px",flexWrap:"wrap",gap:6}}>
        <div style={{color:"#fff",display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:20}}>📊</span><div><div style={{fontSize:16,fontWeight:700}}>Auditoria Cajeros</div><div style={{fontSize:10,opacity:.75}}>{data.sede}{sedes.length>1&&!sedeFilter?` (${sedes.length} sedes)`:""}{hC?" | KPIs":""}</div></div></div>
        <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>{vs.map(v=><button key={v.id} onClick={()=>setView(v.id)} title={vLabels[v.id]} style={{background:view===v.id?"rgba(255,255,255,.22)":"transparent",color:"#fff",border:`1px solid ${view===v.id?"rgba(255,255,255,.4)":"transparent"}`,padding:"6px 10px",borderRadius:7,cursor:"pointer",fontSize:11,fontWeight:view===v.id?600:400}}>{v.l}<span className="nav-lbl" style={{marginLeft:3}}>{vLabels[v.id]}</span></button>)}
          <button onClick={()=>{setShowKPI(true);if(hC){setKC(String(data.kpiCfg.cumple));setKE(String(data.kpiCfg.enProm));}}} style={{background:hC?"rgba(255,200,0,.3)":"transparent",color:"#fff",border:`1px solid ${hC?"rgba(255,200,0,.5)":"transparent"}`,padding:"6px 10px",borderRadius:7,cursor:"pointer",fontSize:11,fontWeight:hC?700:400}}>⚙️</button>
          <button onClick={()=>saveSnapshot(raw)} title="Guardar memoria" style={{background:"transparent",color:"#fff",border:"1px solid transparent",padding:"6px 10px",borderRadius:7,cursor:"pointer",fontSize:11}}>💾</button></div>
        <button onClick={()=>{if(window.confirm("¿Seguro? Se perdera el analisis actual.\nGuarda la memoria (💾) si quieres conservarlo.")){setRaw(null);setView("da");setDateFrom("");setDateTo("");setSedeFilter("");setCcostoFilter([]);}}} style={{padding:"4px 10px",borderRadius:7,border:"1px solid rgba(255,255,255,.25)",background:"rgba(255,255,255,.12)",color:"#fff",fontSize:11,cursor:"pointer"}}>🔄</button></div></nav>
    <div className="aud-content" style={{maxWidth:1280,margin:"0 auto",padding:"16px"}}>
      <FilterBar data={raw} from={dateFrom} to={dateTo} setFrom={setDateFrom} setTo={setDateTo} sedes={sedes} sedeFilter={sedeFilter} setSedeFilter={setSedeFilter} hasBD={hasBD} showRetirados={showRetirados} setShowRetirados={setShowRetirados} ccostos={ccostos} ccostoFilter={ccostoFilter} setCcostoFilter={setCcostoFilter} excSuperv={excSuperv} setExcSuperv={setExcSuperv}/>
      <div style={{marginTop:14}}><ErrBound>
        {view==="da"&&<Dash data={data} showR={setReport}/>}
        {view==="in"&&<Indiv data={data} sel={sel} setSel={setSel} showR={setReport}/>}
        {view==="cm"&&<Cmp data={data}/>}
        {view==="al"&&<Alt data={data}/>}
        {view==="mt"&&<Metas data={data}/>}
        {view==="si"&&<Sim data={data}/>}
        {view==="sc"&&<Score data={data} showR={setReport}/>}
        {view==="rp"&&<Rpt data={data} showR={setReport}/>}
        {view==="pc"&&<PComp data={data}/>}
      </ErrBound></div></div></div>;}
