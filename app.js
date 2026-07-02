/* ============================================================
   PARSER — Funciones puras de lectura del Excel
   ============================================================ */
function stripAccents(s){ return String(s==null?'':s).normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function normHeader(s){ return stripAccents(s).toUpperCase().trim().replace(/\s+/g,' '); }
function isEmptyCell(v){ return v===null||v===undefined||String(v).trim()===''; }
function toNum(v){
  if(v===null||v===undefined) return null;
  if(typeof v==='number') return isNaN(v)?null:v;
  const s=String(v).trim();
  if(s==='') return null;
  if(/^-+$/.test(s)) return 0;
  if(/^#DIV\/0!?$/i.test(s)) return null;
  if(/^N\/?A$/i.test(s)) return null;
  const c=s.replace(/\s/g,'').replace(/\./g,'').replace(',','.');
  let n=parseFloat(c);
  if(isNaN(n)) n=parseFloat(s.replace(/,/g,''));
  return isNaN(n)?null:n;
}
function isTotalLabel(label){
  const n=normHeader(label);
  return /^SUBTOTAL/.test(n)||/^TOTAL(\s|$)/.test(n)||n==='TOTAL GENERAL';
}
const DAY_NAMES=new Set(['DOMINGO','LUNES','MARTES','MIERCOLES','MCOLES','JUEVES','VIERNES','SABADO']);
const DAY_LETTERS=new Set(['D','L','M','W','J','V','S']);
function isDayHeader(h){
  const n=normHeader(h);
  if(DAY_NAMES.has(n)) return true;
  if(n.length===1&&DAY_LETTERS.has(n)) return true;
  return false;
}
function classifyColumns(headers){
  const roles={productIdx:0,planIdx:-1,totalIdx:-1,pctIdx:-1,diffIdx:-1,obsIdx:-1,blockIdx:-1,podaIdx:-1,dayIdxs:[]};
  const seenDays=new Set();
  headers.forEach((h,idx)=>{
    if(idx===0) return;
    const n=normHeader(h);
    if(!n) return;
    if(n==='BLOQUES'){roles.blockIdx=idx;return;}
    if(/^PODA/.test(n)&&roles.podaIdx===-1){roles.podaIdx=idx;return;}
    if(n.includes('%')){if(roles.pctIdx===-1)roles.pctIdx=idx;return;}
    if(n.includes('OBSERV')){roles.obsIdx=idx;return;}
    if(isDayHeader(n)){if(!seenDays.has(n)){seenDays.add(n);roles.dayIdxs.push(idx);}return;}
    if(n.includes('DIF')||n.includes('DESF')){if(roles.diffIdx===-1)roles.diffIdx=idx;return;}
    if(roles.planIdx===-1&&/PLAN|PROGRAMA(?!DOS)|ESTIMADO|PPTO/.test(n)&&!n.includes('ATRASO')&&!n.includes('ADELANTO')){roles.planIdx=idx;return;}
    if(roles.totalIdx===-1&&(/^TOTAL/.test(n)||n==='EJECUTADO'||n.includes('EJECUTADO'))){roles.totalIdx=idx;return;}
  });
  return roles;
}

// Títulos de sección conocidos
const SECTION_TITLES=[
  'PRODUCCION','PREPARACION DE CAMAS','SIEMBRAS PRODUCCION',
  'EJECUCION SEMANAL LABORES CULTURALES','SIEMBRAS ENRAIZAMIENTO Y GERMINACION',
  'EXVITRO','MIPE','MANTENIMIENTO','PROGRAMADOS','LABORATORIO BIOLOGICOS',
];
function matchKnownTitle(rawText,usedTitles){
  const v=normHeader(rawText);
  if(!v) return null;
  for(const t of SECTION_TITLES){
    if(usedTitles.has(t)) continue;
    if(v===t||v.startsWith(t)) return t;
  }
  return null;
}
function rowIsFullyEmpty(row){ return (row||[]).every(c=>isEmptyCell(c)); }
function rowLoneCellText(row){
  const ne=(row||[]).map((c,idx)=>({idx,c})).filter(x=>!isEmptyCell(x.c));
  if(ne.length===1&&ne[0].idx===0&&typeof row[0]==='string') return row[0];
  return null;
}
function maxColUsed(headers){ let max=0; headers.forEach((h,idx)=>{if(h!=null&&String(h).trim()!=='')max=idx;}); return max; }

function parseDetailSheet(aoa){
  const sections=[];
  const usedTitles=new Set();
  let i=0; const N=aoa.length;
  while(i<N){
    const row=aoa[i]||[];
    const loneText=rowLoneCellText(row);
    let title=loneText?matchKnownTitle(loneText,usedTitles):null;
    let isFallbackNew=false;
    if(!title&&loneText&&loneText.trim().length>3){
      const nxt=aoa[i+1]||[];
      const nxtNe=nxt.filter(c=>!isEmptyCell(c)).length;
      if(nxtNe>=2){title=loneText.trim();isFallbackNew=true;}
    }
    if(title){
      if(!isFallbackNew) usedTitles.add(title);
      let hr=i+1;
      while(hr<N&&rowIsFullyEmpty(aoa[hr])) hr++;
      if(hr>=N){i++;continue;}
      const headers=(aoa[hr]||[]).slice();
      const roles=classifyColumns(headers);
      const rows=[];
      let k=hr+1,emptyStreak=0;
      while(k<N){
        const r=aoa[k]||[];
        const loneNext=rowLoneCellText(r);
        const isNextKnown=loneNext?matchKnownTitle(loneNext,usedTitles):null;
        if(isNextKnown) break;
        if(rowIsFullyEmpty(r)){emptyStreak++;if(emptyStreak>=2)break;k++;continue;}
        emptyStreak=0;
        const label=normHeader(r[0]);
        if(label!=='PRODUCTO'&&label!=='LABOR') rows.push(r);
        k++;
      }
      if(roles.obsIdx===-1){
        const used=new Set([0,roles.planIdx,roles.totalIdx,roles.pctIdx,roles.diffIdx,roles.blockIdx,...roles.dayIdxs].filter(x=>x>=0));
        outer:for(let idx=1;idx<30;idx++){
          if(used.has(idx)) continue;
          for(const r of rows){
            const v=r[idx];
            if(typeof v==='string'&&v.trim().length>6&&/[A-Za-zÁÉÍÓÚáéíóúñÑ]/.test(v)&&(v.includes(' ')||v.includes(':'))){
              roles.obsIdx=idx;break outer;
            }
          }
        }
      }
      sections.push({title,headers,roles,rows,rowStart:i,rowEnd:k-1,maxCol:maxColUsed(headers)});
      i=k;
    } else { i++; }
  }
  return sections;
}

function unescapeXml(s){ return String(s||'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&'); }
async function extractEmbeddedImages(zip,targetSheetName){
  const readText=async(path)=>{const f=zip.file(path);if(!f)return null;return f.async('string');};
  try{
    const wbXml=await readText('xl/workbook.xml');
    if(!wbXml) return [];
    const sheetRe=/<sheet\b([^>]*)\/>/g;
    let m,sheetRid=null;
    while((m=sheetRe.exec(wbXml))){
      const attrs=m[1];
      const nameM=/name="([^"]*)"/.exec(attrs);
      const ridM=/r:id="([^"]*)"/.exec(attrs);
      if(nameM&&unescapeXml(nameM[1])===targetSheetName&&ridM){sheetRid=ridM[1];break;}
    }
    if(!sheetRid) return [];
    const wbRelsXml=await readText('xl/_rels/workbook.xml.rels');
    if(!wbRelsXml) return [];
    const relRe=/<Relationship\b([^>]*)\/>/g;
    let sheetTarget=null;
    while((m=relRe.exec(wbRelsXml))){
      const attrs=m[1];
      const idM=/Id="([^"]*)"/.exec(attrs);
      const targetM=/Target="([^"]*)"/.exec(attrs);
      if(idM&&idM[1]===sheetRid&&targetM){sheetTarget=targetM[1];break;}
    }
    if(!sheetTarget) return [];
    const sheetPath='xl/'+sheetTarget.replace(/^\.\.\//,'').replace(/^\/?xl\//,'');
    const sheetFileName=sheetPath.split('/').pop();
    const sheetRelsXml=await readText(`xl/worksheets/_rels/${sheetFileName}.rels`);
    if(!sheetRelsXml) return [];
    let drawingTarget=null;
    const relRe2=/<Relationship\b([^>]*)\/>/g;
    while((m=relRe2.exec(sheetRelsXml))){
      const attrs=m[1];
      if(/Type="[^"]*\/drawing"/.test(attrs)){const targetM=/Target="([^"]*)"/.exec(attrs);if(targetM){drawingTarget=targetM[1];break;}}
    }
    if(!drawingTarget) return [];
    const drawingPath='xl/'+drawingTarget.replace(/^\.\.\//,'').replace(/^\/?xl\//,'');
    const drawingFileName=drawingPath.split('/').pop();
    const drawingRelsXml=await readText(`xl/drawings/_rels/${drawingFileName}.rels`);
    const ridToMedia={};
    if(drawingRelsXml){
      const relRe3=/<Relationship\b([^>]*)\/>/g;
      while((m=relRe3.exec(drawingRelsXml))){
        const attrs=m[1];
        const idM=/Id="([^"]*)"/.exec(attrs);
        const targetM=/Target="([^"]*)"/.exec(attrs);
        if(idM&&targetM&&/\.(png|jpe?g|gif|bmp)$/i.test(targetM[1])){ridToMedia[idM[1]]='xl/'+targetM[1].replace(/^\.\.\//,'').replace(/^\/?xl\//,'');}
      }
    }
    const drawingXml=await readText(drawingPath);
    if(!drawingXml) return [];
    const anchorRe=/<xdr:twoCellAnchor\b[\s\S]*?<\/xdr:twoCellAnchor>/g;
    const anchorsRaw=drawingXml.match(anchorRe)||[];
    const anchors=[];
    for(const a of anchorsRaw){
      const fromM=/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a);
      const toM=/<xdr:to>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a);
      const embedM=/r:embed="([^"]*)"/.exec(a);
      const nameM=/<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(a);
      if(!fromM||!toM||!embedM) continue;
      const mediaPath=ridToMedia[embedM[1]];
      if(!mediaPath) continue;
      anchors.push({name:nameM?unescapeXml(nameM[1]):'Imagen',fromCol:parseInt(fromM[1],10),fromRow:parseInt(fromM[2],10),toCol:parseInt(toM[1],10),toRow:parseInt(toM[2],10),mediaPath});
    }
    const out=[];
    for(const a of anchors){
      const f=zip.file(a.mediaPath);
      if(!f) continue;
      const base64=await f.async('base64');
      const ext=a.mediaPath.split('.').pop().toLowerCase();
      const mime=ext==='jpg'?'jpeg':ext;
      out.push({...a,dataUri:`data:image/${mime};base64,${base64}`});
    }
    return out;
  } catch(e){ return []; }
}
function attachImagesToSections(sections,images){
  sections.forEach(sec=>{sec.images=[];});
  images.forEach(img=>{
    let best=null,bestOverlap=0;
    sections.forEach(sec=>{
      const overlap=Math.min(sec.rowEnd,img.toRow)-Math.max(sec.rowStart,img.fromRow)+1;
      if(overlap>bestOverlap){bestOverlap=overlap;best=sec;}
    });
    if(best) best.images.push(img);
  });
  sections.forEach(sec=>{sec.images.sort((a,b)=>a.fromCol-b.fromCol);});
}

/* ── Hoja resumen ── */
function parseResumenSheet(aoa){
  let headerRowIdx=-1;
  for(let i=0;i<aoa.length;i++){
    if(normHeader((aoa[i]||[])[0])==='PRODUCTO'){headerRowIdx=i;break;}
  }
  if(headerRowIdx===-1) return{headers:[],rows:[],titulo:null,roles:null};
  const headerRow=aoa[headerRowIdx]||[];
  const subRow=aoa[headerRowIdx+1]||[];
  const headers=headerRow.map((h,idx)=>{
    const base=h==null?'':String(h).replace(/\n/g,' ').trim();
    const sub=subRow[idx];
    if(sub!=null&&String(sub).trim()!==''){return base?`${base} - ${String(sub).replace(/\n/g,' ').trim()}`:String(sub).trim();}
    return base;
  });
  const rows=[];
  let i=headerRowIdx+2;
  while(i<aoa.length){
    const r=aoa[i]||[];
    const label=normHeader(r[0]);
    if(!label){i++;if(rowIsFullyEmpty(aoa[i])) break;else continue;}
    rows.push(r);
    if(label==='TOTAL GENERAL') break;
    i++;
  }
  let titulo=null;
  for(const r of aoa.slice(0,4)){if(r&&r[0]){titulo=r[0];break;}}
  const roles=classifyResumenColumns(headers);
  return{headers,rows,titulo,roles};
}
function classifyResumenColumns(headers){
  const find=(re)=>headers.findIndex(h=>re.test(normHeader(h)));
  return{
    productIdx:0, unidadIdx:find(/^UNIDAD/),
    realIdx:find(/EJECUCION REAL.*POSTCOSECHA|EJECUCION REAL DOMINGO/),
    pronosticoCampoIdx:find(/PRONOSTICO VIERNES/),
    totalProyectadoIdx:find(/EJECUCION REAL \+ PRONOSTICO/),
    per1Idx:find(/UNA SEMANA.*PRONOSTICO|PRONOSTICO\s*PER1|PRONOSTICO\s*PER 1/),
    per1PctIdx:find(/%.*PER\s*1(?!\d)/),
    per4Idx:find(/CUATRO SEMANAS/),
    per4PctIdx:find(/%.*PER\s*4/),
    pronosticadorIdx:find(/PROYEC PRONOSTICADOR/),
    pronosticadorPctIdx:find(/%.*PRONOSTICADOR/),
    planFincaIdx:find(/FLORACION PLAN FINCA|PLAN FINCA/),
    planFincaPctIdx:find(/%.*PLAN DE SIEMBRA|%.*PRESUPUESTO|EJECUCION REAL vs PLAN/),
    aprovIdx:find(/APROV/),
    obsIdx:find(/SECADERA|OBSERV/),
  };
}

/* Productos válidos para Corte — solo los de la empresa */
const PRODUCTOS_CORTE=new Set(['CUSHION','DAISY','NOVELTY','POMPON','MICROPOMPON','SNAPDRAGON','SNAPDRAGONS','SOLIDAGO','TRACHELIUM','TRACHELLIUM']);
function isProductoCorte(label){
  const n=normHeader(label);
  if(n==='TOTAL GENERAL') return true; // siempre incluir totales
  return PRODUCTOS_CORTE.has(n);
}
/* Productos válidos para secciones de detalle */
const PRODUCTOS_DETALLE=new Set(['POMPON','MICROPOMPON','SNAPDRAGON','SNAPDRAGONS','SOLIDAGO','TRACHELIUM','TRACHELLIUM']);
function isProductoDetalle(label){
  const n=normHeader(label);
  for(const p of PRODUCTOS_DETALLE){ if(n.includes(p)) return true; }
  return false;
}
/* Herbicidas — sin plan, solo mostrar ejecutado */
function isHerbicida(label){ return normHeader(label).includes('HERBICIDA'); }
/* Mantenimiento — solo estas 3 filas */
const MANT_VALIDAS=new Set(['ESTERILIZACION','CAMBIO DE CUBIERTA','LAVADO CUBIERTA']);
function isMantenimientoValida(label){
  const n=normHeader(label);
  for(const v of MANT_VALIDAS){ if(n.includes(v)) return true; }
  return false;
}

function computeResumenKPIs(resumen){
  const{rows}=resumen;
  const roles=resumen.roles||classifyResumenColumns(resumen.headers);
  const enriched=rows.map(r=>{
    const label=r[0]==null?'':String(r[0]).trim();
    const isTotalGeneral=normHeader(label)==='TOTAL GENERAL';
    const plan=roles.planFincaIdx>=0?toNum(r[roles.planFincaIdx]):null;
    const total=roles.totalProyectadoIdx>=0?toNum(r[roles.totalProyectadoIdx]):null;
    const real=roles.realIdx>=0?toNum(r[roles.realIdx]):null;
    const campo=roles.pronosticoCampoIdx>=0?toNum(r[roles.pronosticoCampoIdx]):null;
    const per1=roles.per1Idx>=0?toNum(r[roles.per1Idx]):null;
    const per4=roles.per4Idx>=0?toNum(r[roles.per4Idx]):null;
    const pronosticador=roles.pronosticadorIdx>=0?toNum(r[roles.pronosticadorIdx]):null;
    const pctPlanFinca=roles.planFincaPctIdx>=0?toNum(r[roles.planFincaPctIdx]):null;
    const pctPer1=roles.per1PctIdx>=0?toNum(r[roles.per1PctIdx]):null;
    const pctPer4=roles.per4PctIdx>=0?toNum(r[roles.per4PctIdx]):null;
    const pctPronosticador=roles.pronosticadorPctIdx>=0?toNum(r[roles.pronosticadorPctIdx]):null;
    const aprov=roles.aprovIdx>=0?toNum(r[roles.aprovIdx]):null;
    const unidad=roles.unidadIdx>=0?r[roles.unidadIdx]:null;
    const obs=roles.obsIdx>=0?(r[roles.obsIdx]!=null?String(r[roles.obsIdx]).trim():''):'';
    return{raw:r,label,isTotalGeneral,plan,total,real,campo,per1,per4,pronosticador,pctPlanFinca,pctPer1,pctPer4,pctPronosticador,aprov,unidad,obs};
  });
  const validRows=enriched.filter(e=>!e.isTotalGeneral&&isProductoCorte(e.label));
  const totalGeneralRow=enriched.find(e=>e.isTotalGeneral)||null;
  return{roles,enriched,validRows,totalGeneralRow};
}

function computeSectionKPIs(section){
  const{roles,rows}=section;
  let sumPlan=0,sumTotal=0,hasPlan=false,hasTotal=false,productRows=0;
  const enriched=rows.map(r=>{
    const label=r[0]==null?'':String(r[0]).trim();
    const plan=roles.planIdx>=0?toNum(r[roles.planIdx]):null;
    const total=roles.totalIdx>=0?toNum(r[roles.totalIdx]):null;
    const isSubtotal=isTotalLabel(label);
    let pct=null;
    if(plan!==null&&plan>0&&total!==null) pct=total/plan;
    else if(roles.pctIdx>=0) pct=toNum(r[roles.pctIdx]);
    // Para mantenimiento: inferir si falta
    if(pct===null&&plan!==null&&plan>0&&total!==null) pct=total/plan;
    const diff=(plan!==null&&total!==null)?(total-plan):(roles.diffIdx>=0?toNum(r[roles.diffIdx]):null);
    const unidad=roles.blockIdx>=0?(r[roles.blockIdx]==null?'':String(r[roles.blockIdx]).trim()):'';
    const poda=roles.podaIdx>=0?toNum(r[roles.podaIdx]):null;
    if(label&&!isSubtotal){
      if(plan!==null){sumPlan+=plan;hasPlan=true;}
      if(total!==null){sumTotal+=total;hasTotal=true;productRows++;}
    }
    return{raw:r,label,plan,total,pct,diff,isSubtotal,unidad,poda};
  });
  const areaPct=(hasPlan&&sumPlan>0)?(sumTotal/sumPlan):null;
  return{plan:hasPlan?sumPlan:null,total:hasTotal?sumTotal:null,pct:areaPct,diff:(hasPlan&&hasTotal)?(sumTotal-sumPlan):null,productRows,enriched,hasData:hasPlan||hasTotal};
}

function computeDayInsights(section){
  const{roles,rows}=section;
  if(!roles.dayIdxs.length) return null;
  const dayTotals=roles.dayIdxs.map(idx=>({idx,total:0,hasData:false,topLabel:null,topVal:-Infinity}));
  rows.forEach(r=>{
    const label=r[0]==null?'':String(r[0]).trim();
    if(!label||isTotalLabel(label)) return;
    dayTotals.forEach(d=>{
      const v=toNum(r[d.idx]);
      if(v===null) return;
      d.hasData=true; d.total+=v;
      if(v>d.topVal){d.topVal=v;d.topLabel=label;}
    });
  });
  const withData=dayTotals.filter(d=>d.hasData);
  if(!withData.length) return null;
  const best=withData.slice().sort((a,b)=>b.total-a.total)[0];
  const worst=withData.slice().sort((a,b)=>a.total-b.total)[0];
  return{dayTotals:withData,best,worst};
}

/* ============================================================
   ESTADO Y UTILIDADES DE FORMATO
   ============================================================ */
const STATE={fileName:'',weekLabel:'',fincaLabel:'',sections:[],resumen:null,resumenKpis:null,pages:[],currentPage:0};

function fmtInt(n){ if(n===null||n===undefined||isNaN(n)) return '—'; return Math.round(n).toLocaleString('es-CO'); }
function fmtNum1(n){ if(n===null||n===undefined||isNaN(n)) return '—'; return n.toLocaleString('es-CO',{maximumFractionDigits:1}); }
function fmtPct(n){ if(n===null||n===undefined||isNaN(n)) return 'Sin dato'; return (n*100).toLocaleString('es-CO',{maximumFractionDigits:1})+'%'; }
function pctTone(n){ if(n===null||n===undefined||isNaN(n)) return 'muted'; if(n>=0.95) return 'ok'; if(n>=0.80) return 'warn'; return 'bad'; }
function toneColor(tone){ return tone==='ok'?'#1E8E5A':tone==='warn'?'#D98B1E':tone==='bad'?'#C0392B':'#B7C0D1'; }
function escapeHtml(s){ return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function titleCaseEs(s){ return String(s||'').toLowerCase().replace(/(^|\s)(\S)/g,(m,sp,c)=>sp+c.toUpperCase()); }
function displayAreaName(title){
  const n=normHeader(title);
  if(n==='PRODUCCION') return 'Corte';
  if(n==='EJECUCION SEMANAL LABORES CULTURALES') return 'Labores Culturales';
  if(n==='SIEMBRAS ENRAIZAMIENTO Y GERMINACION') return 'Enraizamiento y Germinación';
  if(n==='LABORATORIO BIOLOGICOS') return 'Lab. Biológicos';
  return titleCaseEs(title);
}

const ICON_MAP=[
  [/PRODUCCION/,'bi-basket3'],[/PREPARACION DE CAMAS/,'bi-grid-3x3-gap'],[/SIEMBRAS PRODUCCION/,'bi-flower1'],
  [/LABORES CULTURALES/,'bi-scissors'],[/ENRAIZAMIENTO/,'bi-tree'],[/EXVITRO/,'bi-droplet-half'],
  [/MIPE/,'bi-bug'],[/MANTENIMIENTO/,'bi-tools'],[/PROGRAMADOS/,'bi-clipboard-check'],[/LABORATORIO/,'bi-eyedropper'],
];
function iconFor(title){ const n=normHeader(title); for(const[re,icon]of ICON_MAP)if(re.test(n))return icon; return 'bi-bar-chart-line'; }
const PALETTE=['#0D2957','#C9A227','#1E4D9E','#E3956B','#3F7A5C','#9FB6DE','#B08D57','#7A5C9E'];
const DAY_DISPLAY={'DOMINGO':'Domingo','LUNES':'Lunes','MARTES':'Martes','MIERCOLES':'Miércoles','MCOLES':'Miércoles','JUEVES':'Jueves','VIERNES':'Viernes','SABADO':'Sábado','D':'Domingo','L':'Lunes','M':'Martes','W':'Miércoles','J':'Jueves','V':'Viernes','S':'Sábado'};
function dayDisplay(h){ return DAY_DISPLAY[normHeader(h)]||titleCaseEs(h); }

/* ── LECTURA DEL ARCHIVO ── */
function detectSheets(sheetNames){
  let detail=null,resumen=null;
  for(const name of sheetNames){
    const n=name.trim();
    if(/^AHORRO/i.test(n)||/^TURNO/i.test(n)||/^VARIOS/i.test(n)||/^DOMINGO/i.test(n)||/^LUNES/i.test(n)) continue;
    if(!detail&&/^\d{3,5}$/.test(n)) detail=name;
    if(!resumen&&/^\d{4}\s+\d{1,2}$/.test(n)) resumen=name;
  }
  return{detail,resumen};
}

function handleFile(file){
  showLoading('Leyendo archivo…');
  setUploadProgress(15,'Leyendo archivo…');
  const reader=new FileReader();
  reader.onerror=()=>{hideLoading();showUploadError('No se pudo leer el archivo. Verifica que no esté abierto en otro programa.');};
  reader.onload=async(e)=>{
    try{
      setUploadProgress(40,'Analizando estructura del Excel…');
      const data=new Uint8Array(e.target.result);
      const wb=XLSX.read(data,{type:'array',cellDates:false});
      const{detail,resumen}=detectSheets(wb.SheetNames);
      if(!detail) throw new Error('No se encontró la hoja de detalle semanal. Hojas disponibles: '+wb.SheetNames.join(', '));
      setUploadProgress(60,'Detectando áreas y métricas…');
      const detailAoa=XLSX.utils.sheet_to_json(wb.Sheets[detail],{header:1,defval:null,raw:true});
      const sections=parseDetailSheet(detailAoa)
        .filter(sec=>normHeader(sec.title)!=='PROGRAMADOS')
        .map(sec=>({...sec,kpis:computeSectionKPIs(sec),dayInsight:computeDayInsights(sec),images:[]}));
      if(!sections.length) throw new Error('El archivo no tiene el formato esperado dentro de la hoja "'+detail+'".');
      try{
        const zip=await JSZip.loadAsync(data);
        const images=await extractEmbeddedImages(zip,detail);
        attachImagesToSections(sections,images);
      } catch(imgErr){ console.error('Imágenes:',imgErr); }
      let resumenData=null,resumenKpis=null;
      if(resumen){
        setUploadProgress(80,'Cruzando con el resumen ejecutivo…');
        const resumenAoa=XLSX.utils.sheet_to_json(wb.Sheets[resumen],{header:1,defval:null,raw:true});
        resumenData=parseResumenSheet(resumenAoa);
        if(resumenData.rows.length) resumenKpis=computeResumenKPIs(resumenData);
      }
      setUploadProgress(96,'Construyendo el tablero…');
      STATE.fileName=file.name;
      STATE.sections=sections;
      STATE.resumen=resumenData;
      STATE.resumenKpis=resumenKpis;
      STATE.fincaLabel=(resumenData&&resumenData.titulo)?String(resumenData.titulo):'';
      STATE.weekLabel=resumen?resumen.trim():(detail||'');
      STATE.pages=buildPages();
      STATE.currentPage=0;
      setTimeout(()=>{hideLoading();mountApp();},250);
    } catch(err){
      console.error(err); hideLoading();
      showUploadError('No se pudo procesar el archivo: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function buildPages(){
  const pages=[{type:'overview'}];
  const mipeSecs=[];
  for(const sec of STATE.sections){
    if(normHeader(sec.title)==='MIPE') mipeSecs.push({type:'section',section:sec});
    else pages.push({type:'section',section:sec});
  }
  // MIPE va al final
  for(const m of mipeSecs) pages.push(m);
  return pages;
}

/* ── MONTAJE ── */
function mountApp(){
  document.getElementById('uploadScreen').style.display='none';
  document.getElementById('appShell').classList.add('show');
  document.getElementById('fileNameChip').textContent=STATE.fileName;
  document.getElementById('weekChip').textContent=weekLabelDisplay();
  if(STATE.fincaLabel){document.getElementById('fincaChip').style.display='flex';document.getElementById('fincaChipText').textContent=STATE.fincaLabel;}
  buildNav(); renderPage(0);
}
function weekLabelDisplay(){
  const m=STATE.weekLabel.match(/^(\d{4})\s+(\d{1,2})$/);
  if(m) return `Semana ${m[2]} · ${m[1]}`;
  return STATE.weekLabel||'—';
}

/* ── NAVEGACIÓN ── */
function buildNav(){
  const ul=document.getElementById('navList');
  ul.innerHTML='';
  STATE.pages.forEach((pg,idx)=>{
    const li=document.createElement('li');
    const a=document.createElement('div');
    a.className='nav-link-item'+(idx===STATE.currentPage?' active':'');
    a.dataset.idx=idx;
    let icon='bi-speedometer2',label='Panorama general',chip='';
    if(pg.type==='section'){
      icon=iconFor(pg.section.title);
      label=displayAreaName(pg.section.title);
      const pct=pg.section.kpis.pct;
      const isMipe=normHeader(pg.section.title)==='MIPE';
      if(pct!==null&&!isMipe){
        const tone=pctTone(pct);
        chip=`<span class="pct-chip" style="background:${toneColor(tone)}22;color:${toneColor(tone)}">${fmtPct(pct)}</span>`;
      }
    }
    a.innerHTML=`<i class="bi ${icon}"></i><span>${escapeHtml(label)}</span>${chip}`;
    a.addEventListener('click',()=>renderPage(idx));
    li.appendChild(a);
    ul.appendChild(li);
    if(idx===0&&STATE.pages.length>1){
      const lbl=document.createElement('div');
      lbl.className='nav-section-label';
      lbl.textContent='Áreas de la finca';
      ul.appendChild(lbl);
    }
  });
}
function refreshNavActive(){ document.querySelectorAll('.nav-link-item').forEach(el=>el.classList.toggle('active',Number(el.dataset.idx)===STATE.currentPage)); }

/* ── DESPACHO DE PÁGINAS ── */
function renderPage(idx){
  if(idx<0||idx>=STATE.pages.length) return;
  STATE.currentPage=idx;
  const page=STATE.pages[idx];
  const container=document.getElementById('pageArea');
  container.innerHTML='';
  container.classList.remove('slide-fade'); void container.offsetWidth; container.classList.add('slide-fade');
  if(page.type==='overview') renderOverviewPage(container);
  else renderSectionPage(container,page.section);
  refreshNavActive();
  if(window.innerWidth<=900) document.getElementById('sidebar').classList.remove('open');
  window.scrollTo({top:0,behavior:'smooth'});
}
function pageNavHTML(){
  const idx=STATE.currentPage,n=STATE.pages.length;
  return `<div class="page-nav-btns">
    <button class="btn" id="prevPageBtn" ${idx<=0?'disabled':''} title="Área anterior"><i class="bi bi-chevron-left"></i></button>
    <span class="page-dots">${idx+1} / ${n}</span>
    <button class="btn" id="nextPageBtn" ${idx>=n-1?'disabled':''} title="Siguiente área"><i class="bi bi-chevron-right"></i></button>
  </div>`;
}
function wirePageNav(){
  const prev=document.getElementById('prevPageBtn');
  const next=document.getElementById('nextPageBtn');
  if(prev) prev.addEventListener('click',()=>renderPage(STATE.currentPage-1));
  if(next) next.addEventListener('click',()=>renderPage(STATE.currentPage+1));
}
document.addEventListener('keydown',(e)=>{
  if(!document.getElementById('appShell').classList.contains('show')) return;
  if(e.target.tagName==='INPUT') return;
  if(e.key==='ArrowRight') renderPage(STATE.currentPage+1);
  if(e.key==='ArrowLeft') renderPage(STATE.currentPage-1);
});

/* ============================================================
   COMPONENTES VISUALES
   ============================================================ */
function kpiCard({label,value,sub,icon,tone,gold}){
  return `<div class="kpi-card ${gold?'gold':''} ${tone||''}">
    <i class="bi ${icon} kpi-icon"></i>
    <div class="kpi-label">${escapeHtml(label)}</div>
    <div class="kpi-value">${value}</div>
    ${sub?`<div class="kpi-sub">${sub}</div>`:''}
  </div>`;
}
function kpiCardDesfase({diff,unit,rows,tone}){
  const diffTxt=diff!==null?(diff>=0?'+':'')+fmtInt(diff)+'<small style="font-size:.9rem"> '+escapeHtml(unit||'')+'</small>':'—';
  const diffSub=diff!==null?(diff>=0?'por encima del plan':'por debajo del plan'):'';
  // Top 3 labores con mayor desfase negativo
  const worstRows=rows.filter(r=>r.diff!==null&&r.diff<0&&!r.isSubtotal).sort((a,b)=>a.diff-b.diff).slice(0,3);
  const miniList=worstRows.length?`<div class="kpi-desfase-list">${worstRows.map(r=>`<div class="kpi-desfase-item"><span title="${escapeHtml(r.label)}">${escapeHtml(titleCaseEs(r.label))}</span><b style="color:var(--danger)">${fmtInt(r.diff)}</b></div>`).join('')}</div>`:'';
  return `<div class="kpi-card ${tone||''}">
    <i class="bi bi-arrow-left-right kpi-icon"></i>
    <div class="kpi-label">Desfase${unit?' en '+escapeHtml(unit):''}</div>
    <div class="kpi-value">${diffTxt}</div>
    ${diffSub?`<div class="kpi-sub">${diffSub}</div>`:''}
    ${miniList}
  </div>`;
}

/* Strip resumen de sección: promedio, mejor, peor labor */
function sectionSummaryStrip(rows){
  const withPct=rows.filter(r=>r.pct!==null&&!r.isSubtotal);
  if(!withPct.length) return '';
  const avg=withPct.reduce((a,r)=>a+r.pct,0)/withPct.length;
  const best=withPct.slice().sort((a,b)=>b.pct-a.pct)[0];
  const worst=withPct.slice().sort((a,b)=>a.pct-b.pct)[0];
  const avgTone=pctTone(avg);
  return `<div class="section-summary-strip">
    <div class="sss-card">
      <div class="sss-icon" style="background:${toneColor(avgTone)}22"><i class="bi bi-speedometer2" style="color:${toneColor(avgTone)}"></i></div>
      <div><div class="sss-label">Promedio sección</div><div class="sss-value" style="color:${toneColor(avgTone)}">${fmtPct(avg)}</div></div>
    </div>
    <div class="sss-card">
      <div class="sss-icon" style="background:var(--success-soft)"><i class="bi bi-trophy" style="color:var(--success)"></i></div>
      <div><div class="sss-label">Mayor cumplimiento</div><div class="sss-value" style="color:var(--success)">${fmtPct(best.pct)}</div><div class="sss-sub">${escapeHtml(titleCaseEs(best.label))}</div></div>
    </div>
    <div class="sss-card">
      <div class="sss-icon" style="background:var(--danger-soft)"><i class="bi bi-exclamation-diamond" style="color:var(--danger)"></i></div>
      <div><div class="sss-label">Menor cumplimiento</div><div class="sss-value" style="color:var(--danger)">${fmtPct(worst.pct)}</div><div class="sss-sub">${escapeHtml(titleCaseEs(worst.label))}</div></div>
    </div>
  </div>`;
}

/* Barras de labor */
function laborBarList(rows,opts){
  opts=opts||{};
  const planLabel=opts.planLabel||'Plan';
  const execLabel=opts.execLabel||'Ejecutado';
  if(!rows.length) return '<div class="chart-empty"><i class="bi bi-bar-chart"></i>Sin datos esta semana</div>';
  // Calcular el máximo real para saber cuánto sobrepasa el 100%
  const maxPct=Math.max(1,...rows.filter(r=>r.pct!==null).map(r=>r.pct));
  // La barra al 100% ocupa este porcentaje del track
  const scale100=maxPct>1?(100/maxPct)*100:100; // posición de la línea en el track
  const scaleLineHTML=`<div class="bar-scale-line" style="left:${scale100}%"></div><span class="bar-scale-label" style="left:${scale100}%">100%</span>`;
  return '<div class="labor-list">'+rows.map(r=>{
    const tone=pctTone(r.pct);
    const c=toneColor(tone);
    const herbicida=isHerbicida(r.label);
    if(herbicida){
      return `<div class="labor-row">
        <div class="labor-row-head">
          <span class="labor-name">${escapeHtml(titleCaseEs(r.label))}</span>
          <span class="badge-soft" style="background:var(--navy-soft);color:var(--navy-3)">Solo registro</span>
        </div>
        <div class="labor-caption">Ejecutado: <b>${r.total!==null?fmtInt(r.total):'Sin datos'}</b> — sin plan programado</div>
      </div>`;
    }
    // Escala dinámica: si alguna barra supera 100%, reescalar todas
    const w=r.pct!==null?Math.min(100,Math.max((r.pct/Math.max(maxPct,1))*100,r.pct>0?2:0)):0;
    const isOver=r.pct!==null&&r.pct>1;
    const caption=(r.plan!==null||r.total!==null)
      ?`${escapeHtml(planLabel)}: <b>${r.plan!==null?fmtInt(r.plan):'—'}</b> &nbsp;·&nbsp; ${escapeHtml(execLabel)}: <b>${r.total!==null?fmtInt(r.total):'—'}</b>${r.diff!==null?' &nbsp;·&nbsp; Desfase: <b style="color:'+(r.diff>=0?'var(--success)':'var(--danger)')+'">'+( r.diff>=0?'+':'')+fmtInt(r.diff)+'</b>':''}`
      :'';
    return `<div class="labor-row">
      <div class="labor-row-head">
        <span class="labor-name">${escapeHtml(titleCaseEs(r.label))}${r.unidad?` <span class="unit-tag">${escapeHtml(r.unidad)}</span>`:''}</span>
        <span class="badge-soft" style="background:${c}22;color:${c}">${fmtPct(r.pct)}</span>
      </div>
      <div class="bar-wrap" style="padding-top:18px;">
        ${scaleLineHTML}
        ${isOver?`<span class="bar-over-label">+${fmtPct(r.pct-1)} sobre plan</span>`:''}
        <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${c}"></div></div>
      </div>
      ${caption?`<div class="labor-caption">${caption}</div>`:''}
    </div>`;
  }).join('')+'</div>';
}

/* Dona CSS */
function donutCSS(items,opts){
  opts=opts||{};
  const total=items.reduce((a,i)=>a+i.value,0);
  if(!total) return '<div class="chart-empty"><i class="bi bi-pie-chart"></i>Sin ejecución registrada</div>';
  let acc=0;
  const stops=items.map((it,i)=>{const from=acc/total*100;acc+=it.value;const to=acc/total*100;return `${PALETTE[i%PALETTE.length]} ${from}% ${to}%`;}).join(', ');
  const legend=items.map((it,i)=>{
    const tone=pctTone(it.pct);
    const c=it.pct!==null?toneColor(tone):'#8893A6';
    return `<div class="legend-row"><span><span class="legend-dot" style="background:${PALETTE[i%PALETTE.length]}"></span>${escapeHtml(titleCaseEs(it.label))}</span><b style="color:${c}">${fmtPct(it.pct)}</b></div>`;
  }).join('');
  const centerPct=opts.centerPct;
  const centerTxt=centerPct!==undefined&&centerPct!==null?fmtPct(centerPct):fmtInt(total);
  return `<div style="display:flex;gap:22px;align-items:center;flex-wrap:wrap;justify-content:center;">
    <div class="donut-ring" style="background:conic-gradient(${stops})">
      <div class="donut-hole"><b>${centerTxt}</b><small>${escapeHtml(opts.centerLabel||'cumplimiento')}</small></div>
    </div>
    <div class="legend-list" style="flex:1;min-width:180px;">${legend}</div>
  </div>`;
}

/* Insight del día */
function dayInsightHTML(section){
  const di=section.dayInsight;
  if(!di) return '';
  const bestName=dayDisplay(section.headers[di.best.idx]);
  const worstName=dayDisplay(section.headers[di.worst.idx]);
  return `<div class="insight-box">
    <div class="ib-icon"><i class="bi bi-lightbulb"></i></div>
    <div><b>${escapeHtml(bestName)}</b> fue el día de mayor actividad (${fmtInt(di.best.total)}), liderado por <b>${escapeHtml(titleCaseEs(di.best.topLabel))}</b>.
    <b>${escapeHtml(worstName)}</b> registró la menor actividad (${fmtInt(di.worst.total)})${di.worst.topLabel?`, con <b>${escapeHtml(titleCaseEs(di.worst.topLabel))}</b> como el más activo ese día`:''}.
    </div>
  </div>`;
}

/* Galería de imágenes */
function imageGalleryHTML(images){
  if(!images||!images.length) return '';
  return `<div class="chart-card">
    <h6><i class="bi bi-images"></i> Gráficos y tablas del reporte</h6>
    <div class="chart-sub">Imágenes adjuntas en esta sección del Excel</div>
    <div class="embedded-gallery">
      ${images.map(img=>`<a class="embedded-img-wrap" href="${img.dataUri}" target="_blank" rel="noopener" title="Abrir en tamaño completo">
        <img class="embedded-img" src="${img.dataUri}" alt="${escapeHtml(img.name||'Imagen del reporte')}">
      </a>`).join('')}
    </div>
  </div>`;
}

/* Observaciones */
function sectionObsHTML(sec){
  if(sec.roles.obsIdx<0) return `<div class="chart-empty"><i class="bi bi-check2-circle"></i>Sin observaciones esta semana</div>`;
  const items=sec.rows.filter(r=>r[sec.roles.obsIdx]&&String(r[sec.roles.obsIdx]).trim())
    .map(r=>`<div class="obs-item"><b>${escapeHtml(r[0]||'')}:</b> ${escapeHtml(String(r[sec.roles.obsIdx]).trim())}</div>`);
  if(!items.length) return `<div class="chart-empty"><i class="bi bi-check2-circle"></i>Sin observaciones esta semana</div>`;
  return `<div class="obs-list">${items.join('')}</div>`;
}

/* Tabla interactiva */
function buildInteractiveTable(cfg){
  const{wrapId,searchId,exportXlsxId,exportCsvId,headers,fullHeaders,rows,roles,filename,isResumen}=cfg;
  const colIdxs=fullHeaders.map((h,idx)=>({h,idx})).filter(x=>x.h&&String(x.h).trim());
  let sortState={idx:null,dir:1},filterText='';
  function isPctCol(idx){
    if(isResumen) return[roles.per1PctIdx,roles.per4PctIdx,roles.planFincaPctIdx,roles.pronosticadorPctIdx].includes(idx);
    return roles&&roles.pctIdx===idx;
  }
  function cellDisplay(v,idx){
    if(v===null||v===undefined||String(v).trim()==='') return '<span class="cell-muted">—</span>';
    if(/^#DIV\/0!?$/i.test(String(v))) return '<span class="cell-muted">Sin programa</span>';
    if(isPctCol(idx)){const num=toNum(v);if(num===null) return '<span class="cell-muted">Sin programa</span>';const tone=pctTone(num);return `<span class="cell-pct ${tone}">${fmtPct(num)}</span>`;}
    if(typeof v==='number') return v%1===0?fmtInt(v):fmtNum1(v);
    return escapeHtml(v);
  }
  function getFiltered(){
    let data=rows;
    if(filterText){const f=filterText.toLowerCase();data=data.filter(r=>r.some(c=>c!=null&&String(c).toLowerCase().includes(f)));}
    if(sortState.idx!==null){
      data=data.slice().sort((a,b)=>{
        const va=a[sortState.idx],vb=b[sortState.idx];
        const na=toNum(va),nb=toNum(vb);
        let cmp;
        if(na!==null&&nb!==null) cmp=na-nb;
        else cmp=String(va||'').localeCompare(String(vb||''),'es');
        return cmp*sortState.dir;
      });
    }
    return data;
  }
  function draw(){
    const data=getFiltered();
    const wrap=document.getElementById(wrapId);
    const theadHtml=`<tr>${colIdxs.map(c=>{const active=sortState.idx===c.idx;const arrow=active?(sortState.dir===1?'bi-sort-up':'bi-sort-down'):'bi-arrow-down-up';return `<th data-idx="${c.idx}">${escapeHtml(String(c.h).replace(/\r?\n/g,' '))} <i class="bi ${arrow}"></i></th>`;}).join('')}</tr>`;
    const tbodyHtml=data.map(r=>{const label=normHeader(r[0]);const isSub=isTotalLabel(label);return `<tr class="${isSub?'row-subtotal':''}">${colIdxs.map(c=>`<td>${cellDisplay(r[c.idx],c.idx)}</td>`).join('')}</tr>`;}).join('');
    wrap.innerHTML=`<div class="table-responsive-custom"><table class="data-table"><thead>${theadHtml}</thead><tbody>${tbodyHtml||`<tr><td colspan="${colIdxs.length}" class="text-center cell-muted py-3">Sin resultados</td></tr>`}</tbody></table></div>`;
    wrap.querySelectorAll('thead th').forEach(th=>{th.addEventListener('click',()=>{const idx=Number(th.dataset.idx);sortState=sortState.idx===idx?{idx,dir:-sortState.dir}:{idx,dir:1};draw();});});
  }
  draw();
  if(searchId){const inp=document.getElementById(searchId);if(inp)inp.addEventListener('input',()=>{filterText=inp.value;draw();});}
  if(exportXlsxId){const btn=document.getElementById(exportXlsxId);if(btn)btn.addEventListener('click',()=>exportRows(getFiltered(),colIdxs,filename,'xlsx'));}
  if(exportCsvId){const btn=document.getElementById(exportCsvId);if(btn)btn.addEventListener('click',()=>exportRows(getFiltered(),colIdxs,filename,'csv'));}
}
function exportRows(rows,colIdxs,filename,type){
  const headers=colIdxs.map(c=>String(c.h).replace(/\r?\n/g,' '));
  const aoa=[headers,...rows.map(r=>colIdxs.map(c=>{const v=r[c.idx];if(/^#DIV\/0!?$/i.test(String(v)))return null;return v;}))];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Datos');
  const safeName=String(filename).replace(/[^\w\- ]/g,'').replace(/\s+/g,'_').slice(0,60)||'tablero_flores_el_trigal';
  if(type==='csv') XLSX.writeFile(wb,safeName+'.csv',{bookType:'csv'});
  else XLSX.writeFile(wb,safeName+'.xlsx');
}
function tableToolbar(wrapId,searchId,exportXlsxId,exportCsvId,title){
  return `<div class="table-toolbar">
    <h6 class="mb-0"><i class="bi bi-table text-gold"></i> ${escapeHtml(title)}</h6>
    <div class="d-flex gap-2 align-items-center flex-wrap">
      ${searchId?`<div class="search-box"><i class="bi bi-search"></i><input type="text" id="${searchId}" placeholder="Buscar…"></div>`:''}
      <div class="export-btns d-flex gap-2">
        <button class="btn btn-outline-navy" id="${exportXlsxId}"><i class="bi bi-file-earmark-excel me-1"></i>Excel</button>
        <button class="btn btn-outline-navy" id="${exportCsvId}"><i class="bi bi-filetype-csv me-1"></i>CSV</button>
      </div>
    </div>
  </div>`;
}

/* ============================================================
   PÁGINA: PANORAMA GENERAL
   ============================================================ */
function renderOverviewPage(container){
  const secsWithData=STATE.sections.filter(s=>s.kpis.pct!==null&&normHeader(s.title)!=='MIPE');
  const avgPct=secsWithData.length?secsWithData.reduce((a,s)=>a+s.kpis.pct,0)/secsWithData.length:null;
  const best=secsWithData.slice().sort((a,b)=>b.kpis.pct-a.kpis.pct)[0];
  const worst=secsWithData.slice().sort((a,b)=>a.kpis.pct-b.kpis.pct)[0];
  const obsList=[];
  STATE.sections.forEach(sec=>{
    if(sec.roles.obsIdx<0) return;
    sec.rows.forEach(r=>{const txt=r[sec.roles.obsIdx];if(txt&&String(txt).trim())obsList.push({area:sec.title,producto:r[0],texto:String(txt).trim()});});
  });
  // También observaciones de la hoja resumen
  if(STATE.resumenKpis){
    STATE.resumenKpis.enriched.forEach(e=>{if(e.obs&&e.obs.trim())obsList.push({area:'Corte',producto:e.label,texto:e.obs.trim()});});
  }
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Reunión de producción · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi bi-speedometer2 text-navy"></i> Panorama general</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Áreas evaluadas',value:`${secsWithData.length} / ${STATE.sections.length}`,sub:'MIPE se evalúa a detalle',icon:'bi-grid-1x2'})}
      ${kpiCard({label:'Cumplimiento promedio',value:avgPct!==null?fmtPct(avgPct):'—',sub:avgPct!==null?(avgPct>=0.95?'en meta':'bajo meta'):'sin datos',icon:'bi-graph-up-arrow',tone:avgPct!==null?pctTone(avgPct):''})}
      ${kpiCard({label:'Mejor desempeño',value:best?displayAreaName(best.title):'—',sub:best?fmtPct(best.kpis.pct):'',icon:'bi-trophy',gold:true})}
      ${kpiCard({label:'A priorizar',value:worst?displayAreaName(worst.title):'—',sub:worst?fmtPct(worst.kpis.pct):'',icon:'bi-exclamation-diamond',tone:'bad'})}
    </div>
    <div class="row g-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart-steps"></i> Ranking por área</h6>
          <div class="chart-sub">% de ejecución vs. plan de cada área</div>
          <div id="overviewRanking"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-pie-chart"></i> Estado general</h6>
          <div class="chart-sub">Semáforo de cumplimiento con respecto al plan finca</div>
          <div id="overviewDonut"></div>
        </div>
      </div>
    </div>
    <div class="chart-card">
      <h6><i class="bi bi-megaphone"></i> Observaciones destacadas</h6>
      <div class="chart-sub">${obsList.length} nota(s) registradas esta semana</div>
      ${obsList.length?`<div class="obs-list">${obsList.slice(0,10).map(o=>`<div class="obs-item"><b>${escapeHtml(displayAreaName(o.area))} — ${escapeHtml(o.producto||'')}:</b> ${escapeHtml(o.texto)}</div>`).join('')}</div>`:`<div class="chart-empty"><i class="bi bi-check2-circle"></i>Sin observaciones registradas esta semana</div>`}
    </div>`;
  wirePageNav();
  // Ranking
  const elR=document.getElementById('overviewRanking');
  if(!secsWithData.length){elR.innerHTML=`<div class="chart-empty"><i class="bi bi-bar-chart"></i>Sin datos</div>`;}
  else{const sorted=secsWithData.slice().sort((a,b)=>b.kpis.pct-a.kpis.pct);elR.innerHTML=sorted.map((s,i)=>{const tone=pctTone(s.kpis.pct);const c=toneColor(tone);const w=Math.min(100,Math.max(3,s.kpis.pct*100));return `<div class="rank-row"><div class="rank-num">${i+1}</div><div style="flex:1"><div class="rank-name">${escapeHtml(displayAreaName(s.title))}</div><div class="rank-bar-track"><div class="rank-bar-fill" style="width:${w}%;background:${c}"></div></div></div><div class="rank-pct" style="color:${c}">${fmtPct(s.kpis.pct)}</div></div>`;}).join('');}
  // Donut
  const elD=document.getElementById('overviewDonut');
  const onTrack=secsWithData.filter(s=>s.kpis.pct>=0.95).length;
  const warn=secsWithData.filter(s=>s.kpis.pct>=0.80&&s.kpis.pct<0.95).length;
  const bad=secsWithData.filter(s=>s.kpis.pct<0.80).length;
  const items=[
    {label:'En meta (≥95%)',value:onTrack,color:'#1E8E5A'},
    {label:'Atención (80–95%)',value:warn,color:'#D98B1E'},
    {label:'Crítico (<80%)',value:bad,color:'#C0392B'},
  ].filter(i=>i.value>0);
  if(!items.length){elD.innerHTML=`<div class="chart-empty"><i class="bi bi-pie-chart"></i>Sin áreas</div>`;}
  else{
    const tot=items.reduce((a,i)=>a+i.value,0);let acc=0;
    const stops=items.map(it=>{const from=acc/tot*100;acc+=it.value;const to=acc/tot*100;return `${it.color} ${from}% ${to}%`;}).join(', ');
    const avgTxt=avgPct!==null?fmtPct(avgPct):'—';
    const legend=items.map(it=>`<div class="legend-row"><span><span class="legend-dot" style="background:${it.color}"></span>${escapeHtml(it.label)}</span><b>${it.value} área${it.value===1?'':'s'}</b></div>`).join('');
    elD.innerHTML=`<div style="display:flex;gap:22px;align-items:center;flex-wrap:wrap;justify-content:center;"><div class="donut-ring" style="background:conic-gradient(${stops})"><div class="donut-hole"><b>${avgTxt}</b><small>promedio</small></div></div><div class="legend-list" style="flex:1;min-width:170px;">${legend}</div></div>`;
  }
}

/* ============================================================
   DESPACHO DE SECCIONES
   ============================================================ */
function renderSectionPage(container,sec){
  const n=normHeader(sec.title);
  if(n==='PRODUCCION') return renderCortePage(container,sec);
  if(n==='MIPE') return renderMipePage(container,sec);
  if(n==='PREPARACION DE CAMAS') return renderPrepCamasPage(container,sec);
  if(n==='SIEMBRAS PRODUCCION') return renderSiembrasProduccionPage(container,sec);
  if(n==='EJECUCION SEMANAL LABORES CULTURALES') return renderLaboresCulturalesPage(container,sec);
  if(n==='SIEMBRAS ENRAIZAMIENTO Y GERMINACION') return renderEnraizamientoPage(container,sec);
  if(n==='EXVITRO') return renderExvitroPage(container,sec);
  if(n==='MANTENIMIENTO') return renderMantenimientoPage(container,sec);
  if(n==='LABORATORIO BIOLOGICOS') return renderLaboratorioPage(container,sec);
  return renderGenericSectionPage(container,sec);
}

/* ── CORTE (fuente: hoja resumen) ── */
function renderCortePage(container,sec){
  const rk=STATE.resumenKpis;
  // Si no hay hoja resumen, fallback genérico
  if(!rk||!rk.validRows.length){ return renderGenericSectionPage(container,sec); }

  const validRows=rk.validRows;
  const tg=rk.totalGeneralRow;
  const hasAprov=rk.roles.aprovIdx>=0&&validRows.some(r=>r.aprov!==null);

  // KPIs generales del total
  const totalReal=tg?tg.total:null;
  const pctPlanFinca=tg?tg.pctPlanFinca:null;
  const pctPer1=tg?tg.pctPer1:null;

  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi bi-basket3 text-navy"></i> Corte <span class="cell-muted" style="font-size:.9rem;font-weight:500;">(Producción)</span></h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Ejecutado semana',value:totalReal!==null?fmtInt(totalReal):'—',sub:'real + pronóstico · tallos',icon:'bi-basket3',gold:true})}
      ${kpiCard({label:'vs. Plan Finca',value:fmtPct(pctPlanFinca),sub:`ejecutado / plan finca${tg&&tg.plan!==null?' · <b>'+fmtInt(tg.plan)+'</b> tallos':''}`,icon:'bi-flag',tone:pctTone(pctPlanFinca)})}
      ${kpiCard({label:'vs. PER1',value:fmtPct(pctPer1),sub:`ejecutado / pronóstico sem. ant.${tg&&tg.per1!==null?' · <b>'+fmtInt(tg.per1)+'</b> tallos':''}`,icon:'bi-graph-up',tone:pctTone(pctPer1)})}
      ${tg?kpiCard({label:'vs. PER4',value:fmtPct(tg.pctPer4),sub:`ejecutado / pronóstico 4 sem.${tg.per4!==null?' · <b>'+fmtInt(tg.per4)+'</b> tallos':''}`,icon:'bi-calendar4-range',tone:pctTone(tg.pctPer4)}):kpiCard({label:'vs. Pronosticador',value:fmtPct(tg?tg.pctPronosticador:null),sub:'ejecutado / proyección siembras',icon:'bi-calculator',tone:pctTone(tg?tg.pctPronosticador:null)})}
    </div>
    <div class="row g-3 mb-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-pie-chart"></i> Cumplimiento sobre el Plan</h6>
          <div class="chart-sub">Share de tallos por producto principal</div>
          <div id="corteDonut"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del reporte de corte</div>
          ${corteObsHTML(rk)}
        </div>
      </div>
    </div>
    <div class="row g-3 mb-3">
      <div class="col-12">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Cumplimiento por producto</h6>
          <div class="chart-sub">Ejecutado vs. PER1, PER4, Pronosticador y Plan Finca${hasAprov?' · % Aprovechamiento 52 semanas cuando disponible':''}</div>
          <div id="corteProductList"></div>
        </div>
      </div>
    </div>
    <div class="table-card mt-3">
      <div class="table-toolbar">
        <h6 class="mb-0"><i class="bi bi-table text-gold"></i> Tabla resumen completa</h6>
        <div class="d-flex gap-2 align-items-center flex-wrap">
          <div class="search-box"><i class="bi bi-search"></i><input type="text" id="corteSearch" placeholder="Buscar…"></div>
          <div class="export-btns d-flex gap-2">
            <button class="btn btn-outline-navy" id="corteExportXlsx"><i class="bi bi-file-earmark-excel me-1"></i>Excel</button>
            <button class="btn btn-outline-navy" id="corteExportCsv"><i class="bi bi-filetype-csv me-1"></i>CSV</button>
          </div>
        </div>
      </div>
      <div id="corteTableWrap"></div>
    </div>
    ${imageGalleryHTML(sec.images)}`;
  wirePageNav();

  // Lista de productos con multi-barra
  const elList=document.getElementById('corteProductList');
  const prodRows=validRows.filter(r=>!r.isTotalGeneral);
  if(!prodRows.length){elList.innerHTML='<div class="chart-empty"><i class="bi bi-bar-chart"></i>Sin datos de productos</div>';}
  else{
    const maxVal=Math.max(1,...prodRows.map(r=>Math.max(r.plan||0,r.total||0,r.per1||0,r.per4||0,r.pronosticador||0)));

    // Líneas: los horizontes de pronóstico (las barras)
    // La referencia fija es el Ejecutado (r.total)
    const LINES=[
      {key:'per1',          label:'PER1',       color:'#1E4D9E', pctFn:(r)=>r.per1&&r.per1>0?r.total/r.per1:null},
      {key:'per4',          label:'PER4',       color:'#9FB6DE', pctFn:(r)=>r.per4&&r.per4>0?r.total/r.per4:null},
      {key:'pronosticador', label:'Pronostic.', color:'#3F7A5C', pctFn:(r)=>r.pronosticador&&r.pronosticador>0?r.total/r.pronosticador:null},
      {key:'plan',          label:'Plan finca', color:'#C9A227', pctFn:(r)=>r.plan&&r.plan>0?r.total/r.plan:null},
    ];

    function varLabel(pct){
      if(pct===null) return '';
      // pct = ejecutado / horizonte
      // >1: se ejecutó MÁS de lo que predijo el horizonte → +
      // <1: se ejecutó MENOS de lo que predijo el horizonte → -
      const v = pct - 1;
      const sign = v>=0 ? '+' : '';
      const abs = Math.abs(v*100).toLocaleString('es-CO',{maximumFractionDigits:1})+'%';
      return `${sign}${v>=0?'':'-'}${abs} ${v>=0 ? 'sobre lo ejecutado' : 'bajo lo ejecutado'}`;
    }

    elList.innerHTML='<div>'+prodRows.map(r=>{
      // Línea vertical en la posición del Ejecutado
      const ejecutadoPos=r.total!==null?(r.total/maxVal)*100:null;
      const ejecutadoLineHTML=ejecutadoPos!==null
        ?`<div class="corte-plan-line" style="left:${ejecutadoPos}%;background:var(--navy);"></div>`:'';

      const unidadTxt=r.unidad?String(r.unidad):'tallos';
      const ejecutadoChip=r.total!==null
        ?`<span class="corte-plan-chip">Ejecutado: <b>${fmtInt(r.total)}</b> ${escapeHtml(unidadTxt)}</span>`:'';

      const linesHTML=LINES.map(line=>{
        const val=r[line.key];
        if(val===null||val===0) return '';
        const w=Math.max(2,(val/maxVal)*100);
        const pct=line.pctFn(r);
        const tone=pctTone(pct);
        const pctTxt=pct!==null?fmtPct(pct):'';
        const variTxt=varLabel(pct);
        return `<div class="corte-bar-row">
          <div class="corte-bar-label">${escapeHtml(line.label)}</div>
          <div class="corte-bar-area">
            <div class="corte-bar-track">
              ${ejecutadoLineHTML}
              <div class="corte-bar-fill" style="width:${w}%;background:${line.color}"></div>
            </div>
          </div>
          <div class="corte-bar-meta">
            <span class="corte-bar-val">${fmtInt(val)}</span>
            <span class="corte-bar-pct ${tone}">${pctTxt}${variTxt?' · '+variTxt:''}</span>
          </div>
        </div>`;
      }).join('');

      return `<div class="corte-product-block">
        <div class="corte-product-head">
          <div class="corte-product-name">
            ${escapeHtml(titleCaseEs(r.label))}
            <span class="unit-tag">${escapeHtml(unidadTxt)}</span>
          </div>
          ${ejecutadoChip}
        </div>
        ${linesHTML}
        ${hasAprov&&r.aprov!==null?`<div style="margin-top:6px;font-size:.74rem;color:var(--muted);">Aprovechamiento 52 sem: <b style="color:var(--navy)">${fmtPct(r.aprov)}</b></div>`:''}
      </div>`;
    }).join('')+'</div>';
  }

  // Dona — solo productos principales
  const DONUT_OMIT=new Set(['CUSHION','DAISY','NOVELTY']);
  const donutRows=prodRows.filter(r=>r.total&&r.total>0&&!DONUT_OMIT.has(normHeader(r.label))).sort((a,b)=>b.total-a.total);
  document.getElementById('corteDonut').innerHTML=donutCSS(donutRows.map(r=>({label:r.label,value:r.total,pct:r.pctPlanFinca})),{centerPct:pctPlanFinca,centerLabel:'cumplimiento'});

  // Tabla
  buildInteractiveTable({
    wrapId:'corteTableWrap',searchId:'corteSearch',exportXlsxId:'corteExportXlsx',exportCsvId:'corteExportCsv',
    headers:STATE.resumen.headers.filter(h=>h),fullHeaders:STATE.resumen.headers,
    rows:STATE.resumen.rows,roles:rk.roles,isResumen:true,filename:'Corte_'+STATE.weekLabel,
  });
}
function corteObsHTML(rk){
  const items=rk.enriched.filter(e=>e.obs&&e.obs.trim()).map(e=>`<div class="obs-item"><b>${escapeHtml(e.label)}:</b> ${escapeHtml(e.obs)}</div>`);
  if(!items.length) return `<div class="chart-empty"><i class="bi bi-check2-circle"></i>Sin observaciones</div>`;
  return `<div class="obs-list">${items.join('')}</div>`;
}

/* ── PREPARACIÓN DE CAMAS ── */
function renderPrepCamasPage(container,sec){
  const OMIT_PREP=['MICROPOMPON','ASTER PURPLE','LEPIDIUM','AMMI'];
  const k=sec.kpis;
  const rows=k.enriched.filter(e=>{
    if(!e.label||e.isSubtotal) return false;
    const n=normHeader(e.label);
    for(const o of OMIT_PREP) if(n.includes(o)) return false;
    return true;
  });
  // KPIs contextualizados: "camas"
  const sumPlan=rows.filter(r=>r.plan!==null).reduce((a,r)=>a+r.plan,0);
  const sumTotal=rows.filter(r=>r.total!==null).reduce((a,r)=>a+r.total,0);
  const hasPlan=rows.some(r=>r.plan!==null);
  const hasTotal=rows.some(r=>r.total!==null);
  const pct=hasPlan&&sumPlan>0?sumTotal/sumPlan:null;
  const diff=hasPlan&&hasTotal?sumTotal-sumPlan:null;
  const icon='bi-grid-3x3-gap';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> Preparación de Camas</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Camas planificadas',value:hasPlan?fmtInt(sumPlan):'—',sub:'total semana',icon:'bi-flag'})}
      ${kpiCard({label:'Camas preparadas',value:hasTotal?fmtInt(sumTotal):'—',sub:'total ejecutado',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(pct),sub:'camas preparadas vs. planificadas',icon:'bi-speedometer',tone:pctTone(pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Camas preparadas por producto</h6>
          <div class="chart-sub">Plan de siembra vs. total preparado y % cumplimiento</div>
          <div id="prepLaborList"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-pie-chart"></i> Cumplimiento sobre el Plan</h6>
          <div class="chart-sub">Share por producto</div>
          <div id="prepDonut"></div>
        </div>
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del agrónomo</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('prepTableWrap','prepSearch','prepExportXlsx','prepExportCsv','Detalle completo')}
      <div id="prepTableWrap"></div>
    </div>`;
  wirePageNav();
  document.getElementById('prepLaborList').innerHTML=laborBarList(rows,{planLabel:'Plan siembra',execLabel:'Total preparado'});
  const donutRows=rows.filter(r=>r.total&&r.total>0).sort((a,b)=>b.total-a.total);
  document.getElementById('prepDonut').innerHTML=donutCSS(donutRows.map(r=>({label:r.label,value:r.total,pct:r.pct})),{centerPct:pct});
  buildInteractiveTable({wrapId:'prepTableWrap',searchId:'prepSearch',exportXlsxId:'prepExportXlsx',exportCsvId:'prepExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'PrepCamas_'+STATE.weekLabel});
}

/* ── SIEMBRAS PRODUCCIÓN ── */
function renderSiembrasProduccionPage(container,sec){
  const k=sec.kpis;
  const rows=k.enriched.filter(e=>{
    if(!e.label||e.isSubtotal) return false;
    return isProductoDetalle(e.label);
  });
  renderStandardSection(container,sec,{
    title:'Siembras Producción',icon:'bi-flower1',
    planLabel:'Plan siembra',execLabel:'Total siembra',
    rows,extraNote:'Solo productos propios de la empresa',
  });
}

/* ── LABORES CULTURALES ── */
function renderLaboresCulturalesPage(container,sec){
  const k=sec.kpis;
  // Separar herbicidas del resto
  const laboresNormales=k.enriched.filter(e=>e.label&&!e.isSubtotal&&!isHerbicida(e.label)&&e.total!==0&&(e.plan!==null||e.total!==null));
  const herbicidas=k.enriched.filter(e=>e.label&&!e.isSubtotal&&isHerbicida(e.label)&&e.total!==null&&e.total>0);

  const icon='bi-scissors';
  const areaName='Labores Culturales';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> ${escapeHtml(areaName)}</h2>
      </div>
      ${pageNavHTML()}
    </div>
    ${sectionSummaryStrip(laboresNormales)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-8">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Labores con plan</h6>
          <div class="chart-sub">Programa vs. ejecutado y % de cumplimiento</div>
          <div id="lcLaborList"></div>
        </div>
      </div>
      <div class="col-lg-4">
        <div class="chart-card">
          <h6><i class="bi bi-droplet"></i> Herbicidas aplicados</h6>
          <div class="chart-sub">Sin plan programado — solo registro de ejecución</div>
          <div id="lcHerbList"></div>
        </div>
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del agrónomo</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('lcTableWrap','lcSearch','lcExportXlsx','lcExportCsv','Detalle completo')}
      <div id="lcTableWrap"></div>
    </div>`;
  wirePageNav();
  document.getElementById('lcLaborList').innerHTML=laborBarList(laboresNormales,{planLabel:'Programa',execLabel:'Ejecutado'});
  // Herbicidas: lista simple con totales
  const elH=document.getElementById('lcHerbList');
  if(!herbicidas.length){elH.innerHTML='<div class="chart-empty"><i class="bi bi-droplet"></i>Sin herbicidas esta semana</div>';}
  else{elH.innerHTML='<div class="labor-list">'+herbicidas.map(r=>`<div class="labor-row"><div class="labor-row-head"><span class="labor-name">${escapeHtml(titleCaseEs(r.label))}</span><span class="badge-soft" style="background:var(--navy-soft);color:var(--navy-3)">Solo registro</span></div><div class="labor-caption">Ejecutado: <b>${fmtInt(r.total)}</b></div></div>`).join('')+'</div>';}
  buildInteractiveTable({wrapId:'lcTableWrap',searchId:'lcSearch',exportXlsxId:'lcExportXlsx',exportCsvId:'lcExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'LaboresCulturales_'+STATE.weekLabel});
}

/* ── ENRAIZAMIENTO Y GERMINACIÓN ── */
function renderEnraizamientoPage(container,sec){
  const k=sec.kpis;
  const rows=k.enriched.filter(e=>e.label&&!e.isSubtotal&&(e.plan!==null||e.total!==null));
  const icon='bi-tree';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> Enraizamiento y Germinación</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Unidades planificadas',value:k.plan!==null?fmtInt(k.plan):'—',sub:'plan de siembra · total semana',icon:'bi-flag'})}
      ${kpiCard({label:'Unidades ejecutadas',value:k.total!==null?fmtInt(k.total):'—',sub:'total siembra/cosecha',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(k.pct),sub:'ejecutado vs. planificado',icon:'bi-speedometer',tone:pctTone(k.pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Labores por producto</h6>
          <div class="chart-sub">Plan vs. ejecutado y % cumplimiento</div>
          <div id="enrLaborList"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del agrónomo</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    ${imageGalleryHTML(sec.images)}
    <div class="table-card">
      ${tableToolbar('enrTableWrap','enrSearch','enrExportXlsx','enrExportCsv','Detalle completo')}
      <div id="enrTableWrap"></div>
    </div>`;
  wirePageNav();
  document.getElementById('enrLaborList').innerHTML=laborBarList(rows,{planLabel:'Plan siembra',execLabel:'Total siembra'});
  buildInteractiveTable({wrapId:'enrTableWrap',searchId:'enrSearch',exportXlsxId:'enrExportXlsx',exportCsvId:'enrExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'Enraizamiento_'+STATE.weekLabel});
}

/* ── EXVITRO ── */
function renderExvitroPage(container,sec){
  const k=sec.kpis;
  const rows=k.enriched.filter(e=>e.label&&!e.isSubtotal&&(e.plan!==null||e.total!==null));
  const icon='bi-droplet-half';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> Exvitro</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Tallos estimados',value:k.plan!==null?fmtInt(k.plan):'—',sub:'total semana',icon:'bi-flag'})}
      ${kpiCard({label:'Tallos ejecutados',value:k.total!==null?fmtInt(k.total):'—',sub:'total ejecutado',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(k.pct),sub:'ejecutado vs. estimado',icon:'bi-speedometer',tone:pctTone(k.pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Labores Exvitro</h6>
          <div class="chart-sub">Estimado vs. ejecutado. Poda se muestra como dato adicional cuando aplica.</div>
          <div id="exLaborList"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del reporte</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('exTableWrap','exSearch','exExportXlsx','exExportCsv','Detalle completo')}
      <div id="exTableWrap"></div>
    </div>`;
  wirePageNav();
  // Para exvitro: mostrar poda si existe
  const elL=document.getElementById('exLaborList');
  if(!rows.length){elL.innerHTML='<div class="chart-empty"><i class="bi bi-bar-chart"></i>Sin datos</div>';}
  else{
    elL.innerHTML='<div class="labor-list">'+rows.map(r=>{
      const tone=pctTone(r.pct);const c=toneColor(tone);
      const w=r.pct!==null?Math.min(100,Math.max(r.pct*100,r.pct>0?2:0)):0;
      const podaStr=r.poda!==null?` &nbsp;·&nbsp; Poda: <b>${fmtInt(r.poda)}</b>`:'';
      return `<div class="labor-row">
        <div class="labor-row-head"><span class="labor-name">${escapeHtml(titleCaseEs(r.label))}</span><span class="badge-soft" style="background:${c}22;color:${c}">${fmtPct(r.pct)}</span></div>
        <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${c}"></div></div>
        <div class="labor-caption">Estimado: <b>${r.plan!==null?fmtInt(r.plan):'—'}</b> &nbsp;·&nbsp; Ejecutado: <b>${r.total!==null?fmtInt(r.total):'—'}</b>${podaStr}${r.diff!==null?' &nbsp;·&nbsp; Desfase: <b style="color:'+(r.diff>=0?'var(--success)':'var(--danger)')+'">'+( r.diff>=0?'+':'')+fmtInt(r.diff)+'</b>':''}</div>
      </div>`;
    }).join('')+'</div>';
  }
  buildInteractiveTable({wrapId:'exTableWrap',searchId:'exSearch',exportXlsxId:'exExportXlsx',exportCsvId:'exExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'Exvitro_'+STATE.weekLabel});
}

/* ── MIPE ── */
function renderMipePage(container,sec){
  const enr=sec.kpis.enriched;
  // Agrupar variantes RB y calcular % combinado en código
  const rbVariants=[];
  let rbPlanSum=0,rbTotalSum=0;
  const others=[];
  for(let i=0;i<enr.length;i++){
    const row=enr[i];
    if(!row.label) continue;
    if(/^MONITOREO RB/i.test(row.label)){
      const camasRow=(enr[i+1]&&!enr[i+1].label)?enr[i+1]:null;
      const plan=camasRow?camasRow.plan:null;
      const total=camasRow?camasRow.total:null;
      if(plan!==null) rbPlanSum+=plan;
      if(total!==null) rbTotalSum+=total;
      const perDay=sec.roles.dayIdxs.map(idx=>{
        const bVal=row.raw[idx];const cVal=camasRow?camasRow.raw[idx]:null;
        const hasB=bVal!=null&&String(bVal).trim()!=='';const hasC=cVal!=null&&String(cVal).trim()!=='';
        if(!hasB&&!hasC) return null;
        return{day:dayDisplay(sec.headers[idx]),bloques:hasB?String(bVal).trim():null,camas:hasC?toNum(cVal):null};
      }).filter(Boolean);
      rbVariants.push({name:row.label.replace(/^MONITOREO RB\s*/i,'').replace(/[()]/g,'').trim()||row.label,plan,total,perDay});
    } else { others.push(row); }
  }
  const rbPct=rbPlanSum>0?rbTotalSum/rbPlanSum:null;

  const viroide=others.find(r=>normHeader(r.label)==='VIROIDE');
  const directo=others.find(r=>normHeader(r.label)==='MONITOREO DIRECTO');
  const aspirado=others.find(r=>normHeader(r.label)==='ASPIRADO');
  const biologicos=others.find(r=>normHeader(r.label)==='BIOLOGICOS');
  const biologicosObs=(sec.roles.obsIdx>=0&&biologicos)?biologicos.raw[sec.roles.obsIdx]:null;

  // Viroide: determinar si hubo inspección
  const viroideHecho=viroide&&viroide.total!==null&&viroide.total>0;
  const viroideTexto=!viroide?'Sin datos':viroideHecho?`${fmtPct(viroide.pct)} (${fmtNum1(viroide.total)} camas)`:'No se realizó inspección esta semana';

  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi bi-bug text-navy"></i> MIPE <span class="cell-muted" style="font-size:.9rem;font-weight:500;">(Manejo Integrado de Plagas y Enfermedades)</span></h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Monitoreo RB combinado',value:fmtPct(rbPct),sub:`${fmtNum1(rbTotalSum)} / ${fmtInt(rbPlanSum)} camas · Grupo+Exvitro+Variedades`,icon:'bi-clipboard-check',tone:pctTone(rbPct),gold:true})}
      ${kpiCard({label:'Monitoreo directo',value:directo?fmtPct(directo.pct):'—',sub:'cobertura en planta',icon:'bi-eye',tone:directo?pctTone(directo.pct):''})}
      ${kpiCard({label:'Biológicos',value:biologicos&&biologicos.total!==null?fmtInt(biologicos.total)+' válvulas':'—',sub:biologicos&&biologicos.plan!==null?`plan: ${fmtInt(biologicos.plan)}`:'',icon:'bi-droplet',tone:biologicos?pctTone(biologicos.pct):''})}
      ${kpiCard({label:'Viroide (CSVd)',value:viroideTexto,sub:!viroideHecho&&biologicosObs&&biologicosObs.toLowerCase().includes('csvd')?'⚠ Revisión puntual en biológicos':'',icon:'bi-shield-exclamation',tone:viroideHecho?pctTone(viroide.pct):viroide?'warn':''})}
    </div>
    ${biologicosObs?`<div class="insight-box"><div class="ib-icon"><i class="bi bi-flag"></i></div><div><b>Fitosanidad:</b> ${escapeHtml(String(biologicosObs).trim())}</div></div>`:''}
    <div class="chart-card">
      <h6><i class="bi bi-clipboard-data"></i> Monitoreo RB — cumplimiento combinado</h6>
      <div class="chart-sub">% calculado en base a: Plan ${fmtInt(rbPlanSum)} camas · Ejecutado ${fmtNum1(rbTotalSum)} camas</div>
      <div id="mipeRbSummary"></div>
    </div>
    <div class="chart-card">
      <h6><i class="bi bi-geo-alt"></i> Bloques y camas monitoreados por día</h6>
      <div class="chart-sub">Por variante: bloques visitados y camas monitoreadas</div>
      <div id="mipeBlocks"></div>
    </div>
    <div class="chart-card">
      <h6><i class="bi bi-list-check"></i> Otras líneas MIPE</h6>
      <div class="chart-sub">Viroide, monitoreo directo, aspirado y biológicos</div>
      <div id="mipeOthers"></div>
    </div>
    ${imageGalleryHTML(sec.images)}
    <div class="table-card">
      ${tableToolbar('mipeTableWrap','mipeSearch','mipeExportXlsx','mipeExportCsv','Detalle completo')}
      <div id="mipeTableWrap"></div>
    </div>`;
  wirePageNav();

  // RB Summary
  const rbTone=pctTone(rbPct);const rbColor=toneColor(rbTone);const rbW=rbPct!==null?Math.min(100,Math.max(rbPct*100,2)):0;
  document.getElementById('mipeRbSummary').innerHTML=`
    <div class="labor-row">
      <div class="labor-row-head"><span class="labor-name">Monitoreo RB — Grupo, Exvitro y Variedades (combinado)</span><span class="badge-soft" style="background:${rbColor}22;color:${rbColor}">${fmtPct(rbPct)}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${rbW}%;background:${rbColor}"></div></div>
      <div class="labor-caption">Camas monitoreadas: <b>${fmtNum1(rbTotalSum)}</b> · Camas planificadas: <b>${fmtInt(rbPlanSum)}</b></div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-top:16px;">
      ${rbVariants.map(v=>`<div style="background:var(--bg);border-radius:12px;padding:12px 14px;">
        <div style="font-weight:700;font-size:.8rem;color:var(--navy);margin-bottom:4px;">${escapeHtml(v.name)}</div>
        <div style="font-size:.78rem;color:var(--muted);">Camas: <b style="color:var(--navy);">${v.total!==null?fmtNum1(v.total):'—'}</b> / ${v.plan!==null?fmtInt(v.plan):'—'}</div>
      </div>`).join('')}
    </div>`;

  // Bloques por día
  document.getElementById('mipeBlocks').innerHTML='<div class="labor-list">'+rbVariants.map(v=>`
    <div class="labor-row">
      <div class="labor-row-head"><span class="labor-name">${escapeHtml(v.name)}</span></div>
      ${v.perDay.length?v.perDay.map(d=>`<div class="block-day"><span class="day-label">${escapeHtml(d.day)}</span><span>${d.bloques?d.bloques.split(/\s+/).map(b=>`<span class="block-chip">${escapeHtml(b)}</span>`).join(''):''}${d.camas!==null?`<span class="camas-chip">${fmtNum1(d.camas)} camas</span>`:''}</span></div>`).join(''):'<div class="cell-muted" style="font-size:.8rem;">Sin bloques esta semana</div>'}
    </div>`).join('')+'</div>';

  // Otras líneas
  document.getElementById('mipeOthers').innerHTML='<div class="labor-list">'+others.map(r=>{
    const tone=pctTone(r.pct);const c=toneColor(tone);
    const noRealizado=r.total===0&&r.plan!==null&&r.plan>0;
    const w=r.pct!==null?Math.min(100,Math.max(r.pct*100,r.pct>0?2:0)):0;
    if(noRealizado){
      return `<div class="labor-row"><div class="labor-row-head"><span class="labor-name">${escapeHtml(titleCaseEs(r.label))}</span><span class="badge-soft" style="background:var(--bg);color:var(--muted)">No realizado</span></div><div class="labor-caption">No se realizó ${escapeHtml(titleCaseEs(r.label).toLowerCase())} esta semana</div></div>`;
    }
    return `<div class="labor-row">
      <div class="labor-row-head"><span class="labor-name">${escapeHtml(titleCaseEs(r.label))}</span><span class="badge-soft" style="background:${c}22;color:${c}">${fmtPct(r.pct)}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${c}"></div></div>
      <div class="labor-caption">Plan: <b>${r.plan!==null?fmtInt(r.plan):'—'}</b> &nbsp;·&nbsp; Ejecutado: <b>${r.total!==null?fmtInt(r.total):'—'}</b></div>
    </div>`;
  }).join('')+'</div>';

  buildInteractiveTable({wrapId:'mipeTableWrap',searchId:'mipeSearch',exportXlsxId:'mipeExportXlsx',exportCsvId:'mipeExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'MIPE_'+STATE.weekLabel});
}

/* ── MANTENIMIENTO ── */
function renderMantenimientoPage(container,sec){
  const k=sec.kpis;
  // Solo las 3 filas válidas, inferir % en código
  const rows=k.enriched.filter(e=>e.label&&!e.isSubtotal&&isMantenimientoValida(e.label)).map(e=>{
    let pct=e.pct;
    if(pct===null&&e.plan!==null&&e.plan>0&&e.total!==null) pct=e.total/e.plan;
    const diff=e.total!==null&&e.plan!==null?e.total-e.plan:e.diff;
    return{...e,pct,diff};
  });
  // Calcular KPIs del área sumando las filas válidas
  const sumPlan=rows.filter(r=>r.plan!==null).reduce((a,r)=>a+r.plan,0);
  const sumTotal=rows.filter(r=>r.total!==null).reduce((a,r)=>a+r.total,0);
  const hasPlan=rows.some(r=>r.plan!==null);
  const hasTotal=rows.some(r=>r.total!==null);
  const pct=hasPlan&&sumPlan>0?sumTotal/sumPlan:null;
  const diff=hasPlan&&hasTotal?sumTotal-sumPlan:null;
  const icon='bi-tools';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> Mantenimiento</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Bloques programados',value:hasPlan?fmtInt(sumPlan):'—',sub:'total semana',icon:'bi-flag'})}
      ${kpiCard({label:'Bloques ejecutados',value:hasTotal?fmtInt(sumTotal):'—',sub:'total ejecutado',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(pct),sub:'ejecutado vs. programado',icon:'bi-speedometer',tone:pctTone(pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    <div class="row g-3">
      <div class="col-lg-8">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Labores de mantenimiento</h6>
          <div class="chart-sub">Programado vs. ejecutado — % inferido como ejecutado / programado</div>
          <div id="mantLaborList"></div>
        </div>
      </div>
      <div class="col-lg-4">
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del reporte</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('mantTableWrap',null,'mantExportXlsx','mantExportCsv','Detalle completo')}
      <div id="mantTableWrap"></div>
    </div>`;
  wirePageNav();
  document.getElementById('mantLaborList').innerHTML=laborBarList(rows,{planLabel:'Programado',execLabel:'Ejecutado'});
  buildInteractiveTable({wrapId:'mantTableWrap',searchId:null,exportXlsxId:'mantExportXlsx',exportCsvId:'mantExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'Mantenimiento_'+STATE.weekLabel});
}

/* ── LABORATORIO BIOLÓGICOS ── */
function renderLaboratorioPage(container,sec){
  const k=sec.kpis;
  // Detectar % de contaminación en encabezado (última columna suele tenerlo)
  let contaminacion=null;
  sec.headers.forEach(h=>{
    if(h&&/CONTAMINAC/i.test(String(h))){
      const m=String(h).match(/(\d+[\.,]?\d*)%/);
      if(m) contaminacion=m[0];
    }
  });
  // Inferir % en todas las filas
  const rows=k.enriched.filter(e=>e.label&&!e.isSubtotal&&(e.plan!==null||e.total!==null)).map(e=>{
    let pct=e.pct;
    if(pct===null&&e.plan!==null&&e.plan>0&&e.total!==null) pct=e.total/e.plan;
    const diff=e.total!==null&&e.plan!==null?e.total-e.plan:e.diff;
    return{...e,pct,diff};
  });
  const icon='bi-eyedropper';
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> Laboratorio Biológicos</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:'Bolsas estimadas',value:k.plan!==null?fmtInt(k.plan):'—',sub:'total semana',icon:'bi-flag'})}
      ${kpiCard({label:'Bolsas ejecutadas',value:k.total!==null?fmtInt(k.total):'—',sub:'total ejecutado',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(k.pct),sub:'ejecutado vs. estimado',icon:'bi-speedometer',tone:pctTone(k.pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    <div class="row g-3">
      <div class="col-lg-8">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Labores del laboratorio</h6>
          <div class="chart-sub">Estimado vs. total ejecutado — % inferido como Total / Estimado</div>
          <div id="labLaborList"></div>
        </div>
      </div>
      <div class="col-lg-4">
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del reporte</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('labTableWrap',null,'labExportXlsx','labExportCsv','Detalle completo')}
      <div id="labTableWrap"></div>
    </div>`;
  wirePageNav();
  document.getElementById('labLaborList').innerHTML=laborBarList(rows,{planLabel:'Estimado',execLabel:'Total ejecutado'});
  buildInteractiveTable({wrapId:'labTableWrap',searchId:null,exportXlsxId:'labExportXlsx',exportCsvId:'labExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:'Laboratorio_'+STATE.weekLabel});
}

/* ── SECCIÓN ESTÁNDAR REUTILIZABLE ── */
// Unidades por área para contextualizar KPIs
const AREA_UNITS={
  'SIEMBRAS PRODUCCION':'esquejes',
  'SIEMBRAS ENRAIZAMIENTO Y GERMINACION':'unidades',
  'EXVITRO':'tallos/unidades',
  'LABORATORIO BIOLOGICOS':'bolsas',
  'MANTENIMIENTO':'bloques',
  'EJECUCION SEMANAL LABORES CULTURALES':'labores',
};
function unitForArea(title){ return AREA_UNITS[normHeader(title)]||'unidades'; }

function renderStandardSection(container,sec,opts){
  const{title,icon,planLabel,execLabel,rows,extraNote}=opts;
  const k=sec.kpis;
  const unit=opts.unit||unitForArea(sec.title);
  const withTotal=rows.filter(r=>r.total&&r.total>0).sort((a,b)=>b.total-a.total);
  container.innerHTML=`
    <div class="page-header">
      <div>
        <div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div>
        <h2><i class="bi ${icon} text-navy"></i> ${escapeHtml(title)}</h2>
      </div>
      ${pageNavHTML()}
    </div>
    <div class="kpi-row">
      ${kpiCard({label:`${titleCaseEs(unit)} planificad${unit.endsWith('s')?'os':'o'}`,value:k.plan!==null?fmtInt(k.plan):'—',sub:`${escapeHtml(planLabel)} · total semana`,icon:'bi-flag'})}
      ${kpiCard({label:`${titleCaseEs(unit)} ejecutad${unit.endsWith('s')?'os':'o'}`,value:k.total!==null?fmtInt(k.total):'—',sub:`${escapeHtml(execLabel)} · total semana`,icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(k.pct),sub:`${escapeHtml(execLabel.toLowerCase())} vs. ${escapeHtml(planLabel.toLowerCase())}`,icon:'bi-speedometer',tone:pctTone(k.pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-7">
        <div class="chart-card">
          <h6><i class="bi bi-bar-chart"></i> Indicadores por labor / producto</h6>
          <div class="chart-sub">${escapeHtml(planLabel)} vs. ${escapeHtml(execLabel.toLowerCase())} y % cumplimiento</div>
          <div id="stdLaborList"></div>
        </div>
      </div>
      <div class="col-lg-5">
        <div class="chart-card">
          <h6><i class="bi bi-pie-chart"></i> Cumplimiento sobre el Plan</h6>
          <div class="chart-sub">Share de ejecución por labor / producto</div>
          <div id="stdDonut"></div>
        </div>
        <div class="chart-card">
          <h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>
          <div class="chart-sub">Notas del agrónomo</div>
          ${sectionObsHTML(sec)}
        </div>
      </div>
    </div>
    <div class="table-card">
      ${tableToolbar('stdTableWrap','stdSearch','stdExportXlsx','stdExportCsv','Detalle completo')}
      <div id="stdTableWrap"></div>
    </div>
    ${imageGalleryHTML(sec.images)}`;
  wirePageNav();
  document.getElementById('stdLaborList').innerHTML=laborBarList(rows,{planLabel,execLabel});
  document.getElementById('stdDonut').innerHTML=donutCSS(withTotal.map(r=>({label:r.label,value:r.total,pct:r.pct})),{centerPct:k.pct});
  buildInteractiveTable({wrapId:'stdTableWrap',searchId:'stdSearch',exportXlsxId:'stdExportXlsx',exportCsvId:'stdExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:title.replace(/\s+/g,'_')+'_'+STATE.weekLabel});
}

/* ── SECCIÓN GENÉRICA (fallback para secciones nuevas) ── */
function renderGenericSectionPage(container,sec){
  const k=sec.kpis;
  const icon=iconFor(sec.title);
  const areaName=displayAreaName(sec.title);
  if(!k.hasData){
    container.innerHTML=`
      <div class="page-header"><div><div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div><h2><i class="bi ${icon} text-navy"></i> ${escapeHtml(areaName)}</h2></div>${pageNavHTML()}</div>
      <div class="section-empty"><i class="bi bi-inbox"></i>Esta sección no tiene datos numéricos diligenciados esta semana.</div>
      <div class="table-card mt-3"><div class="table-toolbar"><h6 class="mb-0"><i class="bi bi-table text-gold"></i> Datos de la sección</h6></div><div id="secTableWrap"></div></div>`;
    wirePageNav();
    buildInteractiveTable({wrapId:'secTableWrap',searchId:null,exportXlsxId:null,exportCsvId:null,headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:sec.title});
    return;
  }
  const rows=k.enriched.filter(e=>e.label&&!e.isSubtotal);
  const withTotal=rows.filter(r=>r.total&&r.total>0).sort((a,b)=>b.total-a.total);
  container.innerHTML=`
    <div class="page-header"><div><div class="eyebrow">Área · ${escapeHtml(weekLabelDisplay())}</div><h2><i class="bi ${icon} text-navy"></i> ${escapeHtml(areaName)}</h2></div>${pageNavHTML()}</div>
    <div class="kpi-row">
      ${kpiCard({label:'Plan semana',value:k.plan!==null?fmtInt(k.plan):'—',sub:'total planificado',icon:'bi-flag'})}
      ${kpiCard({label:'Total ejecutado',value:k.total!==null?fmtInt(k.total):'—',sub:'total ejecutado',icon:'bi-check2-circle',gold:true})}
      ${kpiCard({label:'% cumplimiento',value:fmtPct(k.pct),sub:'ejecutado vs. plan',icon:'bi-speedometer',tone:pctTone(k.pct)})}
    </div>
    ${sectionSummaryStrip(rows)}
    ${dayInsightHTML(sec)}
    <div class="row g-3">
      <div class="col-lg-7"><div class="chart-card"><h6><i class="bi bi-bar-chart"></i> Indicadores</h6><div id="secLaborList"></div></div></div>
      <div class="col-lg-5">
        <div class="chart-card"><h6><i class="bi bi-pie-chart"></i> Cumplimiento sobre el Plan</h6><div id="secDonut"></div></div>
        <div class="chart-card"><h6><i class="bi bi-chat-square-text"></i> Observaciones</h6>${sectionObsHTML(sec)}</div>
      </div>
    </div>
    <div class="table-card">${tableToolbar('secTableWrap','secSearch','secExportXlsx','secExportCsv','Detalle completo')}<div id="secTableWrap"></div></div>
    ${imageGalleryHTML(sec.images)}`;
  wirePageNav();
  document.getElementById('secLaborList').innerHTML=laborBarList(rows);
  document.getElementById('secDonut').innerHTML=donutCSS(withTotal.map(r=>({label:r.label,value:r.total,pct:r.pct})),{centerPct:k.pct});
  buildInteractiveTable({wrapId:'secTableWrap',searchId:'secSearch',exportXlsxId:'secExportXlsx',exportCsvId:'secExportCsv',headers:sec.headers.filter(h=>h),fullHeaders:sec.headers,rows:sec.rows,roles:sec.roles,filename:sec.title});
}

/* ── EXPORTAR RESUMEN GENERAL ── */
function exportOverviewSummary(){
  const rows=STATE.sections.map(s=>([displayAreaName(s.title),s.kpis.plan,s.kpis.total,s.kpis.pct!==null?+(s.kpis.pct*100).toFixed(1):null,s.kpis.diff]));
  const aoa=[['Área','Plan','Ejecutado','% Cumplimiento','Desviación'],...rows];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Panorama general');
  XLSX.writeFile(wb,`Panorama_${STATE.weekLabel.replace(/\s+/g,'_')}.xlsx`);
}

/* ── UI HANDLERS ── */
function showLoading(text){ document.getElementById('loadingText').textContent=text; document.getElementById('loadingOverlay').classList.add('show'); }
function hideLoading(){ document.getElementById('loadingOverlay').classList.remove('show'); }
function setUploadProgress(pct,text){
  document.getElementById('uploadProgress').classList.add('show');
  document.getElementById('progressBar').style.width=pct+'%';
  document.getElementById('uploadStatusText').textContent=text;
}
function showUploadError(msg){ const box=document.getElementById('uploadError'); document.getElementById('uploadErrorText').textContent=msg; box.classList.add('show'); document.getElementById('uploadProgress').classList.remove('show'); }
function clearUploadError(){ document.getElementById('uploadError').classList.remove('show'); }

function initUploadHandlers(){
  const dz=document.getElementById('dropzone');
  const input=document.getElementById('fileInput');
  dz.addEventListener('click',()=>input.click());
  input.addEventListener('change',()=>{if(input.files[0]){clearUploadError();handleFile(input.files[0]);}});
  ['dragover','dragenter'].forEach(evt=>dz.addEventListener(evt,(e)=>{e.preventDefault();dz.classList.add('dragover');}));
  ['dragleave','drop'].forEach(evt=>dz.addEventListener(evt,(e)=>{e.preventDefault();dz.classList.remove('dragover');}));
  dz.addEventListener('drop',(e)=>{const f=e.dataTransfer.files[0];if(f){clearUploadError();handleFile(f);}else{showUploadError('No se detectó ningún archivo.');}});
  document.getElementById('newFileBtn').addEventListener('click',()=>{
    document.getElementById('appShell').classList.remove('show');
    document.getElementById('uploadScreen').style.display='flex';
    document.getElementById('uploadProgress').classList.remove('show');
    document.getElementById('progressBar').style.width='0%';
    input.value='';
  });
  document.getElementById('sidebarToggle').addEventListener('click',()=>document.getElementById('sidebar').classList.toggle('open'));
  document.getElementById('topExportBtn').addEventListener('click',exportOverviewSummary);
}

initUploadHandlers();

/* ── PWA: registro del Service Worker ── */
if('serviceWorker' in navigator){
  window.addEventListener('load',()=>{
    navigator.serviceWorker.register('sw.js').catch(err=>console.warn('SW no registrado:',err));
  });
}
