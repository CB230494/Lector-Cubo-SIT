/* Procesamiento pesado del Excel en un Web Worker para mantener la interfaz fluida. */
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

let records = [];
let activeIndices = [];
let loadedFileName = '';

const normalizeHeader = (v='') => String(v ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');
const normalizeText = (v='') => String(v ?? '').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
const clean = (v) => v === null || v === undefined ? '' : String(v).trim();

function normalizeSubtypeDisplay(value){
  const raw=clean(value);
  if (normalizeText(raw).replace(/\s+/g,' ').trim()==='RINA') return 'RIÑA';
  return raw;
}

const aliases = {
  id: ['idincidente'],
  origen: ['origen'],
  tipo: ['tipoincidente'],
  subtipo: ['subtipoincidente'],
  diligencia: ['diligenciapolicial'],
  provincia: ['provincia'],
  canton: ['canton'],
  distrito: ['distrito'],
  barrio: ['barrio'],
  incidente: ['incidente'],
  direccion: ['direccion'],
  tipoLugar: ['tipolugar'],
  fechaIngreso: ['fechaingreso'],
  hora: ['hora'],
  fechaFinalizado: ['fechafinalizado'],
  indicativo: ['indicativo'],
  unidad: ['unidad'],
  unidadFormal: ['unidadformal'],
  diferenciaMin: ['diferenciaenminutos'],
  estado: ['estadoincidente'],
  fechaCreado: ['fechacreado'],
  lat: ['latitud'],
  lng: ['longitud'],
  horaIncidente: ['horaincidente'],
  rango: ['rango'],
  anio: ['anio'],
  mes: ['mes'],
  dia: ['dia'],
  numSemana: ['numsemana'],
  fuente: ['fuente'],
  codSubUnidad: ['codsubunidad'],
  producto: ['producto']
};

function parseDate(value){
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0,10);
  if (typeof value === 'number'){
    const d = XLSX.SSF.parse_date_code(value);
    if (d && d.y) return `${String(d.y).padStart(4,'0')}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = clean(value);
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  const d = new Date(s);
  if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  return '';
}

function parseHour(value){
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value)) return value.getHours();
  if (typeof value === 'number'){
    const fraction = value >= 1 ? value % 1 : value;
    return Math.max(0, Math.min(23, Math.floor((fraction * 24) + 1e-8)));
  }
  const s = clean(value);
  const m = s.match(/(?:^|\s)([01]?\d|2[0-3]):[0-5]\d/);
  if (m) return Number(m[1]);
  const h = Number(s);
  if (Number.isFinite(h) && h >= 0 && h <= 23) return Math.floor(h);
  return null;
}

function displayTime(value, fallbackHour=null){
  if (value === null || value === undefined || value === '') return fallbackHour === null ? '' : `${String(fallbackHour).padStart(2,'0')}:00`;
  if (typeof value === 'number'){
    const fraction = value >= 1 ? value % 1 : value;
    const totalMinutes = Math.round(fraction * 24 * 60) % (24*60);
    const h = Math.floor(totalMinutes/60), m = totalMinutes%60;
    return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  }
  const s = clean(value);
  const m = s.match(/([01]?\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?/);
  return m ? `${String(m[1]).padStart(2,'0')}:${m[2]}` : s;
}

function parseCoord(v){
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v === null || v === undefined || v === '') return null;
  const n = Number(String(v).trim().replace(',','.'));
  return Number.isFinite(n) ? n : null;
}


const CR_BOUNDS={south:8.0,north:11.3,west:-86.2,east:-82.35};
const PROVINCE_RULES={
  'SAN JOSE':{lat:9.60,lng:-84.08,maxKm:105},
  'ALAJUELA':{lat:10.35,lng:-84.55,maxKm:190},
  'CARTAGO':{lat:9.75,lng:-83.75,maxKm:105},
  'HEREDIA':{lat:10.15,lng:-84.05,maxKm:110},
  'GUANACASTE':{lat:10.60,lng:-85.35,maxKm:190},
  'PUNTARENAS':{lat:9.45,lng:-84.75,maxKm:360},
  'LIMON':{lat:10.05,lng:-83.20,maxKm:205}
};
function haversineKm(lat1,lng1,lat2,lng2){
  const R=6371, rad=x=>x*Math.PI/180;
  const dLat=rad(lat2-lat1), dLng=rad(lng2-lng1);
  const a=Math.sin(dLat/2)**2+Math.cos(rad(lat1))*Math.cos(rad(lat2))*Math.sin(dLng/2)**2;
  return 2*R*Math.asin(Math.sqrt(a));
}
function validateGeo(lat,lng,provincia){
  if(lat===null || lng===null) return {status:'invalid',label:'Inválida / sin coordenada',reason:'Falta latitud o longitud, o el valor no es numérico.'};
  if(lat<-90 || lat>90 || lng<-180 || lng>180) return {status:'invalid',label:'Inválida',reason:'La coordenada excede los rangos mundiales permitidos.'};
  const looksSwapped=(lng>=8 && lng<=12 && lat<=-82 && lat>=-87);
  if(looksSwapped) return {status:'suspect',label:'Sospechosa',reason:'Posible inversión entre latitud y longitud.'};
  const inCR=lat>=CR_BOUNDS.south && lat<=CR_BOUNDS.north && lng>=CR_BOUNDS.west && lng<=CR_BOUNDS.east;
  if(!inCR) return {status:'suspect',label:'Sospechosa',reason:'Es numéricamente válida, pero cae fuera del área esperada de Costa Rica.'};
  const prov=normalizeText(provincia).replace(/\s+/g,' ').trim();
  const rule=PROVINCE_RULES[prov];
  if(rule){
    const d=haversineKm(lat,lng,rule.lat,rule.lng);
    if(d>rule.maxKm) return {status:'suspect',label:'Sospechosa',reason:`Posible inconsistencia con la provincia indicada (${provincia}); distancia territorial aproximada atípica.`};
  }
  return {status:'valid',label:'Válida en Costa Rica',reason:'Formato y ubicación general compatibles con Costa Rica.'};
}

function isNoUbicado(tipo){
  const n = normalizeText(tipo).replace(/\s+/g,' ');
  return n === 'NO UBICADO / OTROS' || n === 'NO UBICADO/OTROS' || (/NO UBICAD[OA]/.test(n) && /OTROS/.test(n));
}

function getHeaderMap(sheet, range){
  const byNorm = new Map();
  for (let c=range.s.c; c<=range.e.c; c++){
    const cell = sheet[XLSX.utils.encode_cell({r:range.s.r,c})];
    const header = cell ? clean(cell.v) : '';
    if (header) byNorm.set(normalizeHeader(header), c);
  }
  const map = {};
  for (const [key, list] of Object.entries(aliases)){
    let idx = null;
    for (const alias of list){ if (byNorm.has(alias)){ idx = byNorm.get(alias); break; } }
    map[key] = idx;
  }
  return { map, rawHeaders:[...byNorm.keys()] };
}

function cellValue(sheet,row,col){
  if (col === null || col === undefined) return '';
  const cell = sheet[XLSX.utils.encode_cell({r:row,c:col})];
  return cell ? cell.v : '';
}

function uniqOptions(){
  const sets = {unidad:new Set(), unidadFormal:new Set(), provincia:new Set(), canton:new Set(), distrito:new Set(), barrio:new Set(), tipo:new Set(), origen:new Set(), subtipo:new Set()};
  let minDate='', maxDate='';
  for (const r of records){
    for (const key of Object.keys(sets)) if (r[key]) sets[key].add(r[key]);
    if (r.fechaIngreso){ if (!minDate || r.fechaIngreso < minDate) minDate=r.fechaIngreso; if (!maxDate || r.fechaIngreso > maxDate) maxDate=r.fechaIngreso; }
  }
  const collator = new Intl.Collator('es',{sensitivity:'base',numeric:true});
  const out={};
  for (const [k,set] of Object.entries(sets)) out[k]=Array.from(set).sort(collator.compare);
  out.minDate=minDate; out.maxDate=maxDate;
  return out;
}

function matches(r,f){
  if (f.unidad && r.unidad !== f.unidad) return false;
  if (f.unidadFormal && r.unidadFormal !== f.unidadFormal) return false;
  if (f.provincia && r.provincia !== f.provincia) return false;
  if (f.canton && r.canton !== f.canton) return false;
  if (f.distrito && r.distrito !== f.distrito) return false;
  if (f.barrio && !normalizeText(r.barrio).includes(normalizeText(f.barrio))) return false;
  if (f.tipo && r.tipo !== f.tipo) return false;
  if (f.origen && r.origen !== f.origen) return false;
  if (f.subtipo && r.subtipo !== f.subtipo) return false;
  if (f.fechaDesde && (!r.fechaIngreso || r.fechaIngreso < f.fechaDesde)) return false;
  if (f.fechaHasta && (!r.fechaIngreso || r.fechaIngreso > f.fechaHasta)) return false;
  if (f.geoQuality && r.geoStatus !== f.geoQuality) return false;
  return true;
}

function inc(map,key){ if (!key) key='Sin dato'; map.set(key,(map.get(key)||0)+1); }
function topEntries(map,n=12){ return Array.from(map.entries()).sort((a,b)=>b[1]-a[1] || String(a[0]).localeCompare(String(b[0]),'es')).slice(0,n).map(([name,value])=>({name,value})); }

function applyFilter(criteria={}){
  activeIndices=[];
  const hourCounts = Array(24).fill(0);
  const subtypeCounts=new Map(), districtCounts=new Map(), originCounts=new Map();
  const delegations=new Set();
  let geo=0,geoValid=0,geoSuspect=0,geoInvalid=0,noUb=0;
  const tableRows=[];
  for (let i=0;i<records.length;i++){
    const r=records[i];
    if (!matches(r,criteria)) continue;
    activeIndices.push(i);
    if (r.geoStatus==='valid'){geo++;geoValid++;} else if(r.geoStatus==='suspect'){geo++;geoSuspect++;} else {geoInvalid++;}
    if (r.noUbicado) noUb++;
    if (r.hour !== null) hourCounts[r.hour]++;
    inc(subtypeCounts,r.subtipo); inc(districtCounts,r.distrito); inc(originCounts,r.origen);
    if (r.unidadFormal) delegations.add(r.unidadFormal);
    if (tableRows.length < 300){
      tableRows.push({idx:i,id:r.id,fecha:r.fechaIngreso,hora:r.horaTexto,tipo:r.tipo,subtipo:r.subtipo,provincia:r.provincia,canton:r.canton,distrito:r.distrito,barrio:r.barrio,unidadFormal:r.unidadFormal});
    }
  }
  const total=activeIndices.length;
  const max=Math.max(...hourCounts);
  const peakHours=max>0 ? hourCounts.map((v,i)=>v===max?i:null).filter(v=>v!==null) : [];
  postMessage({type:'filtered', stats:{
    total,geo,geoValid,geoSuspect,geoInvalid,noUb,noUbPct:total?noUb/total*100:0,geoPct:total?geo/total*100:0,geoValidPct:total?geoValid/total*100:0,geoSuspectPct:total?geoSuspect/total*100:0,geoInvalidPct:total?geoInvalid/total*100:0,delegationCount:delegations.size,
    hourCounts,peakHours,peakCount:max,subtypeTop:topEntries(subtypeCounts,12),districtTop:topEntries(districtCounts,12),originTop:topEntries(originCounts,10),tableRows
  }});
}

function mapQuery(bounds,limit=25000){
  if (!bounds) return;
  let seen=0,validInBounds=0,suspectInBounds=0;
  const points=[];
  for (const idx of activeIndices){
    const r=records[idx];
    if (r.geoStatus==='invalid' || r.lat===null || r.lng===null) continue;
    if (r.lat < bounds.south || r.lat > bounds.north || r.lng < bounds.west || r.lng > bounds.east) continue;
    seen++;
    if(r.geoStatus==='valid') validInBounds++; else if(r.geoStatus==='suspect') suspectInBounds++;
    const point=[idx,r.lat,r.lng,r.subtipo||r.tipo,r.fechaIngreso,r.horaTexto,r.noUbicado,r.geoStatus,r.geoReason];
    if (points.length < limit) points.push(point);
    else {
      const j=Math.floor(Math.random()*seen);
      if (j<limit) points[j]=point;
    }
  }
  postMessage({type:'mapPoints',points,totalInBounds:seen,validInBounds,suspectInBounds,truncated:seen>limit,limit});
}

self.onmessage = async (e) => {
  const msg=e.data||{};
  try{
    if (msg.type==='load'){
      loadedFileName=msg.fileName||'archivo.xlsx'; records=[]; activeIndices=[];
      postMessage({type:'progress',text:'Leyendo estructura del libro…'});
      const wb=XLSX.read(msg.buffer,{type:'array',cellDates:false});
      const sheetName=wb.SheetNames[0];
      const sheet=wb.Sheets[sheetName];
      if (!sheet || !sheet['!ref']) throw new Error('La primera hoja no contiene datos reconocibles.');
      const range=XLSX.utils.decode_range(sheet['!ref']);
      const {map}=getHeaderMap(sheet,range);
      const required=['id','origen','tipo','subtipo','diligencia','provincia','canton','distrito','barrio','fechaIngreso','unidad','unidadFormal','lat','lng'];
      const missing=required.filter(k=>map[k]===null || map[k]===undefined);
      if (missing.length) throw new Error('Faltan columnas requeridas: '+missing.join(', '));
      const totalRows=Math.max(0,range.e.r-range.s.r);
      for (let r=range.s.r+1;r<=range.e.r;r++){
        if ((r-range.s.r)%12000===0) postMessage({type:'progress',text:`Procesando ${Math.min(r-range.s.r,totalRows).toLocaleString('es-CR')} de ${totalRows.toLocaleString('es-CR')} filas…`});
        const id=clean(cellValue(sheet,r,map.id));
        const tipo=clean(cellValue(sheet,r,map.tipo));
        const fechaRaw=cellValue(sheet,r,map.fechaIngreso);
        const horaRaw=cellValue(sheet,r,map.hora);
        const horaIncRaw=cellValue(sheet,r,map.horaIncidente);
        if (!id && !tipo && fechaRaw==='') continue;
        const hour=parseHour(horaIncRaw)!==null ? parseHour(horaIncRaw) : parseHour(horaRaw);
        const horaTexto=displayTime(horaIncRaw!==''?horaIncRaw:horaRaw,hour);
        const rec={
          id, origen:clean(cellValue(sheet,r,map.origen)), tipo, subtipo:normalizeSubtypeDisplay(cellValue(sheet,r,map.subtipo)), diligencia:clean(cellValue(sheet,r,map.diligencia)),
          provincia:clean(cellValue(sheet,r,map.provincia)), canton:clean(cellValue(sheet,r,map.canton)), distrito:clean(cellValue(sheet,r,map.distrito)), barrio:clean(cellValue(sheet,r,map.barrio)),
          incidente:clean(cellValue(sheet,r,map.incidente)), direccion:clean(cellValue(sheet,r,map.direccion)), tipoLugar:clean(cellValue(sheet,r,map.tipoLugar)),
          fechaIngreso:parseDate(fechaRaw), horaTexto, hour, fechaFinalizado:parseDate(cellValue(sheet,r,map.fechaFinalizado)), indicativo:clean(cellValue(sheet,r,map.indicativo)),
          unidad:clean(cellValue(sheet,r,map.unidad)), unidadFormal:clean(cellValue(sheet,r,map.unidadFormal)), diferenciaMin:clean(cellValue(sheet,r,map.diferenciaMin)), estado:clean(cellValue(sheet,r,map.estado)),
          fechaCreado:parseDate(cellValue(sheet,r,map.fechaCreado)), lat:parseCoord(cellValue(sheet,r,map.lat)), lng:parseCoord(cellValue(sheet,r,map.lng)), rango:clean(cellValue(sheet,r,map.rango)),
          anio:clean(cellValue(sheet,r,map.anio)), mes:clean(cellValue(sheet,r,map.mes)), dia:clean(cellValue(sheet,r,map.dia)), numSemana:clean(cellValue(sheet,r,map.numSemana)),
          fuente:clean(cellValue(sheet,r,map.fuente)), codSubUnidad:clean(cellValue(sheet,r,map.codSubUnidad)), producto:clean(cellValue(sheet,r,map.producto))
        };
        const g=validateGeo(rec.lat,rec.lng,rec.provincia);
        rec.geoStatus=g.status; rec.geoStatusLabel=g.label; rec.geoReason=g.reason;
        rec.noUbicado=isNoUbicado(rec.tipo);
        records.push(rec);
      }
      activeIndices=records.map((_,i)=>i);
      const options=uniqOptions();
      postMessage({type:'loaded',fileName:loadedFileName,total:records.length,options,sheetName,missingOptional:Object.entries(map).filter(([k,v])=>v===null||v===undefined).map(([k])=>k)});
      applyFilter({});
    } else if (msg.type==='filter') applyFilter(msg.criteria||{});
    else if (msg.type==='mapQuery') mapQuery(msg.bounds,msg.limit||25000);
    else if (msg.type==='detail'){
      const r=records[msg.idx];
      postMessage({type:'detail',record:r||null,idx:msg.idx});
    }
  }catch(err){
    postMessage({type:'error',message:err && err.message ? err.message : String(err)});
  }
};
