const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat('es-CR');
const pct = (v) => `${Number(v||0).toLocaleString('es-CR',{minimumFractionDigits:1,maximumFractionDigits:1})}%`;
const esc = (s='') => String(s ?? '').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

let worker=null, dataLoaded=false, currentStats=null, currentFileName='';
let map, markerLayer, mapRequestTimer;
let clockChart, typeChart;
let templateImagePromise=null;

function setLoading(show,title='Procesando…',text='Esto puede tardar unos segundos.'){
  $('loadingOverlay').hidden=!show;
  $('loadingTitle').textContent=title;
  $('loadingText').textContent=text;
}

function initWorker(){
  worker=new Worker('worker.js');
  worker.onmessage=(e)=>{
    const msg=e.data||{};
    if (msg.type==='progress') $('loadingText').textContent=msg.text;
    if (msg.type==='loaded') onLoaded(msg);
    if (msg.type==='filtered') onFiltered(msg.stats);
    if (msg.type==='mapPoints') renderMapPoints(msg);
    if (msg.type==='detail') openDetail(msg.record);
    if (msg.type==='error'){
      setLoading(false);
      alert('No fue posible procesar el archivo. '+msg.message);
    }
  };
}

async function loadExcel(file){
  if (!file) return;
  setLoading(true,'Procesando archivo Excel','Preparando el archivo para análisis…');
  currentFileName=file.name;
  const buffer=await file.arrayBuffer();
  worker.postMessage({type:'load',fileName:file.name,buffer},[buffer]);
}

function onLoaded(msg){
  dataLoaded=true;
  currentFileName=msg.fileName;
  $('fileStatus').classList.add('loaded');
  $('fileStatus').innerHTML=`<span class="status-dot"></span><span>${esc(msg.fileName)} · ${fmt.format(msg.total)} registros</span>`;
  populateSelect('fUnidad',msg.options.unidad,'Todas');
  populateSelect('fUnidadFormal',msg.options.unidadFormal,'Todas');
  populateSelect('fProvincia',msg.options.provincia,'Todas');
  populateSelect('fCanton',msg.options.canton,'Todos');
  populateSelect('fDistrito',msg.options.distrito,'Todos');
  populateSelect('fTipo',msg.options.tipo,'Todos');
  populateSelect('fOrigen',msg.options.origen,'Todos');
  populateSelect('fSubtipo',msg.options.subtipo,'Todos');
  const dl=$('barriosList'); dl.innerHTML='';
  msg.options.barrio.slice(0,1800).forEach(v=>{const o=document.createElement('option');o.value=v;dl.appendChild(o)});
  $('fFechaDesde').min=msg.options.minDate||''; $('fFechaDesde').max=msg.options.maxDate||'';
  $('fFechaHasta').min=msg.options.minDate||''; $('fFechaHasta').max=msg.options.maxDate||'';
  $('uploadPanel').querySelector('h1').textContent='Archivo cargado correctamente';
  $('uploadPanel').querySelector('p').textContent=`${msg.fileName} · ${fmt.format(msg.total)} registros detectados en ${msg.sheetName}. Puede aplicar los filtros y navegar por las secciones.`;
}

function populateSelect(id,values,allLabel){
  const el=$(id); const current=el.value;
  el.innerHTML='';
  const first=document.createElement('option'); first.value=''; first.textContent=allLabel; el.appendChild(first);
  (values||[]).forEach(v=>{ const o=document.createElement('option'); o.value=v;o.textContent=v; el.appendChild(o); });
  if ([...el.options].some(o=>o.value===current)) el.value=current;
}

function criteriaFromUI(){
  return {unidad:$('fUnidad').value,unidadFormal:$('fUnidadFormal').value,fechaDesde:$('fFechaDesde').value,fechaHasta:$('fFechaHasta').value,provincia:$('fProvincia').value,canton:$('fCanton').value,distrito:$('fDistrito').value,barrio:$('fBarrio').value.trim(),tipo:$('fTipo').value,origen:$('fOrigen').value,subtipo:$('fSubtipo').value};
}
function applyFilters(){ if (!dataLoaded) return alert('Primero cargue el archivo Excel.'); setLoading(true,'Aplicando filtros','Recalculando indicadores, gráficos y mapa…'); worker.postMessage({type:'filter',criteria:criteriaFromUI()}); }
function resetFilters(){ ['fUnidad','fUnidadFormal','fProvincia','fCanton','fDistrito','fTipo','fOrigen','fSubtipo'].forEach(id=>$(id).value=''); ['fFechaDesde','fFechaHasta','fBarrio'].forEach(id=>$(id).value=''); if(dataLoaded) applyFilters(); }

function onFiltered(stats){
  currentStats=stats;
  $('kpiTotal').textContent=fmt.format(stats.total);
  $('kpiGeo').textContent=fmt.format(stats.geo);
  $('kpiGeoPct').textContent=`${pct(stats.geoPct)} del total filtrado`;
  $('kpiNoUb').textContent=fmt.format(stats.noUb);
  $('kpiNoUbPct').textContent=`${pct(stats.noUbPct)} del total filtrado`;
  $('kpiPeak').textContent=stats.peakHours.length ? stats.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' / ') : '—';
  $('kpiDelegaciones').textContent=fmt.format(stats.delegationCount);
  $('peakBadge').textContent=stats.peakHours.length ? `${stats.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' · ')} · ${fmt.format(stats.peakCount)} incidentes` : 'Sin datos horarios';
  updateCharts(stats);
  updateTable(stats.tableRows,stats.total);
  updateInsight(stats);
  setLoading(false);
  scheduleMapQuery();
}

function updateInsight(s){
  if (!s.total){ $('insightText').innerHTML='No se encontraron registros para la combinación de filtros seleccionada.'; return; }
  const top=s.typeTop[0]; const district=s.districtTop[0];
  const peak=s.peakHours.length ? s.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' y ') : 'sin hora identificable';
  $('insightText').innerHTML=`El filtro actual contiene <strong>${fmt.format(s.total)} incidentes</strong>. ${fmt.format(s.geo)} cuentan con coordenadas válidas (${pct(s.geoPct)}). El tipo más frecuente es <strong>${esc(top?.name||'Sin dato')}</strong>${top?` con ${fmt.format(top.value)} registros`:''}. La mayor concentración horaria se observa a las <strong>${peak}</strong>. Los registros clasificados como <strong>NO UBICADO / OTROS</strong> representan ${pct(s.noUbPct)} del total. ${district?`El distrito con mayor cantidad dentro del filtro es <strong>${esc(district.name)}</strong> (${fmt.format(district.value)}).`:''}`;
}

function initCharts(){
  clockChart=echarts.init($('clockChart'));
  typeChart=echarts.init($('typeChart'));
  window.addEventListener('resize',()=>{clockChart.resize();typeChart.resize();});
}
function updateCharts(s){
  const hours=Array.from({length:24},(_,i)=>`${String(i).padStart(2,'0')}h`);
  clockChart.setOption({
    tooltip:{trigger:'item',formatter:p=>`${p.name}: ${fmt.format(p.value)} incidentes`},
    polar:{radius:['18%','82%']},angleAxis:{type:'category',data:hours,startAngle:90,axisLabel:{fontSize:10,color:'#667085'},axisLine:{lineStyle:{color:'#d9e0ea'}}},
    radiusAxis:{axisLabel:{show:false},splitLine:{lineStyle:{color:'#edf1f5'}}},
    series:[{type:'bar',coordinateSystem:'polar',data:s.hourCounts,itemStyle:{color:(p)=>s.peakHours.includes(p.dataIndex)?'#d8b55b':'#0b357d',borderRadius:3},roundCap:true,barWidth:'72%'}]
  },true);
  const names=s.typeTop.map(x=>x.name).reverse(), vals=s.typeTop.map(x=>x.value).reverse();
  typeChart.setOption({
    grid:{left:150,right:25,top:12,bottom:25},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},
    xAxis:{type:'value',axisLabel:{color:'#667085'},splitLine:{lineStyle:{color:'#edf1f5'}}},
    yAxis:{type:'category',data:names,axisLabel:{color:'#475467',width:135,overflow:'truncate'}},
    series:[{type:'bar',data:vals,itemStyle:{color:'#0b357d',borderRadius:[0,5,5,0]},barMaxWidth:22}]
  },true);
}

function updateTable(rows,total){
  $('tableCount').textContent=`${fmt.format(rows.length)} mostrados de ${fmt.format(total)}`;
  const body=$('incidentTableBody'); body.innerHTML='';
  if (!rows.length){body.innerHTML='<tr><td colspan="10" class="empty-cell">No hay registros para mostrar.</td></tr>';return;}
  const frag=document.createDocumentFragment();
  rows.forEach(r=>{
    const tr=document.createElement('tr'); tr.dataset.idx=r.idx;
    tr.innerHTML=`<td>${esc(r.id)}</td><td>${esc(r.fecha)}</td><td>${esc(r.hora)}</td><td>${esc(r.tipo)}</td><td>${esc(r.subtipo)}</td><td>${esc(r.provincia)}</td><td>${esc(r.canton)}</td><td>${esc(r.distrito)}</td><td>${esc(r.barrio)}</td><td>${esc(r.unidadFormal)}</td>`;
    tr.addEventListener('click',()=>requestDetail(r.idx)); frag.appendChild(tr);
  }); body.appendChild(frag);
}

function initMap(){
  map=L.map('map',{preferCanvas:true,minZoom:5}).setView([9.75,-83.75],7);

  // Capas sin API key. Se evita CARTO para que GitHub Pages no muestre
  // la marca "API KEY REQUIRED". La capa oscura usa OSM con filtro visual.
  map.createPane('darkTiles');
  map.getPane('darkTiles').classList.add('leaflet-dark-pane');
  map.getPane('darkTiles').style.zIndex=200;

  const normal=L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
    {maxZoom:19,attribution:'Tiles &copy; Esri'}
  ).addTo(map);

  const calles=L.tileLayer(
    'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    {maxZoom:19,attribution:'&copy; OpenStreetMap contributors'}
  );

  const oscuro=L.tileLayer(
    'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    {maxZoom:19,attribution:'&copy; OpenStreetMap contributors',pane:'darkTiles'}
  );

  L.control.layers({'Normal':normal,'Calles':calles,'Oscuro':oscuro},null,{position:'topright'}).addTo(map);
  markerLayer=L.markerClusterGroup({chunkedLoading:true,showCoverageOnHover:false,maxClusterRadius:55,disableClusteringAtZoom:17,removeOutsideVisibleBounds:true});
  map.addLayer(markerLayer);
  map.on('moveend zoomend',scheduleMapQuery);
}
function scheduleMapQuery(){
  if (!dataLoaded || !map) return;
  clearTimeout(mapRequestTimer);
  mapRequestTimer=setTimeout(()=>{
    const b=map.getBounds();
    worker.postMessage({type:'mapQuery',limit:25000,bounds:{north:b.getNorth(),south:b.getSouth(),east:b.getEast(),west:b.getWest()}});
  },220);
}
function renderMapPoints(msg){
  markerLayer.clearLayers();
  const markers=[];
  for (const p of msg.points){
    const [idx,lat,lng,tipo,fecha,hora,noUb]=p;
    const marker=L.circleMarker([lat,lng],{radius:5,weight:1,color:noUb?'#8a1c16':'#fff',fillColor:noUb?'#b3261e':'#0b357d',fillOpacity:.84});
    marker.bindTooltip(`${esc(tipo||'Incidente')}<br>${esc(fecha||'')} ${esc(hora||'')}`,{direction:'top'});
    marker.on('click',()=>requestDetail(idx));
    markers.push(marker);
  }
  markerLayer.addLayers(markers);
  $('mapCounter').textContent=msg.truncated ? `${fmt.format(msg.points.length)} visibles de ${fmt.format(msg.totalInBounds)} en esta vista` : `${fmt.format(msg.totalInBounds)} puntos en esta vista`;
  $('mapNote').textContent=msg.truncated ? `Por rendimiento se muestra una muestra de ${fmt.format(msg.limit)} puntos en esta vista. Acerque el mapa para visualizar de forma puntual todos los incidentes del sector.` : 'Puntos precisos según Latitud y Longitud del archivo. Capas: Normal, Calles y Oscuro.';
}

function requestDetail(idx){ worker.postMessage({type:'detail',idx}); }
function openDetail(r){
  if (!r) return;
  $('modalTitle').textContent=`Incidente ${r.id||''}`;
  const item=(label,value)=>`<div class="detail-item"><span>${esc(label)}</span><strong>${esc(value||'—')}</strong></div>`;
  $('modalBody').innerHTML=`
    ${r.diligencia?'<div class="warning">La diligencia policial puede contener información sensible. Su consulta debe realizarse únicamente conforme a las autorizaciones institucionales aplicables.</div>':''}
    <div class="detail-grid">
      ${item('Tipo',r.tipo)}${item('Subtipo',r.subtipo)}${item('Incidente',r.incidente)}
      ${item('Origen',r.origen)}${item('Estado',r.estado)}${item('Fecha de ingreso',r.fechaIngreso)}
      ${item('Hora',r.horaTexto)}${item('Provincia',r.provincia)}${item('Cantón',r.canton)}
      ${item('Distrito',r.distrito)}${item('Barrio',r.barrio)}${item('Tipo de lugar',r.tipoLugar)}
      ${item('Unidad',r.unidad)}${item('Delegación',r.unidadFormal)}${item('Indicativo',r.indicativo)}
      ${item('Latitud',r.lat)}${item('Longitud',r.lng)}${item('Dirección',r.direccion)}
    </div>
    <div class="diligence"><h3>Diligencia policial</h3>${esc(r.diligencia||'No se registra texto de diligencia policial para este incidente.')}</div>`;
  $('detailModal').showModal();
}

function formatDateLong(date=new Date()){
  const months=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return `${date.getDate()} de ${months[date.getMonth()]} de ${date.getFullYear()}`;
}
function activeFilterLines(){
  const f=criteriaFromUI(); const labels={unidad:'Unidad',unidadFormal:'Delegación',fechaDesde:'Desde',fechaHasta:'Hasta',provincia:'Provincia',canton:'Cantón',distrito:'Distrito',barrio:'Barrio',tipo:'Tipo',origen:'Origen',subtipo:'Subtipo'};
  return Object.entries(f).filter(([,v])=>v).map(([k,v])=>`${labels[k]}: ${v}`);
}
function loadTemplateImage(){
  if (!templateImagePromise) templateImagePromise=new Promise((resolve,reject)=>{const im=new Image();im.onload=()=>resolve(im);im.onerror=reject;im.src='assets/report-template.png';});
  return templateImagePromise;
}
async function generatePdf(){
  if (!currentStats || !dataLoaded) return alert('Primero cargue y analice un archivo.');
  setLoading(true,'Generando informe PDF','Preparando el informe institucional…');
  try{
    const {jsPDF}=window.jspdf; const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
    const template=await loadTemplateImage(); const title=$('reportTitle').value.trim()||'Informe de Incidencia Policial'; const note=$('reportNote').value.trim();
    const addTemplate=(pageNo)=>{
      doc.addImage(template,'PNG',0,0,210,297,undefined,'FAST');
      doc.setTextColor(20,28,45); doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.text(formatDateLong().toUpperCase(),184,31,{align:'right'});
      doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(50,58,72);
      doc.text('Subdirección General de la Fuerza Pública',29,258);
      doc.text('San José, Zapote, Barrio Córdoba, frente al Liceo Dr. José María Castro Madriz',29,262);
      doc.text('Módulo A Daniel Oduber Quirós',29,266);
      doc.setTextColor(16,74,132); doc.text('subdirecion1a@fuerzapublica.go.cr / www.seguridadpublica.go.cr',29,270);
      doc.setTextColor(50,58,72); doc.text('2600-4187',29,274); doc.text('@seguridadcrc',29,278);
      doc.setFontSize(7.5); doc.text(`Página ${pageNo}`,184,281,{align:'right'});
    };
    const header=(subtitle)=>{doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.setFontSize(18);doc.text(title,18,48);doc.setTextColor(102,112,133);doc.setFontSize(9);doc.setFont('helvetica','normal');doc.text(subtitle,18,55);doc.setDrawColor(216,181,91);doc.setLineWidth(1.2);doc.line(18,59,70,59);};
    const card=(x,y,w,label,value,detail)=>{doc.setFillColor(247,249,252);doc.setDrawColor(220,226,236);doc.roundedRect(x,y,w,25,3,3,'FD');doc.setTextColor(102,112,133);doc.setFont('helvetica','bold');doc.setFontSize(7);doc.text(label.toUpperCase(),x+4,y+6);doc.setTextColor(6,36,94);doc.setFontSize(15);doc.text(value,x+4,y+15);doc.setTextColor(102,112,133);doc.setFont('helvetica','normal');doc.setFontSize(6.5);doc.text(detail,x+4,y+21);};

    addTemplate(1); header('Resumen estadístico del filtro seleccionado');
    card(18,67,40,'Incidentes',fmt.format(currentStats.total),'Total filtrado');
    card(61,67,40,'Georreferenciados',fmt.format(currentStats.geo),pct(currentStats.geoPct));
    card(104,67,40,'No ubicado / Otros',fmt.format(currentStats.noUb),pct(currentStats.noUbPct));
    card(147,67,45,'Hora pico',currentStats.peakHours.length?currentStats.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' / '):'—',`${fmt.format(currentStats.peakCount)} registros`);
    doc.setTextColor(52,64,84);doc.setFontSize(9);doc.setFont('helvetica','bold');doc.text('Alcance del análisis',18,103);doc.setFont('helvetica','normal');doc.setFontSize(8);
    const filters=activeFilterLines(); const scope=filters.length?filters.join(' · '):'Sin filtros adicionales: totalidad del archivo cargado.';
    doc.text(doc.splitTextToSize(scope,174),18,109);
    if(note){doc.setFont('helvetica','bold');doc.text('Observación:',18,124);doc.setFont('helvetica','normal');doc.text(doc.splitTextToSize(note,174),18,130);}
    const clockImg=clockChart.getDataURL({type:'png',pixelRatio:2,backgroundColor:'#ffffff'});
    doc.setFont('helvetica','bold');doc.setTextColor(6,36,94);doc.setFontSize(10);doc.text('Data reloj · distribución por hora',18,151);
    doc.addImage(clockImg,'PNG',18,156,82,75,undefined,'FAST');
    const typeImg=typeChart.getDataURL({type:'png',pixelRatio:2,backgroundColor:'#ffffff'});
    doc.text('Principales tipos de incidente',108,151);doc.addImage(typeImg,'PNG',106,158,86,68,undefined,'FAST');

    doc.addPage(); addTemplate(2); header('Clasificación y concentración territorial');
    doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.setFontSize(10);doc.text('Tipos de incidente con mayor frecuencia',18,69);
    doc.autoTable({startY:74,margin:{left:18,right:108},head:[['Tipo de incidente','Cantidad']],body:currentStats.typeTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{fontSize:7,cellPadding:2.2,textColor:[50,64,84]},headStyles:{fillColor:[6,36,94],textColor:255},columnStyles:{1:{halign:'right'}}});
    doc.text('Distritos con mayor frecuencia',108,69);
    doc.autoTable({startY:74,margin:{left:108,right:18},head:[['Distrito','Cantidad']],body:currentStats.districtTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{fontSize:7,cellPadding:2.2,textColor:[50,64,84]},headStyles:{fillColor:[6,36,94],textColor:255},columnStyles:{1:{halign:'right'}}});
    const y=Math.max(doc.lastAutoTable?.finalY||145,150)+10;
    doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.setFontSize(10);doc.text('Indicador de NO UBICADO / OTROS',18,y);
    doc.setFillColor(253,233,231);doc.setDrawColor(230,170,165);doc.roundedRect(18,y+6,174,34,3,3,'FD');doc.setTextColor(179,38,30);doc.setFontSize(18);doc.text(pct(currentStats.noUbPct),25,y+22);doc.setFontSize(8);doc.setTextColor(52,64,84);doc.setFont('helvetica','normal');doc.text(`${fmt.format(currentStats.noUb)} registros clasificados como NO UBICADO / OTROS de un total de ${fmt.format(currentStats.total)} incidentes filtrados.`,62,y+19,{maxWidth:122});
    doc.setTextColor(102,112,133);doc.setFontSize(7.5);doc.text('Fuente: archivo Excel cargado en la herramienta. El cálculo se realiza sobre el filtro activo al momento de generar el informe.',18,y+52,{maxWidth:174});

    doc.addPage(); addTemplate(3); header('Origen de los incidentes y notas metodológicas');
    doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.setFontSize(10);doc.text('Principales orígenes',18,69);
    doc.autoTable({startY:74,margin:{left:18,right:100},head:[['Origen','Cantidad']],body:currentStats.originTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{fontSize:7.5,cellPadding:2.4,textColor:[50,64,84]},headStyles:{fillColor:[6,36,94],textColor:255},columnStyles:{1:{halign:'right'}}});
    doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.setFontSize(10);doc.text('Notas',115,69);
    doc.setFont('helvetica','normal');doc.setTextColor(52,64,84);doc.setFontSize(8);
    const notes=[
      'La georreferenciación utiliza directamente las columnas Latitud y Longitud del archivo.',
      'La distribución horaria prioriza HoraIncidente cuando existe; en caso contrario utiliza la columna Hora.',
      'El porcentaje de NO UBICADO / OTROS se calcula sobre el total de incidentes del filtro activo.',
      'El informe es estadístico. La diligencia policial se consulta únicamente desde el detalle de cada incidente dentro de la aplicación.'
    ];
    let ny=77; notes.forEach((n,i)=>{doc.setFillColor(247,249,252);doc.roundedRect(112,ny-5,80,19,2,2,'F');doc.setTextColor(6,36,94);doc.setFont('helvetica','bold');doc.text(String(i+1),116,ny+1);doc.setTextColor(52,64,84);doc.setFont('helvetica','normal');doc.text(doc.splitTextToSize(n,68),122,ny);ny+=24;});
    doc.setFontSize(7);doc.setTextColor(102,112,133);doc.text(`Archivo analizado: ${currentFileName}`,18,225,{maxWidth:174});

    doc.save(`Informe_Incidencia_${new Date().toISOString().slice(0,10)}.pdf`);
  }catch(err){console.error(err);alert('No fue posible generar el PDF: '+(err.message||err));}
  finally{setLoading(false);}
}

function initTabs(){
  document.querySelectorAll('.tab').forEach(btn=>btn.addEventListener('click',()=>{
    document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active')); btn.classList.add('active');
    document.querySelectorAll('.tab-page').forEach(x=>x.classList.remove('active')); $(`tab-${btn.dataset.tab}`).classList.add('active');
    if(btn.dataset.tab==='mapa') setTimeout(()=>{map.invalidateSize();scheduleMapQuery();},120);
    if(btn.dataset.tab==='dashboard') setTimeout(()=>{clockChart.resize();typeChart.resize();},80);
  }));
}

window.addEventListener('DOMContentLoaded',()=>{
  initWorker(); initTabs(); initCharts(); initMap();
  $('fileInput').addEventListener('change',e=>loadExcel(e.target.files[0]));
  $('applyFilters').addEventListener('click',applyFilters); $('resetFilters').addEventListener('click',resetFilters);
  $('generatePdf').addEventListener('click',generatePdf); $('closeModal').addEventListener('click',()=>$('detailModal').close());
  $('detailModal').addEventListener('click',(e)=>{const r=e.currentTarget.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)e.currentTarget.close();});
});
