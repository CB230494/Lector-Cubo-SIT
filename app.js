const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat('es-CR');
const pct = (v) => `${Number(v||0).toLocaleString('es-CR',{minimumFractionDigits:1,maximumFractionDigits:1})}%`;
const esc = (s='') => String(s ?? '').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

let worker=null, dataLoaded=false, currentStats=null, currentFileName='';
let map, markerLayer, pointLayer, heatLayer, mapRequestTimer;
let mapMode='cluster', lastMapMsg=null;
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
  const top=s.subtypeTop[0]; const district=s.districtTop[0];
  const peak=s.peakHours.length ? s.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' y ') : 'sin hora identificable';
  $('insightText').innerHTML=`El filtro actual contiene <strong>${fmt.format(s.total)} incidentes</strong>. El subtipo más frecuente es <strong>${esc(top?.name||'Sin dato')}</strong>${top?` con ${fmt.format(top.value)} registros`:''}. La mayor concentración horaria se observa a las <strong>${peak}</strong>. Los registros clasificados como <strong>NO UBICADO / OTROS</strong> representan ${pct(s.noUbPct)} del total. ${district?`El distrito con mayor cantidad dentro del filtro es <strong>${esc(district.name)}</strong> (${fmt.format(district.value)}).`:''}`;
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
  const names=s.subtypeTop.map(x=>x.name).reverse(), vals=s.subtypeTop.map(x=>x.value).reverse();
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
  map=L.map('map',{preferCanvas:true,minZoom:5,zoomControl:true}).setView([9.75,-83.75],7);

  map.createPane('darkTiles');
  map.getPane('darkTiles').classList.add('leaflet-dark-pane');
  map.getPane('darkTiles').style.zIndex=200;

  const calles=L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
    {maxZoom:19,attribution:'Tiles &copy; Esri'}
  ).addTo(map);

  const satBase=L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    {maxZoom:19,attribution:'Imagery &copy; Esri'}
  );
  const satLabels=L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
    {maxZoom:19,attribution:'Labels &copy; Esri'}
  );
  const satelite=L.layerGroup([satBase,satLabels]);

  const oscuro=L.tileLayer(
    'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    {maxZoom:19,attribution:'&copy; OpenStreetMap contributors',pane:'darkTiles'}
  );

  L.control.layers({'Calles':calles,'Satélite':satelite,'Oscuro':oscuro},null,{position:'topright',collapsed:false}).addTo(map);

  markerLayer=L.markerClusterGroup({
    chunkedLoading:true,showCoverageOnHover:false,maxClusterRadius:58,disableClusteringAtZoom:17,removeOutsideVisibleBounds:true,
    spiderfyOnMaxZoom:true,zoomToBoundsOnClick:true,
    iconCreateFunction(cluster){
      const n=cluster.getChildCount();
      const size=n<50?'small':n<500?'medium':'large';
      return L.divIcon({html:`<div class="cluster-core"><span>${fmt.format(n)}</span><small>inc.</small></div>`,className:`smart-cluster ${size}`,iconSize:L.point(size==='small'?42:size==='medium'?50:58,size==='small'?42:size==='medium'?50:58)});
    }
  });
  pointLayer=L.layerGroup();
  heatLayer=L.heatLayer([],{radius:24,blur:20,maxZoom:16,minOpacity:.35});
  map.addLayer(markerLayer);
  map.on('moveend zoomend',scheduleMapQuery);

  document.querySelectorAll('.map-mode[data-map-mode]').forEach(btn=>btn.addEventListener('click',()=>{
    mapMode=btn.dataset.mapMode;
    document.querySelectorAll('.map-mode[data-map-mode]').forEach(b=>b.classList.toggle('active',b===btn));
    drawMapMode();
  }));
  $('centerCR').addEventListener('click',()=>map.setView([9.75,-83.75],7,{animate:true}));
}
function scheduleMapQuery(){
  if (!dataLoaded || !map) return;
  clearTimeout(mapRequestTimer);
  mapRequestTimer=setTimeout(()=>{
    const b=map.getBounds();
    worker.postMessage({type:'mapQuery',limit:25000,bounds:{north:b.getNorth(),south:b.getSouth(),east:b.getEast(),west:b.getWest()}});
  },220);
}
function pointStyle(noUb){
  if(noUb) return {radius:6,weight:2,color:'#ffffff',fillColor:'#b3261e',fillOpacity:.9};
  return {radius:5.5,weight:1.5,color:'#ffffff',fillColor:'#0b5cad',fillOpacity:.88};
}
function buildCircleMarker(p){
  const [idx,lat,lng,subtipo,fecha,hora,noUb]=p;
  const marker=L.circleMarker([lat,lng],pointStyle(noUb));
  marker.bindTooltip(`<strong>${esc(subtipo||'Incidente')}</strong><br>${esc(fecha||'')} ${esc(hora||'')}`,{direction:'top',className:'incident-tooltip'});
  marker.on('click',()=>requestDetail(idx));
  return marker;
}
function drawMapMode(){
  if(!map || !lastMapMsg) return;
  [markerLayer,pointLayer,heatLayer].forEach(layer=>{if(map.hasLayer(layer)) map.removeLayer(layer)});
  const pts=lastMapMsg.points||[];
  if(mapMode==='heat'){
    const heat=pts.filter(p=>p[7]!=='invalid').map(p=>[p[1],p[2],p[6]?1.0:.5]);
    heatLayer.setLatLngs(heat); map.addLayer(heatLayer);
  }else if(mapMode==='points'){
    pointLayer.clearLayers();
    const markers=pts.filter(p=>p[7]!=='invalid').map(buildCircleMarker);
    markers.forEach(m=>pointLayer.addLayer(m)); map.addLayer(pointLayer);
  }else{
    markerLayer.clearLayers();
    const markers=pts.filter(p=>p[7]!=='invalid').map(buildCircleMarker);
    markerLayer.addLayers(markers); map.addLayer(markerLayer);
  }
}
function renderMapPoints(msg){
  lastMapMsg=msg;
  drawMapMode();
  $('mapCounter').textContent=msg.truncated ? `${fmt.format(msg.points.length)} visibles de ${fmt.format(msg.totalInBounds)} en esta vista` : `${fmt.format(msg.totalInBounds)} puntos en esta vista`;
  $('mapNote').textContent=msg.truncated ? `Vista limitada a ${fmt.format(msg.limit)} puntos por rendimiento. Acerque el mapa para visualizar mayor detalle.` : `Cambie entre Calles, Satélite y Oscuro o utilice los modos Agrupados, Puntos y Mapa de calor.`;
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
    const {jsPDF}=window.jspdf;
    const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
    const template=await loadTemplateImage();
    const title=$('reportTitle').value.trim()||'Informe de Incidencia Policial';
    const note=$('reportNote').value.trim();
    const BLACK=[30,30,30], DARK=[55,55,55], MID=[105,105,105], LIGHT=[224,224,224], VERYLIGHT=[247,247,247];

    const addTemplate=(pageNo,totalPages=3)=>{
      doc.addImage(template,'PNG',0,0,210,297,undefined,'FAST');
      doc.setTextColor(...DARK); doc.setFont('times','bold'); doc.setFontSize(7);
      doc.text(formatDateLong().toUpperCase(),184,31,{align:'right'});
      doc.setFont('times','normal'); doc.setFontSize(6.3); doc.setTextColor(...DARK);
      doc.text('Subdirección General de la Fuerza Pública',29,258);
      doc.text('San José, Zapote, Barrio Córdoba, frente al Liceo Dr. José María Castro Madriz',29,262);
      doc.text('Módulo A Daniel Oduber Quirós',29,266);
      doc.text('subdirecion1a@fuerzapublica.go.cr / www.seguridadpublica.go.cr',29,270);
      doc.text('2600-4187',29,274); doc.text('@seguridadcrc',29,278);
      doc.setFontSize(7); doc.text(`Página ${pageNo} de ${totalPages}`,184,281,{align:'right'});
    };
    const header=(subtitle)=>{
      doc.setTextColor(...BLACK); doc.setFont('times','bold'); doc.setFontSize(16); doc.text(title,18,47);
      doc.setTextColor(...DARK); doc.setFont('times','normal'); doc.setFontSize(9); doc.text(subtitle,18,54);
      doc.setDrawColor(...MID); doc.setLineWidth(.35); doc.line(18,58,192,58);
    };
    const metric=(x,y,w,label,value,detail)=>{
      doc.setFillColor(255,255,255); doc.setDrawColor(...LIGHT); doc.roundedRect(x,y,w,24,2,2,'FD');
      doc.setTextColor(...MID); doc.setFont('times','bold'); doc.setFontSize(7); doc.text(label.toUpperCase(),x+4,y+6);
      doc.setTextColor(...BLACK); doc.setFontSize(14); doc.text(String(value),x+4,y+15);
      doc.setTextColor(...DARK); doc.setFont('times','normal'); doc.setFontSize(6.7); doc.text(String(detail),x+4,y+21);
    };
    const sectionTitle=(text,y)=>{doc.setTextColor(...BLACK);doc.setFont('times','bold');doc.setFontSize(10);doc.text(text,18,y);};
    const monoBars=(items,x,y,w,h)=>{
      if(!items?.length) return;
      const max=Math.max(...items.map(i=>i.value),1), row=h/items.length;
      items.forEach((it,i)=>{
        const yy=y+i*row;
        const label=String(it.name||'Sin dato');
        doc.setFont('times','normal');doc.setFontSize(6.6);doc.setTextColor(...DARK);
        const short=label.length>30?label.slice(0,29)+'…':label;
        doc.text(short,x,yy+4.2);
        const bx=x+55,bw=Math.max(1,(w-72)*(it.value/max));
        doc.setFillColor(92,92,92);doc.rect(bx,yy+1,bw,4,'F');
        doc.setTextColor(...BLACK);doc.text(fmt.format(it.value),x+w,yy+4.2,{align:'right'});
      });
    };
    const hourBars=(counts,x,y,w,h)=>{
      const max=Math.max(...counts,1), gap=.7, bw=(w-gap*23)/24;
      doc.setDrawColor(...LIGHT);doc.line(x,y+h,x+w,y+h);
      counts.forEach((v,i)=>{
        const bh=(v/max)*(h-8);
        doc.setFillColor(i===currentStats.peakHours[0]?65:130,i===currentStats.peakHours[0]?65:130,i===currentStats.peakHours[0]?65:130);
        doc.rect(x+i*(bw+gap),y+h-bh,bw,bh,'F');
        if(i%3===0){doc.setFont('times','normal');doc.setFontSize(5.7);doc.setTextColor(...MID);doc.text(String(i).padStart(2,'0'),x+i*(bw+gap)+bw/2,y+h+4,{align:'center'});}
      });
    };

    // Página 1
    addTemplate(1); header('Resumen estadístico del filtro seleccionado');
    metric(18,66,41,'Incidentes',fmt.format(currentStats.total),'Total filtrado');
    metric(62,66,41,'Georreferenciados',fmt.format(currentStats.geo),`${pct(currentStats.geoPct)} del total`);
    metric(106,66,41,'No ubicado / Otros',fmt.format(currentStats.noUb),pct(currentStats.noUbPct));
    metric(150,66,42,'Hora pico',currentStats.peakHours.length?currentStats.peakHours.map(h=>`${String(h).padStart(2,'0')}:00`).join(' / '):'—',`${fmt.format(currentStats.peakCount)} registros`);

    sectionTitle('Alcance del análisis',101);
    doc.setFont('times','normal');doc.setFontSize(8);doc.setTextColor(...DARK);
    const filters=activeFilterLines(); const scope=filters.length?filters.join(' · '):'Sin filtros adicionales: totalidad del archivo cargado.';
    let sy=107; const scopeLines=doc.splitTextToSize(scope,174); doc.text(scopeLines,18,sy); sy+=scopeLines.length*4.2+3;
    if(note){doc.setFont('times','bold');doc.text('Observación:',18,sy);sy+=4.5;doc.setFont('times','normal');const nl=doc.splitTextToSize(note,174);doc.text(nl,18,sy);sy+=nl.length*4.2+3;}
    sy=Math.max(sy,130);
    sectionTitle('Distribución de incidentes por hora',sy); hourBars(currentStats.hourCounts,18,sy+7,174,46);
    sy+=62;
    sectionTitle('Principales subtipos de incidente',sy); monoBars(currentStats.subtypeTop.slice(0,8),18,sy+6,174,46);

    // Página 2
    doc.addPage(); addTemplate(2); header('Clasificación y concentración territorial');
    sectionTitle('Subtipos de incidente con mayor frecuencia',68);
    doc.autoTable({startY:73,margin:{left:18,right:108},head:[['Subtipo de incidente','Cantidad']],body:currentStats.subtypeTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{font:'times',fontSize:7,cellPadding:2,textColor:BLACK,lineColor:LIGHT,lineWidth:.15},headStyles:{fillColor:[238,238,238],textColor:BLACK,fontStyle:'bold'},columnStyles:{1:{halign:'right'}},pageBreak:'avoid'});
    const leftEnd=doc.lastAutoTable?.finalY||120;
    sectionTitle('Distritos con mayor frecuencia',68);
    doc.autoTable({startY:73,margin:{left:108,right:18},head:[['Distrito','Cantidad']],body:currentStats.districtTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{font:'times',fontSize:7,cellPadding:2,textColor:BLACK,lineColor:LIGHT,lineWidth:.15},headStyles:{fillColor:[238,238,238],textColor:BLACK,fontStyle:'bold'},columnStyles:{1:{halign:'right'}},pageBreak:'avoid'});
    const rightEnd=doc.lastAutoTable?.finalY||120;
    let y=Math.max(leftEnd,rightEnd)+10;
    if(y>205) y=205;
    sectionTitle('Indicador NO UBICADO / OTROS',y);
    doc.setDrawColor(...LIGHT);doc.setFillColor(255,255,255);doc.roundedRect(18,y+6,174,27,2,2,'FD');
    doc.setTextColor(...BLACK);doc.setFont('times','bold');doc.setFontSize(15);doc.text(pct(currentStats.noUbPct),25,y+20);
    doc.setFont('times','normal');doc.setFontSize(8);doc.setTextColor(...DARK);
    const noUbText=`${fmt.format(currentStats.noUb)} registros clasificados como NO UBICADO / OTROS de un total de ${fmt.format(currentStats.total)} incidentes filtrados.`;
    doc.text(doc.splitTextToSize(noUbText,120),62,y+16);
    doc.setFontSize(7);doc.setTextColor(...MID);doc.text('Fuente: archivo Excel cargado en la herramienta. Cálculo realizado sobre el filtro activo.',18,y+42,{maxWidth:174});

    // Página 3
    doc.addPage(); addTemplate(3); header('Origen de los incidentes y notas metodológicas');
    sectionTitle('Principales orígenes',68);
    doc.autoTable({startY:73,margin:{left:18,right:105},head:[['Origen','Cantidad']],body:currentStats.originTop.map(x=>[x.name,fmt.format(x.value)]),theme:'grid',styles:{font:'times',fontSize:7.5,cellPadding:2.2,textColor:BLACK,lineColor:LIGHT,lineWidth:.15},headStyles:{fillColor:[238,238,238],textColor:BLACK,fontStyle:'bold'},columnStyles:{1:{halign:'right'}},pageBreak:'avoid'});
    sectionTitle('Notas metodológicas',112);
    const notes=[
      'La distribución horaria prioriza HoraIncidente cuando existe; en caso contrario utiliza la columna Hora.',
      'Los gráficos y cuadros de clasificación se presentan por SubtipoIncidente para ofrecer un mayor nivel de detalle.',
      'El porcentaje de NO UBICADO / OTROS se calcula sobre el total de incidentes del filtro activo.',
      'El informe es estadístico. La diligencia policial se consulta únicamente desde el detalle de cada incidente dentro de la aplicación.'
    ];
    doc.setFont('times','normal');doc.setFontSize(8);doc.setTextColor(...DARK);
    let ny=120;
    notes.forEach((n,i)=>{doc.setFont('times','bold');doc.text(`${i+1}.`,18,ny);doc.setFont('times','normal');const lines=doc.splitTextToSize(n,166);doc.text(lines,25,ny);ny+=lines.length*4.3+5;});
    doc.setDrawColor(...LIGHT);doc.line(18,ny+3,192,ny+3);ny+=10;
    doc.setFontSize(7);doc.setTextColor(...MID);doc.text(`Archivo analizado: ${currentFileName}`,18,ny,{maxWidth:174});

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
