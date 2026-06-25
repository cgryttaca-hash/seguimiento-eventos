let dataRows = [];
let currentFilter = 'Actuales';
let selectedCompany = null;
let chartMes = null, chartAnio = null, chartEstado = null;

function showToast(message, type='success'){
  const container = document.getElementById('toastContainer');
  const t = document.createElement('div');
  t.className = 'toast ' + (type==='error'?'error':'success');
  t.textContent = message;
  container.appendChild(t);
  setTimeout(()=>{ if(t.parentNode) t.remove(); }, 3000);
}

function startOfDay(d){ const x = new Date(d); x.setHours(0,0,0,0); return x; }
function endOfDay(d){ const x = new Date(d); x.setHours(23,59,59,999); return x; }

function parseExcelDate(v){
  if(v === undefined || v === null || v === '') return null;
  if (typeof v === 'number') {
    const excelDate = XLSX.SSF.parse_date_code(v);
    return new Date(excelDate.y, excelDate.m - 1, excelDate.d);
  }
  if(v instanceof Date && !isNaN(v)) return startOfDay(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m) return startOfDay(new Date(+m[3], +m[2]-1, +m[1]));
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if(m) return startOfDay(new Date(+m[1], +m[2]-1, +m[3]));
  const d = new Date(s);
  return !isNaN(d) ? startOfDay(d) : null;
}

function parseInputDate(value){
  if(!value) return null;
  const [y,m,d] = value.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function normalizeEstado(v){
  if(!v) return '';
  const s = String(v).toLowerCase();
  if(s.includes('confirmado')) return 'Confirmado';
  if(s.includes('tentativo') || s.includes('pendiente')) return 'Tentativo o pendiente';
  if(s.includes('program')) return 'Programado';
  if(s.includes('curso')) return 'En curso';
  if(s.includes('ejecut')) return 'Ejecutado';
  if(s.includes('final')) return 'Finalizado';
  if(s.includes('cancel')) return 'Cancelado';
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function formatDate(d){
  if(!(d instanceof Date)) return '';
  return d.toLocaleDateString('es-ES',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});
}

function escapeHtml(s){
  if(!s && s!==0) return '';
  return String(s).replace(/[&<>"'`=\/]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;','`':'&#96;','=':'&#61;','/':'&#47;'}[c]));
}

function highlightText(text){
  if(!text) return '';
  return text
    .replace(/(tercer piso)/gi, '<span class="hl-tercer">Tercer piso</span>')
    .replace(/(salon\s*\d*|\bsalon\b)/gi, '<span class="hl-salon">$1</span>');
}

function renderTable(rows){
  const tbody = document.getElementById('tbody'); tbody.innerHTML = '';
  if(dataRows.length === 0){ showToast('Carga un archivo para continuar', 'error'); return; }
  if(rows.length === 0){ showToast('No hay eventos que coincidan con los filtros', 'info'); return; }

  rows.forEach(r => {
    const tr = document.createElement('tr');
    if(r.FECHA instanceof Date){
      const d = startOfDay(r.FECHA);
      const today = startOfDay(new Date());
      const diff = Math.round((d - today) / 86400000);
      if(diff < 0) tr.style.backgroundColor = '#e5e7eb', tr.style.color = '#6b7280';
      else if(diff === 0) tr.style.backgroundColor = '#fff3cd', tr.style.fontWeight = '700';
      else if(diff <= 3) tr.style.backgroundColor = '#dbeafe';
      else tr.style.backgroundColor = '#dcfce7';
    }

    tr.innerHTML = `
<td>${formatDate(r.FECHA)}</td>
<td>${highlightText(escapeHtml(r['ESCENARIO ASIGNADO']))}</td>
<td>${escapeHtml(r['HORARIO DEL EVENTO'])}</td>
<td>${escapeHtml(r['NOMBRE DE LA EMPRESA'])}</td>
<td>${escapeHtml(r['CANTIDAD DE PERSONAS'])}</td>
<td class="horario-ayb">${escapeHtml(r['HORARIO AYB'] || '').replace(/\r?\n/g,'<br>')}</td>
<td class="desc-food">${escapeHtml(r['DESCRIPCION ALIMENTACION'] || '').replace(/\r?\n/g,'<br>')}</td>
<td>${escapeHtml(r['ACOMODACION'])}</td>
<td>${escapeHtml(r['MEDIO DE PAGO'])}</td>
<td>${escapeHtml(r['OBSERVACION'])}</td>
<td>${escapeHtml(r['ESTADO'])}</td>
`;
    tbody.appendChild(tr);
  });
}

function updateDashboard(rows){
  const dash = document.getElementById('dashboard');
  dash.innerHTML = '';
  if(!rows.length){
    dash.innerHTML = `<div class="card"><div class="label">Sin datos</div><div class="value">0</div></div>`;
    return;
  }
  const total = rows.length;
  const pax = rows.reduce((s, r) => s + (parseInt(r['CANTIDAD DE PERSONAS']) || 0), 0);
  const cancelados = rows.filter(x => (x.ESTADO || '').trim() === 'Cancelado').length;

  dash.innerHTML += `<div class="card total"><div class="label">Total eventos</div><div class="value">${total}</div></div>`;
  dash.innerHTML += `<div class="card pax"><div class="label">Total personas</div><div class="value">${pax}</div></div>`;
  dash.innerHTML += `<div class="card cancelado"><div class="label">Cancelados</div><div class="value">${cancelados}</div></div>`;
}

function destroyCharts(){ [chartMes,chartAnio,chartEstado].forEach(c=>c?.destroy()); chartMes=chartAnio=chartEstado=null; }

function updateCharts(rows){
  const cont = document.getElementById('chartContainer');
  cont.style.display = currentFilter === 'Grafica' ? 'grid' : 'none';
  destroyCharts(); if(cont.style.display === 'none') return;
  try{ if(Chart && ChartDataLabels) Chart.register(ChartDataLabels); }catch(e){}

  const xMes={}, xAnio={}, xEstado={}, xEmp={};
  rows.forEach(r=>{
    if(!(r.FECHA instanceof Date)) return;
    const esc = (r['ESCENARIO ASIGNADO']||'').toLowerCase();
    if(esc.includes('tercer piso')) return;
    const m = r.FECHA.toLocaleString('es-ES',{month:'short'});
    const y = r.FECHA.getFullYear();
    xMes[m]=(xMes[m]||0)+1; xAnio[y]=(xAnio[y]||0)+1;
    const e = r.ESTADO||'Sin estado'; xEstado[e]=(xEstado[e]||0)+1;
    const n = (r['NOMBRE DE LA EMPRESA']||'Desconocida').trim();
    if(n) xEmp[n]=(xEmp[n]||0)+1;
  });

  const ordMes = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const datosMes = Object.entries(xMes).sort((a,b)=>ordMes.indexOf(a[0].slice(0,3)) - ordMes.indexOf(b[0].slice(0,3)));
  chartMes = new Chart(document.getElementById('chartMes'),{
    type:'bar', data:{labels:datosMes.map(e=>e[0]),datasets:[{label:'Eventos',data:datosMes.map(e=>e[1]),backgroundColor:'#2563eb'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{title:{display:true,text:'📅 Por mes'},legend:{display:false},datalabels:{display:true}}}
  });

  const datosAnio = Object.entries(xAnio).sort((a,b)=>a[0]-b[0]);
  chartAnio = new Chart(document.getElementById('chartAnio'),{
    type:'line', data:{labels:datosAnio.map(e=>e[0]),datasets:[{label:'Eventos',data:datosAnio.map(e=>e[1]),borderColor:'#f59e0b',backgroundColor:'#f59e0b33',tension:0.3,fill:true}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{title:{display:true,text:'📊 Por año'},legend:{display:false},datalabels:{display:true}}}
  });

  chartEstado = new Chart(document.getElementById('chartEstado'),{
    type:'pie', data:{labels:Object.keys(xEstado),datasets:[{data:Object.values(xEstado),backgroundColor:['#fef3c7','#ffedd5','#fde68a','#dcfce7','#fee2e2','#e2e8f0']}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{title:{display:true,text:'🎯 Por estado'},datalabels:{display:true}}}
  });

  const list = document.getElementById('empresasList'); list.innerHTML='';
  const top = Object.entries(xEmp).sort((a,b)=>b[1]-a[1]).slice(0,5);
  top.forEach(([n,v])=>{
    const li = document.createElement('li');
    li.innerHTML = `<span>${escapeHtml(n)}</span> ${v}`;
    li.onclick=()=>{selectedCompany=selectedCompany===n?null:n;applyFilters();};
    if(selectedCompany===n) li.classList.add('active');
    list.appendChild(li);
  });
}

function applyFilters(){
  const s = document.getElementById('search').value.trim().toLowerCase();
  const df = document.getElementById('dateFrom').value ? startOfDay(parseInputDate(document.getElementById('dateFrom').value)) : null;
  const dt = document.getElementById('dateTo').value ? endOfDay(parseInputDate(document.getElementById('dateTo').value)) : null;
  const hoy = startOfDay(new Date());
  const hideNoFood = document.getElementById('hideNoFood')?.checked;

  let rows = dataRows.filter(r=>{
    let ok = true;
    const fechaOk = r.FECHA instanceof Date;

    if(s){
      const emp = (r['NOMBRE DE LA EMPRESA']||'').toLowerCase();
      const esc = (r['ESCENARIO ASIGNADO']||'').toLowerCase();
      ok = ok && (emp.includes(s) || esc.includes(s));
    }
    if(df) ok = ok && fechaOk && r.FECHA >= df;
    if(dt) ok = ok && fechaOk && r.FECHA <= dt;

    if(currentFilter === 'Actuales') ok = ok && fechaOk && startOfDay(r.FECHA).getTime() >= hoy.getTime();
    if(currentFilter === 'Segundo piso'){
      const esc = (r['ESCENARIO ASIGNADO'] || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      ok = ok && esc.includes('salon') && !esc.includes('tercer piso');
    }
    if(currentFilter === 'Tercer piso'){
      const esc = (r['ESCENARIO ASIGNADO'] || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      ok = ok && esc.includes('tercer piso');
    }
    if(selectedCompany) ok = ok && r['NOMBRE DE LA EMPRESA'] === selectedCompany;
    if(hideNoFood){
      const ayb = String(r['HORARIO AYB'] || '').trim().toUpperCase();
      if(ayb.includes('SIN ALIMENTACION')) ok = false;
    }
    return ok;
  });

  renderTable(rows);
  document.getElementById('graficaAviso').style.display = currentFilter === 'Grafica' ? 'block' : 'none';
  updateDashboard(rows);
  updateCharts(rows);
  document.getElementById('captionFiltro').textContent = `Mostrando ${rows.length} eventos (filtro: ${currentFilter})`;
}

// Carga y procesamiento de archivos Excel
document.getElementById('fileInput').addEventListener('change', e => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = new Uint8Array(ev.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      let raw = [];
      const meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];
      const mesActual = meses[new Date().getMonth()];
      const hojasMes = wb.SheetNames.filter(nombre => meses.some(m => nombre.trim().toUpperCase().includes(m)));
      hojasMes.sort((a,b) => a.toUpperCase().includes(mesActual) ? -1 : b.toUpperCase().includes(mesActual) ? 1 : 0);

      hojasMes.forEach(nombreHoja => {
        const sheet = wb.Sheets[nombreHoja];
        const datos = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
        if (datos.length <= 1) return;
        raw = raw.length === 0 ? datos : [...raw, ...datos.slice(1)];
      });

      const headers = raw[0].map(h => String(h || '').trim().toUpperCase().replace(/\s+/g, ' '));
      const json = raw.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i] ?? '');
        return obj;
      });

      dataRows = json.map(r => ({
        ...r,
        FECHA: parseExcelDate(r['FECHA'] || r['FECHA DEL EVENTO'] || r['FECHA EVENTO'] || r['FECHA_EVENTO'] || r['DIA']),
        ESTADO: normalizeEstado(r['ESTADO'] || r['STATUS'])
      }));

      localStorage.setItem('eventData', JSON.stringify(dataRows));
      showToast('Archivo cargado correctamente');
      selectedCompany = null;
      currentFilter = 'Actuales';
      document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      document.querySelector('.chip[data-filter="Actuales"]').classList.add('active');
      applyFilters();
    } catch (err) {
      console.error(err);
      showToast('Error al leer el archivo', 'error');
    }
  };
  reader.readAsArrayBuffer(f);
  e.target.value = '';
});

// Eventos de interfaz
document.getElementById('search').addEventListener('input', applyFilters);
document.getElementById('hideNoFood')?.addEventListener('change', applyFilters);
document.getElementById('dateFrom').addEventListener('change', applyFilters);
document.getElementById('dateTo').addEventListener('change', applyFilters);
document.getElementById('resetBtn').addEventListener('click',()=>{
  document.getElementById('search').value='';
  document.getElementById('dateFrom').value='';
  document.getElementById('dateTo').value='';
  const chk=document.getElementById('hideNoFood');
  if(chk) chk.checked=true;
  selectedCompany=null; currentFilter='Actuales';
  document.querySelectorAll('.chip').forEach(c=>c.classList.remove('active'));
  document.querySelector('.chip[data-filter="Actuales"]').classList.add('active');
  applyFilters();
});

document.querySelectorAll('.chip').forEach(c=>c.addEventListener('click',()=>{
  document.querySelectorAll('.chip').forEach(x=>x.classList.remove('active'));
  c.classList.add('active'); currentFilter=c.dataset.filter; applyFilters();
}));

window.addEventListener('DOMContentLoaded',()=>{
  const guardado = localStorage.getItem('eventData');
  if(guardado){
    try{
      dataRows = JSON.parse(guardado).map(r=>({...r, FECHA:r.FECHA?parseExcelDate(r.FECHA):null}));
      applyFilters(); showToast('Datos cargados del historial');
    }catch(e){ localStorage.removeItem('eventData'); }
  } else renderTable([]);
});

document.getElementById('btnLoad').addEventListener('click', () => {
  document.getElementById('fileInput').click();
});
