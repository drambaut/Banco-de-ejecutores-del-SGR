// app.js — lógica de la plataforma SAR: fetch a la API y render en el DOM.

const API = "";

function switchTab(tabId, event) {
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  if (event) event.currentTarget.classList.add('active');
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + tabId).classList.add('active');
  if (tabId === 'descriptivo') cargarDescriptivo();
}

function fmtNumero(n) {
  if (n === null || n === undefined) return '—';
  return new Intl.NumberFormat('es-CO').format(n);
}

function fmtMoneda(n) {
  if (n === null || n === undefined) return '—';
  if (n >= 1e12) return '$' + (n / 1e12).toFixed(1) + 'B';
  if (n >= 1e9) return '$' + (n / 1e9).toFixed(1) + 'MM';
  if (n >= 1e6) return '$' + (n / 1e6).toFixed(1) + 'M';
  return '$' + fmtNumero(n);
}

function colorNivel(nivel) {
  return { 'Bajo': 'var(--risk-low)', 'Medio': 'var(--risk-med)',
           'Alto': 'var(--risk-high)', 'Crítico': 'var(--risk-critical)' }[nivel] || '#888';
}

// ============================ PERFIL DE RIESGO ============================

async function buscarPerfil() {
  const codigo = document.getElementById('input-codigo').value.trim();
  const msjError = document.getElementById('mensaje-error');
  msjError.style.display = 'none';
  if (!codigo) return;

  try {
    const resp = await fetch(`${API}/api/perfil/${encodeURIComponent(codigo)}`);
    if (!resp.ok) {
      const err = await resp.json();
      msjError.textContent = err.error || 'No se pudo calcular el perfil.';
      msjError.style.display = 'block';
      document.getElementById('results-section').style.display = 'none';
      return;
    }
    const datos = await resp.json();
    renderPerfil(datos);
  } catch (e) {
    msjError.textContent = 'Error de conexión con el servidor.';
    msjError.style.display = 'block';
  }
}

function renderPerfil(datos) {
  document.getElementById('results-section').style.display = 'flex';

  document.getElementById('res-titulo').textContent = `Perfil de riesgo — ${datos.ejecutor.nombre}`;
  document.getElementById('res-subtitulo').textContent =
    `Código ejecutor: ${datos.ejecutor.codigo_ejecutor} · Metodología ICS v2 (percentil + tau 10%)`;

  document.getElementById('ent-nombre').textContent = datos.ejecutor.nombre;
  document.getElementById('ent-codigo').textContent =
    `Código: ${datos.ejecutor.codigo_ejecutor} · NIT: ${datos.ejecutor.nit || '—'} · ${datos.ejecutor.departamento}`;
  document.getElementById('ent-proyectos').textContent = `${datos.ejecutor.total_proyectos} proyectos`;
  document.getElementById('ent-tipo').textContent = datos.ejecutor.tipo_ejecutor || '—';
  document.getElementById('ent-region').textContent = datos.ejecutor.region || '—';
  document.getElementById('ent-grupo').textContent = datos.ejecutor.tipo_ejecutor
    ? `Grupo ${datos.perfil_riesgo ? datos.perfil_riesgo.grupo_capacidad_institucional : '—'}`
    : '—';

  const pr = datos.perfil_riesgo;
  if (pr) {
    document.getElementById('gauge-pct').textContent = pr.puntaje;
    document.getElementById('gauge-pct').style.color = colorNivel(pr.nivel_4_bandas);
    document.getElementById('gauge-level').textContent = pr.nivel_4_bandas;
    document.getElementById('gauge-level').style.color = colorNivel(pr.nivel_4_bandas);

    // Aguja del gauge: mapea puntaje 0-100 sobre el arco de la media luna.
    // anguloInicio=180 y anguloFin=360 son los extremos reales del arco de fondo
    // en el SVG (centro 100,100 · radio 80), expresados sin cruzar 0° para que la
    // interpolación lineal pase por 270° (arriba) en el punto medio.
    const anguloInicio = 180, anguloFin = 360;
    const angulo = anguloInicio + (anguloFin - anguloInicio) * (pr.puntaje / 100);
    const rad = angulo * Math.PI / 180;
    const cx = 100, cy = 100, largo = 68;
    const x2 = cx + largo * Math.cos(rad);
    const y2 = cy + largo * Math.sin(rad);
    const needle = document.getElementById('gauge-needle');
    needle.setAttribute('x2', x2.toFixed(1));
    needle.setAttribute('y2', y2.toFixed(1));
    needle.setAttribute('stroke', colorNivel(pr.nivel_4_bandas));

    document.getElementById('alert-explicacion').innerHTML =
      `<strong>${pr.nivel_4_bandas}</strong> — TBC de ${(pr.tbc*100).toFixed(1)}%, ` +
      `Factor de Carga ${pr.fc.toFixed(2)} (${pr.n_proyectos} proyectos activos), ` +
      `Penalización ${(pr.pen*100).toFixed(0)}% (${pr.reprogramaciones_no_permitidas} reprogramaciones no permitidas).`;

    document.getElementById('comp-tbc').textContent = (pr.tbc * 100).toFixed(1) + '%';
    document.getElementById('comp-tbc-sub').textContent = 'periodos cumplidos / evaluados';
    document.getElementById('comp-fc').textContent = pr.fc.toFixed(2);
    document.getElementById('comp-fc-sub').textContent = `${pr.n_proyectos} proyectos activos`;
    document.getElementById('comp-pen').textContent = pr.pen.toFixed(2);
    document.getElementById('comp-pen-sub').textContent = `-${pr.descuento_pct_por_reprogramacion}% por reprogramaciones`;
  } else {
    document.getElementById('gauge-pct').textContent = 'N/D';
    document.getElementById('gauge-level').textContent = 'Sin datos';
    document.getElementById('alert-explicacion').textContent =
      'Este ejecutor no tiene suficientes periodos evaluables para calcular el ICS.';
  }

  renderCapacidades(datos.capacidades);
  renderComparables(datos.comparables);
}

function renderCapacidades(cap) {
  const cont = document.getElementById('cap-grid');
  const bloques = [
    { key: 'administrativa', nombre: 'Cap. Administrativa', claseHeader: 'cap-header-admin', claseNombre: 'cap-name-admin', claseScore: 'score-admin', claseBar: 'bar-admin' },
    { key: 'financiera', nombre: 'Cap. Financiera', claseHeader: 'cap-header-fin', claseNombre: 'cap-name-fin', claseScore: 'score-fin', claseBar: 'bar-fin' },
    { key: 'institucional', nombre: 'Cap. Institucional', claseHeader: 'cap-header-inst', claseNombre: 'cap-name-inst', claseScore: 'score-inst', claseBar: 'bar-inst' },
  ];

  cont.innerHTML = bloques.map(b => {
    const info = cap[b.key];
    const disponible = info && info.disponible;
    const score = disponible ? info.score : '—';
    const ancho = disponible ? info.score : 0;
    const variasFilas = disponible
      ? info.variables.map(v => `
          <div class="cap-var-row">
            <span class="var-name">${v.nombre}</span>
            <span class="var-pts ${v.puntos === null ? 'zero' : ''}">${v.puntos === null ? '—' : v.puntos}</span>
          </div>`).join('')
      : `<div class="cap-var-row"><span class="var-name">${info ? info.nota : 'Sin datos disponibles'}</span></div>`;
    const tag = disponible
      ? (info.score >= 66 ? '<span class="cap-level-tag tag-alta">Alta</span>'
        : info.score >= 33 ? '<span class="cap-level-tag tag-media">Media</span>'
        : '<span class="cap-level-tag tag-baja">Baja</span>')
      : '<span class="cap-level-tag tag-baja">Sin datos</span>';

    return `
      <div class="cap-card">
        <div class="cap-header ${b.claseHeader}">
          <span class="cap-name ${b.claseNombre}">${b.nombre}</span>
          <span class="cap-score ${b.claseScore}">${score}</span>
        </div>
        <div class="cap-body">
          <div class="cap-bar-wrap"><div class="cap-bar ${b.claseBar}" style="width:${ancho}%"></div></div>
          <div class="cap-vars">${variasFilas}</div>
          ${tag}
        </div>
      </div>`;
  }).join('');
}

function renderComparables(comparables) {
  const cont = document.getElementById('alt-grid');
  if (!comparables || comparables.length === 0) {
    cont.innerHTML = '<p style="font-size:13px;color:var(--text3);">No hay otras entidades en el mismo departamento con ICS calculado.</p>';
    return;
  }
  // se muestran la 2da y 3ra (si existen), tal como se pidió
  const aMostrar = comparables.slice(1, 3).length ? comparables.slice(1, 3) : comparables.slice(0, 2);
  cont.innerHTML = aMostrar.map(c => `
    <div class="alt-card" onclick="buscarPorCodigoDirecto('${c.codigo_para_buscar}')" style="cursor:pointer;">
      <div class="alt-info">
        <div class="alt-dept">${c.etiqueta}</div>
        <div class="alt-name">${c.nombre}</div>
      </div>
      <div style="display:flex;align-items:center;">
        <div class="alt-score-wrap">
          <div class="alt-score" style="color:${colorNivel(c.nivel.replace('Riesgo ',''))}">${c.puntaje}</div>
          <div class="alt-score-lbl" style="color:${colorNivel(c.nivel.replace('Riesgo ',''))}">${c.nivel}</div>
        </div>
      </div>
    </div>`).join('');
}

function buscarPorCodigoDirecto(codigo) {
  if (!codigo) return;
  document.getElementById('input-codigo').value = codigo;
  buscarPerfil();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// autocompletar mientras se escribe
let debounceTimer;
document.addEventListener('DOMContentLoaded', () => {
  const input = document.getElementById('input-codigo');
  if (input) {
    input.addEventListener('input', () => {
      clearTimeout(debounceTimer);
      const q = input.value.trim();
      if (q.length < 2) return;
      debounceTimer = setTimeout(async () => {
        const resp = await fetch(`${API}/api/buscar?q=${encodeURIComponent(q)}`);
        const datos = await resp.json();
        const lista = document.getElementById('lista-sugerencias');
        lista.innerHTML = datos.resultados
          .map(r => `<option value="${r.codigo_ejecutor}">${r.nombre || ''}</option>`)
          .join('');
      }, 300);
    });
  }
  cargarFiltrosDescriptivo();
});

// ============================ ANÁLISIS DESCRIPTIVO ============================

async function cargarFiltrosDescriptivo() {
  const resp = await fetch(`${API}/api/departamentos`);
  const datos = await resp.json();
  const sel = document.getElementById('filtro-departamento');
  datos.departamentos.forEach(d => {
    const opt = document.createElement('option');
    opt.value = d; opt.textContent = d;
    sel.appendChild(opt);
  });
}

async function cargarDescriptivo() {
  const tipo = document.getElementById('filtro-tipo').value;
  const region = document.getElementById('filtro-region').value;
  const departamento = document.getElementById('filtro-departamento').value;

  const params = new URLSearchParams({ tipo_ejecutor: tipo, region, departamento });
  const resp = await fetch(`${API}/api/descriptivo?${params}`);
  const d = await resp.json();

  document.getElementById('kpi-ejecutores').textContent = fmtNumero(d.total_ejecutores);
  document.getElementById('kpi-proyectos').textContent = fmtNumero(d.total_proyectos);
  document.getElementById('kpi-valor').textContent = fmtMoneda(d.valor_total_proyectos);
  document.getElementById('kpi-promedio').textContent = d.puntaje_promedio;
  document.getElementById('kpi-alto-critico').textContent = d.pct_alto_critico + '%';
  document.getElementById('kpi-alto-critico-sub').textContent =
    `${d.conteo_4_bandas.Alto + d.conteo_4_bandas['Crítico']} ejecutores`;

  // Histograma (bins de 10)
  const maxBin = Math.max(...d.histograma_bins_10);
  const coloresBin = ['var(--risk-low)','var(--risk-low)','var(--risk-low)','#7ccf5b','var(--risk-med)',
                      'var(--risk-med)','var(--risk-high)','var(--risk-high)','var(--risk-critical)','var(--risk-critical)'];
  const etiquetasBin = ['0-10','10-20','20-30','30-40','40-50','50-60','60-70','70-80','80-90','90-100'];
  document.getElementById('histo-bars').innerHTML = d.histograma_bins_10.map((v, i) => `
    <div class="ad-histo-col">
      <div class="ad-histo-count">${v}</div>
      <div class="ad-histo-bar" style="height:${maxBin ? (v/maxBin*100) : 0}%;background:${coloresBin[i]};"></div>
      <div class="ad-histo-label">${etiquetasBin[i]}</div>
    </div>`).join('');

  document.getElementById('chips-4-bandas').innerHTML = Object.entries(d.conteo_4_bandas).map(([nivel, n]) => `
    <div class="ad-chip" style="background:${colorNivel(nivel)}22;color:${colorNivel(nivel)};">
      <span class="ad-chip-count">${n}</span> ${nivel}
    </div>`).join('');

  renderHbarList('hbar-tipo', d.promedio_riesgo_por_tipo);
  renderHbarList('hbar-region', d.promedio_riesgo_por_region);
  renderDonut(d.estado_proyectos);
  cargarRanking();
}

function renderHbarList(idContenedor, datosObj) {
  const entradas = Object.entries(datosObj).sort((a, b) => b[1] - a[1]);
  const cont = document.getElementById(idContenedor);
  cont.innerHTML = entradas.map(([nombre, valor]) => {
    const nivel = valor < 30 ? 'Bajo' : valor < 60 ? 'Medio' : valor < 85 ? 'Alto' : 'Crítico';
    return `
      <div class="ad-hbar-item">
        <div class="ad-hbar-name">${nombre}</div>
        <div class="ad-hbar-track"><div class="ad-hbar-fill" style="width:${valor}%;background:${colorNivel(nivel)};"><span>${valor}</span></div></div>
      </div>`;
  }).join('');
}

function renderDonut(estadoProyectos) {
  const total = Object.values(estadoProyectos).reduce((a, b) => a + b, 0);
  const coloresEstado = {
    'En Ejecución': 'var(--dnp-blue)', 'Terminado': 'var(--risk-low)',
    'Sin Contratar': 'var(--risk-med)', 'Sin Migrar': 'var(--text3)',
  };
  const radio = 65, circunferencia = 2 * Math.PI * radio;
  let acumulado = 0;
  const arcos = Object.entries(estadoProyectos).map(([estado, n]) => {
    const frac = n / total;
    const largo = frac * circunferencia;
    const offset = -acumulado * circunferencia;
    acumulado += frac;
    const color = coloresEstado[estado] || '#999';
    return `<circle cx="85" cy="85" r="${radio}" fill="none" stroke="${color}" stroke-width="24"
              stroke-dasharray="${largo} ${circunferencia}" stroke-dashoffset="${offset}" transform="rotate(-90 85 85)"/>`;
  }).join('');

  document.getElementById('donut-chart').innerHTML = `
    <svg width="170" height="170" viewBox="0 0 170 170">
      <circle cx="85" cy="85" r="${radio}" fill="none" stroke="#e2e6ec" stroke-width="24"/>
      ${arcos}
      <text x="85" y="82" text-anchor="middle" font-size="20" font-weight="800" fill="var(--dnp-navy)" font-family="'IBM Plex Mono',monospace">${fmtNumero(total)}</text>
      <text x="85" y="100" text-anchor="middle" font-size="10" fill="var(--text3)" font-family="'IBM Plex Sans',sans-serif" font-weight="600">PROYECTOS</text>
    </svg>`;

  document.getElementById('donut-legend').innerHTML = Object.entries(estadoProyectos).map(([estado, n]) => `
    <div class="ad-donut-item">
      <div class="ad-donut-dot" style="background:${coloresEstado[estado] || '#999'};"></div>
      <span class="ad-donut-lbl">${estado}</span>
      <span class="ad-donut-val">${fmtNumero(n)} <span style="font-size:10px;color:var(--text3);font-weight:500;">${(n/total*100).toFixed(0)}%</span></span>
    </div>`).join('');
}

async function cargarRanking() {
  const resp = await fetch(`${API}/api/ranking`);
  const d = await resp.json();

  const filaHtml = (f, i, critico) => `
    <tr>
      <td style="font-weight:700;color:${critico ? 'var(--risk-critical)' : 'var(--risk-low)'};">${i + 1}</td>
      <td class="ad-rank-name">${f.nombre_ejecutor}</td>
      <td>${f.tipo_ejecutor || '—'}</td>
      <td>${f.region || '—'}</td>
      <td><span class="ad-rank-score" style="color:${critico ? 'var(--risk-critical)' : 'var(--risk-low)'};">${f.puntaje_riesgo.toFixed(1)}</span></td>
    </tr>`;

  document.querySelector('#tabla-mejores tbody').innerHTML =
    d.mejores.map((f, i) => filaHtml(f, i, false)).join('');
  document.querySelector('#tabla-peores tbody').innerHTML =
    d.peores.map((f, i) => filaHtml(f, i, true)).join('');
}
