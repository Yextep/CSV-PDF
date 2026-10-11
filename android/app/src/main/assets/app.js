'use strict';
const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];
const icon = (name) => `<svg aria-hidden="true"><use href="#i-${name}"/></svg>`;
const esc = (text) => String(text ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let state = {ready:false,busy:false,selection:[],history:[],active:null,themeMode:'system'};
let page = 'convert', kind = 'both', modalSession = null, previousFocus = null, toastTimer;
function send(action, data = {}) {
  if (window.LumaNative) LumaNative.request(action, JSON.stringify(data));
  else if (action === 'theme') applyTheme(data.mode, data.mode === 'system' ? (matchMedia('(prefers-color-scheme:dark)').matches?'dark':'light') : data.mode);
  else toast('Esta acción está disponible en la app Android.');
}
function toast(message) { $('#toast').textContent = message; $('#toast').hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { $('#toast').hidden = true; }, 4500); }
function size(bytes) { return bytes >= 1048576 ? `${(bytes/1048576).toFixed(1)} MB` : `${Math.max(1,Math.round(bytes/1024))} KB`; }
function date(time) { return new Intl.DateTimeFormat('es',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}).format(new Date(time)); }
function isSheet(name) { return /\.(xlsx|xlsm|xltx|xltm|xls|xlt|xlsb|ods|ots|fods|xml)$/i.test(name); }
function selectedCount() { return state.selection.filter(f => kind === 'both' || (kind === 'excel') === isSheet(f.name)).length; }
function options() { return {kind,delimiter:$('#delimiter').value === 'tab'?'\t':$('#delimiter').value,encoding:$('#encoding').value,safe:$('#safe-csv').checked}; }
function applyTheme(mode, resolved) {
  state.themeMode = mode;
  document.documentElement.dataset.theme = resolved;
  $('#theme-toggle').innerHTML = icon(resolved === 'dark'?'sun':'moon');
  $('#theme-toggle').setAttribute('aria-label',resolved === 'dark'?'Activar modo claro':'Activar modo oscuro');
  $$('#appearance button').forEach(b => b.setAttribute('aria-pressed',String(b.dataset.theme === mode)));
}
function navigate(target) {
  if (!['convert','history','settings'].includes(target)) return;
  page = target;
  $$('.page').forEach(p => p.hidden = p.id !== `page-${target}`);
  $$('.bottom-nav button').forEach(b => b.dataset.page === target ? b.setAttribute('aria-current','page') : b.removeAttribute('aria-current'));
  window.scrollTo({top:0,behavior:'instant'});
}
function render() {
  const selected = $('#selection');
  selected.hidden = !state.selection.length;
  selected.innerHTML = `<div class="list-title"><b>${state.selection.length} ${state.selection.length===1?'archivo seleccionado':'archivos seleccionados'}</b><button data-clear ${state.busy?'disabled':''}>Vaciar</button></div><div>${state.selection.map(f => `<div class="file-row"><span class="file-icon ${isSheet(f.name)?'mint':''}">${icon(isSheet(f.name)?'grid':'file')}</span><div class="file-copy"><b>${esc(f.name)}</b><small>${size(f.bytes)} · ${isSheet(f.name)?'Hoja → CSV':'Texto → PDF'}${f.relative && f.relative !== f.name?' · '+esc(f.relative):''}</small></div><button class="remove-file" data-remove="${esc(f.id)}" aria-label="Quitar ${esc(f.name)}" ${state.busy?'disabled':''}>${icon('close')}</button></div>`).join('')}</div>`;
  const count = selectedCount();
  $('#convert').disabled = !state.ready || state.busy || count === 0;
  $('#convert span').textContent = !state.ready?'Preparando Luma…':state.busy?'Trabajando en tus archivos…':count?`Convertir ${count} ${count===1?'archivo':'archivos'}`:'Convertir archivos';
  ['#pick-files','#pick-folder','#demo','#clear-history'].forEach(s => $(s).disabled = state.busy || !state.ready);
  $$('#kind-select button,#delimiter,#encoding,#safe-csv').forEach(b => b.disabled = state.busy);
  const progress = $('#progress');
  progress.hidden = !state.busy;
  if (state.busy) {
    const job = state.active || {};
    const percent = job.total?Math.round((job.processed||0)/job.total*100):0;
    progress.innerHTML = `<div class="progress-heading"><span><i class="busy-dot"></i>${job.label || 'Preparando archivos'}</span><span>${job.total?percent+'%':''}</span></div><div class="progress-track"><i style="width:${percent}%"></i></div><p>${esc(job.current||'Un momento, estamos preparando tu selección.')}</p>${job.total?'<button data-cancel>Detener después del archivo actual</button>':''}`;
  }
  renderHistory();
  if (modalSession) {
    const updated = state.history.find(h => h.id === modalSession);
    if (updated) showSession(updated, false); else closeModal();
  }
}
function renderHistory() {
  const outputs = state.history.reduce((n,h) => n+(h.outputCount||0),0);
  $('#history-stats').innerHTML = `<div class="stat glass"><strong>${state.history.length}</strong><span>conversiones guardadas</span></div><div class="stat glass"><strong>${outputs}</strong><span>archivos creados</span></div>`;
  $('#history-list').innerHTML = state.history.length ? state.history.map(h => `<button class="history-card glass" data-session="${esc(h.id)}"><span class="file-icon mint">${icon('check')}</span><span class="file-copy"><b>${h.outputCount||0} ${(h.outputCount||0)===1?'archivo listo':'archivos listos'}</b><small>${date(h.time)} · ${h.total} ${h.total===1?'original':'originales'}</small><span class="status ${h.errors?'error':h.interrupted||h.cancelled||h.warnings?'warn':''}">${h.interrupted?'Sesión interrumpida':h.cancelled?'Conversión detenida':h.errors?`${h.errors} ${h.errors===1?'archivo con error':'archivos con error'}`:h.warnings?'Listo · revisa los avisos':'Completado'}</span></span><svg class="arrow"><use href="#i-arrow"/></svg></button>`).join('') : `<div class="empty glass"><div class="upload-mark">${icon('clock')}</div><h3>Aquí empieza tu colección</h3><p>Cuando conviertas tus primeros archivos,<br>los encontrarás en este espacio.</p><button class="button secondary" data-navigate="convert">Crear una conversión ${icon('arrow')}</button></div>`;
}
function openModal(html, focus = true) {
  if ($('#modal-wrap').hidden) previousFocus = document.activeElement;
  $('#modal-content').innerHTML = html;
  $('#modal-wrap').hidden = false;
  document.body.style.overflow = 'hidden';
  if (focus) $('.modal').focus();
}
function closeModal() { $('#modal-wrap').hidden = true; document.body.style.overflow = ''; modalSession = null; previousFocus?.focus(); }
function showSession(h, focus = true) {
  modalSession = h.id;
  const files = h.outputs || [];
  openModal(`<span class="eyebrow">TU CONVERSIÓN</span><h2 id="modal-title">${files.length?'Listos para salir.':'Revisemos este lote.'}</h2><p class="modal-subtitle">${date(h.time)} · ${h.total} ${h.total===1?'archivo original':'archivos originales'}</p>${files.length?`<div class="result-banner">${icon('check')} ${files.length} ${files.length===1?'archivo creado':'archivos creados'}. Tus originales siguen intactos.</div><div class="result-actions"><button class="button primary" data-file-action="save" data-name="@zip">${icon('download')}Guardar ZIP</button><button class="button secondary" data-file-action="share" data-name="@zip">${icon('share')}Compartir</button></div>`:''}${h.interrupted||h.cancelled?`<div class="warning-box">${h.interrupted?'La sesión se interrumpió al cerrar la app.':'Detuviste el lote.'} Se conservan los resultados que terminaron. Añade otra vez los archivos pendientes.</div>`:''}<div class="result-group"><h3>Resultados</h3>${files.map(f => `<div class="output-row"><div class="output-name">${esc(f)}</div><div class="output-actions"><button data-file-action="open" data-name="${esc(f)}">Abrir</button><button data-file-action="save" data-name="${esc(f)}">Guardar</button><button data-file-action="share" data-name="${esc(f)}">Compartir</button></div></div>`).join('') || '<p class="body-copy">No se generaron archivos de salida.</p>'}</div>${(h.results||[]).filter(r=>r.error || (r.warnings||[]).length).map(r => `<div class="result-group"><h3>${esc(r.name)}</h3>${r.error?`<div class="error-box"><b>No se convirtió</b>${esc(r.error)}</div>`:''}${(r.warnings||[]).map(w => `<div class="warning-box">${esc(w)}</div>`).join('')}</div>`).join('')}<button class="report-link" data-file-action="save" data-name="@report">${icon('download')} Guardar reporte CSV</button><button class="button danger delete-session" data-delete-session="${esc(h.id)}">${icon('trash')}Eliminar esta conversión</button>`,focus);
}
function confirmDelete(id) {
  modalSession = null;
  openModal(`<span class="eyebrow">ESPACIO DE LUMA</span><h2 id="modal-title">¿Eliminar ${id?'esta conversión':'el historial'}?</h2><p class="body-copy">Se borrarán ${id?'sus resultados privados y su registro':'todos los resultados privados y el historial'}. Los originales y los archivos que ya guardaste fuera de Luma se conservan.</p><div class="confirm-actions"><button class="button secondary" data-dismiss>Conservar</button><button class="button primary" data-confirm-delete="${esc(id||'all')}">Eliminar</button></div>`);
}
window.Luma = {
  onEvent(type, data) {
    if (type === 'state') {
      state = data;
      if (data.options) { kind = data.options.kind || 'both'; $('#delimiter').value = data.options.delimiter === '\t'?'tab':data.options.delimiter || ','; $('#encoding').value = data.options.encoding || 'utf-8-sig'; $('#safe-csv').checked = data.options.safe !== false; }
      $$('#kind-select button').forEach(b => b.setAttribute('aria-pressed',String(b.dataset.kind === kind)));
      applyTheme(data.themeMode || 'system',data.theme || document.documentElement.dataset.theme);
      render();
    } else if (type === 'toast' || type === 'error') toast(data.message);
    else if (type === 'done') { const session = state.history.find(h => h.id === data.id); if (session) showSession(session); }
    else if (type === 'theme') applyTheme(data.mode,data.resolved);
  },
  back() { if (!$('#modal-wrap').hidden) { closeModal(); return true; } if (page !== 'convert') { navigate('convert'); return true; } return false; }
};
document.addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b || b.disabled) return;
  if (b.dataset.page || b.dataset.navigate) navigate(b.dataset.page || b.dataset.navigate);
  if (b.dataset.kind) { kind=b.dataset.kind; $$('#kind-select button').forEach(x=>x.setAttribute('aria-pressed',String(x===b))); send('options',options()); render(); }
  if (b.dataset.theme) send('theme',{mode:b.dataset.theme});
  if (b.dataset.remove) send('remove',{id:b.dataset.remove});
  if (b.hasAttribute('data-clear')) send('clearSelection');
  if (b.hasAttribute('data-cancel')) { send('cancel'); b.disabled=true; b.textContent='Deteniendo al terminar este archivo…'; }
  if (b.dataset.session) showSession(state.history.find(h=>h.id===b.dataset.session));
  if (b.dataset.fileAction) send('file',{session:modalSession,name:b.dataset.name,action:b.dataset.fileAction});
  if (b.dataset.deleteSession) confirmDelete(b.dataset.deleteSession);
  if (b.dataset.confirmDelete) { send('delete',{id:b.dataset.confirmDelete}); closeModal(); }
  if (b.hasAttribute('data-dismiss') || b.classList.contains('modal-close')) closeModal();
});
$('#pick-files').onclick=()=>send('pickFiles'); $('#pick-folder').onclick=()=>send('pickFolder'); $('#demo').onclick=()=>send('demo');
$('#convert').onclick=()=>send('convert',options());
$('#theme-toggle').onclick=()=>send('theme',{mode:document.documentElement.dataset.theme === 'dark'?'light':'dark'});
$('#clear-history').onclick=()=>confirmDelete(null);
['#delimiter','#encoding','#safe-csv'].forEach(s=>$(s).addEventListener('change',()=>send('options',options())));
$('.brand').onclick=e=>{e.preventDefault();navigate('convert');};
$('#modal-wrap').addEventListener('click',e=>{if(e.target===$('#modal-wrap'))closeModal();});
document.addEventListener('keydown',e=>{
  if(e.key==='Escape')window.Luma.back();
  if(e.key==='Tab'&&!$('#modal-wrap').hidden){const items=[...$('.modal').querySelectorAll('button,[tabindex="0"]')].filter(x=>!x.disabled);const first=items[0],last=items[items.length-1];if(e.shiftKey&&(document.activeElement===first||document.activeElement===$('.modal'))){e.preventDefault();last?.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first?.focus();}}
});
render(); send('state');
