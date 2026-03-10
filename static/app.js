// static/app.js
const drop = document.getElementById('drop');
const fileInput = document.getElementById('fileInput');
const list = document.getElementById('list');
const logEl = document.getElementById('log');
const docIdEl = document.getElementById('docId');
const qCountEl = document.getElementById('qCount');
const btnExport = document.getElementById('btnExport');
const categoriaEl = document.getElementById('categoria');

let state = {
  docId: null,
  questoes: [],
  remover: new Set(), // "U#Q#"
};

function log(msg){ logEl.textContent += msg + "\n"; }

function humanId(q){
  return `U${q.unidade_num}Q${q.questao_num}`;
}

function renderQuestoes(){
  list.innerHTML = '';
  state.questoes.forEach(q => {
    const id = humanId(q);
    const checked = state.remover.has(id) ? 'checked' : '';
    const gab = q.gabarito ? q.gabarito : '—';

    const altsHTML = ['A','B','C','D','E'].map(L => {
      const txt = (q.alternativas?.[L] || '').trim() || '<i class="muted">[vazio]</i>';
      return `
        <div class="alt-grid">
          <div><span class="badge text-bg-secondary">${L}</span></div>
          <div>${txt}</div>
        </div>
      `;
    }).join('');

    const card = document.createElement('div');
    card.className = 'card mb-3';
    card.innerHTML = `
      <div class="card-body">
        <div class="d-flex align-items-center gap-2 mb-2">
          <input class="form-check-input me-2" type="checkbox" data-id="${id}" ${checked}>
          <h5 class="card-title mb-0">${q.titulo}</h5>
          <span class="badge text-bg-info badge-gab ms-2">Gabarito: ${gab}</span>
          <span class="badge text-bg-dark badge-gab ms-1 mono">${id}</span>
        </div>
        <div class="mb-2"><div class="muted small">Enunciado</div>${q.enunciado || '<i class="muted">[vazio]</i>'}</div>
        <div class="mt-3">
          <div class="muted small mb-1">Alternativas</div>
          ${altsHTML}
        </div>
      </div>
    `;
    list.appendChild(card);
  });

  // bind checkboxes
  list.querySelectorAll('input[type=checkbox]').forEach(chk => {
    chk.addEventListener('change', (e) => {
      const id = e.target.getAttribute('data-id');
      if (e.target.checked) state.remover.add(id);
      else state.remover.delete(id);
    });
  });
}

function enableExport(can){ btnExport.disabled = !can; }

drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', (e)=>{ e.preventDefault(); drop.classList.add('dragover'); });
drop.addEventListener('dragleave', ()=> drop.classList.remove('dragover'));
drop.addEventListener('drop', (e)=>{
  e.preventDefault();
  drop.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if(file) handleFile(file);
});

fileInput.addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(file) handleFile(file);
});

async function handleFile(file){
  log(`Recebido: ${file.name} (${Math.round(file.size/1024)} KB)`);
  const fd = new FormData();
  fd.append('file', file);

  const res = await fetch('/api/parse', { method: 'POST', body: fd });
  const data = await res.json();
  if(!data.ok){
    log(`ERRO: ${data.error || 'Falha no parse'}`);
    enableExport(false);
    return;
  }

  state.docId = data.doc_id;
  state.questoes = data.questoes || [];
  state.remover = new Set();
  docIdEl.textContent = data.doc_id;
  qCountEl.textContent = data.count;
  renderQuestoes();
  enableExport(true);
  log(`Parse OK — ${data.count} questões carregadas.`);
}

btnExport.addEventListener('click', async ()=>{
  if(!state.docId){ return; }
  const marks = Array.from(state.remover);
  log(`Exportando XML com {REMOVER} em: [${marks.join(', ')}]`);

  const payload = {
    doc_id: state.docId,
    mark_remove: marks,
    categoria: categoriaEl.value || undefined
  };

  const res = await fetch('/api/export', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if(res.ok){
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'moodle.xml';
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
    log('XML baixado com sucesso.');
  } else {
    const err = await res.json().catch(()=>({error:'Falha desconhecida'}));
    log(`ERRO ao exportar: ${err.error || res.statusText}`);
  }
});
