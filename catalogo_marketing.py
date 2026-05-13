<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Catálogo de Ações de Marketing</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f5f5f3;color:#1a1a18;min-height:100vh}
.app{max-width:920px;margin:0 auto;padding:2rem 1.25rem}

/* HERO */
.hero{position:relative;border-radius:14px;overflow:hidden;margin-bottom:1.5rem;padding:28px 28px 24px}
.hero-bg{position:absolute;inset:0;background:#1a1060;z-index:0}
.hero-pattern{position:absolute;inset:0;z-index:1;opacity:.18;width:100%;height:100%}
.hero-content{position:relative;z-index:2;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
.hero-eyebrow{font-size:11px;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#9F99E8;margin-bottom:8px;display:flex;align-items:center;gap:6px}
.hero-title{font-size:26px;font-weight:500;color:#fff;line-height:1.2;margin-bottom:6px}
.hero-title span{color:#A9A3F5}
.hero-sub{font-size:13px;color:#B0ABDA;line-height:1.5}
.hero-badges{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
.hero-badge{background:rgba(255,255,255,.1);border:0.5px solid rgba(255,255,255,.2);border-radius:20px;font-size:11px;color:#D0CCFF;padding:4px 11px;display:inline-flex;align-items:center;gap:5px}
.hero-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}

/* STATS */
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:1.25rem}
.stat-card{background:#fff;border-radius:8px;padding:12px 14px;border:0.5px solid rgba(0,0,0,.1)}
.stat-label{font-size:12px;color:#5f5e5a;margin-bottom:3px}
.stat-value{font-size:22px;font-weight:500}

/* FILTERS */
.filters-bar{background:#fff;border:0.5px solid rgba(0,0,0,.1);border-radius:12px;padding:14px 16px;margin-bottom:1.25rem}
.filters-title{font-size:12px;font-weight:500;color:#5f5e5a;letter-spacing:.05em;margin-bottom:10px}
.filters-row{display:grid;grid-template-columns:2fr 1fr 1fr 1fr;gap:10px}
@media(max-width:640px){.filters-row{grid-template-columns:1fr 1fr}}
.filter-group{display:flex;flex-direction:column;gap:4px}
.filter-label{font-size:12px;color:#5f5e5a}

/* FORM ELEMENTS */
input,select,textarea{width:100%;border:0.5px solid rgba(0,0,0,.15);border-radius:8px;padding:7px 10px;font-size:13px;background:#fff;color:#1a1a18;outline:none;font-family:inherit}
input:focus,select:focus,textarea:focus{border-color:#534AB7;box-shadow:0 0 0 2px rgba(83,74,183,.12)}
textarea{resize:vertical;min-height:72px}

/* BUTTONS */
.btn{border:0.5px solid rgba(0,0,0,.2);border-radius:8px;padding:7px 13px;font-size:13px;background:transparent;color:#1a1a18;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s;font-family:inherit}
.btn:hover{background:#f1efe8}
.btn-primary{background:#534AB7;color:#fff;border-color:#534AB7}
.btn-primary:hover{background:#3C3489;border-color:#3C3489}
.btn-export{background:#0F6E56;color:#fff;border-color:#0F6E56}
.btn-export:hover{background:#085041}
.btn-sm{padding:5px 10px;font-size:12px}

/* ACTIONS HEADER */
.actions-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px}
.actions-count{font-size:13px;color:#5f5e5a}

/* CARDS */
.action-card{background:#fff;border:0.5px solid rgba(0,0,0,.1);border-radius:12px;padding:14px 16px;margin-bottom:10px}
.action-card:hover{border-color:rgba(0,0,0,.2)}
.action-top{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:start}
.action-title{font-size:15px;font-weight:500;margin-bottom:5px}
.action-meta{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:6px}
.tag{font-size:11px;padding:2px 8px;border-radius:20px;font-weight:500}
.tag-tipo{background:#EEEDFE;color:#3C3489}
.tag-empresa{background:#E1F5EE;color:#0F6E56}
.tag-canal{background:#FAEEDA;color:#633806}
.tag-ativo{background:#EAF3DE;color:#27500A}
.tag-planejado{background:#E6F1FB;color:#0C447C}
.tag-encerrado{background:#F1EFE8;color:#5F5E5A}
.tag-objetivo{background:#FBEAF0;color:#72243E}
.action-desc{font-size:13px;color:#5f5e5a;line-height:1.5;margin-bottom:5px}
.action-footer{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;color:#888780;border-top:0.5px solid rgba(0,0,0,.08);padding-top:8px;margin-top:8px;align-items:center}
.action-footer span{display:inline-flex;align-items:center;gap:4px}
.action-controls{display:flex;gap:6px;flex-shrink:0}
.icon-btn{width:30px;height:30px;border:0.5px solid rgba(0,0,0,.12);border-radius:8px;background:transparent;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#5f5e5a;font-size:14px;transition:background .15s}
.icon-btn:hover{background:#f1efe8}
.icon-btn.del:hover{background:#FCEBEB;color:#A32D2D;border-color:#A32D2D}
.resultado-box{font-size:12px;color:#5f5e5a;margin-top:4px;padding:6px 10px;background:#f5f5f3;border-radius:8px}
.attached-files{display:flex;flex-wrap:wrap;gap:5px;margin-top:5px}
.attached-chip{display:inline-flex;align-items:center;gap:4px;background:#f5f5f3;border:0.5px solid rgba(0,0,0,.1);border-radius:20px;font-size:11px;padding:2px 9px;color:#5f5e5a}

/* MODAL */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.38);z-index:100;align-items:flex-start;justify-content:center;padding:40px 16px;overflow-y:auto}
.modal-overlay.open{display:flex}
.modal{background:#fff;border-radius:12px;border:0.5px solid rgba(0,0,0,.1);padding:20px 22px;width:100%;max-width:600px;margin:auto}
.modal-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.modal-title{font-size:16px;font-weight:500;display:flex;align-items:center;gap:8px}
.modal-dot{width:8px;height:8px;border-radius:50%;background:#534AB7;display:inline-block;flex-shrink:0}
.form-section{font-size:11px;font-weight:500;color:#5f5e5a;letter-spacing:.06em;margin:14px 0 8px;padding-bottom:5px;border-bottom:0.5px solid rgba(0,0,0,.08);text-transform:uppercase}
.form-row{margin-bottom:12px}
.form-label{font-size:12px;color:#5f5e5a;margin-bottom:4px;display:block}
.form-grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.form-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
@media(max-width:520px){.form-grid2,.form-grid3{grid-template-columns:1fr}}
.modal-footer{display:flex;gap:8px;justify-content:flex-end;margin-top:16px;border-top:0.5px solid rgba(0,0,0,.08);padding-top:14px}

/* UPLOAD */
.upload-zone{border:1.5px dashed rgba(0,0,0,.18);border-radius:8px;padding:14px 16px;text-align:center;cursor:pointer;transition:background .15s,border-color .15s;background:#f9f9f7}
.upload-zone:hover,.upload-zone.drag{background:#EEEDFE;border-color:#534AB7}
.upload-zone input[type=file]{display:none}
.upload-zone-label{font-size:13px;color:#5f5e5a;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:6px;pointer-events:none}
.upload-icon{font-size:22px;color:#534AB7}
.file-list{margin-top:8px;display:flex;flex-direction:column;gap:5px}
.file-item{display:flex;align-items:center;justify-content:space-between;background:#fff;border:0.5px solid rgba(0,0,0,.1);border-radius:8px;padding:5px 10px;font-size:12px}
.file-item-name{display:flex;align-items:center;gap:6px;color:#5f5e5a}
.file-remove{background:none;border:none;cursor:pointer;color:#888780;font-size:16px;padding:0;line-height:1;display:flex;align-items:center}
.file-remove:hover{color:#A32D2D}

/* MISC */
.budget-color{color:#3C3489;font-weight:500}
.empty-state{text-align:center;padding:3rem 1rem;color:#5f5e5a}
</style>
</head>
<body>
<div class="app">

  <!-- HERO HEADER -->
  <div class="hero">
    <div class="hero-bg"></div>
    <svg class="hero-pattern" viewBox="0 0 900 140" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">
      <defs><pattern id="dots" x="0" y="0" width="28" height="28" patternUnits="userSpaceOnUse"><circle cx="1.5" cy="1.5" r="1.5" fill="#fff"/></pattern></defs>
      <rect width="900" height="140" fill="url(#dots)"/>
      <!-- megafone -->
      <polygon points="60,28 60,92 118,115 118,5" fill="none" stroke="#A9A3F5" stroke-width="2"/>
      <polygon points="118,28 118,92 168,108 168,20" fill="none" stroke="#A9A3F5" stroke-width="1.5" opacity=".5"/>
      <!-- ondas de sinal -->
      <path d="M185 60 Q208 36 208 60 Q208 84 185 60" fill="none" stroke="#9F99E8" stroke-width="2"/>
      <path d="M190 60 Q225 22 225 60 Q225 98 190 60" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".6"/>
      <path d="M196 60 Q244 6 244 60 Q244 114 196 60" fill="none" stroke="#9F99E8" stroke-width="1" opacity=".3"/>
      <!-- gráfico de barras -->
      <rect x="680" y="82" width="16" height="40" rx="3" fill="#534AB7" opacity=".6"/>
      <rect x="702" y="55" width="16" height="67" rx="3" fill="#534AB7" opacity=".8"/>
      <rect x="724" y="65" width="16" height="57" rx="3" fill="#534AB7" opacity=".5"/>
      <rect x="746" y="38" width="16" height="84" rx="3" fill="#A9A3F5" opacity=".7"/>
      <line x1="675" y1="124" x2="770" y2="124" stroke="#A9A3F5" stroke-width="1.5" opacity=".4"/>
      <!-- etiqueta de preço -->
      <rect x="300" y="16" width="72" height="30" rx="6" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".5"/>
      <circle cx="313" cy="31" r="3.5" fill="#9F99E8" opacity=".5"/>
      <line x1="295" y1="31" x2="302" y2="31" stroke="#9F99E8" stroke-width="1.5" opacity=".5"/>
      <!-- carrinho -->
      <path d="M420 95 l10-36 h62 l-10 36 z" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".35"/>
      <circle cx="435" cy="100" r="4.5" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".35"/>
      <circle cx="470" cy="100" r="4.5" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".35"/>
      <!-- estrelas -->
      <text x="510" y="52" font-size="16" fill="#A9A3F5" opacity=".4">&#9733;&#9733;&#9733;&#9733;&#9733;</text>
      <!-- seta tendência -->
      <polyline points="260,108 305,74 338,86 385,48" fill="none" stroke="#9F99E8" stroke-width="2" opacity=".4"/>
      <polygon points="385,48 374,51 380,62" fill="#9F99E8" opacity=".4"/>
      <!-- círculo percent -->
      <circle cx="820" cy="55" r="32" fill="none" stroke="#9F99E8" stroke-width="1.5" opacity=".25"/>
      <text x="808" y="62" font-size="16" fill="#A9A3F5" opacity=".4" font-weight="bold">%</text>
    </svg>

    <div class="hero-content">
      <div>
        <div class="hero-eyebrow">&#128226; Ferramenta de gestão</div>
        <div class="hero-title">Catálogo de Ações de <span>Marketing</span></div>
        <div class="hero-sub">Registro colaborativo de campanhas, promoções e iniciativas</div>
        <div class="hero-badges">
          <span class="hero-badge">&#127978; Varejo supermercadista</span>
          <span class="hero-badge">&#128101; Colaborativo</span>
          <span class="hero-badge">&#9729;&#65039; Salvo localmente</span>
        </div>
      </div>
      <div class="hero-actions">
        <button class="btn btn-export btn-sm" onclick="exportXLSX()">&#8595; Exportar Excel</button>
        <button class="btn btn-primary" onclick="openModal()">&#43; Nova ação</button>
      </div>
    </div>
  </div>

  <!-- STATS -->
  <div class="stats-grid">
    <div class="stat-card"><div class="stat-label">Total de ações</div><div class="stat-value" id="stat-total">0</div></div>
    <div class="stat-card"><div class="stat-label">Ativas</div><div class="stat-value" id="stat-ativas" style="color:#3B6D11">0</div></div>
    <div class="stat-card"><div class="stat-label">Planejadas</div><div class="stat-value" id="stat-plan" style="color:#185FA5">0</div></div>
    <div class="stat-card"><div class="stat-label">Encerradas</div><div class="stat-value" id="stat-enc" style="color:#5F5E5A">0</div></div>
    <div class="stat-card"><div class="stat-label">Budget total</div><div class="stat-value budget-color" id="stat-budget">R$ 0</div></div>
    <div class="stat-card"><div class="stat-label">ROI médio</div><div class="stat-value" id="stat-roi" style="color:#3C3489">—</div></div>
  </div>

  <!-- FILTERS -->
  <div class="filters-bar">
    <div class="filters-title">&#9698; Filtros</div>
    <div class="filters-row">
      <div class="filter-group">
        <span class="filter-label">Busca livre</span>
        <input type="text" id="search" placeholder="Nome, empresa, descrição..." oninput="renderList()">
      </div>
      <div class="filter-group">
        <span class="filter-label">Tipo de ação</span>
        <select id="f-tipo" onchange="renderList()">
          <option value="">Todos os tipos</option>
          <option>Promoção</option><option>Campanha digital</option><option>Evento</option>
          <option>Trade marketing</option><option>Mídia paga</option><option>Email marketing</option>
          <option>Influencer</option><option>Encartes/flyer</option><option>Sazonal</option><option>Outro</option>
        </select>
      </div>
      <div class="filter-group">
        <span class="filter-label">Empresa / Marca</span>
        <select id="f-empresa" onchange="renderList()">
          <option value="">Todas</option>
          <option>Nova Era AM</option><option>Nova Era RR</option><option>Nova Era PVH</option><option>Pátio Gourmet</option>
        </select>
      </div>
      <div class="filter-group">
        <span class="filter-label">Mês de início</span>
        <select id="f-mes" onchange="renderList()">
          <option value="">Todos</option>
          <option value="01">Janeiro</option><option value="02">Fevereiro</option><option value="03">Março</option>
          <option value="04">Abril</option><option value="05">Maio</option><option value="06">Junho</option>
          <option value="07">Julho</option><option value="08">Agosto</option><option value="09">Setembro</option>
          <option value="10">Outubro</option><option value="11">Novembro</option><option value="12">Dezembro</option>
        </select>
      </div>
    </div>
  </div>

  <div class="actions-header">
    <span class="actions-count" id="actions-count">0 ações</span>
    <div style="display:flex;gap:8px">
      <select id="f-status" onchange="renderList()" style="width:auto;font-size:12px;padding:5px 8px">
        <option value="">Todos os status</option>
        <option>Planejado</option><option>Ativo</option><option>Encerrado</option>
      </select>
      <button class="btn btn-sm" onclick="clearFilters()">&#10005; Limpar</button>
    </div>
  </div>

  <div id="list-container"></div>
</div>

<!-- MODAL -->
<div class="modal-overlay" id="modal-overlay" onclick="overlayClick(event)">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title"><span class="modal-dot"></span><span id="modal-title-text">Nova ação de marketing</span></span>
      <button class="icon-btn" onclick="closeModal()" aria-label="Fechar">&#10005;</button>
    </div>

    <div class="form-section">Identificação</div>
    <div class="form-row"><label class="form-label">Nome da ação *</label><input type="text" id="f-nome" placeholder="Ex: Promoção Dia das Mães 2026"></div>
    <div class="form-grid2">
      <div class="form-row"><label class="form-label">Tipo de ação *</label>
        <select id="f-ftipo">
          <option value="">Selecione...</option>
          <option>Promoção</option><option>Campanha digital</option><option>Evento</option>
          <option>Trade marketing</option><option>Mídia paga</option><option>Email marketing</option>
          <option>Influencer</option><option>Encartes/flyer</option><option>Sazonal</option><option>Outro</option>
        </select>
      </div>
      <div class="form-row"><label class="form-label">Empresa / Marca *</label>
        <select id="f-femp">
          <option value="">Selecione...</option>
          <option>Nova Era AM</option><option>Nova Era RR</option><option>Nova Era PVH</option><option>Pátio Gourmet</option>
        </select>
      </div>
    </div>

    <div class="form-section">Período e canal</div>
    <div class="form-grid3">
      <div class="form-row"><label class="form-label">Início</label><input type="date" id="f-ini"></div>
      <div class="form-row"><label class="form-label">Fim</label><input type="date" id="f-fim"></div>
      <div class="form-row"><label class="form-label">Status</label>
        <select id="f-fstatus"><option>Planejado</option><option>Ativo</option><option>Encerrado</option></select>
      </div>
    </div>
    <div class="form-grid2">
      <div class="form-row"><label class="form-label">Canal principal</label>
        <select id="f-canal">
          <option value="">Selecione...</option>
          <option>Instagram</option><option>Facebook</option><option>Google Ads</option><option>WhatsApp</option>
          <option>Email</option><option>Loja física</option><option>TV / Rádio</option>
          <option>Encarte impresso</option><option>App / Site</option><option>Múltiplos canais</option>
        </select>
      </div>
      <div class="form-row"><label class="form-label">Responsável</label><input type="text" id="f-resp" placeholder="Nome do responsável"></div>
    </div>

    <div class="form-section">Objetivo de negócio</div>
    <div class="form-grid2">
      <div class="form-row"><label class="form-label">Objetivo principal</label>
        <select id="f-obj">
          <option value="">Selecione...</option>
          <option>Aumento de vendas</option><option>Captação de clientes</option><option>Retenção / fidelização</option>
          <option>Awareness de marca</option><option>Lançamento de produto</option><option>Queima de estoque</option>
          <option>Engajamento digital</option><option>Tráfego para loja</option><option>Outro</option>
        </select>
      </div>
      <div class="form-row"><label class="form-label">Meta de resultado</label><input type="text" id="f-meta" placeholder="Ex: +15% vendas, 500 cadastros..."></div>
    </div>

    <div class="form-row">
      <label class="form-label">Descrição / mecânica</label>
      <textarea id="f-desc" placeholder="Descreva a mecânica, público-alvo e detalhes da ação..."></textarea>
    </div>

    <div class="form-row">
      <label class="form-label">&#128206; Anexar arquivos</label>
      <div class="upload-zone" id="upload-zone"
           onclick="document.getElementById('file-input').click()"
           ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="dropFiles(event)">
        <input type="file" id="file-input" multiple onchange="addFiles(this.files)">
        <label class="upload-zone-label">
          <span class="upload-icon">&#8679;</span>
          <span>Clique ou arraste arquivos aqui</span>
          <span style="font-size:11px;color:#888780">PDF, imagem, Word, Excel — até 10 arquivos</span>
        </label>
      </div>
      <div class="file-list" id="file-list"></div>
    </div>

    <div class="form-section">Budget e resultados</div>
    <div class="form-grid3">
      <div class="form-row"><label class="form-label">Budget (R$)</label><input type="number" id="f-budget" placeholder="0" min="0" step="100"></div>
      <div class="form-row"><label class="form-label">Receita gerada (R$)</label><input type="number" id="f-receita" placeholder="0" min="0" step="100"></div>
      <div class="form-row"><label class="form-label">ROI (%)</label><input type="number" id="f-roi" placeholder="Auto" step="1"></div>
    </div>
    <div class="form-row">
      <label class="form-label">Resultado alcançado</label>
      <textarea id="f-resultado" placeholder="Resultados, aprendizados e próximos passos..." style="min-height:60px"></textarea>
    </div>

    <div class="modal-footer">
      <button class="btn" onclick="closeModal()">Cancelar</button>
      <button class="btn btn-primary" onclick="saveAction()">&#10003; Salvar ação</button>
    </div>
  </div>
</div>

<script>
const KEY='mkt-acoes-v2';
let actions=[],editId=null,pendingFiles=[];

function load(){
  try{const d=localStorage.getItem(KEY);if(d)actions=JSON.parse(d);}catch(e){actions=[];}
  renderAll();
}
function persist(){try{localStorage.setItem(KEY,JSON.stringify(actions));}catch(e){}}
const fmt=n=>n!=null&&n!==''?'R$ '+Number(n).toLocaleString('pt-BR',{minimumFractionDigits:0,maximumFractionDigits:0}):'—';
const fmtDate=d=>d?d.split('-').reverse().join('/'):'—';
const esc=s=>String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function renderAll(){updateStats();renderList();}

function updateStats(){
  const bud=actions.reduce((s,a)=>s+(Number(a.budget)||0),0);
  const rois=actions.filter(a=>a.roi!==''&&a.roi!=null).map(a=>Number(a.roi));
  document.getElementById('stat-total').textContent=actions.length;
  document.getElementById('stat-ativas').textContent=actions.filter(a=>a.status==='Ativo').length;
  document.getElementById('stat-plan').textContent=actions.filter(a=>a.status==='Planejado').length;
  document.getElementById('stat-enc').textContent=actions.filter(a=>a.status==='Encerrado').length;
  document.getElementById('stat-budget').textContent=bud>0?'R$ '+bud.toLocaleString('pt-BR',{minimumFractionDigits:0}):'R$ 0';
  document.getElementById('stat-roi').textContent=rois.length?Math.round(rois.reduce((a,b)=>a+b,0)/rois.length)+'%':'—';
}

function renderList(){
  const srch=(document.getElementById('search').value||'').toLowerCase();
  const tipo=document.getElementById('f-tipo').value;
  const emp=document.getElementById('f-empresa').value;
  const mes=document.getElementById('f-mes').value;
  const status=document.getElementById('f-status').value;
  let list=actions.filter(a=>{
    if(srch&&![(a.nome||''),(a.empresa||''),(a.desc||''),(a.objetivo||''),(a.resultado||'')].join(' ').toLowerCase().includes(srch))return false;
    if(tipo&&a.tipo!==tipo)return false;
    if(emp&&a.empresa!==emp)return false;
    if(mes&&a.inicio&&a.inicio.split('-')[1]!==mes)return false;
    if(status&&a.status!==status)return false;
    return true;
  }).sort((a,b)=>(b.inicio||'').localeCompare(a.inicio||''));
  document.getElementById('actions-count').textContent=list.length+' ação'+(list.length!==1?'s':'')+' encontrada'+(list.length!==1?'s':'');
  const cont=document.getElementById('list-container');
  if(!list.length){
    cont.innerHTML='<div class="empty-state"><p style="font-size:36px;margin-bottom:10px">&#128230;</p><p>Nenhuma ação encontrada.</p><p style="font-size:12px;margin-top:4px">Cadastre uma nova ação ou ajuste os filtros.</p></div>';
    return;
  }
  cont.innerHTML=list.map(a=>{
    const sc=a.status==='Ativo'?'tag-ativo':a.status==='Encerrado'?'tag-encerrado':'tag-planejado';
    const dr=[a.inicio,a.fim].filter(Boolean).map(fmtDate).join(' &#8594; ');
    const roiVal=a.roi!==''&&a.roi!=null?Number(a.roi):null;
    const roiD=roiVal!=null?(roiVal>=0?'<span style="color:#27500A">+'+Math.round(roiVal)+'% ROI</span>':'<span style="color:#A32D2D">'+Math.round(roiVal)+'% ROI</span>'):'';
    const filesD=(a.files&&a.files.length)?'<div class="attached-files">'+a.files.map(f=>'<span class="attached-chip">&#128206; '+esc(f)+'</span>').join('')+'</div>':'';
    return '<div class="action-card">'
      +'<div class="action-top"><div>'
      +'<div class="action-title">'+esc(a.nome||'Sem título')+'</div>'
      +'<div class="action-meta">'
      +(a.tipo?'<span class="tag tag-tipo">'+esc(a.tipo)+'</span>':'')
      +(a.empresa?'<span class="tag tag-empresa">'+esc(a.empresa)+'</span>':'')
      +(a.canal?'<span class="tag tag-canal">'+esc(a.canal)+'</span>':'')
      +(a.status?'<span class="tag '+sc+'">'+esc(a.status)+'</span>':'')
      +(a.objetivo?'<span class="tag tag-objetivo">'+esc(a.objetivo)+'</span>':'')
      +'</div>'
      +(a.desc?'<div class="action-desc">'+esc(a.desc)+'</div>':'')
      +(a.resultado?'<div class="resultado-box"><strong>Resultado:</strong> '+esc(a.resultado)+'</div>':'')
      +(a.meta?'<div style="font-size:12px;color:#5f5e5a;margin-top:4px">&#8594; Meta: '+esc(a.meta)+'</div>':'')
      +filesD
      +'</div>'
      +'<div class="action-controls">'
      +'<button class="icon-btn" onclick="openModal(\''+a.id+'\')" title="Editar">&#9998;</button>'
      +'<button class="icon-btn del" onclick="del(\''+a.id+'\')" title="Excluir">&#128465;</button>'
      +'</div></div>'
      +'<div class="action-footer">'
      +(dr?'<span>&#128197; '+dr+'</span>':'')
      +(a.budget?'<span>'+fmt(a.budget)+'</span>':'')
      +roiD
      +(a.responsavel?'<span>&#128100; '+esc(a.responsavel)+'</span>':'')
      +'</div></div>';
  }).join('');
}

function openModal(id){
  editId=id||null;pendingFiles=[];
  const a=id?actions.find(x=>x.id===id):null;
  document.getElementById('modal-title-text').textContent=id?'Editar ação':'Nova ação de marketing';
  document.getElementById('f-nome').value=a?.nome||'';
  document.getElementById('f-resp').value=a?.responsavel||'';
  document.getElementById('f-meta').value=a?.meta||'';
  document.getElementById('f-desc').value=a?.desc||'';
  document.getElementById('f-resultado').value=a?.resultado||'';
  document.getElementById('f-ftipo').value=a?.tipo||'';
  document.getElementById('f-femp').value=a?.empresa||'';
  document.getElementById('f-ini').value=a?.inicio||'';
  document.getElementById('f-fim').value=a?.fim||'';
  document.getElementById('f-fstatus').value=a?.status||'Planejado';
  document.getElementById('f-canal').value=a?.canal||'';
  document.getElementById('f-obj').value=a?.objetivo||'';
  document.getElementById('f-budget').value=a?.budget||'';
  document.getElementById('f-receita').value=a?.receita||'';
  document.getElementById('f-roi').value=a?.roi||'';
  if(a?.files)pendingFiles=[...a.files];
  document.getElementById('file-input').value='';
  renderFileList();
  document.getElementById('modal-overlay').classList.add('open');
}
function closeModal(){document.getElementById('modal-overlay').classList.remove('open');editId=null;pendingFiles=[];}
function overlayClick(e){if(e.target===document.getElementById('modal-overlay'))closeModal();}

function addFiles(fileList){
  for(const f of fileList){if(pendingFiles.length<10&&!pendingFiles.includes(f.name))pendingFiles.push(f.name);}
  renderFileList();document.getElementById('file-input').value='';
}
function removeFile(i){pendingFiles.splice(i,1);renderFileList();}
function renderFileList(){
  document.getElementById('file-list').innerHTML=pendingFiles.map((n,i)=>
    '<div class="file-item"><span class="file-item-name">&#128196; '+esc(n)+'</span>'
    +'<button class="file-remove" onclick="removeFile('+i+')" title="Remover">&#10005;</button></div>'
  ).join('');
}
function dragOver(e){e.preventDefault();document.getElementById('upload-zone').classList.add('drag');}
function dragLeave(){document.getElementById('upload-zone').classList.remove('drag');}
function dropFiles(e){e.preventDefault();dragLeave();addFiles(e.dataTransfer.files);}

function calcROI(){
  const b=Number(document.getElementById('f-budget').value)||0;
  const r=Number(document.getElementById('f-receita').value)||0;
  if(b>0&&r>0)document.getElementById('f-roi').value=Math.round(((r-b)/b)*100);
}
document.addEventListener('change',e=>{if(e.target.id==='f-budget'||e.target.id==='f-receita')calcROI();});

function saveAction(){
  const nome=document.getElementById('f-nome').value.trim();
  const tipo=document.getElementById('f-ftipo').value;
  const empresa=document.getElementById('f-femp').value;
  if(!nome||!tipo||!empresa){alert('Preencha ao menos nome, tipo e empresa.');return;}
  const obj={
    id:editId||Date.now().toString(),nome,tipo,empresa,
    inicio:document.getElementById('f-ini').value,
    fim:document.getElementById('f-fim').value,
    status:document.getElementById('f-fstatus').value,
    canal:document.getElementById('f-canal').value,
    responsavel:document.getElementById('f-resp').value.trim(),
    objetivo:document.getElementById('f-obj').value,
    meta:document.getElementById('f-meta').value.trim(),
    desc:document.getElementById('f-desc').value.trim(),
    budget:document.getElementById('f-budget').value,
    receita:document.getElementById('f-receita').value,
    roi:document.getElementById('f-roi').value,
    resultado:document.getElementById('f-resultado').value.trim(),
    files:[...pendingFiles]
  };
  if(editId){const i=actions.findIndex(x=>x.id===editId);if(i>-1)actions[i]=obj;}
  else{actions.unshift(obj);}
  persist();closeModal();renderAll();
}

function del(id){
  if(!confirm('Deseja excluir esta ação? Esta operação não pode ser desfeita.'))return;
  actions=actions.filter(a=>a.id!==id);persist();renderAll();
}
function clearFilters(){
  ['search','f-tipo','f-mes','f-status','f-empresa'].forEach(id=>document.getElementById(id).value='');
  renderList();
}
function exportXLSX(){
  if(!window.XLSX){alert('Aguarde o carregamento e tente novamente.');return;}
  const h=['Nome','Tipo','Empresa/Marca','Status','Canal','Início','Fim','Objetivo','Meta','Responsável','Budget (R$)','Receita (R$)','ROI (%)','Descrição','Resultado','Arquivos anexados'];
  const rows=actions.map(a=>[
    a.nome||'',a.tipo||'',a.empresa||'',a.status||'',a.canal||'',
    a.inicio?fmtDate(a.inicio):'',a.fim?fmtDate(a.fim):'',
    a.objetivo||'',a.meta||'',a.responsavel||'',
    a.budget?Number(a.budget):'',a.receita?Number(a.receita):'',
    a.roi!=null&&a.roi!==''?Number(a.roi):'',
    a.desc||'',a.resultado||'',(a.files||[]).join(', ')
  ]);
  const ws=XLSX.utils.aoa_to_sheet([h,...rows]);
  ws['!cols']=[30,18,22,14,20,12,12,26,24,20,14,14,10,40,40,30].map(w=>({wch:w}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Ações de Marketing');
  XLSX.writeFile(wb,'acoes_marketing_'+new Date().getFullYear()+'.xlsx');
}

load();
</script>
</body>
</html>