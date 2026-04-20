// ╔══════════════════════════════════════════════════════╗
// ║  LM Informática — Portal Interno                     ║
// ║  app.js — Lógica principal                           ║
// ║                                                      ║
// ║  CONEXÕES:                                           ║
// ║  → Systec/Firebird: via API intermediária (api.js)   ║
// ║  → Microsoft 365: via Microsoft Graph API            ║
// ╚══════════════════════════════════════════════════════╝

// ── CONFIGURAÇÃO ──────────────────────────────────────
// ATENÇÃO: Quando a API estiver pronta, substitua a URL abaixo
// pela URL real da sua API hospedada na Azure.
const CONFIG = {
  API_URL: 'http://localhost:3000',   // ← trocar pela URL da Azure quando hospedar
  ATUALIZAR_A_CADA: 60,              // segundos para atualizar os dados automaticamente
  MS_CLIENT_ID: 'SEU_CLIENT_ID_AQUI' // ← preencher após registrar app no Azure AD
};

// ── ESTADO DA APLICAÇÃO ───────────────────────────────
let currentTab = 'aberto';
let currentFilter = 'todos';
let allOS = [];
let intervalId = null;

// ── DADOS DE DEMONSTRAÇÃO ────────────────────────────
// Estes dados simulam o que virá do Systec/Firebird.
// Quando a API estiver pronta, a função loadOS() vai
// buscar os dados reais e substituir estes.
const DEMO_OS = [
  { id: 'OS-4821', cliente: 'João Silva',       telefone: '(51) 99999-1234', equipamento: 'Notebook Dell Inspiron', problema: 'Tela quebrada',          entrada: '2026-04-17', previsao: '2026-04-22', status: 'aberto',    urgente: false, tecnico: 'RC', tecnico_nome: 'Ricardo Costa' },
  { id: 'OS-4820', cliente: 'Pedro Santos',     telefone: '(51) 95555-7890', equipamento: 'Notebook Lenovo',       problema: 'Troca de HD por SSD',     entrada: '2026-04-17', previsao: '2026-04-20', status: 'manutencao', urgente: false, tecnico: 'RC', tecnico_nome: 'Ricardo Costa' },
  { id: 'OS-4819', cliente: 'Maria Oliveira',   telefone: '(51) 98888-5678', equipamento: 'PC Desktop',            problema: 'Não liga',                entrada: '2026-04-16', previsao: '2026-04-19', status: 'aberto',    urgente: true,  tecnico: 'FS', tecnico_nome: 'Felipe Santos' },
  { id: 'OS-4817', cliente: 'Loja Central',     telefone: '(51) 3222-1111',  equipamento: 'Servidor',              problema: 'Configuração de rede',    entrada: '2026-04-16', previsao: '2026-04-22', status: 'manutencao', urgente: false, tecnico: 'LS', tecnico_nome: 'Lucas Silva' },
  { id: 'OS-4815', cliente: 'Empresa ABC Ltda.',telefone: '(51) 3333-4444',  equipamento: 'Impressora HP',         problema: 'Papel emperrado',         entrada: '2026-04-15', previsao: '2026-04-21', status: 'aberto',    urgente: false, tecnico: 'RC', tecnico_nome: 'Ricardo Costa' },
  { id: 'OS-4812', cliente: 'Fernanda Lima',    telefone: '(51) 94444-2345', equipamento: 'iMac',                  problema: 'Limpeza e formatação',    entrada: '2026-04-15', previsao: '2026-04-21', status: 'manutencao', urgente: false, tecnico: 'FS', tecnico_nome: 'Felipe Santos' },
  { id: 'OS-4810', cliente: 'Carlos Pereira',   telefone: '(51) 97777-9012', equipamento: 'Tablet Samsung',        problema: 'Bateria não carrega',     entrada: '2026-04-14', previsao: '2026-04-23', status: 'aberto',    urgente: false, tecnico: 'LS', tecnico_nome: 'Lucas Silva' },
  { id: 'OS-4808', cliente: 'Ana Rodrigues',    telefone: '(51) 96666-3456', equipamento: 'Macbook Air',           problema: 'Sistema travando',        entrada: '2026-04-13', previsao: '2026-04-19', status: 'aberto',    urgente: true,  tecnico: 'FS', tecnico_nome: 'Felipe Santos' },
  { id: 'OS-4805', cliente: 'Roberto Alves',    telefone: '(51) 93333-6789', equipamento: 'Notebook Acer',         problema: 'Teclado substituído',     entrada: '2026-04-12', previsao: '2026-04-15', status: 'pronta',    urgente: false, tecnico: 'RC', tecnico_nome: 'Ricardo Costa', valor: 280.00 },
  { id: 'OS-4802', cliente: 'Sílvia Moura',     telefone: '(51) 92222-0123', equipamento: 'PC Desktop',            problema: 'Formatação Windows 11',   entrada: '2026-04-11', previsao: '2026-04-14', status: 'pronta',    urgente: false, tecnico: 'FS', tecnico_nome: 'Felipe Santos', valor: 150.00 },
  { id: 'OS-4799', cliente: 'Tech Solutions',   telefone: '(51) 3444-5555',  equipamento: 'Impressora',            problema: 'Manutenção preventiva',   entrada: '2026-04-10', previsao: '2026-04-13', status: 'pronta',    urgente: false, tecnico: 'LS', tecnico_nome: 'Lucas Silva',  valor: 90.00  },
];

const DEMO_AGENDA = [
  { dia: 19, mes: 'Abr', diasem: 'Dom', hoje: true, eventos: [
    { titulo: 'Visita técnica — Empresa ABC', hora: '09:00', local: 'Rua das Flores, 123', tipo: 'visita', cor: '#00509d' },
    { titulo: 'Entrega OS-4805 — Roberto Alves', hora: '14:00', local: 'Na loja', tipo: 'entrega', cor: '#ffd500' },
    { titulo: 'Reunião equipe técnica', hora: '17:00', local: 'Sala de reuniões', tipo: 'interno', cor: '#dc2626' },
  ]},
  { dia: 22, mes: 'Abr', diasem: 'Qua', hoje: false, eventos: [
    { titulo: 'Instalação de rede — Loja Central', hora: '08:30', local: 'Av. Principal, 500', tipo: 'visita', cor: '#00509d' },
  ]},
  { dia: 24, mes: 'Abr', diasem: 'Sex', hoje: false, eventos: [
    { titulo: 'Retirada OS-4817 — Loja Central', hora: '10:00', local: 'Na loja', tipo: 'entrega', cor: '#059669' },
  ]},
];

const DEMO_TECNICOS = [
  { iniciais: 'RC', nome: 'Ricardo Costa', os_count: 5, carga: 83 },
  { iniciais: 'FS', nome: 'Felipe Santos', os_count: 4, carga: 67 },
  { iniciais: 'LS', nome: 'Lucas Silva',   os_count: 3, carga: 50 },
];

// ── INIT ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setDate();
  loadOS();
  renderAgenda();
  renderTecnicos();
  // Atualiza automaticamente a cada N segundos
  intervalId = setInterval(loadOS, CONFIG.ATUALIZAR_A_CADA * 1000);
});

// ── DATA E HORA ───────────────────────────────────────
function setDate() {
  const el = document.getElementById('hero-date');
  if (!el) return;
  const opts = { weekday:'long', year:'numeric', month:'long', day:'numeric' };
  const d = new Date().toLocaleDateString('pt-BR', opts);
  el.textContent = d.charAt(0).toUpperCase() + d.slice(1) + ' — Systec conectado';
}

// ── CARREGAR OS ───────────────────────────────────────
// QUANDO A API ESTIVER PRONTA:
// Descomente o bloco "fetch real" e comente o bloco "DEMO".
async function loadOS() {
  try {
    // ── DEMO (remover quando API estiver pronta) ──────
    allOS = DEMO_OS;
    renderOS();
    updateStats();
    // ─────────────────────────────────────────────────

    // ── FETCH REAL (descomentar quando API pronta) ────
    // const res = await fetch(`${CONFIG.API_URL}/api/os`);
    // if (!res.ok) throw new Error('Erro na API');
    // allOS = await res.json();
    // renderOS();
    // updateStats();
    // ─────────────────────────────────────────────────

  } catch (err) {
    console.error('Erro ao carregar OS:', err);
    // Em caso de erro de rede, mantém os dados anteriores
  }
}

// ── RENDERIZAR TABELA ─────────────────────────────────
function renderOS() {
  const tbody = document.getElementById('os-tbody');
  if (!tbody) return;

  const search = (document.getElementById('searchInput')?.value || '').toLowerCase();

  const filtered = allOS.filter(os => {
    const matchTab = currentTab === 'todos' || os.status === currentTab;
    const matchFilter =
      currentFilter === 'todos' ? true :
      currentFilter === 'urgente' ? os.urgente :
      os.status === currentFilter;
    const matchSearch = !search ||
      os.id.toLowerCase().includes(search) ||
      os.cliente.toLowerCase().includes(search) ||
      os.equipamento.toLowerCase().includes(search) ||
      os.problema.toLowerCase().includes(search);
    return matchTab && matchFilter && matchSearch;
  });

  // Counts por tab
  const counts = { aberto: 0, manutencao: 0, pronta: 0 };
  allOS.forEach(os => { if (counts[os.status] !== undefined) counts[os.status]++; });
  const ca = document.getElementById('count-aberto');
  const cm = document.getElementById('count-manutencao');
  const cp = document.getElementById('count-pronta');
  if (ca) ca.textContent = counts.aberto;
  if (cm) cm.textContent = counts.manutencao;
  if (cp) cp.textContent = counts.pronta;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--text-sec);font-size:13px">Nenhuma OS encontrada</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(os => {
    const badge = os.urgente
      ? `<span class="badge badge-urg">! Urgente</span>`
      : os.status === 'aberto'     ? `<span class="badge badge-ab">● Aguardando</span>`
      : os.status === 'manutencao' ? `<span class="badge badge-ma">◐ Manutenção</span>`
      : `<span class="badge badge-pr">✓ Pronta</span>`;

    const previsao = os.status === 'pronta'
      ? `<span style="color:#059669;font-weight:600">${formatDate(os.previsao)}</span>`
      : isOverdue(os.previsao)
      ? `<span style="color:#dc2626;font-weight:600">${formatDate(os.previsao)}</span>`
      : formatDate(os.previsao);

    const valor = os.valor
      ? `<div style="font-size:11px;color:#059669;font-weight:700;margin-top:2px">R$ ${os.valor.toFixed(2).replace('.',',')}</div>`
      : '';

    return `<tr onclick="openOS('${os.id}')">
      <td><div class="os-num">${os.id}</div></td>
      <td><div class="os-client">${os.cliente}</div><div class="os-phone">${os.telefone}</div></td>
      <td><span style="font-size:12px">${os.equipamento}</span><div class="os-phone">${os.problema}</div>${valor}</td>
      <td style="font-size:11px;color:var(--text-sec)">${formatDate(os.entrada)}</td>
      <td style="font-size:11px">${previsao}</td>
      <td>${badge}</td>
      <td><div class="tech-av" title="${os.tecnico_nome}">${os.tecnico}</div></td>
    </tr>`;
  }).join('');
}

// ── RENDERIZAR AGENDA ─────────────────────────────────
function renderAgenda() {
  const el = document.getElementById('agenda-list');
  if (!el) return;

  const tipoTag = {
    visita:  `<span class="badge badge-vis" style="font-size:9px">Visita</span>`,
    entrega: `<span class="badge badge-ent" style="font-size:9px">Entrega</span>`,
    interno: `<span class="badge badge-int" style="font-size:9px">Interno</span>`,
  };

  el.innerHTML = DEMO_AGENDA.map(dia => `
    <div class="agenda-item">
      <div class="day-box ${dia.hoje ? 'today' : ''}">
        <div class="day-num">${dia.dia}</div>
        <div class="day-name">${dia.diasem}</div>
      </div>
      <div class="ev-list">
        ${dia.eventos.map(ev => `
          <div class="ev-item">
            <div class="ev-dot" style="background:${ev.cor}"></div>
            <div class="ev-content">
              <div class="ev-title">${ev.titulo} ${tipoTag[ev.tipo] || ''}</div>
              <div class="ev-time">${ev.hora}${ev.local ? ' — ' + ev.local : ''}</div>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `).join('');
}

// ── RENDERIZAR TÉCNICOS ───────────────────────────────
function renderTecnicos() {
  const el = document.getElementById('tecnicos-list');
  if (!el) return;
  el.innerHTML = DEMO_TECNICOS.map(t => `
    <div class="tec-item">
      <div class="tec-av">${t.iniciais}</div>
      <div class="tec-info">
        <div class="tec-name">${t.nome}</div>
        <div class="tec-detail">${t.os_count} OS em andamento</div>
        <div class="tec-bar"><div class="tec-bar-fill" style="width:${t.carga}%"></div></div>
      </div>
      <span class="badge ${t.os_count >= 5 ? 'badge-urg' : t.os_count >= 4 ? 'badge-ma' : 'badge-ab'}">${t.os_count} OS</span>
    </div>
  `).join('');
}

// ── ATUALIZAR STATS HERO ──────────────────────────────
function updateStats() {
  const ab  = allOS.filter(o => o.status === 'aberto').length;
  const ma  = allOS.filter(o => o.status === 'manutencao').length;
  const pr  = allOS.filter(o => o.status === 'pronta').length;
  const urg = allOS.filter(o => o.urgente).length;
  const el = id => document.getElementById(id);
  if (el('stat-aberto'))    el('stat-aberto').textContent    = ab;
  if (el('stat-manutencao'))el('stat-manutencao').textContent = ma;
  if (el('stat-pronta'))    el('stat-pronta').textContent    = pr;
  if (el('stat-urgente'))   el('stat-urgente').textContent   = urg;
}

// ── CONTROLES ─────────────────────────────────────────
function switchTab(el, tab) {
  document.querySelectorAll('.os-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  currentTab = tab;
  renderOS();
}

function setChip(el, filter) {
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  currentFilter = filter;
  // Sincroniza aba com filtro
  if (['aberto','manutencao','pronta'].includes(filter)) {
    const tab = document.querySelector(`.os-tab[onclick*="${filter}"]`);
    if (tab) switchTab(tab, filter);
  } else {
    renderOS();
  }
}

function filterOS() { renderOS(); }

function showPage(page) {
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
  event.target.classList.add('active');
  // Aqui você pode adicionar navegação entre páginas futuramente
}

function openOS(id) {
  // Aqui você pode abrir um modal com detalhes da OS
  // Por enquanto, apenas mostra no console
  const os = allOS.find(o => o.id === id);
  if (os) console.log('OS selecionada:', os);
}

// ── UTILITÁRIOS ───────────────────────────────────────
function formatDate(dateStr) {
  if (!dateStr) return '—';
  const [y, m, d] = dateStr.split('-');
  return `${d}/${m}/${y}`;
}

function isOverdue(dateStr) {
  if (!dateStr) return false;
  return new Date(dateStr) < new Date();
}
