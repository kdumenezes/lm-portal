// ╔══════════════════════════════════════════════════════════╗
// ║  LM Informática — Portal Interno                         ║
// ║  app.js v3 — agenda login corrigido                      ║
// ╚══════════════════════════════════════════════════════════╝

const CONFIG = {
  CLIENT_ID: '006933b2-4896-4a69-b358-7c3abd3fcb87',
  TENANT_ID: '5c26622f-2878-40e4-ac31-0b8abaace688',
  REDIRECT:  'https://lemon-ocean-03478f91e.7.azurestaticapps.net',
  SCOPES:    ['Calendars.Read', 'Calendars.Read.Shared', 'Group.Read.All', 'User.Read'],
  API_URL:   'https://lminformatica.blob.core.windows.net/portal-data/os.json',
  ATUALIZAR: 300,  // 5 minutos — sincronizado com o agente
};

let msalInstance   = null;
let currentAccount = null;
let currentTab     = 'aberto';
let currentFilter  = 'todos';
let allOS          = [];

// ── DADOS DEMO OS ─────────────────────────────────────────
const DEMO_OS = [
  { id:'OS-4821', cliente:'João Silva',        problema:'Notebook Dell Inspiron — tela quebrada',   entrada:'2026-04-17', previsao:null, status:'aberto',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4820', cliente:'Pedro Santos',      problema:'Notebook Lenovo — troca de HD por SSD',    entrada:'2026-04-17', previsao:null, status:'manutencao', urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4819', cliente:'Maria Oliveira',    problema:'PC Desktop — não liga',                    entrada:'2026-04-16', previsao:null, status:'aberto',     urgente:true,  tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4817', cliente:'Loja Central',      problema:'Servidor — configuração de rede',          entrada:'2026-04-16', previsao:null, status:'manutencao', urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva'   },
  { id:'OS-4815', cliente:'Empresa ABC Ltda.', problema:'Impressora HP — papel emperrado',          entrada:'2026-04-15', previsao:null, status:'aberto',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4812', cliente:'Fernanda Lima',     problema:'iMac — limpeza e formatação',              entrada:'2026-04-15', previsao:null, status:'manutencao', urgente:false, tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4810', cliente:'Carlos Pereira',    problema:'Tablet Samsung — bateria não carrega',     entrada:'2026-04-14', previsao:null, status:'aberto',     urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva'   },
  { id:'OS-4808', cliente:'Ana Rodrigues',     problema:'Macbook Air — sistema travando',           entrada:'2026-04-13', previsao:null, status:'aberto',     urgente:true,  tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4805', cliente:'Roberto Alves',     problema:'Notebook Acer — teclado substituído',      entrada:'2026-04-12', previsao:null, status:'pronta',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4802', cliente:'Sílvia Moura',      problema:'PC Desktop — formatação Windows 11',       entrada:'2026-04-11', previsao:null, status:'pronta',     urgente:false, tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4799', cliente:'Tech Solutions',    problema:'Impressora — manutenção preventiva',       entrada:'2026-04-10', previsao:null, status:'pronta',     urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva'   },
];

const DEMO_TECNICOS = [
  { iniciais:'RC', nome:'Ricardo Costa', os_count:5, carga:83 },
  { iniciais:'FS', nome:'Felipe Santos', os_count:4, carga:67 },
  { iniciais:'LS', nome:'Lucas Silva',   os_count:3, carga:50 },
];

// ── INIT ──────────────────────────────────────────────────
// ── DOMÍNIO PERMITIDO ─────────────────────────────────────
const DOMINIO_PERMITIDO = 'lmrs.com.br';

// ── INIT — Login obrigatório antes de tudo ─────────────────
document.addEventListener('DOMContentLoaded', async () => {
  setDate();

  // Bloqueia o dashboard — só libera após login validado
  showDashboard(false);

  try {
    await initMSAL();
  } catch (err) {
    console.error('Erro MSAL:', err);
    showLoginError('Erro ao conectar com Microsoft. Tente recarregar a página.');
  }
});

// ── CONTROLE DE TELAS ─────────────────────────────────────
function showDashboard(logado) {
  const loginScreen = document.getElementById('login-screen');
  const dashboard   = document.getElementById('dashboard');
  if (loginScreen) loginScreen.style.display = logado ? 'none' : 'flex';
  if (dashboard)   dashboard.style.display   = logado ? 'block' : 'none';
}

function showLoginError(msg) {
  const loading = document.getElementById('login-loading');
  const error   = document.getElementById('login-error');
  const errorMsg = document.getElementById('login-error-msg');
  const btn     = document.getElementById('btn-login');
  if (loading) loading.classList.remove('visible');
  if (error)   error.classList.add('visible');
  if (errorMsg && msg) errorMsg.textContent = msg;
  if (btn) { btn.disabled = false; btn.style.opacity = '1'; }
}

function showLoginLoading() {
  const loading = document.getElementById('login-loading');
  const error   = document.getElementById('login-error');
  const btn     = document.getElementById('btn-login');
  if (loading) loading.classList.add('visible');
  if (error)   error.classList.remove('visible');
  if (btn) { btn.disabled = true; btn.style.opacity = '0.6'; }
}


function setDate() {
  const el = document.getElementById('hero-date');
  if (!el) return;
  const opts = { weekday:'long', year:'numeric', month:'long', day:'numeric' };
  const d = new Date().toLocaleDateString('pt-BR', opts);
  el.textContent = d.charAt(0).toUpperCase() + d.slice(1);
}

// ── MICROSOFT LOGIN ───────────────────────────────────────
async function initMSAL() {
  await loadScript('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.10.0/lib/msal-browser.min.js');

  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CONFIG.CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: CONFIG.REDIRECT,
    },
    cache: { cacheLocation: 'sessionStorage' }
  });

  await msalInstance.initialize();

  // Trata retorno após redirect de login
  const response = await msalInstance.handleRedirectPromise();
  if (response && response.account) {
    await validarELiberar(response.account);
    return;
  }

  // Verifica se ja tem sessao ativa
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await validarELiberar(accounts[0]);
  }
  // Se nao tem sessao, tela de login ja esta visivel
}

// Valida dominio e libera o dashboard
async function validarELiberar(account) {
  const email   = account.username || '';
  const dominio = (email.split('@')[1] || '').toLowerCase();

  if (dominio !== DOMINIO_PERMITIDO) {
    try { await msalInstance.logout({ account }); } catch(e) {}
    showLoginError(
      `Acesso negado. Apenas @${DOMINIO_PERMITIDO} podem acessar este portal. (Conta: ${email})`
    );
    return;
  }

  // Dominio valido — libera o dashboard
  currentAccount = account;
  updateUserUI(account);
  showDashboard(true);

  // Inicia o dashboard
  loadOS();
  renderTecnicos();
  await loadCalendar();
  setInterval(loadOS, CONFIG.ATUALIZAR * 1000);
}

// Chamado pelo botão "Entrar com Microsoft 365"
async function loginM365() {
  showLoginLoading();
  if (!msalInstance) {
    try {
      await initMSAL();
      if (!msalInstance) throw new Error('MSAL nao disponivel');
    } catch(e) {
      showLoginError('Erro ao conectar com Microsoft. Tente recarregar a pagina.');
      return;
    }
  }
  try {
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
  } catch(e) {
    console.error('Erro no login:', e);
    showLoginError('Erro ao iniciar login. Tente novamente.');
  }
}

async function getToken() {
  if (!currentAccount || !msalInstance) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.SCOPES, account: currentAccount
    });
    return r.accessToken;
  } catch {
    await msalInstance.acquireTokenRedirect({ scopes: CONFIG.SCOPES });
    return null;
  }
}

function updateUserUI(account) {
  if (!account) return;
  const nome     = account.name || account.username;
  const nameEl   = document.querySelector('.user-name');
  const avatarEl = document.querySelector('.avatar');
  const dotEl    = document.querySelector('.user-dot');
  if (nameEl)   nameEl.textContent = nome.split(' ')[0];
  if (avatarEl) {
    const parts = nome.trim().split(' ');
    avatarEl.textContent = parts.length >= 2
      ? (parts[0][0] + parts[parts.length-1][0]).toUpperCase()
      : nome.substring(0,2).toUpperCase();
  }
  if (dotEl) dotEl.style.background = '#22c55e';
}

// ── AGENDA MICROSOFT 365 ──────────────────────────────────
async function loadCalendar() {
  const token = await getToken();
  if (!token) { showLoginPrompt(); return; }

  try {
    const agora = new Date();
    const fim   = new Date();
    fim.setDate(fim.getDate() + 14);

    // Busca o calendário do grupo Microsoft 365 "Agenda LM"
    const GROUP_ID = '40d28b1d-bea5-4ca0-adb6-6912e62a3fc8';
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/calendar/calendarView` +
      `?startDateTime=${agora.toISOString()}` +
      `&endDateTime=${fim.toISOString()}` +
      `&$orderby=start/dateTime&$top=20` +
      `&$select=subject,start,end,location,categories,body,organizer`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Graph error ${res.status}`);
    const data = await res.json();
    renderCalendarM365(data.value || []);
  } catch (err) {
    console.error('Erro na agenda Graph:', err);
    renderAgendaDemo();
  }
}

function renderCalendarM365(events) {
  const el = document.getElementById('agenda-list');
  if (!el) return;

  if (events.length === 0) {
    el.innerHTML = `<div style="padding:16px;text-align:center;font-size:12px;color:var(--text-sec)">
      Nenhum evento nos próximos 14 dias
    </div>`;
    return;
  }

  // Agrupa por dia
  const byDay = {};
  events.forEach(ev => {
    const d   = new Date(ev.start.dateTime || ev.start.date);
    const key = d.toDateString();
    if (!byDay[key]) byDay[key] = { date: d, events: [] };
    byDay[key].events.push(ev);
  });

  const hoje   = new Date().toDateString();
  const semana = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  const cor    = cats => {
    if (!cats || !cats.length) return '#00509d';
    const c = (cats[0]||'').toLowerCase();
    if (c.includes('visit') || c.includes('técn')) return '#00509d';
    if (c.includes('entrega') || c.includes('retirad')) return '#059669';
    if (c.includes('urgent') || c.includes('prazo')) return '#dc2626';
    return '#ffd500';
  };

  el.innerHTML = Object.values(byDay).slice(0,4).map(({ date, events }) => `
    <div class="agenda-item">
      <div class="day-box ${date.toDateString()===hoje ? 'today' : ''}">
        <div class="day-num">${date.getDate()}</div>
        <div class="day-name">${semana[date.getDay()]}</div>
      </div>
      <div class="ev-list">
        ${events.map(ev => {
          const horaStr = formatHoraRange(ev);
          const local = ev.location?.displayName || '';
          return `<div class="ev-item" onclick='openEventModal(${JSON.stringify(ev).replace(/'/g,"&#39;")})' style="cursor:pointer">
            <div class="ev-dot" style="background:${cor(ev.categories)}"></div>
            <div class="ev-content">
              <div class="ev-title">${ev.subject}</div>
              <div class="ev-time">${horaStr}${local?' — '+local:''}</div>
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>`).join('');
}

// ── BOTÃO DE LOGIN ────────────────────────────────────────
function showLoginPrompt() {
  const el = document.getElementById('agenda-list');
  if (!el) return;
  el.innerHTML = `
    <div style="padding:24px 16px;text-align:center">
      <div style="width:40px;height:40px;background:var(--lm-light);border-radius:50%;
        display:flex;align-items:center;justify-content:center;margin:0 auto 12px">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#00509d" stroke-width="2">
          <rect x="3" y="4" width="18" height="18" rx="2"/>
          <path d="M16 2v4M8 2v4M3 10h18"/>
        </svg>
      </div>
      <div style="font-size:12px;color:var(--text-sec);margin-bottom:14px;line-height:1.5">
        Conecte sua conta Microsoft<br>para ver a agenda real
      </div>
      <button onclick="loginM365()" style="
        background:var(--lm3);color:white;border:none;border-radius:6px;
        padding:9px 20px;font-size:12px;font-weight:600;
        cursor:pointer;font-family:Barlow,sans-serif;width:100%;
        transition:background 0.15s">
        Entrar com Microsoft 365
      </button>
      <div style="margin-top:8px">
        <button onclick="renderAgendaDemo()" style="
          background:none;border:none;color:var(--text-sec);
          font-size:11px;cursor:pointer;text-decoration:underline;
          font-family:Barlow,sans-serif">
          Ver dados de exemplo
        </button>
      </div>
    </div>`;
}

// ── AGENDA DEMO ───────────────────────────────────────────
function renderAgendaDemo() {
  const el = document.getElementById('agenda-list');
  if (!el) return;
  el.innerHTML = `
    <div style="padding:6px 14px 2px">
      <span style="font-size:10px;color:#d97706;font-weight:600;
        background:#fef3cd;padding:2px 8px;border-radius:20px">
        Dados de demonstração
      </span>
    </div>
    <div class="agenda-item">
      <div class="day-box today">
        <div class="day-num">19</div><div class="day-name">Dom</div>
      </div>
      <div class="ev-list">
        <div class="ev-item">
          <div class="ev-dot" style="background:#00509d"></div>
          <div class="ev-content">
            <div class="ev-title">Visita técnica — Empresa ABC</div>
            <div class="ev-time">09:00 — Rua das Flores, 123</div>
          </div>
        </div>
        <div class="ev-item">
          <div class="ev-dot" style="background:#ffd500"></div>
          <div class="ev-content">
            <div class="ev-title">Entrega OS-4805 — Roberto Alves</div>
            <div class="ev-time">14:00 — Na loja</div>
          </div>
        </div>
        <div class="ev-item">
          <div class="ev-dot" style="background:#dc2626"></div>
          <div class="ev-content">
            <div class="ev-title">Reunião equipe técnica</div>
            <div class="ev-time">17:00 — Sala de reuniões</div>
          </div>
        </div>
      </div>
    </div>
    <div class="agenda-item" style="border:none">
      <div class="day-box">
        <div class="day-num">22</div><div class="day-name">Qua</div>
      </div>
      <div class="ev-list">
        <div class="ev-item">
          <div class="ev-dot" style="background:#00509d"></div>
          <div class="ev-content">
            <div class="ev-title">Instalação de rede — Loja Central</div>
            <div class="ev-time">08:30 — Av. Principal, 500</div>
          </div>
        </div>
      </div>
    </div>`;
}

// ── ORDENS DE SERVIÇO ─────────────────────────────────────
async function loadOS() {
  try {
    // Adiciona timestamp para evitar cache do browser
    const url = `${CONFIG.API_URL}?t=${Date.now()}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();

    // O JSON do agente tem { gerado_em, totais, os: [...] }
    allOS = Array.isArray(data) ? data : (data.os || []);

    // Mostra horário da última sincronização
    if (data.gerado_em) {
      const dt = new Date(data.gerado_em);
      const hora = dt.toLocaleTimeString('pt-BR', { hour:'2-digit', minute:'2-digit' });
      const badge = document.querySelector('.systec-badge');
      if (badge) badge.textContent = `● Systec · atualizado às ${hora}`;
    }

  } catch (err) {
    console.warn('Usando dados demo — erro ao carregar OS:', err);
    allOS = DEMO_OS;
    const badge = document.querySelector('.systec-badge');
    if (badge) { badge.textContent = '⚠ Modo demo'; badge.style.background = '#fef3cd'; badge.style.color = '#92600a'; }
  }
  renderOS();
  updateStats();
}

function renderOS() {
  const tbody  = document.getElementById('os-tbody');
  if (!tbody) return;
  const search = (document.getElementById('searchInput')?.value||'').toLowerCase();

  const filtered = allOS.filter(os => {
    const matchTab    = os.status === currentTab;
    const matchFilter = currentFilter==='todos' ? true
      : currentFilter==='urgente' ? os.urgente
      : os.status===currentFilter;
    const matchSearch = !search
      || os.id.toLowerCase().includes(search)
      || os.cliente.toLowerCase().includes(search)
      || (os.equipamento || '').toLowerCase().includes(search)
      || os.problema.toLowerCase().includes(search);
    return matchTab && matchFilter && matchSearch;
  });

  const counts = { aberto:0, manutencao:0, pronta:0 };
  allOS.forEach(o => { if(counts[o.status]!==undefined) counts[o.status]++; });
  setText('count-aberto',     counts.aberto);
  setText('count-manutencao', counts.manutencao);
  setText('count-pronta',     counts.pronta);

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--text-sec);font-size:13px">Nenhuma OS encontrada</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(os => {
    const badge = os.urgente
      ? `<span class="badge badge-urg">! Urgente</span>`
      : os.status==='aberto'     ? `<span class="badge badge-ab">● Aguardando</span>`
      : os.status==='manutencao' ? `<span class="badge badge-ma">◐ Manutenção</span>`
      : `<span class="badge badge-pr">✓ Pronta</span>`;
    const prevStyle = os.status==='pronta'   ? 'color:#059669;font-weight:600'
      : isOverdue(os.previsao) ? 'color:#dc2626;font-weight:600'
      : 'color:var(--text-sec)';
    const valor = os.valor
      ? `<div style="font-size:11px;color:#059669;font-weight:700;margin-top:2px">R$ ${os.valor.toFixed(2).replace('.',',')}</div>` : '';
    return `<tr onclick="openOS('${os.id}')">
      <td><div class="os-num">${os.id}</div></td>
      <td><div class="os-client">${os.cliente}</div>${os.telefone ? `<div class="os-phone">${os.telefone}</div>` : ''}</td>
      <td><span style="font-size:12px">${os.equipamento || os.problema || '—'}</span>${os.equipamento && os.problema ? `<div class="os-phone">${os.problema}</div>` : ''}${valor}</td>
      <td style="font-size:11px;color:var(--text-sec)">${formatDate(os.entrada)}</td>
      <td style="font-size:11px;${prevStyle}">${formatDate(os.previsao)}</td>
      <td>${badge}</td>
      <td><div class="tech-av" title="${os.tecnico_nome}">${os.tecnico}</div></td>
    </tr>`;
  }).join('');
}

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
      <span class="badge ${t.os_count>=5?'badge-urg':t.os_count>=4?'badge-ma':'badge-ab'}">${t.os_count} OS</span>
    </div>`).join('');
}

function updateStats() {
  setText('stat-aberto',     allOS.filter(o=>o.status==='aberto').length);
  setText('stat-manutencao', allOS.filter(o=>o.status==='manutencao').length);
  setText('stat-pronta',     allOS.filter(o=>o.status==='pronta').length);
  setText('stat-urgente',    allOS.filter(o=>o.urgente).length);
}

// ── CONTROLES ─────────────────────────────────────────────
function switchTab(el, tab) {
  document.querySelectorAll('.os-tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  currentTab = tab;
  renderOS();
}

function setChip(el, filter) {
  document.querySelectorAll('.chip').forEach(c=>c.classList.remove('active'));
  el.classList.add('active');
  currentFilter = filter;
  if (['aberto','manutencao','pronta'].includes(filter)) {
    const tab = document.querySelector(`.os-tab[onclick*="${filter}"]`);
    if (tab) switchTab(tab, filter);
  } else { renderOS(); }
}

function filterOS()  { renderOS(); }
function openOS(id)  { console.log('OS selecionada:', allOS.find(o=>o.id===id)); }

// ── LOGOUT ────────────────────────────────────────────────
async function logout() {
  if (!msalInstance || !currentAccount) {
    showDashboard(false);
    return;
  }
  try {
    // Limpa sessão e volta para tela de login
    await msalInstance.logout({
      account: currentAccount,
      onRedirectNavigate: () => {
        // Evita redirect externo — fica na mesma página
        showDashboard(false);
        currentAccount = null;
        // Reseta avatar
        const av = document.querySelector('.avatar');
        const nm = document.querySelector('.user-name');
        const dt = document.querySelector('.user-dot');
        if (av) av.textContent = 'LM';
        if (nm) nm.textContent = 'Administrador';
        if (dt) dt.style.background = '#6b8279';
        return false; // false = não redireciona para fora
      }
    });
  } catch(e) {
    console.error('Erro logout:', e);
    // Fallback — força volta para tela de login
    sessionStorage.clear();
    showDashboard(false);
    currentAccount = null;
  }
}



// ── NAVEGAÇÃO DE PÁGINAS ──────────────────────────────────
let calWeekOffset = 0;

function showPage(page) {
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
  if (event?.target) event.target.classList.add('active');

  const pageMain   = document.getElementById('page-main');
  const pageAgenda = document.getElementById('page-agenda');
  const searchBar  = document.querySelector('.searchbar');
  const heroEl     = document.querySelector('.hero');

  if (page === 'agenda') {
    if (pageMain)   pageMain.style.display   = 'none';
    if (pageAgenda) pageAgenda.style.display = 'block';
    if (searchBar)  searchBar.style.display  = 'none';
    if (heroEl)     heroEl.style.display     = 'none';
    calWeekOffset = 0;
    renderAgendaPage();
  } else {
    if (pageMain)   pageMain.style.display   = '';
    if (pageAgenda) pageAgenda.style.display = 'none';
    if (searchBar)  searchBar.style.display  = '';
    if (heroEl)     heroEl.style.display     = '';
  }
}

function navCal(dir)  { calWeekOffset += dir; renderAgendaPage(); }
function navCalToday(){ calWeekOffset = 0;    renderAgendaPage(); }

async function renderAgendaPage() {
  const container = document.getElementById('agenda-full-content');
  const subEl     = document.getElementById('agenda-page-sub');
  if (!container) return;

  // Semana atual + offset
  const hoje   = new Date();
  const dow    = hoje.getDay();
  const inicio = new Date(hoje);
  inicio.setDate(hoje.getDate() - dow + (calWeekOffset * 7));
  inicio.setHours(0,0,0,0);
  const fim = new Date(inicio);
  fim.setDate(inicio.getDate() + 6);
  fim.setHours(23,59,59,999);

  const opts = { day:'numeric', month:'long' };
  if (subEl) subEl.textContent =
    `${inicio.toLocaleDateString('pt-BR',opts)} – ${fim.toLocaleDateString('pt-BR',opts)} de ${fim.getFullYear()}`;

  container.innerHTML = `<div style="text-align:center;padding:40px;color:var(--text-sec);font-size:13px">Carregando...</div>`;

  let events = [];
  try {
    const token = await getToken();
    if (!token) throw new Error('sem token');
    const GROUP_ID = '40d28b1d-bea5-4ca0-adb6-6912e62a3fc8';
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/calendar/calendarView` +
      `?startDateTime=${inicio.toISOString()}&endDateTime=${fim.toISOString()}` +
      `&$orderby=start/dateTime&$top=50&$select=subject,start,end,location,categories,body,organizer`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`${res.status}`);
    const data = await res.json();
    events = data.value || [];
  } catch(e) {
    console.warn('Agenda page error:', e);
  }

  const semCurta = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  const hojeStr  = new Date().toDateString();

  // Agrupa por dia (dom → sáb)
  const byDay = {};
  for (let i = 0; i < 7; i++) {
    const d = new Date(inicio);
    d.setDate(inicio.getDate() + i);
    byDay[d.toDateString()] = { date: d, events: [] };
  }
  events.forEach(ev => {
    const d   = new Date(ev.start.dateTime || ev.start.date);
    const key = d.toDateString();
    if (byDay[key]) byDay[key].events.push(ev);
  });

  const cols = Object.values(byDay).map(({ date, events }) => {
    const isHoje = date.toDateString() === hojeStr;
    const evHtml = events.length
      ? events.map(ev => {
          const hora  = formatHoraRange(ev);
          const local = ev.location?.displayName || "";
          const evData = JSON.stringify(ev).replace(/'/g,"\u0027");
          return `<div class="agenda-ev" onclick="openEventModal(JSON.parse(decodeURIComponent('${encodeURIComponent(JSON.stringify(ev))}')))" style="cursor:pointer">
            <div class="agenda-ev-time">${hora}</div>
            <div class="agenda-ev-title">${ev.subject}</div>
            ${local ? `<div class="agenda-ev-local">${local}</div>` : ""}
          </div>`;
        }).join('')
      : `<div class="agenda-empty">Sem eventos</div>`;

    return `<div class="agenda-day-col ${isHoje?'today':''}">
      <div class="agenda-day-header">
        <div class="agenda-day-num">${date.getDate()}</div>
        <div class="agenda-day-name">${semCurta[date.getDay()]}</div>
      </div>
      ${evHtml}
    </div>`;
  }).join('');

  container.innerHTML = `<div class="agenda-week-grid">${cols}</div>`;
}


// ── UTILS ─────────────────────────────────────────────────
function formatDate(s) { if(!s) return '—'; const [y,m,d]=s.split('-'); return `${d}/${m}/${y}`; }
function isOverdue(s)  { return s ? new Date(s)<new Date() : false; }
function setText(id,v) { const el=document.getElementById(id); if(el) el.textContent=v; }

// Converte dateTime do Graph para horário de Brasília
// O Outlook retorna sem 'Z' quando tem fuso próprio — adicionamos Z para forçar UTC
// e depois convertemos para America/Sao_Paulo
function parseEventTime(dateTimeStr, timeZoneStr) {
  if (!dateTimeStr) return null;
  // Se já tem Z ou +, usa direto; senão adiciona Z (é UTC)
  const iso = /Z|[+-]\d{2}:\d{2}$/.test(dateTimeStr)
    ? dateTimeStr
    : dateTimeStr + 'Z';
  return new Date(iso);
}

function formatHora(dateTimeStr, timeZoneStr) {
  const d = parseEventTime(dateTimeStr, timeZoneStr);
  if (!d) return '';
  return d.toLocaleTimeString('pt-BR', {
    hour: '2-digit', minute: '2-digit',
    timeZone: 'America/Sao_Paulo'
  });
}

function formatHoraRange(ev) {
  const hi = formatHora(ev.start?.dateTime, ev.start?.timeZone);
  const hf = formatHora(ev.end?.dateTime,   ev.end?.timeZone);
  if (!hi) return 'Dia inteiro';
  return hf ? `${hi} – ${hf}` : hi;
}

// ── MODAL DE DETALHES DO EVENTO ───────────────────────────
function openEventModal(ev) {
  // Remove modal anterior se existir
  const old = document.getElementById('event-modal');
  if (old) old.remove();

  const hora  = formatHoraRange(ev);
  const local = ev.location?.displayName || '';
  const org   = ev.organizer?.emailAddress?.name || '';

  // Extrai texto do body sem HTML
  let desc = '';
  if (ev.body?.content) {
    const tmp = document.createElement('div');
    tmp.innerHTML = ev.body.content;
    desc = tmp.textContent?.trim() || '';
    if (desc.length > 300) desc = desc.substring(0, 300) + '...';
  }

  const modal = document.createElement('div');
  modal.id = 'event-modal';
  modal.style.cssText = `
    position:fixed;top:0;left:0;width:100%;height:100%;
    background:rgba(0,29,107,0.7);z-index:1000;
    display:flex;align-items:center;justify-content:center;
    padding:20px;
  `;
  modal.onclick = (e) => { if(e.target===modal) modal.remove(); };
  modal.innerHTML = `
    <div style="
      background:white;border-radius:12px;
      width:100%;max-width:480px;overflow:hidden;
      box-shadow:0 20px 60px rgba(0,0,0,0.3)
    ">
      <div style="background:#00296b;padding:18px 20px;display:flex;justify-content:space-between;align-items:flex-start">
        <div>
          <div style="font-size:11px;color:rgba(255,213,0,0.8);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px">Compromisso</div>
          <div style="font-size:17px;font-weight:700;color:white;line-height:1.3">${ev.subject}</div>
        </div>
        <button onclick="document.getElementById('event-modal').remove()" style="
          background:rgba(255,255,255,0.15);border:none;color:white;
          width:28px;height:28px;border-radius:50%;font-size:16px;
          cursor:pointer;display:flex;align-items:center;justify-content:center;
          flex-shrink:0;margin-left:12px
        ">×</button>
      </div>
      <div style="padding:18px 20px">
        <div style="display:grid;gap:12px">
          <div style="display:flex;align-items:flex-start;gap:10px">
            <div style="width:32px;height:32px;background:#e6eef7;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00509d" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            </div>
            <div>
              <div style="font-size:11px;color:#6b8279;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">Horário</div>
              <div style="font-size:14px;font-weight:600;color:#00296b;margin-top:1px">${hora}</div>
            </div>
          </div>
          ${local ? `
          <div style="display:flex;align-items:flex-start;gap:10px">
            <div style="width:32px;height:32px;background:#e6eef7;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00509d" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            </div>
            <div>
              <div style="font-size:11px;color:#6b8279;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">Local</div>
              <div style="font-size:14px;color:#00296b;margin-top:1px">${local}</div>
            </div>
          </div>` : ''}
          ${org ? `
          <div style="display:flex;align-items:flex-start;gap:10px">
            <div style="width:32px;height:32px;background:#e6eef7;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00509d" stroke-width="2"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>
            </div>
            <div>
              <div style="font-size:11px;color:#6b8279;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">Organizador</div>
              <div style="font-size:14px;color:#00296b;margin-top:1px">${org}</div>
            </div>
          </div>` : ''}
          ${desc ? `
          <div style="display:flex;align-items:flex-start;gap:10px">
            <div style="width:32px;height:32px;background:#e6eef7;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00509d" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
            </div>
            <div>
              <div style="font-size:11px;color:#6b8279;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">Descrição</div>
              <div style="font-size:13px;color:#4a6080;margin-top:1px;line-height:1.5">${desc}</div>
            </div>
          </div>` : ''}
        </div>
        <button onclick="document.getElementById('event-modal').remove()" style="
          width:100%;margin-top:18px;background:#00296b;color:white;
          border:none;border-radius:8px;padding:10px;font-size:13px;
          font-weight:600;cursor:pointer;font-family:Barlow,sans-serif
        ">Fechar</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

function loadScript(src) {
  return new Promise((res,rej) => {
    if (document.querySelector(`script[src="${src}"]`)) { res(); return; }
    const s = document.createElement('script');
    s.src=src; s.onload=res; s.onerror=rej;
    document.head.appendChild(s);
  });
}
