// ╔══════════════════════════════════════════════════════════╗
// ║  LM Informática — Portal Interno                         ║
// ║  app.js v5 — agenda login corrigido                      ║
// ╚══════════════════════════════════════════════════════════╝

const CONFIG = {
  CLIENT_ID: '006933b2-4896-4a69-b358-7c3abd3fcb87',
  TENANT_ID: '5c26622f-2878-40e4-ac31-0b8abaace688',
  REDIRECT:  'https://lemon-ocean-03478f91e.7.azurestaticapps.net',
  SCOPES:    ['Calendars.Read', 'Calendars.Read.Shared', 'Group.Read.All', 'User.Read'],
  API_URL:   '',
  ATUALIZAR: 60,
};

let msalInstance   = null;
let currentAccount = null;
let currentTab     = 'aberto';
let currentFilter  = 'todos';
let allOS          = [];

// ── DADOS DEMO OS ─────────────────────────────────────────
const DEMO_OS = [
  { id:'OS-4821', cliente:'João Silva',        telefone:'(51) 99999-1234', equipamento:'Notebook Dell Inspiron', problema:'Tela quebrada',         entrada:'2026-04-17', previsao:'2026-04-22', status:'aberto',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4820', cliente:'Pedro Santos',      telefone:'(51) 95555-7890', equipamento:'Notebook Lenovo',        problema:'Troca de HD por SSD',   entrada:'2026-04-17', previsao:'2026-04-20', status:'manutencao', urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4819', cliente:'Maria Oliveira',    telefone:'(51) 98888-5678', equipamento:'PC Desktop',             problema:'Não liga',              entrada:'2026-04-16', previsao:'2026-04-19', status:'aberto',     urgente:true,  tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4817', cliente:'Loja Central',      telefone:'(51) 3222-1111',  equipamento:'Servidor',               problema:'Configuração de rede',  entrada:'2026-04-16', previsao:'2026-04-22', status:'manutencao', urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva'   },
  { id:'OS-4815', cliente:'Empresa ABC Ltda.', telefone:'(51) 3333-4444',  equipamento:'Impressora HP',          problema:'Papel emperrado',       entrada:'2026-04-15', previsao:'2026-04-21', status:'aberto',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa' },
  { id:'OS-4812', cliente:'Fernanda Lima',     telefone:'(51) 94444-2345', equipamento:'iMac',                   problema:'Limpeza e formatação',  entrada:'2026-04-15', previsao:'2026-04-21', status:'manutencao', urgente:false, tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4810', cliente:'Carlos Pereira',    telefone:'(51) 97777-9012', equipamento:'Tablet Samsung',         problema:'Bateria não carrega',   entrada:'2026-04-14', previsao:'2026-04-23', status:'aberto',     urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva'   },
  { id:'OS-4808', cliente:'Ana Rodrigues',     telefone:'(51) 96666-3456', equipamento:'Macbook Air',            problema:'Sistema travando',      entrada:'2026-04-13', previsao:'2026-04-19', status:'aberto',     urgente:true,  tecnico:'FS', tecnico_nome:'Felipe Santos' },
  { id:'OS-4805', cliente:'Roberto Alves',     telefone:'(51) 93333-6789', equipamento:'Notebook Acer',          problema:'Teclado substituído',   entrada:'2026-04-12', previsao:'2026-04-15', status:'pronta',     urgente:false, tecnico:'RC', tecnico_nome:'Ricardo Costa', valor:280.00 },
  { id:'OS-4802', cliente:'Sílvia Moura',      telefone:'(51) 92222-0123', equipamento:'PC Desktop',             problema:'Formatação Windows 11', entrada:'2026-04-11', previsao:'2026-04-14', status:'pronta',     urgente:false, tecnico:'FS', tecnico_nome:'Felipe Santos', valor:150.00 },
  { id:'OS-4799', cliente:'Tech Solutions',    telefone:'(51) 3444-5555',  equipamento:'Impressora',             problema:'Manutenção preventiva', entrada:'2026-04-10', previsao:'2026-04-13', status:'pronta',     urgente:false, tecnico:'LS', tecnico_nome:'Lucas Silva',   valor:90.00  },
];

const DEMO_TECNICOS = [
  { iniciais:'RC', nome:'Felipe Costa', os_count:5, carga:83 },
  { iniciais:'FS', nome:'Felipe Santos', os_count:4, carga:67 },
  { iniciais:'LS', nome:'Lucas Silva',   os_count:3, carga:50 },
];

// ── INIT ──────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  setDate();
  loadOS();
  renderTecnicos();

  // Mostra botão de login imediatamente — não espera MSAL carregar
  showLoginPrompt();

  // Carrega MSAL em paralelo
  initMSAL().catch(err => {
    console.warn('MSAL não carregou, agenda em modo demo:', err);
    renderAgendaDemo();
  });

  setInterval(loadOS, CONFIG.ATUALIZAR * 1000);
});

// ── DATA ──────────────────────────────────────────────────
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
    currentAccount = response.account;
  }

  // Verifica se já há conta logada na sessão
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    updateUserUI(currentAccount);
    await loadCalendar();
  } else {
    // Nenhuma conta — mostra botão de login
    showLoginPrompt();
  }
}

// Chamado pelo botão "Entrar com Microsoft 365"
async function loginM365() {
  if (!msalInstance) {
    // MSAL ainda não carregou — tenta inicializar agora
    try {
      await initMSAL();
      if (!msalInstance) throw new Error('MSAL não disponível');
    } catch(e) {
      alert('Erro ao conectar com Microsoft. Tente recarregar a página.');
      return;
    }
  }
  try {
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
  } catch(e) {
    console.error('Erro no login:', e);
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
      `&$select=subject,start,end,location,categories`,
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
          const hora = ev.start.dateTime
            ? new Date(ev.start.dateTime).toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})
            : 'Dia inteiro';
          const local = ev.location?.displayName || '';
          return `<div class="ev-item">
            <div class="ev-dot" style="background:${cor(ev.categories)}"></div>
            <div class="ev-content">
              <div class="ev-title">${ev.subject}</div>
              <div class="ev-time">${hora}${local?' — '+local:''}</div>
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
    allOS = CONFIG.API_URL
      ? await fetch(`${CONFIG.API_URL}/api/os`).then(r => r.json())
      : DEMO_OS;
  } catch {
    allOS = DEMO_OS;
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
      || os.equipamento.toLowerCase().includes(search)
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
      <td><div class="os-client">${os.cliente}</div><div class="os-phone">${os.telefone}</div></td>
      <td><span style="font-size:12px">${os.equipamento}</span><div class="os-phone">${os.problema}</div>${valor}</td>
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
function showPage()  { document.querySelectorAll('.nav-link').forEach(l=>l.classList.remove('active')); if(event?.target) event.target.classList.add('active'); }

// ── UTILS ─────────────────────────────────────────────────
function formatDate(s) { if(!s) return '—'; const [y,m,d]=s.split('-'); return `${d}/${m}/${y}`; }
function isOverdue(s)  { return s ? new Date(s)<new Date() : false; }
function setText(id,v) { const el=document.getElementById(id); if(el) el.textContent=v; }
function loadScript(src) {
  return new Promise((res,rej) => {
    if (document.querySelector(`script[src="${src}"]`)) { res(); return; }
    const s = document.createElement('script');
    s.src=src; s.onload=res; s.onerror=rej;
    document.head.appendChild(s);
  });
}
