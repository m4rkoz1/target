// SSW URLs
const SSW_LOGIN_URL = 'https://sistema.ssw.inf.br/bin/ssw0422';
const SSW_TELA_475_URL = 'https://sistema.ssw.inf.br/bin/ssw0094';

// State
const state = {
  rows: [],
  isRunning: false,
  isPaused: false,
  isWaitingLogin: false,
  currentIndex: 0,
  speedDelay: 1000 // default 1.0s
};

// Target Sample Data from BASE - TARGET.xlsx (Planilha2)
const TARGET_DEFAULT_DATA = [
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7022', serie: '1', ncompra: '10426238', modelo: '98', valor: '65,36', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '65,36', historico: 'S/N - KIY2259 - TAG 10426238' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7022', serie: '1', ncompra: '10426197', modelo: '98', valor: '13,20', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '13,20', historico: 'S/N - KMR9I57 - TAG 10426197' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7022', serie: '1', ncompra: '10426192', modelo: '98', valor: '65,36', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '65,36', historico: 'S/N - LCQ6I75 - TAG 10426192' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7022', serie: '1', ncompra: '10426189', modelo: '98', valor: '30,80', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '30,80', historico: 'S/N - CXA4E12 - TAG 10426189' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7022', serie: '1', ncompra: '10426184', modelo: '98', valor: '13,20', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '13,20', historico: 'S/N - LHV2419 - TAG 10426184' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7564', serie: '1', ncompra: '2266661', modelo: '98', valor: '0,33', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '0,33', historico: 'TARIFAS DE COBRANCA - TAG 10426238' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7564', serie: '1', ncompra: '321101', modelo: '98', valor: '0,07', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '0,07', historico: 'TARIFAS DE COBRANCA - TAG 10426197' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7564', serie: '1', ncompra: '236158', modelo: '98', valor: '0,33', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '0,33', historico: 'TARIFAS DE COBRANCA - TAG 10426192' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7564', serie: '1', ncompra: '6651251', modelo: '98', valor: '0,15', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '0,15', historico: 'TARIFAS DE COBRANCA - TAG 10426189' },
  { unidade: 'FSP', fornecedor: '14821124000142', evento: '7564', serie: '1', ncompra: '261450', modelo: '98', valor: '0,07', emissao: '170926', entrada: '180926', vencimento: '170926', pagamento: '170926', competencia: '0926', vlr_parcela: '0,07', historico: 'TARIFAS DE COBRANCA - TAG 10426184' }
];

// BroadcastChannel for cross-tab communication with SSW tab
const sswChannel = new BroadcastChannel('ssw_automation_channel');

// DOM Elements
const pasteZone = document.getElementById('pasteZone');
const tableBody = document.getElementById('tableBody');
const terminalLog = document.getElementById('terminalLog');
const rowCounterBadge = document.getElementById('rowCounterBadge');
const statTotal = document.getElementById('statTotal');
const statPending = document.getElementById('statPending');
const statSuccess = document.getElementById('statSuccess');
const statErrors = document.getElementById('statErrors');
const btnStart = document.getElementById('btnStart');
const btnPause = document.getElementById('btnPause');
const btnStop = document.getElementById('btnStop');
const btnClear = document.getElementById('btnClear');
const btnLoadExample = document.getElementById('btnLoadExample');
const btnClearLog = document.getElementById('btnClearLog');
const btnExportLog = document.getElementById('btnExportLog');
const btnToggleBrowser = document.getElementById('btnToggleBrowser');
const btnBrowserText = document.getElementById('btnBrowserText');
const speedSelect = document.getElementById('speedSelect');
const fileInput = document.getElementById('fileInput');
const btnUploadExcel = document.getElementById('btnUploadExcel');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  addLog('Dashboard pronto para uso.', 'system');
});

// Setup Listeners
function setupEventListeners() {
  // Global Paste
  window.addEventListener('paste', handlePaste);

  // Paste Zone Click Focus
  pasteZone.addEventListener('click', () => pasteZone.focus());

  // Buttons
  btnStart.addEventListener('click', startAutomation);
  btnPause.addEventListener('click', togglePause);
  btnStop.addEventListener('click', stopAutomation);
  btnClear.addEventListener('click', clearTable);
  btnClearLog.addEventListener('click', () => { terminalLog.innerHTML = ''; });
  btnExportLog.addEventListener('click', exportLogCSV);
  btnLoadExample.addEventListener('click', loadDefaultData);
  btnToggleBrowser.addEventListener('click', toggleVisibleBrowser);

  // Speed selection
  speedSelect.addEventListener('change', (e) => {
    const val = e.target.value;
    state.speedDelay = val === 'turbo' ? 500 : (val === 'safe' ? 1800 : 1000);
    addLog(`Velocidade ajustada para: ${val.toUpperCase()} (${state.speedDelay}ms)`, 'info');
  });

  // File Upload
  btnUploadExcel.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });
  fileInput.addEventListener('change', handleFileUpload);

  // Cross-tab message listener from SSW Tab
  sswChannel.onmessage = (event) => {
    handleSSWMessage(event.data);
  };
}

// Formatters
function formatExcelSerialToDate(val) {
  if (!val) return '';
  const num = parseFloat(String(val).replace(',', '.'));
  if (!isNaN(num) && num > 30000 && num < 60000) {
    // Excel date serial
    const date = new Date(Math.round((num - 25569) * 86400 * 1000));
    const d = String(date.getUTCDate()).padStart(2, '0');
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const y = String(date.getUTCFullYear()).slice(-2);
    return `${d}${m}${y}`;
  }
  // Check DD/MM/YYYY or DD/MM/YY
  const dateParts = String(val).match(/(\d{2})\/(\d{2})\/(\d{2,4})/);
  if (dateParts) {
    const y = dateParts[3].slice(-2);
    return `${dateParts[1]}${dateParts[2]}${y}`;
  }
  // Check DDMMAA (already 6 digits)
  const cleanDigits = String(val).replace(/\D/g, '');
  if (cleanDigits.length === 6) return cleanDigits;
  return String(val).trim();
}

function formatExcelSerialToCompetencia(val) {
  if (!val) return '0926';
  const num = parseFloat(String(val).replace(',', '.'));
  if (!isNaN(num) && num > 30000 && num < 60000) {
    const date = new Date(Math.round((num - 25569) * 86400 * 1000));
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const y = String(date.getUTCFullYear()).slice(-2);
    return `${m}${y}`;
  }
  const dateParts = String(val).match(/(\d{2})\/(\d{2})\/(\d{2,4})/);
  if (dateParts) {
    return `${dateParts[2]}${dateParts[3].slice(-2)}`;
  }
  const clean = String(val).replace(/\D/g, '');
  if (clean.length === 4) return clean;
  return String(val).trim();
}

function formatCurrencyBR(val) {
  if (val === undefined || val === null || val === '') return '0,00';
  const cleanStr = String(val).replace('R$', '').trim();
  let num;
  if (cleanStr.includes(',') && !cleanStr.includes('.')) {
    num = parseFloat(cleanStr.replace(',', '.'));
  } else if (cleanStr.includes('.') && cleanStr.includes(',')) {
    num = parseFloat(cleanStr.replace(/\./g, '').replace(',', '.'));
  } else {
    num = parseFloat(cleanStr);
  }
  if (isNaN(num)) return cleanStr;
  return num.toFixed(2).replace('.', ',');
}

// Paste Handler (Ctrl+V)
function handlePaste(e) {
  const clipboardData = e.clipboardData || window.clipboardData;
  if (!clipboardData) return;
  const text = clipboardData.getData('text');
  if (!text || !text.trim()) return;

  parsePastedText(text);
  e.preventDefault();
}

// Parse text from Excel (Tab-separated rows)
function parsePastedText(text) {
  const lines = text.trim().split(/\r\n|\n|\r/);
  if (!lines.length) return;

  let addedCount = 0;
  lines.forEach((line) => {
    if (!line.trim()) return;
    const cols = line.split('\t').map(c => c.trim());
    
    // Ignore header row if pasted
    if (cols[0].toLowerCase().includes('unidade') || cols[1]?.toLowerCase().includes('fornecedor')) {
      return;
    }

    // Sequence mapping:
    // 0: Unidade
    // 1: Fornecedor (ou Evento se ordem invertida)
    // 2: Evento (ou Fornecedor)
    // 3: Série
    // 4: Nº Compra
    // 5: Modelo
    // 6: Valor
    // 7: Data Emissão
    // 8: Data Entrada
    // 9: Data Vencimento
    // 10: Data Pagamento
    // 11: Competência
    // 12: Valor da Parcela
    // 13: Histórico

    let unidade = cols[0] || 'FSP';
    let col1 = cols[1] || '';
    let col2 = cols[2] || '';

    let fornecedor = '';
    let evento = '';

    // Auto-detect CNPJ vs Evento (CNPJ tem 14 dígitos, Evento tem 4 dígitos)
    if (col1.replace(/\D/g, '').length >= 11) {
      fornecedor = col1.replace(/\D/g, '');
      evento = col2.replace(/\D/g, '') || '7022';
    } else if (col2.replace(/\D/g, '').length >= 11) {
      fornecedor = col2.replace(/\D/g, '');
      evento = col1.replace(/\D/g, '') || '7022';
    } else {
      fornecedor = col1 || '14821124000142';
      evento = col2 || '7022';
    }

    const serie = cols[3] || '1';
    const ncompra = cols[4] || '';
    const modelo = cols[5] || '98';
    const valor = formatCurrencyBR(cols[6]);
    const emissao = formatExcelSerialToDate(cols[7]) || '170926';
    const entrada = formatExcelSerialToDate(cols[8]) || '180926';
    const vencimento = formatExcelSerialToDate(cols[9]) || emissao;
    const pagamento = formatExcelSerialToDate(cols[10]) || emissao;
    const competencia = formatExcelSerialToCompetencia(cols[11]) || '0926';
    const vlr_parcela = cols[12] ? formatCurrencyBR(cols[12]) : valor;
    const historico = cols[13] || `S/N - TAG ${ncompra}`;

    if (ncompra || cols.length >= 4) {
      state.rows.push({
        id: Date.now() + Math.random(),
        unidade,
        fornecedor,
        evento,
        serie,
        ncompra,
        modelo,
        valor,
        emissao,
        entrada,
        vencimento,
        pagamento,
        competencia,
        vlr_parcela,
        historico,
        status: 'pending', // pending, running, success, error
        launchNumber: '-',
        errorMsg: ''
      });
      addedCount++;
    }
  });

  if (addedCount > 0) {
    addLog(`Importadas ${addedCount} linhas com sucesso da área de transferência.`, 'success');
    renderTable();
    updateStats();
  } else {
    addLog('Nenhuma linha válida identificada no texto colado.', 'warning');
  }
}

// Load default Target spreadsheet data
function loadDefaultData() {
  state.rows = TARGET_DEFAULT_DATA.map((item, idx) => ({
    id: Date.now() + idx,
    ...item,
    status: idx === 0 ? 'success' : 'pending',
    launchNumber: idx === 0 ? 'FSP236457' : '-',
    errorMsg: ''
  }));
  addLog('Base oficial do TARGET carregada com 10 registros.', 'info');
  renderTable();
  updateStats();
}

// Render Table
function renderTable() {
  if (state.rows.length === 0) {
    tableBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="18">
          <div class="empty-state">
            <div class="empty-state-icon">📊</div>
            <h3>Nenhum dado importado</h3>
            <p>Copie as células do Excel e aperte <strong>Ctrl + V</strong> em qualquer lugar, ou clique em "Carregar Planilha TARGET".</p>
          </div>
        </td>
      </tr>
    `;
    rowCounterBadge.textContent = '0 linhas';
    return;
  }

  rowCounterBadge.textContent = `${state.rows.length} linhas`;

  tableBody.innerHTML = state.rows.map((row, idx) => {
    let statusBadge = '';
    let rowClass = '';

    switch (row.status) {
      case 'running':
        statusBadge = '<span class="badge badge-running">Processando</span>';
        rowClass = 'active-row';
        break;
      case 'success':
        statusBadge = '<span class="badge badge-success">Concluído</span>';
        rowClass = 'success-row';
        break;
      case 'error':
        statusBadge = '<span class="badge badge-error">Erro</span>';
        rowClass = 'error-row';
        break;
      default:
        statusBadge = '<span class="badge badge-pending">Pendente</span>';
    }

    return `
      <tr class="${rowClass}" id="row-${row.id}">
        <td>${idx + 1}</td>
        <td>${statusBadge}</td>
        <td class="launch-number">${row.launchNumber || '-'}</td>
        <td>${row.unidade}</td>
        <td>${row.fornecedor}</td>
        <td>${row.evento}</td>
        <td>${row.serie}</td>
        <td><strong>${row.ncompra}</strong></td>
        <td>${row.modelo}</td>
        <td>R$ ${row.valor}</td>
        <td>${row.emissao}</td>
        <td>${row.entrada}</td>
        <td>${row.vencimento}</td>
        <td>${row.pagamento}</td>
        <td>${row.competencia}</td>
        <td>R$ ${row.vlr_parcela}</td>
        <td title="${row.historico}">${row.historico}</td>
        <td>
          <button class="btn-row-del" onclick="deleteRow(${row.id})" title="Remover linha">✕</button>
        </td>
      </tr>
    `;
  }).join('');
}

// Delete row
window.deleteRow = function(id) {
  state.rows = state.rows.filter(r => r.id !== id);
  renderTable();
  updateStats();
};

// Clear Table
function clearTable() {
  if (state.isRunning) return;
  state.rows = [];
  renderTable();
  updateStats();
  addLog('Tabela limpa.', 'info');
}

// Update stats
function updateStats() {
  const total = state.rows.length;
  const pending = state.rows.filter(r => r.status === 'pending').length;
  const success = state.rows.filter(r => r.status === 'success').length;
  const errors = state.rows.filter(r => r.status === 'error').length;

  statTotal.textContent = total;
  statPending.textContent = pending;
  statSuccess.textContent = success;
  statErrors.textContent = errors;
}

// Add Log Entry
function addLog(msg, type = 'info') {
  const now = new Date();
  const time = now.toTimeString().split(' ')[0];
  const entry = document.createElement('div');
  entry.className = `log-entry ${type}`;
  entry.textContent = `[${time}] ${msg}`;
  terminalLog.appendChild(entry);
  terminalLog.scrollTop = terminalLog.scrollHeight;
}

// Start Automation
function startAutomation() {
  const pendingRows = state.rows.filter(r => r.status === 'pending');
  if (pendingRows.length === 0) {
    addLog('Não há registros pendentes para processar.', 'warning');
    return;
  }

  // Abre janela do SSW se não estiver aberta
  if (!sswBrowserWindow || sswBrowserWindow.closed) {
    state.isWaitingLogin = true;
    addLog('Acessando o SSW... Verificando login do usuário.', 'system');
    sswBrowserWindow = window.open(
      SSW_LOGIN_URL,
      'SSW_475_WINDOW',
      'width=1280,height=800,menubar=no,toolbar=yes,location=yes,status=yes,scrollbars=yes,resizable=yes'
    );
    updateConnectionStatus('waiting_login', 'Aguardando Login no SSW');
    addLog('🔑 Faça o login no SSW na janela que abriu. Assim que você entrar, a automação iniciará automaticamente!', 'warning');
  }

  state.isRunning = true;
  state.isPaused = false;
  btnStart.disabled = true;
  btnPause.disabled = false;
  btnStop.disabled = false;

  addLog(`Automação engatilhada para ${pendingRows.length} registros pendentes...`, 'system');

  // Se já estiver logado / tela pronta, envia imediatamente
  if (!state.isWaitingLogin) {
    sendNextRowToSSW();
  }
}

// Send Next Row to SSW
function sendNextRowToSSW() {
  if (!state.isRunning || state.isPaused) return;

  const nextIndex = state.rows.findIndex(r => r.status === 'pending');
  if (nextIndex === -1) {
    // Finished all
    state.isRunning = false;
    btnStart.disabled = false;
    btnPause.disabled = true;
    btnStop.disabled = true;
    addLog('🎉 TODOS OS REGISTROS FORAM PROCESSADOS COM SUCESSO!', 'success');
    return;
  }

  state.currentIndex = nextIndex;
  const row = state.rows[nextIndex];
  row.status = 'running';
  renderTable();
  updateStats();

  addLog(`[Linha ${nextIndex + 1}/${state.rows.length}] Enviando Nº Compra ${row.ncompra} para SSW Tela 475...`, 'info');

  // Send message to SSW Tab
  sswChannel.postMessage({
    type: 'PROCESS_ROW',
    payload: {
      index: nextIndex,
      row: row,
      delay: state.speedDelay
    }
  });
}

// Update Connection Status Pill
function updateConnectionStatus(type, text) {
  const dot = document.querySelector('#connectionStatus .status-dot');
  const txt = document.getElementById('connectionText');
  txt.textContent = text;

  dot.className = 'status-dot';
  if (type === 'online') dot.classList.add('online');
  else if (type === 'waiting_login') dot.classList.add('waiting');
  else if (type === 'error') dot.classList.add('error');
}

// Handle Messages from SSW tab
function handleSSWMessage(data) {
  if (!data || !data.type) return;

  switch (data.type) {
    case 'SSW_LOGIN_REQUIRED': {
      state.isWaitingLogin = true;
      updateConnectionStatus('waiting_login', 'Aguardando Login');
      addLog('🔑 Tela de login do SSW detectada. Faça login para iniciar.', 'warning');
      break;
    }

    case 'SSW_LOGGED_IN': {
      addLog('🎉 Login no SSW realizado com sucesso! Redirecionando para a Tela 475...', 'success');
      updateConnectionStatus('online', 'Login Concluído');
      break;
    }

    case 'SSW_READY': {
      updateConnectionStatus('online', 'Conectado à Tela 475');
      addLog('✔ Conexão estabelecida com a Tela 475 do SSW.', 'success');
      
      // Se estava aguardando login para iniciar a automação, inicia agora!
      if (state.isRunning && state.isWaitingLogin) {
        state.isWaitingLogin = false;
        addLog('🚀 Login concluído! Iniciando o processamento das despesas...', 'system');
        setTimeout(sendNextRowToSSW, 800);
      }
      break;
    }

    case 'ROW_SUCCESS': {
      const { index, launchNumber } = data.payload;
      if (state.rows[index]) {
        state.rows[index].status = 'success';
        state.rows[index].launchNumber = launchNumber || 'OK';
        addLog(`✔ Linha ${index + 1} gravada com sucesso! Lançamento: ${launchNumber}`, 'success');
        renderTable();
        updateStats();
      }
      setTimeout(sendNextRowToSSW, state.speedDelay);
      break;
    }

    case 'ROW_ERROR': {
      const { index, errorMsg } = data.payload;
      if (state.rows[index]) {
        state.rows[index].status = 'error';
        state.rows[index].errorMsg = errorMsg;
        addLog(`✖ Erro na linha ${index + 1}: ${errorMsg}. Prosseguindo para a próxima...`, 'error');
        renderTable();
        updateStats();
      }
      setTimeout(sendNextRowToSSW, state.speedDelay);
      break;
    }
  }
}

// Pause / Toggle
function togglePause() {
  state.isPaused = !state.isPaused;
  btnPause.innerHTML = state.isPaused ? '<span class="btn-icon">▶</span> Continuar' : '<span class="btn-icon">⏸</span> Pausar';
  addLog(state.isPaused ? 'Automação pausada pelo operador.' : 'Automação retomada.', 'warning');
  if (!state.isPaused && state.isRunning) {
    sendNextRowToSSW();
  }
}

// Stop
function stopAutomation() {
  state.isRunning = false;
  state.isPaused = false;
  btnStart.disabled = false;
  btnPause.disabled = true;
  btnStop.disabled = true;
  btnPause.innerHTML = '<span class="btn-icon">⏸</span> Pausar';
  
  if (state.rows[state.currentIndex]?.status === 'running') {
    state.rows[state.currentIndex].status = 'pending';
  }
  renderTable();
  updateStats();
  addLog('Automação interrompida.', 'warning');
}

// Toggle / Open Visible Browser Window for SSW 475
let sswBrowserWindow = null;
function toggleVisibleBrowser() {
  const sswUrl = 'https://sistema.ssw.inf.br/bin/ssw0094';
  
  if (!sswBrowserWindow || sswBrowserWindow.closed) {
    // Open visible window side-by-side or popup
    sswBrowserWindow = window.open(
      sswUrl,
      'SSW_475_WINDOW',
      'width=1280,height=800,menubar=no,toolbar=yes,location=yes,status=yes,scrollbars=yes,resizable=yes'
    );
    btnBrowserText.textContent = 'Modo Visível: Ativo';
    btnToggleBrowser.style.borderColor = '#10b981';
    btnToggleBrowser.style.color = '#10b981';
    btnToggleBrowser.style.boxShadow = '0 0 12px rgba(16, 185, 129, 0.3)';
    addLog('👁️ Modo Visível ativado! Janela do SSW 475 aberta e visível na tela.', 'success');
  } else {
    // Focus existing window
    sswBrowserWindow.focus();
    addLog('👁️ Janela visível do SSW 475 focada na tela.', 'info');
  }
}

// Export Log to CSV
function exportLogCSV() {
  if (state.rows.length === 0) {
    addLog('Nenhum dado para exportar.', 'warning');
    return;
  }

  const headers = ['Linha', 'Status', 'No_Lancamento', 'Unidade', 'Fornecedor', 'Evento', 'Serie', 'No_Compra', 'Modelo', 'Valor', 'Emissao', 'Entrada', 'Vencimento', 'Pagamento', 'Competencia', 'Vlr_Parcela', 'Historico', 'Erro'];
  const rowsCSV = state.rows.map((r, i) => [
    i + 1,
    r.status,
    r.launchNumber || '',
    r.unidade,
    r.fornecedor,
    r.evento,
    r.serie,
    r.ncompra,
    r.modelo,
    r.valor,
    r.emissao,
    r.entrada,
    r.vencimento,
    r.pagamento,
    r.competencia,
    r.vlr_parcela,
    `"${r.historico.replace(/"/g, '""')}"`,
    `"${(r.errorMsg || '').replace(/"/g, '""')}"`
  ].join(';'));

  const csvContent = '\uFEFF' + [headers.join(';'), ...rowsCSV].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Relatorio_Lancamentos_SSW_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  addLog('Relatório CSV exportado com sucesso.', 'success');
}

// Handle File Upload (.txt / .csv / paste dump)
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const content = evt.target.result;
    parsePastedText(content);
  };
  reader.readAsText(file);
}
