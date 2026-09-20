console.log('[SSW Bot] Automação Target carregada na página:', window.location.href);

const channel = new BroadcastChannel('ssw_automation_channel');

// Detecta se está na tela de login ssw0422
if (window.location.href.includes('ssw0422')) {
  console.log('[SSW Bot] Tela de login detectada.');
  channel.postMessage({ type: 'SSW_LOGIN_REQUIRED' });

  // Monitora envio do formulário de login para redirecionar para tela 475
  window.addEventListener('submit', () => {
    console.log('[SSW Bot] Formulário de login enviado. Aguardando autenticação...');
    sessionStorage.setItem('ssw_auto_redirect_475', 'true');
  });

  // Também monitora links de login
  document.addEventListener('click', (e) => {
    const target = e.target.closest('a, button, input[type="submit"]');
    if (target) {
      sessionStorage.setItem('ssw_auto_redirect_475', 'true');
    }
  });
} else if (window.location.href.includes('ssw0094')) {
  // Tela 475 carregada e pronta
  sessionStorage.removeItem('ssw_auto_redirect_475');
  channel.postMessage({ type: 'SSW_READY' });
} else {
  // Outra tela do sistema pós-login: redireciona para a tela 475 se estava em fluxo de login
  if (sessionStorage.getItem('ssw_auto_redirect_475') === 'true') {
    sessionStorage.removeItem('ssw_auto_redirect_475');
    channel.postMessage({ type: 'SSW_LOGGED_IN' });
    setTimeout(() => {
      window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
    }, 500);
  }
}

// Escuta comandos do Dashboard
channel.onmessage = async (event) => {
  const data = event.data;
  if (!data || data.type !== 'PROCESS_ROW') return;

  const { index, row, delay } = data.payload;
  console.log(`[SSW Bot] Iniciando processamento da linha ${index + 1}:`, row);

  try {
    await executeRow(index, row, delay || 1000);
  } catch (err) {
    console.error('[SSW Bot] Erro durante execução:', err);
    channel.postMessage({
      type: 'ROW_ERROR',
      payload: { index, errorMsg: err.message || 'Falha no processamento' }
    });
    setTimeout(() => {
      window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
    }, 1500);
  }
};

// Executa uma linha completa no SSW
async function executeRow(index, row, baseDelay) {
  // 1. Tela 475 Parte 1
  let unidadeInput = document.getElementById('3');
  let chaveInput = document.getElementById('chave_nfe');
  let eventoInput = document.getElementById('5');

  if (!unidadeInput || !chaveInput) {
    window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
    throw new Error('Retornando para tela 475...');
  }

  // Preenche Parte 1
  unidadeInput.focus();
  unidadeInput.value = row.unidade;
  if (typeof window.getFilial === 'function') window.getFilial('3');

  await sleep(200);

  chaveInput.focus();
  chaveInput.value = row.fornecedor;
  if (typeof window.getEventoChave === 'function') window.getEventoChave(row.fornecedor);

  await sleep(200);

  eventoInput.focus();
  eventoInput.value = row.evento;
  if (typeof window.getEvento === 'function') window.getEvento();

  await sleep(300);

  // Botão ► (id 6)
  const btnAvancar = document.getElementById('6');
  if (btnAvancar) {
    btnAvancar.click();
  } else if (typeof window.ajaxEnvia === 'function') {
    window.ajaxEnvia('INC', 1);
  }

  // Aguarda Parte 2 carregar
  const part2Ready = await waitForElement('lnk_grava_lancto', 6000);
  if (!part2Ready) {
    const alertPopup = document.querySelector('.aviso, [id*="aviso"], [id*="alerta"]');
    const msg = alertPopup ? alertPopup.innerText : 'Parte 2 não carregou a tempo';
    throw new Error(msg);
  }

  await sleep(baseDelay * 0.4);

  // 2. Preenche Parte 2
  const fSerie = document.getElementById('4');
  if (fSerie) fSerie.value = row.serie;

  const fCompra = document.getElementById('5');
  if (fCompra) fCompra.value = row.ncompra;

  const fModelo = document.getElementById('7');
  if (fModelo) {
    fModelo.value = row.modelo || '98';
    if (typeof window.busca_modelo === 'function') window.busca_modelo('7');
  }

  const fValor = document.getElementById('15');
  if (fValor) {
    fValor.value = row.valor;
    if (typeof window.calc_valor_parcela === 'function') window.calc_valor_parcela();
  }

  const fEmissao = document.getElementById('16');
  if (fEmissao) fEmissao.value = row.emissao;

  const fEntrada = document.getElementById('17');
  if (fEntrada) fEntrada.value = row.entrada;

  const fVcto = document.getElementById('data_vcto');
  if (fVcto) fVcto.value = row.vencimento;

  const fPgto = document.getElementById('data_pgto');
  if (fPgto) fPgto.value = row.pagamento;

  const fComp = document.getElementById('mes_competencia');
  if (fComp) {
    fComp.value = row.competencia;
    if (typeof window.myDoDateTime === 'function') window.myDoDateTime(fComp);
  }

  const fVlrParc = document.getElementById('vlr_parcela');
  if (fVlrParc) {
    fVlrParc.value = row.vlr_parcela || row.valor;
    if (typeof window.calc_valor_final === 'function') window.calc_valor_final();
  }

  const fHist = document.getElementById('historico');
  if (fHist) fHist.value = row.historico;

  await sleep(baseDelay * 0.4);

  // Clica em Gravar lançamento
  const btnGravar = document.getElementById('lnk_grava_lancto');
  if (btnGravar) {
    btnGravar.click();
  } else if (typeof window.verifEnvia === 'function') {
    if (typeof window.consiste_nota === 'function') window.consiste_nota();
    if (typeof window.consiste_parcela === 'function') window.consiste_parcela('INC2');
    window.verifEnvia('INC2', 0);
  }

  // Aguarda confirmação
  await sleep(baseDelay * 0.8);
  dismissPopups();

  const launchNumber = extractLaunchNumber();

  channel.postMessage({
    type: 'ROW_SUCCESS',
    payload: { index, launchNumber }
  });

  await sleep(baseDelay * 0.5);
  window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
}

function extractLaunchNumber() {
  const text = document.body.innerText;
  const match = text.match(/N[ºo]\s*do\s*lan[çc]amento:\s*([A-Z0-9]+)/i);
  if (match && match[1]) return match[1].trim();
  const match2 = text.match(/(FSP\d+)/i);
  if (match2) return match2[1].trim();
  return 'Lançado (OK)';
}

function dismissPopups() {
  const btns = document.querySelectorAll('button, a, input[type="button"]');
  for (let b of btns) {
    const txt = (b.innerText || b.value || '').trim();
    if (txt === 'OK' || txt.includes('Continuar') || txt === 'Sim') {
      b.click();
      break;
    }
  }
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function waitForElement(id, timeout = 5000) {
  return new Promise(resolve => {
    const startTime = Date.now();
    const interval = setInterval(() => {
      if (document.getElementById(id)) {
        clearInterval(interval);
        resolve(true);
      } else if (Date.now() - startTime > timeout) {
        clearInterval(interval);
        resolve(false);
      }
    }, 100);
  });
}
