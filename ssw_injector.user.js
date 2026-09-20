// ==UserScript==
// @name         Target SSW 475 Automação
// @namespace    https://sistema.ssw.inf.br/
// @version      1.0
// @description  Automação ultraleve de despesas para SSW Tela 475
// @match        https://sistema.ssw.inf.br/*
// @grant        none
// ==/UserScript==

(function() {
  'use strict';

  console.log('[SSW Bot] Automação Target carregada na página:', window.location.href);

  const channel = new BroadcastChannel('ssw_automation_channel');

  // Detecta se está na tela de login ssw0422
  if (window.location.href.includes('ssw0422')) {
    console.log('[SSW Bot] Tela de login detectada.');
    channel.postMessage({ type: 'SSW_LOGIN_REQUIRED' });

    window.addEventListener('submit', () => {
      sessionStorage.setItem('ssw_auto_redirect_475', 'true');
    });

    document.addEventListener('click', (e) => {
      const target = e.target.closest('a, button, input[type="submit"]');
      if (target) {
        sessionStorage.setItem('ssw_auto_redirect_475', 'true');
      }
    });
  } else if (window.location.href.includes('ssw0094')) {
    sessionStorage.removeItem('ssw_auto_redirect_475');
    channel.postMessage({ type: 'SSW_READY' });
  } else {
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
      // Retorna para tela 475
      setTimeout(() => {
        window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
      }, 1500);
    }
  };

  // Executa uma linha completa no SSW
  async function executeRow(index, row, baseDelay) {
    // 1. Garantir que estamos na tela 475 Parte 1
    let unidadeInput = document.getElementById('3');
    let chaveInput = document.getElementById('chave_nfe');
    let eventoInput = document.getElementById('5');

    // Se não estiver na tela 1, redireciona
    if (!unidadeInput || !chaveInput) {
      window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
      throw new Error('Aguardando retorno para a tela 475 inicial');
    }

    // Preenche Parte 1
    unidadeInput.focus();
    unidadeInput.value = row.unidade;
    if (typeof window.getFilial === 'function') {
      window.getFilial('3');
    }

    await sleep(200);

    chaveInput.focus();
    chaveInput.value = row.fornecedor;
    if (typeof window.getEventoChave === 'function') {
      window.getEventoChave(row.fornecedor);
    }

    await sleep(200);

    eventoInput.focus();
    eventoInput.value = row.evento;
    if (typeof window.getEvento === 'function') {
      window.getEvento();
    }

    await sleep(300);

    // Clica na seta ► (id 6)
    const btnAvancar = document.getElementById('6');
    if (btnAvancar) {
      btnAvancar.click();
    } else if (typeof window.ajaxEnvia === 'function') {
      window.ajaxEnvia('INC', 1);
    }

    // Aguarda Parte 2 carregar
    const part2Ready = await waitForElement('lnk_grava_lancto', 6000);
    if (!part2Ready) {
      // Verifica se houve aviso de bloqueio/alerta
      const alertPopup = document.querySelector('.aviso, [id*="aviso"], [id*="alerta"]');
      const msg = alertPopup ? alertPopup.innerText : 'Parte 2 não carregou a tempo';
      throw new Error(msg);
    }

    await sleep(baseDelay * 0.4);

    // 2. Preenche Parte 2
    // Série (id 4)
    const fSerie = document.getElementById('4');
    if (fSerie) fSerie.value = row.serie;

    // Nº Compra (id 5)
    const fCompra = document.getElementById('5');
    if (fCompra) fCompra.value = row.ncompra;

    // Modelo (id 7)
    const fModelo = document.getElementById('7');
    if (fModelo) {
      fModelo.value = row.modelo || '98';
      if (typeof window.busca_modelo === 'function') window.busca_modelo('7');
    }

    // Valor (id 15)
    const fValor = document.getElementById('15');
    if (fValor) {
      fValor.value = row.valor;
      if (typeof window.calc_valor_parcela === 'function') window.calc_valor_parcela();
    }

    // Data Emissão (id 16)
    const fEmissao = document.getElementById('16');
    if (fEmissao) fEmissao.value = row.emissao;

    // Data Entrada (id 17)
    const fEntrada = document.getElementById('17');
    if (fEntrada) fEntrada.value = row.entrada;

    // Data Vencimento (id data_vcto)
    const fVcto = document.getElementById('data_vcto');
    if (fVcto) fVcto.value = row.vencimento;

    // Data Pagamento (id data_pgto)
    const fPgto = document.getElementById('data_pgto');
    if (fPgto) fPgto.value = row.pagamento;

    // Competência (id mes_competencia)
    const fComp = document.getElementById('mes_competencia');
    if (fComp) {
      fComp.value = row.competencia;
      if (typeof window.myDoDateTime === 'function') window.myDoDateTime(fComp);
    }

    // Valor da Parcela (id vlr_parcela)
    const fVlrParc = document.getElementById('vlr_parcela');
    if (fVlrParc) {
      fVlrParc.value = row.vlr_parcela || row.valor;
      if (typeof window.calc_valor_final === 'function') window.calc_valor_final();
    }

    // Histórico (id historico)
    const fHist = document.getElementById('historico');
    if (fHist) fHist.value = row.historico;

    await sleep(baseDelay * 0.4);

    // Clica em Gravar Lançamento
    const btnGravar = document.getElementById('lnk_grava_lancto');
    if (btnGravar) {
      btnGravar.click();
    } else if (typeof window.verifEnvia === 'function') {
      if (typeof window.consiste_nota === 'function') window.consiste_nota();
      if (typeof window.consiste_parcela === 'function') window.consiste_parcela('INC2');
      window.verifEnvia('INC2', 0);
    }

    // Aguarda confirmação e captura o número do lançamento
    await sleep(baseDelay * 0.8);

    // Lida com eventuais modais de aviso/confirmação
    dismissPopups();

    // Captura o número do lançamento na tela
    let launchNumber = extractLaunchNumber();

    // Notifica o Dashboard do sucesso
    channel.postMessage({
      type: 'ROW_SUCCESS',
      payload: { index, launchNumber }
    });

    // Retorna para a tela 475 para o próximo registro
    await sleep(baseDelay * 0.5);
    window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
  }

  // Extrai número do lançamento gerado
  function extractLaunchNumber() {
    const text = document.body.innerText;
    const match = text.match(/N[ºo]\s*do\s*lan[çc]amento:\s*([A-Z0-9]+)/i);
    if (match && match[1]) {
      return match[1].trim();
    }
    const match2 = text.match(/(FSP\d+)/i);
    if (match2) return match2[1].trim();
    return 'Lançado (OK)';
  }

  // Fecha popups automáticos se aparecerem
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

  // Helpers
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

  // Floating button inside SSW to open or trigger dashboard
  function injectFloatingWidget() {
    if (document.getElementById('ssw-target-widget')) return;
    const div = document.createElement('div');
    div.id = 'ssw-target-widget';
    div.style.cssText = `
      position: fixed;
      bottom: 20px;
      right: 20px;
      z-index: 999999;
      background: #111827;
      color: #fff;
      padding: 10px 16px;
      border-radius: 24px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.5);
      border: 1px solid #2563eb;
      font-family: Arial, sans-serif;
      font-size: 12px;
      font-weight: bold;
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: pointer;
    `;
    div.innerHTML = `<span style="color:#10b981;">●</span> Bot SSW 475 Ativo`;
    div.onclick = () => {
      alert('Automação SSW 475 está ativa e conectada ao Dashboard!\nCole seus dados no Dashboard para iniciar.');
    };
    document.body.appendChild(div);
  }

  window.addEventListener('load', () => {
    setTimeout(injectFloatingWidget, 1000);
  });
})();
