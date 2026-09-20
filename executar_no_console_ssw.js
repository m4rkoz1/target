/**
 * TARGET SSW 475 - SCRIPT DE EXECUÇÃO DIRETA POR ID (SEM PRINTS / SEM PIXELS)
 * 
 * Como usar:
 * 1. Abra o SSW na Tela 475: https://sistema.ssw.inf.br/bin/ssw0094
 * 2. Abra o Console do Chrome (F12 -> Console)
 * 3. Cole este script e aperte ENTER!
 * 
 * O script manipula os elementos diretamente por ID com velocidade máxima.
 */

(async function executarAutomacaoTarget() {
  console.log("%c🚀 INICIANDO AUTOMAÇÃO TARGET SSW 475 DIRETO POR ID...", "color: #10b981; font-weight: bold; font-size: 14px;");

  // Dados da planilha BASE - TARGET.xlsx
  const registros = [
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

  const sleep = (ms) => new Promise(res => setTimeout(res, ms));

  function setFieldValue(id, val) {
    const el = document.getElementById(id);
    if (!el) return false;
    el.focus();
    el.value = val;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  function waitForElement(id, timeout = 6000) {
    return new Promise(resolve => {
      const start = Date.now();
      const intv = setInterval(() => {
        if (document.getElementById(id)) {
          clearInterval(intv);
          resolve(true);
        } else if (Date.now() - start > timeout) {
          clearInterval(intv);
          resolve(false);
        }
      }, 100);
    });
  }

  function dismissPopups() {
    const btns = document.querySelectorAll('button, a, input[type="button"]');
    for (let b of btns) {
      const txt = (b.innerText || b.value || '').trim();
      if (txt === 'OK' || txt.includes('Continuar') || txt.includes('OK') || txt === 'Sim') {
        b.click();
        break;
      }
    }
  }

  // Verifica qual linha executar a partir do sessionStorage
  let currentIdx = parseInt(sessionStorage.getItem('ssw_target_current_row') || '0', 10);

  if (currentIdx >= registros.length) {
    console.log("%c🎉 TODOS OS LANÇAMENTOS FORAM CONCLUÍDOS!", "color: #10b981; font-weight: bold; font-size: 16px;");
    sessionStorage.removeItem('ssw_target_current_row');
    return;
  }

  const row = registros[currentIdx];
  console.log(`%c[LINHA ${currentIdx + 1}/${registros.length}] Processando Compra: ${row.ncompra} - R$ ${row.valor}`, "color: #3b82f6; font-weight: bold;");

  // 1. Parte 1 (Tela 475 inicial)
  const fUnidade = document.getElementById('3');
  const fChave = document.getElementById('chave_nfe');
  const fEvento = document.getElementById('5');
  const btn6 = document.getElementById('6');

  if (fUnidade && fChave && fEvento) {
    setFieldValue('3', row.unidade);
    if (typeof window.getFilial === 'function') window.getFilial('3');
    await sleep(150);

    setFieldValue('chave_nfe', row.fornecedor);
    if (typeof window.getEventoChave === 'function') window.getEventoChave(row.fornecedor);
    await sleep(150);

    setFieldValue('5', row.evento);
    if (typeof window.getEvento === 'function') window.getEvento();
    await sleep(200);

    if (btn6) btn6.click();
    else if (typeof window.ajaxEnvia === 'function') window.ajaxEnvia('INC', 1);

    // Aguarda Parte 2
    const ok = await waitForElement('lnk_grava_lancto', 6000);
    if (!ok) {
      console.error("Erro ao carregar a 2ª parte para a linha", currentIdx + 1);
      return;
    }
    await sleep(300);
  }

  // 2. Parte 2
  setFieldValue('4', row.serie);
  setFieldValue('5', row.ncompra);
  setFieldValue('7', row.modelo || '98');
  if (typeof window.busca_modelo === 'function') window.busca_modelo('7');

  setFieldValue('15', row.valor);
  if (typeof window.calc_valor_parcela === 'function') window.calc_valor_parcela();

  setFieldValue('16', row.emissao);
  setFieldValue('17', row.entrada);
  setFieldValue('data_vcto', row.vencimento);
  setFieldValue('data_pgto', row.pagamento);
  
  setFieldValue('mes_competencia', row.competencia);
  if (typeof window.myDoDateTime === 'function') window.myDoDateTime(document.getElementById('mes_competencia'));

  setFieldValue('vlr_parcela', row.vlr_parcela || row.valor);
  if (typeof window.calc_valor_final === 'function') window.calc_valor_final();

  setFieldValue('historico', row.historico);
  await sleep(300);

  // 3. Gravar Lançamento
  const btnGravar = document.getElementById('lnk_grava_lancto');
  if (btnGravar) {
    btnGravar.click();
  } else if (typeof window.verifEnvia === 'function') {
    if (typeof window.consiste_nota === 'function') window.consiste_nota();
    if (typeof window.consiste_parcela === 'function') window.consiste_parcela('INC2');
    window.verifEnvia('INC2', 0);
  }

  await sleep(600);
  dismissPopups();
  await sleep(400);

  // Captura retorno
  const bodyText = document.body.innerText;
  const match = bodyText.match(/N[ºo]\s*do\s*lan[çc]amento:\s*([A-Z0-9]+)/i) || bodyText.match(/(FSP\d+)/i);
  const lancamentoNum = match ? match[1] : 'Gravado';

  console.log(`%c✔ LINHA ${currentIdx + 1} CONCLUÍDA! Lançamento: ${lancamentoNum}`, "color: #10b981; font-weight: bold;");

  // Incrementa e prepara próxima linha
  sessionStorage.setItem('ssw_target_current_row', String(currentIdx + 1));
  await sleep(500);

  // Retorna para tela 475 e continua automaticamente
  window.location.href = 'https://sistema.ssw.inf.br/bin/ssw0094';
})();
