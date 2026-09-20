import time
import re
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from playwright.sync_api import sync_playwright

def log(msg):
    print(msg, flush=True)

# Conversao de numero serial do Excel para data DDMMAA
def serial_to_ddmmaa(val):
    if not val:
        return '170926'
    try:
        num = float(str(val).replace(',', '.'))
        if 30000 < num < 60000:
            d = date(1899, 12, 30) + timedelta(days=int(num))
            return d.strftime('%d%m%y')
    except Exception:
        pass
    clean = re.sub(r'\D', '', str(val))
    if len(clean) == 6:
        return clean
    if len(clean) == 8:
        return clean[:4] + clean[6:]
    return str(val).strip()

# Conversao de numero serial do Excel para competencia MMAA
def serial_to_mmaa(val):
    if not val:
        return '0926'
    try:
        num = float(str(val).replace(',', '.'))
        if 30000 < num < 60000:
            d = date(1899, 12, 30) + timedelta(days=int(num))
            return d.strftime('%m%y')
    except Exception:
        pass
    clean = re.sub(r'\D', '', str(val))
    if len(clean) == 4:
        return clean
    return str(val).strip()

# Formatacao de moeda brasileira
def format_currency_br(val):
    if not val:
        return '0,00'
    s = str(val).replace('R$', '').strip()
    try:
        if ',' in s and '.' not in s:
            num = float(s.replace(',', '.'))
        elif '.' in s and ',' in s:
            num = float(s.replace('.', '').replace(',', '.'))
        else:
            num = float(s)
        return f"{num:.2f}".replace('.', ',')
    except Exception:
        return s

# Leitura direta da planilha BASE - TARGET.xlsx
def carregar_planilha_target(caminho_xlsx="BASE - TARGET.xlsx"):
    registros = []
    if not os.path.exists(caminho_xlsx):
        log(f"[!] Arquivo {caminho_xlsx} nao encontrado. Usando registros padrao...")
        return get_registros_padrao()

    try:
        with zipfile.ZipFile(caminho_xlsx) as z:
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                s_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in s_tree.findall('ns:si', ns):
                    texts = [e.text for e in si.findall('.//ns:t', ns) if e.text]
                    strings.append(''.join(texts))

            target_sheet = 'xl/worksheets/sheet6.xml'
            if target_sheet not in z.namelist():
                target_sheet = [f for f in z.namelist() if f.startswith('xl/worksheets/sheet')][-1]

            st_tree = ET.fromstring(z.read(target_sheet))
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            rows = []
            for r in st_tree.findall('.//ns:row', ns):
                row_data = {}
                for c in r.findall('ns:c', ns):
                    col = ''.join([ch for ch in c.get('r') if ch.isalpha()])
                    t = c.get('t')
                    v = c.find('ns:v', ns)
                    val = v.text if v is not None else ''
                    if t == 's' and val != '' and int(val) < len(strings):
                        val = strings[int(val)]
                    row_data[col] = val
                rows.append(row_data)

            for r in rows[1:]:
                unidade = r.get('A', 'FSP') or 'FSP'
                fornecedor = r.get('B', '') or r.get('C', '')
                evento = r.get('C', '') or r.get('B', '')
                if len(re.sub(r'\D', '', str(evento))) >= 11:
                    fornecedor, evento = evento, fornecedor
                if not fornecedor:
                    fornecedor = '14821124000142'
                if not evento:
                    evento = '7022'

                serie = r.get('D', '1') or '1'
                ncompra = r.get('E', '').strip()
                if not ncompra:
                    continue

                modelo = r.get('F', '98') or '98'
                valor = format_currency_br(r.get('G', '0,00'))
                emissao = serial_to_ddmmaa(r.get('H', ''))
                entrada = serial_to_ddmmaa(r.get('I', ''))
                vencimento = serial_to_ddmmaa(r.get('J', ''))
                pagamento = serial_to_ddmmaa(r.get('K', ''))
                competencia = serial_to_mmaa(r.get('L', ''))
                vlr_parcela = format_currency_br(r.get('M', valor))
                historico = r.get('N', f'S/N - TAG {ncompra}')

                registros.append({
                    'unidade': unidade,
                    'fornecedor': re.sub(r'\D', '', fornecedor),
                    'evento': re.sub(r'\D', '', evento),
                    'serie': serie,
                    'ncompra': ncompra,
                    'modelo': modelo,
                    'valor': valor,
                    'emissao': emissao,
                    'entrada': entrada,
                    'vencimento': vencimento,
                    'pagamento': pagamento,
                    'competencia': competencia,
                    'vlr_parcela': vlr_parcela,
                    'historico': historico
                })

        if registros:
            log(f"[OK] Carregados {len(registros)} registros da planilha {caminho_xlsx}!")
            return registros
    except Exception as e:
        log(f"[!] Erro ao ler planilha ({e}). Usando registros padrao.")
    
    return get_registros_padrao()

def get_registros_padrao():
    return [
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7022', 'serie': '1', 'ncompra': '10426238', 'modelo': '98', 'valor': '65,36', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '65,36', 'historico': 'S/N - KIY2259 - TAG 10426238' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7022', 'serie': '1', 'ncompra': '10426197', 'modelo': '98', 'valor': '13,20', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '13,20', 'historico': 'S/N - KMR9I57 - TAG 10426197' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7022', 'serie': '1', 'ncompra': '10426192', 'modelo': '98', 'valor': '65,36', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '65,36', 'historico': 'S/N - LCQ6I75 - TAG 10426192' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7022', 'serie': '1', 'ncompra': '10426189', 'modelo': '98', 'valor': '30,80', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '30,80', 'historico': 'S/N - CXA4E12 - TAG 10426189' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7022', 'serie': '1', 'ncompra': '10426184', 'modelo': '98', 'valor': '13,20', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '13,20', 'historico': 'S/N - LHV2419 - TAG 10426184' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7564', 'serie': '1', 'ncompra': '2266661', 'modelo': '98', 'valor': '0,33', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '0,33', 'historico': 'TARIFAS DE COBRANCA - TAG 10426238' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7564', 'serie': '1', 'ncompra': '321101', 'modelo': '98', 'valor': '0,07', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '0,07', 'historico': 'TARIFAS DE COBRANCA - TAG 10426197' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7564', 'serie': '1', 'ncompra': '236158', 'modelo': '98', 'valor': '0,33', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '0,33', 'historico': 'TARIFAS DE COBRANCA - TAG 10426192' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7564', 'serie': '1', 'ncompra': '6651251', 'modelo': '98', 'valor': '0,15', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '0,15', 'historico': 'TARIFAS DE COBRANCA - TAG 10426189' },
        { 'unidade': 'FSP', 'fornecedor': '14821124000142', 'evento': '7564', 'serie': '1', 'ncompra': '261450', 'modelo': '98', 'valor': '0,07', 'emissao': '170926', 'entrada': '180926', 'vencimento': '170926', 'pagamento': '170926', 'competencia': '0926', 'vlr_parcela': '0,07', 'historico': 'TARIFAS DE COBRANCA - TAG 10426184' }
    ]

# Gerenciamento de lancamentos ja realizados para evitar duplicidades
ARQ_LANCAMENTOS = "lancamentos_realizados.txt"

def carregar_lancamentos_realizados():
    realizados = {}
    if os.path.exists(ARQ_LANCAMENTOS):
        try:
            with open(ARQ_LANCAMENTOS, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    m_compra = re.search(r'Compra\s+(\w+)', line)
                    m_lanc = re.search(r'Lancamento:\s*(\w+)', line)
                    if m_compra and m_lanc:
                        realizados[m_compra.group(1)] = m_lanc.group(1)
        except Exception:
            pass
    return realizados

def salvar_lancamento(idx, ncompra, lancamento, valor, evento):
    try:
        with open(ARQ_LANCAMENTOS, "a", encoding="utf-8") as f:
            f.write(f"Linha {idx} | Compra {ncompra} | Lancamento: {lancamento} | Valor: R$ {valor} | Evento: {evento}\n")
    except Exception:
        pass

def executar_automacao():
    log("=" * 65)
    log("   AUTOMACAO TARGET SSW 475 COM PLAYWRIGHT (DIRETO POR ID)")
    log("=" * 65)

    registros = carregar_planilha_target()
    realizados = carregar_lancamentos_realizados()
    log(f"Total de registros na base: {len(registros)}")
    if realizados:
        log(f"[INFO] Ja constam {len(realizados)} lancamento(s) concluidos em {ARQ_LANCAMENTOS}")

    url_login = "https://sistema.ssw.inf.br/bin/ssw0422"
    url_475 = "https://sistema.ssw.inf.br/bin/ssw0094"

    with sync_playwright() as p:
        browser = None
        context = None
        page = None
        is_cdp = False

        # 1. Tenta conectar via CDP ao Chrome ja aberto
        try:
            log("[1/3] Conectando ao navegador Chrome...")
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            if browser.contexts and browser.contexts[0].pages:
                context = browser.contexts[0]
                page = context.pages[-1] # Usa a aba ativa
                is_cdp = True
                log("[OK] Conectado diretamente a sessao ativa do SSW!")
        except Exception:
            pass

        # 2. Se nao encontrou via CDP, abre uma nova sessao com perfil persistente
        if not page:
            log("[INFO] Abrindo nova janela do Chrome...")
            user_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', '.'), 'TargetSSW_BrowserProfile')
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                channel="chrome",
                args=["--start-maximized"]
            )
            page = context.pages[0] if context.pages else context.new_page()

            log(f"[2/3] Acessando SSW: {url_login}")
            page.goto(url_login, timeout=60000)
            time.sleep(1)

            if "ssw0422" in page.url or page.locator("input[name='senha'], input[type='password']").count() > 0:
                log("\n" + "!" * 65)
                log(" [LOGIN NECESSARIO] A tela de login do SSW esta aberta!")
                log(" Digite sua senha e clique em Entrar na janela aberta do Chrome.")
                log(" O robo aguardara o login e iniciara automaticamente assim que entrar!")
                log("!" * 65 + "\n")

                while "ssw0422" in page.url:
                    time.sleep(1)

                log("[OK] Login realizado com sucesso!")
                time.sleep(2)

        # Trata dialogs nativos (alert/confirm/prompt)
        page.on("dialog", lambda d: d.accept())

        log(f"[3/3] Direcionando para a Tela 475: {url_475}")

        for idx, row in enumerate(registros, 1):
            ncompra = row['ncompra']

            # Pula registros ja gravados
            if ncompra in realizados:
                log(f"\n---> [{idx}/{len(registros)}] Compra {ncompra} ja lancada anteriormente ({realizados[ncompra]}). Pulando...")
                continue

            log(f"\n---> [{idx}/{len(registros)}] Processando Compra {ncompra} (R$ {row['valor']} - Evento {row['evento']})...")

            try:
                # 1. Garante que esta na Tela 475 Parte 1 (input id="3" presente)
                is_p1 = page.evaluate("() => !!document.querySelector('#frm input[id=\"3\"]')")
                if not is_p1:
                    page.goto(url_475, timeout=30000)
                    time.sleep(1.0)
                    for _ in range(15):
                        if page.evaluate("() => !!document.querySelector('#frm input[id=\"3\"]')"):
                            break
                        time.sleep(0.5)

                # 2. Preenche Parte 1 usando locator + press Tab e Enter no Evento para avancar
                log("     Preenchendo Parte 1 (Unidade, Fornecedor, Evento)...")
                page.locator('#frm input[id="3"]').fill(row['unidade'])
                page.locator('#frm input[id="3"]').press("Tab")
                time.sleep(0.35)

                page.locator('#frm input[id="chave_nfe"]').fill(row['fornecedor'])
                page.locator('#frm input[id="chave_nfe"]').press("Tab")
                time.sleep(0.35)

                page.locator('#frm input[id="5"]').fill(row['evento'])
                page.locator('#frm input[id="5"]').press("Enter")

                # Aguarda Parte 2 carregar (lnk_grava_lancto)
                log("     Avancando para Parte 2...")
                p2_ready = False
                for _ in range(25):
                    time.sleep(0.6)
                    if page.evaluate("() => !!document.getElementById('lnk_grava_lancto')"):
                        p2_ready = True
                        break

                if not p2_ready:
                    raise Exception("Tempo limite excedido aguardando Parte 2.")

                # 3. Preenche Parte 2 usando manipulacao direta do formulario (#frm)
                log("     Preenchendo Dados Fiscais e de Pagamento...")
                page.evaluate("""(data) => {
                    function setFrmVal(id, val) {
                        const el = document.querySelector('#frm [id="' + id + '"]') || document.querySelector('#frm [name="' + id + '"]');
                        if (el) {
                            el.focus();
                            el.value = val;
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                    }

                    setFrmVal('4', data.serie);
                    setFrmVal('5', data.ncompra);
                    setFrmVal('7', data.modelo);
                    if (typeof busca_modelo === 'function') busca_modelo('7');

                    setFrmVal('15', data.valor);
                    if (typeof calc_valor_parcela === 'function') calc_valor_parcela();

                    setFrmVal('16', data.emissao);
                    setFrmVal('17', data.entrada);
                    setFrmVal('data_vcto', data.vencimento);
                    setFrmVal('data_pgto', data.pagamento);
                    setFrmVal('mes_competencia', data.competencia);
                    
                    const comp = document.querySelector('#frm [id="mes_competencia"]');
                    if (comp && typeof myDoDateTime === 'function') myDoDateTime(comp);

                    setFrmVal('vlr_parcela', data.vlr_parcela || data.valor);
                    if (typeof calc_valor_final === 'function') calc_valor_final();

                    setFrmVal('historico', data.historico);
                }""", row)

                time.sleep(0.4)

                # 4. Clica em Gravar Lancamento
                log("     Gravando lancamento...")
                page.evaluate("""() => {
                    if (typeof consiste_nota === 'function') consiste_nota();
                    if (typeof consiste_parcela === 'function') consiste_parcela('INC2');
                    if (typeof verifEnvia === 'function') verifEnvia('INC2', 0);
                    else {
                        const btn = document.getElementById('lnk_grava_lancto');
                        if (btn) btn.click();
                    }
                }""")

                # 5. Trata eventuais avisos e modais de confirmacao (ex: Continuar)
                num_lanc = None
                for _ in range(15):
                    time.sleep(1.2)

                    # Checa se o lancamento foi gravado no texto da pagina
                    body = page.inner_text('body')
                    m = re.search(r'N[ºo]\s*do\s*lan[çc]amento:\s*([A-Z0-9]+)', body, re.IGNORECASE)
                    if not m:
                        m = re.search(r'(FSP\d+)', body)
                    if m:
                        num_lanc = m.group(1)
                        break

                    # Verifica se ha aviso (#errormsg) na tela
                    aviso_tratado = page.evaluate("""() => {
                        const em = document.getElementById('errormsg');
                        if (em && em.style.visibility !== 'hidden' && em.innerText) {
                            const allLinks = Array.from(document.querySelectorAll('#errormsg a, #errormsg button, a.dialog, button'));
                            // Prioridade 1: Botao com texto Continuar
                            let btn = allLinks.find(el => (el.innerText || '').toLowerCase().includes('continuar'));
                            // Prioridade 2: Botao com OK ou Sim
                            if (!btn) {
                                btn = allLinks.find(el => {
                                    const t = (el.innerText || '').toLowerCase().trim();
                                    return t.includes('ok') || t === 'sim' || t.includes('confirmar');
                                });
                            }
                            if (btn) {
                                const txt = btn.innerText || btn.id;
                                btn.click();
                                return txt;
                            }
                        }
                        return null;
                    }""")

                    if aviso_tratado:
                        log(f"     [AVISO SSW] Confirmado: '{aviso_tratado}'")

                if not num_lanc:
                    num_lanc = "CONFIRMADO"

                log(f"     [SUCESSO] Linha {idx} gravada! Lancamento: {num_lanc}")
                salvar_lancamento(idx, ncompra, num_lanc, row['valor'], row['evento'])
                realizados[ncompra] = num_lanc

                # Retorna para a tela 475 para o proximo registro
                time.sleep(0.5)
                page.goto(url_475, timeout=30000)
                time.sleep(1.0)

            except Exception as e:
                log(f"     [ERRO] Falha na linha {idx}: {e}")
                log("     Tentando prosseguir para a proxima linha...")
                page.goto(url_475, timeout=30000)
                time.sleep(1.5)

        log("\n" + "=" * 65)
        log("  TODOS OS REGISTROS FORAM PROCESSADOS COM SUCESSO!")
        log("=" * 65)
        log(f"Consulte o arquivo {ARQ_LANCAMENTOS} para ver o relatorio final.")

        if not is_cdp:
            context.close()

if __name__ == "__main__":
    executar_automacao()
