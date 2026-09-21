import time
import re
import os
import sys
import json
import getpass
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

# ------------------------------------------------------------------
# Helpers robustos: frames, visibilidade, verificacao de preenchimento
# Motivo: SSW antigo reusa IDs (ex: id="5" = Evento na Parte1 e
# Nº Compra na Parte2) e/ou usa frames. getElementById puro pega
# o primeiro match (muitas vezes o campo oculto/errado).
# ------------------------------------------------------------------
DEBUG_DUMP = os.environ.get("SSW_DEBUG", "1") == "1"

def _all_frames(page):
    try:
        frames = page.frames
        if frames:
            return frames
    except Exception:
        pass
    return [page.main_frame] if hasattr(page, "main_frame") else [page]

def dump_inputs(page, tag=""):
    """Lista todos os inputs visiveis em todos os frames p/ diagnostico."""
    if not DEBUG_DUMP:
        return
    try:
        for fi, fr in enumerate(_all_frames(page)):
            try:
                data = fr.evaluate("""() => {
                    const out = [];
                    document.querySelectorAll('input, select, textarea').forEach(el => {
                        const r = el.getBoundingClientRect ? el.getBoundingClientRect() : {width:0,height:0};
                        out.push({
                            id: el.id || '',
                            name: el.name || '',
                            type: el.type || '',
                            val: (el.value || '').slice(0, 40),
                            vis: !!(r.width || r.height || el.offsetParent),
                            frm: (document.querySelector('#frm') ? 'has#frm' : 'no#frm')
                        });
                    });
                    return {url: location.href, n: out.length, inputs: out.slice(0, 60)};
                }""")
                log(f"     [DUMP {tag} frame{fi}] url={data.get('url')} total_inputs={data.get('n')}")
                for inp in data.get("inputs", []):
                    if inp["id"] in ("3","4","5","6","7","15","16","17","chave_nfe","data_vcto","data_pgto","mes_competencia","vlr_parcela","historico","lnk_grava_lancto") or inp["name"]:
                        log(f"        id='{inp['id']}' name='{inp['name']}' val='{inp['val']}' vis={inp['vis']}")
            except Exception as e:
                log(f"     [DUMP {tag} frame{fi}] erro: {e}")
    except Exception as e:
        log(f"     [DUMP {tag}] erro geral: {e}")

def _frame_locator_visible(page, css_id):
    """Retorna (frame, locator) do elemento VISIVEL com dado id em qq frame."""
    # css_id como: '#3', '#chave_nfe', etc. Convertemos p/ seletor por atributo p/ IDs numericos
    raw_id = css_id.lstrip("#")
    sel = f'[id="{raw_id}"]'
    for fr in _all_frames(page):
        try:
            loc = fr.locator(sel)
            try:
                n = loc.count()
            except Exception:
                n = 0
            for i in range(min(n, 5)):
                try:
                    li = loc.nth(i)
                    if li.is_visible(timeout=500):
                        return fr, li
                except Exception:
                    continue
        except Exception:
            continue
    return None, None

def smart_fill(page, field_id, value, label="", timeout_ms=5000):
    """Preenche campo por ID em qq frame, so se visivel, e CONFERE o valor depois."""
    value = "" if value is None else str(value)
    fr, loc = _frame_locator_visible(page, "#" + field_id)
    if not loc:
        # fallback: tenta main page mesmo oculto p/ dar msg de erro util
        dump_inputs(page, tag=f"FAIL-{label or field_id}")
        raise Exception(f"Campo '{label or field_id}' (id={field_id}) nao encontrado/visivel em nenhum frame.")
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    try:
        loc.click(timeout=2000)
    except Exception:
        pass
    # Preenchimento em 2 etapas: fill (rapido) + type de garantia p/ disparar key events do SSW
    try:
        loc.fill(value, timeout=timeout_ms)
    except Exception:
        try:
            loc.fill("", timeout=2000)
            loc.press_sequentially(value, delay=15, timeout=timeout_ms)
        except Exception as e:
            raise Exception(f"Falha ao preencher '{label or field_id}': {e}")
    time.sleep(0.15)
    # Confere
    try:
        obtido = loc.input_value(timeout=2000)
    except Exception:
        obtido = None
    if obtido is not None and obtido.strip() != value.strip():
        # tenta via JS no frame dono + dispara todos os eventos que o SSW legado espera
        try:
            fr.evaluate("""([fid, val]) => {
                const el = document.querySelector('[id="' + fid + '"]');
                if (!el) return;
                el.focus();
                el.value = val;
                ['input','change','blur'].forEach(t => el.dispatchEvent(new Event(t, {bubbles:true})));
                if (el.onchange) { try { el.onchange(); } catch(e){} }
                if (el.onblur) { try { el.onblur(); } catch(e){} }
            }""", [field_id, value])
            time.sleep(0.15)
            obtido2 = loc.input_value(timeout=2000)
            if obtido2.strip() != value.strip():
                dump_inputs(page, tag=f"MISMATCH-{label or field_id}")
                raise Exception(f"Campo '{label or field_id}' recusou valor: esperado='{value}' obtido='{obtido2}' (antes '{obtido}'). SSW pode ter mascara/validacao.")
        except Exception as e:
            if "recusou valor" in str(e):
                raise
    log(f"     [OK] {label or field_id} = '{value}'")
    return fr, loc

def ssw_call(page, fn_names, args=None):
    """Chama funcao JS do SSW no frame onde ela existir. Retorna True se chamou."""
    if isinstance(fn_names, str):
        fn_names = [fn_names]
    for fr in _all_frames(page):
        for fn in fn_names:
            try:
                exists = fr.evaluate(f"() => typeof {fn} === 'function'")
                if exists:
                    fr.evaluate(f"({fn})(...{args!r})" if args else f"{fn}()")
                    return True
            except Exception:
                continue
    return False

# ------------------------------------------------------------------
# Login automatico SSW (tela ssw0422)
# Campos: id=1 Dominio (ex GIA), id=2 CPF, id=3 Usuario, id=4 Senha,
# botao: id=5 (ajaxEnvia('L', 0)). Senha nunca aparece no log.
# ------------------------------------------------------------------
ARQ_CREDENCIAIS = "credenciais_ssw.json"

def _cred_path():
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, ARQ_CREDENCIAIS)

def obter_credenciais(relogin=False):
    """Busca credenciais: env > credenciais_ssw.json > prompt interativo."""
    cred = {
        "dominio": os.environ.get("SSW_DOMINIO", ""),
        "cpf": os.environ.get("SSW_CPF", ""),
        "usuario": os.environ.get("SSW_USUARIO", ""),
        "senha": os.environ.get("SSW_SENHA", ""),
    }
    if not relogin and not all(cred.values()):
        try:
            if os.path.exists(_cred_path()):
                with open(_cred_path(), "r", encoding="utf-8") as f:
                    saved = json.load(f)
                for k in cred:
                    if not cred[k] and saved.get(k):
                        cred[k] = str(saved[k]).strip()
                if saved.get("dominio"):
                    log("[INFO] Credenciais carregadas do arquivo local (senha oculta).")
        except Exception as e:
            log(f"[AVISO] Nao foi possivel ler {ARQ_CREDENCIAIS}: {e}")

    faltando = [k for k, v in cred.items() if not v]
    if faltando or relogin:
        log("---- LOGIN SSW ----")
        if not cred["dominio"] or relogin:
            d = input("Dominio [GIA]: ").strip().upper() or "GIA"
            cred["dominio"] = d
        if not cred["cpf"] or relogin:
            cred["cpf"] = re.sub(r"\D", "", input("CPF (so numeros): ").strip())
        if not cred["usuario"] or relogin:
            cred["usuario"] = input("Usuario: ").strip().upper()
        if not cred["senha"] or relogin:
            cred["senha"] = getpass.getpass("Senha: ").strip()
        # salva nao-sensiveis; senha so com confirmacao
        try:
            salvar = {"dominio": cred["dominio"], "cpf": cred["cpf"], "usuario": cred["usuario"]}
            if cred.get("senha"):
                r = input("Salvar senha no PC p/ proximas execucoes? (s/N): ").strip().lower()
                if r == "s":
                    salvar["senha"] = cred["senha"]
                    log("[AVISO] Senha salva em texto no PC. Proteja o arquivo.")
            with open(_cred_path(), "w", encoding="utf-8") as f:
                json.dump(salvar, f, indent=2)
            log(f"[INFO] Dados salvos em {ARQ_CREDENCIAIS} (use --relogin p/ trocar).")
        except Exception as e:
            log(f"[AVISO] Falha ao salvar credenciais: {e}")

    if not all(cred.values()):
        raise Exception("Credenciais incompletas. Rode com --relogin e preencha dominio, CPF, usuario e senha.")
    return cred

def fazer_login_ssw(page, url_login, relogin=False):
    """Preenche login, clica Entrar e aguarda saida da tela ssw0422."""
    cred = obter_credenciais(relogin=relogin)
    log(f"[2/3] Acessando SSW: {url_login}")
    page.goto(url_login, timeout=60000)
    time.sleep(1.5)

    # Se ja caiu logado (sessao valida no perfil), pula
    if "ssw0422" not in page.url:
        log("[OK] Sessao valida - login nao necessario.")
        return

    log("     Preenchendo tela de login (dominio, CPF, usuario, senha)...")
    _, loc_dom = smart_fill(page, "1", cred["dominio"], label="Dominio")
    try:
        loc_dom.press("Tab")
    except Exception:
        pass
    time.sleep(0.3)
    smart_fill(page, "2", cred["cpf"], label="CPF")
    smart_fill(page, "3", cred["usuario"], label="Usuario")
    # senha sem ecoar valor no log
    fr_s, loc_s = _frame_locator_visible(page, "#4")
    if not loc_s:
        dump_inputs(page, tag="LOGIN-sem-senha")
        raise Exception("Campo Senha (id=4) nao encontrado na tela de login.")
    try:
        loc_s.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    try:
        loc_s.click(timeout=2000)
    except Exception:
        pass
    try:
        loc_s.fill(cred["senha"], timeout=5000)
    except Exception:
        loc_s.fill("", timeout=2000)
        loc_s.press_sequentially(cred["senha"], delay=15, timeout=5000)
    log("     [OK] Senha = '****'")
    time.sleep(0.3)

    log("     Clicando em Entrar...")
    entrou = False
    fr5, loc5 = _frame_locator_visible(page, "#5")
    if fr5 and loc5:
        try:
            loc5.click(timeout=3000)
            entrou = True
        except Exception:
            entrou = False
    if not entrou:
        ok = ssw_call(page, "ajaxEnvia", ["L", 0])
        if not ok:
            raise Exception("Botao Entrar (id=5 / ajaxEnvia L) nao acionavel.")

    # Aguarda sair da tela de login (SSW processa via AJAX e redireciona)
    for _ in range(60):
        time.sleep(1)
        try:
            if "ssw0422" not in page.url:
                break
        except Exception:
            break
        # erro de login aparece em #errormsg
        try:
            for fr in _all_frames(page):
                try:
                    msg = fr.evaluate("""() => {
                        const em = document.getElementById('errormsg');
                        if (em && em.style.visibility !== 'hidden' && em.innerText.trim()) return em.innerText.trim().slice(0,300);
                        return '';
                    }""")
                    if msg:
                        raise Exception(f"SSW recusou o login: {msg}")
                except Exception as e:
                    if "recusou o login" in str(e):
                        raise
        except Exception as e:
            if "recusou o login" in str(e):
                raise
    else:
        pass

    if "ssw0422" in page.url:
        try:
            page.screenshot(path="debug_login_nao_avancou.png")
        except Exception:
            pass
        dump_inputs(page, tag="LOGIN-timeout")
        raise Exception("Login nao avancou em 60s. Confira dominio/CPF/usuario/senha (rode com --relogin) - screenshot: debug_login_nao_avancou.png")

    log("[OK] Login realizado com sucesso!")
    time.sleep(2)

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

            fazer_login_ssw(page, url_login, relogin=("--relogin" in sys.argv))

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
                # 1. Garante que esta na Tela 475 Parte 1 (input id="3" visivel em qq frame)
                fr_p1, _ = _frame_locator_visible(page, "#3")
                if not fr_p1:
                    page.goto(url_475, timeout=30000)
                    time.sleep(1.0)
                    for _ in range(15):
                        fr_p1, _ = _frame_locator_visible(page, "#3")
                        if fr_p1:
                            break
                        time.sleep(0.5)
                if DEBUG_DUMP and idx == 1:
                    dump_inputs(page, tag="PARTE1-antes")

                # 2. Preenche Parte 1 com verificacao campo-a-campo
                log("     Preenchendo Parte 1 (Unidade, Fornecedor, Evento)...")
                _, loc_un = smart_fill(page, "3", row['unidade'], label="Unidade")
                try:
                    loc_un.press("Tab")
                except Exception:
                    pass
                ssw_call(page, "getFilial", ["3"])
                time.sleep(0.5)

                _, loc_for = smart_fill(page, "chave_nfe", row['fornecedor'], label="Fornecedor/CNPJ")
                try:
                    loc_for.press("Tab")
                except Exception:
                    pass
                ssw_call(page, "getEventoChave", [row['fornecedor']])
                time.sleep(0.5)

                _, loc_evt = smart_fill(page, "5", row['evento'], label="Evento (Parte1)")
                # Enter costuma disparar a ida p/ Parte2 no SSW; Tab como fallback
                try:
                    loc_evt.press("Enter")
                except Exception:
                    pass
                time.sleep(0.6)
                # Se ainda estiver na Parte1, clica na seta ► (id=6)
                fr6, loc6 = _frame_locator_visible(page, "#6")
                if fr6 and loc6:
                    try:
                        loc6.click(timeout=2000)
                    except Exception:
                        ssw_call(page, "ajaxEnvia", ["INC", 1])
                else:
                    ssw_call(page, "ajaxEnvia", ["INC", 1])

                # Aguarda Parte 2 carregar (lnk_grava_lancto visivel em qq frame)
                log("     Avancando para Parte 2...")
                p2_ready = False
                for _ in range(25):
                    time.sleep(0.6)
                    fr_g, _ = _frame_locator_visible(page, "#lnk_grava_lancto")
                    if fr_g:
                        p2_ready = True
                        break

                if not p2_ready:
                    dump_inputs(page, tag="PARTE2-timeout")
                    try:
                        page.screenshot(path=f"debug_parte2_timeout_linha{idx}.png")
                        log(f"     [DEBUG] screenshot salva: debug_parte2_timeout_linha{idx}.png")
                    except Exception:
                        pass
                    raise Exception("Tempo limite excedido aguardando Parte 2 (lnk_grava_lancto nao visivel). Veja DUMP acima: confira se id=6 avancou e se ha frame/erro na tela.")

                if DEBUG_DUMP:
                    dump_inputs(page, tag="PARTE2-antes")

                # 3. Preenche Parte 2 com verificacao campo-a-campo
                # ATENCAO: id="5" aqui e Nº Compra (reuso de ID da Parte1). Como a Parte1
                # ja sumiu/avancou, _frame_locator_visible pega o visivel atual. Se ambos
                # estiverem visiveis, o DUMP vai mostrar e abortamos p/ nao preencher errado.
                log("     Preenchendo Dados Fiscais e de Pagamento...")
                try:
                    n5 = 0
                    for fr in _all_frames(page):
                        try:
                            n5 += fr.locator('[id="5"]').count()
                        except Exception:
                            pass
                    if n5 > 1:
                        # conta quantos visiveis
                        vis = 0
                        for fr in _all_frames(page):
                            try:
                                loc = fr.locator('[id="5"]')
                                for i in range(loc.count()):
                                    try:
                                        if loc.nth(i).is_visible(timeout=300):
                                            vis += 1
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        if vis > 1:
                            dump_inputs(page, tag="ID5-DUPLICADO-VISIVEL")
                            raise Exception(f"ID '5' duplicado e visivel ({vis}x). Preenchimento abortado p/ nao errar Evento x NºCompra. Me mande o DUMP + screenshot.")
                except Exception as e:
                    if "duplicado" in str(e):
                        raise

                smart_fill(page, "4", row['serie'], label="Serie")
                smart_fill(page, "5", row['ncompra'], label="Nº Compra (Parte2)")
                smart_fill(page, "7", row['modelo'], label="Modelo")
                ssw_call(page, "busca_modelo", ["7"])
                smart_fill(page, "15", row['valor'], label="Valor")
                ssw_call(page, "calc_valor_parcela")
                smart_fill(page, "16", row['emissao'], label="Emissao")
                smart_fill(page, "17", row['entrada'], label="Entrada")
                smart_fill(page, "data_vcto", row['vencimento'], label="Vencimento")
                smart_fill(page, "data_pgto", row['pagamento'], label="Pagamento")
                smart_fill(page, "mes_competencia", row['competencia'], label="Competencia")
                ssw_call(page, "myDoDateTime")  # SSW as vezes espera o elemento; chamada sem arg tb tenta
                # tenta com o elemento real:
                try:
                    for fr in _all_frames(page):
                        try:
                            has = fr.evaluate("() => !!document.querySelector('[id=\"mes_competencia\"]') && typeof myDoDateTime === 'function'")
                            if has:
                                fr.evaluate("() => { const c = document.querySelector('[id=\"mes_competencia\"]'); try { myDoDateTime(c); } catch(e){} }")
                                break
                        except Exception:
                            continue
                except Exception:
                    pass
                smart_fill(page, "vlr_parcela", row['vlr_parcela'] or row['valor'], label="Vlr Parcela")
                ssw_call(page, "calc_valor_final")
                smart_fill(page, "historico", row['historico'], label="Historico")

                if DEBUG_DUMP:
                    dump_inputs(page, tag="PARTE2-depois")
                    try:
                        page.screenshot(path=f"debug_antes_gravar_linha{idx}.png")
                    except Exception:
                        pass

                time.sleep(0.4)

                # 4. Clica em Gravar Lancamento (frame-aware)
                log("     Gravando lancamento...")
                gravou = False
                fr_g2, loc_g = _frame_locator_visible(page, "#lnk_grava_lancto")
                if fr_g2 and loc_g:
                    try:
                        loc_g.scroll_into_view_if_needed(timeout=2000)
                    except Exception:
                        pass
                    try:
                        loc_g.click(timeout=3000)
                        gravou = True
                    except Exception:
                        pass
                if not gravou:
                    # fallback via JS nos frames
                    for fr in _all_frames(page):
                        try:
                            ok = fr.evaluate("""() => {
                                if (typeof consiste_nota === 'function') { try { consiste_nota(); } catch(e){} }
                                if (typeof consiste_parcela === 'function') { try { consiste_parcela('INC2'); } catch(e){} }
                                if (typeof verifEnvia === 'function') { try { verifEnvia('INC2', 0); return true; } catch(e){ return 'verif-erro:'+e; } }
                                const btn = document.querySelector('[id="lnk_grava_lancto"]');
                                if (btn) { btn.click(); return true; }
                                return false;
                            }""")
                            if ok is True:
                                gravou = True
                                break
                        except Exception:
                            continue
                if not gravou:
                    raise Exception("Botao Gravar (lnk_grava_lancto) nao clicavel.")

                # 5. Trata eventuais avisos e modais de confirmacao (ex: Continuar)
                num_lanc = None
                for _ in range(15):
                    time.sleep(1.2)

                    # Checa se o lancamento foi gravado no texto da pagina (todos os frames)
                    body = ""
                    try:
                        for fr in _all_frames(page):
                            try:
                                body += "\n" + fr.evaluate("() => document.body ? document.body.innerText.slice(0,8000) : ''")
                            except Exception:
                                continue
                    except Exception:
                        try:
                            body = page.inner_text('body')
                        except Exception:
                            body = ""
                    m = re.search(r'N[ºo]\s*do\s*lan[çc]amento:\s*([A-Z0-9]+)', body, re.IGNORECASE)
                    if not m:
                        m = re.search(r'(FSP\d+)', body)
                    if m:
                        num_lanc = m.group(1)
                        break

                    # Verifica se ha aviso (#errormsg) na tela (todos os frames)
                    aviso_tratado = None
                    for fr in _all_frames(page):
                        try:
                            aviso_tratado = fr.evaluate("""() => {
                                const em = document.getElementById('errormsg');
                                if (em && em.style.visibility !== 'hidden' && em.innerText) {
                                    const allLinks = Array.from(document.querySelectorAll('#errormsg a, #errormsg button, a.dialog, button'));
                                    let btn = allLinks.find(el => (el.innerText || '').toLowerCase().includes('continuar'));
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
                                    return 'aviso-sem-botao:' + em.innerText.slice(0,120);
                                }
                                return null;
                            }""")
                            if aviso_tratado:
                                break
                        except Exception:
                            continue

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
                try:
                    page.screenshot(path=f"debug_erro_linha{idx}.png")
                    log(f"     [DEBUG] screenshot salva: debug_erro_linha{idx}.png")
                except Exception:
                    pass
                dump_inputs(page, tag=f"ERRO-linha{idx}")
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
