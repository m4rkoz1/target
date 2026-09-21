"""Diagnostico visual da Tela 475 - deixa o navegador ABERTO p/ voce ver."""
import time, os
from playwright.sync_api import sync_playwright

URL_LOGIN = "https://sistema.ssw.inf.br/bin/ssw0422"
URL_475 = "https://sistema.ssw.inf.br/bin/ssw0094"

def log(m):
    print(m, flush=True)

with sync_playwright() as p:
    prof = os.path.join(os.environ.get("LOCALAPPDATA", "."), "TargetSSW_BrowserProfile")
    ctx = p.chromium.launch_persistent_context(prof, headless=False, channel="chrome", args=["--start-maximized"])
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.on("dialog", lambda d: d.accept())

    log("1) Abrindo login: " + URL_LOGIN)
    page.goto(URL_LOGIN, timeout=60000)
    time.sleep(2)
    log(f"   URL atual: {page.url} | titulo: {page.title()!r}")
    page.screenshot(path="diag_01_login.png")
    log("   Screenshot: diag_01_login.png")

    log("2) FACA O LOGIN na janela que abriu. O script vai detectar sozinho...")
    for _ in range(300):
        time.sleep(1)
        if "ssw0422" not in page.url:
            break
    log(f"   Login OK? URL agora: {page.url}")
    time.sleep(2)

    log("3) Indo para Tela 475: " + URL_475)
    resp = None
    try:
        resp = page.goto(URL_475, timeout=30000)
        log(f"   HTTP status do goto: {resp.status if resp else 'sem-resposta'}")
    except Exception as e:
        log(f"   ERRO no goto: {e}")
    time.sleep(3)
    log(f"   URL atual: {page.url}")
    log(f"   Titulo: {page.title()!r}")
    try:
        txt = page.evaluate("() => document.body ? document.body.innerText.slice(0,500) : 'SEM-BODY'")
        log(f"   Texto (500 chars): {txt!r}")
        n_inputs = page.evaluate("() => document.querySelectorAll('input').length")
        log(f"   Inputs no frame principal: {n_inputs}")
        log(f"   Frames: {len(page.frames)}")
        for i, fr in enumerate(page.frames):
            try:
                log(f"     frame{i}: {fr.url}")
            except Exception:
                pass
    except Exception as e:
        log(f"   ERRO ao ler DOM: {e}")
    page.screenshot(path="diag_02_475.png", full_page=False)
    log("   Screenshot: diag_02_475.png")
    log("4) Navegador ficara ABERTO. Me mande: status, titulo, texto acima + os 2 prints.")
    log("   Para fechar, feche a janela manualmente.")
    # mantem aberto
    for _ in range(600):
        time.sleep(1)
