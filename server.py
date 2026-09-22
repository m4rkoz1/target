"""TARGET SSW 475 - Central de Controle (backend + frontend estatico).

Serve o dashboard em http://localhost:5000 e expoe a API /api/* para:
  - credenciais do login SSW (sem nunca devolver a senha)
  - upload/selecao da planilha (.xlsx)
  - iniciar/parar a automacao Playwright com log em tempo real
Somente biblioteca padrao do Python.
"""
import datetime
import http.server
import json
import os
import re
import socketserver
import subprocess
import sys
import threading
import webbrowser
from collections import deque

PORT = 5000
FROZEN = bool(getattr(sys, "frozen", False))

def _exe_dir():
    if FROZEN:
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def _recurso(nome):
    """Arquivo empacotado (dashboard, base modelo): _MEIPASS no exe, pasta local no fonte."""
    if FROZEN:
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
        return os.path.join(base, nome)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), nome)

BASE_DIR = _exe_dir()
DIRECTORY = _recurso("dashboard")
ARQ_CRED = os.path.join(BASE_DIR, "credenciais_ssw.json")
ARQ_PLANILHA_ATIVA = os.path.join(BASE_DIR, "planilha_ativa.txt")
PLANILHA_PADRAO = os.path.join(BASE_DIR, "BASE - TARGET.xlsx")
PASTA_UPLOADS = os.path.join(BASE_DIR, "uploads")

# No exe: garante a planilha modelo ao lado do executavel (usuario pode substituir)
if FROZEN and not os.path.exists(PLANILHA_PADRAO):
    try:
        import shutil
        modelo = _recurso("BASE - TARGET.xlsx")
        if os.path.exists(modelo):
            shutil.copyfile(modelo, PLANILHA_PADRAO)
    except Exception:
        pass


def _json(handler, obj, status=200):
    body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def _ler_credenciais():
    try:
        if os.path.exists(ARQ_CRED):
            with open(ARQ_CRED, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _planilha_ativa():
    try:
        if os.path.exists(ARQ_PLANILHA_ATIVA):
            with open(ARQ_PLANILHA_ATIVA, "r", encoding="utf-8") as f:
                p = f.read().strip()
            if p and os.path.exists(p):
                return p
    except Exception:
        pass
    return PLANILHA_PADRAO if os.path.exists(PLANILHA_PADRAO) else ""


class Estado:
    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None
        self.started_at = None
        self.exit_code = None
        self.modo = ""
        self.linhas = deque(maxlen=5000)
        self.seq = 0
        self.senha_memoria = ""
        self.planilha = _planilha_ativa()

    def push_log(self, linha):
        with self.lock:
            self.seq += 1
            self.linhas.append({"seq": self.seq, "txt": linha.rstrip("\n")})
            try:
                with open(os.path.join(BASE_DIR, "automacao_log.txt"), "a", encoding="utf-8") as f:
                    f.write(linha if linha.endswith("\n") else linha + "\n")
            except Exception:
                pass
            return self.seq

    def logs_desde(self, cursor):
        with self.lock:
            novas = [l for l in self.linhas if l["seq"] > cursor]
            return list(novas), self.seq

    def status(self):
        with self.lock:
            running = self.proc is not None and self.proc.poll() is None
            if self.proc is not None and not running and self.exit_code is None:
                self.exit_code = self.proc.poll()
            cred = _ler_credenciais()
            cpf = str(cred.get("cpf", ""))
            return {
                "running": running,
                "pid": self.proc.pid if self.proc else None,
                "exit_code": self.exit_code,
                "started_at": self.started_at,
                "modo": self.modo,
                "planilha": self.planilha,
                "planilha_nome": os.path.basename(self.planilha) if self.planilha else "",
                "credenciais": {
                    "configurado": bool(cred.get("dominio") and cred.get("cpf") and cred.get("usuario")),
                    "dominio": cred.get("dominio", ""),
                    "usuario": cred.get("usuario", ""),
                    "cpf_masc": ("***" + cpf[-4:]) if len(cpf) >= 4 else "",
                    "senha_memoria": bool(self.senha_memoria),
                    "senha_salva": bool(cred.get("senha")),
                },
            }


ESTADO = Estado()


def _ler_dados_planilha():
    """Reusa o parser da automacao (sem abrir navegador)."""
    try:
        sys.path.insert(0, BASE_DIR)
        import automacao_playwright as auto
        regs = auto.carregar_planilha_target(ESTADO.planilha or "BASE - TARGET.xlsx")
        real = auto.carregar_lancamentos_realizados()
        return regs, real
    except Exception as e:
        return None, {"erro": str(e)}


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def log_message(self, *args):
        pass

    # ---------- GET ----------
    def do_GET(self):
        if self.path == "/api/status":
            return _json(self, ESTADO.status())
        if self.path.startswith("/api/logs"):
            cursor = 0
            if "cursor=" in self.path:
                try:
                    cursor = int(self.path.split("cursor=")[1].split("&")[0])
                except Exception:
                    cursor = 0
            novas, seq = ESTADO.logs_desde(cursor)
            return _json(self, {"lines": novas, "cursor": seq})
        if self.path == "/api/planilha":
            regs, real = _ler_dados_planilha()
            if regs is None:
                return _json(self, {"ok": False, "erro": real.get("erro")}, status=500)
            return _json(self, {
                "ok": True,
                "arquivo": os.path.basename(ESTADO.planilha),
                "total": len(regs),
                "concluidos": real,
                "registros": regs,
            })
        return super().do_GET()

    # ---------- POST ----------
    def do_POST(self):
        if self.path == "/api/credentials":
            return self.api_credentials()
        if self.path == "/api/upload":
            return self.api_upload()
        if self.path == "/api/colagem":
            return self.api_colagem()
        if self.path == "/api/start":
            return self.api_start()
        if self.path == "/api/stop":
            return self.api_stop()
        return _json(self, {"ok": False, "erro": "rota desconhecida"}, status=404)

    def _read_json(self):
        try:
            n = int(self.headers.get("Content-Length", 0))
        except Exception:
            n = 0
        raw = self.rfile.read(n) if n else b"{}"
        try:
            return json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            return {}

    def api_credentials(self):
        data = self._read_json()
        dominio = str(data.get("dominio", "")).strip().upper() or "GIA"
        cpf = re.sub(r"\D", "", str(data.get("cpf", "")))
        usuario = str(data.get("usuario", "")).strip().upper()
        senha = str(data.get("senha", "")).strip()
        salvar_senha = bool(data.get("salvar_senha"))
        if not (dominio and cpf and usuario):
            return _json(self, {"ok": False, "erro": "Preencha dominio, CPF e usuario."}, status=400)
        atual = _ler_credenciais()
        novo = {"dominio": dominio, "cpf": cpf, "usuario": usuario}
        if senha and salvar_senha:
            novo["senha"] = senha
        else:
            if atual.get("senha") and not senha:
                novo["senha"] = atual["senha"]
        try:
            with open(ARQ_CRED, "w", encoding="utf-8") as f:
                json.dump(novo, f, indent=2)
        except Exception as e:
            return _json(self, {"ok": False, "erro": f"Falha ao salvar: {e}"}, status=500)
        if senha:
            ESTADO.senha_memoria = senha
        ESTADO.push_log(f"[WEB] Credenciais atualizadas ({dominio}/{usuario}).")
        return _json(self, {"ok": True, "senha_memoria": bool(ESTADO.senha_memoria)})

    def api_upload(self):
        ctype = self.headers.get("Content-Type", "")
        m = re.search(r"boundary=([^;]+)", ctype)
        if not m:
            return _json(self, {"ok": False, "erro": "Envie multipart/form-data."}, status=400)
        try:
            n = int(self.headers.get("Content-Length", 0))
        except Exception:
            n = 0
        if n <= 0 or n > 25 * 1024 * 1024:
            return _json(self, {"ok": False, "erro": "Arquivo vazio ou maior que 25MB."}, status=400)
        try:
            corpo = self.rfile.read(n)
        except Exception as e:
            return _json(self, {"ok": False, "erro": f"Falha ao ler upload: {e}"}, status=400)
        boundary = ("--" + m.group(1).strip().strip('"')).encode("latin-1")
        partes = corpo.split(boundary)
        arquivo_bytes, nome = None, ""
        for p in partes:
            if b'filename="' not in p:
                continue
            try:
                head, dados = p.split(b"\r\n\r\n", 1)
            except ValueError:
                continue
            if b'\r\n--' in dados:
                dados = dados.rsplit(b"\r\n", 1)[0]
                if dados.endswith(b"--"):
                    dados = dados[:-2]
            fn = re.search(br'filename="([^"]+)"', head)
            nome = os.path.basename((fn.group(1).decode("utf-8", "ignore") if fn else "planilha.xlsx"))
            arquivo_bytes = dados
            break
        if not arquivo_bytes:
            return _json(self, {"ok": False, "erro": "Campo 'arquivo' ausente."}, status=400)
        if not nome.lower().endswith((".xlsx", ".xls")):
            return _json(self, {"ok": False, "erro": "Envie um .xlsx."}, status=400)
        nome = re.sub(r"[^A-Za-z0-9_.\- ]+", "_", nome)
        os.makedirs(PASTA_UPLOADS, exist_ok=True)
        destino = os.path.join(PASTA_UPLOADS, nome)
        try:
            with open(destino, "wb") as f:
                f.write(arquivo_bytes)
            with open(ARQ_PLANILHA_ATIVA, "w", encoding="utf-8") as f:
                f.write(destino)
            ESTADO.planilha = destino
        except Exception as e:
            return _json(self, {"ok": False, "erro": f"Falha ao salvar: {e}"}, status=500)
        ESTADO.push_log(f"[WEB] Planilha ativa: {nome}")
        return _json(self, {"ok": True, "arquivo": nome})

    def api_colagem(self):
        """Recebe texto colado do Excel (TSV/CSV), salva como base ativa."""
        data = self._read_json()
        texto = str(data.get("texto", "") or "")
        linhas = [l for l in texto.replace("\r\n", "\n").replace("\r", "\n").split("\n") if l.strip()]
        if not linhas:
            return _json(self, {"ok": False, "erro": "Cole ao menos uma linha da planilha."}, status=400)
        os.makedirs(PASTA_UPLOADS, exist_ok=True)
        nome = "colagem_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".tsv"
        destino = os.path.join(PASTA_UPLOADS, nome)
        try:
            with open(destino, "w", encoding="utf-8") as f:
                f.write("\n".join(linhas) + "\n")
            with open(ARQ_PLANILHA_ATIVA, "w", encoding="utf-8") as f:
                f.write(destino)
            ESTADO.planilha = destino
        except Exception as e:
            return _json(self, {"ok": False, "erro": f"Falha ao salvar: {e}"}, status=500)
        regs, real = _ler_dados_planilha()
        if regs is None:
            return _json(self, {"ok": False, "erro": real.get("erro")}, status=500)
        ESTADO.push_log(f"[WEB] Colagem recebida: {len(regs)} linha(s) valida(s) ({nome}).")
        return _json(self, {"ok": True, "arquivo": nome, "total": len(regs),
                            "concluidos": real, "registros": regs})

    def api_start(self):
        st = ESTADO.status()
        if st["running"]:
            return _json(self, {"ok": False, "erro": "Automacao ja em execucao."}, status=409)
        if not st["credenciais"]["configurado"] and not ESTADO.senha_memoria:
            pass  # a automacao pede no terminal se faltar algo
        data = self._read_json()
        try:
            limite = int(data.get("limite") or 0)
        except Exception:
            limite = 0
        visivel = data.get("visivel", True)
        if isinstance(visivel, str):
            visivel = visivel.lower() not in ("0", "false", "nao", "não", "off")
        else:
            visivel = bool(visivel)
        if FROZEN:
            cmd = [sys.executable, "--worker",
                   "--planilha", ESTADO.planilha or PLANILHA_PADRAO]
        else:
            cmd = [sys.executable, os.path.join(BASE_DIR, "automacao_playwright.py"),
                   "--planilha", ESTADO.planilha or PLANILHA_PADRAO]
        if limite > 0:
            cmd += ["--limite", str(limite)]
        if not visivel:
            cmd += ["--headless"]
        env = dict(os.environ)
        if ESTADO.senha_memoria:
            env["SSW_SENHA"] = ESTADO.senha_memoria
        try:
            proc = subprocess.Popen(cmd, cwd=BASE_DIR, env=env,
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    text=True, bufsize=1)
        except Exception as e:
            return _json(self, {"ok": False, "erro": f"Falha ao iniciar: {e}"}, status=500)
        with ESTADO.lock:
            ESTADO.proc = proc
            ESTADO.exit_code = None
            ESTADO.started_at = datetime.datetime.now().isoformat(timespec="seconds")
            ESTADO.modo = "visivel" if visivel else "oculto"
        ESTADO.push_log(f"[WEB] Automacao iniciada (PID {proc.pid}, limite={limite or 'todos'}, modo={'visivel' if visivel else 'oculto'}).")
        threading.Thread(target=_drenar_saida, args=(proc,), daemon=True).start()
        return _json(self, {"ok": True, "pid": proc.pid})

    def api_stop(self):
        with ESTADO.lock:
            proc = ESTADO.proc
        if proc is None or proc.poll() is not None:
            return _json(self, {"ok": True, "mensagem": "Nada em execucao."})
        try:
            _matar_arvore(proc)
        except Exception as e:
            return _json(self, {"ok": False, "erro": str(e)}, status=500)
        ESTADO.push_log("[WEB] Parada solicitada pela interface (worker + navegador encerrados).")
        return _json(self, {"ok": True})


def _matar_arvore(proc):
    """Encerra o worker E todos os filhos (Chrome do robo)."""
    try:
        if proc is None or proc.poll() is not None:
            return
        if os.name == "nt":
            try:
                subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10)
                return
            except Exception:
                pass
        try:
            proc.terminate()
            proc.wait(timeout=8)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
    except Exception:
        pass


def _encerrar_tudo(origem="saida"):
    try:
        with ESTADO.lock:
            proc = ESTADO.proc
        if proc is not None and proc.poll() is None:
            print(f"\n[ENCERRAR:{origem}] Parando automacao (PID {proc.pid}) e navegador...")
            _matar_arvore(proc)
            print("[ENCERRAR] Processos finalizados.")
    except Exception:
        pass


def _instalar_ctrl_handler():
    """Fecha tudo ao fechar a janela do console / Ctrl+C / logoff (Windows)."""
    try:
        import atexit
        atexit.register(_encerrar_tudo, "atexit")
    except Exception:
        pass
    if os.name != "nt":
        return
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        PHANDLER = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)

        def _handler(ctrl_type):
            # 0=Ctrl+C, 2=fechar janela, 5=logoff, 6=shutdown
            if ctrl_type in (0, 2, 5, 6):
                _encerrar_tudo(f"console-{ctrl_type}")
                return False  # deixa o comportamento padrao seguir (fechar)
            return False

        _ref = PHANDLER(_handler)
        kernel32.SetConsoleCtrlHandler(_ref, True)
        globals()["_ctrl_handler_ref"] = _ref  # evita GC
    except Exception:
        pass


def _drenar_saida(proc):
    try:
        for linha in proc.stdout:
            ESTADO.push_log(linha)
    except Exception:
        pass
    try:
        proc.wait()
        with ESTADO.lock:
            ESTADO.exit_code = proc.poll()
        ESTADO.push_log(f"[WEB] Processo finalizado (exit={proc.poll()}).")
    except Exception:
        pass


def main():
    os.makedirs(DIRECTORY, exist_ok=True)
    _instalar_ctrl_handler()
    print("=" * 60)
    print("  TARGET SSW 475 - CENTRAL DE CONTROLE")
    print("=" * 60)
    print(f" Interface: http://localhost:{PORT}")
    print(" Feche ESTA janela para encerrar tudo (automacao + navegador).")
    print("=" * 60)
    webbrowser.open(f"http://localhost:{PORT}")
    with socketserver.ThreadingTCPServer(("", PORT), Handler) as httpd:
        httpd.allow_reuse_address = True
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            _encerrar_tudo("finally")
            print("Servidor finalizado.")


if __name__ == "__main__":
    if "--worker" in sys.argv:
        sys.argv.remove("--worker")
        sys.path.insert(0, BASE_DIR)
        import automacao_playwright as auto
        auto.executar_automacao()
    else:
        main()
