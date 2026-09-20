import http.server
import socketserver
import webbrowser
import os
import sys

PORT = 5000
DIRECTORY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard")

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

def main():
    print("=" * 60)
    print("  TARGET SSW 475 - DASHBOARD DE AUTOMACAO DE DESPESAS")
    print("=" * 60)
    print(f" Servidor iniciado em: http://localhost:{PORT}")
    print(" Abrindo navegador padrao...")
    print(" Para encerrar, feche esta janela ou pressione Ctrl + C")
    print("=" * 60)

    url = f"http://localhost:{PORT}"
    webbrowser.open(url)

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor finalizado.")
            sys.exit(0)

if __name__ == "__main__":
    main()
