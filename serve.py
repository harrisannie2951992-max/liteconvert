#!/usr/bin/env python3
"""轻转 LiteConvert 本地服务器 — python3 serve.py [端口]"""
import http.server, mimetypes, os, socketserver, sys, webbrowser

mimetypes.add_type('text/javascript', '.mjs')
mimetypes.add_type('application/manifest+json', '.webmanifest')

os.chdir(os.path.dirname(os.path.abspath(__file__)))
port = int(sys.argv[1]) if len(sys.argv) > 1 else 8973

class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        super().end_headers()
    def log_message(self, *a):
        pass

with socketserver.TCPServer(('127.0.0.1', port), Handler) as httpd:
    url = f'http://127.0.0.1:{port}/'
    print(f'轻转 LiteConvert 已启动: {url}  (Ctrl+C 退出)')
    try:
        webbrowser.open(url)
    except Exception:
        pass
    httpd.serve_forever()
