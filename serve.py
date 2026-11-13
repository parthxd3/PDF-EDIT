"""Tiny local web server for XD3 PDF Editor (used by Start PDF Editor.bat and Run Tests.bat).
Like `python -m http.server`, but every file is sent with Cache-Control: no-cache, so an updated editor shows at once instead of
the browser keeping an old copy for a while.   Usage: python serve.py [port]"""
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

class Handler(SimpleHTTPRequestHandler):
    extensions_map = { **SimpleHTTPRequestHandler.extensions_map, '.js': 'text/javascript', '.webmanifest': 'application/manifest+json', '.gz': 'application/gzip', '.ttf': 'font/ttf', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream', '.wasm': 'application/wasm', '.mjs': 'text/javascript' }
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        # cross-origin isolation: lets the offline AI use several processor cores (everything the page loads is its own, so nothing breaks)
        self.send_header('Cross-Origin-Opener-Policy', 'same-origin')
        self.send_header('Cross-Origin-Embedder-Policy', 'require-corp')
        super().end_headers()
    def translate_path(self, path):  # downloads/ is served from dist/downloads, so "Download for Windows" also works on this local copy
        import os
        p = path.split('?')[0].split('#')[0]
        if p.startswith('/downloads/') and not os.path.isdir('downloads'):
            path = '/dist' + path
        return super().translate_path(path)
    def log_message(self, fmt, *args):  # keep the window quiet
        pass
    def do_PUT(self):  # the self-test drops files it saved into tests/out/ (only there), so Word can check them
        import os, re
        name = self.path.split('?')[0]
        if not re.fullmatch(r'/tests/out/[\w.-]+', name):
            self.send_error(403); return
        os.makedirs('tests/out', exist_ok=True)
        with open(name.lstrip('/'), 'wb') as f: f.write(self.rfile.read(int(self.headers.get('Content-Length', 0))))
        self.send_response(204); self.end_headers()

port = int(sys.argv[1]) if len(sys.argv) > 1 else 5391
ThreadingHTTPServer(('127.0.0.1', port), Handler).serve_forever()
