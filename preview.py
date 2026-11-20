"""在本机启动影视术语标注台。"""
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

PUBLIC = Path(__file__).resolve().parent / "dist"

class Preview(SimpleHTTPRequestHandler):
    def list_directory(self, path):
        self.send_error(404)
        return None

    def send_head(self):
        path = Path(self.translate_path(urlsplit(self.path).path)).resolve()
        if not path.is_relative_to(PUBLIC):
            self.send_error(404)
            return None
        return super().send_head()

    def end_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

if __name__ == "__main__":
    with ThreadingHTTPServer(("127.0.0.1",8792), partial(Preview, directory=str(PUBLIC))) as server:
        print("http://127.0.0.1:8792/",flush=True)
        server.serve_forever()

