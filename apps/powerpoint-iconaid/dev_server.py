from __future__ import annotations

import argparse
import http.server
import ssl
from urllib.parse import unquote, urlparse
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
ALLOWED_PATHS = (
    "/apps/powerpoint-iconaid/",
    "/shared/iconaid/catalog.json",
    "/shared/icons/sa_elements.png",
)


class IconAidRequestHandler(http.server.SimpleHTTPRequestHandler):
    def is_allowed(self) -> bool:
        path = unquote(urlparse(self.path).path)
        return any(path == allowed or path.startswith(allowed) for allowed in ALLOWED_PATHS)

    def do_GET(self) -> None:
        if not self.is_allowed():
            self.send_error(404, "IconAid dev server only serves task-pane assets")
            return
        super().do_GET()

    def do_HEAD(self) -> None:
        if not self.is_allowed():
            self.send_error(404, "IconAid dev server only serves task-pane assets")
            return
        super().do_HEAD()


def main() -> None:
    parser = argparse.ArgumentParser(description="Serve the IconAid Office add-in over HTTPS.")
    parser.add_argument("--cert", required=True)
    parser.add_argument("--key", required=True)
    parser.add_argument("--port", type=int, default=3000)
    args = parser.parse_args()

    handler = lambda *args, **kwargs: IconAidRequestHandler(*args, directory=str(ROOT), **kwargs)
    server = http.server.ThreadingHTTPServer(("127.0.0.1", args.port), handler)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=args.cert, keyfile=args.key)
    server.socket = context.wrap_socket(server.socket, server_side=True)
    print(f"Serving IconAid at https://localhost:{args.port}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
