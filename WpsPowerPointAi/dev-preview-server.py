import mimetypes
import os
import socket
import time
import urllib.error
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse


ROOT = os.path.dirname(os.path.abspath(__file__))
TEMP_ROOT = os.path.abspath(os.environ.get("TEMP") or os.environ.get("TMP") or ROOT)
LOG_PATH = os.path.join(ROOT, "preview-server.log")
DOCMEE_BASE_URL = os.environ.get("WENDUODUO_DOCMEE_BASE_URL", "https://test.docmee.cn").rstrip("/")
HOP_BY_HOP_HEADERS = {
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailer",
    "transfer-encoding",
    "upgrade",
    "host",
    "origin",
    "referer",
    "content-length",
    "accept-encoding",
}
ORIGINAL_GETADDRINFO = socket.getaddrinfo


def ipv4_first_getaddrinfo(*args, **kwargs):
    infos = ORIGINAL_GETADDRINFO(*args, **kwargs)
    return sorted(infos, key=lambda info: 0 if info[0] == socket.AF_INET else 1)


socket.getaddrinfo = ipv4_first_getaddrinfo


def write_log(message):
    line = time.strftime("%Y-%m-%d %H:%M:%S") + " " + message + "\n"
    with open(LOG_PATH, "a", encoding="utf-8") as log_file:
        log_file.write(line)


class WpsPreviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/__preview":
            self.serve_preview(parsed.query)
            return
        if parsed.path == "/__download":
            self.serve_download_proxy(parsed.query)
            return
        super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path.startswith("/__docmee_proxy/"):
            self.serve_docmee_proxy(parsed)
            return
        if parsed.path != "/__log":
            self.send_error(404, "Not found")
            return

        length = int(self.headers.get("Content-Length") or "0")
        body = self.rfile.read(length).decode("utf-8", errors="replace") if length else ""
        write_log("CLIENT " + body)
        self.send_response(204)
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, token, companyId, lang, Accept")
        self.end_headers()

    def proxy_headers(self):
        headers = {}
        for name, value in self.headers.items():
            if name.lower() in HOP_BY_HOP_HEADERS:
                continue
            headers[name] = value
        headers.setdefault("User-Agent", "Wenduoduo-WPS-Preview/1.0")
        return headers

    def stream_upstream(self, request):
        try:
            with urllib.request.urlopen(request, timeout=180) as response:
                self.send_response(response.status)
                for name, value in response.headers.items():
                    if name.lower() in HOP_BY_HOP_HEADERS:
                        continue
                    self.send_header(name, value)
                self.end_headers()
                while True:
                    chunk = response.read(64 * 1024)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
                    self.wfile.flush()
        except urllib.error.HTTPError as error:
            data = error.read()
            self.send_response(error.code)
            content_type = error.headers.get("Content-Type") or "text/plain; charset=utf-8"
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as error:
            message = ("Docmee proxy failed: " + str(error)).encode("utf-8", errors="replace")
            write_log("PROXY ERROR " + str(error))
            self.send_response(502)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(message)))
            self.end_headers()
            self.wfile.write(message)

    def serve_docmee_proxy(self, parsed):
        target_path = parsed.path[len("/__docmee_proxy"):]
        if not target_path.startswith("/"):
            target_path = "/" + target_path
        target_url = DOCMEE_BASE_URL + target_path
        if parsed.query:
            target_url += "?" + parsed.query

        length = int(self.headers.get("Content-Length") or "0")
        body = self.rfile.read(length) if length else b""
        write_log(f"PROXY POST {target_url} bytes={len(body)}")
        request = urllib.request.Request(
            target_url,
            data=body,
            headers=self.proxy_headers(),
            method="POST",
        )
        self.stream_upstream(request)

    def is_safe_external_url(self, raw_url):
        parsed = urlparse(raw_url)
        if parsed.scheme not in ("http", "https"):
            return False
        host = (parsed.hostname or "").lower()
        if not host:
            return False
        if host in ("localhost", "127.0.0.1", "::1") or host.startswith("127."):
            return False
        if host.startswith("10.") or host.startswith("192.168."):
            return False
        if host.startswith("172."):
            parts = host.split(".")
            if len(parts) > 1 and parts[1].isdigit() and 16 <= int(parts[1]) <= 31:
                return False
        return True

    def serve_download_proxy(self, query):
        values = parse_qs(query)
        raw_url = values.get("url", [""])[0]
        target_url = unquote(raw_url)
        if not target_url or not self.is_safe_external_url(target_url):
            self.send_error(400, "Unsupported download URL")
            return

        write_log(f"DOWNLOAD {target_url}")
        request = urllib.request.Request(
            target_url,
            headers={"User-Agent": "Wenduoduo-WPS-Preview/1.0"},
            method="GET",
        )
        self.stream_upstream(request)

    def serve_preview(self, query):
        values = parse_qs(query)
        raw_path = values.get("path", [""])[0]
        file_path = os.path.abspath(unquote(raw_path))
        file_name = os.path.basename(file_path)

        is_preview = file_name.startswith("wenduoduoAI_preview_") and file_name.lower().endswith(".png")
        is_temp_child = os.path.commonpath([TEMP_ROOT, file_path]) == TEMP_ROOT

        if not raw_path or not is_preview or not is_temp_child or not os.path.isfile(file_path):
            write_log(f"PREVIEW 404 raw={raw_path!r} path={file_path!r} is_preview={is_preview} is_temp_child={is_temp_child} exists={os.path.isfile(file_path)}")
            self.send_error(404, "Preview image not found")
            return

        mime_type = mimetypes.guess_type(file_path)[0] or "image/png"
        with open(file_path, "rb") as preview_file:
            data = preview_file.read()

        write_log(f"PREVIEW 200 path={file_path!r} bytes={len(data)}")
        self.send_response(200)
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    write_log("START root={!r} temp={!r}".format(ROOT, TEMP_ROOT))
    server = ThreadingHTTPServer(("127.0.0.1", 3889), WpsPreviewHandler)
    print("Serving WPS plugin with preview support on http://127.0.0.1:3889/")
    server.serve_forever()


if __name__ == "__main__":
    main()
