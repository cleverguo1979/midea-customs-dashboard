#!/usr/bin/env python3
"""
华东口岸报关看板 - 本地服务器
启动后可通过本机浏览器和局域网内其他设备访问
"""
import http.server
import socket
import os
import sys

DEFAULT_PORT = 8888

def find_free_port(start_port):
    """找到一个可用端口"""
    for port in range(start_port, start_port + 100):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(('', port))
            s.close()
            return port
        except OSError:
            continue
    return start_port  # fallback

def get_local_ip():
    """获取本机局域网IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '127.0.0.1'

def get_local_hostname():
    """获取本机名"""
    try:
        return socket.gethostname() + '.local'
    except:
        return 'localhost'

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=os.path.dirname(os.path.abspath(__file__)), **kwargs)

    def end_headers(self):
        # Disable caching for development — always load latest
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        print(f"  {args[0]}")

if __name__ == '__main__':
    PORT = find_free_port(DEFAULT_PORT)
    local_ip = get_local_ip()
    hostname = get_local_hostname()

    print()
    print("  ╔══════════════════════════════════════════╗")
    print("  ║    📊 华东口岸报关看板 — 服务已启动   ║")
    print("  ╠══════════════════════════════════════════╣")
    print("  ║                                          ║")
    print(f"  ║  本机访问:                               ║")
    print(f"  ║  http://localhost:{PORT}                ║")
    print(f"  ║  http://127.0.0.1:{PORT}                ║")
    print(f"  ║                                          ║")
    print(f"  ║  局域网访问 (手机/平板/其他电脑):        ║")
    print(f"  ║  http://{local_ip}:{PORT}        ║")
    # Pad the IP line
    ip_url = f"http://{local_ip}:{PORT}"
    padding = 42 - len(ip_url)
    if padding > 0:
        print(f"  ║  {ip_url}{' ' * padding}║")
    print(f"  ║                                          ║")
    print(f"  ║  按 Ctrl+C 停止服务                      ║")
    print("  ╚══════════════════════════════════════════╝")
    print()

    with http.server.HTTPServer(('0.0.0.0', PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n  服务已停止\n")
            sys.exit(0)
