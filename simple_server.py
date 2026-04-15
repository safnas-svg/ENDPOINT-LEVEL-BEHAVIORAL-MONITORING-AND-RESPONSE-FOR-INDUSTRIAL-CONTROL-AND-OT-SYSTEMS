#!/usr/bin/env python3
"""
SWAT Dashboard - Simple HTTP Server
Serves dashboard.html with proper error handling and CORS
"""
import http.server
import socketserver
import os
import sys

class SimpleHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        """Handle GET requests"""
        # Default path behavior
        if self.path == '/':
            self.path = '/dashboard_v2.html'
        
        # Check if file exists
        filepath = os.path.join(os.getcwd(), self.path.lstrip('/'))
        
        if os.path.isfile(filepath):
            # Serve the file
            try:
                with open(filepath, 'rb') as f:
                    self.send_response(200)
                    
                    # Set content type
                    if filepath.endswith('.html'):
                        self.send_header('Content-Type', 'text/html; charset=utf-8')
                    elif filepath.endswith('.js'):
                        self.send_header('Content-Type', 'application/javascript')
                    elif filepath.endswith('.css'):
                        self.send_header('Content-Type', 'text/css')
                    else:
                        self.send_header('Content-Type', 'application/octet-stream')
                    
                    # CORS headers
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.send_header('Cache-Control', 'no-cache')
                    
                    data = f.read()
                    self.send_header('Content-Length', len(data))
                    self.end_headers()
                    self.wfile.write(data)
                    print(f"[SERVE] {self.path} -> {len(data)} bytes")
            except Exception as e:
                self.send_error(500, str(e))
        else:
            # File not found - try dashboard_v2.html
            if self.path != '/dashboard_v2.html':
                try:
                    with open('dashboard_v2.html', 'rb') as f:
                        self.send_response(200)
                        self.send_header('Content-Type', 'text/html; charset=utf-8')
                        self.send_header('Access-Control-Allow-Origin', '*')
                        data = f.read()
                        self.send_header('Content-Length', len(data))
                        self.end_headers()
                        self.wfile.write(data)
                        print(f"[FALLBACK] Served dashboard_v2.html for {self.path}")
                except:
                    self.send_error(404, "dashboard_v2.html not found")
            else:
                self.send_error(404, f"File not found: {self.path}")
    
    def log_message(self, format, *args):
        print(f"[HTTP] {format % args}")

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)) or '.')
    
    PORT = 3000
    
    print("\n" + "=" * 70)
    print("THREATDETECT - DASHBOARD HTTP SERVER")
    print("=" * 70)
    print(f"\nServer root: {os.getcwd()}")
    print(f"Serving files on port {PORT}")
    print(f"Dashboard URL: http://localhost:{PORT}")
    print(f"Backend API: http://localhost:5000")
    
    try:
        with socketserver.TCPServer(("", PORT), SimpleHandler) as httpd:
            print(f"\n[SUCCESS] Server started on port {PORT}")
            print("Press Ctrl+C to stop\n")
            httpd.serve_forever()
    except OSError as e:
        if e.errno == 48 or e.errno == 98:  # Address already in use
            print(f"\n[ERROR] Port {PORT} is already in use!")
            print("Kill existing process: taskkill /F /FI \"COMMAND eq python.exe\"")
            sys.exit(1)
        else:
            print(f"\n[ERROR] {e}")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n[STOP] Server stopped")
        sys.exit(0)

if __name__ == '__main__':
    main()
