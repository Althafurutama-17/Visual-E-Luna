"""
SMART LOCAL HTTP SERVER FOR E-LUNA DASHBOARD (BPBD JAWA BARAT)
- Serves static files on port 3456
- Disables HTTP cache (Cache-Control: no-cache) so browser always sees latest files
- Auto-detects changes to 'database_logistik_bpbd_jabar.xlsx' and syncs to 'data.js'
- Provides /api/sync and /api/status endpoints
"""

import http.server
import socketserver
import os
import json
import time
import datetime
from sync_excel_to_web import sync_data, EXCEL_FILE, TARGET_JS, TARGET_JSON

PORT = 3456
last_mtime = 0

class SmartElunaHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Mencegah browser caching agar perubahan file Excel atau JS langsung tampil
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        self.send_header('Access-Control-Allow-Origin', '*')
        super().end_headers()

    def do_GET(self):
        global last_mtime
        
        # Cek apakah file Excel telah diperbarui di disk
        if os.path.exists(EXCEL_FILE):
            try:
                mtime = os.path.getmtime(EXCEL_FILE)
                if mtime > last_mtime:
                    last_mtime = mtime
                    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Terdeteksi perubahan pada {EXCEL_FILE}, melakukan auto-sync...")
                    sync_data()
            except Exception as e:
                print("Error checking Excel file mtime:", e)

        # Endpoint API Status
        if self.path.startswith('/api/status'):
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            
            file_exists = os.path.exists(EXCEL_FILE)
            status_data = {
                "server": "E-LUNA Localhost Smart Server",
                "excel_exists": file_exists,
                "excel_file": EXCEL_FILE,
                "excel_size": os.path.getsize(EXCEL_FILE) if file_exists else 0,
                "last_modified": datetime.datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime('%Y-%m-%d %H:%M:%S') if file_exists else None,
                "timestamp": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            self.wfile.write(json.dumps(status_data).encode('utf-8'))
            return

        # Endpoint API Sync Manual
        if self.path.startswith('/api/sync'):
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Memproses permintaan sinkronisasi manual dari browser...")
            sync_data()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            resp = {"status": "success", "message": "Excel database synced to data.js"}
            self.wfile.write(json.dumps(resp).encode('utf-8'))
            return

        # Default handler untuk file statis
        super().do_GET()

def run_server():
    global last_mtime
    if os.path.exists(EXCEL_FILE):
        last_mtime = os.path.getmtime(EXCEL_FILE)

    # Allow address reuse
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), SmartElunaHandler) as httpd:
        print(f"=== E-LUNA SMART LOCAL SERVER AKTIF ===")
        print(f"URL: http://localhost:{PORT}")
        print(f"Pemantauan Otomatis: {EXCEL_FILE}")
        httpd.serve_forever()

if __name__ == '__main__':
    run_server()
