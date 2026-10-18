import http.server
import socketserver
import webbrowser
import os
import sys
import json
import subprocess

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class SnaprintHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # Enable CORS for local cross-origin file/port requests
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200, "ok")
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/save':
            try:
                content_len = int(self.headers.get('Content-Length', 0))
                post_body = self.rfile.read(content_len)
                payload = json.loads(post_body.decode('utf-8'))

                year = int(payload.get('year', 2026))
                pl_data = payload.get('profit_loss')
                monthly_data = payload.get('monthly')

                # 1. Save specific year's laba_rugi JSON file
                target_json = ""
                if year == 2024:
                    target_json = os.path.join(DIRECTORY, "laba_rugi_2024.json")
                elif year == 2025:
                    target_json = os.path.join(DIRECTORY, "2025", "laba_rugi_2025.json")
                elif year == 2026:
                    target_json = os.path.join(DIRECTORY, "2026", "laba_rugi_2026.json")

                if target_json and pl_data:
                    with open(target_json, "w", encoding="utf-8") as f:
                        json.dump(pl_data, f, ensure_ascii=False, indent=2)

                # 2. Update sales_data_master.json
                master_path = os.path.join(DIRECTORY, "sales_data_master.json")
                if os.path.exists(master_path):
                    with open(master_path, "r", encoding="utf-8") as f:
                        master = json.load(f)

                    yr_key = f"data_{year}"
                    if yr_key in master and pl_data:
                        master[yr_key]["profit_loss"] = pl_data
                    if yr_key in master and monthly_data:
                        master[yr_key]["monthly"] = monthly_data

                    # Update profit_loss_comparison summary
                    idx = 0 if year == 2024 else (1 if year == 2025 else 2)
                    if "profit_loss_comparison" in master and len(master["profit_loss_comparison"]["summary"]) > idx:
                        s = master["profit_loss_comparison"]["summary"][idx]
                        if pl_data:
                            s["revenue"] = pl_data.get("revenue", {}).get("net_revenue", s["revenue"])
                            s["cogs"] = pl_data.get("cogs", {}).get("total_cogs", s["cogs"])
                            s["cogs_pct"] = pl_data.get("cogs", {}).get("cogs_pct", s["cogs_pct"])
                            s["gross_profit"] = pl_data.get("gross_profit", {}).get("amount", s["gross_profit"])
                            s["gross_margin_pct"] = pl_data.get("gross_profit", {}).get("margin_pct", s["gross_margin_pct"])
                            s["opex"] = pl_data.get("operating_expenses", {}).get("total", s["opex"])
                            s["opex_pct"] = pl_data.get("operating_expenses", {}).get("pct_of_revenue", s["opex_pct"])
                            s["net_profit"] = pl_data.get("net_profit", {}).get("amount", s["net_profit"])
                            s["net_margin_pct"] = pl_data.get("net_profit", {}).get("margin_pct", s["net_margin_pct"])
                            s["avg_monthly_net_profit"] = pl_data.get("net_profit", {}).get("avg_monthly", s["avg_monthly_net_profit"])

                    with open(master_path, "w", encoding="utf-8") as f:
                        json.dump(master, f, ensure_ascii=False, indent=2)

                # 3. Synchronize SQLite Database & SQL script
                sql_builder = os.path.join(DIRECTORY, "build_sql_database.py")
                if os.path.exists(sql_builder):
                    subprocess.run([sys.executable, sql_builder], cwd=DIRECTORY, check=False)

                # Send response
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                resp = {
                    "success": True,
                    "message": f"Data tahun {year} berhasil disimpan ke file folder dan database SQL!",
                    "year": year
                }
                self.wfile.write(json.dumps(resp).encode('utf-8'))
                return

            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                err_resp = {"success": False, "error": str(e)}
                self.wfile.write(json.dumps(err_resp).encode('utf-8'))
                return

        super().do_POST()

def run_server():
    os.chdir(DIRECTORY)
    port = PORT
    while True:
        try:
            with socketserver.TCPServer(("", port), SnaprintHandler) as httpd:
                url = f"http://localhost:{port}/index.html"
                print("=" * 64)
                print("  SNAPRINT DIGITAL - SISTEM DASHBOARD PENJUALAN & LABA RUGI")
                print("  Terintegrasi: 2024, 2025, 2026 (s/d Agustus)")
                print("=" * 64)
                print(f"  Server Aktif di     : {url}")
                print(f"  API Simpan Perubahan: http://localhost:{port}/api/save")
                print("  Tekan Ctrl+C untuk menghentikan server.")
                print("=" * 64)
                webbrowser.open(url)
                httpd.serve_forever()
        except OSError:
            port += 1

if __name__ == "__main__":
    try:
        run_server()
    except KeyboardInterrupt:
        print("\nServer dihentikan.")
        sys.exit(0)
