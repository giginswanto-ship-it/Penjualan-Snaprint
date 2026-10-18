import sqlite3
import sys

def query(sql):
    conn = sqlite3.connect("snaprint_sales.db")
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    cols = [desc[0] for desc in cur.description]
    
    # Calculate column widths
    widths = [len(c) for c in cols]
    for r in rows:
        for i, val in enumerate(r):
            widths[i] = max(widths[i], len(str(val)))
            
    header = " | ".join(c.ljust(widths[i]) for i, c in enumerate(cols))
    sep = "-+-".join("-" * widths[i] for i in range(len(cols)))
    print(header)
    print(sep)
    for r in rows:
        print(" | ".join(str(val).ljust(widths[i]) for i, val in enumerate(r)))
    print(f"\n({len(rows)} baris data)")
    conn.close()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        query(" ".join(sys.argv[1:]))
    else:
        print("--- REKAP KINERJA TAHUNAN ---")
        query("SELECT * FROM v_annual_performance;")
        print("\n--- KOMPARASI 3 TAHUN (JAN - AGU) ---")
        query("SELECT * FROM v_comparison_3_years_8m;")
