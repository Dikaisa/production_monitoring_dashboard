"""
Generate data produksi SIMULASI (dummy) untuk project portofolio.
Data ini BUKAN data asli perusahaan. Struktur kolom meniru format rekap
produksi harian di lini perakitan, nama part & line sudah dianonimkan.

Periode : 1 Mei 2026 - 31 Oktober 2026
Aturan  : Minggu & 17 Agustus libur, Sabtu masuk opsional (~40%),
          lembur acak di hari kerja biasa.
"""
import numpy as np
import pandas as pd
from pathlib import Path

rng = np.random.default_rng(42)

PARTS = {
    # part : (line tujuan, planning dasar per hari, NG rate dasar)
    "Part A": ("Line 3", 400, 0.010),
    "Part B": ("Line 2", 300, 0.012),
    "Part C": ("Line 2", 300, 0.011),
    "Part D": ("Line 2", 250, 0.028),   # part yang paling sering NG
    "Part E": ("Line 2", 200, 0.012),
}
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
LIBUR_NASIONAL = {pd.Timestamp("2026-08-17")}

rows = []
stok = {p: int(base * 0.3) for p, (_, base, _) in PARTS.items()}

for tgl in pd.date_range("2026-05-01", "2026-10-31", freq="D"):
    dow = tgl.dayofweek
    if dow == 6 or tgl in LIBUR_NASIONAL:
        continue
    if dow == 5:
        if rng.random() > 0.4:
            continue
        tipe = "Sabtu"
        lembur = "Tidak"
    else:
        tipe = "Hari Kerja"
        lembur = "Ya" if rng.random() < 0.25 else "Tidak"

    for part, (line, base, ng_base) in PARTS.items():
        plan = base * (0.5 if tipe == "Sabtu" else 1.0)
        plan = int(round(plan * rng.normal(1.0, 0.08) / 10) * 10)

        perf = rng.normal(0.97, 0.05)
        if lembur == "Ya":
            perf += 0.08
        if tipe == "Sabtu":
            perf -= 0.06
        # skenario: Part D sempat ada masalah mesin di Agustus
        if part == "Part D" and tgl.month == 8:
            perf -= 0.12
        produksi = max(0, int(plan * perf))

        ng_rate = ng_base * (1.6 if (part == "Part D" and tgl.month == 8) else 1.0)
        ng = int(rng.binomial(produksi, ng_rate))
        ok = produksi - ng

        supplier = 0
        if part in ("Part C", "Part E") and rng.random() < 0.3:
            supplier = int(rng.integers(3, 8) * 10)

        stok_awal = stok[part]
        tersedia = stok_awal + supplier + ok
        stok[part] = max(0, tersedia - plan)

        rows.append({
            "tanggal": tgl.date(),
            "hari": HARI[dow],
            "tipe_hari": tipe,
            "lembur": lembur,
            "part": part,
            "line_tujuan": line,
            "planning": plan,
            "stok_awal": stok_awal,
            "stok_supplier": supplier,
            "produksi": produksi,
            "ng": ng,
        })

df = pd.DataFrame(rows)
out = Path(__file__).resolve().parents[1] / "data" / "produksi_harian.csv"
out.parent.mkdir(exist_ok=True)
df.to_csv(out, index=False)
print(f"{len(df)} baris -> {out}")
