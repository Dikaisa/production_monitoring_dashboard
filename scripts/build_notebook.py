"""Susun notebook analisis (dijalankan lewat nbconvert supaya output ikut tersimpan)."""
import nbformat as nbf
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
nb = nbf.v4.new_notebook()
md, code = nbf.v4.new_markdown_cell, nbf.v4.new_code_cell

nb.cells = [
md("""# Analisis Data Produksi Harian Lini Perakitan

**Tujuan:** menemukan part yang paling bermasalah, pola pencapaian target per bulan dan per tipe hari, serta kondisi stok, lalu memberi rekomendasi.

> ⚠️ Data di notebook ini **simulasi** (dibuat dengan `scripts/generate_data.py`), meniru struktur rekap produksi yang pernah saya buat saat bekerja sebagai operator produksi. Nama perusahaan, line, dan part sudah dianonimkan.

**Alur:** Load & cek data → Feature engineering → Analisis dengan SQL → Visualisasi → Insight & rekomendasi"""),

code("""import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

sns.set_theme(style="whitegrid")
NAVY = "#1F3864"
Path("../images").mkdir(exist_ok=True)

df = pd.read_csv("../data/produksi_harian.csv", parse_dates=["tanggal"])
df.head()"""),

md("## 1. Cek kualitas data"),
code("""print("Ukuran data :", df.shape)
print("Periode     :", df.tanggal.min().date(), "s.d.", df.tanggal.max().date())
print("Hari kerja  :", df.tanggal.nunique())
print("Duplikat (tanggal+part):", df.duplicated(["tanggal", "part"]).sum())
print()
print(df.isna().sum())"""),
code("""# validasi logika: NG tidak boleh lebih besar dari produksi, angka tidak boleh negatif
num = ["planning", "stok_awal", "stok_supplier", "produksi", "ng"]
assert (df[num] >= 0).all().all()
assert (df.ng <= df.produksi).all()
df[num].describe().round(1)"""),

md("""## 2. Feature engineering

Rumus sama dengan template Excel:

| Kolom | Rumus |
|---|---|
| `ok` | produksi − ng |
| `total_tersedia` | stok_awal + stok_supplier + ok |
| `kekurangan` | max(0, planning − total_tersedia) |
| `stok_akhir` | max(0, total_tersedia − planning) |
| `pencapaian` | ok / planning |"""),
code("""df["ok"] = df.produksi - df.ng
df["total_tersedia"] = df.stok_awal + df.stok_supplier + df.ok
df["kekurangan"] = (df.planning - df.total_tersedia).clip(lower=0)
df["stok_akhir"] = (df.total_tersedia - df.planning).clip(lower=0)
df["pencapaian"] = df.ok / df.planning
df["bulan"] = df.tanggal.dt.to_period("M").astype(str)
df["kategori_hari"] = df.tipe_hari.where(df.tipe_hari == "Sabtu",
                                         "Hari Kerja " + df.lembur.map({"Ya": "(lembur)", "Tidak": "(normal)"}))
df.head()"""),

md("## 3. Analisis dengan SQL\n\nData dimasukkan ke SQLite in-memory supaya bisa di-query seperti di database."),
code("""con = sqlite3.connect(":memory:")
df.assign(tanggal=df.tanggal.dt.strftime("%Y-%m-%d")).to_sql("produksi", con, index=False)

def q(sql):
    return pd.read_sql(sql, con)"""),
md("**3.1 Ringkasan per bulan**"),
code("""q('''
SELECT bulan,
       COUNT(DISTINCT tanggal)                       AS hari_kerja,
       SUM(planning)                                 AS planning,
       SUM(ok)                                       AS ok,
       ROUND(100.0 * SUM(ok) / SUM(planning), 1)     AS pencapaian_pct,
       ROUND(100.0 * SUM(ng) / SUM(produksi), 2)     AS ng_rate_pct,
       SUM(kekurangan)                               AS kekurangan
FROM produksi
GROUP BY bulan
ORDER BY bulan
''')"""),
md("**3.2 Ranking part berdasarkan NG rate dan kekurangan**"),
code("""q('''
SELECT part,
       ROUND(100.0 * SUM(ng) / SUM(produksi), 2)          AS ng_rate_pct,
       SUM(kekurangan)                                    AS total_kekurangan,
       SUM(CASE WHEN kekurangan > 0 THEN 1 ELSE 0 END)    AS hari_kurang,
       ROUND(100.0 * SUM(CASE WHEN kekurangan > 0 THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_hari_kurang
FROM produksi
GROUP BY part
ORDER BY ng_rate_pct DESC
''')"""),
md("**3.3 Pencapaian per kategori hari**"),
code("""q('''
SELECT kategori_hari,
       COUNT(DISTINCT tanggal)                    AS jumlah_hari,
       ROUND(100.0 * SUM(ok) / SUM(planning), 1)  AS pencapaian_pct
FROM produksi
GROUP BY kategori_hari
ORDER BY pencapaian_pct DESC
''')"""),
md("**3.4 Stok akhir terakhir per part** (cek overstock)"),
code("""q('''
SELECT part, stok_akhir, tanggal
FROM produksi
WHERE tanggal = (SELECT MAX(tanggal) FROM produksi)
ORDER BY stok_akhir DESC
''')"""),

md("## 4. Visualisasi"),
code("""bulanan = df.groupby("bulan").agg(planning=("planning", "sum"), ok=("ok", "sum"))
bulanan["pencapaian"] = bulanan.ok / bulanan.planning

fig, ax = plt.subplots(figsize=(9, 4))
ax.plot(bulanan.index, bulanan.pencapaian * 100, marker="o", color=NAVY, lw=2.5)
for x, y in zip(bulanan.index, bulanan.pencapaian * 100):
    ax.annotate(f"{y:.1f}%", (x, y), textcoords="offset points", xytext=(0, 8), ha="center")
ax.set_title("Pencapaian Target per Bulan", fontweight="bold")
ax.set_ylabel("Pencapaian (%)")
ax.set_ylim(90, 100)
plt.tight_layout()
plt.savefig("../images/pencapaian_bulanan.png", dpi=150)
plt.show()"""),
code("""heat = df.pivot_table(index="part", columns="bulan", values=["ok", "planning"], aggfunc="sum")
heat = (heat["ok"] / heat["planning"] * 100).round(1)

fig, ax = plt.subplots(figsize=(9, 3.8))
sns.heatmap(heat, annot=True, fmt=".1f", cmap="RdYlGn", vmin=85, vmax=100, cbar_kws={"label": "%"}, ax=ax)
ax.set_title("Pencapaian (%) per Part per Bulan", fontweight="bold")
ax.set_xlabel(""); ax.set_ylabel("")
plt.tight_layout()
plt.savefig("../images/heatmap_part_bulan.png", dpi=150)
plt.show()"""),
code("""per_part = df.groupby("part").agg(ng=("ng", "sum"), produksi=("produksi", "sum"), kekurangan=("kekurangan", "sum"))
per_part["ng_rate"] = per_part.ng / per_part.produksi * 100
warna = ["#C00000" if p == per_part.ng_rate.idxmax() else "#A6A6A6" for p in per_part.index]

fig, axes = plt.subplots(1, 2, figsize=(11, 4))
axes[0].bar(per_part.index, per_part.ng_rate, color=warna)
axes[0].set_title("NG Rate per Part (%)", fontweight="bold")
for i, v in enumerate(per_part.ng_rate):
    axes[0].text(i, v + 0.05, f"{v:.2f}%", ha="center")
axes[1].bar(per_part.index, per_part.kekurangan, color=warna)
axes[1].set_title("Total Kekurangan per Part (pcs)", fontweight="bold")
for i, v in enumerate(per_part.kekurangan):
    axes[1].text(i, v + 30, f"{v:,}", ha="center")
plt.tight_layout()
plt.savefig("../images/ng_kekurangan_part.png", dpi=150)
plt.show()"""),
code("""kat = df.groupby("kategori_hari").agg(ok=("ok", "sum"), planning=("planning", "sum"))
kat["pencapaian"] = kat.ok / kat.planning * 100
kat = kat.sort_values("pencapaian")

fig, ax = plt.subplots(figsize=(8, 3.2))
ax.barh(kat.index, kat.pencapaian, color=NAVY)
ax.axvline(100, color="grey", ls="--", lw=1)
for i, v in enumerate(kat.pencapaian):
    ax.text(v + 0.3, i, f"{v:.1f}%", va="center")
ax.set_xlim(80, 108)
ax.set_title("Pencapaian per Kategori Hari", fontweight="bold")
plt.tight_layout()
plt.savefig("../images/pencapaian_tipe_hari.png", dpi=150)
plt.show()"""),
code("""fig, ax = plt.subplots(figsize=(10, 4))
for part, g in df.groupby("part"):
    ax.plot(g.tanggal, g.stok_akhir, label=part, lw=1.6)
ax.set_title("Tren Stok Akhir per Part", fontweight="bold")
ax.set_ylabel("pcs")
ax.legend(ncol=5, loc="upper left")
plt.tight_layout()
plt.savefig("../images/tren_stok.png", dpi=150)
plt.show()"""),

md("""## 5. Insight & Rekomendasi

| # | Temuan | Rekomendasi |
|---|---|---|
| 1 | **Part D** punya NG rate ~3,1%, sekitar 3x part lain, dan kekurangan di ~68% hari kerja. | Prioritaskan perbaikan kualitas di proses Part D (cek mesin/jig, analisis jenis NG). |
| 2 | **Agustus** jadi bulan terendah karena pencapaian Part D anjlok ke ~82% (skenario masalah mesin). September juga turun, tapi merata di semua part. | Pasang alert otomatis kalau pencapaian part di bawah 90% supaya masalah ketahuan lebih cepat, bukan di akhir bulan. |
| 3 | **Lembur** menaikkan pencapaian ke >100%, sedangkan **Sabtu masuk** paling rendah (~88%). | Pakai lembur untuk part yang sedang kurang. Evaluasi efektivitas Sabtu masuk (jumlah orang, setup mesin). |
| 4 | **Part C & E** tidak pernah kurang, tapi stok terus naik sampai >1.000 pcs karena tambahan dari supplier. | Kurangi atau jadwalkan ulang order supplier untuk Part C & E supaya stok tidak menumpuk. |

**Catatan:** karena datanya simulasi, angka di atas untuk mendemonstrasikan alur analisis, bukan kondisi pabrik yang sebenarnya."""),
]
nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
out = ROOT / "notebooks" / "analisis_produksi.ipynb"
out.parent.mkdir(exist_ok=True)
nbf.write(nb, out)
print("saved", out)
