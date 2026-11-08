# UUI × Meta — Presentasi Strategi 1 Juta Followers

Paket presentasi untuk pimpinan Universitas Ubudiyah Indonesia: rencana 18 bulan
Meta Ads, pertumbuhan followers, dan mesin engagement di Facebook, Instagram,
Threads, dan WhatsApp Channel.

Tersedia dalam **tiga format**:

| Format | File | Untuk apa |
|---|---|---|
| Web (HTML) | `index.html` | Presentasi interaktif, kalkulator, animasi |
| PowerPoint | `UUI-Meta-Ads-2026.pptx` | Rapat, editing, presentasi dengan speaker notes |
| PDF | `UUI-Meta-Ads-2026.pdf` | Dibagikan atau dicetak, 37 halaman |
| Naskah | `NARASI-PRESENTASI.md` | Narasi per slide, timing, Q&A |

**Buka `index.html` di browser** — tanpa internet, build, atau server. Semua aset lokal.

---

## Publikasi GitHub Pages

Presentasi online: **https://mahendartea.github.io/metaads/**

- Workflow: `.github/workflows/pages.yml`
- Setiap **push ke `main`** otomatis membangun dan men-deploy situs.
- Deploy manual: tab **Actions** → *Deploy presentation to GitHub Pages* → **Run workflow**.
- Sumber Pages: **Settings → Pages → Build and deployment → GitHub Actions**.
- Unduhan langsung:
  - PDF: `/UUI-Meta-Ads-2026.pdf`
  - PowerPoint: `/UUI-Meta-Ads-2026.pptx`

Workflow hanya mempublikasikan file yang dibutuhkan situs (`index.html`,
`styles.css`, `app.js`, `assets/`) plus PDF dan PPTX. Folder `tools/`, naskah
internal, dan file sumber tidak ikut dipublikasikan.

---

## Isi presentasi (37 slide)

| Bab | Slide | Isi |
|---|---|---|
| Pembuka | 1–3 | Cover, agenda, ringkasan eksekutif |
| 01 Fondasi | 4–6 | Ekosistem Meta, arsitektur aset, tata kelola akun |
| 02 Meta Ads | 7–15 | 10 langkah pendaftaran, verifikasi, pixel/CAPI, struktur kampanye, checklist go-live |
| 03 Konten | 16–21 | Pilar konten, ritme posting, anatomi Reels, playbook engagement, kalender |
| 04 1 Juta | 22–28 | Matematika target, kalkulator, alokasi platform, kurva, anggaran, funnel, loop viral |
| 05 Eksekusi | 29–35 | Rencana 90 hari, tim, KPI, ritme kerja, risiko, tata kelola |
| Penutup | 36–37 | 5 keputusan, lampiran standar eksekusi |

---

## Navigasi versi web

| Aksi | Kontrol |
|---|---|
| Slide berikutnya | `→` `Space` `PageDown` atau klik sisi kanan slide |
| Slide sebelumnya | `←` `PageUp` atau klik sisi kiri slide |
| Slide pertama / terakhir | `Home` / `End` |
| Ringkasan semua slide | `O` atau tombol grid |
| Catatan pembicara | `N` atau tombol catatan |
| Layar penuh | `F` |
| Cetak / simpan PDF | `P` atau `Ctrl/Cmd + P` |
| Tutup panel | `Esc` |

Geser (swipe) juga berfungsi di layar sentuh. Setiap slide punya URL sendiri,
mis. `index.html#s23` untuk membuka slide 23.

---

## Fitur versi web

- **Kalkulator target interaktif** di slide 23 — ubah followers awal, target, dan durasi.
- **Catatan pembicara** untuk 37 slide, tersembunyi sampai ditekan `N`.
- **Mode ringkasan** (`O`) untuk melompat ke slide mana pun.
- **Grafik pertumbuhan** beranimasi saat slide dibuka.
- **Skala otomatis** mengikuti ukuran jendela, rasio tetap 16:9.
- **Siap cetak** — 37 halaman 1280×720, tanpa elemen UI.

---

## Versi PowerPoint

`UUI-Meta-Ads-2026.pptx` — 16:9, 37 slide, teks dan bentuk dapat diedit.

- **Speaker notes sudah terisi** untuk setiap slide, diambil dari
  `NARASI-PRESENTASI.md`. Buka *View → Notes Page* atau *Presenter View*.
- Semua tabel, kartu, dan diagram dibuat dari shape PowerPoint, bukan gambar,
  sehingga bisa diubah langsung.
- Grafik pertumbuhan dibuat dari shape vektor; bukan gambar.
- Perbedaan dari versi web: kalkulator bersifat ilustrasi statis (tidak
  interaktif), dan animasi transisi tidak disertakan.
- Font: **Arial**, agar konsisten di macOS dan Windows. Ukuran teks sudah
  disesuaikan; hindari memperbesar teks melebihi 10% agar tidak meluber.

### Mengekspor PPTX ke PDF

Buka di PowerPoint → *File → Export → PDF*, atau gunakan PowerPoint untuk
mencetak ke PDF. Hasilnya 37 halaman 16:9.

---

## Naskah narasi

`NARASI-PRESENTASI.md` berisi:

- peta waktu 37 slide (±51 menit);
- narasi pembicara per slide (bukan untuk dibaca kata per kata);
- poin kunci dan transisi antar slide;
- jalur ringkas 30 menit dan 15 menit;
- delapan pertanyaan yang paling mungkin muncul beserta jawabannya.

Naskah ini juga tertanam sebagai **speaker notes** di file PPTX.

---

## Struktur file

```
metaads/
├── index.html                    # presentasi web (37 slide)
├── styles.css                    # desain web, layout, aturan cetak
├── app.js                        # navigasi, catatan, kalkulator, grafik
├── assets/
│   ├── uui-logo.png              # logo UUI (transparan)
│   ├── meta-logo-white.png       # logo Meta putih (untuk PPTX)
│   ├── meta-logo-white.svg       # logo Meta putih (untuk web)
│   ├── 2021_Facebook_icon.png    # ikon platform untuk PPTX
│   ├── Instagram_logo_2016.png
│   └── WhatsApp.png
├── UUI-Meta-Ads-2026.pptx        # versi PowerPoint
├── UUI-Meta-Ads-2026.pdf         # ekspor PDF versi web
├── NARASI-PRESENTASI.md          # naskah pembicara 37 slide
├── tools/
│   ├── build_pptx.py             # skrip pembuat PPTX
│   ├── bg.html                   # sumber background
│   ├── bg-content.png            # background slide isi
│   ├── bg-cover.png              # background cover/penutup
│   └── bg-divider.png            # background pembatas bab
└── README.md
```

---

## Mengubah isi

### Versi web

- **Teks slide** ada di `index.html`, satu `<section class="slide">` per halaman.
- **Warna** ada di blok `:root` pada `styles.css`
  (`--red: #ce2020` sesuai logo UUI, `--blue: #0a7cff` untuk Meta).
- **Angka target** di slide 23 dihitung otomatis; angka statis (1.000.000,
  987.600, 1.805/hari) perlu disesuaikan manual bila baseline berubah.
- **Catatan pembicara** ada di array `SPEAKER` pada `app.js`.
- **Logo** bisa diganti dengan menimpa file di `assets/`.

### Versi PowerPoint

Edit langsung di PowerPoint, atau bangun ulang dari skrip:

```bash
python3 -m venv .venv
.venv/bin/pip install python-pptx
.venv/bin/python tools/build_pptx.py
```

Skrip memakai koordinat 1280×720 yang dipetakan ke slide 13,33 × 7,5 inci.
Naskah dari `NARASI-PRESENTASI.md` otomatis dimasukkan ke speaker notes.

> Catatan: `tools/build_pptx.py` membutuhkan Python 3.9+ dan `python-pptx`.
> File PPTX yang sudah jadi tidak memerlukan dependensi apa pun.

---

## Asumsi penting

Semua angka audiens, biaya, dan konversi adalah **asumsi kerja** berbasis
benchmark umum, bukan jaminan hasil:

- Baseline 12.400 followers gabungan (perlu audit akun riil).
- Biaya per follower berbayar Rp 2.000 (rentang Rp 1.200–3.500).
- 30% followers dari media berbayar, 70% organik.
- Anggaran rekomendasi Rp 35 jt/bulan media + Rp 9 jt produksi + Rp 2 jt tools.

Ganti dengan data aktual setelah 14 hari audit. Target 1 juta dalam 18 bulan
adalah target agresif; deck juga menyiapkan skenario 24 bulan sebagai cadangan.

---

Disusun oleh **Direktorat Cyber Development Center & Media UBOnTV UUI** — November 2026.
