/* ============================================================
   UUI × Meta — presentation engine
   ============================================================ */

(function () {
  "use strict";

  const stage = document.getElementById("stage");
  const slides = Array.from(document.querySelectorAll(".slide"));
  const total = slides.length;
  const progress = document.getElementById("progress");
  const counter = document.getElementById("counter");
  const ui = document.getElementById("ui");
  const notes = document.getElementById("notes");
  const notesBody = document.getElementById("notesBody");
  const notesSlide = document.getElementById("notesSlide");
  const overview = document.getElementById("overview");
  const ovGrid = document.getElementById("ovGrid");
  const toast = document.getElementById("toast");

  let index = 0;
  let overviewOpen = false;
  let notesOpen = false;
  let idleTimer = null;

  /* ---------------- speaker notes ---------------- */
  const SPEAKER = [
    "Buka dengan framing: ini rencana kerja 18 bulan, bukan sekadar presentasi ide. Tiga angka pembuka — 1 juta audiens, 4 platform Meta, 18 bulan.",
    "Jelaskan alur lima bab. Tegaskan bahwa bab 2 bersifat teknis dan bisa dieksekusi minggu ini, sementara bab 4 adalah inti ambisi pimpinan.",
    "Tiga angka kunci. Tekankan 70/30: iklan mempercepat, bukan menggantikan. Akui bahwa target 18 bulan agresif dan sampaikan skenario 24 bulan sebagai cadangan.",
    "Transisi ke fondasi. Analogi: sebelum mengecat dinding, pastikan pondasi dan kunci rumah ada di tangan yang benar.",
    "Untuk setiap platform, sebutkan satu pekerjaan utamanya. Hindari tergoda menjadikan semua platform sebagai tempat pengumuman yang sama.",
    "Tunjukkan diagram. Tiga aturan: aset milik institusi, akun iklan per tujuan, akses lewat undangan. Ini mencegah kehilangan aset saat staf berganti.",
    "Transisi ke bagian teknis. Beri tahu audience bahwa 10 langkah ini bisa diselesaikan dalam 7–14 hari.",
    "Bacakan peta 10 langkah secara cepat, lalu masuk ke detail di dua slide berikutnya. Jangan membacakan semua teks.",
    "Tekankan verifikasi bisnis dan domain. Sebutkan bahwa ini penyebab paling umum kampanye tidak bisa jalan atau akun dibatasi.",
    "Tekankan pemisahan akun iklan dan keamanan akses. Ceritakan risiko akun pribadi staf yang menjadi pemilik aset.",
    "Tunjukkan checklist dokumen. Minta komitmen: dokumen dikumpulkan dalam 7 hari agar verifikasi tidak menjadi hambatan.",
    "Jelaskan peran dan prinsip akses minimum. Sebut rencana kontinjensi bila akun dibatasi: dua admin aktif.",
    "Jelaskan pixel di browser dan CAPI di server. Analogi: dua saksi untuk satu peristiwa, agar data tidak hilang.",
    "Bahas pemilihan tujuan. Tegaskan bahwa 'Boost Post' bukan strategi, hanya tombol pintas.",
    "Gunakan checklist ini sebagai gerbang go-live. Sebut aturan batas aman belanja dan penggantian kreatif.",
    "Transisi ke konten. Kutipan pembuka: algoritma tidak jatuh cinta pada kampus, manusia jatuh cinta pada manusia.",
    "Jelaskan lima pilar dan proporsi 70/30. Beri contoh nyata untuk masing-masing pilar agar tim punya gambaran produksi.",
    "Bahas ritme. Tekankan produksi batch dan bank konten 14 hari sebagai kunci keberlanjutan.",
    "Bedah Reels dari hook sampai CTA. Tunjukkan bahwa 3 detik pertama adalah investasi terbesar.",
    "Jelaskan engagement sebagai percakapan, bukan metrik. Golden hour 30 menit adalah kebiasaan yang paling berdampak.",
    "Tunjukkan kalender mingguan. Tekankan bahwa tema berbeda tiap hari mencegah kelelahan audiens dan tim.",
    "Transisi ke inti: matematika 1 juta. Minta perhatian penuh karena bagian ini menentukan anggaran.",
    "Tunjukkan hitungan dan demo kalkulator. Minta angka baseline riil dari tim agar proyeksi akurat.",
    "Jelaskan alokasi dan alasan di baliknya. Instagram menanggung beban akuisisi terbesar; WhatsApp menjaga retensi.",
    "Telusuri kurva per fase. Tekankan bahwa pertumbuhan tidak linear dan justru makin cepat di akhir bila fondasi benar.",
    "Tunjukkan tiga skenario anggaran. Rekomendasi Rp 35 juta media per bulan. Minta keputusan prinsip, detail bisa dibahas kemudian.",
    "Jelaskan funnel bulan ke-12. Tekankan bahwa 15 juta tayangan wajar bila konten dan iklan bekerja bersama.",
    "Bahas lima loop. Yang paling murah dan paling berkelanjutan adalah loop kampus melalui kreator mahasiswa.",
    "Transisi ke eksekusi. Kutipan: strategi tanpa ritme hanyalah daftar keinginan.",
    "Bahas 90 hari. Sebut gerbang keputusan hari ke-90: bila pertumbuhan di bawah 8.000/bulan, ubah pesan, bukan menambah uang.",
    "Jelaskan tim 8,5 FTE dan pembagian peran. Tekankan pemisahan produksi dan persetujuan.",
    "Bahas delapan KPI. Tegaskan bahwa follower bukan satu-satunya ukuran; kualitas audiens dan pendaftar lebih penting.",
    "Jelaskan ritme harian sampai kuartalan. Tunjukkan bahwa laporan mingguan harus muat satu layar.",
    "Bahas risiko utama: pembatasan akun, follower palsu, dan krisis reputasi. Tunjukkan mitigasinya.",
    "Bahas tata kelola dan kepatuhan. Sebut UU PDP secara eksplisit karena melibatkan data calon mahasiswa.",
    "Tutup dengan lima keputusan. Minta persetujuan anggaran, penunjukan pemilik akun, dan tenggat dokumen 7 hari.",
    "Lampiran untuk tim eksekusi: penamaan, UTM, dan checklist mingguan. Tidak perlu dibacakan, cukup ditunjukkan."
  ];

  /* ---------------- scaling ---------------- */
  function resize() {
    const notesW = notesOpen ? Math.min(340, window.innerWidth * 0.34) : 0;
    const availW = window.innerWidth - notesW;
    const scale = Math.min(availW / 1280, window.innerHeight / 720);
    stage.style.left = notesOpen ? "calc(50% - " + notesW / 2 + "px)" : "50%";
    stage.style.transform = "translate(-50%, -50%) scale(" + scale + ")";
  }
  window.addEventListener("resize", resize);
  resize();

  /* ---------------- navigation ---------------- */
  function go(n, push) {
    n = Math.max(0, Math.min(total - 1, n));
    if (n === index && slides[index].classList.contains("active")) return;

    const prev = slides[index];
    index = n;

    slides.forEach((s, i) => {
      s.classList.toggle("active", i === index);
      if (i !== index) s.classList.remove("leaving");
    });
    if (prev && prev !== slides[index]) {
      prev.classList.add("leaving");
      setTimeout(function () { prev.classList.remove("leaving"); }, 600);
    }

    const current = slides[index];
    current.classList.remove("leaving");

    progress.style.width = ((index + 1) / total) * 100 + "%";
    counter.textContent = String(index + 1).padStart(2, "0") + " / " + total;

    updateNotes();
    updateOverviewCurrent();

    if (push !== false) {
      history.replaceState(null, "", "#s" + (index + 1));
    }

    // restart chart animation
    if (current.querySelector("#growthChart")) animateChart();
  }

  function next() { go(index + 1); }
  function prev() { go(index - 1); }

  document.getElementById("btnNext").addEventListener("click", next);
  document.getElementById("btnPrev").addEventListener("click", prev);
  document.getElementById("btnPrint").addEventListener("click", function () { window.print(); });
  document.getElementById("btnFull").addEventListener("click", toggleFull);
  document.getElementById("btnNotes").addEventListener("click", toggleNotes);
  document.getElementById("btnOverview").addEventListener("click", toggleOverview);
  document.getElementById("notesClose").addEventListener("click", toggleNotes);

  document.addEventListener("keydown", function (e) {
    if (e.target && e.target.tagName === "INPUT") {
      if (e.key === "Escape") e.target.blur();
      return;
    }
    const k = e.key;
    if (k === "ArrowRight" || k === "PageDown" || k === " " || k === "Enter") {
      if (e.target.tagName === "INPUT") return;
      e.preventDefault(); next();
    } else if (k === "ArrowLeft" || k === "PageUp") {
      e.preventDefault(); prev();
    } else if (k === "Home") { e.preventDefault(); go(0); }
    else if (k === "End") { e.preventDefault(); go(total - 1); }
    else if (k === "o" || k === "O") { e.preventDefault(); toggleOverview(); }
    else if (k === "n" || k === "N") { e.preventDefault(); toggleNotes(); }
    else if (k === "f" || k === "F") { e.preventDefault(); toggleFull(); }
    else if (k === "p" || k === "P") { e.preventDefault(); window.print(); }
    else if (k === "Escape") {
      if (overviewOpen) toggleOverview(false);
      if (notesOpen) toggleNotes(false);
    }
  });

  /* touch */
  let tx = 0, ty = 0;
  stage.addEventListener("touchstart", function (e) {
    tx = e.changedTouches[0].clientX; ty = e.changedTouches[0].clientY;
  }, { passive: true });
  stage.addEventListener("touchend", function (e) {
    const dx = e.changedTouches[0].clientX - tx;
    const dy = e.changedTouches[0].clientY - ty;
    if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) {
      dx < 0 ? next() : prev();
    }
  }, { passive: true });

  /* click on right/left edges of stage */
  stage.addEventListener("click", function (e) {
    if (e.target.closest("input, button, a")) return;
    const rect = stage.getBoundingClientRect();
    const x = (e.clientX - rect.left) / rect.width;
    if (x > 0.82) next();
    else if (x < 0.1) prev();
  });

  /* ---------------- notes ---------------- */
  function updateNotes() {
    notesSlide.textContent = "Slide " + (index + 1) + " · " + (slides[index].dataset.title || "");
    notesBody.innerHTML = "<p>" + (SPEAKER[index] || "—") + "</p>";
  }

  function toggleNotes(force) {
    notesOpen = force === undefined ? !notesOpen : force;
    notes.classList.toggle("open", notesOpen);
    document.body.classList.toggle("notes-open", notesOpen);
    resize();
    if (notesOpen) updateNotes();
  }

  /* ---------------- overview ---------------- */
  function buildOverview() {
    ovGrid.innerHTML = "";
    slides.forEach(function (s, i) {
      const b = document.createElement("button");
      b.className = "ov-item";
      b.innerHTML = '<span class="n">' + String(i + 1).padStart(2, "0") + '</span>' +
        '<span class="t">' + (s.dataset.title || "Slide") + '</span>' +
        '<span class="s">' + (s.dataset.section || "") + '</span>';
      b.addEventListener("click", function () {
        go(i);
        toggleOverview(false);
      });
      ovGrid.appendChild(b);
    });
  }

  function updateOverviewCurrent() {
    const items = ovGrid.querySelectorAll(".ov-item");
    items.forEach(function (it, i) { it.classList.toggle("current", i === index); });
  }

  function toggleOverview(force) {
    overviewOpen = force === undefined ? !overviewOpen : force;
    overview.classList.toggle("open", overviewOpen);
    if (overviewOpen) updateOverviewCurrent();
  }

  /* ---------------- fullscreen ---------------- */
  function toggleFull() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(function () {});
    } else {
      document.exitFullscreen().catch(function () {});
    }
  }

  /* ---------------- toast / idle ---------------- */
  let toastTimer;
  function showToast(msg) {
    toast.textContent = msg;
    toast.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toast.classList.remove("show"); }, 2200);
  }

  function wake() {
    ui.classList.remove("idle");
    clearTimeout(idleTimer);
    idleTimer = setTimeout(function () { ui.classList.add("idle"); }, 3800);
  }
  document.addEventListener("mousemove", wake);
  document.addEventListener("touchstart", wake, { passive: true });
  wake();

  /* ---------------- calculator ---------------- */
  const nf = new Intl.NumberFormat("id-ID");
  const startEl = document.getElementById("calcStart");
  const targetEl = document.getElementById("calcTarget");
  const monthsEl = document.getElementById("calcMonths");
  const outDay = document.getElementById("outDay");
  const outMonth = document.getElementById("outMonth");
  const outPaid = document.getElementById("outPaid");
  const badge = document.getElementById("calcBadge");

  function numOf(el) {
    return Number(String(el.value).replace(/[^0-9]/g, "")) || 0;
  }

  function calc() {
    if (!startEl) return;
    const start = Math.max(0, numOf(startEl));
    const target = Math.max(1, numOf(targetEl));
    const months = Math.max(1, Number(monthsEl.value) || 1);
    const net = Math.max(0, target - start);
    const days = months * 30.4;
    const perDay = net / days;
    const perMonth = net / months;
    const paid = net * 0.3;

    outDay.textContent = nf.format(Math.round(perDay));
    outMonth.textContent = nf.format(Math.round(perMonth));
    outPaid.textContent = nf.format(Math.round(paid));

    let label;
    if (perDay > 1500) label = "Sangat agresif · butuh mesin konten harian + paid scale";
    else if (perDay > 800) label = "Agresif · konsistensi harian adalah syarat";
    else if (perDay > 300) label = "Realistis · fokus pada kualitas dan retensi";
    else label = "Santai · waktu untuk membangun fondasi";
    badge.textContent = label;
  }
  [startEl, targetEl, monthsEl].forEach(function (el) {
    if (el) el.addEventListener("input", calc);
  });
  [startEl, targetEl].forEach(function (el) {
    if (!el) return;
    el.addEventListener("focus", function () { el.select(); });
    el.addEventListener("blur", function () {
      el.value = nf.format(numOf(el));
      calc();
    });
  });
  calc();

  /* ---------------- growth chart ---------------- */
  const CH = { W: 1000, H: 300, L: 58, R: 26, T: 26, B: 40 };

  function smoothPath(pts) {
    let d = "M" + pts[0][0] + "," + pts[0][1];
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] || pts[i];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[i + 2] || p2;
      const c1x = p1[0] + (p2[0] - p0[0]) / 6;
      const c1y = p1[1] + (p2[1] - p0[1]) / 6;
      const c2x = p2[0] - (p3[0] - p1[0]) / 6;
      const c2y = p2[1] - (p3[1] - p1[1]) / 6;
      d += " C" + c1x + "," + c1y + " " + c2x + "," + c2y + " " + p2[0] + "," + p2[1];
    }
    return d;
  }

  function xAt(m) { return CH.L + (m / 18) * (CH.W - CH.L - CH.R); }
  function yAt(v) { return CH.T + (1 - v / 1000) * (CH.H - CH.T - CH.B); }

  function buildChart() {
    const svg = document.getElementById("growthChart");
    if (!svg) return;
    const data = [[0, 12.4], [1, 15], [3, 25], [6, 95], [9, 220], [12, 420], [15, 680], [18, 1000]];
    const pts = data.map(function (d) { return [xAt(d[0]), yAt(d[1])]; });
    const line = smoothPath(pts);
    const area = line + " L" + xAt(18) + "," + yAt(0) + " L" + xAt(0) + "," + yAt(0) + " Z";

    let grid = "";
    [0, 250, 500, 750, 1000].forEach(function (v) {
      const y = yAt(v);
      grid += '<line x1="' + CH.L + '" y1="' + y + '" x2="' + (CH.W - CH.R) + '" y2="' + y +
        '" stroke="rgba(255,255,255,.08)" stroke-dasharray="' + (v === 0 ? "0" : "3 5") + '"/>';
      const label = v === 1000 ? "1 jt" : v === 0 ? "0" : v + " rb";
      grid += '<text x="' + (CH.L - 12) + '" y="' + (y + 3.5) + '" text-anchor="end" fill="#6f7b93" font-size="9.5" font-family="inherit">' + label + '</text>';
    });

    [0, 3, 6, 9, 12, 15, 18].forEach(function (m) {
      const x = xAt(m);
      grid += '<line x1="' + x + '" y1="' + CH.T + '" x2="' + x + '" y2="' + (CH.H - CH.B) + '" stroke="rgba(255,255,255,.05)"/>';
      grid += '<text x="' + x + '" y="' + (CH.H - 16) + '" text-anchor="middle" fill="#6f7b93" font-size="9.5" font-family="inherit">' + (m === 0 ? "Mulai" : "Bln " + m) + '</text>';
    });

    const bands = [
      { x1: xAt(0), x2: xAt(6), label: "FONDASI" },
      { x1: xAt(6), x2: xAt(12), label: "AKSELERASI" },
      { x1: xAt(12), x2: xAt(18), label: "SKALA" }
    ].map(function (b) {
      return '<rect x="' + b.x1 + '" y="' + CH.T + '" width="' + (b.x2 - b.x1) + '" height="' + (CH.H - CH.T - CH.B) + '" fill="rgba(255,255,255,.018)"/>' +
        '<text x="' + (b.x1 + 10) + '" y="' + (CH.T + 14) + '" fill="rgba(255,255,255,.28)" font-size="8.5" letter-spacing="2" font-family="inherit">' + b.label + '</text>';
    }).join("");

    const labels = [[3, "25 rb"], [6, "95 rb"], [9, "220 rb"], [12, "420 rb"], [15, "680 rb"], [18, "1 juta"]];
    let dots = "";
    labels.forEach(function (l) {
      const x = xAt(l[0]), y = yAt(data.filter(function (d) { return d[0] === l[0]; })[0][1]);
      dots += '<circle cx="' + x + '" cy="' + y + '" r="4.5" fill="#070b14" stroke="#ff5a6a" stroke-width="2"/>';
      dots += '<text x="' + x + '" y="' + (y - 12) + '" text-anchor="middle" fill="#ffffff" font-size="10.5" font-weight="600" font-family="inherit">' + l[1] + '</text>';
    });

    svg.innerHTML =
      '<defs>' +
        '<linearGradient id="lineGrad" x1="0" y1="0" x2="1" y2="0">' +
          '<stop offset="0" stop-color="#ff4d5e"/><stop offset="0.55" stop-color="#c04dff"/><stop offset="1" stop-color="#38b6ff"/>' +
        '</linearGradient>' +
        '<linearGradient id="areaGrad" x1="0" y1="0" x2="0" y2="1">' +
          '<stop offset="0" stop-color="#ff4d5e" stop-opacity="0.20"/><stop offset="1" stop-color="#ff4d5e" stop-opacity="0"/>' +
        '</linearGradient>' +
      '</defs>' +
      bands + grid +
      '<path class="area" d="' + area + '" fill="url(#areaGrad)" opacity="0"/>' +
      '<path class="line" d="' + line + '" fill="none" stroke="url(#lineGrad)" stroke-width="3" stroke-linecap="round"/>' +
      dots +
      '<circle cx="' + xAt(18) + '" cy="' + yAt(1000) + '" r="7" fill="rgba(56,182,255,.18)"/>' +
      '<text x="' + (xAt(18) - 4) + '" y="' + (yAt(1000) + 26) + '" text-anchor="end" fill="#38b6ff" font-size="9.5" letter-spacing="1.5" font-family="inherit">TARGET 18 BULAN</text>';

    const lineEl = svg.querySelector(".line");
    const areaEl = svg.querySelector(".area");
    if (lineEl) {
      const len = lineEl.getTotalLength();
      lineEl.style.strokeDasharray = len;
      lineEl.style.strokeDashoffset = len;
      lineEl.style.transition = "stroke-dashoffset 1.8s cubic-bezier(.3,.8,.3,1)";
      areaEl.style.transition = "opacity 1.2s ease .5s";
      svg._line = lineEl;
      svg._area = areaEl;
    }
  }

  function animateChart() {
    const svg = document.getElementById("growthChart");
    if (!svg || !svg._line) return;
    // reset then play
    svg._line.style.transition = "none";
    svg._line.style.strokeDashoffset = svg._line.getTotalLength();
    svg._area.style.opacity = 0;
    void svg._line.getBoundingClientRect();
    requestAnimationFrame(function () {
      svg._line.style.transition = "stroke-dashoffset 1.8s cubic-bezier(.3,.8,.3,1)";
      svg._line.style.strokeDashoffset = 0;
      svg._area.style.opacity = 1;
    });
  }

  /* ---------------- boot ---------------- */
  buildOverview();
  buildChart();
  updateNotes();

  // page numbers in slide footers
  slides.forEach(function (s, i) {
    const f = s.querySelector(".fnum");
    if (f) f.textContent = String(i + 1).padStart(2, "0") + " / " + total;
  });

  // deep link
  const m = /#s(\d+)/.exec(window.location.hash);
  const startIndex = m ? Math.min(total, Math.max(1, Number(m[1]))) - 1 : 0;

  slides.forEach(function (s, i) { s.classList.toggle("active", i === startIndex); });
  index = startIndex;
  progress.style.width = ((index + 1) / total) * 100 + "%";
  counter.textContent = String(index + 1).padStart(2, "0") + " / " + total;
  updateOverviewCurrent();
  updateNotes();
  if (slides[index].querySelector("#growthChart")) animateChart();

  // first-visit hint
  if (!sessionStorage.getItem("uui-deck-visited")) {
    sessionStorage.setItem("uui-deck-visited", "1");
    setTimeout(function () { showToast("Gunakan ← → untuk berpindah slide · tekan O untuk ringkasan"); }, 800);
  }

  window._deck = { go: go, next: next, prev: prev, total: total };
})();
