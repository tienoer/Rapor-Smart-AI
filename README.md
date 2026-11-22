<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>RAPOR SMART AI — Dari Data Menjadi Aksi, Dari Rapor Menjadi Perubahan</title>

  <!-- Google Fonts: Plus Jakarta Sans & JetBrains Mono -->
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['"Plus Jakarta Sans"', 'sans-serif'],
            mono: ['"JetBrains Mono"', 'monospace'],
          },
          colors: {
            emerald: {
              850: '#064e3b',
              900: '#043927',
              950: '#022c22',
            },
            gold: {
              50: '#fffbeb',
              100: '#fef3c7',
              200: '#fde68a',
              400: '#fbbf24',
              500: '#f59e0b',
              600: '#d97706',
              700: '#b45309',
            }
          }
        }
      }
    };
  </script>

  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>

  <!-- Chart.js for Interactive Analytics -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>

  <!-- SheetJS (XLSX) for Real Excel Import & Export -->
  <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>

  <!-- PDF.js for Real PDF Document Reading -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

  <!-- html2pdf.js for Real PDF Report Export -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>

  <style>
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: #f8fafc;
      color: #0f172a;
    }
    /* Custom subtle scrollbar */
    ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    ::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    ::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 9999px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }
    .gold-gradient {
      background: linear-gradient(135deg, #d97706 0%, #f59e0b 100%);
    }
    .emerald-gradient {
      background: linear-gradient(135deg, #043927 0%, #065f46 100%);
    }
    @media print {
      .no-print {
        display: none !important;
      }
      body {
        background: #ffffff !important;
      }
    }
  </style>
</head>
<body class="antialiased selection:bg-emerald-800 selection:text-white">

  <!-- ===================================================================== -->
  <!-- TOP DEMO MODE BANNER (SHOWN ONLY WHEN DEMO SIMULATION IS ACTIVE)      -->
  <!-- ===================================================================== -->
  <div id="demoBanner" class="hidden bg-amber-500 text-emerald-950 px-4 py-2 text-xs sm:text-sm font-semibold flex flex-wrap items-center justify-between gap-2 z-50 relative shadow-sm">
    <div class="flex items-center gap-2">
      <i data-lucide="flask-conical" class="w-4 h-4 shrink-0"></i>
      <span><strong>MODE SIMULASI / DATA CONTOH AKTIF:</strong> Data yang tampil adalah data simulasi pelatihan, bukan data resmi sekolah Anda.</span>
    </div>
    <button onclick="promptResetToCleanState()" class="bg-emerald-950 text-white hover:bg-emerald-900 px-3 py-1 rounded-lg text-xs font-semibold transition">
      Kosongkan &amp; Input Data Sekolah Asli
    </button>
  </div>

  <!-- ===================================================================== -->
  <!-- MAIN APPLICATION SHELL                                                -->
  <!-- ===================================================================== -->
  <div class="min-h-screen flex flex-col lg:flex-row">

    <!-- MOBILE HEADER -->
    <header class="lg:hidden emerald-gradient text-white px-4 py-3 flex items-center justify-between sticky top-0 z-40 shadow-md">
      <div class="flex items-center gap-3">
        <button onclick="toggleMobileSidebar()" class="p-2 rounded-xl bg-white/10 hover:bg-white/20 transition" aria-label="Buka Menu">
          <i data-lucide="menu" class="w-5 h-5"></i>
        </button>
        <div class="flex items-center gap-2">
          <!-- Custom SVG Emblem -->
          <svg class="w-8 h-8 shrink-0" viewBox="0 0 40 40" fill="none">
            <rect width="40" height="40" rx="10" fill="#D97706" fill-opacity="0.2" stroke="#F59E0B" stroke-width="1.5"/>
            <path d="M11 27V15L20 10L29 15V27L20 32L11 27Z" stroke="#FBBF24" stroke-width="2"/>
            <path d="M16 22L19 25L25 17" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
          <div>
            <div class="font-extrabold tracking-tight text-sm leading-tight">RAPOR SMART AI</div>
            <div class="text-[10px] text-amber-300 font-medium">Perencanaan Berbasis Data</div>
          </div>
        </div>
      </div>
      <div class="flex items-center gap-2">
        <button onclick="toggleAssistantDrawer()" class="p-2 rounded-xl bg-amber-500 text-emerald-950 font-bold text-xs flex items-center gap-1 shadow">
          <i data-lucide="sparkles" class="w-4 h-4"></i>
          <span>AI</span>
        </button>
      </div>
    </header>

    <!-- SIDEBAR BACKDROP (MOBILE) -->
    <div id="sidebarBackdrop" onclick="toggleMobileSidebar()" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-40 hidden lg:hidden"></div>

    <!-- LEFT NAVIGATION SIDEBAR -->
    <aside id="appSidebar" class="fixed lg:sticky top-0 left-0 z-50 h-screen w-72 emerald-gradient text-white flex flex-col justify-between transition-transform duration-300 -translate-x-full lg:translate-x-0 shadow-2xl shrink-0">
      <div class="overflow-y-auto flex-1">
        <!-- Brand Header -->
        <div class="p-5 border-b border-white/10">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-3">
              <svg class="w-10 h-10 shrink-0" viewBox="0 0 40 40" fill="none">
                <rect width="40" height="40" rx="11" fill="#D97706" fill-opacity="0.22" stroke="#F59E0B" stroke-width="1.5"/>
                <path d="M11 26V15L20 10L29 15V26L20 31L11 26Z" stroke="#FBBF24" stroke-width="2"/>
                <path d="M15.5 21.5L18.8 24.8L25 16.8" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
              <div>
                <h1 class="font-extrabold text-base tracking-tight text-white">RAPOR SMART AI</h1>
                <p class="text-[11px] text-amber-300 leading-snug font-medium">Dari Data Menjadi Aksi,<br/>Dari Rapor Menjadi Perubahan.</p>
              </div>
            </div>
            <button onclick="toggleMobileSidebar()" class="lg:hidden p-1.5 rounded-lg bg-white/10 text-white">
              <i data-lucide="x" class="w-4 h-4"></i>
            </button>
          </div>

          <!-- Active Year & School Mini Pill -->
          <div class="mt-4 bg-black/25 rounded-xl p-3 border border-white/10">
            <div class="flex items-center justify-between text-[11px] text-emerald-200 mb-1">
              <span>Tahun Perencanaan</span>
              <span id="sidebarCloudBadge" class="inline-flex items-center gap-1 text-amber-300 font-semibold">
                <span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span> Lokal &amp; File
              </span>
            </div>
            <div class="flex items-center gap-2">
              <select id="globalYearSelect" onchange="switchPlanningYear(this.value)" class="w-full bg-emerald-950/90 border border-amber-500/40 text-white text-xs font-bold rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-amber-400">
              </select>
              <button onclick="openNewYearModal()" title="Tambah Tahun Perencanaan" class="p-1.5 rounded-lg bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-400/30 transition">
                <i data-lucide="plus" class="w-4 h-4"></i>
              </button>
            </div>
          </div>
        </div>

        <!-- Navigation Menu -->
        <nav class="p-3 space-y-1 text-xs font-medium" id="sidebarNav">
          <div class="px-3 pt-2 pb-1 text-[10px] font-bold uppercase tracking-wider text-emerald-300/80">Ringkasan Eksekutif</div>
          <button data-tab="dashboard" onclick="navigateTo('dashboard')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition text-left">
            <i data-lucide="layout-dashboard" class="w-4 h-4 shrink-0"></i>
            <span>Dashboard Utama</span>
          </button>

          <div class="px-3 pt-3 pb-1 text-[10px] font-bold uppercase tracking-wider text-emerald-300/80">Alur Perencanaan PBD</div>
          <button data-tab="input" onclick="navigateTo('input')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="file-input" class="w-4 h-4 shrink-0"></i><span>1. Input Data Rapor</span></span>
            <span id="badge-input" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="analisis" onclick="navigateTo('analisis')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="bar-chart-3" class="w-4 h-4 shrink-0"></i><span>2. Analisis Rapor</span></span>
            <span id="badge-analisis" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="akar" onclick="navigateTo('akar')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="git-fork" class="w-4 h-4 shrink-0"></i><span>3. Akar Masalah</span></span>
            <span id="badge-akar" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="prioritas" onclick="navigateTo('prioritas')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="list-ordered" class="w-4 h-4 shrink-0"></i><span>4. Penentuan Prioritas</span></span>
            <span id="badge-prioritas" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="rekomendasi" onclick="navigateTo('rekomendasi')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="lightbulb" class="w-4 h-4 shrink-0"></i><span>5. Rekomendasi Program</span></span>
            <span id="badge-rekomendasi" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="rkjm" onclick="navigateTo('rkjm')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="calendar-range" class="w-4 h-4 shrink-0"></i><span>6. Penyusunan RKJM</span></span>
            <span id="badge-rkjm" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="rkt" onclick="navigateTo('rkt')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="calendar-check-2" class="w-4 h-4 shrink-0"></i><span>7. Penyusunan RKT</span></span>
            <span id="badge-rkt" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>
          <button data-tab="arkas" onclick="navigateTo('arkas')" class="nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left">
            <span class="flex items-center gap-3"><i data-lucide="wallet" class="w-4 h-4 shrink-0"></i><span>8. Rancangan ARKAS</span></span>
            <span id="badge-arkas" class="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10">0</span>
          </button>

          <div class="px-3 pt-3 pb-1 text-[10px] font-bold uppercase tracking-wider text-emerald-300/80">Integrasi, Monev &amp; Ekspor</div>
          <button data-tab="keterhubungan" onclick="navigateTo('keterhubungan')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition text-left">
            <i data-lucide="network" class="w-4 h-4 shrink-0"></i>
            <span>9. Hubungan Antardokumen</span>
          </button>
          <button data-tab="monev" onclick="navigateTo('monev')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition text-left">
            <i data-lucide="activity" class="w-4 h-4 shrink-0"></i>
            <span>10. Monitoring &amp; Evaluasi</span>
          </button>
          <button data-tab="ekspor" onclick="navigateTo('ekspor')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition text-left">
            <i data-lucide="file-down" class="w-4 h-4 shrink-0"></i>
            <span>11. Laporan &amp; Ekspor</span>
          </button>
          <button data-tab="manajemen" onclick="navigateTo('manajemen')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition text-left">
            <i data-lucide="database" class="w-4 h-4 shrink-0"></i>
            <span>12. Manajemen Data</span>
          </button>
        </nav>
      </div>

      <!-- Sidebar Footer Actions -->
      <div class="p-3 border-t border-white/10 bg-black/20 space-y-2">
        <button onclick="openUserGuideModal()" class="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-xl bg-white/10 hover:bg-white/15 text-xs font-semibold text-white transition">
          <i data-lucide="book-open-check" class="w-4 h-4 text-amber-300"></i>
          <span>Panduan Pengguna Baru</span>
        </button>
        <div class="flex items-center justify-between text-[10px] text-emerald-300/80 px-1">
          <span id="autoSaveIndicator">Tersimpan otomatis di sesi</span>
          <span>v2.5 PBD</span>
        </div>
      </div>
    </aside>

    <!-- MAIN CONTENT WORKSPACE -->
    <main class="flex-1 flex flex-col min-w-0">

      <!-- DESKTOP TOP UTILITY BAR -->
      <div class="no-print bg-white border-b border-slate-200 px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-3 sticky top-0 z-30">
        <div class="flex items-center gap-3 min-w-0">
          <div class="w-2.5 h-8 rounded-full bg-emerald-800 shrink-0 hidden sm:block"></div>
          <div class="min-w-0">
            <div class="flex items-center gap-2 flex-wrap">
              <h2 id="headerSchoolName" class="font-bold text-slate-900 text-sm sm:text-base truncate">Satuan Pendidikan Belum Diisi</h2>
              <span id="headerNpsnBadge" class="text-[11px] font-mono font-semibold px-2 py-0.5 rounded-md bg-slate-100 text-slate-600 border border-slate-200">NPSN: -</span>
              <span id="headerJenjangBadge" class="text-[11px] font-semibold px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 border border-emerald-200">Jenjang: -</span>
            </div>
            <p id="headerSubtitle" class="text-xs text-slate-500 truncate">Silakan lengkapi identitas sekolah &amp; data Rapor Pendidikan untuk memulai perencanaan.</p>
          </div>
        </div>

        <!-- Global Quick Actions -->
        <div class="flex items-center gap-2 flex-wrap">
          <button id="btnToggleDemo" onclick="toggleDemoSimulation()" class="px-3 py-2 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-900 text-xs font-semibold flex items-center gap-1.5 transition">
            <i data-lucide="play-circle" class="w-4 h-4 text-amber-600"></i>
            <span id="btnToggleDemoLabel">Muat Data Demo</span>
          </button>

          <button onclick="exportFullWorkspaceJSON()" class="px-3 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold flex items-center gap-1.5 transition" title="Simpan cadangan file proyek (.json)">
            <i data-lucide="save" class="w-4 h-4 text-emerald-700"></i>
            <span class="hidden sm:inline">Simpan Proyek</span>
          </button>

          <button onclick="toggleAssistantDrawer()" class="px-3.5 py-2 rounded-xl gold-gradient hover:opacity-95 text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm transition">
            <i data-lucide="sparkles" class="w-4 h-4"></i>
            <span>ASISTEN RAPOR SMART AI</span>
          </button>
        </div>
      </div>

      <!-- DYNAMIC MODULE CONTAINER -->
      <div id="moduleContainer" class="flex-1 p-4 sm:p-6 lg:p-8 max-w-7xl w-full mx-auto space-y-6">
        <!-- Rendered dynamically by renderActiveModule() -->
      </div>

      <!-- FOOTER -->
      <footer class="no-print border-t border-slate-200 bg-white px-6 py-4 text-xs text-slate-500 flex flex-col sm:flex-row items-center justify-between gap-2">
        <div>
          <span class="font-bold text-emerald-900">RAPOR SMART AI</span> — <em>"Dari Data Menjadi Aksi, Dari Rapor Menjadi Perubahan."</em>
        </div>
        <div class="text-center sm:text-right">
          Alat bantu perencanaan berbasis data (PBD) untuk Kepala Sekolah, Guru, Bendahara BOS, Operator &amp; TPS. Seluruh hasil AI wajib diverifikasi tim sekolah.
        </div>
      </footer>
    </main>
  </div>

  <!-- ===================================================================== -->
  <!-- SLIDE-OVER CHATBOT: ASISTEN RAPOR SMART AI                            -->
  <!-- ===================================================================== -->
  <div id="aiAssistantBackdrop" onclick="toggleAssistantDrawer()" class="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 hidden"></div>
  <aside id="aiAssistantDrawer" class="fixed top-0 right-0 z-50 h-screen w-full sm:w-[440px] bg-white shadow-2xl border-l border-slate-200 flex flex-col translate-x-full transition-transform duration-300">
    <!-- Assistant Header -->
    <div class="emerald-gradient text-white p-4 flex items-center justify-between shrink-0">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-amber-400/20 border border-amber-300/40 flex items-center justify-center">
          <i data-lucide="bot" class="w-5 h-5 text-amber-300"></i>
        </div>
        <div>
          <h3 class="font-bold text-sm tracking-tight flex items-center gap-1.5">
            ASISTEN RAPOR SMART AI
            <span class="text-[10px] px-1.5 py-0.5 rounded bg-amber-400 text-emerald-950 font-extrabold">AKTIF</span>
          </h3>
          <p class="text-[11px] text-emerald-200">Konsultan Perencanaan Sekolah &amp; ARKAS Berbasis Data</p>
        </div>
      </div>
      <button onclick="toggleAssistantDrawer()" class="p-2 rounded-xl bg-white/10 hover:bg-white/20 text-white transition">
        <i data-lucide="x" class="w-4 h-4"></i>
      </button>
    </div>

    <!-- Grounding Notice -->
    <div class="bg-amber-50 border-b border-amber-200 px-4 py-2 text-[11px] text-amber-900 flex items-start gap-2 shrink-0">
      <i data-lucide="shield-check" class="w-4 h-4 text-amber-600 shrink-0 mt-0.5"></i>
      <span>Jawaban didasarkan secara ketat pada data sekolah yang telah Anda input. Asisten tidak mengarang angka dan akan meminta konfirmasi jika data belum lengkap.</span>
    </div>

    <!-- Chat History Messages -->
    <div id="aiChatMessages" class="flex-1 overflow-y-auto p-4 space-y-3 bg-slate-50">
      <!-- Initial greeting injected by JS -->
    </div>

    <!-- Quick Prompt Chips -->
    <div class="px-3 py-2 bg-white border-t border-slate-100 flex gap-1.5 overflow-x-auto shrink-0">
      <button onclick="sendQuickAiPrompt('Jelaskan kondisi Rapor Pendidikan sekolah kami saat ini berdasarkan data yang ada.')" class="shrink-0 text-[11px] px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-900 font-semibold border border-emerald-200 transition">
        Ringkas Rapor
      </button>
      <button onclick="sendQuickAiPrompt('Bantu identifikasi akar masalah utama dan pertanyaan refleksi untuk diskusi guru.')" class="shrink-0 text-[11px] px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-900 font-semibold border border-emerald-200 transition">
        Akar Masalah
      </button>
      <button onclick="sendQuickAiPrompt('Periksa keterhubungan antara Rapor Pendidikan, RKJM, RKT, dan Rancangan ARKAS kami.')" class="shrink-0 text-[11px] px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-900 font-semibold border border-emerald-200 transition">
        Cek Benang Merah
      </button>
      <button onclick="sendQuickAiPrompt('Evaluasi serapan anggaran dan progres pelaksanaan program di Monev.')" class="shrink-0 text-[11px] px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-900 font-semibold border border-emerald-200 transition">
        Evaluasi Monev
      </button>
    </div>

    <!-- Chat Input Box -->
    <form onsubmit="handleAiChatSubmit(event)" class="p-3 bg-white border-t border-slate-200 flex items-end gap-2 shrink-0">
      <textarea id="aiChatInput" rows="2" placeholder="Tanyakan analisis rapor, RKJM, RKT, atau rancangan anggaran..." class="flex-1 resize-none rounded-xl border border-slate-300 px-3 py-2 text-xs focus:outline-none focus:border-emerald-700 focus:ring-2 focus:ring-emerald-700/15" onkeydown="if(event.key==='Enter' && !event.shiftKey){event.preventDefault(); handleAiChatSubmit(event);}"></textarea>
      <button type="submit" id="aiChatSendBtn" class="p-3 rounded-xl emerald-gradient text-white hover:opacity-95 transition shrink-0 shadow-sm" title="Kirim Pertanyaan">
        <i data-lucide="send" class="w-4 h-4"></i>
      </button>
    </form>
  </aside>

  <!-- ===================================================================== -->
  <!-- UNIVERSAL MODAL CONTAINER (NO BROWSER ALERT/CONFIRM USED)             -->
  <!-- ===================================================================== -->
  <div id="universalModalBackdrop" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 hidden items-center justify-center p-4 overflow-y-auto">
    <div id="universalModalBox" class="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl overflow-hidden my-auto transform transition-all">
      <!-- Populated dynamically -->
    </div>
  </div>

  <!-- TOAST NOTIFICATION CONTAINER -->
  <div id="toastContainer" class="fixed bottom-4 right-4 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

  <!-- HIDDEN PRINT/EXPORT STAGING AREA -->
  <div id="exportStagingArea" class="hidden"></div>

  <!-- ===================================================================== -->
  <!-- APPLICATION LOGIC & FIREBASE/GEMINI INTEGRATION                       -->
  <!-- ===================================================================== -->
  <script type="module">
    import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
    import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
    import { getFirestore, doc, setDoc, getDoc } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

    // Configure PDF.js worker
    if (window.pdfjsLib) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    // Optional Cloud Persistence Setup (activated ONLY if environment provides __firebase_config)
    window.cloudSync = {
      enabled: false,
      db: null,
      auth: null,
      user: null,
      appId: typeof __app_id !== 'undefined' ? __app_id : 'rapor-smart-ai',
      async init() {
        try {
          if (typeof __firebase_config === 'undefined' || !__firebase_config) {
            return;
          }
          const firebaseConfig = typeof __firebase_config === 'string' ? JSON.parse(__firebase_config) : __firebase_config;
          const app = initializeApp(firebaseConfig);
          this.auth = getAuth(app);
          this.db = getFirestore(app);

          // Rule 3: Auth before queries
          if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
            await signInWithCustomToken(this.auth, __initial_auth_token);
          } else {
            await signInAnonymously(this.auth);
          }

          onAuthStateChanged(this.auth, async (user) => {
            if (user) {
              this.user = user;
              this.enabled = true;
              window.updateCloudBadge('connected');
              await this.loadFromCloud();
            }
          });
        } catch (err) {
          console.warn('Cloud sync tidak tersedia, menggunakan sesi lokal & ekspor file:', err);
          window.updateCloudBadge('local');
        }
      },
      async saveToCloud(stateObj) {
        if (!this.enabled || !this.user || !this.db) return;
        try {
          window.updateCloudBadge('syncing');
          // Rule 1: Strict private user path
          const docRef = doc(this.db, 'artifacts', this.appId, 'users', this.user.uid, 'rapor_smart_state', 'workspace');
          await setDoc(docRef, {
            payload: JSON.stringify(stateObj),
            updatedAt: new Date().toISOString()
          });
          window.updateCloudBadge('connected');
        } catch (err) {
          console.warn('Gagal menyinkronkan ke cloud:', err);
          window.updateCloudBadge('error');
        }
      },
      async loadFromCloud() {
        if (!this.enabled || !this.user || !this.db) return;
        try {
          const docRef = doc(this.db, 'artifacts', this.appId, 'users', this.user.uid, 'rapor_smart_state', 'workspace');
          const snap = await getDoc(docRef);
          if (snap.exists() && snap.data().payload) {
            const parsed = JSON.parse(snap.data().payload);
            if (parsed && parsed.yearsData) {
              window.appState = parsed;
              window.renderAll();
              window.showToast('Data perencanaan sekolah dipulihkan dari basis data.', 'success');
            }
          }
        } catch (err) {
          console.warn('Gagal memuat dari cloud:', err);
        }
      }
    };

    window.addEventListener('DOMContentLoaded', () => {
      window.cloudSync.init();
    });
  </script>

  <script>
    // =========================================================================
    // 1. CONSTANTS, REFERENCE CATALOGS & INITIAL EMPTY STATE
    // =========================================================================
    const DIMENSI_RAPOR = [
      "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
      "Dimensi C: Kompetensi dan Kinerja GTK",
      "Dimensi D: Mutu dan Relevansi Pembelajaran",
      "Dimensi E: Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel"
    ];

    const KATALOG_INDIKATOR_STANDAR = [
      { kode: "A.1", nama: "Kemampuan Literasi", dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "A.2", nama: "Kemampuan Numerasi", dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "A.3", nama: "Karakter", dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "C.1", nama: "Proporsi GTK Bersertifikat", dimensi: "Dimensi C: Kompetensi dan Kinerja GTK" },
      { kode: "C.3", nama: "Pengalaman Pelatihan GTK", dimensi: "Dimensi C: Kompetensi dan Kinerja GTK" },
      { kode: "D.1", nama: "Kualitas Pembelajaran", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.2", nama: "Refleksi dan Perbaikan Pembelajaran oleh Guru", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.3", nama: "Kepemimpinan Instruksional", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.4", nama: "Iklim Keamanan Satuan Pendidikan", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.6", nama: "Iklim Kesetaraan Gender", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.8", nama: "Iklim Kebinekaan", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "D.10", nama: "Iklim Inklusivitas", dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran" },
      { kode: "E.1", nama: "Partisipasi Warga Satuan Pendidikan", dimensi: "Dimensi E: Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" },
      { kode: "E.2", nama: "Proporsi Pemanfaatan Sumber Daya Sekolah untuk Peningkatan Mutu", dimensi: "Dimensi E: Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" },
      { kode: "E.3", nama: "Pemanfaatan TIK untuk Pengelolaan Anggaran", dimensi: "Dimensi E: Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" }
    ];

    // Standar Komponen Penggunaan Dana BOSP sesuai Juknis Kemendikbudristek/Kemendikdasmen
    const KOMPONEN_BOSP_STANDAR = [
      "Penerimaan Peserta Didik Baru (PPDB)",
      "Pengembangan Perpustakaan dan/atau Layanan Pojok Baca",
      "Pelaksanaan Kegiatan Pembelajaran dan Ekstrakurikuler",
      "Pelaksanaan Kegiatan Asesmen dan Evaluasi Pembelajaran",
      "Pelaksanaan Administrasi Kegiatan Satuan Pendidikan",
      "Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
      "Pembiayaan Langganan Daya dan Jasa",
      "Pemeliharaan Sarana dan Prasarana Satuan Pendidikan",
      "Penyediaan Alat Multimedia Pembelajaran",
      "Pembayaran Honor",
      "Komponen Lainnya (Perlu Verifikasi Juknis BOSP & SSH Daerah)"
    ];

    const BULAN_LIST = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

    function createEmptyYearData(year = "2026") {
      const yNum = parseInt(year, 10) || 2026;
      return {
        school: {
          nama: "",
          npsn: "",
          alamat: "",
          kecamatan: "",
          kabupaten: "",
          provinsi: "",
          jenjang: "",
          tahunRapor: String(yNum - 1),
          tahunPerencanaan: String(yNum),
          tahunAjaran: `${yNum}/${yNum + 1}`,
          periodeRkjmMulai: String(yNum),
          periodeRkjmAkhir: String(yNum + 3)
        },
        indicators: [],
        rootCauses: {},
        priorities: {},
        recommendations: [],
        rkjm: [],
        rkt: [],
        arkas: [],
        monev: {},
        history: [
          {
            id: uid(),
            timestamp: new Date().toLocaleString('id-ID'),
            modul: "Sistem",
            aksi: "Inisialisasi Tahun Perencanaan",
            keterangan: `Ruang kerja tahun anggaran ${yNum} siap diisi data sekolah.`
          }
        ]
      };
    }

    // Main Application State (Clean by default - never presents demo data as real school data)
    window.appState = {
      activeTab: "dashboard",
      activeYear: "2026",
      isDemoMode: false,
      analisisFilter: "semua",
      selectedAkarIndicatorId: null,
      selectedTraceIndicatorId: null,
      yearsData: {
        "2026": createEmptyYearData("2026")
      },
      chatMessages: [
        {
          role: "assistant",
          text: "Halo Bapak/Ibu Kepala Sekolah dan Tim Pengembang Sekolah! Saya **ASISTEN RAPOR SMART AI**.\n\nSaya siap membantu mengolah data Rapor Pendidikan menjadi Analisis Akar Masalah, Prioritas Pembenahan, RKJM, RKT, hingga Rancangan Anggaran ARKAS.\n\nSilakan mulai dengan mengisi **Identitas Sekolah & Data Rapor Pendidikan** pada menu **1. Input Data Rapor**, atau klik **Muat Data Demo** di pojok kanan atas jika ingin mencoba simulasi terlebih dahulu."
        }
      ]
    };

    // Chart instances registry
    let chartInstances = {};

    // =========================================================================
    // 2. UTILITY FUNCTIONS & HELPERS
    // =========================================================================
    function uid() {
      return 'id-' + Math.random().toString(36).substring(2, 9) + '-' + Date.now().toString(36).slice(-4);
    }

    function getCurrentData() {
      if (!window.appState.yearsData[window.appState.activeYear]) {
        window.appState.yearsData[window.appState.activeYear] = createEmptyYearData(window.appState.activeYear);
      }
      return window.appState.yearsData[window.appState.activeYear];
    }

    function logHistory(modul, aksi, keterangan) {
      const data = getCurrentData();
      data.history.unshift({
        id: uid(),
        timestamp: new Date().toLocaleString('id-ID'),
        modul,
        aksi,
        keterangan
      });
      if (data.history.length > 100) data.history.pop();
    }

    function triggerAutoSave(modul, aksi, keterangan) {
      if (modul && aksi) {
        logHistory(modul, aksi, keterangan || "");
      }
      const el = document.getElementById('autoSaveIndicator');
      if (el) {
        el.textContent = `Tersimpan ${new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`;
      }
      if (window.cloudSync && window.cloudSync.enabled) {
        window.cloudSync.saveToCloud(window.appState);
      }
      updateSidebarBadges();
      updateTopHeader();
    }

    function formatRupiah(angka) {
      const num = Number(angka) || 0;
      return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(num);
    }

    function escapeHtml(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    function classifyCategory(nilai, explicitKat) {
      if (explicitKat && ["Baik", "Sedang", "Kurang"].includes(explicitKat)) return explicitKat;
      const v = parseFloat(nilai);
      if (isNaN(v)) return "Belum Terverifikasi";
      if (v >= 70) return "Baik";
      if (v >= 50) return "Sedang";
      return "Kurang";
    }

    function needsImprovement(ind) {
      const kat = classifyCategory(ind.nilai, ind.kategori);
      const delta = parseFloat(ind.perubahan);
      return kat === "Kurang" || kat === "Sedang" || (!isNaN(delta) && delta < 0);
    }

    function getCategoryBadgeHtml(kat) {
      if (kat === "Baik") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200"><span class="w-1.5 h-1.5 rounded-full bg-emerald-600"></span>Baik</span>`;
      }
      if (kat === "Sedang") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-800 border border-amber-200"><span class="w-1.5 h-1.5 rounded-full bg-amber-500"></span>Sedang</span>`;
      }
      if (kat === "Kurang") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-rose-100 text-rose-800 border border-rose-200"><span class="w-1.5 h-1.5 rounded-full bg-rose-600"></span>Kurang</span>`;
      }
      return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-slate-100 text-slate-600 border border-slate-200">Perlu Konfirmasi</span>`;
    }

    function getDeltaBadgeHtml(deltaVal) {
      const d = parseFloat(deltaVal);
      if (isNaN(d)) {
        return `<span class="text-xs text-slate-400 italic">Tidak ada data</span>`;
      }
      if (d > 0) {
        return `<span class="inline-flex items-center gap-0.5 text-xs font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md"><i data-lucide="trending-up" class="w-3.5 h-3.5"></i>+${d.toFixed(2)}</span>`;
      }
      if (d < 0) {
        return `<span class="inline-flex items-center gap-0.5 text-xs font-bold text-rose-700 bg-rose-50 px-2 py-0.5 rounded-md"><i data-lucide="trending-down" class="w-3.5 h-3.5"></i>${d.toFixed(2)}</span>`;
      }
      return `<span class="inline-flex items-center gap-0.5 text-xs font-semibold text-slate-600 bg-slate-100 px-2 py-0.5 rounded-md">0.00 (Tetap)</span>`;
    }

    // =========================================================================
    // 3. CUSTOM TOAST & MODAL SYSTEM (STRICTLY NO ALERT / CONFIRM)
    // =========================================================================
    function showToast(message, type = 'info') {
      const container = document.getElementById('toastContainer');
      const id = uid();
      const styles = {
        success: 'bg-emerald-900 text-white border-emerald-700',
        warning: 'bg-amber-600 text-white border-amber-500',
        error: 'bg-rose-800 text-white border-rose-700',
        info: 'bg-slate-900 text-white border-slate-700'
      };
      const icons = {
        success: 'check-circle-2',
        warning: 'alert-triangle',
        error: 'x-circle',
        info: 'info'
      };
      const div = document.createElement('div');
      div.id = id;
      div.className = `pointer-events-auto flex items-start gap-2.5 px-4 py-3 rounded-xl shadow-lg border text-xs font-medium transition-all ${styles[type] || styles.info}`;
      div.innerHTML = `
        <i data-lucide="${icons[type] || 'info'}" class="w-4 h-4 shrink-0 mt-0.5"></i>
        <div class="flex-1 leading-relaxed">${escapeHtml(message)}</div>
      `;
      container.appendChild(div);
      if (window.lucide) window.lucide.createIcons();
      setTimeout(() => {
        div.style.opacity = '0';
        div.style.transform = 'translateY(8px)';
        setTimeout(() => div.remove(), 300);
      }, 3800);
    }

    function openCustomModal(htmlContent, maxWidthClass = 'max-w-2xl') {
      const backdrop = document.getElementById('universalModalBackdrop');
      const box = document.getElementById('universalModalBox');
      box.className = `bg-white rounded-2xl shadow-2xl border border-slate-200 w-full ${maxWidthClass} overflow-hidden my-auto`;
      box.innerHTML = htmlContent;
      backdrop.classList.remove('hidden');
      backdrop.classList.add('flex');
      if (window.lucide) window.lucide.createIcons();
    }

    function closeCustomModal() {
      const backdrop = document.getElementById('universalModalBackdrop');
      backdrop.classList.add('hidden');
      backdrop.classList.remove('flex');
    }

    function showConfirmDialog({ title, message, confirmText = "Ya, Lanjutkan", danger = false, onConfirm }) {
      const btnClass = danger
        ? "bg-rose-600 hover:bg-rose-700 text-white"
        : "emerald-gradient hover:opacity-95 text-white";
      openCustomModal(`
        <div class="p-6">
          <div class="flex items-start gap-3">
            <div class="w-10 h-10 rounded-xl ${danger ? 'bg-rose-100 text-rose-700' : 'bg-amber-100 text-amber-700'} flex items-center justify-center shrink-0">
              <i data-lucide="${danger ? 'trash-2' : 'help-circle'}" class="w-5 h-5"></i>
            </div>
            <div>
              <h3 class="font-bold text-slate-900 text-base">${escapeHtml(title)}</h3>
              <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">${escapeHtml(message)}</p>
            </div>
          </div>
          <div class="mt-6 flex items-center justify-end gap-2">
            <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-50 text-xs font-semibold">Batal</button>
            <button id="modalConfirmActionBtn" class="px-4 py-2 rounded-xl ${btnClass} text-xs font-bold shadow-sm">${escapeHtml(confirmText)}</button>
          </div>
        </div>
      `, 'max-w-md');
      document.getElementById('modalConfirmActionBtn').onclick = () => {
        closeCustomModal();
        if (typeof onConfirm === 'function') onConfirm();
      };
    }

    // =========================================================================
    // 4. GEMINI AI ENGINE (WITH EXPONENTIAL BACKOFF & GRACEFUL FALLBACK)
    // =========================================================================
    async function callGeminiAPI(promptText, systemInstructionText, jsonSchema = null) {
      const apiKey = ""; // Runtime-provided in Canvas environment
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

      const payload = {
        contents: [{ parts: [{ text: promptText }] }],
        systemInstruction: {
          parts: [{ text: systemInstructionText }]
        }
      };

      if (jsonSchema) {
        payload.generationConfig = {
          responseMimeType: "application/json",
          responseSchema: jsonSchema
        };
      }

      const delays = [1000, 2000, 4000];
      for (let attempt = 0; attempt <= delays.length; attempt++) {
        try {
          const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });

          if (!response.ok) {
            if (attempt < delays.length && (response.status === 429 || response.status >= 500)) {
              await new Promise(res => setTimeout(res, delays[attempt]));
              continue;
            }
            throw new Error(`HTTP ${response.status}`);
          }

          const result = await response.json();
          const text = result?.candidates?.[0]?.content?.parts?.[0]?.text;
          if (!text) throw new Error("Respons AI kosong");
          return jsonSchema ? JSON.parse(text) : text;
        } catch (err) {
          if (attempt < delays.length) {
            await new Promise(res => setTimeout(res, delays[attempt]));
            continue;
          }
          throw err;
        }
      }
    }

    // =========================================================================
    // 5. NAVIGATION, TOP HEADER & SIDEBAR STATE
    // =========================================================================
    function toggleMobileSidebar() {
      const sidebar = document.getElementById('appSidebar');
      const backdrop = document.getElementById('sidebarBackdrop');
      const isClosed = sidebar.classList.contains('-translate-x-full');
      if (isClosed) {
        sidebar.classList.remove('-translate-x-full');
        backdrop.classList.remove('hidden');
      } else {
        sidebar.classList.add('-translate-x-full');
        backdrop.classList.add('hidden');
      }
    }

    function navigateTo(tabName) {
      window.appState.activeTab = tabName;
      // Close mobile sidebar if open
      const sidebar = document.getElementById('appSidebar');
      const backdrop = document.getElementById('sidebarBackdrop');
      if (sidebar && !sidebar.classList.contains('-translate-x-full') && window.innerWidth < 1024) {
        sidebar.classList.add('-translate-x-full');
        backdrop.classList.add('hidden');
      }
      renderAll();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function updateCloudBadge(status) {
      const el = document.getElementById('sidebarCloudBadge');
      if (!el) return;
      if (status === 'connected') {
        el.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-emerald-400"></span> Cloud Tersinkron`;
      } else if (status === 'syncing') {
        el.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400 animate-ping"></span> Menyimpan...`;
      } else if (status === 'error') {
        el.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-rose-400"></span> Offline (Lokal)`;
      } else {
        el.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span> Lokal &amp; File`;
      }
    }

    function updateYearSelector() {
      const sel = document.getElementById('globalYearSelect');
      if (!sel) return;
      const years = Object.keys(window.appState.yearsData).sort();
      sel.innerHTML = years.map(y => `<option value="${y}" ${y === window.appState.activeYear ? 'selected' : ''}>Tahun ${y}</option>`).join('');
    }

    function switchPlanningYear(year) {
      window.appState.activeYear = year;
      renderAll();
      showToast(`Beralih ke Tahun Perencanaan ${year}.`, 'info');
    }

    function updateTopHeader() {
      const data = getCurrentData();
      const s = data.school;
      document.getElementById('headerSchoolName').textContent = s.nama ? s.nama : "Satuan Pendidikan Belum Diisi";
      document.getElementById('headerNpsnBadge').textContent = `NPSN: ${s.npsn || '-'}`;
      document.getElementById('headerJenjangBadge').textContent = `Jenjang: ${s.jenjang || '-'}`;
      document.getElementById('headerSubtitle').textContent = s.nama
        ? `${s.kecamatan ? 'Kec. ' + s.kecamatan + ', ' : ''}${s.kabupaten || ''} • Rapor ${s.tahunRapor || '-'} → Perencanaan ${s.tahunPerencanaan || window.appState.activeYear}`
        : "Silakan lengkapi identitas sekolah & data Rapor Pendidikan untuk memulai perencanaan.";

      // Demo banner & toggle label
      const demoBanner = document.getElementById('demoBanner');
      const btnDemoLabel = document.getElementById('btnToggleDemoLabel');
      if (window.appState.isDemoMode) {
        demoBanner.classList.remove('hidden');
        if (btnDemoLabel) btnDemoLabel.textContent = "Keluar Mode Demo";
      } else {
        demoBanner.classList.add('hidden');
        if (btnDemoLabel) btnDemoLabel.textContent = "Muat Data Demo";
      }
    }

    function updateSidebarBadges() {
      const data = getCurrentData();
      const problemInds = data.indicators.filter(needsImprovement);
      const map = {
        'badge-input': data.indicators.length,
        'badge-analisis': data.indicators.length,
        'badge-akar': Object.keys(data.rootCauses).length,
        'badge-prioritas': problemInds.length,
        'badge-rekomendasi': data.recommendations.length,
        'badge-rkjm': data.rkjm.length,
        'badge-rkt': data.rkt.length,
        'badge-arkas': data.arkas.length
      };
      Object.entries(map).forEach(([id, count]) => {
        const el = document.getElementById(id);
        if (el) el.textContent = count;
      });

      // Highlight active tab
      document.querySelectorAll('.nav-btn').forEach(btn => {
        const isActive = btn.getAttribute('data-tab') === window.appState.activeTab;
        if (isActive) {
          btn.className = "nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left bg-amber-400 text-emerald-950 font-bold shadow-sm";
        } else {
          btn.className = "nav-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl transition text-left text-emerald-100 hover:bg-white/10 hover:text-white";
        }
      });
    }

    function renderAll() {
      updateYearSelector();
      updateTopHeader();
      updateSidebarBadges();
      renderActiveModule();
      if (window.lucide) window.lucide.createIcons();
    }

    // =========================================================================
    // 6. MODULE RENDERER DISPATCHER
    // =========================================================================
    function renderActiveModule() {
      // Destroy existing Chart.js instances before re-rendering DOM
      Object.values(chartInstances).forEach(ch => { try { ch.destroy(); } catch (e) {} });
      chartInstances = {};

      const container = document.getElementById('moduleContainer');
      const tab = window.appState.activeTab;

      switch (tab) {
        case 'dashboard': renderDashboard(container); break;
        case 'input': renderInputModule(container); break;
        case 'analisis': renderAnalisisModule(container); break;
        case 'akar': renderAkarMasalahModule(container); break;
        case 'prioritas': renderPrioritasModule(container); break;
        case 'rekomendasi': renderRekomendasiModule(container); break;
        case 'rkjm': renderRkjmModule(container); break;
        case 'rkt': renderRktModule(container); break;
        case 'arkas': renderArkasModule(container); break;
        case 'keterhubungan': renderKeterhubunganModule(container); break;
        case 'monev': renderMonevModule(container); break;
        case 'ekspor': renderEksporModule(container); break;
        case 'manajemen': renderManajemenModule(container); break;
        default: renderDashboard(container);
      }
    }

    // =========================================================================
    // MODULE B: DASHBOARD UTAMA
    // =========================================================================
    function renderDashboard(container) {
      const data = getCurrentData();
      const s = data.school;
      const totalInd = data.indicators.length;
      const perluDibenahi = data.indicators.filter(needsImprovement).length;
      const totalPrioritas = data.recommendations.length || perluDibenahi;
      const totalRkt = data.rkt.length;
      const totalAnggaran = data.arkas.reduce((sum, item) => sum + (Number(item.jumlahAnggaran) || 0), 0);

      // Calculate completion status across 8 stages
      const stages = [
        { label: "Identitas Sekolah", done: Boolean(s.nama && s.npsn) },
        { label: "Data Rapor Pendidikan", done: totalInd > 0 },
        { label: "Analisis Akar Masalah", done: Object.keys(data.rootCauses).length > 0 },
        { label: "Penentuan Prioritas", done: perluDibenahi > 0 },
        { label: "Rekomendasi Program", done: data.recommendations.length > 0 },
        { label: "Dokumen RKJM", done: data.rkjm.length > 0 },
        { label: "Dokumen RKT", done: data.rkt.length > 0 },
        { label: "Rancangan ARKAS", done: data.arkas.length > 0 }
      ];
      const completedStages = stages.filter(st => st.done).length;
      const progressPct = Math.round((completedStages / stages.length) * 100);

      container.innerHTML = `
        <!-- Hero School Status Banner -->
        <div class="emerald-gradient rounded-2xl p-6 sm:p-7 text-white shadow-lg relative overflow-hidden">
          <div class="absolute right-0 top-0 w-64 h-64 bg-amber-400/10 rounded-full -translate-y-1/2 translate-x-1/3 pointer-events-none"></div>
          <div class="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-6">
            <div class="space-y-2 max-w-2xl">
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-white/10 border border-amber-400/30 text-amber-300 text-xs font-semibold">
                <i data-lucide="award" class="w-3.5 h-3.5"></i>
                <span>Dashboard Eksekutif Perencanaan Berbasis Data (PBD)</span>
              </div>
              <h2 class="text-xl sm:text-2xl font-extrabold tracking-tight">
                ${s.nama ? escapeHtml(s.nama) : 'Selamat Datang di RAPOR SMART AI'}
              </h2>
              <p class="text-xs sm:text-sm text-emerald-100 leading-relaxed">
                ${s.nama
                  ? `NPSN <strong>${escapeHtml(s.npsn || '-')}</strong> • Jenjang <strong>${escapeHtml(s.jenjang || '-')}</strong> • Tahun Ajaran <strong>${escapeHtml(s.tahunAjaran || '-')}</strong> • Tahun Anggaran <strong>${escapeHtml(s.tahunPerencanaan || window.appState.activeYear)}</strong>`
                  : 'Belum ada data satuan pendidikan yang dimasukkan. Mulai dengan menginput data Rapor Pendidikan (PDF/Excel/Manual) atau muat simulasi demo untuk menjelajahi fitur.'}
              </p>
              <div class="pt-2 flex flex-wrap gap-2">
                <button onclick="navigateTo('input')" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 font-bold text-xs flex items-center gap-1.5 shadow hover:opacity-95 transition">
                  <i data-lucide="file-plus-2" class="w-4 h-4"></i>
                  <span>${totalInd > 0 ? 'Kelola Data Rapor' : 'Input Data Rapor Sekarang'}</span>
                </button>
                ${totalInd > 0 ? `
                <button onclick="navigateTo('analisis')" class="px-4 py-2 rounded-xl bg-white/15 hover:bg-white/25 text-white font-semibold text-xs flex items-center gap-1.5 transition">
                  <i data-lucide="bar-chart-3" class="w-4 h-4 text-amber-300"></i>
                  <span>Mulai Analisis</span>
                </button>` : ''}
              </div>
            </div>

            <!-- Data Completion Gauge Card -->
            <div class="bg-black/25 border border-white/15 rounded-2xl p-4 sm:p-5 w-full lg:w-80 shrink-0">
              <div class="flex items-center justify-between text-xs mb-1.5">
                <span class="text-emerald-200 font-medium">Status Kelengkapan Alur PBD</span>
                <span class="font-extrabold text-amber-300">${progressPct}% (${completedStages}/8)</span>
              </div>
              <div class="w-full h-2.5 bg-emerald-950 rounded-full overflow-hidden mb-3">
                <div class="h-full gold-gradient rounded-full transition-all duration-500" style="width: ${progressPct}%"></div>
              </div>
              <div class="grid grid-cols-2 gap-1.5 text-[11px]">
                ${stages.map(st => `
                  <div class="flex items-center gap-1.5 ${st.done ? 'text-emerald-200' : 'text-white/50'}">
                    <i data-lucide="${st.done ? 'check-circle-2' : 'circle'}" class="w-3.5 h-3.5 ${st.done ? 'text-amber-400' : 'text-white/30'} shrink-0"></i>
                    <span class="truncate">${st.label}</span>
                  </div>
                `).join('')}
              </div>
            </div>
          </div>
        </div>

        <!-- KPI Summary Cards -->
        <div class="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
          <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs hover:border-emerald-300 transition">
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-semibold text-slate-500">Indikator Perlu Dibenahi</span>
              <div class="w-9 h-9 rounded-xl bg-rose-50 text-rose-700 flex items-center justify-center">
                <i data-lucide="alert-circle" class="w-5 h-5"></i>
              </div>
            </div>
            <div class="text-2xl font-extrabold text-slate-900">${perluDibenahi} <span class="text-xs font-medium text-slate-400">/ ${totalInd} indikator</span></div>
            <div class="mt-2 text-xs text-slate-500">Capaian Sedang, Kurang, atau Menurun</div>
          </div>

          <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs hover:border-emerald-300 transition">
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-semibold text-slate-500">Program Prioritas &amp; RKJM</span>
              <div class="w-9 h-9 rounded-xl bg-amber-50 text-amber-700 flex items-center justify-center">
                <i data-lucide="target" class="w-5 h-5"></i>
              </div>
            </div>
            <div class="text-2xl font-extrabold text-slate-900">${data.rkjm.length > 0 ? data.rkjm.length : totalPrioritas} <span class="text-xs font-medium text-slate-400">program</span></div>
            <div class="mt-2 text-xs text-slate-500">${data.recommendations.length} Rekomendasi • ${data.rkjm.length} Masuk RKJM</div>
          </div>

          <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs hover:border-emerald-300 transition">
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-semibold text-slate-500">Kegiatan Tahunan (RKT)</span>
              <div class="w-9 h-9 rounded-xl bg-emerald-50 text-emerald-800 flex items-center justify-center">
                <i data-lucide="calendar-check" class="w-5 h-5"></i>
              </div>
            </div>
            <div class="text-2xl font-extrabold text-slate-900">${totalRkt} <span class="text-xs font-medium text-slate-400">kegiatan</span></div>
            <div class="mt-2 text-xs text-slate-500">Terjadwal di Tahun ${escapeHtml(s.tahunPerencanaan || window.appState.activeYear)}</div>
          </div>

          <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs hover:border-emerald-300 transition">
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-semibold text-slate-500">Total Rancangan ARKAS</span>
              <div class="w-9 h-9 rounded-xl bg-emerald-900 text-amber-300 flex items-center justify-center">
                <i data-lucide="banknote" class="w-5 h-5"></i>
              </div>
            </div>
            <div class="text-xl font-extrabold text-emerald-900 truncate">${formatRupiah(totalAnggaran)}</div>
            <div class="mt-2 text-xs text-slate-500">${data.arkas.length} Rincian Uraian Belanja</div>
          </div>
        </div>

        <!-- Charts & Workflow Pipeline -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <!-- Bar Chart Capaian Indikator -->
          <div class="lg:col-span-2 bg-white rounded-2xl p-5 sm:p-6 border border-slate-200 shadow-xs flex flex-col">
            <div class="flex items-center justify-between mb-4">
              <div>
                <h3 class="font-bold text-slate-900 text-sm sm:text-base">Grafik Capaian Indikator Rapor Pendidikan</h3>
                <p class="text-xs text-slate-500">Visualisasi skor capaian tiap indikator yang telah diinput</p>
              </div>
              <button onclick="navigateTo('analisis')" class="text-xs font-semibold text-emerald-800 hover:underline flex items-center gap-1">
                <span>Detail Analisis</span>
                <i data-lucide="arrow-right" class="w-3.5 h-3.5"></i>
              </button>
            </div>
            ${totalInd > 0 ? `
              <div class="relative flex-1 min-h-[280px]">
                <canvas id="dashboardBarChart"></canvas>
              </div>
            ` : `
              <div class="flex-1 min-h-[260px] rounded-xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center p-6 text-center">
                <i data-lucide="bar-chart-2" class="w-10 h-10 text-slate-300 mb-2"></i>
                <p class="text-sm font-semibold text-slate-600">Belum Ada Grafik Capaian Indikator</p>
                <p class="text-xs text-slate-400 max-w-sm mt-1">Masukkan data Rapor Pendidikan melalui dokumen PDF, Excel, atau formulir manual untuk menampilkan grafik otomatis.</p>
                <button onclick="navigateTo('input')" class="mt-4 px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Input Data Rapor</button>
              </div>
            `}
          </div>

          <!-- Doughnut Komposisi Kategori & Quick Flow -->
          <div class="bg-white rounded-2xl p-5 sm:p-6 border border-slate-200 shadow-xs flex flex-col justify-between">
            <div>
              <h3 class="font-bold text-slate-900 text-sm sm:text-base">Distribusi Kategori Capaian</h3>
              <p class="text-xs text-slate-500 mb-4">Proporsi indikator Baik, Sedang, dan Kurang</p>
              ${totalInd > 0 ? `
                <div class="relative h-52">
                  <canvas id="dashboardDoughnutChart"></canvas>
                </div>
              ` : `
                <div class="h-52 rounded-xl bg-slate-50 border border-slate-200 flex items-center justify-center text-xs text-slate-400 text-center px-4">
                  Menunggu input indikator Rapor Pendidikan
                </div>
              `}
            </div>

            <div class="mt-5 pt-4 border-t border-slate-100">
              <div class="text-xs font-bold text-slate-700 mb-2">Akses Cepat Tahapan:</div>
              <div class="grid grid-cols-2 gap-2">
                <button onclick="navigateTo('akar')" class="px-3 py-2 rounded-xl bg-slate-50 hover:bg-emerald-50 border border-slate-200 text-xs font-semibold text-slate-700 hover:text-emerald-900 text-left flex items-center gap-1.5 transition">
                  <i data-lucide="git-fork" class="w-3.5 h-3.5 text-emerald-700"></i> Akar Masalah
                </button>
                <button onclick="navigateTo('rkjm')" class="px-3 py-2 rounded-xl bg-slate-50 hover:bg-emerald-50 border border-slate-200 text-xs font-semibold text-slate-700 hover:text-emerald-900 text-left flex items-center gap-1.5 transition">
                  <i data-lucide="calendar-range" class="w-3.5 h-3.5 text-emerald-700"></i> Susun RKJM
                </button>
                <button onclick="navigateTo('rkt')" class="px-3 py-2 rounded-xl bg-slate-50 hover:bg-emerald-50 border border-slate-200 text-xs font-semibold text-slate-700 hover:text-emerald-900 text-left flex items-center gap-1.5 transition">
                  <i data-lucide="calendar-check-2" class="w-3.5 h-3.5 text-emerald-700"></i> Susun RKT
                </button>
                <button onclick="navigateTo('arkas')" class="px-3 py-2 rounded-xl bg-slate-50 hover:bg-emerald-50 border border-slate-200 text-xs font-semibold text-slate-700 hover:text-emerald-900 text-left flex items-center gap-1.5 transition">
                  <i data-lucide="wallet" class="w-3.5 h-3.5 text-amber-600"></i> Rancangan ARKAS
                </button>
              </div>
            </div>
          </div>
        </div>
      `;

      // Render Charts if indicators exist
      if (totalInd > 0 && window.Chart) {
        const barCtx = document.getElementById('dashboardBarChart');
        if (barCtx) {
          const labels = data.indicators.map(i => (i.kode ? `${i.kode} ` : '') + (i.nama.length > 22 ? i.nama.slice(0, 22) + '…' : i.nama));
          const scores = data.indicators.map(i => parseFloat(i.nilai) || 0);
          const bgColors = data.indicators.map(i => {
            const kat = classifyCategory(i.nilai, i.kategori);
            if (kat === 'Baik') return '#047857';
            if (kat === 'Sedang') return '#f59e0b';
            return '#e11d48';
          });
          chartInstances.dashBar = new Chart(barCtx, {
            type: 'bar',
            data: {
              labels,
              datasets: [{
                label: 'Nilai Capaian',
                data: scores,
                backgroundColor: bgColors,
                borderRadius: 6,
                maxBarThickness: 36
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: {
                legend: { display: false }
              },
              scales: {
                y: { beginAtZero: true, max: 100, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false }, ticks: { font: { size: 10 } } }
              }
            }
          });
        }

        const doughCtx = document.getElementById('dashboardDoughnutChart');
        if (doughCtx) {
          let baik = 0, sedang = 0, kurang = 0;
          data.indicators.forEach(i => {
            const k = classifyCategory(i.nilai, i.kategori);
            if (k === 'Baik') baik++;
            else if (k === 'Sedang') sedang++;
            else kurang++;
          });
          chartInstances.dashDough = new Chart(doughCtx, {
            type: 'doughnut',
            data: {
              labels: ['Baik', 'Sedang', 'Kurang'],
              datasets: [{
                data: [baik, sedang, kurang],
                backgroundColor: ['#047857', '#f59e0b', '#e11d48'],
                borderWidth: 2,
                borderColor: '#ffffff'
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              cutout: '64%',
              plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }
              }
            }
          });
        }
      }
    }

    // =========================================================================
    // MODULE C: INPUT DATA RAPOR PENDIDIKAN (PDF, EXCEL, & MANUAL FORM)
    // =========================================================================
    function renderInputModule(container) {
      const data = getCurrentData();
      const s = data.school;
      const unverifiedCount = data.indicators.filter(i => !i.terverifikasi).length;

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">1. Input Data Identitas Sekolah &amp; Rapor Pendidikan</h2>
            <p class="text-xs text-slate-500">Masukkan data melalui unggah PDF Rapor Pendidikan, file Excel, atau formulir manual terstruktur.</p>
          </div>
          <div class="flex items-center gap-2">
            <button onclick="navigateTo('analisis')" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <span>Mulai Analisis</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- Section 1: Identitas Satuan Pendidikan -->
        <div class="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-xs">
          <div class="flex items-center justify-between mb-4">
            <h3 class="font-bold text-slate-900 text-sm sm:text-base flex items-center gap-2">
              <i data-lucide="school" class="w-4 h-4 text-emerald-800"></i>
              <span>A. Identitas Satuan Pendidikan</span>
            </h3>
            ${(!s.nama || !s.npsn) ? `<span class="text-[11px] font-semibold px-2.5 py-1 rounded-lg bg-amber-50 text-amber-800 border border-amber-200">Perlu Dilengkapi</span>` : `<span class="text-[11px] font-semibold px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">Terisi</span>`}
          </div>

          <form id="schoolIdentityForm" onsubmit="saveSchoolIdentity(event)" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Nama Sekolah <span class="text-rose-600">*</span></label>
              <input type="text" name="nama" value="${escapeHtml(s.nama)}" placeholder="Contoh: SMP Negeri 1 Nusantara" required class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">NPSN <span class="text-rose-600">*</span></label>
              <input type="text" name="npsn" value="${escapeHtml(s.npsn)}" placeholder="8 digit NPSN" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Jenjang Pendidikan <span class="text-rose-600">*</span></label>
              <select name="jenjang" required class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none">
                <option value="">-- Pilih Jenjang --</option>
                ${["PAUD / TK", "SD / MI", "SMP / MTs", "SMA / MA", "SMK", "SLB", "PKBM / Kesetaraan"].map(j => `<option value="${j}" ${s.jenjang === j ? 'selected' : ''}>${j}</option>`).join('')}
              </select>
            </div>
            <div class="sm:col-span-2 lg:col-span-3">
              <label class="block font-semibold text-slate-700 mb-1">Alamat Lengkap Sekolah</label>
              <input type="text" name="alamat" value="${escapeHtml(s.alamat)}" placeholder="Jalan, Desa/Kelurahan, Kode Pos" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Kecamatan</label>
              <input type="text" name="kecamatan" value="${escapeHtml(s.kecamatan)}" placeholder="Nama Kecamatan" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Kabupaten / Kota</label>
              <input type="text" name="kabupaten" value="${escapeHtml(s.kabupaten)}" placeholder="Nama Kabupaten/Kota" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Provinsi</label>
              <input type="text" name="provinsi" value="${escapeHtml(s.provinsi)}" placeholder="Nama Provinsi" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Tahun Rapor Pendidikan</label>
              <input type="text" name="tahunRapor" value="${escapeHtml(s.tahunRapor)}" placeholder="Contoh: 2025" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Tahun Perencanaan / Anggaran</label>
              <input type="text" name="tahunPerencanaan" value="${escapeHtml(s.tahunPerencanaan || window.appState.activeYear)}" placeholder="Contoh: 2026" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Tahun Ajaran</label>
              <input type="text" name="tahunAjaran" value="${escapeHtml(s.tahunAjaran)}" placeholder="Contoh: 2026/2027" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>

            <div class="sm:col-span-2 lg:col-span-3 flex justify-end pt-2">
              <button type="submit" class="px-5 py-2.5 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
                <i data-lucide="check" class="w-4 h-4"></i>
                <span>Simpan Identitas Sekolah</span>
              </button>
            </div>
          </form>
        </div>

        <!-- Section 2: Tiga Metode Input Data Rapor Pendidikan -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-5">
          <!-- Method 1: Upload PDF -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-xl bg-rose-50 text-rose-700 flex items-center justify-center mb-3">
                <i data-lucide="file-text" class="w-5 h-5"></i>
              </div>
              <h4 class="font-bold text-slate-900 text-sm">1. Unggah Dokumen PDF Rapor</h4>
              <p class="text-xs text-slate-500 mt-1 leading-relaxed">
                Sistem akan membaca teks PDF Rapor Pendidikan secara otomatis. Bagian yang tidak terbaca akan ditandai untuk Anda konfirmasi (tanpa mengarang data).
              </p>
            </div>
            <div class="mt-4 space-y-2">
              <label class="block w-full cursor-pointer rounded-xl border-2 border-dashed border-slate-300 hover:border-emerald-700 bg-slate-50 hover:bg-emerald-50/40 p-3 text-center transition">
                <span class="text-xs font-semibold text-emerald-900 flex items-center justify-center gap-1.5">
                  <i data-lucide="upload" class="w-4 h-4"></i> Pilih File PDF (.pdf)
                </span>
                <input type="file" accept=".pdf" onchange="handlePdfUpload(event)" class="hidden" />
              </label>
            </div>
          </div>

          <!-- Method 2: Upload Excel -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-800 flex items-center justify-center mb-3">
                <i data-lucide="file-spreadsheet" class="w-5 h-5"></i>
              </div>
              <h4 class="font-bold text-slate-900 text-sm">2. Unggah Langsung Excel Rapor Pendidikan</h4>
              <p class="text-xs text-slate-500 mt-1 leading-relaxed">
                Bisa langsung mengunggah file unduhan resmi Rapor Pendidikan (.xlsx / .xls / .csv) tanpa ubah format, maupun menggunakan template kami.
              </p>
            </div>
            <div class="mt-4 space-y-2">
              <label class="block w-full cursor-pointer rounded-xl border-2 border-dashed border-emerald-400 hover:border-emerald-700 bg-emerald-50/40 hover:bg-emerald-50 p-3 text-center transition">
                <span class="text-xs font-bold text-emerald-900 flex items-center justify-center gap-1.5">
                  <i data-lucide="upload" class="w-4 h-4"></i> Unggah File Excel Rapor (.xlsx/.xls)
                </span>
                <input type="file" accept=".xlsx,.xls,.csv" onchange="handleExcelUpload(event)" class="hidden" />
              </label>
              <button onclick="downloadRaporExcelTemplate()" class="w-full py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-semibold flex items-center justify-center gap-1 transition">
                <i data-lucide="download" class="w-3.5 h-3.5"></i> Unduh Template Alternatif
              </button>
            </div>
          </div>

          <!-- Method 3: Input Manual & Katalog Indikator -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-xl bg-amber-50 text-amber-700 flex items-center justify-center mb-3">
                <i data-lucide="clipboard-edit" class="w-5 h-5"></i>
              </div>
              <h4 class="font-bold text-slate-900 text-sm">3. Input Manual &amp; Katalog PBD</h4>
              <p class="text-xs text-slate-500 mt-1 leading-relaxed">
                Pilih indikator standar Rapor Pendidikan atau tambahkan indikator secara manual sesuai lembar hasil Rapor Pendidikan sekolah Anda.
              </p>
            </div>
            <div class="mt-4 space-y-2">
              <button onclick="openBulkCatalogPickerModal()" class="w-full py-2.5 px-3 rounded-xl emerald-gradient hover:opacity-95 text-white text-xs font-bold flex items-center justify-center gap-1.5 shadow-2xs transition">
                <i data-lucide="list-checks" class="w-4 h-4 text-amber-300"></i> Pilih dari Katalog Indikator
              </button>
              <button onclick="openEditIndicatorModal()" class="w-full py-2 px-3 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-semibold flex items-center justify-center gap-1.5 transition">
                <i data-lucide="plus" class="w-3.5 h-3.5"></i> Tambah Indikator Manual
              </button>
            </div>
          </div>
        </div>

        <!-- Section 3: Tabel Indikator Rapor Pendidikan yang Diinput -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          <div class="p-4 sm:p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
            <div>
              <h3 class="font-bold text-slate-900 text-sm sm:text-base">B. Daftar Indikator Rapor Pendidikan (${data.indicators.length})</h3>
              <p class="text-xs text-slate-500">Pastikan nilai capaian dan kategori setiap indikator telah sesuai dengan dokumen resmi sekolah.</p>
            </div>
            ${unverifiedCount > 0 ? `
              <span class="text-xs font-bold px-3 py-1 rounded-full bg-amber-100 text-amber-900 border border-amber-300">
                ${unverifiedCount} Indikator Perlu Konfirmasi
              </span>
            ` : ''}
          </div>

          ${data.indicators.length > 0 ? `
            <div class="overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase text-[10px]">
                    <th class="py-3 px-4">Indikator &amp; Dimensi</th>
                    <th class="py-3 px-3 text-center">Nilai</th>
                    <th class="py-3 px-3 text-center">Kategori</th>
                    <th class="py-3 px-3 text-center">Perubahan</th>
                    <th class="py-3 px-3">Peringkat / Rekomendasi Rapor</th>
                    <th class="py-3 px-3 text-center">Verifikasi</th>
                    <th class="py-3 px-4 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100 align-top">
                  ${data.indicators.map(ind => {
                    const kat = classifyCategory(ind.nilai, ind.kategori);
                    return `
                      <tr class="hover:bg-slate-50 ${!ind.terverifikasi ? 'bg-amber-50/40' : ''}">
                        <td class="py-3 px-4 min-w-[200px]">
                          <div class="font-bold text-slate-900">${ind.kode ? `<span class="font-mono text-emerald-800 mr-1">[${escapeHtml(ind.kode)}]</span>` : ''}${escapeHtml(ind.nama)}</div>
                          <div class="text-[11px] text-slate-500 mt-0.5">${escapeHtml(ind.dimensi || '')}</div>
                        </td>
                        <td class="py-3 px-3 text-center font-mono font-bold text-sm">
                          ${ind.nilai !== '' ? escapeHtml(ind.nilai) : '<span class="text-rose-600 text-xs font-sans">Kosong</span>'}
                        </td>
                        <td class="py-3 px-3 text-center">${getCategoryBadgeHtml(kat)}</td>
                        <td class="py-3 px-3 text-center">${getDeltaBadgeHtml(ind.perubahan)}</td>
                        <td class="py-3 px-3 min-w-[180px] text-slate-600">
                          ${ind.peringkat ? `<div class="font-semibold text-slate-800 text-[11px]">${escapeHtml(ind.peringkat)}</div>` : ''}
                          <div class="text-[11px] line-clamp-2">${escapeHtml(ind.rekomendasiRapor || '-')}</div>
                        </td>
                        <td class="py-3 px-3 text-center whitespace-nowrap">
                          ${ind.terverifikasi
                            ? `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-emerald-50 text-emerald-800 border border-emerald-200"><i data-lucide="check-circle-2" class="w-3.5 h-3.5"></i> Terkonfirmasi</span>`
                            : `<button onclick="verifyIndicatorRow('${ind.id}')" class="px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-emerald-950 font-bold text-[11px] transition">Konfirmasi ✓</button>`}
                        </td>
                        <td class="py-3 px-4 text-right whitespace-nowrap">
                          <button onclick="openEditIndicatorModal('${ind.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit Indikator"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteIndicatorRow('${ind.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          ` : `
            <div class="p-10 text-center">
              <i data-lucide="file-spreadsheet" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
              <h4 class="font-bold text-slate-900 text-sm">Belum Ada Indikator Rapor Pendidikan yang Diinput</h4>
              <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Silakan unggah file Excel/PDF Rapor Pendidikan Anda atau pilih indikator dari Katalog PBD di atas.</p>
            </div>
          `}
        </div>
      `;
    }

    function saveSchoolIdentity(event) {
      event.preventDefault();
      const fd = new FormData(event.target);
      const data = getCurrentData();
      data.school = {
        ...data.school,
        nama: String(fd.get('nama') || '').trim(),
        npsn: String(fd.get('npsn') || '').trim(),
        jenjang: String(fd.get('jenjang') || '').trim(),
        alamat: String(fd.get('alamat') || '').trim(),
        kecamatan: String(fd.get('kecamatan') || '').trim(),
        kabupaten: String(fd.get('kabupaten') || '').trim(),
        provinsi: String(fd.get('provinsi') || '').trim(),
        tahunRapor: String(fd.get('tahunRapor') || '').trim(),
        tahunPerencanaan: String(fd.get('tahunPerencanaan') || window.appState.activeYear).trim(),
        tahunAjaran: String(fd.get('tahunAjaran') || '').trim()
      };
      triggerAutoSave("Input Data", "Simpan Identitas Sekolah", `Menyimpan identitas ${data.school.nama} (NPSN: ${data.school.npsn}).`);
      renderAll();
      showToast("Identitas satuan pendidikan berhasil disimpan.", "success");
    }

    function openBulkCatalogPickerModal() {
      const data = getCurrentData();
      const existingCodes = new Set(data.indicators.map(i => i.kode));

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="font-bold text-slate-900 text-base">Pilih Indikator Standar Rapor Pendidikan</h3>
              <p class="text-xs text-slate-500">Centang indikator yang ingin ditambahkan ke tabel input, lalu isi nilai capaiannya.</p>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveBulkCatalogSelection(event)" class="mt-4 space-y-2 text-xs">
            <div class="max-h-72 overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-xl">
              ${KATALOG_INDIKATOR_STANDAR.map(std => {
                const already = existingCodes.has(std.kode);
                return `
                  <label class="flex items-start gap-3 p-3 hover:bg-slate-50 cursor-pointer ${already ? 'opacity-60 bg-slate-50' : ''}">
                    <input type="checkbox" name="selectedCodes" value="${std.kode}" ${already ? 'disabled checked' : ''} class="mt-0.5 rounded text-emerald-700" />
                    <div class="flex-1">
                      <div class="font-bold text-slate-900">[${std.kode}] ${escapeHtml(std.nama)} ${already ? '<span class="text-[10px] text-emerald-700 font-semibold">(Sudah ada)</span>' : ''}</div>
                      <div class="text-[11px] text-slate-500">${escapeHtml(std.dimensi)}</div>
                    </div>
                  </label>
                `;
              }).join('')}
            </div>
            <div class="pt-3 flex justify-end gap-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Tambahkan Indikator Terpilih</button>
            </div>
          </form>
        </div>
      `, 'max-w-2xl');
    }

    function saveBulkCatalogSelection(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const selected = fd.getAll('selectedCodes');
      const data = getCurrentData();
      let added = 0;

      selected.forEach(kode => {
        const std = KATALOG_INDIKATOR_STANDAR.find(s => s.kode === kode);
        if (std && !data.indicators.some(i => i.kode === kode)) {
          data.indicators.push({
            id: uid(),
            kode: std.kode,
            nama: std.nama,
            dimensi: std.dimensi,
            nilai: "",
            kategori: "Belum Terverifikasi",
            perubahan: "",
            peringkat: "",
            rekomendasiRapor: "",
            terverifikasi: false
          });
          added++;
        }
      });

      closeCustomModal();
      if (added > 0) {
        triggerAutoSave("Input Data", "Tambah dari Katalog", `Menambahkan ${added} indikator standar PBD.`);
        renderAll();
        showToast(`${added} indikator ditambahkan. Silakan klik ikon pensil untuk mengisi nilai capaiannya.`, "success");
      }
    }

    function openEditIndicatorModal(indicatorId = null) {
      const data = getCurrentData();
      const existing = indicatorId ? data.indicators.find(i => i.id === indicatorId) : null;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">${existing ? 'Edit Indikator Rapor Pendidikan' : 'Tambah Indikator Rapor Pendidikan'}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveIndicatorModal(event, '${indicatorId || ''}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kode Indikator</label>
                <input type="text" name="kode" value="${escapeHtml(existing?.kode || '')}" placeholder="Contoh: A.1" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold text-slate-700 mb-1">Nama Indikator <span class="text-rose-600">*</span></label>
                <input type="text" name="nama" value="${escapeHtml(existing?.nama || '')}" required placeholder="Contoh: Kemampuan Literasi" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Dimensi Rapor Pendidikan</label>
              <select name="dimensi" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${DIMENSI_RAPOR.map(d => `<option value="${d}" ${existing?.dimensi === d ? 'selected' : ''}>${d}</option>`).join('')}
              </select>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Nilai Capaian (0–100) <span class="text-rose-600">*</span></label>
                <input type="number" step="0.01" min="0" max="100" name="nilai" value="${existing?.nilai !== undefined ? existing.nilai : ''}" required placeholder="Contoh: 64.50" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kategori Capaian</label>
                <select name="kategori" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  <option value="">Otomatis dari Nilai</option>
                  <option value="Baik" ${existing?.kategori === 'Baik' ? 'selected' : ''}>Baik</option>
                  <option value="Sedang" ${existing?.kategori === 'Sedang' ? 'selected' : ''}>Sedang</option>
                  <option value="Kurang" ${existing?.kategori === 'Kurang' ? 'selected' : ''}>Kurang</option>
                </select>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Perubahan Tahun Lalu</label>
                <input type="number" step="0.01" name="perubahan" value="${existing?.perubahan !== undefined ? existing.perubahan : ''}" placeholder="Contoh: -3.20 atau +4.10" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Peringkat / Posisi (Opsional)</label>
              <input type="text" name="peringkat" value="${escapeHtml(existing?.peringkat || '')}" placeholder="Contoh: Peringkat Menengah (41-60%) Kab/Kota" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Rekomendasi dari Rapor Pendidikan</label>
              <textarea name="rekomendasiRapor" rows="2" placeholder="Salin catatan benahi/rekomendasi dari Rapor Pendidikan jika ada..." class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(existing?.rekomendasiRapor || '')}</textarea>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan &amp; Konfirmasi</button>
            </div>
          </form>
        </div>
      `, 'max-w-2xl');
    }

    function saveIndicatorModal(e, indicatorId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const nilaiRaw = fd.get('nilai');
      const nilaiNum = nilaiRaw !== '' ? parseFloat(nilaiRaw) : '';
      const katRaw = fd.get('kategori');
      const perubRaw = fd.get('perubahan');
      const perubNum = perubRaw !== '' ? parseFloat(perubRaw) : '';

      const obj = {
        id: indicatorId || uid(),
        kode: String(fd.get('kode') || '').trim().toUpperCase(),
        nama: String(fd.get('nama') || '').trim(),
        dimensi: String(fd.get('dimensi') || DIMENSI_RAPOR[0]),
        nilai: nilaiNum,
        kategori: katRaw || classifyCategory(nilaiNum),
        perubahan: perubNum,
        peringkat: String(fd.get('peringkat') || '').trim(),
        rekomendasiRapor: String(fd.get('rekomendasiRapor') || '').trim(),
        terverifikasi: nilaiNum !== ''
      };

      if (indicatorId) {
        const idx = data.indicators.findIndex(i => i.id === indicatorId);
        if (idx !== -1) data.indicators[idx] = obj;
      } else {
        data.indicators.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("Input Data", "Simpan Indikator", `Menyimpan indikator "${obj.nama}" (Nilai: ${obj.nilai}).`);
      renderAll();
      showToast("Data indikator berhasil disimpan & dikonfirmasi.", "success");
    }

    function verifyIndicatorRow(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;
      if (ind.nilai === '' || isNaN(parseFloat(ind.nilai))) {
        openEditIndicatorModal(indicatorId);
        showToast("Nilai capaian masih kosong. Silakan isi nilai terlebih dahulu sebelum konfirmasi.", "warning");
        return;
      }
      ind.terverifikasi = true;
      ind.kategori = classifyCategory(ind.nilai, ind.kategori);
      triggerAutoSave("Input Data", "Konfirmasi Indikator", `Mengonfirmasi keakuratan indikator "${ind.nama}".`);
      renderAll();
      showToast(`Indikator "${ind.nama}" telah dikonfirmasi.`, "success");
    }

    function deleteIndicatorRow(indicatorId) {
      showConfirmDialog({
        title: "Hapus Indikator Rapor?",
        message: "Indikator ini akan dihapus dari daftar Rapor Pendidikan tahun aktif.",
        confirmText: "Ya, Hapus",
        danger: true,
        onConfirm: () => {
          const data = getCurrentData();
          data.indicators = data.indicators.filter(i => i.id !== indicatorId);
          delete data.rootCauses[indicatorId];
          delete data.priorities[indicatorId];
          triggerAutoSave("Input Data", "Hapus Indikator", "Menghapus indikator dari tabel Rapor Pendidikan.");
          renderAll();
          showToast("Indikator dihapus.", "info");
        }
      });
    }

    // Real Excel Upload Handler (Supports Official Kemendikdasmen Rapor Pendidikan Multi-Sheet Excel & Custom Templates)
    window._pendingExcelWorkbook = null;
    window._pendingExcelFileName = "";
    window._pendingExcelSchoolMeta = {};

    function normalizeOfficialCategoryLabel(rawLabel, numericScore) {
      const txt = String(rawLabel || "").toLowerCase().trim();
      if (txt) {
        if (txt.includes("jauh di bawah") || txt.includes("di bawah") || txt.includes("kurang") || txt.includes("perlu intervensi")) {
          return "Kurang";
        }
        if (txt.includes("sedang") || txt.includes("cukup")) {
          return "Sedang";
        }
        if (txt.includes("baik") || txt.includes("di atas") || txt.includes("mencapai") || txt.includes("membudaya")) {
          return "Baik";
        }
      }
      return classifyCategory(numericScore, rawLabel);
    }

    function parseIndonesianDeltaValue(rawDelta) {
      if (rawDelta === null || rawDelta === undefined || rawDelta === "") return "";
      if (typeof rawDelta === "number") return isNaN(rawDelta) ? "" : rawDelta;
      const str = String(rawDelta).trim();
      const lower = str.toLowerCase();
      if (lower === "-" || lower === "tetap" || lower.includes("tidak berubah")) return 0;
      const numMatch = str.replace(/,/g, '.').match(/[+-]?\d+(?:\.\d+)?/);
      if (!numMatch) return "";
      let val = parseFloat(numMatch[0]);
      if (isNaN(val)) return "";
      if (lower.includes("turun")) val = -Math.abs(val);
      else if (lower.includes("naik")) val = Math.abs(val);
      return val;
    }

    function parseIndonesianScoreValue(rawScore) {
      if (rawScore === null || rawScore === undefined || rawScore === "") return NaN;
      if (typeof rawScore === "number") return rawScore;
      const cleaned = String(rawScore).replace(/%/g, '').replace(/,/g, '.').trim();
      const match = cleaned.match(/^-?\d+(?:\.\d+)?$/);
      if (!match) return NaN;
      return parseFloat(match[0]);
    }

    function extractSchoolMetadataFromWorkbook(workbook) {
      const meta = {};
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const maxScan = Math.min(rows.length, 25);
        for (let r = 0; r < maxScan; r++) {
          const row = rows[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || "").trim();
            const lower = cell.toLowerCase();
            // Look for key-value in same cell ("NPSN: 20301234") or adjacent cell
            const nextVal = String(row[c + 1] || row[c + 2] || "").trim();

            if (!meta.npsn) {
              const inlineNpsn = cell.match(/npsn\s*[:\-]?\s*(\d{8})/i);
              if (inlineNpsn) meta.npsn = inlineNpsn[1];
              else if (lower === "npsn" || lower.startsWith("npsn")) {
                const adjNpsn = nextVal.match(/(\d{8})/);
                if (adjNpsn) meta.npsn = adjNpsn[1];
              }
            }
            if (!meta.nama && (lower === "nama satuan pendidikan" || lower === "nama sekolah" || lower === "satuan pendidikan")) {
              if (nextVal && nextVal.length > 2 && !nextVal.toLowerCase().includes("npsn")) {
                meta.nama = nextVal.replace(/^[:\-]\s*/, "");
              }
            }
            if (!meta.kecamatan && (lower === "kecamatan" || lower.startsWith("kecamatan"))) {
              const val = cell.includes(":") ? cell.split(":")[1].trim() : nextVal.replace(/^[:\-]\s*/, "");
              if (val) meta.kecamatan = val;
            }
            if (!meta.kabupaten && (lower.includes("kabupaten") || lower.includes("kab/kota") || lower.includes("kota"))) {
              const val = cell.includes(":") ? cell.split(":")[1].trim() : nextVal.replace(/^[:\-]\s*/, "");
              if (val && val.length > 2) meta.kabupaten = val;
            }
            if (!meta.provinsi && lower.startsWith("provinsi")) {
              const val = cell.includes(":") ? cell.split(":")[1].trim() : nextVal.replace(/^[:\-]\s*/, "");
              if (val && val.length > 2) meta.provinsi = val;
            }
          }
        }
      }
      return meta;
    }

    function parseExcelSheetForRapor(sheet) {
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      if (!rows || rows.length === 0) return { headerRowIndex: -1, items: [] };

      // 1. Dynamically locate the header row (scans first 35 rows)
      let headerRowIndex = -1;
      let colMap = {
        kode: -1,
        nama: -1,
        dimensi: -1,
        nilai: -1,
        kategori: -1,
        perubahan: -1,
        peringkat: -1,
        rekomendasi: -1
      };

      const maxHeaderScan = Math.min(rows.length, 35);
      for (let r = 0; r < maxHeaderScan; r++) {
        const row = rows[r] || [];
        const normalizedCells = row.map(c => String(c || "").toLowerCase().replace(/\s+/g, " ").trim());

        const namaIdx = normalizedCells.findIndex(t =>
          t === "indikator" ||
          t === "nama indikator" ||
          t.includes("nama indikator") ||
          (t.includes("indikator") && !t.includes("kode") && !t.includes("no") && !t.includes("id")) ||
          t === "sub indikator" ||
          t === "aspek"
        );

        const hasMetricCol = normalizedCells.some(t =>
          t.includes("skor") || t.includes("nilai") || t.includes("capaian") || t.includes("label") || t.includes("kategori")
        );

        if (namaIdx !== -1 && hasMetricCol) {
          headerRowIndex = r;
          colMap.nama = namaIdx;

          colMap.kode = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t === "kode" || t.includes("kode indikator") || t === "no indikator" || t === "id indikator" || t === "no")
          );
          colMap.dimensi = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t.includes("dimensi") || t.includes("standar"))
          );
          // Match score column while avoiding "perubahan skor"
          colMap.nilai = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx &&
            !t.includes("perubahan") &&
            !t.includes("selisih") &&
            !t.includes("tahun lalu") &&
            !t.includes("label") &&
            (t.includes("skor") || t.includes("nilai") || t === "capaian" || t.includes("capaian sekolah") || t.includes("indeks"))
          );
          colMap.kategori = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t.includes("label") || t.includes("kategori") || t.includes("predikat") || t.includes("status capaian"))
          );
          colMap.perubahan = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t.includes("perubahan") || t.includes("delta") || t.includes("selisih") || t.includes("kenaikan") || t.includes("tren"))
          );
          colMap.peringkat = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t.includes("peringkat") || t.includes("posisi") || t.includes("kuintil") || t.includes("persentil"))
          );
          colMap.rekomendasi = normalizedCells.findIndex((t, idx) =>
            idx !== namaIdx && (t.includes("rekomendasi") || t.includes("benahi") || t.includes("inspirasi") || t.includes("kegiatan"))
          );
          break;
        }
      }

      const items = [];
      let currentDimensiContext = DIMENSI_RAPOR[0];

      if (headerRowIndex !== -1) {
        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawNama = String(row[colMap.nama] || "").trim();
          const rawKode = colMap.kode !== -1 ? String(row[colMap.kode] || "").trim() : "";

          if (!rawNama) continue;
          // Skip repeated headers
          if (rawNama.toLowerCase() === "indikator" || rawNama.toLowerCase() === "nama indikator") continue;

          // Detect if row is a Dimension section header (e.g. "Dimensi A: Mutu dan Relevansi...")
          if (rawNama.toLowerCase().startsWith("dimensi ")) {
            currentDimensiContext = rawNama;
            continue;
          }

          const rawNilai = colMap.nilai !== -1 ? row[colMap.nilai] : "";
          const rawKat = colMap.kategori !== -1 ? String(row[colMap.kategori] || "").trim() : "";
          const rawDelta = colMap.perubahan !== -1 ? row[colMap.perubahan] : "";
          const rawPeringkat = colMap.peringkat !== -1 ? String(row[colMap.peringkat] || "").trim() : "";
          const rawRekom = colMap.rekomendasi !== -1 ? String(row[colMap.rekomendasi] || "").trim() : "";
          const rawDim = colMap.dimensi !== -1 ? String(row[colMap.dimensi] || "").trim() : "";

          const nilaiNum = parseIndonesianScoreValue(rawNilai);
          const hasValidScore = !isNaN(nilaiNum) && nilaiNum >= 0 && nilaiNum <= 100;
          const deltaNum = parseIndonesianDeltaValue(rawDelta);

          // Skip pure group title rows that have no score, no category, no delta, and no recommendation
          if (!hasValidScore && !rawKat && deltaNum === "" && !rawRekom && !rawKode) {
            continue;
          }

          // Extract inline code if embedded in indicator name e.g. "A.1 Kemampuan Literasi"
          let finalKode = rawKode;
          let finalNama = rawNama;
          const inlineCodeMatch = rawNama.match(/^([A-E]\.\d+(?:\.\d+)?)\s+(.+)$/i);
          if (inlineCodeMatch) {
            if (!finalKode) finalKode = inlineCodeMatch[1].toUpperCase();
            finalNama = inlineCodeMatch[2].trim();
          }

          const matchedStd = KATALOG_INDIKATOR_STANDAR.find(s =>
            (finalKode && s.kode.toLowerCase() === finalKode.toLowerCase()) ||
            s.nama.toLowerCase() === finalNama.toLowerCase()
          );

          const finalKat = hasValidScore
            ? normalizeOfficialCategoryLabel(rawKat, nilaiNum)
            : (rawKat ? normalizeOfficialCategoryLabel(rawKat, NaN) : "Belum Terverifikasi");

          items.push({
            id: uid(),
            kode: finalKode || matchedStd?.kode || "",
            nama: finalNama,
            dimensi: rawDim || matchedStd?.dimensi || currentDimensiContext,
            nilai: hasValidScore ? nilaiNum : "",
            kategori: finalKat,
            perubahan: deltaNum,
            peringkat: rawPeringkat,
            rekomendasiRapor: rawRekom,
            terverifikasi: hasValidScore
          });
        }
      }

      // 2. Fallback scan if standard header row wasn't found (matches official indicator names directly)
      if (items.length === 0) {
        for (let r = 0; r < rows.length; r++) {
          const row = rows[r] || [];
          const rowTextCells = row.map(c => String(c || "").trim());
          for (const std of KATALOG_INDIKATOR_STANDAR) {
            const matchIdx = rowTextCells.findIndex(cell =>
              cell.toLowerCase() === std.nama.toLowerCase() ||
              cell.toLowerCase().includes(`${std.kode.toLowerCase()} ${std.nama.toLowerCase()}`)
            );
            if (matchIdx !== -1 && !items.some(it => it.kode === std.kode)) {
              // Look to the right cells for numeric score and category label
              let foundScore = NaN;
              let foundKat = "";
              for (let c = matchIdx + 1; c < Math.min(row.length, matchIdx + 6); c++) {
                const candidateNum = parseIndonesianScoreValue(row[c]);
                if (isNaN(foundScore) && !isNaN(candidateNum) && candidateNum >= 0 && candidateNum <= 100) {
                  foundScore = candidateNum;
                } else if (!foundKat && typeof row[c] === "string" && row[c].trim().length > 2) {
                  const lbl = row[c].trim().toLowerCase();
                  if (lbl.includes("baik") || lbl.includes("sedang") || lbl.includes("kurang") || lbl.includes("kompetensi")) {
                    foundKat = row[c].trim();
                  }
                }
              }
              const valid = !isNaN(foundScore);
              items.push({
                id: uid(),
                kode: std.kode,
                nama: std.nama,
                dimensi: std.dimensi,
                nilai: valid ? foundScore : "",
                kategori: valid ? normalizeOfficialCategoryLabel(foundKat, foundScore) : "Belum Terverifikasi",
                perubahan: "",
                peringkat: "",
                rekomendasiRapor: "",
                terverifikasi: valid
              });
            }
          }
        }
      }

      return { headerRowIndex, items };
    }

    function handleExcelUpload(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      if (!window.XLSX) {
        showToast("Pustaka pembaca Excel belum siap. Silakan coba lagi.", "error");
        return;
      }

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const workbook = XLSX.read(e.target.result, { type: 'array' });
          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            showToast("File Excel tidak memiliki lembar kerja (sheet) yang valid.", "warning");
            return;
          }

          window._pendingExcelWorkbook = workbook;
          window._pendingExcelFileName = file.name;
          window._pendingExcelSchoolMeta = extractSchoolMetadataFromWorkbook(workbook);

          // Score all sheets and pick the sheet with the most Rapor Pendidikan indicators (skipping 'Panduan')
          let bestSheet = workbook.SheetNames[0];
          let bestCount = -1;

          workbook.SheetNames.forEach(name => {
            const parsed = parseExcelSheetForRapor(workbook.Sheets[name]);
            const isPanduan = name.toLowerCase().includes("panduan") || name.toLowerCase().includes("petunjuk");
            const effectiveScore = isPanduan ? parsed.items.length - 100 : parsed.items.length;
            if (effectiveScore > bestCount) {
              bestCount = effectiveScore;
              bestSheet = name;
            }
          });

          openExcelImportPreviewModal(bestSheet);
        } catch (err) {
          console.error(err);
          showToast("Gagal membaca file Excel. Pastikan format .xlsx / .xls / .csv tidak rusak.", "error");
        }
        event.target.value = "";
      };
      reader.readAsArrayBuffer(file);
    }

    function openExcelImportPreviewModal(selectedSheetName) {
      const wb = window._pendingExcelWorkbook;
      if (!wb) return;

      const sheet = wb.Sheets[selectedSheetName];
      const parsed = parseExcelSheetForRapor(sheet);
      const items = parsed.items;
      const meta = window._pendingExcelSchoolMeta || {};
      const hasMeta = Boolean(meta.nama || meta.npsn || meta.kabupaten || meta.provinsi);

      const sheetOptionsHtml = wb.SheetNames.map(name => {
        const count = parseExcelSheetForRapor(wb.Sheets[name]).items.length;
        return `<option value="${escapeHtml(name)}" ${name === selectedSheetName ? 'selected' : ''}>${escapeHtml(name)} (${count} indikator terdeteksi)</option>`;
      }).join('');

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[90vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="font-bold text-slate-900 text-base">Pratinjau &amp; Konfirmasi Impor Excel Rapor Pendidikan</h3>
              <p class="text-xs text-slate-500">File: <strong>${escapeHtml(window._pendingExcelFileName)}</strong></p>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>

          <div class="mt-4 space-y-4 text-xs">
            <!-- Sheet Selector -->
            <div class="bg-slate-50 border border-slate-200 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <div class="font-bold text-slate-800">Pilih Lembar Kerja (Sheet) Data Rapor:</div>
                <div class="text-[11px] text-slate-500">Sistem otomatis melewati sheet "Panduan" dan memilih sheet dengan indikator terbanyak.</div>
              </div>
              <select onchange="openExcelImportPreviewModal(this.value)" class="rounded-xl border border-slate-300 bg-white px-3 py-2 font-semibold text-slate-900 sm:min-w-[260px]">
                ${sheetOptionsHtml}
              </select>
            </div>

            <!-- Detected School Metadata -->
            ${hasMeta ? `
              <div class="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
                <label class="flex items-start gap-2.5 cursor-pointer">
                  <input type="checkbox" id="applyDetectedSchoolMeta" checked class="mt-0.5 rounded text-emerald-700" />
                  <div>
                    <div class="font-bold text-emerald-950">Terapkan Identitas Sekolah yang Terdeteksi dari Excel:</div>
                    <div class="text-emerald-900 mt-0.5">
                      ${meta.nama ? `Nama: <strong>${escapeHtml(meta.nama)}</strong> • ` : ''}
                      ${meta.npsn ? `NPSN: <strong>${escapeHtml(meta.npsn)}</strong> • ` : ''}
                      ${meta.kabupaten ? `${escapeHtml(meta.kabupaten)}` : ''}
                    </div>
                  </div>
                </label>
              </div>
            ` : ''}

            <!-- Preview Table -->
            ${items.length > 0 ? `
              <div>
                <div class="flex items-center justify-between mb-2">
                  <span class="font-bold text-slate-800">Ditemukan ${items.length} Indikator Rapor Pendidikan:</span>
                  <div class="flex items-center gap-3">
                    <label class="inline-flex items-center gap-1.5 cursor-pointer">
                      <input type="radio" name="excelImportMode" value="append" checked class="text-emerald-700" />
                      <span>Tambahkan</span>
                    </label>
                    <label class="inline-flex items-center gap-1.5 cursor-pointer">
                      <input type="radio" name="excelImportMode" value="replace" class="text-emerald-700" />
                      <span>Ganti Data Lama</span>
                    </label>
                  </div>
                </div>
                <div class="max-h-64 overflow-y-auto border border-slate-200 rounded-xl">
                  <table class="w-full text-left border-collapse text-[11px]">
                    <thead class="bg-slate-100 text-slate-700 font-bold sticky top-0">
                      <tr>
                        <th class="py-2 px-3">Kode</th>
                        <th class="py-2 px-3">Nama Indikator</th>
                        <th class="py-2 px-3 text-center">Nilai</th>
                        <th class="py-2 px-3 text-center">Kategori</th>
                        <th class="py-2 px-3 text-center">Perubahan</th>
                      </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                      ${items.map(it => `
                        <tr>
                          <td class="py-2 px-3 font-mono font-bold text-emerald-800">${escapeHtml(it.kode || '-')}</td>
                          <td class="py-2 px-3 font-semibold text-slate-900">${escapeHtml(it.nama)}</td>
                          <td class="py-2 px-3 text-center font-mono font-bold">${it.nilai !== '' ? it.nilai : '<span class="text-rose-600">Kosong</span>'}</td>
                          <td class="py-2 px-3 text-center">${escapeHtml(it.kategori)}</td>
                          <td class="py-2 px-3 text-center font-mono">${it.perubahan !== '' ? it.perubahan : '-'}</td>
                        </tr>
                      `).join('')}
                    </tbody>
                  </table>
                </div>
              </div>
            ` : `
              <div class="p-6 rounded-xl bg-amber-50 border border-amber-200 text-center text-amber-950">
                <p class="font-bold">Tidak ditemukan tabel indikator pada sheet "${escapeHtml(selectedSheetName)}".</p>
                <p class="mt-1">Silakan pilih nama sheet lain pada pilihan di atas (misalnya sheet <strong>Laporan Rapor Pendidikan</strong> atau <strong>Capaian</strong>).</p>
              </div>
            `}
          </div>

          <div class="mt-5 pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
            <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 text-slate-700 text-xs font-semibold">Batal</button>
            ${items.length > 0 ? `
              <button data-sheet="${escapeHtml(selectedSheetName)}" onclick="confirmExcelImport(this.getAttribute('data-sheet'))" class="px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold shadow-sm">
                Konfirmasi &amp; Impor (${items.length} Indikator)
              </button>
            ` : ''}
          </div>
        </div>
      `, 'max-w-3xl');
    }

    function confirmExcelImport(sheetName) {
      const wb = window._pendingExcelWorkbook;
      if (!wb || !wb.Sheets[sheetName]) return;

      const parsed = parseExcelSheetForRapor(wb.Sheets[sheetName]);
      const data = getCurrentData();

      // Apply school metadata if checked
      const applyMeta = document.getElementById('applyDetectedSchoolMeta')?.checked;
      const meta = window._pendingExcelSchoolMeta || {};
      if (applyMeta) {
        if (meta.nama && !data.school.nama) data.school.nama = meta.nama;
        if (meta.npsn && !data.school.npsn) data.school.npsn = meta.npsn;
        if (meta.kecamatan && !data.school.kecamatan) data.school.kecamatan = meta.kecamatan;
        if (meta.kabupaten && !data.school.kabupaten) data.school.kabupaten = meta.kabupaten;
        if (meta.provinsi && !data.school.provinsi) data.school.provinsi = meta.provinsi;
      }

      const mode = document.querySelector('input[name="excelImportMode"]:checked')?.value || 'append';
      if (mode === 'replace') {
        data.indicators = parsed.items;
      } else {
        parsed.items.forEach(item => {
          data.indicators.push(item);
        });
      }

      const unverifiedCount = parsed.items.filter(i => !i.terverifikasi).length;
      window._pendingExcelWorkbook = null;
      closeCustomModal();
      triggerAutoSave("Input Data", "Impor Excel Rapor", `Mengimpor ${parsed.items.length} indikator dari sheet "${sheetName}" (${window._pendingExcelFileName}).`);
      renderAll();

      if (unverifiedCount > 0) {
        showToast(`Berhasil mengimpor ${parsed.items.length} indikator. ${unverifiedCount} indikator belum memiliki nilai angka dan ditandai untuk Anda lengkapi.`, "warning");
      } else {
        showToast(`Berhasil mengimpor ${parsed.items.length} indikator Rapor Pendidikan dari Excel!`, "success");
      }
    }

    function downloadRaporExcelTemplate() {
      if (!window.XLSX) return;
      const sampleRows = [
        {
          "Kode Indikator": "A.1",
          "Nama Indikator": "Kemampuan Literasi",
          "Dimensi": "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
          "Nilai Capaian": "",
          "Kategori Capaian": "",
          "Perubahan Tahun Lalu": "",
          "Peringkat / Posisi": "",
          "Rekomendasi Rapor Pendidikan": ""
        },
        {
          "Kode Indikator": "A.2",
          "Nama Indikator": "Kemampuan Numerasi",
          "Dimensi": "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
          "Nilai Capaian": "",
          "Kategori Capaian": "",
          "Perubahan Tahun Lalu": "",
          "Peringkat / Posisi": "",
          "Rekomendasi Rapor Pendidikan": ""
        },
        {
          "Kode Indikator": "D.1",
          "Nama Indikator": "Kualitas Pembelajaran",
          "Dimensi": "Dimensi D: Mutu dan Relevansi Pembelajaran",
          "Nilai Capaian": "",
          "Kategori Capaian": "",
          "Perubahan Tahun Lalu": "",
          "Peringkat / Posisi": "",
          "Rekomendasi Rapor Pendidikan": ""
        }
      ];
      const ws = XLSX.utils.json_to_sheet(sampleRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template_Rapor_Pendidikan");
      XLSX.writeFile(wb, "Template_Input_Rapor_Pendidikan.xlsx");
      showToast("Template Excel berhasil diunduh. Isi nilai sesuai Rapor Pendidikan sekolah Anda.", "success");
    }

    // Real PDF Upload & Text Extraction Handler (using PDF.js)
    async function handlePdfUpload(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      if (!window.pdfjsLib) {
        showToast("Pustaka pembaca PDF belum siap. Silakan coba lagi.", "error");
        return;
      }

      showToast("Membaca dokumen PDF Rapor Pendidikan...", "info");
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = "";
        for (let p = 1; p <= pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const content = await page.getTextContent();
          const strings = content.items.map(item => item.str);
          fullText += strings.join(" ") + "\n";
        }

        const data = getCurrentData();
        // Attempt to extract NPSN if school NPSN is empty
        const npsnMatch = fullText.match(/NPSN\s*[:\-]?\s*(\d{8})/i);
        if (npsnMatch && !data.school.npsn) {
          data.school.npsn = npsnMatch[1];
        }

        // Match standard Rapor Pendidikan indicators and look for nearby numeric score (without inventing missing numbers!)
        let extractedCount = 0;
        let needsConfirmCount = 0;

        KATALOG_INDIKATOR_STANDAR.forEach(std => {
          const escapedName = std.nama.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const regex = new RegExp(`(?:${std.kode}\\s+)?${escapedName}[\\s\\S]{0,120}?(\\d{1,3}(?:[.,]\\d{1,2})?)`, 'i');
          const match = fullText.match(regex);

          if (match) {
            const parsedScore = parseFloat(match[1].replace(',', '.'));
            const validScore = !isNaN(parsedScore) && parsedScore >= 0 && parsedScore <= 100;

            // Check if already exists
            const already = data.indicators.some(i => i.kode === std.kode);
            if (!already) {
              data.indicators.push({
                id: uid(),
                kode: std.kode,
                nama: std.nama,
                dimensi: std.dimensi,
                nilai: validScore ? parsedScore : "",
                kategori: validScore ? classifyCategory(parsedScore) : "Belum Terverifikasi",
                perubahan: "",
                peringkat: "",
                rekomendasiRapor: "",
                terverifikasi: false // Always require human confirmation on PDF OCR/text extraction
              });
              extractedCount++;
              needsConfirmCount++;
            }
          }
        });

        triggerAutoSave("Input Data", "Ekstraksi PDF Rapor", `Mengekstrak ${extractedCount} indikator dari ${file.name}`);
        renderAll();

        if (extractedCount > 0) {
          openCustomModal(`
            <div class="p-6">
              <div class="flex items-start gap-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center shrink-0">
                  <i data-lucide="file-check-2" class="w-5 h-5"></i>
                </div>
                <div>
                  <h3 class="font-bold text-slate-900 text-base">Hasil Pembacaan PDF Memerlukan Konfirmasi</h3>
                  <p class="text-xs text-slate-600 mt-1 leading-relaxed">
                    Ditemukan <strong>${extractedCount} indikator</strong> dari dokumen <em>${escapeHtml(file.name)}</em> (${pdf.numPages} halaman).
                    Karena format tata letak PDF Rapor Pendidikan dapat bervariasi, sistem <strong>tidak mengarang</strong> angka perubahan tahun lalu atau peringkat yang tidak terbaca jelas.
                  </p>
                  <p class="text-xs text-amber-800 bg-amber-50 border border-amber-200 rounded-xl p-3 mt-3 font-medium">
                    Silakan periksa setiap baris indikator di tabel, lengkapi nilai yang kosong jika ada, lalu klik tombol <strong>Konfirmasi</strong>.
                  </p>
                </div>
              </div>
              <div class="mt-5 flex justify-end">
                <button onclick="closeCustomModal()" class="px-5 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Mengerti, Saya Periksa</button>
              </div>
            </div>
          `, 'max-w-lg');
        } else {
          openCustomModal(`
            <div class="p-6">
              <div class="flex items-start gap-3">
                <div class="w-10 h-10 rounded-xl bg-rose-100 text-rose-700 flex items-center justify-center shrink-0">
                  <i data-lucide="file-warning" class="w-5 h-5"></i>
                </div>
                <div>
                  <h3 class="font-bold text-slate-900 text-base">Teks Indikator PDF Tidak Terbaca Otomatis</h3>
                  <p class="text-xs text-slate-600 mt-1 leading-relaxed">
                    Dokumen <em>${escapeHtml(file.name)}</em> kemungkinan berupa PDF hasil pindaian gambar (scan) atau memiliki struktur tabel yang tidak standar. Sesuai prinsip integritas data, sistem <strong>tidak mengarang data</strong>.
                  </p>
                  <p class="text-xs text-slate-700 mt-2">
                    Saran: Gunakan file <strong>Unduhan Excel Rapor Pendidikan</strong> dari platform Kemendikdasmen, atau pilih indikator dari <strong>Katalog Indikator PBD</strong> lalu input nilainya secara manual.
                  </p>
                </div>
              </div>
              <div class="mt-5 flex justify-end gap-2">
                <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 text-xs font-semibold">Tutup</button>
                <button onclick="closeCustomModal(); openBulkCatalogPickerModal();" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Pilih dari Katalog Indikator</button>
              </div>
            </div>
          `, 'max-w-lg');
        }
      } catch (err) {
        console.error(err);
        showToast("Terjadi kendala saat membaca file PDF.", "error");
      }
      event.target.value = "";
    }

    // =========================================================================
    // MODULE D: ANALISIS RAPOR PENDIDIKAN
    // =========================================================================
    function renderAnalisisModule(container) {
      const data = getCurrentData();
      const inds = data.indicators;

      if (inds.length === 0) {
        container.innerHTML = `
          <div class="bg-white rounded-2xl border border-slate-200 p-10 text-center">
            <i data-lucide="bar-chart-3" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Data Rapor untuk Dianalisis</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Masukkan data Rapor Pendidikan terlebih dahulu pada Modul 1 agar sistem dapat mengelompokkan capaian dan mengidentifikasi indikator prioritas.</p>
            <button onclick="navigateTo('input')" class="mt-4 px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Ke Modul 1: Input Data Rapor</button>
          </div>
        `;
        return;
      }

      const baikList = inds.filter(i => classifyCategory(i.nilai, i.kategori) === 'Baik');
      const sedangList = inds.filter(i => classifyCategory(i.nilai, i.kategori) === 'Sedang');
      const kurangList = inds.filter(i => classifyCategory(i.nilai, i.kategori) === 'Kurang');
      const turunList = inds.filter(i => !isNaN(parseFloat(i.perubahan)) && parseFloat(i.perubahan) < 0);
      const naikList = inds.filter(i => !isNaN(parseFloat(i.perubahan)) && parseFloat(i.perubahan) > 0);
      const prioritasList = inds.filter(needsImprovement);

      const filter = window.appState.analisisFilter || 'semua';
      let filtered = inds;
      if (filter === 'baik') filtered = baikList;
      else if (filter === 'sedang') filtered = sedangList;
      else if (filter === 'kurang') filtered = kurangList;
      else if (filter === 'turun') filtered = turunList;
      else if (filter === 'naik') filtered = naikList;
      else if (filter === 'prioritas') filtered = prioritasList;

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">2. Analisis Capaian Rapor Pendidikan</h2>
            <p class="text-xs text-slate-500">Pengelompokan otomatis capaian, tren perubahan tahun lalu, dan indikator prioritas pembenahan.</p>
          </div>
          <div class="flex items-center gap-2">
            <button onclick="navigateTo('akar')" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <span>Lanjut ke Akar Masalah</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- 6 Classification Cards -->
        <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
          <button onclick="setAnalisisFilter('baik')" class="p-4 rounded-2xl border text-left transition ${filter === 'baik' ? 'bg-emerald-900 text-white border-emerald-900 shadow' : 'bg-white border-slate-200 hover:border-emerald-400'}">
            <div class="text-[11px] font-semibold ${filter === 'baik' ? 'text-emerald-200' : 'text-slate-500'}">Capaian Baik</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'baik' ? 'text-white' : 'text-emerald-700'}">${baikList.length}</div>
          </button>
          <button onclick="setAnalisisFilter('sedang')" class="p-4 rounded-2xl border text-left transition ${filter === 'sedang' ? 'bg-amber-600 text-white border-amber-600 shadow' : 'bg-white border-slate-200 hover:border-amber-400'}">
            <div class="text-[11px] font-semibold ${filter === 'sedang' ? 'text-amber-100' : 'text-slate-500'}">Capaian Sedang</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'sedang' ? 'text-white' : 'text-amber-600'}">${sedangList.length}</div>
          </button>
          <button onclick="setAnalisisFilter('kurang')" class="p-4 rounded-2xl border text-left transition ${filter === 'kurang' ? 'bg-rose-700 text-white border-rose-700 shadow' : 'bg-white border-slate-200 hover:border-rose-400'}">
            <div class="text-[11px] font-semibold ${filter === 'kurang' ? 'text-rose-100' : 'text-slate-500'}">Capaian Rendah</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'kurang' ? 'text-white' : 'text-rose-700'}">${kurangList.length}</div>
          </button>
          <button onclick="setAnalisisFilter('turun')" class="p-4 rounded-2xl border text-left transition ${filter === 'turun' ? 'bg-rose-900 text-white border-rose-900 shadow' : 'bg-white border-slate-200 hover:border-rose-400'}">
            <div class="text-[11px] font-semibold ${filter === 'turun' ? 'text-rose-200' : 'text-slate-500'}">Mengalami Penurunan</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'turun' ? 'text-white' : 'text-rose-700'}">${turunList.length}</div>
          </button>
          <button onclick="setAnalisisFilter('naik')" class="p-4 rounded-2xl border text-left transition ${filter === 'naik' ? 'bg-emerald-800 text-white border-emerald-800 shadow' : 'bg-white border-slate-200 hover:border-emerald-400'}">
            <div class="text-[11px] font-semibold ${filter === 'naik' ? 'text-emerald-200' : 'text-slate-500'}">Mengalami Peningkatan</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'naik' ? 'text-white' : 'text-emerald-700'}">${naikList.length}</div>
          </button>
          <button onclick="setAnalisisFilter('prioritas')" class="p-4 rounded-2xl border text-left transition ${filter === 'prioritas' ? 'bg-slate-900 text-amber-300 border-slate-900 shadow' : 'bg-amber-50/70 border-amber-200 hover:border-amber-400'}">
            <div class="text-[11px] font-bold ${filter === 'prioritas' ? 'text-amber-200' : 'text-amber-900'}">Prioritas Dibenahi</div>
            <div class="text-2xl font-extrabold mt-1 ${filter === 'prioritas' ? 'text-amber-300' : 'text-amber-700'}">${prioritasList.length}</div>
          </button>
        </div>

        <!-- Plain-Language Narrative Explanation for School Team -->
        <div class="bg-emerald-50/80 border border-emerald-200 rounded-2xl p-5">
          <h3 class="font-bold text-emerald-950 text-sm flex items-center gap-2">
            <i data-lucide="book-open" class="w-4 h-4 text-emerald-800"></i>
            <span>Ringkasan Makna Data untuk Kepala Sekolah &amp; Dewan Guru</span>
          </h3>
          <div class="mt-2 text-xs text-emerald-950/90 space-y-1.5 leading-relaxed">
            <p>
              Dari total <strong>${inds.length} indikator</strong> yang dianalisis, satuan pendidikan memiliki <strong>${baikList.length} indikator berkategori Baik</strong> (aset kekuatan sekolah yang perlu dipertahankan), <strong>${sedangList.length} indikator berkategori Sedang</strong>, dan <strong>${kurangList.length} indikator berkategori Kurang</strong>.
            </p>
            ${turunList.length > 0 ? `
              <p class="text-rose-900 font-medium">
                Perhatian khusus diperlukan pada <strong>${turunList.length} indikator yang mengalami penurunan skor</strong> dibanding tahun sebelumnya (${turunList.map(i => i.nama).join(', ')}), karena penurunan menandakan adanya praktik pembelajaran atau pengelolaan yang melemah.
              </p>` : ''}
            <p>
              Sebanyak <strong>${prioritasList.length} indikator</strong> ditetapkan sebagai fokus utama pembenahan untuk ditelusuri akar masalahnya pada tahap berikutnya.
            </p>
          </div>
        </div>

        <!-- Required Table: Nama Indikator | Nilai | Kategori | Perubahan | Status | Prioritas -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          <div class="p-4 sm:p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
            <div class="flex items-center gap-2 flex-wrap">
              <span class="text-xs font-bold text-slate-700">Filter Tampilan:</span>
              <button onclick="setAnalisisFilter('semua')" class="px-3 py-1 rounded-lg text-xs font-semibold ${filter === 'semua' ? 'bg-emerald-900 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}">Semua (${inds.length})</button>
              <button onclick="setAnalisisFilter('prioritas')" class="px-3 py-1 rounded-lg text-xs font-semibold ${filter === 'prioritas' ? 'bg-amber-600 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}">Perlu Pembenahan (${prioritasList.length})</button>
            </div>
            <span class="text-xs text-slate-500">Menampilkan ${filtered.length} indikator</span>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse text-xs">
              <thead>
                <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase text-[11px]">
                  <th class="py-3 px-4">Nama Indikator</th>
                  <th class="py-3 px-4 text-center">Nilai</th>
                  <th class="py-3 px-4 text-center">Kategori</th>
                  <th class="py-3 px-4 text-center">Perubahan</th>
                  <th class="py-3 px-4 text-center">Status</th>
                  <th class="py-3 px-4 text-center">Prioritas</th>
                  <th class="py-3 px-4 text-right">Tindak Lanjut</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100">
                ${filtered.map(ind => {
                  const kat = classifyCategory(ind.nilai, ind.kategori);
                  const isPrio = needsImprovement(ind);
                  const delta = parseFloat(ind.perubahan);
                  let statusTren = "Stabil / Belum Ada Data";
                  if (!isNaN(delta)) {
                    statusTren = delta < 0 ? "Menurun (Waspada)" : (delta > 0 ? "Meningkat" : "Tetap");
                  }
                  return `
                    <tr class="hover:bg-slate-50 transition">
                      <td class="py-3 px-4">
                        <div class="font-bold text-slate-900">${ind.kode ? `<span class="font-mono text-emerald-800 mr-1">[${escapeHtml(ind.kode)}]</span>` : ''}${escapeHtml(ind.nama)}</div>
                        <div class="text-[11px] text-slate-500">${escapeHtml(ind.dimensi || '')}</div>
                      </td>
                      <td class="py-3 px-4 text-center font-mono font-bold text-sm">${ind.nilai !== '' ? escapeHtml(ind.nilai) : '-'}</td>
                      <td class="py-3 px-4 text-center">${getCategoryBadgeHtml(kat)}</td>
                      <td class="py-3 px-4 text-center">${getDeltaBadgeHtml(ind.perubahan)}</td>
                      <td class="py-3 px-4 text-center font-medium text-slate-700">${statusTren}</td>
                      <td class="py-3 px-4 text-center">
                        ${isPrio
                          ? `<span class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-amber-100 text-amber-900 border border-amber-300">Prioritas Pembenahan</span>`
                          : `<span class="px-2.5 py-1 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-600">Pertahankan Mutu</span>`}
                      </td>
                      <td class="py-3 px-4 text-right">
                        <button onclick="window.appState.selectedAkarIndicatorId='${ind.id}'; navigateTo('akar');" class="px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-900 font-bold text-[11px] transition">
                          Bedah Akar Masalah →
                        </button>
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    }

    function setAnalisisFilter(f) {
      window.appState.analisisFilter = f;
      renderActiveModule();
      if (window.lucide) window.lucide.createIcons();
    }

    // =========================================================================
    // MODULE E: ANALISIS AKAR MASALAH (5 WHY & SEBAB-AKIBAT, DUGAAN VS FAKTA)
    // =========================================================================
    function renderAkarMasalahModule(container) {
      const data = getCurrentData();
      const problemInds = data.indicators.filter(needsImprovement);
      const allInds = data.indicators;

      if (allInds.length === 0) {
        container.innerHTML = `
          <div class="bg-white rounded-2xl border border-slate-200 p-10 text-center">
            <i data-lucide="git-fork" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Indikator Rapor Pendidikan</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Input data Rapor Pendidikan terlebih dahulu sebelum melakukan analisis akar masalah.</p>
            <button onclick="navigateTo('input')" class="mt-4 px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Ke Modul 1: Input Data</button>
          </div>
        `;
        return;
      }

      const targetList = problemInds.length > 0 ? problemInds : allInds;
      if (!window.appState.selectedAkarIndicatorId || !targetList.some(i => i.id === window.appState.selectedAkarIndicatorId)) {
        window.appState.selectedAkarIndicatorId = targetList[0].id;
      }

      const activeInd = targetList.find(i => i.id === window.appState.selectedAkarIndicatorId);
      const rc = data.rootCauses[activeInd.id] || null;

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">3. Analisis Akar Masalah Mendalam (5 Why &amp; Sebab-Akibat)</h2>
            <p class="text-xs text-slate-500">Membedakan secara tegas antara <strong>Dugaan/Hipotesis AI</strong> dengan <strong>Fakta Terkonfirmasi Sekolah</strong>.</p>
          </div>
          <div class="flex items-center gap-2">
            <button onclick="generateAiRootCause('${activeInd.id}')" id="btnGenerateAkarAi" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <i data-lucide="sparkles" class="w-4 h-4"></i>
              <span>Analisis Akar Masalah dengan AI</span>
            </button>
            <button onclick="navigateTo('prioritas')" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <span>Penentuan Prioritas</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- Indicator Selector Pills -->
        <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
          <div class="text-xs font-bold text-slate-700 mb-2">Pilih Indikator yang Akan Dibedah Akar Masalahnya:</div>
          <div class="flex flex-wrap gap-2">
            ${targetList.map(ind => {
              const isSelected = ind.id === activeInd.id;
              const hasRc = Boolean(data.rootCauses[ind.id]);
              return `
                <button onclick="window.appState.selectedAkarIndicatorId='${ind.id}'; renderActiveModule(); if(window.lucide) window.lucide.createIcons();"
                  class="px-3.5 py-2 rounded-xl text-xs font-semibold border transition flex items-center gap-2 ${isSelected ? 'emerald-gradient text-white border-emerald-900 shadow-sm' : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'}">
                  <span>${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ''}${escapeHtml(ind.nama)}</span>
                  ${hasRc ? `<i data-lucide="check-circle-2" class="w-3.5 h-3.5 ${isSelected ? 'text-amber-300' : 'text-emerald-600'}"></i>` : ''}
                </button>
              `;
            }).join('')}
          </div>
        </div>

        <!-- Active Indicator Context Header -->
        <div class="bg-slate-900 text-white rounded-2xl p-5 flex flex-wrap items-center justify-between gap-4">
          <div>
            <div class="text-[11px] text-amber-300 font-semibold uppercase tracking-wider">${escapeHtml(activeInd.dimensi)}</div>
            <h3 class="text-base sm:text-lg font-extrabold mt-0.5">${activeInd.kode ? `[${escapeHtml(activeInd.kode)}] ` : ''}${escapeHtml(activeInd.nama)}</h3>
            <p class="text-xs text-slate-300 mt-1">Nilai Capaian: <strong class="text-white">${activeInd.nilai !== '' ? activeInd.nilai : 'Belum diisi'}</strong> • Kategori: <strong class="text-amber-300">${classifyCategory(activeInd.nilai, activeInd.kategori)}</strong> • Perubahan: <strong>${activeInd.perubahan !== '' ? activeInd.perubahan : 'Tidak ada data'}</strong></p>
          </div>
          <div class="flex gap-2">
            <button onclick="openEditRootCauseModal('${activeInd.id}')" class="px-3.5 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-white text-xs font-semibold flex items-center gap-1.5">
              <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
              <span>${rc ? 'Edit Manual Akar Masalah' : 'Isi Manual Akar Masalah'}</span>
            </button>
          </div>
        </div>

        ${rc ? `
          <!-- 7 Structured Root Cause Components -->
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
            <!-- 1. Masalah Utama & 6. Dampak terhadap Peserta Didik -->
            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
              <div>
                <span class="text-[11px] font-bold uppercase tracking-wider text-rose-700 bg-rose-50 px-2.5 py-0.5 rounded-md">1. Masalah Utama</span>
                <p class="text-xs sm:text-sm font-semibold text-slate-900 mt-2 leading-relaxed">${escapeHtml(rc.masalahUtama)}</p>
              </div>
              <div class="pt-3 border-t border-slate-100">
                <span class="text-[11px] font-bold uppercase tracking-wider text-amber-800 bg-amber-50 px-2.5 py-0.5 rounded-md">6. Dampak terhadap Peserta Didik</span>
                <p class="text-xs text-slate-700 mt-2 leading-relaxed">${escapeHtml(rc.dampakSiswa)}</p>
              </div>
              <div class="pt-3 border-t border-slate-100">
                <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800 bg-emerald-50 px-2.5 py-0.5 rounded-md">3. Bukti / Data Pendukung (Rapor &amp; Sekolah)</span>
                <p class="text-xs text-slate-700 mt-2 leading-relaxed">${escapeHtml(rc.buktiPendukung)}</p>
              </div>
            </div>

            <!-- Metode 5 Why -->
            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
              <div class="flex items-center justify-between mb-3">
                <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-900 bg-emerald-100 px-2.5 py-0.5 rounded-md">Analisis Berjenjang (Metode 5 Why)</span>
                <span class="text-[11px] text-slate-400">Penelusuran Sebab-Akibat</span>
              </div>
              <div class="space-y-2.5">
                ${(rc.fiveWhys || []).map((w, idx) => `
                  <div class="flex items-start gap-2.5 text-xs">
                    <span class="w-6 h-6 rounded-lg bg-emerald-900 text-amber-300 font-bold flex items-center justify-center shrink-0 text-[11px]">W${idx + 1}</span>
                    <div class="flex-1 bg-slate-50 rounded-xl p-2.5 border border-slate-200/80 text-slate-800 leading-relaxed">${escapeHtml(w)}</div>
                  </div>
                `).join('')}
              </div>
            </div>
          </div>

          <!-- 2. Kemungkinan Akar Penyebab (Distinguishing Dugaan AI vs Fakta Terkonfirmasi) -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-xs">
            <div class="flex flex-wrap items-center justify-between gap-2 mb-4">
              <div>
                <h4 class="font-bold text-slate-900 text-sm">2. Kemungkinan Akar Penyebab (Verifikasi Tim Sekolah)</h4>
                <p class="text-xs text-slate-500">Klik tombol konfirmasi untuk mengubah status dari <strong>Dugaan AI</strong> menjadi <strong>Fakta Terkonfirmasi Sekolah</strong> setelah didiskusikan guru.</p>
              </div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
              ${(rc.akarPenyebab || []).map((ak, idx) => {
                const isFact = ak.status === 'fakta_sekolah';
                return `
                  <div class="rounded-xl p-4 border ${isFact ? 'bg-emerald-50/70 border-emerald-300' : 'bg-amber-50/50 border-amber-200'} flex flex-col justify-between gap-3">
                    <div>
                      <div class="flex items-center justify-between gap-2 mb-2">
                        <span class="text-[10px] font-bold uppercase px-2 py-0.5 rounded bg-white border border-slate-200 text-slate-600">${escapeHtml(ak.kategori || 'Pembelajaran & GTK')}</span>
                        ${isFact
                          ? `<span class="inline-flex items-center gap-1 text-[11px] font-extrabold text-emerald-800 bg-emerald-200/70 px-2 py-0.5 rounded-full"><i data-lucide="badge-check" class="w-3.5 h-3.5"></i> FAKTA TERKONFIRMASI SEKOLAH</span>`
                          : `<span class="inline-flex items-center gap-1 text-[11px] font-bold text-amber-800 bg-amber-200/70 px-2 py-0.5 rounded-full"><i data-lucide="help-circle" class="w-3.5 h-3.5"></i> DUGAAN / HIPOTESIS AI</span>`}
                      </div>
                      <p class="text-xs text-slate-800 font-medium leading-relaxed">${escapeHtml(ak.teks)}</p>
                    </div>
                    <div class="flex justify-end">
                      <button onclick="toggleRootCauseFactStatus('${activeInd.id}', ${idx})" class="px-3 py-1.5 rounded-lg text-[11px] font-bold transition ${isFact ? 'bg-white border border-slate-300 text-slate-700 hover:bg-slate-50' : 'emerald-gradient text-white hover:opacity-95'}">
                        ${isFact ? 'Kembalikan ke Hipotesis' : 'Konfirmasi Sebagai Fakta Sekolah ✓'}
                      </button>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          </div>

          <!-- 4. Pertanyaan Refleksi Guru, 5. Data Tambahan, 7. Alternatif Solusi -->
          <div class="grid grid-cols-1 lg:grid-cols-3 gap-5">
            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
              <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wider text-emerald-900 mb-3 flex items-center gap-1.5">
                <i data-lucide="message-square-question" class="w-4 h-4"></i>
                <span>4. Pertanyaan Refleksi Diskusi Guru</span>
              </h4>
              <ul class="space-y-2 text-xs text-slate-700 list-disc pl-4">
                ${(rc.pertanyaanRefleksi || []).map(q => `<li class="leading-relaxed">${escapeHtml(q)}</li>`).join('')}
              </ul>
            </div>

            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
              <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wider text-amber-800 mb-3 flex items-center gap-1.5">
                <i data-lucide="folder-search" class="w-4 h-4"></i>
                <span>5. Data Tambahan Perlu Dikumpulkan</span>
              </h4>
              <ul class="space-y-2 text-xs text-slate-700 list-disc pl-4">
                ${(rc.dataTambahan || []).map(d => `<li class="leading-relaxed">${escapeHtml(d)}</li>`).join('')}
              </ul>
            </div>

            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
              <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wider text-emerald-800 mb-3 flex items-center gap-1.5">
                <i data-lucide="check-square" class="w-4 h-4"></i>
                <span>7. Alternatif Solusi Pembenahan</span>
              </h4>
              <ul class="space-y-2 text-xs text-slate-700 list-disc pl-4">
                ${(rc.alternatifSolusi || []).map(sol => `<li class="leading-relaxed">${escapeHtml(sol)}</li>`).join('')}
              </ul>
            </div>
          </div>
        ` : `
          <div class="bg-white rounded-2xl border-2 border-dashed border-slate-300 p-10 text-center">
            <div class="w-12 h-12 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center mx-auto mb-3">
              <i data-lucide="sparkles" class="w-6 h-6"></i>
            </div>
            <h4 class="font-bold text-slate-900 text-base">Indikator Ini Belum Memiliki Analisis Akar Masalah</h4>
            <p class="text-xs text-slate-500 max-w-lg mx-auto mt-1 leading-relaxed">
              Klik <strong>Analisis Akar Masalah dengan AI</strong> untuk menyusun analisis 5 Why, dugaan penyebab, pertanyaan refleksi guru, dan alternatif solusi berbasis data indikator ini, atau isi secara manual.
            </p>
            <div class="mt-5 flex flex-wrap justify-center gap-2">
              <button onclick="generateAiRootCause('${activeInd.id}')" class="px-5 py-2.5 rounded-xl gold-gradient text-emerald-950 font-extrabold text-xs flex items-center gap-1.5 shadow-sm">
                <i data-lucide="sparkles" class="w-4 h-4"></i> Analisis dengan AI Sekarang
              </button>
              <button onclick="openEditRootCauseModal('${activeInd.id}')" class="px-5 py-2.5 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs">
                Isi Manual Formulir
              </button>
            </div>
          </div>
        `}
      `;
    }

    async function generateAiRootCause(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;

      const btn = document.getElementById('btnGenerateAkarAi');
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i><span>Menganalisis Data...</span>`;
        if (window.lucide) window.lucide.createIcons();
      }

      const systemPrompt = `Anda adalah pakar Perencanaan Berbasis Data (PBD) Kemendikdasmen Indonesia. Tugas Anda menganalisis akar masalah Rapor Pendidikan secara objektif, empatik, dan realistis untuk tim sekolah.
ATURAN KETAT:
1. Gunakan HANYA data indikator yang diberikan pengguna. Jangan mengarang fakta internal sekolah yang tidak ada.
2. Seluruh akar penyebab yang Anda rumuskan adalah HIPOTESIS / DUGAAN AWAL yang harus dikonfirmasi guru di sekolah.
3. Gunakan Bahasa Indonesia formal-edukatif yang mudah dipahami kepala sekolah dan guru.`;

      const userPrompt = `Lakukan analisis akar masalah mendalam untuk indikator Rapor Pendidikan berikut:
- Sekolah: ${data.school.nama || 'Satuan Pendidikan'} (${data.school.jenjang || 'Jenjang belum diisi'})
- Dimensi: ${ind.dimensi}
- Indikator: ${ind.kode} ${ind.nama}
- Nilai Capaian: ${ind.nilai !== '' ? ind.nilai : 'Belum tersedia'}
- Kategori: ${classifyCategory(ind.nilai, ind.kategori)}
- Perubahan dibanding tahun lalu: ${ind.perubahan !== '' ? ind.perubahan : 'Tidak ada data'}
- Rekomendasi Rapor Pendidikan: ${ind.rekomendasiRapor || 'Belum dicantumkan'}

Susun respons terstruktur mencakup: Masalah Utama, 5 langkah analisis 5 Why, 3-4 kemungkinan akar penyebab (kategorikan: Pembelajaran, Kompetensi GTK, Sarpras/Media, atau Manajemen Sekolah), Bukti Pendukung dari data Rapor di atas, 3 Pertanyaan Refleksi untuk diskusi KKG/MGMP/Kombel guru, 3 Data Tambahan internal sekolah yang perlu dicek, Dampak terhadap Peserta Didik, dan 3 Alternatif Solusi konkret.`;

      const schema = {
        type: "OBJECT",
        properties: {
          masalahUtama: { type: "STRING" },
          fiveWhys: { type: "ARRAY", items: { type: "STRING" } },
          akarPenyebab: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                kategori: { type: "STRING" },
                teks: { type: "STRING" }
              },
              required: ["kategori", "teks"]
            }
          },
          buktiPendukung: { type: "STRING" },
          pertanyaanRefleksi: { type: "ARRAY", items: { type: "STRING" } },
          dataTambahan: { type: "ARRAY", items: { type: "STRING" } },
          dampakSiswa: { type: "STRING" },
          alternatifSolusi: { type: "ARRAY", items: { type: "STRING" } }
        },
        required: ["masalahUtama", "fiveWhys", "akarPenyebab", "buktiPendukung", "pertanyaanRefleksi", "dataTambahan", "dampakSiswa", "alternatifSolusi"]
      };

      try {
        const result = await callGeminiAPI(userPrompt, systemPrompt, schema);
        data.rootCauses[indicatorId] = {
          indicatorId,
          masalahUtama: result.masalahUtama,
          fiveWhys: result.fiveWhys || [],
          akarPenyebab: (result.akarPenyebab || []).map(a => ({
            kategori: a.kategori,
            teks: a.teks,
            status: 'dugaan_ai'
          })),
          buktiPendukung: result.buktiPendukung,
          pertanyaanRefleksi: result.pertanyaanRefleksi || [],
          dataTambahan: result.dataTambahan || [],
          dampakSiswa: result.dampakSiswa,
          alternatifSolusi: result.alternatifSolusi || []
        };
        triggerAutoSave("Akar Masalah", "Analisis AI 5-Why", `Menyusun hipotesis akar masalah untuk ${ind.kode} ${ind.nama}`);
        renderAll();
        showToast("Analisis akar masalah selesai. Silakan verifikasi hipotesis AI bersama tim guru.", "success");
      } catch (err) {
        // Fallback deterministic analytical synthesis when offline/API unavailable
        const scoreInfo = ind.nilai !== '' ? `skor ${ind.nilai} (${classifyCategory(ind.nilai, ind.kategori)})` : 'capaian yang perlu pembenahan';
        const deltaInfo = !isNaN(parseFloat(ind.perubahan)) ? ` dengan perubahan ${ind.perubahan} dari tahun sebelumnya` : '';
        data.rootCauses[indicatorId] = {
          indicatorId,
          masalahUtama: `Capaian indikator ${ind.nama} masih berada pada ${scoreInfo}${deltaInfo}, sehingga belum mencapai standar optimal yang diharapkan pada ${ind.dimensi}.`,
          fiveWhys: [
            `Why 1: Capaian ${ind.nama} peserta didik belum optimal (${scoreInfo}).`,
            `Why 2: Proses pembelajaran dan asesmen formatif di kelas belum sepenuhnya terdiferensiasi sesuai kebutuhan belajar siswa.`,
            `Why 3: Guru membutuhkan penguatan praktik baik melalui komunitas belajar (Kombel) dalam merancang pembelajaran aktif.`,
            `Why 4: Pemanfaatan sumber belajar, buku bacaan/media ajar, dan refleksi pembelajaran pasca-supervisi belum berjalan rutin.`,
            `Why 5: Perencanaan program peningkatan mutu dan alokasi anggaran sekolah belum sepenuhnya terfokus pada akar pembenahan indikator ini.`
          ],
          akarPenyebab: [
            { kategori: "Pembelajaran", teks: `Pembelajaran berpusat pada peserta didik dan asesmen diagnostik awal untuk ${ind.nama} belum diterapkan secara merata di semua kelas.`, status: "dugaan_ai" },
            { kategori: "Kompetensi GTK", teks: `Optimalisasi Komunitas Belajar (Kombel) sekolah untuk berbagi praktik baik terkait ${ind.nama} masih perlu ditingkatkan.`, status: "dugaan_ai" },
            { kategori: "Sarpras & Media", teks: `Ketersediaan dan pemanfaatan media pembelajaran pendukung ${ind.nama} di kelas belum merata.`, status: "dugaan_ai" }
          ],
          buktiPendukung: `Data Rapor Pendidikan menunjukkan ${ind.nama} memperoleh ${scoreInfo}${deltaInfo}.${ind.rekomendasiRapor ? ' Catatan Rapor: ' + ind.rekomendasiRapor : ''}`,
          pertanyaanRefleksi: [
            `Strategi pembelajaran apa yang selama ini sudah berjalan baik dan mana yang masih menyulitkan siswa pada aspek ${ind.nama}?`,
            `Bagaimana hasil asesmen diagnostik/sumatif internal sekolah mengonfirmasi capaian Rapor Pendidikan ini?`,
            `Dukungan pelatihan atau sarana apa yang paling mendesak dibutuhkan guru di kelas?`
          ],
          dataTambahan: [
            `Rekap hasil asesmen diagnostik awal dan nilai rapor semester peserta didik.`,
            `Catatan hasil supervisi akademik kepala sekolah di kelas.`,
            `Daftar pemanfaatan koleksi perpustakaan / media ajar oleh guru dan siswa.`
          ],
          dampakSiswa: `Peserta didik berisiko mengalami hambatan dalam menguasai kompetensi dasar pada aspek ${ind.nama}, yang dapat memengaruhi kesiapan belajar pada jenjang berikutnya.`,
          alternatifSolusi: [
            `Mengaktifkan Komunitas Belajar (Kombel) guru untuk bedah modul ajar dan praktik baik peningkatan ${ind.nama}.`,
            `Melaksanakan asesmen diagnostik berkala dan pendampingan belajar bagi peserta didik yang membutuhkan.`,
            `Menyediakan bahan bacaan, alat peraga, atau media pembelajaran kontekstual pendukung ${ind.nama}.`
          ]
        };
        triggerAutoSave("Akar Masalah", "Analisis Terstruktur", `Menyusun kerangka akar masalah untuk ${ind.kode} ${ind.nama}`);
        renderAll();
        showToast("Analisis kerangka akar masalah disusun (Mode Offline/Terstruktur). Silakan tinjau dan konfirmasi fakta sekolah.", "info");
      }
    }

    function toggleRootCauseFactStatus(indicatorId, causeIdx) {
      const data = getCurrentData();
      const rc = data.rootCauses[indicatorId];
      if (!rc || !rc.akarPenyebab[causeIdx]) return;
      const item = rc.akarPenyebab[causeIdx];
      item.status = item.status === 'fakta_sekolah' ? 'dugaan_ai' : 'fakta_sekolah';
      triggerAutoSave("Akar Masalah", "Verifikasi Fakta", `Mengubah status akar penyebab menjadi ${item.status === 'fakta_sekolah' ? 'Fakta Sekolah' : 'Hipotesis'}`);
      renderActiveModule();
      if (window.lucide) window.lucide.createIcons();
      showToast(item.status === 'fakta_sekolah' ? "Dikonfirmasi sebagai Fakta Sekolah." : "Dikembalikan ke status Hipotesis AI.", "success");
    }

    function openEditRootCauseModal(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      const rc = data.rootCauses[indicatorId] || {
        masalahUtama: "",
        fiveWhys: ["", "", "", "", ""],
        akarPenyebab: [{ kategori: "Pembelajaran", teks: "", status: "fakta_sekolah" }],
        buktiPendukung: "",
        pertanyaanRefleksi: [""],
        dataTambahan: [""],
        dampakSiswa: "",
        alternatifSolusi: [""]
      };

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Analisis Akar Masalah — ${escapeHtml(ind.nama)}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRootCauseModal(event, '${indicatorId}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">1. Masalah Utama</label>
              <textarea name="masalahUtama" rows="2" required class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.masalahUtama)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">2. Akar Penyebab Terkonfirmasi (Satu baris per penyebab)</label>
              <textarea name="akarPenyebabLines" rows="3" required placeholder="Tulis satu akar penyebab per baris..." class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml((rc.akarPenyebab || []).map(a => a.teks).join('\n'))}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Metode 5 Why (Satu langkah per baris, maks 5 baris)</label>
              <textarea name="fiveWhysLines" rows="4" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml((rc.fiveWhys || []).join('\n'))}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">3. Bukti / Data Pendukung</label>
              <textarea name="buktiPendukung" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.buktiPendukung)}</textarea>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">4. Pertanyaan Refleksi Guru (per baris)</label>
                <textarea name="refleksiLines" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml((rc.pertanyaanRefleksi || []).join('\n'))}</textarea>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">5. Data Tambahan Diperlukan (per baris)</label>
                <textarea name="dataTambahanLines" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml((rc.dataTambahan || []).join('\n'))}</textarea>
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">6. Dampak terhadap Peserta Didik</label>
              <textarea name="dampakSiswa" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.dampakSiswa)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">7. Alternatif Solusi (per baris)</label>
              <textarea name="solusiLines" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml((rc.alternatifSolusi || []).join('\n'))}</textarea>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Analisis</button>
            </div>
          </form>
        </div>
      `, 'max-w-3xl');
    }

    function saveRootCauseModal(e, indicatorId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const splitLines = (str) => String(str || '').split('\n').map(s => s.trim()).filter(Boolean);

      data.rootCauses[indicatorId] = {
        indicatorId,
        masalahUtama: fd.get('masalahUtama').trim(),
        akarPenyebab: splitLines(fd.get('akarPenyebabLines')).map(teks => ({
          kategori: "Terkonfirmasi Tim Sekolah",
          teks,
          status: "fakta_sekolah"
        })),
        fiveWhys: splitLines(fd.get('fiveWhysLines')).slice(0, 5),
        buktiPendukung: fd.get('buktiPendukung').trim(),
        pertanyaanRefleksi: splitLines(fd.get('refleksiLines')),
        dataTambahan: splitLines(fd.get('dataTambahanLines')),
        dampakSiswa: fd.get('dampakSiswa').trim(),
        alternatifSolusi: splitLines(fd.get('solusiLines'))
      };

      closeCustomModal();
      triggerAutoSave("Akar Masalah", "Simpan Manual", "Menyimpan hasil penyuntingan analisis akar masalah.");
      renderAll();
      showToast("Analisis akar masalah berhasil disimpan.", "success");
    }

    // =========================================================================
    // MODULE F: PENENTUAN PRIORITAS PEMBENAHAN
    // =========================================================================
    function computeDefaultPriorityForIndicator(ind) {
      const kat = classifyCategory(ind.nilai, ind.kategori);
      const delta = parseFloat(ind.perubahan);

      const skorCapaian = kat === 'Kurang' ? 5 : (kat === 'Sedang' ? 4 : 2);
      const skorTren = (!isNaN(delta) && delta < -2) ? 5 : ((!isNaN(delta) && delta < 0) ? 4 : 3);
      const skorDampak = (ind.kode && (ind.kode.startsWith('A.') || ind.kode === 'D.1' || ind.kode === 'D.4')) ? 5 : 4;
      const skorUrgensi = kat === 'Kurang' ? 5 : 4;
      const skorKemampuan = 4; // Feasible by school internal effort & BOSP

      const totalSkor = skorCapaian + skorTren + skorDampak + skorUrgensi + skorKemampuan;
      let alasan = `Indikator ${ind.nama} berkategori ${kat}`;
      if (!isNaN(delta) && delta < 0) alasan += ` dan mengalami penurunan ${delta} poin dari tahun lalu`;
      alasan += `, berdampak langsung pada kualitas belajar peserta didik serta dapat dibenahi melalui program sekolah.`;

      return {
        indicatorId: ind.id,
        skorCapaian,
        skorTren,
        skorDampak,
        skorUrgensi,
        skorKemampuan,
        totalSkor,
        peringkatManual: null,
        alasanPrioritas: alasan
      };
    }

    function getRankedPriorities() {
      const data = getCurrentData();
      const targetInds = data.indicators.filter(needsImprovement);
      const list = targetInds.map(ind => {
        if (!data.priorities[ind.id]) {
          data.priorities[ind.id] = computeDefaultPriorityForIndicator(ind);
        }
        const p = data.priorities[ind.id];
        p.totalSkor = Number(p.skorCapaian) + Number(p.skorTren) + Number(p.skorDampak) + Number(p.skorUrgensi) + Number(p.skorKemampuan);
        return { ind, prio: p };
      });

      list.sort((a, b) => {
        if (a.prio.peringkatManual && b.prio.peringkatManual) {
          return a.prio.peringkatManual - b.prio.peringkatManual;
        }
        if (a.prio.peringkatManual) return -1;
        if (b.prio.peringkatManual) return 1;
        return b.prio.totalSkor - a.prio.totalSkor;
      });

      return list;
    }

    function renderPrioritasModule(container) {
      const ranked = getRankedPriorities();

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">4. Penentuan Prioritas Pembenahan</h2>
            <p class="text-xs text-slate-500">Pemeringkatan multi-kriteria (Skala 1–5 per kriteria) dengan fitur penyesuaian manual oleh tim sekolah.</p>
          </div>
          <div class="flex items-center gap-2">
            <button onclick="navigateTo('rekomendasi')" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <i data-lucide="sparkles" class="w-4 h-4"></i>
              <span>Buat Rekomendasi Program</span>
            </button>
          </div>
        </div>

        ${ranked.length > 0 ? `
          <div class="space-y-4">
            ${ranked.map(({ ind, prio }, idx) => {
              const rankNum = idx + 1;
              return `
                <div class="bg-white rounded-2xl border ${rankNum <= 3 ? 'border-amber-300 shadow-sm' : 'border-slate-200'} p-5 transition">
                  <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-slate-100">
                    <div class="flex items-start gap-3.5">
                      <div class="w-11 h-11 rounded-2xl ${rankNum === 1 ? 'gold-gradient text-emerald-950' : (rankNum <= 3 ? 'emerald-gradient text-white' : 'bg-slate-100 text-slate-700')} font-extrabold text-base flex flex-col items-center justify-center shrink-0">
                        <span class="text-[9px] uppercase font-bold opacity-80">Rank</span>
                        <span class="leading-none">#${rankNum}</span>
                      </div>
                      <div>
                        <div class="flex items-center gap-2 flex-wrap">
                          <h3 class="font-bold text-slate-900 text-sm sm:text-base">${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ''}${escapeHtml(ind.nama)}</h3>
                          ${getCategoryBadgeHtml(classifyCategory(ind.nilai, ind.kategori))}
                          ${prio.peringkatManual ? `<span class="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-200">Urutan Manual</span>` : ''}
                        </div>
                        <p class="text-xs text-slate-600 mt-1 leading-relaxed"><strong>Alasan Prioritas:</strong> ${escapeHtml(prio.alasanPrioritas)}</p>
                      </div>
                    </div>

                    <div class="flex items-center gap-3 shrink-0">
                      <div class="text-right">
                        <div class="text-[11px] text-slate-500">Total Skor</div>
                        <div class="text-xl font-extrabold text-emerald-900 font-mono">${prio.totalSkor}<span class="text-xs text-slate-400 font-normal">/25</span></div>
                      </div>
                      <button onclick="openEditPriorityModal('${ind.id}')" class="px-3.5 py-2 rounded-xl border border-slate-300 hover:bg-slate-50 text-xs font-bold text-slate-700 flex items-center gap-1.5">
                        <i data-lucide="sliders-horizontal" class="w-3.5 h-3.5 text-emerald-700"></i>
                        <span>Ubah Skor / Urutan</span>
                      </button>
                    </div>
                  </div>

                  <!-- 5 Criteria Breakdown Pills -->
                  <div class="mt-3 grid grid-cols-2 sm:grid-cols-5 gap-2 text-[11px]">
                    <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                      <div class="text-slate-500">1. Tingkat Capaian</div>
                      <div class="font-bold text-slate-900 mt-0.5">${prio.skorCapaian} / 5</div>
                    </div>
                    <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                      <div class="text-slate-500">2. Tren Perubahan</div>
                      <div class="font-bold text-slate-900 mt-0.5">${prio.skorTren} / 5</div>
                    </div>
                    <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                      <div class="text-slate-500">3. Dampak ke Siswa</div>
                      <div class="font-bold text-slate-900 mt-0.5">${prio.skorDampak} / 5</div>
                    </div>
                    <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                      <div class="text-slate-500">4. Urgensi Masalah</div>
                      <div class="font-bold text-slate-900 mt-0.5">${prio.skorUrgensi} / 5</div>
                    </div>
                    <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                      <div class="text-slate-500">5. Kemampuan Sekolah</div>
                      <div class="font-bold text-slate-900 mt-0.5">${prio.skorKemampuan} / 5</div>
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        ` : `
          <div class="bg-white rounded-2xl border border-slate-200 p-10 text-center">
            <i data-lucide="list-ordered" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Indikator Prioritas</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Input data Rapor Pendidikan terlebih dahulu agar sistem dapat menghitung peringkat prioritas pembenahan.</p>
            <button onclick="navigateTo('input')" class="mt-4 px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Ke Modul 1: Input Data</button>
          </div>
        `}
      `;
    }

    function openEditPriorityModal(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      const prio = data.priorities[indicatorId] || computeDefaultPriorityForIndicator(ind);

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Sesuaikan Prioritas — ${escapeHtml(ind.nama)}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="savePriorityModal(event, '${indicatorId}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-semibold text-slate-700 mb-1">1. Tingkat Capaian (1–5)</label>
                <input type="number" min="1" max="5" name="skorCapaian" value="${prio.skorCapaian}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">2. Tren Perubahan (1–5)</label>
                <input type="number" min="1" max="5" name="skorTren" value="${prio.skorTren}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">3. Dampak terhadap Peserta Didik (1–5)</label>
                <input type="number" min="1" max="5" name="skorDampak" value="${prio.skorDampak}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">4. Urgensi Masalah (1–5)</label>
                <input type="number" min="1" max="5" name="skorUrgensi" value="${prio.skorUrgensi}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">5. Kemampuan Sekolah Membenahi (1–5)</label>
                <input type="number" min="1" max="5" name="skorKemampuan" value="${prio.skorKemampuan}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-indigo-800 mb-1">Peringkat Manual (Opsional, 1 = Tertinggi)</label>
                <input type="number" min="1" max="50" name="peringkatManual" value="${prio.peringkatManual || ''}" placeholder="Kosongkan untuk otomatis" class="w-full rounded-xl border border-indigo-300 px-3 py-2" />
              </div>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Alasan Prioritas (Justifikasi Tim Sekolah)</label>
              <textarea name="alasanPrioritas" rows="3" required class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(prio.alasanPrioritas)}</textarea>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Prioritas</button>
            </div>
          </form>
        </div>
      `);
    }

    function savePriorityModal(e, indicatorId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const manualVal = fd.get('peringkatManual');
      const sc = Number(fd.get('skorCapaian'));
      const st = Number(fd.get('skorTren'));
      const sd = Number(fd.get('skorDampak'));
      const su = Number(fd.get('skorUrgensi'));
      const sk = Number(fd.get('skorKemampuan'));

      data.priorities[indicatorId] = {
        indicatorId,
        skorCapaian: sc,
        skorTren: st,
        skorDampak: sd,
        skorUrgensi: su,
        skorKemampuan: sk,
        totalSkor: sc + st + sd + su + sk,
        peringkatManual: manualVal ? Number(manualVal) : null,
        alasanPrioritas: fd.get('alasanPrioritas').trim()
      };

      closeCustomModal();
      triggerAutoSave("Penentuan Prioritas", "Ubah Skor/Urutan", "Menyesuaikan pembobotan atau urutan prioritas indikator.");
      renderAll();
      showToast("Peringkat prioritas berhasil diperbarui.", "success");
    }

    // =========================================================================
    // MODULE G: REKOMENDASI PROGRAM
    // =========================================================================
    function renderRekomendasiModule(container) {
      const data = getCurrentData();
      const ranked = getRankedPriorities();

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">5. Rekomendasi Program Pembenahan</h2>
            <p class="text-xs text-slate-500">Rekomendasi program realistis berbasis indikator prioritas &amp; akar masalah. Wajib ditinjau dan disetujui tim sekolah.</p>
          </div>
          <div class="flex flex-wrap items-center gap-2">
            <button onclick="generateAllProgramRecommendations()" id="btnGenerateRekomAi" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
              <i data-lucide="sparkles" class="w-4 h-4"></i>
              <span>Buat Rekomendasi dengan AI</span>
            </button>
            <button onclick="openEditRecommendationModal()" class="px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold flex items-center gap-1.5">
              <i data-lucide="plus" class="w-4 h-4"></i>
              <span>Tambah Manual</span>
            </button>
            ${data.recommendations.length > 0 ? `
              <button onclick="syncRecommendationsToRkjm()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm hover:opacity-95 transition">
                <span>Susun RKJM</span>
                <i data-lucide="arrow-right" class="w-4 h-4"></i>
              </button>` : ''}
          </div>
        </div>

        ${data.recommendations.length > 0 ? `
          <div class="space-y-4">
            ${data.recommendations.map((rec, idx) => {
              const isApproved = rec.statusPersetujuan === 'disetujui';
              return `
                <div class="bg-white rounded-2xl border ${isApproved ? 'border-emerald-300' : 'border-slate-200'} shadow-xs overflow-hidden">
                  <div class="p-5 bg-slate-50/70 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <div class="flex items-center gap-2 flex-wrap">
                        <span class="text-xs font-extrabold px-2.5 py-0.5 rounded-lg bg-emerald-900 text-amber-300">Rekomendasi #${idx + 1}</span>
                        <span class="text-xs font-bold text-emerald-800">${escapeHtml(rec.indicatorNama)}</span>
                        ${isApproved
                          ? `<span class="inline-flex items-center gap-1 text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800"><i data-lucide="check-circle-2" class="w-3.5 h-3.5"></i> Disetujui Tim Sekolah</span>`
                          : `<span class="inline-flex items-center gap-1 text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-900"><i data-lucide="clock" class="w-3.5 h-3.5"></i> Menunggu Tinjauan Tim</span>`}
                      </div>
                      <h3 class="font-extrabold text-slate-900 text-base mt-1.5">${escapeHtml(rec.bentukKegiatan)}</h3>
                    </div>

                    <div class="flex items-center gap-2 flex-wrap">
                      <button onclick="toggleRecommendationApproval('${rec.id}')" class="px-3 py-1.5 rounded-xl text-xs font-bold transition ${isApproved ? 'bg-white border border-slate-300 text-slate-700' : 'emerald-gradient text-white shadow-xs'}">
                        ${isApproved ? 'Batalkan Persetujuan' : 'Setujui Program ✓'}
                      </button>
                      <button onclick="openEditRecommendationModal('${rec.id}')" class="p-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700" title="Edit Rekomendasi">
                        <i data-lucide="pencil" class="w-3.5 h-3.5"></i>
                      </button>
                      <button onclick="deleteRecommendation('${rec.id}')" class="p-2 rounded-xl border border-slate-200 bg-white hover:bg-rose-50 text-rose-700" title="Hapus">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                      </button>
                    </div>
                  </div>

                  <div class="p-5 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Permasalahan</div>
                      <p class="text-slate-800 mt-1 leading-relaxed">${escapeHtml(rec.permasalahan)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Tujuan Perbaikan</div>
                      <p class="text-slate-800 mt-1 leading-relaxed">${escapeHtml(rec.tujuanPerbaikan)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Sasaran Program</div>
                      <p class="text-slate-800 mt-1 leading-relaxed">${escapeHtml(rec.sasaran)}</p>
                    </div>
                    <div class="md:col-span-2">
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Langkah Pelaksanaan</div>
                      <p class="text-slate-800 mt-1 whitespace-pre-line leading-relaxed">${escapeHtml(rec.langkahPelaksanaan)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Indikator Keberhasilan &amp; Target</div>
                      <p class="text-slate-800 mt-1 leading-relaxed"><strong>Indikator:</strong> ${escapeHtml(rec.indikatorKeberhasilan)}<br/><strong>Target:</strong> ${escapeHtml(rec.targetCapaian)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Penanggung Jawab &amp; Waktu</div>
                      <p class="text-slate-800 mt-1"><strong>PJ:</strong> ${escapeHtml(rec.penanggungJawab)}<br/><strong>Waktu:</strong> ${escapeHtml(rec.waktuPelaksanaan)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Kebutuhan Sumber Daya</div>
                      <p class="text-slate-800 mt-1">${escapeHtml(rec.sumberDaya)}</p>
                    </div>
                    <div>
                      <div class="font-bold text-slate-500 uppercase text-[10px]">Perkiraan Anggaran</div>
                      <p class="text-emerald-900 font-extrabold text-sm mt-1">${formatRupiah(rec.perkiraanAnggaran)}</p>
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        ` : `
          <div class="bg-white rounded-2xl border-2 border-dashed border-slate-300 p-10 text-center">
            <div class="w-12 h-12 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center mx-auto mb-3">
              <i data-lucide="lightbulb" class="w-6 h-6"></i>
            </div>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Rekomendasi Program Pembenahan</h3>
            <p class="text-xs text-slate-500 max-w-lg mx-auto mt-1 leading-relaxed">
              ${ranked.length > 0
                ? `Sistem mendeteksi <strong>${ranked.length} indikator prioritas</strong>. Klik tombol <strong>Buat Rekomendasi dengan AI</strong> untuk merancang program pembenahan yang realistis dan terukur.`
                : `Silakan input data Rapor Pendidikan terlebih dahulu pada Modul 1 agar sistem dapat merancang rekomendasi program.`}
            </p>
            ${ranked.length > 0 ? `
              <button onclick="generateAllProgramRecommendations()" class="mt-5 px-5 py-2.5 rounded-xl gold-gradient text-emerald-950 font-extrabold text-xs inline-flex items-center gap-1.5 shadow-sm">
                <i data-lucide="sparkles" class="w-4 h-4"></i> Buat Rekomendasi Program Sekarang
              </button>` : ''}
          </div>
        `}
      `;
    }

    async function generateAllProgramRecommendations() {
      const data = getCurrentData();
      const ranked = getRankedPriorities();
      if (ranked.length === 0) {
        showToast("Belum ada indikator prioritas. Silakan isi data Rapor Pendidikan terlebih dahulu.", "warning");
        return;
      }

      const btn = document.getElementById('btnGenerateRekomAi');
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i><span>Menyusun Rekomendasi...</span>`;
        if (window.lucide) window.lucide.createIcons();
      }

      const topPriorities = ranked.slice(0, 6);
      const summaryContext = topPriorities.map(({ ind, prio }, i) => {
        const rc = data.rootCauses[ind.id];
        const confirmedFacts = rc ? rc.akarPenyebab.filter(a => a.status === 'fakta_sekolah').map(a => a.teks).join('; ') : '';
        return `${i + 1}. [${ind.kode}] ${ind.nama} (Nilai: ${ind.nilai}, Kategori: ${classifyCategory(ind.nilai, ind.kategori)}, Skor Prioritas: ${prio.totalSkor}). Akar masalah: ${confirmedFacts || (rc ? rc.masalahUtama : 'Perlu penguatan pembelajaran dan kompetensi GTK')}`;
      }).join('\n');

      const systemPrompt = `Anda adalah konsultan Perencanaan Berbasis Data (PBD) dan ARKAS sekolah di Indonesia. Susun rekomendasi program perbaikan sekolah yang sangat realistis, operasional, hemat biaya namun berdampak tinggi sesuai kondisi sekolah di Indonesia. Gunakan Bahasa Indonesia formal.`;

      const userPrompt = `Berdasarkan daftar indikator prioritas sekolah berikut (${data.school.nama || 'Satuan Pendidikan'}, Jenjang: ${data.school.jenjang || 'Sekolah'}):
${summaryContext}

Buat 1 rekomendasi program konkret untuk setiap indikator di atas (maksimal ${topPriorities.length} program).`;

      const schema = {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            priorityIndex: { type: "INTEGER" },
            permasalahan: { type: "STRING" },
            tujuanPerbaikan: { type: "STRING" },
            sasaran: { type: "STRING" },
            bentukKegiatan: { type: "STRING" },
            langkahPelaksanaan: { type: "STRING" },
            penanggungJawab: { type: "STRING" },
            indikatorKeberhasilan: { type: "STRING" },
            targetCapaian: { type: "STRING" },
            waktuPelaksanaan: { type: "STRING" },
            sumberDaya: { type: "STRING" },
            perkiraanAnggaran: { type: "NUMBER" }
          },
          required: ["priorityIndex", "permasalahan", "tujuanPerbaikan", "sasaran", "bentukKegiatan", "langkahPelaksanaan", "penanggungJawab", "indikatorKeberhasilan", "targetCapaian", "waktuPelaksanaan", "sumberDaya", "perkiraanAnggaran"]
        }
      };

      try {
        const aiList = await callGeminiAPI(userPrompt, systemPrompt, schema);
        aiList.forEach((item, idx) => {
          const target = topPriorities[item.priorityIndex - 1] || topPriorities[idx];
          if (!target) return;
          data.recommendations.push({
            id: uid(),
            indicatorId: target.ind.id,
            indicatorNama: `${target.ind.kode ? '[' + target.ind.kode + '] ' : ''}${target.ind.nama}`,
            permasalahan: item.permasalahan,
            tujuanPerbaikan: item.tujuanPerbaikan,
            sasaran: item.sasaran,
            bentukKegiatan: item.bentukKegiatan,
            langkahPelaksanaan: item.langkahPelaksanaan,
            penanggungJawab: item.penanggungJawab,
            indikatorKeberhasilan: item.indikatorKeberhasilan,
            targetCapaian: item.targetCapaian,
            waktuPelaksanaan: item.waktuPelaksanaan,
            sumberDaya: item.sumberDaya,
            perkiraanAnggaran: Number(item.perkiraanAnggaran) || 0,
            statusPersetujuan: 'menunggu',
            sumberAi: true
          });
        });
        triggerAutoSave("Rekomendasi Program", "Generate Rekomendasi AI", `Menyusun ${aiList.length} rekomendasi program berbasis prioritas.`);
        renderAll();
        showToast("Rekomendasi program berhasil disusun! Silakan tinjau dan setujui setiap program.", "success");
      } catch (err) {
        // Structured deterministic fallback grounded in user's actual indicators
        topPriorities.forEach(({ ind }) => {
          const rc = data.rootCauses[ind.id];
          data.recommendations.push({
            id: uid(),
            indicatorId: ind.id,
            indicatorNama: `${ind.kode ? '[' + ind.kode + '] ' : ''}${ind.nama}`,
            permasalahan: rc ? rc.masalahUtama : `Capaian ${ind.nama} (${ind.nilai || 'perlu penguatan'}) masih membutuhkan peningkatan sistematis.`,
            tujuanPerbaikan: `Meningkatkan kualitas proses pembelajaran dan capaian peserta didik pada indikator ${ind.nama}.`,
            sasaran: `Seluruh guru mata pelajaran/kelas dan peserta didik di ${data.school.nama || 'satuan pendidikan'}.`,
            bentukKegiatan: `Optimalisasi Komunitas Belajar (Kombel) & Penguatan Pembelajaran Kontekstual ${ind.nama}`,
            langkahPelaksanaan: `1. Pemetaan kebutuhan belajar peserta didik dan refleksi guru.\n2. Lokakarya penyusunan perangkat ajar & asesmen formatif di Kombel.\n3. Implementasi pembelajaran terdiferensiasi di kelas.\n4. Evaluasi berkala dan tindak lanjut hasil belajar.`,
            penanggungJawab: `Wakil Kepala Sekolah Bidang Kurikulum & Koordinator Kombel`,
            indikatorKeberhasilan: `Minimal 85% guru menerapkan perangkat ajar hasil Kombel dan skor ${ind.nama} meningkat.`,
            targetCapaian: `Kenaikan capaian ${ind.nama} minimal +5 poin pada Rapor Pendidikan berikutnya`,
            waktuPelaksanaan: `Triwulan I – Triwulan IV Tahun ${data.school.tahunPerencanaan || window.appState.activeYear}`,
            sumberDaya: `Narasumber internal/pengawas pembina, modul ajar, bahan bacaan/media pembelajaran, ATK Kombel`,
            perkiraanAnggaran: 3500000,
            statusPersetujuan: 'menunggu',
            sumberAi: true
          });
        });
        triggerAutoSave("Rekomendasi Program", "Susun Rekomendasi Terstruktur", `Menyusun ${topPriorities.length} rekomendasi program berbasis indikator prioritas.`);
        renderAll();
        showToast("Rekomendasi program berhasil disusun berdasarkan indikator prioritas Anda.", "success");
      }
    }

    function toggleRecommendationApproval(recId) {
      const data = getCurrentData();
      const rec = data.recommendations.find(r => r.id === recId);
      if (rec) {
        rec.statusPersetujuan = rec.statusPersetujuan === 'disetujui' ? 'menunggu' : 'disetujui';
        triggerAutoSave("Rekomendasi Program", "Persetujuan Program", `Status "${rec.bentukKegiatan}" diubah menjadi ${rec.statusPersetujuan}.`);
        renderAll();
        showToast(rec.statusPersetujuan === 'disetujui' ? "Program disetujui tim sekolah." : "Persetujuan dibatalkan.", "success");
      }
    }

    function deleteRecommendation(recId) {
      showConfirmDialog({
        title: "Hapus Rekomendasi Program?",
        message: "Rekomendasi ini akan dihapus dari daftar.",
        confirmText: "Ya, Hapus",
        danger: true,
        onConfirm: () => {
          const data = getCurrentData();
          data.recommendations = data.recommendations.filter(r => r.id !== recId);
          triggerAutoSave("Rekomendasi Program", "Hapus Program", "Menghapus rekomendasi program.");
          renderAll();
          showToast("Rekomendasi dihapus.", "info");
        }
      });
    }

    function openEditRecommendationModal(recId = null) {
      const data = getCurrentData();
      const rec = recId ? data.recommendations.find(r => r.id === recId) : null;
      const inds = data.indicators;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">${rec ? 'Edit Rekomendasi Program' : 'Tambah Rekomendasi Program'}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRecommendationModal(event, '${recId || ''}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Indikator Rapor Terkait <span class="text-rose-600">*</span></label>
              <select name="indicatorId" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${inds.map(i => `<option value="${i.id}" ${rec?.indicatorId === i.id ? 'selected' : ''}>${i.kode ? '[' + i.kode + '] ' : ''}${escapeHtml(i.nama)}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Nama / Bentuk Program Kegiatan <span class="text-rose-600">*</span></label>
              <input type="text" name="bentukKegiatan" value="${escapeHtml(rec?.bentukKegiatan || '')}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Permasalahan</label>
                <textarea name="permasalahan" rows="2" required class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rec?.permasalahan || '')}</textarea>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Tujuan Perbaikan</label>
                <textarea name="tujuanPerbaikan" rows="2" required class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rec?.tujuanPerbaikan || '')}</textarea>
              </div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Sasaran Program</label>
                <input type="text" name="sasaran" value="${escapeHtml(rec?.sasaran || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Penanggung Jawab</label>
                <input type="text" name="penanggungJawab" value="${escapeHtml(rec?.penanggungJawab || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Langkah Pelaksanaan</label>
              <textarea name="langkahPelaksanaan" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rec?.langkahPelaksanaan || '')}</textarea>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Indikator Keberhasilan</label>
                <input type="text" name="indikatorKeberhasilan" value="${escapeHtml(rec?.indikatorKeberhasilan || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Target Capaian</label>
                <input type="text" name="targetCapaian" value="${escapeHtml(rec?.targetCapaian || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Waktu Pelaksanaan</label>
                <input type="text" name="waktuPelaksanaan" value="${escapeHtml(rec?.waktuPelaksanaan || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kebutuhan Sumber Daya</label>
                <input type="text" name="sumberDaya" value="${escapeHtml(rec?.sumberDaya || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Perkiraan Anggaran (Rp)</label>
                <input type="number" min="0" name="perkiraanAnggaran" value="${rec?.perkiraanAnggaran || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Program</button>
            </div>
          </form>
        </div>
      `, 'max-w-3xl');
    }

    function saveRecommendationModal(e, recId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const indId = fd.get('indicatorId');
      const ind = data.indicators.find(i => i.id === indId);

      const obj = {
        id: recId || uid(),
        indicatorId: indId,
        indicatorNama: ind ? `${ind.kode ? '[' + ind.kode + '] ' : ''}${ind.nama}` : 'Indikator Rapor',
        permasalahan: fd.get('permasalahan').trim(),
        tujuanPerbaikan: fd.get('tujuanPerbaikan').trim(),
        sasaran: fd.get('sasaran').trim(),
        bentukKegiatan: fd.get('bentukKegiatan').trim(),
        langkahPelaksanaan: fd.get('langkahPelaksanaan').trim(),
        penanggungJawab: fd.get('penanggungJawab').trim(),
        indikatorKeberhasilan: fd.get('indikatorKeberhasilan').trim(),
        targetCapaian: fd.get('targetCapaian').trim(),
        waktuPelaksanaan: fd.get('waktuPelaksanaan').trim(),
        sumberDaya: fd.get('sumberDaya').trim(),
        perkiraanAnggaran: Number(fd.get('perkiraanAnggaran')) || 0,
        statusPersetujuan: 'disetujui',
        sumberAi: false
      };

      if (recId) {
        const idx = data.recommendations.findIndex(r => r.id === recId);
        if (idx !== -1) data.recommendations[idx] = obj;
      } else {
        data.recommendations.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("Rekomendasi Program", "Simpan Program", `Menyimpan program "${obj.bentukKegiatan}".`);
      renderAll();
      showToast("Rekomendasi program berhasil disimpan.", "success");
    }

    // =========================================================================
    // MODULE H: PENYUSUNAN RKJM (RENCANA KERJA JANGKA MENENGAH)
    // =========================================================================
    function syncRecommendationsToRkjm() {
      const data = getCurrentData();
      if (data.recommendations.length === 0) {
        showToast("Belum ada rekomendasi program untuk diturunkan ke RKJM.", "warning");
        return;
      }

      let added = 0;
      data.recommendations.forEach(rec => {
        const exists = data.rkjm.some(p => p.recommendationId === rec.id);
        if (!exists) {
          const ind = data.indicators.find(i => i.id === rec.indicatorId);
          const rc = data.rootCauses[rec.indicatorId];
          const confirmedCauses = rc ? rc.akarPenyebab.filter(a => a.status === 'fakta_sekolah').map(a => a.teks).join('; ') : '';

          data.rkjm.push({
            id: uid(),
            indicatorId: rec.indicatorId,
            recommendationId: rec.id,
            dimensi: ind ? ind.dimensi : DIMENSI_RAPOR[0],
            indikatorNama: rec.indicatorNama,
            akarMasalah: confirmedCauses || (rc ? rc.masalahUtama : rec.permasalahan),
            tujuanStrategis: rec.tujuanPerbaikan,
            namaProgram: rec.bentukKegiatan,
            targetTahun1: rec.targetCapaian || "Tercapai 75%",
            targetTahun2: "Tercapai 85%",
            targetTahun3: "Tercapai 92%",
            targetTahun4: "Tercapai 100% (Kategori Baik)",
            indikatorKeberhasilan: rec.indikatorKeberhasilan,
            penanggungJawab: rec.penanggungJawab,
            sumberPendanaan: "BOSP Reguler"
          });
          added++;
        }
      });

      triggerAutoSave("RKJM", "Turunkan ke RKJM", `Menurunkan ${added} program prioritas ke dalam matriks RKJM.`);
      navigateTo('rkjm');
      showToast(added > 0 ? `${added} program berhasil diturunkan ke RKJM dan terhubung dengan Rapor Pendidikan.` : "Seluruh rekomendasi sudah terhubung di RKJM.", "success");
    }

    function renderRkjmModule(container) {
      const data = getCurrentData();
      const s = data.school;
      const yStart = parseInt(s.periodeRkjmMulai || window.appState.activeYear, 10) || 2026;
      const yEnd = parseInt(s.periodeRkjmAkhir || (yStart + 3), 10) || (yStart + 3);

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">6. Penyusunan Rencana Kerja Jangka Menengah (RKJM)</h2>
            <p class="text-xs text-slate-500">Periode Perencanaan: <strong>Tahun ${yStart} – ${yEnd}</strong> • Terhubung langsung dengan Dimensi &amp; Indikator Rapor Pendidikan.</p>
          </div>
          <div class="flex flex-wrap items-center gap-2">
            <button onclick="openRkjmPeriodModal()" class="px-3 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold flex items-center gap-1.5">
              <i data-lucide="calendar" class="w-3.5 h-3.5 text-emerald-700"></i>
              <span>Atur Periode (${yStart}-${yEnd})</span>
            </button>
            <button onclick="syncRecommendationsToRkjm()" class="px-3.5 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
              <span>Tarik dari Rekomendasi</span>
            </button>
            <button onclick="openEditRkjmModal()" class="px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold flex items-center gap-1">
              <i data-lucide="plus" class="w-3.5 h-3.5"></i> Tambah Program
            </button>
            ${data.rkjm.length > 0 ? `
              <button onclick="syncRkjmToRkt()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm">
                <span>Susun RKT</span>
                <i data-lucide="arrow-right" class="w-4 h-4"></i>
              </button>` : ''}
          </div>
        </div>

        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          ${data.rkjm.length > 0 ? `
            <div class="overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-emerald-950 text-white font-bold uppercase text-[10px]">
                    <th class="py-3 px-3">Dimensi &amp; Indikator Dibenahi</th>
                    <th class="py-3 px-3">Akar Masalah</th>
                    <th class="py-3 px-3">Tujuan Strategis &amp; Program Pengembangan</th>
                    <th class="py-3 px-3 text-center">Target ${yStart}</th>
                    <th class="py-3 px-3 text-center">Target ${yStart + 1}</th>
                    <th class="py-3 px-3 text-center">Target ${yStart + 2}</th>
                    <th class="py-3 px-3 text-center">Target ${yStart + 3}</th>
                    <th class="py-3 px-3">Indikator Keberhasilan</th>
                    <th class="py-3 px-3">PJ &amp; Sumber Dana</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-200">
                  ${data.rkjm.map(p => `
                    <tr class="hover:bg-slate-50 align-top">
                      <td class="py-3 px-3 min-w-[170px]">
                        <span class="inline-block px-2 py-0.5 rounded bg-emerald-50 text-emerald-900 font-bold text-[11px] border border-emerald-200">${escapeHtml(p.indikatorNama)}</span>
                        <div class="text-[10px] text-slate-500 mt-1">${escapeHtml(p.dimensi)}</div>
                      </td>
                      <td class="py-3 px-3 min-w-[170px] text-slate-700">${escapeHtml(p.akarMasalah)}</td>
                      <td class="py-3 px-3 min-w-[200px]">
                        <div class="font-bold text-slate-900">${escapeHtml(p.namaProgram)}</div>
                        <div class="text-[11px] text-slate-600 mt-1">Tujuan: ${escapeHtml(p.tujuanStrategis)}</div>
                      </td>
                      <td class="py-3 px-3 text-center bg-amber-50/40 font-medium text-slate-800 min-w-[95px]">${escapeHtml(p.targetTahun1)}</td>
                      <td class="py-3 px-3 text-center text-slate-700 min-w-[95px]">${escapeHtml(p.targetTahun2)}</td>
                      <td class="py-3 px-3 text-center text-slate-700 min-w-[95px]">${escapeHtml(p.targetTahun3)}</td>
                      <td class="py-3 px-3 text-center text-slate-700 min-w-[95px]">${escapeHtml(p.targetTahun4)}</td>
                      <td class="py-3 px-3 min-w-[150px] text-slate-700">${escapeHtml(p.indikatorKeberhasilan)}</td>
                      <td class="py-3 px-3 min-w-[130px]">
                        <div class="font-semibold text-slate-800">${escapeHtml(p.penanggungJawab)}</div>
                        <span class="inline-block mt-1 px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] font-semibold">${escapeHtml(p.sumberPendanaan)}</span>
                      </td>
                      <td class="py-3 px-3 text-right whitespace-nowrap">
                        <button onclick="openEditRkjmModal('${p.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit RKJM"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                        <button onclick="deleteRkjm('${p.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          ` : `
            <div class="p-10 text-center">
              <i data-lucide="calendar-range" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
              <h3 class="font-bold text-slate-900 text-base">Matriks RKJM Masih Kosong</h3>
              <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Tarik program dari Modul Rekomendasi Program atau tambahkan program pengembangan sekolah secara manual.</p>
              <div class="mt-4 flex justify-center gap-2">
                <button onclick="syncRecommendationsToRkjm()" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold">Tarik dari Rekomendasi</button>
                <button onclick="openEditRkjmModal()" class="px-4 py-2 rounded-xl border border-slate-300 text-xs font-bold">Tambah Manual</button>
              </div>
            </div>
          `}
        </div>
      `;
    }

    function openRkjmPeriodModal() {
      const s = getCurrentData().school;
      openCustomModal(`
        <div class="p-6">
          <h3 class="font-bold text-slate-900 text-base mb-3">Atur Periode Perencanaan RKJM (4 Tahun)</h3>
          <form onsubmit="saveRkjmPeriod(event)" class="space-y-3 text-xs">
            <div class="grid grid-cols-2 gap-3">
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Tahun Mulai</label>
                <input type="number" name="mulai" value="${escapeHtml(s.periodeRkjmMulai || window.appState.activeYear)}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Tahun Berakhir</label>
                <input type="number" name="akhir" value="${escapeHtml(s.periodeRkjmAkhir || (Number(window.appState.activeYear) + 3))}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div class="flex justify-end gap-2 pt-3">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Periode</button>
            </div>
          </form>
        </div>
      `, 'max-w-md');
    }

    function saveRkjmPeriod(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const s = getCurrentData().school;
      s.periodeRkjmMulai = fd.get('mulai');
      s.periodeRkjmAkhir = fd.get('akhir');
      closeCustomModal();
      triggerAutoSave("RKJM", "Atur Periode", `Periode RKJM diatur ke ${s.periodeRkjmMulai}-${s.periodeRkjmAkhir}`);
      renderAll();
    }

    function openEditRkjmModal(rkjmId = null) {
      const data = getCurrentData();
      const existing = rkjmId ? data.rkjm.find(p => p.id === rkjmId) : null;
      const inds = data.indicators;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">${existing ? 'Edit Program RKJM' : 'Tambah Program RKJM'}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRkjmModal(event, '${rkjmId || ''}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Indikator Rapor Pendidikan yang Dibenahi <span class="text-rose-600">*</span></label>
              <select name="indicatorId" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${inds.length > 0
                  ? inds.map(i => `<option value="${i.id}" ${existing?.indicatorId === i.id ? 'selected' : ''}>${i.kode ? '[' + i.kode + '] ' : ''}${escapeHtml(i.nama)}</option>`).join('')
                  : `<option value="">-- Input Indikator Rapor Terlebih Dahulu --</option>`}
              </select>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Nama Program Pengembangan Sekolah <span class="text-rose-600">*</span></label>
              <input type="text" name="namaProgram" value="${escapeHtml(existing?.namaProgram || '')}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Akar Masalah</label>
                <textarea name="akarMasalah" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(existing?.akarMasalah || '')}</textarea>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Tujuan Strategis</label>
                <textarea name="tujuanStrategis" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(existing?.tujuanStrategis || '')}</textarea>
              </div>
            </div>
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Target Tahun 1</label>
                <input type="text" name="targetTahun1" value="${escapeHtml(existing?.targetTahun1 || '')}" class="w-full rounded-xl border border-slate-300 px-2.5 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Target Tahun 2</label>
                <input type="text" name="targetTahun2" value="${escapeHtml(existing?.targetTahun2 || '')}" class="w-full rounded-xl border border-slate-300 px-2.5 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Target Tahun 3</label>
                <input type="text" name="targetTahun3" value="${escapeHtml(existing?.targetTahun3 || '')}" class="w-full rounded-xl border border-slate-300 px-2.5 py-2" />
              </div>
              <div>
                <label class="block font-semibold text-slate-700 mb-1">Target Tahun 4</label>
                <input type="text" name="targetTahun4" value="${escapeHtml(existing?.targetTahun4 || '')}" class="w-full rounded-xl border border-slate-300 px-2.5 py-2" />
              </div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Indikator Keberhasilan</label>
                <input type="text" name="indikatorKeberhasilan" value="${escapeHtml(existing?.indikatorKeberhasilan || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Penanggung Jawab</label>
                <input type="text" name="penanggungJawab" value="${escapeHtml(existing?.penanggungJawab || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Sumber Pendanaan</label>
                <select name="sumberPendanaan" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${["BOSP Reguler", "BOSP Kinerja", "APBD / Daerah", "Komite / Swadaya Masyarakat", "Lainnya"].map(sd => `<option value="${sd}" ${existing?.sumberPendanaan === sd ? 'selected' : ''}>${sd}</option>`).join('')}
                </select>
              </div>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan RKJM</button>
            </div>
          </form>
        </div>
      `, 'max-w-3xl');
    }

    function saveRkjmModal(e, rkjmId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const indId = fd.get('indicatorId');
      const ind = data.indicators.find(i => i.id === indId);

      const obj = {
        id: rkjmId || uid(),
        indicatorId: indId,
        recommendationId: rkjmId ? data.rkjm.find(p => p.id === rkjmId)?.recommendationId : null,
        dimensi: ind ? ind.dimensi : DIMENSI_RAPOR[0],
        indikatorNama: ind ? `${ind.kode ? '[' + ind.kode + '] ' : ''}${ind.nama}` : 'Umum',
        akarMasalah: fd.get('akarMasalah').trim(),
        tujuanStrategis: fd.get('tujuanStrategis').trim(),
        namaProgram: fd.get('namaProgram').trim(),
        targetTahun1: fd.get('targetTahun1').trim(),
        targetTahun2: fd.get('targetTahun2').trim(),
        targetTahun3: fd.get('targetTahun3').trim(),
        targetTahun4: fd.get('targetTahun4').trim(),
        indikatorKeberhasilan: fd.get('indikatorKeberhasilan').trim(),
        penanggungJawab: fd.get('penanggungJawab').trim(),
        sumberPendanaan: fd.get('sumberPendanaan')
      };

      if (rkjmId) {
        const idx = data.rkjm.findIndex(p => p.id === rkjmId);
        if (idx !== -1) data.rkjm[idx] = obj;
      } else {
        data.rkjm.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("RKJM", "Simpan Program RKJM", `Menyimpan program RKJM "${obj.namaProgram}".`);
      renderAll();
      showToast("Program RKJM berhasil disimpan.", "success");
    }

    function deleteRkjm(rkjmId) {
      showConfirmDialog({
        title: "Hapus Program RKJM?",
        message: "Menghapus program RKJM ini akan memutuskan tautannya dari kegiatan RKT terkait.",
        confirmText: "Ya, Hapus",
        danger: true,
        onConfirm: () => {
          const data = getCurrentData();
          data.rkjm = data.rkjm.filter(p => p.id !== rkjmId);
          triggerAutoSave("RKJM", "Hapus Program RKJM", "Menghapus program RKJM.");
          renderAll();
          showToast("Program RKJM dihapus.", "info");
        }
      });
    }

    // =========================================================================
    // MODULE I: PENYUSUNAN RKT (RENCANA KERJA TAHUNAN) & KALENDER KEGIATAN
    // =========================================================================
    function syncRkjmToRkt() {
      const data = getCurrentData();
      if (data.rkjm.length === 0) {
        showToast("Susun program RKJM terlebih dahulu sebelum menurunkan ke RKT.", "warning");
        return;
      }

      let added = 0;
      data.rkjm.forEach(prog => {
        const exists = data.rkt.some(k => k.rkjmId === prog.id);
        if (!exists) {
          const rec = data.recommendations.find(r => r.id === prog.recommendationId);
          data.rkt.push({
            id: uid(),
            rkjmId: prog.id,
            indicatorId: prog.indicatorId,
            namaProgram: prog.namaProgram,
            namaKegiatan: `Pelaksanaan ${prog.namaProgram}`,
            latarBelakang: `Menindaklanjuti capaian ${prog.indikatorNama} (${prog.akarMasalah})`,
            tujuan: prog.tujuanStrategis,
            sasaran: rec?.sasaran || "Guru dan Peserta Didik",
            indikatorKeberhasilan: prog.indikatorKeberhasilan,
            target: prog.targetTahun1 || "Terlaksana 100%",
            bulanPelaksanaan: [2, 3, 7, 8, 10], // Default Feb, Mar, Jul, Agu, Okt
            penanggungJawab: prog.penanggungJawab,
            kebutuhanAnggaran: rec?.perkiraanAnggaran || 0,
            sumberDana: prog.sumberPendanaan || "BOSP Reguler",
            buktiPelaksanaan: "Undangan, Daftar Hadir, Notulen, Foto Kegiatan, Laporan Hasil"
          });
          added++;
        }
      });

      triggerAutoSave("RKT", "Turunkan RKJM ke RKT", `Menurunkan ${added} program RKJM menjadi kegiatan operasional RKT.`);
      navigateTo('rkt');
      showToast(added > 0 ? `${added} kegiatan RKT berhasil disusun dari program RKJM.` : "Seluruh program RKJM sudah memiliki kegiatan RKT.", "success");
    }

    function renderRktModule(container) {
      const data = getCurrentData();
      const tahun = data.school.tahunPerencanaan || window.appState.activeYear;

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">7. Penyusunan Rencana Kerja Tahunan (RKT) — Tahun ${escapeHtml(tahun)}</h2>
            <p class="text-xs text-slate-500">Penjabaran operasional program RKJM menjadi kegiatan tahunan beserta kalender jadwal pelaksanaan.</p>
          </div>
          <div class="flex flex-wrap items-center gap-2">
            <button onclick="syncRkjmToRkt()" class="px-3.5 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
              <span>Tarik dari RKJM</span>
            </button>
            <button onclick="openEditRktModal()" class="px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold flex items-center gap-1">
              <i data-lucide="plus" class="w-3.5 h-3.5"></i> Tambah Kegiatan RKT
            </button>
            ${data.rkt.length > 0 ? `
              <button onclick="syncRktToArkas()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm">
                <span>Buat Rancangan ARKAS</span>
                <i data-lucide="arrow-right" class="w-4 h-4"></i>
              </button>` : ''}
          </div>
        </div>

        ${data.rkt.length > 0 ? `
          <!-- Tabel RKT -->
          <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
            <div class="p-4 border-b border-slate-200 flex items-center justify-between">
              <h3 class="font-bold text-slate-900 text-sm">Matriks Rencana Kerja Tahunan (${data.rkt.length} Kegiatan)</h3>
              <span class="text-xs text-slate-500">Total Kebutuhan Anggaran: <strong class="text-emerald-900">${formatRupiah(data.rkt.reduce((s, k) => s + getRktEffectiveBudget(k), 0))}</strong></span>
            </div>
            <div class="overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase text-[10px]">
                    <th class="py-3 px-3">Program (RKJM) &amp; Nama Kegiatan</th>
                    <th class="py-3 px-3">Latar Belakang &amp; Tujuan</th>
                    <th class="py-3 px-3">Sasaran</th>
                    <th class="py-3 px-3">Indikator Keberhasilan &amp; Target</th>
                    <th class="py-3 px-3">PJ &amp; Bukti Pelaksanaan</th>
                    <th class="py-3 px-3 text-right">Anggaran &amp; Sumber</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100 align-top">
                  ${data.rkt.map(k => {
                    const effBudget = getRktEffectiveBudget(k);
                    return `
                      <tr class="hover:bg-slate-50">
                        <td class="py-3 px-3 min-w-[200px]">
                          <div class="text-[10px] font-bold text-emerald-800 uppercase">${escapeHtml(k.namaProgram)}</div>
                          <div class="font-bold text-slate-900 mt-0.5">${escapeHtml(k.namaKegiatan)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[210px] text-slate-700">
                          <div><span class="font-semibold">Latar Belakang:</span> ${escapeHtml(k.latarBelakang)}</div>
                          <div class="mt-1"><span class="font-semibold">Tujuan:</span> ${escapeHtml(k.tujuan)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[130px] text-slate-700">${escapeHtml(k.sasaran)}</td>
                        <td class="py-3 px-3 min-w-[160px] text-slate-700">
                          <div>${escapeHtml(k.indikatorKeberhasilan)}</div>
                          <div class="mt-1 font-semibold text-emerald-800">Target: ${escapeHtml(k.target)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[160px] text-slate-700">
                          <div class="font-semibold text-slate-900">${escapeHtml(k.penanggungJawab)}</div>
                          <div class="text-[11px] text-slate-500 mt-1">Bukti: ${escapeHtml(k.buktiPelaksanaan)}</div>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <div class="font-mono font-bold text-slate-900">${formatRupiah(effBudget)}</div>
                          <span class="inline-block mt-1 px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 text-[10px] font-semibold">${escapeHtml(k.sumberDana)}</span>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <button onclick="openEditRktModal('${k.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit Kegiatan"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteRkt('${k.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <!-- Kalender Kegiatan Tahunan (12 Bulan) -->
          <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
            <div class="p-4 sm:p-5 border-b border-slate-200 flex items-center justify-between">
              <div>
                <h3 class="font-bold text-slate-900 text-sm">Kalender Kegiatan Tahunan RKT (${escapeHtml(tahun)})</h3>
                <p class="text-xs text-slate-500">Klik kotak bulan pada baris kegiatan untuk mengaktifkan/menonaktifkan jadwal pelaksanaan bulanan.</p>
              </div>
            </div>
            <div class="overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-emerald-950 text-white text-[11px]">
                    <th class="py-2.5 px-4 min-w-[220px]">Nama Kegiatan RKT</th>
                    ${BULAN_LIST.map((b, idx) => `<th class="py-2.5 px-2 text-center w-12">${b}</th>`).join('')}
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${data.rkt.map(k => `
                    <tr class="hover:bg-slate-50">
                      <td class="py-2.5 px-4 font-semibold text-slate-800">${escapeHtml(k.namaKegiatan)}</td>
                      ${BULAN_LIST.map((_, idx) => {
                        const monthNum = idx + 1;
                        const active = (k.bulanPelaksanaan || []).includes(monthNum);
                        return `
                          <td class="py-2 px-1 text-center">
                            <button onclick="toggleRktMonth('${k.id}', ${monthNum})" class="w-7 h-7 rounded-lg mx-auto flex items-center justify-center transition ${active ? 'bg-emerald-700 text-amber-300 font-bold shadow-xs' : 'bg-slate-100 hover:bg-slate-200 text-slate-400'}" title="${BULAN_LIST[idx]}">
                              ${active ? '✓' : '·'}
                            </button>
                          </td>
                        `;
                      }).join('')}
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>
        ` : `
          <div class="bg-white rounded-2xl border-2 border-dashed border-slate-300 p-10 text-center">
            <i data-lucide="calendar-check-2" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Kegiatan RKT</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Tarik program dari RKJM untuk menurunkannya menjadi kegiatan tahunan, atau tambah kegiatan secara manual.</p>
            <div class="mt-4 flex justify-center gap-2">
              <button onclick="syncRkjmToRkt()" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold">Tarik dari RKJM</button>
              <button onclick="openEditRktModal()" class="px-4 py-2 rounded-xl border border-slate-300 text-xs font-bold">Tambah Manual</button>
            </div>
          </div>
        `}
      `;
    }

    function getRktEffectiveBudget(rktItem) {
      const data = getCurrentData();
      const linkedArkas = data.arkas.filter(a => a.rktId === rktItem.id);
      if (linkedArkas.length > 0) {
        return linkedArkas.reduce((s, a) => s + (Number(a.jumlahAnggaran) || 0), 0);
      }
      return Number(rktItem.kebutuhanAnggaran) || 0;
    }

    function toggleRktMonth(rktId, monthNum) {
      const data = getCurrentData();
      const k = data.rkt.find(item => item.id === rktId);
      if (!k) return;
      if (!Array.isArray(k.bulanPelaksanaan)) k.bulanPelaksanaan = [];
      if (k.bulanPelaksanaan.includes(monthNum)) {
        k.bulanPelaksanaan = k.bulanPelaksanaan.filter(m => m !== monthNum);
      } else {
        k.bulanPelaksanaan.push(monthNum);
        k.bulanPelaksanaan.sort((a, b) => a - b);
      }
      triggerAutoSave("RKT", "Ubah Jadwal Bulan", `Menyesuaikan jadwal bulan kegiatan "${k.namaKegiatan}".`);
      renderActiveModule();
    }

    function openEditRktModal(rktId = null) {
      const data = getCurrentData();
      const existing = rktId ? data.rkt.find(k => k.id === rktId) : null;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">${existing ? 'Edit Kegiatan RKT' : 'Tambah Kegiatan RKT'}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRktModal(event, '${rktId || ''}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Program RKJM Induk <span class="text-rose-600">*</span></label>
              <select name="rkjmId" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${data.rkjm.length > 0
                  ? data.rkjm.map(p => `<option value="${p.id}" ${existing?.rkjmId === p.id ? 'selected' : ''}>${escapeHtml(p.namaProgram)} (${escapeHtml(p.indikatorNama)})</option>`).join('')
                  : `<option value="">-- Susun Program RKJM Terlebih Dahulu --</option>`}
              </select>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Nama Kegiatan Tahunan (RKT) <span class="text-rose-600">*</span></label>
              <input type="text" name="namaKegiatan" value="${escapeHtml(existing?.namaKegiatan || '')}" required class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Latar Belakang</label>
                <textarea name="latarBelakang" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(existing?.latarBelakang || '')}</textarea>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Tujuan Kegiatan</label>
                <textarea name="tujuan" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(existing?.tujuan || '')}</textarea>
              </div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Sasaran</label>
                <input type="text" name="sasaran" value="${escapeHtml(existing?.sasaran || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Indikator Keberhasilan</label>
                <input type="text" name="indikatorKeberhasilan" value="${escapeHtml(existing?.indikatorKeberhasilan || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Target</label>
                <input type="text" name="target" value="${escapeHtml(existing?.target || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Penanggung Jawab</label>
                <input type="text" name="penanggungJawab" value="${escapeHtml(existing?.penanggungJawab || '')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Perkiraan Anggaran (Rp)</label>
                <input type="number" min="0" name="kebutuhanAnggaran" value="${existing?.kebutuhanAnggaran || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${["BOSP Reguler", "BOSP Kinerja", "APBD / Daerah", "Komite / Swadaya", "Lainnya"].map(sd => `<option value="${sd}" ${existing?.sumberDana === sd ? 'selected' : ''}>${sd}</option>`).join('')}
                </select>
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Bukti Pelaksanaan (Output)</label>
              <input type="text" name="buktiPelaksanaan" value="${escapeHtml(existing?.buktiPelaksanaan || 'Daftar Hadir, Notulen, Foto Dokumentasi, Laporan Kegiatan')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan RKT</button>
            </div>
          </form>
        </div>
      `, 'max-w-3xl');
    }

    function saveRktModal(e, rktId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rkjmId = fd.get('rkjmId');
      const prog = data.rkjm.find(p => p.id === rkjmId);

      const obj = {
        id: rktId || uid(),
        rkjmId,
        indicatorId: prog?.indicatorId || null,
        namaProgram: prog?.namaProgram || 'Program Pengembangan Sekolah',
        namaKegiatan: fd.get('namaKegiatan').trim(),
        latarBelakang: fd.get('latarBelakang').trim(),
        tujuan: fd.get('tujuan').trim(),
        sasaran: fd.get('sasaran').trim(),
        indikatorKeberhasilan: fd.get('indikatorKeberhasilan').trim(),
        target: fd.get('target').trim(),
        bulanPelaksanaan: rktId ? (data.rkt.find(k => k.id === rktId)?.bulanPelaksanaan || [3, 8]) : [3, 8],
        penanggungJawab: fd.get('penanggungJawab').trim(),
        kebutuhanAnggaran: Number(fd.get('kebutuhanAnggaran')) || 0,
        sumberDana: fd.get('sumberDana'),
        buktiPelaksanaan: fd.get('buktiPelaksanaan').trim()
      };

      if (rktId) {
        const idx = data.rkt.findIndex(k => k.id === rktId);
        if (idx !== -1) data.rkt[idx] = obj;
      } else {
        data.rkt.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("RKT", "Simpan Kegiatan RKT", `Menyimpan kegiatan "${obj.namaKegiatan}".`);
      renderAll();
      showToast("Kegiatan RKT berhasil disimpan.", "success");
    }

    function deleteRkt(rktId) {
      showConfirmDialog({
        title: "Hapus Kegiatan RKT?",
        message: "Kegiatan RKT dan rincian anggaran ARKAS yang tertaut padanya akan dihapus.",
        confirmText: "Ya, Hapus",
        danger: true,
        onConfirm: () => {
          const data = getCurrentData();
          data.rkt = data.rkt.filter(k => k.id !== rktId);
          data.arkas = data.arkas.filter(a => a.rktId !== rktId);
          delete data.monev[rktId];
          triggerAutoSave("RKT", "Hapus Kegiatan RKT", "Menghapus kegiatan RKT beserta rincian anggaran terkait.");
          renderAll();
          showToast("Kegiatan RKT dihapus.", "info");
        }
      });
    }

    // =========================================================================
    // MODULE J: RANCANGAN ARKAS (TERHUBUNG DENGAN RKT & BOSP)
    // =========================================================================
    function syncRktToArkas() {
      const data = getCurrentData();
      if (data.rkt.length === 0) {
        showToast("Susun kegiatan RKT terlebih dahulu sebelum membuat Rancangan ARKAS.", "warning");
        return;
      }

      let added = 0;
      data.rkt.forEach(keg => {
        const hasLines = data.arkas.some(a => a.rktId === keg.id);
        if (!hasLines) {
          const est = Number(keg.kebutuhanAnggaran) || 1500000;
          data.arkas.push({
            id: uid(),
            rktId: keg.id,
            rkjmId: keg.rkjmId,
            indicatorId: keg.indicatorId,
            namaProgram: keg.namaProgram,
            namaKegiatan: keg.namaKegiatan,
            komponenPembiayaan: "Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
            uraianBelanja: `Penggandaan Modul/Bahan Ajar & ATK Pelaksanaan ${keg.namaKegiatan}`,
            volume: 1,
            satuan: "Paket",
            hargaSatuan: est,
            jumlahAnggaran: est,
            sumberDana: keg.sumberDana || "BOSP Reguler",
            waktuPelaksanaan: "Triwulan II & III",
            statusVerifikasi: "Perlu Verifikasi SSH/Kode Rekening"
          });
          added++;
        }
      });

      triggerAutoSave("ARKAS", "Buat Rancangan ARKAS", `Menurunkan ${added} kegiatan RKT ke rincian Rancangan ARKAS.`);
      navigateTo('arkas');
      showToast(added > 0 ? `${added} rancangan belanja ARKAS dibuat. Silakan sesuaikan volume, satuan, dan harga satuan sesuai SSH daerah.` : "Seluruh kegiatan RKT sudah memiliki rincian di Rancangan ARKAS.", "success");
    }

    function renderArkasModule(container) {
      const data = getCurrentData();
      const tahun = data.school.tahunPerencanaan || window.appState.activeYear;
      const totalAll = data.arkas.reduce((s, a) => s + (Number(a.jumlahAnggaran) || 0), 0);

      // Compute 3 Recaps: Per Program, Per Kegiatan, Per Komponen Pembiayaan
      const recapProgram = {};
      const recapKegiatan = {};
      const recapKomponen = {};

      data.arkas.forEach(item => {
        const amt = Number(item.jumlahAnggaran) || 0;
        recapProgram[item.namaProgram] = (recapProgram[item.namaProgram] || 0) + amt;
        recapKegiatan[item.namaKegiatan] = (recapKegiatan[item.namaKegiatan] || 0) + amt;
        recapKomponen[item.komponenPembiayaan] = (recapKomponen[item.komponenPembiayaan] || 0) + amt;
      });

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">8. Rancangan Anggaran ARKAS — Tahun Anggaran ${escapeHtml(tahun)}</h2>
            <p class="text-xs text-slate-500">Perencanaan anggaran terhubung langsung dengan kegiatan RKT dan mengacu pada Komponen Penggunaan Dana BOSP.</p>
          </div>
          <div class="flex flex-wrap items-center gap-2">
            <button onclick="syncRktToArkas()" class="px-3.5 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center gap-1.5 shadow-sm">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
              <span>Tarik dari RKT</span>
            </button>
            <button onclick="openEditArkasModal()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm">
              <i data-lucide="plus-circle" class="w-4 h-4"></i>
              <span>Tambah Uraian Belanja</span>
            </button>
          </div>
        </div>

        <!-- Disclaimer Banner (Mandatory Guardrail) -->
        <div class="bg-amber-50 border border-amber-300 rounded-2xl p-4 flex items-start gap-3 text-xs text-amber-950">
          <i data-lucide="shield-alert" class="w-5 h-5 text-amber-600 shrink-0 mt-0.5"></i>
          <div>
            <p class="font-bold">Catatan Kepatuhan ARKAS &amp; BOSP:</p>
            <p class="mt-0.5 leading-relaxed">
              Dokumen ini adalah <strong>Rancangan Kertas Kerja Anggaran</strong> berbasis Rapor Pendidikan sebagai bahan input manual ke aplikasi ARKAS resmi Kemendikdasmen. Aplikasi ini <strong>tidak membuat kode rekening secara sembarangan</strong> dan <strong>tidak terhubung/terkirim otomatis ke server ARKAS resmi</strong>. Bendahara BOS &amp; Kepala Sekolah wajib memverifikasi kecocokan kode rekening serta Standar Satuan Harga (SSH/SBU) Pemda setempat sebelum pengesahan.
            </p>
          </div>
        </div>

        <!-- Total & 3 Recap Cards -->
        <div class="grid grid-cols-1 lg:grid-cols-4 gap-4">
          <div class="emerald-gradient text-white rounded-2xl p-5 shadow-sm flex flex-col justify-between">
            <div>
              <div class="text-xs text-emerald-200 font-semibold">Rekap Total Rancangan Anggaran</div>
              <div class="text-2xl font-extrabold text-amber-300 mt-1 font-mono">${formatRupiah(totalAll)}</div>
            </div>
            <div class="mt-4 pt-3 border-t border-white/15 text-[11px] text-emerald-100">
              ${data.arkas.length} item belanja pada ${Object.keys(recapKegiatan).length} kegiatan RKT
            </div>
          </div>

          <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
            <div class="text-xs font-bold text-slate-800 mb-2 flex items-center justify-between">
              <span>Rekap per Program</span>
              <span class="text-[10px] text-slate-400">${Object.keys(recapProgram).length}</span>
            </div>
            <div class="space-y-1.5 max-h-28 overflow-y-auto text-[11px]">
              ${Object.entries(recapProgram).map(([k, v]) => `
                <div class="flex justify-between gap-2 border-b border-slate-100 pb-1">
                  <span class="truncate text-slate-600" title="${escapeHtml(k)}">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-slate-900 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join('') || '<span class="text-slate-400">Belum ada data</span>'}
            </div>
          </div>

          <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
            <div class="text-xs font-bold text-slate-800 mb-2 flex items-center justify-between">
              <span>Rekap per Kegiatan RKT</span>
              <span class="text-[10px] text-slate-400">${Object.keys(recapKegiatan).length}</span>
            </div>
            <div class="space-y-1.5 max-h-28 overflow-y-auto text-[11px]">
              ${Object.entries(recapKegiatan).map(([k, v]) => `
                <div class="flex justify-between gap-2 border-b border-slate-100 pb-1">
                  <span class="truncate text-slate-600" title="${escapeHtml(k)}">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-slate-900 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join('') || '<span class="text-slate-400">Belum ada data</span>'}
            </div>
          </div>

          <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
            <div class="text-xs font-bold text-slate-800 mb-2 flex items-center justify-between">
              <span>Rekap Komponen BOSP</span>
              <span class="text-[10px] text-slate-400">${Object.keys(recapKomponen).length}</span>
            </div>
            <div class="space-y-1.5 max-h-28 overflow-y-auto text-[11px]">
              ${Object.entries(recapKomponen).map(([k, v]) => `
                <div class="flex justify-between gap-2 border-b border-slate-100 pb-1">
                  <span class="truncate text-slate-600" title="${escapeHtml(k)}">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-emerald-800 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join('') || '<span class="text-slate-400">Belum ada data</span>'}
            </div>
          </div>
        </div>

        <!-- Tabel Rincian ARKAS -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          ${data.arkas.length > 0 ? `
            <div class="overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-emerald-950 text-white font-bold uppercase text-[10px]">
                    <th class="py-3 px-3">Program &amp; Kegiatan</th>
                    <th class="py-3 px-3">Komponen Pembiayaan BOSP</th>
                    <th class="py-3 px-3">Uraian Belanja</th>
                    <th class="py-3 px-3 text-center">Vol</th>
                    <th class="py-3 px-3 text-center">Satuan</th>
                    <th class="py-3 px-3 text-right">Harga Satuan</th>
                    <th class="py-3 px-3 text-right">Jumlah (Vol × Harga)</th>
                    <th class="py-3 px-3">Sumber &amp; Waktu</th>
                    <th class="py-3 px-3">Verifikasi Standar</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-200 align-top">
                  ${data.arkas.map(item => {
                    const isVerified = item.statusVerifikasi === 'Terverifikasi SSH & Juknis BOSP';
                    return `
                      <tr class="hover:bg-slate-50">
                        <td class="py-3 px-3 min-w-[170px]">
                          <div class="text-[10px] font-bold text-emerald-800">${escapeHtml(item.namaProgram)}</div>
                          <div class="font-semibold text-slate-900 mt-0.5">${escapeHtml(item.namaKegiatan)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[170px] text-slate-700">${escapeHtml(item.komponenPembiayaan)}</td>
                        <td class="py-3 px-3 min-w-[180px] font-medium text-slate-900">${escapeHtml(item.uraianBelanja)}</td>
                        <td class="py-3 px-3 text-center font-mono font-bold">${item.volume}</td>
                        <td class="py-3 px-3 text-center text-slate-600">${escapeHtml(item.satuan)}</td>
                        <td class="py-3 px-3 text-right font-mono">${formatRupiah(item.hargaSatuan)}</td>
                        <td class="py-3 px-3 text-right font-mono font-extrabold text-emerald-900">${formatRupiah(item.jumlahAnggaran)}</td>
                        <td class="py-3 px-3 whitespace-nowrap">
                          <span class="inline-block px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] font-semibold">${escapeHtml(item.sumberDana)}</span>
                          <div class="text-[11px] text-slate-500 mt-1">${escapeHtml(item.waktuPelaksanaan)}</div>
                        </td>
                        <td class="py-3 px-3">
                          <button onclick="toggleArkasVerify('${item.id}')" class="text-[10px] font-bold px-2 py-1 rounded-lg border transition ${isVerified ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-amber-50 text-amber-900 border-amber-300'}">
                            ${isVerified ? '✓ Terverifikasi SSH' : '⚠ Perlu Verifikasi'}
                          </button>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <button onclick="openEditArkasModal('${item.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit Belanja"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteArkas('${item.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
                <tfoot>
                  <tr class="bg-slate-100 font-extrabold text-slate-900">
                    <td colspan="6" class="py-3 px-4 text-right uppercase text-xs">Total Rancangan Anggaran ARKAS:</td>
                    <td class="py-3 px-3 text-right font-mono text-sm text-emerald-900">${formatRupiah(totalAll)}</td>
                    <td colspan="3"></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          ` : `
            <div class="p-10 text-center">
              <i data-lucide="wallet" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
              <h3 class="font-bold text-slate-900 text-base">Belum Ada Rincian Anggaran ARKAS</h3>
              <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Tarik daftar kegiatan dari Modul RKT atau tambahkan uraian belanja baru.</p>
              <div class="mt-4 flex justify-center gap-2">
                <button onclick="syncRktToArkas()" class="px-4 py-2 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold">Tarik dari RKT</button>
                <button onclick="openEditArkasModal()" class="px-4 py-2 rounded-xl border border-slate-300 text-xs font-bold">Tambah Manual</button>
              </div>
            </div>
          `}
        </div>
      `;
    }

    function toggleArkasVerify(arkasId) {
      const data = getCurrentData();
      const item = data.arkas.find(a => a.id === arkasId);
      if (item) {
        item.statusVerifikasi = item.statusVerifikasi === 'Terverifikasi SSH & Juknis BOSP'
          ? 'Perlu Verifikasi SSH/Kode Rekening'
          : 'Terverifikasi SSH & Juknis BOSP';
        triggerAutoSave("ARKAS", "Verifikasi Belanja", `Status verifikasi "${item.uraianBelanja}" diperbarui.`);
        renderAll();
      }
    }

    function openEditArkasModal(arkasId = null) {
      const data = getCurrentData();
      const existing = arkasId ? data.arkas.find(a => a.id === arkasId) : null;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">${existing ? 'Edit Rincian Belanja ARKAS' : 'Tambah Uraian Belanja ARKAS'}</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveArkasModal(event, '${arkasId || ''}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Kegiatan RKT Induk <span class="text-rose-600">*</span></label>
              <select name="rktId" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${data.rkt.length > 0
                  ? data.rkt.map(k => `<option value="${k.id}" ${existing?.rktId === k.id ? 'selected' : ''}>${escapeHtml(k.namaKegiatan)} — [${escapeHtml(k.namaProgram)}]</option>`).join('')
                  : `<option value="">-- Susun Kegiatan RKT Terlebih Dahulu --</option>`}
              </select>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Komponen Pembiayaan BOSP (Juknis Kemendikdasmen) <span class="text-rose-600">*</span></label>
              <select name="komponenPembiayaan" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${KOMPONEN_BOSP_STANDAR.map(c => `<option value="${c}" ${existing?.komponenPembiayaan === c ? 'selected' : ''}>${c}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Uraian Belanja Barang / Jasa <span class="text-rose-600">*</span></label>
              <input type="text" name="uraianBelanja" value="${escapeHtml(existing?.uraianBelanja || '')}" required placeholder="Contoh: Konsumsi Rapat Kombel Guru (Snack & Makan Siang)" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Volume <span class="text-rose-600">*</span></label>
                <input type="number" step="0.01" min="0.01" id="arkasVolInput" oninput="recalcArkasPreview()" name="volume" value="${existing?.volume || 1}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Satuan <span class="text-rose-600">*</span></label>
                <input type="text" name="satuan" value="${escapeHtml(existing?.satuan || 'Paket')}" required placeholder="OK / Paket / Rim / Buah" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Harga Satuan (Rp) <span class="text-rose-600">*</span></label>
                <input type="number" min="0" id="arkasPriceInput" oninput="recalcArkasPreview()" name="hargaSatuan" value="${existing?.hargaSatuan || 0}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div class="bg-emerald-50 border border-emerald-200 rounded-xl p-3 flex items-center justify-between">
              <span class="font-semibold text-emerald-900">Jumlah Anggaran Otomatis (Volume × Harga Satuan):</span>
              <span id="arkasTotalPreview" class="font-mono font-extrabold text-sm text-emerald-950">${formatRupiah((existing?.volume || 1) * (existing?.hargaSatuan || 0))}</span>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${["BOSP Reguler", "BOSP Kinerja", "APBD / Daerah", "Komite / Swadaya", "Lainnya"].map(sd => `<option value="${sd}" ${existing?.sumberDana === sd ? 'selected' : ''}>${sd}</option>`).join('')}
                </select>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Waktu Pelaksanaan</label>
                <input type="text" name="waktuPelaksanaan" value="${escapeHtml(existing?.waktuPelaksanaan || 'Triwulan II')}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Status Verifikasi Standar</label>
                <select name="statusVerifikasi" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  <option value="Perlu Verifikasi SSH/Kode Rekening" ${existing?.statusVerifikasi !== 'Terverifikasi SSH & Juknis BOSP' ? 'selected' : ''}>Perlu Verifikasi SSH/Kode Rekening</option>
                  <option value="Terverifikasi SSH & Juknis BOSP" ${existing?.statusVerifikasi === 'Terverifikasi SSH & Juknis BOSP' ? 'selected' : ''}>Terverifikasi SSH & Juknis BOSP</option>
                </select>
              </div>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Rincian ARKAS</button>
            </div>
          </form>
        </div>
      `, 'max-w-2xl');
    }

    function recalcArkasPreview() {
      const v = parseFloat(document.getElementById('arkasVolInput')?.value) || 0;
      const p = parseFloat(document.getElementById('arkasPriceInput')?.value) || 0;
      const el = document.getElementById('arkasTotalPreview');
      if (el) el.textContent = formatRupiah(v * p);
    }

    function saveArkasModal(e, arkasId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rktId = fd.get('rktId');
      const keg = data.rkt.find(k => k.id === rktId);
      const vol = parseFloat(fd.get('volume')) || 0;
      const harga = parseFloat(fd.get('hargaSatuan')) || 0;
      const jumlah = Math.round(vol * harga);

      const obj = {
        id: arkasId || uid(),
        rktId,
        rkjmId: keg?.rkjmId || null,
        indicatorId: keg?.indicatorId || null,
        namaProgram: keg?.namaProgram || 'Program Pengembangan Sekolah',
        namaKegiatan: keg?.namaKegiatan || 'Kegiatan RKT',
        komponenPembiayaan: fd.get('komponenPembiayaan'),
        uraianBelanja: fd.get('uraianBelanja').trim(),
        volume: vol,
        satuan: fd.get('satuan').trim(),
        hargaSatuan: harga,
        jumlahAnggaran: jumlah,
        sumberDana: fd.get('sumberDana'),
        waktuPelaksanaan: fd.get('waktuPelaksanaan').trim(),
        statusVerifikasi: fd.get('statusVerifikasi')
      };

      if (arkasId) {
        const idx = data.arkas.findIndex(a => a.id === arkasId);
        if (idx !== -1) data.arkas[idx] = obj;
      } else {
        data.arkas.push(obj);
      }

      // Sync RKT effective budget
      if (keg) {
        keg.kebutuhanAnggaran = getRktEffectiveBudget(keg);
      }

      closeCustomModal();
      triggerAutoSave("ARKAS", "Simpan Uraian Belanja", `Menyimpan belanja "${obj.uraianBelanja}" sebesar ${formatRupiah(jumlah)}.`);
      renderAll();
      showToast("Rincian belanja ARKAS berhasil disimpan.", "success");
    }

    function deleteArkas(arkasId) {
      showConfirmDialog({
        title: "Hapus Uraian Belanja ARKAS?",
        message: "Baris uraian belanja ini akan dihapus dari Rancangan ARKAS.",
        confirmText: "Ya, Hapus",
        danger: true,
        onConfirm: () => {
          const data = getCurrentData();
          data.arkas = data.arkas.filter(a => a.id !== arkasId);
          triggerAutoSave("ARKAS", "Hapus Belanja", "Menghapus baris uraian belanja ARKAS.");
          renderAll();
          showToast("Uraian belanja dihapus.", "info");
        }
      });
    }

    // =========================================================================
    // MODULE K: HUBUNGAN ANTARDOKUMEN (BENANG MERAH PENELUSURAN END-TO-END)
    // =========================================================================
    function renderKeterhubunganModule(container) {
      const data = getCurrentData();
      const inds = data.indicators;

      if (inds.length === 0) {
        container.innerHTML = `
          <div class="bg-white rounded-2xl border border-slate-200 p-10 text-center">
            <i data-lucide="network" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Peta Keterhubungan Belum Tersedia</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Masukkan data Rapor Pendidikan terlebih dahulu untuk menelusuri keterhubungan antar dokumen.</p>
            <button onclick="navigateTo('input')" class="mt-4 px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Ke Modul 1: Input Data</button>
          </div>
        `;
        return;
      }

      if (!window.appState.selectedTraceIndicatorId || !inds.some(i => i.id === window.appState.selectedTraceIndicatorId)) {
        const prioFirst = inds.find(needsImprovement);
        window.appState.selectedTraceIndicatorId = prioFirst ? prioFirst.id : inds[0].id;
      }

      const activeInd = inds.find(i => i.id === window.appState.selectedTraceIndicatorId);
      const rc = data.rootCauses[activeInd.id];
      const ranked = getRankedPriorities();
      const rankEntry = ranked.findIndex(r => r.ind.id === activeInd.id);
      const prioData = data.priorities[activeInd.id];
      const linkedRecs = data.recommendations.filter(r => r.indicatorId === activeInd.id);
      const linkedRkjm = data.rkjm.filter(p => p.indicatorId === activeInd.id);
      const linkedRkt = data.rkt.filter(k => k.indicatorId === activeInd.id || linkedRkjm.some(p => p.id === k.rkjmId));
      const linkedArkas = data.arkas.filter(a => a.indicatorId === activeInd.id || linkedRkt.some(k => k.id === a.rktId));
      const totalLinkedBudget = linkedArkas.reduce((s, a) => s + (Number(a.jumlahAnggaran) || 0), 0);

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">9. Hubungan Antardokumen (Benang Merah Perencanaan)</h2>
            <p class="text-xs text-slate-500">Penelusuran utuh: Rapor Pendidikan → Analisis → Akar Masalah → Prioritas → Rekomendasi → RKJM → RKT → ARKAS.</p>
          </div>
        </div>

        <!-- Visual Flowchart Banner -->
        <div class="bg-emerald-950 text-white rounded-2xl p-4 overflow-x-auto">
          <div class="flex items-center justify-between min-w-[740px] gap-2 text-[11px] font-bold">
            <div class="px-3 py-2 rounded-xl bg-white/10 border border-amber-400/40 text-amber-300 text-center flex-1">1. Rapor Pendidikan</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">2. Analisis Masalah</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">3. Akar Masalah</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">4. Prioritas</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">5. Rekomendasi</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">6. RKJM</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl bg-white/10 text-center flex-1">7. RKT</div>
            <i data-lucide="arrow-right" class="w-4 h-4 text-amber-400 shrink-0"></i>
            <div class="px-3 py-2 rounded-xl gold-gradient text-emerald-950 font-extrabold text-center flex-1">8. ARKAS</div>
          </div>
        </div>

        <!-- Selector Indikator untuk Ditelusuri -->
        <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
          <label class="block text-xs font-bold text-slate-700 mb-2">Pilih Indikator Rapor Pendidikan untuk Menelusuri Keterhubungannya:</label>
          <select onchange="window.appState.selectedTraceIndicatorId=this.value; renderActiveModule(); if(window.lucide) window.lucide.createIcons();" class="w-full sm:max-w-xl rounded-xl border border-slate-300 px-3.5 py-2.5 text-xs font-bold text-slate-900">
            ${inds.map(i => `<option value="${i.id}" ${i.id === activeInd.id ? 'selected' : ''}>${i.kode ? '[' + i.kode + '] ' : ''}${escapeHtml(i.nama)} — Nilai: ${i.nilai || '-'} (${classifyCategory(i.nilai, i.kategori)})</option>`).join('')}
          </select>
        </div>

        <!-- Vertical Traceability Tree -->
        <div class="space-y-4">
          <!-- Node 1 & 2: Rapor & Analisis -->
          <div class="bg-white rounded-2xl border-l-4 border-emerald-700 border border-slate-200 p-5 shadow-xs">
            <div class="text-[11px] font-extrabold uppercase text-emerald-800">Tahap 1 &amp; 2 — Capaian Rapor Pendidikan &amp; Status Analisis</div>
            <div class="mt-1 flex flex-wrap items-center justify-between gap-2">
              <h3 class="font-extrabold text-slate-900 text-base">${activeInd.kode ? `[${escapeHtml(activeInd.kode)}] ` : ''}${escapeHtml(activeInd.nama)}</h3>
              <div class="flex items-center gap-2">
                <span class="font-mono font-bold text-sm">Nilai: ${activeInd.nilai || '-'}</span>
                ${getCategoryBadgeHtml(classifyCategory(activeInd.nilai, activeInd.kategori))}
                ${getDeltaBadgeHtml(activeInd.perubahan)}
              </div>
            </div>
            <p class="text-xs text-slate-500 mt-1">${escapeHtml(activeInd.dimensi)}</p>
          </div>

          <!-- Node 3 & 4: Akar Masalah & Prioritas -->
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="bg-white rounded-2xl border-l-4 border-amber-500 border border-slate-200 p-5 shadow-xs">
              <div class="text-[11px] font-extrabold uppercase text-amber-800">Tahap 3 — Akar Masalah Teridentifikasi</div>
              ${rc ? `
                <p class="text-xs font-semibold text-slate-900 mt-1.5">${escapeHtml(rc.masalahUtama)}</p>
                <ul class="mt-2 space-y-1 text-xs text-slate-600 list-disc pl-4">
                  ${(rc.akarPenyebab || []).map(a => `<li>${escapeHtml(a.teks)} <span class="text-[10px] font-bold">(${a.status === 'fakta_sekolah' ? 'Fakta Sekolah' : 'Hipotesis AI'})</span></li>`).join('')}
                </ul>
              ` : `<p class="text-xs text-rose-600 mt-1.5 italic">Belum dilakukan analisis akar masalah untuk indikator ini.</p>`}
            </div>

            <div class="bg-white rounded-2xl border-l-4 border-amber-600 border border-slate-200 p-5 shadow-xs">
              <div class="text-[11px] font-extrabold uppercase text-amber-800">Tahap 4 — Peringkat Prioritas Pembenahan</div>
              ${rankEntry !== -1 && prioData ? `
                <div class="mt-1.5 flex items-center gap-3">
                  <span class="px-3 py-1 rounded-xl bg-amber-100 text-amber-950 font-extrabold text-sm">Peringkat #${rankEntry + 1}</span>
                  <span class="text-xs font-bold text-slate-700">Skor Multi-Kriteria: ${prioData.totalSkor} / 25</span>
                </div>
                <p class="text-xs text-slate-600 mt-2">${escapeHtml(prioData.alasanPrioritas)}</p>
              ` : `<p class="text-xs text-slate-500 mt-1.5">Indikator ini berada pada kategori Baik / tidak masuk daftar prioritas pembenahan.</p>`}
            </div>
          </div>

          <!-- Node 5 & 6: Rekomendasi & RKJM -->
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="bg-white rounded-2xl border-l-4 border-emerald-600 border border-slate-200 p-5 shadow-xs">
              <div class="text-[11px] font-extrabold uppercase text-emerald-800">Tahap 5 — Rekomendasi Program (${linkedRecs.length})</div>
              ${linkedRecs.length > 0 ? linkedRecs.map(r => `
                <div class="mt-2 p-3 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                  <div class="font-bold text-slate-900">${escapeHtml(r.bentukKegiatan)}</div>
                  <div class="text-slate-600 mt-0.5">Target: ${escapeHtml(r.targetCapaian)}</div>
                </div>
              `).join('') : `<p class="text-xs text-rose-600 mt-1.5 italic">Belum ada rekomendasi program tertaut.</p>`}
            </div>

            <div class="bg-white rounded-2xl border-l-4 border-emerald-700 border border-slate-200 p-5 shadow-xs">
              <div class="text-[11px] font-extrabold uppercase text-emerald-800">Tahap 6 — Program RKJM (${linkedRkjm.length})</div>
              ${linkedRkjm.length > 0 ? linkedRkjm.map(p => `
                <div class="mt-2 p-3 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                  <div class="font-bold text-slate-900">${escapeHtml(p.namaProgram)}</div>
                  <div class="text-slate-600 mt-0.5">Tujuan Strategis: ${escapeHtml(p.tujuanStrategis)}</div>
                </div>
              `).join('') : `<p class="text-xs text-rose-600 mt-1.5 italic">Belum diturunkan ke RKJM.</p>`}
            </div>
          </div>

          <!-- Node 7 & 8: RKT & ARKAS -->
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="bg-white rounded-2xl border-l-4 border-emerald-800 border border-slate-200 p-5 shadow-xs">
              <div class="text-[11px] font-extrabold uppercase text-emerald-900">Tahap 7 — Kegiatan Tahunan RKT (${linkedRkt.length})</div>
              ${linkedRkt.length > 0 ? linkedRkt.map(k => `
                <div class="mt-2 p-3 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                  <div class="font-bold text-slate-900">${escapeHtml(k.namaKegiatan)}</div>
                  <div class="text-slate-600 mt-0.5">PJ: ${escapeHtml(k.penanggungJawab)} • Sumber: ${escapeHtml(k.sumberDana)}</div>
                </div>
              `).join('') : `<p class="text-xs text-rose-600 mt-1.5 italic">Belum diturunkan ke kegiatan RKT.</p>`}
            </div>

            <div class="bg-white rounded-2xl border-l-4 border-amber-500 border border-slate-200 p-5 shadow-xs">
              <div class="flex items-center justify-between">
                <div class="text-[11px] font-extrabold uppercase text-amber-800">Tahap 8 — Rancangan Belanja ARKAS (${linkedArkas.length} item)</div>
                <span class="font-mono font-extrabold text-xs text-emerald-900">${formatRupiah(totalLinkedBudget)}</span>
              </div>
              ${linkedArkas.length > 0 ? linkedArkas.map(a => `
                <div class="mt-2 p-3 rounded-xl bg-amber-50/50 border border-amber-200/70 text-xs flex justify-between gap-2">
                  <div>
                    <div class="font-bold text-slate-900">${escapeHtml(a.uraianBelanja)}</div>
                    <div class="text-[11px] text-slate-500">${a.volume} ${escapeHtml(a.satuan)} × ${formatRupiah(a.hargaSatuan)}</div>
                  </div>
                  <div class="font-mono font-bold text-emerald-900 shrink-0">${formatRupiah(a.jumlahAnggaran)}</div>
                </div>
              `).join('') : `<p class="text-xs text-rose-600 mt-1.5 italic">Belum memiliki rincian anggaran di ARKAS.</p>`}
            </div>
          </div>
        </div>
      `;
    }

    // =========================================================================
    // MODULE L: MONITORING DAN EVALUASI (MONEV)
    // =========================================================================
    function renderMonevModule(container) {
      const data = getCurrentData();
      const rkList = data.rkt;

      if (rkList.length === 0) {
        container.innerHTML = `
          <div class="bg-white rounded-2xl border border-slate-200 p-10 text-center">
            <i data-lucide="activity" class="w-12 h-12 text-slate-300 mx-auto mb-3"></i>
            <h3 class="font-bold text-slate-900 text-base">Belum Ada Kegiatan RKT untuk Dipantau</h3>
            <p class="text-xs text-slate-500 max-w-md mx-auto mt-1">Susun kegiatan RKT terlebih dahulu agar Tim Pengembang Sekolah dapat memantau progres pelaksanaan dan serapan anggarannya.</p>
            <button onclick="navigateTo('rkt')" class="mt-4 px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Ke Modul 7: Penyusunan RKT</button>
          </div>
        `;
        return;
      }

      // Ensure monev entries exist
      rkList.forEach(k => {
        if (!data.monev[k.id]) {
          data.monev[k.id] = {
            rktId: k.id,
            statusPelaksanaan: 'belum',
            persentaseKegiatan: 0,
            realisasiAnggaran: 0,
            hambatan: '',
            buktiKegiatan: '',
            evaluasiHasil: '',
            tindakLanjut: ''
          };
        }
      });

      const belumCount = rkList.filter(k => data.monev[k.id].statusPelaksanaan === 'belum').length;
      const berjalanCount = rkList.filter(k => data.monev[k.id].statusPelaksanaan === 'berjalan').length;
      const selesaiCount = rkList.filter(k => data.monev[k.id].statusPelaksanaan === 'selesai').length;
      const avgProgress = Math.round(rkList.reduce((s, k) => s + (Number(data.monev[k.id].persentaseKegiatan) || 0), 0) / rkList.length);
      const totalPagu = rkList.reduce((s, k) => s + getRktEffectiveBudget(k), 0);
      const totalRealisasi = rkList.reduce((s, k) => s + (Number(data.monev[k.id].realisasiAnggaran) || 0), 0);
      const pctSerapan = totalPagu > 0 ? Math.round((totalRealisasi / totalPagu) * 100) : 0;

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">10. Monitoring dan Evaluasi (Monev) Program</h2>
            <p class="text-xs text-slate-500">Pemantauan keterlaksanaan kegiatan RKT, serapan anggaran ARKAS, hambatan, dan tindak lanjut perbaikan.</p>
          </div>
          <button onclick="exportMonevPdf()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center gap-1.5 shadow-sm">
            <i data-lucide="file-down" class="w-4 h-4"></i>
            <span>Unduh Laporan Monev (PDF)</span>
          </button>
        </div>

        <!-- Status Summary Cards -->
        <div class="grid grid-cols-2 lg:grid-cols-5 gap-3">
          <div class="bg-white rounded-2xl p-4 border border-slate-200">
            <div class="text-[11px] font-semibold text-slate-500">Belum Dilaksanakan</div>
            <div class="text-2xl font-extrabold text-slate-700 mt-1">${belumCount}</div>
          </div>
          <div class="bg-white rounded-2xl p-4 border border-amber-200 bg-amber-50/40">
            <div class="text-[11px] font-semibold text-amber-800">Sedang Berjalan</div>
            <div class="text-2xl font-extrabold text-amber-700 mt-1">${berjalanCount}</div>
          </div>
          <div class="bg-white rounded-2xl p-4 border border-emerald-200 bg-emerald-50/40">
            <div class="text-[11px] font-semibold text-emerald-800">Program Selesai</div>
            <div class="text-2xl font-extrabold text-emerald-800 mt-1">${selesaiCount}</div>
          </div>
          <div class="bg-white rounded-2xl p-4 border border-slate-200">
            <div class="text-[11px] font-semibold text-slate-500">Rata-rata Realisasi Fisik</div>
            <div class="text-2xl font-extrabold text-emerald-900 mt-1">${avgProgress}%</div>
          </div>
          <div class="col-span-2 lg:col-span-1 bg-emerald-950 text-white rounded-2xl p-4">
            <div class="text-[11px] font-semibold text-emerald-200">Realisasi Anggaran</div>
            <div class="text-xl font-extrabold text-amber-300 mt-1">${pctSerapan}%</div>
            <div class="text-[10px] text-emerald-200 truncate">${formatRupiah(totalRealisasi)} / ${formatRupiah(totalPagu)}</div>
          </div>
        </div>

        <!-- Grafik Perkembangan Program -->
        <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
          <h3 class="font-bold text-slate-900 text-sm mb-3">Grafik Perkembangan Realisasi Kegiatan vs Serapan Anggaran (%)</h3>
          <div class="relative h-64">
            <canvas id="monevProgressChart"></canvas>
          </div>
        </div>

        <!-- Tabel Monitoring & Evaluasi Interaktif -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse text-xs">
              <thead>
                <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase text-[10px]">
                  <th class="py-3 px-4">Kegiatan RKT</th>
                  <th class="py-3 px-3 text-center">Status</th>
                  <th class="py-3 px-3 text-center">Realisasi Kegiatan</th>
                  <th class="py-3 px-3 text-right">Realisasi Anggaran</th>
                  <th class="py-3 px-3">Hambatan &amp; Bukti</th>
                  <th class="py-3 px-3">Evaluasi &amp; Tindak Lanjut</th>
                  <th class="py-3 px-4 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100 align-top">
                ${rkList.map(k => {
                  const m = data.monev[k.id];
                  const pagu = getRktEffectiveBudget(k);
                  const pctAng = pagu > 0 ? Math.round(((Number(m.realisasiAnggaran) || 0) / pagu) * 100) : 0;
                  const statusBadges = {
                    belum: `<span class="px-2.5 py-1 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600">Belum Dilaksanakan</span>`,
                    berjalan: `<span class="px-2.5 py-1 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800">Sedang Berjalan</span>`,
                    selesai: `<span class="px-2.5 py-1 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">Selesai</span>`
                  };
                  return `
                    <tr class="hover:bg-slate-50">
                      <td class="py-3 px-4 min-w-[180px]">
                        <div class="font-bold text-slate-900">${escapeHtml(k.namaKegiatan)}</div>
                        <div class="text-[11px] text-slate-500 mt-0.5">PJ: ${escapeHtml(k.penanggungJawab)}</div>
                      </td>
                      <td class="py-3 px-3 text-center whitespace-nowrap">${statusBadges[m.statusPelaksanaan] || statusBadges.belum}</td>
                      <td class="py-3 px-3 text-center min-w-[120px]">
                        <div class="font-mono font-bold text-slate-900">${m.persentaseKegiatan}%</div>
                        <div class="w-full h-1.5 bg-slate-200 rounded-full mt-1 overflow-hidden">
                          <div class="h-full bg-emerald-700 rounded-full" style="width:${Math.min(100, m.persentaseKegiatan)}%"></div>
                        </div>
                      </td>
                      <td class="py-3 px-3 text-right whitespace-nowrap">
                        <div class="font-mono font-bold text-slate-900">${formatRupiah(m.realisasiAnggaran)}</div>
                        <div class="text-[10px] text-slate-500">${pctAng}% dari ${formatRupiah(pagu)}</div>
                      </td>
                      <td class="py-3 px-3 min-w-[180px] text-slate-700">
                        <div><span class="font-semibold">Hambatan:</span> ${escapeHtml(m.hambatan || '-')}</div>
                        <div class="mt-1"><span class="font-semibold">Bukti:</span> ${escapeHtml(m.buktiKegiatan || '-')}</div>
                      </td>
                      <td class="py-3 px-3 min-w-[180px] text-slate-700">
                        <div><span class="font-semibold">Evaluasi:</span> ${escapeHtml(m.evaluasiHasil || '-')}</div>
                        <div class="mt-1 text-emerald-800"><span class="font-semibold">Tindak Lanjut:</span> ${escapeHtml(m.tindakLanjut || '-')}</div>
                      </td>
                      <td class="py-3 px-4 text-right whitespace-nowrap">
                        <button onclick="openEditMonevModal('${k.id}')" class="px-3 py-1.5 rounded-xl emerald-gradient text-white text-[11px] font-bold">
                          Update Monev
                        </button>
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;

      // Render Monev Chart
      const ctx = document.getElementById('monevProgressChart');
      if (ctx && window.Chart) {
        const labels = rkList.map(k => k.namaKegiatan.length > 24 ? k.namaKegiatan.slice(0, 24) + '…' : k.namaKegiatan);
        const fisikData = rkList.map(k => Number(data.monev[k.id].persentaseKegiatan) || 0);
        const angData = rkList.map(k => {
          const pagu = getRktEffectiveBudget(k);
          return pagu > 0 ? Math.min(100, Math.round(((Number(data.monev[k.id].realisasiAnggaran) || 0) / pagu) * 100)) : 0;
        });

        chartInstances.monevBar = new Chart(ctx, {
          type: 'bar',
          data: {
            labels,
            datasets: [
              { label: 'Realisasi Kegiatan (%)', data: fisikData, backgroundColor: '#047857', borderRadius: 5 },
              { label: 'Realisasi Anggaran (%)', data: angData, backgroundColor: '#f59e0b', borderRadius: 5 }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              y: { beginAtZero: true, max: 100 }
            }
          }
        });
      }
    }

    function openEditMonevModal(rktId) {
      const data = getCurrentData();
      const keg = data.rkt.find(k => k.id === rktId);
      const m = data.monev[rktId];
      const pagu = getRktEffectiveBudget(keg);

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Update Monitoring &amp; Evaluasi</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <p class="text-xs font-bold text-emerald-800 mt-2">${escapeHtml(keg.namaKegiatan)} (Pagu Anggaran: ${formatRupiah(pagu)})</p>
          <form onsubmit="saveMonevModal(event, '${rktId}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Status Pelaksanaan</label>
                <select name="statusPelaksanaan" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  <option value="belum" ${m.statusPelaksanaan === 'belum' ? 'selected' : ''}>Belum Dilaksanakan</option>
                  <option value="berjalan" ${m.statusPelaksanaan === 'berjalan' ? 'selected' : ''}>Sedang Berjalan</option>
                  <option value="selesai" ${m.statusPelaksanaan === 'selesai' ? 'selected' : ''}>Selesai</option>
                </select>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Realisasi Kegiatan (%)</label>
                <input type="number" min="0" max="100" name="persentaseKegiatan" value="${m.persentaseKegiatan}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Realisasi Anggaran (Rp)</label>
                <input type="number" min="0" name="realisasiAnggaran" value="${m.realisasiAnggaran}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Hambatan Pelaksanaan</label>
              <textarea name="hambatan" rows="2" placeholder="Catat kendala waktu, teknis, atau partisipasi..." class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.hambatan)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Bukti Kegiatan (Nomor SPJ / Tautan / Dokumen Fisik)</label>
              <input type="text" name="buktiKegiatan" value="${escapeHtml(m.buktiKegiatan)}" placeholder="Contoh: Laporan Kegiatan No. 04/2026 & SPJ BOSP Tw 2" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Evaluasi Hasil (Dampak terhadap Mutu Pembelajaran)</label>
              <textarea name="evaluasiHasil" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.evaluasiHasil)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Rekomendasi Tindak Lanjut</label>
              <textarea name="tindakLanjut" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.tindakLanjut)}</textarea>
            </div>
            <div class="pt-3 flex justify-end gap-2 border-t border-slate-100">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Laporan Monev</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveMonevModal(e, rktId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      data.monev[rktId] = {
        rktId,
        statusPelaksanaan: fd.get('statusPelaksanaan'),
        persentaseKegiatan: Number(fd.get('persentaseKegiatan')) || 0,
        realisasiAnggaran: Number(fd.get('realisasiAnggaran')) || 0,
        hambatan: fd.get('hambatan').trim(),
        buktiKegiatan: fd.get('buktiKegiatan').trim(),
        evaluasiHasil: fd.get('evaluasiHasil').trim(),
        tindakLanjut: fd.get('tindakLanjut').trim()
      };
      closeCustomModal();
      triggerAutoSave("Monitoring & Evaluasi", "Update Progres", "Memperbarui realisasi kegiatan dan anggaran Monev.");
      renderAll();
      showToast("Data Monitoring & Evaluasi berhasil disimpan.", "success");
    }

    // =========================================================================
    // MODULE M: LAPORAN DAN EKSPOR (PDF, WORD .DOC, EXCEL .XLSX)
    // =========================================================================
    function renderEksporModule(container) {
      const data = getCurrentData();
      const s = data.school;

      const exportCards = [
        {
          no: 1,
          title: "Hasil Analisis Rapor Pendidikan",
          format: "PDF (.pdf)",
          badgeClass: "bg-rose-100 text-rose-800",
          desc: "Laporan lengkap capaian seluruh indikator, kategori, tren perubahan tahun lalu, dan status prioritas.",
          btnLabel: "Unduh PDF Analisis",
          action: "exportAnalisisPdf()"
        },
        {
          no: 2,
          title: "Dokumen Analisis Akar Masalah",
          format: "Word (.doc)",
          badgeClass: "bg-blue-100 text-blue-800",
          desc: "Dokumen naratif & tabel bedah 5-Why, hipotesis vs fakta sekolah, pertanyaan refleksi guru, dan solusi.",
          btnLabel: "Unduh Word Akar Masalah",
          action: "exportAkarMasalahWord()"
        },
        {
          no: 3,
          title: "Daftar Prioritas Pembenahan",
          format: "Excel (.xlsx)",
          badgeClass: "bg-emerald-100 text-emerald-800",
          desc: "Tabel pemeringkatan 5 kriteria prioritas beserta skor total dan alasan penetapan prioritas.",
          btnLabel: "Unduh Excel Prioritas",
          action: "exportPrioritasExcel()"
        },
        {
          no: 4,
          title: "Dokumen Rencana Kerja Jangka Menengah (RKJM)",
          format: "Word & Excel",
          badgeClass: "bg-amber-100 text-amber-900",
          desc: "Matriks 4 tahunan yang menghubungkan Dimensi Rapor, Akar Masalah, Tujuan Strategis, dan Target Tahunan.",
          dual: true,
          wordAction: "exportRkjmWord()",
          excelAction: "exportRkjmExcel()"
        },
        {
          no: 5,
          title: "Dokumen Rencana Kerja Tahunan (RKT)",
          format: "Word & Excel",
          badgeClass: "bg-amber-100 text-amber-900",
          desc: "Rincian kegiatan tahunan, tujuan, sasaran, jadwal bulanan, penanggung jawab, dan sumber dana.",
          dual: true,
          wordAction: "exportRktWord()",
          excelAction: "exportRktExcel()"
        },
        {
          no: 6,
          title: "Rancangan Anggaran ARKAS (Kertas Kerja)",
          format: "Excel (.xlsx)",
          badgeClass: "bg-emerald-100 text-emerald-800",
          desc: "Rincian belanja per program, kegiatan, komponen BOSP, volume, satuan, harga satuan, dan rekapitulasi.",
          btnLabel: "Unduh Excel ARKAS",
          action: "exportArkasExcel()"
        },
        {
          no: 7,
          title: "Laporan Monitoring dan Evaluasi (Monev)",
          format: "PDF (.pdf)",
          badgeClass: "bg-rose-100 text-rose-800",
          desc: "Laporan capaian fisik kegiatan, realisasi serapan anggaran, hambatan, evaluasi hasil, dan tindak lanjut.",
          btnLabel: "Unduh PDF Monev",
          action: "exportMonevPdf()"
        }
      ];

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">11. Pusat Laporan &amp; Ekspor Dokumen Resmi</h2>
            <p class="text-xs text-slate-500">Seluruh dokumen dilengkapi kop identitas sekolah, tahun perencanaan, tabel rapi, dan penomoran halaman.</p>
          </div>
          <button onclick="window.print()" class="px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold flex items-center gap-1.5">
            <i data-lucide="printer" class="w-4 h-4"></i>
            <span>Cetak Layar</span>
          </button>
        </div>

        ${(!s.nama || !s.npsn) ? `
          <div class="bg-amber-50 border border-amber-300 rounded-2xl p-4 text-xs text-amber-950 flex items-center justify-between gap-3">
            <span><strong>Perhatian:</strong> Nama Sekolah dan NPSN belum dilengkapi. Lengkapi terlebih dahulu agar tercetak rapi pada header dokumen ekspor.</span>
            <button onclick="navigateTo('input')" class="shrink-0 px-3 py-1.5 rounded-lg bg-amber-600 text-white font-bold">Lengkapi Identitas</button>
          </div>
        ` : ''}

        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          ${exportCards.map(c => `
            <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
              <div>
                <div class="flex items-center justify-between mb-3">
                  <span class="w-7 h-7 rounded-lg bg-emerald-900 text-amber-300 font-bold text-xs flex items-center justify-center">${c.no}</span>
                  <span class="text-[11px] font-bold px-2.5 py-0.5 rounded-full ${c.badgeClass}">${c.format}</span>
                </div>
                <h3 class="font-bold text-slate-900 text-sm">${c.title}</h3>
                <p class="text-xs text-slate-500 mt-1.5 leading-relaxed">${c.desc}</p>
              </div>
              <div class="mt-5 pt-3 border-t border-slate-100">
                ${c.dual ? `
                  <div class="grid grid-cols-2 gap-2">
                    <button onclick="${c.wordAction}" class="py-2 px-3 rounded-xl bg-blue-900 hover:bg-blue-800 text-white text-xs font-bold flex items-center justify-center gap-1.5 transition">
                      <i data-lucide="file-text" class="w-3.5 h-3.5"></i> Word
                    </button>
                    <button onclick="${c.excelAction}" class="py-2 px-3 rounded-xl emerald-gradient hover:opacity-95 text-white text-xs font-bold flex items-center justify-center gap-1.5 transition">
                      <i data-lucide="file-spreadsheet" class="w-3.5 h-3.5"></i> Excel
                    </button>
                  </div>
                ` : `
                  <button onclick="${c.action}" class="w-full py-2.5 px-4 rounded-xl emerald-gradient hover:opacity-95 text-white text-xs font-bold flex items-center justify-center gap-1.5 transition">
                    <i data-lucide="download" class="w-4 h-4 text-amber-300"></i>
                    <span>${c.btnLabel}</span>
                  </button>
                `}
              </div>
            </div>
          `).join('')}
        </div>
      `;
    }

    // Helper: Official School Header HTML for Word & PDF Exports
    function buildOfficialDocumentHeader(docTitle) {
      const data = getCurrentData();
      const s = data.school;
      const demoNote = window.appState.isDemoMode
        ? `<div style="background:#fef3c7;border:1px solid #f59e0b;padding:6px 10px;font-size:10pt;font-weight:bold;color:#78350f;margin-bottom:10px;">CATATAN: DOKUMEN INI MENGGUNAKAN DATA SIMULASI / CONTOH PELATIHAN</div>`
        : '';
      return `
        ${demoNote}
        <div style="border-bottom:3px solid #064e3b;padding-bottom:10px;margin-bottom:14px;font-family:Arial,sans-serif;">
          <div style="font-size:14pt;font-weight:bold;color:#064e3b;text-transform:uppercase;">${escapeHtml(s.nama || 'SATUAN PENDIDIKAN (NAMA SEKOLAH BELUM DIISI)')}</div>
          <div style="font-size:9.5pt;color:#334155;margin-top:2px;">
            NPSN: <strong>${escapeHtml(s.npsn || '-')}</strong> | Jenjang: <strong>${escapeHtml(s.jenjang || '-')}</strong> | Alamat: ${escapeHtml(s.alamat || '-')}${s.kecamatan ? ', Kec. ' + escapeHtml(s.kecamatan) : ''}${s.kabupaten ? ', ' + escapeHtml(s.kabupaten) : ''}
          </div>
          <div style="font-size:9.5pt;color:#334155;margin-top:2px;">
            Tahun Rapor Pendidikan: <strong>${escapeHtml(s.tahunRapor || '-')}</strong> | Tahun Perencanaan / Anggaran: <strong>${escapeHtml(s.tahunPerencanaan || window.appState.activeYear)}</strong>
          </div>
        </div>
        <h2 style="font-family:Arial,sans-serif;font-size:13pt;color:#0f172a;text-align:center;margin:10px 0 14px 0;text-transform:uppercase;">${escapeHtml(docTitle)}</h2>
      `;
    }

    function downloadWordHtml(filename, innerHtml) {
      const fullDoc = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>${escapeHtml(filename)}</title>
          <style>
            @page { size: A4 landscape; margin: 1.5cm; }
            body { font-family: Arial, sans-serif; font-size: 10pt; color: #1e293b; }
            table { width: 100%; border-collapse: collapse; margin-top: 8px; }
            th { background-color: #064e3b; color: #ffffff; font-weight: bold; padding: 6px 8px; border: 1px solid #cbd5e1; font-size: 9pt; text-align: left; }
            td { padding: 6px 8px; border: 1px solid #cbd5e1; font-size: 9pt; vertical-align: top; }
          </style>
        </head>
        <body>
          ${innerHtml}
          <div style="margin-top:18px;font-size:8.5pt;color:#64748b;text-align:right;">
            Dicetak melalui RAPOR SMART AI • Tanggal: ${new Date().toLocaleDateString('id-ID')} • Halaman 1
          </div>
        </body>
        </html>
      `;
      const blob = new Blob(['\ufeff', fullDoc], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast(`Dokumen Word "${filename}" berhasil diunduh.`, "success");
    }

    function downloadPdfFromHtml(filename, innerHtml) {
      if (!window.html2pdf) {
        showToast("Pustaka PDF sedang dimuat, silakan coba sesaat lagi.", "warning");
        return;
      }
      const wrapper = document.createElement('div');
      wrapper.style.padding = '16px';
      wrapper.style.fontFamily = 'Arial, sans-serif';
      wrapper.style.color = '#0f172a';
      wrapper.innerHTML = `
        <style>
          table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 8.5pt; }
          th { background-color: #064e3b; color: #ffffff; padding: 6px; border: 1px solid #94a3b8; text-align: left; }
          td { padding: 6px; border: 1px solid #cbd5e1; vertical-align: top; }
        </style>
        ${innerHtml}
      `;

      const opt = {
        margin: [10, 10, 14, 10],
        filename: filename,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
      };

      html2pdf().set(opt).from(wrapper).toPdf().get('pdf').then(function (pdf) {
        const totalPages = pdf.internal.getNumberOfPages();
        for (let i = 1; i <= totalPages; i++) {
          pdf.setPage(i);
          pdf.setFontSize(8);
          pdf.setTextColor(100);
          pdf.text(`RAPOR SMART AI — Halaman ${i} dari ${totalPages}`, pdf.internal.pageSize.getWidth() - 65, pdf.internal.pageSize.getHeight() - 6);
        }
      }).save().then(() => {
        showToast(`Dokumen PDF "${filename}" berhasil diunduh.`, "success");
      });
    }

    // 1. Hasil Analisis Rapor Pendidikan (PDF)
    function exportAnalisisPdf() {
      const data = getCurrentData();
      const rows = data.indicators.map((ind, idx) => `
        <tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(ind.kode)} ${escapeHtml(ind.nama)}</td>
          <td>${escapeHtml(ind.dimensi)}</td>
          <td style="text-align:center;font-weight:bold;">${ind.nilai !== '' ? ind.nilai : '-'}</td>
          <td style="text-align:center;">${classifyCategory(ind.nilai, ind.kategori)}</td>
          <td style="text-align:center;">${ind.perubahan !== '' ? ind.perubahan : '-'}</td>
          <td>${needsImprovement(ind) ? 'Prioritas Pembenahan' : 'Pertahankan Mutu'}</td>
        </tr>
      `).join('');

      const html = `
        ${buildOfficialDocumentHeader('LAPORAN HASIL ANALISIS RAPOR PENDIDIKAN')}
        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>Nama Indikator</th>
              <th>Dimensi</th>
              <th>Nilai</th>
              <th>Kategori</th>
              <th>Perubahan</th>
              <th>Status Prioritas</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="7">Belum ada data indikator.</td></tr>'}</tbody>
        </table>
      `;
      downloadPdfFromHtml(`Hasil_Analisis_Rapor_Pendidikan_${data.school.tahunPerencanaan || window.appState.activeYear}.pdf`, html);
    }

    // 2. Dokumen Analisis Akar Masalah (Word)
    function exportAkarMasalahWord() {
      const data = getCurrentData();
      const entries = Object.values(data.rootCauses);
      const bodyHtml = entries.map((rc, idx) => {
        const ind = data.indicators.find(i => i.id === rc.indicatorId);
        return `
          <div style="margin-bottom:16px;border:1px solid #cbd5e1;padding:10px;">
            <div style="font-weight:bold;font-size:11pt;color:#064e3b;">${idx + 1}. Indikator: ${escapeHtml(ind ? `${ind.kode} ${ind.nama}` : rc.indicatorId)}</div>
            <p><strong>Masalah Utama:</strong> ${escapeHtml(rc.masalahUtama)}</p>
            <p><strong>Akar Penyebab:</strong></p>
            <ul>${(rc.akarPenyebab || []).map(a => `<li>[${a.status === 'fakta_sekolah' ? 'FAKTA TERKONFIRMASI SEKOLAH' : 'HIPOTESIS AI'}] ${escapeHtml(a.teks)}</li>`).join('')}</ul>
            <p><strong>Analisis 5 Why:</strong> ${(rc.fiveWhys || []).map((w, i) => `(${i + 1}) ${escapeHtml(w)}`).join(' → ')}</p>
            <p><strong>Bukti Pendukung:</strong> ${escapeHtml(rc.buktiPendukung)}</p>
            <p><strong>Pertanyaan Refleksi Guru:</strong> ${(rc.pertanyaanRefleksi || []).map(q => escapeHtml(q)).join('; ')}</p>
            <p><strong>Dampak ke Peserta Didik:</strong> ${escapeHtml(rc.dampakSiswa)}</p>
            <p><strong>Alternatif Solusi:</strong> ${(rc.alternatifSolusi || []).map(s => escapeHtml(s)).join('; ')}</p>
          </div>
        `;
      }).join('');

      const html = `
        ${buildOfficialDocumentHeader('DOKUMEN ANALISIS AKAR MASALAH (METODE 5-WHY & SEBAB-AKIBAT)')}
        ${bodyHtml || '<p>Belum ada data analisis akar masalah.</p>'}
      `;
      downloadWordHtml(`Analisis_Akar_Masalah_${data.school.tahunPerencanaan || window.appState.activeYear}.doc`, html);
    }

    // 3. Daftar Prioritas Pembenahan (Excel)
    function exportPrioritasExcel() {
      const data = getCurrentData();
      const s = data.school;
      const ranked = getRankedPriorities();
      const aoa = [
        ["DAFTAR PRIORITAS PEMBENAHAN RAPOR PENDIDIKAN"],
        [`Sekolah: ${s.nama || '-'} | NPSN: ${s.npsn || '-'} | Tahun Perencanaan: ${s.tahunPerencanaan || window.appState.activeYear}`],
        [],
        ["Peringkat", "Kode", "Nama Indikator", "Nilai", "Kategori", "Skor Capaian", "Skor Tren", "Skor Dampak", "Skor Urgensi", "Skor Kemampuan", "Total Skor", "Alasan Prioritas"]
      ];
      ranked.forEach(({ ind, prio }, idx) => {
        aoa.push([
          idx + 1,
          ind.kode,
          ind.nama,
          ind.nilai,
          classifyCategory(ind.nilai, ind.kategori),
          prio.skorCapaian,
          prio.skorTren,
          prio.skorDampak,
          prio.skorUrgensi,
          prio.skorKemampuan,
          prio.totalSkor,
          prio.alasanPrioritas
        ]);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Prioritas_Pembenahan");
      XLSX.writeFile(wb, `Daftar_Prioritas_Pembenahan_${s.tahunPerencanaan || window.appState.activeYear}.xlsx`);
      showToast("Daftar Prioritas Pembenahan (Excel) berhasil diunduh.", "success");
    }

    // 4a. Dokumen RKJM (Word)
    function exportRkjmWord() {
      const data = getCurrentData();
      const s = data.school;
      const yStart = parseInt(s.periodeRkjmMulai || window.appState.activeYear, 10) || 2026;
      const rows = data.rkjm.map((p, idx) => `
        <tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(p.dimensi)}<br/><strong>${escapeHtml(p.indikatorNama)}</strong></td>
          <td>${escapeHtml(p.akarMasalah)}</td>
          <td><strong>${escapeHtml(p.namaProgram)}</strong><br/>Tujuan: ${escapeHtml(p.tujuanStrategis)}</td>
          <td>${escapeHtml(p.targetTahun1)}</td>
          <td>${escapeHtml(p.targetTahun2)}</td>
          <td>${escapeHtml(p.targetTahun3)}</td>
          <td>${escapeHtml(p.targetTahun4)}</td>
          <td>${escapeHtml(p.indikatorKeberhasilan)}</td>
          <td>${escapeHtml(p.penanggungJawab)} (${escapeHtml(p.sumberPendanaan)})</td>
        </tr>
      `).join('');

      const html = `
        ${buildOfficialDocumentHeader(`RENCANA KERJA JANGKA MENENGAH (RKJM) PERIODE ${yStart}–${yStart + 3}`)}
        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>Dimensi &amp; Indikator</th>
              <th>Akar Masalah</th>
              <th>Program &amp; Tujuan Strategis</th>
              <th>Target ${yStart}</th>
              <th>Target ${yStart + 1}</th>
              <th>Target ${yStart + 2}</th>
              <th>Target ${yStart + 3}</th>
              <th>Indikator Keberhasilan</th>
              <th>PJ &amp; Sumber Dana</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="10">Belum ada data RKJM.</td></tr>'}</tbody>
        </table>
      `;
      downloadWordHtml(`Dokumen_RKJM_${yStart}_${yStart + 3}.doc`, html);
    }

    // 4b. Dokumen RKJM (Excel)
    function exportRkjmExcel() {
      const data = getCurrentData();
      const s = data.school;
      const yStart = parseInt(s.periodeRkjmMulai || window.appState.activeYear, 10) || 2026;
      const aoa = [
        [`RENCANA KERJA JANGKA MENENGAH (RKJM) PERIODE ${yStart}-${yStart + 3}`],
        [`Sekolah: ${s.nama || '-'} | NPSN: ${s.npsn || '-'} | Jenjang: ${s.jenjang || '-'}`],
        [],
        ["No", "Dimensi Rapor Pendidikan", "Indikator Dibenahi", "Akar Masalah", "Tujuan Strategis", "Program Pengembangan Sekolah", `Target ${yStart}`, `Target ${yStart + 1}`, `Target ${yStart + 2}`, `Target ${yStart + 3}`, "Indikator Keberhasilan", "Penanggung Jawab", "Sumber Pendanaan"]
      ];
      data.rkjm.forEach((p, idx) => {
        aoa.push([idx + 1, p.dimensi, p.indikatorNama, p.akarMasalah, p.tujuanStrategis, p.namaProgram, p.targetTahun1, p.targetTahun2, p.targetTahun3, p.targetTahun4, p.indikatorKeberhasilan, p.penanggungJawab, p.sumberPendanaan]);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Matriks_RKJM");
      XLSX.writeFile(wb, `Dokumen_RKJM_${yStart}_${yStart + 3}.xlsx`);
      showToast("Dokumen RKJM (Excel) berhasil diunduh.", "success");
    }

    // 5a. Dokumen RKT (Word)
    function exportRktWord() {
      const data = getCurrentData();
      const s = data.school;
      const tahun = s.tahunPerencanaan || window.appState.activeYear;
      const rows = data.rkt.map((k, idx) => {
        const jadwalStr = (k.bulanPelaksanaan || []).map(m => BULAN_LIST[m - 1]).join(', ');
        return `
          <tr>
            <td>${idx + 1}</td>
            <td>${escapeHtml(k.namaProgram)}</td>
            <td><strong>${escapeHtml(k.namaKegiatan)}</strong><br/>Latar Belakang: ${escapeHtml(k.latarBelakang)}</td>
            <td>${escapeHtml(k.tujuan)}</td>
            <td>${escapeHtml(k.sasaran)}</td>
            <td>${escapeHtml(k.indikatorKeberhasilan)}<br/>Target: ${escapeHtml(k.target)}</td>
            <td>${escapeHtml(jadwalStr || '-')}</td>
            <td>${escapeHtml(k.penanggungJawab)}</td>
            <td style="text-align:right;">${formatRupiah(getRktEffectiveBudget(k))}<br/>(${escapeHtml(k.sumberDana)})</td>
            <td>${escapeHtml(k.buktiPelaksanaan)}</td>
          </tr>
        `;
      }).join('');

      const html = `
        ${buildOfficialDocumentHeader(`RENCANA KERJA TAHUNAN (RKT) TAHUN ${tahun}`)}
        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>Nama Program</th>
              <th>Nama Kegiatan &amp; Latar Belakang</th>
              <th>Tujuan</th>
              <th>Sasaran</th>
              <th>Indikator &amp; Target</th>
              <th>Jadwal</th>
              <th>PJ</th>
              <th>Kebutuhan Anggaran</th>
              <th>Bukti Pelaksanaan</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="10">Belum ada data RKT.</td></tr>'}</tbody>
        </table>
      `;
      downloadWordHtml(`Dokumen_RKT_Tahun_${tahun}.doc`, html);
    }

    // 5b. Dokumen RKT (Excel)
    function exportRktExcel() {
      const data = getCurrentData();
      const s = data.school;
      const tahun = s.tahunPerencanaan || window.appState.activeYear;
      const aoa = [
        [`RENCANA KERJA TAHUNAN (RKT) TAHUN ${tahun}`],
        [`Sekolah: ${s.nama || '-'} | NPSN: ${s.npsn || '-'} | Tahun Ajaran: ${s.tahunAjaran || '-'}`],
        [],
        ["No", "Nama Program (RKJM)", "Nama Kegiatan", "Latar Belakang", "Tujuan", "Sasaran", "Indikator Keberhasilan", "Target", "Jadwal Pelaksanaan", "Penanggung Jawab", "Kebutuhan Anggaran (Rp)", "Sumber Dana", "Bukti Pelaksanaan"]
      ];
      data.rkt.forEach((k, idx) => {
        const jadwalStr = (k.bulanPelaksanaan || []).map(m => BULAN_LIST[m - 1]).join(', ');
        aoa.push([idx + 1, k.namaProgram, k.namaKegiatan, k.latarBelakang, k.tujuan, k.sasaran, k.indikatorKeberhasilan, k.target, jadwalStr, k.penanggungJawab, getRktEffectiveBudget(k), k.sumberDana, k.buktiPelaksanaan]);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Matriks_RKT");
      XLSX.writeFile(wb, `Dokumen_RKT_Tahun_${tahun}.xlsx`);
      showToast("Dokumen RKT (Excel) berhasil diunduh.", "success");
    }

    // 6. Rancangan Anggaran ARKAS (Excel)
    function exportArkasExcel() {
      const data = getCurrentData();
      const s = data.school;
      const tahun = s.tahunPerencanaan || window.appState.activeYear;
      const totalAll = data.arkas.reduce((sum, a) => sum + (Number(a.jumlahAnggaran) || 0), 0);

      const aoa = [
        [`RANCANGAN KERTAS KERJA ANGGARAN (BAHAN INPUT ARKAS) TAHUN ANGGARAN ${tahun}`],
        [`Sekolah: ${s.nama || '-'} | NPSN: ${s.npsn || '-'} | Jenjang: ${s.jenjang || '-'}`],
        ["CATATAN: Dokumen ini adalah rancangan perencanaan berbasis Rapor Pendidikan sebagai bahan input ke aplikasi ARKAS resmi. Wajib verifikasi kode rekening & SSH daerah."],
        [],
        ["No", "Program", "Kegiatan (RKT)", "Komponen Pembiayaan BOSP", "Uraian Belanja", "Volume", "Satuan", "Harga Satuan (Rp)", "Jumlah Anggaran (Rp)", "Sumber Dana", "Waktu Pelaksanaan", "Status Verifikasi"]
      ];
      data.arkas.forEach((a, idx) => {
        aoa.push([idx + 1, a.namaProgram, a.namaKegiatan, a.komponenPembiayaan, a.uraianBelanja, a.volume, a.satuan, a.hargaSatuan, a.jumlahAnggaran, a.sumberDana, a.waktuPelaksanaan, a.statusVerifikasi]);
      });
      aoa.push(["", "", "", "", "", "", "", "TOTAL ANGGARAN", totalAll, "", "", ""]);

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Rancangan_ARKAS");
      XLSX.writeFile(wb, `Rancangan_Anggaran_ARKAS_${tahun}.xlsx`);
      showToast("Rancangan Anggaran ARKAS (Excel) berhasil diunduh.", "success");
    }

    // 7. Laporan Monitoring & Evaluasi (PDF)
    function exportMonevPdf() {
      const data = getCurrentData();
      const s = data.school;
      const tahun = s.tahunPerencanaan || window.appState.activeYear;
      const statusLabel = { belum: "Belum Dilaksanakan", berjalan: "Sedang Berjalan", selesai: "Selesai" };

      const rows = data.rkt.map((k, idx) => {
        const m = data.monev[k.id] || { statusPelaksanaan: 'belum', persentaseKegiatan: 0, realisasiAnggaran: 0, hambatan: '-', buktiKegiatan: '-', evaluasiHasil: '-', tindakLanjut: '-' };
        const pagu = getRktEffectiveBudget(k);
        return `
          <tr>
            <td>${idx + 1}</td>
            <td><strong>${escapeHtml(k.namaKegiatan)}</strong><br/>PJ: ${escapeHtml(k.penanggungJawab)}</td>
            <td>${statusLabel[m.statusPelaksanaan] || 'Belum'}</td>
            <td style="text-align:center;">${m.persentaseKegiatan}%</td>
            <td style="text-align:right;">${formatRupiah(m.realisasiAnggaran)}<br/><small>Pagu: ${formatRupiah(pagu)}</small></td>
            <td>${escapeHtml(m.hambatan || '-')}</td>
            <td>${escapeHtml(m.buktiKegiatan || '-')}</td>
            <td>${escapeHtml(m.evaluasiHasil || '-')}</td>
            <td>${escapeHtml(m.tindakLanjut || '-')}</td>
          </tr>
        `;
      }).join('');

      const html = `
        ${buildOfficialDocumentHeader(`LAPORAN MONITORING DAN EVALUASI PROGRAM TAHUN ${tahun}`)}
        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>Kegiatan RKT</th>
              <th>Status</th>
              <th>Realisasi Fisik</th>
              <th>Realisasi Anggaran</th>
              <th>Hambatan</th>
              <th>Bukti Kegiatan</th>
              <th>Evaluasi Hasil</th>
              <th>Rekomendasi Tindak Lanjut</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="9">Belum ada data Monitoring & Evaluasi.</td></tr>'}</tbody>
        </table>
      `;
      downloadPdfFromHtml(`Laporan_Monev_Program_${tahun}.pdf`, html);
    }

    // =========================================================================
    // MODULE O: MANAJEMEN DATA, MULTI-TAHUN, DUPLIKASI & RIWAYAT PERUBAHAN
    // =========================================================================
    function renderManajemenModule(container) {
      const data = getCurrentData();
      const years = Object.keys(window.appState.yearsData).sort();

      container.innerHTML = `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold text-slate-900">12. Manajemen Data, Multi-Tahun &amp; Riwayat Perubahan</h2>
            <p class="text-xs text-slate-500">Kelola beberapa tahun perencanaan, duplikasi program tahun sebelumnya, impor/ekspor proyek, dan audit log.</p>
          </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-5">
          <!-- Multi-Tahun & Duplikasi -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
              <i data-lucide="copy-plus" class="w-4 h-4 text-emerald-800"></i>
              <span>Tahun Perencanaan &amp; Duplikasi</span>
            </h3>
            <p class="text-xs text-slate-500 leading-relaxed">
              Simpan beberapa tahun perencanaan secara terpisah atau salin struktur program RKJM/RKT dari tahun sebelumnya ke tahun aktif (${window.appState.activeYear}).
            </p>
            <div class="space-y-2">
              <button onclick="openNewYearModal()" class="w-full py-2.5 px-4 rounded-xl emerald-gradient text-white text-xs font-bold flex items-center justify-center gap-1.5">
                <i data-lucide="calendar-plus" class="w-4 h-4"></i> Buat Tahun Perencanaan Baru
              </button>
              <button onclick="openDuplicateYearModal()" class="w-full py-2.5 px-4 rounded-xl border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-800 text-xs font-bold flex items-center justify-center gap-1.5">
                <i data-lucide="copy" class="w-4 h-4 text-amber-600"></i> Duplikasi Program Tahun Lain
              </button>
            </div>
            <div class="pt-2 border-t border-slate-100 text-xs">
              <span class="font-semibold text-slate-600">Tahun Tersedia:</span>
              <div class="flex flex-wrap gap-1.5 mt-1.5">
                ${years.map(y => `
                  <button onclick="switchPlanningYear('${y}')" class="px-2.5 py-1 rounded-lg font-mono font-bold text-xs ${y === window.appState.activeYear ? 'bg-emerald-900 text-amber-300' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'}">${y}</button>
                `).join('')}
              </div>
            </div>
          </div>

          <!-- Ekspor / Impor Cadangan File Proyek (.json) -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
              <i data-lucide="hard-drive-download" class="w-4 h-4 text-emerald-800"></i>
              <span>Simpan &amp; Buka Kembali Proyek (.JSON)</span>
            </h3>
            <p class="text-xs text-slate-500 leading-relaxed">
              Unduh seluruh data identitas sekolah, analisis, RKJM, RKT, ARKAS, dan Monev dalam satu file cadangan agar dapat dibuka kembali kapan saja di perangkat lain.
            </p>
            <div class="space-y-2">
              <button onclick="exportFullWorkspaceJSON()" class="w-full py-2.5 px-4 rounded-xl gold-gradient text-emerald-950 text-xs font-extrabold flex items-center justify-center gap-1.5 shadow-xs">
                <i data-lucide="download" class="w-4 h-4"></i> Unduh File Proyek (.JSON)
              </button>
              <label class="block w-full cursor-pointer py-2.5 px-4 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-bold text-center transition">
                <span class="flex items-center justify-center gap-1.5">
                  <i data-lucide="folder-up" class="w-4 h-4 text-emerald-700"></i> Buka / Impor File Proyek (.JSON)
                </span>
                <input type="file" accept=".json" onchange="importFullWorkspaceJSON(event)" class="hidden" />
              </label>
            </div>
          </div>

          <!-- Reset & Mode Demo -->
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
              <i data-lucide="shield" class="w-4 h-4 text-rose-700"></i>
              <span>Mode Simulasi &amp; Reset Data</span>
            </h3>
            <p class="text-xs text-slate-500 leading-relaxed">
              Gunakan Mode Simulasi untuk latihan tim sekolah, atau kosongkan seluruh data pada tahun aktif untuk memulai input data sekolah asli dari awal.
            </p>
            <div class="space-y-2">
              <button onclick="toggleDemoSimulation()" class="w-full py-2.5 px-4 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-950 text-xs font-bold flex items-center justify-center gap-1.5">
                <i data-lucide="flask-conical" class="w-4 h-4 text-amber-600"></i>
                <span>${window.appState.isDemoMode ? 'Keluar & Kosongkan Data Demo' : 'Aktifkan Mode Simulasi / Demo'}</span>
              </button>
              <button onclick="promptResetToCleanState()" class="w-full py-2.5 px-4 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-800 border border-rose-200 text-xs font-bold flex items-center justify-center gap-1.5">
                <i data-lucide="trash-2" class="w-4 h-4"></i> Reset Bersih Data Tahun ${window.appState.activeYear}
              </button>
            </div>
          </div>
        </div>

        <!-- Riwayat Perubahan (Audit Trail) -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
          <div class="p-4 sm:p-5 border-b border-slate-200 flex items-center justify-between">
            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
              <i data-lucide="history" class="w-4 h-4 text-emerald-800"></i>
              <span>Riwayat Perubahan &amp; Aktivitas Tim (${data.history.length})</span>
            </h3>
          </div>
          <div class="max-h-80 overflow-y-auto divide-y divide-slate-100 text-xs">
            ${data.history.map(h => `
              <div class="px-5 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-1 hover:bg-slate-50">
                <div>
                  <span class="font-bold text-emerald-900">[${escapeHtml(h.modul)}]</span>
                  <span class="font-semibold text-slate-800 ml-1">${escapeHtml(h.aksi)}</span>
                  <span class="text-slate-500 ml-1">— ${escapeHtml(h.keterangan)}</span>
                </div>
                <span class="text-[11px] font-mono text-slate-400 shrink-0">${escapeHtml(h.timestamp)}</span>
              </div>
            `).join('')}
          </div>
        </div>
      `;
    }

    function openNewYearModal() {
      const nextY = String((parseInt(window.appState.activeYear, 10) || 2026) + 1);
      openCustomModal(`
        <div class="p-6">
          <h3 class="font-bold text-slate-900 text-base mb-2">Tambah Tahun Perencanaan Baru</h3>
          <p class="text-xs text-slate-500 mb-4">Buat ruang kerja perencanaan terpisah untuk tahun anggaran berikutnya.</p>
          <form onsubmit="createNewPlanningYear(event)" class="space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Tahun Perencanaan / Anggaran</label>
              <input type="number" min="2023" max="2035" name="newYear" value="${nextY}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
            </div>
            <label class="flex items-center gap-2 text-slate-700 cursor-pointer">
              <input type="checkbox" name="copySchool" checked class="rounded text-emerald-700" />
              <span>Salin Identitas Sekolah dari tahun ${window.appState.activeYear}</span>
            </label>
            <div class="flex justify-end gap-2 pt-3">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Buat Tahun Baru</button>
            </div>
          </form>
        </div>
      `, 'max-w-md');
    }

    function createNewPlanningYear(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const y = String(fd.get('newYear')).trim();
      const copySchool = fd.get('copySchool') === 'on';

      if (!window.appState.yearsData[y]) {
        const fresh = createEmptyYearData(y);
        if (copySchool) {
          const curSchool = getCurrentData().school;
          fresh.school = {
            ...curSchool,
            tahunRapor: String(parseInt(y, 10) - 1),
            tahunPerencanaan: y,
            tahunAjaran: `${y}/${parseInt(y, 10) + 1}`
          };
        }
        window.appState.yearsData[y] = fresh;
      }
      window.appState.activeYear = y;
      closeCustomModal();
      triggerAutoSave("Manajemen Data", "Tambah Tahun Perencanaan", `Membuat dan beralih ke tahun perencanaan ${y}.`);
      renderAll();
      showToast(`Tahun perencanaan ${y} aktif.`, "success");
    }

    function openDuplicateYearModal() {
      const years = Object.keys(window.appState.yearsData).filter(y => y !== window.appState.activeYear);
      if (years.length === 0) {
        showToast("Belum ada tahun perencanaan lain untuk diduplikasi. Buat tahun baru terlebih dahulu.", "warning");
        return;
      }
      openCustomModal(`
        <div class="p-6">
          <h3 class="font-bold text-slate-900 text-base mb-2">Duplikasi Program dari Tahun Sebelumnya</h3>
          <p class="text-xs text-slate-500 mb-4">Menyalin program RKJM dan kegiatan RKT dari tahun sumber ke tahun aktif <strong>${window.appState.activeYear}</strong>.</p>
          <form onsubmit="executeDuplicateYear(event)" class="space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Pilih Tahun Sumber</label>
              <select name="sourceYear" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono">
                ${years.map(y => `<option value="${y}">Tahun ${y}</option>`).join('')}
              </select>
            </div>
            <div class="flex justify-end gap-2 pt-3">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Duplikasi Sekarang</button>
            </div>
          </form>
        </div>
      `, 'max-w-md');
    }

    function executeDuplicateYear(e) {
      e.preventDefault();
      const srcYear = new FormData(e.target).get('sourceYear');
      const src = window.appState.yearsData[srcYear];
      const dst = getCurrentData();
      if (!src) return;

      src.rkjm.forEach(p => dst.rkjm.push({ ...JSON.parse(JSON.stringify(p)), id: uid() }));
      src.rkt.forEach(k => dst.rkt.push({ ...JSON.parse(JSON.stringify(k)), id: uid() }));

      closeCustomModal();
      triggerAutoSave("Manajemen Data", "Duplikasi Program", `Menduplikasi RKJM & RKT dari tahun ${srcYear} ke ${window.appState.activeYear}.`);
      renderAll();
      showToast(`Berhasil menduplikasi program dari tahun ${srcYear}.`, "success");
    }

    function exportFullWorkspaceJSON() {
      const dataStr = JSON.stringify(window.appState, null, 2);
      const blob = new Blob([dataStr], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const schoolSlug = (getCurrentData().school.nama || 'Sekolah').replace(/[^a-zA-Z0-9]/g, '_');
      a.href = url;
      a.download = `Proyek_RAPOR_SMART_AI_${schoolSlug}_${window.appState.activeYear}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast("File proyek (.json) berhasil disimpan ke perangkat Anda.", "success");
    }

    function importFullWorkspaceJSON(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const parsed = JSON.parse(e.target.result);
          if (!parsed || !parsed.yearsData) {
            throw new Error("Format file tidak valid");
          }
          window.appState = parsed;
          triggerAutoSave("Manajemen Data", "Impor File Proyek", `Memulihkan ruang kerja dari file ${file.name}`);
          renderAll();
          showToast("Proyek RAPOR SMART AI berhasil dibuka kembali.", "success");
        } catch (err) {
          showToast("File JSON tidak valid atau bukan cadangan RAPOR SMART AI.", "error");
        }
        event.target.value = "";
      };
      reader.readAsText(file);
    }

    // =========================================================================
    // DEMO SIMULATION MODE (STRICTLY SEPARATE & CLEARLY LABELED)
    // =========================================================================
    function toggleDemoSimulation() {
      if (window.appState.isDemoMode) {
        promptResetToCleanState();
        return;
      }
      showConfirmDialog({
        title: "Aktifkan Mode Simulasi / Data Contoh?",
        message: "Mode ini memuat data contoh simulasi pelatihan (bukan data resmi sekolah) agar Tim Pengembang Sekolah dapat mencoba seluruh alur dari Rapor Pendidikan hingga ARKAS.",
        confirmText: "Ya, Muat Simulasi",
        onConfirm: () => {
          loadSimulationDataset();
        }
      });
    }

    function promptResetToCleanState() {
      showConfirmDialog({
        title: "Kosongkan & Reset Data Tahun Ini?",
        message: "Seluruh data pada tahun perencanaan aktif akan dikosongkan sehingga Anda dapat menginput data resmi sekolah Anda dari awal.",
        confirmText: "Ya, Kosongkan Data",
        danger: true,
        onConfirm: () => {
          window.appState.isDemoMode = false;
          window.appState.yearsData[window.appState.activeYear] = createEmptyYearData(window.appState.activeYear);
          window.appState.selectedAkarIndicatorId = null;
          window.appState.selectedTraceIndicatorId = null;
          triggerAutoSave("Sistem", "Reset Data Bersih", "Mengosongkan data untuk input satuan pendidikan asli.");
          renderAll();
          showToast("Ruang kerja dikosongkan dan siap diisi data sekolah Anda.", "success");
        }
      });
    }

    function loadSimulationDataset() {
      const y = window.appState.activeYear;
      const indLit = uid(), indNum = uid(), indKar = uid(), indPemb = uid(), indAman = uid(), indBineka = uid();
      const rec1 = uid(), rec2 = uid(), rec3 = uid();
      const rkjm1 = uid(), rkjm2 = uid(), rkjm3 = uid();
      const rkt1 = uid(), rkt2 = uid(), rkt3 = uid();

      window.appState.isDemoMode = true;
      window.appState.yearsData[y] = {
        school: {
          nama: "SMP NEGERI CONTOH NUSANTARA (DATA SIMULASI)",
          npsn: "00000000",
          alamat: "Jl. Pendidikan Merdeka No. 45 (Simulasi)",
          kecamatan: "Cilacap Tengah",
          kabupaten: "Kabupaten Cilacap",
          provinsi: "Jawa Tengah",
          jenjang: "SMP / MTs",
          tahunRapor: String(parseInt(y, 10) - 1),
          tahunPerencanaan: y,
          tahunAjaran: `${y}/${parseInt(y, 10) + 1}`,
          periodeRkjmMulai: y,
          periodeRkjmAkhir: String(parseInt(y, 10) + 3)
        },
        indicators: [
          {
            id: indLit,
            kode: "A.1",
            nama: "Kemampuan Literasi",
            dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
            nilai: 62.50,
            kategori: "Sedang",
            perubahan: -4.20,
            peringkat: "Peringkat Menengah (41-60%) Kab/Kota",
            rekomendasiRapor: "Peningkatan kompetensi guru dalam pembelajaran literasi lintas mata pelajaran dan pemanfaatan buku bacaan bermutu.",
            terverifikasi: true
          },
          {
            id: indNum,
            kode: "A.2",
            nama: "Kemampuan Numerasi",
            dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
            nilai: 48.00,
            kategori: "Kurang",
            perubahan: -6.50,
            peringkat: "Peringkat Bawah (61-80%) Kab/Kota",
            rekomendasiRapor: "Penguatan pembelajaran numerasi kontekstual dan asesmen diagnostik berkala melalui Komunitas Belajar.",
            terverifikasi: true
          },
          {
            id: indKar,
            kode: "A.3",
            nama: "Karakter",
            dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
            nilai: 74.80,
            kategori: "Baik",
            perubahan: 2.15,
            peringkat: "Peringkat Atas (1-20%) Kab/Kota",
            rekomendasiRapor: "Pertahankan pembiasaan Profil Pelajar Pancasila melalui kokurikuler.",
            terverifikasi: true
          },
          {
            id: indPemb,
            kode: "D.1",
            nama: "Kualitas Pembelajaran",
            dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran",
            nilai: 58.40,
            kategori: "Sedang",
            perubahan: -1.80,
            peringkat: "Peringkat Menengah (41-60%) Kab/Kota",
            rekomendasiRapor: "Penerapan pembelajaran interaktif dan umpan balik konstruktif di kelas.",
            terverifikasi: true
          },
          {
            id: indAman,
            kode: "D.4",
            nama: "Iklim Keamanan Satuan Pendidikan",
            dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran",
            nilai: 66.00,
            kategori: "Sedang",
            perubahan: 1.20,
            peringkat: "Peringkat Menengah Atas (21-40%)",
            rekomendasiRapor: "Penguatan Tim Pencegahan dan Penanganan Kekerasan (TPPK) sekolah.",
            terverifikasi: true
          },
          {
            id: indBineka,
            kode: "D.8",
            nama: "Iklim Kebinekaan",
            dimensi: "Dimensi D: Mutu dan Relevansi Pembelajaran",
            nilai: 78.20,
            kategori: "Baik",
            perubahan: 3.40,
            peringkat: "Peringkat Atas (1-20%)",
            rekomendasiRapor: "Pertahankan budaya toleransi dan penghargaan keberagaman.",
            terverifikasi: true
          }
        ],
        rootCauses: {
          [indNum]: {
            indicatorId: indNum,
            masalahUtama: "Capaian Kemampuan Numerasi berada pada kategori Kurang (48.00) dan turun 6.50 poin dibanding tahun lalu, khususnya pada domain aljabar dan pemecahan masalah kontekstual.",
            fiveWhys: [
              "Why 1: Siswa kesulitan menyelesaikan soal penalaran numerasi berbasis konteks kehidupan sehari-hari.",
              "Why 2: Pembelajaran matematika dan IPA masih dominan berfokus pada hafalan rumus prosedural.",
              "Why 3: Guru belum rutin melakukan asesmen diagnostik awal untuk memetakan tingkat kesiapan numerasi siswa.",
              "Why 4: Komunitas Belajar (Kombel) sekolah belum memiliki agenda bedah soal numerasi dan berbagi praktik pembelajaran aktif.",
              "Why 5: Ketersediaan alat peraga matematika kontekstual dan modul penguatan numerasi lintas mapel masih terbatas."
            ],
            akarPenyebab: [
              { kategori: "Pembelajaran", teks: "Pembelajaran di kelas masih berfokus pada latihan soal prosedural daripada penalaran kontekstual.", status: "fakta_sekolah" },
              { kategori: "Kompetensi GTK", teks: "Belum optimalnya praktik berbagi ajar numerasi lintas mata pelajaran di Komunitas Belajar sekolah.", status: "fakta_sekolah" },
              { kategori: "Sarpras & Media", teks: "Keterbatasan alat peraga numerasi dan bahan ajar kontekstual di kelas.", status: "dugaan_ai" }
            ],
            buktiPendukung: "Skor Rapor Pendidikan A.2 = 48.00 (Kurang, turun -6.50) serta rekap nilai PHB Matematika kelas VII & VIII semester ganjil.",
            pertanyaanRefleksi: [
              "Bagian kompetensi numerasi mana yang paling banyak belum dikuasai siswa berdasarkan hasil ulangan harian?",
              "Bagaimana guru non-matematika (IPA, IPS, Bahasa) dapat mengintegrasikan pembacaan data/grafik dalam pembelajaran?",
              "Alat peraga sederhana apa yang paling dibutuhkan guru untuk memvisualisasikan konsep abstrak?"
            ],
            dataTambahan: [
              "Hasil asesmen diagnostik numerasi awal semester bagi seluruh peserta didik.",
              "Instrumen observasi supervisi kelas terkait penggunaan masalah kontekstual."
            ],
            dampakSiswa: "Peserta didik mengalami kesulitan berpikir logis-kuantitatif dan rendahnya kepercayaan diri saat menghadapi soal pemecahan masalah.",
            alternatifSolusi: [
              "Lokakarya Kombel guru menyusun perangkat ajar & asesmen diagnostik numerasi kontekstual.",
              "Klinik numerasi (pendampingan belajar terstruktur) bagi siswa dengan kemampuan dasar rendah.",
              "Pengadaan alat peraga matematika dan modul penguatan numerasi."
            ]
          }
        },
        priorities: {},
        recommendations: [
          {
            id: rec1,
            indicatorId: indNum,
            indicatorNama: "[A.2] Kemampuan Numerasi",
            permasalahan: "Capaian Kemampuan Numerasi kategori Kurang (48.00, turun -6.50 poin) akibat pembelajaran yang masih prosedural dan belum meratanya asesmen diagnostik.",
            tujuanPerbaikan: "Meningkatkan kemampuan penalaran numerasi peserta didik dan kompetensi guru dalam pembelajaran numerasi kontekstual.",
            sasaran: "32 Guru Mata Pelajaran & 420 Peserta Didik Kelas VII–IX",
            bentukKegiatan: "Lokakarya Kombel Pembelajaran Numerasi Kontekstual & Klinik Numerasi Siswa",
            langkahPelaksanaan: "1. Pelaksanaan asesmen diagnostik numerasi awal.\n2. Lokakarya 2 hari di Kombel penyusunan modul ajar numerasi lintas mapel.\n3. Pelaksanaan Klinik Numerasi 2x seminggu bagi siswa yang membutuhkan.\n4. Evaluasi progres tengah dan akhir semester.",
            penanggungJawab: "Waka Kurikulum & Koordinator Kombel",
            indikatorKeberhasilan: "90% guru menerapkan soal penalaran kontekstual dan skor numerasi meningkat minimal +8 poin.",
            targetCapaian: "Naik dari kategori Kurang (48.00) ke kategori Sedang (>56.00)",
            waktuPelaksanaan: `Februari – November ${y}`,
            sumberDaya: "Narasumber Pengawas/Fasilitator Daerah, ATK Lokakarya, Modul Numerasi, Alat Peraga Matematika",
            perkiraanAnggaran: 6500000,
            statusPersetujuan: "disetujui",
            sumberAi: true
          },
          {
            id: rec2,
            indicatorId: indLit,
            indicatorNama: "[A.1] Kemampuan Literasi",
            permasalahan: "Kemampuan Literasi kategori Sedang (62.50, turun -4.20 poin) karena pemanfaatan buku bacaan non-teks pelajaran belum optimal.",
            tujuanPerbaikan: "Menumbuhkan budaya baca dan meningkatkan kemampuan memahami teks informasi serta sastra.",
            sasaran: "Seluruh Peserta Didik & Pengelola Perpustakaan Sekolah",
            bentukKegiatan: "Revitalisasi Pojok Baca Kelas & Program 15 Menit Membaca Bermakna",
            langkahPelaksanaan: "1. Pengadaan buku bacaan pengayaan fiksi & nonfiksi.\n2. Penataan pojok baca di setiap ruang kelas.\n3. Pembiasaan jurnal baca mingguan siswa.",
            penanggungJawab: "Kepala Perpustakaan & Guru Bahasa Indonesia",
            indikatorKeberhasilan: "100% kelas memiliki pojok baca aktif dan tingkat peminjaman buku naik 50%.",
            targetCapaian: "Skor Literasi naik ke kategori Baik (>=70.00)",
            waktuPelaksanaan: `Maret – Oktober ${y}`,
            sumberDaya: "Koleksi buku bacaan bermutu, rak pojok baca, jurnal literasi siswa",
            perkiraanAnggaran: 8000000,
            statusPersetujuan: "disetujui",
            sumberAi: true
          },
          {
            id: rec3,
            indicatorId: indPemb,
            indicatorNama: "[D.1] Kualitas Pembelajaran",
            permasalahan: "Kualitas Pembelajaran kategori Sedang (58.40, turun -1.80) memerlukan penguatan umpan balik konstruktif dan supervisi akademik.",
            tujuanPerbaikan: "Meningkatkan kualitas interaksi kelas dan pembelajaran berdiferensiasi.",
            sasaran: "Seluruh Pendidik Satuan Pendidikan",
            bentukKegiatan: "Supervisi Akademik Klinis & Berbagi Praktik Baik Pembelajaran Interaktif",
            langkahPelaksanaan: "1. Penyusunan jadwal supervisi pra-observasi, observasi, dan pasca-observasi.\n2. Refleksi bersama di Kombel bulanan.",
            penanggungJawab: "Kepala Sekolah & Tim Supervisi",
            indikatorKeberhasilan: "100% guru disupervisi dan menindaklanjuti umpan balik pembelajaran.",
            targetCapaian: "Skor Kualitas Pembelajaran meningkat ke >=65.00",
            waktuPelaksanaan: `Triwulan I – IV ${y}`,
            sumberDaya: "Instrumen supervisi, ATK administrasi Kombel",
            perkiraanAnggaran: 2500000,
            statusPersetujuan: "menunggu",
            sumberAi: true
          }
        ],
        rkjm: [
          {
            id: rkjm1,
            indicatorId: indNum,
            recommendationId: rec1,
            dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
            indikatorNama: "[A.2] Kemampuan Numerasi",
            akarMasalah: "Pembelajaran masih prosedural dan belum meratanya asesmen diagnostik numerasi.",
            tujuanStrategis: "Mewujudkan pembelajaran numerasi kontekstual yang meningkatkan daya nalar kuantitatif siswa.",
            namaProgram: "Penguatan Ekosistem Numerasi & Klinik Belajar Siswa",
            targetTahun1: "Skor >56 (Sedang)",
            targetTahun2: "Skor >63 (Sedang Tinggi)",
            targetTahun3: "Skor >70 (Baik)",
            targetTahun4: "Skor >75 (Baik Mantap)",
            indikatorKeberhasilan: "Peningkatan skor Rapor Pendidikan A.2 setiap tahun",
            penanggungJawab: "Waka Kurikulum",
            sumberPendanaan: "BOSP Reguler"
          },
          {
            id: rkjm2,
            indicatorId: indLit,
            recommendationId: rec2,
            dimensi: "Dimensi A: Mutu dan Relevansi Hasil Belajar Peserta Didik",
            indikatorNama: "[A.1] Kemampuan Literasi",
            akarMasalah: "Terbatasnya bahan bacaan pengayaan di kelas dan pemanfaatan perpustakaan.",
            tujuanStrategis: "Membangun budaya literasi sekolah dan kecakapan membaca pemahaman siswa.",
            namaProgram: "Gerakan Literasi Sekolah & Revitalisasi Pojok Baca",
            targetTahun1: "Skor >=70 (Baik)",
            targetTahun2: "Skor >=74 (Baik)",
            targetTahun3: "Skor >=78 (Baik)",
            targetTahun4: "Skor >=82 (Baik)",
            indikatorKeberhasilan: "Capaian A.1 masuk dan bertahan di kategori Baik",
            penanggungJawab: "Kepala Perpustakaan",
            sumberPendanaan: "BOSP Reguler"
          }
        ],
        rkt: [
          {
            id: rkt1,
            rkjmId: rkjm1,
            indicatorId: indNum,
            namaProgram: "Penguatan Ekosistem Numerasi & Klinik Belajar Siswa",
            namaKegiatan: "Lokakarya Kombel Pembelajaran Numerasi Kontekstual & Klinik Numerasi",
            latarBelakang: "Skor Kemampuan Numerasi turun -6.50 poin menjadi 48.00 (Kurang).",
            tujuan: "Meningkatkan kompetensi guru merancang asesmen & pembelajaran numerasi kontekstual.",
            sasaran: "32 Guru & 420 Peserta Didik",
            indikatorKeberhasilan: "90% guru menerapkan pembelajaran kontekstual & skor numerasi naik",
            target: "2 kali lokakarya guru & 24 sesi klinik numerasi",
            bulanPelaksanaan: [2, 3, 4, 8, 9, 10],
            penanggungJawab: "Waka Kurikulum",
            kebutuhanAnggaran: 6500000,
            sumberDana: "BOSP Reguler",
            buktiPelaksanaan: "Daftar Hadir, Modul Ajar Numerasi, Foto Kegiatan, Laporan Hasil Klinik"
          },
          {
            id: rkt2,
            rkjmId: rkjm2,
            indicatorId: indLit,
            namaProgram: "Gerakan Literasi Sekolah & Revitalisasi Pojok Baca",
            namaKegiatan: "Pengadaan Buku Bacaan Bermutu & Penataan Pojok Baca Kelas",
            latarBelakang: "Skor Kemampuan Literasi turun -4.20 poin menjadi 62.50 (Sedang).",
            tujuan: "Menyediakan akses buku bacaan non-teks pelajaran di setiap kelas.",
            sasaran: "12 Rombongan Belajar (Kelas VII–IX)",
            indikatorKeberhasilan: "12 kelas memiliki pojok baca aktif",
            target: "200 eksemplar buku bacaan baru & 12 rak pojok baca",
            bulanPelaksanaan: [3, 4, 7, 8],
            penanggungJawab: "Kepala Perpustakaan",
            kebutuhanAnggaran: 8000000,
            sumberDana: "BOSP Reguler",
            buktiPelaksanaan: "BAST Buku, Inventaris Perpustakaan, Foto Pojok Baca"
          }
        ],
        arkas: [
          {
            id: uid(),
            rktId: rkt1,
            rkjmId: rkjm1,
            indicatorId: indNum,
            namaProgram: "Penguatan Ekosistem Numerasi & Klinik Belajar Siswa",
            namaKegiatan: "Lokakarya Kombel Pembelajaran Numerasi Kontekstual & Klinik Numerasi",
            komponenPembiayaan: "Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
            uraianBelanja: "Konsumsi Rapat/Lokakarya Kombel Guru (35 org x 2 hari)",
            volume: 70,
            satuan: "Orang/Hari",
            hargaSatuan: 35000,
            jumlahAnggaran: 2450000,
            sumberDana: "BOSP Reguler",
            waktuPelaksanaan: "Triwulan I & III",
            statusVerifikasi: "Terverifikasi SSH & Juknis BOSP"
          },
          {
            id: uid(),
            rktId: rkt1,
            rkjmId: rkjm1,
            indicatorId: indNum,
            namaProgram: "Penguatan Ekosistem Numerasi & Klinik Belajar Siswa",
            namaKegiatan: "Lokakarya Kombel Pembelajaran Numerasi Kontekstual & Klinik Numerasi",
            komponenPembiayaan: "Pelaksanaan Kegiatan Pembelajaran dan Ekstrakurikuler",
            uraianBelanja: "Pengadaan Alat Peraga Matematika Kontekstual & Cetak Modul Klinik",
            volume: 1,
            satuan: "Paket",
            hargaSatuan: 4050000,
            jumlahAnggaran: 4050000,
            sumberDana: "BOSP Reguler",
            waktuPelaksanaan: "Triwulan II",
            statusVerifikasi: "Perlu Verifikasi SSH/Kode Rekening"
          },
          {
            id: uid(),
            rktId: rkt2,
            rkjmId: rkjm2,
            indicatorId: indLit,
            namaProgram: "Gerakan Literasi Sekolah & Revitalisasi Pojok Baca",
            namaKegiatan: "Pengadaan Buku Bacaan Bermutu & Penataan Pojok Baca Kelas",
            komponenPembiayaan: "Pengembangan Perpustakaan dan/atau Layanan Pojok Baca",
            uraianBelanja: "Buku Bacaan Pengayaan Fiksi & Nonfiksi Bermutu (Perpusnas/Kemendikdasmen)",
            volume: 160,
            satuan: "Eksemplar",
            hargaSatuan: 50000,
            jumlahAnggaran: 8000000,
            sumberDana: "BOSP Reguler",
            waktuPelaksanaan: "Triwulan II",
            statusVerifikasi: "Terverifikasi SSH & Juknis BOSP"
          }
        ],
        monev: {
          [rkt1]: {
            rktId: rkt1,
            statusPelaksanaan: "berjalan",
            persentaseKegiatan: 60,
            realisasiAnggaran: 2450000,
            hambatan: "Penyesuaian jadwal klinik numerasi agar tidak berbenturan dengan latihan ekstrakurikuler sore.",
            buktiKegiatan: "Laporan Lokakarya Kombel Tahap 1 & Daftar Hadir Guru",
            evaluasiHasil: "Guru mulai menyusun soal pemantik numerasi di awal pembelajaran.",
            tindakLanjut: "Menjadwalkan klinik numerasi pada jam literasi/numerasi pagi hari."
          },
          [rkt2]: {
            rktId: rkt2,
            statusPelaksanaan: "selesai",
            persentaseKegiatan: 100,
            realisasiAnggaran: 8000000,
            hambatan: "Tidak ada hambatan berarti.",
            buktiKegiatan: "BAST Pengadaan Buku & Foto Pojok Baca 12 Kelas",
            evaluasiHasil: "Antusiasme membaca siswa meningkat signifikan pada jam istirahat.",
            tindakLanjut: "Mengadakan lomba resensi buku antar kelas pada semester genap."
          }
        },
        history: [
          {
            id: uid(),
            timestamp: new Date().toLocaleString('id-ID'),
            modul: "Mode Simulasi",
            aksi: "Muat Data Contoh",
            keterangan: "Memuat dataset simulasi lengkap untuk demonstrasi fitur RAPOR SMART AI."
          }
        ]
      };

      triggerAutoSave("Mode Simulasi", "Aktifkan Simulasi", "Data contoh simulasi berhasil dimuat.");
      renderAll();
      showToast("Mode Simulasi aktif! Anda dapat menjelajahi seluruh alur dari Dashboard hingga ARKAS & Monev.", "success");
    }

    // =========================================================================
    // MODULE N: CHATBOT ASISTEN RAPOR SMART AI (STRICTLY DATA-GROUNDED)
    // =========================================================================
    function toggleAssistantDrawer() {
      const drawer = document.getElementById('aiAssistantDrawer');
      const backdrop = document.getElementById('aiAssistantBackdrop');
      const isClosed = drawer.classList.contains('translate-x-full');
      if (isClosed) {
        drawer.classList.remove('translate-x-full');
        backdrop.classList.remove('hidden');
        renderChatMessages();
      } else {
        drawer.classList.add('translate-x-full');
        backdrop.classList.add('hidden');
      }
    }

    function renderChatMessages() {
      const box = document.getElementById('aiChatMessages');
      if (!box) return;
      box.innerHTML = window.appState.chatMessages.map(msg => {
        const isUser = msg.role === 'user';
        // Basic bold & newline formatting
        const formatted = escapeHtml(msg.text)
          .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
          .replace(/\n/g, '<br/>');
        return `
          <div class="flex ${isUser ? 'justify-end' : 'justify-start'}">
            <div class="max-w-[88%] rounded-2xl px-3.5 py-2.5 text-xs leading-relaxed shadow-2xs ${isUser ? 'emerald-gradient text-white rounded-br-sm' : 'bg-white border border-slate-200 text-slate-800 rounded-bl-sm'}">
              ${!isUser ? `<div class="text-[10px] font-extrabold text-emerald-800 mb-1 flex items-center gap-1"><i data-lucide="sparkles" class="w-3 h-3 text-amber-500"></i> ASISTEN RAPOR SMART AI</div>` : ''}
              <div>${formatted}</div>
            </div>
          </div>
        `;
      }).join('');
      box.scrollTop = box.scrollHeight;
      if (window.lucide) window.lucide.createIcons();
    }

    function sendQuickAiPrompt(text) {
      const input = document.getElementById('aiChatInput');
      if (input) {
        input.value = text;
        handleAiChatSubmit(new Event('submit'));
      }
    }

    async function handleAiChatSubmit(e) {
      e.preventDefault();
      const input = document.getElementById('aiChatInput');
      const userText = input.value.trim();
      if (!userText) return;

      input.value = "";
      window.appState.chatMessages.push({ role: 'user', text: userText });
      window.appState.chatMessages.push({ role: 'assistant', text: "⏳ Sedang memeriksa data Rapor Pendidikan & perencanaan sekolah Anda..." });
      renderChatMessages();

      const data = getCurrentData();
      const totalBudget = data.arkas.reduce((s, a) => s + (Number(a.jumlahAnggaran) || 0), 0);

      // Build grounded snapshot of the user's current workspace
      const groundedContext = {
        modeSimulasi: window.appState.isDemoMode,
        sekolah: data.school,
        jumlahIndikator: data.indicators.length,
        indikator: data.indicators.map(i => ({
          kode: i.kode,
          nama: i.nama,
          dimensi: i.dimensi,
          nilai: i.nilai,
          kategori: classifyCategory(i.nilai, i.kategori),
          perubahan: i.perubahan,
          terverifikasi: i.terverifikasi
        })),
        akarMasalahTersedia: Object.keys(data.rootCauses).length,
        jumlahRekomendasi: data.recommendations.length,
        jumlahProgramRKJM: data.rkjm.length,
        jumlahKegiatanRKT: data.rkt.length,
        totalRancanganARKAS: totalBudget,
        rincianARKAS: data.arkas.map(a => ({ kegiatan: a.namaKegiatan, komponen: a.komponenPembiayaan, uraian: a.uraianBelanja, jumlah: a.jumlahAnggaran })),
        monev: data.monev
      };

      const systemPrompt = `Anda adalah ASISTEN RAPOR SMART AI, asisten cerdas perencanaan sekolah berbasis data (PBD) untuk Kepala Sekolah, Guru, Bendahara BOS, Operator, dan Tim Pengembang Sekolah di Indonesia.
PRINSIP WAJIB:
1. Gunakan HANYA data JSON konteks sekolah yang diberikan di bawah ini sebagai dasar fakta.
2. JANGAN PERNAH mengarang angka indikator, nama sekolah, atau nominal anggaran yang tidak ada dalam data.
3. Jika data sekolah masih kosong atau belum lengkap untuk menjawab pertanyaan, sampaikan secara sopan dan minta konfirmasi/input dari pengguna.
4. Bedakan antara hipotesis/saran dengan data aktual sekolah.
5. Gunakan Bahasa Indonesia yang jelas, ramah, terstruktur (poin-poin singkat), dan praktis.`;

      const promptWithContext = `DATA AKTUAL SEKOLAH SAAT INI:\n${JSON.stringify(groundedContext, null, 2)}\n\nPERTANYAAN TIM SEKOLAH:\n${userText}`;

      try {
        const reply = await callGeminiAPI(promptWithContext, systemPrompt);
        window.appState.chatMessages.pop(); // remove loading placeholder
        window.appState.chatMessages.push({ role: 'assistant', text: reply });
        renderChatMessages();
      } catch (err) {
        window.appState.chatMessages.pop();
        // Grounded offline deterministic assistant response
        let fallbackReply = "";
        if (data.indicators.length === 0) {
          fallbackReply = "Saat ini **data Identitas Sekolah dan Indikator Rapor Pendidikan masih kosong**.\n\nAgar saya dapat memberikan analisis yang akurat tanpa mengarang data, silakan:\n1. Buka menu **1. Input Data Rapor** untuk mengunggah PDF/Excel Rapor Pendidikan atau mengisi formulir manual.\n2. Atau klik **Muat Data Demo** di bagian atas jika ingin mencoba simulasi terlebih dahulu.";
        } else {
          const prio = data.indicators.filter(needsImprovement);
          fallbackReply = `Berdasarkan data sekolah **${data.school.nama || '(Nama sekolah belum diisi)'}** (Tahun ${data.school.tahunPerencanaan || window.appState.activeYear}):\n\n` +
            `• **Total Indikator Diinput:** ${data.indicators.length} indikator (${prio.length} indikator memerlukan pembenahan: ${prio.map(i => i.nama).join(', ') || 'tidak ada'}).\n` +
            `• **Status Perencanaan:** ${Object.keys(data.rootCauses).length} analisis akar masalah, ${data.rkjm.length} program RKJM, ${data.rkt.length} kegiatan RKT, dan total Rancangan ARKAS **${formatRupiah(totalBudget)}**.\n\n` +
            `*(Catatan: Layanan AI daring sedang tidak terhubung/offline, sehingga ringkasan ini disusun langsung dari basis data lokal Anda. Silakan lengkapi bagian yang masih kosong pada modul terkait.)*`;
        }
        window.appState.chatMessages.push({ role: 'assistant', text: fallbackReply });
        renderChatMessages();
      }
    }

    // =========================================================================
    // MODULE P: PANDUAN PENGGUNAAN UNTUK PENGGUNA BARU
    // =========================================================================
    function openUserGuideModal() {
      openCustomModal(`
        <div class="p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div class="flex items-center gap-2.5">
              <div class="w-9 h-9 rounded-xl emerald-gradient text-amber-300 flex items-center justify-center">
                <i data-lucide="book-open-check" class="w-5 h-5"></i>
              </div>
              <div>
                <h3 class="font-extrabold text-slate-900 text-base">Panduan Penggunaan RAPOR SMART AI</h3>
                <p class="text-xs text-slate-500">Alur kerja praktis bagi Kepala Sekolah, Guru, Bendahara BOS, &amp; Operator</p>
              </div>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>

          <div class="mt-4 space-y-3 text-xs text-slate-700 leading-relaxed">
            <div class="p-3.5 rounded-xl bg-emerald-50 border border-emerald-200">
              <p class="font-bold text-emerald-950">Prinsip Utama Aplikasi:</p>
              <p class="mt-1">Seluruh tahapan perencanaan saling terhubung dari hulu ke hilir. Perubahan pada Rapor Pendidikan dan RKJM akan terpantau hingga ke kegiatan RKT dan rincian belanja ARKAS.</p>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div class="p-3 rounded-xl border border-slate-200">
                <div class="font-bold text-emerald-900">Langkah 1 — Input Data Rapor</div>
                <p class="mt-1">Isi identitas sekolah (Nama, NPSN, Jenjang) lalu unggah PDF/Excel Rapor Pendidikan atau isi formulir manual. Konfirmasi setiap angka yang diimpor.</p>
              </div>
              <div class="p-3 rounded-xl border border-slate-200">
                <div class="font-bold text-emerald-900">Langkah 2 &amp; 3 — Analisis &amp; Akar Masalah</div>
                <p class="mt-1">Tinjau indikator berkategori Sedang/Kurang/Turun. Gunakan analisis 5-Why lalu konfirmasi hipotesis AI menjadi Fakta Sekolah bersama dewan guru.</p>
              </div>
              <div class="p-3 rounded-xl border border-slate-200">
                <div class="font-bold text-emerald-900">Langkah 4 &amp; 5 — Prioritas &amp; Rekomendasi</div>
                <p class="mt-1">Sesuaikan bobot prioritas pembenahan, lalu hasilkan rekomendasi program kegiatan yang realistis dan setujui program terpilih.</p>
              </div>
              <div class="p-3 rounded-xl border border-slate-200">
                <div class="font-bold text-emerald-900">Langkah 6, 7 &amp; 8 — RKJM, RKT &amp; ARKAS</div>
                <p class="mt-1">Turunkan program ke matriks RKJM (4 tahun), RKT (jadwal 12 bulan), dan Rancangan ARKAS (Volume × Harga Satuan mengacu BOSP).</p>
              </div>
            </div>

            <div class="p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-amber-950">
              <strong>Tips Keamanan Data:</strong> Gunakan tombol <strong>Simpan Proyek (.JSON)</strong> di pojok kanan atas secara berkala untuk menyimpan arsip file perencanaan ke komputer Anda sehingga dapat dibuka kembali kapan saja.
            </div>
          </div>

          <div class="mt-5 flex justify-end">
            <button onclick="closeCustomModal()" class="px-5 py-2.5 rounded-xl emerald-gradient text-white text-xs font-bold">Siap Menggunakan Aplikasi</button>
          </div>
        </div>
      `, 'max-w-2xl');
    }

    // =========================================================================
    // INITIALIZATION ON LOAD
    // =========================================================================
    window.addEventListener('DOMContentLoaded', () => {
      renderAll();
    });
  </script>
</body>
</html>
