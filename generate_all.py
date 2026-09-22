import os
import subprocess
from create_docx import md_to_docx

repo_dir = "/root/.openclaw/workspace/lesson-plan-pembuatan-konten-visual"

lessons_data = [
    {
        "day": 1,
        "title": "Pengantar Content Creator & Ekosistem Media Sosial",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Memahami peran, etika, dan peluang karier sebagai Content Creator di media sosial.",
            "Mengidentifikasi jenis-jenis platform media sosial (Instagram, TikTok, YouTube, LinkedIn) dan karakteristik audiensnya.",
            "Memahami alur kerja dasar pembuatan konten visual profesional."
        ],
        "materi": [
            "Pengenalan Dunia Content Creator & Industri Kreatif Digital",
            "Anatomi Platform Media Sosial & Algoritma Konten Visual",
            "Etika, Hak Cipta (Copyright), dan Lisensi Asset Media",
            "Struktur Kerja & Tools Utama Pembuat Konten Visual"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (20 menit)", "deskripsi": "Fasilitator membuka sesi, menyapa peserta, melakukan ice breaking, dan menyampaikan apersepsi mengenai lanskap media sosial saat ini."},
            {"tahap": "Kegiatan Inti - Teori (60 menit)", "deskripsi": "Penyampaian materi tentang profesi Content Creator, karakteristik platform (Instagram Carousel vs TikTok Reels vs Shorts), dan etika hak cipta gambar/audio."},
            {"tahap": "Kegiatan Inti - Praktik & Diskusi (80 menit)", "deskripsi": "Peserta melakukan bedah akun (account audit) dari 3 creator sukses di bidang berbeda dan menganalisis elemen visual yang digunakan."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Refleksi pembelajaran, sesi tanya jawab, dan pemberian kuis singkat/Test Tertulis Hari 01."}
        ],
        "asesmen": "Kuis Tertulis Hari 01 (Pilihan Ganda/Esai) & Lembar Refleksi Analisis Akun Creator."
    },
    {
        "day": 2,
        "title": "Teknik Fotografi Smartphone untuk Konten Visual",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menguasai fitur-fitur kamera smartphone (ISO, Shutter Speed, Focus, Gridlines, HDR).",
            "Menerapkan komposisi foto dasar (Rule of Thirds, Leading Lines, Framing, Symmetry) untuk sosial media.",
            "Memanfaatkan pencahayaan alami (Natural Light) dan pencahayaan buatan secara efektif."
        ],
        "materi": [
            "Pengenalan Hardware & Setting Kamera Smartphone",
            "Prinsip Komposisi Fotografi Media Sosial",
            "Teknik Pencahayaan (Lighting Setup: Continuous Light, Softbox, Natural Window Light)",
            "Angle Pengambilan Gambar (Eye Level, High Angle, Low Angle, Flatlay/Top-down)"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Review materi pertemuan 1 dan penjelasan tujuan pembelajaran teknik fotografi smartphone."},
            {"tahap": "Kegiatan Inti - Demonstrasi (45 menit)", "deskripsi": "Demonstrasi pengaturan manual kamera hp dan contoh hasil foto berdasarkan komposisi Rule of Thirds & Flatlay."},
            {"tahap": "Kegiatan Inti - Praktik Hands-on (100 menit)", "deskripsi": "Praktik pengambilan foto produk/objek menggunakan hp masing-masing dengan 3 teknik angle dan komposisi berbeda (Jobsheet Hari 02)."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Review hasil foto peserta, umpan balik dari instruktur, dan evaluasi tertulis."}
        ],
        "asesmen": "Jobsheet Hari 02 (Praktik Fotografi Produk/Lifestyle) & Test Evaluasi Fotografi Smartphone."
    },
    {
        "day": 3,
        "title": "Teknik Pengambilan Video dengan Smartphone",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Memahami prinsip pergerakan kamera (Pan, Tilt, Zoom, Track, Panning, Dolly shot).",
            "Menguasai teknik stabilisasi video hp (Handheld technique, Gimbal, Tripod).",
            "Merencanakan dan mengambil footage video pendek yang menarik untuk Reels/TikTok."
        ],
        "materi": [
            "Prinsip Dasar Sinematografi Smartphone (Resolution 1080p vs 4K, Frame Rate 24fps/30fps/60fps)",
            "Jenis Shot (B-Roll, A-Roll, Close-Up, Medium Shot, Wide Shot)",
            "Teknik Pergerakan Kamera (Camera Movement & Stabilization)",
            "Manajemen File Footage & Storyboarding Ringkas"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pembahasan konsep video pendek di media sosial dan pentingnya variasi shot."},
            {"tahap": "Kegiatan Inti - Teori & Penjelasan (45 menit)", "deskripsi": "Penjelasan mengenai frame rate, resolusi, jenis shot (Establishing shot hingga Detail shot), dan pergerakan kamera."},
            {"tahap": "Kegiatan Inti - Praktik Pengambilan Gambar (100 menit)", "deskripsi": "Peserta melakukan simulasi pengambilan B-Roll dan A-Roll sesuai panduan Jobsheet Hari 03."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Evaluasi bersama footage hasil rekam peserta dan pengisian Test Tertulis Hari 03."}
        ],
        "asesmen": "Jobsheet Hari 03 (Tugas Shooting Video Pendek B-Roll) & Kuis Test Tertulis Hari 03."
    },
    {
        "day": 4,
        "title": "Dasar Copywriting & Psikologi Pembaca",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Memahami psikologi audiens media sosial dan pemicu emosi (Emotional Triggers).",
            "Menguasai formula penulisan naskah/caption populer (AIDA, PAS, BAB).",
            "Menulis Hook, Body, dan Call to Action (CTA) yang dapat meningkatkan engagemen."
        ],
        "materi": [
            "Pengenalan Psikologi Pembaca & Target Audiens",
            "Formula Copywriting Utama: AIDA (Attention, Interest, Desire, Action) & PAS (Problem, Agitate, Solution)",
            "Teknik Pembuatan Hook (Judul/3 Detik Pertama) yang Menarik",
            "Penyusunan Call to Action (CTA) dan Caption Struktural"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Diskusi interaktif mengenai caption atau iklan sosmed yang paling berkesan bagi peserta."},
            {"tahap": "Kegiatan Inti - Teori (60 menit)", "deskripsi": "Pemaparan materi formula copywriting AIDA/PAS, perbedaan naskah visual vs caption, serta psikologi warna/kata penarik."},
            {"tahap": "Kegiatan Inti - Latihan Penulisan (85 menit)", "deskripsi": "Latihan membuat 5 variasi Hook dan naskah copywriting iklan produk lokal sesuai Jobsheet Hari 04."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Presentasi singkat karya tulisan peserta, masukan instruktur, dan pengerjaan Test Tertulis Hari 04."}
        ],
        "asesmen": "Jobsheet Hari 04 (Latihan Penulisan Copywriting AIDA & Hook) & Kuis Test Tertulis Hari 04."
    },
    {
        "day": 5,
        "title": "Design Thinking dalam Pembuatan Konten Visual",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menerapkan 5 tahapan Design Thinking (Empathize, Define, Ideate, Prototype, Test) dalam pembuatan konten.",
            "Membuat User Persona dan Map Kebutuhan Audiens.",
            "Menghasilkan ide konten yang solutif, kreatif, dan relevan dengan audiens sasaran."
        ],
        "materi": [
            "Konsep Dasar Design Thinking untuk Media Sosial",
            "Tahap Empathize: Riset & Analisis Kebutuhan Audiens (User Persona)",
            "Tahap Define & Ideate: Problem Statement & Brainstorming Ide Konten",
            "Tahap Prototype & Test: Validasi Ide Konten Sebelum Produksi Massal"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pengenalan pentingnya pemikiran terstruktur berbasis emosi manusia (Design Thinking) dalam pembuatan konten."},
            {"tahap": "Kegiatan Inti - Ideasi Berkelompok (120 menit)", "deskripsi": "Peserta dibagi menjadi kelompok kecil untuk memetakan masalah audiens, merumuskan User Persona, dan melakukan brainstorming 10 ide konten visual."},
            {"tahap": "Penutup & Presentasi (45 menit)", "deskripsi": "Presentasi ide ideasi kelompok, umpan balik silang antarkelompok, serta penyimpulan materi."}
        ],
        "asesmen": "Lembar Kerja User Persona & Pemetaan Ide Konten Solutif (Design Thinking Canvas)."
    },
    {
        "day": 6,
        "title": "Prinsip Desain Grafis: Warna, Tipografi & Layout",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Memahami elemen dasar desain grafis (Garis, Bentuk, Warna, Tekstur, Ruang).",
            "Menguasai Teori Warna (Color Wheel, Harmoni Warna, Psikologi Warna).",
            "Menerapkan Hierarki Tipografi dan Prinsip Layout (Alignment, Contrast, Proximity, Balance)."
        ],
        "materi": [
            "Elemen & Prinsip Dasar Desain Grafis",
            "Teori & Psikologi Warna (Color Palette, Contrast Ratio)",
            "Tipografi (Serif, Sans-Serif, Font Pairing, Hierarchy, Kerning, Leading)",
            "Prinsip Layout & Compositional Grid"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pembedahan contoh desain buruk vs desain profesional untuk memahami pentingnya teori visual."},
            {"tahap": "Kegiatan Inti - Teori & Studi Kasus (60 menit)", "deskripsi": "Pemaparan materi roda warna, kombinasi font, dan tata letak hierarki teks."},
            {"tahap": "Kegiatan Inti - Praktik Layout (85 menit)", "deskripsi": "Peserta mengerjakan Jobsheet Hari 06 Desain Grafis & Warna serta Tipografi & Layout (membuat moodboard warna dan penataan teks)."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Review karya peserta, pembahasan kesalahan umum layout, dan pengerjaan Test Tertulis Hari 06."}
        ],
        "asesmen": "Jobsheet Hari 06 (Desain Grafis & Warna; Tipografi & Layout) & Test Tertulis Hari 06."
    },
    {
        "day": 7,
        "title": "Canva Basic, Branding & Identitas Visual",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menguasai antarmuka, navasi, dan fitur-fitur dasar Canva.",
            "Memahami konsep Brand Identity (Logo, Color Palette, Typography Style, Brand Guidelines).",
            "Membangun Brand Kit awal untuk kebutuhan media sosial di Canva."
        ],
        "materi": [
            "Pengenalan Workspace Canva (Tools, Templates, Elements, Canvas Size)",
            "Dasar Branding & Identitas Visual Perusahaan/Personal Brand",
            "Penyusunan Brand Kit (Logo, Palette Warna, Kombinasi Font di Canva)",
            "Teknik Ekspor Asset & Format File (PNG, JPG, SVG, PDF)"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pentingnya konsistensi visual branding di seluruh channel media sosial."},
            {"tahap": "Kegiatan Inti - Demonstrasi Canva (45 menit)", "deskripsi": "Instruktur mendemonstrasikan eksplorasi alat Canva, pengaturan Brand Kit, dan pembuatan aset pertama."},
            {"tahap": "Kegiatan Inti - Praktik Mandiri (100 menit)", "deskripsi": "Peserta mengerjakan Jobsheet Hari 07 Canva Basic dan Jobsheet Branding & Identitas Visual."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Pemeriksaan Brand Kit peserta dan pengerjaan Test Tertulis Hari 07."}
        ],
        "asesmen": "Jobsheet Hari 07 (Canva Basic & Brand Kit Design) serta Test Tertulis Hari 07."
    },
    {
        "day": 8,
        "title": "Desain Konten Visual Media Sosial & Advanced Canva",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menggunakan fitur-fitur lanjutan Canva (Background Remover, Magic Studio, Frame, Animation, Transparency).",
            "Merancang konten Microblog / Carousel Instagram yang berkesinambungan (Seamless Carousel).",
            "Membuat desain poster promosional dan Story Instagram berdaya konversi tinggi."
        ],
        "materi": [
            "Fitur Advanced Canva (Magic Edit, Background Remover, Grids & Frames, Ruler & Guides)",
            "Teknik Merancang Carousel Seamless Instagram (Multi-page Layout)",
            "Desain Promotional Banner & Story Engagement (Poll, Quiz Frame)",
            "Standardisasi Ukuran Feed (1:1, 4:5) dan Story/Reels (9:16)"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Tren desain visual media sosial saat ini: Carousel edukatif & visual promosi interaktif."},
            {"tahap": "Kegiatan Inti - Demonstrasi Fitur Advanced (45 menit)", "deskripsi": "Tutorial pembuatan Seamless Carousel di Canva dan penggunaan Magic Tools."},
            {"tahap": "Kegiatan Inti - Praktik (100 menit)", "deskripsi": "Peserta membuat 1 set Seamless Carousel (5 slide) sesuai skenario Jobsheet Hari 08."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Pemeriksaan hasil desain slide carousel peserta dan pengerjaan Test Tertulis Hari 08."}
        ],
        "asesmen": "Jobsheet Hari 08 (Advanced Canva & Konten Visual Sosmed) & Kuis Test Tertulis Hari 08."
    },
    {
        "day": 9,
        "title": "Prototipe Desain & Narasi Visual (Storytelling)",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menyusun alur narasi visual (Visual Storytelling) yang menyentuh emosi audiens.",
            "Membuat prototipe desain konten lengkap (Feed Carousel + Story + Cover Video).",
            "Melakukan pengujian awal prototipe (Usability & Appeal Test) kepada sesama peserta."
        ],
        "materi": [
            "Prinsip Narasi Visual (Visual Storytelling Arc: Hook, Conflict, Resolution, Action)",
            "Prototyping Konten Visual (Mockup Display & Visual Consistency Check)",
            "Teknik Feedback & Peer Review Hasil Prototipe Desain",
            "Iterasi & Penyempurnaan Prototipe Konten Visual"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pentingnya alur cerita (story arc) dalam mempertahankan perhatian audiens."},
            {"tahap": "Kegiatan Inti - Penyusunan Prototipe (105 menit)", "deskripsi": "Peserta menyusun prototipe narasi visual lengkap berdasarkan Jobsheet Hari 09 (Prototipe Desain & Narasi Visual)."},
            {"tahap": "Kegiatan Inti - Uji Praktik & Peer Review (40 menit)", "deskripsi": "Pelaksanaan Test Praktik Hari 09: Mempresentasikan prototipe konten kepada teman sebaya untuk mendapat masukan."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Penyimpulan evaluasi dan pengerjaan Test Tertulis Hari 09."}
        ],
        "asesmen": "Jobsheet Hari 09 (Prototipe Desain & Narasi Visual), Test Praktik Hari 09, dan Test Tertulis Hari 09."
    },
    {
        "day": 10,
        "title": "Pengambilan Video Project (Shooting Day)",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menerapkan seluruh teknik fotografi & sinematografi dalam skenario produksi nyata.",
            "Mengambil klip A-Roll dan B-Roll sesuai dengan script/storyboard yang direncanakan.",
            "Mengelola proses produksi video lapangan secara mandiri dan disiplin."
        ],
        "materi": [
            "Pengecekan Pre-Production (Equipment Check, Storyboard Review, Sound Check)",
            "Eksekusi Shooting Video (Pengambilan Talent, Produk, B-Roll Ambient, Voice Over Directing)",
            "Manajemen File Media Simpanan (Offloading & Organizing Footage)",
            "Quality Assurance Video Lapangan (Focus, Exposure, Audio Levels Check)"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Briefing keselamatan kerja, pembagian area shooting, dan review checklist shot list."},
            {"tahap": "Kegiatan Inti - Shooting Lapangan / Studio (135 menit)", "deskripsi": "Peserta melaksanakan produksi pengambilan video mandiri/kelompok sesuai skenario Test Praktik Hari 10."},
            {"tahap": "Kegiatan Inti - Review Asset (15 menit)", "deskripsi": "Memastikan semua shotlist terisi dan tidak ada file rusak/blur."},
            {"tahap": "Penutup (15 menit)", "deskripsi": "Penutupan sesi shooting, pemindahan file ke penyimpanan kerja, dan pemberian arahan untuk sesi editing."}
        ],
        "asesmen": "Test Praktik Hari 10 (Penilaian Keberhasilan Pengambilan Footage Video Project)."
    },
    {
        "day": 11,
        "title": "Editing Video Pendek dengan CapCut (Pengenalan & Basic Editing)",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menguasai antarmuka, timeline, dan alur kerja aplikasi CapCut (Mobile/Desktop).",
            "Melakukan pemotongan video (Trim, Split, Cut on Action, Rough Cut).",
            "Menambahkan musik latar (Backsound), sound effect (SFX), dan Teks/Subtitle otomatis."
        ],
        "materi": [
            "Pengenalan Interface CapCut (Timeline, Media Import, Tracks, Tools Bar)",
            "Teknik Potong Video (Split, Trim, Speed Curve Basics, Ratio 9:16)",
            "Pengaturan Audio (Background Music Balancing, Fade In/Out, Sound Effects)",
            "Auto Captions / Text Overlay & Basic Transitions"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pengenalan pentingnya post-production editing dalam membentuk irama konten video pendek."},
            {"tahap": "Kegiatan Inti - Tutorial CapCut (45 menit)", "deskripsi": "Instruktur memperagakan langkah mengimpor footage hasil shooting Hari 10, cutting, dan pemberian musik."},
            {"tahap": "Kegiatan Inti - Praktik Editing (100 menit)", "deskripsi": "Peserta mengedit video pertama mereka sesuai arahan Jobsheet Hari 11."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Review hasil draf video awal peserta dan pengerjaan Test Tertulis Hari 11."}
        ],
        "asesmen": "Jobsheet Hari 11 (Editing Video CapCut) & Kuis Test Tertulis Hari 11."
    },
    {
        "day": 12,
        "title": "Audio Enhancement, Motion Graphics & Advanced CapCut",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menerapkan teknik Advanced CapCut (Keyframe Animation, Masking, Chroma Key/Green Screen).",
            "Melakukan ekualisasi dan pembersihan noise audio (Noise Reduction & Voice Over Balancing).",
            "Membuat efek motion graphics dan teks dinamis untuk meningkatkan nilai estetika video."
        ],
        "materi": [
            "Fitur Advanced CapCut: Keyframe Animation (Zoom in/out halus, Movement)",
            "Teknik Masking, Overlay, dan Chroma Key (Green Screen Removal)",
            "Audio Mixing & Noise Reduction untuk Suara Jernih",
            "Motion Graphics Teks, Stiker Animasi, dan Tracking Motion"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Demonstrasi perbedaan video biasa vs video beranimasi motion graphics yang dinamis."},
            {"tahap": "Kegiatan Inti - Tutorial Advanced (45 menit)", "deskripsi": "Penjelasan penggunaan Keyframe untuk animasi kustom dan teknik masking di CapCut."},
            {"tahap": "Kegiatan Inti - Praktik (100 menit)", "deskripsi": "Peserta menerapkan keyframe, efek motion graphics, dan pembersihan audio sesuai Jobsheet Hari 12."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Showcase singkat beberapa karya video peserta dan pengerjaan Test Tertulis Hari 12."}
        ],
        "asesmen": "Jobsheet Hari 12 (Advanced CapCut & Audio Motion Graphics) serta Test Tertulis Hari 12."
    },
    {
        "day": 13,
        "title": "Integrasi Canva-CapCut & Teknik Ekspor Video Optimal",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Mengintegrasikan aset grafis dari Canva (PNG Transparent/Overlays) ke dalam projek CapCut.",
            "Memahami parameter rendering dan ekspor video (Bitrate, Resolution, FPS, Codec).",
            "Mengekspor video tanpa watermark dengan kualitas jernih siap diunggah ke platform sosmed."
        ],
        "materi": [
            "Workflow Integrasi Canva ke CapCut (Exporting Assets as PNG Transparent & Video Elements)",
            "Desain Frame, Lower Thirds, & End Screen di Canva untuk Diimpor ke CapCut",
            "Pemahaman Setting Ekspor (1080p, 60fps vs 30fps, High Bitrate, Color Space)",
            "Troubleshooting Video Compression di Instagram & TikTok"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Memahami pentingnya alur kerja antar-aplikasi untuk mempermudah produksi konten profesional."},
            {"tahap": "Kegiatan Inti - Tutorial & Praktik (120 menit)", "deskripsi": "Peserta mendesain aset Lower Third & Frame di Canva lalu diintegrasikan ke video CapCut, dilanjutkan ekspor video sesuai Jobsheet Hari 13."},
            {"tahap": "Penutup (45 menit)", "deskripsi": "Review video akhir yang sudah diekspor dan pengerjaan Test Tertulis Hari 13."}
        ],
        "asesmen": "Jobsheet Hari 13 (Integrasi Canva-CapCut & Ekspor Video PDF) & Test Tertulis Hari 13."
    },
    {
        "day": 14,
        "title": "Project Marketplace, Monetisasi Konten & Link-in-Bio",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Memahami strategi monetisasi konten (Affiliate Marketing, Endorsement, Digital Product).",
            "Merancang halaman Link-in-Bio (Lynk.id / Milkshake / Taplink) yang menarik konversi.",
            "Membuat konten visual promosi khusus untuk produk marketplace/affiliate."
        ],
        "materi": [
            "Pengenalan Ekosistem Monetisasi Content Creator & Affiliate Marketing",
            "Desain & Optimalisasi Landing Page Link-in-Bio",
            "Strategi Visual Promosi Produk Marketplace (Shopee/TikTok Shop Content)",
            "Manajemen Call-to-Action & Tracking Conversion Link"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pemaparan potensi pendapatan seorang Content Creator melalui affiliate & marketplace."},
            {"tahap": "Kegiatan Inti - Pembahasan Slide & Praktik (120 menit)", "deskripsi": "Penyampaian materi Slide Hari 14, pengerjaan Jobsheet Link-in-bio dan Jobsheet Marketplace & Monetisasi."},
            {"tahap": "Penutup (45 menit)", "deskripsi": "Pemeriksaan tautan Link-in-bio milik peserta dan pengerjaan Test Tertulis Hari 14."}
        ],
        "asesmen": "Jobsheet Hari 14 (Project Marketplace & Link-in-Bio) & Test Tertulis Hari 14."
    },
    {
        "day": 15,
        "title": "Canva Content Planner, Audiens & Tone of Voice",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menentukan Tone of Voice (Gaya Bahasa/Sikap) merek yang sesuai dengan audiens sasaran.",
            "Menggunakan fitur Content Planner & Scheduling di Canva.",
            "Menyusun Kalender Konten (Content Calendar) bulanan yang teratur."
        ],
        "materi": [
            "Definisi & Formulasi Tone of Voice (Formal, Kasual, Humoris, Edukatif)",
            "Penyusunan Content Matrix & Content Pillar (Edukasi, Hiburan, Promosi, Soft-selling)",
            "Penggunaan Canva Content Planner (Scheduling & Direct Publishing)",
            "Manajemen Workflow Publishing & Batch Content Creation"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Mengapa konsistensi jadwal posting dan tone of voice sangat penting bagi pertumbuhan akun."},
            {"tahap": "Kegiatan Inti - Diskusi Slide & Praktik (120 menit)", "deskripsi": "Penyampaian Slide Hari 15, pengerjaan Jobsheet Audiens & Tone of Voice serta Jobsheet Content Planner & Brand Kit."},
            {"tahap": "Penutup (45 menit)", "deskripsi": "Presentasi Kalender Konten bulanan peserta dan pengerjaan Test Tertulis Hari 15."}
        ],
        "asesmen": "Jobsheet Hari 15 (Audiens & Tone of Voice, Content Planner) & Test Tertulis Hari 15."
    },
    {
        "day": 16,
        "title": "Storytelling lanjutan, SEO Copywriting & Email Marketing",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menerapkan teknik SEO Copywriting (Keyword Optimization) pada caption dan bio media sosial.",
            "Menyusun naskah Storytelling mendalam untuk memperkuat ikatan emosional (Engagement).",
            "Memahami dasar penulisan Email Marketing untuk penawaran produk visual."
        ],
        "materi": [
            "Dasar-dasar SEO Media Sosial (Instagram SEO, TikTok Search Keyword Integration)",
            "Teknik Storytelling Lanjutan (Hero's Journey, Relatable Struggle Stories)",
            "Struktur Penulisan Email Marketing (Subject Line, Body Text, CTA Button)",
            "Integrasi Narasi Konten Visual dengan Saluran Email & Copywriting"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Penjelasan pentingnya ditemukannya konten melalui mesin pencari sosmed (SEO)."},
            {"tahap": "Kegiatan Inti - Praktik Penulisan (120 menit)", "deskripsi": "Pengerjaan Jobsheet Hari 16 Storytelling & Headline, Jobsheet SEO Copywriting, dan Jobsheet Email Marketing."},
            {"tahap": "Penutup (45 menit)", "deskripsi": "Review bersama naskah email & penulisan caption ber-SEO, diikuti Test Tertulis Hari 16."}
        ],
        "asesmen": "Jobsheet Hari 16 (SEO Copywriting, Storytelling & Email Marketing) serta Test Tertulis Hari 16."
    },
    {
        "day": 17,
        "title": "Pemanfaatan AI untuk Copywriting, Storyboard & Personal Branding",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Menggunakan Generative AI (ChatGPT/Claude/Gemini) untuk riset ide, ideasi storyboard, dan naskah konten.",
            "Menyusun Prompt Engineering yang presisi untuk hasil copywriting yang berkualitas tinggi.",
            "Membangun strategi Personal Branding diri/brand menggunakan bantuan AI."
        ],
        "materi": [
            "Pengenalan Tools AI untuk Content Creator (ChatGPT, Midjourney/DALL-E, Gemini)",
            "Teknik Prompt Engineering untuk Penulisan Copywriting & Outline Konten",
            "Ideasi & Pembuatan Storyboard Visual Otomatis Berbasis AI",
            "Strategi Personal Branding & Citra Diri Digital di Media Sosial"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pengenalan revolusi AI dalam mempercepat alur kerja pembuatan konten tanpa mengorbankan kualitas."},
            {"tahap": "Kegiatan Inti - Pemaparan Slide & Eksperimen AI (120 menit)", "deskripsi": "Pembedahan Lembar Informasi Hari 17 & Slide, dilanjutkan pengerjaan Jobsheet AI Copywriting, Ideasi Storyboard AI, dan Personal Branding."},
            {"tahap": "Penutup (45 menit)", "deskripsi": "Demonstrasi prompt AI terbaik peserta dan pengerjaan Test Tertulis Hari 17."}
        ],
        "asesmen": "Jobsheet Hari 17 (AI Copywriting, Ideasi Storyboard AI, Personal Branding) & Test Tertulis Hari 17."
    },
    {
        "day": 18,
        "title": "Final Project: Perencanaan & Produksi Konten Visual Terintegrasi",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Merencanakan Kampanye Konten Visual Terintegrasi (Terdiri dari Carousel Feed, Reels/Shorts Video, dan Story).",
            "Menerapkan seluruh pengetahuan dari Hari 1 - 17 dalam satu projek komprehensif.",
            "Melakukan eksekusi produksi (shooting, grafis, copywriting) secara mandiri."
        ],
        "materi": [
            "Briefing Final Project: Pembuatan Kampanye Konten Visual Produk/Brand",
            "Penyusunan Project Proposal (Target Audiens, Concept, Storyboard, Asset List)",
            "Eksekusi Produksi Asset Visual (Foto, Desain Canva, Video Editing CapCut)",
            "Penerapan AI & Copywriting pada Final Campaign Assets"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (20 menit)", "deskripsi": "Penjelasan petunjuk teknis, rubrik penilaian, dan tenggat waktu Final Project."},
            {"tahap": "Kegiatan Inti - Work Sessi / Workshop (140 menit)", "deskripsi": "Peserta bekerja secara intensif menyelesaikan perencanaan dan draf aset konten terintegrasi dengan bimbingan instruktur."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Check-point kemajuan projek setiap peserta dan pengarahan untuk sesi polishing di pertemuan 19."}
        ],
        "asesmen": "Draf Perencanaan Final Project & Lembar Kemajuan Kerja (Progress Checklist)."
    },
    {
        "day": 19,
        "title": "Final Project: Finishing, Polishing & Quality Control",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Melakukan proses editing akhir (finishing) pada seluruh aset gambar dan video.",
            "Melakukan Quality Control (QC) terhadap kesesuaian brand, ketepatan tata bahasa, dan estetika visual.",
            "Menyiapkan berkas presentasi dan portofolio digital."
        ],
        "materi": [
            "Prinsip Quality Control Konten Digital (Visual Polish, Typo Check, Color Consistency, Audio Clarity)",
            "Penyusunan Dokumen Portofolio Karya Konten Visual",
            "Teknik Presentasi & Pitching Karya Kreatif",
            "Persiapan Media & Berkas Ujian Akhir"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pengarahan standar kualitas (Quality Control) sebelum karya dipublikasikan atau dipresentasikan."},
            {"tahap": "Kegiatan Inti - Polishing & Assembly (145 menit)", "deskripsi": "Peserta melakukan finishing akhir pada aset visual, menyusun slide presentasi, dan mengunggah hasil ke repositori/drive portofolio."},
            {"tahap": "Penutup (20 menit)", "deskripsi": "Simulasi singkat presentasi dan pengecekan kelengkapan berkas akhir peserta."}
        ],
        "asesmen": "Lembar Quality Control Portofolio & Kelengkapan Aset Final Project."
    },
    {
        "day": 20,
        "title": "Presentasi Final Project, Evaluasi Portofolio & Penutupan Pelatihan",
        "durasi": "4 JP (180 Menit)",
        "tujuan": [
            "Mempresentasikan hasil kampanye konten visual terintegrasi di hadapan penguji dan rekan peserta.",
            "Menerima masukan dan evaluasi konstruktif terhadap karya portofolio yang dihasilkan.",
            "Merumuskan rencana pengembangan karier/usaha pasca-pelatihan di bidang pembuatan konten visual."
        ],
        "materi": [
            "Sesi Presentasi Karya Final Project (Showcasing Campaign Assets)",
            "Evaluasi & Umpan Balik Instruktur/Penguji (Review Portofolio)",
            "Refleksi Hasil Belajar 20 Pertemuan Pelatihan",
            "Rencana Aksi Karier (Action Plan Content Creator/Digital Marketer)"
        ],
        "kegiatan": [
            {"tahap": "Pendahuluan (15 menit)", "deskripsi": "Pembukaan sesi ujian akhir presentasi dan pengarahan tata tertib presentasi."},
            {"tahap": "Kegiatan Inti - Presentasi & Uji Portofolio (140 menit)", "deskripsi": "Setiap peserta mempresentasikan kampanye konten visualnya selama 7-10 menit dilanjutkan tanya jawab dan umpan balik penguji."},
            {"tahap": "Penutup & Ceremonial (25 menit)", "deskripsi": "Pengumuman hasil evaluasi, penyerahan apresiasi karya terbaik, dan penutupan resmi pelatihan."}
        ],
        "asesmen": "Rubrik Penilaian Presentasi Final Project & Evaluasi Portofolio Komprehensif (Ujian Akhir Pelatihan)."
    }
]

def generate_md(lesson):
    d = lesson['day']
    tujuan_str = "\n".join([f"{idx+1}. {t}" for idx, t in enumerate(lesson['tujuan'])])
    materi_str = "\n".join([f"- {m}" for m in lesson['materi']])
    
    kegiatan_rows = ""
    for k in lesson['kegiatan']:
        kegiatan_rows += f"| {k['tahap']} | {k['deskripsi']} |\n"
        
    md_content = f"""# LESSON PLAN / RENCANA PELAKSANAAN PEMBELAJARAN (RPP)

## PERTEMUAN {d:02d}: {lesson['title'].upper()}

---

### I. INFORMASI UMUM
- **Nama Pelatihan:** Pembuatan Konten Visual untuk Media Sosial
- **Pertemuan Ke-:** {d:02d} dari 20 Pertemuan
- **Alokasi Waktu:** {lesson['durasi']}
- **Target Peserta:** Pemula / Content Creator Muda / Pelaku UMKM / Designer Pemula
- **Modul / Referensi:** Materi Google Drive Hari {d:02d} (Jobsheet, Test Tertulis, Slide & Modul SKKNI)

---

### II. CAPAIAN & TUJUAN PEMBELAJARAN
Setelah mengikuti sesi pembelajaran ini, peserta didik diharapkan mampu:
{tujuan_str}

---

### III. POKOK BAHASAN / MATERI PEMBELAJARAN
{materi_str}

---

### IV. METODE, MEDIA & ALAT PEMBELAJARAN
- **Metode Pembelajaran:** Ceramah Interaktif, Demonstrasi Praktik, Problem-Based Learning (PBL), Workshop Hands-on, dan Peer Review.
- **Media Pembelajaran:** Slide Presentasi, Contoh Konten Visual, Lembar Jobsheet Hari {d:02d}, dan Kuis Interactive.
- **Alat & Perangkat:** Smartphone (Android/iOS), Laptop, Aplikasi Canva / CapCut, Koneksi Internet, Tripod, dan Lighting.

---

### V. SKENARIO / KEGIATAN PEMBELAJARAN

| Tahapan Kegiatan | Deskripsi Aktivitas Pembelajaran |
| :--- | :--- |
{kegiatan_rows}

---

### VI. PENILAIAN & ASESMEN (EVALUASI)
- **Jenis Asesmen:** Formatif & Sumatif
- **Bentuk Asesmen:** {lesson['asesmen']}
- **Rubrik Penilaian:**
  - **Sikap & Kehadiran (15%):** Kedisiplinan, keaktifan, dan etika kerja.
  - **Pemahaman Teori (35%):** Hasil Test Tertulis / Kuis Hari {d:02d}.
  - **Keterampilan Praktik (50%):** Kualitas penyelesaian Jobsheet / Karya Konten Visual.

---

### VII. TUGAS MANDIRI & TINDAK LANJUT
1. Menyelesaikan seluruh instruksi kerja pada **Jobsheet Hari {d:02d}**.
2. Mengunggah atau mendokumentasikan hasil latihan ke folder portofolio digital masing-masing.
3. Membaca dan mempelajari materi pengantar untuk Pertemuan {d+1:02d} berikutnya.
"""
    return md_content

def main():
    os.makedirs(repo_dir, exist_ok=True)
    
    # Write README.md
    readme_content = """# Lesson Plan Pelatihan Pembuatan Konten Visual untuk Media Sosial (20 Pertemuan)

Repository ini berisi **Lesson Plan / Rencana Pelaksanaan Pembelajaran (RPP)** komprehensif sebanyak **20 Pertemuan** untuk program pelatihan **Pembuatan Konten Visual untuk Media Sosial**. 

Materi dan struktur kurikulum disusun berdasarkan dokumen standar industri, materi SKKNI, serta sumber pembelajaran pada repositori Google Drive pelatihan.

## 📚 Daftar Struktur Pertemuan (20 Lesson Plans)

| Pertemuan | Judul Pembelajaran | Format File |
| :---: | :--- | :---: |
| **Pertemuan 01** | Pengantar Content Creator & Ekosistem Media Sosial | [.md](./pertemuan-01.md) \| [.docx](./pertemuan-01.docx) |
| **Pertemuan 02** | Teknik Fotografi Smartphone untuk Konten Visual | [.md](./pertemuan-02.md) \| [.docx](./pertemuan-02.docx) |
| **Pertemuan 03** | Teknik Pengambilan Video dengan Smartphone | [.md](./pertemuan-03.md) \| [.docx](./pertemuan-03.docx) |
| **Pertemuan 04** | Dasar Copywriting & Psikologi Pembaca | [.md](./pertemuan-04.md) \| [.docx](./pertemuan-04.docx) |
| **Pertemuan 05** | Design Thinking dalam Pembuatan Konten Visual | [.md](./pertemuan-05.md) \| [.docx](./pertemuan-05.docx) |
| **Pertemuan 06** | Prinsip Desain Grafis: Warna, Tipografi & Layout | [.md](./pertemuan-06.md) \| [.docx](./pertemuan-06.docx) |
| **Pertemuan 07** | Canva Basic, Branding & Identitas Visual | [.md](./pertemuan-07.md) \| [.docx](./pertemuan-07.docx) |
| **Pertemuan 08** | Desain Konten Visual Media Sosial & Advanced Canva | [.md](./pertemuan-08.md) \| [.docx](./pertemuan-08.docx) |
| **Pertemuan 09** | Prototipe Desain & Narasi Visual (Storytelling) | [.md](./pertemuan-09.md) \| [.docx](./pertemuan-09.docx) |
| **Pertemuan 10** | Pengambilan Video Project (Shooting Day) | [.md](./pertemuan-10.md) \| [.docx](./pertemuan-10.docx) |
| **Pertemuan 11** | Editing Video Pendek dengan CapCut (Basic Editing) | [.md](./pertemuan-11.md) \| [.docx](./pertemuan-11.docx) |
| **Pertemuan 12** | Audio Enhancement, Motion Graphics & Advanced CapCut | [.md](./pertemuan-12.md) \| [.docx](./pertemuan-12.docx) |
| **Pertemuan 13** | Integrasi Canva-CapCut & Teknik Ekspor Video Optimal | [.md](./pertemuan-13.md) \| [.docx](./pertemuan-13.docx) |
| **Pertemuan 14** | Project Marketplace, Monetisasi Konten & Link-in-Bio | [.md](./pertemuan-14.md) \| [.docx](./pertemuan-14.docx) |
| **Pertemuan 15** | Canva Content Planner, Audiens & Tone of Voice | [.md](./pertemuan-15.md) \| [.docx](./pertemuan-15.docx) |
| **Pertemuan 16** | Storytelling Lanjutan, SEO Copywriting & Email Marketing | [.md](./pertemuan-16.md) \| [.docx](./pertemuan-16.docx) |
| **Pertemuan 17** | Pemanfaatan AI untuk Copywriting, Storyboard & Personal Branding | [.md](./pertemuan-17.md) \| [.docx](./pertemuan-17.docx) |
| **Pertemuan 18** | Final Project: Perencanaan & Produksi Konten Visual Terintegrasi | [.md](./pertemuan-18.md) \| [.docx](./pertemuan-18.docx) |
| **Pertemuan 19** | Final Project: Finishing, Polishing & Quality Control | [.md](./pertemuan-19.md) \| [.docx](./pertemuan-19.docx) |
| **Pertemuan 20** | Presentasi Final Project, Evaluasi Portofolio & Penutupan Pelatihan | [.md](./pertemuan-20.md) \| [.docx](./pertemuan-20.docx) |

## 🛠️ Format Dokumen
Setiap pertemuan disediakan dalam 2 format dokumen:
1. **Markdown (`.md`):** Cocok untuk dokumentasi web, GitHub, dan pembacaan online cepat.
2. **Microsoft Word (`.docx`):** Diformat rapi dengan tata letak profesional, tabel berwarna, dan siap cetak untuk keperluan administratif/instruktur.

---
*Dibuat secara otomatis & terstruktur oleh Arrau AI Assistant untuk Arry (2026).*
"""
    with open(os.path.join(repo_dir, "README.md"), "w", encoding="utf-8") as f:
        f.write(readme_content)
        
    for l in lessons_data:
        d = l['day']
        md_file = os.path.join(repo_dir, f"pertemuan-{d:02d}.md")
        docx_file = os.path.join(repo_dir, f"pertemuan-{d:02d}.docx")
        
        md_content = generate_md(l)
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_content)
            
        md_to_docx(md_file, docx_file)
        print(f"Generated Pertemuan {d:02d}: {md_file} & {docx_file}")

if __name__ == "__main__":
    main()
