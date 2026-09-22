import os
import docx

REPO_DIR = "/root/.openclaw/workspace/lesson-plan-pembuatan-konten-visual"
MASTER_DOCX = "/root/.openclaw/workspace/lessonplan/master_lessonplan.docx"

all_meetings = [
    {
        "no": 1,
        "judul": "Pengantar Content Creator & Ekosistem Media Sosial",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam memahami profesi content creator, karakteristik platform media sosial, dan etika hak cipta konten visual secara tepat.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point, Slide Pengantar & Video Contoh Konten",
        "alat_list": "Alat Proyektor dan Layar, Perangkat Komputer / HP, Papan Tulis, Spidol dan Penghapus",
        "bahan_list": "Flashdisk, Modul Pengantar Content Creator, Form Kuis Hari 01",
        "persiapan": "Persiapan ruangan laboratorium TIK\nPersiapan alat peraga & file presentasi",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction dan tata tertib laboratorium TIK\n4. Review pemicu: Peluang karier dan peran Content Creator di industri digital saat ini\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Pengenalan Profesi Content Creator & Lanskap Sosmed",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Peran dan etika Content Creator di media sosial\n- Karakteristik platform (Instagram, TikTok, YouTube, LinkedIn) & algoritma dasarnya\n- Anatomi konten visual: Hook, Body, dan Call-to-Action (CTA)\nInstruktur memberikan kesempatan peserta untuk mengajukan pertanyaan.",
        "sub1_media": "Power Point & Video Contoh Konten",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Etika, Copyright, & Lisensi Aset Visual",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Aturan hak cipta (Copyright), lisensi bebas royalti (CC0/Unsplash/Freepik), dan Fair Use\n- Cara mencari aset visual gratis tanpa melanggar etika\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Browser Internet & Presentation Slide",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Audit Akun Creator & Identifikasi Konten",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Mengidentifikasi 3 akun creator sukses berdasarkan pilar kontennya\n- Menganalisis elemen visual Hook & CTA yang digunakan\n- Instruktur berkeliling memberikan bimbingan dan umpan balik personal.",
        "sub3_media": "PC Lab / HP & Lembar Jobsheet 01",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan bahwa profesi content creator memerlukan pemahaman etika hak cipta, pemetaan platform, dan struktur visual yang konsisten.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan pertanyaan kuis interaktif mengenai etika hak cipta dan karakteristik platform.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai hasil identifikasi audit akun creator peserta berdasarkan rubrik Jobsheet 01.",
        "evaluasi_media": "Lembar Jobsheet 01 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Mencari 1 contoh konten viral di TikTok/Instagram dan menganalisis struktur Hook-nya."
    },
    {
        "no": 2,
        "judul": "Teknik Fotografi Smartphone untuk Konten Visual",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengoperasikan fitur kamera smartphone, menerapkan komposisi foto (Rule of Thirds, Flatlay), dan pencahayaan secara optimal.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point, Roda Warna, & Contoh Foto Produk Studio",
        "alat_list": "Alat Proyektor dan Layar, Smartphone Kamera, Continuous Light / Softbox, Papan Tulis",
        "bahan_list": "Flashdisk, Objek Produk Latihan, Background Foto (Jobsheet 02)",
        "persiapan": "Persiapan mini studio foto / area pencahayaan lab TIK\nPersiapan alat peraga objek foto produk",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction penggunaan peralatan lab\n4. Apersepsi: Dampak foto produk jernih vs foto buram pada tingkat penjualan online\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Pengaturan Kamera Smartphone & Prinsip Komposisi",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Pengaturan fokus, exposure lock, gridlines, dan resolusi kamera HP\n- Prinsip komposisi: Rule of Thirds, Leading Lines, Framing, & Flatlay (Top-Down)\n- Teknik pencahayaan alami (Natural Light) vs lampu studio buatan\nInstruktur memberikan kesempatan peserta untuk mengajukan pertanyaan.",
        "sub1_media": "Power Point & Smartphone Master Demo",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Shooting Angle & Lighting Setup",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Pengambilan foto angle Eye Level, High Angle, Low Angle, dan Flatlay\n- Pengaturan arah cahaya utama (Key Light) dan peredam bayangan (Reflector)\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Lighting Studio Set & Objek Produk",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Fotografi Produk & Lifestyle HP",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Mengambil minimal 3 variasi foto produk dengan komposisi & angle berbeda\n- Mengatur pencahayaan agar tidak terdapat bayangan keras\n- Instruktur berkeliling memberikan umpan balik personal.",
        "sub3_media": "Smartphone Peserta & Jobsheet 02",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan foto HP berkualitas ditentukan oleh kontrol pencahayaan, pemilihan angle yang tepat, dan kebersihan lensa.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan kuis tebak teknik komposisi dan angle dari sampel gambar.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai 3 karya foto produk HP peserta berdasarkan rubrik penilaian Jobsheet 02.",
        "evaluasi_media": "Lembar Jobsheet 02 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Mengambil 1 foto flatlay produk di rumah memanfaatkan cahaya jendela (Natural Window Light)."
    },
    {
        "no": 3,
        "judul": "Teknik Pengambilan Video dengan Smartphone",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menguasai pergerakan kamera (Camera Movement), stabilisasi video, dan pengambilan variasi shot (A-Roll & B-Roll) menggunakan smartphone.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point, Video Contoh Movement & Shot Range",
        "alat_list": "Alat Proyektor dan Layar, Smartphone, Tripod, Gimbal Stabilizer, Papan Tulis",
        "bahan_list": "Flashdisk, Modul Sinematografi HP, Shotlist Template (Jobsheet 03)",
        "persiapan": "Persiapan area praktik video shooting\nPersiapan alat stabilizer & tripod",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction pergerakan kamera\n4. Apersepsi: Mengapa video Reels/TikTok butuh variasi shot dinamis agar tidak dibaikan\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Dasar Sinematografi HP & Variasi Shot",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Resolusi video (1080p vs 4K) dan Frame Rate (24fps, 30fps, 60fps)\n- Variasi jenis shot: Wide Shot, Medium Shot, Close-Up, dan Detail Shot\n- Perbedaan A-Roll (Subjek Utama) dan B-Roll (Video Pendukung)\nInstruktur membuka sesi tanya jawab.",
        "sub1_media": "Power Point & Video Sample Shot",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Teknik Pergerakan Kamera & Stabilisasi",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Teknik Panning, Tilting, Dolly In/Out, dan Orbit shot secara halus (Handheld)\n- Penggunaan Tripod dan Gimbal Stabilizer untuk meredam goncangan\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Gimbal Stabilizer & Smartphone Master",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Shooting B-Roll Video Pendek",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Mengambil 5 klip B-Roll berdurasi 5 detik dengan variasi shot & movement berbeda\n- Memastikan footage tidak shaking dan fokus terkunci\n- Instruktur berkeliling memberikan masukan.",
        "sub3_media": "Smartphone Peserta & Jobsheet 03",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan variasi shot dan pergerakan kamera yang stabil adalah kunci utama menarik perhatian penonton video pendek.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan pertanyaan seputar penentuan frame rate dan jenis shot sesuai skenario.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai 5 klip B-Roll peserta berdasarkan kejelasan fokus dan kestabilan gerakan.",
        "evaluasi_media": "Lembar Jobsheet 03 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Merekam 3 klip B-Roll kegiatan sehari-hari dengan teknik Panning."
    },
    {
        "no": 4,
        "judul": "Dasar Copywriting & Psikologi Pembaca",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menerapkan formula copywriting (AIDA, PAS), menyusun Hook visual penarik perhatian, dan Call-to-Action (CTA) pada naskah konten.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point, Contoh Caption Viral & Slide Headlines",
        "alat_list": "Alat Proyektor dan Layar, Perangkat Komputer / HP, Papan Tulis",
        "bahan_list": "Flashdisk, Modul Copywriting, Lembar Kerja (Jobsheet 04)",
        "persiapan": "Persiapan slide formula copywriting\nPersiapan sampel naskah caption berbagai brand",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Pentingnya 3 detik pertama (Hook) agar audiens tidak melakukan swipe-away\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Psikologi Pembaca & Formula AIDA / PAS",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Emosional triggers dan pemicu psikologis pembaca di media sosial\n- Formula AIDA: Attention, Interest, Desire, Action\n- Formula PAS: Problem, Agitate, Solution\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Slide Anatomi Teks",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Teknik Merancang Hook & Call-to-Action (CTA)",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- 5 Jenis Hook: Pertanyaan Penasaran, Angka/Listicle, Kontroversi Halus, Solusi Masalah, & Stating Fact\n- Penulisan CTA yang jelas (Save, Share, Click Link in Bio)\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Power Point & Latihan Papan Tulis",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Penulisan Naskah Copywriting",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun 3 variasi Hook headline dan 1 naskah caption produk sesuai formula AIDA\n- Menentukan CTA yang relevan dengan tujuan promosi\n- Instruktur berkeliling memberikan koreksi naskah.",
        "sub3_media": "PC Lab / Lembar Jobsheet 04",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan copywriting yang efektif menggabungkan empati masalah pembaca dengan pesan singkat yang terstruktur.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta peserta mengidentifikasi formula (AIDA/PAS) pada contoh naskah iklan.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai naskah copywriting AIDA buatan peserta sesuai standar rubrik Jobsheet 04.",
        "evaluasi_media": "Lembar Jobsheet 04 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Memperbaiki caption 1 postingan IG pribadi menggunakan formula AIDA."
    },
    {
        "no": 5,
        "judul": "Design Thinking dalam Pembuatan Konten Visual",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menerapkan 5 tahap Design Thinking (Empathize, Define, Ideate, Prototype, Test) untuk menghasilkan ide konten visual solutif.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point, Canvas Design Thinking, & User Persona Chart",
        "alat_list": "Alat Proyektor dan Layar, Papan Tulis, Sticky Notes, Spidol Warna",
        "bahan_list": "Flashdisk, Modul Design Thinking, Lembar Kerja Persona (Jobsheet 05)",
        "persiapan": "Persiapan meja kelompok workshop\nPersiapan perlengkapan sticky notes & papan ideasi",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction & dinamika kelompok\n4. Apersepsi: Mengapa konten yang bagus secara estetik sering gagal karena tidak menjawab kebutuhan audiens\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Konsep Design Thinking & Pemetaan User Persona",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- 5 Tahapan Design Thinking untuk perancangan konten\n- Tahap Empathize: Mengidentifikasi masalah, keluhan, dan harapan audiens sasaran\n- Menyusun User Persona (Demografi, Geografi, Psikografi, Pain Points)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Template Persona",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Tahap Define, Ideate, & Storyboard Konten",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Merumuskan Problem Statement (How Might We...)\n- Teknik Brainstorming ide konten dengan Mind Mapping & Crazy Eights\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Papan Tulis & Sticky Notes",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri / Kelompok: Workshop Ideasi Konten",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun 1 User Persona target pembeli produk\n- Menghasilkan minimal 5 ide konten solutif berbasis masalah persona tersebut\n- Instruktur berkeliling memfasilitasi ideasi.",
        "sub3_media": "Sticky Notes & Lembar Kerja 05",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan bahwa Design Thinking memastikan setiap konten yang dibuat memiliki tujuan yang jelas bagi audiens.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta kelompok mempresentasikan 1 ide konten hasil ideasi cepat.",
        "aplikasi_media": "Presentasi Kelompok",
        "evaluasi": "Instruktur menilai kerapian dan ketepatan User Persona serta keterhubungan ide konten.",
        "evaluasi_media": "Lembar Kerja 05 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Merumuskan 3 pertanyaan wawancara singkat ke calon audiens."
    },
    {
        "no": 6,
        "judul": "Prinsip Desain Grafis: Warna, Tipografi & Layout",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengaplikasikan teori warna, hierarki tipografi, dan prinsip layout (Alignment, Contrast, Proximity) pada sketsa desain grafis.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Gambar dan Video Contoh Layout",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer Papan Tulis, Spidol dan Penghapus",
        "bahan_list": "Flashdisk Buku Sketsa Pensil Software Desain (Canva/Corel/Photoshop)",
        "persiapan": "Persiapan ruangan lab komputer\nPersiapan alat peraga sketsa & layout",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction dan rules\n4. Mengadakan apersepsi review tata letak\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Teori Warna & Tipografi dalam Layout",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Harmoni warna (Monokromatik, Komplementer) dan psikologi warna\n- Klasifikasi font dan hierarki keterbacaan teks\n- Alignment, Margin, Gridlines, dan White Space\nInstruktur menanyakan pemahaman peserta.",
        "sub1_media": "Power Point Gambar dan Video",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Brainstorming & Pembuatan Thumbnail Layout",
        "sub2_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Mengumpulkan dan mengelompokkan konten teks & gambar\n- Menentukan prioritas konten & focal point\n- Teknik membuat thumbnail sketsa kasar beberapa alternatif layout\nInstruktur membuka sesi tanya jawab.",
        "sub2_media": "Power Point Gambar dan Video",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Membuat Sketsa Thumbnail Layout",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Membuat minimal 3 alternatif sketsa thumbnail di buku sketsa\n- Menentukan margin, grid, dan hirarki teks sesuai design brief\n- Memilih 1 sketsa terbaik sebagai acuan\nInstruktur berkeliling memberikan pendampingan.",
        "sub3_media": "Buku Sketsa Pensil Lembar Jobsheet 06",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan konsep sketsa dan prinsip layout yang seimbang adalah fondasi utama sebelum masuk software digital.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan pertanyaan kuis interaktif mengenai penentuan prinsip layout.",
        "aplikasi_media": "Quizizz.com",
        "evaluasi": "Instruktur membagikan lembar evaluasi / menilai sketsa layout peserta.",
        "evaluasi_media": "Lembar Jobsheet 06 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Mengembangkan thumbnail terpilih menjadi prototipe digital 1 layout."
    },
    {
        "no": 7,
        "judul": "Canva Basic, Branding & Identitas Visual",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengoperasikan fitur dasar Canva, mengatur Brand Kit (Logo, Palette Warna, Font), dan menerapkan identitas visual merek secara konsisten.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point & Workspace Canva Live Demo",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / Workstation Lab Papan Tulis",
        "bahan_list": "Flashdisk File Logo Brand Palet Warna Brand Kit Assets (Jobsheet 07)",
        "persiapan": "Persiapan ruangan lab TIK & akun Canva peserta\nPersiapan sampel aset identitas visual",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Konsistensi warna dan logo brand pada Feed Instagram\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Pengenalan Interface Canva & Elemen Identitas Visual",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Navigasi antarmuka Canva: Templates, Elements, Uploads, Text, Layers\n- Elemen Identitas Visual: Logo, Palette Warna Utama, Font Style, & Graphic Elements\n- Menentukan dimensi kanvas yang tepat (1:1, 4:5, 9:16)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Canva Web Interface",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Setup Brand Kit & Template Pertama",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Cara mengonfigurasi Brand Kit di Canva (Upload logo, set HEX warna, tentukan font)\n- Merancang 1 template postingan Feed sederhana berbasis Brand Kit\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Canva App & Screen Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Membuat Brand Kit & Template Desain Canva",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun Brand Kit pribadi/UMKM pada akun Canva masing-masing\n- Mendesain 1 template Feed Instagram 1080x1080px yang konsisten\n- Instruktur berkeliling memberikan masukan perbaikan.",
        "sub3_media": "PC Lab / Canva & Jobsheet 07",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan Brand Kit di Canva mempermudah pembuatan desain yang konsisten dan efisien.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan kuis tebak elemen identitas visual brand terkenal.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai kelengkapan Brand Kit dan kerapian template Canva peserta.",
        "evaluasi_media": "Lembar Jobsheet 07 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Membuat 2 variasi warna latar belakang pada template Canva."
    },
    {
        "no": 8,
        "judul": "Desain Konten Visual Media Sosial & Advanced Canva",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam memanfaatkan fitur advanced Canva (Magic Tools, Background Remover, Seamless Carousel) untuk memproduksi konten visual interaktif.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Demo Fitur Advanced Canva & Example Carousel",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / Workstation Lab Papan Tulis",
        "bahan_list": "Flashdisk Asset Foto Produk Gambar Mentah Modul Advanced Canva",
        "persiapan": "Persiapan ruangan lab TIK & koneksi internet stabil\nPersiapan file gambar latihan olah background",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Daya tarik Seamless Carousel di Instagram yang meningkatkan swipe rate\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Fitur Advanced Canva & Magic Studio",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Menggunakan Background Remover, Magic Eraser, dan Magic Expand\n- Penggunaan Grid, Frames, Transparency, dan Position Alignment\n- Teknik menyambung garis/elemen antar slide (Seamless Effect)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Canva Live Demo",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Merancang Seamless Carousel Instagram",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Pengaturan ukuran kanvas panjang (misal 5400x1080px) dan garis bantu (Rulers & Guides)\n- Menata alur baca dari Cover, Slide Isi, hingga CTA Slide\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Canva App & Screen Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Mendesain 5-Slide Seamless Carousel",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Membuat 1 set Seamless Carousel Instagram (5 slide) untuk materi edukasi/promosi\n- Menggunakan fitur olah gambar Background Remover pada subjek utama\n- Instruktur berkeliling memberikan pendampingan.",
        "sub3_media": "PC Lab / Canva & Jobsheet 08",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan fitur advanced Canva mempercepat pembuatan desain komersial berstandar tinggi.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan kuis identifikasi fitur Canva yang paling tepat untuk masalah desain tertentu.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai Seamless Carousel buatan peserta berdasarkan kesinambungan visual dan keterbacaan.",
        "evaluasi_media": "Lembar Jobsheet 08 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Ekspor carousel dalam format PNG dan uji posting di feed akun pengujian."
    },
    {
        "no": 9,
        "judul": "Prototipe Desain & Narasi Visual (Storytelling)",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam merancang prototipe desain grafis terintegrasi (Feed + Story + Banner) berbasis alur narasi visual (Storytelling) yang persuasif.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Sample Prototipe Kampanye Visual",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / Workstation Lab Papan Tulis",
        "bahan_list": "Flashdisk Asset Foto Campaign Template Prototipe (Jobsheet 09)",
        "persiapan": "Persiapan ruangan lab TIK & file mockup display\nPersiapan panduan rubrik Test Praktik 09",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Menghubungkan cerita emosional dengan bentuk prototipe tampilan desain\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Prinsip Narasi Visual (Visual Storytelling Arc)",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Struktur narasi visual: Hook (Pembuka), Conflict (Masalah), Climax (Solusi), Resolution (CTA)\n- Konsistensi tone warna dan emosi dari slide pertama hingga slide terakhir\n- Pengertian prototipe desain dan penyusunan mockup kampanye\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Slide Storytelling",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Penyusunan Prototipe Kampanye Visual",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Menggabungkan elemen Feed Carousel dan Instagram Story menjadi satu kesatuan prototipe kampanye\n- Melakukan self-check keterbacaan naskah cerita dan daya tarik gambar\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Canva / Figma & Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri & Uji Praktik Prototipe Narasi Visual",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun prototipe narasi visual lengkap (3 slide Feed + 1 Story Promosi)\n- Melakukan uji praktik peer-review presentasi singkat prototipe ke teman sebaya\n- Instruktur berkeliling memberikan masukan perbaikan.",
        "sub3_media": "PC Lab & Jobsheet/Test Praktik 09",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan prototipe desain dengan narasi yang kuat mampu menyampaikan nilai produk secara menyentuh emosi audiens.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta beberapa peserta mempresentasikan narasi cerita desainnya.",
        "aplikasi_media": "Presentasi Singkat",
        "evaluasi": "Instruktur menilai prototipe dan keterhubungan narasi cerita sesuai lembar Test Praktik Hari 09.",
        "evaluasi_media": "Lembar Test Praktik 09 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Merapikan susunan tata letak prototipe berdasarkan masukan peer review."
    },
    {
        "no": 10,
        "judul": "Pengambilan Video Project (Shooting Day)",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam melaksanakan prosedur produksi video lapangan/studio, mengambil klip A-Roll & B-Roll sesuai shotlist, serta mengelola asset secara disiplin.",
        "metode": "Praktikum Fieldwork Simulasi",
        "media_list": "Media Shotlist Board Storyboard Guidance Sheet",
        "alat_list": "Alat Smartphone Tripod Lighting Reflector Clip-on Mic Papan Tulis",
        "bahan_list": "Flashdisk / Storage Card Check-list Peralatan Form Test Praktik 10",
        "persiapan": "Persiapan lokasi shooting (Indoor Studio / Outdoor Area)\nPersiapan pengecekan baterai & ruang penyimpanan HP",
        "pendahuluan_text": "1. Memberi salam dan pengarahan lokasi shooting\n2. Mengabsen dan membagi kelompok/peran produksi\n3. Menjelaskan safety induction dan aturan shooting lapangan\n4. Review shotlist dan target klip video yang harus diambil\n5. Menyampaikan tujuan Test Praktik Hari 10",
        "sub1_judul": "Briefing Pra-Produksi & Pengecekan Alat",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Pengecekan kelengkapan kamera HP, kebersihan lensa, memori, dan audio clip-on\n- Review shotlist: Kepastian jumlah klip A-Roll (Talent) dan B-Roll (Detail Produk)\n- Pembagian waktu dan alur pengambilan gambar di area shooting\nInstruktur membuka sesi tanya jawab teknis.",
        "sub1_media": "Shotlist Sheet & Equipment Checklist",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Eksekusi Production Shooting Lapangan",
        "sub2_uraian": "Instruktur mendampingi dan mengawasi peserta selama:\n- Pengambilan video talent berbicara (A-Roll) dengan audio jernih\n- Pengambilan pergerakan kamera untuk variasi B-Roll lingkungan/produk\n- Pengawasan pencahayaan dan penguncian fokus kamera HP.",
        "sub2_media": "Area Shooting & Studio Set",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Shooting Mandiri & Verifikasi Footage",
        "sub3_uraian": "Instruktur mendampingi peserta untuk:\n- Menyelesaikan seluruh shotlist kebutuhan video project\n- Melakukan pengecekan ulang (playback) kelayakan fokus, pencahayaan, dan suara klip\n- Mengelompokkan file footage ke folder projek.",
        "sub3_media": "Smartphone Peserta & Test Praktik 10",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan kedisiplinan mengecek shotlist di lapangan mencegah kekurangan footage saat proses editing.",
        "kesimpulan_media": "Papan Pengarahan",
        "aplikasi": "Instruktur memeriksa kerapian struktur folder simpanan hasil shooting di HP/Laptop peserta.",
        "aplikasi_media": "File Explorer",
        "evaluasi": "Instruktur menilai kinerja dan kualitas kelayakan hasil rekam video peserta berdasarkan rubrik Test Praktik Hari 10.",
        "evaluasi_media": "Lembar Test Praktik 10",
        "tugas": "Instruktur memberikan tugas tambahan: Meng-copy seluruh file video mentah ke laptop dan membuat cadangan di Google Drive."
    },
    {
        "no": 11,
        "judul": "Editing Video Pendek dengan CapCut (Pengenalan & Basic Editing)",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengoperasikan aplikasi CapCut, melakukan pemotongan video (Trim/Split), serta menambahkan audio, subtitle otomatis, dan transisi dasar.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Demo Interface CapCut & Video Latihan",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis Earphone/Headset",
        "bahan_list": "Flashdisk Footage Video Hari 10 Backsound No-Copyright Modul CapCut",
        "persiapan": "Persiapan instalasi CapCut PC/Mobile di lab\nPersiapan aset musik latar bebas hak cipta",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction penggunaan earphone/headset\n4. Apersepsi: Peran ritme editing dalam menjaga ritme ketertarikan penonton\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Pengenalan Interface CapCut & Import Media",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Navigasi workspace CapCut: Timeline, Tracks, Preview Window, & Tools Bar\n- Mengatur rasio kanvas vertikal (9:16 untuk Reels/Shorts/TikTok)\n- Mengimpor klip A-Roll dan B-Roll hasil shooting Hari 10\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & CapCut Live Demo",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Teknik Pemotongan (Rough Cut) & Pengaturan Audio",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Menggunakan alat Split dan Trim untuk membuang bagian bernada salah / jeda kosong\n- Memasukkan musik latar, mengatur balance volume (Backsound vs Voice Over), dan Fade In/Out\n- Menambahkan Auto Captions / Teks Subtitle otomatis\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "CapCut Software & Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Editing Video Pendek (Rough Cut ke Fine Cut)",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun klip video pendek berdurasi 15-30 detik berbasis footage Hari 10\n- Mengatur ritme potongan gambar sesuai alur cerita dan lagu latar\n- Instruktur berkeliling memberikan bimbingan teknis.",
        "sub3_media": "PC / HP Peserta & Jobsheet 11",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan dasar editing video yang baik berfokus pada kebersihan potongan (clean cut) dan kejelasan audio.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan kuis singkat seputar penggunaan rasio video dan tombol fungsi CapCut.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai hasil draf video pendek buatan peserta berdasarkan ketepatan potongan dan keselarasan audio.",
        "evaluasi_media": "Lembar Jobsheet 11 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Menyelaraskan teks subtitle otomatis agar tidak ada ejaan yang keliru."
    },
    {
        "no": 12,
        "judul": "Audio Enhancement, Motion Graphics & Advanced CapCut",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menerapkan teknik Advanced CapCut (Keyframe Animation, Masking, Chroma Key) serta melakukan perbaikan kualitas audio (Noise Reduction).",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Video Contoh Motion Graphics & Keyframe Animation",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis Earphone/Headset",
        "bahan_list": "Flashdisk Sound Effect (SFX) Collection Overlay Elements Modul Advanced Editing",
        "persiapan": "Persiapan laboratorium komputer & jaringan latihan\nPersiapan pustaka efek suara (SFX) & stiker animasi",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Efek gerakan teks dinamis dan SFX dalam meningkatkan nilai estetika video\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Keyframe Animation & Chroma Key / Masking",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Konsep Keyframe: Membuat pergerakan kustom (Zoom in, Zoom out, Pan) pada gambar/teks\n- Fitur Masking dan Overlay untuk efek reaksi / gambar menumpuk\n- Fitur Chroma Key (Pembersih Latar Belakang Green Screen)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & CapCut Advanced Demo",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Pembersihan Audio (Noise Reduction) & Sound Effects",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Penggunaan fitur Noise Reduction untuk menghilangkan suara bising latar belakang\n- Penempatan Sound Effects (Whoosh, Pop, Bell) pada titik transisi teks/gambar\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "CapCut Software & Headset Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Penerapan Keyframe & SFX pada Video",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menambahkan minimal 3 gerakan Keyframe animasi pada elemen gambar/teks di video\n- Memasukkan SFX penegas pada judul dan membersihkan audio suara talent\n- Instruktur berkeliling melakukan bimbingan.",
        "sub3_media": "PC / HP Peserta & Jobsheet 12",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan animasi keyframe dan SFX yang terukur akan memberi kesan video profesional tanpa berlebihan.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur memberikan kuis identifikasi penggunaan keyframe pada contoh video klip.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai kehalusan gerakan animasi keyframe dan kejernihan audio peserta.",
        "evaluasi_media": "Lembar Jobsheet 12 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Menambahkan 1 efek penegas visual pada kata kunci utama di video."
    },
    {
        "no": 13,
        "judul": "Integrasi Canva-CapCut & Teknik Ekspor Video Optimal",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengintegrasikan aset grafis buatan Canva ke dalam timeline CapCut serta mengekspor video dengan parameter bitrate & resolusi yang optimal.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Tutorial Workflow Canva-CapCut",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis",
        "bahan_list": "Flashdisk Asset Canva (PNG Transparent) Modul Export Setting (Jobsheet 13)",
        "persiapan": "Persiapan file integrasi dua aplikasi\nPersiapan sampel hasil ekspor video tajam vs buram",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Penyebab video yang sudah diedit bagus menjadi pecah saat diunggah ke sosmed\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Workflow Integrasi Canva ke CapCut",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Merancang elemen pendukung di Canva (Lower Third, Frame Nominator, End Screen, Stickers)\n- Ekspor aset Canva dengan latar transparan (PNG Transparent)\n- Mengimpor dan mengatur posisi overlay aset Canva di timeline CapCut\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Interface Canva/CapCut",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Parameter Rendering & Ekspor Video Tajam",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Pemilihan resolusi (1080p vs 4K) dan Frame Rate (30fps vs 60fps)\n- Pengaturan Bitrate (Recommended vs Higher Bitrate) untuk mencegah kompresi berlebih di Instagram/TikTok\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "CapCut Export Settings & Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Integrasi Aset & Ekspor Final Video",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Mendesain Lower Third nama talent di Canva, lalu memasangkannya ke video CapCut\n- Mengekspor video final tanpa watermark dengan resolusi 1080p 30fps\n- Instruktur berkeliling mendampingi proses rendering.",
        "sub3_media": "PC / HP Peserta & Jobsheet 13",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan integrasi Canva-CapCut memperkaya estetika grafis, sedangkan parameter rendering yang tepat menjaga kejernihan visual.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta peserta membandingkan ukuran file dan kualitas gambar dari 2 setting bitrate berbeda.",
        "aplikasi_media": "File Property Comparison",
        "evaluasi": "Instruktur menilai kelancaran integrasi aset Canva dan kejernihan file video final buatan peserta.",
        "evaluasi_media": "Lembar Jobsheet 13 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Mengunggah file video final ke Google Drive portofolio."
    },
    {
        "no": 14,
        "judul": "Project Marketplace, Monetisasi Konten & Link-in-Bio",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam merancang strategi monetisasi konten (Affiliate/Marketplace), merancang halaman Link-in-Bio yang menarik konversi, dan membuat aset visual promosinya.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Slide Marketplace & Monetisasi Examples",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis",
        "bahan_list": "Flashdisk Modul Monetisasi Link-in-Bio Platform Tools (Jobsheet 14)",
        "persiapan": "Persiapan akun platform link-in-bio (Lynk.id/Milkshake/Taplink)\nPersiapan slide presentasi strategi monetisasi",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Peluang penghasilan tambahan creator dari komisi affiliate & produk digital\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Ekosistem Monetisasi Creator & Affiliate Marketing",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Model monetisasi: Affiliate Links, Brand Endorsement, Product Sales, & Digital Download\n- Pentingnya jembatan rujukan (Landing Page Link-in-Bio)\n- Psikologi tata letak tombol penawaran agar menghasilkan klik (CTR tinggi)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Slide Marketplace",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Desain Link-in-Bio & Banner Promosi",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Membangun dan mengedit halaman Link-in-Bio (Menata foto profil, deskripsi, dan tombol tautan)\n- Merancang banner gambar promosi produk untuk dimasukkan ke halaman tautan\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Lynk.id / Canva & Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Membuat Halaman Link-in-Bio & Aset Promosi",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Membuat 1 halaman Link-in-Bio aktif lengkap dengan minimal 3 tombol penawaran\n- Mendesain 1 banner promosi visual produk penarik klik\n- Instruktur berkeliling memberikan evaluasi tampilan.",
        "sub3_media": "PC / HP Peserta & Jobsheet 14",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan halaman Link-in-Bio yang rapi dan menarik merupakan kunci mengonversi penonton konten menjadi pembeli.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur menguji keaktifan seluruh tautan tombol Link-in-Bio peserta.",
        "aplikasi_media": "Browser Test",
        "evaluasi": "Instruktur menilai estetika dan fungsi halaman Link-in-Bio buatan peserta sesuai rubrik Jobsheet 14.",
        "evaluasi_media": "Lembar Jobsheet 14 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Menautkan link-in-bio pada profil bio Instagram/TikTok pribadi."
    },
    {
        "no": 15,
        "judul": "Canva Content Planner, Audiens & Tone of Voice",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menentukan Tone of Voice merek, menyusun matriks kalender konten (Content Calendar), dan mengoperasikan fitur Canva Content Planner.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Slide Audiens & Tone of Voice Matrix",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis",
        "bahan_list": "Flashdisk Template Content Calendar (Excel/Sheets/Canva) Modul Jobsheet 15",
        "persiapan": "Persiapan template kalender konten bulanan\nPersiapan slide matriks pilar konten",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Konsistensi jadwal posting dan karakter suara brand dalam membangun reputasi\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Tone of Voice & Matriks Pilar Konten",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Menentukan karakter persona merek (Formal, Friendly, Energetic, Humorous)\n- 4 Pilar Konten Utama: Educate, Entertain, Inspire, & Promote\n- Menentukan proporsi pembagian pilar konten dalam satu bulan\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Slide Tone of Voice",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Demonstrasi Content Calendar & Canva Planner",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Mengisi jadwal tayang konten pada template Content Calendar bulanan\n- Menggunakan fitur Canva Content Planner untuk mepenjadwalan posting otomatis ke media sosial\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Canva Content Planner & Sheets",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Menyusun Kalender Konten Bulanan",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun skema Kalender Konten untuk 1 minggu penuh (7 postingan) lengkap dengan pilar, format, dan caption\n- Mengatur jadwal draf pada Canva Content Planner\n- Instruktur berkeliling memberikan masukan.",
        "sub3_media": "PC Lab & Jobsheet 15",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan perencanaan kalender konten menjaga konsistensi publikasi tanpa harus membuat konten secara terburu-buru.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta peserta menyebutkan contoh penyesuaian kata-kata sesuai Tone of Voice tertentu.",
        "aplikasi_media": "Quizizz.com / Form",
        "evaluasi": "Instruktur menilai variasi pilar dan kejelasan jadwal pada Kalender Konten peserta.",
        "evaluasi_media": "Lembar Jobsheet 15 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Menyiapkan 2 draf desain gambar untuk jadwal posting hari berikutnya."
    },
    {
        "no": 16,
        "judul": "Storytelling Lanjutan, SEO Copywriting & Email Marketing",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menerapkan teknik SEO Copywriting (Keyword Integration) pada bio/caption sosmed serta menyusun naskah Storytelling dan Email Marketing.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Contoh Caption SEO & Format Email Marketing",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis",
        "bahan_list": "Flashdisk Modul SEO Sosmed & Email Copywriting (Jobsheet 16)",
        "persiapan": "Persiapan laboratorium komputer & koneksi internet riset kata kunci\nPersiapan sampel naskah email penawaran",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Cara kerja mesin pencari Instagram/TikTok (SEO) dalam mendatangkan penonton organik\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "SEO Media Sosial & Optimalisasi Kata Kunci",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Konsep SEO Sosmed: Kata kunci pada Nama Akun, Bio, Caption, dan Alt Text\n- Cara melakukan riset kata kunci pencarian populer di TikTok & Instagram Search Bar\n- Menempatkan kata kunci secara alami dalam naskah tanpa merusak alur baca\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Live Search Bar Demo",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Storytelling Lanjutan & Struktur Email Marketing",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Alur cerita naratif (Hero's Journey) untuk membangun kedekatan emosional dengan audiens\n- Anatomi naskah Email Marketing: Subject Line penarik perhatian, Body Text, dan Button CTA\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Power Point & Sample Email Editor",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Penulisan Caption SEO & Naskah Email",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun 1 caption ber-SEO mengandung 3 kata kunci utama hasil riset\n- Menulis 1 naskah email penawaran singkat berbasis narasi storytelling\n- Instruktur berkeliling melakukan koreksi teks.",
        "sub3_media": "PC Lab & Jobsheet 16",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan SEO sosmed membantu konten ditemukan audiens baru, sedangkan email storytelling memperkuat loyalitas audiens lama.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta peserta menunjukkan kata kunci utama pada naskah buatan rekannya.",
        "aplikasi_media": "Peer Analysis",
        "evaluasi": "Instruktur menilai ketepatan kata kunci SEO dan struktur naskah email peserta sesuai Jobsheet 16.",
        "evaluasi_media": "Lembar Jobsheet 16 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Memperbaiki kata kunci pada deskripsi profil bio akun sosmed."
    },
    {
        "no": 17,
        "judul": "Pemanfaatan AI untuk Copywriting, Storyboard & Personal Branding",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mengoperasikan Generative AI (ChatGPT/Gemini/Canva AI) dengan teknik Prompt Engineering untuk merancang naskah, ideasi storyboard, dan personal branding.",
        "metode": "Ceramah\nDemonstrasi\nPraktikum",
        "media_list": "Media Power Point Lembar Informasi AI Prompt Guide & Live Demo AI Tools",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Papan Tulis",
        "bahan_list": "Flashdisk Lembar Informasi Hari 17 Modul Prompt Engineering (Jobsheet 17)",
        "persiapan": "Persiapan akun ChatGPT / Gemini di komputer lab\nPersiapan slide panduan struktur prompt AI",
        "pendahuluan_text": "1. Memberi salam dan Memperkenalkan diri\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Apersepsi: Menggunakan AI sebagai asisten kreatif mempercepat riset dan ideasi konten 10x lebih cepat\n5. Menyampaikan judul unit dan tujuan pembelajaran",
        "sub1_judul": "Pengenalan Tools AI & Teknik Prompt Engineering",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Peran Generative AI (ChatGPT, Gemini, Claude, Canva Magic Write) dalam industri kreatif\n- Formula Prompting Efektif: Role + Context + Task + Format + Constraint\n- Etika, batasan, dan validasi kebenaran informasi hasil keluaran AI (Cek Fakta)\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Power Point & Lembar Informasi 17",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Ideasi Storyboard & Personal Branding Berbasis AI",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Menggunakan AI untuk menghasilkan 10 ide konten visual dan outline storyboard video pendek\n- Memanfaatkan AI dalam memetakan posisi Personal Branding diri di media sosial\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "ChatGPT / Gemini Web Interface",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Penulisan Prompt AI Copywriting & Storyboard",
        "sub3_uraian": "Instruktur memberikan jobsheet dan mendampingi peserta untuk:\n- Menyusun prompt untuk menghasilkan 3 variasi naskah caption dan ide visual storyboard video\n- Melakukan penyuntingan manusia (Humanizing Text) agar naskah tidak kaku\n- Instruktur berkeliling memberikan pendampingan.",
        "sub3_media": "PC Lab / AI Tools & Jobsheet 17",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan AI adalah alat pembantu yang ampuh, tetapi sentuhan kreativitas dan kurasi manusia tetap menentukan hasil akhir.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur meminta peserta memperlihatkan prompt AI terbaik yang menghasilkan respon paling presisi.",
        "aplikasi_media": "Screen Sharing / Live Show",
        "evaluasi": "Instruktur menilai struktur prompt dan hasil akhir naskah AI peserta sesuai Jobsheet 17.",
        "evaluasi_media": "Lembar Jobsheet 17 & Rubrik",
        "tugas": "Instruktur memberikan tugas tambahan: Menyimpan daftar 5 prompt AI favorit untuk pustaka pribadi."
    },
    {
        "no": 18,
        "judul": "Final Project: Perencanaan & Produksi Konten Visual Terintegrasi",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam menyusun proposal perencanaan kampanye konten terintegrasi (Carousel, Reels Video, & Story) dan memulai eksekusi produksi secara mandiri.",
        "metode": "Workshop Praktikum Bimbingan Mandiri",
        "media_list": "Media Power Point Briefing Final Project & Rubrik Ujian",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Workstation Lab Papan Tulis",
        "bahan_list": "Flashdisk Template Project Plan Check-list Produksi Final Project",
        "persiapan": "Persiapan lembar pengarahan Ujian Akhir / Final Project\nPersiapan tempat bimbingan individu instruktur",
        "pendahuluan_text": "1. Memberi salam dan pengarahan Final Project\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction & aturan Ujian Praktik Akhir\n4. Penjelasan Briefing Final Project: Membuat kampanye visual terintegrasi (Feed + Video + Story)\n5. Menyampaikan kriteria penilaian & tenggat waktu",
        "sub1_judul": "Penyusunan Perencanaan Kampanye & Design Brief",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Menyusun sasaran kampanye, identifikasi target audiens, dan pilar pesan utama\n- Menentukan alur narasi visual dan daftar kebutuhan aset (Foto, Grafis Canva, Video CapCut)\n- Memvalidasi rencana projek dengan bimbingan instruktur\nInstruktur membuka sesi tanya jawab.",
        "sub1_media": "Project Briefing Sheet & Project Plan Template",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Eksekusi Produksi Aset Grafis (Feed Carousel & Story)",
        "sub2_uraian": "Instruktur mendampingi peserta selama:\n- Merancang 1 set Seamless Carousel (5 slide) berbasis Brand Kit di Canva\n- Mendesain 2 aset Instagram Story pendukung promo\n- Instruktur memberikan masukan bimbingan per individu.",
        "sub2_media": "Canva / Workstation Lab",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Eksekusi Produksi Aset Video Pendek (Editing CapCut)",
        "sub3_uraian": "Instruktur mendampingi peserta selama:\n- Mengedit klip video pendek (Reels/Shorts 15-30 detik) mengintegrasikan audio, subtitle, dan animasi\n- Memastikan seluruh aset visual sesuai dengan panduan identitas merek\n- Instruktur memantau kemajuan kerja.",
        "sub3_media": "CapCut / Workstation Lab",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan perencanaan kampanye yang terstruktur mempermudah eksekusi produksi aset visual terintegrasi.",
        "kesimpulan_media": "Papan Informasi",
        "aplikasi": "Instruktur mengimbau peserta mencatat poin progres kerja pada Check-list Final Project.",
        "aplikasi_media": "Progress Check-list",
        "evaluasi": "Instruktur melakukan pemeriksaan draf awal perencanaan dan aset visual kampanye peserta.",
        "evaluasi_media": "Lembar Bimbingan Final Project",
        "tugas": "Instruktur memberikan tugas tambahan: Melanjutkan finishing aset video dan grafis untuk persiapam polishing di pertemua 19."
    },
    {
        "no": 19,
        "judul": "Final Project: Finishing, Polishing & Quality Control",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam melakukan Quality Control (QC) aset visual (typo, warna, resolusi, audio) serta menyusun slide presentasi portofolio kampanye.",
        "metode": "Workshop Praktikum Bimbingan Mandiri",
        "media_list": "Media Power Point Checklist Quality Control & Sample Presentation Slide",
        "alat_list": "Alat Proyektor dan Layar Perangkat Komputer / HP Workstation Lab Papan Tulis",
        "bahan_list": "Flashdisk Template Presentation Portofolio PDF Export Tools",
        "persiapan": "Persiapan lembar periksa Quality Control (QC)\nPersiapan jaringan pengumpulan portofolio",
        "pendahuluan_text": "1. Memberi salam dan pengarahan sesi finishing\n2. Mengabsen peserta pelatihan\n3. Menjelaskan safety induction lab TIK\n4. Penjelasan pentingnya Quality Control sebelum karya dipublikasikan/dipresentasikan\n5. Menyampaikan target penyelesaian portofolio hari ini",
        "sub1_judul": "Prosedur Quality Control (QC) & Polishing Aset",
        "sub1_uraian": "Instruktur menjelaskan kepada peserta terkait:\n- Pengecekan ejaan teks (Typo Check), kontras warna, dan konsistensi font\n- Pengecekan kejernihan ekspor video, keselarasan subtitle, dan kejelasan audio\n- Melakukan perbaikan akhir (Polishing) pada aset yang belum sempurna\nInstruktur memberikan kesempatan peserta bertanya.",
        "sub1_media": "Checklist QC Sheet & Sample Review",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Penyusunan Slide Presentasi Portofolio Kampanye",
        "sub2_uraian": "Instruktur mendemonstrasikan kepada peserta terkait:\n- Menyusun slide presentasi (Latar belakang brand, Target audiens, Hasil karya Feed/Video/Story, & Link-in-Bio)\n- Teknik menyampaikan proses kreatif secara singkat dan percaya diri\n- Tanya jawab interaktif bersama peserta.",
        "sub2_media": "Power Point / Canva Presentation",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Praktik Mandiri: Finishing & Pengunggahan Berkas Portofolio",
        "sub3_uraian": "Instruktur mendampingi peserta untuk:\n- Menyelesaikan seluruh berkas final project dan slide presentasi\n- Mengunggah file karya final (MD/DOCX/PDF/Video) ke repository/folder portofolio\n- Instruktur memverifikasi kelengkapan berkas setiap peserta.",
        "sub3_media": "PC Lab & Google Drive / GitHub",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan Quality Control yang teliti menjamin profesionalisme dan kualitas karya portofolio desainer.",
        "kesimpulan_media": "Power Point",
        "aplikasi": "Instruktur melakukan simulasi singkat urutan pemanggilan presentasi ujian esok hari.",
        "aplikasi_media": "Agenda Sheet",
        "evaluasi": "Instruktur mengecek dan menandatangani lembar kelengkapan berkas kampanye visual peserta.",
        "evaluasi_media": "Lembar Verifikasi Berkas QC",
        "tugas": "Instruktur memberikan tugas tambahan: Berlatih membaca slide presentasi di rumah agar siap tampil percaya diri di pertemuan 20."
    },
    {
        "no": 20,
        "judul": "Presentasi Final Project, Evaluasi Portofolio & Penutupan Pelatihan",
        "unit": "M.74DKV13.012.2",
        "unit_title": "Menciptakan Karya Desain",
        "tujuan": "Setelah selesai mengikuti pelatihan ini peserta kompeten dalam mempresentasi karya kampanye konten visual di depan penguji, menerima umpan balik evaluasi, dan merumuskan rencana aksi karier content creator.",
        "metode": "Presentasi Ujian Evaluasi Ceremonial",
        "media_list": "Media Slide Presentasi Peserta & Rubrik Penilaian Ujian Akhir",
        "alat_list": "Alat Proyektor dan Layar Sound System Pointer Workstation Penguji Papan Tulis",
        "bahan_list": "Flashdisk Lembar Penilaian Penguji Sertifikat Pelatihan Form Evaluasi Program",
        "persiapan": "Persiapan ruang presentasi & sistem audio proyektor\nPersiapan berkas penilaian penguji instruktur",
        "pendahuluan_text": "1. Memberi salam dan pembukaan Ujian Presentasi Akhir\n2. Mengabsen dan mengecek kesiapan peserta\n3. Menjelaskan tata tertib dan durasi waktu presentasi (7 menit presentasi + 3 menit tanya jawab)\n4. Pengarahan rubrik penilaian oleh tim penguji\n5. Memulai sesi ujian presentasi",
        "sub1_judul": "Sesi Presentasi Karya Final Project (Sesi 1)",
        "sub1_uraian": "Pelaksanaan presentasi peserta kelompok 1 di hadapan penguji:\n- Pemaparan konsep kampanye, pertunjukkan aset visual Carousel, Video Reels, dan Link-in-Bio\n- Sesi tanya jawab dan masukan evaluasi dari instruktur penguji\n- Penguji mencatat nilai pada rubrik ujian.",
        "sub1_media": "Slide Presentasi Peserta & Projector",
        "sub1_waktu": "35 menit",
        "sub2_judul": "Sesi Presentasi Karya Final Project (Sesi 2)",
        "sub2_uraian": "Pelaksanaan presentasi peserta kelompok 2 di hadapan penguji:\n- Pemaparan konsep kampanye, pertunjukkan aset visual Carousel, Video Reels, dan Link-in-Bio\n- Sesi tanya jawab dan masukan evaluasi dari instruktur penguji\n- Penguji mencatat nilai pada rubrik ujian.",
        "sub2_media": "Slide Presentasi Peserta & Projector",
        "sub2_waktu": "35 menit",
        "sub3_judul": "Evaluasi Portofolio, Action Plan & Penutupan",
        "sub3_uraian": "Instruktur memimpin sesi akhir:\n- Penyampaian evaluasi umum terhadap seluruh karya portofolio peserta\n- Pengisian Rencana Aksi Karier (Action Plan Content Creator / Desainer Grafis)\n- Penutupan resmi program pelatihan dan foto bersama.",
        "sub3_media": "Form Rencana Aksi & Berkas Penutupan",
        "sub3_waktu": "35 menit",
        "kesimpulan": "Instruktur menyimpulkan seluruh 20 pertemuan telah membekali peserta dengan keterampilan lengkap pembuatan konten visual media sosial.",
        "kesimpulan_media": "Slide Penutupan",
        "aplikasi": "Peserta menyerahkan lembar komitmen Rencana Aksi Karier pasca-pelatihan.",
        "aplikasi_media": "Action Plan Sheet",
        "evaluasi": "Penguji merekapitulasi total nilai kelulusan pelatihan peserta.",
        "evaluasi_media": "Lembar Rekapitulasi Nilai Akhir",
        "tugas": "Instruktur memberikan pesan penutup: Terus konsisten mengunggah karya ke portofolio digital dan media sosial."
    }
]

def make_md_content(m):
    p_text = m['pendahuluan_text'].replace('\n', '<br>')
    s1_text = m['sub1_uraian'].replace('\n', '<br>')
    s2_text = m['sub2_uraian'].replace('\n', '<br>')
    s3_text = m['sub3_uraian'].replace('\n', '<br>')

    out = "# 📚 LESSON PLAN (Rencana Penyajian)\n"
    out += "> **Program Pelatihan:** Pembuatan Konten Visual untuk Media Sosial  \n"
    out += f"> **Kejuruan:** TIK | **Unit Kompetensi:** {m['unit']} - Menciptakan Karya Desain  \n\n"
    out += f"![Status](https://img.shields.io/badge/Status-Draft-orange) ![Pertemuan](https://img.shields.io/badge/Penyajian-Ke--{m['no']}-blue) ![Waktu](https://img.shields.io/badge/Total_Waktu-180_Menit-green)\n\n"
    out += "---\n\n"
    out += "## 📌 Metadata Pelatihan\n\n"
    out += "| Parameter | Detail |\n"
    out += "| :--- | :--- |\n"
    out += f"| **Unit Kompetensi** | {m['unit']} — Menciptakan Karya Desain (72 JP) |\n"
    out += f"| **Penyajian Ke-** | {m['no']} dari 20 |\n"
    out += f"| **Judul Materi** | **{m['judul']}** |\n"
    out += "| **Alokasi Waktu** | 180 Menit |\n"
    out += "| **Instruktur** | Arry Trie Adhytia |\n"
    out += "| **Lokasi / Tanggal** | BPVP Belitung / September 2026 |\n\n"
    out += "---\n\n"
    out += "## 🎯 Tujuan & Manfaat\n\n"
    out += "> [!NOTE]\n"
    out += "> **Tujuan Pembelajaran:**  \n"
    out += f"> {m['tujuan']}\n\n"
    out += "- **Manfaat:** Peserta mampu menguasai keterampilan praktis sesuai standar kualifikasi nasional industri kreatif & kejuruan TIK.\n\n"
    out += "---\n\n"
    out += "## ⏱️ Alur & Pembagian Waktu (180 Menit)\n\n"
    out += "| No | Tahapan | Waktu | Aktivitas Instruktur & Peserta | Alat / Media |\n"
    out += "| :-: | :--- | :-: | :--- | :--- |\n"
    out += f"| **1** | **Pendahuluan** | 10' | {p_text} | Slide PPT, Projector |\n"
    out += f"| **2** | **Penyajian Konsep** | 35' | {s1_text} | {m['sub1_media']} |\n"
    out += f"| **3** | **Demonstrasi Instruktur** | 35' | {s2_text} | {m['sub2_media']} |\n"
    out += f"| **4** | **Praktik Mandiri** | 35' | {s3_text} | {m['sub3_media']} |\n"
    out += f"| **5** | **Kesimpulan** | 5' | {m['kesimpulan']} | {m['kesimpulan_media']} |\n"
    out += f"| **6** | **Aplikasi & Kuis** | 20' | {m['aplikasi']} | {m['aplikasi_media']} |\n"
    out += f"| **7** | **Evaluasi (Jobsheet)** | 30' | {m['evaluasi']} | {m['evaluasi_media']} |\n"
    out += f"| **8** | **Tugas Tambahan & Penutup** | 10' | {m['tugas']}<br>- Kebersihan lab (*5R*) & Doa Penutup | Form Pengumpulan |\n\n"
    out += "---\n\n"
    out += "> [!IMPORTANT]\n"
    out += "> **Catatan Instruktur:**  \n"
    out += "> Pastikan seluruh peserta aktif berpartisipasi dan mematuhi tata tertib keselamatan kerja TIK selama berada di laboratorium komputer.\n"
    return out

def generate_docx(m):
    doc = docx.Document(MASTER_DOCX)
    t = doc.tables[0]

    # Header Row 0 - 2
    t.rows[0].cells[2].text = 'Pembuatan Konten Visual untuk Media Sosial'
    t.rows[0].cells[4].text = 'Kejuruan'
    t.rows[0].cells[6].text = ':'
    t.rows[0].cells[7].text = 'Teknologi Informasi dan Komunikasi (TIK)'

    t.rows[1].cells[2].text = m['unit']
    t.rows[1].cells[4].text = 'Penyajian ke'
    t.rows[1].cells[6].text = ':'
    t.rows[1].cells[7].text = str(m['no'])

    t.rows[2].cells[2].text = m['unit_title']
    t.rows[2].cells[4].text = 'Waktu'
    t.rows[2].cells[6].text = ':'
    t.rows[2].cells[7].text = '180 menit'

    # Judul & Tujuan (Row 3 & 4)
    t.rows[3].cells[0].text = m['judul']
    t.rows[4].cells[0].text = m['tujuan']

    # Metode & Media (Row 5 - 8)
    t.rows[5].cells[0].text = 'Metode Mengajar'
    t.rows[5].cells[3].text = m['metode']
    t.rows[5].cells[9].text = 'Media Pembelajaran\n(Media, Alat, Bahan)'

    t.rows[6].cells[2].text = m['media_list']
    t.rows[7].cells[1].text = m['alat_list']
    t.rows[8].cells[0].text = m['bahan_list']

    # Persiapan (Row 9)
    t.rows[9].cells[0].text = 'Persiapan'
    t.rows[9].cells[8].text = m['persiapan']

    # Pendahuluan (Row 10 & 11)
    t.rows[10].cells[0].text = m['pendahuluan_text']
    t.rows[10].cells[3].text = 'Perkiraan Waktu'
    t.rows[11].cells[2].text = '10 menit'
    t.rows[11].cells[3].text = 'Penyajian'

    # Headers Sub-tabel (Row 12)
    t.rows[12].cells[2].text = 'Sub Judul'
    t.rows[12].cells[5].text = 'Uraian Kegiatan'
    t.rows[12].cells[7].text = 'Media Pembelajaran'

    # Sub Judul 1 (Row 13) - 35 menit
    t.rows[13].cells[0].text = 'Perkiraan Waktu'
    t.rows[13].cells[1].text = m['sub1_judul']
    t.rows[13].cells[4].text = m['sub1_uraian']
    t.rows[13].cells[6].text = m['sub1_media']
    t.rows[13].cells[9].text = m['sub1_waktu']

    # Sub Judul 2 (Row 14) - 35 menit
    t.rows[14].cells[0].text = m['sub2_judul']
    t.rows[14].cells[3].text = m['sub2_uraian']
    t.rows[14].cells[6].text = m['sub2_media']
    t.rows[14].cells[8].text = m['sub2_waktu']
    t.rows[14].cells[9].text = m['sub3_judul']

    # Sub Judul 3 (Row 15) - 35 menit
    t.rows[15].cells[0].text = m['sub3_judul']
    t.rows[15].cells[2].text = m['sub3_uraian']
    t.rows[15].cells[4].text = m['sub3_media']
    t.rows[15].cells[7].text = m['sub3_waktu']
    t.rows[15].cells[8].text = 'Kesimpulan ( Summary)'
    t.rows[15].cells[9].text = 'Kesimpulan ( Summary)'

    # Kesimpulan (Row 16) - 5 menit
    t.rows[16].cells[0].text = 'Kesimpulan ( Summary)'
    t.rows[16].cells[1].text = m['kesimpulan']
    t.rows[16].cells[3].text = m['kesimpulan_media']
    t.rows[16].cells[6].text = '5 menit'
    t.rows[16].cells[7].text = 'Aplikasi'

    # Aplikasi (Row 17) - 20 menit
    t.rows[17].cells[0].text = m['aplikasi']
    t.rows[17].cells[2].text = m['aplikasi_media']
    t.rows[17].cells[5].text = '20 menit'
    t.rows[17].cells[6].text = 'Evaluasi'
    t.rows[17].cells[9].text = m['evaluasi']

    # Evaluasi & Tugas (Row 18) - 30 menit
    t.rows[18].cells[0].text = m['evaluasi']
    t.rows[18].cells[1].text = m['evaluasi_media']
    t.rows[18].cells[4].text = '30 menit'
    t.rows[18].cells[5].text = 'Tugas Tambahan'
    t.rows[18].cells[8].text = m['tugas']

    # Penutup (Row 19 & 20) - 5 menit
    t.rows[19].cells[0].text = m['tugas']
    t.rows[19].cells[3].text = '5 menit'
    t.rows[19].cells[4].text = 'Pemberesan Kelas'
    t.rows[19].cells[7].text = 'Merapikan perangkat PC/HP dan ruang lab (5R)\nBerdoa dan memberikan salam penutup'

    t.rows[20].cells[2].text = '5 menit'

    no_str = f"{m['no']:02d}"
    docx_path = os.path.join(REPO_DIR, f"pertemuan-{no_str}.docx")
    doc.save(docx_path)
    
    md_path = os.path.join(REPO_DIR, f"pertemuan-{no_str}.md")
    with open(md_path, 'w', encoding='utf-8') as f:\n        f.write(make_md_content(m))\n\n    print(f"Generated Meeting {no_str} (DOCX & MD)")

if __name__ == '__main__':
    for m in all_meetings:
        generate_docx(m)
EOF
python3 /root/.openclaw/workspace/lesson-plan-pembuatan-konten-visual/generate_master_format.py
