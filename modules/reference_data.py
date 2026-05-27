"""
Reference data for Coretax BP21, BPMP, BPA1, and BPPU.
All data sourced from the REF sheets of official Coretax Excel templates.
"""

# =============================================================================
# BP21 — Kode Objek Pajak (32 codes)
# Format: (code, name, deemed, tariff_type)
# =============================================================================
BP21_TAX_OBJECTS = [
    ("21-100-35", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Bulanan", 100, "TER"),
    ("21-100-10", "Honorarium atau Imbalan kepada Anggota Dewan Komisaris atau Dewan Pengawas yang Menerima Imbalan secara Tidak Teratur", 100, "TER"),
    ("21-100-27", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Bulanan yang Mendapat Fasilitas di Daerah Tertentu", 100, "TER"),
    ("21-100-37", "Penghasilan yang Diterima atau Diperoleh Pegawai Tetap di Daerah Tertentu yang Tidak Memenuhi Persyaratan Fasilitas", 100, "TER"),
    ("21-100-07", "Imbalan kepada Tenaga Ahli (Pengacara, Akuntan, Arsitek, Dokter, Konsultan, Notaris, PPAT, Penilai, Aktuaris)", 50, "PS17"),
    ("21-100-18", "Imbalan kepada Penasihat, Pengajar, Pelatih, Penceramah, Penyuluh, dan Moderator", 50, "PS17"),
    ("21-100-19", "Imbalan kepada Pengarang, Peneliti, Penerjemah", 50, "PS17"),
    ("21-100-20", "Imbalan kepada Pemberi Jasa dalam Segala Bidang", 50, "PS17"),
    ("21-100-21", "Imbalan kepada Agen Iklan", 50, "PS17"),
    ("21-100-22", "Imbalan kepada Pengawas atau Pengelola Proyek", 50, "PS17"),
    ("21-100-23", "Imbalan kepada Pembawa Pesanan atau yang Menemukan Langganan atau yang Menjadi Perantara", 50, "PS17"),
    ("21-100-06", "Imbalan kepada Petugas Penjaja Barang Dagangan", 50, "PS17"),
    ("21-100-05", "Imbalan kepada Agen Asuransi", 50, "PS17"),
    ("21-100-04", "Imbalan kepada Distributor Perusahaan Pemasaran Berjenjang atau Penjualan Langsung dan Kegiatan Sejenis Lainnya", 50, "PS17"),
    ("21-100-30", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Harian, Mingguan, Satuan dan Borongan dengan Penghasilan Bruto lebih dari Rp2.500.000 Sehari", 50, "PS17"),
    ("21-100-31", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Harian, Mingguan, Satuan dan Borongan dengan Penghasilan Bruto lebih dari Rp2.500.000 Sehari yang Mendapat Fasilitas di Daerah Tertentu", 50, "PS17"),
    ("21-100-12", "Uang Manfaat Pensiun atau Penghasilan Sejenisnya yang diambil sebagian oleh Peserta Program Pensiun yang Masih Berstatus sebagai Pegawai", 100, "PS17"),
    ("21-100-36", "Imbalan kepada Peserta Perlombaan dalam Segala Bidang", 100, "PS17"),
    ("21-100-14", "Imbalan kepada Peserta Rapat, Konferensi, Sidang, Pertemuan, Kunjungan Kerja, Seminar, Lokakarya, atau Pertunjukan, atau Kegiatan Tertentu Lainnya", 100, "PS17"),
    ("21-100-15", "Imbalan kepada Peserta atau Anggota dalam Suatu Kepanitiaan sebagai Penyelenggara Kegiatan Tertentu", 100, "PS17"),
    ("21-100-16", "Imbalan kepada Peserta Pendidikan, Pelatihan, dan Magang", 100, "PS17"),
    ("21-100-17", "Imbalan kepada Peserta Kegiatan Lainnya", 100, "PS17"),
    ("21-100-25", "Penghasilan berupa Uang Pesangon, Uang Manfaat Pensiun, THT, atau JHT yang Terutang atau Dibayarkan pada Tahun Ketiga dan Tahun-Tahun Berikutnya", 100, "PS17"),
    ("21-100-33", "Imbalan kepada Pemain Musik, Pembawa Acara, Penyanyi, Pelawak, Bintang Film/Sinetron/Iklan, Sutradara, Kru Film, Foto Model, Peragawan/Peragawati, Pemain Drama, Penari, Pemahat, Pelukis, Pembuat/Pencipta Konten (Influencer, Selebgram, Blogger, Vlogger), dan Seniman Lainnya", 50, "PS17"),
    ("21-100-34", "Imbalan yang Diterima oleh Olahragawan", 50, "PS17"),
    ("21-100-24", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Harian, Mingguan, Satuan dan Borongan dengan Penghasilan Bruto sampai dengan Rp2.500.000 Sehari", 100, "HARIAN"),
    ("21-100-29", "Upah Pegawai Tidak Tetap yang Dibayarkan secara Harian, Mingguan, Satuan dan Borongan dengan Penghasilan Bruto sampai dengan Rp2.500.000 Sehari yang Mendapat Fasilitas di Daerah Tertentu", 100, "HARIAN"),
    ("21-402-04", "Honor atau Imbalan Lain yang Dibebankan kepada APBN atau APBD yang Diterima oleh PNS Golongan I dan II, Anggota TNI/POLRI Golongan Tamtama dan Bintara, dan Pensiunannya", 100, "0"),
    ("21-402-02", "Honor atau Imbalan Lain yang Dibebankan kepada APBN atau APBD yang Diterima oleh PNS Golongan III, Anggota TNI/POLRI Golongan Perwira Pertama, dan Pensiunannya", 100, "5"),
    ("21-402-03", "Honor atau Imbalan Lain yang Dibebankan kepada APBN atau APBD yang Diterima oleh Pejabat Negara, PNS Golongan IV, Anggota TNI/POLRI Golongan Perwira Menengah dan Tinggi, dan Pensiunannya", 100, "15"),
    ("21-401-01", "Uang Pesangon yang Dibayarkan Sekaligus", 100, "PESANGON"),
    ("21-401-02", "Uang Manfaat Pensiun, Tunjangan Hari Tua, atau Jaminan Hari Tua yang Dibayarkan Sekaligus", 100, "PENSIUN"),
]

# Lookup dict: code -> (name, deemed, tariff_type)
BP21_TAX_OBJECT_LOOKUP = {
    code: (name, deemed, tariff_type)
    for code, name, deemed, tariff_type in BP21_TAX_OBJECTS
}

# =============================================================================
# BPMP — Kode Objek Pajak (3 codes, all TER)
# =============================================================================
BPMP_TAX_OBJECTS = [
    ("21-100-01", "Penghasilan yang diterima oleh Pegawai Tetap termasuk PNS, Anggota TNI, Anggota POLRI, atau Pejabat Negara", 100, "TER"),
    ("21-100-02", "Penghasilan yang diterima oleh Penerima Pensiun secara teratur", 100, "TER"),
    ("21-100-32", "Penghasilan yang diterima oleh Pegawai Tetap yang menerima fasilitas di daerah tertentu", 100, "TER"),
]

BPMP_TAX_OBJECT_LOOKUP = {
    code: (name, deemed, tariff_type)
    for code, name, deemed, tariff_type in BPMP_TAX_OBJECTS
}

# =============================================================================
# BPA1 — Kode Objek Pajak (3 codes, same as BPMP)
# =============================================================================
BPA1_TAX_OBJECTS = [
    ("21-100-01", "Penghasilan yang diterima oleh Pegawai Tetap termasuk Pegawai Negeri Sipil, Anggota Tentara Nasional Indonesia, Anggota Polisi Republik Indonesia atau Pejabat Negara"),
    ("21-100-02", "Penghasilan yang diterima oleh Penerima Pensiun secara teratur"),
    ("21-100-32", "Penghasilan yang diterima oleh Pegawai tetap yang menerima fasilitas di daerah tertentu"),
]

BPA1_TAX_OBJECT_LOOKUP = {code: name for code, name in BPA1_TAX_OBJECTS}

# =============================================================================
# BPPU — Kode Objek Pajak (~200 codes, fixed rates)
# Format: (code, name, rate)
# =============================================================================
BPPU_TAX_OBJECTS = [
    ("24-101-01", "Dividen", 15),
    ("24-104-05", "Jasa Aktuaris", 2),
    ("28-409-25", "Pekerjaan Konstruksi Terintegrasi yang Dilakukan oleh Penyedia Jasa yang Memiliki Sertifikat Badan Usaha", 2.65),
    ("28-409-26", "Pekerjaan Konstruksi Terintegrasi yang Dilakukan oleh Penyedia Jasa yang Tidak Memiliki Sertifikat Badan Usaha", 4),
    ("28-409-27", "Jasa Konsultansi Konstruksi yang Dilakukan oleh Penyedia Jasa yang Memiliki Sertifikat Badan Usaha atau Sertifikat Kompetensi Kerja untuk Usaha Orang Perseorangan", 3.5),
    ("28-409-28", "Jasa Konsultansi Konstruksi yang Dilakukan oleh Penyedia Jasa yang Tidak Memiliki Sertifikat Badan Usaha atau Sertifikat Kompetensi Kerja untuk Usaha Orang Perseorangan", 6),
    ("24-104-06", "Jasa Akuntansi, Pembukuan, dan Atestasi Laporan Keuangan", 2),
    ("28-417-01", "Bunga Simpanan yang Dibayarkan oleh Koperasi kepada Anggota Wajib Pajak Orang Pribadi (bunga sampai dengan Rp240.000,00)", 0),
    ("28-417-02", "Bunga Simpanan yang Dibayarkan oleh Koperasi kepada Anggota Wajib Pajak Orang Pribadi (bunga di atas Rp240.000,00)", 10),
    ("28-419-01", "Dividen yang Diterima/Diperoleh Wajib Pajak Orang Pribadi Dalam Negeri", 10),
    ("28-423-01", "Pemotongan atau pemungutan PPh atas penjualan barang atau penyerahan jasa yang dilakukan oleh Wajib Pajak dengan peredaran bruto tertentu sesuai dengan PP No. 23/2018 atau PP No. 55/2022", 0.5),
    ("28-423-02", "Pemotongan atau pemungutan PPh atas transaksi pembelian yang dilakukan oleh Wajib Pajak dengan peredaran bruto tertentu sesuai dengan PP No. 55/2022", 0.5),
    ("28-410-02", "Imbalan yang Dibayarkan/Terutang kepada Perusahaan Pelayaran Dalam Negeri", 1.2),
    ("28-411-02", "Imbalan Charter Kapal Laut dan/atau Pesawat Udara yang Dibayarkan/Terutang kepada Perusahaan Pelayaran dan/atau Penerbangan Luar Negeri melalui BUT di Indonesia", 2.64),
    ("29-101-01", "Imbalan Charter Pesawat Udara yang Dibayarkan/Terutang kepada Perusahaan Penerbangan Dalam Negeri oleh Pemotong Pajak", 1.8),
    ("28-421-01", "Uplift Hulu Migas", 20),
    ("24-104-07", "Jasa Hukum", 2),
    ("28-421-02", "Participating Interest Eksplorasi Hulu Migas", 5),
    ("28-421-03", "Participating Interest Eksploitasi Hulu Migas", 7),
    ("22-900-01", "Pembayaran atas Pembelian Barang dan/atau Bahan untuk Kegiatan Usahanya oleh BUMN/Badan Usaha Tertentu", 1.5),
    ("22-100-07", "Penjualan Hasil Produksi Kepada Distributor di Dalam Negeri oleh Badan Usaha/Industri Tertentu (Industri Semen)", 0.25),
    ("22-100-08", "Penjualan Hasil Produksi Kepada Distributor di Dalam Negeri oleh Badan Usaha/Industri Tertentu (Industri Baja)", 0.3),
    ("24-104-08", "Jasa Arsitektur", 2),
    ("22-100-09", "Penjualan Hasil Produksi Kepada Distributor di Dalam Negeri oleh Badan Usaha/Industri Tertentu (Industri Otomotif)", 0.45),
    ("22-100-10", "Penjualan Hasil Produksi Kepada Distributor di Dalam Negeri oleh Badan Usaha/Industri Tertentu (Industri Farmasi)", 0.3),
    ("22-100-11", "Penjualan Hasil Produksi Kepada Distributor di Dalam Negeri oleh Badan Usaha/Industri Tertentu (Industri Kertas)", 0.1),
    ("22-405-01", "Pembelian Bahan-bahan Berupa Hasil Kehutanan, Perkebunan, Pertanian, Peternakan, dan Perikanan untuk Keperluan Industri atau Ekspor oleh Badan Usaha Industri atau Eksportir", 0.25),
    ("24-104-09", "Jasa Pertambangan dan Jasa Penunjang Pertambangan Selain Migas", 2),
    ("22-100-12", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Kertas)", 1.5),
    ("22-100-13", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Semen)", 1.5),
    ("22-100-14", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Baja)", 1.5),
    ("22-100-15", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Otomotif)", 1.5),
    ("22-100-16", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Farmasi)", 1.5),
    ("22-100-17", "Pembelian Barang yang Dilakukan oleh DJPB, Bendahara Pemerintah, BUMN/BUMD (Barang Lainnya)", 1.5),
    ("24-104-10", "Jasa Penunjang di Bidang Penerbangan dan Bandar Udara", 2),
    ("22-100-02", "Pembelian Barang yang Dilakukan oleh Bank Devisa dan DJBC", 2.5),
    ("22-100-03", "Penjualan Hasil Produksi Migas oleh Pertamina dan Badan Usaha Selain Pertamina (Bahan Bakar Minyak)", 0.25),
    ("22-100-04", "Penjualan Hasil Produksi Migas oleh Pertamina dan Badan Usaha Selain Pertamina (Bahan Bakar Gas)", 0.3),
    ("22-100-05", "Penjualan Hasil Produksi Migas oleh Pertamina dan Badan Usaha Selain Pertamina (Pelumas)", 0.3),
    ("22-100-06", "Penjualan Hasil Produksi Migas oleh Pertamina dan Badan Usaha Selain Pertamina (Non-Subsidi dan Gas LPG)", 0.3),
    ("24-104-11", "Jasa Penebangan Hutan", 2),
    ("22-402-01", "Penjualan Barang yang Tergolong Sangat Mewah yang Dilakukan oleh Wajib Pajak Badan", 5),
    ("22-100-01", "Impor dengan Angka Pengenal Impor (API)", 2.5),
    ("22-100-18", "Impor Tanpa Angka Pengenal Impor (Non-API)", 7.5),
    ("22-100-19", "Impor yang Tidak Dikuasai", 7.5),
    ("22-100-20", "Impor Kedelai, Gandum dan Tepung Terigu dengan API", 0.5),
    ("24-104-12", "Jasa Pengolahan Limbah", 2),
    ("24-104-13", "Jasa Penyedia Tenaga Kerja (Outsourcing Services)", 2),
    ("24-104-14", "Jasa Perantara dan/atau Keagenan", 2),
    ("24-104-15", "Jasa di Bidang Perdagangan Surat-Surat Berharga, Kecuali yang Dilakukan Oleh Bursa Efek, KSEI dan KPEI", 2),
    ("24-104-16", "Jasa Kustodian/Penyimpanan/Penitipan, Kecuali yang Dilakukan Oleh KSEI", 2),
    ("24-104-17", "Jasa Pengisian Suara (Dubbing) dan/atau Sulih Suara", 2),
    ("24-104-18", "Jasa Mixing Film", 2),
    ("24-104-19", "Jasa Pembuatan Sarana Promosi Film, Iklan, Poster, Photo, Slide, Klise, Banner, Pamphlet, Baliho dan Folder", 2),
    ("24-104-20", "Jasa Sehubungan dengan Software atau Hardware atau Sistem Komputer, Termasuk Perawatan, Pemeliharaan dan Perbaikan", 2),
    ("24-104-21", "Jasa Pembuatan dan/atau Pengelolaan Website", 2),
    ("24-104-22", "Jasa Internet Termasuk Sambungannya", 2),
    ("24-104-23", "Jasa Penyimpanan, Pengolahan dan/atau Penyaluran Data, Informasi, dan/atau Program", 2),
    ("24-104-24", "Jasa Instalasi/Pemasangan Mesin, Peralatan, Listrik, Telepon, Air, Gas, AC dan/atau TV Kabel, Selain yang Dilakukan oleh Wajib Pajak yang Ruang Lingkupnya di Bidang Konstruksi dan Mempunyai Izin dan/atau Sertifikasi Sebagai Pengusaha Konstruksi", 2),
    ("24-104-25", "Jasa Perawatan/Perbaikan/Pemeliharaan Mesin, Peralatan, Listrik, Telepon, Air, Gas, Ac dan/atau Tv Kabel, Selain yang Dilakukan oleh Wajib Pajak yang Ruang Lingkupnya di Bidang Konstruksi dan Mempunyai Izin dan/atau Sertifikasi Sebagai Pengusaha Konstruksi", 2),
    ("24-104-26", "Jasa Perawatan Kendaraan dan/atau Alat Transportasi Darat, Laut dan Udara", 2),
    ("24-104-27", "Jasa Maklon", 2),
    ("24-104-28", "Jasa Penyelidikan dan Keamanan", 2),
    ("24-104-29", "Jasa Penyelenggara Kegiatan Atau Event Organizer", 2),
    ("24-104-30", "Jasa Penyediaan Tempat dan/atau Waktu Dalam Media Massa, Media Luar Ruang Atau Media Lain Untuk Penyampaian Informasi, dan/atau Jasa Periklanan", 2),
    ("24-100-01", "Hadiah, Penghargaan, Bonus dan Lainnya Selain yang Telah Dipotong PPh Pasal 21 Ayat (1) Huruf E UU PPh", 15),
    ("24-104-31", "Jasa Pembasmian Hama", 2),
    ("24-104-32", "Jasa Kebersihan Atau Cleaning Service", 2),
    ("24-104-33", "Jasa Sedot Septic Tank", 2),
    ("24-104-34", "Jasa Pemeliharaan Kolam", 2),
    ("24-104-35", "Jasa Katering Atau Tata Boga", 2),
    ("24-104-36", "Jasa Freight Forwarding", 2),
    ("24-104-37", "Jasa Logistik", 2),
    ("24-104-38", "Jasa Pengurusan Dokumen", 2),
    ("24-104-39", "Jasa Pengepakan", 2),
    ("24-104-40", "Jasa Loading dan Unloading", 2),
    ("24-100-02", "Sewa dan Penghasilan Lain Sehubungan Dengan Penggunaan Harta Kecuali Sewa Tanah dan/atau Bangunan yang Telah Dikenai PPh Pasal 4 Ayat (2) UU PPh.", 2),
    ("24-104-41", "Jasa Laboratorium dan/atau Pengujian Kecuali yang Dilakukan oleh Lembaga atau Institusi Pendidikan Dalam Rangka Penelitian Akademis", 2),
    ("24-104-42", "Jasa Pengelolaan Parkir", 2),
    ("24-104-43", "Jasa Penyondiran Tanah", 2),
    ("24-104-44", "Jasa Penyiapan dan/atau Pengolahan Lahan", 2),
    ("24-104-45", "Jasa Pembibitan dan/atau Penanaman Bibit", 2),
    ("24-104-46", "Jasa Pemeliharaan Tanaman", 2),
    ("24-104-47", "Jasa Pemanenan", 2),
    ("24-104-48", "Jasa Pengolahan Hasil Pertanian, Perkebunan, Perikanan, Peternakan dan/atau Perhutanan", 2),
    ("24-104-49", "Jasa Dekorasi", 2),
    ("24-104-50", "Jasa Pencetakan/Penerbitan", 2),
    ("24-104-01", "Jasa Teknik", 2),
    ("24-104-51", "Jasa Penerjemahan", 2),
    ("24-104-52", "Jasa Pengangkutan/Ekspedisi Kecuali Yang Telah Diatur Dalam Pasal 15 Undang-Undang Pajak Penghasilan", 2),
    ("24-104-53", "Jasa Pelayanan Pelabuhan", 2),
    ("24-104-54", "Jasa Pengangkutan Melalui Jalur Pipa", 2),
    ("24-104-55", "Jasa Pengelolaan Penitipan Anak", 2),
    ("24-104-56", "Jasa Pelatihan dan/atau Kursus", 2),
    ("24-104-57", "Jasa Pengiriman dan Pengisian Uang Ke Atm", 2),
    ("24-104-58", "Jasa Sertifikasi", 2),
    ("24-104-59", "Jasa Survey", 2),
    ("24-104-60", "Jasa Tester", 2),
    ("24-104-02", "Jasa Manajemen", 2),
    ("24-104-61", "Jasa Selain Jasa-Jasa Tersebut di Atas yang Pembayarannya Dibebankan pada APBN atau APBD", 2),
    ("24-104-62", "Jasa Penyelenggaraan Layanan Transaksi Pembayaran Terkait dengan Distribusi Token Oleh Penyelenggara Distribusi", 2),
    ("24-104-63", "Jasa Pemasaran dengan Media Voucer Oleh Penyelenggara Voucer", 2),
    ("24-104-64", "Jasa Penyelenggaraan Layanan Transaksi Pembayaran Terkait dengan Distribusi Voucer Oleh Penyelenggara Voucer dan Penyelenggara Distribusi", 2),
    ("24-104-65", "Jasa Penyelenggaraan Program Loyalitas dan Penghargaan Pelanggan (Consumer Loyalty/Reward Program) Oleh Penyelenggara Voucer", 2),
    ("24-105-01", "Bunga Pinjaman yang Diterima Wajib Pajak Dalam Negeri dan Bentuk Usaha Tetap Melalui Layanan Pinjam Meminjam Uang Berbasis Teknologi Informasi", 15),
    ("28-404-01", "Bunga Tabungan dan Bunga Deposito yang Ditempatkan di Dalam Negeri yang Dananya Bersumber Selain dari DHE", 20),
    ("28-404-02", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang IDR bersumber dari DHE tenor 1 bulan)", 7.5),
    ("28-404-03", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang IDR bersumber dari DHE tenor 3 bulan)", 5),
    ("28-404-04", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang IDR bersumber dari DHE tenor 6 bulan atau lebih)", 0),
    ("24-104-03", "Jasa Konsultan", 2),
    ("28-404-05", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang USD bersumber dari DHE tenor 1 bulan)", 10),
    ("28-404-06", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang USD bersumber dari DHE tenor 3 bulan)", 7.5),
    ("28-404-07", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang USD bersumber dari DHE tenor 6 bulan)", 2.5),
    ("28-404-08", "Bunga Deposito yang Ditempatkan di Dalam Negeri (mata uang USD bersumber dari DHE tenor lebih 6 bulan)", 0),
    ("28-404-09", "Bunga Deposito/Tabungan yang Ditempatkan di Luar Negeri Melalui Bank yang Didirikan atau Bertempat Kedudukan di Indonesia atau Cabang Bank Luar Negeri di Indonesia", 20),
    ("28-404-10", "Diskonto Sertifikat Bank Indonesia", 20),
    ("28-404-11", "Jasa Giro", 20),
    ("28-401-01", "Bunga Obligasi, Surat Utang Negara, atau Obligasi Daerah yang Diterima Wajib Pajak Dalam Negeri dan Bentuk Usaha Tetap", 15),
    ("28-401-06", "Bunga Obligasi yang Diterima Wajib Pajak Dalam Negeri dan Bentuk Usaha Tetap", 10),
    ("28-401-03", "Bunga Obligasi yang Diterima Wajib Pajak Dalam Negeri dan Bentuk Usaha Tetap yang diadministrasikan oleh BI", 10),
    ("24-104-04", "Jasa Penilai (Appraisal)", 2),
    ("28-401-04", "Diskonto Surat Perbendaharaan Negara yang Diterima Wajib Pajak Dalam Negeri dan Bentuk Usaha Tetap", 20),
    ("28-401-05", "Diskonto Surat Perbendaharaan Negara yang Diterima Wajib Pajak Penduduk/Berkedudukan di Luar Negeri", 20),
    ("28-407-01", "Transaksi Penjualan Saham di Bursa Efek (Saham Pendiri)", 0.5),
    ("28-406-01", "Transaksi Penjualan Saham di Bursa Efek (Bukan Saham Pendiri)", 0.1),
    ("28-408-01", "Transaksi Penjualan Saham Milik Perusahaan Modal Ventura Tidak di Bursa Efek", 0.1),
    ("28-403-02", "Persewaan Tanah dan/atau Bangunan", 10),
    ("28-405-01", "Hadiah Undian", 25),
    ("28-409-22", "Pekerjaan Konstruksi yang Dilakukan oleh Penyedia Jasa yang Memiliki Sertifikat Badan Usaha Kualifikasi Kecil atau Sertifikat Kompetensi Kerja untuk Usaha Orang Perseorangan", 1.75),
    ("28-409-23", "Pekerjaan Konstruksi yang Dilakukan oleh Penyedia Jasa yang Tidak Memiliki Sertifikat Badan Usaha Atau Sertifikat Kompetensi Kerja untuk Usaha Orang Perseorangan", 4),
    ("28-409-24", "Pekerjaan Konstruksi yang Dilakukan oleh Penyedia Jasa yang Memiliki Sertifikat Selain Sertifikat Badan Usaha Kualifikasi Kecil atau Sertifikat Kompetensi Kerja untuk Usaha Orang Perseorangan", 2.65),
    ("24-104-29", "Jasa Perawatan/Perbaikan/Pemeliharaan Mesin, Peralatan, Listrik, Telepon, Air, Gas, Ac dan/atau Tv Kabel, Selain Yang Dilakukan Oleh Wajib Pajak Yang Ruang Lingkupnya di Bidang Konstruksi", 2),
    ("24-104-30", "Jasa Perawatan Kendaraan dan/atau Alat Transportasi Darat, Laut dan Udara", 2),
    ("24-104-31", "Jasa Maklon", 2),
    ("24-104-32", "Jasa Penyelidikan dan Keamanan", 2),
    ("24-104-33", "Jasa Penyelenggara Kegiatan Atau Event Organizer", 2),
    ("24-104-34", "Jasa Penyediaan Tempat dan/atau Waktu Dalam Media Massa, Media Luar Ruang Atau Media Lain Untuk Penyampaian Informasi, dan/atau Jasa Periklanan", 2),
]

# Lookup dict: code -> (name, rate)
BPPU_TAX_OBJECT_LOOKUP = {
    code: (name, rate) for code, name, rate in BPPU_TAX_OBJECTS
}

# =============================================================================
# Fasilitas Perpajakan
# =============================================================================
BP21_FACILITIES = {
    "N/A": "Tanpa Fasilitas",
    "TaxExAr21": "Surat Keterangan Bebas (SKB) Pemotongan PPh Pasal 21",
    "DTP": "PPh Ditanggung Pemerintah (DTP)",
    "ETC": "Fasilitas Lainnya",
}

BPMP_FACILITIES = {
    "N/A": "Tanpa Fasilitas",
    "DTP": "PPh Ditanggung Pemerintah (DTP)",
    "ETC": "Fasilitas Lainnya",
}

BPA1_FACILITIES = {
    "N/A": "Tanpa Fasilitas",
    "DTP": "PPh Ditanggung Pemerintah (DTP)",
    "ETC": "Fasilitas Lainnya",
}

BPPU_FACILITIES = {
    "N/A": "Tanpa Fasilitas",
    "TaxExAr22": "SKB Pemotongan PPh Pasal 22",
    "TaxExAr23": "SKB Pemotongan PPh Pasal 23",
    "TaxExIntDep": "SKB Pemotongan PPh atas Bunga atas Deposito Berjangka dan Tabungan",
    "TaxExIntPhtb": "SKB Pemotongan PPh atas Pengalihan Hak atas Tanah dan Bangunan",
    "DTP": "PPh Ditanggung Pemerintah (DTP)",
    "PP23": "Surat Keterangan PP 23/2018",
    "ETC": "Fasilitas Lainnya",
}

# =============================================================================
# Dokumen Referensi (BP21 & BPPU)
# =============================================================================
DOCUMENTS = {
    "Announcement": "Pengumuman",
    "CommercialInvoice": "Surat Tagihan",
    "Contract": "Kontrak",
    "CurrentAccount": "Jasa Giro",
    "Decree": "Decree",
    "DeedOfEngagement": "Akta Perjanjian",
    "DeedOfGeneral": "Akta RUPS",
    "Other": "Lainnya",
    "OtherFacilityDoc": "Dokumen Fasilitas Lainnya",
    "PaymentProof": "Bukti Pembayaran",
    "StatementLetter": "Surat Pernyataan",
    "TaxInvoice": "Faktur Pajak",
    "TaxRegulationDoc": "Dokumen Perpajakan",
    "TradeConfirmation": "Trade Confirmation",
}

# =============================================================================
# Status PTKP
# =============================================================================
PTKP_STATUS_LIST = [
    "TK/0", "TK/1", "TK/2", "TK/3",
    "K/0", "K/1", "K/2", "K/3",
    "HB/0", "HB/1", "HB/2", "HB/3",
]

# PTKP categories for TER rate lookup
PTKP_CATEGORY_A = {"TK/0", "TK/1", "K/0", "HB/0", "HB/1"}
PTKP_CATEGORY_B = {"TK/2", "TK/3", "K/1", "K/2", "HB/2", "HB/3"}
PTKP_CATEGORY_C = {"K/3"}

# =============================================================================
# Status Pegawai (BPMP & BPA1)
# =============================================================================
COUNTERPART_OPTIONS = ["Resident", "Foreign"]

# =============================================================================
# BPA1-specific options
# =============================================================================
STATUS_OF_WITHHOLDING = ["FullYear", "PartialYear", "Annualized"]
WORK_FOR_SECOND_EMPLOYER = ["Yes", "No"]
GROSS_UP_OPTIONS = ["Yes", "No"]

# =============================================================================
# BPPU-specific: Opsi Pembayaran (Instansi Pemerintah)
# =============================================================================
GOV_TREASURER_OPTIONS = ["N/A", "Imprest", "Direct"]

# =============================================================================
# Column definitions for data editor
# =============================================================================
BP21_COLUMNS = [
    ("Masa Pajak", "TaxPeriodMonth", "int"),
    ("Tahun Pajak", "TaxPeriodYear", "int"),
    ("NPWP/NIK", "CounterpartTin", "str"),
    ("ID TKU Penerima", "IDPlaceOfBusinessActivityOfIncomeRecipient", "str"),
    ("Status PTKP", "StatusTaxExemption", "select"),
    ("Fasilitas", "TaxCertificate", "select"),
    ("Kode Objek Pajak", "TaxObjectCode", "select"),
    ("Penghasilan", "Gross", "float"),
    ("Deemed", "Deemed", "float"),
    ("Tarif", "Rate", "float"),
    ("Jenis Dok. Referensi", "Document", "select"),
    ("Nomor Dok. Referensi", "DocumentNumber", "str"),
    ("Tanggal Dok. Referensi", "DocumentDate", "date"),
    ("ID TKU Pemotong", "IDPlaceOfBusinessActivity", "str"),
    ("Tanggal Pemotongan", "WithholdingDate", "date"),
]

BPMP_COLUMNS = [
    ("Masa Pajak", "TaxPeriodMonth", "int"),
    ("Tahun Pajak", "TaxPeriodYear", "int"),
    ("Status Pegawai", "CounterpartOpt", "select"),
    ("NPWP/NIK/TIN", "CounterpartTin", "str"),
    ("Nomor Passport", "CounterpartPassport", "str"),
    ("Status PTKP", "StatusTaxExemption", "select"),
    ("Posisi", "Position", "str"),
    ("Fasilitas", "TaxCertificate", "select"),
    ("Kode Objek Pajak", "TaxObjectCode", "select"),
    ("Penghasilan Kotor", "Gross", "float"),
    ("Tarif", "Rate", "float"),
    ("ID TKU", "IDPlaceOfBusinessActivity", "str"),
    ("Tanggal Pemotongan", "WithholdingDate", "date"),
]

BPA1_COLUMNS = [
    ("Pemberi Kerja Selanjutnya", "WorkForSecondEmployer", "select"),
    ("Masa Pajak Awal", "TaxPeriodMonthStart", "int"),
    ("Masa Pajak Akhir", "TaxPeriodMonthEnd", "int"),
    ("Tahun Pajak", "TaxPeriodYear", "int"),
    ("WNI/WNA", "CounterpartOpt", "select"),
    ("No. Paspor", "CounterpartPassport", "str"),
    ("NPWP", "CounterpartTin", "str"),
    ("Status PTKP", "TaxExemptOpt", "select"),
    ("Posisi", "CounterpartPosition", "str"),
    ("Kode Objek Pajak", "TaxObjectCode", "select"),
    ("Status Bukti Potong", "StatusOfWithholding", "select"),
    ("Jumlah Bulan Bekerja", "NumberOfMonths", "int"),
    ("Gaji", "SalaryPensionJhtTht", "float"),
    ("Opsi Gross Up", "GrossUpOpt", "select"),
    ("Tunjangan PPh", "IncomeTaxBenefit", "float"),
    ("Tunjangan Lainnya / Lembur", "OtherBenefit", "float"),
    ("Honorarium", "Honorarium", "float"),
    ("Asuransi", "InsurancePaidByEmployer", "float"),
    ("Natura", "Natura", "float"),
    ("Tantiem/Bonus/THR", "TantiemBonusThr", "float"),
    ("Iuran Pensiun/THT/JHT", "PensionContributionJhtThtFee", "float"),
    ("Zakat", "Zakat", "float"),
    ("No. BP Sebelumnya", "PrevWhTaxSlip", "str"),
    ("Fasilitas", "TaxCertificate", "select"),
    ("PPh Pasal 21", "Article21IncomeTax", "float"),
    ("ID TKU Pemotong", "IDPlaceOfBusinessActivity", "str"),
    ("Tanggal Pemotongan", "WithholdingDate", "date"),
]

BPPU_COLUMNS = [
    ("Masa Pajak", "TaxPeriodMonth", "int"),
    ("Tahun Pajak", "TaxPeriodYear", "int"),
    ("NPWP", "CounterpartTin", "str"),
    ("ID TKU Penerima", "IDPlaceOfBusinessActivityOfIncomeRecipient", "str"),
    ("Fasilitas", "TaxCertificate", "select"),
    ("Kode Objek Pajak", "TaxObjectCode", "select"),
    ("DPP", "TaxBase", "float"),
    ("Tarif", "Rate", "float"),
    ("Jenis Dok. Referensi", "Document", "select"),
    ("Nomor Dok. Referensi", "DocumentNumber", "str"),
    ("Tanggal Dok. Referensi", "DocumentDate", "date"),
    ("ID TKU Pemotong", "IDPlaceOfBusinessActivity", "str"),
    ("Opsi Pembayaran (IP)", "GovTreasurerOpt", "select"),
    ("Nomor SP2D (IP)", "SP2DNumber", "str"),
    ("Tanggal Pemotongan", "WithholdingDate", "date"),
]
