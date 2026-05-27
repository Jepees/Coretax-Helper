# Coretax Helper

**Coretax Helper** adalah aplikasi web berbasis [Streamlit](https://streamlit.io/) yang dirancang khusus untuk mempermudah konversi data dari *template* Excel perpajakan menjadi format XML yang diwajibkan oleh sistem **Coretax** Direktorat Jenderal Pajak (DJP).

Aplikasi ini diciptakan sebagai alternatif *cross-platform* yang dapat dijalankan di Windows, macOS, maupun Linux—menjadi solusi bagi pengguna yang tidak dapat menggunakan fitur Macro/VBA bawaan dari file Excel DJP.

## ✨ Fitur Utama
- **Mendukung 4 Tipe Dokumen Utama**: BP21, BPMP, BPA1, dan BPPU.
- **Validasi Otomatis**: Memeriksa kelengkapan kolom dan referensi kode objek pajak sebelum konversi.
- **UI/UX Modern & Intuitif**: Antarmuka *drag-and-drop* yang mudah digunakan.
- **Privasi Terjaga**: Pemrosesan *file* dilakukan langsung di memori, tanpa ada *file* sensitif yang disimpan secara permanen di server.

## 🚀 Cara Penggunaan
1. **Upload File**: *Drag-and-drop* file Excel (format `.xlsx`) yang sudah diisi sesuai *template* resmi ke *card* yang sesuai.
2. **Review Data**: Periksa ringkasan jumlah baris, validasi NPWP, dan pastikan tidak ada *error*.
3. **Download XML**: Klik tombol **Download XML**, dan file siap untuk diunggah (di-import) ke dalam menu eBupot di aplikasi Coretax DJP.

## 🛠️ Instalasi Lokal (Developer)

Jika Anda ingin menjalankan aplikasi ini di komputer lokal Anda:

1. **Clone repositori**
   ```bash
   git clone https://github.com/Jepees/Coretax-Helper.git
   cd Coretax-Helper
   ```

2. **Buat Virtual Environment** (Sangat disarankan)
   ```bash
   python -m venv venv
   # Pengguna Windows:
   venv\Scripts\activate
   # Pengguna Mac/Linux:
   source venv/bin/activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Jalankan Aplikasi**
   ```bash
   python -m streamlit run app.py
   ```
   Aplikasi akan terbuka otomatis di browser pada alamat `http://localhost:8501`.

## ⚠️ Disclaimer
Aplikasi ini dikembangkan oleh pihak ketiga dan **tidak terafiliasi dengan Direktorat Jenderal Pajak (DJP)** Kementerian Keuangan RI. Segala bentuk kesalahan data atau kerugian yang ditimbulkan akibat penggunaan aplikasi ini menjadi tanggung jawab pengguna. Selalu pastikan kebenaran data pada file XML sebelum melakukan *upload* ke sistem resmi.

---
*Made with ❤️ by Jeepes for JP*
