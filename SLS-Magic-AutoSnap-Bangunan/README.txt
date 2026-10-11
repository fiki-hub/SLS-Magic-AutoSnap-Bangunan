SLS MAGIC AUTOSNAP BANGUNAN v2.5.1 — DESA / SLS FILTER + TOTAL TITIK

PEMBARUAN
- Dropdown kiri hanya menampilkan KODE DESA (kode hasil pencocokan level_6_full_code dengan referensi desa).
- Dropdown kanan hanya menampilkan nama SLS asli dari level_6_name, dibatasi oleh kode desa terpilih.
- Peta QGIS menyembunyikan fitur titik lain pada layer sumber saat workflow aktif; filter asli dipulihkan ketika dialog ditutup.
- Tombol TOTAL TITIK SLS AKTIF di panel kanan memperlihatkan jumlah fitur yang sesuai dengan pilihan kode desa dan SLS.
- Tabel Manual dan Otomatis memuat kolom FID, no_bang, Kode desa, Nama SLS, flag, status/saran, dan koordinat/skor.
- Kolom flag workflow bernilai 1 untuk fitur yang belum dipetakan, lalu 0 setelah target ditetapkan. Layer sumber tidak diedit.
- Mempertahankan workflow Manual/Otomatis, Review, Load Style, dan pembuatan layer hasil baru.

INSTALASI
1. Tutup QGIS.
2. Hapus folder lama SLS_Magic_AutoSnap_Bangunan dari profil QGIS/python/plugins.
3. Ekstrak folder SLS_Magic_AutoSnap_Bangunan dari ZIP ke python/plugins/.
4. Pastikan folder tidak bertingkat dua.
5. Aktifkan plugin, pilih layer titik, pilih kode desa, lalu pilih nama SLS.

CATATAN
- Pemetaan kode desa menggunakan level_6_full_code dan referensi kode desa yang disediakan.
- Filter peta memerlukan provider yang mendukung subset string (GeoPackage mendukungnya). Jika provider lain tidak mendukung, tabel tetap terfilter tetapi fitur di kanvas mungkin tidak bisa disembunyikan satu per satu.
- Pengujian langsung pada instalasi QGIS pengguna belum dilakukan.


PERBAIKAN v2.5.1
- Output hasil AutoSnap kini meminta lokasi penyimpanan GeoPackage (.gpkg) dan memuat kembali layer permanen ke proyek; tidak lagi mengandalkan layer memory sebagai output akhir.
- Menu utama SLS MAGIC menambahkan Audit Laporan Titik. Pilih layer acuan dan group/layer pekerjaan untuk melihat total titik.
- Ekspor audit saat ini berupa CSV berformat Excel-compatible, bukan workbook XLSX native.
- Layer sumber tidak diedit. Uji runtime QGIS langsung belum dapat dilakukan di lingkungan ini.


Versi 2.5.3: struktur menu dirapikan menjadi Auto Geotagging (Manual/Otomatis), Toolbox (Style Layer dan perintah Load Style yang sudah ada), serta Audit Laporan Titik. Logika pemrosesan inti dipertahankan.

PEMBARUAN v2.5.4
- Ekspor Audit Laporan Titik kini menghasilkan workbook Excel .xlsx asli (tanpa dependensi eksternal), bukan CSV.
- Versi plugin dan judul jendela diperbarui ke 2.5.4.
- Menu Auto Geotagging (Manual/Otomatis), Toolbox (Style Layer/Load Style), dan Audit Laporan Titik tetap dipertahankan.
- Validasi sintaks Python berhasil. Pengujian runtime di QGIS belum dilakukan di lingkungan ini.


PEMBARUAN v2.5.8 — STRUKTUR UI
- Menu perintah diratakan ke menu utama SLS MAGIC agar mudah diakses tanpa submenu bertumpuk.
- Urutan: buka AutoSnap, Auto Geotagging Manual, Auto Geotagging Otomatis, Toolbox — Style Layer, Load Style — Titik Bangunan, Audit Laporan Titik.
- Panel layer titik, kode desa, dan SLS di dialog utama tetap dipertahankan; progress bar proses yang sudah ada juga tetap digunakan.
- Logika pemrosesan dan fungsi perintah lama tidak diubah.
- Validasi sintaks Python berhasil; pengujian runtime QGIS langsung belum dilakukan.


Versi 2.5.8: Saat SLS aktif dipilih, plugin otomatis mencentang raster georeference yang nama layernya cocok dengan level_6_full_code pada SLS tersebut dan menonaktifkan raster georeference lainnya. Saat berpindah SLS, pilihan raster ikut berpindah. Visibilitas awal dipulihkan saat dialog ditutup.


Versi 2.5.8: memperbaiki pencocokan raster georeference melalui nama layer dan nama file sumber; layer SLS aktif dicentang, raster georeference SLS lain dilepas centangnya, serta grup induk diaktifkan bila diperlukan agar raster terlihat.


Versi 2.6.0: Sinkronisasi raster georeference kini mencocokkan nama SLS aktif saja. Kode desa dan level_6_full_code tidak dipakai untuk pencocokan; akhiran _WSS dan ekstensi raster diabaikan.


PEMBARUAN v2.6.6 — PERUBAHAN UI NYATA
- Dialog utama kini memiliki command bar dengan popup Auto Geotagging (Manual/Otomatis), Toolbox, dan Audit Laporan Titik.
- Tab Manual/Otomatis disembunyikan dari tampilan dan dikendalikan melalui popup; widget dan logika kerja tab tetap dipakai.
- Tombol Style Layer dipindahkan dari panel progres ke popup Toolbox.
- Menu SLS MAGIC dipasang langsung pada menu bar QGIS.
- Fungsi pemrosesan yang sudah ada dipertahankan; sintaks Python diperiksa, tetapi uji runtime QGIS tetap diperlukan.


PATCH v2.6.6 - RASTER GEOREFERENCE SELECTION
- Raster SLS selector now lists rasters from the Georeference layer-tree group even when raster names lack WSS/georeference keywords.
- Automatic matching uses level_6_fullcode and normalized raster names, including *_WSS.
- Manual Raster SLS selection mirrors into the Auto Bangunan Raster Referensi selector.
- Improved spacing and selector widths for easier use.


Versi 2.6.8: tombol Choose untuk kedua pemilih raster Georeference, aktivasi layer terpilih di QGIS, dan acuan audit dikunci ke layer Kecamatan/Geotaging Kecamatan.


Versi 2.6.8: tombol Choose untuk kedua pemilih raster Georeference, aktivasi layer terpilih di QGIS, dan acuan audit dikunci ke layer Kecamatan/Geotaging Kecamatan.


PEMBARUAN v2.6.9 — PILIH RASTER GEOREFERENCE MANUAL
- Pilihan raster yang ditentukan pengguna tidak lagi ditimpa otomatis oleh pencocokan nama SLS.
- Pilih raster pada dropdown Raster SLS atau tekan Choose, lalu pilih raster dari daftar layer di grup Georeference.
- Raster terpilih akan dicentang; raster lain yang berada di grup Georeference akan dilepas centangnya.
- Layer aktif QGIS diarahkan ke raster terpilih.
- Visibilitas raster awal tetap dipulihkan saat dialog plugin ditutup.
- Pemeriksaan sintaks Python dilakukan; pengujian langsung di QGIS pengguna belum dilakukan.


PEMBARUAN v2.7.0 — AUDIT LAPORAN TITIK
- Sumber titik acuan BPS Pusat diprioritaskan ke layer Kecamatan, termasuk layer bernama KdKec_...
- Pilihan layer acuan tidak dikunci; pengguna dapat mengganti layer melalui dropdown.
- Semua layer titik valid di project tersedia sebagai pilihan jika diperlukan.
- Tombol Hitung ulang, Ekspor audit Excel, dan Tutup diberi warna yang lebih tegas dan kontras.
- Pemeriksaan sintaks Python berhasil; belum diuji langsung pada instalasi QGIS pengguna.


PEMBARUAN v2.7.1 — ACUAN TITIK PUSAT BPS TETAP
- File KdKec_3210200.gpkg dibundel bersama plugin.
- Audit Laporan Titik otomatis memuat layer "KdKec_3210200 (Acuan Titik Pusat BPS)" bila belum tersedia di proyek.
- Layer acuan bawaan menjadi pilihan awal setiap kali dialog audit dibuka, bukan mengikuti layer aktif yang berubah-ubah.
- Dropdown acuan tetap aktif sehingga pengguna tidak terkunci dan masih dapat memilih layer titik lain secara manual.
- Tombol Hitung ulang dan Ekspor audit Excel memakai warna kontras.
- Sintaks Python diperiksa; pengujian runtime di QGIS belum dilakukan.


PEMBARUAN v2.7.2 — PERBAIKAN ERROR RASTER DAN KOORDINAT TARGET
- Pemilih Georeference memprioritaskan raster valid; jika tidak ada raster QGIS yang terdeteksi,
  layer valid lain di dalam grup Georeference tetap dapat dipilih.
- Perbaikan posisi kolom saat target manual disimpan: status ditulis pada kolom Status dan
  koordinat ditulis pada kolom Koordinat target, bukan kolom Kode desa/Nama SLS.
- Acuan titik pusat BPS KdKec_3210200.gpkg tetap disertakan.
- Pemeriksaan sintaks Python dilakukan; pengujian langsung pada instalasi QGIS pengguna belum dilakukan.


PEMBARUAN v2.7.3 — MATCH NOMOR RASTER & KONVERSI GEOTIFF
- Auto Geotagging diarahkan untuk mencocokkan no_bang (001/0001) dengan label digit raster (1)
  menggunakan OCR Tesseract, bukan memilih piksel gelap sebagai target.
- OCR memerlukan modul Python pytesseract dan aplikasi Tesseract OCR yang dapat dijalankan
  oleh lingkungan Python QGIS. Jika belum tersedia, plugin menampilkan peringatan.
- Hasil OCR masuk ke kolom Saran otomatis dan Skor gelap (kolom skor menampilkan confidence OCR
  dan jarak kandidat); kolom Kode desa dan Nama SLS tidak ditimpa.
- Menambahkan Toolbox > Konversi Georeference PNG/PGW/XML → GeoTIFF.
- Mendukung PNG dengan worldfile PGW/WLD dan metadata XML pendamping yang dapat dibaca GDAL,
  serta ekspor semua raster valid di dalam grup layer terpilih ke file TIF terpisah.
- GeoTIFF hasil konversi otomatis ditambahkan ke proyek; periksa CRS/georeferensi sebelum digunakan.
- Pemeriksaan sintaks Python dilakukan; pengujian runtime langsung di QGIS belum dilakukan.

- Label skor pada tabel Auto Geotagging diubah menjadi 'Skor OCR / jarak' agar tidak menyiratkan pencarian noktah gelap.


PEMBARUAN v2.7.4 — TESSERACT OCR + PYTESSERACT
- OCR memakai aplikasi Tesseract OCR dan modul Python pytesseract.
- Plugin otomatis mencari tesseract.exe melalui PATH, TESSERACT_CMD, dan beberapa lokasi standar Windows.
- Plugin memeriksa versi Tesseract sebelum memulai analisis dan menampilkan pesan spesifik jika dependensi belum siap.
- Untuk instalasi Tesseract di lokasi khusus, set variabel lingkungan TESSERACT_CMD ke path lengkap tesseract.exe.
- Nomor bangunan 001/0001 dinormalisasi menjadi angka 1 agar cocok dengan label satu digit pada raster.
- Pemeriksaan sintaks Python dilakukan; pengujian runtime harus dilakukan di instalasi QGIS pengguna.


Versi 2.8.2: menambahkan tombol Catatan Titik Bangunan pada panel utama. Pilih layer titik aktif, klik titik bermasalah di peta, lalu isi status dan catatan. Kolom STATUS_CATATAN dan CATATAN_TITIK harus sudah ada pada layer.


PEMBARUAN v2.8.3
- Catatan Titik Bangunan: toleransi klik diperbaiki, titik terdekat dicari dalam radius layar, dan kolom STATUS_CATATAN/CATATAN_TITIK dibuat otomatis bila layer mendukung pengeditan atribut.
- Tombol jendela custom menggunakan warna netral, bukan kelompok tombol merah. Minimize dan maximize/restore tetap tersedia sebelum tombol tutup.
- Menu SLS MAGIC > Cek Update GitHub memeriksa rilis terbaru dan membuka halaman rilis. Instalasi otomatis belum dilakukan; tutup QGIS sebelum mengganti plugin.
