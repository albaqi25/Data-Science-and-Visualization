# 🛒 DQLab Hackathon: Retail Crisis & Recovery

Repositori ini berisi solusi script Python untuk menyelesaikan studi kasus **Online Hackathon: Retail Crisis & Recovery Visualization Challenge** dari DQLab.

## 📝 Deskripsi Proyek
Proyek ini bertujuan untuk membantu manajemen toko retail mengatasi penurunan penjualan dengan cara menggali wawasan tersembunyi dari data transaksi. Script ini secara otomatis mencari produk **"Rising Star"** (produk dengan tren naik yang terabaikan) dan merancang **"Potential Packaging"** (strategi *bundling* produk).

## 🔍 Fitur Analisis
1. **Identifikasi Rising Star:** Menggunakan metode *Moving Average* (MA) 3-hari untuk menyaring produk yang mengalami tren kenaikan penjualan konsisten minimal 12 hari berturut-turut.
2. **Market Basket Analysis:** Menggunakan Algoritma Apriori untuk mencari rekomendasi produk pendamping (*cross-selling*) berdasarkan metrik *Lift* > 2 dan *Support* > 1%.
3. **Visualisasi Data:** Menghasilkan dua grafik Matplotlib untuk membandingkan produk Rising Star dengan *Top 3* penjualan (berbasis Indeks 100 dan nilai aktual).
4. **Automated Reporting:** Mengekspor hasil analisis secara rapi ke dalam file Excel terformat.

## 💻 Teknologi yang Digunakan
* **Bahasa:** Python
* **Data Manipulasi:** Pandas
* **Algoritma Apriori:** mlxtend
* **Visualisasi:** Matplotlib
* **Ekspor Data:** Openpyxl

## 🚀 Cara Menjalankan Script
1. Pastikan dataset transaksi berada di dalam folder yang sama.
2. Instal *library* yang dibutuhkan dengan perintah: `pip install pandas matplotlib mlxtend openpyxl`
3. Eksekusi program melalui terminal: `python solusi-retail.py`
4. Hasil analisis akan otomatis muncul di folder Anda dalam bentuk file `.xlsx` dan `.png`.
