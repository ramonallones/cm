CERUTUMURAH ORDER MANAGER - GOOGLE SHEETS VERSION

DEPLOY ONLINE
1. Buat Google Spreadsheet baru.
2. Buka Extensions > Apps Script.
3. Salin isi google-apps-script/Code.gs ke editor Apps Script.
4. Jalankan fungsi setupSheets_ satu kali dan izinkan akses.
5. Deploy > New deployment > Web app.
6. Execute as: Me. Who has access: Anyone.
7. Salin URL Web App lalu ganti URL_APPS_SCRIPT di config.js.
8. Upload isi folder ini ke repository GitHub dan aktifkan GitHub Pages.

API yang tersedia:
- GET getDashboard, getProducts, getOrders
- POST addProduct, updateProduct, nonaktifProduct, addStock, saveOrder

Mode fallback:
Jika API_URL masih "URL_APPS_SCRIPT", aplikasi tetap memakai data lokal browser.

Frontend tetap web statis HTML, CSS, dan JavaScript tanpa Node.js atau framework.

Cara pakai:
1. Extract ZIP.
2. Klik index.html.

DATABASE PRODUK
Database produk utama berbentuk Excel:
database/database-produk-cerutumurah.xlsx

Cara edit produk:
1. Buka database/database-produk-cerutumurah.xlsx
2. Edit/tambah produk di Excel
3. Simpan file Excel
4. Di aplikasi buka Master Produk
5. Klik Import Excel Produk
6. Pilih file Excel tersebut

EXPORT
- Export Produk menghasilkan file Excel .xls
- Export Order menghasilkan file Excel .xls
File .xls bisa dibuka di Microsoft Excel.

Catatan:
Data kerja tetap tersimpan di browser PC itu.
Gunakan Backup JSON rutin untuk cadangan aplikasi.

Total produk awal: 493


UPDATE V2
- Order tersimpan bisa dihapus satu per satu.
- Order tersimpan bisa diceklis lalu dihapus massal.
- Order tersimpan bisa diceklis lalu export Excel hanya yang dipilih.
