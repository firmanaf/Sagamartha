# Sagamartha HR Enterprise Desktop Production-Ready v4

Perubahan utama:
- setiap karyawan sekarang dapat mengisi multi target dalam 1 hari
- ditambahkan field:
  - Judul Proyek
  - Penjelasan Target
- logika anti-duplikasi diubah:
  - multi target per hari diperbolehkan
  - hanya data yang benar-benar identik yang ditolak
- daftar pegawai dan role sudah disesuaikan dengan data terbaru pengguna

Role mapping dalam aplikasi:
- Admin -> admin
- Supervisor -> atasan
- Pegawai -> karyawan

Catatan akun:
- seluruh user bawaan memakai password awal: 12345678
- seluruh user dipaksa ganti password saat login pertama
- file user juga disertakan terpisah agar mudah direview

Catatan supervisor:
- pegawai dibagi otomatis ke daftar supervisor yang tersedia
- bila ingin mapping bawahan tertentu secara spesifik, edit file JSON user
