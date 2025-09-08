# Tujuan
Project ini bertujuan untuk mengotomatiskan alur kerja manual yang repetitif dalam pembuatan tiket, yang secara historis memakan waktu dan rentan terhadap kesalahan manusia. Dengan memanfaatkan Python dan Selenium WebDriver, sebuah bot cerdas dikembangkan untuk meniru interaksi pengguna secara presisi, mulai dari login, navigasi, pengisian formulir dinamis, hingga verifikasi keberhasilan tugas. Hasilnya adalah peningkatan efisiensi operasional yang signifikan, pengurangan human error, dan penghematan waktu kerja yang berharga.

# Fitur Utama

End-to-end otomatis: login → buka halaman DP → isi form → verifikasi sukses.

Tahan banting: auto-retry (maks 3x) + refresh & re-navigasi kalau halaman/elemen rewel.

Pilih opsi yang tepat: dropdown dinamis (model Vue) dipilih pakai pencocokan teks exact.

Verifikasi akurat: pantau respons API (bukan pop-up UI) biar kepastian suksesnya valid.

Konfigurasi rapi & aman: .env untuk kredensial/URL/proxy; gampang dipindah lingkungan.

Logging lengkap: tiap sesi ada file log buat debug & audit.


# Alur Kerja

Muat config & data

Start WebDriver + login

Buka halaman DP

Loop setiap Kode DP

Isi form (dropdown dinamis = pilih opsi exact)

Cek respons API create-ticket → sukses/gagal

Retry kalau gagal (maks 3x) → lanjut ke item berikutnya

Tulis laporan akhir + simpan log

# Catatan Teknis Penting

Dropdown dinamis: pilih <li> pada menu dengan normalize-space()='<teks yang sama persis>' → menghindari salah pilih hasil parsial.

Retry cerdas: tiap gagal akan refresh + re-navigasi sebelum coba lagi.

Verifikasi via API: injeksi JS untuk “mengintip” respons create-ticket; status sukses dibaca langsung dari serve
