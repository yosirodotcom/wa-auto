import pywhatkit
import pandas as pd
import time
import pyautogui
from datetime import datetime
import os


def get_waktu_sapaan():
    """Fungsi untuk menentukan Pagi/Siang/Sore/Malam"""
    jam = datetime.now().hour
    if 3 <= jam < 11:
        return "Pagi"
    elif 11 <= jam < 15:
        return "Siang"
    elif 15 <= jam < 18:
        return "Sore"
    else:
        return "Malam"


def kirim_pesan_bulk(file_path):
    """Fungsi utama pengiriman pesan"""

    # Cek apakah file ada sebelum lanjut
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' tidak ditemukan.")
        return

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error membaca Excel: {e}")
        return

    total_data = len(df)
    print(f"--- Memulai Proses Pengiriman ke {total_data} Kontak ---\n")

    for index, row in df.iterrows():
        # Ambil data per baris
        nomor = str(row['Nomor'])
        sapaan = row['Sapaan']  # Kolom: Bapak/Ibu/Kak
        nama = row['Nama']  # Kolom: Nama orang
        pesan_isi = row['Pesan']  # Kolom: Isi pesan

        # Validasi Nomor HP
        if nomor.startswith('0'):
            nomor = '+62' + nomor[1:]
        elif not nomor.startswith('+'):
            nomor = '+62' + nomor

        # Logic Greeting Otomatis
        waktu = get_waktu_sapaan()  # Pagi/Siang/Sore

        # Format: "Selamat Pagi Bapak Budi,"
        greeting = f"Selamat {waktu} {sapaan} {nama},"

        # Gabungkan semua (greeting + enter + pesan excel)
        pesan_final = f"{greeting}\n\n{pesan_isi}"

        print(f"[{index + 1}/{total_data}] Mengirim ke {nama} ({nomor})...")

        try:
            # Kirim pesan
            pywhatkit.sendwhatmsg_instantly(
                phone_no=nomor,
                message=pesan_final,
                wait_time=20,  # Waktu tunggu loading WA Web (detik)
                tab_close=True,  # Tutup tab setelah kirim? Ya.
                close_time=5  # Tutup tab setelah X detik
            )

            # Jeda agar aman dari deteksi spam
            time.sleep(8)

        except Exception as e:
            print(f"--> Gagal mengirim ke {nama}: {e}")

    print("\n--- Selesai! Semua tugas telah dijalankan. ---")


# --- BAGIAN MAIN ---
# Kode di bawah ini hanya akan jalan jika file ini dijalankan langsung,
# bukan jika di-import sebagai modul.
if __name__ == "__main__":
    nama_file = 'data_pengiriman.xlsx'

    # Konfirmasi user sebelum mulai (opsional, biar tidak kaget)
    input(f"Pastikan WA Web sudah login. Tekan ENTER untuk mulai membaca '{nama_file}'...")

    kirim_pesan_bulk(nama_file)