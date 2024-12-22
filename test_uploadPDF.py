import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
import mysql.connector
import pandas as pd
from pandasgui import show
import tabula
import fitz

# ======= Bagian DATABASE ==========
def masukkan_ke_database(data_laporan):
    conn = mysql.connector.connect(
        host='localhost',
        user='root',
        password='',
        database='db_keuangan'
    )
    
    cursor = conn.cursor()
    sql = "INSERT INTO tb_laporan_keuangan (kode_emiten, nama_emiten, tahun, quartal, grup_laporan_keuangan, item, nilai, notes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"

    for i in range(len(data_laporan)):
        cursor.execute(sql, (informasi_tambahan[0], informasi_tambahan[1], informasi_tambahan[3], informasi_tambahan[2], grup_laporan_keuangan[0], data_laporan.iloc[i][0], data_laporan.iloc[i][2], data_laporan.iloc[i][1]))
        conn.commit()

    print(cursor.rowcount, "Data berhasil disimpan")
    
    cursor.close()
    conn.close()

# ======= Bagian PDF ==========
def baca_pdf(nama_file, halaman):
    area = [160, 0, 715, 595]
    tables = tabula.read_pdf(nama_file, pages=halaman, multiple_tables=True, area=area)
    return tables

def baca_pdf_atas(nama_file, halaman):
    area = [70, 0, 150, 595]
    tables = tabula.read_pdf(nama_file, pages=halaman, multiple_tables=True, area=area)
    return tables

def gabung_tabel(tables):
    df = pd.concat(tables, ignore_index=True)
    df = df.dropna(axis=1, how='all')  # Menghapus kolom yang seluruhnya kosong
    return df

def bersihkan_header(df):
    df = df[4:].reset_index(drop=True)
    return df

def hilangkan_baris_kosong(df):
    df = df.dropna(how='all').reset_index(drop=True)
    return df

def set_header(df):
    header = ["Aset", "Catatan/Notes", "2023", "2022", "Assets"]
    df.columns = header
    return df

def hilangkan_nan(df):
    df.fillna("-", inplace=True)
    return df

def cari_halaman_mengandung_teks(nama_file, kata_kunci):
    halaman_ditemukan = []
    doc = fitz.open(nama_file)
    
    for halaman_num in range(doc.page_count):
        halaman = doc[halaman_num]
        teks = halaman.get_text()
        if kata_kunci in teks:
            halaman_ditemukan.append(halaman_num + 1)  # fitz menggunakan indeks 0, jadi tambahkan 1 untuk halaman yang benar

    doc.close()

    # Hapus halaman yang non berurutan
    i = 1
    while i < len(halaman_ditemukan):
        if halaman_ditemukan[i] - halaman_ditemukan[i - 1] > 1:
            del halaman_ditemukan[i]
        else:
            i += 1
    return halaman_ditemukan

def pilih_file_pdf():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    for file_path in file_paths:
        if file_path not in selected_files_pdf:
            selected_files_pdf.append(file_path)
            file_listbox.insert(tk.END, f"PDF: {file_path}")

def proses_files_pdf():
    proses_files_excel()
    kata_kunci = ["LAPORAN POSISI KEUANGAN", "LAPORAN LABA RUGI DAN PENGHASILAN", "LAPORAN ARUS KAS KONSOLIDASIAN"]
    for kata in kata_kunci:
        if(kata == "LAPORAN POSISI KEUANGAN"):
            grup_laporan_keuangan[0] = "Laporan Neraca"
        elif(kata == "LAPORAN LABA RUGI DAN PENGHASILAN"):
            grup_laporan_keuangan[0] = "Laporan Laba Rugi"
        else:
            grup_laporan_keuangan[0] = "Laporan Arus Kas"
        
        all_tables = []
        
        for file_path in selected_files_pdf:
            halaman_ditemukan = cari_halaman_mengandung_teks(file_path, kata)
            if not halaman_ditemukan:
                messagebox.showwarning("Tidak Ditemukan", f"Kata kunci '{kata}' tidak ditemukan dalam file {file_path}.")
                continue

            print(halaman_ditemukan)

            for halaman in halaman_ditemukan:
                tables = baca_pdf(file_path, halaman)
                if tables:
                    df = pd.concat(tables, ignore_index=True)
                    df = bersihkan_header(df)
                    df = hilangkan_baris_kosong(df)
                    df = set_header(df)
                    all_tables.append(df)

        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df = hilangkan_nan(final_df)
            final_df.to_excel("output.xlsx", index=False)

            show(final_df)
            masukkan_ke_database(final_df)
            messagebox.showinfo("Proses Selesai", "File PDF {kata} telah diproses dan data disimpan ke database.")
        else:
            messagebox.showwarning("Proses Dibatalkan", "Tidak ada tabel yang ditemukan untuk diproses.")

    # Menghapus seluruh path jika sudah diproses
    file_listbox.delete(0, tk.END)
    selected_files_pdf.clear()
    selected_files_excel.clear()
    informasi_tambahan.clear()

# ======= Bagian EXCEL ==========
def pilih_file_excel():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    for file_path in file_paths:
        if file_path not in selected_files_excel:
            selected_files_excel.append(file_path)
            file_listbox.insert(tk.END, f"Excel: {file_path}")

def proses_files_excel():
    # Untuk mencari informasi umum
    for file_path in selected_files_excel:
        try:
            # Informasi umum biasanya ada di sheet "1000000"
            df = pd.read_excel(file_path, sheet_name="1000000") 
            for i in range(1, len(df)):
                if(df.iloc[i][0] == "Kode entitas"):
                    kode_emiten = df.iloc[i][1]
                if(df.iloc[i][0] == "Nama entitas"):
                    nama_emiten = df.iloc[i][1]
                if(df.iloc[i][0] == "Periode penyampaian laporan keuangan"):
                    if(df.iloc[i][1] == "Kuartal I / First Quarter"):
                        quartal = 1
                    elif (df.iloc[i][1] == "Kuartal II / Second Quarter"):
                        quartal = 2
                    elif (df.iloc[i][1] == "Kuartal III / Third Quarter"):
                        quartal = 3
                    else:
                        quartal = 4
                if(df.iloc[i][0] == "Tanggal awal periode berjalan"):
                    tahun = int(str(df.iloc[i][1])[:4])
                
            informasi_tambahan.append(kode_emiten)
            informasi_tambahan.append(nama_emiten)
            informasi_tambahan.append(quartal)
            informasi_tambahan.append(tahun)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel: {e}")

# ======= Kepentingan UI ==========
def hapus_file():
    # Dapatkan indeks item yang dipilih di Listbox
    selected_index = file_listbox.curselection()
    if not selected_index:
        messagebox.showwarning("Peringatan", "Pilih file yang ingin dihapus.")
        return

    # Dapatkan teks item yang dipilih
    selected_item = file_listbox.get(selected_index)

    # Hapus item dari Listbox
    file_listbox.delete(selected_index)

    # Hapus dari daftar selected_files_pdf atau selected_files_excel
    if selected_item.startswith("PDF: "):
        file_path = selected_item.replace("PDF: ", "")
        if file_path in selected_files_pdf:
            selected_files_pdf.remove(file_path)
    elif selected_item.startswith("Excel: "):
        file_path = selected_item.replace("Excel: ", "")
        if file_path in selected_files_excel:
            selected_files_excel.remove(file_path)

# ============ Membuat UI menggunakan tkinter ==============
root = tk.Tk()
root.title("PDF dan Excel Uploader dan Processor")

selected_files_pdf = []   # Menyimpan daftar file PDF yang dipilih
selected_files_excel = [] # Menyimpan daftar file Excel yang dipilih
informasi_tambahan = [] # Informasi kode emiten, nama emite, quartal, dan tahun
grup_laporan_keuangan = ["Kosong"] # Informasi grup laporan keuangan

# Tombol untuk menambahkan file PDF
tambah_file_pdf_button = tk.Button(root, text="Tambah File PDF", command=pilih_file_pdf)
tambah_file_pdf_button.pack(pady=10)

# Tombol untuk menambahkan file Excel
tambah_file_excel_button = tk.Button(root, text="Tambah File Excel", command=pilih_file_excel)
tambah_file_excel_button.pack(pady=10)

# Listbox untuk menampilkan daftar file yang dipilih
file_listbox = Listbox(root, width=160, height=20)
file_listbox.pack(pady=10)

# Tombol untuk menghapus file yang dipilih
hapus_file_button = tk.Button(root, text="Hapus File yang Dipilih", command=hapus_file)
hapus_file_button.pack(pady=10)

# Tombol untuk memproses file PDF
proses_pdf_button = tk.Button(root, text="PROSES FILE!", command=proses_files_pdf)
proses_pdf_button.pack(pady=10)

root.mainloop()