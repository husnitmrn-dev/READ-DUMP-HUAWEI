# Huawei Dump Viewer

Aplikasi desktop sederhana untuk mencari Site ID pada workbook dump Huawei dan mengekspor hasilnya ke XLSX.

## Instalasi

Gunakan Python 3.10 atau yang lebih baru.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Menjalankan

Letakkan workbook dump di folder yang sama dengan `BACA_DUMP_HW.py`. Nama workbook utama yang dicari aplikasi adalah:

`DUMP_Parshing_ALL SITE JABO CCSI & WPCList.xlsx`

Lalu jalankan:

```powershell
python BACA_DUMP_HW.py
```

File Excel tidak disertakan di repository karena ukurannya besar dan dapat berisi data operasional. Jangan unggah workbook tersebut ke repository publik.

## Upload ke GitHub

Buat repository kosong di GitHub, lalu jalankan dari folder proyek:

```powershell
git init
git add BACA_DUMP_HW.py README.md requirements.txt .gitignore
git commit -m "Initial Huawei dump viewer"
git branch -M main
git remote add origin https://github.com/USERNAME/NAMA-REPOSITORY.git
git push -u origin main
```

Ganti `USERNAME/NAMA-REPOSITORY` dengan alamat repository GitHub Anda.
