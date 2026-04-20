# 🛡️ PROJECT NAVY - Enterprise Audit Monitoring System

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-Latest-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-Corporate-green.svg)]()

> **Corporate Edition** - Pertamina Style Concept  
> Sistem monitoring audit terintegrasi dengan UI spektakuler Navy Blue & Red Accent

---

## 📋 Table of Contents

1. [Overview](#-overview)
2. [Features](#-features)
3. [Tech Stack](#-tech-stack)
4. [Installation](#-installation)
5. [Quick Start](#-quick-start)
6. [Project Structure](#-project-structure)
7. [User Guide](#-user-guide)
8. [Configuration](#-configuration)
9. [Building Executable](#-building-executable)
10. [Troubleshooting](#-troubleshooting)
11. [API Reference](#-api-reference)

---

## 📖 Overview

**Project Dashboard** adalah sistem monitoring audit terintegrasi berbasis desktop yang dirancang untuk lingkungan korporat. Sistem ini mengadopsi arsitektur **Serverless-Local** dimana Microsoft Excel berfungsi sebagai pusat data (database), dikombinasikan dengan antarmuka modern berbasis Python Streamlit.

### 🎯 Tujuan Sistem
- Menggantikan proses manual makro VBA dengan solusi yang lebih stabil
- Menyediakan UI/UX spektakuler dengan tema Corporate (Navy Blue & Red Accent)
- Mengotomatisasi proses reminder email via Outlook Desktop Client
- Memudahkan monitoring status dokumen audit secara real-time

---

## ✨ Features

### 📊 Executive Dashboard
- **KPI Cards**: Total Dokumen, Outstanding, Need to Review, Done
- **Interactive Charts**: Bar Chart per Fungsi & Donut Chart Completion Rate
- **Smart Filtering**: Filter berdasarkan Divisi/Fungsi
- **Overdue Alerts**: Notifikasi dokumen yang melewati deadline

### 📂 Document Control Center (CRUD)
- **View**: Tabel data dengan fitur search & filter
- **Create**: Form input dengan validasi Doc. ID unik
- **Update**: Update status dan link Microsoft Form
- **Delete**: Hapus data dengan konfirmasi keamanan

### 📧 Smart Email Automation
- **Auto-Filter**: Otomatis menampilkan dokumen Outstanding/Need to Review
- **Email Composer**: Field To, CC, Subject, Body yang bisa diedit
- **HTML Template**: Template email profesional siap pakai
- **Outlook Integration**: Generate email langsung ke Outlook Desktop
- **Auto-Log**: Update Status Reminder setelah email dibuat

### ⚙️ Master Data Management
- **Master Fungsi**: Kelola list divisi dengan data editor
- **User Management**: CRUD user dengan role Admin/User

---

## 🛠 Tech Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.10+ |
| **Frontend** | Streamlit (Latest) |
| **Data Processing** | pandas, openpyxl |
| **Visualization** | Plotly Express |
| **Automation** | pywin32 (win32com.client) |
| **Packaging** | PyInstaller |

---

## 📦 Installation

### Prerequisites
- Python 3.10 atau lebih baru
- Microsoft Outlook Desktop (untuk fitur email)
- Windows OS (untuk integrasi Outlook)

### Step 1: Clone/Download Project
```bash
# Jika menggunakan git
git clone <repository-url>
cd VBA2

# Atau download dan ekstrak ZIP
```

### Step 2: Create Virtual Environment (Recommended)
```bash
# Buat virtual environment
python -m venv .venv

# Aktivasi (Windows PowerShell)
.\.venv\Scripts\Activate.ps1

# Aktivasi (Windows CMD)
.\.venv\Scripts\activate.bat
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 4: Verify Installation
```bash
# Test import modules
python -c "import streamlit; import pandas; import plotly; print('All modules OK!')"
```

---

## 🚀 Quick Start

### Option 1: Run dengan Python
```bash
# Dari folder project
streamlit run app.py
```

### Option 2: Run dengan Launcher Script
```bash
python run_app.py
```

### Option 3: Run dengan Custom Theme
```bash
streamlit run app.py --theme.primaryColor="#003366" --theme.backgroundColor="#F4F6F9"
```

Aplikasi akan terbuka di browser pada `http://localhost:8501`

---

## ☁️ Deploy ke Streamlit Cloud (Public URL)

### 1) Pastikan file deploy tersedia
Project ini sudah disiapkan untuk Streamlit Cloud dengan:
- `streamlit_app.py` (entrypoint cloud)
- `.streamlit/config.toml` (konfigurasi server + tema)
- `requirements.txt` (dependency Linux-safe, `pywin32` hanya untuk Windows)

### 2) Push project ke GitHub
```bash
git init
git add .
git commit -m "prepare streamlit cloud deploy"
git branch -M main
git remote add origin https://github.com/<username>/<repo>.git
git push -u origin main
```

### 3) Deploy dari Streamlit Cloud
1. Buka: https://share.streamlit.io/
2. Klik **New app**
3. Isi form:
    - **Repository**: `<username>/<repo>`
    - **Branch**: `main`
    - **Main file path**: `streamlit_app.py`
4. Klik **Deploy**

### 4) Catatan penting
- Fitur Outlook (`pywin32`) hanya berjalan di Windows desktop, sehingga di Streamlit Cloud fitur ini tidak aktif.
- Penyimpanan `project.xlsx` di cloud bersifat non-persisten (dapat reset saat app restart/redeploy).
- Jika butuh data permanen multi-user, pindahkan storage ke database/cloud storage.

### 5) Setup Supabase (recommended untuk production)

1. Buat project di Supabase: https://supabase.com/
2. Ambil connection string PostgreSQL (Transaction Pooler) dari **Project Settings → Database**.
3. Jalankan SQL schema di file `supabase_schema.sql` melalui **SQL Editor** Supabase.
4. Di Streamlit Cloud, buka app settings → **Secrets**, lalu isi:

```toml
SUPABASE_DB_URL = "postgresql://postgres.<project-ref>:<password>@aws-0-<region>.pooler.supabase.com:6543/postgres"
```

5. Redeploy app. Setelah `SUPABASE_DB_URL` terisi, aplikasi otomatis pakai Supabase.
6. Jika `SUPABASE_DB_URL` tidak ada, aplikasi fallback ke mode Excel lokal.

#### Migrasi data Excel lama ke Supabase (one-time)
Jalankan dari terminal lokal setelah `SUPABASE_DB_URL` diset:

```bash
python migrate_excel_to_supabase.py
```

Script akan memindahkan sheet Document, Users, Master Fungsi, dan Audit Trail ke Supabase.

### 6) Konfigurasi Email di cloud
- Integrasi Outlook Desktop (`pywin32`) tidak didukung di Streamlit Cloud.
- Untuk environment cloud, gunakan Microsoft Graph API atau SMTP Office 365 sebagai pengganti.
- Placeholder konfigurasi email cloud tersedia di `.streamlit/secrets.toml.example`.

---

## 📁 Project Structure

```
VBA2/
│
├── 📄 app.py                    # Main UI & Navigation Logic
├── 📄 utils.py                  # Backend Logic (Excel I/O, Outlook)
├── 📄 run_app.py                # Launcher Script
├── 📄 requirements.txt          # Python Dependencies
├── 📄 README.md                 # Documentation (this file)
├── 📄 Audit&MonitoringDashboardSystem.md  # Original Specification
│
├── 📊 project.xlsx              # Excel Database (auto-created)
│
└── 📁 Berkas/                   # Additional files
    ├── asli.xlsm
    ├── project.xlsm
    └── req.txt
```

### File Descriptions

| File | Description |
|------|-------------|
| `app.py` | Main Streamlit application dengan semua modul UI |
| `utils.py` | Backend functions: Excel operations, Outlook automation, helpers |
| `run_app.py` | Launcher script untuk easy start dan PyInstaller packaging |
| `requirements.txt` | List semua Python dependencies |
| `project.xlsx` | Excel database (auto-generated jika tidak ada) |

---

## 📘 User Guide

### 📊 Module 1: Executive Dashboard

1. **Akses Dashboard**: Klik menu "📊 Executive Dashboard" di sidebar
2. **View KPI**: Lihat statistik dokumen di metric cards
3. **Filter Data**: Gunakan filter "Fungsi/Divisi" untuk menyaring data
4. **Analyze Charts**: Hover chart untuk detail data

### 📂 Module 2: Document Control

#### Menambah Dokumen Baru
1. Klik tab "➕ Add Document"
2. Isi semua field yang wajib (ditandai *)
3. Gunakan ID yang di-suggest atau buat custom
4. Klik "💾 Simpan Dokumen"

#### Mengupdate Dokumen
1. Klik tab "✏️ Update Document"
2. Pilih dokumen dari dropdown
3. Update status atau link form
4. Klik "💾 Update Dokumen"

#### Menghapus Dokumen
1. Klik tab "🗑️ Delete Document"
2. Pilih dokumen dari dropdown
3. Centang konfirmasi
4. Klik "🗑️ Hapus Dokumen"

### 📧 Module 3: Email Automation

1. **Pilih Dokumen**: Pilih dokumen Outstanding/Need to Review
2. **Review Email Fields**: Edit To, CC, Subject jika perlu
3. **Edit Body**: Modifikasi template email HTML
4. **Preview**: Centang "Preview" untuk melihat tampilan
5. **Generate**: Klik "📤 Generate Outlook Email"
6. **Send di Outlook**: Review dan kirim email di Outlook

### ⚙️ Module 4: Master Data

#### Kelola Divisi
1. Klik tab "🏢 Master Fungsi/Divisi"
2. Edit langsung di tabel (tambah/hapus/edit)
3. Klik "💾 Simpan Perubahan Fungsi"

#### Kelola User
1. Klik tab "👥 User Management"
2. Edit langsung di tabel
3. Klik "💾 Simpan Perubahan User"

---

## ⚙️ Configuration

### Database Schema

#### Sheet: Document Master Data
| Column | Type | Description |
|--------|------|-------------|
| Doc. ID Number | String (PK) | Nomor unik dokumen |
| Document Description | String | Judul dokumen |
| Fungsi | String | Divisi pemilik |
| Doc. Request Date | Date | Tanggal request |
| Doc. Submission Deadline | Date | Deadline |
| Days To Go | Integer | Auto-calculated |
| Email - Auditee1 | String | Email penerima (To) |
| Recepient - cc | String | Email CC |
| Document Status | Enum | Outstanding/Need to Review/Done |
| Remarks | String | Link Microsoft Form |
| Status Reminder | String | Log pengiriman email |

#### Sheet: Users
| Column | Description |
|--------|-------------|
| Nama | Nama lengkap user |
| Email | Alamat email |
| Role | Admin / User |
| Fungsi | Divisi user |

#### Sheet: Master Fungsi
| Column | Description |
|--------|-------------|
| Nama Fungsi | Nama divisi/fungsi |

### Color Theme (Pertamina Style)
```python
PRIMARY_NAVY = "#003366"    # Sidebar & Header
ACCENT_RED = "#DA291C"      # Call-to-Action, Alerts
SUCCESS_GREEN = "#469B43"   # Done status
WARNING_AMBER = "#FFC107"   # Need to Review
BACKGROUND = "#F4F6F9"      # Light grey background
```

---

## 📦 Building Executable

### Build dengan PyInstaller

```bash
# Install PyInstaller (sudah ada di requirements.txt)
pip install pyinstaller

# Build executable
pyinstaller --onefile --windowed --name "ProjectNavy" --icon=icon.ico run_app.py

# Executable akan ada di folder dist/
```

### Build dengan Additional Files
```bash
pyinstaller --onefile --windowed \
    --name "ProjectNavy" \
    --add-data "app.py;." \
    --add-data "utils.py;." \
    --add-data "project.xlsx;." \
    run_app.py
```

### Running the Executable
1. Copy semua file ke folder yang sama:
   - `ProjectNavy.exe`
   - `app.py`
   - `utils.py`
   - `project.xlsx`
2. Jalankan `ProjectNavy.exe`

---

## 🔧 Troubleshooting

### ❌ Error: File Excel Sedang Dibuka

**Problem**: `PermissionError: File Excel sedang dibuka di aplikasi lain`

**Solution**: 
1. Tutup file Excel yang sedang terbuka
2. Coba lagi operasi yang gagal

### ❌ Error: Outlook Tidak Ditemukan

**Problem**: `Outlook Desktop tidak ditemukan atau tidak aktif`

**Solution**:
1. Pastikan Microsoft Outlook Desktop terinstall
2. Buka Outlook dan pastikan sudah ter-configure
3. Restart aplikasi

### ❌ Error: Module Not Found

**Problem**: `ModuleNotFoundError: No module named 'xxx'`

**Solution**:
```bash
# Reinstall dependencies
pip install -r requirements.txt
```

### ❌ Streamlit Port Already in Use

**Problem**: `Address already in use`

**Solution**:
```bash
# Run di port berbeda
streamlit run app.py --server.port 8502
```

### ❌ Excel File Corrupt

**Problem**: File project.xlsx corrupt atau tidak readable

**Solution**:
1. Backup file yang corrupt
2. Hapus file `project.xlsx`
3. Restart aplikasi (file baru akan auto-generate)

---

## 📚 API Reference

### utils.py Functions

#### Data Loading
```python
load_documents() -> pd.DataFrame
# Load semua dokumen dari Excel

load_users() -> pd.DataFrame
# Load semua user dari Excel

load_fungsi() -> pd.DataFrame
# Load master fungsi dari Excel

get_fungsi_list() -> List[str]
# Get list fungsi untuk dropdown
```

#### Data Saving
```python
save_documents(df: pd.DataFrame) -> Tuple[bool, str]
# Simpan DataFrame dokumen ke Excel

save_users(df: pd.DataFrame) -> Tuple[bool, str]
# Simpan DataFrame users ke Excel

save_fungsi(df: pd.DataFrame) -> Tuple[bool, str]
# Simpan DataFrame fungsi ke Excel
```

#### Document Operations
```python
add_document(doc_data: dict) -> Tuple[bool, str]
# Tambah dokumen baru (validasi ID unik)

update_document(doc_id: str, updates: dict) -> Tuple[bool, str]
# Update dokumen berdasarkan ID

delete_document(doc_id: str) -> Tuple[bool, str]
# Hapus dokumen berdasarkan ID
```

#### Email Functions
```python
get_documents_for_reminder() -> pd.DataFrame
# Get dokumen yang perlu reminder (Outstanding/Need to Review)

generate_email_subject(doc_data: dict) -> str
# Generate subject email otomatis

generate_email_body(doc_data: dict) -> str
# Generate HTML body email

generate_outlook_email(to, cc, subject, body_html, doc_id) -> Tuple[bool, str]
# Generate dan tampilkan email di Outlook
```

#### Helper Functions
```python
calculate_days_to_go(deadline) -> Optional[int]
# Hitung selisih hari ke deadline

get_status_color(status: str) -> str
# Get warna berdasarkan status

generate_doc_id(prefix: str) -> str
# Generate ID dokumen otomatis

get_document_statistics() -> dict
# Get statistik untuk dashboard
```

---

## 📝 Changelog

### v1.0.0 (2026-02-18)
- Initial release
- Executive Dashboard dengan KPI cards dan charts
- Document Control Center (CRUD)
- Smart Email Automation dengan Outlook integration
- Master Data Management
- Pertamina Corporate Style UI

---

## 👥 Author

**Project Navy Development Team**

---

## 📄 License

This project is proprietary software developed for corporate use.

---

<div align="center">

**🛡️ PROJECT NAVY**

*Enterprise Audit Monitoring System*

Corporate Edition © 2026

</div>


---

<div align="center">

[![Made by CocoTech](https://img.shields.io/badge/Made%20by-CocoTech%20Studio-7C3AED?style=flat-square&logo=vercel&logoColor=white)](https://cocotech.studio)

*Built with passion · [cocotech.studio](https://cocotech.studio)*

</div>