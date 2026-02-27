🛡️ PROJECT NAVY: ENTERPRISE AUDIT MONITORING SYSTEM
Corporate Edition (Pertamina Style Concept)
1. Executive Summary
Project Navy adalah sistem monitoring audit terintegrasi berbasis desktop (Standalone Executable) yang dirancang untuk lingkungan korporat. Sistem ini mengadopsi arsitektur Serverless-Local di mana Microsoft Excel (project.xlsx) berfungsi sebagai pusat data (database), dikombinasikan dengan antarmuka modern berbasis Python Streamlit.
Sistem ini difokuskan untuk menggantikan proses manual makro VBA dengan solusi yang lebih stabil, aman, memiliki UI/UX spektakuler (Navy Blue & Red Accent), serta kemampuan otomatisasi email via Outlook Desktop Client.

2. Technical Stack & Architecture
Agar AI Agent dapat membangun sistem ini dengan presisi, gunakan spesifikasi teknis berikut:
Core Language: Python 3.10+
Frontend Framework: Streamlit (Versi Terbaru)
Data Engine: pandas (Data Manipulation) & openpyxl (Excel I/O - Preserve Formatting)
Visualization: Plotly Express (Interactive & Responsive Charts)
Automation: pywin32 (win32com.client) untuk integrasi Native Outlook.
Packaging: PyInstaller (untuk kompilasi menjadi .exe portable).

3. Database Schema Specification (project.xlsx)
Aplikasi WAJIB membaca dan menulis ke file Excel ini. Jika file/sheet tidak ditemukan, aplikasi harus memilik fitur Auto-Healing (membuat sheet baru secara otomatis).
A. Sheet: Document Master Data (Transaksi Utama)
Sheet ini menampung seluruh data monitoring dokumen audit.
| Column Name | Data Type | Description |
| :--- | :--- | :--- |
| Doc. ID Number | String (PK) | Nomor unik dokumen (Contoh: Fin-2026-001). |
| Document Description | String | Judul/Deskripsi dokumen audit. |
| Fungsi | String | Divisi pemilik dokumen (Lookup ke Master Fungsi). |
| Doc. Request Date | Date | Tanggal permintaan dokumen. |
| Doc. Submission Deadline | Date | Tanggal jatuh tempo. |
| Days To Go2 | Integer | Formula Python: Deadline - Today. |
| Email - Auditee1 | String | Email utama penerima (To). |
| Recepient - cc | String | Email tembusan (CC), dipisah koma. |
| Document Status | String (Enum) | Pilihan: Outstanding, Need to Review, Done. |
| Remarks | String | Link Output Microsoft Form (Respon Auditee). |
| Status Reminder | String | Log audit pengiriman email (Contoh: "Sent: 2026-02-18"). |
B. Sheet: Users (Manajemen User)


Column Name
Description
Nama
Nama lengkap user.
Email
Alamat email user.
Role
Hak akses (Admin/User).
Fungsi
Divisi user tersebut.

C. Sheet: Master Fungsi (Manajemen Divisi)
Column Name
Description
Nama Fungsi
Daftar nama divisi/fungsi untuk dropdown (Contoh: Finance, HC, Operation).

4. UI/UX Design System (Corporate & Spectacular)
AI Agent harus menerapkan Custom CSS Injection untuk mencapai tampilan "High-End Corporate".
🎨 Color Palette (Pertamina DNA)
Primary (Navy Blue): #003366 (Sidebar & Header)
Accent (Red): #DA291C (Call-to-Action Buttons, Alert Badges)
Background: #F4F6F9 (Light Grey - Professional Clean Look)
Success: #469B43 (Green Pertamina)
Warning: #FFC107 (Amber)
🖥️ Layout & Components
Sidebar: Logo Perusahaan di atas, Menu Navigasi (Radio Button/Option Menu), Copyright Footer.
Metric Cards: Kotak metrik dengan border kiri berwarna (Left Border Accent), bayangan halus (Soft Shadow), dan indikator Delta.
Typography: Font modern (Segoe UI / Roboto), Judul H1/H2 berwarna Navy Blue.

5. Functional Modules (Fitur Detail)
📊 Module 1: Executive Dashboard (The "Wow" Factor)
KPI Cards: Tampilkan 4 kartu utama: Total Dokumen, Outstanding (Merah), Need to Review (Kuning), Done (Hijau).
Interactive Charts:
Bar Chart: Status Dokumen per Fungsi (Grouped Bar).
Donut Chart: Persentase penyelesaian audit.
Smart Filter: Multiselect untuk memfilter dashboard berdasarkan "Fungsi/Divisi".
📂 Module 2: Document Control Center (CRUD)
View: Tabel data (st.dataframe) dengan fitur sort & search bawaan.
Create (Add): Form input lengkap. Validasi Doc. ID Number tidak boleh duplikat.
Update: Fitur untuk mengupdate status dan mengisi kolom Remarks.
Note: Kolom Remarks harus dilabeli sebagai "Link Respon Microsoft Form".
Delete: Tombol hapus data dengan konfirmasi keamanan.
📧 Module 3: Smart Email Automation (Outlook Integration)
Fitur paling krusial. Sistem TIDAK mengirim email langsung (SMTP), melainkan membuka Popup Outlook Desktop.
Logic Filter: Otomatis menampilkan daftar dokumen yang statusnya "Outstanding" ATAU "Need to Review".
UI Compose Email (Terpisah & Rapi):
Dropdown: Pilih No. Dokumen yang ingin di-remind.
Field "To": Terisi otomatis dari database (Bisa diedit).
Field "CC": Terisi otomatis dari database (Bisa diedit, terpisah dari "To").
Field "Subject": Auto-generate: REMINDER: [Status] - [Nama Dokumen].
Text Area "Body": Template HTML default yang sopan (Yth. Rekan...), bisa diedit user.
Action Button: Tombol "Generate Outlook Email".
Backend Logic: Gunakan win32com.client.Dispatch('outlook.application'). Buat item email, isi To, CC, Subject, Body, lalu panggil .Display() (Bukan .Send()).
Post-Action: Update kolom Status Reminder di Excel menjadi "Sent: [Tanggal Hari Ini]".
⚙️ Module 4: Master Data Management
Master Divisi: Gunakan st.data_editor agar user bisa menambah/hapus/edit nama Divisi langsung di tabel (Excel-like experience).
User Profile: CRUD sederhana untuk menambah user yang berhak mengakses (jika nanti dikembangkan ke login system).

6. Implementation Guide for AI Agent (Instruction Manual)
Prompt untuk AI:
"Buatkan saya kode Python lengkap (app.py dan utils.py) berdasarkan spesifikasi README.md di atas. Pastikan kode menangani error PermissionError jika file Excel sedang terbuka. Gunakan styling CSS yang didefinisikan untuk membuat tampilan 'Pertamina Corporate Style'. Jangan lupa buatkan fungsi generate_outlook_email menggunakan pywin32."
Struktur File yang Diharapkan:
Project_Navy/
│
├── project.xlsx            # Database (Wajib)
├── app.py                  # Main UI & Navigation Logic
├── utils.py                # Backend Logic (Load Data, Save Data, Outlook Trigger)
├── requirements.txt        # Library: streamlit, pandas, openpyxl, plotly, pywin32
└── run_app.py              # Launcher Script untuk PyInstaller


