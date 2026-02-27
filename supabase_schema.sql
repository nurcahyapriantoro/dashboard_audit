-- Jalankan di Supabase SQL Editor

CREATE TABLE IF NOT EXISTS documents (
    doc_id TEXT PRIMARY KEY,
    document_description TEXT,
    fungsi TEXT,
    doc_request_date TIMESTAMP NULL,
    doc_submission_deadline TIMESTAMP NULL,
    days_to_go INTEGER NULL,
    email_auditee1 TEXT,
    recipient_cc TEXT,
    document_status TEXT,
    remarks TEXT,
    status_reminder TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS users_data (
    id BIGSERIAL PRIMARY KEY,
    nama TEXT,
    email TEXT,
    role TEXT,
    fungsi TEXT
);

CREATE TABLE IF NOT EXISTS master_fungsi (
    nama_fungsi TEXT PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS audit_logs (
    id BIGSERIAL PRIMARY KEY,
    timestamp TIMESTAMPTZ,
    username TEXT,
    action TEXT,
    doc_id TEXT,
    details TEXT
);

INSERT INTO master_fungsi (nama_fungsi)
VALUES
    ('Finance & Risk Management'),
    ('HC & Business Support'),
    ('Operation'),
    ('Legal & Compliance'),
    ('Information Technology'),
    ('Marketing & Sales'),
    ('Corporate Planning')
ON CONFLICT (nama_fungsi) DO NOTHING;
