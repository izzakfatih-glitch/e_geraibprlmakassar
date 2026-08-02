"""
Aplikasi Web Penggabung Proposal PKKPRL
=========================================
Alur: Upload 2 PDF -> halaman Review (preview dokumen penuh + form koreksi
data) -> klik "Generate Dokumen Final" -> file Word diunduh.

HTML ditanam langsung di dalam file ini (tidak pakai folder templates/)
supaya tidak ada masalah TemplateNotFound di berbagai platform hosting.

CARA MENJALANKAN (LOKAL / TES):
    pip install -r requirements.txt --break-system-packages
    export ANTHROPIC_API_KEY="sk-ant-..."   # opsional, untuk fallback ekstraksi
    python3 app.py
    -> buka http://localhost:5000 di browser
"""
import os
import uuid
import shutil
import traceback
from flask import Flask, request, render_template_string, send_file, after_this_request
import mammoth

from extract import extract_proposal_with_fallback, extract_laporan_with_fallback
from generate_docx import build_document
from review_fields import FIELD_GROUPS, form_field_name, get_value, apply_form_values
import job_store

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
JOBS_DIR = os.path.join(BASE_DIR, "jobs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(JOBS_DIR, exist_ok=True)

MAX_CONTENT_LENGTH = 30 * 1024 * 1024  # 30 MB batas unggah per file gabungan

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

print(f"[startup] BASE_DIR = {BASE_DIR}")
print(f"[startup] Isi BASE_DIR = {os.listdir(BASE_DIR)}")

BASE_CSS = """
:root { --navy:#1F4E79; --bg:#f4f6f8; }
* { box-sizing: border-box; }
body {
  font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  background: var(--bg); margin: 0; padding: 0; color: #222;
}
.wrap { max-width: 640px; margin: 0 auto; padding: 32px 20px 60px; }
.wrap.wide { max-width: 900px; }
h1 { color: var(--navy); font-size: 22px; margin-bottom: 4px; }
p.sub { color: #555; margin-top: 0; font-size: 14px; }
.card { background: #fff; border-radius: 12px; padding: 24px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-top: 20px; }
label { display:block; font-weight:600; margin-bottom:6px; margin-top:14px; font-size: 13px; color:#333; }
input[type=file] { display:block; width:100%; padding: 10px; border: 1px dashed #aaa;
                    border-radius: 8px; background:#fafafa; font-size: 13px; }
input[type=text] { display:block; width:100%; padding: 9px 10px; border: 1px solid #ccc;
                    border-radius: 6px; font-size: 13px; }
button, .btn { margin-top: 20px; width: 100%; background: var(--navy); color:#fff;
       border:none; padding: 14px; border-radius: 8px; font-size: 15px;
       font-weight: 600; cursor:pointer; display:block; text-align:center; text-decoration:none; }
button:hover, .btn:hover { background:#163a5c; }
.note { font-size: 12px; color:#777; margin-top: 16px; line-height:1.5; }
.flash { background:#fff3cd; border:1px solid #ffe08a; padding:12px 14px;
         border-radius:8px; margin-top:16px; font-size:13px; color:#7a5b00; }
.spinner { display:none; text-align:center; margin-top:18px; font-size:13px; color:var(--navy); }
details { border:1px solid #e2e2e2; border-radius:8px; margin-top:12px; padding: 4px 14px 12px; background:#fbfbfd; }
summary { cursor:pointer; font-weight:700; color:var(--navy); padding:10px 0; font-size:14px; }
.preview-box { border:1px solid #ddd; border-radius:10px; background:#fff; padding: 30px 40px;
               max-height: 720px; overflow-y: auto; font-size: 14px; line-height:1.5; }
.preview-box table { border-collapse: collapse; width:100%; margin: 10px 0; }
.preview-box td { border:1px solid #999; padding:5px 8px; vertical-align: top; }
.preview-box img { max-width: 100%; height:auto; margin: 8px 0; }
.preview-box p { margin: 6px 0; }
.top-actions { position: sticky; top:0; background: var(--bg); padding: 10px 0; z-index:5; }
"""

LANDING_CSS = """
:root {
  --navy:#123A63; --blue:#1E63C7; --blue2:#2F7FE0; --bg:#eef3f8;
  --line:#e3e9f0; --ink:#1c2b3a; --muted:#5b6b7c; --green:#1f9d55; --green-bg:#eafaf0;
}
* { box-sizing: border-box; }
html, body { margin:0; padding:0; }
body { font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif; background: var(--bg); color: var(--ink); }
a { text-decoration:none; }

/* ---- Header ---- */
.site-header { background:#fff; border-bottom:1px solid var(--line); padding:18px 40px;
  display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:20px; }
.brand { display:flex; align-items:center; gap:14px; }
.brand-badge { width:44px; height:44px; border-radius:50%; flex:none;
  background: conic-gradient(from 220deg, var(--blue2), var(--navy), var(--blue2));
  display:flex; align-items:center; justify-content:center; }
.brand-badge svg { width:22px; height:22px; }
.brand-text .t1 { font-size:19px; font-weight:800; color:var(--navy); line-height:1.1; }
.brand-text .t2 { font-size:11.5px; font-weight:700; color:var(--blue); letter-spacing:.03em; }
.header-right { display:flex; align-items:center; gap:18px; }
.gov-block { display:flex; align-items:center; gap:12px; }
.gov-badge { width:52px; height:52px; border-radius:50%; flex:none; background:#f0f4f9;
  border:1px solid var(--line); display:flex; align-items:center; justify-content:center; padding:6px; }
.brand-logo-img { height:76px; width:auto; display:block; }
.gov-badge img { max-width:100%; max-height:100%; object-fit:contain; }
.gov-badge.rect { width:auto; height:48px; min-width:48px; border-radius:8px; padding:5px 10px;
  background:#fff; border:1px solid var(--line); }
.gov-badge.rect img { height:100%; width:auto; }
.gov-text { font-size:12px; line-height:1.35; color:var(--navy); font-weight:700; max-width:170px; }
.header-divider { width:1px; align-self:stretch; background:var(--line); }
.brl-text .t1 { font-size:15px; font-weight:800; color:var(--navy); line-height:1.1; }
.brl-text .t2 { font-size:10.5px; font-weight:700; color:var(--muted); letter-spacing:.05em; }

.navbar { display:flex; align-items:center; gap:26px; flex-wrap:wrap; }
.nav-link { display:flex; align-items:center; gap:6px; font-size:14px; font-weight:700;
  color:#334862; padding:8px 2px; border-bottom:2px solid transparent; white-space:nowrap; }
.nav-link svg { width:16px; height:16px; }
.nav-link:hover { color:var(--blue); }
.nav-link.active { color:var(--blue); border-bottom-color:var(--blue); }
.login-btn { display:flex; align-items:center; gap:8px; font-size:13.5px; font-weight:700;
  color:var(--blue); border:1.5px solid #cfe0f5; background:#f3f8ff; padding:9px 18px;
  border-radius:999px; white-space:nowrap; }
.login-btn svg { width:16px; height:16px; }
.login-btn:hover { background:#e8f1fd; }
@media (max-width: 1180px) { .navbar { order:3; width:100%; justify-content:center; padding-top:10px;
  border-top:1px solid var(--line); } }

/* ---- Hero ---- */
.hero { position:relative; overflow:hidden; background:linear-gradient(135deg,#eaf2fb 0%,#cfe1f6 55%,#a9cdec 100%);
  padding:36px 32px 50px; }
.hero-image { padding:0; }
.hero-image .hero-inner { display:block; max-width:1600px; margin:0 auto; }
.hero-banner-img { width:100%; height:auto; display:block; }
.hero-inner { max-width:1600px; margin:0 auto; display:flex; align-items:center; gap:36px; flex-wrap:wrap; }

.hero-logo { flex:0 0 auto; }
.hero-logo img { height:150px; width:auto; display:block; }
@media (max-width: 1300px) { .hero-logo img { height:110px; } }
@media (max-width: 700px) { .hero-logo { width:100%; text-align:center; } .hero-logo img { height:90px; } }

.hero-copy { flex:1 1 380px; min-width:280px; border-left:3px solid #ffb020; padding-left:24px; }
.hero-copy h1 { font-size:30px; font-weight:800; color:var(--navy); margin:0 0 10px; letter-spacing:-.01em; }
.hero-copy p { font-size:14.5px; color:#33495e; max-width:560px; line-height:1.6; margin:0 0 22px; }
.feature-row { display:flex; gap:22px; flex-wrap:wrap; }
.feature { display:flex; align-items:center; gap:10px; }
.feature .ic { width:38px; height:38px; border-radius:10px; background:#fff; flex:none;
  display:flex; align-items:center; justify-content:center; box-shadow:0 2px 6px rgba(18,58,99,.12); }
.feature .ic svg { width:19px; height:19px; color:var(--blue); }
.feature span { font-size:12.5px; font-weight:700; color:var(--navy); line-height:1.25; display:block; max-width:110px; }

.hero-art { flex:0 0 auto; display:flex; justify-content:center; }
.illust-wrap { position:relative; width:300px; }
.illust-screen { background:linear-gradient(160deg,#1c4faa,#0d2a5c); border-radius:16px;
  aspect-ratio:4/3; display:flex; align-items:center; justify-content:center;
  box-shadow:0 20px 40px rgba(10,30,60,.28); position:relative; overflow:hidden; }
.illust-stand { width:110px; height:10px; background:#c3cad2; border-radius:0 0 8px 8px; margin:0 auto; }
.illust-ai-badge { position:absolute; top:-14px; right:-10px; z-index:2;
  background:linear-gradient(135deg,#ffc857,#ff8a00); color:#fff; font-weight:900; font-size:17px;
  width:46px; height:46px; border-radius:12px; display:flex; align-items:center; justify-content:center;
  box-shadow:0 8px 16px rgba(255,140,0,.35); transform:rotate(-6deg); }
.illust-mag { position:absolute; bottom:-10px; right:18px; width:44px; height:44px; border-radius:50%;
  background:#fff; display:flex; align-items:center; justify-content:center; box-shadow:0 6px 14px rgba(18,58,99,.25); }
.illust-mag svg { width:22px; height:22px; color:var(--blue); }
@media (max-width: 700px) { .illust-wrap { width:220px; } }

/* ---- Main grid ---- */
.main-wrap { max-width:1600px; margin:-24px auto 0; padding:0 40px 40px; position:relative; z-index:2; }
.grid { display:grid; grid-template-columns:1fr 1fr 340px; gap:26px; align-items:start; }
@media (max-width: 980px) { .grid { grid-template-columns:1fr; } }

.upload-card { background:#fff; border-radius:16px; padding:24px; box-shadow:0 6px 24px rgba(18,58,99,.08); }
.step-head { display:flex; align-items:center; gap:12px; margin-bottom:6px; }
.step-num { width:26px; height:26px; border-radius:50%; background:var(--blue); color:#fff;
  font-size:13px; font-weight:800; display:flex; align-items:center; justify-content:center; flex:none; }
.step-icon { width:44px; height:44px; border-radius:12px; background:#eaf1fc; flex:none;
  display:flex; align-items:center; justify-content:center; }
.step-icon svg { width:22px; height:22px; color:var(--blue); }
.step-title { font-size:15.5px; font-weight:800; color:var(--navy); }
.step-desc { font-size:12.5px; color:var(--muted); margin:2px 0 16px; }

.dropzone { border:2px dashed #b9cbe0; border-radius:12px; background:#f7fafd; padding:26px 14px;
  text-align:center; cursor:pointer; transition:.15s; }
.dropzone:hover, .dropzone.dragover { border-color:var(--blue); background:#eef5fd; }
.dropzone svg { width:30px; height:30px; color:var(--blue); margin-bottom:8px; }
.dropzone .dz-title { font-size:13.5px; font-weight:700; color:var(--navy); }
.dropzone .dz-sub { font-size:12px; color:var(--muted); margin-top:2px; }
.dropzone .dz-max { font-size:11px; color:var(--muted); margin-top:8px; }
.dropzone input[type=file] { display:none; }

.file-chip { display:none; margin-top:12px; align-items:center; justify-content:space-between;
  background:var(--green-bg); border:1px solid #cdeedb; border-radius:10px; padding:9px 12px; }
.file-chip.show { display:flex; }
.file-chip .fc-left { display:flex; align-items:center; gap:8px; min-width:0; }
.file-chip .fc-left svg { width:16px; height:16px; color:var(--green); flex:none; }
.file-chip .fc-name { font-size:12.5px; font-weight:600; color:#1c2b3a; overflow:hidden;
  text-overflow:ellipsis; white-space:nowrap; max-width:160px; }
.file-chip .fc-size { font-size:11.5px; color:var(--muted); flex:none; margin-left:8px; }
.file-chip .fc-remove { background:none; border:none; cursor:pointer; color:var(--muted);
  font-size:15px; line-height:1; padding:2px 4px; }

.gen-btn { grid-column:1 / span 2; }
@media (max-width: 980px) { .gen-btn { grid-column:1; } }
.gen-btn button { width:100%; background:linear-gradient(90deg,var(--blue2),var(--navy)); color:#fff;
  border:none; padding:16px; border-radius:12px; font-size:15.5px; font-weight:800; cursor:pointer;
  display:flex; align-items:center; justify-content:center; gap:8px; box-shadow:0 6px 18px rgba(30,99,199,.3); }
.gen-btn button svg { width:20px; height:20px; flex:none; }
.gen-btn button:hover { filter:brightness(1.05); }
.gen-btn button:disabled { opacity:.55; cursor:not-allowed; }
.gen-btn .gen-note { text-align:center; font-size:11.5px; color:var(--muted); margin-top:10px; }
.gen-btn .spinner { display:none; text-align:center; font-size:12.5px; color:var(--blue); margin-top:10px; font-weight:700; }

.flow-card { background:#fff; border-radius:16px; padding:22px 20px; box-shadow:0 6px 24px rgba(18,58,99,.08); }
.flow-card h3 { font-size:14px; font-weight:800; color:var(--blue); margin:0 0 16px; }
.flow-step { display:flex; gap:12px; position:relative; padding-bottom:22px; }
.flow-step:last-child { padding-bottom:0; }
.flow-step::before { content:""; position:absolute; left:15px; top:34px; bottom:0; width:2px;
  background:repeating-linear-gradient(to bottom, #c9d6e6 0 4px, transparent 4px 8px); }
.flow-step:last-child::before { display:none; }
.flow-dot { width:32px; height:32px; border-radius:50%; flex:none; display:flex; align-items:center;
  justify-content:center; font-weight:800; font-size:13px; color:#fff; z-index:1; }
.flow-dot svg { width:16px; height:16px; }
.flow-step:nth-of-type(1) .flow-dot { background:var(--blue); }
.flow-step:nth-of-type(2) .flow-dot { background:#3fa7d6; }
.flow-step:nth-of-type(3) .flow-dot { background:#7e5bd6; }
.flow-step:nth-of-type(4) .flow-dot { background:linear-gradient(135deg,#ffc857,#ff8a00); }
.flow-body .ft { font-size:13px; font-weight:800; color:var(--navy); }
.flow-body .fd { font-size:11.5px; color:var(--muted); margin-top:2px; line-height:1.4; }

.error-banner { grid-column:1/-1; background:#fff3cd; border:1px solid #ffe08a; padding:12px 16px;
  border-radius:10px; font-size:13px; color:#7a5b00; margin-bottom:2px; }

/* ---- Trust strip ---- */
.trust-strip { max-width:1600px; margin:0 auto 40px; padding:0 40px; }
.trust-inner { background:#fff; border-radius:16px; padding:22px 26px; box-shadow:0 6px 24px rgba(18,58,99,.06);
  display:flex; align-items:center; flex-wrap:wrap; gap:26px; justify-content:space-between; }
.trust-item { display:flex; align-items:center; gap:10px; min-width:170px; }
.trust-item .ic { width:38px; height:38px; border-radius:10px; background:#eaf1fc; flex:none;
  display:flex; align-items:center; justify-content:center; }
.trust-item .ic svg { width:18px; height:18px; color:var(--blue); }
.trust-item .tt { font-size:12.5px; font-weight:800; color:var(--navy); }
.trust-item .td { font-size:11px; color:var(--muted); }
.trust-brand { display:flex; align-items:center; gap:10px; }
.trust-brand .tt { font-size:14px; font-weight:800; color:var(--navy); }
.trust-brand .td { font-size:11px; color:var(--muted); max-width:220px; }
"""

ICONS = {
  "doc": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>',
  "wave": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12c2-3 4-3 6 0s4 3 6 0 4-3 6 0"/><path d="M2 18c2-3 4-3 6 0s4 3 6 0 4-3 6 0"/></svg>',
  "upload": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/></svg>',
  "bolt": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2 3 14h7l-1 8 10-12h-7z"/></svg>',
  "shield": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>',
  "shield-check": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/></svg>',
  "lock": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="10" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>',
  "check-circle": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="m22 4-10 10-3-3"/></svg>',
  "x": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18M6 6l12 12"/></svg>',
  "cloud": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z"/></svg>',
  "gear": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.6 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.6a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9c.14.35.4.65.73.83.3.17.65.26 1 .26H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',
  "download": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v12m0 0-4-4m4 4 4-4"/><path d="M4 20h16"/></svg>',
  "boat": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 18h20l-2 4H4z"/><path d="M4 18V9l8-6 8 6v9"/><path d="M12 3v15"/></svg>',
  "home": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M9 22V12h6v10"/></svg>',
  "chevron-down": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>',
  "chart-bar": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><rect x="7" y="12" width="3" height="6"/><rect x="12" y="8" width="3" height="10"/><rect x="17" y="5" width="3" height="13"/></svg>',
  "user": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"/></svg>',
  "search": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg>',
  "book": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>',
  "life-buoy": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/><path d="m4.9 4.9 4.24 4.24m5.72 5.72 4.24 4.24m0-14.2-4.24 4.24m-5.72 5.72L4.9 19.1"/></svg>',
}

HEADER_HTML = """
<header class="site-header">
  <div class="brand">
    <div class="gov-badge"><img src="/static/logo-kkp.png" alt="Kementerian Kelautan dan Perikanan"></div>
    <div class="header-divider"></div>
    <div class="gov-badge rect"><img src="/static/logo-djprl.png" alt="DJPRL"></div>
    <div class="header-divider"></div>
    <div class="gov-text">BALAI PENATAAN<br>RUANG LAUT (BPRL)<br>MAKASSAR</div>
  </div>

  <nav class="navbar">
    <a href="/" class="nav-link active">""" + ICONS["home"] + """ Beranda</a>
    <a href="#" class="nav-link" onclick="return false;">Layanan """ + ICONS["chevron-down"] + """</a>
    <a href="#" class="nav-link" onclick="return false;">Informasi """ + ICONS["chevron-down"] + """</a>
    <a href="#" class="nav-link" onclick="return false;">""" + ICONS["book"] + """ Panduan</a>
    <a href="#" class="nav-link" onclick="return false;">""" + ICONS["life-buoy"] + """ Bantuan</a>
    <a href="#" class="nav-link" onclick="return false;">""" + ICONS["chart-bar"] + """ Laporan</a>
  </nav>

  <a href="#" class="login-btn" onclick="return false;">""" + ICONS["user"] + """ Login</a>
</header>
"""

UPLOAD_HTML = """<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>e-GeRAI KKPRL &mdash; Generate &amp; Asistensi Dokumen KKPRL</title>
<meta name="description" content="Platform layanan digital terintegrasi untuk Konsultasi, Asistensi, Pendampingan, Informasi &amp; Generate Dokumen KKPRL secara cepat, tepat, efisien dan efektif.">
<meta property="og:title" content="e-GeRAI KKPRL &mdash; Generate &amp; Asistensi Dokumen KKPRL">
<meta property="og:description" content="Platform layanan digital terintegrasi untuk Konsultasi, Asistensi, Pendampingan, Informasi &amp; Generate Dokumen KKPRL secara cepat, tepat, efisien dan efektif.">
<meta property="og:type" content="website">
<style>""" + LANDING_CSS + """</style></head>
<body>
""" + HEADER_HTML + """

<section class="hero hero-image">
  <div class="hero-inner">
    <img src="/static/hero-banner.png" alt="Generate &amp; Asistensi Dokumen KKPRL" class="hero-banner-img">
  </div>
</section>

<div class="main-wrap">
  <form method="POST" action="/review" enctype="multipart/form-data" id="genForm">
    <div class="grid">
      {% if error %}<div class="error-banner">\u26A0 {{ error }}</div>{% endif %}

      <div class="upload-card">
        <div class="step-head">
          <div class="step-num">1</div>
          <div class="step-icon">""" + ICONS["doc"] + """</div>
        </div>
        <div class="step-title">Draft Proposal PKKPRL (PDF)</div>
        <div class="step-desc">Unggah file PDF proposal yang akan digabungkan. Belum punya file-nya?
        <a href="/proposal-manual" style="color:var(--blue); font-weight:700;">Isi Formulir di sini</a>.</div>
        <div class="dropzone" id="dz1">
          """ + ICONS["cloud"] + """
          <div class="dz-title">Drag &amp; Drop PDF di sini</div>
          <div class="dz-sub">atau klik untuk memilih file</div>
          <div class="dz-max">Maksimum 10 MB</div>
          <input type="file" name="proposal" id="proposal" accept="application/pdf" required>
        </div>
        <div class="file-chip" id="chip1">
          <div class="fc-left">""" + ICONS["check-circle"] + """<span class="fc-name" id="name1"></span></div>
          <span class="fc-size" id="size1"></span>
          <button type="button" class="fc-remove" data-target="1">""" + ICONS["x"] + """</button>
        </div>
      </div>

      <div class="upload-card">
        <div class="step-head">
          <div class="step-num">2</div>
          <div class="step-icon">""" + ICONS["wave"] + """</div>
        </div>
        <div class="step-title">Laporan Kondisi Eksisting / Hidro-Oseanografi (PDF)</div>
        <div class="step-desc">Unggah file PDF laporan hidro-oseanografi.</div>
        <div class="dropzone" id="dz2">
          """ + ICONS["cloud"] + """
          <div class="dz-title">Drag &amp; Drop PDF di sini</div>
          <div class="dz-sub">atau klik untuk memilih file</div>
          <div class="dz-max">Maksimum 10 MB</div>
          <input type="file" name="laporan" id="laporan" accept="application/pdf" required>
        </div>
        <div class="file-chip" id="chip2">
          <div class="fc-left">""" + ICONS["check-circle"] + """<span class="fc-name" id="name2"></span></div>
          <span class="fc-size" id="size2"></span>
          <button type="button" class="fc-remove" data-target="2">""" + ICONS["x"] + """</button>
        </div>
      </div>

      <div class="flow-card">
        <h3>Alur Proses</h3>
        <div class="flow-step">
          <div class="flow-dot">""" + ICONS["cloud"] + """</div>
          <div class="flow-body"><div class="ft">Upload Proposal</div><div class="fd">Unggah file PDF Proposal PKKPRL</div></div>
        </div>
        <div class="flow-step">
          <div class="flow-dot">""" + ICONS["wave"] + """</div>
          <div class="flow-body"><div class="ft">Upload Laporan</div><div class="fd">Unggah file PDF Laporan Hidro-Oseanografi</div></div>
        </div>
        <div class="flow-step">
          <div class="flow-dot">""" + ICONS["gear"] + """</div>
          <div class="flow-body"><div class="ft">Generate Dokumen</div><div class="fd">Sistem menggabungkan dokumen secara otomatis</div></div>
        </div>
        <div class="flow-step">
          <div class="flow-dot">""" + ICONS["download"] + """</div>
          <div class="flow-body"><div class="ft">Download Dokumen</div><div class="fd">Dokumen Word siap diunduh dan diedit</div></div>
        </div>
      </div>

      <div class="gen-btn">
        <button type="submit">""" + ICONS["bolt"] + """ Generate &amp; Gabungkan Dokumen Word</button>
        <div class="gen-note">Sistem akan memproses dan membuat dokumen Word final secara otomatis</div>
        <div class="spinner" id="spinner">\u23F3 Memproses dokumen, mohon tunggu...</div>
      </div>
    </div>
  </form>
</div>

<script>
function fmtSize(bytes) {
  if (bytes >= 1024*1024) return (bytes/(1024*1024)).toFixed(2) + " MB";
  return Math.ceil(bytes/1024) + " KB";
}
function setupDropzone(n) {
  var dz = document.getElementById('dz' + n);
  var input = dz.querySelector('input[type=file]');
  var chip = document.getElementById('chip' + n);
  var nameEl = document.getElementById('name' + n);
  var sizeEl = document.getElementById('size' + n);

  function showFile(file) {
    if (!file) { chip.classList.remove('show'); dz.style.display = 'block'; return; }
    nameEl.textContent = file.name;
    nameEl.title = file.name;
    sizeEl.textContent = fmtSize(file.size);
    chip.classList.add('show');
    dz.style.display = 'none';
  }

  dz.addEventListener('click', function(e) {
    if (e.target.closest('.file-chip')) return;
    input.click();
  });
  input.addEventListener('change', function() {
    if (input.files && input.files[0]) showFile(input.files[0]);
  });
  dz.addEventListener('dragover', function(e) { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', function() { dz.classList.remove('dragover'); });
  dz.addEventListener('drop', function(e) {
    e.preventDefault();
    dz.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      input.files = e.dataTransfer.files;
      showFile(input.files[0]);
    }
  });
  document.querySelector('.fc-remove[data-target="' + n + '"]').addEventListener('click', function() {
    input.value = '';
    dz.style.display = 'block';
    chip.classList.remove('show');
  });
}
setupDropzone(1);
setupDropzone(2);

document.getElementById('genForm').addEventListener('submit', function() {
  document.getElementById('spinner').style.display = 'block';
});
</script>
</body></html>"""


REVIEW_CSS = """
.review-hero { background:linear-gradient(135deg,#eaf2fb 0%,#cfe1f6 55%,#a9cdec 100%);
  padding:26px 32px 34px; }
.review-hero h1 { font-size:24px; font-weight:800; color:var(--navy); margin:0 0 6px; }
.review-hero p { font-size:13.5px; color:#33495e; margin:0; max-width:640px; line-height:1.5; }
.review-wrap { max-width:1600px; margin:-16px auto 40px; padding:0 40px; position:relative; z-index:2; }
.review-grid { display:grid; grid-template-columns:1.15fr .85fr; gap:22px; align-items:start; }
@media (max-width: 980px) { .review-grid { grid-template-columns:1fr; } }

.sticky-bar { position:sticky; top:0; z-index:6; background:var(--bg); padding:14px 0 10px; }
.sticky-bar button { width:100%; background:linear-gradient(90deg,var(--blue2),var(--navy)); color:#fff;
  border:none; padding:15px; border-radius:12px; font-size:14.5px; font-weight:800; cursor:pointer;
  display:flex; align-items:center; justify-content:center; gap:8px; box-shadow:0 6px 18px rgba(30,99,199,.28); }
.sticky-bar button svg { width:19px; height:19px; flex:none; }
.sticky-bar button:hover { filter:brightness(1.05); }
.back-link { display:inline-flex; align-items:center; gap:6px; font-size:12.5px; font-weight:700;
  color:var(--blue); margin-top:14px; }

.review-card { background:#fff; border-radius:16px; padding:22px 24px; box-shadow:0 6px 24px rgba(18,58,99,.08);
  margin-bottom:18px; }
.review-card h3 { font-size:14.5px; font-weight:800; color:var(--navy); margin:0 0 14px;
  display:flex; align-items:center; gap:8px; }
.review-card h3 svg { width:18px; height:18px; color:var(--blue); }

.acc-item { border:1px solid var(--line); border-radius:12px; margin-bottom:10px; overflow:hidden; }
.acc-item:last-child { margin-bottom:0; }
.acc-item summary { cursor:pointer; list-style:none; padding:13px 16px; font-size:13.5px; font-weight:700;
  color:var(--navy); background:#f7fafd; display:flex; align-items:center; justify-content:space-between; }
.acc-item summary::-webkit-details-marker { display:none; }
.acc-item summary::after { content:"\\25BE"; color:var(--blue); font-size:12px; transition:.15s; }
.acc-item[open] summary::after { transform:rotate(180deg); }
.acc-body { padding:14px 16px 4px; }
.field-row { margin-bottom:12px; }
.field-row label { display:block; font-size:12px; font-weight:700; color:var(--muted); margin-bottom:5px; }
.field-row input[type=text] { width:100%; padding:9px 11px; border:1px solid #d3dde7; border-radius:8px;
  font-size:13px; color:var(--ink); background:#fff; }
.field-row input[type=text]:focus { outline:none; border-color:var(--blue); box-shadow:0 0 0 3px rgba(47,127,224,.15); }

.preview-panel { position:sticky; top:80px; }
.preview-box { border:1px solid var(--line); border-radius:12px; background:#fff; padding:24px 28px;
  max-height:70vh; overflow-y:auto; font-size:13.5px; line-height:1.55; }
.preview-box table { border-collapse:collapse; width:100%; margin:10px 0; }
.preview-box td { border:1px solid #c7d1db; padding:5px 8px; vertical-align:top; }
.preview-box img { max-width:100%; height:auto; margin:8px 0; }
.preview-box p { margin:6px 0; }

.checkbox-row { display:flex; align-items:center; gap:10px; margin-bottom:12px; }
.checkbox-row input[type=checkbox] { width:18px; height:18px; accent-color:var(--blue); flex:none; }
.checkbox-row label { font-size:13px; font-weight:600; color:var(--ink); }
.manual-hero { background:linear-gradient(135deg,#eaf2fb 0%,#cfe1f6 55%,#a9cdec 100%); padding:26px 32px 34px; }
.manual-hero h1 { font-size:24px; font-weight:800; color:var(--navy); margin:0 0 6px; }
.manual-hero p { font-size:13.5px; color:#33495e; margin:0; max-width:640px; line-height:1.5; }
.manual-upload-card { background:#fff; border-radius:16px; padding:22px 24px; box-shadow:0 6px 24px rgba(18,58,99,.08); margin-bottom:18px; }

.file-field-row { margin-bottom:16px; }
.file-field-row label { display:block; font-size:12.5px; font-weight:700; color:var(--navy); margin-bottom:3px; }
.file-field-row .ff-hint { font-size:11px; color:var(--muted); margin-bottom:6px; }
.file-field-row input[type=file] { width:100%; font-size:12.5px; padding:8px; border:1px solid #d3dde7;
  border-radius:8px; background:#f7fafd; }
.field-row textarea { width:100%; padding:9px 11px; border:1px solid #d3dde7; border-radius:8px;
  font-size:12.5px; font-family:monospace; color:var(--ink); background:#fff; resize:vertical; min-height:80px; }
.field-row textarea:focus { outline:none; border-color:var(--blue); box-shadow:0 0 0 3px rgba(47,127,224,.15); }
.field-example { font-size:11px; color:var(--muted); margin-top:4px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.field-example .ex-text { font-style:italic; }
.field-example .ex-fill { font-size:10.5px; font-weight:700; color:var(--blue); background:#eaf1fc;
  border:1px solid #cfe0f5; border-radius:6px; padding:2px 8px; cursor:pointer; white-space:nowrap; }
.field-example .ex-fill:hover { background:#dcebfa; }

.img-paste-zone { border:2px dashed #b9cbe0; border-radius:10px; background:#f7fafd; padding:14px;
  cursor:pointer; transition:.15s; outline:none; }
.img-paste-zone:hover, .img-paste-zone:focus, .img-paste-zone.dragover { border-color:var(--blue); background:#eef5fd; }
.img-paste-placeholder { font-size:11.5px; color:var(--muted); text-align:center; }
.img-paste-placeholder svg { width:18px; height:18px; display:block; margin:0 auto 4px; color:var(--blue); }
.img-preview-list { display:flex; flex-wrap:wrap; gap:8px; margin-top:8px; }
.img-preview-list:empty { margin-top:0; }
.img-thumb { position:relative; width:64px; height:64px; border-radius:8px; overflow:hidden;
  border:1px solid var(--line); flex:none; }
.img-thumb img { width:100%; height:100%; object-fit:cover; display:block; }
.img-thumb .img-remove { position:absolute; top:2px; right:2px; width:18px; height:18px; border-radius:50%;
  background:rgba(0,0,0,.55); color:#fff; border:none; font-size:11px; line-height:1; cursor:pointer;
  display:flex; align-items:center; justify-content:center; }
"""


EXAMPLE_HINTS = {
    ("prop", "Nama Pemohon"): "Andi Wijaya, S.T., M.M.",
    ("prop", "Jabatan Pemohon"): "Direktur Utama",
    ("prop", "Nama Perusahaan/Instansi"): "PT. Bahari Sejahtera Makassar",
    ("prop", "NIB"): "1234567891234",
    ("prop", "NPWP"): "01.234.567.8-901.000",
    ("prop", "Nomor Telepon Selular"): "081234567890",
    ("prop", "Surat Elektronik"): "info@baharisejahteramks.co.id",
    ("prop", "Jenis Kegiatan"): "Pembangunan Dermaga dan Fasilitas Wisata Bahari",
    ("prop", "Nama Perairan"): "Selat Makassar",
    ("prop", "Luas Kebutuhan Ruang"): "5,2 Ha",
    ("prop", "KBLI"): "50121 - Angkutan Laut Wisata Dalam Negeri",
    ("prop", "Tanggal Penyusunan"): "02 Agustus 2026",
    ("prop_loc", "0"): "Desa Bontolebang",
    ("prop_loc", "1"): "Kecamatan Ujung Tanah",
    ("prop_loc", "2"): "Kabupaten Pangkajene dan Kepulauan",
    ("prop_loc", "3"): "Sulawesi Selatan",
    ("prop", "investasi"): "15000000000",
    ("prop", "tenaga_kerja"): "35",
    ("prop", "tenaga_kerja_asing"): "0",
    ("prop", "mangrove_spesies"): "Rhizophora mucronata",
    ("prop", "mangrove_persen"): "65",
    ("prop", "mangrove_kondisi"): "baik/rapat",
    ("prop", "desa_luas_ha"): "250",
    ("prop", "desa_penduduk"): "3400",
}


def render_manual_form_page(error=None):
    prop_groups = []
    for group_title, fields in FIELD_GROUPS:
        prop_fields = [f for f in fields if f[0] in ("prop", "prop_loc")]
        if prop_fields:
            prop_groups.append((group_title, prop_fields))

    groups_html = []
    for i, (group_title, fields) in enumerate(prop_groups):
        rows = []
        for source, key, label in fields:
            fname = form_field_name(source, key)
            example = EXAMPLE_HINTS.get((source, key), "")
            example_html = ""
            if example:
                example_html = (
                    f'<div class="field-example">Contoh: <span class="ex-text">{example}</span>'
                    f'<button type="button" class="ex-fill" data-target="{fname}">Pakai contoh ini</button></div>'
                )
            rows.append(
                f'<div class="field-row"><label>{label}</label>'
                f'<input type="text" name="{fname}" id="{fname}" placeholder="Isi {label.lower()}">'
                f'{example_html}</div>'
            )
        groups_html.append(
            f'<details class="acc-item"{" open" if i == 0 else ""}>'
            f"<summary>{group_title}</summary>"
            f'<div class="acc-body">{"".join(rows)}</div></details>'
        )

    error_html = f'<div class="error-banner" style="margin-bottom:16px;">\u26A0 {error}</div>' if error else ""

    return """<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Isi Formulir Draft Proposal &mdash; e-GeRAI KKPRL</title>
<style>""" + LANDING_CSS + REVIEW_CSS + """</style></head>
<body>
""" + HEADER_HTML + """

<section class="manual-hero">
  <h1>\U0001F4DD Isi Formulir Draft Proposal PKKPRL</h1>
  <p>Belum punya file Draft Proposal PKKPRL siap pakai? Isi data di bawah ini secara manual, lalu unggah
  Laporan Kondisi Eksisting / Hidro-Oseanografi (PDF). Sistem akan menggabungkan otomatis menjadi
  1 dokumen Word final &mdash; sama seperti alur upload 2 PDF.</p>
</section>

<div class="review-wrap">
  """ + error_html + """
  <form method="POST" action="/proposal-manual" enctype="multipart/form-data" id="manualForm">
    <div class="manual-upload-card">
      <h3>""" + ICONS["wave"] + """ Laporan Kondisi Eksisting / Hidro-Oseanografi (PDF)</h3>
      <div class="dropzone" id="dzManual">
        """ + ICONS["cloud"] + """
        <div class="dz-title">Drag &amp; Drop PDF di sini</div>
        <div class="dz-sub">atau klik untuk memilih file</div>
        <div class="dz-max">Maksimum 10 MB</div>
        <input type="file" name="laporan" id="laporanManual" accept="application/pdf" required>
      </div>
      <div class="file-chip" id="chipManual">
        <div class="fc-left">""" + ICONS["check-circle"] + """<span class="fc-name" id="nameManual"></span></div>
        <span class="fc-size" id="sizeManual"></span>
        <button type="button" class="fc-remove" data-target="Manual">""" + ICONS["x"] + """</button>
      </div>
    </div>

    <div class="sticky-bar">
      <button type="submit">""" + ICONS["bolt"] + """ Proses &amp; Lanjut ke Tinjau Data</button>
    </div>

    <div class="review-card">
      <h3>""" + ICONS["doc"] + """ Data Draft Proposal PKKPRL</h3>
      """ + "".join(groups_html) + """

      <details class="acc-item">
        <summary>Status Kegiatan</summary>
        <div class="acc-body">
          <div class="checkbox-row"><input type="checkbox" name="non_reklamasi" id="cb1"><label for="cb1">Kegiatan tanpa reklamasi</label></div>
          <div class="checkbox-row"><input type="checkbox" name="kegiatan_berusaha" id="cb2"><label for="cb2">Termasuk kegiatan berusaha</label></div>
          <div class="checkbox-row"><input type="checkbox" name="non_strategis" id="cb3"><label for="cb3">Termasuk kegiatan non-strategis nasional</label></div>
        </div>
      </details>

      <details class="acc-item">
        <summary>Titik Koordinat Batas Area (Opsional)</summary>
        <div class="acc-body">
          <div class="field-row">
            <label>Satu titik per baris, format: Nomor | Longitude | Latitude</label>
            <textarea name="koordinat_manual" rows="4" placeholder="1 | 106&deg;49&#39;30.5&quot; BT | 06&deg;54&#39;52.5&quot; LS&#10;2 | 106&deg;49&#39;35.2&quot; BT | 06&deg;54&#39;48.1&quot; LS"></textarea>
          </div>
        </div>
      </details>

      <details class="acc-item">
        <summary>Lampiran Gambar (Opsional)</summary>
        <div class="acc-body">
          <div class="ff-hint" style="margin-bottom:14px;">Klik kotak di bawah lalu tekan <b>Ctrl+V</b> untuk paste gambar dari clipboard (misal screenshot), atau klik untuk pilih file dari komputer.</div>
          <div class="file-field-row">
            <label>Peta Rencana Tapak (Site Plan)</label>
            <div class="img-paste-zone" tabindex="0" data-field="img_siteplan" data-multiple="false">
              <input type="file" name="img_siteplan" accept="image/*" style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
          <div class="file-field-row">
            <label>Peta Lokasi &amp; Sebaran Titik Koordinat</label>
            <div class="img-paste-zone" tabindex="0" data-field="img_peta_lokasi" data-multiple="false">
              <input type="file" name="img_peta_lokasi" accept="image/*" style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
          <div class="file-field-row">
            <label>Foto Kondisi Perairan &amp; Garis Pantai <span style="font-weight:400;color:var(--muted);">(bisa lebih dari 1)</span></label>
            <div class="img-paste-zone" tabindex="0" data-field="img_foto_pantai" data-multiple="true">
              <input type="file" name="img_foto_pantai" accept="image/*" multiple style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file (bisa banyak)</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
          <div class="file-field-row">
            <label>Foto Kondisi Mangrove</label>
            <div class="img-paste-zone" tabindex="0" data-field="img_foto_mangrove" data-multiple="false">
              <input type="file" name="img_foto_mangrove" accept="image/*" style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
          <div class="file-field-row">
            <label>Foto Survei Terumbu Karang</label>
            <div class="img-paste-zone" tabindex="0" data-field="img_foto_karang_insitu" data-multiple="false">
              <input type="file" name="img_foto_karang_insitu" accept="image/*" style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
          <div class="file-field-row">
            <label>Peta Rencana Pola Ruang Wilayah</label>
            <div class="img-paste-zone" tabindex="0" data-field="img_peta_pola_ruang" data-multiple="false">
              <input type="file" name="img_peta_pola_ruang" accept="image/*" style="display:none">
              <div class="img-paste-placeholder">""" + ICONS["upload"] + """Klik lalu Ctrl+V, atau klik untuk pilih file</div>
              <div class="img-preview-list"></div>
            </div>
          </div>
        </div>
      </details>
    </div>
  </form>

  <a href="/" class="back-link">&larr; Kembali ke halaman utama (unggah 2 PDF)</a>
</div>

<script>
function fmtSize(bytes) {
  if (bytes >= 1024*1024) return (bytes/(1024*1024)).toFixed(2) + " MB";
  return Math.ceil(bytes/1024) + " KB";
}
(function() {
  var dz = document.getElementById('dzManual');
  var input = document.getElementById('laporanManual');
  var chip = document.getElementById('chipManual');
  var nameEl = document.getElementById('nameManual');
  var sizeEl = document.getElementById('sizeManual');
  function showFile(file) {
    if (!file) { chip.classList.remove('show'); dz.style.display = 'block'; return; }
    nameEl.textContent = file.name;
    nameEl.title = file.name;
    sizeEl.textContent = fmtSize(file.size);
    chip.classList.add('show');
    dz.style.display = 'none';
  }
  dz.addEventListener('click', function(e) { if (!e.target.closest('.file-chip')) input.click(); });
  input.addEventListener('change', function() { if (input.files[0]) showFile(input.files[0]); });
  dz.addEventListener('dragover', function(e) { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', function() { dz.classList.remove('dragover'); });
  dz.addEventListener('drop', function(e) {
    e.preventDefault(); dz.classList.remove('dragover');
    if (e.dataTransfer.files[0]) { input.files = e.dataTransfer.files; showFile(input.files[0]); }
  });
  document.querySelector('.fc-remove[data-target="Manual"]').addEventListener('click', function() {
    input.value = ''; dz.style.display = 'block'; chip.classList.remove('show');
  });
})();

// Tombol "Pakai contoh ini" -> isi input teks dengan contoh, siap diedit
document.querySelectorAll('.ex-fill').forEach(function(btn) {
  btn.addEventListener('click', function() {
    var target = document.getElementById(btn.dataset.target);
    var exampleText = btn.previousElementSibling.textContent;
    if (target) { target.value = exampleText; target.focus(); }
  });
});

// Paste-zone gambar: klik untuk pilih file, Ctrl+V untuk paste dari clipboard, drag & drop, multi-file
document.querySelectorAll('.img-paste-zone').forEach(function(zone) {
  var input = zone.querySelector('input[type=file]');
  var previewList = zone.querySelector('.img-preview-list');
  var placeholder = zone.querySelector('.img-paste-placeholder');
  var multiple = zone.dataset.multiple === 'true';
  var files = [];

  function refreshInput() {
    var dt = new DataTransfer();
    files.forEach(function(f) { dt.items.add(f); });
    input.files = dt.files;
  }

  function renderPreviews() {
    previewList.innerHTML = '';
    files.forEach(function(file, idx) {
      var thumb = document.createElement('div');
      thumb.className = 'img-thumb';
      var img = document.createElement('img');
      img.src = URL.createObjectURL(file);
      var removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'img-remove';
      removeBtn.innerHTML = '&times;';
      removeBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        files.splice(idx, 1);
        refreshInput();
        renderPreviews();
      });
      thumb.appendChild(img);
      thumb.appendChild(removeBtn);
      previewList.appendChild(thumb);
    });
    placeholder.style.display = (files.length && !multiple) ? 'none' : 'block';
  }

  function addFile(file) {
    if (!file || file.type.indexOf('image/') !== 0) return;
    if (multiple) { files.push(file); } else { files = [file]; }
    refreshInput();
    renderPreviews();
  }

  zone.addEventListener('click', function(e) {
    if (e.target.closest('.img-remove')) return;
    input.click();
  });
  zone.addEventListener('paste', function(e) {
    var items = (e.clipboardData || window.clipboardData).items;
    for (var i = 0; i < items.length; i++) {
      if (items[i].type.indexOf('image/') === 0) { addFile(items[i].getAsFile()); }
    }
  });
  input.addEventListener('change', function() {
    for (var i = 0; i < input.files.length; i++) { addFile(input.files[i]); }
  });
  zone.addEventListener('dragover', function(e) { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', function() { zone.classList.remove('dragover'); });
  zone.addEventListener('drop', function(e) {
    e.preventDefault(); zone.classList.remove('dragover');
    for (var i = 0; i < e.dataTransfer.files.length; i++) { addFile(e.dataTransfer.files[i]); }
  });
});
</script>
</body></html>"""


def render_review_page(job_id, prop_data, lap_data, preview_html, error=None):
    groups_html = []
    for i, (group_title, fields) in enumerate(FIELD_GROUPS):
        rows = []
        for source, key, label in fields:
            value = get_value(source, key, prop_data, lap_data)
            fname = form_field_name(source, key)
            value_escaped = (value or "").replace('"', "&quot;")
            rows.append(
                f'<div class="field-row"><label>{label}</label>'
                f'<input type="text" name="{fname}" value="{value_escaped}"></div>'
            )
        groups_html.append(
            f'<details class="acc-item"{" open" if i == 0 else ""}>'
            f"<summary>{group_title}</summary>"
            f'<div class="acc-body">{"".join(rows)}</div></details>'
        )

    error_html = f'<div class="error-banner" style="margin-bottom:16px;">\u26A0 {error}</div>' if error else ""

    return """<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tinjau &amp; Koreksi Data &mdash; e-GeRAI KKPRL</title>
<style>""" + LANDING_CSS + REVIEW_CSS + """</style></head>
<body>
""" + HEADER_HTML + """

<section class="review-hero">
  <h1>\U0001F4DD Tinjau &amp; Koreksi Data</h1>
  <p>Periksa hasil ekstraksi di bawah, lalu bandingkan dengan pratinjau dokumen di sebelah kanan.
  Koreksi kolom yang salah, lalu klik "Generate Dokumen Final &amp; Unduh".</p>
</section>

<div class="review-wrap">
  """ + error_html + """
  <form method="POST" action="/finalize">
    <input type="hidden" name="job_id" value=\"""" + job_id + """\">
    <div class="review-grid">
      <div>
        <div class="sticky-bar">
          <button type="submit">""" + ICONS["check-circle"] + """ Generate Dokumen Final &amp; Unduh</button>
        </div>
        <div class="review-card">
          <h3>""" + ICONS["doc"] + """ Data Hasil Ekstraksi (bisa dikoreksi)</h3>
          """ + "".join(groups_html) + """
        </div>
      </div>

      <div class="preview-panel">
        <div class="review-card">
          <h3>""" + ICONS["doc"] + """ Pratinjau Dokumen Lengkap</h3>
          <div class="preview-box">""" + preview_html + """</div>
        </div>
      </div>
    </div>
  </form>

  <a href="/" class="back-link">&larr; Unggah ulang dokumen lain</a>
</div>
</body></html>"""


@app.route("/", methods=["GET"])
def index():
    return render_template_string(UPLOAD_HTML, error=None)


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}


@app.route("/review", methods=["POST"])
def review():
    proposal_file = request.files.get("proposal")
    laporan_file = request.files.get("laporan")

    if not proposal_file or not laporan_file or proposal_file.filename == "" or laporan_file.filename == "":
        return render_template_string(UPLOAD_HTML, error="Mohon unggah kedua file PDF (proposal & laporan)."), 400
    if not proposal_file.filename.lower().endswith(".pdf") or not laporan_file.filename.lower().endswith(".pdf"):
        return render_template_string(UPLOAD_HTML, error="Kedua file harus berformat PDF."), 400

    job_store.cleanup_old_jobs(JOBS_DIR)

    job_id = uuid.uuid4().hex[:12]
    tmp_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(tmp_dir, exist_ok=True)
    proposal_path = os.path.join(tmp_dir, "proposal.pdf")
    laporan_path = os.path.join(tmp_dir, "laporan.pdf")
    proposal_file.save(proposal_path)
    laporan_file.save(laporan_path)

    try:
        prop_data, prop_images = extract_proposal_with_fallback(proposal_path, log=lambda *_: None)
        lap_data, lap_images = extract_laporan_with_fallback(laporan_path, log=lambda *_: None)

        job_store.save_job(JOBS_DIR, job_id, prop_data, prop_images, lap_data, lap_images)

        preview_docx_path = os.path.join(tmp_dir, "preview.docx")
        build_document(prop_data, prop_images, lap_data, lap_images, preview_docx_path)
        with open(preview_docx_path, "rb") as f:
            preview_html = mammoth.convert_to_html(f).value
    except Exception:
        traceback.print_exc()
        shutil.rmtree(tmp_dir, ignore_errors=True)
        job_store.delete_job(JOBS_DIR, job_id)
        return render_template_string(
            UPLOAD_HTML,
            error="Terjadi kesalahan saat memproses dokumen. Pastikan kedua file adalah PDF yang valid.",
        ), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return render_review_page(job_id, prop_data, lap_data, preview_html)


@app.route("/proposal-manual", methods=["GET"])
def proposal_manual_form():
    return render_manual_form_page()


@app.route("/proposal-manual", methods=["POST"])
def proposal_manual_submit():
    laporan_file = request.files.get("laporan")
    if not laporan_file or laporan_file.filename == "":
        return render_manual_form_page(error="Mohon unggah file PDF Laporan Hidro-Oseanografi."), 400
    if not laporan_file.filename.lower().endswith(".pdf"):
        return render_manual_form_page(error="File Laporan harus berformat PDF."), 400

    job_store.cleanup_old_jobs(JOBS_DIR)

    job_id = uuid.uuid4().hex[:12]
    tmp_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(tmp_dir, exist_ok=True)
    laporan_path = os.path.join(tmp_dir, "laporan.pdf")
    laporan_file.save(laporan_path)

    try:
        prop_data = {}
        prop_data, _ = apply_form_values(request.form, prop_data, {})
        prop_data["non_reklamasi"] = "non_reklamasi" in request.form
        prop_data["kegiatan_berusaha"] = "kegiatan_berusaha" in request.form
        prop_data["non_strategis"] = "non_strategis" in request.form

        koordinat = []
        for line in request.form.get("koordinat_manual", "").splitlines():
            parts = [p.strip() for p in line.split("|")]
            if len(parts) == 3 and any(parts):
                koordinat.append(parts)
        prop_data["koordinat"] = koordinat

        def file_ext(filename):
            ext = os.path.splitext(filename)[1].lstrip(".").lower()
            return ext if ext in ("jpg", "jpeg", "png", "webp", "gif", "bmp") else "jpg"

        prop_images = []
        single_image_fields = [
            ("img_siteplan", "siteplan"),
            ("img_peta_lokasi", "peta_lokasi"),
            ("img_foto_mangrove", "foto_mangrove"),
            ("img_foto_karang_insitu", "foto_karang_insitu"),
            ("img_peta_pola_ruang", "peta_pola_ruang"),
        ]
        for field_name, tag in single_image_fields:
            f = request.files.get(field_name)
            if f and f.filename:
                prop_images.append({"tag": tag, "bytes": f.read(), "ext": file_ext(f.filename)})
        for f in request.files.getlist("img_foto_pantai"):
            if f and f.filename:
                prop_images.append({"tag": "foto_pantai", "bytes": f.read(), "ext": file_ext(f.filename)})

        lap_data, lap_images = extract_laporan_with_fallback(laporan_path, log=lambda *_: None)

        job_store.save_job(JOBS_DIR, job_id, prop_data, prop_images, lap_data, lap_images)

        preview_docx_path = os.path.join(tmp_dir, "preview.docx")
        build_document(prop_data, prop_images, lap_data, lap_images, preview_docx_path)
        with open(preview_docx_path, "rb") as f:
            preview_html = mammoth.convert_to_html(f).value
    except Exception:
        traceback.print_exc()
        shutil.rmtree(tmp_dir, ignore_errors=True)
        job_store.delete_job(JOBS_DIR, job_id)
        return render_manual_form_page(error="Terjadi kesalahan saat memproses. Pastikan file Laporan adalah PDF yang valid."), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return render_review_page(job_id, prop_data, lap_data, preview_html)


@app.route("/finalize", methods=["POST"])
def finalize():
    job_id = request.form.get("job_id", "")
    loaded = job_store.load_job(JOBS_DIR, job_id)
    if not loaded:
        return render_template_string(
            UPLOAD_HTML,
            error="Sesi review sudah kedaluwarsa atau tidak ditemukan. Mohon unggah ulang dokumennya.",
        ), 400

    prop_data, prop_images, lap_data, lap_images = loaded
    prop_data, lap_data = apply_form_values(request.form, prop_data, lap_data)

    output_path = os.path.join(OUTPUT_DIR, f"Proposal_Final_{job_id}.docx")
    try:
        build_document(prop_data, prop_images, lap_data, lap_images, output_path)
    except Exception:
        traceback.print_exc()
        return render_template_string(
            UPLOAD_HTML, error="Terjadi kesalahan saat membuat dokumen final. Silakan coba lagi."
        ), 500
    finally:
        job_store.delete_job(JOBS_DIR, job_id)

    @after_this_request
    def cleanup(response):
        try:
            os.remove(output_path)
        except OSError:
            pass
        return response

    perusahaan = prop_data.get("Nama Perusahaan/Instansi", "PKKPRL").replace(" ", "_").replace(".", "")
    download_name = f"Proposal_Teknis_PKKPRL_{perusahaan}.docx"

    return send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, threaded=True, debug=False)
