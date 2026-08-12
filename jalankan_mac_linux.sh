# Jangan pernah commit file ini ke GitHub -- berisi API key rahasia
.env

# Folder sementara aplikasi (otomatis terisi & dibersihkan saat runtime)
uploads/*
outputs/*
drafts/*
jobs/*
!uploads/.gitkeep
!outputs/.gitkeep
!drafts/.gitkeep
!jobs/.gitkeep

# Python
__pycache__/
*.pyc
*.pyo

# Riwayat pemakaian & draft (data lokal, jangan ikut ter-commit)
history.jsonl
