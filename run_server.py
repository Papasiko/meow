"""Run the Organizational Evaluation app with Waitress on Windows/Linux."""
from __future__ import annotations

import os
from waitress import serve
from app import app

host = os.environ.get("HOST", "0.0.0.0")
port = int(os.environ.get("PORT", "5000"))
threads = int(os.environ.get("WAITRESS_THREADS", "8"))

print("=" * 64)
print("سامانه ارزیابی عملکرد — نسخه اصلاح‌شده 2026.10.30")
print(f"Listening address: {host}:{port}")
print("Build verification: http://127.0.0.1:%s/api/health" % port)
print("Local health check: http://127.0.0.1:%s/api/health" % port)
print("برای توقف سرور، در همین پنجره Ctrl+C بزنید.")
print("قبل از باز کردن دسترسی شبکه، رمز پیش‌فرض تمام کاربران را تغییر دهید.")
print("=" * 64)
serve(app, host=host, port=port, threads=threads)
