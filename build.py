"""Đóng gói mã nguồn + file mẫu thành file HTML chạy offline.

Chạy:  python build.py
- index.html          : bản công khai, chỉ nhúng mẫu và lịch DỮ LIỆU GIẢ trong mau/.
- ho-so-khao-sat.html : bản dùng trên máy, nhúng mẫu thật trong mau-that/ và lịch thật.
                        Chỉ tạo khi có các file đó; cả hai đều nằm trong .gitignore.
"""
import base64
import json
from pathlib import Path

WEB = Path(__file__).resolve().parent


def assets(folder, sample):
    return {
        "nk": folder / "nhat-ky.docx",
        "bb": folder / "bien-ban-nghiem-thu.docx",
        "sample": sample,
    }


PUBLIC = assets(WEB / "mau", WEB / "mau" / "lich-mau.xlsx")
LOCAL = assets(WEB / "mau-that", WEB.parent / "5. NHAT KY KS.xlsx")


def build(files, name):
    data = {k: base64.b64encode(p.read_bytes()).decode("ascii") for k, p in files.items()}
    src = WEB / "src"
    html = (src / "index.html").read_text(encoding="utf-8")
    parts = {
        "/*__ASSETS__*/": "const ASSETS = " + json.dumps(data) + ";",
        "/*__CORE__*/": "\n".join((src / n).read_text(encoding="utf-8") for n in ("core.js", "info.js")),
        "/*__APP__*/": (src / "app.js").read_text(encoding="utf-8"),
    }
    for marker, code in parts.items():
        html = html.replace(marker, code, 1)
    out = WEB / name
    out.write_text(html, encoding="utf-8")
    print(f"OK: {out} ({out.stat().st_size // 1024} KB)")


def main():
    build(PUBLIC, "index.html")
    if all(p.exists() for p in LOCAL.values()):
        build(LOCAL, "ho-so-khao-sat.html")


if __name__ == "__main__":
    main()
