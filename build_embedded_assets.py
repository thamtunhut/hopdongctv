# -*- coding: utf-8 -*-
"""
Nhung toan bo file nhi phan (PDF template, font, docx template) thanh base64
ngay trong 1 file JS (js/embeddedAssets.js), de index.html chay duoc bang
cach double-click mo truc tiep (file://) ma khong can static server -
vi trinh duyet chan fetch() doc file cuc bo duoi file://, nhung <script src>
nap file JS thi khong bi chan.

Chay lai script nay moi khi thay doi file trong assets/ hoac templates/.
"""
import base64
import json
import os

HERE = os.path.dirname(__file__)

FILES = {
    "Template_CTV_new.pdf": os.path.join(HERE, "assets", "Template_CTV_new.pdf"),
    "times.ttf": os.path.join(HERE, "assets", "times.ttf"),
    "Mau_02.docx": os.path.join(HERE, "templates", "Mau_02.docx"),
    "BBNT_1_item.docx": os.path.join(HERE, "templates", "BBNT_1_item.docx"),
    "BBNT_2_item.docx": os.path.join(HERE, "templates", "BBNT_2_item.docx"),
    "BBNT_3_item.docx": os.path.join(HERE, "templates", "BBNT_3_item.docx"),
    "BBNT_4_item.docx": os.path.join(HERE, "templates", "BBNT_4_item.docx"),
}


def main():
    data = {}
    for name, path in FILES.items():
        with open(path, "rb") as f:
            data[name] = base64.b64encode(f.read()).decode("ascii")

    out_path = os.path.join(HERE, "js", "embeddedAssets.js")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("// Tu dong sinh boi build_embedded_assets.py - KHONG sua tay.\n")
        f.write("window.EMBEDDED_ASSETS_B64 = ")
        json.dump(data, f)
        f.write(";\n")

    total = sum(len(v) for v in data.values())
    print(f"OK: {out_path} ({total / 1024:.0f} KB base64, {len(data)} files)")


if __name__ == "__main__":
    main()
