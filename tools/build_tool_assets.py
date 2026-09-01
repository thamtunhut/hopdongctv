# -*- coding: utf-8 -*-
"""
Sinh anh nen (PNG, base64) cho tung trang cua template hop dong hien tai, cong voi
font times.ttf (base64), dung cho coord-editor.html - mini tool canh chinh toa do
bang tay. Chay lai script nay moi khi doi file template PDF.
"""
import base64
import json
import os
import fitz  # PyMuPDF

HERE = os.path.dirname(__file__)
WEBAPP_DIR = os.path.dirname(HERE)

TEMPLATE_PDF = os.path.join(WEBAPP_DIR, "assets", "Template_CTV_new.pdf")
FONT_FILE = os.path.join(WEBAPP_DIR, "assets", "times.ttf")

SCALE = 1.6  # px per pt trong anh nen (dung khop voi coord-editor.js)


def main():
    doc = fitz.open(TEMPLATE_PDF)
    pages = []
    page_sizes = []
    matrix = fitz.Matrix(SCALE, SCALE)
    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        png_b64 = base64.b64encode(pix.tobytes("png")).decode("ascii")
        pages.append(f"data:image/png;base64,{png_b64}")
        page_sizes.append({"width": page.rect.width, "height": page.rect.height})

    with open(FONT_FILE, "rb") as f:
        font_b64 = base64.b64encode(f.read()).decode("ascii")

    data = {
        "scale": SCALE,
        "pages": pages,
        "pageSizes": page_sizes,
        "fontBase64": font_b64,
    }

    out_path = os.path.join(HERE, "js", "toolAssets.js")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("// Tu dong sinh boi build_tool_assets.py - KHONG sua tay.\n")
        f.write("window.TOOL_ASSETS = ")
        json.dump(data, f)
        f.write(";\n")

    print(f"OK: {out_path} ({len(pages)} trang, scale={SCALE})")


if __name__ == "__main__":
    main()
