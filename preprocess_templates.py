# -*- coding: utf-8 -*-
"""
Tien xu ly cac file .docx (Mau_02.docx, BBNT_1..4_item.docx):
- Chuyen Word field code MERGEFIELD (fldChar begin/instrText/separate/w:t/fldChar end)
  thanh 1 run van ban dang {TenTruong} de docxtemplater doc duoc.
- Go bo <w:mailMerge> trong word/settings.xml - day la thu pham gay ra hop thoai canh
  bao "Opening this document will run the following SQL command..." khi mo file bang
  Word, vi template goc tung duoc ket noi that voi data.xlsx qua Mail Merge cua Word.
- Rieng Mau_02.docx:
  + Bao loop {#rows} / {/rows} quanh dong bang chua field nguoi ban de lap lai theo
    tung dong du lieu Excel.
  + Xoa 9 dong trong con sot lai cua bang (thiet ke cu cho toi da 10 nguoi/bang, moi
    dong chi co san field Ngay_NT va cac o con lai de trong tay) - neu khong xoa, cac
    dong nay se hien chu "undefined" vi field Ngay_NT o do nam NGOAI vong lap va
    khong duoc cap du lieu.
  + Chen tag {Tong_M2} / {Tong_M2_chu} vao 2 cho trong (o dang nhieu khoang trang lien
    tiep) tren dong "Tong gia tri hang hoa, dich vu mua vao: ... (So tien bang chu ...)"
    de tu dong dien tong tien + so tien bang chu thay vi de trong nhu truoc.

Chay 1 lan, ghi ket qua vao webapp/templates/*.docx. Khong sua file goc.
"""
import zipfile
import shutil
import re
import os

SRC_DIR = os.path.join(os.path.dirname(__file__), "source_docx")
OUT_DIR = os.path.join(os.path.dirname(__file__), "templates")

# Mau_02.docx nam san trong thu muc goc cua repo (ban goc, chua qua Word mail-merge)
ROOT_DIR = os.path.dirname(os.path.dirname(__file__))

FILES = {
    "Mau_02.docx": os.path.join(ROOT_DIR, "Mau_02.docx"),
    "BBNT_1_item.docx": os.path.join(SRC_DIR, "BBNT_1_item.docx"),
    "BBNT_2_item.docx": os.path.join(SRC_DIR, "BBNT_2_item.docx"),
    "BBNT_3_item.docx": os.path.join(SRC_DIR, "BBNT_3_item.docx"),
    "BBNT_4_item.docx": os.path.join(SRC_DIR, "BBNT_4_item.docx"),
}

RUN_RE = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.DOTALL)
RPR_RE = re.compile(r"<w:rPr>.*?</w:rPr>", re.DOTALL)
FLDCHAR_RE = re.compile(r'<w:fldChar\s+w:fldCharType="(begin|separate|end)"\s*/>')
INSTRTEXT_RE = re.compile(r"<w:instrText[^>]*>(.*?)</w:instrText>", re.DOTALL)
MERGEFIELD_RE = re.compile(r'MERGEFIELD\s+"?([^"\s]+)"?')


def convert_fields_to_tags(xml: str) -> str:
    """Tim tung cum run tao thanh 1 field (begin..end) va thay bang 1 run {Tag}."""
    runs = list(RUN_RE.finditer(xml))
    out = []
    last_end = 0
    i = 0
    while i < len(runs):
        run_text = runs[i].group(0)
        fld = FLDCHAR_RE.search(run_text)
        if fld and fld.group(1) == "begin":
            # gom instrText tu run nay den truoc fldChar end
            group_start = runs[i].start()
            rpr_match = RPR_RE.search(run_text)
            rpr = rpr_match.group(0) if rpr_match else ""
            instr_parts = []
            j = i
            group_end = None
            while j < len(runs):
                rj = runs[j].group(0)
                m_instr = INSTRTEXT_RE.search(rj)
                if m_instr:
                    instr_parts.append(m_instr.group(1))
                fld_j = FLDCHAR_RE.search(rj)
                if fld_j and fld_j.group(1) == "end":
                    group_end = runs[j].end()
                    break
                j += 1
            if group_end is None:
                # khong tim thay end, bo qua an toan
                i += 1
                continue
            field_code = "".join(instr_parts)
            m = MERGEFIELD_RE.search(field_code)
            tag = m.group(1) if m else None

            out.append(xml[last_end:group_start])
            if tag:
                out.append(
                    "<w:r>{rpr}<w:t xml:space=\"preserve\">{{{tag}}}</w:t></w:r>".format(
                        rpr=rpr, tag=tag
                    )
                )
            # neu khong nhan dien duoc tag thi bo hang field nay (hiem khi xay ra)
            last_end = group_end
            i = j + 1
            continue
        i += 1
    out.append(xml[last_end:])
    return "".join(out)


def wrap_seller_row_with_loop(xml: str) -> str:
    """Tim dong <w:tr> chua field 'Họ_và_tên' (sau khi da convert sang {Họ_và_tên})
    va bao {#rows} o dau o dau tien, {/rows} o cuoi o cuoi cung trong dong do."""
    anchor = "{Họ_và_tên}"
    idx = xml.find(anchor)
    if idx == -1:
        raise RuntimeError("Khong tim thay {Họ_và_tên} sau khi convert field - kiem tra lai template")

    tr_start = xml.rfind("<w:tr ", 0, idx)
    if tr_start == -1:
        tr_start = xml.rfind("<w:tr>", 0, idx)
    tr_end = xml.find("</w:tr>", idx)
    if tr_start == -1 or tr_end == -1:
        raise RuntimeError("Khong xac dinh duoc bien <w:tr> chua dong nguoi ban")
    tr_end += len("</w:tr>")

    row_xml = xml[tr_start:tr_end]

    # Chen {#rows} vao dau noi dung text cua o dau tien (truoc <w:t> dau tien trong hang)
    # Luu y: phai phan biet <w:t>/<w:t xml:space=...> voi <w:tr>/<w:tc>/<w:tbl> (deu bat dau bang "w:t")
    WT_RE = r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)'
    first_t = re.search(WT_RE, row_xml, re.DOTALL)
    if not first_t:
        raise RuntimeError("Khong tim thay <w:t> nao trong dong nguoi ban de chen {#rows}")
    new_row = (
        row_xml[: first_t.start(2)]
        + "{#rows}"
        + first_t.group(2)
        + row_xml[first_t.end(2) : ]
    )

    # Chen {/rows} vao cuoi noi dung text cua o cuoi cung trong hang (sau khi da chen {#rows})
    last_t = None
    for last_t in re.finditer(WT_RE, new_row, re.DOTALL):
        pass
    if not last_t:
        raise RuntimeError("Khong tim thay <w:t> cuoi cung trong dong nguoi ban de chen {/rows}")
    new_row = (
        new_row[: last_t.start(2)]
        + last_t.group(2)
        + "{/rows}"
        + new_row[last_t.end(2) : ]
    )

    return xml[:tr_start] + new_row + xml[tr_end:]


MAILMERGE_RE = re.compile(r"<w:mailMerge>.*?</w:mailMerge>", re.DOTALL)


def strip_mail_merge(settings_xml: str) -> str:
    """Xoa ket noi Mail Merge (OLE DB toi file Excel goc) khoi settings.xml de Word
    khong con hoi 'run SQL command' khi mo file nua."""
    return MAILMERGE_RE.sub("", settings_xml)


def remove_dead_rows(xml: str) -> str:
    """Xoa 9 <w:tr> con sot lai giua dong da wrap loop va cuoi bang (thiet ke cu
    cho toi da 10 nguoi/bang, chi con field Ngay_NT don doc, khong con dung nua)."""
    rows_close = xml.find("{/rows}")
    if rows_close == -1:
        raise RuntimeError("Khong tim thay {/rows} - can goi sau wrap_seller_row_with_loop")

    tr_end_after_loop = xml.find("</w:tr>", rows_close)
    if tr_end_after_loop == -1:
        raise RuntimeError("Khong tim thay </w:tr> ket thuc dong da wrap loop")
    tr_end_after_loop += len("</w:tr>")

    tbl_close = xml.find("</w:tbl>", tr_end_after_loop)
    if tbl_close == -1:
        raise RuntimeError("Khong tim thay </w:tbl> sau dong da wrap loop")

    between = xml[tr_end_after_loop:tbl_close]
    if "<w:tbl" in between:
        raise RuntimeError("Vung du kien xoa co chua bang long - dung lai de kiem tra bang tay")

    return xml[:tr_end_after_loop] + xml[tbl_close:]


BLANK_RUN_RE = re.compile(r"(<w:r[ >].*?</w:r>)", re.DOTALL)
BLANK_T_RE = re.compile(r"(<w:t[^>]*>)(\s{8,})(</w:t>)")


def insert_total_tags(xml: str) -> str:
    """Tim 2 o trong (nhieu khoang trang lien tiep) tren dong 'Tong gia tri hang hoa,
    dich vu mua vao: ___ (So tien bang chu ___)' - la phan van ban tinh (khong phai
    MERGEFIELD) nen khong duoc convert_fields_to_tags xu ly - va chen tag moi vao."""
    anchor = xml.find("{/rows}")
    if anchor == -1:
        raise RuntimeError("Khong tim thay {/rows} - can goi sau wrap_seller_row_with_loop")

    matches = []
    for run_m in BLANK_RUN_RE.finditer(xml, anchor):
        t_m = BLANK_T_RE.search(run_m.group(1))
        if t_m:
            # offset tuyet doi cua noi dung khoang trang trong toan bo xml
            abs_start = run_m.start(1) + t_m.start(2)
            abs_end = run_m.start(1) + t_m.end(2)
            matches.append((abs_start, abs_end))

    if len(matches) < 2:
        raise RuntimeError(
            f"Khong tim du 2 cho trong de dien Tong_M2/Tong_M2_chu (tim thay {len(matches)}) - "
            "kiem tra lai template Mau_02.docx"
        )

    (s1, e1), (s2, e2) = matches[0], matches[1]
    # Thay tu cuoi len dau de khong lam lech offset cua vi tri con lai.
    xml = xml[:s2] + "{Tong_M2_chu}" + xml[e2:]
    xml = xml[:s1] + "{Tong_M2}" + xml[e1:]
    return xml


def process_docx(src_path: str, dst_path: str, add_loop: bool):
    with zipfile.ZipFile(src_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        settings_xml = zin.read("word/settings.xml").decode("utf-8")

    doc_xml = convert_fields_to_tags(doc_xml)

    if add_loop:
        doc_xml = wrap_seller_row_with_loop(doc_xml)
        doc_xml = remove_dead_rows(doc_xml)
        doc_xml = insert_total_tags(doc_xml)

    remaining = MERGEFIELD_RE.search(doc_xml)
    if remaining:
        raise RuntimeError(f"Van con MERGEFIELD chua convert het trong {src_path}: {remaining.group(0)}")

    settings_xml = strip_mail_merge(settings_xml)

    with zipfile.ZipFile(src_path, "r") as zin, zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc_xml.encode("utf-8")
            elif item.filename == "word/settings.xml":
                data = settings_xml.encode("utf-8")
            zout.writestr(item, data)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for fname, src in FILES.items():
        dst = os.path.join(OUT_DIR, fname)
        add_loop = fname == "Mau_02.docx"
        process_docx(src, dst, add_loop)
        print(f"OK: {fname} -> {dst}")


if __name__ == "__main__":
    main()
