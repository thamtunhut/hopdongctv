// Ve de text len template hop dong PDF theo toa do co dinh.
// Toa do duoi day do tren "Template_CTV_new.pdf" (mau moi do Legal cung cap 09/2026,
// khac han mau cu template_co_chuky.pdf ve bo cuc/kich thuoc trang A4) - xac dinh bang
// cach do vi tri label/gach chan thuc te tren file (PyMuPDF), khong con dung lai toa do cu.
(function (global) {
  const { PDFDocument, rgb } = PDFLib;

  // Toa do da duoc canh chinh bang tay qua tools/coord-editor.html (09/2026).
  // Cac truong co the co noi dung dai (dia chi, chi nhanh, cong viec, dia diem, so tien
  // bang chu) dat type:"paragraph" de tu dong xuong dong; "width" la be rong khung wrap.
  // Ho_ten o trang ky ten (#3) dung align:"center" de tu can giua trong khung [x, x+width]
  // du ten dai ngan khac nhau, thay vi 1 toa do x co dinh chi dung cho 1 do dai ten cu the.
  const COORDINATES = {
    Ho_ten: [
      { page: 0, x: 222, y: 543 }, // "BEN CUNG UNG DICH VU :"
      { page: 1, x: 373, y: 563 }, // bang thong tin lien he, cot Ben Cung Ung Dich Vu
      { page: 3, x: 282, y: 289, width: 198, align: "center" }, // trang ky ten - can giua tren gach chan
    ],
    Ngay1: [{ page: 0, x: 358, y: 720 }],
    Thang1: [{ page: 0, x: 424, y: 720 }],
    Nam1: [{ page: 0, x: 472, y: 720 }],
    DD: [{ page: 0, x: 162, y: 517 }],
    MM: [{ page: 0, x: 195, y: 517 }],
    YYYY: [{ page: 0, x: 222, y: 517 }],
    Quoc_tich: [{ page: 0, x: 378, y: 517 }],
    CCCD: [{ page: 0, x: 157, y: 483 }],
    D1: [{ page: 0, x: 308, y: 483.5 }],
    M1: [{ page: 0, x: 329, y: 483 }],
    Y1: [{ page: 0, x: 348.5, y: 483.5 }],
    Noi_cap: [{ page: 0, type: "paragraph", x: 425, y: 482, width: 118, height: 34 }],
    Noi_o: [{ page: 0, type: "paragraph", x: 158, y: 449, width: 370, height: 34 }],
    MST: [{ page: 0, x: 158, y: 423 }],
    STK: [{ page: 0, x: 157, y: 396 }],
    Bank: [{ page: 0, x: 157, y: 369 }],
    Chi_nhanh: [{ page: 0, type: "paragraph", x: 392, y: 368, width: 136, height: 34 }],
    SDT: [
      { page: 0, x: 158, y: 343 },
      { page: 1, x: 389, y: 512 },
    ],
    Cong_viec: [{ page: 0, type: "paragraph", x: 95, y: 171, width: 426, height: 60 }],
    // Template moi chi co 1 dong "Thoi han thuc hien Dich vu:" (khong tach 2 o nhu mau cu)
    // nen ghep Time + Time2 thanh "tu ngay - den ngay" ngay trong generateContractPdf.
    Time_Range: [{ page: 1, x: 240, y: 719 }],
    Location: [{ page: 1, type: "paragraph", x: 235, y: 637, width: 293, height: 34 }],
    So_tien: [{ page: 1, x: 159, y: 447.5 }],
    Bang_chu: [{ page: 1, type: "paragraph", x: 150, y: 420, width: 375, height: 34 }],
    // 3 truong moi (09/2026): khop truc tiep ten cot Excel "San pham"/"PIC"/"Email PIC".
    "Sản phẩm": [{ page: 1, x: 255, y: 750 }], // muc 1.2 "... san pham tro choi (Game) cua VNG:"
    PIC: [{ page: 1, x: 153, y: 563 }], // bang thong tin lien he, cot VNG, "Ho va ten:"
    "Email PIC": [{ page: 1, x: 134, y: 537 }], // bang thong tin lien he, cot VNG, "Email:"
  };

  const FONT_SIZE = 12;
  const LEADING = FONT_SIZE + 2;

  function wrapText(text, font, size, maxWidth) {
    const words = String(text).split(/\s+/).filter(Boolean);
    const lines = [];
    let current = "";
    for (const word of words) {
      const trial = current ? `${current} ${word}` : word;
      if (font.widthOfTextAtSize(trial, size) > maxWidth && current) {
        lines.push(current);
        current = word;
      } else {
        current = trial;
      }
    }
    if (current) lines.push(current);
    return lines;
  }

  // reportlab Paragraph.drawOn neo (x,y) la GOC DUOI TRAI cua noi dung DA WRAP THUC TE
  // (khong phai cua khung "height" khai bao) - chieu cao khung trong COORDINATES chi la
  // gioi han de tinh wrap, khong dung de neo vi tri hay cat bot dong.
  function drawParagraph(page, text, font, props) {
    const lines = wrapText(text, font, FONT_SIZE, props.width);
    const actualHeight = lines.length * LEADING;
    let y = props.y + actualHeight - FONT_SIZE;
    for (const line of lines) {
      page.drawText(line, { x: props.x, y, size: FONT_SIZE, font, color: rgb(0, 0, 0) });
      y -= LEADING;
    }
  }

  // Luu y hieu nang: PDFDocument.load() tren cung 1 file template goi lap lai nhieu lan
  // (moi nguoi 1 lan) bi pdf-lib xu ly cham dan (co the toi vai giay/lan sau ~2-3 lan goi).
  // Vi vay chi load() file goc DUY NHAT 1 LAN roi dung copyPages() de nhan ban cho tung
  // nguoi - cach nay nhanh on dinh (vai ms/lan) thay vi load() lai tu dau moi lan.
  async function loadAssets() {
    const pdfBytes = getEmbeddedAsset("Template_CTV_new.pdf");
    const fontBytes = getEmbeddedAsset("times.ttf");
    const srcDoc = await PDFDocument.load(pdfBytes);
    return { srcDoc, fontBytes };
  }

  async function generateContractPdf(row, assets) {
    const pdfDoc = await PDFDocument.create();
    pdfDoc.registerFontkit(fontkit);
    const copiedPages = await pdfDoc.copyPages(assets.srcDoc, assets.srcDoc.getPageIndices());
    copiedPages.forEach((p) => pdfDoc.addPage(p));
    const font = await pdfDoc.embedFont(assets.fontBytes, { subset: true });
    const pages = pdfDoc.getPages();

    const effectiveRow = { ...row };
    const timeRange = [row.Time, row.Time2].map((v) => (v === undefined || v === null ? "" : String(v).trim())).filter(Boolean);
    if (timeRange.length) effectiveRow.Time_Range = timeRange.join(" - ");

    for (const [key, locations] of Object.entries(COORDINATES)) {
      const raw = effectiveRow[key];
      const value = raw === undefined || raw === null ? "" : String(raw).trim();
      if (!value) continue;

      for (const loc of locations) {
        const page = pages[loc.page];
        if (!page) continue;
        if (loc.type === "paragraph") {
          drawParagraph(page, value, font, loc);
        } else if (loc.align === "center") {
          const textWidth = font.widthOfTextAtSize(value, FONT_SIZE);
          const x = loc.x + Math.max(0, (loc.width - textWidth) / 2);
          page.drawText(value, { x, y: loc.y, size: FONT_SIZE, font, color: rgb(0, 0, 0) });
        } else {
          page.drawText(value, { x: loc.x, y: loc.y, size: FONT_SIZE, font, color: rgb(0, 0, 0) });
        }
      }
    }

    return pdfDoc.save();
  }

  global.PdfContract = { COORDINATES, loadAssets, generateContractPdf };
})(window);
