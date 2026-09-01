// Dien cac mau Word (BBNT_N_item.docx, Mau_02.docx) bang docxtemplater.
// Ten tag trong template (dat theo Word field code cu) khac ten cot trong data.xlsx
// o dau underscore/dau cach, nen can 1 bang anh xa tag -> cot Excel.
(function (global) {
  const FIELD_MAP = {
    Ho_ten: "Ho_ten",
    Họ_và_tên: "Ho_ten",
    CCCD: "CCCD",
    Số_CCCD: "CCCD",
    MST: "MST",
    STK: "STK",
    Bank: "Bank",
    Chi_nhanh: "Chi_nhanh",
    Số_điện_thoại: "SDT",
    Địa_chỉ: "Địa chỉ",
    Ngày_cấp_CCCD: "Ngày cấp CCCD",
    Nơi_cấp_CCCD: "Nơi cấp CCCD",
    Công_việc: "Cong_viec",
    Link_nghiệm_thu: "Link nghiệm thu",
    So_tien: "So_tien",
    Tong: "Tong",
    Tổng_tiền_bằng_chữ: "Bang_chu",
    Ngay1: "Ngay1",
    Thang1: "Thang1",
    Nam1: "Nam1",
    Ngay3: "Ngay3",
    Thang3: "Thang3",
    Nam3: "Nam3",
    Ngay_NT: "Ngay_NT",
    Items_1: "Items 1",
    Items_2: "Items 2",
    Items_3: "Items 3",
    Items_4: "Items 4",
    Don_gia_1: "Don_gia_1",
    Don_gia_2: "Don_gia_2",
    Don_gia_3: "Don_gia_3",
    Don_gia_4: "Don_gia_4",
    TTPIT_1: "TT+PIT_1",
    TTPIT_2: "TT+PIT_2",
    TTPIT_3: "TT+PIT_3",
    TTPIT_4: "TT+PIT_4",
  };

  const BBNT_TAGS = {
    1: ["Bank", "CCCD", "Chi_nhanh", "Công_việc", "Don_gia_1", "Ho_ten", "Items_1", "Link_nghiệm_thu", "MST", "Nam1", "Nam3", "Ngay1", "Ngay3", "Ngày_cấp_CCCD", "Nơi_cấp_CCCD", "STK", "So_tien", "TTPIT_1", "Thang1", "Thang3", "Tong", "Tổng_tiền_bằng_chữ", "Địa_chỉ"],
    2: ["Bank", "CCCD", "Chi_nhanh", "Công_việc", "Don_gia_1", "Don_gia_2", "Ho_ten", "Items_1", "Items_2", "Link_nghiệm_thu", "MST", "Nam1", "Nam3", "Ngay1", "Ngay3", "Ngày_cấp_CCCD", "Nơi_cấp_CCCD", "STK", "So_tien", "TTPIT_1", "TTPIT_2", "Thang1", "Thang3", "Tong", "Tổng_tiền_bằng_chữ", "Địa_chỉ"],
    3: ["Bank", "CCCD", "Chi_nhanh", "Công_việc", "Don_gia_1", "Don_gia_2", "Don_gia_3", "Ho_ten", "Items_1", "Items_2", "Items_3", "Link_nghiệm_thu", "MST", "Nam1", "Nam3", "Ngay1", "Ngay3", "Ngày_cấp_CCCD", "Nơi_cấp_CCCD", "STK", "So_tien", "TTPIT_1", "TTPIT_2", "TTPIT_3", "Thang1", "Thang3", "Tong", "Tổng_tiền_bằng_chữ", "Địa_chỉ"],
    4: ["Bank", "CCCD", "Chi_nhanh", "Công_việc", "Don_gia_1", "Don_gia_2", "Don_gia_3", "Don_gia_4", "Ho_ten", "Items_1", "Items_2", "Items_3", "Items_4", "Link_nghiệm_thu", "MST", "Nam1", "Nam3", "Ngay1", "Ngay3", "Ngày_cấp_CCCD", "Nơi_cấp_CCCD", "STK", "So_tien", "TTPIT_1", "TTPIT_2", "TTPIT_3", "TTPIT_4", "Thang1", "Thang3", "Tong", "Tổng_tiền_bằng_chữ", "Địa_chỉ"],
  };

  // Tag trong vong lap {#rows}...{/rows} (lap lai theo tung nguoi).
  const MAU02_ROW_TAGS = ["Công_việc", "Họ_và_tên", "Ngay_NT", "So_tien", "Số_CCCD", "Số_điện_thoại", "Địa_chỉ"];
  // Tag nam NGOAI vong lap (dong ngay lap bang ke o dau trang) - dung chung cho ca bang,
  // lay tu dong dau tien vi day la 1 gia tri cho toan bo Mau 02, khong phai theo tung nguoi.
  const MAU02_TOP_TAGS = ["Ngay3", "Thang3", "Nam3"];

  function buildTagData(row, tags) {
    const data = {};
    for (const tag of tags) {
      const col = FIELD_MAP[tag] || tag;
      const v = row[col];
      data[tag] = v === undefined || v === null ? "" : String(v);
    }
    return data;
  }

  function countItems(row) {
    let n = 0;
    for (let i = 1; i <= 4; i++) {
      const v = row[`Items ${i}`];
      if (v !== undefined && v !== null && String(v).trim() !== "") n = i;
      else break;
    }
    return n;
  }

  function loadTemplateBytes(name) {
    return getEmbeddedAsset(name);
  }

  function renderDocx(templateBytes, data) {
    const zip = new PizZip(templateBytes);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{", end: "}" },
    });
    doc.render(data);
    return doc.getZip().generate({ type: "arraybuffer" });
  }

  // Tra ve { bytes, itemCount } hoac null neu dong khong co hang muc cong viec nao.
  async function generateBbnt(row) {
    const n = countItems(row);
    if (n === 0) return null;
    const templateName = `BBNT_${n}_item.docx`;
    const bytes = await loadTemplateBytes(templateName);
    const data = buildTagData(row, BBNT_TAGS[n]);
    return { bytes: renderDocx(bytes, data), itemCount: n };
  }

  async function generateMau02(rows) {
    const bytes = await loadTemplateBytes("Mau_02.docx");
    const rowsData = rows.map((row) => buildTagData(row, MAU02_ROW_TAGS));

    // Ngay lap bang ke la 1 gia tri chung cho ca van ban - lay tu dong dau tien.
    const topData = rows.length ? buildTagData(rows[0], MAU02_TOP_TAGS) : {};

    // Tong_M2 / Tong_M2_chu la gia tri TINH TOAN (cong don gia cua tat ca nguoi trong
    // dot nay), khong co san trong Excel nen phai tu tinh, khac voi Bang_chu tung nguoi.
    const total = rows.reduce((sum, row) => sum + VnNumberWords.parseAmount(row["So_tien"]), 0);
    topData.Tong_M2 = VnNumberWords.formatAmount(total);
    topData.Tong_M2_chu = VnNumberWords.numberToVietnameseWords(total);

    return renderDocx(bytes, { rows: rowsData, ...topData });
  }

  global.DocxFill = { FIELD_MAP, countItems, generateBbnt, generateMau02 };
})(window);
