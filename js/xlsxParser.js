// Doc file data.xlsx (client-side) va tra ve mang object theo dung ten cot goc.
(function (global) {
  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function formatDate(d) {
    return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
  }

  // Excel cell co the la Date (voi cellDates:true) -> chuan hoa thanh chuoi dd/mm/yyyy.
  // Cac cot da la chuoi dd/mm/yyyy san (vi du 'Ngay cap CCCD') khong bi dong nay.
  function normalizeValue(v) {
    if (v instanceof Date) return formatDate(v);
    if (v === null || v === undefined) return "";
    return v;
  }

  async function parseExcelFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

    const rows = rawRows
      .map((r) => {
        const out = {};
        for (const key of Object.keys(r)) {
          if (!key || key.startsWith("__EMPTY")) continue;
          out[key] = normalizeValue(r[key]);
        }
        return out;
      })
      // Bo qua nhung dong hoan toan trong (sheet co dinh dang keo dai qua het du lieu that).
      .filter((row) => Object.values(row).some((v) => String(v).trim() !== ""));

    return { sheetName, rows };
  }

  global.ExcelParser = { parseExcelFile };
})(window);
