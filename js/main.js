(function () {
  const fileInput = document.getElementById("fileInput");
  const dropZone = document.getElementById("dropZone");
  const previewSection = document.getElementById("previewSection");
  const previewBody = document.getElementById("previewBody");
  const generateBtn = document.getElementById("generateBtn");
  const statusEl = document.getElementById("status");
  const logEl = document.getElementById("log");

  let currentRows = [];

  function log(msg) {
    const line = document.createElement("div");
    line.textContent = msg;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function safeName(row, index) {
    const name = (row["Ho_ten"] || row["Họ và tên"] || "").toString().trim();
    return (name || `Dong_${index + 2}`).replace(/\s+/g, "_").replace(/\//g, "-");
  }

  function renderPreview(rows) {
    previewBody.innerHTML = "";
    rows.forEach((row, i) => {
      const tr = document.createElement("tr");
      const name = (row["Ho_ten"] || "").toString().trim();
      const items = DocxFill.countItems(row);
      const soTien = row["So_tien"] || "";
      const warnings = [];
      if (!name) warnings.push("thiếu Họ tên");
      if (items === 0) warnings.push("không có hạng mục công việc (bỏ qua BBNT)");
      if (!row["CCCD"]) warnings.push("thiếu CCCD");

      tr.innerHTML = `
        <td>${i + 2}</td>
        <td>${name || "(trống)"}</td>
        <td>${items}</td>
        <td>${soTien}</td>
        <td class="${warnings.length ? "warn" : "ok"}">${warnings.length ? warnings.join("; ") : "OK"}</td>
      `;
      previewBody.appendChild(tr);
    });
    previewSection.hidden = false;
    generateBtn.disabled = rows.length === 0;
  }

  async function handleFile(file) {
    setStatus(`Đang đọc ${file.name}...`);
    logEl.innerHTML = "";
    try {
      const { rows } = await ExcelParser.parseExcelFile(file);
      currentRows = rows;
      renderPreview(rows);
      setStatus(`Đã đọc ${rows.length} dòng từ ${file.name}.`);
    } catch (err) {
      console.error(err);
      setStatus(`Lỗi đọc file: ${err.message}`);
    }
  }

  fileInput.addEventListener("change", (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });

  ["dragover", "dragenter"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropZone.classList.add("dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropZone.classList.remove("dragover");
    })
  );
  dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  dropZone.addEventListener("click", () => fileInput.click());

  generateBtn.addEventListener("click", async () => {
    generateBtn.disabled = true;
    const zip = new JSZip();
    try {
      setStatus("Đang tải template PDF...");
      const pdfAssets = await PdfContract.loadAssets();

      const hopDongFolder = zip.folder("HopDong");
      const bbntFolder = zip.folder("BBNT");
      const mauSoFolder = zip.folder("MauSo02");

      const mau02Rows = [];

      for (let i = 0; i < currentRows.length; i++) {
        const row = currentRows[i];
        const name = safeName(row, i);
        setStatus(`Đang xử lý ${i + 1}/${currentRows.length}: ${name}`);

        try {
          const pdfBytes = await PdfContract.generateContractPdf(row, pdfAssets);
          hopDongFolder.file(`HopDong_${name}.pdf`, pdfBytes);
          log(`✓ Hợp đồng: ${name}`);
        } catch (err) {
          console.error(err);
          log(`✗ Lỗi tạo hợp đồng cho ${name}: ${err.message}`);
        }

        try {
          const bbnt = await DocxFill.generateBbnt(row);
          if (bbnt) {
            bbntFolder.file(`BBNT_${name}.docx`, bbnt.bytes);
            log(`✓ BBNT (${bbnt.itemCount} hạng mục): ${name}`);
          } else {
            log(`… Bỏ qua BBNT cho ${name} (không có hạng mục công việc)`);
          }
        } catch (err) {
          console.error(err);
          log(`✗ Lỗi tạo BBNT cho ${name}: ${err.message}`);
        }

        const hoTen = (row["Ho_ten"] || "").toString().trim();
        if (hoTen) mau02Rows.push(row);
      }

      if (mau02Rows.length > 0) {
        setStatus("Đang tạo Mẫu 02 (bảng kê chung)...");
        try {
          const mau02Bytes = await DocxFill.generateMau02(mau02Rows);
          mauSoFolder.file("Mau_02_BangKe.docx", mau02Bytes);
          log(`✓ Mẫu 02: liệt kê ${mau02Rows.length} người`);
        } catch (err) {
          console.error(err);
          log(`✗ Lỗi tạo Mẫu 02: ${err.message}`);
        }
      }

      setStatus("Đang nén file zip...");
      const zipBlob = await zip.generateAsync({ type: "blob" });
      saveAs(zipBlob, "HoSo_CTV.zip");
      setStatus(`Hoàn thành! Đã tạo hồ sơ cho ${currentRows.length} người.`);
    } catch (err) {
      console.error(err);
      setStatus(`Lỗi: ${err.message}`);
    } finally {
      generateBtn.disabled = false;
    }
  });
})();
