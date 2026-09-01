(function () {
  const SCALE = TOOL_ASSETS.scale;
  const FONT_SIZE_PT = 12;
  const LEADING_PT = 14;

  // Gia tri mau, dung khi chua upload data.xlsx (lay tu vi du thuc te cua nguoi dung).
  const SAMPLE_DEFAULTS = {
    Ho_ten: "ĐOÀN QUỐC VIỆT",
    Ngay1: "02", Thang1: "09", Nam1: "2026",
    DD: "01", MM: "11", YYYY: "2003",
    Quoc_tich: "Việt Nam",
    CCCD: "001203009547",
    D1: "14", M1: "04", Y1: "2021",
    Noi_cap: "CỤC CẢNH SÁT QLHC VỀ TTXH",
    Noi_o: "614 nhà N3a KDT Trung Hòa Nhân Chính, Phường Yên Hòa, Hà Nội",
    MST: "001203009547",
    STK: "5880111038888",
    Bank: "MB Bank",
    Chi_nhanh: "THANH XUÂN, HÀ NỘI",
    SDT: "0947635868",
    Cong_viec: "Thực hiện giám sát thi đấu các đội tuyển tại 2026 PMPL VN FALL",
    Time_Range: "02/09/2026 - 06/09/2026",
    Location: "NTĐ Tương Mai, Hà Nội: Ngõ 104 Phố Nguyễn An Ninh, phường Tương Mai, quận Hoàng Mai, Hà Nội.",
    So_tien: "2,500,000",
    Bang_chu: "Hai triệu năm trăm nghìn đồng",
    "Sản phẩm": "PUBG Mobile",
    PIC: "NGUYỄN HỒNG TỐ NHƯ",
    "Email PIC": "nhunht2@vng.com.vn",
  };

  // Chen @font-face tu font nhung san (khop voi font that dung trong PDF).
  const styleEl = document.createElement("style");
  styleEl.textContent = `@font-face { font-family: "TimesPreview"; src: url(data:font/ttf;base64,${TOOL_ASSETS.fontBase64}) format("truetype"); }`;
  document.head.appendChild(styleEl);

  const ORIGINAL_SNAPSHOT = JSON.parse(JSON.stringify(PdfContract.COORDINATES));
  const state = {
    coords: PdfContract.COORDINATES, // mutate truc tiep de "Xem thu PDF that" dung ngay tọa do moi
    activePage: 0,
    excelRows: [],
    previewRow: null,
  };
  let selected = null; // { key, idx }
  const elementMap = new Map(); // `${key}|${idx}` -> element

  const pageTabsEl = document.getElementById("pageTabs");
  const pageContainerEl = document.getElementById("pageContainer");
  const fieldPanelEl = document.getElementById("fieldPanel");
  const dropZoneEl = document.getElementById("dropZone");
  const fileInputEl = document.getElementById("fileInput");
  const rowSelectEl = document.getElementById("rowSelect");
  const exportOutputEl = document.getElementById("exportOutput");

  function previewValues() {
    const row = state.previewRow;
    const values = {};
    for (const key of Object.keys(SAMPLE_DEFAULTS)) {
      if (key === "Time_Range") continue;
      const v = row ? row[key] : undefined;
      values[key] = v !== undefined && v !== null && String(v).trim() !== "" ? String(v) : SAMPLE_DEFAULTS[key];
    }
    const t1 = row ? row["Time"] : null;
    const t2 = row ? row["Time2"] : null;
    const parts = [t1, t2].map((v) => (v ? String(v).trim() : "")).filter(Boolean);
    values.Time_Range = parts.length ? parts.join(" - ") : SAMPLE_DEFAULTS.Time_Range;
    return values;
  }

  function buildPageTabs() {
    pageTabsEl.innerHTML = "";
    TOOL_ASSETS.pageSizes.forEach((_, i) => {
      const btn = document.createElement("button");
      btn.textContent = `Trang ${i + 1}`;
      btn.className = i === state.activePage ? "active" : "";
      btn.addEventListener("click", () => {
        state.activePage = i;
        buildPageTabs();
        renderPage();
      });
      pageTabsEl.appendChild(btn);
    });
  }

  function wrapPreviewText(text, widthPt, fontSizePt) {
    // Dung canvas measureText voi font that de uoc luong xuong dong tuong tu wrapText() trong pdfContract.js.
    const canvas = wrapPreviewText._canvas || (wrapPreviewText._canvas = document.createElement("canvas"));
    const ctx = canvas.getContext("2d");
    ctx.font = `${fontSizePt * SCALE}px TimesPreview, "Times New Roman", serif`;
    const words = String(text).split(/\s+/).filter(Boolean);
    const lines = [];
    let current = "";
    const maxWidthPx = widthPt * SCALE;
    for (const word of words) {
      const trial = current ? `${current} ${word}` : word;
      if (ctx.measureText(trial).width > maxWidthPx && current) {
        lines.push(current);
        current = word;
      } else {
        current = trial;
      }
    }
    if (current) lines.push(current);
    return lines;
  }

  function renderPage() {
    pageContainerEl.innerHTML = "";
    elementMap.clear();
    const size = TOOL_ASSETS.pageSizes[state.activePage];
    pageContainerEl.style.width = size.width * SCALE + "px";
    pageContainerEl.style.height = size.height * SCALE + "px";
    pageContainerEl.style.backgroundImage = `url(${TOOL_ASSETS.pages[state.activePage]})`;

    const values = previewValues();

    for (const key of Object.keys(state.coords)) {
      state.coords[key].forEach((loc, idx) => {
        if (loc.page !== state.activePage) return;
        const el = document.createElement("div");
        el.className = "field" + (loc.type === "paragraph" ? " paragraph" : "");
        el.dataset.key = key;
        el.dataset.idx = String(idx);
        el.style.left = loc.x * SCALE + "px";
        el.style.bottom = loc.y * SCALE + "px";
        el.style.fontSize = FONT_SIZE_PT * SCALE + "px";

        if (loc.type === "paragraph") {
          el.style.width = loc.width * SCALE + "px";
          el.style.lineHeight = LEADING_PT * SCALE + "px";
          const lines = wrapPreviewText(values[key] || "", loc.width, FONT_SIZE_PT);
          el.textContent = lines.join(" ");
          // hien thi dung so dong bang <br> de nguoi dung thay chinh xac so dong se xuong
          el.innerHTML = lines.map((l) => escapeHtml(l)).join("<br/>");
        } else if (loc.align === "center") {
          el.classList.add("centered");
          el.style.width = loc.width * SCALE + "px";
          el.textContent = values[key] || "";
        } else {
          el.textContent = values[key] || "";
        }

        const label = state.coords[key].length > 1 ? `${key} #${idx + 1}` : key;
        el.title = label;

        attachDrag(el, key, idx);
        el.addEventListener("pointerdown", (e) => e.stopPropagation());
        el.addEventListener("click", (e) => {
          e.stopPropagation();
          selectField(key, idx);
        });

        pageContainerEl.appendChild(el);
        elementMap.set(`${key}|${idx}`, el);
      });
    }

    if (selected) {
      const el = elementMap.get(`${selected.key}|${selected.idx}`);
      if (el) el.classList.add("selected");
    }
  }

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function attachDrag(el, key, idx) {
    el.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      selectField(key, idx);
      const loc = state.coords[key][idx];
      const startX = e.clientX;
      const startY = e.clientY;
      const startPtX = loc.x;
      const startPtY = loc.y;
      el.setPointerCapture(e.pointerId);

      function onMove(ev) {
        const dxPt = (ev.clientX - startX) / SCALE;
        const dyPt = (ev.clientY - startY) / SCALE;
        loc.x = Math.round((startPtX + dxPt) * 2) / 2;
        loc.y = Math.round((startPtY - dyPt) * 2) / 2;
        el.style.left = loc.x * SCALE + "px";
        el.style.bottom = loc.y * SCALE + "px";
        updateFieldPanel();
      }
      function onUp() {
        el.releasePointerCapture(e.pointerId);
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
      }
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    });
  }

  function selectField(key, idx) {
    document.querySelectorAll(".field.selected").forEach((e) => e.classList.remove("selected"));
    selected = { key, idx };
    const el = elementMap.get(`${key}|${idx}`);
    if (el) el.classList.add("selected");
    updateFieldPanel();
  }

  function updateFieldPanel() {
    if (!selected) {
      fieldPanelEl.innerHTML = '<p class="muted">Bấm vào 1 nhãn trên trang để chỉnh.</p>';
      return;
    }
    const loc = state.coords[selected.key][selected.idx];
    const label = state.coords[selected.key].length > 1 ? `${selected.key} (#${selected.idx + 1})` : selected.key;
    const isParagraph = loc.type === "paragraph";
    const isCentered = loc.align === "center";
    const hasWidth = isParagraph || isCentered;
    fieldPanelEl.innerHTML = `
      <label>Trường</label>
      <input type="text" value="${label}" disabled />
      <label>Trang</label>
      <select id="fp_page">
        ${TOOL_ASSETS.pageSizes.map((_, i) => `<option value="${i}" ${i === loc.page ? "selected" : ""}>Trang ${i + 1}</option>`).join("")}
      </select>
      <label>X (pt)${isCentered ? " — mép trái khung canh giữa" : ""}</label>
      <input type="number" id="fp_x" value="${loc.x}" step="0.5" />
      <label>Y (pt)</label>
      <input type="number" id="fp_y" value="${loc.y}" step="0.5" />
      ${hasWidth ? `<label>Chiều rộng khung (pt)</label><input type="number" id="fp_width" value="${loc.width}" step="1" />` : ""}
      <label><input type="checkbox" id="fp_center" ${isCentered ? "checked" : ""} style="width:auto;display:inline-block;margin-right:0.4rem;" />Căn giữa trong khung (chỉ áp dụng cho trường 1 dòng)</label>
    `;
    document.getElementById("fp_page").addEventListener("change", (e) => {
      loc.page = Number(e.target.value);
      state.activePage = loc.page;
      buildPageTabs();
      renderPage();
    });
    document.getElementById("fp_x").addEventListener("input", (e) => {
      loc.x = Number(e.target.value) || 0;
      renderPage();
    });
    document.getElementById("fp_y").addEventListener("input", (e) => {
      loc.y = Number(e.target.value) || 0;
      renderPage();
    });
    if (hasWidth) {
      document.getElementById("fp_width").addEventListener("input", (e) => {
        loc.width = Number(e.target.value) || 1;
        renderPage();
      });
    }
    document.getElementById("fp_center").addEventListener("change", (e) => {
      if (isParagraph) {
        e.target.checked = false;
        alert("Trường dạng xuống dòng (paragraph) không dùng căn giữa - bỏ qua.");
        return;
      }
      if (e.target.checked) {
        loc.align = "center";
        if (!loc.width) loc.width = 150;
      } else {
        delete loc.align;
      }
      renderPage();
      updateFieldPanel();
    });
  }

  window.addEventListener("keydown", (e) => {
    if (!selected) return;
    const tag = document.activeElement && document.activeElement.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
    const loc = state.coords[selected.key][selected.idx];
    const step = e.shiftKey ? 5 : 1;
    let handled = true;
    if (e.key === "ArrowLeft") loc.x -= step;
    else if (e.key === "ArrowRight") loc.x += step;
    else if (e.key === "ArrowUp") loc.y += step;
    else if (e.key === "ArrowDown") loc.y -= step;
    else handled = false;
    if (handled) {
      e.preventDefault();
      renderPage();
      updateFieldPanel();
    }
  });

  document.getElementById("pageContainer").addEventListener("click", () => {
    selected = null;
    document.querySelectorAll(".field.selected").forEach((e) => e.classList.remove("selected"));
    updateFieldPanel();
  });

  document.getElementById("resetBtn").addEventListener("click", () => {
    for (const key of Object.keys(ORIGINAL_SNAPSHOT)) {
      ORIGINAL_SNAPSHOT[key].forEach((loc, idx) => {
        // Thay the hoan toan (khong dung Object.assign) de xoa luon cac thuoc tinh
        // nguoi dung moi them (vd align/width khi bat "can giua") ma ban goc khong co.
        state.coords[key][idx] = JSON.parse(JSON.stringify(loc));
      });
    }
    selected = null;
    renderPage();
    updateFieldPanel();
  });

  // Tra ve key nguyen dang (vd Ho_ten) neu la identifier JS hop le, nguoc lai bao bang
  // dau nhay kep (vd "Sản phẩm", "Email PIC") - thieu buoc nay se xuat ra code loi cu phap.
  function formatKey(key) {
    return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? key : JSON.stringify(key);
  }

  function formatCoordinates(coords) {
    const lines = ["  const COORDINATES = {"];
    for (const key of Object.keys(coords)) {
      const locsStr = coords[key].map((loc) => {
        const parts = [`page: ${loc.page}`];
        if (loc.type) parts.push(`type: "${loc.type}"`);
        parts.push(`x: ${loc.x}`, `y: ${loc.y}`);
        if (loc.type === "paragraph") {
          parts.push(`width: ${loc.width}`, `height: ${loc.height !== undefined ? loc.height : 40}`);
        } else if (loc.align === "center") {
          parts.push(`width: ${loc.width}`, `align: "center"`);
        }
        return `{ ${parts.join(", ")} }`;
      });
      const inner = locsStr.length > 1 ? "[\n      " + locsStr.join(",\n      ") + ",\n    ]" : "[" + locsStr[0] + "]";
      lines.push(`    ${formatKey(key)}: ${inner},`);
    }
    lines.push("  };");
    return lines.join("\n");
  }

  document.getElementById("exportBtn").addEventListener("click", () => {
    exportOutputEl.value = formatCoordinates(state.coords);
    document.getElementById("copyBtn").disabled = false;
  });

  document.getElementById("copyBtn").addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(exportOutputEl.value);
      const btn = document.getElementById("copyBtn");
      const old = btn.textContent;
      btn.textContent = "Đã sao chép!";
      setTimeout(() => (btn.textContent = old), 1500);
    } catch (err) {
      exportOutputEl.select();
      document.execCommand("copy");
    }
  });

  // --- Upload data.xlsx de xem truoc bang du lieu that ---
  async function handleFile(file) {
    const { rows } = await ExcelParser.parseExcelFile(file);
    state.excelRows = rows;
    rowSelectEl.innerHTML = '<option value="">(dùng dữ liệu mẫu)</option>';
    rows.forEach((row, i) => {
      const name = (row["Ho_ten"] || "").toString().trim() || `Dòng ${i + 2}`;
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = name;
      rowSelectEl.appendChild(opt);
    });
    rowSelectEl.disabled = false;
  }

  fileInputEl.addEventListener("change", (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });
  dropZoneEl.addEventListener("click", () => fileInputEl.click());
  ["dragover", "dragenter"].forEach((evt) =>
    dropZoneEl.addEventListener(evt, (e) => {
      e.preventDefault();
      dropZoneEl.classList.add("dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropZoneEl.addEventListener(evt, (e) => {
      e.preventDefault();
      dropZoneEl.classList.remove("dragover");
    })
  );
  dropZoneEl.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  rowSelectEl.addEventListener("change", (e) => {
    const idx = e.target.value;
    state.previewRow = idx === "" ? null : state.excelRows[Number(idx)];
    renderPage();
  });

  // --- Xem thu PDF that (render bang pdf-lib voi toa do hien tai) ---
  document.getElementById("testRenderBtn").addEventListener("click", async () => {
    const btn = document.getElementById("testRenderBtn");
    btn.disabled = true;
    btn.textContent = "Đang tạo PDF...";
    try {
      const row = state.previewRow || SAMPLE_DEFAULTS;
      const assets = await PdfContract.loadAssets();
      const bytes = await PdfContract.generateContractPdf(row, assets);
      const blob = new Blob([bytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const wrap = document.getElementById("pdfPreviewWrap");
      const iframe = document.getElementById("pdfPreview");
      if (iframe.src) URL.revokeObjectURL(iframe.src);
      iframe.src = url;
      wrap.hidden = false;
      wrap.scrollIntoView({ behavior: "smooth" });
    } catch (err) {
      alert("Lỗi tạo PDF: " + err.message);
      console.error(err);
    } finally {
      btn.disabled = false;
      btn.textContent = "Xem thử PDF thật";
    }
  });

  document.getElementById("closePreviewBtn").addEventListener("click", () => {
    document.getElementById("pdfPreviewWrap").hidden = true;
  });

  buildPageTabs();
  renderPage();
})();
