// Doc so tien ra chu tieng Viet (dung cho tong tien Mau 02, vi day la gia tri TINH TOAN
// (tong nhieu nguoi cong lai) nen khong the lay san tu cot Excel nhu Bang_chu tung nguoi.
(function (global) {
  const DIGIT_WORDS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
  const GROUP_NAMES = ["", " nghìn", " triệu", " tỷ"];

  function readGroup(n, isFirstGroupOfNumber) {
    const hundreds = Math.floor(n / 100);
    const tens = Math.floor((n % 100) / 10);
    const units = n % 10;
    const parts = [];

    if (hundreds > 0) {
      parts.push(DIGIT_WORDS[hundreds] + " trăm");
    } else if (!isFirstGroupOfNumber && n > 0) {
      parts.push("không trăm");
    }

    if (tens === 0) {
      if (units > 0 && (hundreds > 0 || !isFirstGroupOfNumber)) {
        parts.push("lẻ " + DIGIT_WORDS[units]);
      } else if (units > 0) {
        parts.push(DIGIT_WORDS[units]);
      }
    } else if (tens === 1) {
      parts.push("mười");
      if (units === 5) parts.push("lăm");
      else if (units === 1) parts.push("một");
      else if (units > 0) parts.push(DIGIT_WORDS[units]);
    } else {
      parts.push(DIGIT_WORDS[tens] + " mươi");
      if (units === 1) parts.push("mốt");
      else if (units === 5) parts.push("lăm");
      else if (units > 0) parts.push(DIGIT_WORDS[units]);
    }

    return parts.join(" ");
  }

  // Doc mot so nguyen khong am ra chu, ket thuc bang "dong" (vd 2500000 -> "Hai triệu năm trăm nghìn đồng").
  function numberToVietnameseWords(num) {
    const n0 = Math.floor(Math.abs(Number(num) || 0));
    if (n0 === 0) return "Không đồng";

    const groups = [];
    let n = n0;
    while (n > 0) {
      groups.push(n % 1000);
      n = Math.floor(n / 1000);
    }

    let firstNonZeroIdx = -1;
    for (let i = groups.length - 1; i >= 0; i--) {
      if (groups[i] > 0) {
        firstNonZeroIdx = i;
        break;
      }
    }

    const words = [];
    for (let i = groups.length - 1; i >= 0; i--) {
      if (groups[i] === 0) continue;
      words.push(readGroup(groups[i], i === firstNonZeroIdx) + GROUP_NAMES[i]);
    }

    let result = words.join(" ").replace(/\s+/g, " ").trim();
    result = result.charAt(0).toUpperCase() + result.slice(1);
    return result + " đồng";
  }

  // Chuyen "2,500,000" hoac 2500000 -> 2500000 (bo qua ky tu khong phai so).
  function parseAmount(v) {
    if (v === undefined || v === null || v === "") return 0;
    const n = Number(String(v).replace(/[^0-9]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // "2500000" -> "2,500,000"
  function formatAmount(n) {
    return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  global.VnNumberWords = { numberToVietnameseWords, parseAmount, formatAmount };
})(window);
