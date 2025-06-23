let excelData = [];
const viewPdfBtn = document.getElementById("viewPDF");

function numberToThaiText(number) {
  const numText = ['ศูนย์','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'];
  const rankText = ['','สิบ','ร้อย','พัน','หมื่น','แสน','ล้าน'];

  number = parseFloat(number).toFixed(2);
  const [intPart, decPart] = number.split('.');

  function convert(numStr) {
    let result = '';
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i]);
      if (digit === 0) continue;
      if (i === len - 1 && digit === 1 && len > 1) {
        result += 'เอ็ด';
      } else if (i === len - 2 && digit === 2) {
        result += 'ยี่';
      } else if (i === len - 2 && digit === 1) {
        result += '';
      } else {
        result += numText[digit];
      }
      result += rankText[len - i - 1];
    }
    return result;
  }

  let text = convert(intPart) + 'บาท';
  if (decPart === '00') {
    text += 'ถ้วน';
  } else {
    text += convert(decPart) + 'สตางค์';
  }
  return text;
}


document.getElementById("excelFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file || !file.name.endsWith(".xlsx")) {
    alert("กรุณาเลือกไฟล์ .xlsx เท่านั้น");
    return;
  }

  document.getElementById("fileName").textContent = "ไฟล์ที่เลือก: " + file.name;
  document.getElementById("fileName").style.display = "block";

  const reader = new FileReader();
  reader.onload = function (evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
    console.log("โหลดข้อมูลจาก Excel:", excelData);

    if (excelData.length > 0) {
      viewPdfBtn.style.display = "inline-block";
    }
  };
  reader.readAsArrayBuffer(file);
});

viewPdfBtn.addEventListener("click", async function () {
  if (excelData.length === 0) {
    alert("ไม่มีข้อมูลสำหรับสร้าง PDF");
    return;
  }

  const templateResponse = await fetch("Bill-template1.html");
  const templateHtml = await templateResponse.text();

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "A5",
  });

  // ✅ ใช้ for...of + await เพื่อให้รอการ render
  for (let i = 0; i < excelData.length; i++) {
    const data = excelData[i];
    if (i > 0) doc.addPage();


     let filledHtml = templateHtml;

    // ✅ แปลง TotalPrice เป็นตัวหนังสือ
    const totalPriceText = numberToThaiText(data.TotalPrice);
    filledHtml = filledHtml.replace(/{{TotalPriceText}}/g, totalPriceText);

    // ✅ สร้าง BranchInfo
    const rawBranchId = data.BranchIdCust;
    const rawBranchName = data.BranchNameCust;

    let branchInfo = "";
    if (isValidValue(rawBranchId) && isValidValue(rawBranchName)) {
      const paddedBranchId = String(rawBranchId).padStart(5, "0");
      branchInfo = ` สาขาที่ : ${paddedBranchId} ${rawBranchName}`;
    }
    data.BranchInfo = branchInfo;


    function isValidValue(val) {
      if (val === null || val === undefined) return false;

      if (typeof val === "number") return true;

      if (typeof val === "string") {
        const cleaned = val.trim().toLowerCase();
        return cleaned !== "" && cleaned !== "null" && cleaned !== "undefined";
      }

      return false;
    }

    function safeValue(val) {
      if (val === null || val === undefined) return "";
      if (typeof val === "string" && val.trim().toLowerCase() === "null") return "";
      if (typeof val === "string" && val.trim().toLowerCase() === "undefined") return "";
      return val;
    }

    
    function formatCurrency(value) {
    const num = parseFloat(value);
    if (isNaN(num)) return ""; // ถ้าไม่ใช่ตัวเลขให้คืนค่าว่าง
    return num.toLocaleString("th-TH", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
  });
}
    const priceFields = ["UnitPrice", "Price", "Disc", "SumPrice", "Vat", "TotalPrice"];

    Object.entries(data).forEach(([key, value]) => {
    const regex = new RegExp(`{{${key}}}`, "g");

     let cleanValue = safeValue(value);

      if (priceFields.includes(key)) {
        cleanValue = formatCurrency(cleanValue);
      }

      filledHtml = filledHtml.replace(regex,cleanValue);
    });

    const container = document.createElement("div");
    container.innerHTML = filledHtml;
    container.style.position = "absolute";
    container.style.left = "-9999px";
    document.body.appendChild(container);
    
    const canvas = await html2canvas(container, { scale: 3 });
    const imgData = canvas.toDataURL("image/png");
    const margin = 5; // เว้นขอบ 5mm รอบด้าน
    const pdfWidth = doc.internal.pageSize.getWidth() - margin * 2;
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    doc.addImage(imgData, "PNG", margin, margin, pdfWidth, pdfHeight);


    container.remove();
  }

  // ✅ เปิด PDF หลังจาก render เสร็จทุกหน้า
  const blob = doc.output("blob");
  const url = URL.createObjectURL(blob);

  const win = window.open("", "_blank");
  if (!win) {
    alert("กรุณาอนุญาต popup");
    return;
  }

  const iframe = win.document.createElement("iframe");
  iframe.style.width = "100%";
  iframe.style.height = "100%";
  iframe.style.border = "none";
  iframe.src = url;
  win.document.body.style.margin = "0";
  win.document.body.appendChild(iframe);
});
