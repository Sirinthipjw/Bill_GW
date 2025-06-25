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

  const uniqueCONums = [...new Set(excelData.map(row => row.CONum))];
  // ✅ ใช้ for...of + await เพื่อให้รอการ render

for (const conum of uniqueCONums) {
   const allItems = excelData.filter(row => row.CONum === conum);
  const data = allItems[0]; // ใช้ข้อมูลหลักจากรายการแรก

  if (uniqueCONums.indexOf(conum) > 0) doc.addPage();

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

     // ✅ แยกสินค้าหลัก/ของแถมตาม CONum
    // const allItems = excelData.filter(row => row.CONum === data.CONum);
    let mainProducts = [];
    let freeProducts = [];

    if (allItems.length === 1) {
      // กรณีมีแค่ 1 รายการ → ให้เป็นสินค้าหลัก
      mainProducts = allItems;
      freeProducts = [];
    } else {
      // กรณีมีหลายรายการ → แยกสินค้าหลัก/ของแถมตาม PromoCode
      mainProducts = allItems.filter(item => !!item.PromoCode && item.PromoCode.toString().trim() !== "");
      freeProducts = allItems.filter(item => !item.PromoCode || item.PromoCode.toString().trim() === "");
    }


    const freeProductNames = freeProducts.map(item => item.ProductList).join(", ");
    filledHtml = filledHtml.replace(/{{FreeProductList}}/g, freeProductNames);

    const mainProduct = mainProducts[0] || {};
    filledHtml = filledHtml.replace(/{{ProductList}}/g, mainProduct.ProductList || "");
    filledHtml = filledHtml.replace(/{{Qty}}/g, mainProduct.Qty || "");
    filledHtml = filledHtml.replace(/{{UnitPrice_Formatted}}/g, formatCurrency(mainProduct.UnitPrice));
    filledHtml = filledHtml.replace(/{{Price_Formatted}}/g, formatCurrency(mainProduct.Price));

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
      if (typeof val === "string" && val.trim().toLowerCase() === "-") return "";
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
    function formatDate(dateStr) {
      if (!dateStr) return "";
      const date = new Date(dateStr);
      const localDate = new Date(date.getTime() + (7 * 60 * 60 * 1000)); // ปรับเป็นเวลาไทย
      const day = String(localDate.getDate()).padStart(2, "0");
      const month = String(localDate.getMonth() + 1).padStart(2, "0");
      const year = localDate.getFullYear();
      return `${day}/${month}/${year}`;
    }

    const priceFields = ["UnitPrice", "Price", "Disc", "SumPrice", "Vat", "TotalPrice"];

    // Object.entries(data).forEach(([key, value]) => {
    // const regex = new RegExp(`{{${key}}}`, "g");

    //  let cleanValue = safeValue(value);

    //   if (priceFields.includes(key)) {
    //     cleanValue = formatCurrency(cleanValue);
    //   }

    //   filledHtml = filledHtml.replace(regex,cleanValue);
    // });
    Object.entries(data).forEach(([key, value]) => {
      let cleanValue = safeValue(value);
      const regexBase = new RegExp(`{{${key}}}`, "g");
      filledHtml = filledHtml.replace(regexBase, cleanValue);
     
      // เพิ่มเติม: ถ้าเป็นจำนวนเงิน แปลงเป็นรูปแบบไทย + ข้อความ
      if (priceFields.includes(key)) {
        const formattedValue = formatCurrency(cleanValue);
        filledHtml = filledHtml.replace(new RegExp(`{{${key}_Formatted}}`, "g"), formattedValue);

        const textValue = numberToThaiText(cleanValue);
        filledHtml = filledHtml.replace(new RegExp(`{{${key}_Text}}`, "g"), textValue);
      }
      if (!safeValue(data.Model)) {
        filledHtml = filledHtml.replace(/<div[^>]*class=["']model-line["'][^>]*>.*?<\/div>\s*/g, "แบบรถ : ");
      }

      if (!safeValue(data.Color)) {
        filledHtml = filledHtml.replace(/<div[^>]*class=["']color-line["'][^>]*>.*?<\/div>\s*/g, '<div class="color-line">สี : </div>');
      }

      if (!safeValue(data.CertNum)) {
        filledHtml = filledHtml.replace(/<div[^>]*class=["']CertNum-line["'][^>]*>.*?<\/div>\s*/g, '<div class="CertNum-line">หมายเลขเครื่อง : </div>');
      }

      if (!safeValue(data.SeriNum)) {
        filledHtml = filledHtml.replace(/<div[^>]*class=["']SeriNum-line["'][^>]*>.*?<\/div>\s*/g, '<div class="SeriNum-line">หมายเลขตัวถัง : </div>');
      }

      if (!safeValue(data.RefCust)) {
        filledHtml = filledHtml.replace(/<div[^>]*class=["']RefCust-line["'][^>]*>.*?<\/div>\s*/g, '<div class="RefCust-line">&nbsp;</div>');
      }

      filledHtml = filledHtml.replace(/{{Date}}/g, formatDate(data.Date));
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
