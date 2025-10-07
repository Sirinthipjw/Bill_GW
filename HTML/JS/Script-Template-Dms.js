let excelData = [];

const viewPdfBtn = document.getElementById("viewPDF");
const dropZone = document.getElementById("dropZone");
const rawDate = document.getElementById("create-date");
const fileName = document.getElementById("fileName");

// แปลงตัวเลขเป็นตัวหนังสือ
function numberToThaiText(number) {
  const numText = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
  const rankText = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"];

  if (isNaN(number) || number === null || number === undefined || number === "") {
    return "";
  }

  number = parseFloat(number).toFixed(2);
  const [intPart, decPart] = number.split(".");

  function convert(numStr) {
    let result = "";
    if (typeof numStr !== "string" || !numStr.trim().length) return "";
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i]);
      if (isNaN(digit) || digit === 0) continue;

      if (i === len - 1 && digit === 1 && len > 1) {
        result += "เอ็ด";
      } else if (i === len - 2 && digit === 2) {
        result += "ยี่";
      } else if (i === len - 2 && digit === 1) {
        result += "";
      } else {
        result += numText[digit];
      }
      result += rankText[len - i - 1];
    }
    return result;
  }

  let text = convert(intPart) + "บาท";
  if (decPart === "00") {
    text += "ถ้วน";
  } else {
    text += convert(decPart) + "สตางค์";
  }

  return text;
}

document.getElementById("uploadBtn").addEventListener("click", function () {
    document.getElementById("excelFile").click();
  });

// อ่าน file csv
document.getElementById("excelFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file || !file.name.endsWith(".csv")) {
    alert("กรุณาเลือกไฟล์ .csv เท่านั้น");
    return;
  }

  document.getElementById("fileName").textContent =
    "Your Select File : " + file.name;
  document.getElementById("fileName").style.display = "block";
  document.querySelector(".BoxShow").style.display = "block";

  const reader = new FileReader();
  reader.onload = function (evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
    

    function cleanKeys(row) {
      const cleaned = {};
      for (const key in row) {
        const cleanKey = key.replace(/^\uFEFF/, ""); // ลบ BOM ที่ต้น key
        cleaned[cleanKey] = row[key];
      }
      return cleaned;
    }
    console.log("โหลดข้อมูลจาก Excel (ล้าง BOM แล้ว):", excelData);
    excelData = excelData.map(cleanKeys);
    if (excelData.length > 0) {
      viewPdfBtn.style.display = "inline-block"; // ✅ แสดงปุ่มเมื่อมีข้อมูล
    }
  };
  reader.readAsArrayBuffer(file);
  const fileInput = document.getElementById("fileName");
  // ✅ Event: เลือกไฟล์ด้วยปุ่ม
  fileInput.addEventListener("change", function (e) {
    const file = e.target.files[0];
    handleExcelFile(file);
  });
});

document.addEventListener("DOMContentLoaded", function () {
    const billType = document.getElementById("billType");
    const viewPDF = document.getElementById("viewPDF");

    if (billType) {
        billType.addEventListener("change", updateBillTitle);
    }

    if (viewPDF) {
        viewPDF.addEventListener("click", function () {
          const element = document.getElementById("billTemplate");
          html2pdf().from(element).save("Bill.pdf");
        });
      }
});



viewPdfBtn.addEventListener("click", async function () {
  
    if (excelData.length === 0) {
        alert("ไม่มีข้อมูลสำหรับสร้าง PDF");
        return;
    }
    const templateResponse = await fetch("/Bill-template-dms.html");
    const templateHtml = await templateResponse.text();

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "A4",
    });

        for (let i = 0; i < excelData.length; i++) {
            const data = excelData[i];
            let filledHtml = templateHtml;

            const totalPages = excelData.length;

            let price = data.Price
            filledHtml = filledHtml.replace(/{{Price}}/g, formatPrice(price));
            // ✅ แปลง TotalPrice เป็นตัวหนังสือ
            const totalPriceText = numberToThaiText(data.Price);
            filledHtml = filledHtml.replace(/{{TotalPriceText}}/g, totalPriceText);

            filledHtml = filledHtml.replace(
                /<span class="page-number">.*?<\/span>/g,
                `<span class="page-number">Page ${i + 1} of ${totalPages}</span>`
            );

            const today = new Date();

            const formattedDate = today.toLocaleDateString("en-GB", { day:"2-digit", month:"2-digit", year:"numeric" });
            filledHtml = filledHtml.replace(
                /<strong>วันที่ \/ Date\. :<\/strong>\s*\d{2}\/\d{2}\/\d{4}/g,
                `<strong>วันที่ / Date. :</strong> ${formattedDate}`
            );

            const docNo = generateDocNo(data.Site, data.BranchCom, data.Custpo);

            filledHtml = filledHtml.replace(
            /<strong>เลขที่ \/ No\. :<\/strong>\s*\S+/g,
            `<strong>เลขที่ / No. :</strong> ${docNo}`
            );
            
            // ✅ สร้างข้อมูลสาขา (ตัวอย่าง)
            const rawBranchId = data.BranchIdCust;
            const rawBranchName = "สำนักงานใหญ่";
            let branchInfo = "";
            let branchIDc = "";
            if (rawBranchId && rawBranchName) {
            const paddedBranchId = String(rawBranchId).padStart(5, "0");
            branchInfo = ` สาขาที่ : ${paddedBranchId} ${rawBranchName}`;
            branchIDc = `${paddedBranchId}`
            }
            filledHtml = filledHtml.replace(/{{BranchInfo}}/g, branchInfo);
            filledHtml = filledHtml.replace(/{{branchIDc}}/g, branchIDc);

            // ✅ แทนค่าข้อมูลทั่วไปทั้งหมด
            for (const key in data) {
            const regex = new RegExp(`{{${key}}}`, "g");
            filledHtml = filledHtml.replace(regex, data[key] ?? "");
            }

            // ✅ แปลง HTML เป็นภาพ
            const container = document.createElement("div");
            container.innerHTML = filledHtml;
            container.style.position = "absolute";
            container.style.left = "-9999px";
            document.body.appendChild(container);

            const canvas = await html2canvas(container, { scale: 3 });
            const imgData = canvas.toDataURL("image/jpeg");

            const margin = 5;
            const pdfWidth = doc.internal.pageSize.getWidth() - margin * 2;
            const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
            if (i > 0) doc.addPage();
            doc.addImage(imgData, "JPEG", margin, margin, pdfWidth, pdfHeight);

            container.remove();
            console.log(`✅ สร้างหน้า PDF สำหรับแถวที่ ${i + 1}/${excelData.length}`);

            
        }
        function formatPrice(value) {
            const number = parseFloat(value);
            if (isNaN(number)) return value; // ถ้าไม่ใช่ตัวเลข ให้คืนค่าเดิม
            return number.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }

        function extractBranchCode(branchCom) {
            if (!branchCom) return "000";
            const parts = branchCom.trim().split(/\s+/); // "00000 สำนักงานใหญ่" → ["00000", "สำนักงานใหญ่"]
            const rawCode = parts[0];
            return rawCode.slice(-3); // เอา 3 หลักท้าย เช่น "000"
        }

        function extractCustPoLast4(custPo) {
            if (!custPo) return "0000";
            const digits = custPo.replace(/\D/g, ""); // เอาเฉพาะตัวเลข เช่น "1VO-25080001" → "25080001"
            return digits.slice(-4).padStart(4, "0"); // เอา 4 ตัวท้าย เช่น "0001"
        }

        function getCompanyIdFromSite(site) {
            const map = {
                GM: "1",
                GT: "2",
                AC: "3",
                AP: "4",
                WL: "6",
                WM: "A",
                SM: "9",
                GC: "G",
                MG: "M",
                GL: "L",
            };
            return map[site] || "0"; // ถ้า site ไม่มีใน map ให้ใช้ "0" เป็นค่าเริ่มต้น
        }
       function generateDocNo(site, branchCom, custPo) {
            const companyId = getCompanyIdFromSite(site);
            // ✅ 1. ปี (ค.ศ. 2 หลักท้าย)
            const year = new Date().getFullYear().toString().slice(-2);

            // ✅ 2. เดือน (1–12)
            const month = new Date().getMonth() + 1;

            // ✅ 3. แปลงเดือนเป็นเลขหรืออักษร
            let monthCode;
            if (month <= 9) {
                monthCode = month.toString(); // ใช้เลขตรง ๆ
            } else if (month === 10) {
                monthCode = "X";
            } else if (month === 11) {
                monthCode = "Y";
            } else if (month === 12) {
                monthCode = "Z";
            }

            // ✅ 4. branchCom ให้มีความยาว 3 หลัก เช่น "5" → "005"
            const branchCode = extractBranchCode(branchCom);

            // ✅ 5. Custpo เอา 4 ตัวท้าย (ถ้ามีไม่ถึง เติม 0 ข้างหน้า)
            const custPoLast4 = extractCustPoLast4(custPo);

            // ✅ 6. รวมเป็นเลขที่
            const docNo = `${companyId}M${branchCode}${year}${monthCode}${custPoLast4}`;
            return docNo;
        }
    

   
  const blob = doc.output("blob");
  const url = URL.createObjectURL(blob);
  // doc.save("output.pdf"); 
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


    setTimeout(() => {
    location.reload();
  }, 10000);
});








