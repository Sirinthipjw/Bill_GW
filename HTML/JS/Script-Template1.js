let excelData = [];

const viewPdfBtn = document.getElementById("viewPDF");
const dropZone = document.getElementById("dropZone");
const rawDate = document.getElementById("create-date");
const fileName = document.getElementById("fileName");

// function numberToThaiText(number) {
//   const numText = [
//     "ศูนย์",
//     "หนึ่ง",
//     "สอง",
//     "สาม",
//     "สี่",
//     "ห้า",
//     "หก",
//     "เจ็ด",
//     "แปด",
//     "เก้า",
//   ];
//   const rankText = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"];

//   number = parseFloat(number).toFixed(2);
//   const [intPart, decPart] = number.split(".");

//   function convert(numStr) {
//     let result = "";
//     const len = numStr.length;
//     for (let i = 0; i < len; i++) {
//       const digit = parseInt(numStr[i]);
//       if (digit === 0) continue;
//       if (i === len - 1 && digit === 1 && len > 1) {
//         result += "เอ็ด";
//       } else if (i === len - 2 && digit === 2) {
//         result += "ยี่";
//       } else if (i === len - 2 && digit === 1) {
//         result += "";
//       } else {
//         result += numText[digit];
//       }
//       result += rankText[len - i - 1];
//     }
//     return result;
//   }

//   let text = convert(intPart) + "บาท";
//   if (decPart === "00") {
//     text += "ถ้วน";
//   } else {
//     text += convert(decPart) + "สตางค์";
//   }
//   return text;
// }

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

      function updateBillTitle() {
        const billTitle = document.getElementById("billTitle");
        if (!billTitle) return;

        const selected = billType ? billType.value : "";
        if (selected === "rc1") {
          billTitle.textContent = "ใบสำคัญรับเงิน";
        } else if (selected === "rc2") {
          billTitle.textContent = "ใบเสร็จรับเงิน/ใบกำกับภาษี";
        } else if (selected === "rc3") {
          billTitle.textContent = "ใบรับเงินอื่นๆ";
        } else {
          billTitle.textContent = "กรุณาเลือกประเภทใบเสร็จ";
        }
      }

      if (billType) {
        billType.addEventListener("change", updateBillTitle);
      }

      if (viewPDF) {
        viewPDF.addEventListener("click", function () {
          updateBillTitle(); // อัปเดต title ก่อน export
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
  // if (fileName) {
  //   fileName.textContent = "";
  // }
  // viewPdfBtn.disabled = true;
  // setTimeout(() => {
  //   location.reload(); // รีเฟรชหน้า
  // }, 1500);

  const templateResponse = await fetch("/Bill-template-dms.html");
  const templateHtml = await templateResponse.text();

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "A4",
  });


  //ใช้ conum เป็นตัวกรอง
  // const uniqueCONums = [...new Set(excelData.map((row) => row.CONum))];
  // // ✅ ใช้ for...of + await เพื่อให้รอการ render

  // for (const conum of uniqueCONums) {
  //   const allItems = excelData.filter((row) => row.CONum === conum);
  //   const InvGroup = excelData.filter((row) => row.Invoice_Group === InvGroup);
  //   const data = allItems[0]; // ใช้ข้อมูลหลักจากรายการแรก

  //   if (uniqueCONums.indexOf(conum) > 0) doc.addPage();

  //   let filledHtml = templateHtml;

  //   // ✅ แปลง TotalPrice เป็นตัวหนังสือ
  //   const totalPriceText = numberToThaiText(data.TotalPrice);
  //   filledHtml = filledHtml.replace(/{{TotalPriceText}}/g, totalPriceText);

  //   // ✅ สร้าง BranchInfo
  //   const rawBranchId = data.BranchIdCust;
  //   // const rawBranchName = data.BranchNameCust;
  //    const rawBranchName = "สำนักงานใหญ่";

  //   let branchInfo = "";
  //   if (isValidValue(rawBranchId) && isValidValue(rawBranchName)) {
  //     const paddedBranchId = String(rawBranchId).padStart(5, "0");
  //     branchInfo = ` สาขาที่ : ${paddedBranchId} ${rawBranchName}`;
  //   }
  //   data.BranchInfo = branchInfo;

  //   // ✅ แยกสินค้าหลัก/ของแถมตาม CONum
  //   // const allItems = excelData.filter(row => row.CONum === data.CONum);
  //   let mainProducts = [];
  //   let freeProducts = [];

  //   if (allItems.length === 1) {
  //     // กรณีมีแค่ 1 รายการ → ให้เป็นสินค้าหลัก
  //     mainProducts = allItems;
  //     freeProducts = [];
  //   } else {
  //     // กรณีมีหลายรายการ → แยกสินค้าหลัก/ของแถมตาม PromoCode
  //     mainProducts = allItems.filter(
  //       (item) => !!item.PromoCode && item.PromoCode.toString().trim() !== ""
  //     );
  //     freeProducts = allItems.filter(
  //       (item) => !item.PromoCode || item.PromoCode.toString().trim() === ""
  //     );
  //   }

  const uniqueGroups = [
    ...new Set(excelData.map(row => `${row.CONum}__${row.Invoice_Group}`))
  ];

  for (const groupKey of uniqueGroups) {
    const [conum, invoiceGroup] = groupKey.split("__");

    const allItems = excelData.filter(
      row => row.CONum === conum && row.Invoice_Group === invoiceGroup
    );

    if (allItems.length === 0) {
    console.warn(`ไม่พบข้อมูลสำหรับ CONum=${conum}, Invoice_Group=${invoiceGroup}`);
    continue; // ข้ามไปกลุ่มถัดไป
    }

    



    const data = allItems[0];

  if (uniqueGroups.indexOf(groupKey) > 0) doc.addPage();

  let filledHtml = templateHtml;

  filledHtml = filledHtml.replace(/{{dayInv}}/g, String(data.dayInv).padStart(2, '0'));
  filledHtml = filledHtml.replace(/{{monthInv}}/g, String(data.monthInv).padStart(2, '0'));

  // แปลง TotalPrice เป็นตัวหนังสือ
  const totalPriceText = numberToThaiText(data.TotalPrice);
  filledHtml = filledHtml.replace(/{{TotalPriceText}}/g, totalPriceText);

  // สร้าง BranchInfo
  const rawBranchId = data.BranchIdCust;
  const rawBranchName = "สำนักงานใหญ่";

  let branchInfo = "";
  if (isValidValue(rawBranchId) && isValidValue(rawBranchName)) {
    const paddedBranchId = String(rawBranchId).padStart(5, "0");
    branchInfo = ` สาขาที่ : ${paddedBranchId} ${rawBranchName}`;
  }
  data.BranchInfo = branchInfo;

  let mainProducts = [];
  let freeProducts = [];

  // ตรวจว่ามีสินค้าที่ผูกกับ PromoCode หรือไม่
  const hasPromo = allItems.some(item => !!item.PromoCode && item.PromoCode.toString().trim() !== "");

  if (hasPromo) {
    // กรณีมีของแถม
    mainProducts = allItems.filter(
      item => !!item.PromoCode && item.PromoCode.toString().trim() !== ""
    );
    freeProducts = allItems.filter(
      item => !item.PromoCode || item.PromoCode.toString().trim() === ""
    );
  } else {
    // กรณีไม่มีของแถม
    // เลือกรายการที่มี SeriNum ซ้ำกันและไม่มี PromoCode เป็นรายการหลัก
    mainProducts = allItems.filter(item => !item.PromoCode || item.PromoCode.toString().trim() === "");

    // ไม่มีของแถมเพราะไม่มี PromoCode
    freeProducts = [];
  }

  // แสดงของแถม (ถ้ามี)
  const freeProductNames = freeProducts.map(item => item.ProductList).join(", ");
  filledHtml = filledHtml.replace(/{{FreeProductList}}/g, freeProductNames);

  if (mainProducts.length > 0) {
    // สร้าง HTML แถวสำหรับแต่ละสินค้า
    const mainProductRows = mainProducts.map(item => {
      let qty = parseFloat(item.qtyItem) || 0;
      let qtyRounded = (qty % 1 >= 0.5) ? Math.ceil(qty) : Math.floor(qty);

      return `
        <tr>
          <td class="text-left">${item.ProductList}</td>
          <td>${qtyRounded}</td>
          <td class="text-right">${formatCurrency(item.ItemPrice)}</td>
          <td class="text-right">${formatCurrency(item.qtytotalItem)}</td>
        </tr>
      `;
    }).join("");

    filledHtml = filledHtml.replace(/{{ProductRows}}/g, mainProductRows);
  } else {
    filledHtml = filledHtml.replace(/{{ProductRows}}/g, "");
  }


  const totalVat = mainProducts.length > 0 ? (mainProducts[0].Vat || 0) : 0;
  // เช็คเงื่อนไขภาษี
  let vatLabel = "";
  let vatValue = "";

  if (totalVat !== 0) {
    vatLabel = "ภาษีมูลค่าเพิ่ม 7%";
    vatValue = formatCurrency(totalVat);
  } else {
    vatLabel = "ภาษีมูลค่าเพิ่ม NON VAT";
    vatValue = formatCurrency(0);
  }

  // แทนค่าใน HTML
  filledHtml = filledHtml.replace(/{{Vat_Label}}/g, vatLabel);
  filledHtml = filledHtml.replace(/{{Vat_Formatted}}/g, vatValue);
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
      if (typeof val === "string" && val.trim().toLowerCase() === "null")
        return "";
      if (typeof val === "string" && val.trim().toLowerCase() === "undefined")
        return "";
      if (typeof val === "string" && val.trim().toLowerCase() === "-")
        return "";
      if (typeof val === "string" && val.trim().toLowerCase() === "")
        return "";
      if (typeof val === "string" && val.trim().toLowerCase() === " ")
        return "";
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

    //TimeStamp
    // function formatDate(dateStr) {
    //   if (!dateStr) return "";
    //   const date = new Date(dateStr);
    //   const localDate = new Date(date.getTime() + 7 * 60 * 60 * 1000); // ปรับเป็นเวลาไทย
    //   const day = String(localDate.getDate()).padStart(2, "0");
    //   const month = String(localDate.getMonth() + 1).padStart(2, "0");
    //   const year = localDate.getFullYear();
    //   return `${day}/${month}/${year}`;
    // }

   // 05/27/2025 9:17 AM format
    // function formatDateOnly(datetimeStr) {
    // if (!datetimeStr) return "";
    // const datePart = datetimeStr.split(' ')[0];
    // const [month, day, year] = datePart.split('/');
    // if (!month || !day || !year) return "";
    // return `${day}/${month}/${year}`;
    // }

    //Serial
  //   function excelSerialToDate(serial) {
  //   const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  //   const msPerDay = 24 * 60 * 60 * 1000;
  //   const date = new Date(excelEpoch.getTime() + serial * msPerDay);

  //   const day = String(date.getUTCDate()).padStart(2, '0');
  //   const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  //   const year = date.getUTCFullYear();
  //   const hour = String(date.getUTCHours()).padStart(2, '0');
  //   const minute = String(date.getUTCMinutes()).padStart(2, '0');

   
  //   return `${day}/${month}/${year} `;
  // }

  //format DD/MM/YYYY
 


   
    const priceFields = [
      "UnitPrice",
      "Price",
      "Disc",
      "SumPrice",
      "Vat",
      "TotalPrice",
    ];

    Object.entries(data).forEach(([key, value]) => {
      let cleanValue = safeValue(value);
      const regexBase = new RegExp(`{{${key}}}`, "g");
      filledHtml = filledHtml.replace(regexBase, cleanValue);

      // เพิ่มเติม: ถ้าเป็นจำนวนเงิน แปลงเป็นรูปแบบไทย + ข้อความ
      if (priceFields.includes(key)) {
        const formattedValue = formatCurrency(cleanValue);
        filledHtml = filledHtml.replace(
          new RegExp(`{{${key}_Formatted}}`, "g"),
          formattedValue
        );

        const textValue = numberToThaiText(cleanValue);
        filledHtml = filledHtml.replace(
          new RegExp(`{{${key}_Text}}`, "g"),
          textValue
        );
      }
      if (!safeValue(data.Model)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']model-line["'][^>]*>.*?<\/div>\s*/g,
          "แบบรถ : "
        );
      }

      if (!safeValue(data.Color)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']color-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="color-line">สี : </div>'
        );
      }

      if (!safeValue(data.CertNum)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']CertNum-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="CertNum-line">หมายเลขเครื่อง : </div>'
        );
      }

      if (!safeValue(data.SeriNum)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']SeriNum-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="SeriNum-line">หมายเลขตัวถัง : </div>'
        );
      }

    

       if (!safeValue(data.ContractNum)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']ContractNum-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="ContractNum-line">สัญญาเลขที่ : </div>'
        );
      }

    if (!safeValue(data.Comment)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']Comment-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="Comment-line">หมายเหตุ : </div>'
        );
      }
 
      if (!safeValue(data.RefCust)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']RefCust-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="RefCust-line">&nbsp;</div>'
        );
      }

      if (!safeValue(data.TaxCust)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']TaxCust-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="TaxCust-line">เลขประจำตัวผู้เสียภาษีอากร : </div>'
        );
      }

      if (!safeValue(data.PayType)) {
        filledHtml = filledHtml.replace(
          /<div[^>]*class=["']PayType-line["'][^>]*>.*?<\/div>\s*/g,
          '<div class="PayType-line">ชำระโดย : </div>'
        );
      }

  // filledHtml = filledHtml.replace(/{{CreateDate}}/g, excelSerialToDate(data.CreateDate));

   //filledHtml = filledHtml.replace(/{{CreateDate}}/g, formatDateOnly(data.CreateDate));

    });

    const container = document.createElement("div");
    container.innerHTML = filledHtml;
    container.style.position = "absolute";
    container.style.left = "-9999px";
    document.body.appendChild(container);


    const canvas = await html2canvas(container, {
    scale: 3
    });


    // const imgData = canvas.toDataURL("image/png");
    const imgData = canvas.toDataURL("image/jpeg");
    const margin = 5; // เว้นขอบ 5mm รอบด้าน
    const pdfWidth = doc.internal.pageSize.getWidth() - margin * 2;
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    doc.addImage(imgData, "JPEG", margin, margin, pdfWidth, pdfHeight);

    container.remove();
    console.log("HTML Length (innerHTML):", container.innerHTML.length);
  }
  

 
  // ✅ เปิด PDF หลังจาก render เสร็จทุกหน้า
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
