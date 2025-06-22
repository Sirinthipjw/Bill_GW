let excelData = [];
const viewPdfBtn = document.getElementById("viewPDF");

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

    // ✅ แทนที่ข้อมูลใน Template
    let filledHtml = templateHtml;
    Object.entries(data).forEach(([key, value]) => {
      const regex = new RegExp(`{{${key}}}`, "g");
      filledHtml = filledHtml.replace(regex, value || "");
    });

    const container = document.createElement("div");
    container.innerHTML = filledHtml;
    container.style.position = "absolute";
    container.style.left = "-9999px";
    document.body.appendChild(container);
    
    const canvas = await html2canvas(container, { scale: 2 });
    const imgData = canvas.toDataURL("image/png");
    const pdfWidth = doc.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

    doc.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
    

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
