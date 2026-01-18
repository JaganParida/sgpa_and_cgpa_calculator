const validGradePoints = { O: 10, E: 9, A: 8, B: 7, C: 6, D: 5, F: 2, M: 0 };
let workbookData = [];

function parseCredit(val) {
  if (!val) return 0;
  return val
    .toString()
    .split("+")
    .reduce((acc, curr) => acc + parseFloat(curr || 0), 0);
}

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    workbookData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    alert("✅ Excel Data Synced Successfully.");
  };
  reader.readAsArrayBuffer(file);
}

function calculateSGPAFromSheet() {
  const regNo = document.getElementById("regno-input").value.trim();
  const sem = document.getElementById("semester-number").value;
  const reportDiv = document.getElementById("report-output");
  const downloadBtn = document.getElementById("download-btn");

  if (!regNo || workbookData.length === 0) {
    alert("Please upload file and enter Registration Number.");
    return;
  }

  const studentRows = workbookData.filter(
    (row) => String(row["Reg_No"]).trim() === regNo,
  );
  if (studentRows.length === 0) {
    reportDiv.innerHTML =
      "<div style='color:red; text-align:center; padding:20px; font-weight:bold;'>❌ No data found for this Reg No.</div>";
    downloadBtn.style.display = "none";
    return;
  }

  let totalWeightedPoints = 0;
  let totalSGPACredits = 0;
  let creditsCleared = 0;
  let sGradeFound = false;
  const studentName = studentRows[0]["Name"];

  const tableRowsHTML = studentRows
    .map((row, index) => {
      const grade = String(row["Grade"]).trim().toUpperCase();
      const credit = parseCredit(row["Credits"]);

      if (grade === "S") sGradeFound = true;

      if (grade !== "R" && grade !== "S") {
        const points =
          validGradePoints[grade] !== undefined ? validGradePoints[grade] : 0;
        totalWeightedPoints += points * credit;
        totalSGPACredits += credit;
        if (grade !== "F" && grade !== "M") {
          creditsCleared += credit;
        }
      }

      return `
            <tr>
                <td>${index + 1}</td>
                <td>${row["Subject_Code"] || ""}</td>
                <td style="text-align:left;">${row["Subject_Name"] || ""}</td>
                <td>${row["Type"] || ""}</td>
                <td>${credit.toFixed(1)}</td>
                <td>${grade}</td>
            </tr>`;
    })
    .join("");

  const finalSGPA =
    totalSGPACredits > 0
      ? (totalWeightedPoints / totalSGPACredits).toFixed(2)
      : "0.00";
  const sWarning = sGradeFound
    ? `<div style="color: #856404; background-color: #fff3cd; border: 1px solid #ffeeba; padding: 10px; margin-bottom: 15px; border-radius: 5px; text-align: center; font-weight: bold;">⚠️ There is any logical issue due to S grade point please recheck like this.</div>`
    : "";

  reportDiv.innerHTML = `
        <div id="grade-sheet-container" class="grade-sheet">
            <div class="sheet-header">
                <h2>Centurion University of Technology and Management</h2>
                <h3 class="sheet-title">Semester Grade Sheet</h3>
            </div>
            <div class="student-meta-grid">
                <div class="meta-item"><span>Regd. No:</span> <strong>${regNo}</strong></div>
                <div class="meta-item"><span>Name:</span> <strong>${studentName}</strong></div>
                <div class="meta-item"><span>Semester:</span> <strong>Sem ${sem}</strong></div>
            </div>
            ${sWarning}
            <div class="table-responsive">
                <table class="result-table">
                    <thead>
                        <tr>
                            <th>SL.NO</th>
                            <th>SUB.CODE</th>
                            <th>SUBJECT</th>
                            <th>TYPE</th>
                            <th>CREDIT</th>
                            <th>GRADE</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRowsHTML}
                    </tbody>
                </table>
            </div>
            <div class="sheet-footer">
                <div class="footer-info">
                    <p>Credits Cleared: ${creditsCleared}</p>
                    <p>Generated On: ${new Date().toLocaleDateString("en-GB")}</p>
                </div>
                <div class="sgpa-badge">
                    <span class="label">SGPA</span>
                    <span class="value">${finalSGPA}</span>
                </div>
            </div>
        </div>
    `;

  document.getElementById("sgpa-result").innerText = finalSGPA;
  downloadBtn.style.display = "block";
}

function downloadReport() {
  const element = document.getElementById("grade-sheet-container");
  const originalWidth = element.style.width;
  element.style.width = "900px";

  html2canvas(element, {
    scale: 3.5,
    useCORS: true,
    windowWidth: 1200,
  }).then((canvas) => {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "mm", "a4");
    const imgData = canvas.toDataURL("image/jpeg", 1.0);
    const width = pdf.internal.pageSize.getWidth();
    const height = (canvas.height * width) / canvas.width;

    pdf.addImage(imgData, "JPEG", 0, 0, width, height, undefined, "FAST");
    pdf.save(`SGPA_Report_${document.getElementById("regno-input").value}.pdf`);
    element.style.width = originalWidth;
  });
}

function addCgpaRow() {
  const div = document.createElement("div");
  div.className = "cgpa-row";
  div.innerHTML = `<input type="number" class="cgpa-sgpa" placeholder="Enter SGPA" step="0.01" style="width:100%; margin-top:10px;"/>`;
  document.getElementById("cgpa-entries").appendChild(div);
}

function calculateCGPA() {
  const inputs = document.querySelectorAll(".cgpa-sgpa");
  let sum = 0,
    count = 0;
  inputs.forEach((i) => {
    if (i.value) {
      sum += parseFloat(i.value);
      count++;
    }
  });
  document.getElementById("cgpa-calc-res").innerText =
    "Resulting CGPA: " + (count > 0 ? (sum / count).toFixed(2) : "--");
}

document
  .getElementById("excel-file")
  .addEventListener("change", handleFileUpload);
document
  .getElementById("calculate-btn")
  .addEventListener("click", calculateSGPAFromSheet);
document
  .getElementById("download-btn")
  .addEventListener("click", downloadReport);
