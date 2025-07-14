const gradePointMap = { O: 10, E: 9, A: 8, B: 7, C: 6, D: 5, F: 0 };
let workbookData = [];

function truncate(num, digits) {
  const factor = Math.pow(10, digits);
  return Math.floor(num * factor) / factor;
}

function parseCredit(rawCredit) {
  if (!rawCredit) return 0;
  return rawCredit
    .toString()
    .split("+")
    .reduce((sum, val) => sum + parseFloat(val || 0), 0);
}

function calculateSGPAFromSheet() {
  const regInput = document.getElementById("regno-input").value.trim();
  const semInput = document.getElementById("semester-number").value.trim();
  const resultsDiv = document.getElementById("subject-list");
  const sgpaDisplay = document.getElementById("sgpa-result");
  const feedbackEl = document.getElementById("sgpa-feedback");

  resultsDiv.innerHTML = "";
  sgpaDisplay.textContent = "SGPA: 0.00";
  feedbackEl.textContent = "⬅ Add grades to get feedback";

  if (!regInput || workbookData.length === 0) {
    alert("Please enter registration number and upload the Excel file.");
    return;
  }

  const studentRows = workbookData.filter((row) => {
    return String(row["Reg_No"]).trim() === regInput;
  });

  if (studentRows.length === 0) {
    resultsDiv.innerHTML = `<p class="error">❌ No data found for Reg No: <b>${regInput}</b></p>`;
    return;
  }

  const studentName = studentRows[0]["Name"] || "Unknown";
  let totalObtained = 0;
  let totalMax = 0;

  const rowsHTML = studentRows
    .map((sub, index) => {
      const grade = sub["Grade"]?.toString().trim().toUpperCase();
      const credit = parseCredit(sub["Credits"]);
      const subjectCode = sub["Subject_Code"] || "";
      const subjectName = sub["Subject_Name"] || "";
      const point = gradePointMap[grade];

      if (isNaN(point) || credit === 0) return "";

      const earned = point * credit;
      const max = 10 * credit;

      totalObtained += earned;
      totalMax += max;

      return `
      <tr>
        <td class="cell">${index + 1}</td>
        <td class="cell">${subjectCode}</td>
        <td class="cell">${subjectName}</td>
        <td class="cell">${credit}</td>
        <td class="cell">${grade}</td>
      </tr>
    `;
    })
    .filter(Boolean)
    .join("");

  const sgpa = totalMax ? truncate((totalObtained / totalMax) * 10, 4) : 0.0;

  resultsDiv.innerHTML = `
    <div class="info">
      <h2 style="text-align:center;">SGPA Report</h2>
      <p><strong>Name:</strong> ${studentName}</p>
      <p><strong>College:</strong> CUTM, BBSR</p>
      <p><strong>Semester:</strong> ${semInput}</p>
      <div style="overflow-x:auto;">
      <table class="modern-table" style="margin:auto; font-family:'Arial'; font-size:14px; border-collapse:collapse; border:1px solid #000;">
        <thead style="background:#f0f0f0;">
          <tr>
            <th class="cell">Sl No</th>
            <th class="cell">Subject Code</th>
            <th class="cell">Subject Name</th>
            <th class="cell">Credits</th>
            <th class="cell">Grade</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHTML}
        </tbody>
      </table>
      </div>
    </div>
  `;

  sgpaDisplay.textContent = `SGPA: ${sgpa}`;
  updateFeedback(sgpa);
}

function updateFeedback(sgpa) {
  const feedbackEl = document.getElementById("sgpa-feedback");
  if (isNaN(sgpa) || sgpa === 0) {
    feedbackEl.textContent = "⬅ Add grades to get feedback";
    return;
  }
  if (sgpa >= 9) feedbackEl.textContent = "🔥 Excellent! Keep it up!";
  else if (sgpa >= 8) feedbackEl.textContent = "✅ Great work, aim for 9+!";
  else if (sgpa >= 7)
    feedbackEl.textContent = "👍 Good, try to push a little higher.";
  else feedbackEl.textContent = "⚠️ You can improve. Focus next term.";
}

function handleFileUpload(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    workbookData = json;
    alert("✅ Excel data loaded. Now enter your registration number.");
  };
  reader.readAsArrayBuffer(file);
}

function calculateCGPA() {
  const sgpaInputs = document.querySelectorAll(".cgpa-sgpa");
  let total = 0,
    count = 0;
  sgpaInputs.forEach((input) => {
    const val = parseFloat(input.value);
    if (!isNaN(val)) {
      total += val;
      count++;
    }
  });
  const cgpa = count > 0 ? (total / count).toFixed(4) : "--";
  document.getElementById("cgpa-result").textContent = `CGPA: ${cgpa}`;
}

function addCgpaRow() {
  const container = document.getElementById("cgpa-entries");
  const row = document.createElement("div");
  row.className = "cgpa-row";
  row.innerHTML = `
    <label for="cgpa-semester">Semester:</label>
    <select class="cgpa-semester">
      <option value="">Select Semester</option>
      ${[1, 2, 3, 4, 5, 6, 7, 8]
        .map((i) => `<option value="${i}">Semester ${i}</option>`)
        .join("")}
    </select>
    <input type="number" class="cgpa-sgpa" placeholder="Enter SGPA" min="0" max="10" step="0.0001" />
  `;
  container.appendChild(row);
}

document.addEventListener("DOMContentLoaded", () => {
  document
    .getElementById("excel-file")
    ?.addEventListener("change", handleFileUpload);
  document
    .getElementById("calculate-btn")
    ?.addEventListener("click", calculateSGPAFromSheet);
  document
    .getElementById("calculate-cgpa-btn")
    ?.addEventListener("click", calculateCGPA);
  document
    .getElementById("download-btn")
    ?.addEventListener("click", downloadReportAsPDF);
});

function downloadReportAsPDF() {
  const resultsDiv = document.getElementById("subject-list");
  if (!resultsDiv) return;

  html2canvas(resultsDiv, {
    scale: 2,
    useCORS: true,
  }).then((canvas) => {
    const imgData = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "pt",
      format: "a4",
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const imgWidth = canvas.width / 2;
    const imgHeight = canvas.height / 2;
    const x = (pageWidth - imgWidth) / 2;

    pdf.setFont("Helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text("SGPA Report", pageWidth / 2, 40, { align: "center" });
    pdf.addImage(imgData, "PNG", x, 60, imgWidth, imgHeight);
    pdf.save("SGPA_Report.pdf");
  });
}
