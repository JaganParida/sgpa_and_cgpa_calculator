const validGradePoints = {
  O: 10,
  E: 9,
  A: 8,
  B: 7,
  C: 6,
  D: 5,
  R: 0,
  F: 0,
  M: 0,
  S: 0,
};
let workbookData = [];

function openModal() {
  document.getElementById("formula-modal").classList.add("open");
}
function closeModal() {
  document.getElementById("formula-modal").classList.remove("open");
}
document
  .getElementById("formula-modal")
  .addEventListener("click", function (e) {
    if (e.target === this) closeModal();
  });

function openExcelModal() {
  document.getElementById("excel-modal").classList.add("open");
}
function closeExcelModal() {
  document.getElementById("excel-modal").classList.remove("open");
}
document.getElementById("excel-modal").addEventListener("click", function (e) {
  if (e.target === this) closeExcelModal();
});

function toggleMenu() {
  const nav = document.getElementById("nav-menu");
  const icon = document.getElementById("menu-icon");
  nav.classList.toggle("active");

  if (nav.classList.contains("active")) {
    icon.classList.remove("ri-menu-3-line");
    icon.classList.add("ri-close-line");
  } else {
    icon.classList.remove("ri-close-line");
    icon.classList.add("ri-menu-3-line");
  }
}

function closeMenu() {
  const nav = document.getElementById("nav-menu");
  const icon = document.getElementById("menu-icon");
  nav.classList.remove("active");
  icon.classList.remove("ri-close-line");
  icon.classList.add("ri-menu-3-line");
}

function switchTab(tabId) {
  document.getElementById("sgpa-section").style.display = "none";
  document.getElementById("cgpa-section").style.display = "none";
  document.getElementById("tab-sgpa").classList.remove("active");
  document.getElementById("tab-cgpa").classList.remove("active");

  const section = document.getElementById(tabId + "-section");
  section.style.display = "block";
  document.getElementById("tab-" + tabId).classList.add("active");

  document
    .querySelectorAll(".nav-link")
    .forEach((el) => el.classList.remove("active"));
  document.getElementById("nav-" + tabId).classList.add("active");
}

document.getElementById("excel-file").addEventListener("change", function (e) {
  const fileName = e.target.files[0]
    ? e.target.files[0].name
    : "Click to upload .xlsx file";
  document.getElementById("file-name-display").innerText = fileName;

  if (e.target.files[0]) {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      workbookData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    };
    reader.readAsArrayBuffer(e.target.files[0]);
  }
});

document.getElementById("calculate-btn").addEventListener("click", function () {
  const regNo = document.getElementById("regno-input").value.trim();
  const sem = document.getElementById("semester-number").value;
  const reportDiv = document.getElementById("report-output");

  if (!regNo || workbookData.length === 0) {
    alert("Please upload Excel and enter Reg No.");
    return;
  }

  const studentRows = workbookData.filter(
    (row) => String(row["Reg_No"]).trim() === regNo,
  );

  if (studentRows.length === 0) {
    alert("No data found for this Reg No.");
    return;
  }

  let totalPoints = 0,
    totalCredits = 0,
    creditsCleared = 0;
  let backlogs = [];
  const studentName = studentRows[0]["Name"];
  const watermarkUrl = "cutm_text.jpg";

  const rowsHTML = studentRows
    .map((row, i) => {
      const grade = String(row["Grade"]).trim().toUpperCase();
      const credit = parseCredit(row["Credits"]);
      const subject = row["Subject_Name"] || "Unknown";
      const type = row["Type"] || "";

      if (["F", "M", "S"].includes(grade)) backlogs.push(subject);
      let points =
        validGradePoints[grade] !== undefined ? validGradePoints[grade] : 0;

      if (grade !== "R" && grade !== "S") {
        totalPoints += points * credit;
        totalCredits += credit;
        if (!["F", "M"].includes(grade)) creditsCleared += credit;
      }
      if (grade === "R") creditsCleared += credit;

      return `<tr>
            <td>${i + 1}</td>
            <td>${row["Subject_Code"] || ""}</td>
            <td style="text-align:left;">${subject}</td>
            <td>${type}</td>
            <td>${credit.toFixed(1)}</td>
            <td>${grade}</td>
        </tr>`;
    })
    .join("");

  const sgpa =
    totalCredits > 0 ? (totalPoints / totalCredits).toFixed(2) : "0.00";
  const feedback =
    backlogs.length > 0
      ? `<div class="feedback-box" style="border-color:#ef4444; background:#fef2f2; color:#b91c1c;"><strong>Backlogs:</strong> ${backlogs.join(", ")}</div>`
      : `<div class="feedback-box"><strong>Congratulations!</strong> You have passed all subjects.<br>Excellent performance! 🚀</div>`;

  reportDiv.innerHTML = `
        <div id="grade-sheet" class="grade-sheet">
            <img src="${watermarkUrl}" class="watermark" onerror="this.style.display='none'">
            <div class="sheet-header">
                <h2 class="univ-name">Centurion University of Technology and<br>Management</h2>
                <div class="doc-title">Semester Grade Sheet</div>
                <div class="header-line"></div>
            </div>
            <div style="display:grid; grid-template-columns:100px 1fr; margin-bottom:20px; font-size:14px; position:relative; z-index:2; text-align:left;">
                <div style="color:#444">Regd. No:</div> <div style="font-weight:bold">${regNo}</div>
                <div style="color:#444">Name:</div> <div style="font-weight:bold">${studentName}</div>
                <div style="color:#444">Semester:</div> <div style="font-weight:bold">Sem ${sem}</div>
            </div>
            <table class="result-table">
                <thead><tr><th>SL.NO</th><th>SUB.CODE</th><th>SUBJECT</th><th>TYPE</th><th>CREDIT</th><th>GRADE</th></tr></thead>
                <tbody>${rowsHTML}</tbody>
            </table>
            ${feedback}
            <div class="sheet-footer">
                <div style="text-align:left;">
                    <p style="margin:5px 0;">Credits Cleared: <strong>${creditsCleared}</strong></p>
                    <p style="margin:5px 0; color:#666; font-size:12px;">Generated On: ${new Date().toLocaleDateString("en-GB")}</p>
                </div>
                <div class="sgpa-box">
                    <span style="display:block; font-size:12px; letter-spacing:1px;">SGPA</span>
                    <span class="sgpa-value">${sgpa}</span>
                </div>
            </div>
        </div>`;

  document.getElementById("download-actions").style.display = "flex";
  reportDiv.scrollIntoView({ behavior: "smooth" });
});

function addCgpaRow() {
  const div = document.createElement("div");
  div.className = "cgpa-row";
  div.innerHTML = `
    <div class="input-with-icon">
        <i class="ri-hashtag"></i>
        <input type="number" class="cgpa-sgpa" placeholder="SGPA" step="0.01">
    </div>
    <div class="input-with-icon">
        <i class="ri-coin-line"></i>
        <input type="number" class="cgpa-credit" placeholder="Credits" step="0.5">
    </div>`;
  document.getElementById("cgpa-entries").appendChild(div);
}

function calculateCGPA() {
  const sgpas = document.querySelectorAll(".cgpa-sgpa");
  const credits = document.querySelectorAll(".cgpa-credit");
  let num = 0,
    den = 0;

  sgpas.forEach((inp, i) => {
    const s = parseFloat(inp.value);
    const c = parseFloat(credits[i].value);
    if (!isNaN(s) && !isNaN(c)) {
      num += s * c;
      den += c;
    }
  });

  document.getElementById("cgpa-result-value").innerText =
    den > 0 ? (num / den).toFixed(2) : "0.00";
}

function parseCredit(val) {
  if (!val) return 0;
  return val
    .toString()
    .split("+")
    .reduce((a, c) => a + parseFloat(c || 0), 0);
}

document.getElementById("download-btn").addEventListener("click", () => {
  const element = document.getElementById("grade-sheet");
  html2canvas(element, {
    scale: 2,
    onclone: (clonedDoc) => {
      const feedback = clonedDoc.querySelector(".feedback-box");
      if (feedback) feedback.style.display = "none";
    },
  }).then((canvas) => {
    const pdf = new window.jspdf.jsPDF("p", "mm", "a4");
    const imgWidth = 210;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    pdf.addImage(
      canvas.toDataURL("image/png"),
      "PNG",
      0,
      0,
      imgWidth,
      imgHeight,
    );
    pdf.save("GradeSheet.pdf");
  });
});

document.getElementById("download-photo-btn").addEventListener("click", () => {
  const element = document.getElementById("grade-sheet");
  html2canvas(element, {
    scale: 3,
    onclone: (clonedDoc) => {
      const feedback = clonedDoc.querySelector(".feedback-box");
      if (feedback) feedback.style.display = "none";
    },
  }).then((canvas) => {
    const a = document.createElement("a");
    a.download = "GradeSheet.jpg";
    a.href = canvas.toDataURL("image/jpeg");
    a.click();
  });
});
