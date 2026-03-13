const fileInput = document.getElementById("fileInput");
const sessionCodeInput = document.getElementById("sessionCode");
const courseSheetSelect = document.getElementById("courseSheetSelect");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("status");
const summaryPanel = document.getElementById("summary");
const resultTable = document.getElementById("resultTable");

if (fileInput && sessionCodeInput && courseSheetSelect && downloadBtn && clearBtn && statusText && summaryPanel && resultTable) {
  const errorSheetKey = "__ERRORS__";
  const outputHeaders = ["CERTIFICATE NO", "NAME", "ENROLLMENT NO", "GRADE", "COURSE NAME", "SC CODE", "FILE NAME"];
  let extractedRows = [];
  let courseSheetRows = new Map();
  let errorRows = [];
  let activeCourseCode = "";
  let loadedFileCount = 0;
  let uploadHistory = [];

  fileInput.addEventListener("change", async (event) => {
    const files = Array.from(event.target.files || []);

    if (!files.length) {
      return;
    }

    const sessionCode = normalizeSessionCode(sessionCodeInput.value);
    if (!sessionCode) {
      statusText.textContent = "Please enter Term Code first (example: JUNE25 or DEC25).";
      fileInput.value = "";
      return;
    }

    let batchAdded = 0;
    let batchSucceeded = 0;
    let batchSkipped = 0;

    for (const file of files) {
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        const rows = extractRowsFromWorkbook(workbook, file.name, sessionCode);

        if (!rows.length) {
          batchSkipped += 1;
          uploadHistory.unshift({
            fileName: file.name,
            status: "Skipped",
            detail: "No usable records found",
          });
          continue;
        }

        extractedRows.push(...rows);
        batchAdded += rows.length;
        batchSucceeded += 1;
        loadedFileCount += 1;

        uploadHistory.unshift({
          fileName: file.name,
          status: "Added",
          detail: `${rows.length} record(s) extracted`,
        });
      } catch (error) {
        batchSkipped += 1;
        uploadHistory.unshift({
          fileName: file.name,
          status: "Error",
          detail: "Could not process this file",
        });
        console.error(error);
      }
    }

    uploadHistory = uploadHistory.slice(0, 20);
    rebuildPreparedSheets();
    renderTable();
    renderSummary();

    downloadBtn.disabled = extractedRows.length === 0;
    clearBtn.disabled = extractedRows.length === 0;

    if (!extractedRows.length) {
      statusText.textContent = "No usable data extracted yet.";
    } else {
      statusText.textContent = `Added ${batchSucceeded} file(s), skipped ${batchSkipped}. Added ${batchAdded} record(s). Total records: ${extractedRows.length}.`;
    }

    fileInput.value = "";
  });

  courseSheetSelect.addEventListener("change", () => {
    activeCourseCode = normalizeText(courseSheetSelect.value);
    renderTable();
    renderSummary();
  });

  downloadBtn.addEventListener("click", () => {
    if (!extractedRows.length) {
      return;
    }

    const outputBook = XLSX.utils.book_new();
    const usedSheetNames = new Set();
    const sortedCourseCodes = Array.from(courseSheetRows.keys()).sort((a, b) => a.localeCompare(b));

    sortedCourseCodes.forEach((courseCode) => {
      const rows = courseSheetRows.get(courseCode) || [];
      const sheetName = safeUniqueSheetName(courseCode || "UNKNOWN", usedSheetNames);
      const aoa = [outputHeaders, ...rows];
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(outputBook, worksheet, sheetName);
    });

    if (errorRows.length) {
      const errorSheetName = safeUniqueSheetName("errors", usedSheetNames);
      const errorSheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...errorRows]);
      XLSX.utils.book_append_sheet(outputBook, errorSheet, errorSheetName);
    }

    XLSX.writeFile(outputBook, "certificate_data.xlsx");
  });

  clearBtn.addEventListener("click", () => {
    extractedRows = [];
    courseSheetRows = new Map();
    errorRows = [];
    activeCourseCode = "";
    loadedFileCount = 0;
    uploadHistory = [];
    renderCourseSheetSelect();
    renderTable();
    renderSummary();
    downloadBtn.disabled = true;
    clearBtn.disabled = true;
    statusText.textContent = "All extracted data cleared.";
  });

  function extractRowsFromWorkbook(workbook, sourceFileName, sessionCode) {
    const rows = [];
    let serial = 1;

    workbook.SheetNames.forEach((sheetName) => {
      const yearType = detectSheetYearType(sheetName);
      if (yearType === "first") {
        return;
      }

      const worksheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      const headerRowIndex = findHeaderRowIndex(aoa);

      if (headerRowIndex < 0) {
        return;
      }

      const headerRow = aoa[headerRowIndex] || [];
      const gradeCol = findGradeColumnIndex(headerRow);
      const courseName = extractCourseName(aoa, headerRowIndex);
      const scCode = extractScCode(aoa, headerRowIndex);
      const sheetCourseCode = extractCourseCodeFromSheetName(sheetName);
      const courseCode = sheetCourseCode || "UNKNOWN";
      const hasPaperRow = isPaperRow(aoa[headerRowIndex + 1] || []);
      const startRow = hasPaperRow ? headerRowIndex + 2 : headerRowIndex + 1;

      let sheetStarted = false;

      for (let r = startRow; r < aoa.length; r += 1) {
        const row = aoa[r] || [];
        const name = normalizeText(row[0]);
        const enrollmentRaw = normalizeText(row[1]);
        const firstTwoCols = `${name} ${enrollmentRaw}`.toUpperCase();

        if (
          firstTwoCols.includes("DATE OF COMMENCEMENT") ||
          firstTwoCols.includes("DATE OF COMPLETION") ||
          firstTwoCols.includes("DATE OF PUBLICATION") ||
          firstTwoCols.includes("DATE OF ISSUE") ||
          firstTwoCols.includes("PAPER CODE")
        ) {
          break;
        }

        if (!name && !enrollmentRaw) {
          if (sheetStarted) {
            break;
          }
          continue;
        }

        if (!name || !enrollmentRaw) {
          continue;
        }

        sheetStarted = true;
        const grade = gradeCol >= 0 ? normalizeText(row[gradeCol]) : "";
        const certificateNo = buildCertificateNo(courseCode, sessionCode, scCode, serial);

        rows.push([
          certificateNo,
          name,
          normalizeEnrollment(enrollmentRaw),
          grade,
          courseName,
          scCode,
          sourceFileName,
        ]);

        serial += 1;
      }
    });

    return rows;
  }

  function groupRowsByCourseCode(rows) {
    const grouped = new Map();

    rows.forEach((row) => {
      const certificateNo = normalizeText(row[0]);
      const courseCode = extractCourseCodeFromCertificateNo(certificateNo) || "UNKNOWN";

      if (!grouped.has(courseCode)) {
        grouped.set(courseCode, []);
      }

      grouped.get(courseCode).push(row);
    });

    return grouped;
  }

  function splitRowsByGrade(rows) {
    const validRows = [];
    const missingGradeRows = [];

    rows.forEach((row) => {
      const grade = normalizeText(row[3]).toUpperCase();
      if (grade && grade !== "--") {
        validRows.push(row);
      } else {
        missingGradeRows.push(row);
      }
    });

    return { validRows, missingGradeRows };
  }

  function extractCourseCodeFromCertificateNo(certificateNo) {
    const text = normalizeText(certificateNo).toUpperCase();
    if (!text) {
      return "";
    }

    const match = text.match(/^C-([^/]+)/);
    return match ? match[1] : "";
  }

  function safeUniqueSheetName(rawName, usedNames) {
    const cleaned = normalizeText(rawName)
      .replace(/[\\/?*\[\]:]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const base = (cleaned || "COURSE").slice(0, 31);
    let candidate = base;
    let count = 1;

    while (usedNames.has(candidate.toLowerCase())) {
      const suffix = `_${count}`;
      const maxBase = Math.max(1, 31 - suffix.length);
      candidate = `${base.slice(0, maxBase)}${suffix}`;
      count += 1;
    }

    usedNames.add(candidate.toLowerCase());
    return candidate;
  }

  function buildCertificateNo(courseCode, sessionCode, scCode, serial) {
    const cleanCourse = normalizeText(courseCode).toUpperCase() || "UNKNOWN";
    const cleanSession = normalizeSessionCode(sessionCode) || "SESSION";
    const cleanSc = normalizeText(scCode).toUpperCase() || "NA";
    const sequence = String(serial).padStart(3, "0");

    return `C-${cleanCourse}/${cleanSession}/${cleanSc}/${sequence}`;
  }

  function detectSheetYearType(sheetName) {
    const text = normalizeText(sheetName).toUpperCase();
    if (!text) {
      return "unknown";
    }

    if (/\b(2ND|SECOND)\s*YEAR\b/.test(text)) {
      return "second";
    }

    if (/\b(1ST|FIRST)\s*YEAR\b/.test(text)) {
      return "first";
    }

    return "unknown";
  }

  function extractCourseCodeFromSheetName(sheetName) {
    const cleaned = normalizeText(sheetName)
      .toUpperCase()
      .replace(/\(\s*(1ST|FIRST|2ND|SECOND)\s*YEAR\s*\)/g, " ")
      .replace(/\b(1ST|FIRST|2ND|SECOND)\s*YEAR\b/g, " ")
      .trim();

    if (!cleaned) {
      return "";
    }

    const tokens = cleaned.split(/[^A-Z0-9-]+/).filter((token) => token);
    for (const token of tokens) {
      if (/^[A-Z]{3,}(-[A-Z])?$/.test(token)) {
        return token;
      }
    }

    return "";
  }

  function findHeaderRowIndex(rows) {
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const col1 = normalizeText(row[0]).toUpperCase();
      const col2 = normalizeText(row[1]).toUpperCase();

      if (col1.includes("NAME") && col2.includes("ENROLLMENT")) {
        return i;
      }
    }

    return -1;
  }

  function findGradeColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
      const label = normalizeText(headerRow[i]).toUpperCase();
      if (label.includes("GRADE")) {
        return i;
      }
    }

    return -1;
  }

  function extractCourseName(rows, headerRowIndex) {
    for (let i = 0; i <= headerRowIndex; i += 1) {
      const row = rows[i] || [];
      const line = row
        .map((cell) => normalizeText(cell))
        .filter((cell) => cell)
        .join(" ");

      const upper = line.toUpperCase();
      const marker = "COURSE -";
      const pos = upper.indexOf(marker);
      if (pos >= 0) {
        return line.slice(pos + marker.length).trim();
      }
    }

    return "";
  }

  function extractScCode(rows, headerRowIndex) {
    for (let i = 0; i <= headerRowIndex; i += 1) {
      const row = rows[i] || [];
      const line = row
        .map((cell) => normalizeText(cell))
        .filter((cell) => cell)
        .join(" ");

      const upper = line.toUpperCase();
      if (!upper.includes("SC WITH CODE")) {
        continue;
      }

      const codeMatch = line.match(/\(([A-Za-z]-\d+)\)\s*$/);
      if (codeMatch) {
        return normalizeText(codeMatch[1]).toUpperCase();
      }

      const looseMatch = line.match(/\b([A-Za-z]-\d+)\b/);
      if (looseMatch) {
        return normalizeText(looseMatch[1]).toUpperCase();
      }
    }

    return "";
  }

  function isPaperRow(row) {
    return row.some((cell) => normalizeText(cell).toUpperCase().includes("PAPER"));
  }

  function rebuildPreparedSheets() {
    const split = splitRowsByGrade(extractedRows);
    courseSheetRows = groupRowsByCourseCode(split.validRows);
    errorRows = split.missingGradeRows;
    const availableCourseCodes = Array.from(courseSheetRows.keys()).sort((a, b) => a.localeCompare(b));

    if (activeCourseCode === errorSheetKey && errorRows.length) {
      renderCourseSheetSelect();
      return;
    }

    if (!activeCourseCode || !courseSheetRows.has(activeCourseCode)) {
      if (availableCourseCodes.length) {
        activeCourseCode = availableCourseCodes[0];
      } else if (errorRows.length) {
        activeCourseCode = errorSheetKey;
      } else {
        activeCourseCode = "";
      }
    }

    renderCourseSheetSelect();
  }

  function renderCourseSheetSelect() {
    courseSheetSelect.innerHTML = "";

    if (!courseSheetRows.size && !errorRows.length) {
      const emptyOption = document.createElement("option");
      emptyOption.value = "";
      emptyOption.textContent = "No course sheets";
      courseSheetSelect.appendChild(emptyOption);
      courseSheetSelect.disabled = true;
      return;
    }

    const courseCodes = Array.from(courseSheetRows.keys()).sort((a, b) => a.localeCompare(b));
    courseCodes.forEach((courseCode) => {
      const option = document.createElement("option");
      option.value = courseCode;
      option.textContent = `${courseCode} (${(courseSheetRows.get(courseCode) || []).length})`;
      if (courseCode === activeCourseCode) {
        option.selected = true;
      }
      courseSheetSelect.appendChild(option);
    });

    if (errorRows.length) {
      const errorOption = document.createElement("option");
      errorOption.value = errorSheetKey;
      errorOption.textContent = `errors (${errorRows.length})`;
      if (activeCourseCode === errorSheetKey) {
        errorOption.selected = true;
      }
      courseSheetSelect.appendChild(errorOption);
    }

    courseSheetSelect.disabled = false;
  }

  function renderTable() {
    if (!extractedRows.length || (!courseSheetRows.size && !errorRows.length)) {
      resultTable.innerHTML = "";
      return;
    }

    const rowsToRender = activeCourseCode === errorSheetKey
      ? errorRows
      : ((activeCourseCode && courseSheetRows.get(activeCourseCode)) || []);
    if (!rowsToRender.length) {
      resultTable.innerHTML = "";
      return;
    }

    const tableHead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    outputHeaders.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headerRow.appendChild(th);
    });

    tableHead.appendChild(headerRow);

    const tableBody = document.createElement("tbody");
    rowsToRender.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = normalizeText(value);
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });

    resultTable.replaceChildren(tableHead, tableBody);
  }

  function renderSummary() {
    if (!extractedRows.length) {
      summaryPanel.innerHTML = "";
      return;
    }

    const recentUploads = uploadHistory
      .slice(0, 8)
      .map((item) => `<li><strong>${escapeHtml(item.status)}:</strong> ${escapeHtml(item.fileName)} - ${escapeHtml(item.detail)}</li>`)
      .join("");

    summaryPanel.innerHTML = `
      <h2>Extraction Summary</h2>
      <ul>
        <li><strong>Loaded files:</strong> ${loadedFileCount}</li>
        <li><strong>Total records:</strong> ${extractedRows.length}</li>
        <li><strong>Course sheets prepared:</strong> ${courseSheetRows.size}</li>
        <li><strong>Error rows (missing grade):</strong> ${errorRows.length}</li>
        <li><strong>Viewing sheet:</strong> ${escapeHtml(activeCourseCode === errorSheetKey ? "errors" : (activeCourseCode || "None"))}</li>
      </ul>
      <h2>Recent Uploads</h2>
      <ul>${recentUploads || "<li>No uploads yet.</li>"}</ul>
    `;
  }

  function normalizeText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function normalizeSessionCode(value) {
    return normalizeText(value).toUpperCase().replace(/\s+/g, "");
  }

  function normalizeEnrollment(value) {
    const text = normalizeText(value).replace(/\s+/g, "");

    if (/^\d+\.0+$/.test(text)) {
      return text.replace(/\.0+$/, "");
    }

    return text;
  }

  function escapeHtml(text) {
    return normalizeText(text)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }
}
