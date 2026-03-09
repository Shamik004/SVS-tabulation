const fileInput = document.getElementById("fileInput");
const sessionCodeInput = document.getElementById("sessionCode");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("status");
const summaryPanel = document.getElementById("summary");
const resultTable = document.getElementById("resultTable");

if (fileInput && sessionCodeInput && downloadBtn && clearBtn && statusText && summaryPanel && resultTable) {
  const outputHeaders = ["CERTIFICATE NO", "NAME", "ENROLLMENT NO", "GRADE", "COURSE NAME", "SC CODE", "FILE NAME"];
  let extractedRows = [];
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

  downloadBtn.addEventListener("click", () => {
    if (!extractedRows.length) {
      return;
    }

    const outputBook = XLSX.utils.book_new();
    const aoa = [outputHeaders, ...extractedRows];
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(outputBook, worksheet, "Certificates");
    XLSX.writeFile(outputBook, "certificate_data.xlsx");
  });

  clearBtn.addEventListener("click", () => {
    extractedRows = [];
    loadedFileCount = 0;
    uploadHistory = [];
    renderTable();
    renderSummary();
    downloadBtn.disabled = true;
    clearBtn.disabled = true;
    statusText.textContent = "All extracted data cleared.";
  });

  function extractRowsFromWorkbook(workbook, sourceFileName, sessionCode) {
    const rows = [];
    const fileCourseCode = extractCourseCodeFromFileName(sourceFileName);
    let serial = 1;

    workbook.SheetNames.forEach((sheetName) => {
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
      const courseCode = fileCourseCode || sheetCourseCode || "UNKNOWN";
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

  function buildCertificateNo(courseCode, sessionCode, scCode, serial) {
    const cleanCourse = normalizeText(courseCode).toUpperCase() || "UNKNOWN";
    const cleanSession = normalizeSessionCode(sessionCode) || "SESSION";
    const cleanSc = normalizeText(scCode).toUpperCase() || "NA";
    const sequence = String(serial).padStart(3, "0");

    return `C-${cleanCourse}/${cleanSession}/${cleanSc}/${sequence}`;
  }

  function extractCourseCodeFromFileName(fileName) {
    const base = normalizeText(fileName).toUpperCase().replace(/\.[^.]+$/, "");
    if (!base) {
      return "";
    }

    const strongMatch = base.match(/\b[A-Z]{3,}-[A-Z]\b/);
    if (strongMatch) {
      return strongMatch[0];
    }

    const tokens = base.split(/[^A-Z0-9-]+/).filter((token) => token);
    const stopWords = new Set(["TABULATION", "SHEET", "YEAR", "FIRST", "SECOND", "ALL", "DECEMBER", "JUNE", "MARKSHEET", "FILE", "PIMT", "KHATRA"]);

    for (const token of tokens) {
      if (stopWords.has(token)) {
        continue;
      }

      if (/^[A-Z]{4,}(-[A-Z])?$/.test(token)) {
        return token;
      }
    }

    return "";
  }

  function extractCourseCodeFromSheetName(sheetName) {
    const codeRaw = normalizeText(sheetName).split("(")[0].trim().replace(/[.\-\s_]+$/, "");
    const compact = codeRaw.replace(/\s+/g, "").toUpperCase();

    if (/^[A-Z]{4,}(-[A-Z])?$/.test(compact)) {
      return compact;
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

  function renderTable() {
    if (!extractedRows.length) {
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
    extractedRows.forEach((row) => {
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
