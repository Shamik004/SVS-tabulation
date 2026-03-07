const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("status");
const summaryPanel = document.getElementById("summary");
const excelTable = document.getElementById("excelTable");

if (fileInput && sheetSelect && downloadBtn && clearBtn && statusText && summaryPanel && excelTable) {
  let processedSheets = {};
  let activeOutputSheet = "";
  let loadedFileCount = 0;
  let uploadHistory = [];

  fileInput.addEventListener("change", async (event) => {
    const files = Array.from(event.target.files || []);

    if (!files.length) {
      return;
    }

    persistCurrentSheetEdits();

    let batchAddedRows = 0;
    let batchSucceededFiles = 0;
    let batchSkippedFiles = 0;

    for (const file of files) {
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        const result = processWorkbook(workbook);

        if (!Object.keys(result.outputSheets).length) {
          batchSkippedFiles += 1;
          uploadHistory.unshift({
            fileName: file.name,
            status: "Skipped",
            detail: "No usable 2nd-year/lateral sheets found",
          });
          continue;
        }

        const merge = mergeIntoProcessedSheets(result.outputSheets, file.name);
        batchAddedRows += merge.addedRows;
        batchSucceededFiles += 1;
        loadedFileCount += 1;

        uploadHistory.unshift({
          fileName: file.name,
          status: "Added",
          detail: `${Object.keys(result.outputSheets).length} sheet(s), ${merge.addedRows} row(s)`,
        });

        merge.warnings.forEach((warning) => {
          uploadHistory.unshift({
            fileName: file.name,
            status: "Notice",
            detail: warning,
          });
        });
      } catch (error) {
        batchSkippedFiles += 1;
        uploadHistory.unshift({
          fileName: file.name,
          status: "Error",
          detail: "Could not process this file",
        });
        console.error(error);
      }
    }

    uploadHistory = uploadHistory.slice(0, 20);
    refreshSheetSelectorAndView();
    renderSummary();

    if (!Object.keys(processedSheets).length) {
      statusText.textContent = "No usable data has been loaded yet.";
    } else {
      statusText.textContent = `Added ${batchSucceededFiles} file(s), skipped ${batchSkippedFiles}. Added ${batchAddedRows} row(s). Total loaded files: ${loadedFileCount}.`;
    }

    fileInput.value = "";
  });

  sheetSelect.addEventListener("change", (event) => {
    persistCurrentSheetEdits();
    activeOutputSheet = event.target.value;
    renderSheet(activeOutputSheet);
  });

  downloadBtn.addEventListener("click", () => {
    if (!activeOutputSheet) {
      return;
    }

    persistCurrentSheetEdits();

    const outputBook = XLSX.utils.book_new();
    const usedSheetNames = new Set();

    Object.keys(processedSheets).forEach((sheetName) => {
      const rows = processedSheets[sheetName];
      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      const safeName = safeUniqueSheetName(sheetName, usedSheetNames);
      XLSX.utils.book_append_sheet(outputBook, worksheet, safeName);
    });

    XLSX.writeFile(outputBook, "combined_marksheet_data.xlsx");
  });

  clearBtn.addEventListener("click", () => {
    processedSheets = {};
    activeOutputSheet = "";
    loadedFileCount = 0;
    uploadHistory = [];
    resetView();
    statusText.textContent = "All accumulated data cleared.";
  });

  function processWorkbook(workbook) {
    const grouped = {};

    workbook.SheetNames.forEach((sheetName) => {
      const identity = parseSheetIdentity(sheetName);

      if (!identity) {
        return;
      }

      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      const parsedSheet = parseTabulationSheet(rows, identity);

      if (!parsedSheet.students.length) {
        return;
      }

      if (!grouped[identity.code]) {
        grouped[identity.code] = {
          first: null,
          second: null,
          lateral: [],
        };
      }

      if (identity.type === "first") {
        grouped[identity.code].first = parsedSheet;
      } else if (identity.type === "second") {
        grouped[identity.code].second = parsedSheet;
      } else if (identity.type === "lateral") {
        grouped[identity.code].lateral.push(parsedSheet);
      }
    });

    const outputSheets = {};

    Object.keys(grouped)
      .sort()
      .forEach((code) => {
        const group = grouped[code];

        if (group.second) {
          const merged = buildSecondYearSheet(code, group.first, group.second);
          outputSheets[`${code} 2ND YEAR`] = merged.rows;
        }

        if (group.lateral.length) {
          group.lateral.forEach((lateralSheet, idx) => {
            const lateral = buildLateralSheet(lateralSheet);
            const sheetName = group.lateral.length > 1 ? `${code} LATERAL ${idx + 1}` : `${code} LATERAL`;
            outputSheets[sheetName] = lateral.rows;
          });
        }
      });

    return { outputSheets };
  }

  function mergeIntoProcessedSheets(newSheets, fileName) {
    let addedRows = 0;
    const warnings = [];

    Object.entries(newSheets).forEach(([sheetName, rows]) => {
      if (!rows.length) {
        return;
      }

      const incoming = cloneRows(rows);

      if (!processedSheets[sheetName]) {
        processedSheets[sheetName] = incoming;
        addedRows += Math.max(0, incoming.length - 1);
        return;
      }

      const existingHeader = processedSheets[sheetName][0] || [];
      const incomingHeader = incoming[0] || [];

      if (headersMatch(existingHeader, incomingHeader)) {
        processedSheets[sheetName].push(...incoming.slice(1));
        addedRows += Math.max(0, incoming.length - 1);
        return;
      }

      const fallbackName = uniqueSheetKey(`${sheetName} (${shortFileTag(fileName)})`);
      processedSheets[fallbackName] = incoming;
      addedRows += Math.max(0, incoming.length - 1);
      warnings.push(`Header mismatch for ${sheetName}; stored as ${fallbackName}`);
    });

    return { addedRows, warnings };
  }

  function headersMatch(left, right) {
    if (left.length !== right.length) {
      return false;
    }

    for (let i = 0; i < left.length; i += 1) {
      if (normalizeText(left[i]) !== normalizeText(right[i])) {
        return false;
      }
    }

    return true;
  }

  function cloneRows(rows) {
    return rows.map((row) => [...row]);
  }

  function uniqueSheetKey(baseName) {
    if (!processedSheets[baseName]) {
      return baseName;
    }

    let index = 2;
    let candidate = `${baseName} ${index}`;

    while (processedSheets[candidate]) {
      index += 1;
      candidate = `${baseName} ${index}`;
    }

    return candidate;
  }

  function shortFileTag(fileName) {
    const base = fileName.replace(/\.[^.]+$/, "");
    return base.length > 16 ? `${base.slice(0, 16)}...` : base;
  }

  function safeUniqueSheetName(originalName, usedSheetNames) {
    const base = safeSheetName(originalName);

    if (!usedSheetNames.has(base)) {
      usedSheetNames.add(base);
      return base;
    }

    let index = 2;
    let candidate = safeSheetName(`${base} ${index}`);

    while (usedSheetNames.has(candidate)) {
      index += 1;
      candidate = safeSheetName(`${base} ${index}`);
    }

    usedSheetNames.add(candidate);
    return candidate;
  }

  function refreshSheetSelectorAndView() {
    const sheetNames = Object.keys(processedSheets);

    if (!sheetNames.length) {
      resetView();
      return;
    }

    if (!sheetNames.includes(activeOutputSheet)) {
      activeOutputSheet = sheetNames[0];
    }

    populateSheetList(sheetNames);
    sheetSelect.value = activeOutputSheet;
    renderSheet(activeOutputSheet);

    sheetSelect.disabled = false;
    downloadBtn.disabled = false;
    clearBtn.disabled = false;
  }

  function renderSummary() {
    const sheetNames = Object.keys(processedSheets);

    if (!sheetNames.length) {
      summaryPanel.innerHTML = "";
      return;
    }

    const totalRows = sheetNames.reduce((sum, sheetName) => {
      const rows = processedSheets[sheetName] || [];
      return sum + Math.max(0, rows.length - 1);
    }, 0);

    const sheetRows = sheetNames
      .map((sheetName) => {
        const rows = processedSheets[sheetName] || [];
        return `<li><strong>${escapeHtml(sheetName)}:</strong> ${Math.max(0, rows.length - 1)} row(s)</li>`;
      })
      .join("");

    const recentUploads = uploadHistory
      .slice(0, 8)
      .map((item) => `<li><strong>${escapeHtml(item.status)}:</strong> ${escapeHtml(item.fileName)} - ${escapeHtml(item.detail)}</li>`)
      .join("");

    summaryPanel.innerHTML = `
      <h2>Processing Summary</h2>
      <ul>
        <li><strong>Loaded files:</strong> ${loadedFileCount}</li>
        <li><strong>Output sheets:</strong> ${sheetNames.length}</li>
        <li><strong>Total rows:</strong> ${totalRows}</li>
      </ul>
      <h2>Rows Per Sheet</h2>
      <ul>${sheetRows}</ul>
      <h2>Recent Uploads</h2>
      <ul>${recentUploads || "<li>No uploads yet.</li>"}</ul>
    `;
  }

  function parseSheetIdentity(sheetName) {
    const upper = sheetName.toUpperCase();
    const codeRaw = sheetName.split("(")[0].trim().replace(/[.\-\s_]+$/, "");

    if (!codeRaw) {
      return null;
    }

    let type = "";

    if (upper.includes("LATERAL")) {
      type = "lateral";
    } else if (upper.includes("2ND") || upper.includes("SECOND")) {
      type = "second";
    } else if (upper.includes("1ST") || upper.includes("FIRST")) {
      type = "first";
    }

    if (!type) {
      return null;
    }

    return {
      code: codeRaw.replace(/\s+/g, ""),
      type,
    };
  }

  function parseTabulationSheet(rows, identity) {
    const headerRowIndex = findHeaderRowIndex(rows);

    if (headerRowIndex < 0) {
      return {
        code: identity.code,
        type: identity.type,
        paperLabels: [],
        students: [],
        scName: "",
        rollCode: "",
      };
    }

    const headerRow = rows[headerRowIndex] || [];
    const paperRow = rows[headerRowIndex + 1] || [];
    const paperLabels = extractPaperLabels(paperRow);
    const extraColumns = getExtraColumnIndexes(headerRow);
    const scDetails = extractScDetails(rows, headerRowIndex);
    const students = extractStudents(rows, headerRowIndex, paperLabels, extraColumns, scDetails);

    return {
      code: identity.code,
      type: identity.type,
      paperLabels,
      students,
      scName: scDetails.name,
      rollCode: scDetails.rollCode,
    };
  }

  function getExtraColumnIndexes(headerRow) {
    return {
      grandTotal: findColumnIndex(headerRow, ["GRAND", "TOTAL"]),
      grade: findColumnIndex(headerRow, ["GRADE"]),
      roll: findColumnIndex(headerRow, ["ROLL"]),
      er: findColumnIndex(headerRow, ["ER"], ["NUMBER"]),
      mSlNo: findColumnIndex(headerRow, ["M", "SL", "NO"]),
      result: findColumnIndex(headerRow, ["RESULT"]),
    };
  }

  function findColumnIndex(headerRow, requiredTokens, fallbackTokens = null) {
    const upperRow = headerRow.map((cell) => normalizeText(cell).toUpperCase());

    for (let i = 0; i < upperRow.length; i += 1) {
      const value = upperRow[i];
      if (!value) {
        continue;
      }

      const matches = requiredTokens.every((token) => value.includes(token));
      if (matches) {
        return i;
      }
    }

    if (fallbackTokens) {
      for (let i = 0; i < upperRow.length; i += 1) {
        const value = upperRow[i];
        if (!value) {
          continue;
        }

        const matches = fallbackTokens.every((token) => value.includes(token));
        if (matches) {
          return i;
        }
      }
    }

    return -1;
  }

  function extractScDetails(rows, headerRowIndex) {
    for (let i = 0; i <= headerRowIndex; i += 1) {
      const row = rows[i] || [];
      const cell = normalizeText(row[0]);
      const upper = cell.toUpperCase();

      if (!upper.includes("SC WITH CODE")) {
        continue;
      }

      const afterDash = cell.includes("-") ? cell.split("-").slice(1).join("-").trim() : cell;
      const rollCodeMatch = afterDash.match(/\(([A-Za-z]-\d+)\)\s*$/);
      return {
        name: cleanScName(afterDash),
        rollCode: rollCodeMatch ? normalizeText(rollCodeMatch[1]).toUpperCase() : "",
      };
    }

    return { name: "", rollCode: "" };
  }

  function cleanScName(value) {
    const text = normalizeText(value);

    if (!text) {
      return "";
    }

    return text.replace(/\s*\([^)]+\)\s*$/, "").trim();
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

  function extractPaperLabels(row) {
    const labels = [];

    for (let col = 2; col <= 60; col += 7) {
      const rawLabel = normalizeText(row[col]);

      if (!rawLabel) {
        continue;
      }

      labels.push(rawLabel.replace(/\s+/g, " ").toUpperCase());
    }

    return labels;
  }

  function extractStudents(rows, headerRowIndex, paperLabels, extraColumns, scDetails) {
    const students = [];

    for (let rowIndex = headerRowIndex + 2; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
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
        if (students.length) {
          break;
        }

        continue;
      }

      if (!name || !enrollmentRaw) {
        continue;
      }

      const student = {
        name,
        enrollment: normalizeEnrollment(enrollmentRaw),
        papers: {},
        extra: {
          roll: scDetails.rollCode || getCellByIndex(row, extraColumns.roll),
          er: getCellByIndex(row, extraColumns.er, { roundNumeric: true }),
          mSlNo: getCellByIndex(row, extraColumns.mSlNo),
          grade: getCellByIndex(row, extraColumns.grade),
          result: getCellByIndex(row, extraColumns.result),
          grandTotal: getCellByIndex(row, extraColumns.grandTotal, { roundNumeric: true }),
          scName: scDetails.name,
        },
      };

      paperLabels.forEach((paperLabel, idx) => {
        const baseCol = 2 + idx * 7;
        student.papers[paperLabel] = {
          assignmentFm100: roundNumericText(row[baseCol]),
          assignment: roundNumericText(row[baseCol + 1]),
          teeFm100: roundNumericText(row[baseCol + 2]),
          tee70: roundNumericText(row[baseCol + 3]),
          total: roundNumericText(row[baseCol + 4]),
          cnc: normalizeText(row[baseCol + 6]),
        };
      });

      students.push(student);
    }

    return students;
  }

  function getCellByIndex(row, index, options = {}) {
    if (index < 0) {
      return "";
    }

    if (options.roundNumeric) {
      return roundNumericText(row[index]);
    }

    return normalizeText(row[index]);
  }

  function buildSecondYearSheet(code, firstYear, secondYear) {
    const firstLabels = firstYear ? firstYear.paperLabels : [];
    const secondLabels = secondYear.paperLabels;

    const headers = [
      "NAME",
      "ENROLLMENT NO",
      "SC NAME",
      "ROLL",
      "ER",
      "M. SL. NO.",
      "GRADE",
      "RESULT",
      "GRAND TOTAL",
    ];

    appendPaperHeaders(headers, firstLabels, "1ST YEAR");
    appendPaperHeaders(headers, secondLabels, "2ND YEAR");
    headers.push("REMARK");

    const rows = [headers];
    const firstExactMap = new Map();
    const firstEnrollmentMap = new Map();

    if (firstYear) {
      firstYear.students.forEach((student) => {
        firstExactMap.set(createStudentKey(student.name, student.enrollment), student);

        const enrollment = normalizeEnrollment(student.enrollment);
        if (!firstEnrollmentMap.has(enrollment)) {
          firstEnrollmentMap.set(enrollment, []);
        }
        firstEnrollmentMap.get(enrollment).push(student);
      });
    }

    secondYear.students.forEach((student) => {
      const exactKey = createStudentKey(student.name, student.enrollment);
      const firstStudent = firstExactMap.get(exactKey);
      const enrollmentMatches = firstEnrollmentMap.get(normalizeEnrollment(student.enrollment)) || [];

      const row = [
        student.name,
        student.enrollment,
        student.extra.scName,
        student.extra.roll,
        student.extra.er,
        student.extra.mSlNo,
        student.extra.grade,
        student.extra.result,
        student.extra.grandTotal,
      ];

      firstLabels.forEach((label) => {
        const marks = firstStudent ? firstStudent.papers[label] : null;
        pushPaperValues(row, marks);
      });

      secondLabels.forEach((label) => {
        const marks = student.papers[label];
        pushPaperValues(row, marks);
      });

      if (firstStudent) {
        row.push("Merged from 1st + 2nd year using exact name+enrollment match");
      } else if (enrollmentMatches.length) {
        row.push("Enrollment matched in 1st year but name is different");
      } else {
        row.push("No matching 1st-year record");
      }

      rows.push(row);
    });

    return {
      code,
      rows,
    };
  }

  function buildLateralSheet(lateralSheet) {
    const headers = [
      "NAME",
      "ENROLLMENT NO",
      "SC NAME",
      "ROLL",
      "ER",
      "M. SL. NO.",
      "GRADE",
      "RESULT",
      "GRAND TOTAL",
    ];

    appendPaperHeaders(headers, lateralSheet.paperLabels, "LATERAL");
    headers.push("REMARK");

    const rows = [headers];

    lateralSheet.students.forEach((student) => {
      const row = [
        student.name,
        student.enrollment,
        student.extra.scName,
        student.extra.roll,
        student.extra.er,
        student.extra.mSlNo,
        student.extra.grade,
        student.extra.result,
        student.extra.grandTotal,
      ];

      lateralSheet.paperLabels.forEach((label) => {
        const marks = student.papers[label];
        pushPaperValues(row, marks);
      });

      row.push("Lateral record");
      rows.push(row);
    });

    return {
      rows,
    };
  }

  function appendPaperHeaders(headers, paperLabels, yearLabel) {
    paperLabels.forEach((label) => {
      headers.push(`${yearLabel} ${label} ASSIGNMENT (F.M 100)`);
      headers.push(`${yearLabel} ${label} ASSIGNMENT`);
      headers.push(`${yearLabel} ${label} T.E.E (F.M 100)`);
      headers.push(`${yearLabel} ${label} T.E.E 70`);
      headers.push(`${yearLabel} ${label} TOTAL F.M 100`);
      headers.push(`${yearLabel} ${label} C/NC`);
    });
  }

  function pushPaperValues(row, marks) {
    if (!marks) {
      row.push("", "", "", "", "", "");
      return;
    }

    row.push(marks.assignmentFm100);
    row.push(marks.assignment);
    row.push(marks.teeFm100);
    row.push(marks.tee70);
    row.push(marks.total);
    row.push(marks.cnc);
  }

  function createStudentKey(name, enrollment) {
    return `${normalizeEnrollment(enrollment)}|${normalizeName(name)}`;
  }

  function roundNumericText(value) {
    const text = normalizeText(value);

    if (!text) {
      return "";
    }

    const compact = text.replace(/,/g, "");
    if (!/^[+-]?\d+(\.\d+)?$/.test(compact)) {
      return text;
    }

    return String(Math.round(Number(compact)));
  }

  function populateSheetList(sheetNames) {
    sheetSelect.innerHTML = "";

    sheetNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      sheetSelect.appendChild(option);
    });
  }

  function renderSheet(sheetName) {
    const rows = processedSheets[sheetName] || [];

    if (!rows.length) {
      excelTable.innerHTML = "<tbody><tr><td>No data available.</td></tr></tbody>";
      return;
    }

    const tableHead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    rows[0].forEach((value) => {
      const th = document.createElement("th");
      th.textContent = normalizeText(value);
      headerRow.appendChild(th);
    });

    tableHead.appendChild(headerRow);

    const tableBody = document.createElement("tbody");

    rows.slice(1).forEach((row) => {
      const tr = document.createElement("tr");
      const colCount = rows[0].length;

      for (let col = 0; col < colCount; col += 1) {
        const td = document.createElement("td");
        td.contentEditable = "true";
        td.textContent = normalizeText(row[col]);
        tr.appendChild(td);
      }

      tableBody.appendChild(tr);
    });

    excelTable.replaceChildren(tableHead, tableBody);
  }

  function persistCurrentSheetEdits() {
    if (!activeOutputSheet || !excelTable.querySelector("tr")) {
      return;
    }

    processedSheets[activeOutputSheet] = extractTableData();
  }

  function extractTableData() {
    const rows = [];
    const rowElements = excelTable.querySelectorAll("tr");

    rowElements.forEach((row) => {
      const cells = row.querySelectorAll("th, td");
      const values = Array.from(cells, (cell) => normalizeText(cell.textContent));
      rows.push(values);
    });

    return rows;
  }

  function resetView() {
    sheetSelect.innerHTML = "";
    sheetSelect.disabled = true;
    downloadBtn.disabled = true;
    clearBtn.disabled = true;
    excelTable.innerHTML = "";
    summaryPanel.innerHTML = "";
  }

  function normalizeText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function normalizeName(value) {
    return normalizeText(value).toUpperCase().replace(/\s+/g, " ").trim();
  }

  function normalizeEnrollment(value) {
    const text = normalizeText(value).replace(/\s+/g, "");

    if (/^\d+\.0+$/.test(text)) {
      return text.replace(/\.0+$/, "");
    }

    return text;
  }

  function safeSheetName(name) {
    return name.replace(/[\\/?*\[\]:]/g, " ").slice(0, 31);
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

