const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("status");
const summaryPanel = document.getElementById("summary");
const resultTable = document.getElementById("resultTable");

if (fileInput && sheetSelect && downloadBtn && clearBtn && statusText && summaryPanel && resultTable) {
  // courseSheets is keyed by a canonical (case-insensitive) course key
  // value shape: { displayName: string, records: Array<Record> }
  let courseSheets = {};
  let activeCourseKey = "";
  let loadedFileCount = 0;
  let uploadHistory = [];

  fileInput.addEventListener("change", async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) {
      return;
    }

    let batchAdded = 0;
    let batchSucceeded = 0;
    let batchSkipped = 0;

    for (const file of files) {
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        const extracted = extractFromWorkbook(workbook, file.name);

        if (!extracted.length) {
          batchSkipped += 1;
          uploadHistory.unshift({ fileName: file.name, status: "Skipped", detail: "No usable CLL data found" });
          continue;
        }

        extracted.forEach((record) => {
          const rawCourse = normalizeText(record.course) || "UNSPECIFIED COURSE";
          const courseKey = normalizeCourseKey(rawCourse);

          if (!courseSheets[courseKey]) {
            courseSheets[courseKey] = {
              displayName: rawCourse,
              records: [],
            };
          }

          courseSheets[courseKey].records.push(record);
        });

        batchAdded += extracted.length;
        batchSucceeded += 1;
        loadedFileCount += 1;
        uploadHistory.unshift({ fileName: file.name, status: "Added", detail: `${extracted.length} record(s) extracted` });
      } catch (error) {
        batchSkipped += 1;
        uploadHistory.unshift({ fileName: file.name, status: "Error", detail: "Could not process this file" });
        console.error(error);
      }
    }

    uploadHistory = uploadHistory.slice(0, 20);
    refreshSheetSelectorAndView();
    renderSummary();

    const totalRecords = getTotalRecordCount();
    downloadBtn.disabled = totalRecords === 0;
    clearBtn.disabled = totalRecords === 0;

    if (!totalRecords) {
      statusText.textContent = "No usable CLL data extracted yet.";
    } else {
      statusText.textContent = `Added ${batchSucceeded} file(s), skipped ${batchSkipped}. Added ${batchAdded} record(s). Total records: ${totalRecords}.`;
    }

    fileInput.value = "";
  });

  sheetSelect.addEventListener("change", (event) => {
    activeCourseKey = event.target.value;
    renderTable();
  });

  downloadBtn.addEventListener("click", () => {
    const courseKeys = Object.keys(courseSheets);
    if (!courseKeys.length) {
      return;
    }

    const headers = buildHeaders();
    const outputBook = XLSX.utils.book_new();

    // Keep case-insensitive set to avoid Excel sheet-name collisions.
    const usedSheetNames = new Set();

    courseKeys.forEach((courseKey) => {
      const group = courseSheets[courseKey];
      const rows = (group.records || []).map((record) => buildRow(record));
      const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      const safeName = safeUniqueSheetName(group.displayName, usedSheetNames);
      XLSX.utils.book_append_sheet(outputBook, worksheet, safeName);
    });

    XLSX.writeFile(outputBook, "cll_tabulation_data.xlsx");
  });

  clearBtn.addEventListener("click", () => {
    courseSheets = {};
    activeCourseKey = "";
    loadedFileCount = 0;
    uploadHistory = [];
    resetView();
    statusText.textContent = "All extracted data cleared.";
  });

  function extractFromWorkbook(workbook, sourceFileName) {
    const output = [];

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      const headerInfo = findStudentHeaderRow(rows);
      if (!headerInfo) {
        return;
      }

      const centreName = findValueAfterLabel(rows, /CENT(RE|ER)\s*NAME/i, headerInfo.rowIndex);
      const centreCode = findValueAfterLabel(rows, /CENT(RE|ER)\s*CODE/i, headerInfo.rowIndex);
      const course = findValueAfterLabel(rows, /COURSE\s*-/i, headerInfo.rowIndex, /COURSE\s*DURATION/i);

      const paperNames = extractPaperNames(rows, headerInfo.rowIndex);
      const students = extractStudents(rows, headerInfo);

      students.forEach((student) => {
        const paperData = {};

        for (let p = 1; p <= 6; p += 1) {
          paperData[p] = {
            name: normalizeText(paperNames[p]),
            mark: normalizeText(student.paperMarks[p]),
          };
        }

        output.push({
          name: student.name,
          uin: student.uin,
          papers: paperData,
          marksObtained: student.marksObtained,
          fullMarks: student.fullMarks,
          percentMarks: student.percentMarks,
          grade: student.grade,
          centreName,
          centreCode,
          course,
          fileName: sourceFileName,
          sheetName,
        });
      });
    });

    return output;
  }

  function findStudentHeaderRow(rows) {
    for (let r = 0; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const line = row.map((cell) => normalizeText(cell).toUpperCase()).join(" ");

      if (!line.includes("UIN")) {
        continue;
      }

      if (!line.includes("STUDENT") && !line.includes("STUDENTS") && !line.includes("NAME")) {
        continue;
      }

      const columns = {
        name: -1,
        uin: -1,
        marksObtained: -1,
        fullMarks: -1,
        percentMarks: -1,
        grade: -1,
        paper: {},
      };

      row.forEach((cell, idx) => {
        const norm = simplifyToken(cell);

        if (norm === "NAME" || norm.includes("STUDENTSNAME") || norm.includes("STUDENTNAME")) {
          columns.name = idx;
        }

        if (norm === "UIN" || norm.includes("UIN")) {
          columns.uin = idx;
        }

        const paperMatch = norm.match(/^P-?(\d)$/);
        if (paperMatch) {
          const paperNo = Number(paperMatch[1]);
          if (paperNo >= 1 && paperNo <= 6) {
            columns.paper[paperNo] = idx;
          }
        }

        if (norm.includes("MARKSOBTAINED") || norm === "OBTAINED") {
          columns.marksObtained = idx;
        }

        if (norm.includes("FULLMARKS")) {
          columns.fullMarks = idx;
        }

        if (norm.includes("OFMARKS") || norm.includes("PERCENTMARKS") || norm.includes("%OFMARKS")) {
          columns.percentMarks = idx;
        }

        if (norm.includes("GRADE")) {
          columns.grade = idx;
        }
      });

      if (columns.name >= 0 && columns.uin >= 0) {
        return { rowIndex: r, columns };
      }
    }

    return null;
  }

  function extractPaperNames(rows, maxRowExclusive) {
    const names = {};
    for (let p = 1; p <= 6; p += 1) {
      names[p] = "";
    }

    for (let r = 0; r < maxRowExclusive; r += 1) {
      const row = rows[r] || [];

      for (let c = 0; c < row.length; c += 1) {
        const text = normalizeText(row[c]);
        const match = text.match(/^PAPER[-\s]?(\d)\s*[:\-]?\s*(.*)$/i);

        if (!match) {
          continue;
        }

        const paperNo = Number(match[1]);
        if (paperNo < 1 || paperNo > 6) {
          continue;
        }

        let name = normalizeText(match[2]);

        if (!name) {
          for (let j = c + 1; j < row.length; j += 1) {
            const candidate = normalizeText(row[j]);
            if (candidate) {
              name = candidate;
              break;
            }
          }
        }

        if (name && !names[paperNo]) {
          names[paperNo] = name;
        }
      }
    }

    return names;
  }

  function extractStudents(rows, headerInfo) {
    const students = [];
    const { rowIndex, columns } = headerInfo;
    let started = false;

    for (let r = rowIndex + 1; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const name = normalizeText(row[columns.name]);
      const uin = normalizeText(row[columns.uin]);

      if (!name && !uin) {
        if (started) {
          break;
        }
        continue;
      }

      if (!name || !uin) {
        continue;
      }

      started = true;
      const paperMarks = {};

      for (let p = 1; p <= 6; p += 1) {
        const colIndex = columns.paper[p];
        paperMarks[p] = colIndex === undefined ? "" : normalizeText(row[colIndex]);
      }

      students.push({
        name,
        uin,
        paperMarks,
        marksObtained: columns.marksObtained >= 0 ? normalizeText(row[columns.marksObtained]) : "",
        fullMarks: columns.fullMarks >= 0 ? normalizeText(row[columns.fullMarks]) : "",
        percentMarks: columns.percentMarks >= 0 ? roundNumericText(row[columns.percentMarks]) : "",
        grade: columns.grade >= 0 ? normalizeText(row[columns.grade]) : "",
      });
    }

    return students;
  }

  function findValueAfterLabel(rows, labelRegex, maxRowExclusive, excludeRegex = null) {
    for (let r = 0; r < Math.min(maxRowExclusive, rows.length); r += 1) {
      const row = rows[r] || [];

      for (let c = 0; c < row.length; c += 1) {
        const label = normalizeText(row[c]);
        if (!label || !labelRegex.test(label)) {
          continue;
        }

        if (excludeRegex && excludeRegex.test(label)) {
          continue;
        }

        for (let j = c + 1; j < row.length; j += 1) {
          const value = normalizeText(row[j]);
          if (value) {
            return value;
          }
        }

        const afterDash = label.split("-").slice(1).join("-").trim();
        if (afterDash) {
          return afterDash;
        }
      }
    }

    return "";
  }

  function buildHeaders() {
    const headers = ["NAME", "UIN"];

    for (let p = 1; p <= 6; p += 1) {
      headers.push(`P-${p} NAME`);
      headers.push(`P-${p}`);
    }

    headers.push(
      "MARKS OBTAINED",
      "FULL MARKS",
      "% MARKS",
      "GRADE",
      "CENTRE NAME",
      "CENTRE CODE",
      "COURSE",
      "FILE NAME",
      "SHEET NAME"
    );

    return headers;
  }

  function buildRow(record) {
    const row = [
      safeExcelText(record.name),
      safeExcelText(record.uin),
    ];

    for (let p = 1; p <= 6; p += 1) {
      const paper = record.papers[p] || { name: "", mark: "" };
      row.push(safeExcelText(paper.name || ""));
      row.push(safeExcelText(paper.mark || ""));
    }

    row.push(
      safeExcelText(record.marksObtained),
      safeExcelText(record.fullMarks),
      safeExcelText(record.percentMarks),
      safeExcelText(record.grade),
      safeExcelText(record.centreName),
      safeExcelText(record.centreCode),
      safeExcelText(record.course),
      safeExcelText(record.fileName),
      safeExcelText(record.sheetName)
    );

    return row;
  }

  function getTotalRecordCount() {
    return Object.values(courseSheets).reduce((sum, group) => sum + (group.records || []).length, 0);
  }

  function refreshSheetSelectorAndView() {
    const courseKeys = Object.keys(courseSheets);

    if (!courseKeys.length) {
      resetView();
      return;
    }

    if (!courseKeys.includes(activeCourseKey)) {
      activeCourseKey = courseKeys[0];
    }

    populateSheetList(courseKeys);
    sheetSelect.value = activeCourseKey;
    renderTable();

    sheetSelect.disabled = false;
    downloadBtn.disabled = false;
    clearBtn.disabled = false;
  }

  function populateSheetList(courseKeys) {
    sheetSelect.innerHTML = "";

    courseKeys.forEach((key) => {
      const option = document.createElement("option");
      option.value = key;
      option.textContent = courseSheets[key].displayName;
      sheetSelect.appendChild(option);
    });
  }

  function renderTable() {
    const group = courseSheets[activeCourseKey];
    const courseRecords = group ? group.records : [];

    if (!courseRecords.length) {
      resultTable.innerHTML = "";
      return;
    }

    const headers = buildHeaders();

    const tableHead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headerRow.appendChild(th);
    });

    tableHead.appendChild(headerRow);

    const tableBody = document.createElement("tbody");
    courseRecords.forEach((record) => {
      const tr = document.createElement("tr");
      const row = buildRow(record);

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
    const courseKeys = Object.keys(courseSheets);
    if (!courseKeys.length) {
      summaryPanel.innerHTML = "";
      return;
    }

    const totalRecords = getTotalRecordCount();
    const sheetSet = new Set();

    courseKeys.forEach((key) => {
      (courseSheets[key].records || []).forEach((record) => {
        sheetSet.add(`${record.fileName} :: ${record.sheetName}`);
      });
    });

    const rowsPerCourse = courseKeys
      .map((key) => `<li><strong>${escapeHtml(courseSheets[key].displayName)}:</strong> ${(courseSheets[key].records || []).length} record(s)</li>`)
      .join("");

    const recentUploads = uploadHistory
      .slice(0, 8)
      .map((item) => `<li><strong>${escapeHtml(item.status)}:</strong> ${escapeHtml(item.fileName)} - ${escapeHtml(item.detail)}</li>`)
      .join("");

    summaryPanel.innerHTML = `
      <h2>Extraction Summary</h2>
      <ul>
        <li><strong>Loaded files:</strong> ${loadedFileCount}</li>
        <li><strong>Total records:</strong> ${totalRecords}</li>
        <li><strong>Course sheets created:</strong> ${courseKeys.length}</li>
        <li><strong>Total file-sheet combinations:</strong> ${sheetSet.size}</li>
      </ul>
      <h2>Rows Per Course</h2>
      <ul>${rowsPerCourse}</ul>
      <h2>Recent Uploads</h2>
      <ul>${recentUploads || "<li>No uploads yet.</li>"}</ul>
    `;
  }

  function resetView() {
    sheetSelect.innerHTML = "";
    sheetSelect.disabled = true;
    downloadBtn.disabled = true;
    clearBtn.disabled = true;
    resultTable.innerHTML = "";
    summaryPanel.innerHTML = "";
  }

  function safeSheetName(name) {
    return safeExcelText(name)
      .replace(/[\\/?*\[\]:]/g, " ")
      .replace(/^[\s'"]+|[\s'"]+$/g, "")
      .slice(0, 31) || "Sheet";
  }

  function safeUniqueSheetName(name, usedSheetNames) {
    const base = safeSheetName(name);

    if (!usedSheetNames.has(base.toLowerCase())) {
      usedSheetNames.add(base.toLowerCase());
      return base;
    }

    let index = 2;
    while (true) {
      const suffix = ` ${index}`;
      const candidate = `${base.slice(0, Math.max(1, 31 - suffix.length)).trimEnd()}${suffix}`;
      const candidateKey = candidate.toLowerCase();

      if (!usedSheetNames.has(candidateKey)) {
        usedSheetNames.add(candidateKey);
        return candidate;
      }

      index += 1;
    }
  }

  function normalizeText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function normalizeCourseKey(value) {
    return normalizeText(value).replace(/\s+/g, " ").toUpperCase();
  }

  // Remove control characters that make Excel mark files as corrupted.
  function safeExcelText(value) {
    return normalizeText(value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "");
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

  function simplifyToken(value) {
    return normalizeText(value)
      .toUpperCase()
      .replace(/\s+/g, "")
      .replace(/[^A-Z0-9%-]/g, "");
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
