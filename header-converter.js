const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const downloadBtn = document.getElementById("downloadBtn");
const statusText = document.getElementById("status");
const summaryPanel = document.getElementById("summary");
const excelTable = document.getElementById("excelTable");

if (fileInput && sheetSelect && downloadBtn && statusText && summaryPanel && excelTable) {
  let convertedSheets = {};
  let conversionSummary = [];
  let activeSheet = "";
  let sourceName = "";

  fileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    resetView();
    sourceName = file.name;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const result = convertWorkbook(workbook);

      if (!Object.keys(result.sheets).length) {
        throw new Error("No convertible sheets found.");
      }

      convertedSheets = result.sheets;
      conversionSummary = result.summary;
      activeSheet = Object.keys(convertedSheets)[0];

      populateSheetList(Object.keys(convertedSheets));
      sheetSelect.value = activeSheet;
      renderSheet(activeSheet);
      renderSummary(conversionSummary);

      sheetSelect.disabled = false;
      downloadBtn.disabled = false;
      statusText.textContent = `Converted ${file.name}: ${Object.keys(convertedSheets).length} sheet(s) ready.`;
    } catch (error) {
      resetView();
      statusText.textContent = "Could not convert this file. Please upload a valid combined processed workbook.";
      console.error(error);
    }
  });

  sheetSelect.addEventListener("change", (event) => {
    activeSheet = event.target.value;
    renderSheet(activeSheet);
  });

  downloadBtn.addEventListener("click", () => {
    if (!activeSheet) {
      return;
    }

    const outputBook = XLSX.utils.book_new();
    const usedNames = new Set();

    Object.keys(convertedSheets).forEach((sheetName) => {
      const rows = convertedSheets[sheetName];
      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      const safeName = safeUniqueSheetName(sheetName, usedNames);
      XLSX.utils.book_append_sheet(outputBook, worksheet, safeName);
    });

    const base = sourceName.replace(/\.[^.]+$/, "");
    XLSX.writeFile(outputBook, `${base}_renamed_headers.xlsx`);
  });

  function convertWorkbook(workbook) {
    const sheets = {};
    const summary = [];

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      const converted = convertSheet(rows, sheetName);

      if (!converted.rows.length) {
        return;
      }

      sheets[sheetName] = converted.rows;
      summary.push({
        sheet: sheetName,
        type: converted.type,
        papers: converted.paperCount,
        rows: Math.max(0, converted.rows.length - 1),
      });
    });

    return { sheets, summary };
  }

  function convertSheet(rows, sheetName) {
    if (!rows.length) {
      return { rows: [], paperCount: 0, type: "unknown" };
    }

    const headers = (rows[0] || []).map((cell) => normalizeText(cell));
    const upperHeaders = headers.map((h) => h.toUpperCase());

    const idx = {
      name: findHeaderIndex(upperHeaders, ["NAME"]),
      enrollment: findHeaderIndex(upperHeaders, ["ENROLLMENT NO"]),
      er: findHeaderIndex(upperHeaders, ["ER"], ["NUMBER"]),
      mSlNo: findHeaderIndex(upperHeaders, ["M. SL. NO."]),
      grade: findHeaderIndex(upperHeaders, ["GRADE"]),
      result: findHeaderIndex(upperHeaders, ["RESULT"]),
      roll: findHeaderIndex(upperHeaders, ["ROLL"]),
      scName: findHeaderIndex(upperHeaders, ["SC NAME"]),
      grandTotal: findHeaderIndex(upperHeaders, ["GRAND TOTAL"]),
    };

    const paperBlocks = findPaperBlocks(upperHeaders);
    if (!paperBlocks.length || idx.name < 0 || idx.enrollment < 0) {
      return { rows: [], paperCount: 0, type: "unknown" };
    }

    const isLateral = sheetName.toUpperCase().includes("LATERAL");
    const convertedHeaders = buildConvertedHeaders(paperBlocks.length, isLateral);
    const outputRows = [convertedHeaders];

    for (let r = 1; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const name = normalizeText(row[idx.name]);
      const enrollment = normalizeEnrollment(row[idx.enrollment]);

      if (!name && !enrollment) {
        continue;
      }

      const statusValues = paperBlocks.map((block) => normalizeText(row[block.base + 5]));
      const marksValues = [];
      let tmo = 0;
      let tmoCount = 0;

      paperBlocks.forEach((block) => {
        const totalNum = toNumber(row[block.base + 4]);
        if (totalNum !== null) {
          tmo += totalNum;
          tmoCount += 1;
        }

        marksValues.push(roundNumericText(row[block.base]));
        marksValues.push(roundNumericText(row[block.base + 1]));
        marksValues.push(roundNumericText(row[block.base + 2]));
        marksValues.push(roundNumericText(row[block.base + 3]));
        marksValues.push(roundNumericText(row[block.base + 4]));
      });

      const gtNum = toNumber(row[idx.grandTotal]);
      const gtText = gtNum === null ? "" : String(Math.round(gtNum));
      const tmoText = tmoCount ? String(Math.round(tmo)) : "";
      const ytText = isLateral && gtNum !== null && tmoCount ? String(Math.round(gtNum - tmo)) : "";
      const code = extractCenterCode(row[idx.roll], row[idx.mSlNo], row[idx.scName]) || normalizeText(row[idx.roll]);

      const convertedRow = [
        name,
        enrollment,
        roundNumericText(row[idx.er]),
        normalizeText(row[idx.mSlNo]),
        ...statusValues,
        normalizeText(row[idx.result]),
        normalizeText(row[idx.grade]),
        code,
        code,
        normalizeText(row[idx.scName]),
        ...marksValues,
      ];

      if (isLateral) {
        convertedRow.push(tmoText, ytText, gtText);
      } else {
        convertedRow.push(tmoText);
      }

      outputRows.push(convertedRow);
    }

    return {
      rows: outputRows,
      paperCount: paperBlocks.length,
      type: isLateral ? "lateral" : "2nd-year",
    };
  }

  function extractCenterCode(...values) {
    for (const value of values) {
      const text = normalizeText(value).toUpperCase();
      if (!text) {
        continue;
      }

      const direct = text.match(/\b[A-Z]{1,3}-\d{1,4}\b/);
      if (direct) {
        return direct[0];
      }

      const slashCode = text.match(/\/([A-Z]{1,3}-\d{1,4})\//);
      if (slashCode) {
        return slashCode[1];
      }

      const parenCode = text.match(/\(([A-Z]{1,3}-\d{1,4})\)/);
      if (parenCode) {
        return parenCode[1];
      }
    }

    return "";
  }

  function findHeaderIndex(upperHeaders, tokens, fallbackTokens = null) {
    for (let i = 0; i < upperHeaders.length; i += 1) {
      const value = upperHeaders[i];
      if (!value) {
        continue;
      }

      if (tokens.every((token) => value.includes(token))) {
        return i;
      }
    }

    if (fallbackTokens) {
      for (let i = 0; i < upperHeaders.length; i += 1) {
        const value = upperHeaders[i];
        if (!value) {
          continue;
        }

        if (fallbackTokens.every((token) => value.includes(token))) {
          return i;
        }
      }
    }

    return -1;
  }

  function findPaperBlocks(upperHeaders) {
    const blocks = [];

    for (let i = 0; i < upperHeaders.length; i += 1) {
      const value = upperHeaders[i];
      const match = value.match(/PAPER\s*(\d+)\s+ASSIGNMENT\s*\(F\.M 100\)/);
      if (!match) {
        continue;
      }

      blocks.push({
        paperNo: Number(match[1]),
        base: i,
      });
    }

    blocks.sort((a, b) => a.paperNo - b.paperNo);
    return blocks;
  }

  function buildConvertedHeaders(paperCount, isLateral) {
    const headers = ["N", "EN", "NO", "SLN"];

    for (let i = 0; i < paperCount; i += 1) {
      headers.push(i === 0 ? "C" : `C${i}`);
    }

    headers.push("RESULT", "G", "R", "SCC", "SCN");

    for (let i = 0; i < paperCount; i += 1) {
      const index = i + 1;
      const letter = String.fromCharCode(65 + i);
      const taHeader = i === 0 ? "TA" : `TA${i}`;

      headers.push(`A${index}`);
      headers.push(`${letter}30`);
      headers.push(`OB${index}`);
      headers.push(`${letter}70`);
      headers.push(taHeader);
    }

    if (isLateral) {
      headers.push("TMO", "YT", "GT");
    } else {
      headers.push("TMO");
    }

    return headers;
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

  function renderSummary(summaryItems) {
    const list = summaryItems
      .map((item) => `<li><strong>${escapeHtml(item.sheet)}:</strong> ${item.type}, ${item.papers} paper(s), ${item.rows} row(s)</li>`)
      .join("");

    summaryPanel.innerHTML = `<h2>Conversion Summary</h2><ul>${list}</ul>`;
  }

  function renderSheet(sheetName) {
    const rows = convertedSheets[sheetName] || [];

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
        td.textContent = normalizeText(row[col]);
        tr.appendChild(td);
      }

      tableBody.appendChild(tr);
    });

    excelTable.replaceChildren(tableHead, tableBody);
  }

  function resetView() {
    convertedSheets = {};
    conversionSummary = [];
    activeSheet = "";
    sheetSelect.innerHTML = "";
    sheetSelect.disabled = true;
    downloadBtn.disabled = true;
    excelTable.innerHTML = "";
    summaryPanel.innerHTML = "";
  }

  function normalizeText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function normalizeEnrollment(value) {
    const text = normalizeText(value).replace(/\s+/g, "");
    if (/^\d+\.0+$/.test(text)) {
      return text.replace(/\.0+$/, "");
    }

    return text;
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

  function toNumber(value) {
    const text = normalizeText(value);
    if (!text) {
      return null;
    }

    const compact = text.replace(/,/g, "");
    if (!/^[+-]?\d+(\.\d+)?$/.test(compact)) {
      return null;
    }

    return Number(compact);
  }

  function safeSheetName(name) {
    return normalizeText(name).replace(/[\\/?*\[\]:]/g, " ").slice(0, 31);
  }

  function safeUniqueSheetName(name, usedNames) {
    const base = safeSheetName(name) || "Sheet";

    if (!usedNames.has(base)) {
      usedNames.add(base);
      return base;
    }

    let index = 2;
    while (true) {
      const suffix = ` ${index}`;
      const prefix = base.slice(0, Math.max(1, 31 - suffix.length)).trimEnd();
      const candidate = `${prefix}${suffix}`;

      if (!usedNames.has(candidate)) {
        usedNames.add(candidate);
        return candidate;
      }

      index += 1;
    }
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
