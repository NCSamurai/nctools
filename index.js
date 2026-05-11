document.addEventListener("DOMContentLoaded", () => {
  let html5QrCode = null;
  let scannerRunning = false;
  const state = {
    batchNum: Number(localStorage.getItem("der2_batch")) || 1,
    workbook: null,
    sheets: new Set(),
  };

  // --- Reset batch if date changed ---
  function checkBatchDate() {
    const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
    const storedDate = localStorage.getItem("der2_date");

    if (storedDate !== today) {
      state.batchNum = 1;
      localStorage.setItem("der2_batch", "1");
      localStorage.setItem("der2_date", today);
    } else {
      state.batchNum = Number(localStorage.getItem("der2_batch")) || 1;
    }
  }

  checkBatchDate();
  loadSession();
  renderPreviews();
  updateDownloadButton();
  updateProductDropdown();
  saveSession();

  // --- Helpers ---
  function todayLabel() {
    const d = new Date();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}-${dd}-${yy}`;
  }

  function sheetFilename() {
    return `DER2 Report ${todayLabel()}_Batch ${state.batchNum}.xlsx`;
  }

  function numericCheck(s) {
    if (!s) return true;
    return !/^\d+$/.test(s);
  }

  function formatData(cd_data, time_data) {
    let code_date = "";
    if (cd_data.length === 6) {
      code_date = `${cd_data.slice(0, 2)}-${cd_data.slice(
        2,
        4,
      )}-${cd_data.slice(4)}`;
    } else if (cd_data.length === 5) {
      if (cd_data.slice(0, 2).startsWith("0")) {
        code_date = `${cd_data.slice(0, 2)}-${cd_data.slice(
          2,
          3,
        )}-${cd_data.slice(3)}`;
      } else {
        code_date = `${cd_data.slice(0, 1)}-${cd_data.slice(
          1,
          3,
        )}-${cd_data.slice(3)}`;
      }
    } else {
      code_date = `${cd_data.slice(0, 1)}-${cd_data.slice(
        1,
        2,
      )}-${cd_data.slice(2)}`;
    }

    let timestamp = "";
    if (time_data.length === 4) {
      timestamp = `${time_data.slice(0, 2)}:${time_data.slice(2)}`;
    } else {
      timestamp = `${time_data.slice(0, 1)}:${time_data.slice(1)}`;
    }
    return { code_date, timestamp };
  }

  function initWorkbook() {
    if (!state.workbook) {
      state.workbook = new ExcelJS.Workbook();
    }
  }

  function getOrCreateSheet(name) {
    initWorkbook();
    let ws = state.workbook.getWorksheet(name);
    if (!ws) {
      ws = state.workbook.addWorksheet(name);
      state.sheets.add(name);

      const base = [
        "Ticket",
        "Product Number",
        "Code Date",
        "Cases",
        "Line",
        "Timestamp",
      ];
      const headers =
        name === "Duplicates"
          ? [
              "Ticket",
              "New Ticket",
              "Product Number",
              "Code Date",
              "Cases",
              "Line",
              "Timestamp",
            ]
          : base;

      ws.addRow(headers);
      ws.getRow(1).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 15 };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF4F81BD" },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
    }
    return ws;
  }

  function findColumnIndexByHeader(ws, headerName) {
    const row = ws.getRow(1);
    for (let col = 1; col <= row.cellCount; col++) {
      if (row.getCell(col).value === headerName) return col;
    }
    return null;
  }

  function ensureNewTicketColumn(ws) {
    const hasNew = findColumnIndexByHeader(ws, "New Ticket");
    if (hasNew) return;
    const headerRow = ws.getRow(1);
    headerRow.splice(2, 0, "New Ticket");
    for (let r = 2; r <= ws.rowCount; r++) {
      ws.getRow(r).splice(2, 0, "");
    }
  }

  function addRowWithStyles(ws, rowValues, wrongSet) {
    const row = ws.addRow(rowValues);
    const even = row.number % 2 === 0;
    const fillEven = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2F2F2" },
    };
    const fillOdd = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD7EBFF" },
    };
    const highlight = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFC107" },
    };

    row.eachCell((cell) => {
      cell.font = { size: 14 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      const header = ws.getRow(1).getCell(cell.col).value;
      cell.fill = wrongSet.has(header) ? highlight : even ? fillEven : fillOdd;
    });

    ws.columns.forEach((col) => {
      let maxLen = 0;
      col.eachCell({ includeEmpty: true }, (cell) => {
        const v = cell.value == null ? "" : String(cell.value);
        maxLen = Math.max(maxLen, v.length);
      });
      col.width = Math.max(10, maxLen + 8);
    });
  }

  // --- Preview rendering ---
  function renderPreviews() {
    const tabs = document.getElementById("tabs");
    const btns = document.getElementById("tab_buttons");
    const content = document.getElementById("tab_content");
    btns.innerHTML = "";
    content.innerHTML = "";

    const sheetNames = Array.from(state.sheets);
    tabs.style.display = "block";

    if (sheetNames.length === 0) {
      const msg = document.createElement("div");
      msg.style.textAlign = "center";
      msg.style.color = "#666";
      msg.style.fontStyle = "italic";
      msg.textContent = "No tickets yet. Add one to see the preview.";
      content.appendChild(msg);
      return;
    }

    sheetNames.forEach((name, idx) => {
      const b = document.createElement("button");
      b.textContent = name;
      b.className = idx === 0 ? "active button" : "";
      b.addEventListener("click", () => {
        document
          .querySelectorAll(".tab-buttons button")
          .forEach((el) => el.classList.remove("active"));
        b.classList.add("active");
        renderSheetTable(name);
      });
      btns.appendChild(b);
    });

    renderSheetTable(sheetNames[0]);
  }

  function renderSheetTable(name) {
    const content = document.getElementById("tab_content");
    content.innerHTML = "";
    const ws = state.workbook.getWorksheet(name);
    const table = document.createElement("table");

    const thead = document.createElement("thead");
    const trHead = document.createElement("tr");
    ws.getRow(1).eachCell((cell) => {
      const th = document.createElement("th");
      th.textContent = cell.value;
      trHead.appendChild(th);
    });
    // Extra header for actions
    const thAction = document.createElement("th");
    thAction.textContent = "Actions";
    trHead.appendChild(thAction);
    thead.appendChild(trHead);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let r = 2; r <= ws.rowCount; r++) {
      const tr = document.createElement("tr");
      const row = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        const td = document.createElement("td");
        td.textContent = row.getCell(c).value ?? "";
        tr.appendChild(td);
      }

      // Add remove button
      const tdAction = document.createElement("td");
      const btnRemove = document.createElement("button");
      btnRemove.textContent = "Remove";
      btnRemove.className = " button button-sm button-clear";
      btnRemove.addEventListener("click", () => {
        showRemoveRowModal(name, r);
      });
      tdAction.appendChild(btnRemove);
      tr.appendChild(tdAction);

      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    content.appendChild(table);
  }

  let pendingRemove = null;

  function showRemoveRowModal(sheetName, rowIndex) {
    pendingRemove = { sheetName, rowIndex };
    document.getElementById("removeRowModal").style.display = "block";
  }

  document.getElementById("confirmRemoveRow").addEventListener("click", () => {
    if (pendingRemove) {
      const ws = state.workbook.getWorksheet(pendingRemove.sheetName);
      ws.spliceRows(pendingRemove.rowIndex, 1); // delete row
      pendingRemove = null;
      renderPreviews();
      updateDownloadButton();
      updateProductDropdown();
      saveSession();
    }
    document.getElementById("removeRowModal").style.display = "none";
  });

  document.getElementById("cancelRemoveRow").addEventListener("click", () => {
    pendingRemove = null;
    document.getElementById("removeRowModal").style.display = "none";
  });

  // --- Error and success messages ---
  function showErrors(list) {
    const container = document.getElementById("errors");
    container.innerHTML = "";
    list.forEach((msg) => {
      const div = document.createElement("div");
      div.className = "error";
      div.textContent = msg;
      container.appendChild(div);
      fadeInMessage(container);
    });
  }

  function clearErrors() {
    document.getElementById("errors").innerHTML = "";
  }

  // function flashSuccess() {
  //   const s = document.getElementById("success");
  //   s.style.display = "block";
  //   setTimeout(() => {
  //     s.style.display = "none";
  //   }, 1500);
  //   updateProductDropdown();
  // }

  saveSession();

  function updateDownloadButton() {
    const btn = document.getElementById("download_btn");
    btn.disabled = state.sheets.size === 0;
    btn.textContent = btn.disabled
      ? "Download DER2 Report"
      : "Download DER2 Report";
  }

  // --- Persistence helpers ---
  function saveSession() {
    if (!state.workbook) return;

    const allData = {};
    state.sheets.forEach((sheetName) => {
      const ws = state.workbook.getWorksheet(sheetName);
      if (!ws) return;
      allData[sheetName] = [];
      for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r).values.slice(1);
        // collect wrong headers from cell fills if needed
        const wrong = [];
        ws.getRow(r).eachCell((cell) => {
          const header = ws.getRow(1).getCell(cell.col).value;
          if (
            cell.fill &&
            cell.fill.fgColor &&
            cell.fill.fgColor.argb === "FFFFC107"
          ) {
            wrong.push(header);
          }
        });
        allData[sheetName].push({ values: row, wrong });
      }
    });

    const payload = {
      batchNum: state.batchNum,
      date: new Date().toISOString().slice(0, 10),
      data: allData,
    };

    localStorage.setItem("der2_data", JSON.stringify(payload));
  }

  function loadSession() {
    const raw = localStorage.getItem("der2_data");
    if (!raw) return;

    const payload = JSON.parse(raw);
    const today = new Date().toISOString().slice(0, 10);

    if (payload.batchNum === state.batchNum && payload.date === today) {
      for (const [sheetName, rows] of Object.entries(payload.data)) {
        const ws = getOrCreateSheet(sheetName);
        rows.forEach((rowObj) => {
          const wrongSet = new Set(rowObj.wrong || []);
          addRowWithStyles(ws, rowObj.values, wrongSet);
        });
      }
      renderPreviews();
      updateDownloadButton();
      updateProductDropdown();
      saveSession();
    }
  }

  // --- UI bindings ---
  const issueTypeEl = document.getElementById("issue_type");
  const rowNewTicket = document.getElementById("row_new_ticket");
  const rowWrongInfo = document.getElementById("row_wrong_info");

  function updateConditionalFields() {
    const val = issueTypeEl.value;
    rowNewTicket.style.display = val === "Duplicates" ? "" : "none";
    rowWrongInfo.style.display = val === "Wrong Information" ? "" : "none";
    // Show/hide Scan New Ticket button
    const scanNewBtn = document.getElementById("scan_new_ticket");
    if (scanNewBtn) {
      scanNewBtn.style.display = val === "Duplicates" ? "inline-block" : "none";
    }
  }
  issueTypeEl.addEventListener("change", updateConditionalFields);
  updateConditionalFields();

  // Clear inline buttons
  document.getElementById("clear-ticket").addEventListener("click", () => {
    document.getElementById("ticket-num").value = "";
  });
  document.getElementById("clear_new_ticket").addEventListener("click", () => {
    document.getElementById("new_ticket").value = "";
  });
  document.getElementById("clear_prod").addEventListener("click", () => {
    document.getElementById("prod_num").value = "";
  });
  document.getElementById("clear_code").addEventListener("click", () => {
    document.getElementById("code_date").value = "";
  });
  document.getElementById("clear_cases").addEventListener("click", () => {
    document.getElementById("case_count").value = "";
  });
  document.getElementById("clear_line").addEventListener("click", () => {
    document.getElementById("line_num").value = "";
  });
  document.getElementById("clear_time").addEventListener("click", () => {
    document.getElementById("timestamp").value = "";
  });
  document
    .getElementById("product_dropdown")
    .addEventListener("change", (e) => {
      const selected = e.target.value;
      if (selected) {
        document.getElementById("prod_num").value = selected;
      }
    });

  // Clear all
  document.getElementById("clear_all").addEventListener("click", () => {
    [
      "ticket-num",
      "new_ticket",
      "prod_num",
      "code_date",
      "case_count",
      "line_num",
      "timestamp",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = "";
    });
    document
      .querySelectorAll(".wrong_chk")
      .forEach((el) => (el.checked = false));
    clearErrors();
  });

  // --- Batch modal logic ---
  const modal = document.getElementById("batchModal");
  const confirmBtn = document.getElementById("confirmBatch");
  const cancelBtn = document.getElementById("cancelBatch");

  document.getElementById("new_batch").addEventListener("click", () => {
    modal.style.display = "block"; // show confirmation modal
  });

  cancelBtn.addEventListener("click", () => {
    modal.style.display = "none"; // hide confirmation modal
  });

  confirmBtn.addEventListener("click", () => {
    modal.style.display = "none"; // hide confirmation modal

    state.batchNum += 1;
    localStorage.setItem("der2_batch", String(state.batchNum));
    state.workbook = null;
    state.sheets.clear();
    localStorage.removeItem("der2_data"); // clear saved tickets
    renderPreviews();
    updateDownloadButton();
    updateProductDropdown();
    saveSession();

    // Show the Batch Started modal
    const startedModal = document.getElementById("batchStartedModal");
    const closeBtn = document.getElementById("closeBatchStarted");
    document.getElementById("batchNumberDisplay").textContent = state.batchNum;
    startedModal.style.display = "block";

    // Manual close
    closeBtn.onclick = () => {
      startedModal.style.display = "none";
    };

    // Auto-close after 3 seconds
    setTimeout(() => {
      startedModal.style.display = "none";
    }, 3000);
  });

  // --- Add ticket logic ---
  document.getElementById("add-ticket-btn").addEventListener("click", () => {
    clearErrors();
    const issue = issueTypeEl.value;
    const ticket = document.getElementById("ticket-num").value.trim();
    const newTicket = document.getElementById("new_ticket").value.trim();
    const prod = document.getElementById("prod_num").value.trim();
    const codeDateRaw = document.getElementById("code_date").value.trim();
    const cases = document.getElementById("case_count").value.trim();
    const line = document.getElementById("line_num").value.trim();
    const timestampRaw = document.getElementById("timestamp").value.trim();

    let wrongSet = new Set();
    if (issue === "Wrong Information") {
      document
        .querySelectorAll(".wrong_chk:checked")
        .forEach((chk) => wrongSet.add(chk.value));
    }

    const errs = [];
    if (numericCheck(ticket) || ticket.length !== 10)
      errs.push("Ticket must be numeric and 10 digits.");
    if (issue === "Duplicates") {
      if (newTicket && (numericCheck(newTicket) || newTicket.length !== 10)) {
        errs.push("New Ticket must be numeric and 10 digits.");
      }
    }
    if (numericCheck(prod) || prod.length < 12 || prod.length > 14) {
      errs.push("Product Number must be numeric and 12–14 digits.");
    }
    if (
      numericCheck(codeDateRaw) ||
      codeDateRaw.length < 4 ||
      codeDateRaw.length > 6
    ) {
      errs.push("Code Date must be numeric and 4 to 6 digits.");
    }
    if (numericCheck(cases)) errs.push("Cases must be numeric.");
    if (numericCheck(line)) errs.push("Line must be numeric.");
    if (
      numericCheck(timestampRaw) ||
      timestampRaw.length < 3 ||
      timestampRaw.length > 4
    ) {
      errs.push("Timestamp must be numeric and 3 to 4 digits.");
    }

    if (errs.length) {
      showErrors(errs);
      return;
    }

    const { code_date, timestamp } = formatData(codeDateRaw, timestampRaw);
    const ws = getOrCreateSheet(issue);

    const ticketColIdx = findColumnIndexByHeader(ws, "Ticket");
    const existingTickets = new Set();
    for (let r = 2; r <= ws.rowCount; r++) {
      const v = ws.getRow(r).getCell(ticketColIdx).value;
      existingTickets.add(String(v));
    }
    if (existingTickets.has(String(ticket))) {
      // Save the new values temporarily
      const newRowData = {
        ticket,
        newTicket,
        prod,
        code_date,
        cases,
        line,
        timestamp,
        wrongSet,
        issue,
      };

      // Show overwrite modal
      const overwriteModal = document.getElementById("overwriteModal");
      overwriteModal.style.display = "block";

      // Wire up buttons
      const confirmBtn = document.getElementById("confirmOverwrite");
      const newHandler = () => {
        overwriteModal.style.display = "none";
        overwriteTicket(newRowData);
        confirmBtn.removeEventListener("click", newHandler);
      };
      confirmBtn.addEventListener("click", newHandler);
      document.getElementById("cancelOverwrite").onclick = () => {
        overwriteModal.style.display = "none";
        confirmBtn.replaceWith(confirmBtn.cloneNode(true)); // removes all listeners
      };

      return; // stop normal add flow
    }

    if (issue === "Duplicates") ensureNewTicketColumn(ws);

    const headers = [];
    ws.getRow(1).eachCell((cell) => headers.push(cell.value));
    const valuesByHeader = {
      Ticket: ticket,
      "New Ticket": newTicket || "",
      "Product Number": prod,
      "Code Date": code_date,
      Cases: cases,
      Line: line,
      Timestamp: timestamp,
    };
    const rowVals = headers.map((h) => valuesByHeader[h] ?? "");

    addRowWithStyles(ws, rowVals, wrongSet);
    renderPreviews();
    updateDownloadButton();
    // flashSuccess();
    updateProductDropdown();
    saveSession();
  });

  function reapplyRowStyles(ws) {
    const fillEven = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2F2F2" },
    };
    const fillOdd = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD7EBFF" },
    };
    const highlight = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFC107" },
    };

    const headerRow = ws.getRow(1);

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const even = r % 2 === 0;
      row.eachCell((cell) => {
        cell.font = { size: 14 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        const header = headerRow.getCell(cell.col).value;
        cell.fill = even ? fillEven : fillOdd;
      });
    }

    // Recalculate column widths
    ws.columns.forEach((col) => {
      let maxLen = 0;
      col.eachCell({ includeEmpty: true }, (cell) => {
        const v = cell.value == null ? "" : String(cell.value);
        maxLen = Math.max(maxLen, v.length);
      });
      col.width = Math.max(10, maxLen + 8);
    });
  }

  // --- Overwrite ticket ---
  function overwriteTicket(data) {
    const ws = getOrCreateSheet(data.issue);
    const ticketColIdx = findColumnIndexByHeader(ws, "Ticket");

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (String(row.getCell(ticketColIdx).value) === data.ticket) {
        const headers = [];
        ws.getRow(1).eachCell((cell) => headers.push(cell.value));
        const valuesByHeader = {
          Ticket: data.ticket,
          "New Ticket": data.newTicket || "",
          "Product Number": data.prod,
          "Code Date": data.code_date,
          Cases: data.cases,
          Line: data.line,
          Timestamp: data.timestamp,
        };
        const rowVals = headers.map((h) => valuesByHeader[h] ?? "");

        // Replace row content at same index
        ws.spliceRows(r, 1);
        addRowWithStyles(ws, rowVals, data.wrongSet || new Set());

        renderPreviews();
        updateDownloadButton();
        // flashSuccess();
        updateProductDropdown();
        saveSession();
        return;
      }
    }
  }

  // --- Download workbook ---
  document
    .getElementById("download_btn")
    .addEventListener("click", async () => {
      if (!state.workbook || state.sheets.size === 0) return;
      const buffer = await state.workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(blob, sheetFilename());
    });

  function updateProductDropdown() {
    const dropdown = document.getElementById("product_dropdown");
    dropdown.innerHTML = "";

    const optDefault = document.createElement("option");
    optDefault.value = "";
    optDefault.textContent = "-- none yet --";
    dropdown.appendChild(optDefault);

    if (!state.workbook) return;

    const seen = new Set();

    state.sheets.forEach((sheetName) => {
      const ws = state.workbook.getWorksheet(sheetName);
      if (!ws) return;
      const prodColIdx = findColumnIndexByHeader(ws, "Product Number");
      if (!prodColIdx) return;

      for (let r = 2; r <= ws.rowCount; r++) {
        const val = ws.getRow(r).getCell(prodColIdx).value;
        if (val) seen.add(String(val));
      }
    });

    Array.from(seen).forEach((prod) => {
      const opt = document.createElement("option");
      opt.value = prod;
      opt.textContent = prod;
      dropdown.appendChild(opt);
    });
  }

  // Initialize UI
  updateDownloadButton();

  function startBarcodeScan(targetInputId) {
    const scannerModal = document.getElementById("scannerModal");
    const scannerEl = document.getElementById("scanner");
    let scanLocked = false;

    scannerModal.style.display = "block";
    Quagga.offDetected(); // clear previous listeners if any

    Quagga.init(
      {
        inputStream: {
          type: "LiveStream",
          target: scannerEl,
          constraints: {
            facingMode: "environment",
          },
        },
        decoder: {
          readers: ["code_128_reader"],
        },
        locate: true,
      },
      (err) => {
        if (err) {
          console.error("Quagga init error:", err);
          return;
        }
        Quagga.start();
      },
    );

    Quagga.onDetected((result) => {
      if (scanLocked) return;

      const code = result?.codeResult?.code;
      if (!code || !/^\d{6,10}$/.test(code)) return; // only accept 6–10 digit codes

      scanLocked = true;
      document.getElementById(targetInputId).value = code;

      Quagga.stop();
      Quagga.offDetected();
      scannerModal.style.display = "none";
    });
  }

  // Wire buttons
  document.getElementById("scan-ticket").addEventListener("click", () => {
    startBarcodeScan("ticket-num");
  });

  document.getElementById("scan_new_ticket").addEventListener("click", () => {
    startBarcodeScan("new_ticket");
  });

  document.getElementById("cancelScan").addEventListener("click", () => {
    Quagga.stop();
    Quagga.offDetected();
    document.getElementById("scannerModal").style.display = "none";
  });
});

// Functions

function toEWM(val) {
  const ewmStart = "10044710";
  const digits = String(val).slice(0, 9);
  const weights = [3, 1, 3, 1, 3, 1, 3, 1, 3];

  const total =
    29 +
    digits.split("").reduce((sum, d, i) => sum + parseInt(d) * weights[i], 0);
  const checkDigit = (10 - (total % 10)) % 10;

  return `${ewmStart}${digits + String(checkDigit)}`;
}

function convertTicket() {
  const ticketInput = document
    .getElementById("ticket-num")
    .value.trim()
    .split("\n");
  const multiInput = document.getElementById("multi").value.trim().split("\n");

  const inputs = [...new Set([...ticketInput, ...multiInput])].filter(Boolean);

  for (let ticket of inputs) {
    ticket = ticket.trim();
    const notValid = isNaN(parseInt(ticket));

    const isValid =
      ticket.length == 10 || ticket.length == 18 || ticket.length == 20;

    if (!isValid || notValid) {
      alert("Enter a valid ticket");
      return;
    }

    let ewmTicket;
    if (ticket.length == 18) ewmTicket = ticket;
    else if (ticket.length == 20) ewmTicket = ticket.slice(2);
    else ewmTicket = toEWM(ticket);

    const matricsTicket = `${ticket.slice(-10, -1)}0`;

    addTicket(matricsTicket, ewmTicket);
  }
}

function addTicket(matrics, ewm) {
  const table = document.getElementById("ticketTable");
  const tr = document.createElement("tr");

  const codeMatrics = document.createElement("td");
  codeMatrics.innerHTML = `<button class="barcodeBtns" data-value="${matrics}">Show</button>`;

  const codeEwm = document.createElement("td");
  codeEwm.innerHTML = `<button class="barcodeBtns" data-value="${ewm}">Show</button>`;

  const tdMatrics = document.createElement("td");
  tdMatrics.setAttribute("class", "matrics");
  tdMatrics.textContent = matrics;

  const tdEwm = document.createElement("td");
  tdEwm.setAttribute("class", "ewm");
  tdEwm.textContent = ewm;

  tr.appendChild(codeMatrics);
  tr.appendChild(tdMatrics);
  tr.appendChild(tdEwm);
  tr.appendChild(codeEwm);

  table.appendChild(tr);
}

function copyMatrics() {
  const values = [...document.querySelectorAll(".matrics")].map(
    (el) => el.innerHTML,
  );
  const tickets = values.join("\n");

  navigator.clipboard.writeText(tickets);
}

function copyEWM() {
  const values = [...document.querySelectorAll(".ewm")].map(
    (el) => el.innerHTML,
  );
  const tickets = values.join("\n");

  navigator.clipboard.writeText(tickets);
}

function clearTable() {
  const table = document.getElementById("ticketTable");
  table.innerHTML = "";

  const tr = document.createElement("tr");

  const codeMatrics = document.createElement("th");
  codeMatrics.textContent = "Barcode";

  const thMatrics = document.createElement("th");
  thMatrics.textContent = "Matrics";

  const thEwm = document.createElement("th");
  thEwm.textContent = "EWM";

  const codeEwm = document.createElement("th");
  codeEwm.textContent = "Barcode";

  tr.appendChild(codeMatrics);
  tr.appendChild(thMatrics);
  tr.appendChild(thEwm);
  tr.appendChild(codeEwm);

  table.appendChild(tr);
}

function startBarcodeScan(targetInputId) {
  const scannerModal = document.getElementById("scannerModal");
  const scannerEl = document.getElementById("scanner");
  let scanLocked = false;

  scannerModal.style.display = "block";
  Quagga.offDetected(); // clear previous listeners if any

  Quagga.init(
    {
      inputStream: {
        type: "LiveStream",
        target: scannerEl,
        constraints: {
          facingMode: "environment",
        },
      },
      decoder: {
        readers: ["code_128_reader"],
      },
      locate: true,
    },
    (err) => {
      if (err) {
        console.error("Quagga init error:", err);
        return;
      }
      Quagga.start();
    },
  );

  Quagga.onDetected((result) => {
    if (scanLocked) return;

    const code = result?.codeResult?.code;
    if (!code || !/^\d{6,10}$/.test(code)) return; // only accept 6–10 digit codes

    scanLocked = true;
    document.getElementById(targetInputId).value = code;

    Quagga.stop();
    Quagga.offDetected();
    scannerModal.style.display = "none";
  });
}

function createBarCode(val) {
  document.getElementById("showBarcode").style.display = "block";

  JsBarcode("#barcode", val, {
    format: "CODE128",
    lineColor: "#000",
    width: 4,
    height: 60,
    displayValue: true,
  });
}

function flashMessage(msg) {
  msg.classList.add("show");

  setTimeout(() => {
    msg.classList.remove("show");
  }, 2000);
}

function fadeInMessage(msg) {
  msg.classList.add("fade");

  setTimeout(() => {
    msg.classList.remove("fade");
    document.getElementById("errors").innerHTML = "";
  }, 15000);

}

// Converter Listeners

if (window.location.pathname === "/ticket-converter") {
  document.getElementById("toggleSwitch").addEventListener("click", () => {
    const currentState = document
      .getElementById("multiContainer")
      .getAttribute("style");

    const isOn = currentState === "display: block;";
    const isOff = currentState === "display: none;";

    if (isOn) {
      document
        .getElementById("multiContainer")
        .setAttribute("style", "display: none;");
      document.getElementById("multi").value = "";
      document.getElementById("ticket-num").removeAttribute("disabled");
    } else {
      document
        .getElementById("multiContainer")
        .setAttribute("style", "display: block;");
      document.getElementById("ticket-num").value = "";
      document.getElementById("ticket-num").setAttribute("disabled", true);
    }
  });
  document.getElementById("convertTicket").addEventListener("click", () => {
    convertTicket();
  });
  document.getElementById("hideBarcode").addEventListener("click", () => {
    document.getElementById("showBarcode").style.display = "none";
  });
  document.addEventListener("click", (e) => {
    if (e.target.classList.contains("barcodeBtns")) {
      e.target.classList.remove("active");
      void e.offsetWidth;
      e.target.classList.add("active");

      const ticket = e.target.dataset.value;
      createBarCode(ticket);
    }
  });
  document.getElementById("clearTable").addEventListener("click", () => {
    clearTable();
  });
  document.getElementById("copyMatricsBtn").addEventListener("click", () => {
    const button = document.getElementById("copyMatricsBtn");
    button.classList.remove("active");
    void button.offsetWidth;
    button.classList.add("active");
    copyMatrics();
  });
  document.getElementById("copyEWMBtn").addEventListener("click", () => {
    const button = document.getElementById("copyEWMBtn");
    button.classList.remove("active");
    void button.offsetWidth;
    button.classList.add("active");
    copyEWM();
  });
}

// Site-wide Listeners

document.getElementById("clear-ticket").addEventListener("click", () => {
  document.getElementById("ticket-num").value = "";
  // document.getElementById("ticket-multi").value = "";
});
document.getElementById("scan-ticket").addEventListener("click", () => {
  startBarcodeScan("ticket-num");
});
document.getElementById("cancelScan").addEventListener("click", () => {
  Quagga.stop();
  Quagga.offDetected();
  document.getElementById("scannerModal").style.display = "none";
});
document.getElementById("add-ticket-btn").addEventListener("click", () => {
  const msg = document.getElementById("add-ticket");
  flashMessage(msg);
});
