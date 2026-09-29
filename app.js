/**
 * 물리치료 현황 (PT Daily Log) PWA
 * Excel-Style Web Application
 */

// Initial Sample Data extracted directly from the user's Excel spreadsheet image
const INITIAL_SAMPLE_DATA = {
  "2026-09-29": [
    { no: 3, gender: "F", chartNo: "3387", name: "신재례", part: "양무", prescription: "사지 ( HP / Laser / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 9, gender: "F", chartNo: "15307", name: "김시호", part: "오무", prescription: "학생 ( HP / Laser )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 6, gender: "M", chartNo: "6780", name: "이춘식", part: "왼 고관절", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "Lt.femur *핀(MW X)", date: "2026.09.29" },
    { no: 4, gender: "M", chartNo: "16133", name: "최의규", part: "Lt. Heel", prescription: "사지 ( HP / Laser / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 10, gender: "F", chartNo: "15520", name: "정화자", part: "왼 손목", prescription: "사지 ( HP / Laser / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 8, gender: "M", chartNo: "12208", name: "천진우", part: "왼 손목", prescription: "학생 ( HP / Laser )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 1, gender: "F", chartNo: "12500", name: "양명자", part: "왼 엉", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 10, gender: "M", chartNo: "4079", name: "정서우", part: "Lt. Thigh", prescription: "학생 ( HP / Laser )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: 4, gender: "M", chartNo: "8064", name: "이영덕", part: "허리", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: "", gender: "F", chartNo: "14358", name: "박서하", part: "Rt. heel", prescription: "X", extra: "충격파", writer: "S", memo: "충완", specialNote: "", date: "2026.09.29" },
    { no: 9, gender: "M", chartNo: "5285", name: "전기운", part: "허리", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.29" },
    { no: "", gender: "M", chartNo: "16155", name: "전재윤", part: "목", prescription: "척추 ( HP / 자기장 / ICT )", extra: "충격파", writer: "S", memo: "충완", specialNote: "", date: "2026.09.29" },
    { no: "", gender: "M", chartNo: "16156", name: "김관웅", part: "우 발바닥", prescription: "사지 ( HP / Laser / ICT )", extra: "윈백", writer: "S", memo: "충완", specialNote: "", date: "2026.09.29" },
    { no: 1, gender: "", chartNo: "", name: "김열중", part: "발목", prescription: "사지 ( HP / Laser / ICT )", extra: "이온", writer: "S", memo: "이온 완", specialNote: "", date: "2026.09.29" }
  ],
  "2026-09-28": [
    { no: 1, gender: "F", chartNo: "105", name: "김경희", part: "오 손목", prescription: "사지 ( HP / Laser / ICT )", extra: "이온", writer: "S", memo: "이온 완", specialNote: "", date: "2026.09.28" },
    { no: 4, gender: "F", chartNo: "10948", name: "오성희", part: "허리", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.28" },
    { no: "", gender: "M", chartNo: "16056", name: "김민성", part: "허리", prescription: "척추 ( HP / 자기장 / ICT )", extra: "충격파", writer: "S", memo: "충완", specialNote: "", date: "2026.09.28" },
    { no: 9, gender: "F", chartNo: "16122", name: "김현신", part: "오 등", prescription: "척추 ( HP / 자기장 / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.28" },
    { no: 3, gender: "F", chartNo: "15889", name: "이연진", part: "오 발등", prescription: "사지 ( HP / Laser / ICT )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.28" },
    { no: "", gender: "M", chartNo: "16154", name: "김세윤", part: "뗀무", prescription: "항냉 ( ICE / Laser )", extra: "", writer: "S", memo: "", specialNote: "", date: "2026.09.28" }
  ]
};

const STORAGE_KEY = "PT_APP_DATA_STORAGE_V1";
const DEFAULT_WRITER = "S";
const BASE_ROW_NUMBER = 1; // 행 번호 1부터 시작

class PTApp {
  constructor() {
    this.dataStore = this.loadDataStore();
    this.currentDate = this.getTodayString();
    this.activeCell = null; // { rowIdx, colKey }
    this.selectedRowIdx = null;

    this.cacheElements();
    this.bindEvents();
    this.initPWA();

    // Start with today's date and auto-focus
    this.setDate(this.currentDate, true);
  }

  // Helper: Get local Date string YYYY-MM-DD
  getTodayString() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  // Load from LocalStorage or initialize with sample
  loadDataStore() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error("Failed to parse localStorage data", e);
    }
    // Save sample data on first run
    localStorage.setItem(STORAGE_KEY, JSON.stringify(INITIAL_SAMPLE_DATA));
    return JSON.parse(JSON.stringify(INITIAL_SAMPLE_DATA));
  }

  saveDataStore() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataStore));
      this.showSaveIndicator("저장 완료됨");
      this.updateSidebarStats();
      this.renderRecentDays();
    } catch (e) {
      console.error("Save error", e);
      this.showSaveIndicator("저장 실패", true);
    }
  }

  showSaveIndicator(text, isError = false) {
    if (!this.elSaveStatus) return;
    this.elSaveStatus.textContent = text;
    this.elSaveStatus.className = "save-status" + (isError ? " unsaved" : "");
    if (!isError) {
      setTimeout(() => {
        if (this.elSaveStatus) this.elSaveStatus.textContent = "저장 완료됨";
      }, 1500);
    }
  }

  cacheElements() {
    this.elDatePicker = document.getElementById("datePicker");
    this.elDateLabel = document.getElementById("dateDisplayLabel");
    this.elBtnPrevDay = document.getElementById("btnPrevDay");
    this.elBtnNextDay = document.getElementById("btnNextDay");
    this.elBtnGoToday = document.getElementById("btnGoToday");
    this.elBtnPreview = document.getElementById("btnPreview");
    this.elBtnExportCsv = document.getElementById("btnExportCsv");

    this.elTableBody = document.getElementById("tableBody");
    this.elSheetContainer = document.getElementById("sheetContainer");
    this.elCellAddress = document.getElementById("cellAddressDisplay");
    this.elFormulaInput = document.getElementById("formulaInput");
    this.elSelectedCellCoords = document.getElementById("selectedCellCoords");

    this.elBtnAddRow = document.getElementById("btnAddRow");
    this.elBtnDeleteSelected = document.getElementById("btnDeleteSelected");
    this.elBtnBottomAddRow = document.getElementById("btnBottomAddRow");
    this.elTabAddRow = document.getElementById("tabAddRow");
    this.elTabClearEmpty = document.getElementById("tabClearEmpty");
    this.elTabBackup = document.getElementById("tabBackup");

    this.elSearchInput = document.getElementById("searchInput");
    this.elBtnClearSearch = document.getElementById("btnClearSearch");

    // Sidebar
    this.elSidebarDateTag = document.getElementById("sidebarDateTag");
    this.elStatTotalCount = document.getElementById("statTotalCount");
    this.elStatMaleCount = document.getElementById("statMaleCount");
    this.elStatFemaleCount = document.getElementById("statFemaleCount");
    this.elStatShockwave = document.getElementById("statShockwave");
    this.elStatIon = document.getElementById("statIon");
    this.elStatWinback = document.getElementById("statWinback");
    this.elStatExtraOther = document.getElementById("statExtraOther");
    this.elStatPrescExtremity = document.getElementById("statPrescExtremity");
    this.elStatPrescSpine = document.getElementById("statPrescSpine");
    this.elStatPrescOther = document.getElementById("statPrescOther");
    this.elRecentDaysList = document.getElementById("recentDaysList");

    // Tabs
    this.elCurrentSheetTab = document.getElementById("currentSheetTab");
    this.elSheetTabTitle = document.getElementById("sheetTabTitle");
    this.elBtnQuickNewDay = document.getElementById("btnQuickNewDay");
    this.elSaveStatus = document.getElementById("saveStatus");

    // Preview Modal
    this.elPreviewModal = document.getElementById("previewModal");
    this.elPreviewModalDate = document.getElementById("previewModalDate");
    this.elPrintDateFull = document.getElementById("printDateFull");
    this.elPrintTotalCount = document.getElementById("printTotalCount");
    this.elPrintMaleCount = document.getElementById("printMaleCount");
    this.elPrintFemaleCount = document.getElementById("printFemaleCount");
    this.elPrintShockwave = document.getElementById("printShockwave");
    this.elPrintIon = document.getElementById("printIon");
    this.elPrintWinback = document.getElementById("printWinback");
    this.elPrintTableBody = document.getElementById("printTableBody");
    this.elPrintGeneratedTime = document.getElementById("printGeneratedTime");
    this.elBtnPrintConfirm = document.getElementById("btnPrintConfirm");
    this.elBtnClosePreview = document.getElementById("btnClosePreview");

    // Backup Modal
    this.elBackupModal = document.getElementById("backupModal");
    this.elBtnCloseBackup = document.getElementById("btnCloseBackup");
    this.elBtnDownloadBackup = document.getElementById("btnDownloadBackup");
    this.elFileRestore = document.getElementById("fileRestore");
    this.elBtnClearAllData = document.getElementById("btnClearAllData");
  }

  bindEvents() {
    // Date Navigation
    this.elDatePicker.addEventListener("change", (e) => {
      if (e.target.value) this.setDate(e.target.value);
    });

    this.elBtnPrevDay.addEventListener("click", () => this.shiftDay(-1));
    this.elBtnNextDay.addEventListener("click", () => this.shiftDay(1));
    this.elBtnGoToday.addEventListener("click", () => this.setDate(this.getTodayString(), true));

    // Rows management
    this.elBtnAddRow.addEventListener("click", () => this.addNewRow(true));
    this.elBtnBottomAddRow.addEventListener("click", () => this.addNewRow(true));
    this.elTabAddRow.addEventListener("click", () => this.addNewRow(true));
    this.elBtnDeleteSelected.addEventListener("click", () => this.deleteSelectedRow());

    this.elTabClearEmpty.addEventListener("click", () => this.removeEmptyRows());
    this.elTabBackup.addEventListener("click", () => this.openBackupModal());

    // Formula Input Sync
    this.elFormulaInput.addEventListener("input", (e) => {
      if (!this.activeCell) return;
      const { rowIdx, colKey } = this.activeCell;
      const rows = this.getCurrentRows();
      if (rows[rowIdx]) {
        rows[rowIdx][colKey] = e.target.value;
        const cellEl = document.querySelector(`.excel-cell[data-row="${rowIdx}"][data-col="${colKey}"]`);
        if (cellEl) {
          const inputEl = cellEl.querySelector("input");
          if (inputEl) inputEl.value = e.target.value;
          else cellEl.textContent = e.target.value;
        }
        this.saveDataStore();
      }
    });

    // Quick Chips (Prescription & Extra)
    document.querySelectorAll(".quick-chips .chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        const type = chip.getAttribute("data-type");
        const val = chip.getAttribute("data-val");
        this.applyQuickChip(type, val);
      });
    });

    // Search
    this.elSearchInput.addEventListener("input", () => this.handleSearch());
    this.elBtnClearSearch.addEventListener("click", () => {
      this.elSearchInput.value = "";
      this.handleSearch();
    });

    // Preview & Print Modal
    this.elBtnPreview.addEventListener("click", () => this.openPreviewModal());
    this.elBtnClosePreview.addEventListener("click", () => this.closePreviewModal());
    this.elBtnPrintConfirm.addEventListener("click", () => window.print());

    // Export CSV
    this.elBtnExportCsv.addEventListener("click", () => this.exportCurrentDayCsv());

    // Backup & Restore
    this.elBtnCloseBackup.addEventListener("click", () => this.closeBackupModal());
    this.elBtnDownloadBackup.addEventListener("click", () => this.downloadFullBackup());
    this.elFileRestore.addEventListener("change", (e) => this.handleRestoreFile(e));
    this.elBtnClearAllData.addEventListener("click", () => this.clearCurrentDayData());

    // Keyboard Shortcuts
    document.addEventListener("keydown", (e) => this.handleGlobalKeyDown(e));

    // Close modal on escape
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        this.closePreviewModal();
        this.closeBackupModal();
      }
    });
  }

  // Get or initialize rows for a given date
  getCurrentRows() {
    if (!this.dataStore[this.currentDate]) {
      // Create empty rows for today/selected date
      this.dataStore[this.currentDate] = this.createDefaultEmptyRows();
    }
    return this.dataStore[this.currentDate];
  }

  createDefaultEmptyRows(count = 10) {
    const formattedDate = this.currentDate.replace(/-/g, ".");
    const rows = [];
    for (let i = 1; i <= count; i++) {
      rows.push({
        no: "",
        gender: "",
        chartNo: "",
        name: "",
        part: "",
        prescription: "",
        extra: "",
        writer: DEFAULT_WRITER,
        memo: "",
        specialNote: "",
        date: formattedDate
      });
    }
    return rows;
  }

  setDate(dateStr, autoFocusFirstEmpty = false) {
    this.currentDate = dateStr;
    this.elDatePicker.value = dateStr;

    // Format display: e.g., 2026.09.29 (화)
    const [y, m, d] = dateStr.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);
    const daysKor = ["일", "월", "화", "수", "목", "금", "토"];
    const dayLabel = daysKor[dateObj.getDay()] || "";
    const dateFormatted = `${y}.${String(m).padStart(2, "0")}.${String(d).padStart(2, "0")}`;

    this.elDateLabel.textContent = `${dateFormatted} (${dayLabel})`;
    this.elSidebarDateTag.textContent = dateFormatted;
    this.elSheetTabTitle.textContent = dateFormatted;

    this.renderTable();
    this.updateSidebarStats();
    this.renderRecentDays();

    if (autoFocusFirstEmpty) {
      setTimeout(() => this.focusFirstEmptyCell(), 80);
    }
  }

  shiftDay(deltaDays) {
    const [y, m, d] = this.currentDate.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);
    dateObj.setDate(dateObj.getDate() + deltaDays);
    const nextY = dateObj.getFullYear();
    const nextM = String(dateObj.getMonth() + 1).padStart(2, "0");
    const nextD = String(dateObj.getDate()).padStart(2, "0");
    this.setDate(`${nextY}-${nextM}-${nextD}`);
  }

  // Render main Excel table
  renderTable() {
    const rows = this.getCurrentRows();
    this.elTableBody.innerHTML = "";

    const colKeys = ["no", "gender", "chartNo", "name", "part", "prescription", "extra", "writer", "memo", "specialNote", "date"];
    const colLetters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

    rows.forEach((row, rowIdx) => {
      const tr = document.createElement("tr");
      tr.className = "excel-row";
      tr.dataset.rowIdx = rowIdx;

      // Row Number Header (1, 2, 3...)
      const excelRowNum = BASE_ROW_NUMBER + rowIdx;
      const thNum = document.createElement("th");
      thNum.className = "row-num";
      thNum.textContent = excelRowNum;
      thNum.title = `행 번호: ${excelRowNum}`;
      tr.appendChild(thNum);

      // Columns
      colKeys.forEach((key, colIdx) => {
        const td = document.createElement("td");
        td.className = `excel-cell cell-${key}`;
        td.dataset.row = rowIdx;
        td.dataset.col = key;
        td.dataset.colLetter = colLetters[colIdx];
        td.dataset.excelRow = excelRowNum;

        const val = row[key] || "";

        // Custom render for Gender: Quick click toggle or select
        if (key === "gender") {
          td.textContent = val;
          if (val === "F") td.classList.add("f");
          if (val === "M") td.classList.add("m");
          td.title = "클릭하여 F/M 변경 가능";
        } else {
          // Normal text content or editable
          td.textContent = val;
        }

        // Cell Click handler
        td.addEventListener("click", (e) => {
          this.selectCell(rowIdx, key, td);
        });

        // Cell Double Click to inline edit
        td.addEventListener("dblclick", (e) => {
          this.startInlineEdit(rowIdx, key, td);
        });

        tr.appendChild(td);
      });

      // Delete action column
      const tdDel = document.createElement("td");
      tdDel.className = "excel-cell cell-del-action";
      const btnDel = document.createElement("button");
      btnDel.type = "button";
      btnDel.className = "btn-row-del";
      btnDel.textContent = "✕";
      btnDel.title = "이 행 삭제";
      btnDel.addEventListener("click", (e) => {
        e.stopPropagation();
        this.deleteRow(rowIdx);
      });
      tdDel.appendChild(btnDel);
      tr.appendChild(tdDel);

      this.elTableBody.appendChild(tr);
    });

    // If active cell was set, restore highlight if valid
    if (this.activeCell) {
      const targetCell = document.querySelector(`.excel-cell[data-row="${this.activeCell.rowIdx}"][data-col="${this.activeCell.colKey}"]`);
      if (targetCell) {
        this.highlightCell(targetCell);
      }
    }
  }

  // Select and focus cell like Excel
  selectCell(rowIdx, colKey, cellElement) {
    this.activeCell = { rowIdx, colKey };
    this.selectedRowIdx = rowIdx;

    // Highlight row
    document.querySelectorAll(".excel-row").forEach((r) => r.classList.remove("active-row"));
    const rowEl = cellElement.closest("tr");
    if (rowEl) rowEl.classList.add("active-row");

    this.highlightCell(cellElement);

    // Update Formula Bar
    const colLetter = cellElement.dataset.colLetter;
    const excelRow = cellElement.dataset.excelRow;
    const cellAddress = `${colLetter}${excelRow}`;

    this.elCellAddress.textContent = cellAddress;
    this.elSelectedCellCoords.textContent = `${cellAddress} (${colKey})`;

    const rows = this.getCurrentRows();
    const cellValue = rows[rowIdx] ? (rows[rowIdx][colKey] || "") : "";
    this.elFormulaInput.value = cellValue;

    // If Gender column, toggle on click
    if (colKey === "gender") {
      const current = rows[rowIdx].gender;
      let next = "";
      if (!current || current === "") next = "F";
      else if (current === "F") next = "M";
      else if (current === "M") next = "";
      rows[rowIdx].gender = next;
      cellElement.textContent = next;
      cellElement.classList.remove("f", "m");
      if (next === "F") cellElement.classList.add("f");
      if (next === "M") cellElement.classList.add("m");
      this.elFormulaInput.value = next;
      this.saveDataStore();
      return;
    }

    // Direct single click turns into editable input for fast data entry
    this.startInlineEdit(rowIdx, colKey, cellElement);
  }

  highlightCell(cellElement) {
    document.querySelectorAll(".excel-cell").forEach((c) => c.classList.remove("cell-focused"));
    cellElement.classList.add("cell-focused");
  }

  startInlineEdit(rowIdx, colKey, cellElement) {
    // If already editing
    if (cellElement.querySelector("input") || cellElement.querySelector("select")) return;

    const rows = this.getCurrentRows();
    const initialVal = rows[rowIdx] ? (rows[rowIdx][colKey] || "") : "";

    // Clear content and place input
    cellElement.textContent = "";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-input-element";
    input.value = initialVal;

    // Auto-complete list suggestion support for Part and Prescription
    if (colKey === "part") {
      input.setAttribute("list", "partPresets");
      this.ensureDatalists();
    } else if (colKey === "prescription") {
      input.setAttribute("list", "prescPresets");
      this.ensureDatalists();
    } else if (colKey === "extra") {
      input.setAttribute("list", "extraPresets");
      this.ensureDatalists();
    }

    cellElement.appendChild(input);
    input.focus();
    input.select();

    // Input events
    input.addEventListener("input", (e) => {
      const val = e.target.value;
      rows[rowIdx][colKey] = val;
      this.elFormulaInput.value = val;
      this.saveDataStore();
    });

    const commitAndBlur = () => {
      const finalVal = input.value.trim();
      rows[rowIdx][colKey] = finalVal;
      cellElement.textContent = finalVal;
      if (colKey === "gender") {
        cellElement.classList.remove("f", "m");
        if (finalVal === "F") cellElement.classList.add("f");
        if (finalVal === "M") cellElement.classList.add("m");
      }
      this.saveDataStore();
    };

    input.addEventListener("blur", () => {
      commitAndBlur();
    });

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commitAndBlur();
        // Move to next row in same column
        this.navigateCell(rowIdx + 1, colKey);
      } else if (e.key === "Tab") {
        e.preventDefault();
        commitAndBlur();
        if (e.shiftKey) {
          this.navigateCol(rowIdx, colKey, -1);
        } else {
          this.navigateCol(rowIdx, colKey, 1);
        }
      }
    });
  }

  ensureDatalists() {
    if (document.getElementById("partPresets")) return;
    const datalistPart = document.createElement("datalist");
    datalistPart.id = "partPresets";
    datalistPart.innerHTML = `
      <option value="허리"></option>
      <option value="목"></option>
      <option value="오 손목"></option>
      <option value="왼 손목"></option>
      <option value="양무"></option>
      <option value="오무"></option>
      <option value="뗀무"></option>
      <option value="왼 고관절"></option>
      <option value="오 고관절"></option>
      <option value="오 등"></option>
      <option value="오 발등"></option>
      <option value="우 발바닥"></option>
      <option value="Lt. Heel"></option>
      <option value="Rt. heel"></option>
      <option value="발목"></option>
      <option value="Lt. Thigh"></option>
    `;
    document.body.appendChild(datalistPart);

    const datalistPresc = document.createElement("datalist");
    datalistPresc.id = "prescPresets";
    datalistPresc.innerHTML = `
      <option value="사지 ( HP / Laser / ICT )"></option>
      <option value="척추 ( HP / 자기장 / ICT )"></option>
      <option value="학생 ( HP / Laser )"></option>
      <option value="항냉 ( ICE / Laser )"></option>
      <option value="X"></option>
    `;
    document.body.appendChild(datalistPresc);

    const datalistExtra = document.createElement("datalist");
    datalistExtra.id = "extraPresets";
    datalistExtra.innerHTML = `
      <option value="충격파"></option>
      <option value="이온"></option>
      <option value="윈백"></option>
      <option value="도수치료"></option>
      <option value="견인"></option>
    `;
    document.body.appendChild(datalistExtra);
  }

  navigateCell(targetRowIdx, colKey) {
    const rows = this.getCurrentRows();
    if (targetRowIdx >= rows.length) {
      // Auto add new row if hitting enter at bottom!
      this.addNewRow(false);
    }
    const nextCell = document.querySelector(`.excel-cell[data-row="${targetRowIdx}"][data-col="${colKey}"]`);
    if (nextCell) {
      this.selectCell(targetRowIdx, colKey, nextCell);
    }
  }

  navigateCol(rowIdx, currentColKey, direction) {
    const colOrder = ["no", "gender", "chartNo", "name", "part", "prescription", "extra", "writer", "memo", "specialNote", "date"];
    const curIdx = colOrder.indexOf(currentColKey);
    let nextIdx = curIdx + direction;
    let nextRowIdx = rowIdx;

    if (nextIdx >= colOrder.length) {
      nextIdx = 0;
      nextRowIdx += 1;
    } else if (nextIdx < 0) {
      nextIdx = colOrder.length - 1;
      nextRowIdx = Math.max(0, nextRowIdx - 1);
    }

    const rows = this.getCurrentRows();
    if (nextRowIdx >= rows.length) {
      this.addNewRow(false);
    }

    const nextColKey = colOrder[nextIdx];
    const nextCell = document.querySelector(`.excel-cell[data-row="${nextRowIdx}"][data-col="${nextColKey}"]`);
    if (nextCell) {
      this.selectCell(nextRowIdx, nextColKey, nextCell);
    }
  }

  focusFirstEmptyCell() {
    const rows = this.getCurrentRows();
    let targetRow = 0;
    for (let i = 0; i < rows.length; i++) {
      if (!rows[i].name && !rows[i].chartNo) {
        targetRow = i;
        break;
      }
    }
    // Prefer focusing on chartNo or name
    const cell = document.querySelector(`.excel-cell[data-row="${targetRow}"][data-col="chartNo"]`) ||
                 document.querySelector(`.excel-cell[data-row="${targetRow}"][data-col="name"]`);
    if (cell) {
      cell.scrollIntoView({ behavior: "smooth", block: "center" });
      this.selectCell(targetRow, "chartNo", cell);
    }
  }

  addNewRow(andFocus = true) {
    const rows = this.getCurrentRows();
    const formattedDate = this.currentDate.replace(/-/g, ".");
    const newRow = {
      no: "",
      gender: "",
      chartNo: "",
      name: "",
      part: "",
      prescription: "",
      extra: "",
      writer: DEFAULT_WRITER,
      memo: "",
      specialNote: "",
      date: formattedDate
    };
    rows.push(newRow);
    this.saveDataStore();
    this.renderTable();

    if (andFocus) {
      const newIdx = rows.length - 1;
      const targetCell = document.querySelector(`.excel-cell[data-row="${newIdx}"][data-col="chartNo"]`);
      if (targetCell) {
        targetCell.scrollIntoView({ behavior: "smooth", block: "center" });
        this.selectCell(newIdx, "chartNo", targetCell);
      }
    }
  }

  deleteRow(rowIdx) {
    const rows = this.getCurrentRows();
    if (rows.length <= 1) {
      // Just clear
      rows[0] = {
        no: "", gender: "", chartNo: "", name: "", part: "",
        prescription: "", extra: "", writer: DEFAULT_WRITER,
        memo: "", specialNote: "", date: this.currentDate.replace(/-/g, ".")
      };
    } else {
      rows.splice(rowIdx, 1);
    }
    this.saveDataStore();
    this.renderTable();
  }

  deleteSelectedRow() {
    if (this.selectedRowIdx === null) {
      alert("삭제할 행이나 셀을 먼저 선택해주세요.");
      return;
    }
    this.deleteRow(this.selectedRowIdx);
    this.selectedRowIdx = null;
    this.activeCell = null;
  }

  removeEmptyRows() {
    const rows = this.getCurrentRows();
    const filtered = rows.filter((r) => r.name || r.chartNo || r.part || r.prescription || r.extra);
    if (filtered.length === 0) {
      this.dataStore[this.currentDate] = this.createDefaultEmptyRows(5);
    } else {
      this.dataStore[this.currentDate] = filtered;
    }
    this.saveDataStore();
    this.renderTable();
    alert("빈 행이 깔끔하게 정리되었습니다.");
  }

  applyQuickChip(type, val) {
    const rows = this.getCurrentRows();
    let targetRow = this.selectedRowIdx !== null ? this.selectedRowIdx : 0;
    if (!rows[targetRow]) {
      this.addNewRow(false);
      targetRow = rows.length - 1;
    }

    if (type === "prescription") {
      rows[targetRow].prescription = val;
    } else if (type === "extra") {
      // Append or set extra
      const cur = rows[targetRow].extra || "";
      rows[targetRow].extra = cur ? `${cur}, ${val}` : val;
    }

    this.saveDataStore();
    this.renderTable();

    // Select the modified cell
    const colKey = type === "prescription" ? "prescription" : "extra";
    const cellEl = document.querySelector(`.excel-cell[data-row="${targetRow}"][data-col="${colKey}"]`);
    if (cellEl) this.selectCell(targetRow, colKey, cellEl);
  }

  // Update right sidebar statistics
  updateSidebarStats() {
    const rows = this.getCurrentRows().filter((r) => r.name || r.chartNo || r.part);
    const totalCount = rows.length;
    let maleCount = 0;
    let femaleCount = 0;
    let shockwaveCount = 0;
    let ionCount = 0;
    let winbackCount = 0;
    let extraOtherCount = 0;
    let prescExtremity = 0;
    let prescSpine = 0;
    let prescOther = 0;

    rows.forEach((r) => {
      const g = (r.gender || "").toUpperCase();
      if (g === "M") maleCount++;
      else if (g === "F") femaleCount++;

      const extra = (r.extra || "").trim();
      if (extra.includes("충격파")) shockwaveCount++;
      if (extra.includes("이온")) ionCount++;
      if (extra.includes("윈백")) winbackCount++;
      if (extra && !extra.includes("충격파") && !extra.includes("이온") && !extra.includes("윈백")) {
        extraOtherCount++;
      }

      const presc = (r.prescription || "").trim();
      if (presc.includes("사지")) prescExtremity++;
      else if (presc.includes("척추")) prescSpine++;
      else if (presc && presc !== "X") prescOther++;
    });

    this.elStatTotalCount.textContent = totalCount;
    this.elStatMaleCount.textContent = maleCount;
    this.elStatFemaleCount.textContent = femaleCount;

    this.elStatShockwave.textContent = shockwaveCount;
    this.elStatIon.textContent = ionCount;
    this.elStatWinback.textContent = winbackCount;
    this.elStatExtraOther.textContent = extraOtherCount;

    this.elStatPrescExtremity.textContent = prescExtremity;
    this.elStatPrescSpine.textContent = prescSpine;
    this.elStatPrescOther.textContent = prescOther;
  }

  renderRecentDays() {
    const dates = Object.keys(this.dataStore).sort().reverse();
    this.elRecentDaysList.innerHTML = "";

    dates.slice(0, 10).forEach((d) => {
      const badge = document.createElement("span");
      badge.className = "recent-day-badge" + (d === this.currentDate ? " active" : "");
      badge.textContent = d.replace(/-/g, ".");
      badge.addEventListener("click", () => this.setDate(d));
      this.elRecentDaysList.appendChild(badge);
    });
  }

  handleSearch() {
    const q = this.elSearchInput.value.trim().toLowerCase();
    this.elBtnClearSearch.style.display = q ? "block" : "none";

    const rows = document.querySelectorAll(".excel-row");
    rows.forEach((rowEl) => {
      if (!q) {
        rowEl.style.display = "";
        return;
      }
      const text = rowEl.textContent.toLowerCase();
      rowEl.style.display = text.includes(q) ? "" : "none";
    });
  }

  // --- Preview & Print Modal ---
  openPreviewModal() {
    const rows = this.getCurrentRows().filter((r) => r.name || r.chartNo || r.part || r.prescription || r.extra);

    const [y, m, d] = this.currentDate.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);
    const daysKor = ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"];
    const dayLabel = daysKor[dateObj.getDay()] || "";
    const dateTitle = `${y}년 ${m}월 ${d}일`;

    this.elPreviewModalDate.textContent = dateTitle;
    this.elPrintDateFull.textContent = `${dateTitle} (${dayLabel})`;

    let maleCount = 0;
    let femaleCount = 0;
    let shockwave = 0;
    let ion = 0;
    let winback = 0;

    this.elPrintTableBody.innerHTML = "";

    if (rows.length === 0) {
      const emptyTr = document.createElement("tr");
      emptyTr.innerHTML = `<td colspan="10" style="text-align:center; padding: 20px; color: #888;">해당 날짜에 등록된 물리치료 환자 데이터가 없습니다.</td>`;
      this.elPrintTableBody.appendChild(emptyTr);
    } else {
      rows.forEach((r, idx) => {
        const g = (r.gender || "").toUpperCase();
        if (g === "M") maleCount++;
        if (g === "F") femaleCount++;
        if ((r.extra || "").includes("충격파")) shockwave++;
        if ((r.extra || "").includes("이온")) ion++;
        if ((r.extra || "").includes("윈백")) winback++;

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td style="text-align:center;">${r.no || idx + 1}</td>
          <td style="text-align:center; font-weight:bold;">${r.gender || ""}</td>
          <td style="text-align:right; font-family:monospace;">${r.chartNo || ""}</td>
          <td style="text-align:center; font-weight:bold;">${r.name || ""}</td>
          <td>${r.part || ""}</td>
          <td>${r.prescription || ""}</td>
          <td style="text-align:center; font-weight:bold;">${r.extra || ""}</td>
          <td style="text-align:center;">${r.writer || DEFAULT_WRITER}</td>
          <td style="text-align:center;">${r.memo || ""}</td>
          <td style="color:#b30000;">${r.specialNote || ""}</td>
        `;
        this.elPrintTableBody.appendChild(tr);
      });
    }

    this.elPrintTotalCount.textContent = rows.length;
    this.elPrintMaleCount.textContent = maleCount;
    this.elPrintFemaleCount.textContent = femaleCount;
    this.elPrintShockwave.textContent = shockwave;
    this.elPrintIon.textContent = ion;
    this.elPrintWinback.textContent = winback;

    const now = new Date();
    this.elPrintGeneratedTime.textContent = now.toLocaleString("ko-KR");

    this.elPreviewModal.style.display = "flex";
  }

  closePreviewModal() {
    this.elPreviewModal.style.display = "none";
  }

  // --- CSV Export ---
  exportCurrentDayCsv() {
    const rows = this.getCurrentRows().filter((r) => r.name || r.chartNo || r.part || r.prescription || r.extra);
    if (rows.length === 0) {
      alert("내보낼 데이터가 없습니다.");
      return;
    }

    const headers = ["No.", "G", "차트No.", "성함", "부위", "처방", "추가 사항", "작성", "메모", "특이 사항", "날짜"];
    let csvContent = "\uFEFF"; // UTF-8 BOM for Excel
    csvContent += headers.map((h) => `"${h}"`).join(",") + "\n";

    rows.forEach((r, idx) => {
      const line = [
        r.no || idx + 1,
        r.gender || "",
        r.chartNo || "",
        r.name || "",
        r.part || "",
        r.prescription || "",
        r.extra || "",
        r.writer || DEFAULT_WRITER,
        r.memo || "",
        r.specialNote || "",
        r.date || this.currentDate.replace(/-/g, ".")
      ];
      csvContent += line.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",") + "\n";
    });

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `물리치료현황_${this.currentDate}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  // --- Backup & Restore Modal ---
  openBackupModal() {
    this.elBackupModal.style.display = "flex";
  }

  closeBackupModal() {
    this.elBackupModal.style.display = "none";
  }

  downloadFullBackup() {
    const dataStr = JSON.stringify(this.dataStore, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `PT현황_전체백업_${this.getTodayString()}.json`;
    link.click();
    URL.revokeObjectURL(url);
  }

  handleRestoreFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (typeof parsed === "object" && parsed !== null) {
          this.dataStore = parsed;
          this.saveDataStore();
          this.setDate(this.currentDate);
          alert("백업 파일이 성공적으로 복원되었습니다.");
          this.closeBackupModal();
        } else {
          alert("올바르지 않은 백업 파일 형식입니다.");
        }
      } catch (err) {
        alert("JSON 파일을 읽는 중 오류가 발생했습니다.");
      }
    };
    reader.readAsText(file);
  }

  clearCurrentDayData() {
    if (confirm(`정말로 ${this.currentDate} 날짜의 데이터를 모두 초기화하시겠습니까?`)) {
      this.dataStore[this.currentDate] = this.createDefaultEmptyRows(10);
      this.saveDataStore();
      this.renderTable();
      this.closeBackupModal();
    }
  }

  // Global Keyboard Shortcuts
  handleGlobalKeyDown(e) {
    if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") {
      return;
    }
    // Ctrl+P or Cmd+P -> Preview modal
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "p") {
      e.preventDefault();
      this.openPreviewModal();
      return;
    }
    // Ctrl+S or Cmd+S -> Quick Save
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      this.saveDataStore();
      return;
    }
    // Alt + Left / Right -> Date switch
    if (e.altKey && e.key === "ArrowLeft") {
      e.preventDefault();
      this.shiftDay(-1);
    } else if (e.altKey && e.key === "ArrowRight") {
      e.preventDefault();
      this.shiftDay(1);
    }
  }

  initPWA() {
    if ("serviceWorker" in navigator) {
      window.addEventListener("load", () => {
        navigator.serviceWorker.register("./sw.js")
          .then((reg) => console.log("PT Status PWA Service Worker Registered", reg.scope))
          .catch((err) => console.log("SW Registration Failed", err));
      });
    }
  }
}

// Instantiate on DOMContentLoaded
window.addEventListener("DOMContentLoaded", () => {
  window.ptApp = new PTApp();
});
