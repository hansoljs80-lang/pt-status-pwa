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
const SUPABASE_CONFIG_KEY = "PT_SUPABASE_CONFIG_V1";
const DEFAULT_WRITER = "S";
const BASE_ROW_NUMBER = 1; // 행 번호 1부터 시작
const DEFAULT_ROW_COUNT = 150; // 기본 하루 150개 행

class PTApp {
  constructor() {
    this.dataStore = this.loadDataStore();
    this.currentDate = this.getTodayString();
    this.activeCell = null; // { rowIdx, colKey }
    this.selectedRowIdx = null;
    this.selectedColKey = null;
    this.sortState = { colKey: null, direction: "asc" };
    this.supabaseClient = null;
    this.supabaseSyncTimer = null;

    this.cacheElements();
    this.bindEvents();
    this.initColumnResizing();
    this.initSupabase();
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
      this.scheduleSupabaseSync();
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

    // Supabase Modal & Cloud Elements
    this.elBtnSupabase = document.getElementById("btnSupabase");
    this.elSupabaseStatusLabel = document.getElementById("supabaseStatusLabel");
    this.elSupabaseModal = document.getElementById("supabaseModal");
    this.elSupabaseModalStatus = document.getElementById("supabaseModalStatus");
    this.elBtnCloseSupabase = document.getElementById("btnCloseSupabase");
    this.elSbUrlInput = document.getElementById("sbUrlInput");
    this.elSbKeyInput = document.getElementById("sbKeyInput");
    this.elBtnSaveSupabase = document.getElementById("btnSaveSupabase");
    this.elBtnDisconnectSupabase = document.getElementById("btnDisconnectSupabase");
    this.elSupabaseManualSyncBox = document.getElementById("supabaseManualSyncBox");
    this.elBtnPushToCloud = document.getElementById("btnPushToCloud");
    this.elBtnPullFromCloud = document.getElementById("btnPullFromCloud");
  }

  bindEvents() {
    // Supabase Cloud Sync Modal
    if (this.elBtnSupabase) {
      this.elBtnSupabase.addEventListener("click", () => this.openSupabaseModal());
    }
    if (this.elBtnCloseSupabase) {
      this.elBtnCloseSupabase.addEventListener("click", () => this.closeSupabaseModal());
    }
    if (this.elBtnSaveSupabase) {
      this.elBtnSaveSupabase.addEventListener("click", () => this.saveSupabaseConfig());
    }
    if (this.elBtnDisconnectSupabase) {
      this.elBtnDisconnectSupabase.addEventListener("click", () => this.disconnectSupabase());
    }
    if (this.elBtnPushToCloud) {
      this.elBtnPushToCloud.addEventListener("click", () => this.pushToCloud(this.currentDate, true));
    }
    if (this.elBtnPullFromCloud) {
      this.elBtnPullFromCloud.addEventListener("click", () => this.pullFromCloud(this.currentDate, true));
    }

    // Column Headers Click (Select Entire Column)
    document.querySelectorAll(".col-headers-row th.col-letter").forEach((th) => {
      th.addEventListener("click", (e) => {
        if (e.target.classList.contains("col-resizer")) return;
        const colKey = th.dataset.col;
        const colLetter = th.dataset.colLetter;
        if (colKey) this.selectEntireColumn(colKey, colLetter);
      });
    });

    // Business Headers Click (Sort Column)
    document.querySelectorAll(".business-headers-row th.b-header").forEach((th) => {
      th.addEventListener("click", () => {
        const colKey = th.dataset.col;
        if (colKey && colKey !== "del") this.sortByColumn(colKey);
      });
    });

    // Corner Header Click (Select All Sheet)
    const cornerHeader = document.getElementById("cornerHeader");
    if (cornerHeader) {
      cornerHeader.addEventListener("click", () => this.selectAllCells());
    }

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
        this.closeSupabaseModal();
      }
    });
  }

  // Get or initialize rows for a given date (guarantee minimum DEFAULT_ROW_COUNT = 150 rows)
  getCurrentRows() {
    if (!this.dataStore[this.currentDate]) {
      this.dataStore[this.currentDate] = this.createDefaultEmptyRows(DEFAULT_ROW_COUNT);
    } else {
      const rows = this.dataStore[this.currentDate];
      if (rows.length < DEFAULT_ROW_COUNT) {
        const formattedDate = this.currentDate.replace(/-/g, ".");
        const diff = DEFAULT_ROW_COUNT - rows.length;
        for (let i = 0; i < diff; i++) {
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
      }
    }
    return this.dataStore[this.currentDate];
  }

  createDefaultEmptyRows(count = DEFAULT_ROW_COUNT) {
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

    // Pull from Supabase cloud if connected
    if (this.supabaseClient) {
      this.pullFromCloud(dateStr, false);
    }

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
      thNum.title = `행 ${excelRowNum} 클릭하여 행 전체 선택`;
      thNum.addEventListener("click", () => {
        this.selectEntireRow(rowIdx);
      });
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
        this.updateActiveHeaders(this.activeCell.rowIdx, this.activeCell.colKey);
      }
    }
  }

  // Select and focus cell like Excel
  selectCell(rowIdx, colKey, cellElement) {
    this.activeCell = { rowIdx, colKey };
    this.selectedRowIdx = rowIdx;

    this.clearHeaderSelections();
    this.updateActiveHeaders(rowIdx, colKey);

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

  // Update active highlighted headers matching currently focused cell
  updateActiveHeaders(rowIdx, colKey) {
    document.querySelectorAll(".col-letter, .b-header, .row-num").forEach((el) => {
      el.classList.remove("header-active");
    });

    const activeColLetter = document.querySelector(`.col-letter[data-col="${colKey}"]`);
    const activeBHeader = document.querySelector(`.b-header[data-col="${colKey}"]`);
    if (activeColLetter) activeColLetter.classList.add("header-active");
    if (activeBHeader) activeBHeader.classList.add("header-active");

    const rowTr = document.querySelector(`tr[data-row-idx="${rowIdx}"]`);
    if (rowTr) {
      const activeRowNum = rowTr.querySelector(".row-num");
      if (activeRowNum) activeRowNum.classList.add("header-active");
    }
  }

  clearHeaderSelections() {
    document.querySelectorAll(".col-letter, .b-header, .row-num, .corner-header").forEach((el) => {
      el.classList.remove("selected");
    });
    document.querySelectorAll(".excel-cell").forEach((el) => {
      el.classList.remove("col-selected", "row-selected", "all-selected");
    });
  }

  selectEntireColumn(colKey, colLetter) {
    this.clearHeaderSelections();
    this.selectedColKey = colKey;
    this.selectedRowIdx = null;

    // Highlight Column Header
    const colTh = document.querySelector(`.col-letter[data-col="${colKey}"]`);
    const bTh = document.querySelector(`.b-header[data-col="${colKey}"]`);
    if (colTh) colTh.classList.add("selected");
    if (bTh) bTh.classList.add("selected");

    // Select all cells in this column
    document.querySelectorAll(`.excel-cell[data-col="${colKey}"]`).forEach((cell) => {
      cell.classList.add("col-selected");
    });

    this.elCellAddress.textContent = `${colLetter}:${colLetter}`;
    this.elSelectedCellCoords.textContent = `${colLetter}열 전체 선택 (${colKey})`;
    this.elFormulaInput.value = "";
  }

  selectEntireRow(rowIdx) {
    this.clearHeaderSelections();
    this.selectedRowIdx = rowIdx;
    this.selectedColKey = null;

    const rowTr = document.querySelector(`tr[data-row-idx="${rowIdx}"]`);
    if (rowTr) {
      const rowNumTh = rowTr.querySelector(".row-num");
      if (rowNumTh) rowNumTh.classList.add("selected");
      rowTr.querySelectorAll(".excel-cell").forEach((cell) => {
        cell.classList.add("row-selected");
      });
    }

    const excelRowNum = BASE_ROW_NUMBER + rowIdx;
    this.elCellAddress.textContent = `${excelRowNum}:${excelRowNum}`;
    this.elSelectedCellCoords.textContent = `${excelRowNum}행 전체 선택`;
    this.elFormulaInput.value = "";
  }

  selectAllCells() {
    this.clearHeaderSelections();
    const cornerHeader = document.getElementById("cornerHeader");
    if (cornerHeader) cornerHeader.classList.add("selected");

    document.querySelectorAll(".excel-cell").forEach((cell) => {
      cell.classList.add("all-selected");
    });

    this.elCellAddress.textContent = "1:전체";
    this.elSelectedCellCoords.textContent = "전체 시트 선택";
    this.elFormulaInput.value = "";
  }

  sortByColumn(colKey) {
    const rows = this.getCurrentRows();
    if (this.sortState.colKey === colKey) {
      this.sortState.direction = this.sortState.direction === "asc" ? "desc" : "asc";
    } else {
      this.sortState.colKey = colKey;
      this.sortState.direction = "asc";
    }

    const dir = this.sortState.direction;

    // Update sort indicators in header
    document.querySelectorAll(".b-header .sort-indicator").forEach((ind) => {
      ind.textContent = "";
    });
    const currentTh = document.querySelector(`.b-header[data-col="${colKey}"] .sort-indicator`);
    if (currentTh) {
      currentTh.textContent = dir === "asc" ? " ▲" : " ▼";
    }

    // Separate rows with data and blank rows so blank rows always stay at bottom
    const dataRows = rows.filter((r) => r.name || r.chartNo || r.part || r.no);
    const emptyRows = rows.filter((r) => !r.name && !r.chartNo && !r.part && !r.no);

    dataRows.sort((a, b) => {
      let valA = (a[colKey] || "").toString().trim();
      let valB = (b[colKey] || "").toString().trim();

      // Check if numeric
      const numA = parseFloat(valA);
      const numB = parseFloat(valB);
      if (!isNaN(numA) && !isNaN(numB) && String(numA) === valA && String(numB) === valB) {
        return dir === "asc" ? numA - numB : numB - numA;
      }

      return dir === "asc" ? valA.localeCompare(valB, "ko") : valB.localeCompare(valA, "ko");
    });

    this.dataStore[this.currentDate] = [...dataRows, ...emptyRows];
    this.saveDataStore();
    this.renderTable();
  }

  initColumnResizing() {
    let activeTh = null;
    let startX = 0;
    let startWidth = 0;

    document.addEventListener("mousedown", (e) => {
      if (e.target.classList.contains("col-resizer")) {
        e.preventDefault();
        e.stopPropagation();
        const resizer = e.target;
        activeTh = resizer.closest("th");
        startX = e.pageX;
        startWidth = activeTh.offsetWidth;
        resizer.classList.add("resizing");
        document.body.style.cursor = "col-resize";
        document.body.style.userSelect = "none";
      }
    });

    document.addEventListener("mousemove", (e) => {
      if (!activeTh) return;
      const diff = e.pageX - startX;
      const newWidth = Math.max(35, startWidth + diff);
      activeTh.style.width = `${newWidth}px`;
    });

    document.addEventListener("mouseup", () => {
      if (activeTh) {
        const resizer = activeTh.querySelector(".col-resizer");
        if (resizer) resizer.classList.remove("resizing");
        activeTh = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
      }
    });
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
    const formattedDate = this.currentDate.replace(/-/g, ".");
    while (filtered.length < DEFAULT_ROW_COUNT) {
      filtered.push({
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
    this.dataStore[this.currentDate] = filtered;
    this.saveDataStore();
    this.renderTable();
    alert("데이터가 정리되었으며 기본 150행이 유지됩니다.");
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
      this.dataStore[this.currentDate] = this.createDefaultEmptyRows(DEFAULT_ROW_COUNT);
      this.saveDataStore();
      this.renderTable();
      this.closeBackupModal();
    }
  }

  // --- Supabase Cloud Sync Methods ---
  initSupabase() {
    try {
      const saved = localStorage.getItem(SUPABASE_CONFIG_KEY);
      if (saved) {
        const { url, key } = JSON.parse(saved);
        if (url && key && window.supabase) {
          this.supabaseClient = window.supabase.createClient(url, key);
          this.updateSupabaseUI(true);
          // Pull latest cloud data for today
          this.pullFromCloud(this.currentDate, false);
          return;
        }
      }
    } catch (e) {
      console.error("Supabase init error:", e);
    }
    this.updateSupabaseUI(false);
  }

  updateSupabaseUI(connected) {
    if (this.elSupabaseStatusLabel) {
      this.elSupabaseStatusLabel.textContent = connected ? "☁️ 클라우드 연결됨" : "슈파베이스 연동";
    }
    if (this.elBtnSupabase) {
      this.elBtnSupabase.classList.toggle("connected", connected);
    }
    if (this.elSupabaseModalStatus) {
      this.elSupabaseModalStatus.textContent = connected ? "연결됨 (실시간 동기화 중)" : "미연결";
      this.elSupabaseModalStatus.style.background = connected ? "#e2efda" : "#f1f3f5";
      this.elSupabaseModalStatus.style.color = connected ? "#274e13" : "#666";
    }
    if (this.elBtnDisconnectSupabase) {
      this.elBtnDisconnectSupabase.style.display = connected ? "inline-block" : "none";
    }
    if (this.elSupabaseManualSyncBox) {
      this.elSupabaseManualSyncBox.style.display = connected ? "block" : "none";
    }
  }

  openSupabaseModal() {
    try {
      const saved = localStorage.getItem(SUPABASE_CONFIG_KEY);
      if (saved) {
        const { url, key } = JSON.parse(saved);
        if (this.elSbUrlInput) this.elSbUrlInput.value = url || "";
        if (this.elSbKeyInput) this.elSbKeyInput.value = key || "";
      }
    } catch (e) {}
    if (this.elSupabaseModal) this.elSupabaseModal.style.display = "flex";
  }

  closeSupabaseModal() {
    if (this.elSupabaseModal) this.elSupabaseModal.style.display = "none";
  }

  async saveSupabaseConfig() {
    const url = (this.elSbUrlInput.value || "").trim();
    const key = (this.elSbKeyInput.value || "").trim();

    if (!url || !key) {
      alert("Supabase Project URL과 Anon Key를 모두 입력해주세요.");
      return;
    }

    if (!window.supabase) {
      alert("Supabase SDK를 로드할 수 없습니다. 인터넷 연결을 확인해주세요.");
      return;
    }

    try {
      const testClient = window.supabase.createClient(url, key);
      // Test query to check table
      const { error } = await testClient.from("pt_daily_records").select("date").limit(1);
      if (error && (error.code === "PGRST204" || (error.message && error.message.includes("relation") && error.message.includes("does not exist")))) {
        alert("Supabase 연결은 확인되었으나, 'pt_daily_records' 테이블이 없습니다.\n하단의 SQL 스크립트를 Supabase SQL Editor에서 실행해주세요!");
      }

      localStorage.setItem(SUPABASE_CONFIG_KEY, JSON.stringify({ url, key }));
      this.supabaseClient = testClient;
      this.updateSupabaseUI(true);
      alert("Supabase 클라우드 동기화가 성공적으로 활성화되었습니다!\n지금부터 모든 기록이 자동 동기화됩니다.");
      this.closeSupabaseModal();

      // Push current local data to cloud
      this.pushToCloud(this.currentDate, false);
    } catch (err) {
      console.error("Supabase connect error:", err);
      alert("연결 중 오류가 발생했습니다: " + (err.message || err));
    }
  }

  disconnectSupabase() {
    if (confirm("Supabase 클라우드 연결을 해제하시겠습니까? (로컬 데이터는 안전하게 유지됩니다)")) {
      localStorage.removeItem(SUPABASE_CONFIG_KEY);
      this.supabaseClient = null;
      this.updateSupabaseUI(false);
      if (this.elSbUrlInput) this.elSbUrlInput.value = "";
      if (this.elSbKeyInput) this.elSbKeyInput.value = "";
      this.closeSupabaseModal();
      alert("연결이 해제되었습니다.");
    }
  }

  scheduleSupabaseSync() {
    if (!this.supabaseClient) return;
    if (this.supabaseSyncTimer) clearTimeout(this.supabaseSyncTimer);
    this.supabaseSyncTimer = setTimeout(() => {
      this.pushToCloud(this.currentDate, false);
    }, 1200);
  }

  async pushToCloud(dateStr, showNotice = false) {
    if (!this.supabaseClient) {
      if (showNotice) alert("Supabase가 연결되어 있지 않습니다. 상단 [슈파베이스 연동]을 눌러 설정해주세요.");
      return;
    }

    try {
      const rows = this.dataStore[dateStr] || this.getCurrentRows();
      const meaningfulCount = rows.filter((r) => r.name || r.chartNo).length;

      const { error } = await this.supabaseClient.from("pt_daily_records").upsert({
        date: dateStr,
        rows_data: rows,
        total_count: meaningfulCount,
        updated_at: new Date().toISOString()
      });

      if (error) {
        console.warn("Cloud push warning:", error);
        if (showNotice) alert("클라우드 업로드 실패: " + error.message);
      } else {
        this.showSaveIndicator("클라우드 동기화 완료");
        if (showNotice) alert(`${dateStr} 데이터가 Supabase 클라우드에 성공적으로 저장되었습니다!`);
      }
    } catch (err) {
      console.error("Cloud push exception:", err);
      if (showNotice) alert("클라우드 통신 오류: " + err.message);
    }
  }

  async pullFromCloud(dateStr, showNotice = false) {
    if (!this.supabaseClient) {
      if (showNotice) alert("Supabase가 연결되어 있지 않습니다.");
      return;
    }

    try {
      const { data, error } = await this.supabaseClient
        .from("pt_daily_records")
        .select("rows_data, updated_at")
        .eq("date", dateStr)
        .maybeSingle();

      if (error) {
        console.warn("Cloud pull warning:", error);
        if (showNotice) alert("클라우드 데이터 가져오기 실패: " + error.message);
        return;
      }

      if (data && Array.isArray(data.rows_data) && data.rows_data.length > 0) {
        this.dataStore[dateStr] = data.rows_data;
        // Ensure minimum 150 rows
        this.getCurrentRows();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataStore));
        if (this.currentDate === dateStr) {
          this.renderTable();
          this.updateSidebarStats();
          this.renderRecentDays();
        }
        this.showSaveIndicator("클라우드 데이터 수신됨");
        if (showNotice) alert(`${dateStr} 클라우드 최신 데이터를 성공적으로 불러왔습니다!`);
      } else {
        if (showNotice) alert(`${dateStr} 일자의 클라우드 데이터가 아직 없습니다.`);
      }
    } catch (err) {
      console.error("Cloud pull exception:", err);
      if (showNotice) alert("클라우드 통신 오류: " + err.message);
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
      return;
    } else if (e.altKey && e.key === "ArrowRight") {
      e.preventDefault();
      this.shiftDay(1);
      return;
    }

    // Delete or Backspace when row/col is selected
    if (e.key === "Delete" || e.key === "Backspace") {
      if (this.selectedRowIdx !== null) {
        e.preventDefault();
        this.deleteRow(this.selectedRowIdx);
        this.selectedRowIdx = null;
        return;
      }
      if (this.selectedColKey !== null) {
        e.preventDefault();
        const rows = this.getCurrentRows();
        rows.forEach((r) => { r[this.selectedColKey] = ""; });
        this.saveDataStore();
        this.renderTable();
        return;
      }
    }

    // Escape -> Clear selections
    if (e.key === "Escape") {
      this.clearHeaderSelections();
      document.querySelectorAll(".cell-focused").forEach((c) => c.classList.remove("cell-focused"));
      this.activeCell = null;
      this.selectedRowIdx = null;
      this.selectedColKey = null;
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
