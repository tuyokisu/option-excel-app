(function () {
  "use strict";

  var api = window.OptionExcel;
  var rows = [];
  var tbody = document.getElementById("optionRows");
  var optionTableWrap = document.getElementById("optionTableWrap");
  var specNameInput = document.getElementById("specName");
  var basePriceInput = document.getElementById("basePrice");
  var inputRowsSelect = document.getElementById("inputRows");
  var errorBox = document.getElementById("errors");
  var statusBox = document.getElementById("status");
  var optionCount = document.getElementById("optionCount");
  var selectedNosInput = document.getElementById("selectedNos");
  var previewBase = document.getElementById("previewBase");
  var previewOptionTotal = document.getElementById("previewOptionTotal");
  var previewGrandTotal = document.getElementById("previewGrandTotal");
  var previewSpecStatus = document.getElementById("previewSpecStatus");
  var previewSelectedNos = document.getElementById("previewSelectedNos");
  var previewSelectedNames = document.getElementById("previewSelectedNames");
  var previewEmpty = document.getElementById("previewEmpty");
  var previewRows = document.getElementById("previewRows");
  var previewTableWrap = document.getElementById("previewTableWrap");
  var doneDialog = document.getElementById("doneDialog");
  var doneDialogMessage = document.getElementById("doneDialogMessage");
  var doneDialogOk = document.getElementById("doneDialogOk");
  var linkSaveFolderButton = document.getElementById("linkSaveFolder");
  var saveFolderStatus = document.getElementById("saveFolderStatus");
  var savedFilePicker = document.getElementById("savedFilePicker");
  var savedFileToggle = document.getElementById("savedFileToggle");
  var savedFileSearch = document.getElementById("savedFileSearch");
  var savedFileList = document.getElementById("savedFileList");
  var loadSavedFileButton = document.getElementById("loadSavedFile");
  var savedFileNames = [];
  var selectedSavedFile = "";
  var excelFilePicker = document.getElementById("excelFilePicker");
  var excelFileToggle = document.getElementById("excelFileToggle");
  var excelFileSearch = document.getElementById("excelFileSearch");
  var excelFileList = document.getElementById("excelFileList");
  var loadExcelFileButton = document.getElementById("loadExcelFile");
  var excelFileNames = [];
  var selectedExcelFile = "";
  var exportRowInput = document.getElementById("exportRow");
  var exportTargetStatus = document.getElementById("exportTargetStatus");
  var exportSelectedFile = document.getElementById("exportSelectedFile");
  var exportSelectedName = document.getElementById("exportSelectedName");
  var writeCalcRowButton = document.getElementById("writeCalcRow");
  var chooseExportExcelButton = document.getElementById("chooseExportExcel");
  var exportFileHandle = null;
  var SAVE_DIR_DB = "option-excel-app";
  var SAVE_DIR_STORE = "handles";
  var APP_DIR_KEY = "app-root";
  var EXPORT_FILE_KEY = "export-xlsx";
  var appDirHandle = null;
  var saveDirHandle = null;
  var excelDirHandle = null;
  var saveApi = false;
  var saveApiBase = "";
  var lastPreview = null;
  var copyFeedbackTimer = 0;
  var popupCanCloseAt = 0;

  function nextOptionNo(list) {
    var max = 0;
    (list || rows).forEach(function (row) {
      var n = Number(String(row && row.no != null ? row.no : "").trim());
      if (Number.isInteger(n) && n > max) {
        max = n;
      }
    });
    return String(max + 1);
  }

  function emptyRow(list) {
    return { no: nextOptionNo(list), name: "", price: "" };
  }

  function starterRows() {
    var list = [];
    var i;
    for (i = 1; i <= 5; i += 1) {
      list.push({ no: String(i), name: "", price: "" });
    }
    return list;
  }

  function isFilled(row) {
    return row.name.trim() || String(row.price).trim();
  }

  function setStatus(message) {
    statusBox.textContent = message || "";
    statusBox.hidden = !message;
  }

  function formatMoney(value) {
    if (value == null || !Number.isFinite(Number(value))) {
      return "—";
    }
    return Number(value).toLocaleString("ja-JP");
  }

  function showDonePopup(message) {
    doneDialogMessage.textContent = message;
    popupCanCloseAt = Date.now() + 1200;
    doneDialog.hidden = false;
    doneDialog.classList.add("is-open");
    function swallow(event) {
      if (Date.now() < popupCanCloseAt) {
        event.preventDefault();
        event.stopPropagation();
      }
    }
    window.addEventListener("click", swallow, true);
    window.addEventListener("pointerup", swallow, true);
    window.setTimeout(function () {
      window.removeEventListener("click", swallow, true);
      window.removeEventListener("pointerup", swallow, true);
    }, 1200);
  }

  function hideDonePopup(event) {
    if (Date.now() < popupCanCloseAt) {
      if (event) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }
    doneDialog.classList.remove("is-open");
    doneDialog.hidden = true;
  }

  function copyText(text) {
    var value = String(text == null ? "" : text);
    if (!value) {
      return Promise.reject(new Error("コピーする値がありません。"));
    }
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function" && window.isSecureContext) {
      return navigator.clipboard.writeText(value);
    }
    return new Promise(function (resolve, reject) {
      var area = document.createElement("textarea");
      area.value = value;
      area.setAttribute("readonly", "");
      area.style.position = "fixed";
      area.style.left = "-9999px";
      document.body.appendChild(area);
      area.select();
      var ok = false;
      try {
        ok = document.execCommand("copy");
      } catch (error) {
        ok = false;
      }
      area.remove();
      if (ok) {
        resolve();
      } else {
        reject(new Error("コピーできませんでした。"));
      }
    });
  }

  function hideCopyFeedback() {
    document.querySelectorAll(".copy-feedback").forEach(function (el) {
      el.hidden = true;
    });
  }

  function showCopyFeedback(button) {
    hideCopyFeedback();
    var note = button.parentElement && button.parentElement.querySelector(".copy-feedback");
    if (!note) {
      note = document.createElement("p");
      note.className = "copy-feedback";
      note.textContent = "コピーしました";
      button.insertAdjacentElement("afterend", note);
    }
    note.hidden = false;
    window.clearTimeout(copyFeedbackTimer);
    copyFeedbackTimer = window.setTimeout(function () {
      note.hidden = true;
    }, 1800);
  }

  function copyValue(text, button) {
    if (!text) {
      showErrors(["コピーする値がありません。"]);
      return;
    }
    copyText(text).then(function () {
      showErrors([]);
      showCopyFeedback(button);
    }).catch(function () {
      showErrors(["コピーできませんでした。"]);
    });
  }

  function previewCopyText(key, item) {
    if (key === "base") {
      return lastPreview ? String(lastPreview.basePrice) : "";
    }
    if (key === "optionTotal") {
      return lastPreview ? String(lastPreview.optionTotal) : "";
    }
    if (key === "grandTotal") {
      return lastPreview ? String(lastPreview.grandTotal) : "";
    }
    if (key === "selectedNos") {
      return lastPreview && lastPreview.selectedNosText ? lastPreview.selectedNosText : "";
    }
    if (key === "selectedNames") {
      return lastPreview && lastPreview.selectedWithNames ? lastPreview.selectedWithNames : "";
    }
    if (key === "row" && item) {
      if (item.missing) {
        return item.no + " " + item.name;
      }
      return item.no + " " + item.name + " " + item.price;
    }
    return "";
  }

  function updatePreview() {
    if (!api || typeof api.previewOptionSelection !== "function") {
      return;
    }
    var preview = api.previewOptionSelection({
      basePrice: basePriceInput.value,
      inputRows: Number(inputRowsSelect.value),
      selectedNos: selectedNosInput.value,
      options: rows,
    });
    var specLabel = specNameInput.value.trim() || "（仕様名なし）";
    previewSpecStatus.textContent =
      specLabel + " のオプション一覧 " + preview.masterCount + "件で計算";
    lastPreview = preview;
    previewBase.textContent = formatMoney(preview.basePrice);
    previewOptionTotal.textContent = formatMoney(preview.optionTotal);
    previewGrandTotal.textContent = formatMoney(preview.grandTotal);
    previewSelectedNos.textContent = preview.selectedNosText || "（まだ選んでいません）";
    previewSelectedNames.textContent = preview.selectedWithNames || "（まだ選んでいません）";
    previewRows.replaceChildren();
    preview.items.forEach(function (item) {
      var tr = document.createElement("tr");
      var noCell = document.createElement("td");
      noCell.textContent = item.no;
      var nameCell = document.createElement("td");
      nameCell.textContent = item.name;
      var priceCell = document.createElement("td");
      priceCell.textContent = item.missing ? "—" : formatMoney(item.price);
      var actionCell = document.createElement("td");
      var copyWrap = document.createElement("div");
      copyWrap.className = "copy-row";
      var copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "ghost copy-btn";
      copyBtn.textContent = "コピー";
      copyBtn.addEventListener("click", function () {
        copyValue(previewCopyText("row", item), copyBtn);
      });
      var copied = document.createElement("p");
      copied.className = "copy-feedback";
      copied.textContent = "コピーしました";
      copied.hidden = true;
      copyWrap.appendChild(copyBtn);
      copyWrap.appendChild(copied);
      actionCell.appendChild(copyWrap);
      tr.appendChild(noCell);
      tr.appendChild(nameCell);
      tr.appendChild(priceCell);
      tr.appendChild(actionCell);
      previewRows.appendChild(tr);
    });
    previewTableWrap.classList.toggle("compact", preview.items.length >= 11);
    if (preview.masterCount === 0) {
      previewEmpty.textContent = "この仕様書にはまだオプション一覧がありません。JSONを読み込むか、上の一覧を入力してください。";
      previewEmpty.hidden = false;
    } else if (preview.items.length === 0) {
      previewEmpty.textContent = "選ぶオプションNoを入力すると、Noと品名が表示されます。";
      previewEmpty.hidden = false;
    } else {
      previewEmpty.hidden = true;
    }
  }

  function showErrors(messages) {
    errorBox.replaceChildren();
    if (!messages || messages.length === 0) {
      errorBox.hidden = true;
      return;
    }
    var list = document.createElement("ul");
    messages.forEach(function (message) {
      var item = document.createElement("li");
      item.textContent = message;
      list.appendChild(item);
    });
    errorBox.appendChild(list);
    errorBox.hidden = false;
    if (typeof errorBox.scrollIntoView === "function") {
      errorBox.scrollIntoView({ block: "nearest" });
    }
  }

  function updateCount() {
    var count = rows.filter(isFilled).length;
    optionCount.textContent = "入力中: " + count + "件（空行はExcelに入りません）";
  }

  function addField(tr, field, label, value, extra) {
    var cell = document.createElement("td");
    var input = document.createElement("input");
    input.type = "text";
    input.value = value;
    input.setAttribute("aria-label", label);
    input.autocomplete = "off";
    input.spellcheck = false;
    if (extra) {
      extra(input);
    }
    input.addEventListener("input", function () {
      field.row[field.key] = input.value;
      setStatus("");
      updateCount();
      updatePreview();
    });
    cell.appendChild(input);
    tr.appendChild(cell);
    return input;
  }

  function render(focusIndex) {
    tbody.replaceChildren();
    rows.forEach(function (row, index) {
      var tr = document.createElement("tr");

      addField(tr, { row: row, key: "no" }, (index + 1) + " オプションNo", row.no, function (input) {
        input.className = "option-no";
      });
      var nameInput = addField(tr, { row: row, key: "name" }, (index + 1) + " 品名", row.name);
      addField(tr, { row: row, key: "price" }, (index + 1) + " 値段", row.price, function (input) {
        input.inputMode = "decimal";
        input.className = "price";
      });

      var actionCell = document.createElement("td");
      var removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.className = "ghost";
      removeButton.textContent = "削除";
      removeButton.addEventListener("click", function () {
        rows.splice(index, 1);
        if (rows.length === 0) {
          rows.push(emptyRow());
        }
        setStatus("");
        render();
        updatePreview();
      });
      actionCell.appendChild(removeButton);
      tr.appendChild(actionCell);
      tbody.appendChild(tr);

      if (focusIndex === index) {
        nameInput.focus();
      }
    });
    optionTableWrap.classList.toggle("compact", rows.length >= 11);
    updateCount();
  }

  function currentPayload() {
    return {
      specName: specNameInput.value,
      basePrice: basePriceInput.value,
      inputRows: Number(inputRowsSelect.value),
      options: rows.map(function (row) {
        return {
          no: isFilled(row) ? String(row.no == null ? "" : row.no).trim() : "",
          name: row.name,
          price: row.price,
        };
      }),
    };
  }

  function downloadBlob(blob, filename) {
    var link = document.createElement("a");
    var url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function apiUrl(path) {
    return saveApiBase + path;
  }

  function jsonFilename() {
    var excelName = api.sanitizeFilename(specNameInput.value.trim() || "オプション");
    return excelName.replace(/\.xlsx$/i, ".json");
  }

  function supportsSaveFolder() {
    return typeof window.showDirectoryPicker === "function" && typeof indexedDB !== "undefined";
  }

  function openSaveDirDb() {
    return new Promise(function (resolve, reject) {
      var request = indexedDB.open(SAVE_DIR_DB, 1);
      request.onupgradeneeded = function () {
        request.result.createObjectStore(SAVE_DIR_STORE);
      };
      request.onsuccess = function () {
        resolve(request.result);
      };
      request.onerror = function () {
        reject(request.error || new Error("IndexedDBを開けませんでした。"));
      };
    });
  }

  function storeHandle(key, handle) {
    return openSaveDirDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(SAVE_DIR_STORE, "readwrite");
        tx.objectStore(SAVE_DIR_STORE).put(handle, key);
        tx.oncomplete = function () {
          resolve();
        };
        tx.onerror = function () {
          reject(tx.error || new Error("保存フォルダ情報を記録できませんでした。"));
        };
      });
    });
  }

  function readStoredHandle(key) {
    return openSaveDirDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(SAVE_DIR_STORE, "readonly");
        var request = tx.objectStore(SAVE_DIR_STORE).get(key);
        request.onsuccess = function () {
          resolve(request.result || null);
        };
        request.onerror = function () {
          reject(request.error || new Error("保存フォルダ情報を読み込めませんでした。"));
        };
      });
    });
  }

  async function ensureSaveDirPermission(handle, mode) {
    if (!handle || typeof handle.queryPermission !== "function") {
      return false;
    }
    var opts = { mode: mode || "readwrite" };
    if ((await handle.queryPermission(opts)) === "granted") {
      return true;
    }
    if (typeof handle.requestPermission !== "function") {
      return false;
    }
    return (await handle.requestPermission(opts)) === "granted";
  }

  function fillFileSelect(files) {
    savedFileNames = Array.isArray(files) ? files.slice() : [];
    if (selectedSavedFile && savedFileNames.indexOf(selectedSavedFile) === -1) {
      selectedSavedFile = "";
    }
    renderSavedFileList();
  }

  function savedFileToggleLabel() {
    return selectedSavedFile || "保存済みを選ぶ";
  }

  function closeSavedFileMenu() {
    savedFilePicker.classList.remove("is-open");
    savedFileToggle.setAttribute("aria-expanded", "false");
  }

  function openSavedFileMenu() {
    savedFilePicker.classList.add("is-open");
    savedFileToggle.setAttribute("aria-expanded", "true");
    renderSavedFileList();
    savedFileSearch.focus();
  }

  function renderSavedFileList() {
    savedFileToggle.textContent = savedFileToggleLabel();
    var keyword = savedFileSearch.value.trim().toLowerCase();
    var matched = savedFileNames.filter(function (name) {
      if (!keyword) {
        return true;
      }
      return name.toLowerCase().indexOf(keyword) !== -1;
    });
    savedFileList.replaceChildren();
    if (matched.length === 0) {
      var empty = document.createElement("li");
      empty.className = "empty";
      empty.textContent = savedFileNames.length
        ? "キーワードに一致する保存はありません"
        : "保存済みファイルはまだありません";
      savedFileList.appendChild(empty);
      return;
    }
    matched.forEach(function (name) {
      var item = document.createElement("li");
      item.textContent = name;
      item.setAttribute("role", "option");
      if (name === selectedSavedFile) {
        item.className = "is-selected";
      }
      item.addEventListener("click", function () {
        selectedSavedFile = name;
        savedFileSearch.value = "";
        renderSavedFileList();
        closeSavedFileMenu();
      });
      savedFileList.appendChild(item);
    });
  }

  function fillExcelSelect(files) {
    excelFileNames = Array.isArray(files) ? files.slice() : [];
    if (selectedExcelFile && excelFileNames.indexOf(selectedExcelFile) === -1) {
      selectedExcelFile = "";
    }
    renderExcelFileList();
  }

  function excelFileToggleLabel() {
    return selectedExcelFile || "Excelを選ぶ";
  }

  function closeExcelFileMenu() {
    excelFilePicker.classList.remove("is-open");
    excelFileToggle.setAttribute("aria-expanded", "false");
  }

  function openExcelFileMenu() {
    excelFilePicker.classList.add("is-open");
    excelFileToggle.setAttribute("aria-expanded", "true");
    renderExcelFileList();
    excelFileSearch.focus();
  }

  function updateExportTargetStatus() {
    var name = exportFileHandle && exportFileHandle.name ? exportFileHandle.name : "";
    if (exportSelectedName) {
      exportSelectedName.textContent = name || "未選択";
    }
    if (exportSelectedFile) {
      exportSelectedFile.classList.toggle("is-ready", Boolean(name));
      exportSelectedFile.classList.toggle("is-empty", !name);
    }
    if (chooseExportExcelButton) {
      chooseExportExcelButton.textContent = name ? "選び直す" : "PCのExcelを選ぶ";
    }
    if (exportTargetStatus) {
      exportTargetStatus.textContent = name
        ? "このファイルに書き込みます（読み込み用のExcelとは別です）"
        : "まだ選んでいません。「PCのExcelを選ぶ」を押してください。";
    }
  }

  async function useExportFileHandle(handle) {
    if (!handle) {
      return false;
    }
    if (!/\.xlsx$/i.test(handle.name || "")) {
      showErrors(["書き込み先は .xlsx を選んでください。"]);
      return false;
    }
    var ok = await ensureSaveDirPermission(handle, "readwrite");
    if (!ok) {
      return false;
    }
    exportFileHandle = handle;
    await storeHandle(EXPORT_FILE_KEY, handle);
    updateExportTargetStatus();
    return true;
  }

  async function chooseExportExcelFile() {
    if (typeof window.showOpenFilePicker !== "function") {
      showErrors(["Edge または Chrome で index.html を開いてください。"]);
      return;
    }
    try {
      var handles = await window.showOpenFilePicker({
        id: "option-excel-export",
        multiple: false,
        types: [
          {
            description: "Excel",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            },
          },
        ],
      });
      if (await useExportFileHandle(handles && handles[0])) {
        showErrors([]);
        setStatus("書き込み先に設定しました: " + exportFileHandle.name);
      }
    } catch (error) {
      if (error && error.name !== "AbortError") {
        showErrors([error.message || "書き込み先のExcelを選べませんでした。"]);
      }
    }
  }

  async function restoreExportExcelFile() {
    try {
      var stored = await readStoredHandle(EXPORT_FILE_KEY);
      if (stored) {
        await useExportFileHandle(stored);
      }
    } catch (error) {
      exportFileHandle = null;
    }
    updateExportTargetStatus();
  }

  function renderExcelFileList() {
    excelFileToggle.textContent = excelFileToggleLabel();
    var keyword = excelFileSearch.value.trim().toLowerCase();
    var matched = excelFileNames.filter(function (name) {
      if (!keyword) {
        return true;
      }
      return name.toLowerCase().indexOf(keyword) !== -1;
    });
    excelFileList.replaceChildren();
    if (matched.length === 0) {
      var empty = document.createElement("li");
      empty.className = "empty";
      empty.textContent = excelFileNames.length
        ? "キーワードに一致するExcelはありません"
        : "Excelファイルはまだありません";
      excelFileList.appendChild(empty);
      return;
    }
    matched.forEach(function (name) {
      var item = document.createElement("li");
      item.textContent = name;
      item.setAttribute("role", "option");
      if (name === selectedExcelFile) {
        item.className = "is-selected";
      }
      item.addEventListener("click", function () {
        selectedExcelFile = name;
        excelFileSearch.value = "";
        renderExcelFileList();
        closeExcelFileMenu();
      });
      excelFileList.appendChild(item);
    });
  }

  function setSaveFolderUiLinked(name) {
    saveFolderStatus.textContent =
      "保存先はアプリフォルダ内です。途中保存は saveデータ、Excelは 仕様書一覧。" +
      (name ? "（" + name + "）" : "");
    linkSaveFolderButton.hidden = saveApi;
    linkSaveFolderButton.textContent = "アプリフォルダを変更";
    savedFilePicker.hidden = false;
    loadSavedFileButton.hidden = false;
    excelFilePicker.hidden = false;
    loadExcelFileButton.hidden = false;
  }

  function setSaveFolderUiUnlinked(message) {
    saveFolderStatus.textContent =
      message ||
      "初回だけ、index.html が入っているこのアプリのフォルダを選んでください。bat は不要です。";
    linkSaveFolderButton.hidden = false;
    linkSaveFolderButton.textContent = "このアプリのフォルダを指定";
    savedFilePicker.hidden = true;
    closeSavedFileMenu();
    loadSavedFileButton.hidden = true;
    excelFilePicker.hidden = true;
    closeExcelFileMenu();
    loadExcelFileButton.hidden = true;
  }

  async function listSavedJsonFiles(handle) {
    var names = [];
    for await (var entry of handle.values()) {
      if (entry.kind === "file" && /\.json$/i.test(entry.name)) {
        names.push(entry.name);
      }
    }
    names.sort(function (a, b) {
      return a.localeCompare(b, "ja");
    });
    return names;
  }

  async function listExcelFiles(handle) {
    var names = [];
    for await (var entry of handle.values()) {
      if (entry.kind === "file" && /\.xlsx$/i.test(entry.name)) {
        names.push(entry.name);
      }
    }
    names.sort(function (a, b) {
      return a.localeCompare(b, "ja");
    });
    return names;
  }

  async function refreshExcelFileList() {
    if (saveApi) {
      var excelRes = await fetch(apiUrl("/api/excels"), { cache: "no-store" });
      if (!excelRes.ok) {
        throw new Error("仕様書一覧 の取得ができませんでした。");
      }
      var excelData = await excelRes.json();
      fillExcelSelect(excelData.files || []);
      return;
    }
    if (!excelDirHandle) {
      fillExcelSelect([]);
      return;
    }
    var excelOk = await ensureSaveDirPermission(excelDirHandle, "read");
    if (!excelOk) {
      fillExcelSelect([]);
      return;
    }
    fillExcelSelect(await listExcelFiles(excelDirHandle));
  }

  async function refreshSavedFileList() {
    if (saveApi) {
      var res = await fetch(apiUrl("/api/saves"), { cache: "no-store" });
      if (!res.ok) {
        throw new Error("saveデータ の一覧を取得できませんでした。");
      }
      var data = await res.json();
      fillFileSelect(data.files || []);
      await refreshExcelFileList();
      return;
    }
    if (!saveDirHandle) {
      fillFileSelect([]);
      fillExcelSelect([]);
      return;
    }
    var ok = await ensureSaveDirPermission(saveDirHandle, "read");
    if (!ok) {
      appDirHandle = null;
      saveDirHandle = null;
      excelDirHandle = null;
      setSaveFolderUiUnlinked("フォルダへのアクセスが切れました。もう一度、このアプリのフォルダを指定してください。");
      return;
    }
    fillFileSelect(await listSavedJsonFiles(saveDirHandle));
    await refreshExcelFileList();
  }

  async function useAppDirHandle(handle) {
    if (!handle) {
      return false;
    }
    var ok = await ensureSaveDirPermission(handle, "readwrite");
    if (!ok) {
      return false;
    }
    try {
      await handle.getFileHandle("index.html");
    } catch (error) {
      showErrors(["index.html があるフォルダ（このアプリのフォルダ）を選んでください。"]);
      return false;
    }
    appDirHandle = handle;
    saveDirHandle = await handle.getDirectoryHandle("saveデータ", { create: true });
    excelDirHandle = await handle.getDirectoryHandle("仕様書一覧", { create: true });
    await storeHandle(APP_DIR_KEY, handle);
    setSaveFolderUiLinked(handle.name);
    await refreshSavedFileList();
    return true;
  }

  async function chooseAppFolder() {
    if (!supportsSaveFolder()) {
      showErrors(["Edge または Chrome で index.html を開いてください。"]);
      return false;
    }
    try {
      var handle = await window.showDirectoryPicker({
        id: "option-excel-app-root",
        mode: "readwrite",
      });
      if (await useAppDirHandle(handle)) {
        showErrors([]);
        setStatus("このアプリのフォルダを使えました。saveデータ と 仕様書一覧 に保存します。");
        return true;
      }
    } catch (error) {
      if (error && error.name !== "AbortError") {
        showErrors([error.message || "フォルダを指定できませんでした。"]);
      }
    }
    return false;
  }

  async function ensureAppFolder() {
    if (saveApi) {
      return true;
    }
    if (appDirHandle && saveDirHandle && excelDirHandle) {
      if (await ensureSaveDirPermission(appDirHandle, "readwrite")) {
        return true;
      }
    }
    try {
      var stored = await readStoredHandle(APP_DIR_KEY);
      if (stored && (await useAppDirHandle(stored))) {
        return true;
      }
    } catch (error) {
      // ignore
    }
    return chooseAppFolder();
  }

  async function writeExcelToHandle(filename, buffer) {
    var fileHandle = await excelDirHandle.getFileHandle(filename, { create: true });
    var writable = await fileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();
  }

  function chooseSaveFolder() {
    return chooseAppFolder();
  }

  async function probeSaveApi() {
    var bases = [""];
    if (window.location.protocol === "file:") {
      bases = ["http://127.0.0.1:8765"];
    } else {
      bases.push("http://127.0.0.1:8765");
    }
    var i;
    for (i = 0; i < bases.length; i += 1) {
      try {
        var res = await fetch(bases[i] + "/api/saves", { cache: "no-store" });
        if (!res.ok) {
          continue;
        }
        var data = await res.json();
        saveApi = true;
        saveApiBase = bases[i];
        fillFileSelect(data.files || []);
        setSaveFolderUiLinked("saveデータ");
        refreshExcelFileList().catch(function () {});
        return true;
      } catch (error) {
        // try next
      }
    }
    saveApi = false;
    saveApiBase = "";
    return false;
  }

  async function initSaveFolder() {
    if (await probeSaveApi()) {
      return;
    }
    if (!supportsSaveFolder()) {
      setSaveFolderUiUnlinked("Edge または Chrome で index.html を開いてください。");
      linkSaveFolderButton.hidden = true;
      return;
    }
    try {
      var stored = await readStoredHandle(APP_DIR_KEY);
      if (stored && (await useAppDirHandle(stored))) {
        return;
      }
    } catch (error) {
      // ignore restore errors
    }
    setSaveFolderUiUnlinked();
  }

  function buildDocumentData() {
    var options = [];
    rows.forEach(function (row, index) {
      if (!isFilled(row)) {
        return;
      }
      var price = api.normalizePriceText(row.price);
      var parsed = /^\d+(\.\d+)?$/.test(price) ? Number(price) : row.price.trim();
      options.push({
        no: String(row.no == null ? "" : row.no).trim(),
        name: row.name.trim(),
        price: parsed,
      });
    });

    var baseText = api.normalizePriceText(basePriceInput.value);
    var baseStored = /^\d+(\.\d+)?$/.test(baseText) ? Number(baseText) : basePriceInput.value.trim();

    return {
      specName: specNameInput.value.trim(),
      basePrice: baseStored,
      inputRows: Number(inputRowsSelect.value),
      selectedNos: selectedNosInput.value.trim(),
      options: options,
    };
  }

  async function writeDocumentToSaveFolder(documentData) {
    var filename = jsonFilename();
    var blob = new Blob([JSON.stringify(documentData, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    var fileHandle = await saveDirHandle.getFileHandle(filename, { create: true });
    var writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return filename;
  }

  async function saveList() {
    showErrors([]);
    var documentData = buildDocumentData();
    if (
      !documentData.specName &&
      documentData.options.length === 0 &&
      (documentData.basePrice === "" || documentData.basePrice == null)
    ) {
      showErrors(["保存する仕様名、本体価格、またはオプションがありません。"]);
      return;
    }
    if (!saveApi) {
      await probeSaveApi();
    }
    if (saveApi) {
      try {
        var apiName = jsonFilename();
        var put = await fetch(apiUrl("/api/saves/" + encodeURIComponent(apiName)), {
          method: "PUT",
          headers: { "Content-Type": "application/json; charset=utf-8" },
          body: JSON.stringify(documentData, null, 2),
        });
        if (!put.ok) {
          throw new Error("saveデータ フォルダへ保存できませんでした。");
        }
        await refreshSavedFileList();
        selectedSavedFile = apiName;
        renderSavedFileList();
        setStatus("saveデータ に保存しました: " + apiName);
        showDonePopup("保存が完了しました。\n" + apiName);
        return;
      } catch (error) {
        showErrors([error && error.message ? error.message : "saveデータ フォルダへ保存できませんでした。"]);
        return;
      }
    }
    if (!(await ensureAppFolder())) {
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
      return;
    }
    try {
      var savedName = await writeDocumentToSaveFolder(documentData);
      await refreshSavedFileList();
      selectedSavedFile = savedName;
      renderSavedFileList();
      setStatus("saveデータ に保存しました: " + savedName);
      showDonePopup("保存が完了しました。\n" + savedName);
    } catch (error) {
      showErrors([error && error.message ? error.message : "saveデータ フォルダへ保存できませんでした。"]);
    }
  }

  function applyLoaded(data) {
    if (!data || typeof data !== "object" || Array.isArray(data) || !Array.isArray(data.options)) {
      showErrors(["一覧ファイルの形式が正しくありません。"]);
      return;
    }

    var inputRows = Number(data.inputRows);
    if (api.ALLOWED_INPUT_ROWS.indexOf(inputRows) === -1) {
      inputRows = Number(inputRowsSelect.value);
    }
    if (api.ALLOWED_INPUT_ROWS.indexOf(inputRows) === -1) {
      inputRows = 300;
    }

    var nextRows = [];
    for (var i = 0; i < data.options.length; i += 1) {
      var item = data.options[i];
      if (!item || typeof item !== "object") {
        showErrors(["一覧ファイルの " + (i + 1) + " 件目が正しくありません。"]);
        return;
      }
      nextRows.push({
        no: item.no == null ? "" : String(item.no).trim(),
        name: item.name == null ? "" : String(item.name),
        price: item.price == null ? "" : String(item.price),
      });
    }

    if (nextRows.length === 0) {
      nextRows = starterRows();
    } else {
      var last = nextRows[nextRows.length - 1];
      if (isFilled(last)) {
        nextRows.push(emptyRow(nextRows));
        nextRows.push(emptyRow(nextRows));
        nextRows.push(emptyRow(nextRows));
      } else {
        nextRows.push(emptyRow(nextRows));
      }
    }

    specNameInput.value = data.specName == null ? "" : String(data.specName);
    basePriceInput.value = data.basePrice == null ? "" : String(data.basePrice);
    selectedNosInput.value = data.selectedNos == null ? "" : String(data.selectedNos);
    inputRowsSelect.value = String(inputRows);
    rows = nextRows;
    var filled = rows.filter(isFilled).length;
    var nextNo = filled + 1;
    var continueIndex = 0;
    for (var j = 0; j < rows.length; j += 1) {
      if (!isFilled(rows[j])) {
        continueIndex = j;
        break;
      }
      continueIndex = rows.length - 1;
    }
    showErrors([]);
    setStatus(
      filled
        ? "一覧を読み込みました。次は " + nextNo + " 番から続けられます。"
        : "一覧を読み込みました。"
    );
    render(continueIndex);
    updatePreview();
  }

  function excelNameToSpec(filename) {
    return String(filename || "")
      .replace(/\.(xlsx|xlsm)$/i, "")
      .replace(/_オプション$/, "")
      .trim();
  }

  async function writeSelectedCalcRow() {
    showErrors([]);
    if (!lastPreview || !lastPreview.selectedNosText) {
      showErrors(["先に選ぶオプションNoを入力してください。"]);
      return;
    }
    var row = Number(exportRowInput && exportRowInput.value);
    if (!Number.isInteger(row) || row < 1) {
      showErrors(["書き込む行は 1 以上の整数にしてください。"]);
      return;
    }
    if (!api || typeof api.writeCalcRowToXlsx !== "function") {
      showErrors(["Excelへの書き込みに必要なファイルがありません。"]);
      return;
    }
    try {
      if (!exportFileHandle) {
        showErrors(["書き込み先は、読み込み用とは別のExcelを「PCのExcelを選ぶ」から指定してください。"]);
        return;
      }
      if (!(await useExportFileHandle(exportFileHandle))) {
        showErrors(["書き込み先のExcelへのアクセスが切れました。もう一度選んでください。"]);
        return;
      }
      var name = exportFileHandle.name;
      var file = await exportFileHandle.getFile();
      var source = await file.arrayBuffer();
      var next = await api.writeCalcRowToXlsx(source, row, {
        a: lastPreview.basePrice,
        b: lastPreview.selectedNosText,
        c: lastPreview.optionTotal,
      });
      var writable = await exportFileHandle.createWritable();
      await writable.write(next);
      await writable.close();
      setStatus(name + " の " + row + " 行目（A〜C）に書き込みました。");
      showDonePopup("Excelへ書き込みました。\n" + name + " の " + row + " 行目");
    } catch (error) {
      showErrors([error && error.message ? error.message : "Excelへ書き込めませんでした。"]);
    }
  }

  function applyExcelBuffer(buffer, filename) {
    if (!api || typeof api.readXlsxOptions !== "function") {
      showErrors(["Excelの読み込みに必要なファイルがありません。"]);
      return Promise.resolve();
    }
    return api.readXlsxOptions(buffer).then(function (parsed) {
      applyLoaded({
        specName: parsed.specName || specNameInput.value.trim() || excelNameToSpec(filename),
        basePrice: parsed.basePrice !== "" ? parsed.basePrice : basePriceInput.value,
        inputRows: Number(inputRowsSelect.value),
        selectedNos: selectedNosInput.value,
        options: parsed.options,
      });
      setStatus("Excelから " + parsed.options.length + " 件を入力しました。");
    });
  }

  async function loadSelectedExcelFile() {
    var name = selectedExcelFile;
    if (!name) {
      showErrors(["読み込むExcelを選んでください。"]);
      return;
    }
    if (!saveApi) {
      await probeSaveApi();
    }
    if (saveApi) {
      try {
        var excelRes = await fetch(apiUrl("/api/excels/" + encodeURIComponent(name)), { cache: "no-store" });
        if (!excelRes.ok) {
          throw new Error("仕様書一覧 から読み込めませんでした。");
        }
        await applyExcelBuffer(await excelRes.arrayBuffer(), name);
      } catch (error) {
        showErrors([error && error.message ? error.message : "Excelを読み込めませんでした。"]);
      }
      return;
    }
    if (!(await ensureAppFolder()) || !excelDirHandle) {
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
      return;
    }
    try {
      var ok = await ensureSaveDirPermission(excelDirHandle, "read");
      if (!ok) {
        showErrors(["仕様書一覧 フォルダへのアクセスが切れました。もう一度指定してください。"]);
        return;
      }
      var fileHandle = await excelDirHandle.getFileHandle(name);
      var file = await fileHandle.getFile();
      await applyExcelBuffer(await file.arrayBuffer(), name);
    } catch (error) {
      showErrors([error && error.message ? error.message : "Excelを読み込めませんでした。"]);
    }
  }

  function loadList(file) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
        applyLoaded(JSON.parse(String(reader.result)));
      } catch (error) {
        showErrors(["一覧ファイルを読み込めませんでした。JSON形式か確認してください。"]);
      }
    };
    reader.onerror = function () {
      showErrors(["一覧ファイルを読み込めませんでした。"]);
    };
    reader.readAsText(file, "utf-8");
  }

  async function loadSelectedSavedFile() {
    var name = selectedSavedFile;
    if (!name) {
      showErrors(["読み込む保存済みファイルを選んでください。"]);
      return;
    }
    if (saveApi) {
      try {
        var apiRes = await fetch(apiUrl("/api/saves/" + encodeURIComponent(name)), { cache: "no-store" });
        if (!apiRes.ok) {
          throw new Error("saveデータ から読み込めませんでした。");
        }
        applyLoaded(await apiRes.json());
      } catch (error) {
        showErrors([error && error.message ? error.message : "saveデータ から読み込めませんでした。"]);
      }
      return;
    }
    if (!(await ensureAppFolder()) || !saveDirHandle) {
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
      return;
    }
    try {
      var ok = await ensureSaveDirPermission(saveDirHandle, "read");
      if (!ok) {
        saveDirHandle = null;
        setSaveFolderUiUnlinked("saveデータ フォルダへのアクセスが切れました。もう一度指定してください。");
        showErrors(["saveデータ から読み込めませんでした。"]);
        return;
      }
      var fileHandle = await saveDirHandle.getFileHandle(name);
      var file = await fileHandle.getFile();
      loadList(file);
    } catch (error) {
      showErrors([error && error.message ? error.message : "saveデータ から読み込めませんでした。"]);
    }
  }

  async function buildExcel() {
    showErrors([]);
    setStatus("");

    if (!api || typeof api.buildOptionWorkbook !== "function") {
      showErrors(["必要なファイルを読み込めませんでした。フォルダ一式を同じ場所に置いて、index.html を開いてください。"]);
      return;
    }

    var checked = api.validateOptionInput(currentPayload());
    if (!checked.ok) {
      showErrors(checked.errors);
      return;
    }

    try {
      var workbook = api.buildOptionWorkbook(checked);
      var buffer = await workbook.xlsx.writeBuffer();
      var filename = api.sanitizeFilename(checked.specName);
      if (!saveApi) {
        await probeSaveApi();
      }
      if (saveApi) {
        var put = await fetch(apiUrl("/api/excels/" + encodeURIComponent(filename)), {
          method: "PUT",
          headers: {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
          body: buffer,
        });
        if (!put.ok) {
          throw new Error("仕様書一覧 フォルダへ保存できませんでした。");
        }
        try {
          await refreshExcelFileList();
          selectedExcelFile = filename;
          renderExcelFileList();
        } catch (refreshError) {
          // 保存自体は成功している
        }
        setStatus("仕様書一覧 に保存しました: " + filename);
        showDonePopup("Excelの作成が完了しました。\n" + filename);
        return;
      }
      if (await ensureAppFolder()) {
        await writeExcelToHandle(filename, buffer);
        try {
          await refreshExcelFileList();
          selectedExcelFile = filename;
          renderExcelFileList();
        } catch (refreshError) {
          // 保存自体は成功している
        }
        setStatus("仕様書一覧 に保存しました: " + filename);
        showDonePopup("Excelの作成が完了しました。\n" + filename);
        return;
      }
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。Excelは Downloads には保存しません。"]);
    } catch (error) {
      showErrors([error && error.message ? error.message : "Excelを作成できませんでした。"]);
    }
  }

  document.getElementById("addRow").addEventListener("click", function () {
    rows.push(emptyRow());
    setStatus("");
    render(rows.length - 1);
    updatePreview();
  });

  linkSaveFolderButton.addEventListener("click", function () {
    chooseSaveFolder();
  });
  loadSavedFileButton.addEventListener("click", function () {
    loadSelectedSavedFile();
  });
  loadExcelFileButton.addEventListener("click", function () {
    loadSelectedExcelFile();
  });
  if (writeCalcRowButton) {
    writeCalcRowButton.addEventListener("click", function () {
      writeSelectedCalcRow();
    });
  }
  if (chooseExportExcelButton) {
    chooseExportExcelButton.addEventListener("click", function () {
      chooseExportExcelFile();
    });
  }
  savedFileToggle.addEventListener("click", function () {
    if (savedFilePicker.classList.contains("is-open")) {
      closeSavedFileMenu();
    } else {
      closeExcelFileMenu();
      openSavedFileMenu();
    }
  });
  excelFileToggle.addEventListener("click", function () {
    if (excelFilePicker.classList.contains("is-open")) {
      closeExcelFileMenu();
    } else {
      closeSavedFileMenu();
      openExcelFileMenu();
    }
  });
  savedFileSearch.addEventListener("input", function () {
    renderSavedFileList();
  });
  excelFileSearch.addEventListener("input", function () {
    renderExcelFileList();
  });
  document.addEventListener("click", function (event) {
    if (doneDialog.classList.contains("is-open")) {
      return;
    }
    if (savedFilePicker && !savedFilePicker.contains(event.target)) {
      closeSavedFileMenu();
    }
    if (excelFilePicker && !excelFilePicker.contains(event.target)) {
      closeExcelFileMenu();
    }
  });
  document.addEventListener("keydown", function (event) {
    if (event.key !== "Escape") {
      return;
    }
    if (doneDialog.classList.contains("is-open")) {
      hideDonePopup(event);
      return;
    }
    closeSavedFileMenu();
    closeExcelFileMenu();
  });
  document.getElementById("saveJson").addEventListener("click", function () {
    saveList();
  });
  document.getElementById("buildExcel").addEventListener("click", function () {
    buildExcel();
  });

  specNameInput.addEventListener("input", function () {
    setStatus("");
    updatePreview();
  });
  basePriceInput.addEventListener("input", function () {
    setStatus("");
    updatePreview();
  });
  selectedNosInput.addEventListener("input", function () {
    updatePreview();
  });
  inputRowsSelect.addEventListener("change", function () {
    setStatus("");
    updatePreview();
  });
  doneDialogOk.addEventListener("click", hideDonePopup);

  document.querySelectorAll(".copy-btn[data-copy-key]").forEach(function (button) {
    button.addEventListener("click", function () {
      copyValue(previewCopyText(button.getAttribute("data-copy-key")), button);
    });
  });

  rows = starterRows();
  render();
  updatePreview();
  initSaveFolder();
  restoreExportExcelFile();
})();
