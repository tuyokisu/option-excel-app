(function () {
  "use strict";

  var api = window.OptionExcel;
  var rows = [];
  var tbody = document.getElementById("optionRows");
  var optionTableWrap = document.getElementById("optionTableWrap");
  var specNameInput = document.getElementById("specName");
  var modelNameInput = document.getElementById("modelName");
  var drawingNoInput = document.getElementById("drawingNo");
  var destinationInput = document.getElementById("destination");
  var workNoInput = document.getElementById("workNo");
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
  var deleteJsonButton = document.getElementById("deleteJson");
  var savedFileNames = [];
  var selectedSavedFile = "";
  var loadedJsonPath = "";
  var savedBrowseFolder = null;
  var pendingSaveDocument = null;
  var saveFolderNames = [];
  var saveFolderDialog = document.getElementById("saveFolderDialog");
  var saveFolderSearch = document.getElementById("saveFolderSearch");
  var saveFolderChoices = document.getElementById("saveFolderChoices");
  var saveFolderCancel = document.getElementById("saveFolderCancel");
  var ROOT_FOLDER_LABEL = "（フォルダなし）";
  var excelFilePicker = document.getElementById("excelFilePicker");
  var excelFileToggle = document.getElementById("excelFileToggle");
  var excelFileSearch = document.getElementById("excelFileSearch");
  var excelFileList = document.getElementById("excelFileList");
  var loadExcelFileButton = document.getElementById("loadExcelFile");
  var excelFileNames = [];
  var extraExcelFolders = [];
  var extraSavedFolders = [];
  var selectedExcelFile = "";
  var excelBrowseFolder = null;
  var exportRowInput = document.getElementById("exportRow");
  var exportTargetStatus = document.getElementById("exportTargetStatus");
  var exportSelectedFile = document.getElementById("exportSelectedFile");
  var exportSelectedName = document.getElementById("exportSelectedName");
  var writeCalcRowButton = document.getElementById("writeCalcRow");
  var chooseExportExcelButton = document.getElementById("chooseExportExcel");
  var chooseMarkExcelButton = document.getElementById("chooseMarkExcel");
  var loadMarkExcelButton = document.getElementById("loadMarkExcel");
  var markExcelSelected = document.getElementById("markExcelSelected");
  var markExcelSelectedName = document.getElementById("markExcelSelectedName");
  var markExcelStatus = document.getElementById("markExcelStatus");
  var markExcelFile = null;
  var markExcelInput = document.getElementById("markExcelInput");
  var exportFileHandle = null;
  var SAVE_DIR_DB = "option-excel-app";
  var SAVE_DIR_STORE = "handles";
  var APP_DIR_KEY = "app-root";
  var EXPORT_FILE_KEY = "export-xlsx";
  var SAVE_FOLDER_NAME = "saveデータ(読み込むよう）";
  var EXCEL_LOAD_FOLDER_NAME = "読み込み用excel";
  var EXCEL_WRITE_FOLDER_NAME = "書き込むエクセル";
  var appDirHandle = null;
  var saveDirHandle = null;
  var excelDirHandle = null;
  var excelWriteDirHandle = null;
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
    var specLabel = specNameInput.value.trim() || "（図番番号なし）";
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
      modelName: modelNameInput ? modelNameInput.value : "",
      drawingNo: drawingNoInput ? drawingNoInput.value : "",
      destination: destinationInput ? destinationInput.value : "",
      workNo: workNoInput ? workNoInput.value : "",
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

  function modelFolderName() {
    if (!api || typeof api.sanitizeFolderName !== "function") {
      return String(modelNameInput && modelNameInput.value ? modelNameInput.value : "").trim();
    }
    return api.sanitizeFolderName(modelNameInput ? modelNameInput.value : "");
  }

  function jsonRelativePath(folder) {
    var name = jsonFilename();
    var dir = folder == null ? modelFolderName() : String(folder);
    return dir ? dir + "/" + name : name;
  }

  function excelRelativePath(filename) {
    var folder = modelFolderName();
    return folder ? folder + "/" + filename : filename;
  }

  function folderFromSavedPath(rel) {
    var text = String(rel || "");
    var at = text.lastIndexOf("/");
    return at === -1 ? "" : text.slice(0, at);
  }

  async function nestedFileHandle(root, rel, create, createDirs) {
    var parts = String(rel || "").split("/").filter(Boolean);
    if (parts.length === 0) {
      throw new Error("ファイル名が正しくありません。");
    }
    var makeFile = Boolean(create);
    var makeDir = createDirs == null ? makeFile : Boolean(createDirs);
    var dir = root;
    var i;
    for (i = 0; i < parts.length - 1; i += 1) {
      dir = await dir.getDirectoryHandle(parts[i], { create: makeDir });
    }
    return {
      dir: dir,
      name: parts[parts.length - 1],
      handle: await dir.getFileHandle(parts[parts.length - 1], { create: makeFile }),
    };
  }

  function isDirEntry(entry) {
    return entry && (entry.kind === "directory" || typeof entry.getFile !== "function");
  }

  function isFileEntry(entry) {
    return entry && (entry.kind === "file" || typeof entry.getFile === "function");
  }

  async function forEachDirEntry(handle, visit) {
    if (!handle) {
      return;
    }
    try {
      if (typeof handle.entries === "function") {
        for await (var pair of handle.entries()) {
          await visit(pair[1], pair[0]);
        }
        return;
      }
      if (typeof handle.values === "function") {
        for await (var entry of handle.values()) {
          await visit(entry, entry.name);
        }
      }
    } catch (error) {
      // 一部のフォルダは列挙できない
    }
  }

  function isListedExcelName(name) {
    return /\.(xlsx|xlsm|xls)$/i.test(name || "") && !/^~\$/.test(name || "");
  }

  async function listFilesRecursive(handle, prefix, matchName, depth) {
    var names = [];
    var level = depth || 0;
    await forEachDirEntry(handle, async function (entry, name) {
      if (isDirEntry(entry) && level < 2) {
        var nested = await listFilesRecursive(entry, prefix + name + "/", matchName, level + 1);
        names = names.concat(nested);
        return;
      }
      if (isFileEntry(entry) && matchName(name)) {
        names.push(prefix + name);
      }
    });
    names.sort(function (a, b) {
      return a.localeCompare(b, "ja");
    });
    return names;
  }

  async function listJsonFilesRecursive(handle, prefix) {
    return listFilesRecursive(handle, prefix || "", function (name) {
      return /\.json$/i.test(name);
    }, 0);
  }

  async function listExcelFilesRecursive(handle, prefix) {
    return listFilesRecursive(handle, prefix || "", isListedExcelName, 0);
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

  function askPermission(handle, mode) {
    if (!handle || typeof handle.requestPermission !== "function") {
      return Promise.resolve(false);
    }
    return handle
      .requestPermission({ mode: mode || "readwrite" })
      .then(function (state) {
        return state === "granted";
      })
      .catch(function () {
        return false;
      });
  }

  function fsAccessMessage(error, fallback) {
    if (error && (error.name === "NotAllowedError" || error.name === "SecurityError")) {
      return "ブラウザがファイルの読み取りを拒否しました。ページを再読み込みし、「このアプリのフォルダを指定」でもう一度許可してから読み込んでください。";
    }
    return (error && error.message) || fallback;
  }

  async function readHandleFile(handle) {
    try {
      return await handle.getFile();
    } catch (error) {
      var blocked = error && (error.name === "NotAllowedError" || error.name === "SecurityError");
      if (!blocked) {
        throw error;
      }
      if ((await askPermission(handle, "read")) || (await askPermission(handle, "readwrite"))) {
        return handle.getFile();
      }
      throw error;
    }
  }

  async function ensureSaveDirPermission(handle, mode) {
    if (!handle || typeof handle.queryPermission !== "function") {
      return false;
    }
    var modes = mode === "read" ? ["readwrite", "read"] : [mode || "readwrite"];
    var i;
    for (i = 0; i < modes.length; i += 1) {
      try {
        if ((await handle.queryPermission({ mode: modes[i] })) === "granted") {
          return true;
        }
      } catch (error) {
        // 次の許可状態を見る
      }
    }
    return askPermission(handle, modes[0]);
  }

  function uniqueFolders(paths, extraFolders) {
    var folders = [];
    var hasRoot = false;
    function addFolder(folder) {
      if (!folder || folders.indexOf(folder) !== -1) {
        return;
      }
      folders.push(folder);
    }
    (paths || []).forEach(function (name) {
      var slash = String(name).indexOf("/");
      if (slash === -1) {
        hasRoot = true;
        return;
      }
      addFolder(name.slice(0, slash));
    });
    (extraFolders || []).forEach(addFolder);
    folders.sort(function (a, b) {
      return a.localeCompare(b, "ja");
    });
    return { folders: folders, hasRoot: hasRoot };
  }

  async function listDirectoryNames(handle) {
    var names = [];
    await forEachDirEntry(handle, async function (entry, name) {
      if (isDirEntry(entry) && name) {
        names.push(name);
      }
    });
    names.sort(function (a, b) {
      return a.localeCompare(b, "ja");
    });
    return names;
  }

  function filesInFolder(paths, folder) {
    if (folder === "") {
      return (paths || []).filter(function (name) {
        return String(name).indexOf("/") === -1;
      });
    }
    var prefix = folder + "/";
    return (paths || []).filter(function (name) {
      return String(name).indexOf(prefix) === 0;
    });
  }

  function applyBrowseFolder(folder) {
    savedBrowseFolder = folder;
    if (folder && modelNameInput) {
      modelNameInput.value = folder;
    }
  }

  function fillFileSelect(files, folders) {
    savedFileNames = Array.isArray(files) ? files.slice() : [];
    extraSavedFolders = Array.isArray(folders) ? folders.slice() : [];
    if (selectedSavedFile && savedFileNames.indexOf(selectedSavedFile) === -1) {
      selectedSavedFile = "";
    }
    renderSavedFileList();
  }

  function savedFileToggleLabel() {
    if (selectedSavedFile) {
      return selectedSavedFile;
    }
    if (savedBrowseFolder) {
      return savedBrowseFolder + " のフォルダ";
    }
    if (savedBrowseFolder === "") {
      return ROOT_FOLDER_LABEL;
    }
    return "保存済みを選ぶ";
  }

  function closeSavedFileMenu() {
    savedFilePicker.classList.remove("is-open");
    savedFileToggle.setAttribute("aria-expanded", "false");
  }

  function openSavedFileMenu() {
    savedBrowseFolder = null;
    savedFileSearch.value = "";
    savedFilePicker.classList.add("is-open");
    savedFileToggle.setAttribute("aria-expanded", "true");
    refreshSavedFileList()
      .catch(function () {})
      .then(function () {
        if (!savedFilePicker.classList.contains("is-open")) {
          return;
        }
        renderSavedFileList();
        savedFileSearch.focus();
      });
  }

  function addPickerEmpty(list, text) {
    var empty = document.createElement("li");
    empty.className = "empty";
    empty.textContent = text;
    list.appendChild(empty);
  }

  function renderSavedFileList() {
    savedFileToggle.textContent = savedFileToggleLabel();
    var keyword = savedFileSearch.value.trim().toLowerCase();
    savedFileList.replaceChildren();
    savedFileSearch.placeholder =
      savedBrowseFolder === null ? "フォルダ名で検索" : "ファイル名で検索";

    if (savedBrowseFolder === null) {
      var grouped = uniqueFolders(savedFileNames, extraSavedFolders);
      var folders = grouped.folders.filter(function (folder) {
        return !keyword || folder.toLowerCase().indexOf(keyword) !== -1;
      });
      if (grouped.hasRoot && (!keyword || ROOT_FOLDER_LABEL.toLowerCase().indexOf(keyword) !== -1)) {
        folders = folders.concat([""]);
      }
      if (folders.length === 0) {
        addPickerEmpty(
          savedFileList,
          savedFileNames.length || extraSavedFolders.length
            ? "キーワードに一致するフォルダはありません"
            : "保存済みファイルはまだありません"
        );
        return;
      }
      folders.forEach(function (folder) {
        var item = document.createElement("li");
        item.className = "folder";
        item.textContent = folder || ROOT_FOLDER_LABEL;
        item.setAttribute("role", "option");
        if (folder && folder === folderFromSavedPath(selectedSavedFile)) {
          item.classList.add("is-selected");
        }
        item.addEventListener("click", function (event) {
          event.stopPropagation();
          savedBrowseFolder = folder;
          savedFileSearch.value = "";
          renderSavedFileList();
        });
        savedFileList.appendChild(item);
      });
      return;
    }

    var back = document.createElement("li");
    back.className = "nav";
    back.textContent = "← フォルダ一覧へ";
    back.addEventListener("click", function (event) {
      event.stopPropagation();
      savedBrowseFolder = null;
      savedFileSearch.value = "";
      renderSavedFileList();
    });
    savedFileList.appendChild(back);

    var files = filesInFolder(savedFileNames, savedBrowseFolder).filter(function (name) {
      if (!keyword) {
        return true;
      }
      return name.toLowerCase().indexOf(keyword) !== -1;
    });
    if (files.length === 0) {
      addPickerEmpty(savedFileList, "このフォルダに保存はありません");
      return;
    }
    files.forEach(function (name) {
      var item = document.createElement("li");
      var folder = folderFromSavedPath(name);
      item.textContent = name.indexOf("/") === -1 ? name : name.slice(name.lastIndexOf("/") + 1);
      item.setAttribute("role", "option");
      if (name === selectedSavedFile) {
        item.className = "is-selected";
      }
      item.addEventListener("click", function (event) {
        event.stopPropagation();
        selectedSavedFile = name;
        applyBrowseFolder(folder);
        savedFileSearch.value = "";
        renderSavedFileList();
        closeSavedFileMenu();
      });
      savedFileList.appendChild(item);
    });
  }

  function fillExcelSelect(files, folders) {
    excelFileNames = Array.isArray(files) ? files.slice() : [];
    extraExcelFolders = Array.isArray(folders) ? folders.slice() : [];
    if (selectedExcelFile && excelFileNames.indexOf(selectedExcelFile) === -1) {
      selectedExcelFile = "";
    }
    renderExcelFileList();
  }

  function excelFileToggleLabel() {
    if (selectedExcelFile) {
      return selectedExcelFile;
    }
    if (excelBrowseFolder) {
      return excelBrowseFolder + " のフォルダ";
    }
    if (excelBrowseFolder === "") {
      return ROOT_FOLDER_LABEL;
    }
    return "Excelを選ぶ";
  }

  function closeExcelFileMenu() {
    excelFilePicker.classList.remove("is-open");
    excelFileToggle.setAttribute("aria-expanded", "false");
  }

  function openExcelFileMenu() {
    excelBrowseFolder = null;
    excelFileSearch.value = "";
    excelFilePicker.classList.add("is-open");
    excelFileToggle.setAttribute("aria-expanded", "true");
    refreshExcelFileList()
      .catch(function () {})
      .then(function () {
        if (!excelFilePicker.classList.contains("is-open")) {
          return;
        }
        renderExcelFileList();
        excelFileSearch.focus();
      });
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
    if (!/\.xlsx$/i.test(handle.name || "") && !/\.xls$/i.test(handle.name || "")) {
      showErrors(["書き込み先は .xlsx か .xls を選んでください。"]);
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
              "application/vnd.ms-excel": [".xls"],
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

  function updateMarkExcelStatus() {
    var name = markExcelFile && markExcelFile.name ? markExcelFile.name : "";
    if (markExcelSelectedName) {
      markExcelSelectedName.textContent = name || "未選択";
    }
    if (markExcelSelected) {
      markExcelSelected.classList.toggle("is-ready", Boolean(name));
      markExcelSelected.classList.toggle("is-empty", !name);
    }
    if (chooseMarkExcelButton) {
      chooseMarkExcelButton.textContent = name ? "選び直す" : "PCのExcelを選ぶ";
    }
    if (markExcelStatus) {
      markExcelStatus.textContent = name
        ? name + " を選択しました。"
        : "まだ選んでいません。番号の横に ○ があるExcelを選ぶか、下で番号を手入力してください。";
    }
  }

  function useMarkExcelFile(file) {
    if (!file) {
      return;
    }
    markExcelFile = file;
    updateMarkExcelStatus();
    showErrors([]);
    setStatus("仕様Excelを選びました: " + file.name);
    loadMarkedOptionNos();
  }

  async function loadMarkedOptionNos() {
    showErrors([]);
    if (!markExcelFile) {
      showErrors(["先に、丸の付いた仕様Excelを「PCのExcelを選ぶ」から指定してください。"]);
      return;
    }
    if (!api || typeof api.readMarkedOptionNos !== "function") {
      showErrors(["丸付き番号の読み込みに必要なファイルがありません。"]);
      return;
    }
    try {
      var source = await markExcelFile.arrayBuffer();
      var parsed = await api.readMarkedOptionNos(source);
      var nos = Array.isArray(parsed) ? parsed : parsed.nos || [];
      if (rows.filter(isFilled).length === 0) {
        showErrors(["先にオプション一覧（オプション表）を読み込むか入力してください。丸付きExcelは番号だけ使います。"]);
        return;
      }
      selectedNosInput.value = nos.join(",");
      if (markExcelStatus) {
        markExcelStatus.textContent =
          nos.length + "件の番号を読みました: " + nos.join(", ") + "。下の「選ぶオプションNo」に入れています。";
      }
      setStatus("丸付き番号を読み込みました: " + nos.join(", "));
      updatePreview();
    } catch (error) {
      showErrors([error && error.message ? error.message : "丸付き番号を読み込めませんでした。"]);
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
    excelFileList.replaceChildren();
    excelFileSearch.placeholder =
      excelBrowseFolder === null ? "フォルダ名で検索" : "ファイル名で検索";

    if (excelBrowseFolder === null) {
      var grouped = uniqueFolders(excelFileNames, extraExcelFolders);
      var folders = grouped.folders.filter(function (folder) {
        return !keyword || folder.toLowerCase().indexOf(keyword) !== -1;
      });
      if (grouped.hasRoot && (!keyword || ROOT_FOLDER_LABEL.toLowerCase().indexOf(keyword) !== -1)) {
        folders = folders.concat([""]);
      }
      if (folders.length === 0) {
        addPickerEmpty(
          excelFileList,
          excelFileNames.length ? "キーワードに一致するフォルダはありません" : "Excelファイルはまだありません"
        );
        return;
      }
      folders.forEach(function (folder) {
        var item = document.createElement("li");
        item.className = "folder";
        item.textContent = folder || ROOT_FOLDER_LABEL;
        item.setAttribute("role", "option");
        item.addEventListener("click", function (event) {
          event.stopPropagation();
          excelBrowseFolder = folder;
          if (folder && modelNameInput) {
            modelNameInput.value = folder;
          }
          excelFileSearch.value = "";
          renderExcelFileList();
        });
        excelFileList.appendChild(item);
      });
      return;
    }

    var back = document.createElement("li");
    back.className = "nav";
    back.textContent = "← フォルダ一覧へ";
    back.addEventListener("click", function (event) {
      event.stopPropagation();
      excelBrowseFolder = null;
      excelFileSearch.value = "";
      renderExcelFileList();
    });
    excelFileList.appendChild(back);

    var files = filesInFolder(excelFileNames, excelBrowseFolder).filter(function (name) {
      if (!keyword) {
        return true;
      }
      return name.toLowerCase().indexOf(keyword) !== -1;
    });
    if (files.length === 0) {
      addPickerEmpty(excelFileList, "このフォルダに一致するExcelはありません");
      return;
    }
    files.forEach(function (name) {
      var item = document.createElement("li");
      item.textContent = name.indexOf("/") === -1 ? name : name.slice(name.lastIndexOf("/") + 1);
      item.setAttribute("role", "option");
      if (name === selectedExcelFile) {
        item.className = "is-selected";
      }
      item.addEventListener("click", function () {
        selectedExcelFile = name;
        excelBrowseFolder = folderFromSavedPath(name);
        excelFileSearch.value = "";
        renderExcelFileList();
        closeExcelFileMenu();
      });
      excelFileList.appendChild(item);
    });
  }

  async function existingDirHandle(parent, name) {
    try {
      return await parent.getDirectoryHandle(name);
    } catch (error) {
      throw new Error(
        "「" + name + "」フォルダが見つかりません。新しく作らず、あらかじめ用意したフォルダを使います。"
      );
    }
  }

  function setSaveFolderUiLinked(name) {
    saveFolderStatus.textContent =
      "JSONは " +
      SAVE_FOLDER_NAME +
      "、Excelの読み込みは " +
      EXCEL_LOAD_FOLDER_NAME +
      "、Excelの作成は " +
      EXCEL_WRITE_FOLDER_NAME +
      " です。" +
      (name ? "（" + name + "）" : "");
    linkSaveFolderButton.hidden = saveApi;
    linkSaveFolderButton.textContent = "アプリフォルダを変更";
    savedFilePicker.hidden = false;
    loadSavedFileButton.hidden = false;
    if (deleteJsonButton) {
      deleteJsonButton.hidden = false;
    }
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
    if (deleteJsonButton) {
      deleteJsonButton.hidden = true;
    }
    excelFilePicker.hidden = true;
    closeExcelFileMenu();
    loadExcelFileButton.hidden = true;
  }

  async function listSavedJsonFiles(handle) {
    return listJsonFilesRecursive(handle, "");
  }

  async function listExcelFiles(handle) {
    return listExcelFilesRecursive(handle, "");
  }

  async function refreshExcelFileList() {
    if (saveApi) {
      var excelRes = await fetch(apiUrl("/api/excels"), { cache: "no-store" });
      if (!excelRes.ok) {
        throw new Error(EXCEL_LOAD_FOLDER_NAME + " の取得ができませんでした。");
      }
      var excelData = await excelRes.json();
      fillExcelSelect(excelData.files || [], excelData.folders || []);
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
    fillExcelSelect(await listExcelFiles(excelDirHandle), await listDirectoryNames(excelDirHandle));
  }

  async function refreshSavedFileList() {
    if (saveApi) {
      var res = await fetch(apiUrl("/api/saves"), { cache: "no-store" });
      if (!res.ok) {
        throw new Error(SAVE_FOLDER_NAME + " の一覧を取得できませんでした。");
      }
      var data = await res.json();
      fillFileSelect(data.files || [], data.folders || []);
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
      excelWriteDirHandle = null;
      setSaveFolderUiUnlinked("フォルダへのアクセスが切れました。もう一度、このアプリのフォルダを指定してください。");
      return;
    }
    fillFileSelect(await listSavedJsonFiles(saveDirHandle), await listDirectoryNames(saveDirHandle));
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
    saveDirHandle = await existingDirHandle(handle, SAVE_FOLDER_NAME);
    excelDirHandle = await existingDirHandle(handle, EXCEL_LOAD_FOLDER_NAME);
    excelWriteDirHandle = await existingDirHandle(handle, EXCEL_WRITE_FOLDER_NAME);
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
        setStatus(
          "このアプリのフォルダを使えました。JSONは " +
            SAVE_FOLDER_NAME +
            "、Excelは " +
            EXCEL_LOAD_FOLDER_NAME +
            " / " +
            EXCEL_WRITE_FOLDER_NAME +
            " です。"
        );
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
    if (appDirHandle && saveDirHandle && excelDirHandle && excelWriteDirHandle) {
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
    var nested = await nestedFileHandle(excelWriteDirHandle, filename, true);
    var writable = await nested.handle.createWritable();
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
        fillFileSelect(data.files || [], data.folders || []);
        setSaveFolderUiLinked(SAVE_FOLDER_NAME);
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
      modelName: modelNameInput ? modelNameInput.value.trim() : "",
      drawingNo: drawingNoInput ? drawingNoInput.value.trim() : "",
      destination: destinationInput ? destinationInput.value.trim() : "",
      workNo: workNoInput ? workNoInput.value.trim() : "",
      basePrice: baseStored,
      inputRows: Number(inputRowsSelect.value),
      selectedNos: selectedNosInput.value.trim(),
      options: options,
    };
  }

  function updateOverwriteButton() {
    var button = document.getElementById("overwriteJson");
    if (!button) {
      return;
    }
    button.disabled = !loadedJsonPath;
    button.title = loadedJsonPath
      ? loadedJsonPath + " を上書きします"
      : "読み込んだJSON、またはsaveしたファイルを上書きします";
  }

  function rememberLoadedJson(name) {
    loadedJsonPath = name || "";
    if (loadedJsonPath) {
      selectedSavedFile = loadedJsonPath;
    }
    updateOverwriteButton();
  }

  async function writeDocumentToRelativePath(documentData, rel) {
    var blob = new Blob([JSON.stringify(documentData, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    var nested = await nestedFileHandle(saveDirHandle, rel, true, false);
    var writable = await nested.handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return rel;
  }

  async function saveDocumentAt(documentData, rel, overwrite) {
    if (saveApi) {
      var put = await fetch(apiUrl("/api/saves/" + encodeURIComponent(rel)), {
        method: "PUT",
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify(documentData, null, 2),
      });
      if (!put.ok) {
        throw new Error(SAVE_FOLDER_NAME + " へ保存できませんでした。");
      }
    } else {
      if (!(await ensureAppFolder()) || !saveDirHandle) {
        throw new Error("このアプリのフォルダ（index.html がある場所）を指定してください。");
      }
      await writeDocumentToRelativePath(documentData, rel);
    }
    rememberLoadedJson(rel);
    await refreshSavedFileList();
    renderSavedFileList();
    setStatus((overwrite ? "上書きしました: " : SAVE_FOLDER_NAME + " に保存しました: ") + rel);
    showDonePopup((overwrite ? "上書きしました。\n" : "保存が完了しました。\n") + rel);
  }

  async function listSaveFolders() {
    if (saveApi) {
      var res = await fetch(apiUrl("/api/saves"), { cache: "no-store" });
      if (!res.ok) {
        throw new Error(SAVE_FOLDER_NAME + " のフォルダ一覧を取得できませんでした。");
      }
      var data = await res.json();
      return Array.isArray(data.folders) ? data.folders : [];
    }
    if (!saveDirHandle) {
      return [];
    }
    return listDirectoryNames(saveDirHandle);
  }

  function closeSaveFolderDialog() {
    pendingSaveDocument = null;
    saveFolderNames = [];
    if (!saveFolderDialog) {
      return;
    }
    saveFolderDialog.classList.remove("is-open");
    saveFolderDialog.hidden = true;
  }

  function renderSaveFolderChoices() {
    if (!saveFolderChoices) {
      return;
    }
    var keyword = saveFolderSearch ? saveFolderSearch.value.trim().toLowerCase() : "";
    saveFolderChoices.replaceChildren();
    var folders = saveFolderNames.filter(function (folder) {
      return !keyword || String(folder).toLowerCase().indexOf(keyword) !== -1;
    });
    if (folders.length === 0) {
      addPickerEmpty(
        saveFolderChoices,
        saveFolderNames.length ? "キーワードに一致するフォルダはありません" : "保存先のフォルダはまだありません"
      );
      return;
    }
    folders.forEach(function (folder) {
      var item = document.createElement("li");
      item.className = "folder";
      item.textContent = folder;
      item.setAttribute("role", "option");
      item.addEventListener("click", function (event) {
        event.stopPropagation();
        commitSave(folder);
      });
      saveFolderChoices.appendChild(item);
    });
  }

  function openSaveFolderDialog(folders) {
    saveFolderNames = folders.slice().sort(function (a, b) {
      return String(a).localeCompare(String(b), "ja");
    });
    if (saveFolderSearch) {
      saveFolderSearch.value = "";
    }
    renderSaveFolderChoices();
    saveFolderDialog.hidden = false;
    saveFolderDialog.classList.add("is-open");
    if (saveFolderSearch) {
      saveFolderSearch.focus();
    }
  }

  async function commitSave(folder) {
    var documentData = pendingSaveDocument;
    closeSaveFolderDialog();
    if (!documentData || !folder) {
      return;
    }
    showErrors([]);
    try {
      await saveDocumentAt(documentData, jsonRelativePath(folder), false);
    } catch (error) {
      showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " へ保存できませんでした。"]);
    }
  }

  async function overwriteSave() {
    showErrors([]);
    closeSavedFileMenu();
    if (!loadedJsonPath) {
      showErrors(["上書きするファイルがありません。先にJSONを読み込むか、saveするで保存してください。"]);
      return;
    }
    var documentData = buildDocumentData();
    if (
      !documentData.specName &&
      documentData.options.length === 0 &&
      (documentData.basePrice === "" || documentData.basePrice == null)
    ) {
      showErrors(["保存する図番番号、本体価格、またはオプションがありません。"]);
      return;
    }
    if (!window.confirm("「" + loadedJsonPath + "」を上書きします。よろしいですか？")) {
      return;
    }
    if (!saveApi) {
      await probeSaveApi();
    }
    try {
      await saveDocumentAt(documentData, loadedJsonPath, true);
    } catch (error) {
      showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " へ上書きできませんでした。"]);
    }
  }

  async function saveList() {
    showErrors([]);
    closeSavedFileMenu();
    var documentData = buildDocumentData();
    if (
      !documentData.specName &&
      documentData.options.length === 0 &&
      (documentData.basePrice === "" || documentData.basePrice == null)
    ) {
      showErrors(["保存する図番番号、本体価格、またはオプションがありません。"]);
      return;
    }
    if (!saveApi) {
      await probeSaveApi();
    }
    if (!saveApi && !(await ensureAppFolder())) {
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
      return;
    }
    try {
      pendingSaveDocument = documentData;
      openSaveFolderDialog(await listSaveFolders());
    } catch (error) {
      pendingSaveDocument = null;
      showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " のフォルダ一覧を取得できませんでした。"]);
    }
  }

  function applyLoaded(data) {
    if (!data || typeof data !== "object" || Array.isArray(data) || !Array.isArray(data.options)) {
      showErrors(["一覧ファイルの形式が正しくありません。"]);
      return false;
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
        return false;
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
    if (modelNameInput) {
      var loadedModel = data.modelName == null ? "" : String(data.modelName);
      if (!loadedModel.trim() && data.modelFolder) {
        loadedModel = String(data.modelFolder);
      }
      modelNameInput.value = loadedModel;
    }
    if (drawingNoInput) {
      drawingNoInput.value = data.drawingNo == null ? "" : String(data.drawingNo);
    }
    if (destinationInput) {
      destinationInput.value = data.destination == null ? "" : String(data.destination);
    }
    if (workNoInput) {
      workNoInput.value = data.workNo == null ? "" : String(data.workNo);
    }
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
    return true;
  }

  function excelNameToSpec(filename) {
    return String(filename || "")
      .replace(/\.(xlsx|xlsm)$/i, "")
      .replace(/_オプション$/, "")
      .trim();
  }

  async function writeExportBytes(bytes) {
    var data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    var original = exportFileHandle.name;
    var nextName =
      api && typeof api.safeExportFilename === "function"
        ? api.safeExportFilename(original, data)
        : original;
    var handle = exportFileHandle;
    if (nextName !== original) {
      if (typeof exportFileHandle.getParent !== "function") {
        throw new Error(
          "「" +
            original +
            "」は拡張子が .xls のままでは、新しいExcel形式を保存できません。Excelで .xlsx として保存してから書き込んでください。"
        );
      }
      var parent = await exportFileHandle.getParent();
      if (!(await ensureSaveDirPermission(parent, "readwrite"))) {
        throw new Error("同じフォルダへ .xlsx を保存できませんでした。もう一度フォルダの許可を出してください。");
      }
      handle = await parent.getFileHandle(nextName, { create: true });
      exportFileHandle = handle;
      await storeHandle(EXPORT_FILE_KEY, handle);
      updateExportTargetStatus();
    }
    var writable = await handle.createWritable();
    await writable.write(data);
    await writable.close();
    return { name: handle.name, original: original, renamed: nextName !== original };
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
      var file = await readHandleFile(exportFileHandle);
      var source = await file.arrayBuffer();
      var next = await api.writeCalcRowToXlsx(source, row, {
        a: lastPreview.basePrice,
        b: lastPreview.selectedNosText,
        c: lastPreview.optionTotal,
        modelName: modelNameInput ? modelNameInput.value.trim() : "",
        drawingNo: drawingNoInput ? drawingNoInput.value.trim() : "",
        destination: destinationInput ? destinationInput.value.trim() : "",
        workNo: workNoInput ? workNoInput.value.trim() : "",
      });
      var saved = await writeExportBytes(next);
      setStatus(saved.name + " の " + row + " 行目（E・G・H）に書き込みました。");
      showDonePopup(
        saved.renamed
          ? "Excelへ書き込みました。\n" +
              saved.original +
              " は形式が合わないため、" +
              saved.name +
              " に保存しました。\n" +
              row +
              " 行目"
          : "Excelへ書き込みました。\n" + saved.name + " の " + row + " 行目"
      );
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
      loadedJsonPath = "";
      updateOverwriteButton();
      applyLoaded({
        specName:
          parsed.specName ||
          specNameInput.value.trim() ||
          excelNameToSpec(String(filename || "").split("/").pop()),
        basePrice: parsed.basePrice !== "" ? parsed.basePrice : basePriceInput.value,
        inputRows: Number(inputRowsSelect.value),
        selectedNos: selectedNosInput.value,
        options: parsed.options,
        modelFolder: folderFromSavedPath(filename),
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
          throw new Error(EXCEL_LOAD_FOLDER_NAME + " から読み込めませんでした。");
        }
        await applyExcelBuffer(await excelRes.arrayBuffer(), name);
      } catch (error) {
        showErrors([fsAccessMessage(error, "Excelを読み込めませんでした。")]);
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
        showErrors([EXCEL_LOAD_FOLDER_NAME + " へのアクセスが切れました。もう一度指定してください。"]);
        return;
      }
      var fileHandle = (await nestedFileHandle(excelDirHandle, name, false)).handle;
      var file = await readHandleFile(fileHandle);
      await applyExcelBuffer(await file.arrayBuffer(), name);
    } catch (error) {
      showErrors([fsAccessMessage(error, "Excelを読み込めませんでした。")]);
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

  async function removeEmptyFolder(root, folderName) {
    if (!folderName || typeof root.removeEntry !== "function") {
      return;
    }
    try {
      await root.removeEntry(folderName);
    } catch (error) {
      // 中にファイルが残っていればそのまま
    }
  }

  async function deleteSelectedSavedFile() {
    var name = selectedSavedFile;
    if (!name) {
      showErrors(["消す保存済みファイルを選んでください。"]);
      return;
    }
    if (!window.confirm("「" + name + "」を削除します。よろしいですか？")) {
      return;
    }
    showErrors([]);
    if (!saveApi) {
      await probeSaveApi();
    }
    if (saveApi) {
      try {
        var del = await fetch(apiUrl("/api/saves/" + encodeURIComponent(name)), { method: "DELETE" });
        if (!del.ok) {
          throw new Error(SAVE_FOLDER_NAME + " から削除できませんでした。");
        }
        selectedSavedFile = "";
        if (loadedJsonPath === name) {
          rememberLoadedJson("");
        }
        await refreshSavedFileList();
        setStatus("JSONを削除しました: " + name);
        showDonePopup("JSONを削除しました。\n" + name);
      } catch (error) {
        showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " から削除できませんでした。"]);
      }
      return;
    }
    if (!(await ensureAppFolder()) || !saveDirHandle) {
      showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
      return;
    }
    try {
      var ok = await ensureSaveDirPermission(saveDirHandle, "readwrite");
      if (!ok) {
        showErrors([SAVE_FOLDER_NAME + " へのアクセスが切れました。もう一度指定してください。"]);
        return;
      }
      var nested = await nestedFileHandle(saveDirHandle, name, false);
      await nested.dir.removeEntry(nested.name);
      await removeEmptyFolder(saveDirHandle, folderFromSavedPath(name));
      selectedSavedFile = "";
      if (loadedJsonPath === name) {
        rememberLoadedJson("");
      }
      await refreshSavedFileList();
      setStatus("JSONを削除しました: " + name);
      showDonePopup("JSONを削除しました。\n" + name);
    } catch (error) {
      showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " から削除できませんでした。"]);
    }
  }

  async function loadSavedJsonFromApi(name) {
    var apiRes = await fetch(apiUrl("/api/saves/" + encodeURIComponent(name)), { cache: "no-store" });
    if (!apiRes.ok) {
      throw new Error(SAVE_FOLDER_NAME + " から読み込めませんでした。");
    }
    var loaded = await apiRes.json();
    loaded.modelFolder = folderFromSavedPath(name);
    if (applyLoaded(loaded)) {
      rememberLoadedJson(name);
    }
  }

  function readSavedJsonFile(file, name) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
          var loaded = JSON.parse(String(reader.result));
          loaded.modelFolder = folderFromSavedPath(name);
          if (applyLoaded(loaded)) {
            rememberLoadedJson(name);
          }
      } catch (parseError) {
        showErrors(["一覧ファイルを読み込めませんでした。JSON形式か確認してください。"]);
      }
    };
    reader.onerror = function () {
      showErrors(["一覧ファイルを読み込めませんでした。"]);
    };
    reader.readAsText(file, "utf-8");
  }

  async function loadSelectedSavedFile(permissionPromise) {
    var name = selectedSavedFile;
    if (!name) {
      showErrors(["読み込む保存済みファイルを選んでください。"]);
      return;
    }
    showErrors([]);
    if (!saveApi) {
      await probeSaveApi();
    }
    if (saveApi) {
      try {
        await loadSavedJsonFromApi(name);
      } catch (error) {
        showErrors([error && error.message ? error.message : SAVE_FOLDER_NAME + " から読み込めませんでした。"]);
      }
      return;
    }
    if (saveDirHandle && permissionPromise) {
      await permissionPromise;
    }
    if (!(saveDirHandle && (await ensureSaveDirPermission(saveDirHandle, "readwrite")))) {
      if (!(await ensureAppFolder()) || !saveDirHandle) {
        showErrors(["このアプリのフォルダ（index.html がある場所）を指定してください。"]);
        return;
      }
    }
    try {
      var nested = await nestedFileHandle(saveDirHandle, name, false);
      var file = await readHandleFile(nested.handle);
      readSavedJsonFile(file, name);
    } catch (error) {
      if (!saveApi && (await probeSaveApi())) {
        try {
          await loadSavedJsonFromApi(name);
          return;
        } catch (apiError) {
          error = apiError;
        }
      }
      showErrors([fsAccessMessage(error, SAVE_FOLDER_NAME + " から読み込めませんでした。")]);
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
      var filename = excelRelativePath(api.sanitizeFilename(checked.specName));
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
          throw new Error(EXCEL_WRITE_FOLDER_NAME + " へ保存できませんでした。");
        }
        setStatus(EXCEL_WRITE_FOLDER_NAME + " に保存しました: " + filename);
        showDonePopup("Excelの作成が完了しました。\n" + filename);
        return;
      }
      if (await ensureAppFolder()) {
        await writeExcelToHandle(filename, buffer);
        setStatus(EXCEL_WRITE_FOLDER_NAME + " に保存しました: " + filename);
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
    var permissionPromise =
      !saveApi && saveDirHandle ? askPermission(saveDirHandle, "readwrite") : Promise.resolve(false);
    loadSelectedSavedFile(permissionPromise);
  });
  if (deleteJsonButton) {
    deleteJsonButton.addEventListener("click", function () {
      deleteSelectedSavedFile();
    });
  }
  loadExcelFileButton.addEventListener("click", function () {
    if (!saveApi && excelDirHandle) {
      askPermission(excelDirHandle, "readwrite");
    }
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
  if (chooseMarkExcelButton) {
    chooseMarkExcelButton.addEventListener("click", function () {
      if (!markExcelInput) {
        showErrors(["このブラウザではExcelを選べません。"]);
        return;
      }
      markExcelInput.value = "";
      markExcelInput.click();
    });
  }
  if (markExcelInput) {
    markExcelInput.addEventListener("change", function () {
      var file = markExcelInput.files && markExcelInput.files[0];
      useMarkExcelFile(file);
    });
  }
  if (loadMarkExcelButton) {
    loadMarkExcelButton.addEventListener("click", function () {
      loadMarkedOptionNos();
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
    if (!event.target || !event.target.isConnected) {
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
    if (saveFolderDialog && saveFolderDialog.classList.contains("is-open")) {
      closeSaveFolderDialog();
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
  var overwriteJsonButton = document.getElementById("overwriteJson");
  if (overwriteJsonButton) {
    overwriteJsonButton.addEventListener("click", function () {
      overwriteSave();
    });
  }
  updateOverwriteButton();
  if (saveFolderCancel) {
    saveFolderCancel.addEventListener("click", function (event) {
      event.stopPropagation();
      closeSaveFolderDialog();
    });
  }
  if (saveFolderSearch) {
    saveFolderSearch.addEventListener("input", function () {
      renderSaveFolderChoices();
    });
  }
  if (saveFolderDialog) {
    saveFolderDialog.addEventListener("click", function (event) {
      if (event.target === saveFolderDialog) {
        closeSaveFolderDialog();
      }
    });
  }
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
  updateMarkExcelStatus();
})();
