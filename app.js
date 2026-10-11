(function () {
  "use strict";

  var api = window.OptionExcel;
  var rows = [];
  var tbody = document.getElementById("optionRows");
  var specNameInput = document.getElementById("specName");
  var basePriceInput = document.getElementById("basePrice");
  var inputRowsSelect = document.getElementById("inputRows");
  var errorBox = document.getElementById("errors");
  var statusBox = document.getElementById("status");
  var optionCount = document.getElementById("optionCount");
  var loadInput = document.getElementById("loadJson");
  var linkSaveFolderButton = document.getElementById("linkSaveFolder");
  var saveFolderStatus = document.getElementById("saveFolderStatus");
  var savedFileSelect = document.getElementById("savedFileSelect");
  var loadSavedFileButton = document.getElementById("loadSavedFile");
  var refreshSavedFilesButton = document.getElementById("refreshSavedFiles");
  var SAVE_DIR_DB = "option-excel-app";
  var SAVE_DIR_STORE = "handles";
  var APP_DIR_KEY = "app-root";
  var appDirHandle = null;
  var saveDirHandle = null;
  var excelDirHandle = null;
  var saveApi = false;
  var saveApiBase = "";

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
      });
      actionCell.appendChild(removeButton);
      tr.appendChild(actionCell);
      tbody.appendChild(tr);

      if (focusIndex === index) {
        nameInput.focus();
      }
    });
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
    savedFileSelect.replaceChildren();
    var placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = files.length ? "保存済みを選ぶ" : "保存済みファイルはまだありません";
    savedFileSelect.appendChild(placeholder);
    files.forEach(function (name) {
      var option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      savedFileSelect.appendChild(option);
    });
  }

  function setSaveFolderUiLinked(name) {
    saveFolderStatus.textContent =
      "保存先はアプリフォルダ内です。途中保存は saveデータ、Excelは 仕様書一覧。" +
      (name ? "（" + name + "）" : "");
    linkSaveFolderButton.hidden = saveApi;
    linkSaveFolderButton.textContent = "アプリフォルダを変更";
    savedFileSelect.hidden = false;
    loadSavedFileButton.hidden = false;
    refreshSavedFilesButton.hidden = false;
  }

  function setSaveFolderUiUnlinked(message) {
    saveFolderStatus.textContent =
      message ||
      "初回だけ、index.html が入っているこのアプリのフォルダを選んでください。bat は不要です。";
    linkSaveFolderButton.hidden = false;
    linkSaveFolderButton.textContent = "このアプリのフォルダを指定";
    savedFileSelect.hidden = true;
    loadSavedFileButton.hidden = true;
    refreshSavedFilesButton.hidden = true;
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

  async function refreshSavedFileList() {
    if (saveApi) {
      var res = await fetch(apiUrl("/api/saves"), { cache: "no-store" });
      if (!res.ok) {
        throw new Error("saveデータ の一覧を取得できませんでした。");
      }
      var data = await res.json();
      fillFileSelect(data.files || []);
      return;
    }
    if (!saveDirHandle) {
      fillFileSelect([]);
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
        savedFileSelect.value = apiName;
        setStatus("saveデータ に保存しました: " + apiName);
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
      savedFileSelect.value = savedName;
      setStatus("saveデータ に保存しました: " + savedName);
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
      showErrors(["一覧ファイルの入力行数は 100、200、300、500 のいずれかにしてください。"]);
      return;
    }

    var nextRows = [];
    for (var i = 0; i < data.options.length; i += 1) {
      var item = data.options[i];
      if (!item || typeof item !== "object") {
        showErrors(["一覧ファイルの " + (i + 1) + " 件目が正しくありません。"]);
        return;
      }
      nextRows.push({
        no:
          item.no == null || String(item.no).trim() === ""
            ? String(i + 1)
            : String(item.no),
        name: item.name == null ? "" : String(item.name),
        price: item.price == null ? "" : String(item.price),
      });
    }

    if (nextRows.length === 0) {
      nextRows = starterRows();
    } else {
      var last = nextRows[nextRows.length - 1];
      if (isFilled(last)) {
        nextRows.push(emptyRow(nextRows), emptyRow(nextRows), emptyRow(nextRows));
      } else {
        nextRows.push(emptyRow(nextRows));
      }
    }

    specNameInput.value = data.specName == null ? "" : String(data.specName);
    basePriceInput.value = data.basePrice == null ? "" : String(data.basePrice);
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
    var name = savedFileSelect.value;
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
        setStatus("仕様書一覧 に保存しました: " + filename);
        return;
      }
      if (await ensureAppFolder()) {
        await writeExcelToHandle(filename, buffer);
        setStatus("仕様書一覧 に保存しました: " + filename);
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
  });

  linkSaveFolderButton.addEventListener("click", function () {
    chooseSaveFolder();
  });
  refreshSavedFilesButton.addEventListener("click", function () {
    refreshSavedFileList().catch(function (error) {
      showErrors([error && error.message ? error.message : "保存済み一覧を更新できませんでした。"]);
    });
  });
  loadSavedFileButton.addEventListener("click", function () {
    loadSelectedSavedFile();
  });
  document.getElementById("saveJson").addEventListener("click", function () {
    saveList();
  });
  document.getElementById("buildExcel").addEventListener("click", function () {
    buildExcel();
  });

  loadInput.addEventListener("change", function () {
    var file = loadInput.files && loadInput.files[0];
    loadInput.value = "";
    if (file) {
      loadList(file);
    }
  });

  specNameInput.addEventListener("input", function () {
    setStatus("");
  });
  basePriceInput.addEventListener("input", function () {
    setStatus("");
  });
  inputRowsSelect.addEventListener("change", function () {
    setStatus("");
  });

  rows = starterRows();
  render();
  initSaveFolder();
})();
