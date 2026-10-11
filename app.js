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

  function emptyRow() {
    return { name: "", price: "" };
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

      var numberCell = document.createElement("td");
      numberCell.className = "row-no";
      numberCell.textContent = String(index + 1);
      tr.appendChild(numberCell);

      var nameInput = addField(tr, { row: row, key: "name" }, index + 1 + " 品名", row.name);
      addField(tr, { row: row, key: "price" }, index + 1 + " 値段", row.price, function (input) {
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
      options: rows.map(function (row, index) {
        return {
          no: isFilled(row) ? String(index + 1) : "",
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

  function jsonFilename() {
    var excelName = api.sanitizeFilename(specNameInput.value.trim() || "オプション");
    return excelName.replace(/\.xlsx$/i, ".json");
  }

  function saveList() {
    showErrors([]);
    var checked = api.validateOptionInput(currentPayload());
    var options = [];
    rows.forEach(function (row, index) {
      if (!isFilled(row)) {
        return;
      }
      var price = api.normalizePriceText(row.price);
      var parsed = /^\d+(\.\d+)?$/.test(price) ? Number(price) : row.price.trim();
      options.push({
        no: String(index + 1),
        name: row.name.trim(),
        price: parsed,
      });
    });

    var baseText = api.normalizePriceText(basePriceInput.value);
    var baseStored = /^\d+(\.\d+)?$/.test(baseText) ? Number(baseText) : basePriceInput.value.trim();

    if (
      !checked.ok &&
      options.length === 0 &&
      !specNameInput.value.trim() &&
      !basePriceInput.value.trim()
    ) {
      showErrors(["保存する仕様名、本体価格、またはオプションがありません。"]);
      return;
    }

    var documentData = {
      specName: specNameInput.value.trim(),
      basePrice: baseStored,
      inputRows: Number(inputRowsSelect.value),
      options: options,
    };
    var blob = new Blob([JSON.stringify(documentData, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    downloadBlob(blob, jsonFilename());
    setStatus("一覧を保存しました。");
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
        name: item.name == null ? "" : String(item.name),
        price: item.price == null ? "" : String(item.price),
      });
    }

    if (nextRows.length === 0) {
      nextRows = [emptyRow(), emptyRow(), emptyRow(), emptyRow(), emptyRow()];
    }

    specNameInput.value = data.specName == null ? "" : String(data.specName);
    basePriceInput.value = data.basePrice == null ? "" : String(data.basePrice);
    inputRowsSelect.value = String(inputRows);
    rows = nextRows;
    showErrors([]);
    setStatus("一覧を読み込みました。");
    render();
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
      var blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      downloadBlob(blob, api.sanitizeFilename(checked.specName));
      setStatus("Excelを保存しました。");
    } catch (error) {
      showErrors([error && error.message ? error.message : "Excelを作成できませんでした。"]);
    }
  }

  document.getElementById("addRow").addEventListener("click", function () {
    rows.push(emptyRow());
    setStatus("");
    render(rows.length - 1);
  });

  document.getElementById("saveJson").addEventListener("click", saveList);
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

  rows = [emptyRow(), emptyRow(), emptyRow(), emptyRow(), emptyRow()];
  render();
})();
