(function (factory) {
  "use strict";

  var api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }

  if (typeof window !== "undefined") {
    window.buildOptionWorkbook = api.buildOptionWorkbook;
    window.OptionExcel = api;
  }
})(function () {
  "use strict";

  var ALLOWED_INPUT_ROWS = [100, 200, 300, 500];
  var DATA_START_ROW = 8;
  var TOKEN_WIDTH = 20;
  var ILLEGAL_FILENAME = /[\\/:*?"<>|]/g;

  function trimText(value) {
    return String(value == null ? "" : value).trim();
  }

  function normalizePriceText(value) {
    return trimText(value)
      .replace(/[０-９]/g, function (ch) {
        return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
      })
      .replace(/．/g, ".")
      .replace(/，/g, ",")
      .replace(/円/g, "")
      .replace(/,/g, "")
      .replace(/\s+/g, "");
  }

  function parsePrice(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : NaN;
    }
    var text = normalizePriceText(value);
    if (!/^\d+(\.\d+)?$/.test(text)) {
      return NaN;
    }
    return Number(text);
  }

  function sanitizeFolderName(value) {
    var base = trimText(value)
      .replace(ILLEGAL_FILENAME, "_")
      .replace(/[\u0000-\u001f]/g, "")
      .replace(/[. ]+$/g, "")
      .trim();
    if (base === "." || base === "..") {
      return "";
    }
    return base;
  }

  function sanitizeFilename(specName) {
    var base = sanitizeFolderName(specName);
    if (!base) {
      base = "オプション";
    }
    return base + "_オプション.xlsx";
  }

  function validateOptionInput(input) {
    var source = input || {};
    var errors = [];
    var specName = trimText(source.specName);
    var inputRows = Number(source.inputRows);
    var rawRows = Array.isArray(source.options) ? source.options : [];
    var options = [];
    var seen = Object.create(null);

    if (!specName) {
      errors.push("図番番号を入力してください。");
    }

    if (ALLOWED_INPUT_ROWS.indexOf(inputRows) === -1) {
      errors.push("入力行数は 100、200、300、500 のいずれかにしてください。");
    }

    var baseBlank =
      source.basePrice == null ||
      (typeof source.basePrice === "string" && trimText(source.basePrice) === "");
    var basePrice = NaN;
    if (baseBlank) {
      errors.push("本体価格を入力してください。");
    } else {
      basePrice = parsePrice(source.basePrice);
      if (!Number.isFinite(basePrice) || basePrice < 0) {
        errors.push("本体価格は 0 以上の数値にしてください。");
        basePrice = NaN;
      }
    }

    rawRows.forEach(function (row, index) {
      var rowNumber = index + 1;
      var item = row || {};
      var no = trimText(item.no);
      var name = trimText(item.name);
      var priceBlank =
        item.price == null ||
        (typeof item.price === "string" && trimText(item.price) === "");
      var fullyEmpty = !no && !name && priceBlank;

      if (fullyEmpty) {
        return;
      }

      var rowOk = true;

      if (!no) {
        errors.push(rowNumber + "行目: オプションNoを入力してください。");
        rowOk = false;
      }

      if (!name) {
        errors.push(rowNumber + "行目: 品名を入力してください。");
        rowOk = false;
      }

      var price = NaN;
      if (priceBlank) {
        errors.push(rowNumber + "行目: 値段を入力してください。");
        rowOk = false;
      } else {
        price = parsePrice(item.price);
        if (!Number.isFinite(price) || price < 0) {
          errors.push(rowNumber + "行目: 値段は 0 以上の数値にしてください。");
          rowOk = false;
        }
      }

      if (no) {
        if (Object.prototype.hasOwnProperty.call(seen, no)) {
          errors.push(
            rowNumber +
              "行目: オプションNo「" +
              no +
              "」が" +
              seen[no] +
              "行目と重複しています。"
          );
          rowOk = false;
        } else {
          seen[no] = rowNumber;
        }
      }

      if (rowOk) {
        options.push({ no: no, name: name, price: price });
      }
    });

    if (options.length === 0) {
      var hasRowError = errors.some(function (message) {
        return message.indexOf("行目") !== -1;
      });
      if (!hasRowError) {
        errors.push("オプションを1行以上入力してください。");
      }
    }

    return {
      ok: errors.length === 0,
      errors: errors,
      specName: specName,
      inputRows: inputRows,
      basePrice: basePrice,
      options: options,
    };
  }

  function splitOptionCodes(text, maxRows) {
    var limit = Number(maxRows);
    if (ALLOWED_INPUT_ROWS.indexOf(limit) === -1) {
      limit = 300;
    }
    if (trimText(text) === "") {
      return [];
    }
    var normalized = String(text == null ? "" : text)
      .replace(/、/g, ",")
      .replace(/，/g, ",");
    var parts = normalized.split(",");
    var tokens = [];
    var i;
    for (i = 0; i < parts.length && tokens.length < limit; i += 1) {
      tokens.push(trimText(parts[i]));
    }
    return tokens;
  }

  function previewOptionSelection(source) {
    var input = source || {};
    var maxRows = Number(input.inputRows);
    if (ALLOWED_INPUT_ROWS.indexOf(maxRows) === -1) {
      maxRows = 300;
    }
    var baseBlank =
      input.basePrice == null ||
      (typeof input.basePrice === "string" && trimText(input.basePrice) === "");
    var basePrice = baseBlank ? 0 : parsePrice(input.basePrice);
    if (!Number.isFinite(basePrice) || basePrice < 0) {
      basePrice = 0;
    }

    var master = Object.create(null);
    var rawRows = Array.isArray(input.options) ? input.options : [];
    rawRows.forEach(function (row) {
      var item = row || {};
      var no = trimText(item.no);
      var name = trimText(item.name);
      var price = parsePrice(item.price);
      if (!no || !name || !Number.isFinite(price) || price < 0) {
        return;
      }
      if (!Object.prototype.hasOwnProperty.call(master, no)) {
        master[no] = { no: no, name: name, price: price };
      }
    });

    var codes = splitOptionCodes(input.selectedNos, maxRows);
    var items = [];
    var names = [];
    var optionTotal = 0;
    codes.forEach(function (code) {
      if (!code) {
        return;
      }
      if (Object.prototype.hasOwnProperty.call(master, code)) {
        var found = master[code];
        items.push({
          no: found.no,
          name: found.name,
          price: found.price,
          missing: false,
        });
        names.push(found.name);
        optionTotal += found.price;
      } else {
        items.push({
          no: code,
          name: "該当なし",
          price: null,
          missing: true,
        });
        names.push("該当なし");
      }
    });

    var used = items.filter(function (item) {
      return !item.missing;
    });
    var selectedNosText = used
      .map(function (item) {
        return item.no;
      })
      .join(",");
    var selectedWithNames = used
      .map(function (item) {
        return item.no + " " + item.name;
      })
      .join("、");

    return {
      codes: codes.filter(function (code) {
        return code !== "";
      }),
      items: items,
      itemList: names.join(","),
      selectedNosText: selectedNosText,
      selectedWithNames: selectedWithNames,
      optionTotal: optionTotal,
      grandTotal: basePrice + optionTotal,
      basePrice: basePrice,
      masterCount: Object.keys(master).length,
    };
  }

  function priceNumFmt(options, basePrice) {
    var hasDecimal =
      !Number.isInteger(basePrice) ||
      options.some(function (option) {
        return !Number.isInteger(option.price);
      });
    return hasDecimal ? "#,##0.00" : "#,##0";
  }

  function lookupFormula(row, masterLast, column, fallback) {
    var ifError = fallback === "" ? '""' : '"' + fallback + '"';
    return (
      'IF(A' +
      row +
      '="","",IFERROR(INDEX(マスタ!$' +
      column +
      "$2:$" +
      column +
      "$" +
      masterLast +
      ",MATCH(A" +
      row +
      '&"",マスタ!$A$2:$A$' +
      masterLast +
      ",0))," +
      ifError +
      "))"
    );
  }

  function splitFormula(row) {
    var start = (row - DATA_START_ROW) * TOKEN_WIDTH + 1;
    var normalized = 'SUBSTITUTE(SUBSTITUTE(B3&"","、",","),"，",",")';
    return (
      'IF(B3="","",TRIM(MID(SUBSTITUTE(' +
      normalized +
      '&",",",",REPT(" ",' +
      TOKEN_WIDTH +
      "))," +
      start +
      "," +
      TOKEN_WIDTH +
      ")))"
    );
  }

  function joinFormula(row) {
    var source = "B" + row;
    if (row === DATA_START_ROW) {
      return "IF(" + source + '="","",' + source + ")";
    }
    var prev = row - 1;
    return (
      "IF(" +
      source +
      '="",E' +
      prev +
      ",IF(E" +
      prev +
      '="",' +
      source +
      ",E" +
      prev +
      '&","&' +
      source +
      "))"
    );
  }

  var STYLE = {
    header: 1,
    code: 2,
    money: 3,
    moneyDecimal: 4,
    label: 5,
    plain: 6,
    totalLabel: 7,
    totalMoney: 8,
    totalMoneyDecimal: 9,
    nameFormula: 10,
    moneyLeft: 11,
    moneyDecimalLeft: 12,
    totalMoneyLeft: 13,
    totalMoneyDecimalLeft: 14,
    input: 15,
  };

  function xmlText(value) {
    return String(value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function xmlAttr(value) {
    return xmlText(value).replace(/"/g, "&quot;");
  }

  function inlineCell(ref, style, text) {
    return (
      '<c r="' +
      ref +
      '" s="' +
      style +
      '" t="inlineStr"><is><t>' +
      xmlText(text) +
      "</t></is></c>"
    );
  }

  function formulaCell(ref, style, formula) {
    var styleAttr = style ? ' s="' + style + '"' : "";
    return '<c r="' + ref + '"' + styleAttr + "><f>" + xmlText(formula) + "</f></c>";
  }

  function numberCell(ref, style, value) {
    return '<c r="' + ref + '" s="' + style + '"><v>' + value + "</v></c>";
  }

  function emptyCell(ref, style) {
    return '<c r="' + ref + '" s="' + style + '"/>';
  }

  function stylesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="3">' +
      '<font><sz val="11"/><color rgb="FF1C1915"/><name val="Yu Gothic"/><family val="2"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FF1C3330"/><name val="Yu Gothic"/><family val="2"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FF1C1915"/><name val="Yu Gothic"/><family val="2"/></font>' +
      "</fonts>" +
      '<fills count="4">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFDCE8E3"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4D6"/><bgColor indexed="64"/></patternFill></fill>' +
      "</fills>" +
      '<borders count="2">' +
      "<border><left/><right/><top/><bottom/><diagonal/></border>" +
      '<border><left style="thin"><color rgb="FFD0D5DC"/></left><right style="thin"><color rgb="FFD0D5DC"/></right><top style="thin"><color rgb="FFD0D5DC"/></top><bottom style="thin"><color rgb="FFD0D5DC"/></bottom><diagonal/></border>' +
      "</borders>" +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="16">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="3" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="4" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="3" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="4" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="49" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
      "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      "</styleSheet>"
    );
  }

  function masterSheetXml(options, moneyStyle) {
    var masterLast = options.length + 1;
    var rows = [
      '<row r="1">' +
        inlineCell("A1", STYLE.header, "オプションNo") +
        inlineCell("B1", STYLE.header, "品名") +
        inlineCell("C1", STYLE.header, "値段") +
        "</row>",
    ];

    options.forEach(function (option, index) {
      var rowNumber = index + 2;
      rows.push(
        '<row r="' +
          rowNumber +
          '">' +
          inlineCell("A" + rowNumber, STYLE.code, option.no) +
          inlineCell("B" + rowNumber, STYLE.nameFormula, option.name) +
          numberCell("C" + rowNumber, moneyStyle, option.price) +
          "</row>"
      );
    });

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:C' +
      masterLast +
      '"/>' +
      "<sheetViews><sheetView workbookViewId=\"0\">" +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      "</sheetView></sheetViews>" +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      "<cols>" +
      '<col min="1" max="1" width="18" customWidth="1"/>' +
      '<col min="2" max="2" width="36" customWidth="1"/>' +
      '<col min="3" max="3" width="16" customWidth="1"/>' +
      "</cols>" +
      "<sheetData>" +
      rows.join("") +
      "</sheetData></worksheet>"
    );
  }

  function inputSheetXml(specName, inputRows, masterLast, basePrice, moneyStyle, totalStyle) {
    var lastInputRow = DATA_START_ROW + inputRows - 1;
    var optionTotalRow = lastInputRow + 1;
    var grandTotalRow = lastInputRow + 2;
    var sumFormula = "SUM(C" + DATA_START_ROW + ":C" + lastInputRow + ")";
    var headerMoneyStyle = moneyStyle === STYLE.money ? STYLE.moneyLeft : STYLE.moneyDecimalLeft;
    var headerTotalStyle =
      totalStyle === STYLE.totalMoney ? STYLE.totalMoneyLeft : STYLE.totalMoneyDecimalLeft;
    var rows = [
      '<row r="1">' +
        inlineCell("A1", STYLE.label, "図番番号") +
        inlineCell("B1", STYLE.plain, specName) +
        "</row>",
      '<row r="2">' +
        inlineCell("A2", STYLE.label, "本体価格") +
        numberCell("B2", headerMoneyStyle, basePrice) +
        "</row>",
      '<row r="3">' +
        inlineCell("A3", STYLE.label, "オプションNo") +
        emptyCell("B3", STYLE.input) +
        inlineCell("C3", STYLE.plain, "例: 1,2,3") +
        "</row>",
      '<row r="4">' +
        inlineCell("A4", STYLE.label, "品目リスト") +
        formulaCell("B4", STYLE.plain, "E" + lastInputRow) +
        "</row>",
      '<row r="5">' +
        inlineCell("A5", STYLE.totalLabel, "オプション合計") +
        formulaCell("B5", headerTotalStyle, sumFormula) +
        "</row>",
      '<row r="6">' +
        inlineCell("A6", STYLE.totalLabel, "総合計") +
        formulaCell("B6", headerTotalStyle, "B2+B5") +
        "</row>",
      '<row r="7">' +
        inlineCell("A7", STYLE.header, "オプションNo") +
        inlineCell("B7", STYLE.header, "品名") +
        inlineCell("C7", STYLE.header, "値段") +
        "</row>",
    ];

    for (var row = DATA_START_ROW; row <= lastInputRow; row += 1) {
      rows.push(
        '<row r="' +
          row +
          '">' +
          formulaCell("A" + row, STYLE.code, splitFormula(row)) +
          formulaCell("B" + row, STYLE.nameFormula, lookupFormula(row, masterLast, "B", "該当なし")) +
          formulaCell("C" + row, moneyStyle, lookupFormula(row, masterLast, "C", "")) +
          formulaCell("E" + row, 0, joinFormula(row)) +
          "</row>"
      );
    }

    rows.push(
      '<row r="' +
        optionTotalRow +
        '">' +
        inlineCell("A" + optionTotalRow, STYLE.totalLabel, "オプション合計") +
        emptyCell("B" + optionTotalRow, STYLE.totalLabel) +
        formulaCell("C" + optionTotalRow, totalStyle, "B5") +
        "</row>"
    );
    rows.push(
      '<row r="' +
        grandTotalRow +
        '">' +
        inlineCell("A" + grandTotalRow, STYLE.totalLabel, "総合計") +
        emptyCell("B" + grandTotalRow, STYLE.totalLabel) +
        formulaCell("C" + grandTotalRow, totalStyle, "B6") +
        "</row>"
    );

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:E' +
      grandTotalRow +
      '"/>' +
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' +
      '<pane ySplit="7" topLeftCell="A8" activePane="topLeft" state="frozen"/>' +
      '<selection pane="topLeft" activeCell="B3" sqref="B3"/>' +
      "</sheetView></sheetViews>" +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      "<cols>" +
      '<col min="1" max="1" width="24" customWidth="1"/>' +
      '<col min="2" max="2" width="62" customWidth="1"/>' +
      '<col min="3" max="3" width="16" customWidth="1"/>' +
      '<col min="4" max="4" width="3" customWidth="1"/>' +
      '<col min="5" max="5" width="28" customWidth="1" hidden="1"/>' +
      "</cols>" +
      "<sheetData>" +
      rows.join("") +
      "</sheetData></worksheet>"
    );
  }

  function workbookXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<workbookViews><workbookView activeTab="1"/></workbookViews>' +
      "<sheets>" +
      '<sheet name="マスタ" sheetId="1" r:id="rId1"/>' +
      '<sheet name="入力" sheetId="2" r:id="rId2"/>' +
      "</sheets>" +
      '<calcPr calcMode="auto" fullCalcOnLoad="1"/>' +
      "</workbook>"
    );
  }

  function contentTypesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      "</Types>"
    );
  }

  function rootRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      "</Relationships>"
    );
  }

  function workbookRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>"
    );
  }

  function concatBytes(parts) {
    var size = 0;
    var index;
    for (index = 0; index < parts.length; index += 1) {
      size += parts[index].length;
    }
    var out = new Uint8Array(size);
    var offset = 0;
    for (index = 0; index < parts.length; index += 1) {
      out.set(parts[index], offset);
      offset += parts[index].length;
    }
    return out;
  }

  function u16(value) {
    return new Uint8Array([value & 255, (value >> 8) & 255]);
  }

  function u32(value) {
    var number = value >>> 0;
    return new Uint8Array([
      number & 255,
      (number >>> 8) & 255,
      (number >>> 16) & 255,
      (number >>> 24) & 255,
    ]);
  }

  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    var n;
    var k;
    var c;
    for (n = 0; n < 256; n += 1) {
      c = n;
      for (k = 0; k < 8; k += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var crc = 0xffffffff;
    var i;
    for (i = 0; i < bytes.length; i += 1) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function dosDateTime(date) {
    var safe = date || new Date();
    var year = safe.getFullYear();
    if (year < 1980) {
      year = 1980;
    }
    return {
      time: (safe.getHours() << 11) | (safe.getMinutes() << 5) | Math.floor(safe.getSeconds() / 2),
      date: ((year - 1980) << 9) | ((safe.getMonth() + 1) << 5) | safe.getDate(),
    };
  }

  function copyBytes(src) {
    var out = new Uint8Array(src.length);
    out.set(src);
    return out;
  }

  function decodeText(bytes) {
    if (!bytes || !bytes.length) {
      return "";
    }
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
      return new TextDecoder("utf-16le").decode(bytes);
    }
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
      return new TextDecoder("utf-16be").decode(bytes);
    }
    return new TextDecoder("utf-8").decode(bytes);
  }

  function encodeText(text, encoding) {
    if (encoding === "utf-16le") {
      var i;
      var out = new Uint8Array(2 + String(text).length * 2);
      out[0] = 0xff;
      out[1] = 0xfe;
      for (i = 0; i < text.length; i += 1) {
        var code = text.charCodeAt(i);
        out[2 + i * 2] = code & 0xff;
        out[3 + i * 2] = (code >> 8) & 0xff;
      }
      return out;
    }
    return new TextEncoder().encode(text);
  }

  function sniffWorkbook(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    if (bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b) {
      return "xlsx";
    }
    if (
      bytes.length >= 8 &&
      bytes[0] === 0xd0 &&
      bytes[1] === 0xcf &&
      bytes[2] === 0x11 &&
      bytes[3] === 0xe0
    ) {
      return "xls";
    }
    var head = decodeText(bytes.subarray(0, Math.min(bytes.length, 1200)));
    if (
      /urn:schemas-microsoft-com:office:spreadsheet/i.test(head) ||
      /mso-application[^>]*Excel\.Sheet/i.test(head)
    ) {
      return "xmlss";
    }
    return "unknown";
  }

  function buildZip(files) {
    var encoder = new TextEncoder();
    var stamp = dosDateTime(new Date(2026, 9, 11, 12, 0, 0));
    var locals = [];
    var centrals = [];
    var offset = 0;

    files.forEach(function (file) {
      var nameBytes = file.nameBytes ? file.nameBytes : encoder.encode(file.name);
      var data = file.bytes ? file.bytes : encoder.encode(file.text);
      var crc = crc32(data);
      var local = concatBytes([
        u32(0x04034b50),
        u16(20),
        u16(0),
        u16(0),
        u16(stamp.time),
        u16(stamp.date),
        u32(crc),
        u32(data.length),
        u32(data.length),
        u16(nameBytes.length),
        u16(0),
        nameBytes,
        data,
      ]);
      var central = concatBytes([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0),
        u16(0),
        u16(stamp.time),
        u16(stamp.date),
        u32(crc),
        u32(data.length),
        u32(data.length),
        u16(nameBytes.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        nameBytes,
      ]);
      offset += local.length;
      locals.push(local);
      centrals.push(central);
    });

    var centralBlob = concatBytes(centrals);
    var eocd = concatBytes([
      u32(0x06054b50),
      u16(0),
      u16(0),
      u16(files.length),
      u16(files.length),
      u32(centralBlob.length),
      u32(offset),
      u16(0),
    ]);
    return concatBytes(locals.concat([centralBlob, eocd]));
  }

  function buildOptionWorkbook(first, second) {
    var input = first;
    if (
      arguments.length > 1 &&
      (first == null || typeof first === "function" || typeof first.Workbook === "function")
    ) {
      input = second;
    }

    var checked = validateOptionInput(input);
    if (!checked.ok) {
      throw new Error(checked.errors.join("\n"));
    }

    var specName = checked.specName;
    var inputRows = checked.inputRows;
    var options = checked.options;
    var masterLast = options.length + 1;
    var decimal = priceNumFmt(options, checked.basePrice) === "#,##0.00";
    var moneyStyle = decimal ? STYLE.moneyDecimal : STYLE.money;
    var totalStyle = decimal ? STYLE.totalMoneyDecimal : STYLE.totalMoney;
    var bytes = buildZip([
      { name: "[Content_Types].xml", text: contentTypesXml() },
      { name: "_rels/.rels", text: rootRelsXml() },
      { name: "xl/workbook.xml", text: workbookXml() },
      { name: "xl/_rels/workbook.xml.rels", text: workbookRelsXml() },
      { name: "xl/styles.xml", text: stylesXml() },
      { name: "xl/worksheets/sheet1.xml", text: masterSheetXml(options, moneyStyle) },
      {
        name: "xl/worksheets/sheet2.xml",
        text: inputSheetXml(
          specName,
          inputRows,
          masterLast,
          checked.basePrice,
          moneyStyle,
          totalStyle
        ),
      },
    ]);

    return {
      xlsx: {
        writeBuffer: function () {
          return Promise.resolve(bytes);
        },
      },
    };
  }

  if (crc32(new TextEncoder().encode("123456789")) !== 0xcbf43926) {
    throw new Error("CRC32 self-check failed");
  }

  function decodeXml(value) {
    return String(value == null ? "" : value)
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (_, num) {
        return String.fromCharCode(Number(num));
      })
      .replace(/&amp;/g, "&");
  }

  function inflateRaw(bytes) {
    if (typeof DecompressionStream !== "function") {
      return Promise.reject(new Error("このブラウザではExcelを解けません。Edge または Chrome を使ってください。"));
    }
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Response(stream).arrayBuffer().then(function (buffer) {
      return new Uint8Array(buffer);
    });
  }

  function findEocd(bytes) {
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var min = Math.max(0, bytes.length - 65557);
    var i;
    for (i = bytes.length - 22; i >= min; i -= 1) {
      if (view.getUint32(i, true) === 0x06054b50) {
        return i;
      }
    }
    return -1;
  }

  function unzipZip(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var eocd = findEocd(bytes);
    if (eocd < 0) {
      return Promise.reject(new Error("Excelファイルの形式が正しくありません。"));
    }
    var count = view.getUint16(eocd + 10, true);
    var central = view.getUint32(eocd + 16, true);
    var jobs = [];
    var i;
    var offset = central;
    for (i = 0; i < count; i += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) {
        return Promise.reject(new Error("Excelファイルを分解できませんでした。"));
      }
      var method = view.getUint16(offset + 10, true);
      var compressedSize = view.getUint32(offset + 20, true);
      var nameLength = view.getUint16(offset + 28, true);
      var extraLength = view.getUint16(offset + 30, true);
      var commentLength = view.getUint16(offset + 32, true);
      var localOffset = view.getUint32(offset + 42, true);
      var name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
      var localNameLength = view.getUint16(localOffset + 26, true);
      var localExtraLength = view.getUint16(localOffset + 28, true);
      var dataStart = localOffset + 30 + localNameLength + localExtraLength;
      var payload = bytes.subarray(dataStart, dataStart + compressedSize);
      jobs.push({ name: name, method: method, payload: payload });
      offset += 46 + nameLength + extraLength + commentLength;
    }

    return Promise.all(
      jobs.map(function (job) {
        var raw =
          job.method === 0
            ? Promise.resolve(job.payload)
            : job.method === 8
              ? inflateRaw(job.payload)
              : Promise.reject(new Error("未対応のExcel圧縮形式です。"));
        return raw.then(function (data) {
          return [job.name.replace(/^\/+/, ""), copyBytes(data)];
        });
      })
    ).then(function (entries) {
      var files = Object.create(null);
      entries.forEach(function (entry) {
        files[entry[0]] = entry[1];
      });
      return files;
    });
  }

  function unzipXlsx(buffer) {
    return unzipZip(buffer).then(function (files) {
      var textFiles = Object.create(null);
      Object.keys(files).forEach(function (name) {
        if (/\.(xml|rels)$/i.test(name) || name === "[Content_Types].xml") {
          textFiles[name] = decodeText(files[name]);
        }
      });
      return textFiles;
    });
  }

  function xmlFromZip(files, name) {
    return files[name] ? decodeText(files[name]) : "";
  }

  function putXmlInZip(files, name, xml) {
    files[name] = encodeText(xml, "utf-8");
  }

  function parseSharedStrings(xml) {
    var strings = [];
    if (!xml) {
      return strings;
    }
    var siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    var block;
    while ((block = siRe.exec(xml))) {
      var texts = [];
      var tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      var tBlock;
      while ((tBlock = tRe.exec(block[1]))) {
        texts.push(decodeXml(tBlock[1]));
      }
      strings.push(texts.join(""));
    }
    return strings;
  }

  function cellRefParts(ref) {
    var match = /^([A-Z]+)(\d+)$/i.exec(String(ref || ""));
    if (!match) {
      return null;
    }
    return { col: match[1].toUpperCase(), row: Number(match[2]) };
  }

  function colLettersToIndex(letters) {
    var n = 0;
    var i;
    var text = String(letters || "").toUpperCase();
    for (i = 0; i < text.length; i += 1) {
      n = n * 26 + (text.charCodeAt(i) - 64);
    }
    return n;
  }

  function indexToCol(index) {
    var n = Number(index);
    var text = "";
    while (n > 0) {
      var rem = (n - 1) % 26;
      text = String.fromCharCode(65 + rem) + text;
      n = Math.floor((n - 1) / 26);
    }
    return text;
  }

  function normalizeImportedNo(value) {
    var text = trimText(value);
    if (!text) {
      return "";
    }
    if (/^\d+\.0+$/.test(text)) {
      return text.replace(/\.0+$/, "");
    }
    return text;
  }

  function sheetCellValue(cellXml, shared) {
    var typeMatch = /\bt="([^"]+)"/.exec(cellXml);
    var type = typeMatch ? typeMatch[1] : "";
    var inline = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(cellXml);
    if (type === "inlineStr" || (!type && inline && cellXml.indexOf("<is") !== -1)) {
      return inline ? decodeXml(inline[1]) : "";
    }
    var valueMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cellXml);
    if (!valueMatch) {
      return "";
    }
    var raw = decodeXml(valueMatch[1]);
    if (type === "s") {
      var index = Number(raw);
      return shared[index] == null ? "" : String(shared[index]);
    }
    return raw;
  }

  function parseSheetGrid(xml, shared) {
    var grid = Object.create(null);
    if (!xml) {
      return grid;
    }
    var rowRe = /<(?:[\w]+:)?row\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?row>/g;
    var rowMatch;
    while ((rowMatch = rowRe.exec(xml))) {
      var rowAttr = /\br="(\d+)"/.exec(rowMatch[1]);
      var rowNum = rowAttr ? Number(rowAttr[1]) : 0;
      var colIndex = 0;
      var cellRe = /<(?:[\w]+:)?c\b([^/>]*)(\/>|>([\s\S]*?)<\/(?:[\w]+:)?c>)/g;
      var cellMatch;
      while ((cellMatch = cellRe.exec(rowMatch[2]))) {
        var attrs = cellMatch[1];
        var inner = cellMatch[2] === "/>" ? "" : cellMatch[3] || "";
        var refMatch = /\br="([^"]+)"/.exec(attrs);
        var parts = refMatch ? cellRefParts(refMatch[1]) : null;
        if (parts) {
          rowNum = parts.row;
          colIndex = colLettersToIndex(parts.col);
        } else {
          colIndex += 1;
          if (!rowNum) {
            continue;
          }
          parts = { col: indexToCol(colIndex), row: rowNum };
        }
        if (!grid[parts.row]) {
          grid[parts.row] = Object.create(null);
        }
        grid[parts.row][parts.col] = sheetCellValue(attrs + ">" + inner, shared);
      }
    }
    if (Object.keys(grid).length === 0) {
      var cellReAll = /<(?:[\w]+:)?c\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?c>/g;
      var match;
      while ((match = cellReAll.exec(xml))) {
        var ref = /\br="([^"]+)"/.exec(match[1]);
        var cellParts = ref ? cellRefParts(ref[1]) : null;
        if (!cellParts) {
          continue;
        }
        if (!grid[cellParts.row]) {
          grid[cellParts.row] = Object.create(null);
        }
        grid[cellParts.row][cellParts.col] = sheetCellValue(match[1] + ">" + match[2], shared);
      }
    }
    return grid;
  }

  function looksLikeHeader(a, b, c) {
    var text = (a + " " + b + " " + c).replace(/\s+/g, "");
    return /ナンバー|オプションNo|品名|品目|売価|値段|^No$/i.test(text);
  }

  function optionsFromGrid(grid) {
    var rows = Object.keys(grid)
      .map(Number)
      .filter(function (n) {
        return n > 0;
      })
      .sort(function (a, b) {
        return a - b;
      });
    var options = [];
    var started = false;
    rows.forEach(function (rowNum) {
      var row = grid[rowNum] || {};
      var no = trimText(row.A);
      var name = trimText(row.B);
      var price = trimText(row.C);
      if (!started && looksLikeHeader(no, name, price)) {
        return;
      }
      if (!no && !name && !price) {
        return;
      }
      if (no === "仕様名" || no === "図番番号" || no === "本体価格" || no === "オプションNo" || no === "品目リスト") {
        return;
      }
      started = true;
      options.push({ no: normalizeImportedNo(no), name: name, price: price });
    });
    return options;
  }

  function isCircleMark(value) {
    var text = trimText(value).replace(/\s+/g, "");
    if (!text) {
      return false;
    }
    return /^(○|●|◯|〇|◎|◉|⦿|✔|✓|☑|レ|丸)$/.test(text);
  }

  function optionNoFromText(value) {
    var text = normalizeImportedNo(value);
    if (/^\d+$/.test(text)) {
      return text;
    }
    var mixed = /^(\d+)\s*[○●◯〇◎◉⦿✔✓☑レ丸]$/.exec(trimText(value).replace(/\s+/g, " "));
    if (mixed) {
      return mixed[1];
    }
    return "";
  }

  function gridCell(grid, rowNum, colIndex) {
    var row = grid[rowNum] || {};
    return row[indexToCol(colIndex)] || "";
  }

  function usedColumnIndexes(grid) {
    var seen = Object.create(null);
    Object.keys(grid).forEach(function (rowNum) {
      Object.keys(grid[rowNum] || {}).forEach(function (col) {
        if (trimText(grid[rowNum][col])) {
          seen[colLettersToIndex(col)] = true;
        }
      });
    });
    return Object.keys(seen)
      .map(Number)
      .filter(function (n) {
        return n > 0;
      })
      .sort(function (a, b) {
        return a - b;
      });
  }

  function isMarkHeaderText(value) {
    return /^(番号|品名|品目|値段|売価|選択|該当|ナンバー|No)$/i.test(trimText(value));
  }

  function collectRowCells(grid, rowNum, start, limit) {
    var cells = [];
    var col;
    for (col = start; col <= limit; col += 1) {
      var value = trimText(gridCell(grid, rowNum, col));
      if (value) {
        cells.push({ col: col, value: value });
      }
    }
    return cells;
  }

  function rowHasCircle(grid, rowNum, start, limit) {
    var col;
    for (col = start; col <= limit; col += 1) {
      if (isCircleMark(gridCell(grid, rowNum, col))) {
        return true;
      }
    }
    return false;
  }

  function columnBlocks(grid) {
    var cols = usedColumnIndexes(grid);
    var blocks = [];
    if (cols.length === 0) {
      return blocks;
    }
    var start = cols[0];
    var prev = cols[0];
    var i;
    for (i = 1; i < cols.length; i += 1) {
      if (cols[i] - prev >= 2) {
        blocks.push({ start: start, end: prev });
        start = cols[i];
      }
      prev = cols[i];
    }
    blocks.push({ start: start, end: prev });
    return blocks;
  }

  function headerKind(value) {
    var text = trimText(value).replace(/\s+/g, "");
    if (/^(番号|ナンバー|オプションNo|No)$/i.test(text)) {
      return "no";
    }
    if (/^(品名|品目)$/.test(text)) {
      return "name";
    }
    if (/^(値段|売価|価格)$/.test(text)) {
      return "price";
    }
    if (/^(選択|該当)$/.test(text)) {
      return "mark";
    }
    return "";
  }

  function isNumericCell(value) {
    return optionNoFromText(value) !== "" && /^[\d,.\s]+$/.test(trimText(value));
  }

  function nameFromCell(value) {
    var text = trimText(value);
    if (!text || isCircleMark(text) || headerKind(text) || isMarkHeaderText(text)) {
      return "";
    }
    if (isNumericCell(text)) {
      return "";
    }
    return text;
  }

  function blockColumns(grid, start, end) {
    var noCol = start;
    var nameCol = 0;
    var priceCol = 0;
    var headerRows = Object.create(null);
    var rowNum;
    var col;
    for (rowNum = 1; rowNum <= 12; rowNum += 1) {
      var found = false;
      for (col = start; col <= end; col += 1) {
        var kind = headerKind(gridCell(grid, rowNum, col));
        if (!kind) {
          continue;
        }
        found = true;
        if (kind === "no") {
          noCol = col;
        } else if (kind === "name" && !nameCol) {
          nameCol = col;
        } else if (kind === "price" && !priceCol) {
          priceCol = col;
        }
      }
      if (found) {
        headerRows[rowNum] = true;
      }
    }
    if (!nameCol) {
      for (col = noCol + 1; col <= end; col += 1) {
        if (col !== priceCol) {
          nameCol = col;
          break;
        }
      }
    }
    if (!priceCol) {
      for (col = noCol + 1; col <= end; col += 1) {
        if (col === nameCol) {
          continue;
        }
        var numeric = false;
        for (rowNum = 1; rowNum <= 40; rowNum += 1) {
          if (headerRows[rowNum]) {
            continue;
          }
          if (isNumericCell(gridCell(grid, rowNum, col)) && !isCircleMark(gridCell(grid, rowNum, col))) {
            numeric = true;
            break;
          }
        }
        if (numeric) {
          priceCol = col;
          break;
        }
      }
    }
    return { noCol: noCol, nameCol: nameCol, priceCol: priceCol, headerRows: headerRows };
  }

  function optionsInColumnRange(grid, start, limit) {
    var columns = blockColumns(grid, start, limit);
    var rows = Object.keys(grid)
      .map(Number)
      .filter(function (n) {
        return n > 0;
      })
      .sort(function (a, b) {
        return a - b;
      });
    var options = [];
    var nos = [];
    var skip = Object.create(null);
    var seen = Object.create(null);
    rows.forEach(function (rowNum) {
      if (skip[rowNum] || columns.headerRows[rowNum]) {
        return;
      }
      var no = optionNoFromText(gridCell(grid, rowNum, columns.noCol));
      if (!no) {
        return;
      }
      var name = columns.nameCol ? nameFromCell(gridCell(grid, rowNum, columns.nameCol)) : "";
      var price = "";
      if (columns.priceCol && columns.priceCol !== columns.nameCol) {
        var priceValue = parsePrice(gridCell(grid, rowNum, columns.priceCol));
        if (Number.isFinite(priceValue)) {
          price = String(priceValue);
        }
      }
      var marked = rowHasCircle(grid, rowNum, start, limit);
      var nextNo = optionNoFromText(gridCell(grid, rowNum + 1, columns.noCol));
      if (!nextNo && !columns.headerRows[rowNum + 1]) {
        skip[rowNum + 1] = true;
        if (!name && columns.nameCol) {
          name = nameFromCell(gridCell(grid, rowNum + 1, columns.nameCol));
        }
        if (!price && columns.priceCol && columns.priceCol !== columns.nameCol) {
          var nextPrice = parsePrice(gridCell(grid, rowNum + 1, columns.priceCol));
          if (Number.isFinite(nextPrice)) {
            price = String(nextPrice);
          }
        }
        if (rowHasCircle(grid, rowNum + 1, start, limit)) {
          marked = true;
        }
      }
      if (seen[no]) {
        return;
      }
      seen[no] = true;
      options.push({ no: no, name: name, price: price });
      if (marked) {
        nos.push(no);
      }
    });
    return { nos: nos, options: options };
  }

  function leftOptionsFromGrid(grid) {
    var blocks = columnBlocks(grid);
    var options = [];
    var nos = [];
    var seen = Object.create(null);
    if (blocks.length === 0) {
      return { nos: nos, options: options };
    }
    blocks.forEach(function (block) {
      var part = optionsInColumnRange(grid, block.start, block.end);
      part.options.forEach(function (item) {
        if (seen[item.no]) {
          return;
        }
        seen[item.no] = true;
        options.push(item);
      });
      part.nos.forEach(function (no) {
        if (nos.indexOf(no) === -1) {
          nos.push(no);
        }
      });
    });
    return { nos: nos, options: options };
  }

  function markedNosFromGrid(grid) {
    return leftOptionsFromGrid(grid).nos;
  }

  function xmlssSheetGrid(sheetXml) {
    var grid = Object.create(null);
    var rowRe = elementRe("Row");
    var lastRow = 0;
    var rowMatch;
    while ((rowMatch = rowRe.exec(sheetXml || ""))) {
      var rowAttrs = rowMatch[1] || "";
      var inner = rowMatch[3] || "";
      var rowIdx = /\bss:Index="(\d+)"/.exec(rowAttrs);
      lastRow = rowIdx ? Number(rowIdx[1]) : lastRow + 1;
      var col = 0;
      var cellRe = elementRe("Cell");
      var cellMatch;
      while ((cellMatch = cellRe.exec(inner))) {
        var attrs = cellMatch[1] || "";
        var cellInner = cellMatch[3] || "";
        var colIdx = /\bss:Index="(\d+)"/.exec(attrs);
        col = colIdx ? Number(colIdx[1]) : col + 1;
        var data = /<Data\b[^>]*>([\s\S]*?)<\/Data>/.exec(cellInner);
        if (!grid[lastRow]) {
          grid[lastRow] = Object.create(null);
        }
        grid[lastRow][indexToCol(col)] = data ? decodeXml(data[1]) : "";
      }
    }
    return grid;
  }

  function pickBestMarked(list) {
    var best = { nos: [], options: [] };
    list.forEach(function (item) {
      if (item && item.nos && item.nos.length > best.nos.length) {
        best = item;
      }
    });
    return best;
  }

  function readMarkedOptionNos(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var kind = sniffWorkbook(bytes);
    if (kind === "xls") {
      return Promise.reject(
        new Error("丸付き仕様のExcelは .xlsx（またはXML形式の .xls）で選んでください。")
      );
    }
    if (kind === "xmlss") {
      try {
        var xml = decodeText(bytes);
        var sheets = xml.match(/<Worksheet\b[\s\S]*?<\/Worksheet>/g) || [];
        var lists = sheets.map(function (sheet) {
          return leftOptionsFromGrid(xmlssSheetGrid(sheet));
        });
        var parsed = pickBestMarked(lists);
        if (parsed.nos.length === 0) {
          throw new Error("番号の横に ○ が付いた行が見つかりませんでした。");
        }
        return Promise.resolve(parsed);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (kind !== "xlsx") {
      return Promise.reject(new Error("対応しているのは .xlsx と、XML形式のExcelです。"));
    }
    return unzipXlsx(bytes).then(function (files) {
      var shared = parseSharedStrings(files["xl/sharedStrings.xml"]);
      var sheets = parseWorkbookSheets(files);
      if (sheets.length === 0) {
        throw new Error("Excelにシートが見つかりませんでした。");
      }
      var lists = sheets.map(function (sheet) {
        if (!files[sheet.path]) {
          return { nos: [], options: [] };
        }
        return leftOptionsFromGrid(parseSheetGrid(files[sheet.path], shared));
      });
      var parsed = pickBestMarked(lists);
      if (parsed.nos.length === 0) {
        throw new Error("番号の横に ○ が付いた行が見つかりませんでした。");
      }
      return parsed;
    });
  }

  function parseWorkbookSheets(files) {
    var workbook = files["xl/workbook.xml"] || "";
    var rels = files["xl/_rels/workbook.xml.rels"] || "";
    var ridToTarget = Object.create(null);
    var relRe = /<Relationship\b([^>]*)\/?>/g;
    var rel;
    while ((rel = relRe.exec(rels))) {
      var idMatch = /\bId="([^"]+)"/.exec(rel[1]);
      var targetMatch = /\bTarget="([^"]+)"/.exec(rel[1]);
      if (idMatch && targetMatch) {
        var target = targetMatch[1].replace(/^\//, "");
        if (target.indexOf("xl/") !== 0) {
          target = "xl/" + target.replace(/^\.\//, "");
        }
        ridToTarget[idMatch[1]] = target;
      }
    }
    var sheets = [];
    var sheetRe = /<sheet\b([^>]*)\/?>/g;
    var sheet;
    while ((sheet = sheetRe.exec(workbook))) {
      var nameMatch = /\bname="([^"]+)"/.exec(sheet[1]);
      var ridMatch = /\br:id="([^"]+)"/.exec(sheet[1]);
      if (!nameMatch || !ridMatch || !ridToTarget[ridMatch[1]]) {
        continue;
      }
      sheets.push({
        name: decodeXml(nameMatch[1]),
        path: ridToTarget[ridMatch[1]],
      });
    }
    return sheets;
  }

  function readXlsxOptions(buffer) {
    return unzipXlsx(buffer).then(function (files) {
      var shared = parseSharedStrings(files["xl/sharedStrings.xml"]);
      var sheets = parseWorkbookSheets(files);
      if (sheets.length === 0) {
        throw new Error("Excelにシートが見つかりませんでした。");
      }
      var candidates = sheets.filter(function (sheet) {
        return sheet.name !== "入力";
      });
      if (candidates.length === 0) {
        candidates = sheets;
      }
      var namedMaster = candidates.filter(function (sheet) {
        return sheet.name === "マスタ";
      })[0];
      var options = [];
      if (namedMaster) {
        options = optionsFromGrid(parseSheetGrid(files[namedMaster.path], shared));
      }
      if (options.length === 0) {
        var best = [];
        candidates.forEach(function (sheet) {
          var rows = optionsFromGrid(parseSheetGrid(files[sheet.path], shared));
          if (rows.length > best.length) {
            best = rows;
          }
        });
        options = best;
      }
      if (options.length === 0) {
        throw new Error("A列ナンバー、B列品目、C列売価のデータが見つかりませんでした。");
      }

      var specName = "";
      var basePrice = "";
      var inputSheet = sheets.filter(function (sheet) {
        return sheet.name === "入力";
      })[0];
      if (inputSheet && files[inputSheet.path]) {
        var inputGrid = parseSheetGrid(files[inputSheet.path], shared);
        var row1 = inputGrid[1] || {};
        var row2 = inputGrid[2] || {};
        if (trimText(row1.A) === "仕様名" || trimText(row1.A) === "図番番号") {
          specName = trimText(row1.B);
        }
        if (trimText(row2.A) === "本体価格") {
          basePrice = trimText(row2.B);
        }
      }

      return {
        specName: specName,
        basePrice: basePrice,
        options: options,
      };
    });
  }

  function pickWriteSheet(sheets) {
    var named = sheets.filter(function (sheet) {
      return sheet.name === "計算" || sheet.name === "一覧";
    })[0];
    if (named) {
      return named;
    }
    var notMaster = sheets.filter(function (sheet) {
      return sheet.name !== "マスタ";
    });
    return notMaster[0] || sheets[0];
  }

  function calcNumberValue(value) {
    var n = parsePrice(value);
    return Number.isFinite(n) ? n : 0;
  }

  function cellStyleAttr(attrText) {
    var match = /\bs="([^"]*)"/.exec(attrText || "");
    return match ? ' s="' + match[1] + '"' : "";
  }

  function buildPatchedCell(ref, attrText, kind, value) {
    var style = cellStyleAttr(attrText);
    if (kind === "n") {
      return '<c r="' + ref + '"' + style + "><v>" + calcNumberValue(value) + "</v></c>";
    }
    return (
      '<c r="' +
      ref +
      '"' +
      style +
      ' t="inlineStr"><is><t>' +
      xmlText(value) +
      "</t></is></c>"
    );
  }

  var XML_ATTRS = '((?:[^>"/]|"[^"]*"|/(?!>))*)';

  function elementRe(tag) {
    return new RegExp("<" + tag + "\\b" + XML_ATTRS + "(/>|>([\\s\\S]*?)</" + tag + ">)", "g");
  }

  function parseRowCells(inner) {
    var cells = [];
    var re = elementRe("c");
    var match;
    while ((match = re.exec(inner || ""))) {
      var attrs = match[1] || "";
      var refMatch = /\br="([^"]+)"/.exec(attrs);
      var parts = cellRefParts(refMatch ? refMatch[1] : "");
      if (!parts) {
        cells.push({ col: 9999, ref: "", xml: match[0], attrs: attrs });
        continue;
      }
      cells.push({
        col: colLettersToIndex(parts.col),
        ref: parts.col + parts.row,
        xml: match[0],
        attrs: attrs,
      });
    }
    return cells;
  }

  function applyCellsToRowInner(inner, row, patches) {
    var cells = parseRowCells(inner);
    patches.forEach(function (patch) {
      var ref = patch.col + row;
      var col = colLettersToIndex(patch.col);
      var found = cells.filter(function (cell) {
        return cell.ref === ref;
      })[0];
      if (found && /<f\b[^>]*\bt="shared"[^>]*\bref="/.test(found.xml)) {
        throw new Error(
          ref +
            " は、ほかのセルと共有している数式の元になっています。上書きすると数式が壊れるため、書き込む行か列を変えてください。"
        );
      }
      var nextXml = buildPatchedCell(ref, found ? found.attrs : "", patch.kind, patch.value);
      if (found) {
        found.xml = nextXml;
        return;
      }
      cells.push({ col: col, ref: ref, xml: nextXml, attrs: "" });
    });
    cells.sort(function (a, b) {
      return a.col - b.col;
    });
    return cells
      .map(function (cell) {
        return cell.xml;
      })
      .join("");
  }

  function rowOpenTag(attrText, row) {
    var cleaned = String(attrText || "")
      .replace(/\br="[^"]*"/g, "")
      .replace(/\bspans="[^"]*"/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return '<row r="' + row + '"' + (cleaned ? " " + cleaned : "") + ">";
  }

  function insertRowXml(sheetDataInner, rowXml, row) {
    var rowRe = elementRe("row");
    var match;
    var lastIndex = 0;
    var chunks = [];
    var inserted = false;
    while ((match = rowRe.exec(sheetDataInner))) {
      var attrs = match[1] || "";
      var rMatch = /\br="(\d+)"/.exec(attrs);
      var r = rMatch ? Number(rMatch[1]) : 0;
      chunks.push(sheetDataInner.slice(lastIndex, match.index));
      if (!inserted && r > row) {
        chunks.push(rowXml);
        inserted = true;
      }
      chunks.push(match[0]);
      lastIndex = match.index + match[0].length;
    }
    chunks.push(sheetDataInner.slice(lastIndex));
    if (!inserted) {
      chunks.push(rowXml);
    }
    return chunks.join("");
  }

  function upsertSheetCells(sheetXml, row, patches) {
    var xml = sheetXml || "";
    var rowRe = elementRe("row");
    var match;
    var found = null;
    while ((match = rowRe.exec(xml))) {
      var attrs = match[1] || "";
      var rMatch = /\br="(\d+)"/.exec(attrs);
      if (rMatch && Number(rMatch[1]) === row) {
        found = match;
        break;
      }
    }
    var nextRow;
    if (found) {
      var inner = found[2] === "/>" ? "" : found[3] || "";
      nextRow = rowOpenTag(found[1], row) + applyCellsToRowInner(inner, row, patches) + "</row>";
      xml = xml.slice(0, found.index) + nextRow + xml.slice(found.index + found[0].length);
    } else {
      nextRow = rowOpenTag("", row) + applyCellsToRowInner("", row, patches) + "</row>";
      if (xml.indexOf("<sheetData/>") !== -1) {
        xml = xml.replace("<sheetData/>", "<sheetData>" + nextRow + "</sheetData>");
      } else if (/<sheetData\b[\s\S]*<\/sheetData>/.test(xml)) {
        xml = xml.replace(/<sheetData(\s[^>]*)?>([\s\S]*?)<\/sheetData>/, function (_, attrs, inner) {
          return "<sheetData" + (attrs || "") + ">" + insertRowXml(inner, nextRow, row) + "</sheetData>";
        });
      } else {
        throw new Error("Excelシートへ書き込めませんでした。");
      }
    }
    return xml;
  }

  function stripCalcChain(files) {
    Object.keys(files).forEach(function (name) {
      if (/calcChain\.xml$/i.test(name)) {
        delete files[name];
      }
    });
    var typesName = Object.keys(files).filter(function (name) {
      return name.toLowerCase() === "[content_types].xml";
    })[0];
    if (typesName) {
      putXmlInZip(
        files,
        typesName,
        xmlFromZip(files, typesName).replace(/<Override\b[^>]*calcChain\.xml"[^>]*\/>/gi, "")
      );
    }
    Object.keys(files).forEach(function (name) {
      if (!/\.rels$/i.test(name)) {
        return;
      }
      var rels = xmlFromZip(files, name);
      if (rels.indexOf("calcChain") === -1) {
        return;
      }
      putXmlInZip(
        files,
        name,
        rels.replace(/<Relationship\b[^>]*calcChain\.xml"[^>]*\/>/gi, "")
      );
    });
  }

  function xmlssCellStyle(cellXml) {
    var match = /ss:StyleID="([^"]*)"/.exec(cellXml || "");
    return match ? match[1] : "";
  }

  function xmlssBuildCell(index, styleId, kind, value) {
    var style = styleId ? ' ss:StyleID="' + styleId + '"' : "";
    if (kind === "n") {
      return (
        '<Cell ss:Index="' +
        index +
        '"' +
        style +
        '><Data ss:Type="Number">' +
        calcNumberValue(value) +
        "</Data></Cell>"
      );
    }
    return (
      '<Cell ss:Index="' +
      index +
      '"' +
      style +
      '><Data ss:Type="String">' +
      xmlText(value) +
      "</Data></Cell>"
    );
  }

  function upsertXmlssRow(tableInner, row, patches) {
    var rowRe = elementRe("Row");
    var match;
    var implied = 0;
    var found = null;
    while ((match = rowRe.exec(tableInner))) {
      var attrs = match[1] || "";
      var indexMatch = /ss:Index="(\d+)"/.exec(attrs);
      implied = indexMatch ? Number(indexMatch[1]) : implied + 1;
      if (implied === row) {
        found = { match: match, implied: implied };
        break;
      }
    }
    function applyInner(inner) {
      var cellRe = elementRe("Cell");
      var cells = [];
      var cellMatch;
      var col = 0;
      while ((cellMatch = cellRe.exec(inner || ""))) {
        var xml = cellMatch[0];
        var idxMatch = /ss:Index="(\d+)"/.exec(xml);
        col = idxMatch ? Number(idxMatch[1]) : col + 1;
        cells.push({ index: col, xml: xml });
      }
      patches.forEach(function (patch) {
        var existing = cells.filter(function (cell) {
          return cell.index === patch.index;
        })[0];
        var styleId = existing ? xmlssCellStyle(existing.xml) : "";
        var next = xmlssBuildCell(patch.index, styleId, patch.kind, patch.value);
        if (existing) {
          existing.xml = next;
          return;
        }
        cells.push({ index: patch.index, xml: next });
      });
      cells.sort(function (a, b) {
        return a.index - b.index;
      });
      return cells
        .map(function (cell) {
          return cell.xml;
        })
        .join("");
    }
    var nextRow = '<Row ss:Index="' + row + '">' + applyInner(found ? found.match[3] || "" : "") + "</Row>";
    if (!found) {
      return tableInner + nextRow;
    }
    return (
      tableInner.slice(0, found.match.index) +
      nextRow +
      tableInner.slice(found.match.index + found.match[0].length)
    );
  }

  function writeCalcRowToXmlss(bytes, row, values) {
    var encoding =
      bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe ? "utf-16le" : "utf-8";
    var xml = decodeText(bytes);
    var sheets = xml.match(/<Worksheet\b[\s\S]*?<\/Worksheet>/g) || [];
    if (sheets.length === 0) {
      throw new Error("Excelにシートが見つかりませんでした。");
    }
    var target = sheets.filter(function (sheet) {
      return /ss:Name="計算"|ss:Name="一覧"/.test(sheet);
    })[0] || sheets[0];
    var tableMatch = /<Table\b[\s\S]*?<\/Table>/.exec(target);
    if (!tableMatch) {
      throw new Error("Excelシートへ書き込めませんでした。");
    }
    var openEnd = tableMatch[0].indexOf(">") + 1;
    var closeAt = tableMatch[0].lastIndexOf("</Table>");
    var patchedTable =
      tableMatch[0].slice(0, openEnd) +
      upsertXmlssRow(
        tableMatch[0].slice(openEnd, closeAt),
        row,
        calcPatches(values).map(function (patch) {
          return {
            index: colLettersToIndex(patch.col),
            kind: patch.kind,
            value: patch.value,
          };
        })
      ) +
      tableMatch[0].slice(closeAt);
    var patchedSheet = target.slice(0, tableMatch.index) + patchedTable + target.slice(tableMatch.index + tableMatch[0].length);
    var sheetAt = xml.indexOf(target);
    xml = xml.slice(0, sheetAt) + patchedSheet + xml.slice(sheetAt + target.length);
    return encodeText(xml, encoding);
  }

  function loadXlsWriter() {
    if (typeof window !== "undefined" && window.OptionXls) {
      return window.OptionXls;
    }
    if (typeof require === "function") {
      return require("./xls-writer.js");
    }
    throw new Error("xls-writer.js が読み込まれていません。");
  }

  function writeCalcRowToBinaryXls(bytes, row, values) {
    var patches = calcPatches(values).map(function (patch) {
      return {
        col: colLettersToIndex(patch.col) - 1,
        kind: patch.kind,
        value: patch.kind === "n" ? calcNumberValue(patch.value) : String(patch.value),
      };
    });
    return loadXlsWriter().writeXlsRow(bytes, row, patches);
  }

  function calcPatches(values) {
    var source = values || {};
    var patches = [];
    function addText(col, value) {
      var text = value == null ? "" : String(value).trim();
      if (!text) {
        return;
      }
      patches.push({ col: col, kind: "s", value: text });
    }
    addText("A", source.modelName);
    addText("B", source.drawingNo);
    addText("C", source.destination);
    addText("D", source.workNo);
    patches.push({ col: "E", kind: "n", value: source.a != null ? source.a : "" });
    patches.push({ col: "G", kind: "s", value: source.b != null ? String(source.b) : "" });
    patches.push({ col: "H", kind: "n", value: source.c != null ? source.c : "" });
    return patches;
  }

  var CALC_PR_FOLLOWERS = [
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
  ];

  function forceFullCalcOnLoad(workbookXmlText) {
    var xml = workbookXmlText || "";
    var calcRe = /<((?:\w+:)?calcPr)\b([^>]*?)(\/?)>/;
    var calc = calcRe.exec(xml);
    if (calc) {
      var attrs = calc[2].replace(/\s+fullCalcOnLoad="[^"]*"/g, "");
      return (
        xml.slice(0, calc.index) +
        "<" +
        calc[1] +
        attrs +
        ' fullCalcOnLoad="1"' +
        calc[3] +
        ">" +
        xml.slice(calc.index + calc[0].length)
      );
    }
    var root = /<((?:\w+:)?)workbook\b/.exec(xml);
    var prefix = root ? root[1] : "";
    var tag = "<" + prefix + 'calcPr fullCalcOnLoad="1"/>';
    var insertAt = -1;
    CALC_PR_FOLLOWERS.forEach(function (name) {
      var at = xml.search(new RegExp("<" + prefix + name + "\\b"));
      if (at !== -1 && (insertAt === -1 || at < insertAt)) {
        insertAt = at;
      }
    });
    if (insertAt === -1) {
      insertAt = xml.lastIndexOf("</" + prefix + "workbook>");
    }
    if (insertAt === -1) {
      throw new Error("Excelの再計算設定を書き込めませんでした。");
    }
    return xml.slice(0, insertAt) + tag + xml.slice(insertAt);
  }

  function isZipBytes(bytes) {
    var data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
    return data.length >= 2 && data[0] === 0x50 && data[1] === 0x4b;
  }

  function safeExportFilename(originalName, bytes) {
    var name = String(originalName || "export.xlsx");
    if (isZipBytes(bytes) && /\.xls$/i.test(name) && !/\.xlsx$/i.test(name) && !/\.xlsm$/i.test(name)) {
      return name.replace(/\.xls$/i, ".xlsx");
    }
    return name;
  }

  function writeCalcRowToXlsx(buffer, row, values) {
    var rowNum = Number(row);
    if (!Number.isInteger(rowNum) || rowNum < 1 || rowNum > 1000) {
      return Promise.reject(new Error("行番号は 1〜1000 の整数にしてください。"));
    }
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var kind = sniffWorkbook(bytes);
    if (kind === "xls") {
      try {
        return Promise.resolve(writeCalcRowToBinaryXls(bytes, rowNum, values || {}));
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (kind === "xmlss") {
      try {
        return Promise.resolve(writeCalcRowToXmlss(bytes, rowNum, values || {}));
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (kind !== "xlsx") {
      return Promise.reject(new Error("対応しているのは .xlsx と、XML形式のExcelです。"));
    }
    return unzipZip(bytes).then(function (files) {
      var xmlFiles = Object.create(null);
      Object.keys(files).forEach(function (name) {
        if (/\.(xml|rels)$/i.test(name) || name === "[Content_Types].xml") {
          xmlFiles[name] = decodeText(files[name]);
        }
      });
      var sheets = parseWorkbookSheets(xmlFiles);
      if (sheets.length === 0) {
        throw new Error("Excelにシートが見つかりませんでした。");
      }
      var target = pickWriteSheet(sheets);
      if (!files[target.path]) {
        throw new Error("書き込み先シートが見つかりませんでした。");
      }
      putXmlInZip(
        files,
        target.path,
        upsertSheetCells(xmlFromZip(files, target.path), rowNum, calcPatches(values))
      );
      if (files["xl/workbook.xml"]) {
        putXmlInZip(
          files,
          "xl/workbook.xml",
          forceFullCalcOnLoad(xmlFromZip(files, "xl/workbook.xml"))
        );
      }
      stripCalcChain(files);
      return buildZip(
        Object.keys(files).map(function (name) {
          return { name: name, bytes: files[name] };
        })
      );
    });
  }

  return {
    ALLOWED_INPUT_ROWS: ALLOWED_INPUT_ROWS,
    buildOptionWorkbook: buildOptionWorkbook,
    sanitizeFilename: sanitizeFilename,
    sanitizeFolderName: sanitizeFolderName,
    validateOptionInput: validateOptionInput,
    normalizePriceText: normalizePriceText,
    parsePrice: parsePrice,
    splitOptionCodes: splitOptionCodes,
    previewOptionSelection: previewOptionSelection,
    readXlsxOptions: readXlsxOptions,
    readMarkedOptionNos: readMarkedOptionNos,
    writeCalcRowToXlsx: writeCalcRowToXlsx,
    safeExportFilename: safeExportFilename,
    forceFullCalcOnLoad: forceFullCalcOnLoad,
  };
});
