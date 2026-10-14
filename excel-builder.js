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

  function sanitizeFilename(specName) {
    var base = trimText(specName)
      .replace(ILLEGAL_FILENAME, "_")
      .replace(/[\u0000-\u001f]/g, "")
      .replace(/[. ]+$/g, "")
      .trim();
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
      errors.push("仕様名を入力してください。");
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
        inlineCell("A1", STYLE.label, "仕様名") +
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

  function buildZip(files) {
    var encoder = new TextEncoder();
    var stamp = dosDateTime(new Date(2026, 9, 11, 12, 0, 0));
    var locals = [];
    var centrals = [];
    var offset = 0;

    files.forEach(function (file) {
      var nameBytes = encoder.encode(file.name);
      var data = encoder.encode(file.text);
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

  function unzipXlsx(buffer) {
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
          return [job.name.replace(/^\/+/, ""), new TextDecoder("utf-8").decode(data)];
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

  function sheetCellValue(cellXml, shared) {
    var typeMatch = /\bt="([^"]+)"/.exec(cellXml);
    var type = typeMatch ? typeMatch[1] : "";
    if (type === "inlineStr") {
      var inline = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(cellXml);
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
    var cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    var match;
    while ((match = cellRe.exec(xml))) {
      var attrs = match[1];
      var refMatch = /\br="([^"]+)"/.exec(attrs);
      if (!refMatch) {
        continue;
      }
      var parts = cellRefParts(refMatch[1]);
      if (!parts) {
        continue;
      }
      if (!grid[parts.row]) {
        grid[parts.row] = Object.create(null);
      }
      grid[parts.row][parts.col] = sheetCellValue(attrs + ">" + match[2], shared);
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
      if (no === "仕様名" || no === "本体価格" || no === "オプションNo" || no === "品目リスト") {
        return;
      }
      started = true;
      options.push({ no: no, name: name, price: price });
    });
    return options;
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
      var master =
        sheets.filter(function (sheet) {
          return sheet.name === "マスタ";
        })[0] || sheets[0];
      var options = optionsFromGrid(parseSheetGrid(files[master.path], shared));
      if (options.length === 0 && sheets[0] && sheets[0] !== master) {
        options = optionsFromGrid(parseSheetGrid(files[sheets[0].path], shared));
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
        if (trimText(row1.A) === "仕様名") {
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

  return {
    ALLOWED_INPUT_ROWS: ALLOWED_INPUT_ROWS,
    buildOptionWorkbook: buildOptionWorkbook,
    sanitizeFilename: sanitizeFilename,
    validateOptionInput: validateOptionInput,
    normalizePriceText: normalizePriceText,
    parsePrice: parsePrice,
    splitOptionCodes: splitOptionCodes,
    previewOptionSelection: previewOptionSelection,
    readXlsxOptions: readXlsxOptions,
  };
});
