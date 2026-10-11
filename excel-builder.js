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

  return {
    ALLOWED_INPUT_ROWS: ALLOWED_INPUT_ROWS,
    buildOptionWorkbook: buildOptionWorkbook,
    sanitizeFilename: sanitizeFilename,
    validateOptionInput: validateOptionInput,
    normalizePriceText: normalizePriceText,
    parsePrice: parsePrice,
    splitOptionCodes: splitOptionCodes,
    previewOptionSelection: previewOptionSelection,
  };
});
