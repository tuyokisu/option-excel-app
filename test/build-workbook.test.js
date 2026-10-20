"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const {
  buildOptionWorkbook,
  sanitizeFilename,
  sanitizeFolderName,
  validateOptionInput,
  previewOptionSelection,
  readXlsxOptions,
  readMarkedOptionNos,
  writeCalcRowToXlsx,
  safeExportFilename,
  forceFullCalcOnLoad,
} = require("../excel-builder");

const appDir = path.join(__dirname, "..");
const sampleOptions = [
  { no: "01", name: "サンルーフ", price: 50000 },
  { no: "2", name: "カーナビ", price: 80000 },
  { no: "10", name: "フロアマット", price: 12000 },
];

function xmlText(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function nameFormula(row, masterLast) {
  return xmlText(
    `IF(A${row}="","",IFERROR(INDEX(マスタ!$B$2:$B$${masterLast},MATCH(A${row}&"",マスタ!$A$2:$A$${masterLast},0)),"該当なし"))`
  );
}

function priceFormula(row, masterLast) {
  return xmlText(
    `IF(A${row}="","",IFERROR(INDEX(マスタ!$C$2:$C$${masterLast},MATCH(A${row}&"",マスタ!$A$2:$A$${masterLast},0)),""))`
  );
}

function splitFormula(row) {
  const start = (row - 8) * 20 + 1;
  return xmlText(
    `IF(B3="","",TRIM(MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B3&"","、",","),"，",",")&",",",",REPT(" ",20)),${start},20)))`
  );
}

function nameJoin(row) {
  if (row === 8) {
    return xmlText('IF(B8="","",B8)');
  }
  const prev = row - 1;
  return xmlText(`IF(B${row}="",E${prev},IF(E${prev}="",B${row},E${prev}&","&B${row}))`);
}

function unzipStore(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const files = {};
  let offset = 0;

  while (offset + 4 <= bytes.length) {
    const signature = view.getUint32(offset, true);
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break;
    }
    assert.equal(signature, 0x04034b50, "zip local header");
    const method = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const name = new TextDecoder().decode(bytes.subarray(nameStart, nameStart + nameLength));
    const dataStart = nameStart + nameLength + extraLength;
    assert.equal(method, 0, name);
    files[name] = new TextDecoder().decode(bytes.subarray(dataStart, dataStart + compressedSize));
    offset = dataStart + compressedSize;
  }

  return files;
}

async function readWorkbook(spec) {
  const built = buildOptionWorkbook(spec);
  const bytes = await built.xlsx.writeBuffer();
  return unzipStore(bytes);
}

test("sanitizes Windows filenames", () => {
  assert.equal(sanitizeFilename("標準仕様A"), "標準仕様A_オプション.xlsx");
  assert.equal(sanitizeFilename(' a/b:c*?<>|" '), "a_b_c_______オプション.xlsx");
  assert.equal(sanitizeFilename("   "), "オプション_オプション.xlsx");
});

test("sanitizes 機種 folder names", () => {
  assert.equal(sanitizeFolderName("ABC-100"), "ABC-100");
  assert.equal(sanitizeFolderName(" a/b:c "), "a_b_c");
  assert.equal(sanitizeFolderName(".."), "");
  assert.equal(sanitizeFolderName("   "), "");
});

test("validates option rows and keeps leading zeros distinct", () => {
  const ok = validateOptionInput({
    specName: " 標準仕様 ",
    basePrice: "2,500,000",
    inputRows: 300,
    options: [
      { no: "", name: "", price: "" },
      { no: "01", name: "サンルーフ", price: "50,000" },
      { no: "1", name: "ナビ", price: 10.5 },
    ],
  });
  assert.equal(ok.ok, true);
  assert.equal(ok.specName, "標準仕様");
  assert.equal(ok.basePrice, 2500000);
  assert.deepEqual(ok.options, [
    { no: "01", name: "サンルーフ", price: 50000 },
    { no: "1", name: "ナビ", price: 10.5 },
  ]);

  const missingBase = validateOptionInput({
    specName: "仕様",
    inputRows: 100,
    basePrice: "",
    options: [{ no: "01", name: "A", price: 1 }],
  });
  assert.equal(missingBase.ok, false);
  assert.match(missingBase.errors.join("\n"), /本体価格を入力してください/);

  const badBase = validateOptionInput({
    specName: "仕様",
    inputRows: 100,
    basePrice: "abc",
    options: [{ no: "01", name: "A", price: 1 }],
  });
  assert.equal(badBase.ok, false);
  assert.match(badBase.errors.join("\n"), /本体価格は 0 以上/);

  const duplicate = validateOptionInput({
    specName: "仕様",
    basePrice: 1,
    inputRows: 100,
    options: [
      { no: "01", name: "A", price: 1 },
      { no: "01", name: "B", price: 2 },
    ],
  });
  assert.equal(duplicate.ok, false);
  assert.match(duplicate.errors.join("\n"), /2行目/);

  const negative = validateOptionInput({
    specName: "仕様",
    basePrice: 1,
    inputRows: 100,
    options: [{ no: "01", name: "A", price: -1 }],
  });
  assert.equal(negative.ok, false);
  assert.match(negative.errors.join("\n"), /1行目/);
});

test("builds a 100-row workbook and writes the sample file", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: "2,500,000",
    inputRows: 100,
    options: sampleOptions,
  });
  const bytes = await built.xlsx.writeBuffer();
  const outDir = path.join(appDir, "test-output");
  const outFile = path.join(outDir, "sample.xlsx");
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(outFile, bytes);

  const files = unzipStore(fs.readFileSync(outFile));
  const workbook = files["xl/workbook.xml"];
  const master = files["xl/worksheets/sheet1.xml"];
  const input = files["xl/worksheets/sheet2.xml"];
  const styles = files["xl/styles.xml"];

  assert.equal(workbook.includes('name="マスタ"'), true);
  assert.equal(workbook.includes('name="入力"'), true);
  assert.equal(workbook.includes('activeTab="1"'), true);
  assert.equal(workbook.includes('fullCalcOnLoad="1"'), true);
  assert.equal(workbook.includes("OptionList"), false);

  assert.equal(master.includes(">オプションNo</t>"), true);
  assert.equal(master.includes(">品名</t>"), true);
  assert.equal(master.includes(">値段</t>"), true);
  assert.equal(master.includes('ySplit="1"'), true);
  assert.equal(master.includes('<c r="A2" s="2" t="inlineStr"><is><t>01</t></is></c>'), true);
  assert.equal(master.includes("<t>サンルーフ</t>"), true);
  assert.equal(master.includes('<c r="C2" s="3"><v>50000</v></c>'), true);
  assert.equal(master.includes("<t>2</t>"), true);
  assert.equal(master.includes("<t>10</t>"), true);
  assert.equal(master.includes("<v>12000</v>"), true);
  assert.equal(master.includes('r="A5"'), false);
  assert.equal(styles.includes('numFmtId="49"'), true);
  assert.equal(styles.includes('numFmtId="3"'), true);

  assert.equal(input.includes(">図番番号</t>"), true);
  assert.equal(input.includes("<t>標準仕様A</t>"), true);
  assert.equal(input.includes(">本体価格</t>"), true);
  assert.equal(input.includes('<c r="B2" s="11"><v>2500000</v></c>'), true);
  assert.equal(input.includes(">オプション合計</t>"), true);
  assert.equal(input.includes(">総合計</t>"), true);
  assert.equal(input.includes("<f>SUM(C8:C107)</f>"), true);
  assert.equal(input.includes("<f>B2+B5</f>"), true);
  assert.equal(input.includes(">品目リスト</t>"), true);
  assert.equal(input.includes("<t>例: 1,2,3</t>"), true);
  assert.equal(input.includes('<c r="B3" s="15"/>'), true);
  assert.equal(input.includes('<c r="B4" s="6">'), true);
  assert.equal(input.includes('<c r="B5" s="13">'), true);
  assert.equal(input.includes('<c r="B6" s="13">'), true);
  assert.equal(input.includes(">オプションNo</t>"), true);
  assert.equal(input.includes(">品名</t>"), true);
  assert.equal(input.includes('ySplit="7"'), true);
  assert.equal(input.includes('activeCell="B3"'), true);
  assert.equal(input.includes('min="5" max="5" width="28" customWidth="1" hidden="1"'), true);
  assert.equal(input.includes(`<f>${splitFormula(8)}</f>`), true);
  assert.equal(input.includes(`<f>${splitFormula(107)}</f>`), true);
  assert.equal(input.includes(`<f>${nameFormula(8, 4)}</f>`), true);
  assert.equal(input.includes(`<f>${priceFormula(8, 4)}</f>`), true);
  assert.equal(input.includes(`<f>${nameFormula(107, 4)}</f>`), true);
  assert.equal(input.includes(`<f>${priceFormula(107, 4)}</f>`), true);
  assert.equal(input.includes('<c r="A8" s="2">'), true);
  assert.equal(input.includes('<c r="C8" s="3">'), true);
  assert.equal(input.includes(`<f>${nameJoin(8)}</f>`), true);
  assert.equal(input.includes(`<f>${nameJoin(9)}</f>`), true);
  assert.equal(input.includes(`<f>${nameJoin(107)}</f>`), true);
  assert.equal(input.includes("<f>E107</f>"), true);
  assert.equal(input.includes("<f>B5</f>"), true);
  assert.equal(input.includes("<f>B6</f>"), true);
  assert.equal(input.includes('<c r="C108" s="8">'), true);
  assert.equal(input.includes('<c r="C109" s="8">'), true);
  assert.equal(input.includes('r="A110"'), false);
  assert.equal(input.includes('type="list"'), false);
  assert.equal(input.includes("XLOOKUP"), false);
  assert.equal(input.includes("TEXTJOIN"), false);
  assert.equal(input.includes("TEXTSPLIT"), false);
  assert.equal(master.includes("XLOOKUP"), false);
  assert.equal(fs.statSync(outFile).size > 0, true);
});

test("a 300-row workbook ends on the matching row", async () => {
  const files = await readWorkbook({
    specName: "長尺仕様",
    basePrice: 2000000,
    inputRows: 300,
    options: [
      { no: "01", name: "サンルーフ", price: 1500.5 },
      { no: "02", name: "ナビ", price: 80000 },
    ],
  });
  const workbook = files["xl/workbook.xml"];
  const master = files["xl/worksheets/sheet1.xml"];
  const input = files["xl/worksheets/sheet2.xml"];

  assert.equal(workbook.includes("OptionList"), false);
  assert.equal(input.includes("<f>E307</f>"), true);
  assert.equal(input.includes(`<f>${splitFormula(8)}</f>`), true);
  assert.equal(input.includes(`<f>${splitFormula(307)}</f>`), true);
  assert.equal(input.includes(`<f>${nameFormula(8, 3)}</f>`), true);
  assert.equal(input.includes(`<f>${nameFormula(307, 3)}</f>`), true);
  assert.equal(input.includes(`<f>${priceFormula(307, 3)}</f>`), true);
  assert.equal(input.includes(`<f>${nameJoin(307)}</f>`), true);
  assert.equal(input.includes("<t>品目リスト</t>"), true);
  assert.equal(input.includes("<t>オプション合計</t>"), true);
  assert.equal(input.includes("<t>総合計</t>"), true);
  assert.equal(input.includes("<f>SUM(C8:C307)</f>"), true);
  assert.equal(input.includes("<f>B2+B5</f>"), true);
  assert.equal(input.includes('<c r="B2" s="12"><v>2000000</v></c>'), true);
  assert.equal(input.includes('<c r="B3" s="15"/>'), true);
  assert.equal(master.includes('<c r="C2" s="4"><v>1500.5</v></c>'), true);
  assert.equal(input.includes('<c r="C8" s="4">'), true);
  assert.equal(input.includes('<c r="B6" s="14">'), true);
  assert.equal(input.includes('<c r="C308" s="9">'), true);
  assert.equal(input.includes('<c r="C309" s="9">'), true);
  assert.equal(input.includes("TEXTSPLIT"), false);
  assert.equal(input.includes("XLOOKUP"), false);
  assert.equal(input.includes("TEXTJOIN"), false);
});

test("previews Excel totals from comma-separated option numbers", () => {
  const preview = previewOptionSelection({
    basePrice: "2,500,000",
    inputRows: 100,
    selectedNos: "01、2, 99",
    options: sampleOptions,
  });
  assert.equal(preview.basePrice, 2500000);
  assert.equal(preview.optionTotal, 130000);
  assert.equal(preview.grandTotal, 2630000);
  assert.equal(preview.itemList, "サンルーフ,カーナビ,該当なし");
  assert.equal(preview.selectedNosText, "01,2");
  assert.equal(preview.selectedWithNames, "01 サンルーフ、2 カーナビ");
  assert.equal(preview.masterCount, 3);
  assert.equal(preview.items[2].missing, true);
});

test("writes calculation values into a chosen Excel row", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: 100,
    inputRows: 100,
    options: sampleOptions,
  });
  const source = await built.xlsx.writeBuffer();
  const before = unzipStore(source);
  const next = await writeCalcRowToXlsx(source, 4, {
    a: 2500000,
    b: "70,73",
    c: 130000,
  });
  const files = unzipStore(next);
  const input = files["xl/worksheets/sheet2.xml"];
  assert.equal(input.includes("<v>2500000</v>"), true);
  assert.equal(input.includes("70,73"), true);
  assert.equal(input.includes("<v>130000</v>"), true);
  assert.equal(input.includes('r="E4"'), true);
  assert.equal(input.includes('r="G4"'), true);
  assert.equal(input.includes('r="H4"'), true);
  assert.equal(input.includes("品目リスト"), true);
  assert.equal(input.includes("<f>" + nameJoin(8) + "</f>"), true);
  assert.equal(files["xl/styles.xml"], before["xl/styles.xml"]);
  assert.equal(files["xl/worksheets/sheet1.xml"], before["xl/worksheets/sheet1.xml"]);

  const withMeta = await writeCalcRowToXlsx(source, 4, {
    a: 2500000,
    b: "70,73",
    c: 130000,
    modelName: "MX-1",
    drawingNo: "Z-10",
    destination: "国内",
    workNo: "W-9",
  });
  const metaInput = unzipStore(withMeta)["xl/worksheets/sheet2.xml"];
  assert.equal(metaInput.includes("MX-1"), true);
  assert.equal(metaInput.includes("Z-10"), true);
  assert.equal(metaInput.includes("国内"), true);
  assert.equal(metaInput.includes("W-9"), true);
  assert.equal(metaInput.includes('r="A4"'), true);
  assert.equal(metaInput.includes('r="D4"'), true);
});

test("overwriting a row next to an empty bordered cell keeps one cell per address", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: 100,
    inputRows: 100,
    options: sampleOptions,
  });
  const source = await built.xlsx.writeBuffer();
  assert.match(unzipStore(source)["xl/worksheets/sheet2.xml"], /<c r="B3" s="\d+"\/><c r="C3"/);

  const once = await writeCalcRowToXlsx(source, 3, {
    a: 80000,
    b: "1,3",
    c: 9587,
    destination: "国内",
  });
  const twice = await writeCalcRowToXlsx(once, 3, {
    a: 80000,
    b: "1,3,5",
    c: 12000,
    destination: "海外",
  });
  const sheet = unzipStore(twice)["xl/worksheets/sheet2.xml"];
  const row3 = /<row r="3"[^>]*>([\s\S]*?)<\/row>/.exec(sheet)[1];
  const refs = [...row3.matchAll(/<c r="([A-Z]+)3"/g)].map((m) => m[1]);
  assert.deepEqual(refs, ["A", "B", "C", "E", "G", "H"]);
  assert.equal(row3.includes("海外"), true);
  assert.equal(row3.includes("国内"), false);
  assert.equal(row3.includes("1,3,5"), true);
  assert.equal((sheet.match(/<row r="3"/g) || []).length, 1);
});

test("forces Excel to recalculate formulas when the file opens", () => {
  const added = forceFullCalcOnLoad(
    '<workbook xmlns="m"><sheets><sheet name="a"/></sheets><definedNames/><extLst/></workbook>'
  );
  assert.equal(
    added,
    '<workbook xmlns="m"><sheets><sheet name="a"/></sheets><definedNames/><calcPr fullCalcOnLoad="1"/><extLst/></workbook>'
  );
  const updated = forceFullCalcOnLoad(
    '<workbook><sheets/><calcPr calcId="191029" fullCalcOnLoad="0"/></workbook>'
  );
  assert.equal(updated, '<workbook><sheets/><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
  const prefixed = forceFullCalcOnLoad("<x:workbook><x:sheets/></x:workbook>");
  assert.equal(prefixed, '<x:workbook><x:sheets/><x:calcPr fullCalcOnLoad="1"/></x:workbook>');
});

test("written workbook recalculates on open", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: 100,
    inputRows: 100,
    options: sampleOptions,
  });
  const source = await built.xlsx.writeBuffer();
  const next = await writeCalcRowToXlsx(source, 4, { a: 80000, b: "1,3", c: 9587 });
  const workbook = unzipStore(next)["xl/workbook.xml"];
  assert.equal((workbook.match(/<calcPr\b/g) || []).length, 1);
  assert.equal(workbook.includes('fullCalcOnLoad="1"'), true);
});

test("does not keep zip bytes under an .xls name", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: 100,
    inputRows: 100,
    options: sampleOptions,
  });
  const zip = await built.xlsx.writeBuffer();
  assert.equal(safeExportFilename("sample.xls", zip), "sample.xlsx");
  assert.equal(safeExportFilename("sample.xlsx", zip), "sample.xlsx");
});

function readXlsCells(bytes) {
  const xls = require("../xls-writer");
  const stream = xls.findWorkbookEntry(xls.readCfb(bytes)).data;
  const records = xls.parseRecords(stream, 0, stream.length);
  const sstAt = records.findIndex((record) => record.type === xls.RT.SST);
  let end = sstAt + 1;
  while (records[end].type === xls.RT.CONTINUE) {
    end += 1;
  }
  const sst = xls.readSst(records.slice(sstAt, end));
  const sheetPos = records
    .filter((record) => record.type === xls.RT.BOUNDSHEET)
    .map((record) => xls.u32(record.data, 0))[0];
  const cells = {};
  let labelCount = 0;
  let depth = 0;
  for (const record of records) {
    if (record.pos < sheetPos) {
      continue;
    }
    if (record.type === xls.RT.BOF) {
      depth += 1;
    } else if (record.type === xls.RT.EOF) {
      depth -= 1;
      if (depth === 0) {
        break;
      }
    }
    const rw = xls.u16(record.data, 0) + 1;
    const col = String.fromCharCode(65 + xls.u16(record.data, 2));
    if (record.type === 0x0203) {
      cells[col + rw] = new DataView(record.data.buffer, record.data.byteOffset + 6, 8).getFloat64(0, true);
    } else if (record.type === 0x00fd) {
      cells[col + rw] = sst.strings[xls.u32(record.data, 6)].text;
    }
  }
  for (const record of records) {
    if (record.type === 0x00fd) {
      labelCount += 1;
    }
  }
  return { cells, sst, labelCount };
}

const sampleXlsPath = path.join(__dirname, "fixtures", "sample.xls");

test("rebuilding the Excel-made .xls without changes keeps every byte", () => {
  const xls = require("../xls-writer");
  const source = fs.readFileSync(sampleXlsPath);
  const cfb = xls.readCfb(source);
  const stream = xls.findWorkbookEntry(cfb).data;
  const same = xls.rewriteWorkbookStream(stream, 4, [], { keepRecalcId: true });
  assert.equal(Buffer.compare(Buffer.from(same), Buffer.from(stream)), 0);
  const again = xls.readCfb(xls.writeCfb(cfb));
  cfb.entries.forEach((entry, index) => {
    assert.equal(again.entries[index].name, entry.name);
    if (entry.type === 2) {
      assert.equal(Buffer.compare(Buffer.from(again.entries[index].data), Buffer.from(entry.data)), 0);
    }
  });
});

test("writes calculation values into a binary .xls row", async () => {
  const source = fs.readFileSync(sampleXlsPath);
  const before = readXlsCells(source);
  const written = await writeCalcRowToXlsx(source, 4, {
    a: "1,000,000",
    b: "1,2,3",
    c: 50000,
    modelName: "上書き機種",
  });
  assert.equal(written[0], 0xd0);
  assert.equal(safeExportFilename("sample.xls", written), "sample.xls");
  const after = readXlsCells(written);
  assert.equal(after.cells.A4, "上書き機種");
  assert.equal(after.cells.B4, before.cells.B4);
  assert.equal(after.cells.E4, 1000000);
  assert.equal(after.cells.G4, "1,2,3");
  assert.equal(after.cells.H4, 50000);
  assert.equal(after.cells.A40, before.cells.A40);
  assert.equal(after.sst.total, after.labelCount);
});

test("writes a new .xls row twice and keeps long text intact", async () => {
  const source = fs.readFileSync(sampleXlsPath);
  const long = "長".repeat(300) + "ABC";
  let written = await writeCalcRowToXlsx(source, 6, { a: 10, b: "x", c: 20 });
  written = await writeCalcRowToXlsx(written, 6, { a: 30, b: long, c: 40, workNo: "W6" });
  const after = readXlsCells(written);
  assert.equal(after.cells.D6, "W6");
  assert.equal(after.cells.E6, 30);
  assert.equal(after.cells.G6, long);
  assert.equal(after.cells.H6, 40);
  assert.equal(after.sst.total, after.labelCount);
});

test("rejects .xls files that are not Excel 97-2003 workbooks", async () => {
  const ole = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  await assert.rejects(() => writeCalcRowToXlsx(ole, 2, { a: 1, b: "2", c: 3 }));
});

test("writes into XML Spreadsheet .xls without dropping other cells", async () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<?mso-application progid="Excel.Sheet"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Worksheet ss:Name="Sheet1"><Table>' +
    '<Row ss:Index="1"/>' +
    "<Row>" +
    '<Cell ss:StyleID="s21"><Data ss:Type="Number">1</Data></Cell>' +
    '<Cell ss:Index="4"><Data ss:Type="String">keep</Data></Cell>' +
    '<Cell ss:Index="6"><Data ss:Type="String">skip</Data></Cell>' +
    "</Row></Table></Worksheet></Workbook>";
  const next = await writeCalcRowToXlsx(new TextEncoder().encode(xml), 2, {
    a: 2500000,
    b: "70,73",
    c: 130000,
  });
  const text = new TextDecoder().decode(next);
  assert.equal(text.includes("2500000"), true);
  assert.equal(text.includes("70,73"), true);
  assert.equal(text.includes("130000"), true);
  assert.equal(text.includes("keep"), true);
  assert.equal(text.includes("skip"), true);
  assert.equal(text.includes('ss:StyleID="s21"'), true);
  assert.equal((text.match(/<Row\b/g) || []).length, 2);
});

test("keeps non-sequential Excel option numbers", async () => {
  const built = buildOptionWorkbook({
    specName: "欠番仕様",
    basePrice: 1,
    inputRows: 100,
    options: [
      { no: "70", name: "コメ", price: 900 },
      { no: "73", name: "もち", price: 799 },
    ],
  });
  const parsed = await readXlsxOptions(await built.xlsx.writeBuffer());
  assert.deepEqual(
    parsed.options.map((item) => item.no),
    ["70", "73"]
  );
});

test("reads option rows from a generated workbook", async () => {
  const built = buildOptionWorkbook({
    specName: "標準仕様A",
    basePrice: "2500000",
    inputRows: 100,
    options: sampleOptions,
  });
  const bytes = await built.xlsx.writeBuffer();
  const parsed = await readXlsxOptions(bytes);
  assert.equal(parsed.specName, "標準仕様A");
  assert.equal(parsed.basePrice, "2500000");
  assert.deepEqual(parsed.options, [
    { no: "01", name: "サンルーフ", price: "50000" },
    { no: "2", name: "カーナビ", price: "80000" },
    { no: "10", name: "フロアマット", price: "12000" },
  ]);
});

test("reads circled option numbers from a spec Excel", async () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Worksheet ss:Name="仕様"><Table>' +
    '<Row><Cell><Data ss:Type="Number">1</Data></Cell><Cell><Data ss:Type="String">サンルーフ</Data></Cell><Cell><Data ss:Type="String">○</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="Number">2</Data></Cell><Cell><Data ss:Type="String">ナビ</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="Number">3</Data></Cell><Cell><Data ss:Type="String">エアロ</Data></Cell><Cell><Data ss:Type="String">●</Data></Cell></Row>' +
    "</Table></Worksheet></Workbook>";
  const parsed = await readMarkedOptionNos(new TextEncoder().encode(xml));
  assert.deepEqual(parsed.nos, ["1", "3"]);
});

test("reads circled options on both sides of a spec Excel", async () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Worksheet ss:Name="仕様"><Table>' +
    '<Row ss:Index="1">' +
    '<Cell ss:Index="1"><Data ss:Type="Number">1</Data></Cell>' +
    '<Cell ss:Index="2"><Data ss:Type="String">左サンルーフ</Data></Cell>' +
    '<Cell ss:Index="5"><Data ss:Type="Number">9</Data></Cell>' +
    '<Cell ss:Index="6"><Data ss:Type="String">右マット</Data></Cell>' +
    "</Row>" +
    '<Row ss:Index="2">' +
    '<Cell ss:Index="3"><Data ss:Type="String">○</Data></Cell>' +
    '<Cell ss:Index="7"><Data ss:Type="String">○</Data></Cell>' +
    "</Row>" +
    '<Row ss:Index="3">' +
    '<Cell ss:Index="1"><Data ss:Type="Number">2</Data></Cell>' +
    '<Cell ss:Index="2"><Data ss:Type="String">左ナビ</Data></Cell>' +
    '<Cell ss:Index="5"><Data ss:Type="Number">10</Data></Cell>' +
    '<Cell ss:Index="6"><Data ss:Type="String">右エアロ</Data></Cell>' +
    '<Cell ss:Index="7"><Data ss:Type="String">○</Data></Cell>' +
    "</Row>" +
    "</Table></Worksheet></Workbook>";
  const parsed = await readMarkedOptionNos(new TextEncoder().encode(xml));
  assert.deepEqual(parsed.nos, ["1", "9", "10"]);
});

test("reads prices from circled left-side options", async () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Worksheet ss:Name="仕様"><Table>' +
    '<Row ss:Index="1">' +
    '<Cell ss:Index="1"><Data ss:Type="Number">1</Data></Cell>' +
    '<Cell ss:Index="2"><Data ss:Type="String">サンルーフ</Data></Cell>' +
    '<Cell ss:Index="3"><Data ss:Type="Number">50000</Data></Cell>' +
    '<Cell ss:Index="6"><Data ss:Type="Number">9</Data></Cell>' +
    '<Cell ss:Index="7"><Data ss:Type="String">右マット</Data></Cell>' +
    '<Cell ss:Index="8"><Data ss:Type="Number">999</Data></Cell>' +
    "</Row>" +
    '<Row ss:Index="2">' +
    '<Cell ss:Index="4"><Data ss:Type="String">○</Data></Cell>' +
    '<Cell ss:Index="9"><Data ss:Type="String">○</Data></Cell>' +
    "</Row>" +
    '<Row ss:Index="3">' +
    '<Cell ss:Index="1"><Data ss:Type="Number">2</Data></Cell>' +
    '<Cell ss:Index="2"><Data ss:Type="String">ナビ</Data></Cell>' +
    '<Cell ss:Index="3"><Data ss:Type="Number">80000</Data></Cell>' +
    "</Row>" +
    "</Table></Worksheet></Workbook>";
  const parsed = await readMarkedOptionNos(new TextEncoder().encode(xml));
  assert.deepEqual(parsed.nos, ["1", "9"]);
  assert.equal(parsed.options[0].no, "1");
  assert.equal(parsed.options[0].name, "サンルーフ");
  assert.equal(parsed.options[0].price, "50000");
  assert.equal(parsed.options[1].no, "2");
  assert.equal(parsed.options[1].price, "80000");
});

test("ignores a number written in the name column", async () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Worksheet ss:Name="仕様"><Table>' +
    '<Row><Cell><Data ss:Type="String">番号</Data></Cell><Cell><Data ss:Type="String">品名</Data></Cell><Cell><Data ss:Type="String">選択</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="Number">1</Data></Cell><Cell><Data ss:Type="Number">9</Data></Cell><Cell><Data ss:Type="String">○</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="Number">2</Data></Cell><Cell><Data ss:Type="String">ナビ</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="Number">4</Data></Cell><Cell><Data ss:Type="String">エアロ</Data></Cell></Row>' +
    '<Row><Cell><Data ss:Type="String">○</Data></Cell></Row>' +
    "</Table></Worksheet></Workbook>";
  const parsed = await readMarkedOptionNos(new TextEncoder().encode(xml));
  assert.deepEqual(parsed.nos, ["1", "4"]);
  assert.equal(parsed.options[0].name, "");
  assert.equal(parsed.options[1].name, "ナビ");
});

test("the page is a classic offline script", () => {
  const html = fs.readFileSync(path.join(appDir, "index.html"), "utf8");
  assert.equal(/https?:\/\//.test(html), false);
  assert.equal(html.includes("exceljs"), false);
  const xlsWriter = html.indexOf("xls-writer.js");
  const builder = html.indexOf("excel-builder.js");
  const app = html.indexOf("app.js");
  assert.equal(xlsWriter !== -1 && xlsWriter < builder, true);
  assert.equal(builder !== -1 && builder < app, true);
  assert.equal(html.includes('type="module"'), false);
  assert.equal(html.includes('id="basePrice"'), true);
  assert.equal(html.includes('id="modelName"'), true);
  assert.equal(html.includes("図番番号"), true);
  assert.equal(html.includes('id="drawingNo"'), true);
  assert.equal(html.includes('id="destination"'), true);
  assert.equal(html.includes('id="workNo"'), true);
  assert.equal(html.includes("saveデータ(読み込むよう）"), true);
  assert.equal(html.includes('id="saveJson"'), true);
  assert.equal(html.includes('id="overwriteJson"'), true);
  assert.equal(html.includes("上書き保存"), true);
  assert.equal(html.includes('id="saveFolderDialog"'), true);
  assert.equal(html.includes('id="deleteJson"'), true);
  assert.equal(html.includes("JSONを消す"), true);
  assert.equal(html.includes('id="savedFileSearch"'), true);
  assert.equal(html.includes("フォルダ名で検索"), true);
  assert.equal(html.includes("index.html"), true);
  assert.equal(html.includes("読み込み用excel"), true);
  assert.equal(html.includes("書き込むエクセル"), true);
  assert.equal(html.includes('id="selectedNos"'), true);
  assert.equal(html.includes('id="doneDialog"'), true);
  assert.equal(html.includes("仕様書の計算"), true);
  assert.equal(html.includes('id="excelFilePicker"'), true);
  assert.equal(html.includes('id="exportRow"'), true);
  assert.equal(html.includes('id="chooseExportExcel"'), true);
  assert.equal(html.includes('id="exportSelectedName"'), true);
  assert.equal(html.includes("エクセルに書き込む"), true);
  assert.equal(html.includes('id="writeCalcRow"'), true);
  assert.equal(html.includes("該当するオプション"), true);
  assert.equal(html.includes('id="chooseMarkExcel"'), true);
  assert.equal(html.includes('id="loadMarkExcel"'), true);
});
