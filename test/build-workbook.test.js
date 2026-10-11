"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const {
  buildOptionWorkbook,
  sanitizeFilename,
  validateOptionInput,
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

  assert.equal(input.includes(">仕様名</t>"), true);
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

test("the page is a classic offline script", () => {
  const html = fs.readFileSync(path.join(appDir, "index.html"), "utf8");
  assert.equal(/https?:\/\//.test(html), false);
  assert.equal(html.includes("exceljs"), false);
  const builder = html.indexOf("excel-builder.js");
  const app = html.indexOf("app.js");
  assert.equal(builder !== -1 && builder < app, true);
  assert.equal(html.includes('type="module"'), false);
  assert.equal(html.includes('id="basePrice"'), true);
  assert.equal(html.includes("saveデータ"), true);
  assert.equal(html.includes('id="saveJson"'), true);
  assert.equal(html.includes('id="loadJson"'), true);
  assert.equal(html.includes("開く.bat"), true);
  assert.equal(html.includes("仕様書一覧"), true);
});
