const assert = require("assert");
const crypto = require("crypto");
const fs = require("fs");
const vm = require("vm");

class Range {
    constructor(sheet, row, col, rows = 1, cols = 1) { Object.assign(this, { sheet, row, col, rows, cols }); }
    ensure(row, col) {
        while (this.sheet.rows.length <= row) this.sheet.rows.push([]);
        while (this.sheet.rows[row].length <= col) this.sheet.rows[row].push("");
    }
    getValues() {
        return Array.from({ length: this.rows }, (_, i) =>
            Array.from({ length: this.cols }, (_, j) => this.sheet.rows[this.row - 1 + i]?.[this.col - 1 + j] ?? ""));
    }
    getValue() { return this.getValues()[0][0]; }
    setValues(values) {
        values.forEach((row, i) => row.forEach((value, j) => {
            this.ensure(this.row - 1 + i, this.col - 1 + j);
            this.sheet.rows[this.row - 1 + i][this.col - 1 + j] = value;
        }));
        return this;
    }
    setValue(value) { return this.setValues([[value]]); }
    merge() { return this; }
    setBackground() { return this; }
    setBackgrounds() { return this; }
    setFontColor() { return this; }
    setFontWeight() { return this; }
    setFontWeights() { return this; }
    setFontSize() { return this; }
    setHorizontalAlignment() { return this; }
    setNumberFormat() { return this; }
    clearDataValidations() { return this; }
    breakApart() { return this; }
}

class Sheet {
    constructor(name, rows = [], maxColumns = 26) {
        this.name = name;
        this.rows = rows;
        this.hidden = false;
        this.hiddenColumns = new Set();
        this.maxColumns = Math.max(maxColumns, ...rows.map((row) => row.length));
    }
    getName() { return this.name; }
    getLastRow() { return this.rows.length; }
    getRange(row, col, rows = 1, cols = 1) { return new Range(this, row, col, rows, cols); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.rows.length), Math.max(1, ...this.rows.map((row) => row.length))); }
    appendRow(row) { this.rows.push([...row]); }
    clearContents() { this.rows = this.rows.map((row) => row.map(() => "")); return this; }
    clearFormats() { return this; }
    setColumnWidth() { return this; }
    setFrozenRows() { return this; }
    hideSheet() { this.hidden = true; return this; }
    hideColumns(column, count = 1) {
        for (let i = 0; i < count; i++) this.hiddenColumns.add(column + i);
        return this;
    }
    getMaxColumns() { return this.maxColumns; }
    insertColumnAfter() { this.maxColumns++; return this; }
    insertColumnsAfter(_column, count) { this.maxColumns += count; return this; }
    deleteRows(start, count) { this.rows.splice(start - 1, count); return this; }
}

class Book {
    constructor(sheets = []) {
        this.sheets = new Map(sheets.map((sheet) => [sheet.name, sheet]));
        this.sheetReadCounts = {};
    }
    getSheetByName(name) {
        this.sheetReadCounts[name] = (this.sheetReadCounts[name] || 0) + 1;
        return this.sheets.get(name) || null;
    }
    insertSheet(name) { const sheet = new Sheet(name); this.sheets.set(name, sheet); return sheet; }
}

const cacheValues = new Map();
const scriptCache = {
    get: (key) => cacheValues.has(key) ? cacheValues.get(key) : null,
    put: (key, value) => cacheValues.set(key, value),
    remove: (key) => cacheValues.delete(key),
};

const date = "19/07/69";
const order = new Sheet("ออเดอร์-19-07-69", [["ใบจัดออเดอร์"]]);
const stock = new Sheet("ของในสต็อก JUL", [
    ["", "", "", ""],
    ["", "", "", ""],
    ["", date, "", ""],
    ["", "แอปเปิล", 100, "กก"],
]);
const mainBook = new Book([order]);
const stockBook = new Book([stock]);
const context = {
    console: { log: console.log, error: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => "" }) },
    CacheService: { getScriptCache: () => scriptCache },
    SpreadsheetApp: { openById: () => stockBook, flush: () => {} },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    ContentService: {
        MimeType: { JSON: "JSON" },
        createTextOutput: () => ({ setMimeType() { return this; } }),
    },
    Utilities: {
        formatDate: () => "19/07/2026 12:00",
        DigestAlgorithm: { SHA_256: "SHA_256" },
        computeDigest: (_algorithm, value) => Array.from(crypto.createHash("sha256").update(value).digest()),
        base64EncodeWebSafe: (bytes) => Buffer.from(bytes).toString("base64url"),
    },
};
vm.createContext(context);
const appScriptFiles = [
    "00_Config.gs",
    "10_Core.gs",
    "20_Webhook.gs",
    "30_Parsing.gs",
    "40_Stock.gs",
    "50_Orders.gs",
    "60_Cutoff.gs",
    "90_TestHarness.gs",
];
const appScriptSource = appScriptFiles
    .map((file) => fs.readFileSync(`src/${file}`, "utf8"))
    .join("");
vm.runInContext(appScriptSource +
    "\nthis.api={CONFIG,withScriptLock,getStableHash,normalizeDelivDate,parseOrderMessage,processStockUpdate,applyStockUpdatePlan,processOrderUpdate,saveEventJournal,loadEventJournal,updateDeliverySheet,getOrderRounds,findStoreOrders,getOrderSnapshotHash,formatCancelPreview,cancelOrderIfSnapshotMatches,updatePurchaseSummarySheet,prepareStockDeduction,applyStockDeductionPlan,deductStockForCutoff,createCutoffRoundDivider,executeCutoff,saveCutoffJournal,loadCutoffJournal,beginEventLog,setActiveLogContext,setActiveReplyState,finishActiveLog,recordEventFailure,replyToLine,getStoreMappingDictionary,getNoStockProductNames,getOrderMatchNotes,formatStoreMappingLog,onEdit,validateDirectOrderTestInput};", context);

const claspProject = JSON.parse(fs.readFileSync(".clasp.json", "utf8"));
assert.equal(claspProject.rootDir, "src");
assert.deepEqual(claspProject.filePushOrder, appScriptFiles);
assert.equal(context.api.CONFIG.SHEET_ID, "1IUpkB2Cs2cjXVoBm_d9OYiYxhxzz9ReWCftNOlKphVk");
assert.deepEqual(
    context.api.validateDirectOrderTestInput(
        "31/12/99\nTEST-Codex\nแครอท 0.123 กก",
        "test-direct-20260902-01",
    ),
    { eventId: "test-direct-20260902-01", deliveryDate: "31/12/99", storeName: "TEST-Codex" },
);
assert.throws(
    () => context.api.validateDirectOrderTestInput("31/12/99\nร้านจริง\nแครอท 1 กก", "test-direct-20260902-02"),
    /TEST-/,
);
assert.throws(
    () => context.api.validateDirectOrderTestInput("02/09/69\nTEST-Codex\nแครอท 1 กก", "test-direct-20260902-03"),
    /31\/12\/99/,
);
assert.throws(
    () => context.api.validateDirectOrderTestInput("ตัดรอบ 31/12/99", "test-direct-20260902-04"),
    /คำสั่งแอดมิน/,
);

function purchaseRow(round) {
    const rows = mainBook.getSheetByName("ใบซื้อ-19-07-69").getDataRange().getValues();
    const header = rows.findIndex((row) => row[0] === `รอบที่ ${round}`);
    return rows[header + 1];
}

for (let round = 1; round <= 6; round++) {
    order.rows.push(
        ["ลำดับที่ 1", "", "", "", "", ""],
        ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
        [date, "ร้านทดสอบ", "แอปเปิล", round, "กก.", ""],
    );

    context.api.updatePurchaseSummarySheet(mainBook, date);
    stock.rows[3][2] += 10; // สต๊อกเปลี่ยนหลังใบซื้อล่าสุด แต่ก่อนกดตัดรอบ
    const stockBeforeCutoff = stock.rows[3][2];

    const cutoff = context.api.executeCutoff(mainBook, date, `evt-round-${round}`);
    const result = cutoff.deductRes;
    assert.equal(cutoff.res.success, true);
    assert.equal(result.success, true);
    assert.equal(result.deductedCount, 1);
    assert.equal(result.deductedByUnit["กก"], round);

    const row = purchaseRow(round);
    assert.equal(row[3], stockBeforeCutoff - round);
    assert.equal(row[5], round);
    assert.equal(stock.rows[3][2], stockBeforeCutoff - round);
}

assert.throws(
    () => context.api.updatePurchaseSummarySheet(mainBook, "20/07/69", false, true),
    /ไม่พบส่วนสต๊อกวันที่ 20\/07\/69/,
);

// ออเดอร์ปกติต้องยังสร้างใบซื้อได้ แม้หัววันที่ในไฟล์สต๊อกยังไม่มี
mainBook.sheets.set("ออเดอร์-20-07-69", new Sheet("ออเดอร์-20-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["20/07/69", "ร้านทดสอบ", "แอปเปิล", 2, "กก.", ""],
]));
context.api.updatePurchaseSummarySheet(mainBook, "20/07/69");
assert.ok(mainBook.getSheetByName("ใบซื้อ-20-07-69"));

// สต๊อกไม่พอต้องรายงานยอดที่หักได้และยอดขาด ไม่รวมเป็นรายการที่ไม่ได้หัก
stock.rows.push(
    ["", "21/07/69", "", ""],
    ["", "แอปเปิล", 10, "กก"],
);
mainBook.sheets.set("ออเดอร์-21-07-69", new Sheet("ออเดอร์-21-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["21/07/69", "ร้านทดสอบ", "แอปเปิล", 12, "กก.", ""],
]));
const partialResult = context.api.deductStockForCutoff(mainBook, "21/07/69");
assert.equal(partialResult.fullDeductedCount, 0);
assert.equal(partialResult.partialItems.length, 1);
assert.equal(partialResult.partialItems[0].deducted, 10);
assert.equal(partialResult.partialItems[0].shortage, 2);
assert.equal(partialResult.deductedItems[0].name, "แอปเปิล");
assert.equal(partialResult.deductedItems[0].deducted, 10);
assert.equal(partialResult.unmatchedCount, 0);
assert.equal(partialResult.zeroStockCount, 0);

// รายการสต๊อกศูนย์และหน่วยไม่ตรงต้องส่งชื่อกลับไปแสดงใน LINE
stock.rows.push(
    ["", "22/07/69", "", ""],
    ["", "แอปเปิล", 0, "กก"],
);
mainBook.sheets.set("ออเดอร์-22-07-69", new Sheet("ออเดอร์-22-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["22/07/69", "ร้านทดสอบ", "แอปเปิล", 2, "กก.", ""],
    ["22/07/69", "ร้านทดสอบ", "แอปเปิล", 3, "ลูก", ""],
]));
const unavailableResult = context.api.deductStockForCutoff(mainBook, "22/07/69");
assert.equal(unavailableResult.zeroStockItems[0].name, "แอปเปิล");
assert.equal(unavailableResult.zeroStockItems[0].requested, 2);
assert.equal(unavailableResult.unmatchedItems[0].name, "แอปเปิล");
assert.match(unavailableResult.unmatchedItems[0].reason, /หน่วยสต๊อกเป็น กก/);

const logSheet = new Sheet("Logs", [[new Date(), "user", "message", "Processing"]]);
context.api.setActiveLogContext(logSheet, 1);
context.api.finishActiveLog("COMPLETE", "LINE Reply Success");
assert.equal(logSheet.rows[0][3], "COMPLETE; LINE Reply Success");

// รับเฉพาะวันที่จริง และ normalize ปี ค.ศ./พ.ศ. ให้เป็นรูปแบบเดิม DD/MM/YY
assert.equal(context.api.normalizeDelivDate("0/07/69"), null);
assert.equal(context.api.normalizeDelivDate("31/02/69"), null);
assert.equal(context.api.normalizeDelivDate("29/02/68"), null);
assert.equal(context.api.normalizeDelivDate("19/07/2569"), "19/07/69");
assert.equal(context.api.normalizeDelivDate("19/07/2026"), "19/07/69");

// parser ต้องไม่รับวันที่ปลอม จำนวนติดลบ หรือจำนวนศูนย์
const exactMapping = { "แอปเปิล": { name: "แอปเปิล", unit: "กก", noStock: false } };
const fuzzyParsedOrder = context.api.parseOrderMessage(
    "09/08/69\nร้านทดสอบ\nแตงกวาญี่ปุน 0.4 กก",
    { "แตงกวาญี่ปุ่น": { name: "แตงกวาญี่ปุ่น | JAPANESE CUCUMBER", unit: "กก", noStock: false } },
);
assert.equal(fuzzyParsedOrder.items[0].name, "แตงกวาญี่ปุ่น | JAPANESE CUCUMBER");
assert.equal(fuzzyParsedOrder.items[0].isFuzzy, true);
const ambiguousFuzzyOrder = context.api.parseOrderMessage(
    "09/08/69\nร้านทดสอบ\nabf 1 กก",
    {
        abd: { name: "สินค้า A", unit: "กก", noStock: false },
        abe: { name: "สินค้า B", unit: "กก", noStock: false },
    },
);
assert.equal(ambiguousFuzzyOrder.items[0].name, "abf");
assert.equal(ambiguousFuzzyOrder.items[0].isWarning, true);
const prototypeProductOrder = context.api.parseOrderMessage(
    "09/08/69\nร้านทดสอบ\n__proto__ 1 กก\nconstructor 1 กก",
    {},
);
assert.deepEqual(prototypeProductOrder.items.map((item) => item.name), ["__proto__", "constructor"]);
assert.ok(prototypeProductOrder.items.every((item) => item.isWarning && !item.isFuzzy));

// StoreMapping: exact ต่อร้าน, validation, cache, fallback และ UX ต้องเป็นไปตาม requirement
const storeMappingBook = new Book([
    new Sheet("Mapping", [
        ["คำที่ลูกค้าพิมพ์", "ชื่อมาตรฐานในระบบ", "หน่วยมาตรฐาน"],
        ["apple", "แอปเปิลมาตรฐาน", "กก"],
        ["", "สินค้าไม่มี alias", "กล่อง", "x"],
        ["oil kg", "น้ำมันมาตรฐาน", "กก"],
        ["oil gallon", "น้ำมันมาตรฐาน", "แกลลอน"],
    ]),
    new Sheet("StoreMapping", [
        ["ชื่อร้าน", "ชื่อสินค้าใน PO", "ชื่อมาตรฐานในระบบ"],
        ["Dusit Thani", "PO APPLE", "แอปเปิลมาตรฐาน"],
        ["Dusit Thani", "PO BOX", "สินค้าไม่มี alias"],
        ["Dusit Thani", "PO BAD", "ไม่มีจริง"],
        ["Dusit Thani", "PO DUP", "แอปเปิลมาตรฐาน"],
        ["DusitThani", "PO DUP", "สินค้าไม่มี alias"],
        ["Dusit Thani", "PO OIL", "น้ำมันมาตรฐาน"],
        ["Rhapsody", "PO APPLE", "แอปเปิลมาตรฐาน"],
        ["constructor", "__proto__", "แอปเปิลมาตรฐาน"],
    ]),
]);
cacheValues.clear();
const storeMappingDict = context.api.getStoreMappingDictionary(storeMappingBook);
assert.equal(storeMappingBook.sheetReadCounts.StoreMapping, 1);
assert.equal(storeMappingBook.sheetReadCounts.Mapping, 1);
assert.equal(storeMappingDict["ชื่อร้าน"], undefined); // header ต้องไม่ถูกโหลด
const cachedStoreMappingDict = context.api.getStoreMappingDictionary(storeMappingBook);
assert.equal(storeMappingBook.sheetReadCounts.StoreMapping, 1); // cache hit ต้องไม่อ่านชีตซ้ำ
assert.deepEqual(context.api.getNoStockProductNames(storeMappingBook), ["สินค้าไม่มี alias"]);

const prototypeKeyOrder = context.api.parseOrderMessage(
    "09/08/69 constructor\n__proto__ 1 กก",
    {},
    cachedStoreMappingDict,
);
assert.equal(prototypeKeyOrder.items[0].name, "แอปเปิลมาตรฐาน");
assert.equal(prototypeKeyOrder.items[0].mappingSource, "STORE_MAPPING");

const storeMappingFallback = {
    fallback: { name: "สินค้าจาก Mapping", unit: "กก", noStock: false },
    pobad: { name: "สินค้าที่ห้าม fallback", unit: "กก", noStock: false },
};
const absentPrototypeStoreOrder = context.api.parseOrderMessage(
    "09/08/69 __proto__\nfallback 1 กก",
    storeMappingFallback,
    {},
);
assert.equal(absentPrototypeStoreOrder.items[0].name, "สินค้าจาก Mapping");
assert.equal(absentPrototypeStoreOrder.storeMappingSummary.configured, false);
const storeExactOrder = context.api.parseOrderMessage(
    "09/08/69\nDUSIT THANI\npo apple 2 กก",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(storeExactOrder.items[0].name, "แอปเปิลมาตรฐาน");
assert.equal(storeExactOrder.items[0].isFuzzy, false);
assert.equal(storeExactOrder.items[0].isWarning, false);
assert.equal(storeExactOrder.items[0].mappingSource, "STORE_MAPPING");
assert.equal(storeExactOrder.storeMappingSummary.matched, 1);
assert.doesNotMatch(context.api.getOrderMatchNotes(storeExactOrder.items), /🔍|⚠️/);

const blankAliasCanonicalOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO BOX 3",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(blankAliasCanonicalOrder.items[0].name, "สินค้าไม่มี alias");
assert.equal(blankAliasCanonicalOrder.items[0].unit, "กล่อง");

const storeMissOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nfallback 1 กก",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(storeMissOrder.items[0].name, "สินค้าจาก Mapping");
assert.equal(storeMissOrder.items[0].mappingSource, "MAPPING_EXACT");
assert.equal(storeMissOrder.storeMappingSummary.missed, 1);
assert.match(context.api.formatStoreMappingLog(storeMissOrder.storeMappingSummary), /Missed 1/);

const outsideStoreWithMapping = context.api.parseOrderMessage(
    "09/08/69\nร้านทั่วไป\nfallback 1 กก",
    storeMappingFallback,
    storeMappingDict,
);
const outsideStoreWithoutMapping = context.api.parseOrderMessage(
    "09/08/69\nร้านทั่วไป\nfallback 1 กก",
    storeMappingFallback,
);
assert.equal(JSON.stringify(outsideStoreWithMapping), JSON.stringify(outsideStoreWithoutMapping));

const invalidCanonicalOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO BAD 1 กก",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(invalidCanonicalOrder.items[0].name, "PO BAD");
assert.equal(invalidCanonicalOrder.items[0].isWarning, true);
assert.equal(invalidCanonicalOrder.items[0].mappingSource, "STORE_MAPPING_INVALID");
assert.equal(invalidCanonicalOrder.storeMappingSummary.invalid, 1);
assert.match(invalidCanonicalOrder.storeMappingSummary.errors[0].reason, /ไม่มีใน Mapping/);
assert.match(context.api.getOrderMatchNotes(invalidCanonicalOrder.items), /StoreMapping ไม่ถูกต้อง/);
assert.doesNotMatch(context.api.getOrderMatchNotes(invalidCanonicalOrder.items), /ไม่พบใน Mapping/);

const duplicateConflictOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO DUP 1 กก",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(duplicateConflictOrder.items[0].mappingSource, "STORE_MAPPING_INVALID");
assert.match(duplicateConflictOrder.items[0].mappingError, /ซ้ำ/);

const ambiguousUnitOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO OIL 2",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(ambiguousUnitOrder.items[0].mappingSource, "STORE_MAPPING_INVALID");
assert.match(ambiguousUnitOrder.items[0].mappingError, /หลายหน่วย/);
const selectedUnitOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO OIL 2 แกลลอน",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(selectedUnitOrder.items[0].mappingSource, "STORE_MAPPING");
assert.equal(selectedUnitOrder.items[0].unit, "แกลลอน");

const splitStoreOrder = context.api.parseOrderMessage(
    "09/08/69\nDusit Thani\nPO APPLE 1 กก + 2 ถุง",
    storeMappingFallback,
    storeMappingDict,
);
assert.equal(splitStoreOrder.items.length, 2);
assert.ok(splitStoreOrder.items.every((item) => item.mappingSource === "STORE_MAPPING" && !item.isFuzzy && !item.isWarning));

assert.match(
    context.api.formatStoreMappingLog(invalidCanonicalOrder.storeMappingSummary),
    /Dusit Thani \/ PO BAD: ชื่อมาตรฐานไม่มีใน Mapping คอลัมน์ B/,
);
const cappedErrorLog = context.api.formatStoreMappingLog({
    configured: true,
    matched: 0,
    missed: 0,
    invalid: 25,
    errors: Array.from({ length: 25 }, (_, i) => ({ storeName: "Dusit Thani", poName: `PO ${i + 1}`, reason: "ผิด" })),
});
assert.match(cappedErrorLog, /\+5 more/);
assert.doesNotMatch(cappedErrorLog, /PO 25/);

context.api.onEdit({ range: { getSheet: () => storeMappingBook.getSheetByName("StoreMapping") } });
assert.equal(cacheValues.has("store_mapping_dict"), false);
context.api.getStoreMappingDictionary(storeMappingBook);
assert.equal(storeMappingBook.sheetReadCounts.StoreMapping, 3); // onEdit lookup 1 + loader read 1

const noStoreMappingBook = new Book([storeMappingBook.getSheetByName("Mapping")]);
cacheValues.clear();
assert.deepEqual(context.api.getStoreMappingDictionary(noStoreMappingBook), {});
context.api.getStoreMappingDictionary(noStoreMappingBook);
assert.equal(noStoreMappingBook.sheetReadCounts.StoreMapping, 1);

assert.equal(context.api.parseOrderMessage("31/02/69\nร้านทดสอบ\nแอปเปิล 2 กก", exactMapping), null);
assert.equal(context.api.parseOrderMessage("19/07/69\nร้านทดสอบ\nแอปเปิล -2 กก", exactMapping), null);
assert.equal(context.api.parseOrderMessage("19/07/69\nร้านทดสอบ\nแอปเปิล 0 กก", exactMapping), null);
assert.equal(context.api.processStockUpdate("เติมสต๊อก 31/02/69\nแอปเปิล 2 กก", mainBook, exactMapping).success, false);
const beforeWrongUnit = stock.rows[3][2];
const wrongUnitUpdate = context.api.processStockUpdate("เติมสต๊อก 19/07/69\nแอปเปิล 2 ลูก", mainBook, exactMapping);
assert.equal(wrongUnitUpdate.success, false);
assert.match(wrongUnitUpdate.message, /หน่วยที่เติมเป็น ลูก แต่หน่วยสต๊อกเป็น กก/);
assert.equal(stock.rows[3][2], beforeWrongUnit);
const mixedNewUnits = context.api.processStockUpdate(
    "เติมสต๊อก 19/07/69\nส้ม 1 กก\nส้ม 2 ลูก",
    mainBook,
    exactMapping,
);
assert.equal(mixedNewUnits.success, false);
assert.match(mixedNewUnits.message, /มีหลายหน่วยในคำสั่งเดียวกัน/);
const combinedStockUpdate = context.api.processStockUpdate(
    "เติมสต๊อก 19/07/69\nแอปเปิล 1 กก\nแอปเปิล 2 กก",
    mainBook,
    exactMapping,
);
assert.equal(combinedStockUpdate.success, true);
assert.equal(combinedStockUpdate.count, 2);
assert.equal(stock.rows[3][2], beforeWrongUnit + 3);

// เติมสต๊อก event เดิมต้องบวกเพียงครั้งเดียว
const beforeStockEvent = stock.rows[3][2];
const stockEventMessage = "เติมสต๊อก 19/07/69\nแอปเปิล 4 กก";
assert.equal(context.api.processStockUpdate(stockEventMessage, mainBook, exactMapping, "evt-stock-once").success, true);
assert.equal(context.api.processStockUpdate(stockEventMessage, mainBook, exactMapping, "evt-stock-once").success, true);
assert.equal(stock.rows[3][2], beforeStockEvent + 4);
assert.equal(context.api.loadEventJournal(mainBook, "evt-stock-once|STOCK_UPDATE").status, "COMPLETE");
assert.equal(context.api.processStockUpdate(stockEventMessage, mainBook, exactMapping, "evt-stock-intentional-new").success, true);
assert.equal(stock.rows[3][2], beforeStockEvent + 8);

// Resume เติมสต๊อกหลังเขียนแถวแรกแล้ว ต้องไม่บวกแถวแรกซ้ำ
stock.rows.push(
    ["", "31/07/69", "", ""],
    ["", "แอปเปิล", 12, "กก"], // จำลองว่าเขียนจาก 10 เป็น 12 สำเร็จก่อน process ล้ม
    ["", "กล้วย", 20, "กก"],
);
const stockResumeMessage = "เติมสต๊อก 31/07/69\nแอปเปิล 2 กก\nกล้วย 3 กก";
const stockResumePayload = {
    deliveryDate: "31/07/69",
    count: 2,
    commandHash: context.api.getStableHash(stockResumeMessage),
    changes: [
        { lookupKey: "แอปเปิล", name: "แอปเปิล", unit: "กก", wasExisting: true, before: 10, after: 12 },
        { lookupKey: "กล้วย", name: "กล้วย", unit: "กก", wasExisting: true, before: 20, after: 23 },
    ],
};
context.api.saveEventJournal(mainBook, "evt-stock-resume|STOCK_UPDATE", "STOCK_UPDATE", "STARTED", stockResumePayload, "simulated crash");
const resumedStockUpdate = context.api.processStockUpdate(
    stockResumeMessage,
    mainBook,
    exactMapping,
    "evt-stock-resume-new-id",
);
assert.equal(resumedStockUpdate.success, true);
assert.equal(stock.rows[stock.rows.length - 2][2], 12);
assert.equal(stock.rows[stock.rows.length - 1][2], 23);
assert.equal(context.api.processStockUpdate(
    stockResumeMessage,
    mainBook,
    exactMapping,
    "evt-stock-resume-new-id",
).success, true);
assert.equal(stock.rows[stock.rows.length - 2][2], 12);
assert.equal(stock.rows[stock.rows.length - 1][2], 23);

// สถานะจับคู่เก็บในคอลัมน์ G ที่ซ่อนอยู่ จึงไม่เพิ่มคอลัมน์ใน UX เดิม
const deliveryBook = new Book();
context.api.updateDeliverySheet(deliveryBook, "27/07/69", "ร้านทดสอบ", [
    { name: "แอปเปิล", amount: 1, unit: "กก", inputNote: "", isFuzzy: false, isWarning: false },
    { name: "กล้วย", amount: 2, unit: "กก", inputNote: "", isFuzzy: true, isWarning: false },
    { name: "ชื่อใหม่", amount: 3, unit: "กก", inputNote: "", isFuzzy: false, isWarning: true },
], new Date());
const deliverySheet = deliveryBook.getSheetByName("ออเดอร์-27-07-69");
assert.equal(deliverySheet.hiddenColumns.has(7), true);
assert.deepEqual(deliverySheet.rows.slice(-3).map((row) => row[6]), ["EXACT", "FUZZY", "UNMAPPED"]);

// StoreMapping exact ต้องตัดสต๊อกผ่าน flow เดิม; config ผิดต้องเป็น UNMAPPED และไม่ถูกตัด
const storeCutoffDate = "10/08/69";
const storeCutoffStock = new Sheet("ของในสต็อก AUG", [
    ["", storeCutoffDate, "", ""],
    ["", "แอปเปิลมาตรฐาน", 10, "กก"],
    ["", "PO BAD", 10, "กก"],
]);
stockBook.sheets.set(storeCutoffStock.name, storeCutoffStock);
const storeCutoffBook = new Book();
const storeCutoffOrder = context.api.parseOrderMessage(
    `${storeCutoffDate}\nDusit Thani\nPO APPLE 2 กก\nPO BAD 3 กก`,
    storeMappingFallback,
    storeMappingDict,
);
context.api.updateDeliverySheet(
    storeCutoffBook,
    storeCutoffDate,
    storeCutoffOrder.storeName,
    storeCutoffOrder.items,
    new Date(),
);
const storeCutoffSheet = storeCutoffBook.getSheetByName("ออเดอร์-10-08-69");
assert.deepEqual(storeCutoffSheet.rows.slice(-2).map((row) => row[6]), ["EXACT", "UNMAPPED"]);
const storeCutoffResult = context.api.deductStockForCutoff(storeCutoffBook, storeCutoffDate);
assert.equal(storeCutoffResult.deductedCount, 1);
assert.equal(storeCutoffResult.unmatchedCount, 1);
assert.equal(storeCutoffStock.rows[1][2], 8);
assert.equal(storeCutoffStock.rows[2][2], 10);

// ออเดอร์ที่ค้างต้อง resume ข้าม Event ID และ retry Event ID ใหม่ซ้ำได้โดยไม่เพิ่มรายการ
const retryOrderBook = new Book();
const retryItems = [
    { name: "สินค้า A", amount: 1, unit: "กก", inputNote: "", isFuzzy: false, isWarning: false },
    { name: "สินค้า B", amount: 2, unit: "กก", inputNote: "", isFuzzy: false, isWarning: false },
    { name: "สินค้า C", amount: 3, unit: "กก", inputNote: "", isFuzzy: false, isWarning: false },
];
const retryOrderMessage = "27/07/69\nร้าน retry\nสินค้า A 1 กก\nสินค้า B 2 กก\nสินค้า C 3 กก";
const retryOrderPayload = {
    commandHash: context.api.getStableHash(retryOrderMessage),
    deliveryDate: "27/07/69",
    storeName: "ร้าน retry",
    items: retryItems,
    timestamp: new Date().toISOString(),
    rowEventId: "evt-order-original",
};
context.api.saveEventJournal(
    retryOrderBook,
    "evt-order-original|ORDER",
    "ORDER",
    "STARTED",
    retryOrderPayload,
    "simulated crash",
);
context.api.updateDeliverySheet(
    retryOrderBook,
    "27/07/69",
    "ร้าน retry",
    retryItems.slice(0, 2),
    new Date(),
    "evt-order-original",
);
const retryOrderData = { deliveryDate: "27/07/69", storeName: "ร้าน retry", items: retryItems };
context.api.processOrderUpdate(retryOrderMessage, retryOrderBook, retryOrderData, new Date(), "evt-order-new");
context.api.processOrderUpdate(retryOrderMessage, retryOrderBook, retryOrderData, new Date(), "evt-order-new");
const retryOrderRows = retryOrderBook.getSheetByName("ออเดอร์-27-07-69").rows
    .filter((row) => row[1] === "ร้าน retry");
assert.equal(retryOrderRows.length, 3);
assert.deepEqual(retryOrderRows.map((row) => row[8]), [1, 2, 3]);
assert.equal(retryOrderBook.getSheetByName("ออเดอร์-27-07-69").hiddenColumns.has(9), true);
assert.equal(context.api.loadEventJournal(retryOrderBook, "evt-order-original|ORDER").status, "COMPLETE");
assert.equal(context.api.loadEventJournal(retryOrderBook, "evt-order-new|ORDER").status, "COMPLETE");

// ถ้าออเดอร์เปลี่ยนหลัง preview ต้องไม่ลบตามเลขรายการเดิม
const cancelBook = new Book();
context.api.updateDeliverySheet(cancelBook, "30/07/69", "ร้านยกเลิก", retryItems, new Date(), "evt-cancel-source");
const cancelSheet = cancelBook.getSheetByName("ออเดอร์-30-07-69");
const cancelPreviewItems = context.api.findStoreOrders(cancelSheet, "ร้านยกเลิก");
const pendingCancel = {
    deliveryDate: "30/07/69",
    storeName: "ร้านยกเลิก",
    items: cancelPreviewItems,
    snapshotHash: context.api.getOrderSnapshotHash(cancelPreviewItems),
};
cancelSheet.getRange(cancelPreviewItems[0].row, 4).setValue(99);
const staleCancel = context.api.cancelOrderIfSnapshotMatches(cancelBook, pendingCancel, [1]);
assert.equal(staleCancel.success, false);
assert.match(staleCancel.message, /รายการออเดอร์เปลี่ยน/);
assert.equal(staleCancel.stale, true);
assert.equal(staleCancel.items[0].amount, 99);
assert.match(context.api.formatCancelPreview("ร้านยกเลิก", "30/07/69", staleCancel.items, "รายการล่าสุด"), /รายการล่าสุด/);
assert.equal(context.api.findStoreOrders(cancelSheet, "ร้านยกเลิก").length, 3);
const refreshedCancelItems = context.api.findStoreOrders(cancelSheet, "ร้านยกเลิก");
const validCancel = context.api.cancelOrderIfSnapshotMatches(cancelBook, {
    ...pendingCancel,
    snapshotHash: context.api.getOrderSnapshotHash(refreshedCancelItems),
}, [2]);
assert.equal(validCancel.success, true);
assert.equal(validCancel.deletedCount, 1);
assert.equal(context.api.findStoreOrders(cancelSheet, "ร้านยกเลิก").length, 2);

// ใบออเดอร์เก่า 6 คอลัมน์ยังอ่านและตัดรอบต่อได้หลัง deploy
const legacyOrderSheet = new Sheet("ออเดอร์-เก่า", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["27/07/69", "ร้านเดิม", "แอปเปิล", 1, "กก.", ""],
], 6);
const legacyRounds = context.api.getOrderRounds(legacyOrderSheet);
assert.equal(legacyRounds[0].products["แอปเปิล|กก."].total, 1);

// PROCESSING ใหม่ยังไม่ทำซ้ำ; FAILED หรือ PROCESSING ที่ค้างต้อง retry; COMPLETE เท่านั้นที่ทิ้ง
const dedupeBook = new Book();
const webhookEvent = { webhookEventId: "evt-001", message: { id: "msg-001" }, source: { userId: "user-1" } };
const firstEvent = context.api.beginEventLog(dedupeBook, webhookEvent, "ตัดรอบ 19/07/69", new Date());
const secondEvent = context.api.beginEventLog(dedupeBook, webhookEvent, "ตัดรอบ 19/07/69", new Date());
assert.equal(firstEvent.duplicate, false);
assert.equal(secondEvent.duplicate, true);
assert.equal(secondEvent.inProgress, true);
assert.equal(dedupeBook.getSheetByName("Logs").getLastRow(), 1);
dedupeBook.getSheetByName("Logs").getRange(1, 1).setValue(new Date(Date.now() - 10 * 60 * 1000));
const staleEvent = context.api.beginEventLog(dedupeBook, webhookEvent, "ตัดรอบ 19/07/69", new Date());
assert.equal(staleEvent.duplicate, false);
assert.equal(staleEvent.retry, true);
dedupeBook.getSheetByName("Logs").getRange(1, 4).setValue("FAILED; LINE Reply Failed");
const failedEvent = context.api.beginEventLog(dedupeBook, webhookEvent, "ตัดรอบ 19/07/69", new Date());
assert.equal(failedEvent.duplicate, false);
assert.equal(failedEvent.retry, true);
dedupeBook.getSheetByName("Logs").getRange(1, 4).setValue("COMPLETE; LINE Reply Success");
const completedEvent = context.api.beginEventLog(dedupeBook, webhookEvent, "ตัดรอบ 19/07/69", new Date());
assert.equal(completedEvent.duplicate, true);

// LINE reply ล้มแล้ว retry ต้องส่งข้อความเดิมอย่างเดียว ไม่รันคำสั่งตัดรอบซ้ำ
const replyRetryBook = new Book();
const replyRetryEvent = {
    webhookEventId: "evt-reply-retry",
    replyToken: "reply-token",
    message: { id: "msg-reply-retry", type: "text", text: "ตัดรอบ 19/07/69" },
    source: { userId: "user-reply" },
};
const firstReplyLog = context.api.beginEventLog(replyRetryBook, replyRetryEvent, replyRetryEvent.message.text, new Date());
context.api.setActiveLogContext(firstReplyLog.logSheet, firstReplyLog.logRow);
const originalPropertiesService = context.PropertiesService;
const originalUrlFetchApp = context.UrlFetchApp;
const originalReplyOpenById = context.SpreadsheetApp.openById;
context.PropertiesService = { getScriptProperties: () => ({ getProperty: () => "test-token" }) };
context.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 500, getContentText: () => "simulated LINE failure" }) };
const originalReplyMessage = "✅ ตัดรอบสำเร็จ — ข้อความเดิม";
assert.equal(context.api.replyToLine(replyRetryEvent.replyToken, originalReplyMessage), false);
assert.equal(replyRetryBook.getSheetByName("Logs").rows[0][5], "FAILED");
assert.equal(replyRetryBook.getSheetByName("Logs").rows[0][6], originalReplyMessage);

context.SpreadsheetApp.openById = () => replyRetryBook;
context.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => "{}" }) };
context.handleTextEvent(replyRetryEvent);
assert.equal(replyRetryBook.getSheetByName("Logs").rows[0][3].startsWith("COMPLETE"), true);
assert.equal(replyRetryBook.getSheetByName("Logs").rows[0][5], "COMPLETE");
assert.equal(replyRetryBook.getSheetByName("Logs").rows[0][6], originalReplyMessage);
assert.equal(replyRetryBook.getSheetByName("_CutoffJournal"), null);
replyRetryBook.getSheetByName("Logs").getRange(1, 4).setValue("PROCESSING");
assert.equal(
    context.api.beginEventLog(replyRetryBook, replyRetryEvent, replyRetryEvent.message.text, new Date()).duplicate,
    true,
);
context.PropertiesService = originalPropertiesService;
context.UrlFetchApp = originalUrlFetchApp;
context.SpreadsheetApp.openById = originalReplyOpenById;
context.api.setActiveLogContext(null, 0);

// Lock timeout ต้องแยกจาก error ทั่วไปและมีข้อความให้ลองใหม่
const originalGetScriptLock = context.LockService.getScriptLock;
context.LockService.getScriptLock = () => ({ waitLock: () => { throw new Error("timeout"); }, releaseLock: () => {} });
assert.throws(
    () => context.api.withScriptLock(() => true),
    (err) => err.code === "LOCK_TIMEOUT" && /กรุณาลองใหม่/.test(err.message),
);
context.LockService.getScriptLock = originalGetScriptLock;

// Lock timeout ก่อนมี active log ต้องสร้างแถวที่ค้นหาได้
const originalOpenById = context.SpreadsheetApp.openById;
const lockLogBook = new Book();
context.SpreadsheetApp.openById = () => lockLogBook;
context.api.setActiveLogContext(null, 0);
context.api.recordEventFailure({ ...webhookEvent, message: { ...webhookEvent.message, text: "ทดสอบ lock" } }, "Lock Timeout");
assert.match(lockLogBook.getSheetByName("Logs").rows[0][3], /FAILED; Lock Timeout/);
context.SpreadsheetApp.openById = originalOpenById;

// event หนึ่งพังต้องไม่ขวาง event ถัดไปใน webhook เดียวกัน
const handledEvents = [];
const originalHandleTextEvent = context.handleTextEvent;
context.handleTextEvent = (event) => {
    handledEvents.push(event.webhookEventId);
    if (event.webhookEventId === "evt-fail") throw new Error("simulated event failure");
};
context.doPost({ postData: { contents: JSON.stringify({ events: [
    { type: "message", webhookEventId: "evt-fail", replyToken: "reply-1", message: { type: "text", text: "fail" }, source: {} },
    { type: "message", webhookEventId: "evt-next", replyToken: "reply-2", message: { type: "text", text: "next" }, source: {} },
] }) } });
assert.deepEqual(handledEvents, ["evt-fail", "evt-next"]);
context.handleTextEvent = originalHandleTextEvent;

// fuzzy ที่หา Mapping เจอต้องหักได้ รวมกับ exact; unmapped ยังต้องไม่หัก
stock.rows.push(
    ["", "23/07/69", "", ""],
    ["", "แอปเปิล", 9, "กก"],
    ["", "สินค้านอก Mapping", 7, "กก"],
);
const fuzzyAppleIndex = stock.rows.length - 2;
const unmappedStockIndex = stock.rows.length - 1;
mainBook.sheets.set("ออเดอร์-23-07-69", new Sheet("ออเดอร์-23-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", "", ""],
    ["23/07/69", "ร้านทดสอบ", "แอปเปิล", 4, "กก.", "", "FUZZY"],
    ["23/07/69", "ร้านทดสอบ", "แอปเปิล", 1, "กก.", "", "EXACT"],
    ["23/07/69", "ร้านทดสอบ", "แอปเปิล", 2, "ลูก", "", "FUZZY"],
    ["23/07/69", "ร้านทดสอบ", "สินค้านอก Mapping", 2, "กก.", "", "UNMAPPED"],
]));
const fuzzyResult = context.api.deductStockForCutoff(mainBook, "23/07/69");
assert.equal(fuzzyResult.success, true);
assert.equal(fuzzyResult.plan.length, 1);
assert.equal(fuzzyResult.deductedItems[0].deducted, 5);
assert.equal(fuzzyResult.unmatchedCount, 2);
assert.ok(fuzzyResult.unmatchedItems.some((item) => /หน่วยสต๊อกเป็น กก/.test(item.reason)));
assert.ok(fuzzyResult.unmatchedItems.some((item) => /ไม่พบชื่อใน Mapping/.test(item.reason)));
assert.equal(stock.rows[fuzzyAppleIndex][2], 4);
assert.equal(stock.rows[unmappedStockIndex][2], 7);

// regression: แตงกวาชื่อพิมพ์คลาดเคลื่อนแต่ Mapping เจอ ต้องใช้ 29 กก. และรายงานขาด 4 กก.
const fuzzyCucumberName = "แตงกวาญี่ปุ่น | JAPANESE CUCUMBER";
const augustStock = new Sheet("ของในสต็อก AUG", [
    ["", "09/08/69", "", ""],
    ["", fuzzyCucumberName, 29, "กก"],
]);
stockBook.sheets.set(augustStock.name, augustStock);
const fuzzyCucumberOrder = new Sheet("ออเดอร์-09-08-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", "", ""],
    ["09/08/69", "ร้านทดสอบ", fuzzyCucumberName, 33, "กก.", "", "FUZZY"],
]);
mainBook.sheets.set(fuzzyCucumberOrder.name, fuzzyCucumberOrder);
const fuzzyCucumberRound1 = context.api.executeCutoff(mainBook, "09/08/69", "evt-fuzzy-cucumber-1");
assert.equal(fuzzyCucumberRound1.deductRes.unmatchedCount, 0);
assert.equal(fuzzyCucumberRound1.deductRes.deductedItems[0].deducted, 29);
assert.equal(fuzzyCucumberRound1.deductRes.deductedItems[0].shortage, 4);
assert.equal(augustStock.rows[1][2], 0);
assert.equal(fuzzyCucumberOrder.rows.filter((row) => row[0] === "รอบ 2").length, 1);

// รอบถัดไปต้องไม่ทำให้สต๊อก 0 กลับมาเป็น 29
fuzzyCucumberOrder.rows.push(
    ["ลำดับที่ 1", "", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", "", ""],
    ["09/08/69", "ร้านทดสอบ", fuzzyCucumberName, 0.4, "กก.", "", "FUZZY"],
);
const fuzzyCucumberRound2 = context.api.executeCutoff(mainBook, "09/08/69", "evt-fuzzy-cucumber-2");
assert.equal(fuzzyCucumberRound2.deductRes.unmatchedCount, 0);
assert.equal(fuzzyCucumberRound2.deductRes.zeroStockItems[0].requested, 0.4);
assert.equal(augustStock.rows[1][2], 0);
assert.equal(fuzzyCucumberOrder.rows.filter((row) => row[0] === "รอบ 3").length, 1);

// สินค้าชื่อซ้ำใน section เดียวกันต้องหยุด ไม่เดาว่าจะหักแถวใด
stock.rows.push(
    ["", "24/07/69", "", ""],
    ["", "แอปเปิล", 5, "กก"],
    ["", "แอปเปิล", 6, "กก"],
);
mainBook.sheets.set("ออเดอร์-24-07-69", new Sheet("ออเดอร์-24-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["24/07/69", "ร้านทดสอบ", "แอปเปิล", 2, "กก.", ""],
]));
const duplicateStock = context.api.prepareStockDeduction(mainBook, "24/07/69");
assert.equal(duplicateStock.success, false);
assert.match(duplicateStock.message, /สินค้าซ้ำในสต๊อก/);

// apply เฉพาะเซลล์ที่เปลี่ยน: ข้อมูล/สูตรจำลองในแถวอื่นต้องไม่ถูกเขียนทับ
stock.rows.push(
    ["", "25/07/69", "", ""],
    ["", "แอปเปิล", 8, "กก"],
    ["", "สินค้าใช้สูตร", "=FORMULA()", "กก"],
);
const appleRow25 = stock.rows.length - 1;
context.api.applyStockDeductionPlan("25/07/69", [{ row: appleRow25, product: "แอปเปิล", before: 8, after: 5 }]);
assert.equal(stock.rows[appleRow25 - 1][2], 5);
assert.equal(stock.rows[appleRow25][2], "=FORMULA()");
assert.throws(
    () => context.api.applyStockDeductionPlan("25/07/69", [{ row: appleRow25, product: "แอปเปิล", before: 8, after: 3 }]),
    /ถูกแก้ระหว่างตัดรอบ/,
);
assert.equal(stock.rows[appleRow25 - 1][2], 5);

// Resume หลัง Apps Script ล้มกลางการหัก: แถวที่หักไปแล้วต้องไม่ถูกหักซ้ำ
stock.rows.push(
    ["", "26/07/69", "", ""],
    ["", "แอปเปิล", 10, "กก"],
    ["", "กล้วย", 20, "กก"],
);
mainBook.sheets.set("ออเดอร์-26-07-69", new Sheet("ออเดอร์-26-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["26/07/69", "ร้านทดสอบ", "แอปเปิล", 3, "กก.", ""],
    ["26/07/69", "ร้านทดสอบ", "กล้วย", 4, "กก.", ""],
]));
const resumePlan = context.api.prepareStockDeduction(mainBook, "26/07/69");
assert.equal(resumePlan.success, true);
assert.equal(resumePlan.plan.length, 2);
const appleChange = resumePlan.plan.find((item) => item.product === "แอปเปิล");
stock.getRange(appleChange.row, 3).setValue(appleChange.after); // จำลองว่าหักแถวแรกแล้ว process ตาย
context.api.saveCutoffJournal(mainBook, "26/07/69|1", "26/07/69", 1, "evt-resume", "STARTED", { deductRes: resumePlan }, "simulated crash");
const resumed = context.api.executeCutoff(mainBook, "26/07/69", "evt-resume-2");
assert.equal(resumed.deductRes.success, true);
assert.equal(stock.getRange(appleChange.row, 3).getValue(), 7);
const bananaChange = resumePlan.plan.find((item) => item.product === "กล้วย");
assert.equal(stock.getRange(bananaChange.row, 3).getValue(), 16);
assert.equal(mainBook.getSheetByName("_CutoffJournal").hidden, true);
assert.equal(context.api.loadCutoffJournal(mainBook, "26/07/69|1").status, "COMPLETE");
assert.equal(mainBook.getSheetByName("ออเดอร์-26-07-69").rows.filter((row) => row[0] === "รอบ 2").length, 1);

// Resume หลังสร้าง divider แล้วแต่ journal ยังไม่เดินต่อ: ต้องไม่สร้าง divider ซ้ำ
context.api.saveCutoffJournal(mainBook, "26/07/69|1", "26/07/69", 1, "evt-resume-3", "STOCK_DEDUCTED", { deductRes: resumePlan }, "simulated crash after divider");
const resumedAfterDivider = context.api.executeCutoff(mainBook, "26/07/69", "evt-resume-4");
assert.equal(resumedAfterDivider.res.success, true);
assert.equal(resumedAfterDivider.res.alreadyExists, true);
assert.equal(resumedAfterDivider.recoveredRound, undefined);
assert.equal(mainBook.getSheetByName("ออเดอร์-26-07-69").rows.filter((row) => row[0] === "รอบ 2").length, 1);

// ใบซื้อพังหลังปิดรอบ: journal ต้องค้าง ROUND_CLOSED และ retry เฉพาะใบซื้อ
stock.rows.push(
    ["", "29/07/69", "", ""],
    ["", "แอปเปิล", 10, "กก"],
);
const summaryRetryStockRow = stock.rows.length;
mainBook.sheets.set("ออเดอร์-29-07-69", new Sheet("ออเดอร์-29-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["29/07/69", "ร้านทดสอบ", "แอปเปิล", 3, "กก.", "", "EXACT"],
]));
const originalUpdatePurchaseSummary = context.updatePurchaseSummarySheet;
let summaryUpdateCalls = 0;
context.updatePurchaseSummarySheet = (...args) => {
    summaryUpdateCalls++;
    if (summaryUpdateCalls === 2) throw new Error("simulated summary failure");
    return originalUpdatePurchaseSummary(...args);
};
const summaryFailedCutoff = context.api.executeCutoff(mainBook, "29/07/69", "evt-summary-fail");
assert.match(summaryFailedCutoff.summaryError, /simulated summary failure/);
assert.equal(context.api.loadCutoffJournal(mainBook, "29/07/69|1").status, "ROUND_CLOSED");
assert.equal(stock.getRange(summaryRetryStockRow, 3).getValue(), 7);
const coercedJournalSheet = mainBook.getSheetByName("_CutoffJournal");
const coercedJournalRow = coercedJournalSheet.rows.findIndex((row) => row[0] === "29/07/69|1") + 1;
coercedJournalSheet.getRange(coercedJournalRow, 2).setValue(
    vm.runInContext("new Date(2026, 6, 29)", context),
); // จำลอง Google Sheets แปลงข้อความเป็น Date ใน Apps Script realm เดียวกัน
context.updatePurchaseSummarySheet = () => { throw new Error("simulated repeated summary failure"); };
const repeatedSummaryFailure = context.api.executeCutoff(mainBook, "29/07/69", "evt-summary-fail-again");
assert.equal(repeatedSummaryFailure.recoveryRetryFailed, true);
assert.equal(repeatedSummaryFailure.recoveredRound, 1);
assert.match(repeatedSummaryFailure.summaryError, /simulated repeated summary failure/);
assert.equal(context.api.loadCutoffJournal(mainBook, "29/07/69|1").status, "ROUND_CLOSED");
assert.equal(stock.getRange(summaryRetryStockRow, 3).getValue(), 7);
context.updatePurchaseSummarySheet = originalUpdatePurchaseSummary;
const summaryRetriedCutoff = context.api.executeCutoff(mainBook, "29/07/69", "evt-summary-retry");
assert.equal(summaryRetriedCutoff.summaryError, null);
assert.equal(summaryRetriedCutoff.recoveredOnly, true);
assert.equal(summaryRetriedCutoff.recoveredRound, 1);
assert.equal(context.api.loadCutoffJournal(mainBook, "29/07/69|1").status, "COMPLETE");
assert.equal(stock.getRange(summaryRetryStockRow, 3).getValue(), 7);
assert.equal(mainBook.getSheetByName("ออเดอร์-29-07-69").rows.filter((row) => row[0] === "รอบ 2").length, 1);
assert.equal(mainBook.getSheetByName("ออเดอร์-29-07-69").rows.filter((row) => row[0] === "รอบ 3").length, 0);

// ถ้ามีออเดอร์รอบใหม่แล้ว retry ROUND_CLOSED ต้องซ่อมรอบเก่าและตัดรอบใหม่ในคำสั่งเดียว
stock.rows.push(
    ["", "30/07/69", "", ""],
    ["", "แอปเปิล", 10, "กก"],
);
const recoveryFollowUpStockRow = stock.rows.length;
const recoveryFollowUpSheet = new Sheet("ออเดอร์-30-07-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["30/07/69", "ร้านทดสอบ", "แอปเปิล", 3, "กก.", "", "EXACT"],
]);
mainBook.sheets.set(recoveryFollowUpSheet.name, recoveryFollowUpSheet);
let recoverySummaryCalls = 0;
context.updatePurchaseSummarySheet = (...args) => {
    recoverySummaryCalls++;
    if (recoverySummaryCalls === 2) throw new Error("simulated recovery summary failure");
    return originalUpdatePurchaseSummary(...args);
};
const recoveryFailed = context.api.executeCutoff(mainBook, "30/07/69", "evt-recovery-fail");
assert.match(recoveryFailed.summaryError, /simulated recovery summary failure/);
assert.equal(stock.getRange(recoveryFollowUpStockRow, 3).getValue(), 7);
assert.equal(context.api.loadCutoffJournal(mainBook, "30/07/69|1").status, "ROUND_CLOSED");

recoveryFollowUpSheet.rows.push(
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    ["30/07/69", "ร้านทดสอบ", "แอปเปิล", 5, "กก.", "", "EXACT"],
);
let successfulRecoverySummaryCalls = 0;
context.updatePurchaseSummarySheet = (...args) => {
    successfulRecoverySummaryCalls++;
    return originalUpdatePurchaseSummary(...args);
};
const recoveredAndCut = context.api.executeCutoff(mainBook, "30/07/69", "evt-recovery-follow-up");
context.updatePurchaseSummarySheet = originalUpdatePurchaseSummary;
assert.equal(successfulRecoverySummaryCalls, 2); // ซ่อมรอบเก่า 1 ครั้ง + สรุปรอบใหม่หลังปิด 1 ครั้ง
assert.equal(recoveredAndCut.recoveredRound, 1);
assert.equal(recoveredAndCut.res.closedRound, 2);
assert.equal(recoveredAndCut.deductRes.deductedItems[0].deducted, 5);
assert.equal(stock.getRange(recoveryFollowUpStockRow, 3).getValue(), 2);
assert.equal(context.api.loadCutoffJournal(mainBook, "30/07/69|1").status, "COMPLETE");
assert.equal(context.api.loadCutoffJournal(mainBook, "30/07/69|2").status, "COMPLETE");
assert.equal(recoveryFollowUpSheet.rows.filter((row) => row[0] === "รอบ 2").length, 1);
assert.equal(recoveryFollowUpSheet.rows.filter((row) => row[0] === "รอบ 3").length, 1);
const recoveredStockAfterRetry = stock.getRange(recoveryFollowUpStockRow, 3).getValue();
const noDuplicateCutoff = context.api.executeCutoff(mainBook, "30/07/69", "evt-recovery-follow-up");
assert.equal(noDuplicateCutoff.deductRes.success, false);
assert.equal(stock.getRange(recoveryFollowUpStockRow, 3).getValue(), recoveredStockAfterRetry);

// ซ่อมรอบเก่าสำเร็จแต่เตรียมรอบใหม่ล้ม ต้องไม่รายงานว่ารอบใหม่สำเร็จและต้องไม่หักซ้ำ
const followUpFailureDate = "11/08/69";
const followUpFailureStock = stockBook.getSheetByName("ของในสต็อก AUG");
followUpFailureStock.rows.push(
    ["", followUpFailureDate, "", ""],
    ["", "แอปเปิล", 10, "กก"],
);
const followUpFailureStockRow = followUpFailureStock.rows.length;
const followUpFailureSheet = new Sheet("ออเดอร์-11-08-69", [
    ["ใบจัดออเดอร์"],
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    [followUpFailureDate, "ร้านทดสอบ", "แอปเปิล", 2, "กก.", "", "EXACT"],
]);
mainBook.sheets.set(followUpFailureSheet.name, followUpFailureSheet);
let initialFailureCalls = 0;
context.updatePurchaseSummarySheet = (...args) => {
    initialFailureCalls++;
    if (initialFailureCalls === 2) throw new Error("simulated initial summary failure");
    return originalUpdatePurchaseSummary(...args);
};
const initialFollowUpFailure = context.api.executeCutoff(mainBook, followUpFailureDate, "evt-follow-up-initial");
assert.match(initialFollowUpFailure.summaryError, /simulated initial summary failure/);
assert.equal(followUpFailureStock.getRange(followUpFailureStockRow, 3).getValue(), 8);
followUpFailureSheet.rows.push(
    ["ลำดับที่ 1", "", "", "", "", ""],
    ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", ""],
    [followUpFailureDate, "ร้านทดสอบ", "แอปเปิล", 3, "กก.", "", "EXACT"],
);
context.updatePurchaseSummarySheet = originalUpdatePurchaseSummary;
followUpFailureStock.rows.push(["", "แอปเปิล", 99, "กก"]); // ทำให้ prepare รอบใหม่ fail แบบไม่สร้าง journal
const failedFollowUp = context.api.executeCutoff(mainBook, followUpFailureDate, "evt-follow-up-retry");
assert.equal(failedFollowUp.recoveredOnly, true);
assert.equal(failedFollowUp.recoveredRound, 1);
assert.match(failedFollowUp.followUpError, /พบสินค้าซ้ำในสต๊อก/);
assert.equal(context.api.loadCutoffJournal(mainBook, `${followUpFailureDate}|1`).status, "COMPLETE");
assert.equal(context.api.loadCutoffJournal(mainBook, `${followUpFailureDate}|2`), null);
assert.equal(followUpFailureStock.getRange(followUpFailureStockRow, 3).getValue(), 8);
assert.equal(followUpFailureSheet.rows.filter((row) => row[0] === "รอบ 3").length, 0);

// Stress test: 10 รอบ × 20 ร้าน × 15 สินค้า = 3,000 รายการ
const stressStarted = Date.now();
const stressDate = "28/07/69";
const stressProducts = Array.from({ length: 15 }, (_, i) => `สินค้าทดสอบ ${String(i + 1).padStart(2, "0")}`);
const stressStockRows = {};
const expectedStressStock = {};
stock.rows.push(["", stressDate, "", ""]);
for (const product of stressProducts) {
    stock.rows.push(["", product, 100000, "กก"]);
    stressStockRows[product] = stock.rows.length;
    expectedStressStock[product] = 100000;
}

const stressOrder = new Sheet("ออเดอร์-28-07-69", [["ใบจัดออเดอร์"]]);
mainBook.sheets.set(stressOrder.name, stressOrder);

for (let round = 1; round <= 10; round++) {
    const expectedRoundTotals = Object.fromEntries(stressProducts.map((product) => [product, 0]));
    for (let store = 1; store <= 20; store++) {
        stressOrder.rows.push(
            [`ลำดับที่ ${store}`, "", "", "", "", "", ""],
            ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", "", ""],
        );
        stressProducts.forEach((product, productIndex) => {
            const qty = ((round + store + productIndex) % 7) + 1;
            expectedRoundTotals[product] += qty;
            stressOrder.rows.push([stressDate, `ร้านทดสอบ ${store}`, product, qty, "กก.", "", "EXACT"]);
        });
    }

    const cutoff = context.api.executeCutoff(mainBook, stressDate, `evt-stress-${round}`);
    assert.equal(cutoff.res.success, true);
    assert.equal(cutoff.res.closedRound, round);
    assert.equal(cutoff.deductRes.plan.length, stressProducts.length);
    assert.equal(cutoff.deductRes.fullDeductedCount, stressProducts.length);
    assert.equal(cutoff.deductRes.unmatchedCount, 0);
    assert.equal(cutoff.deductRes.zeroStockCount, 0);

    for (const product of stressProducts) {
        expectedStressStock[product] -= expectedRoundTotals[product];
        assert.equal(stock.getRange(stressStockRows[product], 3).getValue(), expectedStressStock[product]);
    }

    const dividerText = `รอบ ${round + 1}`;
    assert.equal(stressOrder.rows.filter((row) => row[0] === dividerText).length, 1);
    assert.equal(context.api.loadCutoffJournal(mainBook, `${stressDate}|${round}`).status, "COMPLETE");

    const summaryRows = mainBook.getSheetByName("ใบซื้อ-28-07-69").rows;
    const roundHeader = summaryRows.findIndex((row) => row[0] === `รอบที่ ${round}`);
    assert.ok(roundHeader >= 0);
    const summaryProducts = Object.fromEntries(
        summaryRows.slice(roundHeader + 1, roundHeader + 1 + stressProducts.length)
            .map((row) => [row[0], row[5]]),
    );
    for (const product of stressProducts) assert.equal(summaryProducts[product], expectedRoundTotals[product]);
}

assert.equal(stressOrder.rows.filter((row) => /^รอบ \d+$/.test(String(row[0]))).length, 10);
const stressElapsedMs = Date.now() - stressStarted;

console.log(`cutoff safety tests + stress 10 rounds/20 stores/15 products: ok (${stressElapsedMs} ms)`);
