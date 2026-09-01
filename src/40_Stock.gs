// ==============================================================================
// 📦 STOCK SYSTEM
// ==============================================================================
function applyStockUpdatePlan(ss, journalKey, payload) {
    const deliveryDate = payload.deliveryDate;
    try {
        const res = ensureStockDateSection(deliveryDate);
        const stockSheet = res.stockSheet;
        const section = res.section;
        const data = stockSheet.getDataRange().getValues();
        const rowMap = Object.create(null);
        for (let i = section.startIndex; i < section.endIndex; i++) {
            const key = String(data[i][1] || "").replace(/\s+/g, "").trim();
            if (!key) continue;
            if (rowMap[key]) throw new Error(`พบสินค้าซ้ำในสต๊อก: ${data[i][1]}`);
            rowMap[key] = { row: i + 1, qty: parseFloat(data[i][2]) || 0, unit: String(data[i][3] || "").trim() };
        }

        const today = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
        for (const change of payload.changes || []) {
            const current = rowMap[change.lookupKey];
            if (current) {
                if (standardizeUnit(current.unit) !== standardizeUnit(change.unit)) {
                    throw new Error(`${change.name}: หน่วยสต๊อกเปลี่ยนเป็น ${current.unit}`);
                }
                if (Math.abs(current.qty - change.after) < 1e-9) continue;
                if (!change.wasExisting || Math.abs(current.qty - change.before) >= 1e-9) {
                    throw new Error(`สต๊อก ${change.name} ถูกแก้ระหว่างเติมสต๊อก`);
                }
                stockSheet.getRange(current.row, 2, 1, 3).setValues([[change.name, change.after, change.unit]]);
                current.qty = change.after;
                continue;
            }

            if (change.wasExisting) throw new Error(`ไม่พบสินค้าเดิมระหว่างเติมสต๊อก: ${change.name}`);
            const insertRow = section.endIndex + 1;
            if (insertRow > stockSheet.getLastRow()) stockSheet.appendRow([today, change.name, change.after, change.unit]);
            else {
                stockSheet.insertRowBefore(insertRow);
                stockSheet.getRange(insertRow, 1, 1, 4).setValues([[today, change.name, change.after, change.unit]]);
            }
            rowMap[change.lookupKey] = { row: insertRow, qty: change.after, unit: change.unit };
            section.endIndex++;
        }
        SpreadsheetApp.flush();
        const result = { success: true, count: payload.count, deliveryDate };
        saveEventJournal(ss, journalKey, "STOCK_UPDATE", "COMPLETE", { ...payload, result }, "");
        return result;
    } catch (err) {
        saveEventJournal(ss, journalKey, "STOCK_UPDATE", "STARTED", payload, err.message);
        return { success: false, count: 0, deliveryDate, message: `เขียนสต๊อกไม่ครบ: ${err.message}` };
    }
}

function processStockUpdate(rawMessage, ss, mappingDict, eventId) {
    const hasDateToken = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/.test(String(rawMessage || ""));
    const parsedDate = parseDateFromText(rawMessage);
    if (hasDateToken && !parsedDate) {
        return { success: false, count: 0, deliveryDate: null, message: "วันที่ไม่ถูกต้อง" };
    }
    const deliveryDate = parsedDate || getTodayStr();
    const commandHash = eventId ? getStableHash(String(rawMessage || "").trim()) : "";
    const eventJournalKey = eventId ? `${eventId}|STOCK_UPDATE` : "";
    let journalKey = eventJournalKey;
    let journal = loadEventJournal(ss, journalKey);
    if (!journal && commandHash) {
        journal = loadPendingEventJournal(ss, "STOCK_UPDATE", commandHash);
        if (journal) journalKey = journal.key;
    }
    if (journal) {
        if (journal.status === "COMPLETE") return journal.payload.result;
        if (journal.status === "STARTED") {
            const sourceJournalKey = journalKey;
            if (eventJournalKey && eventJournalKey !== sourceJournalKey) {
                saveEventJournal(ss, eventJournalKey, "STOCK_UPDATE", "STARTED", journal.payload, "");
                journalKey = eventJournalKey;
            }
            const result = applyStockUpdatePlan(ss, journalKey, journal.payload);
            if (result.success && sourceJournalKey !== journalKey) {
                saveEventJournal(ss, sourceJournalKey, "STOCK_UPDATE", "COMPLETE", { ...journal.payload, result }, "");
            }
            return result;
        }
    }
    let stockSheet, section;
    try {
        const res = ensureStockDateSection(deliveryDate);
        stockSheet = res.stockSheet;
        section = res.section;
    } catch (err) {
        console.error("External Stock Error:", err.message);
        return { success: false, count: 0, deliveryDate, message: err.message };
    }
    if (!stockSheet || !section) return { success: false, count: 0, deliveryDate, message: "ไม่พบส่วนสต๊อกของวันที่ระบุ" };

    let text = normalizeUnits(rawMessage);
    const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
    const UNIT_LIST =
        "กก|กรัม|ขีด|แพ็ค|มัด|กล่อง|ถุง|ขวด|ลัง|ลูก|ผล|แผง|ห่อ|ตัว|ก้อน|ถาด|ม้วน|อัน|ชิ้น|กระบอก|กระป๋อง|กำ|ช่อ|เข่ง|โหล|หัว|เม็ด|ชุด|คู่|ซอง|ลิตร|มล|ฝัก|กระสอบ|ฟอง";
    const ITEM_RE = new RegExp(
        `^(.+?)\\s+(\\d+(?:\\.\\d+)?)\\s*(${UNIT_LIST})?\\s*(.*)$`,
        "i",
    );

    const existingData = stockSheet.getDataRange().getValues();
    const stockMap = Object.create(null);
    const stockQtyMap = Object.create(null);
    const stockUnitMap = Object.create(null);
    for (let i = section.startIndex; i < section.endIndex; i++) {
        const k = String(existingData[i][1]).replace(/\s+/g, "").trim();
        if (k) {
            if (stockMap[k]) {
                return { success: false, count: 0, deliveryDate, message: `พบสินค้าซ้ำในสต๊อก: ${existingData[i][1]}` };
            }
            stockMap[k] = i + 1;
            stockQtyMap[k] = parseFloat(existingData[i][2]) || 0;
            stockUnitMap[k] = String(existingData[i][3] || "").trim();
        }
    }

    const parsedItems = [];
    for (let i = 1; i < lines.length; i++) {
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(lines[i])) continue;
        const m = lines[i].match(ITEM_RE);
        if (!m) continue;

        const rawName = m[1].trim();
        const searchKey = rawName.replace(/\s+/g, "");
        const amount = parseFloat(m[2]);
        if (!Number.isFinite(amount) || amount <= 0) continue;
        const unit = m[3] ? m[3].trim() : null;
        const findResult = fuzzyFindMapping(searchKey, mappingDict);
        if (findResult && findResult.isFuzzy) {
            return { success: false, count: 0, deliveryDate, message: `ชื่อสินค้าไม่ตรง Mapping แบบชัดเจน: ${rawName}` };
        }
        const mapped = findResult ? findResult.mapped : null;
        const finalName = mapped ? mapped.name : rawName;
        const finalUnit = unit || (mapped && mapped.unit) || "หน่วย";
        const lookupKey = finalName.replace(/\s+/g, "");
        const existingUnit = stockUnitMap[lookupKey] || "";
        if (existingUnit && standardizeUnit(existingUnit) !== standardizeUnit(finalUnit)) {
            return {
                success: false,
                count: 0,
                deliveryDate,
                message: `${finalName}: หน่วยที่เติมเป็น ${finalUnit} แต่หน่วยสต๊อกเป็น ${existingUnit}`,
            };
        }
        parsedItems.push({ finalName, finalUnit, lookupKey, amount });
    }
    if (parsedItems.length === 0) {
        return { success: false, count: 0, deliveryDate, message: "ไม่พบรายการสต๊อกที่อ่านได้" };
    }

    // รวมชื่อเดียวกันก่อนเขียน ลดจำนวน Spreadsheet calls และตรวจครบก่อนแก้ข้อมูลจริง
    const combined = Object.create(null);
    for (const item of parsedItems) {
        if (!combined[item.lookupKey]) combined[item.lookupKey] = { ...item };
        else {
            if (standardizeUnit(combined[item.lookupKey].finalUnit) !== standardizeUnit(item.finalUnit)) {
                return {
                    success: false,
                    count: 0,
                    deliveryDate,
                    message: `${item.finalName}: มีหลายหน่วยในคำสั่งเดียวกัน (${combined[item.lookupKey].finalUnit}, ${item.finalUnit})`,
                };
            }
            combined[item.lookupKey].amount += item.amount;
        }
    }

    const changes = Object.values(combined).map((item) => {
        const wasExisting = !!stockMap[item.lookupKey];
        const before = wasExisting ? stockQtyMap[item.lookupKey] : null;
        return {
            lookupKey: item.lookupKey,
            name: item.finalName,
            unit: item.finalUnit,
            wasExisting,
            before,
            after: (before || 0) + item.amount,
        };
    });
    const payload = { deliveryDate, count: parsedItems.length, commandHash, changes };
    saveEventJournal(ss, journalKey, "STOCK_UPDATE", "STARTED", payload, "");
    return applyStockUpdatePlan(ss, journalKey, payload);
}

// ==============================================================================
// 📊 DASHBOARD
// ==============================================================================
function refreshDashboard(ss) {
    let dash = ss.getSheetByName("Dashboard");
    if (!dash) {
        dash = ss.insertSheet("Dashboard");
        ss.setActiveSheet(dash);
        ss.moveActiveSheet(1);
    }
    dash.clearContents();
    dash.clearFormats();

    const today = getTodayStr();

    const allOrders = [];
    ss.getSheets().forEach((sheet) => {
        if (sheet.getName().startsWith("ออเดอร์-")) {
            allOrders.push(...sheet.getDataRange().getValues());
        }
    });
    const summaryByDate = {};
    const dateList = [];

    for (const row of allOrders) {
        const colA = String(row[0] || "").trim();
        if (!colA || colA.startsWith("ใบจัดออเดอร์") || colA === "วันที่" ||
            colA.startsWith("ลำดับที่ ") || colA === "ซื้อเพิ่ม" ||
            colA.match(/^รอบ\s+\d+$/) || colA.startsWith("==========")) continue;
        const delivDateStr = normalizeDelivDate(row[0]);
        if (!delivDateStr) continue;
        const product = String(row[2] || "").trim();
        if (!product) continue;
        const amount = parseFloat(row[3]) || 0;
        const unit = String(row[4] || "");
        const store = String(row[1] || "");

        if (!summaryByDate[delivDateStr]) {
            summaryByDate[delivDateStr] = Object.create(null);
            dateList.push(delivDateStr);
        }
        if (!summaryByDate[delivDateStr][product]) {
            summaryByDate[delivDateStr][product] = { total: 0, unit, lines: [] };
        }
        summaryByDate[delivDateStr][product].total += amount;
        summaryByDate[delivDateStr][product].lines.push({ store, amount, unit });
    }

    const sortedDates = [...new Set(dateList)].sort((a, b) => {
        const da = parseDateStr(a), db = parseDateStr(b);
        if (!da || !db) return 0;
        return da - db;
    });

    const stockMaps = {};
    const getStockMapForDate = (dateStr) => {
        if (stockMaps[dateStr]) return stockMaps[dateStr];
        const stockMap = Object.create(null);
        try {
            const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
            const stockSheet = stockSS.getSheetByName(getMonthlyStockTabName(dateStr));
            if (stockSheet) {
                const stockData = stockSheet.getDataRange().getValues();
                const section = findStockDateSection(stockData, dateStr);
                if (section) {
                    for (let i = section.startIndex; i < section.endIndex; i++) {
                        const name = String(stockData[i][1]).trim();
                        if (name) stockMap[name] = { qty: parseFloat(stockData[i][2]) || 0, unit: String(stockData[i][3] || "") };
                    }
                }
            }
        } catch (err) {
            console.error("Dashboard stock error:", err.message);
        }
        stockMaps[dateStr] = stockMap;
        return stockMap;
    };

    const noStockItems = getNoStockProductNames(ss);
    let r = 1;

    dash.getRange(r, 1, 1, 6).merge().setValue(`🍎 Mr.Fruity — สรุปออเดอร์ทั้งหมด`).setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    r++;
    dash.getRange(r, 1, 1, 6).merge().setValue("อัปเดตล่าสุด: " + Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss  dd/MM/yyyy")).setBackground("#388E3C").setFontColor("#FFFFFF").setFontSize(10).setHorizontalAlignment("center");
    r += 2;

    if (sortedDates.length === 0) {
        dash.getRange(r, 1, 1, 6).merge().setValue("— ยังไม่มีออเดอร์ —").setHorizontalAlignment("center").setFontColor("#9E9E9E");
        return;
    }

    for (const dateStr of sortedDates) {
        const isToday = dateStr === today;
        const dateSummary = summaryByDate[dateStr];
        const productNames = Object.keys(dateSummary).sort();
        const stockMap = getStockMapForDate(dateStr);

        dash.getRange(r, 1, 1, 6).merge().setValue(`📅 วันส่ง: ${dateStr}${isToday ? "  ← วันนี้" : ""}`).setBackground(isToday ? "#0D47A1" : "#1565C0").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
        r++;
        dash.getRange(r, 1, 1, 6).setValues([["รายการสินค้า", "รวมจำนวน", "หน่วย", "สต๊อกคงเหลือ", "ต้องซื้อเพิ่ม", "หมายเหตุ"]]).setBackground("#2E7D32").setFontColor("#FFFFFF").setFontWeight("bold");
        r++;

        // [FIX-4] batch build แถว product แล้ว setValues ครั้งเดียว
        const productRows = [];
        const productBgs = [];
        productNames.forEach((name) => {
            const s = dateSummary[name];
            const stock = stockMap[name];
            const stockQty = stock ? stock.qty : "-";
            const needBuy = stock && !isNaN(stockQty) ? Math.max(0, s.total - stockQty) : s.total;
            const isNoStock = noStockItems.includes(name);
            const remark = isNoStock ? "🛒 สั่งตามออเดอร์" : stock ? "" : "❓ ไม่พบในสต๊อก";
            productRows.push([name, s.total, s.unit, stockQty, isNoStock ? s.total : needBuy, remark]);
            productBgs.push(isNoStock ? "#FFF9C4" : needBuy > 0 ? "#FFEBEE" : "#E8F5E9");
        });
        if (productRows.length) {
            dash.getRange(r, 1, productRows.length, 6).setValues(productRows);
            dash.getRange(r, 1, productRows.length, 6).setBackgrounds(productBgs.map((c) => Array(6).fill(c)));
            r += productRows.length;
        }

        r++;
        dash.getRange(r, 1, 1, 6).merge().setValue(`  รายละเอียดแยกตามร้าน`).setBackground("#1565C0").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10);
        r++;
        dash.getRange(r, 1, 1, 5).setValues([["ร้าน", "รายการ", "จำนวน", "หน่วย", ""]]).setBackground("#1976D2").setFontColor("#FFFFFF").setFontWeight("bold");
        r++;

        // [FIX-4] batch build store breakdown
        const storeRows = [];
        productNames.forEach((name) => {
            dateSummary[name].lines.forEach((line) => {
                storeRows.push([line.store, name, line.amount, line.unit, ""]);
            });
        });
        if (storeRows.length) {
            dash.getRange(r, 1, storeRows.length, 5).setValues(storeRows);
            r += storeRows.length;
        }

        r += 2;
    }

    if (noStockItems.length > 0) {
        dash.getRange(r, 1, 1, 6).merge().setValue("รายการที่ไม่ stock (สั่งตามออเดอร์)").setBackground("#E65100").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11);
        r++;
        const noStockRows = noStockItems.map((name) => ["• " + name, "", "", "", "", ""]);
        dash.getRange(r, 1, noStockRows.length, 6).setValues(noStockRows);
        dash.getRange(r, 1, noStockRows.length, 6).setBackgrounds(noStockRows.map(() => Array(6).fill("#FFF3E0")));
        r += noStockRows.length;
    }

    dash.setColumnWidth(1, 240);
    dash.setColumnWidth(2, 90);
    dash.setColumnWidth(3, 70);
    dash.setColumnWidth(4, 110);
    dash.setColumnWidth(5, 110);
    dash.setColumnWidth(6, 140);
    dash.setFrozenRows(1);
}

// ==============================================================================
// 🔧 ONE-TIME SETUP
// ==============================================================================
function setupOnce() {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    const mapSheet = ss.getSheetByName("Mapping");
    if (mapSheet) {
        mapSheet.getRange(1, 4).setValue("ไม่ stock (ใส่ x)");
        mapSheet.getRange(1, 4).setBackground("#FF6F00").setFontColor("#FFFFFF").setFontWeight("bold");
    }

    ScriptApp.getProjectTriggers().forEach((t) => {
        if (t.getHandlerFunction() === "onEdit") ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();

    Logger.log("✅ setupOnce complete");
}

// ==============================================================================
// 📊 UPDATE PRODUCT DATABASE SHEET ("รายการ")
// ==============================================================================
function updateProductListSheet(ss, monthlyStockMap, monthName) {
    try {
        const dbSheet = ss.getSheetByName("รายการ");
        if (!dbSheet) return;
        const data = dbSheet.getDataRange().getValues();
        if (data.length <= 1) return;

        dbSheet.getRange(1, 4).setValue(`จำนวนคงเหลือ (${monthName})`);

        // [FIX-4] batch write col D ทีเดียว
        const updates = [];
        for (let i = 1; i < data.length; i++) {
            const colA = String(data[i][0]).trim();
            const colB = String(data[i][1]).trim();
            const colC = String(data[i][2]).trim();
            let foundStock = "-";
            if (colA && monthlyStockMap[colA] !== undefined) foundStock = monthlyStockMap[colA];
            else if (colB && monthlyStockMap[colB] !== undefined) foundStock = monthlyStockMap[colB];
            else if (colC && monthlyStockMap[colC] !== undefined) foundStock = monthlyStockMap[colC];
            updates.push([foundStock]);
        }
        if (updates.length > 0) {
            dbSheet.getRange(2, 4, updates.length, 1).setValues(updates);
        }
    } catch (err) {
        console.error("Update Product List Error:", err.message);
    }
}
