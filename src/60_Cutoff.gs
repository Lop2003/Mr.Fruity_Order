// ==============================================================================
// 🛒 PURCHASE SUMMARY SHEET & STOCK DEDUCTION
// ==============================================================================

function getOrderRounds(orderSheet) {
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 3) return [];
    // รองรับใบออเดอร์เก่าที่มีเพียง 6 คอลัมน์ก่อนเพิ่มสถานะจับคู่ภายใน
    const readCols = Math.min(7, orderSheet.getMaxColumns());
    const allData = orderSheet.getRange(1, 1, lastRow, readCols).getValues();

    const rounds = [];
    let currentRoundNum = 1;
    let currentRoundProducts = {};
    const SECTION_PREFIX = "ลำดับที่ ";

    for (let i = 0; i < allData.length; i++) {
        const colA = String(allData[i][0]).trim();
        if (!colA || colA.startsWith("ใบจัดออเดอร์") || colA === "วันที่" || colA.startsWith(SECTION_PREFIX)) continue;

        const roundMatch = colA.match(/รอบ\s+(\d+)/);
        if (colA.startsWith("==========") || roundMatch) {
            // เมื่อเจอป้ายบอกว่าขึ้นรอบใหม่ เช่น "========== รอบที่ 1" หรือ "รอบ 2"
            // ถือว่าจบรอบปัจจุบัน แล้วดันเข้า Array
            if (Object.keys(currentRoundProducts).length > 0 || rounds.length === 0) {
                rounds.push({ round: currentRoundNum, isClosed: true, products: currentRoundProducts });
            }
            currentRoundNum++;
            currentRoundProducts = {};
            continue;
        }

        const product = String(allData[i][2]).trim();
        const amount = parseFloat(allData[i][3]) || 0;
        const unit = String(allData[i][4] || "").trim();
        const spec = String(allData[i][5] || "").trim();
        const matchStatus = readCols >= 7 ? String(allData[i][6] || "").trim() : "";
        const canDeduct = !matchStatus || matchStatus === "EXACT" || matchStatus === "FUZZY";
        if (!product) continue;

        // key = ชื่อ|หน่วย เพื่อแยกแถวเมื่อหน่วยต่างกัน (เช่น กก. vs ลูก)
        const productKey = product + "|" + unit;

        if (!currentRoundProducts[productKey]) {
            currentRoundProducts[productKey] = { total: amount, unit, name: product, specs: [], isSafeToDeduct: canDeduct };
            if (spec) currentRoundProducts[productKey].specs.push(spec);
        } else {
            currentRoundProducts[productKey].total += amount;
            currentRoundProducts[productKey].isSafeToDeduct = currentRoundProducts[productKey].isSafeToDeduct && canDeduct;
            if (spec && !currentRoundProducts[productKey].specs.includes(spec)) {
                currentRoundProducts[productKey].specs.push(spec);
            }
        }
    }

    // รอบสุดท้ายที่ยังไม่ได้ตัด (แม้จะไม่มีสินค้าเลยก็ตาม ก็ดันเข้า Array เพื่อให้แสดงหัวข้อ)
    rounds.push({ round: currentRoundNum, isClosed: false, products: currentRoundProducts });

    return rounds;
}

function getHistoricalStock(summarySheet) {
    const history = {};
    if (!summarySheet) return history;

    const data = summarySheet.getDataRange().getValues();
    let currentParsingRound = null;

    for (let i = 0; i < data.length; i++) {
        const colA = String(data[i][0]).trim();
        const roundMatch = colA.match(/^รอบที่\s+(\d+)/);
        if (roundMatch) {
            currentParsingRound = parseInt(roundMatch[1], 10);
            if (!history[currentParsingRound]) history[currentParsingRound] = {};
            continue;
        }

        if (currentParsingRound && colA && colA !== "สินค้า") {
            // ใช้ key = ชื่อ|หน่วย ให้ตรงกับ productKey ใน getOrderRounds
            const unit = String(data[i][2] || "").trim();
            const historyKey = colA + "|" + unit;
            history[currentParsingRound][historyKey] = {
                remainingStock: data[i][3],     // D
                stockUnit: data[i][4],          // E
                afterStock: data[i][5],         // F
                needBuy: data[i][6]             // G
            };
        }
    }
    return history;
}

function prepareStockDeduction(ss, deliveryDateStr) {
    const orderSheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);
    if (!orderSheet) return { success: false, message: `ไม่พบชีตออเดอร์วันที่ ${deliveryDateStr}` };

    const rounds = getOrderRounds(orderSheet);
    if (rounds.length === 0) return { success: false, message: "ไม่มีออเดอร์ให้ตัดรอบ" };

    const openRound = rounds[rounds.length - 1];
    if (openRound.isClosed) return { success: false, message: "ไม่มีออเดอร์ใหม่ตั้งแต่ตัดรอบครั้งล่าสุด" };
    if (Object.keys(openRound.products).length === 0) {
        return { success: false, message: "ไม่มีออเดอร์ใหม่ตั้งแต่ตัดรอบครั้งล่าสุด" };
    }

    try {
        const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
        const stockSheet = stockSS.getSheetByName(getMonthlyStockTabName(deliveryDateStr));
        if (!stockSheet) return { success: false, message: "ไม่พบชีตสต็อกเดือนนี้" };
        const listData = stockSheet.getDataRange().getValues();
        const section = findStockDateSection(listData, deliveryDateStr);
        if (!stockSheet || !section) return { success: false, message: "ไม่พบชีตสต็อกเดือนนี้" };

        const normalizeUnit = (u) => String(u || "").replace(/\./g, "").trim();

        const stockUnitMap = Object.create(null);
        const stockRowsMap = Object.create(null);
        for (let i = section.startIndex; i < section.endIndex; i++) {
            const prod = String(listData[i][1]).trim();
            const unit = String(listData[i][3]).trim();
            if (prod) {
                stockUnitMap[prod] = unit;
                if (!stockRowsMap[prod]) stockRowsMap[prod] = [];
                stockRowsMap[prod].push(i);
            }
        }

        const stockDeductMap = Object.create(null);
        let unmatchedCount = 0;
        const unmatchedItems = [];
        for (const key of Object.keys(openRound.products)) {
            const p = openRound.products[key];
            const name = p.name || key;
            const stockUnit = stockUnitMap[name] || "";
            if (!p.isSafeToDeduct) {
                unmatchedCount++;
                unmatchedItems.push({
                    name,
                    requested: p.total,
                    orderUnit: p.unit,
                    reason: "ไม่พบชื่อใน Mapping ต้องตรวจสอบ",
                });
            } else if (stockUnit && normalizeUnit(p.unit) === normalizeUnit(stockUnit)) {
                stockDeductMap[name] = (stockDeductMap[name] || 0) + p.total;
            } else {
                unmatchedCount++;
                unmatchedItems.push({
                    name,
                    requested: p.total,
                    orderUnit: p.unit,
                    reason: stockUnit ? `หน่วยสต๊อกเป็น ${stockUnit}` : "ไม่พบสินค้าในสต๊อก",
                });
            }
        }

        let deductedCount = 0;
        let fullDeductedCount = 0;
        let zeroStockCount = 0;
        const deductedItems = [];
        const partialItems = [];
        const zeroStockItems = [];
        const deductedByUnit = {};
        const plan = [];
        for (const product of Object.keys(stockDeductMap)) {
            const productRows = stockRowsMap[product] || [];
            if (productRows.length !== 1) {
                return {
                    success: false,
                    message: productRows.length > 1
                        ? `พบสินค้าซ้ำในสต๊อก: ${product}`
                        : `ไม่พบสินค้าในสต๊อก: ${product}`,
                };
            }
            const i = productRows[0];
            const currentStock = parseFloat(listData[i][2]) || 0;
            const deductQty = stockDeductMap[product] || 0;
            const actualDeduct = Math.min(currentStock, deductQty);
            const afterStock = Math.max(0, currentStock - deductQty);
            if (actualDeduct > 0) {
                plan.push({ row: i + 1, product, before: currentStock, after: afterStock });
            }
            if (actualDeduct > 0) {
                deductedCount++;
                const unit = stockUnitMap[product] || "หน่วย";
                deductedByUnit[unit] = (deductedByUnit[unit] || 0) + actualDeduct;
                deductedItems.push({
                    name: product,
                    requested: deductQty,
                    deducted: actualDeduct,
                    shortage: Math.max(0, deductQty - actualDeduct),
                    unit,
                });
                if (actualDeduct < deductQty) {
                    partialItems.push({
                        name: product,
                        requested: deductQty,
                        deducted: actualDeduct,
                        shortage: deductQty - actualDeduct,
                        unit,
                    });
                } else {
                    fullDeductedCount++;
                }
            } else if (deductQty > 0) {
                zeroStockCount++;
                zeroStockItems.push({
                    name: product,
                    requested: deductQty,
                    unit: stockUnitMap[product] || "หน่วย",
                });
            }
        }

        return { success: true, round: openRound.round, plan, deductedCount, fullDeductedCount, deductedItems, partialItems, unmatchedItems, zeroStockItems, deductedByUnit, unmatchedCount, zeroStockCount };
    } catch (err) {
        return { success: false, message: "Stock Error: " + err.message };
    }
}

function applyStockDeductionPlan(deliveryDateStr, plan) {
    const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
    const stockSheet = stockSS.getSheetByName(getMonthlyStockTabName(deliveryDateStr));
    if (!stockSheet) throw new Error("ไม่พบชีตสต็อกเดือนนี้");

    for (const change of plan || []) {
        const rowValues = stockSheet.getRange(change.row, 2, 1, 2).getValues()[0];
        const currentProduct = String(rowValues[0] || "").trim();
        const currentQty = parseFloat(rowValues[1]) || 0;
        if (currentProduct !== change.product) {
            throw new Error(`ตำแหน่งสินค้าเปลี่ยนระหว่างตัดรอบ: ${change.product}`);
        }
        if (Math.abs(currentQty - change.after) < 1e-9) continue;
        if (Math.abs(currentQty - change.before) >= 1e-9) {
            throw new Error(`สต๊อก ${change.product} ถูกแก้ระหว่างตัดรอบ (คาด ${change.before}, พบ ${currentQty})`);
        }
        stockSheet.getRange(change.row, 3).setValue(change.after);
    }
    SpreadsheetApp.flush();
}

function deductStockForCutoff(ss, deliveryDateStr) {
    const result = prepareStockDeduction(ss, deliveryDateStr);
    if (!result.success) return result;
    try {
        applyStockDeductionPlan(deliveryDateStr, result.plan);
        return result;
    } catch (err) {
        return { success: false, message: "Stock Error: " + err.message };
    }
}

function updatePurchaseSummarySheet(ss, deliveryDateStr, useCache = false, requireStock = false) {
    deliveryDateStr = normalizeDelivDate(deliveryDateStr);
    const summarySheetName = "ใบซื้อ-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);

    const rounds = orderSheet ? getOrderRounds(orderSheet) : [];

    const summarySheet = ss.getSheetByName(summarySheetName);
    const history = getHistoricalStock(summarySheet);

    let externalStockMap = Object.create(null);
    let monthlyStockMap = Object.create(null);
    let tabName = "";

    if (!useCache) {
        try {
            const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
            tabName = getMonthlyStockTabName(deliveryDateStr);
            const listSheet = stockSS.getSheetByName(tabName);
            if (listSheet) {
                const listData = listSheet.getDataRange().getValues();
                for (let i = 0; i < listData.length; i++) {
                    const b = String(listData[i][1]).trim();
                    const cVal = listData[i][2];
                    if (b && cVal !== "") {
                        const parsedQty = parseFloat(cVal);
                        if (!isNaN(parsedQty)) monthlyStockMap[b] = parsedQty;
                    }
                }
                const section = findStockDateSection(listData, deliveryDateStr);
                if (section) {
                    for (let i = section.startIndex; i < section.endIndex; i++) {
                        const colB = String(listData[i][1]).trim();
                        const colC = String(listData[i][2]).trim();
                        const colD = String(listData[i][3]).trim();
                        if (colB) externalStockMap[colB] = { qty: parseFloat(colC) || 0, unit: colD };
                    }
                } else if (requireStock) {
                    throw new Error(`ไม่พบส่วนสต๊อกวันที่ ${deliveryDateStr} ในแท็บ ${tabName}`);
                }
            } else if (requireStock) {
                throw new Error(`ไม่พบแท็บสต๊อก ${tabName}`);
            }
        } catch (err) {
            console.error(err);
            if (requireStock) throw new Error("อ่านสต๊อกไม่สำเร็จ: " + err.message);
        }
    }

    let sheet = summarySheet;
    if (sheet) {
        sheet.getDataRange().breakApart();
        sheet.clearContents();
        sheet.clearFormats();
    } else {
        sheet = ss.insertSheet(summarySheetName);
    }

    const NUM_COLS = 9;
    sheet.getRange(1, 1).setValue(deliveryDateStr).setNumberFormat("@").setBackground("#0D47A1").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
    sheet.getRange(1, 2, 1, NUM_COLS - 1).merge().setValue(`ใบซื้อ ออเดอร์ Horeca รอบส่งวันที่ ${deliveryDateStr}`).setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    sheet.getRange(2, 1, 1, NUM_COLS).setValues([["สินค้า", "รวมลูกค้าสั่ง", "หน่วยลูกค้าสั่ง", "คงเหลือ สตอค", "หน่วย สตอค", "ยอดรวมตัดสตอค", "ของที่ต้องซื้อ", "หน่วย", "สเปค เพิ่มเติม"]]).setBackground("#2E7D32").setFontColor("#FFFFFF").setFontWeight("bold");

    const dataRows = [];
    const bgColors = [];
    const fontWeights = [];

    const reversedRounds = [...rounds].reverse();
    const normalizeUnit = (u) => String(u || "").replace(/\./g, "").trim();

    for (const rData of reversedRounds) {
        const roundNum = rData.round;
        const isClosed = rData.isClosed;

        const dividerText = isClosed ? `รอบที่ ${roundNum}` : `รอบที่ ${roundNum} (ล่าสุด)`;
        dataRows.push([dividerText, "", "", "", "", "", "", "", ""]);
        bgColors.push(Array(NUM_COLS).fill("#ffd966"));
        fontWeights.push(Array(NUM_COLS).fill("bold"));

        const sortedProducts = Object.keys(rData.products).sort();

        // ติดตามชื่อสินค้าที่เจอแล้ว — แถวแรกสีปกติ, แถวถัดไป (ต่างหน่วย) สีเทา
        const seenProductNames = new Set();

        for (const productKey of sortedProducts) {
            const s = rData.products[productKey];
            const productName = s.name || productKey;
            const isDuplicateUnit = seenProductNames.has(productName);
            seenProductNames.add(productName);

            let stockQty = 0, stockUnit = "-", isMatched = false;
            let remainingStock, afterStock, needBuyNum, needBuy;

            if (isClosed && history[roundNum] && history[roundNum][productKey]) {
                const h = history[roundNum][productKey];
                remainingStock = h.remainingStock;
                stockUnit = h.stockUnit;
                afterStock = h.afterStock;
                needBuy = h.needBuy;
                needBuyNum = (needBuy === "ไม่ต้องซื้อ") ? 0 : parseFloat(needBuy);
                isMatched = true;
            } else {
                let currentStockMap = externalStockMap;
                if (useCache && history[roundNum] && history[roundNum][productKey]) {
                    const h = history[roundNum][productKey];
                    const origQty = (parseFloat(h.remainingStock) || 0) + (parseFloat(h.afterStock) || 0);
                    currentStockMap = { [productName]: { qty: origQty, unit: h.stockUnit } };
                }

                // เทียบหน่วยสั่งกับหน่วยสตอคจริง — ตรงกันจึงตัด ไม่ตรงก็ซื้อเสมอ
                if (currentStockMap[productName]) {
                    const stk = currentStockMap[productName];
                    stockUnit = stk.unit;
                    if (normalizeUnit(s.unit) === normalizeUnit(stk.unit)) {
                        stockQty = stk.qty;
                        isMatched = true;
                    }
                    // หน่วยไม่ตรง → stockQty = 0 → ต้องซื้อทั้งหมด
                }
                afterStock = Math.min(s.total, stockQty);
                needBuyNum = s.total - afterStock;
                needBuy = needBuyNum > 0 ? needBuyNum : "ไม่ต้องซื้อ";
                remainingStock = Math.max(0, stockQty - s.total);
            }

            const buyUnit = needBuyNum > 0 ? s.unit : "";
            dataRows.push([productName, s.total, s.unit, remainingStock, stockUnit, afterStock, needBuy, buyUnit, s.specs.join(", ")]);

            // สีเทาเฉพาะแถวที่ 2+ ของชื่อเดียวกัน (หน่วยย่อย เช่น ลูก)
            let rowBg;
            if (isDuplicateUnit) {
                rowBg = Array(NUM_COLS).fill("#E0E0E0");
            } else {
                rowBg = Array(NUM_COLS).fill(needBuyNum > 0 ? "#FFEBEE" : "#E8F5E9");
            }
            if (isMatched) rowBg[0] = "#FFF2CC";
            bgColors.push(rowBg);

            const rowFw = Array(NUM_COLS).fill("normal");
            fontWeights.push(rowFw);
        }
    }

    if (dataRows.length > 0) {
        sheet.getRange(3, 1, dataRows.length, NUM_COLS).setValues(dataRows);
        sheet.getRange(3, 1, dataRows.length, NUM_COLS).setBackgrounds(bgColors);
        sheet.getRange(3, 1, dataRows.length, NUM_COLS).setFontWeights(fontWeights);

        for (let i = 0; i < dataRows.length; i++) {
            if (String(dataRows[i][0]).startsWith("รอบที่")) {
                sheet.getRange(3 + i, 1, 1, NUM_COLS).merge().setHorizontalAlignment("center");
            }
        }
    }

    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 130);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 80);
    sheet.setColumnWidth(9, 200);
    sheet.setFrozenRows(2);

    if (!useCache && tabName) {
        updateProductListSheet(ss, monthlyStockMap, tabName);
    }
}

// ==============================================================================
// ⏱️ สร้างแถวตัดรอบ
// ==============================================================================
function createCutoffRoundDivider(ss, deliveryDateStr, expectedClosedRound) {
    const sheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false };

    const data = sheet.getDataRange().getValues();
    let cutoffCount = 0;
    for (let i = 0; i < data.length; i++) {
        const rowStr = String(data[i][0]).trim();
        if (rowStr.includes("========== รอบ") || rowStr.match(/^รอบ\s+\d+$/)) {
            cutoffCount++
        }
    }

    if (expectedClosedRound && cutoffCount >= expectedClosedRound) {
        return { success: true, round: expectedClosedRound + 1, closedRound: expectedClosedRound, alreadyExists: true };
    }
    if (expectedClosedRound && cutoffCount !== expectedClosedRound - 1) {
        return { success: false, message: `ลำดับรอบไม่ตรง (คาด ${expectedClosedRound - 1}, พบ ${cutoffCount})` };
    }

    const closedRound = expectedClosedRound || cutoffCount + 1;
    const nextRound = closedRound + 1;
    const dividerText = `รอบ ${nextRound}`;

    sheet.appendRow([dividerText, "", "", "", "", ""]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 6)
        .merge()
        .setBackground("#FF6F00") // สีส้มเข้มให้สังเกตง่าย
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setFontSize(14)
        .setHorizontalAlignment("center");

    return { success: true, round: nextRound, closedRound };
}
