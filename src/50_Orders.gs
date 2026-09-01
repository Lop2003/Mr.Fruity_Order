// ==============================================================================
// 📋 PER-DATE SHEET — "ออเดอร์-DD-MM-YY"
// ==============================================================================
function applyOrderUpdatePlan(ss, journalKey, payload) {
    try {
        const result = updateDeliverySheet(
            ss,
            payload.deliveryDate,
            payload.storeName,
            payload.items,
            new Date(payload.timestamp),
            payload.rowEventId,
        );
        saveEventJournal(ss, journalKey, "ORDER", "COMPLETE", { ...payload, result }, "");
        return result;
    } catch (err) {
        saveEventJournal(ss, journalKey, "ORDER", "STARTED", payload, err.message);
        throw err;
    }
}

function processOrderUpdate(rawMessage, ss, orderData, orderTimestamp, eventId) {
    const commandHash = eventId ? getStableHash(String(rawMessage || "").trim()) : "";
    const eventJournalKey = eventId ? `${eventId}|ORDER` : "";
    let journalKey = eventJournalKey;
    let journal = loadEventJournal(ss, journalKey);
    if (!journal && commandHash) {
        journal = loadPendingEventJournal(ss, "ORDER", commandHash);
        if (journal) journalKey = journal.key;
    }
    if (journal) {
        if (journal.status === "COMPLETE") return journal.payload.result;
        if (journal.status === "STARTED") {
            const sourceJournalKey = journalKey;
            if (eventJournalKey && eventJournalKey !== sourceJournalKey) {
                saveEventJournal(ss, eventJournalKey, "ORDER", "STARTED", journal.payload, "");
                journalKey = eventJournalKey;
            }
            const result = applyOrderUpdatePlan(ss, journalKey, journal.payload);
            if (sourceJournalKey !== journalKey) {
                saveEventJournal(ss, sourceJournalKey, "ORDER", "COMPLETE", { ...journal.payload, result }, "");
            }
            return result;
        }
    }

    const payload = {
        commandHash,
        deliveryDate: orderData.deliveryDate,
        storeName: orderData.storeName,
        items: orderData.items,
        timestamp: orderTimestamp.toISOString(),
        rowEventId: eventId || "",
    };
    saveEventJournal(ss, journalKey, "ORDER", "STARTED", payload, "");
    return applyOrderUpdatePlan(ss, journalKey, payload);
}

function updateDeliverySheet(ss, deliveryDateStr, storeName, items, orderTimestamp, eventId) {
    const sheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const NUM_COLS = 6;
    const INTERNAL_COLS = 9;
    const SECTION_LABEL_PREFIX = "ลำดับที่ ";
    const fmtUnit = (u) => u === "กก" ? "กก." : u;
    const COL_HEADERS = ["วันที่", "ชื่อร้าน", "สินค้า", "จำนวน", "หน่วย", "สเปค เพิ่มเติม"];

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        const titleText = `ใบจัดออเดอร์ Horeca รอบส่งวันที่ ${deliveryDateStr}`;
        sheet.getRange(1, 1, 1, NUM_COLS).merge().setValue(titleText).setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 120);
        sheet.setColumnWidth(2, 220);
        sheet.setColumnWidth(3, 260);
        sheet.setColumnWidth(4, 80);
        sheet.setColumnWidth(5, 80);
        sheet.setColumnWidth(6, 200);
    }
    // G-I เก็บสถานะจับคู่และ Event ID ภายใน ไม่เปลี่ยนหน้าตาใบออเดอร์ที่ผู้ใช้เห็น
    if (sheet.getMaxColumns() < INTERNAL_COLS) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), INTERNAL_COLS - sheet.getMaxColumns());
    }
    sheet.hideColumns(7, 3);

    items = items.map((item, index) => ({ ...item, _eventItemIndex: index + 1 }));
    if (eventId && sheet.getLastRow() > 0) {
        const eventRows = sheet.getRange(1, 8, sheet.getLastRow(), 2).getValues();
        const savedIndexes = new Set(
            eventRows
                .filter((row) => String(row[0] || "") === eventId)
                .map((row) => Number(row[1]))
                .filter((index) => index > 0),
        );
        items = items.filter((item) => !savedIndexes.has(item._eventItemIndex));
        if (items.length === 0) return { alreadyExists: true };
    }

    const lastRow = sheet.getLastRow();
    let storeSection = null;
    let isStoreFromPreviousRound = false;

    if (lastRow >= 2) {
        const allDataA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        const allDataB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

        // 1. หาตำแหน่งของบรรทัดตัดรอบล่างสุด เพื่อป้องกันออเดอร์ใหม่ไปโผล่ในรอบเก่า
        let lastRoundIndex = 0;
        for (let i = allDataA.length - 1; i >= 0; i--) {
            const cellA = String(allDataA[i][0]).trim();
            if (cellA.match(/^รอบ\s+\d+$/) || cellA.startsWith("==========") || cellA.startsWith("'==========")) {
                lastRoundIndex = i + 1; // เริ่มหา storeSection ใต้บรรทัดนี้ลงไป
                break;
            }
        }

        // เช็คว่าเคยมีร้านนี้ในรอบก่อนหน้าหรือไม่ (ตั้งแต่ต้นจนถึงบรรทัดตัดรอบ)
        for (let i = 0; i < lastRoundIndex; i++) {
            const cellA = String(allDataA[i][0]).trim();
            if (cellA.startsWith(SECTION_LABEL_PREFIX) || cellA === "ซื้อเพิ่ม") {
                const dataIdx = i + 2;
                if (dataIdx < allDataA.length) {
                    const sectionStore = String(allDataB[dataIdx][0]).trim();
                    if (sectionStore === storeName) {
                        isStoreFromPreviousRound = true;
                        break;
                    }
                }
            }
        }

        // 2. หาชื่อร้านเฉพาะในรอบปัจจุบันเท่านั้น
        for (let i = lastRoundIndex; i < allDataA.length; i++) {
            const cellA = String(allDataA[i][0]).trim();
            if (cellA.startsWith(SECTION_LABEL_PREFIX) || cellA === "ซื้อเพิ่ม") {
                const dataIdx = i + 2;
                if (dataIdx < allDataA.length) {
                    const sectionStore = String(allDataB[dataIdx][0]).trim();
                    if (sectionStore === storeName) {
                        const sectionRow = 2 + i;
                        const headerRow = sectionRow + 1;
                        const dataStartRow = sectionRow + 2;
                        let dataEndRow = lastRow;
                        for (let j = dataIdx + 1; j < allDataA.length; j++) {
                            const nextA = String(allDataA[j][0]).trim();
                            if (nextA.startsWith(SECTION_LABEL_PREFIX) || nextA === "ซื้อเพิ่ม" || nextA.match(/^รอบ\s+\d+$/) || nextA.startsWith("==========") || nextA.startsWith("'==========")) {
                                dataEndRow = 2 + j - 1;
                                break;
                            }
                        }
                        storeSection = { sectionRow, headerRow, dataStartRow, dataEndRow };
                    }
                }
            }
        }
    }

    if (storeSection) {
        let insertRow = storeSection.dataEndRow + 1;
        items.forEach((item) => {
            const currentLast = sheet.getLastRow();
            const addedDateStr = deliveryDateStr + " (เพิ่ม)";
            const rowValues = [
                addedDateStr, storeName, item.name, item.amount, fmtUnit(item.unit), item.inputNote || "",
                item.isFuzzy ? "FUZZY" : item.isWarning ? "UNMAPPED" : "EXACT",
                eventId || "", eventId ? item._eventItemIndex : "",
            ];
            if (insertRow > currentLast) {
                sheet.appendRow(rowValues);
            } else {
                sheet.insertRowBefore(insertRow);
                sheet.getRange(insertRow, 1, 1, INTERNAL_COLS).setValues([rowValues]);
            }
            sheet.getRange(insertRow, 1).setNumberFormat("@");

            if (item.isWarning) {
                sheet.getRange(insertRow, 1, 1, NUM_COLS).setBackground(CONFIG.WARNING_BG).setFontColor(CONFIG.WARNING_FG).setFontWeight("normal").setFontSize(10).setHorizontalAlignment("normal");
            } else {
                let bgColor = "#E3F2FD";
                if (item.isSubItem) bgColor = "#BDBDBD"; // สีเทาเข้ม ถ้าเป็นชิ้นส่วนย่อยที่เพิ่มมา (เช่น ฝัก, กรัม)

                // บังคับสีฟอนต์เป็นสีดำ และเอาตัวหนาออก ป้องกันการสืบทอด Format สีขาวจาก Header ด้านล่าง
                sheet.getRange(insertRow, 1, 1, NUM_COLS).setBackground(bgColor).setFontColor("#000000").setFontWeight("normal").setFontSize(10).setHorizontalAlignment("normal");
                if (item.isFuzzy) {
                    sheet.getRange(insertRow, 3).setBackground("#FFF9C4");
                }
            }
            insertRow++;
        });
    } else {
        let newLabel = SECTION_LABEL_PREFIX + "0";
        let headerBgColor = "#0D47A1"; // สีน้ำเงินเริ่มต้นสำหรับร้านใหม่
        let sectionRow;

        if (isStoreFromPreviousRound) {
            newLabel = "ซื้อเพิ่ม";
            headerBgColor = "#E65100"; // สีส้มเข้มให้เห็นเด่นชัดว่าซื้อเพิ่ม

            // แทรกที่ด้านบนสุดของรอบปัจจุบัน (หลังบรรทัดตัดรอบ)
            const totalNewRows = 2 + items.length; // section header + col headers + data rows
            const currentLastRoundRow = lastRow >= 2 ? (function () {
                const colAAll = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
                for (let i = colAAll.length - 1; i >= 0; i--) {
                    const v = String(colAAll[i][0]).trim();
                    if (v.match(/^รอบ\s+\d+$/) || v.startsWith("==========") || v.startsWith("'==========")) {
                        return 2 + i; // 1-indexed row of the round divider
                    }
                }
                return 1; // ไม่มีตัดรอบเลย → แทรกหลัง title row
            })() : 1;

            sectionRow = currentLastRoundRow + 1;
            sheet.insertRowsBefore(sectionRow, totalNewRows);
            // เคลียร์ format ที่สืบทอดมาจากแถวด้านล่าง (สีน้ำเงินของ ลำดับที่)
            sheet.getRange(sectionRow, 1, totalNewRows, NUM_COLS).clearFormat();
        } else {
            sectionRow = sheet.getLastRow() + 1;
        }

        sheet.getRange(sectionRow, 1, 1, NUM_COLS).merge().setValue(newLabel).setBackground(headerBgColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11);
        const hdrRow = sectionRow + 1;
        sheet.getRange(hdrRow, 1, 1, NUM_COLS).setValues([COL_HEADERS]).setBackground("#2E7D32").setFontColor("#FFFFFF").setFontWeight("bold");

        let dataRow = hdrRow + 1;

        // [FIX-4] batch write data rows ทีเดียว
        const batchRows = items.map((item) => [
            deliveryDateStr, storeName, item.name, item.amount, fmtUnit(item.unit), item.inputNote || "",
            item.isFuzzy ? "FUZZY" : item.isWarning ? "UNMAPPED" : "EXACT",
            eventId || "", eventId ? item._eventItemIndex : "",
        ]);
        if (batchRows.length) {
            sheet.getRange(dataRow, 1, batchRows.length, INTERNAL_COLS).setValues(batchRows);
            sheet.getRange(dataRow, 1, batchRows.length, NUM_COLS).setFontColor("#000000").setFontWeight("normal").setFontSize(10).setHorizontalAlignment("normal");

            items.forEach((item, idx) => {
                sheet.getRange(dataRow + idx, 1).setNumberFormat("@");

                let bgColor = "#FFFFFF";
                if (item.isSubItem) bgColor = "#E0E0E0"; // สีเทาเข้มขึ้น สำหรับหน่วยรอง

                if (item.isWarning) {
                    sheet.getRange(dataRow + idx, 1, 1, NUM_COLS).setBackground(CONFIG.WARNING_BG).setFontColor(CONFIG.WARNING_FG);
                } else {
                    if (item.isSubItem) {
                        sheet.getRange(dataRow + idx, 1, 1, NUM_COLS).setBackground(bgColor);
                    }
                    if (item.isFuzzy) {
                        sheet.getRange(dataRow + idx, 3).setBackground("#FFF9C4");
                    }
                }
            });
        }
    }

    // Renumber section headers (รีเซ็ตเมื่อขึ้นรอบใหม่)
    const finalLastRow = sheet.getLastRow();
    if (finalLastRow >= 2) {
        const colA = sheet.getRange(2, 1, finalLastRow - 1, 1).getValues();
        let sectionNum = 0;
        for (let i = 0; i < colA.length; i++) {
            const val = String(colA[i][0]).trim();
            if (val.match(/^รอบ\s+\d+$/) || val.startsWith("==========") || val.startsWith("'==========")) {
                sectionNum = 0;
            } else if (val.startsWith(SECTION_LABEL_PREFIX)) {
                sectionNum++;
                const newLabel = SECTION_LABEL_PREFIX + sectionNum;
                if (val !== newLabel) {
                    sheet.getRange(2 + i, 1).setValue(newLabel);
                }
            }
        }
    }
    return { alreadyExists: false };
}

// ==============================================================================
// ❌ CANCEL ORDER HELPERS
// ==============================================================================

/** ค้นหารายการสินค้าทั้งหมดของร้านในชีทออเดอร์ */
function findStoreOrders(orderSheet, storeName) {
    if (!orderSheet) return [];
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 3) return [];
    const allData = orderSheet.getRange(1, 1, lastRow, 6).getValues();
    const items = [];
    for (let i = 0; i < allData.length; i++) {
        const colA = String(allData[i][0]).trim();
        const colB = String(allData[i][1]).trim();
        if (!colA || colA.startsWith("ใบจัดออเดอร์") || colA === "วันที่" ||
            colA.startsWith("ลำดับที่ ") || colA === "ซื้อเพิ่ม" ||
            colA.match(/^รอบ\s+\d+$/) || colA.startsWith("==========") || colA.startsWith("'==========")) continue;
        if (colB === storeName) {
            items.push({ row: i + 1, name: String(allData[i][2]).trim(), amount: allData[i][3], unit: String(allData[i][4]).trim() });
        }
    }
    return items;
}

function getOrderSnapshotHash(items) {
    const snapshot = (items || []).map((item) => [item.row, item.name, String(item.amount), item.unit]);
    return getStableHash(JSON.stringify(snapshot));
}

function formatCancelPreview(storeName, deliveryDate, items, notice) {
    const fmtUnit = (unit) => unit === "กก" ? "กก." : unit;
    const itemList = (items || [])
        .map((item, index) => `${index + 1}. ${item.name}  ${item.amount} ${fmtUnit(item.unit)}`)
        .join("\n");
    return `${notice ? notice + "\n\n" : ""}📋 พบออเดอร์ร้าน "${storeName}" วันที่ ${deliveryDate}\n\n${itemList}\n\n` +
        `✏️ พิมพ์ "ยืนยัน" เพื่อยกเลิกทั้งหมด\n` +
        `✏️ พิมพ์ "ยืนยัน 1,3" เพื่อเลือกยกเลิกบางรายการ\n` +
        `✏️ พิมพ์ "ยกเลิกการลบ" เพื่อยกเลิก`;
}

function cancelOrderIfSnapshotMatches(ss, pending, itemIndices) {
    const orderSheet = ss.getSheetByName("ออเดอร์-" + pending.deliveryDate.replace(/\//g, "-"));
    const currentOrders = findStoreOrders(orderSheet, pending.storeName);
    const currentHash = getOrderSnapshotHash(currentOrders);
    const expectedHash = pending.snapshotHash || getOrderSnapshotHash(pending.items);
    if (!expectedHash || currentHash !== expectedHash) {
        return {
            success: false,
            stale: true,
            message: "รายการออเดอร์เปลี่ยนหลังจากแสดงตัวอย่าง",
            items: currentOrders,
            snapshotHash: currentHash,
        };
    }
    return cancelOrder(ss, pending.deliveryDate, pending.storeName, itemIndices);
}

function deleteRowsBatch(sheet, rows) {
    const sorted = Array.from(rows).sort((a, b) => b - a);
    for (let i = 0; i < sorted.length; i++) {
        let start = sorted[i];
        let count = 1;
        while (i + 1 < sorted.length && sorted[i + 1] === start - count) {
            count++;
            i++;
        }
        sheet.deleteRows(start - count + 1, count);
    }
}

/** ยกเลิกออเดอร์ — ลบทั้งหมดหรือเฉพาะบางรายการ */
function cancelOrder(ss, deliveryDateStr, storeName, itemIndices) {
    const orderSheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);
    if (!orderSheet) return { success: false, message: "ไม่พบใบออเดอร์วันที่ " + deliveryDateStr };

    const storeOrders = findStoreOrders(orderSheet, storeName);
    if (storeOrders.length === 0) return { success: false, message: 'ไม่พบออเดอร์ของร้าน "' + storeName + '"' };

    // กำหนดแถวข้อมูลที่ต้องลบ
    const dataRowsToDelete = new Set();
    if (itemIndices && itemIndices.length > 0) {
        itemIndices.filter(idx => idx >= 1 && idx <= storeOrders.length)
            .forEach(idx => dataRowsToDelete.add(storeOrders[idx - 1].row));
    } else {
        storeOrders.forEach(item => dataRowsToDelete.add(item.row));
    }
    const deletedCount = dataRowsToDelete.size;
    if (deletedCount === 0) return { success: false, message: "ไม่พบรายการที่ระบุ" };

    // หา section headers ที่เป็นของร้านนี้
    const lastRow = orderSheet.getLastRow();
    const allData = orderSheet.getRange(1, 1, lastRow, 6).getValues();
    const allRowsToDelete = new Set(dataRowsToDelete);

    for (let i = 0; i < allData.length; i++) {
        const colA = String(allData[i][0]).trim();
        if (!(colA.startsWith("ลำดับที่ ") || colA === "ซื้อเพิ่ม")) continue;
        const dataStart = i + 2; // skip section header + column header
        if (dataStart >= allData.length) continue;
        if (String(allData[dataStart][1]).trim() !== storeName) continue;

        // หาจุดสิ้นสุดของ section
        let dataEnd = allData.length - 1;
        for (let j = dataStart + 1; j < allData.length; j++) {
            const nextA = String(allData[j][0]).trim();
            if (nextA.startsWith("ลำดับที่ ") || nextA === "ซื้อเพิ่ม" ||
                nextA.match(/^รอบ\s+\d+$/) || nextA.startsWith("==========") || nextA.startsWith("'==========")) {
                dataEnd = j - 1;
                break;
            }
        }

        // เช็คว่า section นี้ยังเหลือข้อมูลไหมหลังลบ
        let hasRemaining = false;
        for (let r = dataStart; r <= dataEnd; r++) {
            if (!allRowsToDelete.has(r + 1)) { hasRemaining = true; break; }
        }
        if (!hasRemaining) {
            allRowsToDelete.add(i + 1);       // section header
            allRowsToDelete.add(i + 2);       // column headers
        }
    }

    // ลบจากล่างขึ้นบน
    deleteRowsBatch(orderSheet, allRowsToDelete);

    // Renumber sections
    const finalLastRow = orderSheet.getLastRow();
    if (finalLastRow >= 2) {
        const colA = orderSheet.getRange(2, 1, finalLastRow - 1, 1).getValues();
        let sectionNum = 0;
        for (let i = 0; i < colA.length; i++) {
            const val = String(colA[i][0]).trim();
            if (val.match(/^รอบ\s+\d+$/) || val.startsWith("==========") || val.startsWith("'==========")) {
                sectionNum = 0;
            } else if (val.startsWith("ลำดับที่ ")) {
                sectionNum++;
                const newLabel = "ลำดับที่ " + sectionNum;
                if (val !== newLabel) orderSheet.getRange(2 + i, 1).setValue(newLabel);
            }
        }
    }

    return { success: true, deletedCount };
}
