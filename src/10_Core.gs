// ==============================================================================
// 📝 LOG HELPER
// ==============================================================================
function appendLog(ss, context, action, detail, status, errorMsg) {
    try {
        const logSheet = ss.getSheetByName("Logs");
        if (logSheet) {
            logSheet.appendRow([new Date(), context, action, detail, "-", status, errorMsg || ""]);
        }
    } catch (_) {
        // ถ้า log พัง ก็แค่ console.error แล้วไปต่อ ไม่ throw ซ้ำ
        console.error("appendLog failed:", errorMsg);
    }
}

function withScriptLock(fn) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    } catch (_) {
        const err = new Error("ระบบกำลังประมวลผลคำสั่งอื่น กรุณาลองใหม่อีกครั้ง");
        err.code = "LOCK_TIMEOUT";
        throw err;
    }
    try {
        return fn();
    } finally {
        lock.releaseLock();
    }
}

function getOrCreateSheet(ss, name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
}

function safeSetLog(logSheet, row, value) {
    if (logSheet && row) logSheet.getRange(row, 4).setValue(value);
}

function setActiveLogContext(logSheet, row) {
    ACTIVE_LOG_SHEET = logSheet;
    ACTIVE_LOG_ROW = row;
}

function finishActiveLog(state, suffix) {
    if (!ACTIVE_LOG_SHEET || !ACTIVE_LOG_ROW) return;
    const cell = ACTIVE_LOG_SHEET.getRange(ACTIVE_LOG_ROW, 4);
    const current = String(cell.getValue() || "").trim();
    const detail = /^(PROCESSING|Processing)\b/.test(current)
        ? ""
        : current.replace(/^(COMPLETE|FAILED)\s*;?\s*/, "");
    cell.setValue([state, detail, suffix].filter(Boolean).join("; "));
}

function setActiveReplyState(state, message) {
    if (!ACTIVE_LOG_SHEET || !ACTIVE_LOG_ROW) return;
    try {
        ACTIVE_LOG_SHEET.getRange(ACTIVE_LOG_ROW, 6, 1, 2)
            .setValues([[state, String(message || "")]]);
    } catch (err) {
        console.error("setActiveReplyState failed:", err.message);
    }
}

function recordEventFailure(event, status) {
    if (ACTIVE_LOG_SHEET && ACTIVE_LOG_ROW) {
        finishActiveLog("FAILED", status);
        return;
    }
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        const logSheet = getOrCreateSheet(ss, "Logs");
        const eventId = event.webhookEventId || (event.message && event.message.id) || event.replyToken || "unknown";
        const rawMessage = event.message && event.message.text ? event.message.text.trim() : "";
        logSheet.appendRow([new Date(), getSourceUserId(event), rawMessage, `FAILED; ${status}`, eventId]);
        setActiveLogContext(logSheet, logSheet.getLastRow());
    } catch (err) {
        console.error("recordEventFailure failed:", err.message);
    }
}

function beginEventLog(ss, event, rawMessage, timestamp) {
    return withScriptLock(() => {
        const logSheet = getOrCreateSheet(ss, "Logs");
        const eventId = event.webhookEventId || (event.message && event.message.id) || event.replyToken || `${timestamp.getTime()}-${getSourceUserId(event)}`;
        const lastRow = logSheet.getLastRow();
        if (lastRow > 0) {
            const startRow = Math.max(1, lastRow - 1999);
            const recentRows = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();
            for (let i = recentRows.length - 1; i >= 0; i--) {
                if (String(recentRows[i][4] || "") !== eventId) continue;
                const logRow = startRow + i;
                const status = String(recentRows[i][3] || "");
                if (status.startsWith("COMPLETE") || status.includes("LINE Reply Success")) {
                    return { duplicate: true, eventId, status };
                }
                const replyState = String(recentRows[i][5] || "");
                const replyMessage = String(recentRows[i][6] || "");
                if (replyState === "COMPLETE") {
                    return { duplicate: true, eventId, status: status || "Reply COMPLETE" };
                }
                if (replyState === "FAILED" && replyMessage) {
                    logSheet.getRange(logRow, 1).setValue(timestamp);
                    logSheet.getRange(logRow, 4).setValue("PROCESSING; Retry LINE reply only");
                    return { duplicate: false, retry: true, retryReply: replyMessage, eventId, logSheet, logRow };
                }
                const previousTime = new Date(recentRows[i][0]).getTime();
                const ageMs = timestamp.getTime() - previousTime;
                if (/^PROCESSING\b/i.test(status) && Number.isFinite(previousTime) && ageMs < CONFIG.EVENT_STALE_MS) {
                    return { duplicate: true, inProgress: true, eventId, status };
                }
                logSheet.getRange(logRow, 1).setValue(timestamp);
                logSheet.getRange(logRow, 4).setValue(`PROCESSING; Retry from ${status || "unknown"}`);
                return { duplicate: false, retry: true, eventId, logSheet, logRow };
            }
        }

        logSheet.appendRow([timestamp, getSourceUserId(event), rawMessage, "PROCESSING", eventId]);
        return { duplicate: false, eventId, logSheet, logRow: logSheet.getLastRow() };
    });
}

function getEventJournalSheet(ss) {
    let sheet = ss.getSheetByName("_EventJournal");
    if (sheet) return sheet;
    sheet = ss.insertSheet("_EventJournal");
    sheet.appendRow(["Key", "Type", "Status", "Payload", "Error", "Updated At"]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
    return sheet;
}

function loadEventJournal(ss, key) {
    if (!key) return null;
    const sheet = getEventJournalSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0] || "") !== key) continue;
        try {
            return { sheet, row: i + 2, type: String(rows[i][1] || ""), status: String(rows[i][2] || ""), payload: JSON.parse(String(rows[i][3] || "{}")) };
        } catch (_) {
            throw new Error(`Event journal เสียหาย: ${key}`);
        }
    }
    return null;
}

function saveEventJournal(ss, key, type, status, payload, errorMessage) {
    if (!key) return;
    const sheet = getEventJournalSheet(ss);
    const existing = loadEventJournal(ss, key);
    const values = [[key, type, status, JSON.stringify(payload || {}), errorMessage || "", new Date()]];
    if (existing) sheet.getRange(existing.row, 1, 1, 6).setValues(values);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, 6).setValues(values);
}

function getStableHash(value) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ""));
    return Utilities.base64EncodeWebSafe(digest);
}

function loadPendingEventJournal(ss, type, commandHash) {
    if (!type || !commandHash) return null;
    const sheet = getEventJournalSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][1] || "") !== type || String(rows[i][2] || "") !== "STARTED") continue;
        try {
            const payload = JSON.parse(String(rows[i][3] || "{}"));
            if (payload.commandHash === commandHash) {
                return { sheet, row: i + 2, key: String(rows[i][0] || ""), type, status: "STARTED", payload };
            }
        } catch (_) {
            throw new Error(`Event journal เสียหาย: ${rows[i][0]}`);
        }
    }
    return null;
}

function getCutoffJournalSheet(ss) {
    let sheet = ss.getSheetByName("_CutoffJournal");
    if (sheet) return sheet;
    sheet = ss.insertSheet("_CutoffJournal");
    sheet.appendRow(["Key", "วันที่", "รอบ", "Event ID", "Status", "Payload", "Error", "Updated At"]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
    return sheet;
}

function loadCutoffJournal(ss, key) {
    const sheet = getCutoffJournalSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]) !== key) continue;
        let payload;
        try {
            payload = JSON.parse(String(rows[i][5] || "{}"));
        } catch (_) {
            throw new Error(`Cutoff journal เสียหาย: ${key}`);
        }
        return { sheet, row: i + 2, status: String(rows[i][4] || ""), payload };
    }
    return null;
}

function loadPendingCutoffJournal(ss, deliveryDate) {
    const sheet = getCutoffJournalSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
        const status = String(rows[i][4] || "");
        if (normalizeDelivDate(rows[i][1]) !== deliveryDate || status === "COMPLETE") continue;
        let payload;
        try {
            payload = JSON.parse(String(rows[i][5] || "{}"));
        } catch (_) {
            throw new Error(`Cutoff journal เสียหาย: ${rows[i][0]}`);
        }
        return { sheet, row: i + 2, key: String(rows[i][0]), round: Number(rows[i][2]), status, payload };
    }
    return null;
}

function saveCutoffJournal(ss, key, deliveryDate, round, eventId, status, payload, errorMessage) {
    const sheet = getCutoffJournalSheet(ss);
    const existing = loadCutoffJournal(ss, key);
    const values = [[key, deliveryDate, round, eventId || "", status, JSON.stringify(payload || {}), errorMessage || "", new Date()]];
    const row = existing ? existing.row : sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 2).setNumberFormat("@");
    sheet.getRange(row, 1, 1, 8).setValues(values);
    return row;
}

function hasOpenOrdersForCutoff(ss, targetDate) {
    const orderSheet = ss.getSheetByName("ออเดอร์-" + targetDate.replace(/\//g, "-"));
    if (!orderSheet) return false;
    const rounds = getOrderRounds(orderSheet);
    if (!rounds.length) return false;
    const openRound = rounds[rounds.length - 1];
    return !openRound.isClosed && Object.keys(openRound.products).length > 0;
}

function executeCutoff(ss, targetDate, eventId, allowFollowUp = true, skipPreRefresh = false) {
    let journal = loadPendingCutoffJournal(ss, targetDate);
    const resumedRoundClosed = !!journal && journal.status === "ROUND_CLOSED";
    let deductRes;
    let closedRound;

    if (journal) {
        deductRes = journal.payload.deductRes;
        closedRound = journal.round;
    } else {
        // Refresh รอบที่กำลังเปิดจากสต๊อกจริงก่อนหัก เพื่อไม่ให้ snapshot เก่าถูกแช่แข็ง
        if (!skipPreRefresh) {
            try {
                updatePurchaseSummarySheet(ss, targetDate, false, true);
            } catch (err) {
                return { syncError: err.message };
            }
        }

        deductRes = prepareStockDeduction(ss, targetDate);
        if (!deductRes.success) return { deductRes };
        closedRound = deductRes.round;
        const key = `${targetDate}|${closedRound}`;
        saveCutoffJournal(ss, key, targetDate, closedRound, eventId, "STARTED", { deductRes }, "");
        journal = loadCutoffJournal(ss, key);
        journal.key = key;
        journal.round = closedRound;
    }

    if (!deductRes || !Array.isArray(deductRes.plan)) {
        throw new Error(`Cutoff journal ไม่มีแผนหักสต๊อก: ${journal.key}`);
    }

    if (journal.status === "STARTED") {
        try {
            applyStockDeductionPlan(targetDate, deductRes.plan);
        } catch (err) {
            saveCutoffJournal(ss, journal.key, targetDate, closedRound, eventId, "STARTED", { deductRes }, err.message);
            return { deductRes: { success: false, message: "Stock Error: " + err.message } };
        }
        saveCutoffJournal(ss, journal.key, targetDate, closedRound, eventId, "STOCK_DEDUCTED", { deductRes }, "");
        journal.status = "STOCK_DEDUCTED";
    }

    let res = { success: true, closedRound, round: closedRound + 1 };
    if (journal.status === "STOCK_DEDUCTED") {
        res = createCutoffRoundDivider(ss, targetDate, closedRound);
        if (!res.success) {
            saveCutoffJournal(ss, journal.key, targetDate, closedRound, eventId, "STOCK_DEDUCTED", { deductRes }, res.message || "สร้างแถวตัดรอบไม่สำเร็จ");
            return { deductRes, res };
        }
        saveCutoffJournal(ss, journal.key, targetDate, closedRound, eventId, "ROUND_CLOSED", { deductRes }, "");
        journal.status = "ROUND_CLOSED";
    }

    let summaryError = null;
    if (journal.status === "ROUND_CLOSED") {
        try {
            updatePurchaseSummarySheet(ss, targetDate, false, true);
        } catch (err) {
            summaryError = err.message;
        }
        saveCutoffJournal(ss, journal.key, targetDate, closedRound, eventId, summaryError ? "ROUND_CLOSED" : "COMPLETE", { deductRes }, summaryError || "");
    }

    if (resumedRoundClosed && !summaryError && allowFollowUp) {
        const recoveredResult = {
            deductRes,
            res,
            summaryError: null,
            recoveredOnly: true,
            recoveredRound: closedRound,
        };
        if (!hasOpenOrdersForCutoff(ss, targetDate)) return recoveredResult;

        const followUp = executeCutoff(ss, targetDate, eventId, false, true);
        const followUpError = followUp.syncError
            || (!followUp.deductRes || !followUp.deductRes.success
                ? (followUp.deductRes && followUp.deductRes.message) || "ไม่สามารถเตรียมตัดรอบใหม่ได้"
                : (!followUp.res || !followUp.res.success
                    ? (followUp.res && followUp.res.message) || "ไม่สามารถปิดรอบใหม่ได้"
                    : ""));
        if (followUpError) {
            return {
                ...recoveredResult,
                followUpError,
                followUpDeductRes: followUp.deductRes || null,
                followUpRes: followUp.res || null,
            };
        }
        return { ...followUp, recoveredRound: closedRound };
    }
    return {
        deductRes,
        res,
        summaryError,
        recoveryRetryFailed: resumedRoundClosed && !!summaryError,
        recoveredRound: resumedRoundClosed ? closedRound : undefined,
    };
}

function getSourceUserId(event) {
    return event.source && event.source.userId ? event.source.userId : "unknown";
}

function parseDateFromText(text) {
    const m = String(text || "").match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
    return m ? normalizeDelivDate(m[1]) : null;
}

function isStockDateHeader(row) {
    const colB = String(row[1] || "").trim();
    const colC = String(row[2] || "").trim();
    const colD = String(row[3] || "").trim();
    return !!colB && colC === "" && colD === "" && (
        colB.includes("/") ||
        colB.includes("202") ||
        !!colB.match(/ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\./)
    );
}

function findStockDateSection(listData, deliveryDateStr) {
    let inTargetSection = false;
    let startIndex = -1;

    for (let i = 0; i < listData.length; i++) {
        const rawColB = listData[i][1];
        const colB = String(rawColB).trim();

        if (!inTargetSection) {
            const normDate = normalizeDelivDate(rawColB);
            if (colB === deliveryDateStr || normDate === deliveryDateStr) {
                inTargetSection = true;
                startIndex = i + 1;
            }
            continue;
        }

        if (colB === "") continue;
        if (isStockDateHeader(listData[i])) {
            return { startIndex, endIndex: i };
        }
    }

    return inTargetSection ? { startIndex, endIndex: listData.length } : null;
}

function getOrCreateMonthlyStockSheet(stockSS, deliveryDateStr) {
    const tabName = getMonthlyStockTabName(deliveryDateStr);
    let stockSheet = stockSS.getSheetByName(tabName);
    if (stockSheet) return stockSheet;

    const templateSheet = stockSS.getSheetByName("เทมเพลสสต็อก");
    if (templateSheet) {
        stockSheet = templateSheet.copyTo(stockSS);
        stockSheet.setName(tabName);
        stockSheet.showSheet();
        stockSS.setActiveSheet(stockSheet);
        stockSS.moveActiveSheet(1);
        stockSheet.getRange(3, 1, 1, 4).clearContent();
        return stockSheet;
    }
    return stockSS.insertSheet(tabName);
}

function ensureStockDateSection(deliveryDateStr) {
    const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
    const stockSheet = getOrCreateMonthlyStockSheet(stockSS, deliveryDateStr);
    let data = stockSheet.getDataRange().getValues();
    let section = findStockDateSection(data, deliveryDateStr);
    if (section) return { stockSS, stockSheet, section };

    let targetRow = 3;
    for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][1] || "").trim() !== "") {
            targetRow = i + 2;
            break;
        }
    }
    if (targetRow < 3) targetRow = 3;

    stockSheet.getRange(targetRow, 2)
        .clearDataValidations()
        .setNumberFormat("@")
        .setValue(deliveryDateStr)
        .setBackground("#FFF2CC")
        .setFontWeight("bold");

    SpreadsheetApp.flush();
    data = stockSheet.getDataRange().getValues();
    section = findStockDateSection(data, deliveryDateStr);
    return { stockSS, stockSheet, section };
}
