// ==============================================================================
// 🧪 OWNER-ONLY INTEGRATION TEST
// ==============================================================================

const DIRECT_TEST_DATE = "31/12/99";

// เก็บ TEST- rows ไว้เป็นหลักฐาน; LINE reply failure เป็น expected เพราะไม่มี reply token จริง

function validateDirectTestEventId(eventId) {
    const normalized = String(eventId || "").trim();
    if (!/^test-[A-Za-z0-9_-]{8,80}$/.test(normalized)) {
        throw new Error("Event ID สำหรับ direct test ไม่ถูกต้อง");
    }
    return normalized;
}

function createDirectTestEvent(text, eventId) {
    return {
        type: "message",
        timestamp: Date.now(),
        webhookEventId: eventId,
        source: { type: "user", userId: "U-DIRECT-INTEGRATION-TEST" },
        replyToken: "DIRECT-TEST-NO-LINE-REPLY",
        message: { id: eventId, type: "text", text: String(text) },
    };
}

function validateDirectOrderTestInput(rawMessage, eventId) {
    const text = String(rawMessage || "").trim();
    const normalizedEventId = validateDirectTestEventId(eventId);
    if (/^(ตัดรอบ|เติมสต๊อก|อัปเดตสต๊อก|ยกเลิก|ยืนยัน|คำสั่ง)(?:\s|$)/.test(text)) {
        throw new Error("Direct test ไม่อนุญาตคำสั่งแอดมิน");
    }
    const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    const deliveryDate = normalizeDelivDate(lines[0]);
    const storeName = lines[1] || "";
    if (deliveryDate !== DIRECT_TEST_DATE) {
        throw new Error(`Direct test ใช้ได้เฉพาะวันที่ ${DIRECT_TEST_DATE}`);
    }
    if (!storeName.startsWith("TEST-")) {
        throw new Error('ชื่อร้านสำหรับ direct test ต้องขึ้นต้นด้วย "TEST-"');
    }
    if (lines.length < 3) throw new Error("Direct test ต้องมีสินค้าอย่างน้อย 1 รายการ");

    return { eventId: normalizedEventId, deliveryDate, storeName };
}

function runDirectOrderTest(rawMessage, eventId) {
    const input = validateDirectOrderTestInput(rawMessage, eventId);
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const orderData = parseOrderMessage(
        rawMessage,
        getMappingDictionary(ss),
        getStoreMappingDictionary(ss),
    );
    if (!orderData || orderData.deliveryDate !== input.deliveryDate || orderData.storeName !== input.storeName) {
        throw new Error("Direct test parse ออเดอร์ไม่สำเร็จ");
    }

    handleTextEvent(createDirectTestEvent(rawMessage, input.eventId));

    const orderSheetName = "ออเดอร์-" + input.deliveryDate.replace(/\//g, "-");
    const purchaseSheetName = "ใบซื้อ-" + input.deliveryDate.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);
    const orderRows = orderSheet
        ? orderSheet.getDataRange().getValues()
            .filter((row) => String(row[7] || "") === input.eventId)
            .map((row) => ({
                date: String(row[0] || ""),
                store: String(row[1] || ""),
                product: String(row[2] || ""),
                quantity: row[3],
                unit: String(row[4] || ""),
                matchStatus: String(row[6] || ""),
                itemIndex: row[8],
            }))
        : [];

    const journal = loadEventJournal(ss, `${input.eventId}|ORDER`);
    const logSheet = ss.getSheetByName("Logs");
    const logRows = logSheet && logSheet.getLastRow()
        ? logSheet.getRange(1, 1, logSheet.getLastRow(), 7).getValues()
            .filter((row) => String(row[4] || "") === input.eventId)
        : [];
    const latestLog = logRows.length ? logRows[logRows.length - 1] : [];

    return {
        success: orderRows.length === orderData.items.length &&
            journal && journal.status === "COMPLETE" &&
            !!ss.getSheetByName(purchaseSheetName),
        eventId: input.eventId,
        deliveryDate: input.deliveryDate,
        storeName: input.storeName,
        expectedItemCount: orderData.items.length,
        savedItemCount: orderRows.length,
        orderSheetName,
        purchaseSheetName,
        purchaseSheetCreated: !!ss.getSheetByName(purchaseSheetName),
        journalStatus: journal ? journal.status : "NOT_FOUND",
        logStatus: String(latestLog[3] || "NOT_FOUND"),
        replyState: String(latestLog[5] || "NOT_FOUND"),
        orderRows,
    };
}

function validateDirectCommandTestInput(rawMessage, eventId) {
    const text = String(rawMessage || "").trim();
    const normalizedEventId = validateDirectTestEventId(eventId);
    if (text === `อัปเดตสต๊อก ${DIRECT_TEST_DATE}`) {
        return { eventId: normalizedEventId, type: "STOCK_SYNC" };
    }
    if (text === `ตัดรอบ ${DIRECT_TEST_DATE}`) {
        return { eventId: normalizedEventId, type: "CUTOFF" };
    }
    if (text === "ยกเลิกการลบ") {
        return { eventId: normalizedEventId, type: "CANCEL_ABORT" };
    }
    if (/^ยืนยัน(?:\s+\d+(?:[,\s]+\d+)*)?$/.test(text)) {
        return { eventId: normalizedEventId, type: "CANCEL_CONFIRM" };
    }

    const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    if (lines[0] === `เติมสต๊อก ${DIRECT_TEST_DATE}`) {
        if (lines.length < 2) throw new Error("Direct test เติมสต๊อกต้องมีสินค้า");
        const hasInjectedCommand = lines.slice(1).some(
            (line) => /^(เติมสต๊อก|อัปเดตสต๊อก|ตัดรอบ|ยกเลิก|ยืนยัน|คำสั่ง)(?:\s|$)/.test(line),
        );
        if (hasInjectedCommand) throw new Error("Direct test ไม่อนุญาตคำสั่งซ้อน");
        return { eventId: normalizedEventId, type: "STOCK_FILL" };
    }

    const cancelMatch = text.match(/^ยกเลิก\s+(.+?)\s+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})$/);
    if (cancelMatch) {
        const storeName = cancelMatch[1].trim();
        const deliveryDate = normalizeDelivDate(cancelMatch[2]);
        if (!storeName.startsWith("TEST-") || deliveryDate !== DIRECT_TEST_DATE) {
            throw new Error(`Direct test ยกเลิกได้เฉพาะร้าน TEST- วันที่ ${DIRECT_TEST_DATE}`);
        }
        return { eventId: normalizedEventId, type: "CANCEL_PREVIEW", storeName };
    }

    throw new Error(`Direct test อนุญาตเฉพาะคำสั่งวันที่ ${DIRECT_TEST_DATE}`);
}

function validateDirectPendingCancel(pending) {
    if (!pending || !String(pending.storeName || "").startsWith("TEST-") ||
        normalizeDelivDate(pending.deliveryDate) !== DIRECT_TEST_DATE) {
        throw new Error(`Direct test ยืนยันได้เฉพาะร้าน TEST- วันที่ ${DIRECT_TEST_DATE}`);
    }
    return true;
}

function serializeDirectTestCell(value) {
    if (value instanceof Date) return Utilities.formatDate(value, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    if (value === null || value === undefined) return "";
    return typeof value === "object" ? String(value) : value;
}

function inspectDirectTestState(productNames, eventId) {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const normalizedEventId = eventId ? validateDirectTestEventId(eventId) : "";
    const wantedProducts = new Set((Array.isArray(productNames) ? productNames : [])
        .slice(0, 50)
        .map((name) => String(name || "").replace(/\s+/g, "").trim())
        .filter(Boolean));

    const orderSheetName = "ออเดอร์-" + DIRECT_TEST_DATE.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);
    const orderData = orderSheet ? orderSheet.getDataRange().getValues() : [];
    const testRows = orderData.filter((row) => String(row[1] || "").startsWith("TEST-"));
    const statusCounts = { EXACT: 0, FUZZY: 0, UNMAPPED: 0 };
    testRows.forEach((row) => {
        const status = String(row[6] || "");
        statusCounts[status] = (statusCounts[status] || 0) + 1;
    });

    const purchaseSheetName = "ใบซื้อ-" + DIRECT_TEST_DATE.replace(/\//g, "-");
    const purchaseSheet = ss.getSheetByName(purchaseSheetName);
    const purchaseRows = purchaseSheet
        ? purchaseSheet.getDataRange().getValues()
            .filter((row) => {
                const label = String(row[0] || "");
                return label.startsWith("รอบที่") || wantedProducts.has(label.replace(/\s+/g, ""));
            })
            .slice(0, 250)
            .map((row) => row.slice(0, 9).map(serializeDirectTestCell))
        : [];

    const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
    const stockSheetName = getMonthlyStockTabName(DIRECT_TEST_DATE);
    const stockSheet = stockSS.getSheetByName(stockSheetName);
    const stockData = stockSheet ? stockSheet.getDataRange().getValues() : [];
    const stockSection = stockSheet ? findStockDateSection(stockData, DIRECT_TEST_DATE) : null;
    const stockRows = [];
    if (stockSection) {
        for (let i = stockSection.startIndex; i < stockSection.endIndex; i++) {
            const name = String(stockData[i][1] || "").trim();
            if (!wantedProducts.has(name.replace(/\s+/g, ""))) continue;
            stockRows.push({ row: i + 1, name, quantity: parseFloat(stockData[i][2]) || 0, unit: String(stockData[i][3] || "") });
        }
    }

    const eventJournalSheet = ss.getSheetByName("_EventJournal");
    const eventJournals = eventJournalSheet && normalizedEventId
        ? eventJournalSheet.getDataRange().getValues()
            .filter((row) => String(row[0] || "").startsWith(normalizedEventId + "|"))
            .map((row) => ({ key: String(row[0]), type: String(row[1]), status: String(row[2]), error: String(row[4] || "") }))
        : [];
    const cutoffJournalSheet = ss.getSheetByName("_CutoffJournal");
    const cutoffJournals = cutoffJournalSheet
        ? cutoffJournalSheet.getDataRange().getValues()
            .filter((row) => normalizeDelivDate(row[1]) === DIRECT_TEST_DATE)
            .map((row) => ({ key: String(row[0]), round: Number(row[2]), eventId: String(row[3]), status: String(row[4]), error: String(row[6] || "") }))
        : [];
    const logSheet = ss.getSheetByName("Logs");
    const eventLogs = logSheet && normalizedEventId
        ? logSheet.getDataRange().getValues()
            .filter((row) => String(row[4] || "") === normalizedEventId)
            .map((row) => ({ status: String(row[3] || ""), replyState: String(row[5] || "") }))
        : [];

    return {
        date: DIRECT_TEST_DATE,
        orderSheetName,
        purchaseSheetName,
        stockSheetName,
        orderSheetCreated: !!orderSheet,
        purchaseSheetCreated: !!purchaseSheet,
        stockSectionFound: !!stockSection,
        testStoreCount: new Set(testRows.map((row) => String(row[1]))).size,
        testOrderRowCount: testRows.length,
        currentEventOrderRowCount: normalizedEventId
            ? testRows.filter((row) => String(row[7] || "") === normalizedEventId).length
            : 0,
        roundDividers: orderData.filter((row) => /^รอบ\s+\d+$/.test(String(row[0] || "").trim())).map((row) => String(row[0]).trim()),
        statusCounts,
        stockRows,
        purchaseRows,
        eventJournals,
        cutoffJournals,
        eventLogs,
    };
}

function runDirectCommandTest(rawMessage, eventId, productNames) {
    const input = validateDirectCommandTestInput(rawMessage, eventId);
    if (input.type === "CANCEL_CONFIRM") {
        const pendingJson = PropertiesService.getScriptProperties()
            .getProperty("PENDING_CANCEL_U-DIRECT-INTEGRATION-TEST");
        validateDirectPendingCancel(pendingJson ? JSON.parse(pendingJson) : null);
    }
    handleTextEvent(createDirectTestEvent(rawMessage, input.eventId));
    return inspectDirectTestState(productNames, input.eventId);
}
