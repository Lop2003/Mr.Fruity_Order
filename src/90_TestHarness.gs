// ==============================================================================
// 🧪 OWNER-ONLY INTEGRATION TEST
// ==============================================================================

const DIRECT_TEST_DATE = "31/12/99";

// เก็บ TEST- rows ไว้เป็นหลักฐาน; LINE reply failure เป็น expected เพราะไม่มี reply token จริง

function validateDirectOrderTestInput(rawMessage, eventId) {
    const text = String(rawMessage || "").trim();
    const normalizedEventId = String(eventId || "").trim();
    if (/^(ตัดรอบ|เติมสต๊อก|อัปเดตสต๊อก|ยกเลิก|ยืนยัน|คำสั่ง)(?:\s|$)/.test(text)) {
        throw new Error("Direct test ไม่อนุญาตคำสั่งแอดมิน");
    }
    if (!/^test-[A-Za-z0-9_-]{8,80}$/.test(normalizedEventId)) {
        throw new Error("Event ID สำหรับ direct test ไม่ถูกต้อง");
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

    handleTextEvent({
        type: "message",
        timestamp: Date.now(),
        webhookEventId: input.eventId,
        source: { type: "user", userId: "U-DIRECT-INTEGRATION-TEST" },
        replyToken: "DIRECT-TEST-NO-LINE-REPLY",
        message: { id: input.eventId, type: "text", text: String(rawMessage) },
    });

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
