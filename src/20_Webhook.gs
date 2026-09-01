// ==============================================================================
// 🚀 WEBHOOK
// ==============================================================================
function doPost(e) {
    try {
        const eventData = JSON.parse(e.postData.contents);
        if (!eventData.events || !eventData.events.length) return replySuccess();

        eventData.events.forEach((event) => {
            if (event.type === "message" && event.message && event.message.type === "text") {
                setActiveLogContext(null, 0);
                ACTIVE_EVENT_ID = "";
                try {
                    handleTextEvent(event);
                } catch (err) {
                    if (err.code === "LOCK_TIMEOUT") {
                        recordEventFailure(event, "Lock Timeout");
                        replyToLine(event.replyToken, `⏳ ${err.message}`);
                        return;
                    }
                    recordEventFailure(event, `Error: ${err.message}`);
                    replyToLine(event.replyToken, "❌ ระบบประมวลผลไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
                }
            }
        });
        return replySuccess();
    } catch (err) {
        console.error("Webhook Error:", err.message);
        return ContentService.createTextOutput(
            JSON.stringify({ status: "error", message: err.message }),
        ).setMimeType(ContentService.MimeType.JSON);
    }
}

function handleTextEvent(event) {
        const rawMessage = event.message.text.trim();
        const replyToken = event.replyToken;
        const timestamp = new Date();
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        setActiveLogContext(null, 0);
        ACTIVE_EVENT_ID = "";
        const eventLog = beginEventLog(ss, event, rawMessage, timestamp);
        if (eventLog.duplicate) return replySuccess();
        ACTIVE_EVENT_ID = eventLog.eventId;
        const logSheet = eventLog.logSheet;
        const logRow = eventLog.logRow;
        setActiveLogContext(logSheet, logRow);
        if (eventLog.retryReply) {
            replyToLine(replyToken, eventLog.retryReply);
            return replySuccess();
        }

        // ─── Help: คู่มือคำสั่ง ───────────────────────────────────────────────────
        if (rawMessage.trim() === "คำสั่ง") {
            const helpText = `📋 รวมคำสั่งใช้งานบอท

1️⃣ วิธีสังสินค้า (ขึ้นบรรทัดใหม่ตามที่ระบุ):
27/05/69
ร้านกะเพราปลา
คะน้า 5 กก
แตงกวา 2 กก

2️⃣ แจ้งตัดรอบสั่งของ:
ตัดรอบ 27/05/69

3️⃣ แจ้งอัปเดตสต๊อก:
อัปเดตสต๊อก 27/05/69

4️⃣ ยกเลิกออเดอร์:
ยกเลิก ร้านกะเพราปลา 27/05/69
`;
            safeSetLog(logSheet, logRow, "Help");
            replyToLine(replyToken, helpText);
            return replySuccess();
        }

        // ─── Admin: เติมสต๊อก ────────────────────────────────────────────────────
        if (rawMessage.startsWith("เติมสต๊อก")) {
            const result = withScriptLock(() => {
                const mappingDict = getMappingDictionary(ss);
                const stockResult = processStockUpdate(rawMessage, ss, mappingDict, ACTIVE_EVENT_ID);
                if (!stockResult.success) return stockResult;
                const orderSheetName = "ออเดอร์-" + stockResult.deliveryDate.replace(/\//g, "-");
                if (stockResult.deliveryDate && ss.getSheetByName(orderSheetName)) {
                    updatePurchaseSummarySheet(ss, stockResult.deliveryDate, false, true);
                }
                return stockResult;
            });
            if (!result.success) {
                safeSetLog(logSheet, logRow, `Stock Update Failed: ${result.message}`);
                replyToLine(replyToken, `❌ อัปเดตสต๊อกไม่สำเร็จ: ${result.message}`);
                return replySuccess();
            }
            safeSetLog(logSheet, logRow, `Stock Updated: ${result.count} items`);
            replyToLine(replyToken, `✅ อัปเดตสต๊อกแล้ว ${result.count} รายการ`);
            return replySuccess();
        }

        // ─── Admin: ตัดรอบ ────────────────────────────────────────────────────────
        if (rawMessage.startsWith("ตัดรอบ")) {
            const parts = rawMessage.split(/\s+/);

            if (parts.length < 2) {
                replyToLine(replyToken, `❌ กรุณาระบุวันที่ที่ต้องการตัดรอบด้วยครับ (ตัวอย่าง: ตัดรอบ 27/05/69)`);
                return replySuccess();
            }

            const targetDateInput = parts[1].trim();
            const targetDate = normalizeDelivDate(targetDateInput);

            if (!targetDate) {
                replyToLine(replyToken, `❌ รูปแบบวันที่ไม่ถูกต้อง (ตัวอย่าง: ตัดรอบ 27/05/69)`);
                return replySuccess();
            }

            const result = withScriptLock(() => executeCutoff(ss, targetDate, ACTIVE_EVENT_ID));

            if (result.recoveryRetryFailed) {
                safeSetLog(logSheet, logRow, `Cutoff Recovery Retry Failed: ${targetDate}; Round ${result.recoveredRound}; ${result.summaryError}`);
                replyToLine(
                    replyToken,
                    `⚠️ รอบที่ ${result.recoveredRound} ของวันที่ ${targetDate} ปิดไปแล้วและไม่ได้หักสต๊อกซ้ำ\n` +
                    `แต่ซ่อมใบซื้อยังไม่สำเร็จ: ${result.summaryError}\n` +
                    `กรุณาพิมพ์ตัดรอบอีกครั้ง`,
                );
                return replySuccess();
            }

            if (result.recoveredOnly) {
                if (result.followUpError) {
                    safeSetLog(logSheet, logRow, `Cutoff Recovery Success: ${targetDate}; Recovered ${result.recoveredRound}; Follow-up Failed: ${result.followUpError}`);
                    replyToLine(
                        replyToken,
                        `✅ ซ่อมใบซื้อรอบที่ ${result.recoveredRound} สำหรับวันที่ ${targetDate} สำเร็จแล้ว\n` +
                        `⚠️ แต่ตัดรอบใหม่ยังไม่สมบูรณ์: ${result.followUpError}\n` +
                        `กรุณาพิมพ์ตัดรอบอีกครั้ง ระบบจะทำต่อโดยไม่หักสต๊อกซ้ำ`,
                    );
                } else {
                    safeSetLog(logSheet, logRow, `Cutoff Recovery Success: ${targetDate}; Recovered ${result.recoveredRound}; No New Orders`);
                    replyToLine(
                        replyToken,
                        `✅ ซ่อมใบซื้อรอบที่ ${result.recoveredRound} สำหรับวันที่ ${targetDate} สำเร็จแล้ว\n` +
                        `ℹ️ ไม่มีออเดอร์รอบใหม่ให้ตัด`,
                    );
                }
                return replySuccess();
            }

            if (result.syncError) {
                safeSetLog(logSheet, logRow, `Cutoff Sync Failed: ${targetDate}`);
                replyToLine(replyToken, `❌ ไม่สามารถตัดรอบได้ เพราะอ่านสต๊อกล่าสุดไม่สำเร็จ\n${result.syncError}`);
                return replySuccess();
            }

            if (!result.deductRes.success) {
                safeSetLog(logSheet, logRow, `Cutoff Failed: ${result.deductRes.message}`);
                replyToLine(replyToken, `❌ ${result.deductRes.message}`);
                return replySuccess();
            }

            if (result.res.success) {
                const cutoffStatus = `Cutoff Success: ${targetDate}; Closed ${result.res.closedRound}; Opened ${result.res.round}; ` +
                    `Full ${result.deductRes.fullDeductedCount || 0}; Partial ${(result.deductRes.partialItems || []).length}; ` +
                    `No Stock ${result.deductRes.zeroStockCount || 0}; Unmatched ${result.deductRes.unmatchedCount || 0}` +
                    (result.summaryError ? `; Summary Warning: ${result.summaryError}` : "") +
                    (result.recoveredRound ? `; Recovered ${result.recoveredRound}` : "");
                safeSetLog(logSheet, logRow, cutoffStatus);
                const fmtQty = (qty) => Math.round(qty * 1000) / 1000;
                const deductedItems = result.deductRes.deductedItems || [];
                const zeroStockItems = result.deductRes.zeroStockItems || [];
                const unmatchedItems = result.deductRes.unmatchedItems || [];
                const deductedDetail = `\n\n📋 รายการที่หักจากสต๊อก\n` +
                    (deductedItems.length
                        ? deductedItems
                            .map((item) => `• ${item.name}: ${fmtQty(item.deducted)} ${item.unit}` +
                                (item.shortage > 0 ? ` (ขาด ${fmtQty(item.shortage)} ${item.unit} — ต้องซื้อเพิ่ม)` : ""))
                            .join("\n")
                        : "• ไม่มีรายการที่หักได้");
                const noStockDetail = zeroStockItems.length
                    ? `\n\n📦 ไม่มีในสต๊อก\n` + zeroStockItems
                        .map((item) => `• ${item.name}: ${fmtQty(item.requested)} ${item.unit} — ต้องซื้อเพิ่ม`)
                        .join("\n")
                    : "";
                const reviewLines = unmatchedItems
                    .map((item) => `• ${item.name}: สั่ง ${fmtQty(item.requested)} ${item.orderUnit} (${item.reason})`);
                const reviewDetail = reviewLines.length
                    ? `\n\n⚠️ ต้องตรวจสอบ\n${reviewLines.join("\n")}`
                    : "";
                const stockDetail = deductedDetail + noStockDetail + reviewDetail;
                const summaryWarning = result.summaryError
                    ? `\n⚠️ หักสต๊อกและปิดรอบแล้ว แต่ใบซื้ออัปเดตไม่สำเร็จ: ${result.summaryError}`
                    : "";
                const recoveryNote = result.recoveredRound
                    ? `\n♻️ ซ่อมใบซื้อรอบที่ ${result.recoveredRound} สำเร็จแล้วในคำสั่งเดียวกัน`
                    : "";
                replyToLine(replyToken, `✅ ตัดรอบที่ ${result.res.closedRound} สำหรับวันที่ ${targetDate} เรียบร้อยแล้ว${stockDetail}${summaryWarning}${recoveryNote}`);
            } else {
                safeSetLog(logSheet, logRow, `Cutoff Failed: ${targetDate}; order sheet not found`);
                replyToLine(replyToken, `❌ ไม่สามารถสร้างแถวตัดรอบได้ (ยังไม่มีใบจัดออเดอร์ของวันที่ ${targetDate})`);
            }
            return replySuccess();
        }

        // ─── Admin: อัปเดตสต๊อก ──────────────────────────────────────────────────
        if (rawMessage.startsWith("อัปเดตสต๊อก")) {
            const parts = rawMessage.split(/\s+/);
            const targetDate = parts.length > 1 ? normalizeDelivDate(parts[1].trim()) : null;
            if (!targetDate) {
                replyToLine(replyToken, `❌ กรุณาระบุวันที่ เช่น "อัปเดตสต๊อก 14/04/69"`);
                return replySuccess();
            }
            let syncError = null;
            try {
                withScriptLock(() => updatePurchaseSummarySheet(ss, targetDate, false, true));
            } catch (err) {
                syncError = err.message;
            }
            if (syncError) {
                safeSetLog(logSheet, logRow, `Stock Sync Failed: ${targetDate}; ${syncError}`);
                replyToLine(replyToken, `❌ ซิงค์ใบซื้อและหน้ารายการไม่สำเร็จ\n${syncError}`);
                return replySuccess();
            }
            safeSetLog(logSheet, logRow, `Stock Synced: ${targetDate}`);
            replyToLine(replyToken, `✅ ซิงค์ใบซื้อและหน้ารายการของวันที่ ${targetDate} เรียบร้อยแล้ว!`);
            return replySuccess();
        }

        // ─── Admin: ยกเลิกการลบ ───────────────────────────────────────────────────
        if (rawMessage.trim() === "ยกเลิกการลบ") {
            const props = PropertiesService.getScriptProperties();
            props.deleteProperty("PENDING_CANCEL_" + getSourceUserId(event));
            safeSetLog(logSheet, logRow, "Cancel Aborted");
            replyToLine(replyToken, "✅ ยกเลิกการลบเรียบร้อยแล้ว");
            return replySuccess();
        }

        // ─── Admin: ยืนยัน (ลบออเดอร์ที่ค้างอยู่) ────────────────────────────────
        if (rawMessage.trim().startsWith("ยืนยัน")) {
            const props = PropertiesService.getScriptProperties();
            const pendingKey = "PENDING_CANCEL_" + getSourceUserId(event);
            const pendingJson = props.getProperty(pendingKey);
            if (!pendingJson) {
                replyToLine(replyToken, '❌ ไม่มีรายการที่รอยกเลิก กรุณาพิมพ์ "ยกเลิก ชื่อร้าน วันที่" ก่อน');
                return replySuccess();
            }
            const pending = JSON.parse(pendingJson);

            // Timeout 10 นาที
            if (Date.now() - pending.timestamp > 10 * 60 * 1000) {
                props.deleteProperty(pendingKey);
                replyToLine(replyToken, '❌ หมดเวลายืนยัน กรุณาพิมพ์ "ยกเลิก ชื่อร้าน วันที่" ใหม่');
                return replySuccess();
            }

            // Parse เลขรายการ (ถ้ามี)
            const numPart = rawMessage.trim().replace(/^ยืนยัน\s*/, "").trim();
            let itemIndices = null;
            if (numPart) {
                itemIndices = numPart.split(/[,\s]+/).map(Number).filter(n => n > 0);
                if (itemIndices.length === 0) itemIndices = null;
            }

            const result = withScriptLock(() => {
                const cancelResult = cancelOrderIfSnapshotMatches(ss, pending, itemIndices);
                if (cancelResult.success) {
                    try {
                        updatePurchaseSummarySheet(ss, pending.deliveryDate, false, true);
                    } catch (err) {
                        cancelResult.summaryError = err.message;
                    }
                }
                return cancelResult;
            });
            if (result.stale) {
                props.setProperty(pendingKey, JSON.stringify({
                    ...pending,
                    timestamp: Date.now(),
                    items: result.items,
                    snapshotHash: result.snapshotHash,
                }));
                safeSetLog(logSheet, logRow, `Cancel Preview Refreshed: ${pending.storeName} ${pending.deliveryDate}`);
                replyToLine(replyToken, formatCancelPreview(
                    pending.storeName,
                    pending.deliveryDate,
                    result.items,
                    "⚠️ รายการออเดอร์เปลี่ยนแล้ว นี่คือรายการล่าสุด",
                ));
                return replySuccess();
            }
            props.deleteProperty(pendingKey);

            if (result.success) {
                safeSetLog(logSheet, logRow, `Cancelled: ${pending.storeName} ${pending.deliveryDate} (${result.deletedCount} items)` +
                    (result.summaryError ? `; Summary Warning: ${result.summaryError}` : ""));
                let msg = itemIndices
                    ? `✅ ยกเลิกสำเร็จ! ลบ ${result.deletedCount} รายการจากร้าน "${pending.storeName}" วันที่ ${pending.deliveryDate}`
                    : `✅ ยกเลิกออเดอร์ร้าน "${pending.storeName}" วันที่ ${pending.deliveryDate} ทั้งหมดสำเร็จ! (${result.deletedCount} รายการ)`;
                if (result.summaryError) msg += `\n⚠️ ลบออเดอร์แล้ว แต่ใบซื้ออัปเดตไม่สำเร็จ: ${result.summaryError}`;
                replyToLine(replyToken, msg);
            } else {
                replyToLine(replyToken, `❌ ${result.message}`);
            }
            return replySuccess();
        }

        // ─── Admin: ยกเลิกออเดอร์ (แสดงรายการ + เก็บ state) ─────────────────────
        if (rawMessage.startsWith("ยกเลิก")) {
            const dateMatch = rawMessage.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
            if (!dateMatch) {
                replyToLine(replyToken, '❌ กรุณาระบุวันที่ เช่น "ยกเลิก ร้านABC 04/05/69"');
                return replySuccess();
            }
            const deliveryDate = normalizeDelivDate(dateMatch[1]);
            if (!deliveryDate) { replyToLine(replyToken, "❌ รูปแบบวันที่ไม่ถูกต้อง"); return replySuccess(); }

            const storeName = rawMessage
                .replace(/^ยกเลิก\s*/, "")
                .replace(/\s*\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\s*$/, "").trim();

            if (!storeName) { replyToLine(replyToken, '❌ กรุณาระบุชื่อร้าน เช่น "ยกเลิก ร้านABC 04/05/69"'); return replySuccess(); }

            const orderSheetName = "ออเดอร์-" + deliveryDate.replace(/\//g, "-");
            const orderSheet = ss.getSheetByName(orderSheetName);
            if (!orderSheet) { replyToLine(replyToken, `❌ ไม่พบใบออเดอร์ของวันที่ ${deliveryDate}`); return replySuccess(); }

            const storeOrders = findStoreOrders(orderSheet, storeName);
            if (storeOrders.length === 0) { replyToLine(replyToken, `❌ ไม่พบออเดอร์ของร้าน "${storeName}" ในวันที่ ${deliveryDate}`); return replySuccess(); }

            // เก็บ state รอยืนยัน
            const props = PropertiesService.getScriptProperties();
            props.setProperty("PENDING_CANCEL_" + getSourceUserId(event), JSON.stringify({
                storeName,
                deliveryDate,
                timestamp: Date.now(),
                items: storeOrders,
                snapshotHash: getOrderSnapshotHash(storeOrders),
            }));

            safeSetLog(logSheet, logRow, `Cancel Preview: ${storeName} ${deliveryDate}`);
            replyToLine(replyToken, formatCancelPreview(storeName, deliveryDate, storeOrders));
            return replySuccess();
        }

        // ─── Order ────────────────────────────────────────────────────────────────
        const mappingDict = getMappingDictionary(ss);
        const storeMappingDict = getStoreMappingDictionary(ss);
        const orderData = parseOrderMessage(rawMessage, mappingDict, storeMappingDict);

        if (orderData && orderData.items.length > 0) {
            const matchNotes = getOrderMatchNotes(orderData.items);
            const storeMappingLog = formatStoreMappingLog(orderData.storeMappingSummary);

            const orderResult = withScriptLock(() => {
                processOrderUpdate(rawMessage, ss, orderData, timestamp, ACTIVE_EVENT_ID);

                let stockSectionWarning = null;
                try {
                    ensureStockDateSection(orderData.deliveryDate);
                } catch (err) {
                    // การสร้างหัววันที่ในไฟล์สต๊อกไม่ควรขวางใบซื้อ
                    stockSectionWarning = err.message;
                }

                try {
                    updatePurchaseSummarySheet(ss, orderData.deliveryDate, false, true);
                    return { success: true, stockSectionWarning };
                } catch (err) {
                    return { success: false, message: err.message };
                }
            });

            if (!orderResult.success) {
                safeSetLog(logSheet, logRow, [`Order Saved; Purchase Failed: ${orderResult.message}`, storeMappingLog].filter(Boolean).join("; "));
                replyToLine(replyToken, `⚠️ บันทึกใบออเดอร์แล้ว แต่สร้างใบซื้อไม่สำเร็จ\n${orderResult.message}`);
                return replySuccess();
            }
            const baseSuccessStatus = orderResult.stockSectionWarning
                ? `Success; Stock Section Warning: ${orderResult.stockSectionWarning}`
                : "Success";
            const successStatus = [baseSuccessStatus, storeMappingLog].filter(Boolean).join("; ");
            safeSetLog(logSheet, logRow, successStatus);

            const itemLines = orderData.items
                .map((i) => {
                    const prefix = i.isWarning ? "⚠️ " : i.isFuzzy ? "🔍 " : "• ";
                    const displayUnit = i.unit === "กก" ? "กก." : i.unit;
                    return `${prefix}${i.name}  ${i.amount} ${displayUnit}`;
                })
                .join("\n");
            const stockNote = orderResult.stockSectionWarning
                ? `\n\n⚠️ สร้างใบซื้อแล้ว แต่สร้างหัววันที่ในไฟล์สต๊อกไม่สำเร็จ: ${orderResult.stockSectionWarning}`
                : "";

            replyToLine(
                replyToken,
                `✅ รับออเดอร์แล้วครับ!\n🏪 ${orderData.storeName}\n📅 สั่ง: ${orderData.deliveryDate}\n\nรายการ:\n${itemLines}${matchNotes}${stockNote}`,
            );
        } else {
            safeSetLog(logSheet, logRow, "Failed to Parse");
            replyToLine(replyToken, "❌ ไม่สามารถอ่านออเดอร์ได้ กรุณาตรวจสอบรูปแบบข้อความ");
        }

        return replySuccess();
}

function replySuccess() {
    return ContentService.createTextOutput(
        JSON.stringify({ status: "success" }),
    ).setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================================
// 💬 LINE REPLY
// ==============================================================================
function replyToLine(replyToken, message) {
    LAST_LINE_REPLY_ERROR = "";
    setActiveReplyState("PENDING", message);
    try {
        const token = CONFIG.LINE_TOKEN;
        if (!token) throw new Error("ยังไม่ได้ตั้ง Script Property: LINE_TOKEN");

        const chunks = [];
        let chunk = "";
        for (const line of String(message).split("\n")) {
            const next = chunk ? `${chunk}\n${line}` : line;
            if (next.length > 4900 && chunk) {
                chunks.push(chunk);
                chunk = line;
            } else {
                chunk = next;
            }
        }
        if (chunk) chunks.push(chunk);
        // ponytail: LINE reply รองรับสูงสุด 5 ข้อความ; ถ้าเกินค่อยเปลี่ยนไปใช้ push API
        if (chunks.length > 5) {
            chunks.length = 5;
            chunks[4] = chunks[4].slice(0, 4800) + "\n…รายการที่เหลือดูได้ในใบซื้อ";
        }

        const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
            method: "post",
            headers: {
                "Content-Type": "application/json",
                Authorization: "Bearer " + token,
            },
            payload: JSON.stringify({
                replyToken,
                messages: chunks.map((text) => ({ type: "text", text })),
            }),
            muteHttpExceptions: true,
        });

        const status = response.getResponseCode();
        if (status < 200 || status >= 300) {
            throw new Error(`LINE HTTP ${status}: ${response.getContentText().slice(0, 300)}`);
        }
        const firstLine = String(message).split("\n")[0].slice(0, 160);
        setActiveReplyState("COMPLETE", message);
        finishActiveLog("COMPLETE", firstLine.startsWith("❌") ? `Rejected: ${firstLine}` : "LINE Reply Success");
        return true;
    } catch (err) {
        LAST_LINE_REPLY_ERROR = err.message || String(err);
        setActiveReplyState("FAILED", message);
        finishActiveLog("FAILED", `LINE Reply Failed: ${LAST_LINE_REPLY_ERROR}`);
        console.error("LINE Reply Error:", LAST_LINE_REPLY_ERROR);
        return false;
    }
}
