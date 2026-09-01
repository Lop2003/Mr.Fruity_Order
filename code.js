// ==============================================================================
// 🍎 Order System — Fix-04/07/69
// ==============================================================================

const CONFIG = {
    SHEET_ID: "1JAxtrpjvhxJ7DmtTL59DktJ2K1nFjrmDsBuD4UaGcJI",
    STOCK_FILE_ID: "1GAKGQ_n2OXO2lAGCs4aEpyH9sJYitmHyY76NwH8jo2o",
    get LINE_TOKEN() {
        return PropertiesService.getScriptProperties().getProperty("LINE_TOKEN");
    },
    SECTION_BG: "#1B5E20",
    SECTION_FG: "#FFFFFF",
    WARNING_BG: "#FFCDD2",
    WARNING_FG: "#B71C1C",
    // [FIX-7] cache TTL (วินาที)
    MAPPING_CACHE_TTL: 300,
    // [FIX-7] fuzzy threshold ratio
    FUZZY_RATIO: 0.25,
    FUZZY_MIN_DIST: 1,
    SUBSEQ_MIN_RATIO: 0.6,
    // [FIX-7] lock timeout (ms)
    LOCK_TIMEOUT_MS: 10000,
};

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
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
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
        .setNumberFormat("@")
        .setValue(deliveryDateStr)
        .setBackground("#FFF2CC")
        .setFontWeight("bold");

    SpreadsheetApp.flush();
    data = stockSheet.getDataRange().getValues();
    section = findStockDateSection(data, deliveryDateStr);
    return { stockSS, stockSheet, section };
}

// ==============================================================================
// 🚀 WEBHOOK
// ==============================================================================
function doPost(e) {
    try {
        const eventData = JSON.parse(e.postData.contents);
        if (!eventData.events || !eventData.events.length) return replySuccess();

        eventData.events.forEach((event) => {
            if (event.type === "message" && event.message && event.message.type === "text") {
                handleTextEvent(event);
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
        const logSheet = getOrCreateSheet(ss, "Logs");

        logSheet.appendRow([timestamp, getSourceUserId(event), rawMessage, "Processing"]);
        const logRow = logSheet.getLastRow();

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
            replyToLine(replyToken, helpText);
            return replySuccess();
        }

        // ─── Admin: เติมสต๊อก ────────────────────────────────────────────────────
        if (rawMessage.startsWith("เติมสต๊อก")) {
            const result = withScriptLock(() => {
                const mappingDict = getMappingDictionary(ss);
                const stockResult = processStockUpdate(rawMessage, ss, mappingDict);
                const orderSheetName = "ออเดอร์-" + stockResult.deliveryDate.replace(/\//g, "-");
                if (stockResult.deliveryDate && ss.getSheetByName(orderSheetName)) {
                    updatePurchaseSummarySheet(ss, stockResult.deliveryDate);
                }
                return stockResult;
            });
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

            const result = withScriptLock(() => {
                const deductRes = deductStockForCutoff(ss, targetDate);
                if (!deductRes.success) return { deductRes };

                const res = createCutoffRoundDivider(ss, targetDate);
                if (res.success) updatePurchaseSummarySheet(ss, targetDate, false);
                return { deductRes, res };
            });

            if (!result.deductRes.success) {
                replyToLine(replyToken, `❌ ${result.deductRes.message}`);
                return replySuccess();
            }

            if (result.res.success) {
                safeSetLog(logSheet, logRow, `Cutoff Added: ${targetDate} Round ${result.res.round}`);
                replyToLine(replyToken, `✅ ตัดรอบที่ ${result.res.closedRound} สำหรับวันที่ ${targetDate} และหักสต็อกเรียบร้อยแล้ว`);
            } else {
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
            withScriptLock(() => updatePurchaseSummarySheet(ss, targetDate));
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
                const cancelResult = cancelOrder(ss, pending.deliveryDate, pending.storeName, itemIndices);
                if (cancelResult.success) updatePurchaseSummarySheet(ss, pending.deliveryDate);
                return cancelResult;
            });
            props.deleteProperty(pendingKey);

            if (result.success) {
                safeSetLog(logSheet, logRow, `Cancelled: ${pending.storeName} ${pending.deliveryDate} (${result.deletedCount} items)`);
                const msg = itemIndices
                    ? `✅ ยกเลิกสำเร็จ! ลบ ${result.deletedCount} รายการจากร้าน "${pending.storeName}" วันที่ ${pending.deliveryDate}`
                    : `✅ ยกเลิกออเดอร์ร้าน "${pending.storeName}" วันที่ ${pending.deliveryDate} ทั้งหมดสำเร็จ! (${result.deletedCount} รายการ)`;
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
                storeName, deliveryDate, timestamp: Date.now()
            }));

            const fmtU = (u) => u === "กก" ? "กก." : u;
            const itemList = storeOrders.map((item, idx) => `${idx + 1}. ${item.name}  ${item.amount} ${fmtU(item.unit)}`).join("\n");

            safeSetLog(logSheet, logRow, `Cancel Preview: ${storeName} ${deliveryDate}`);
            replyToLine(replyToken,
                `📋 พบออเดอร์ร้าน "${storeName}" วันที่ ${deliveryDate}\n\n${itemList}\n\n` +
                `✏️ พิมพ์ "ยืนยัน" เพื่อยกเลิกทั้งหมด\n` +
                `✏️ พิมพ์ "ยืนยัน 1,3" เพื่อเลือกยกเลิกบางรายการ\n` +
                `✏️ พิมพ์ "ยกเลิกการลบ" เพื่อยกเลิก`
            );
            return replySuccess();
        }

        // ─── Order ────────────────────────────────────────────────────────────────
        const mappingDict = getMappingDictionary(ss);
        const orderData = parseOrderMessage(rawMessage, mappingDict);

        if (orderData && orderData.items.length > 0) {
            const hasWarning = orderData.items.some((i) => i.isWarning);
            const hasFuzzy = orderData.items.some((i) => i.isFuzzy);

            withScriptLock(() => {
                updateDeliverySheet(ss, orderData.deliveryDate, orderData.storeName, orderData.items, timestamp);
                ensureStockDateSection(orderData.deliveryDate);
                updatePurchaseSummarySheet(ss, orderData.deliveryDate);
            });
            safeSetLog(logSheet, logRow, "Success");

            const itemLines = orderData.items
                .map((i) => {
                    const prefix = i.isWarning ? "⚠️ " : i.isFuzzy ? "🔍 " : "• ";
                    const displayUnit = i.unit === "กก" ? "กก." : i.unit;
                    return `${prefix}${i.name}  ${i.amount} ${displayUnit}`;
                })
                .join("\n");
            const warnNote = hasWarning ? "\n\n⚠️ บางรายการไม่พบใน Mapping กรุณาตรวจสอบ" : "";
            const fuzzyNote = hasFuzzy ? "\n\n🔍 บางรายการใช้การจับคู่อัตโนมัติ กรุณาตรวจสอบความถูกต้อง" : "";

            replyToLine(
                replyToken,
                `✅ รับออเดอร์แล้วครับ!\n🏪 ${orderData.storeName}\n📅 สั่ง: ${orderData.deliveryDate}\n\nรายการ:\n${itemLines}${fuzzyNote}${warnNote}`,
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
    try {
        UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
            method: "post",
            headers: {
                "Content-Type": "application/json",
                Authorization: "Bearer " + CONFIG.LINE_TOKEN,
            },
            payload: JSON.stringify({
                replyToken,
                messages: [{ type: "text", text: message }],
            }),
        });
    } catch (err) {
        console.error("LINE Reply Error:", err.message);
    }
}

// ==============================================================================
// 🧠 ORDER PARSER
// ==============================================================================
function normalizeUnits(text) {
    return (
        text
            .replace(/\r/g, "")
            .replace(/เบอร์\s+(\d+)/g, "เบอร์$1") // รวม "เบอร์ 2" → "เบอร์2" ให้เป็นส่วนของชื่อสินค้า
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:กิโลกรัม|กิโล|โล|kilogram|kilo)(?![a-zA-Z\u0E00-\u0E7F])/gi, "$1 กก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*k\.?g\.?(?![a-zA-Z])/gi, "$1 กก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ก\.ก\.?/g, "$1 กก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*กก\./g, "$1 กก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*กก(?![\u0E00-\u0E7Fa-zA-Z])/g, "$1 กก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:กรัม|gram|gr)\.?(?![a-zA-Z])/gi, "$1 กรัม")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ก\.(?!ก)/g, "$1 กรัม")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ก(?![\u0E00-\u0E7Fa-zA-Z.])/g, "$1 กรัม")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*g(?![\u0E00-\u0E7Fa-zA-Z])/gi, "$1 กรัม")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ขีด/g, "$1 ขีด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:แพ็ค|แพ็ก|แพค|แพก|pack)s?/gi, "$1 แพ็ค")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:มัด|bunch(?:es)?)/gi, "$1 มัด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:กล่อง|box(?:es)?)/gi, "$1 กล่อง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ถุง|bag)s?/gi, "$1 ถุง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ขวด|bottle|btl)s?/gi, "$1 ขวด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ลัง|crate)s?/gi, "$1 ลัง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ลูก/g, "$1 ลูก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ผล/g, "$1 ผล")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:แผง|tray)s?/gi, "$1 แผง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ห่อ/g, "$1 ห่อ")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ตัว/g, "$1 ตัว")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ก้อน/g, "$1 ก้อน")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ถาด/g, "$1 ถาด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ม้วน|roll)s?/gi, "$1 ม้วน")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*อัน/g, "$1 อัน")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ชิ้น|pcs?|pieces?)/gi, "$1 ชิ้น")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*กระบอก/g, "$1 กระบอก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:กระป๋อง|ป๋อง)s?/gi, "$1 กระป๋อง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*กำ(?![\u0E00-\u0E7F])/g, "$1 กำ")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ช่อ/g, "$1 ช่อ")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*เข่ง/g, "$1 เข่ง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:โหล|dozen)/gi, "$1 โหล")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*หัว/g, "$1 หัว")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*เม็ด/g, "$1 เม็ด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ชุด|set)s?/gi, "$1 ชุด")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:คู่|pair)s?/gi, "$1 คู่")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ซอง|sachet)s?/gi, "$1 ซอง")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:ลิตร|liter|litre|lt)\.?(?![a-zA-Z])/gi, "$1 ลิตร")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*l(?![a-zA-Z\u0E00-\u0E7F])/gi, "$1 ลิตร")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*(?:มล|มิลลิลิตร|milliliter|millilitre|ml)\.?(?![a-zA-Z])/gi, "$1 มล")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ฝัก/g, "$1 ฝัก")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*กระสอบ/g, "$1 กระสอบ")
            .replace(/(\d+(?:\.\d+)?)[^\S\n]*ฟอง/g, "$1 ฟอง")
    );
}

function parseOrderMessage(rawMessage, mappingDict) {
    let text = normalizeUnits(rawMessage);
    const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
    if (!lines.length) return null;

    const DATE_RE = /(?<![\/\-\d])(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?![\/\-\d])/;
    const firstLine = lines[0];
    const dateMatch = firstLine.match(DATE_RE);

    let storeName = "";
    let deliveryDateStr = "";

    if (dateMatch) {
        deliveryDateStr = normalizeDelivDate(dateMatch[0]);
        storeName = firstLine
            .replace(dateMatch[0], "")
            .replace(/ส่งวันที่|วันที่|ส่ง/g, "")
            .trim();
        // ถ้าบรรทัดแรกมีแค่วันที่ (ไม่มีชื่อร้าน) → หาชื่อร้านจากบรรทัดถัดไป
        if (!storeName && lines.length > 1) {
            for (let i = 1; i < lines.length; i++) {
                if (!DATE_RE.test(lines[i])) {
                    storeName = lines[i];
                    lines.splice(i, 1); // เอาบรรทัดชื่อร้านออกจากรายการ ไม่ให้ถูก parse เป็นสินค้า
                    break;
                }
            }
        }
    } else {
        storeName = firstLine;
        deliveryDateStr = getTodayStr();
        for (let i = 1; i < lines.length; i++) {
            const dm = lines[i].match(DATE_RE);
            if (dm) {
                deliveryDateStr = normalizeDelivDate(dm[0]);
                break;
            }
        }
    }

    const UNIT_LIST =
        "กก|กรัม|ขีด|แพ็ค|มัด|กล่อง|ถุง|ขวด|ลัง|ลูก|ผล|แผง|ห่อ|ตัว|ก้อน|ถาด|ม้วน|อัน|ชิ้น|กระบอก|กระป๋อง|กำ|ช่อ|เข่ง|โหล|หัว|เม็ด|ชุด|คู่|ซอง|ลิตร|มล|ฝัก|กระสอบ|ฟอง|แกลลอน|หวี|แถว";
    const ITEM_RE = new RegExp(
        `^(.+?)\\s+(\\d+(?:\\.\\d+)?)(?!\\d*[*xX\u00D7])\\s*(${UNIT_LIST})?\\s*(.*)$`,
        "i",
    );
    const EXTRA_QTY_RE = new RegExp(`^(\\+)?\\s*(\\d+(?:\\.\\d+)?)\\s*(${UNIT_LIST})\\s*(.*)$`, "i");

    const items = [];
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i];
        if (DATE_RE.test(line)) continue;

        const m = line.match(ITEM_RE);
        if (!m) continue;

        const rawName = m[1].trim();
        const searchKey = rawName.replace(/\s+/g, "");
        let amount = parseFloat(m[2]);
        let inputUnit = m[3] ? m[3].trim() : null;
        let inputNote = m[4] ? m[4].trim() : "";

        let specKey = "";
        let subItemsData = []; // เก็บสินค้าแยกกรณีหน่วยรอง
        if (inputUnit && inputNote) {
            const extra = inputNote.match(EXTRA_QTY_RE);
            if (extra) {
                const hasPlus = !!extra[1];
                const amount2 = parseFloat(extra[2]);
                const unit2 = extra[3].trim();
                const note2 = extra[4].trim();

                const CONTAINERS = ["แพ็ค", "กล่อง", "ถุง", "ขวด", "ลัง", "แผง", "ห่อ", "ถาด", "ม้วน", "กระบอก", "เข่ง", "ซอง", "กระสอบ", "ฟอง"];
                const WEIGHTS = ["กก", "กรัม", "ขีด", "ลิตร", "มล"];

                let action = "SPLIT";
                // ถ้าไม่มี + และตัวหลังเป็น Container แต่ตัวแรกเป็น Weight -> ทำการสลับตำแหน่ง (SWAP) เป็น spec
                if (!hasPlus && CONTAINERS.includes(unit2) && WEIGHTS.includes(inputUnit)) {
                    action = "SWAP";
                }

                if (action === "SWAP") {
                    const specStr = `${amount} ${inputUnit}`;
                    specKey = `${amount}${inputUnit}`;
                    amount = amount2;
                    inputUnit = unit2;
                    inputNote = [specStr, note2].filter(Boolean).join(" ");
                } else {
                    // SPLIT: แยกเป็น 2 แถวและหาว่าตัวไหนเป็นตัวหลัก (น้ำหนักหรือปริมาณมากกว่า)
                    const UNIT_SCORE = { "กก": 300, "ลิตร": 300, "ขีด": 200, "กรัม": 100, "มล": 100 };
                    let score1 = UNIT_SCORE[inputUnit] || 0;
                    let score2 = UNIT_SCORE[unit2] || 0;

                    let mainAmount = amount, mainUnit = inputUnit;
                    let subAmount = amount2, subUnit = unit2;

                    // สลับให้ตัวที่มี Score มากกว่าเป็นตัวหลักขึ้นก่อน
                    if (score2 > score1) {
                        mainAmount = amount2; mainUnit = unit2;
                        subAmount = amount; subUnit = inputUnit;
                    }

                    amount = mainAmount;
                    inputUnit = mainUnit;
                    inputNote = note2; // กระจาย note ส่วนที่เหลือให้ทั้งคู่

                    subItemsData.push({ amount: subAmount, unit: subUnit, note: note2 });
                }
            }
        }

        const findResult =
            (specKey && fuzzyFindMapping(searchKey + specKey, mappingDict)) ||
            fuzzyFindMapping(searchKey, mappingDict);
        const mapped = findResult ? findResult.mapped : null;
        const isFuzzy = findResult ? findResult.isFuzzy : false;
        const isWarning = !mapped;
        const finalName = mapped ? mapped.name : rawName;
        const finalUnit = inputUnit || (mapped && mapped.unit) || "หน่วย";
        const note = `(${amount} ${finalUnit}${inputNote ? " " + inputNote : ""})`;

        items.push({ name: finalName, amount, unit: finalUnit, note, inputNote, isWarning, isFuzzy, isSubItem: false });

        for (const sub of subItemsData) {
            const subName = finalName;
            const sUnit = sub.unit || "หน่วย";
            const sNoteText = `(${sub.amount} ${sUnit}${sub.note ? " " + sub.note : ""})`;
            items.push({ name: subName, amount: sub.amount, unit: sUnit, note: sNoteText, inputNote: sub.note, isWarning, isFuzzy, isSubItem: true });
        }
    }
    // [FIX] Group items with the same name to handle sub-items typed on separate lines
    const groupedItems = {};
    for (const item of items) {
        if (!groupedItems[item.name]) groupedItems[item.name] = [];
        groupedItems[item.name].push(item);
    }

    const finalItems = [];
    const UNIT_SCORE = { "กก": 300, "ลิตร": 300, "ขีด": 200, "กรัม": 100, "มล": 100 };
    for (const name in groupedItems) {
        const group = groupedItems[name];
        if (group.length > 1) {
            // Sort group so heavier unit comes first
            group.sort((a, b) => {
                const scoreA = UNIT_SCORE[a.unit] || 0;
                const scoreB = UNIT_SCORE[b.unit] || 0;
                return scoreB - scoreA;
            });

            // Mark the first one as main, subsequent as subItem
            group[0].isSubItem = false;
            for (let i = 1; i < group.length; i++) {
                group[i].isSubItem = true;
            }
        }
        finalItems.push(...group);
    }

    return { storeName, deliveryDate: deliveryDateStr, items: finalItems };
}

// ==============================================================================
// 📅 DATE HELPERS
// ==============================================================================
function getTodayStr() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = String((d.getFullYear() + 543) % 100).padStart(2, "0");
    return `${dd}/${mm}/${yy}`;
}

function parseDateStr(str) {
    const parts = str.split("/");
    if (parts.length < 3) return null;
    let [dd, mm, yy] = parts.map(Number);
    if (yy < 100) yy = yy + 2500 - 543;
    return new Date(yy, mm - 1, dd);
}

function getMonthlyStockTabName(deliveryDateStr) {
    const parts = deliveryDateStr.split("/");
    if (parts.length < 2) return "ของในสต็อก";
    const monthNum = parseInt(parts[1], 10);
    const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    if (monthNum >= 1 && monthNum <= 12) {
        return "ของในสต็อก " + monthNames[monthNum - 1];
    }
    return "ของในสต็อก";
}

function isSameDay(d1, d2) {
    return (
        d1.getFullYear() === d2.getFullYear() &&
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate()
    );
}

// ==============================================================================
// 🔍 FUZZY MATCHING
// ==============================================================================
function levenshtein(a, b) {
    const m = a.length, n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;
    const dp = [];
    for (let i = 0; i <= m; i++) {
        dp[i] = [i];
        for (let j = 1; j <= n; j++) {
            dp[i][j] =
                i === 0
                    ? j
                    : a[i - 1] === b[j - 1]
                        ? dp[i - 1][j - 1]
                        : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
        }
    }
    return dp[m][n];
}

function isSubsequence(sub, str) {
    let si = 0;
    for (let i = 0; i < str.length && si < sub.length; i++) {
        if (str[i] === sub[si]) si++;
    }
    return si === sub.length;
}

function fuzzyFindMapping(searchKey, mappingDict) {
    // 1. Exact match
    if (mappingDict[searchKey]) {
        return { mapped: mappingDict[searchKey], isFuzzy: false };
    }
    if (searchKey.length < 3) return null;

    // 2. Substring match
    const subMatches = Object.keys(mappingDict).filter(
        (k) => k.includes(searchKey) || searchKey.includes(k),
    );
    if (subMatches.length === 1) {
        return { mapped: mappingDict[subMatches[0]], isFuzzy: true };
    }

    // 3. Subsequence match
    const seqMatches = Object.keys(mappingDict).filter((k) => {
        if (searchKey.length / k.length < CONFIG.SUBSEQ_MIN_RATIO) return false;
        return isSubsequence(searchKey, k);
    });
    if (seqMatches.length === 1) {
        return { mapped: mappingDict[seqMatches[0]], isFuzzy: true };
    }

    // 4. Levenshtein — [FIX-6] ตัด key ที่ length ต่างกันเกิน threshold ออกก่อน
    const threshold = Math.max(CONFIG.FUZZY_MIN_DIST, Math.floor(searchKey.length * CONFIG.FUZZY_RATIO));
    let bestKey = null, bestDist = Infinity;
    for (const key of Object.keys(mappingDict)) {
        // [FIX-6] skip key ที่ length ต่างกันมากกว่า threshold — ไม่มีทางชนะ Levenshtein
        if (Math.abs(key.length - searchKey.length) > threshold) continue;
        const dist = levenshtein(searchKey, key);
        if (dist < bestDist) {
            bestDist = dist;
            bestKey = key;
        }
    }
    if (bestKey && bestDist <= threshold) {
        return { mapped: mappingDict[bestKey], isFuzzy: true };
    }

    return null;
}

function normalizeDelivDate(rawDelivDate) {
    if (rawDelivDate instanceof Date) {
        const sheetsYear = rawDelivDate.getFullYear();
        let ceYear;
        if (sheetsYear >= 1930 && sheetsYear < 2000) {
            ceYear = (sheetsYear % 100) + 2500 - 543;
        } else {
            ceYear = sheetsYear;
        }
        const dd = String(rawDelivDate.getDate()).padStart(2, "0");
        const mm = String(rawDelivDate.getMonth() + 1).padStart(2, "0");
        const yy = String((ceYear + 543) % 100).padStart(2, "0");
        return `${dd}/${mm}/${yy}`;
    } else {
        const parts = String(rawDelivDate || "").trim().split(/[\/\-]/);
        if (parts.length === 3) {
            return (
                parts[0].trim().padStart(2, "0") +
                "/" +
                parts[1].trim().padStart(2, "0") +
                "/" +
                parts[2].trim().padStart(2, "0")
            );
        }
        return String(rawDelivDate || "").trim() || null;
    }
}

// ==============================================================================
// ✏️ onEdit — [FIX-8] guard e.source null
// ==============================================================================
function onEdit(e) {
    // [FIX-8] guard: e.source อาจ null ถ้า trigger ถูกเรียกโดยไม่มี context
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // ล้าง Cache ของ Mapping ทันทีที่มีการแก้ไขชีต Mapping เพื่อให้บอทรู้จักสินค้าใหม่/ที่แก้ไขทันที
    if (sheetName === "Mapping") {
        const cache = CacheService.getScriptCache();
        if (cache) cache.remove("mapping_dict");
    }

    if (sheetName.startsWith("ออเดอร์-")) {
        const editRow = e.range.getRow();
        if (editRow <= 1) return;
        const datePart = sheetName.replace("ออเดอร์-", "");
        const deliveryDateStr = datePart.replace(/-/g, "/");
        try {
            // [FIX-8] ตรวจ e.source ก่อนใช้
            const ss = e.source;
            if (!ss) return;
            SpreadsheetApp.flush();
            withScriptLock(() => updatePurchaseSummarySheet(ss, deliveryDateStr));
        } catch (err) {
            console.error("onEdit sync purchase summary error:", err.message);
        }
    }
}

// ==============================================================================
// 🗺️ MAPPING DICTIONARY — [FIX-3] CacheService
// ==============================================================================
function standardizeUnit(raw) {
    const u = raw.trim();
    if (!u) return "";
    const normalized = normalizeUnits("999 " + u);
    const m = normalized.match(/^999\s+(.+)/);
    return m ? m[1].trim().replace(/\.$/, "") : u;
}

function getMappingDictionary(ss) {
    // [FIX-3] ลอง hit cache ก่อน
    const cache = CacheService.getScriptCache();
    const cached = cache.get("mapping_dict");
    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (_) {
            // cache เสียหาย → rebuild
        }
    }

    const sheet = ss.getSheetByName("Mapping");
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    const mapDict = {};

    for (let i = 1; i < data.length; i++) {
        const rawAliases = String(data[i][0] || "").trim();
        const name = String(data[i][1] || "").trim();
        const rawUnit = String(data[i][2] || "").trim();
        const noStock =
            String(data[i][3] || "").trim().toLowerCase() === "x";

        if (!rawAliases || !name) continue;

        const firstUnit = rawUnit.split(",")[0].trim();
        const unit = standardizeUnit(firstUnit);
        const aliases = rawAliases.split(",").map((k) => k.trim()).filter(Boolean);

        for (const alias of aliases) {
            const key = alias.replace(/\s+/g, "");
            if (key) mapDict[key] = { name: name || alias, unit, noStock };
        }
    }

    // [FIX-3] เก็บ cache ไว้ CONFIG.MAPPING_CACHE_TTL วินาที
    try {
        cache.put("mapping_dict", JSON.stringify(mapDict), CONFIG.MAPPING_CACHE_TTL);
    } catch (_) {
        // กรณี mapDict ใหญ่เกิน 100KB limit ของ CacheService → ข้ามไป
    }

    return mapDict;
}

// ==============================================================================
// 📦 STOCK SYSTEM
// ==============================================================================
function processStockUpdate(rawMessage, ss, mappingDict) {
    const deliveryDate = parseDateFromText(rawMessage) || getTodayStr();
    let stockSheet, section;
    try {
        const res = ensureStockDateSection(deliveryDate);
        stockSheet = res.stockSheet;
        section = res.section;
    } catch (err) {
        console.error("External Stock Error:", err.message);
        return { count: 0, deliveryDate };
    }
    if (!stockSheet || !section) return { count: 0, deliveryDate };

    let text = normalizeUnits(rawMessage);
    const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
    const today = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    const UNIT_LIST =
        "กก|กรัม|ขีด|แพ็ค|มัด|กล่อง|ถุง|ขวด|ลัง|ลูก|ผล|แผง|ห่อ|ตัว|ก้อน|ถาด|ม้วน|อัน|ชิ้น|กระบอก|กระป๋อง|กำ|ช่อ|เข่ง|โหล|หัว|เม็ด|ชุด|คู่|ซอง|ลิตร|มล|ฝัก|กระสอบ|ฟอง";
    const ITEM_RE = new RegExp(
        `^(.+?)\\s+(\\d+(?:\\.\\d+)?)\\s*(${UNIT_LIST})?\\s*(.*)$`,
        "i",
    );

    const existingData = stockSheet.getDataRange().getValues();
    const stockMap = {};
    const stockQtyMap = {};
    for (let i = section.startIndex; i < section.endIndex; i++) {
        const k = String(existingData[i][1]).replace(/\s+/g, "").trim();
        if (k) {
            stockMap[k] = i + 1;
            stockQtyMap[k] = parseFloat(existingData[i][2]) || 0;
        }
    }

    let count = 0;
    for (let i = 1; i < lines.length; i++) {
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(lines[i])) continue;
        const m = lines[i].match(ITEM_RE);
        if (!m) continue;

        const rawName = m[1].trim();
        const searchKey = rawName.replace(/\s+/g, "");
        const amount = parseFloat(m[2]);
        const unit = m[3] ? m[3].trim() : null;
        const findResult = fuzzyFindMapping(searchKey, mappingDict);
        const mapped = findResult ? findResult.mapped : null;
        const finalName = mapped ? mapped.name : rawName;
        const finalUnit = unit || (mapped && mapped.unit) || "หน่วย";
        const lookupKey = finalName.replace(/\s+/g, "");

        if (stockMap[lookupKey]) {
            const row = stockMap[lookupKey];
            const currentQty = stockQtyMap[lookupKey] || 0;
            stockQtyMap[lookupKey] = currentQty + amount;
            stockSheet
                .getRange(row, 2, 1, 3)
                .setValues([[finalName, stockQtyMap[lookupKey], finalUnit]]);
        } else {
            const insertRow = section.endIndex + 1;
            if (insertRow > stockSheet.getLastRow()) {
                stockSheet.appendRow(["", finalName, amount, finalUnit]);
            } else {
                stockSheet.insertRowBefore(insertRow);
                stockSheet.getRange(insertRow, 2, 1, 3).setValues([[finalName, amount, finalUnit]]);
            }
            stockSheet.getRange(insertRow, 1).setValue(today);
            stockMap[lookupKey] = insertRow;
            stockQtyMap[lookupKey] = amount;
            section.endIndex++;
        }
        count++;
    }
    return { count, deliveryDate };
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
    const mappingData = getMappingDictionary(ss);

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
            summaryByDate[delivDateStr] = {};
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
        const stockMap = {};
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

    const noStockItems = Object.values(mappingData).filter((v) => v.noStock).map((v) => v.name);
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

// ==============================================================================
// 📋 PER-DATE SHEET — "ออเดอร์-DD-MM-YY"
// ==============================================================================
function updateDeliverySheet(ss, deliveryDateStr, storeName, items, orderTimestamp) {
    const sheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const NUM_COLS = 6;
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
            if (insertRow > currentLast) {
                sheet.appendRow([addedDateStr, storeName, item.name, item.amount, fmtUnit(item.unit), item.inputNote || ""]);
            } else {
                sheet.insertRowBefore(insertRow);
                sheet.getRange(insertRow, 1, 1, NUM_COLS).setValues([[addedDateStr, storeName, item.name, item.amount, fmtUnit(item.unit), item.inputNote || ""]]);
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
        const batchRows = items.map((item) => [deliveryDateStr, storeName, item.name, item.amount, fmtUnit(item.unit), item.inputNote || ""]);
        if (batchRows.length) {
            sheet.getRange(dataRow, 1, batchRows.length, NUM_COLS).setValues(batchRows);
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
}

// ==============================================================================
// ❌ CANCEL ORDER HELPERS
// ==============================================================================

/** ค้นหารายการสินค้าทั้งหมดของร้านในชีทออเดอร์ */
function findStoreOrders(orderSheet, storeName) {
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

// ==============================================================================
// 🛒 PURCHASE SUMMARY SHEET & STOCK DEDUCTION
// ==============================================================================

function getOrderRounds(orderSheet) {
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 3) return [];
    const allData = orderSheet.getRange(1, 1, lastRow, 6).getValues();

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
        if (!product) continue;

        // key = ชื่อ|หน่วย เพื่อแยกแถวเมื่อหน่วยต่างกัน (เช่น กก. vs ลูก)
        const productKey = product + "|" + unit;

        if (!currentRoundProducts[productKey]) {
            currentRoundProducts[productKey] = { total: amount, unit, name: product, specs: [] };
            if (spec) currentRoundProducts[productKey].specs.push(spec);
        } else {
            currentRoundProducts[productKey].total += amount;
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

function deductStockForCutoff(ss, deliveryDateStr) {
    const orderSheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);
    if (!orderSheet) return { success: false, message: `ไม่พบชีตออเดอร์วันที่ ${deliveryDateStr}` };

    const rounds = getOrderRounds(orderSheet);
    if (rounds.length === 0) return { success: false, message: "ไม่มีออเดอร์ให้ตัดรอบ" };

    const openRound = rounds[rounds.length - 1];
    if (openRound.isClosed) return { success: false, message: "ไม่มีออเดอร์ใหม่ตั้งแต่ตัดรอบครั้งล่าสุด" };
    // ถ้าไม่มีสินค้าเลยก็ไม่เป็นไร แค่ไม่หักสต็อก
    if (Object.keys(openRound.products).length === 0) return { success: true };

    try {
        const stockSS = SpreadsheetApp.openById(CONFIG.STOCK_FILE_ID);
        const stockSheet = stockSS.getSheetByName(getMonthlyStockTabName(deliveryDateStr));
        if (!stockSheet) return { success: false, message: "ไม่พบชีตสต็อกเดือนนี้" };
        const listData = stockSheet.getDataRange().getValues();
        const section = findStockDateSection(listData, deliveryDateStr);
        if (!stockSheet || !section) return { success: false, message: "ไม่พบชีตสต็อกเดือนนี้" };

        const normalizeUnit = (u) => String(u || "").replace(/\./g, "").trim();

        const stockUnitMap = {};
        for (let i = section.startIndex; i < section.endIndex; i++) {
            const prod = String(listData[i][1]).trim();
            const unit = String(listData[i][3]).trim();
            if (prod) stockUnitMap[prod] = unit;
        }

        const stockDeductMap = {};
        for (const key of Object.keys(openRound.products)) {
            const p = openRound.products[key];
            const name = p.name || key;
            if (stockUnitMap[name] && normalizeUnit(p.unit) === normalizeUnit(stockUnitMap[name])) {
                stockDeductMap[name] = (stockDeductMap[name] || 0) + p.total;
            }
        }

        const qtyRows = [];
        for (let i = section.startIndex; i < section.endIndex; i++) {
            const product = String(listData[i][1]).trim();
            const currentStock = parseFloat(listData[i][2]) || 0;
            const deductQty = product ? stockDeductMap[product] || 0 : 0;
            qtyRows.push([Math.max(0, currentStock - deductQty)]);
        }
        if (qtyRows.length) {
            stockSheet.getRange(section.startIndex + 1, 3, qtyRows.length, 1).setValues(qtyRows);
            SpreadsheetApp.flush();
        }

        return { success: true };
    } catch (err) {
        return { success: false, message: "Stock Error: " + err.message };
    }
}

function updatePurchaseSummarySheet(ss, deliveryDateStr, useCache = false) {
    deliveryDateStr = normalizeDelivDate(deliveryDateStr);
    const summarySheetName = "ใบซื้อ-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheetName = "ออเดอร์-" + deliveryDateStr.replace(/\//g, "-");
    const orderSheet = ss.getSheetByName(orderSheetName);

    const rounds = orderSheet ? getOrderRounds(orderSheet) : [];

    const summarySheet = ss.getSheetByName(summarySheetName);
    const history = getHistoricalStock(summarySheet);

    let externalStockMap = {};
    let monthlyStockMap = {};
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
                }
            }
        } catch (err) { console.error(err); }
    }

    let sheet = summarySheet;
    if (sheet) {
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
function createCutoffRoundDivider(ss, deliveryDateStr) {
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

    const closedRound = cutoffCount + 1;
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
