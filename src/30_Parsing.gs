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

function parseOrderMessage(rawMessage, mappingDict, storeMappingDict) {
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

    const storeKey = normalizeStoreMappingKey(storeName);
    const storeTable = storeMappingDict && Object.prototype.hasOwnProperty.call(storeMappingDict, storeKey)
        ? storeMappingDict[storeKey]
        : null;
    const storeMappingSummary = {
        configured: !!storeTable,
        matched: 0,
        missed: 0,
        invalid: 0,
        errors: [],
    };

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
        if (!Number.isFinite(amount) || amount <= 0) continue;
        let inputUnit = m[3] ? m[3].trim() : null;
        let inputNote = m[4] ? m[4].trim() : "";

        let specKey = "";
        let subItemsData = []; // เก็บสินค้าแยกกรณีหน่วยรอง
        if (inputUnit && inputNote) {
            const extra = inputNote.match(EXTRA_QTY_RE);
            if (extra) {
                const hasPlus = !!extra[1];
                const amount2 = parseFloat(extra[2]);
                if (!Number.isFinite(amount2) || amount2 <= 0) continue;
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

        let mapped = null;
        let isFuzzy = false;
        let isWarning = false;
        let mappingSource = "UNMAPPED";
        let mappingError = "";
        const poKey = normalizeStoreMappingKey(rawName);
        const storeProducts = storeTable && storeTable.products;
        const hasStoreProduct = !!storeProducts && Object.prototype.hasOwnProperty.call(storeProducts, poKey);

        if (hasStoreProduct) {
            const storeProduct = storeProducts[poKey];
            let isValidStoreProduct = storeProduct.valid;
            let storeUnit = storeProduct.units.length === 1 ? storeProduct.units[0] : "";
            mappingError = storeProduct.reason || "";

            if (isValidStoreProduct && storeProduct.units.length > 1) {
                const normalizedInputUnit = inputUnit ? standardizeUnit(inputUnit) : "";
                if (normalizedInputUnit && storeProduct.units.includes(normalizedInputUnit)) {
                    storeUnit = normalizedInputUnit;
                } else {
                    isValidStoreProduct = false;
                    mappingError = `ชื่อมาตรฐานมีหลายหน่วย (${storeProduct.units.join(", ")})`;
                }
            }

            if (isValidStoreProduct) {
                mapped = { name: storeProduct.standardName, unit: storeUnit };
                mappingSource = "STORE_MAPPING";
                storeMappingSummary.matched++;
            } else {
                isWarning = true;
                mappingSource = "STORE_MAPPING_INVALID";
                storeMappingSummary.invalid++;
                storeMappingSummary.errors.push({ storeName, poName: rawName, reason: mappingError });
            }
        } else {
            if (storeTable) storeMappingSummary.missed++;
            const findResult =
                (specKey && fuzzyFindMapping(searchKey + specKey, mappingDict)) ||
                fuzzyFindMapping(searchKey, mappingDict);
            mapped = findResult ? findResult.mapped : null;
            isFuzzy = findResult ? findResult.isFuzzy : false;
            isWarning = !mapped;
            mappingSource = mapped ? (isFuzzy ? "MAPPING_FUZZY" : "MAPPING_EXACT") : "UNMAPPED";
        }

        const finalName = mapped ? mapped.name : rawName;
        const finalUnit = inputUnit || (mapped && mapped.unit) || "หน่วย";
        const note = `(${amount} ${finalUnit}${inputNote ? " " + inputNote : ""})`;

        items.push({ name: finalName, amount, unit: finalUnit, note, inputNote, isWarning, isFuzzy, isSubItem: false, mappingSource, mappingError });

        for (const sub of subItemsData) {
            const subName = finalName;
            const sUnit = sub.unit || "หน่วย";
            const sNoteText = `(${sub.amount} ${sUnit}${sub.note ? " " + sub.note : ""})`;
            items.push({ name: subName, amount: sub.amount, unit: sUnit, note: sNoteText, inputNote: sub.note, isWarning, isFuzzy, isSubItem: true, mappingSource, mappingError });
        }
    }
    // [FIX] Group items with the same name to handle sub-items typed on separate lines
    const groupedItems = Object.create(null);
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

    if (!storeName || !deliveryDateStr || finalItems.length === 0) return null;
    return { storeName, deliveryDate: deliveryDateStr, items: finalItems, storeMappingSummary };
}

function getOrderMatchNotes(items) {
    const hasFuzzy = items.some((item) => item.isFuzzy);
    const hasStoreMappingError = items.some((item) => item.mappingSource === "STORE_MAPPING_INVALID");
    const hasMappingWarning = items.some(
        (item) => item.isWarning && item.mappingSource !== "STORE_MAPPING_INVALID",
    );
    return [
        hasFuzzy ? "\n\n🔍 บางรายการใช้การจับคู่อัตโนมัติ กรุณาตรวจสอบความถูกต้อง" : "",
        hasMappingWarning ? "\n\n⚠️ บางรายการไม่พบใน Mapping กรุณาตรวจสอบ" : "",
        hasStoreMappingError ? "\n\n⚠️ บางรายการตั้งค่า StoreMapping ไม่ถูกต้อง กรุณาตรวจสอบ" : "",
    ].join("");
}

function formatStoreMappingLog(summary) {
    if (!summary || !summary.configured) return "";
    let detail = `StoreMapping: Matched ${summary.matched || 0}; Missed ${summary.missed || 0}; Invalid ${summary.invalid || 0}`;
    if (summary.errors && summary.errors.length) {
        const uniqueErrors = [...new Set(summary.errors
            .map((item) => `${item.storeName} / ${item.poName}: ${item.reason}`))];
        const shownErrors = uniqueErrors.slice(0, 20);
        detail += "; Errors " + shownErrors.join(" | ");
        if (uniqueErrors.length > shownErrors.length) {
            detail += ` | +${uniqueErrors.length - shownErrors.length} more`;
        }
    }
    return detail;
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
    if (Object.prototype.hasOwnProperty.call(mappingDict, searchKey)) {
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
    let bestKey = null, bestDist = Infinity, hasBestTie = false;
    for (const key of Object.keys(mappingDict)) {
        // [FIX-6] skip key ที่ length ต่างกันมากกว่า threshold — ไม่มีทางชนะ Levenshtein
        if (Math.abs(key.length - searchKey.length) > threshold) continue;
        const dist = levenshtein(searchKey, key);
        if (dist < bestDist) {
            bestDist = dist;
            bestKey = key;
            hasBestTie = false;
        } else if (dist === bestDist) {
            hasBestTie = true;
        }
    }
    if (bestKey && !hasBestTie && bestDist <= threshold) {
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
    }

    const match = String(rawDelivDate || "").trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2}|\d{4})$/);
    if (!match) return null;

    const day = Number(match[1]);
    const month = Number(match[2]);
    const rawYear = Number(match[3]);
    const ceYear = rawYear < 100 ? rawYear + 2500 - 543 : rawYear >= 2400 ? rawYear - 543 : rawYear;
    const date = new Date(ceYear, month - 1, day);
    if (date.getFullYear() !== ceYear || date.getMonth() !== month - 1 || date.getDate() !== day) return null;

    const buddhistYear = ceYear + 543;
    return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${String(buddhistYear % 100).padStart(2, "0")}`;
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
        if (cache) {
            cache.remove("mapping_dict");
            cache.remove("store_mapping_dict");
        }
    } else if (sheetName === "StoreMapping") {
        const cache = CacheService.getScriptCache();
        if (cache) cache.remove("store_mapping_dict");
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
            withScriptLock(() => updatePurchaseSummarySheet(ss, deliveryDateStr, false, true));
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

function normalizeStoreMappingKey(value) {
    return String(value || "").trim().replace(/\s+/g, "").toLowerCase();
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
    const mapDict = Object.create(null);

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

function getStoreMappingDictionary(ss) {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("store_mapping_dict");
    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (_) {
            // cache เสียหาย → rebuild
        }
    }

    const storeSheet = ss.getSheetByName("StoreMapping");
    if (!storeSheet) {
        try {
            cache.put("store_mapping_dict", "{}", CONFIG.MAPPING_CACHE_TTL);
        } catch (_) {
            // cache ใช้ไม่ได้ → ทำงานเดิมต่อ
        }
        return {};
    }

    const canonicalByName = Object.create(null);
    const mappingSheet = ss.getSheetByName("Mapping");
    if (mappingSheet) {
        const mappingRows = mappingSheet.getDataRange().getValues();
        for (let i = 1; i < mappingRows.length; i++) {
            const name = String(mappingRows[i][1] || "").trim();
            if (!name) continue;
            const primaryUnit = standardizeUnit(String(mappingRows[i][2] || "").split(",")[0].trim());
            if (!canonicalByName[name]) canonicalByName[name] = { name, units: [] };
            if (primaryUnit && !canonicalByName[name].units.includes(primaryUnit)) {
                canonicalByName[name].units.push(primaryUnit);
            }
        }
    }

    const storeRows = storeSheet.getDataRange().getValues();
    const storeDict = Object.create(null);
    for (let i = 1; i < storeRows.length; i++) {
        const storeName = String(storeRows[i][0] || "").trim();
        const poName = String(storeRows[i][1] || "").trim();
        const standardName = String(storeRows[i][2] || "").trim();
        const storeKey = normalizeStoreMappingKey(storeName);
        const poKey = normalizeStoreMappingKey(poName);
        if (!storeKey || !poKey) continue;

        if (!Object.prototype.hasOwnProperty.call(storeDict, storeKey)) {
            storeDict[storeKey] = { products: Object.create(null) };
        }
        const products = storeDict[storeKey].products;
        const existing = products[poKey];
        if (existing && existing.standardName !== standardName) {
            products[poKey] = {
                valid: false,
                standardName: "",
                units: [],
                reason: "StoreMapping ซ้ำและชี้ไปคนละชื่อมาตรฐาน",
            };
            continue;
        }
        if (existing) continue;

        const canonical = canonicalByName[standardName];
        products[poKey] = canonical
            ? { valid: true, standardName, units: canonical.units, reason: "" }
            : {
                valid: false,
                standardName,
                units: [],
                reason: standardName
                    ? "ชื่อมาตรฐานไม่มีใน Mapping คอลัมน์ B"
                    : "ชื่อมาตรฐานใน StoreMapping ว่าง",
            };
    }

    try {
        cache.put("store_mapping_dict", JSON.stringify(storeDict), CONFIG.MAPPING_CACHE_TTL);
    } catch (_) {
        // cache ใหญ่เกิน limit → ใช้ค่าที่อ่านได้ในรอบนี้
    }
    return storeDict;
}

function getNoStockProductNames(ss) {
    const sheet = ss.getSheetByName("Mapping");
    if (!sheet) return [];
    const names = new Set();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
        const name = String(rows[i][1] || "").trim();
        const noStock = String(rows[i][3] || "").trim().toLowerCase() === "x";
        if (name && noStock) names.add(name);
    }
    return [...names];
}
