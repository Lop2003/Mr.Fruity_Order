// ==============================================================================
// 🍎 Order System — Fix-04/07/69
// ==============================================================================

const CONFIG = {
    SHEET_ID: "1IUpkB2Cs2cjXVoBm_d9OYiYxhxzz9ReWCftNOlKphVk",
    STOCK_FILE_ID: "1I0_DZs4jCYSarG4haw1aB01mQYJ1sFany9qlbyo0tbc",
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
    EVENT_STALE_MS: 5 * 60 * 1000,
};

let LAST_LINE_REPLY_ERROR = "";
let ACTIVE_LOG_SHEET = null;
let ACTIVE_LOG_ROW = 0;
let ACTIVE_EVENT_ID = "";
