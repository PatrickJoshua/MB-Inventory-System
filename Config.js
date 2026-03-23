/**
 * Environment Configuration for Store Inventory Management System
 * 
 * This file contains all environment-specific configuration values.
 * Use getConfig(env) to retrieve configuration for the specified environment.
 * 
 * Supported environments: 'PRD' (production), 'DEV' (development)
 */

const ENV = {
    PRD: {
        // ===== Inventory Spreadsheets =====
        INVENTORY_URL_RF: 'https://docs.google.com/spreadsheets/d/1XxOw-t7q60ULv59GzABMn4uEouFZlZrYtMOMsoLtAtA/edit',
        INVENTORY_URL_PCGH: 'https://docs.google.com/spreadsheets/d/1rwW-JantrQTuEg2Uzez6XR0xcvceOFbMVuke1t2Idak/edit',

        // ===== Archive Spreadsheets =====
        ARCHIVE_URL_RF: 'https://docs.google.com/spreadsheets/d/1-31Kf3SsdMhNu1D9ziwBvqlhI6LUTlMgO3PD-vUTNkM/edit',
        ARCHIVE_URL_PCGH: 'https://docs.google.com/spreadsheets/d/1QasIoOag67V1I_ePVT_-LaENPTA0Ah1cPIu7uo-q7XE/edit',

        // ===== PO & Cash Management =====
        PO_URL: 'https://docs.google.com/spreadsheets/d/10IGAlAy_4LqyFgi3UNAVHBDy9hp69yJOd6oQVhL4JLk/edit',
<<<<<<< Updated upstream
        ARCHIVE_PO_URL: 'https://docs.google.com/spreadsheets/d/1XcJc4mK3EQ7Ag7wV9K25fqTkX-uWhlJrpRxkLyzMJEY/edit',
=======
        ARCHIVE_PO_URL: 'https://docs.google.com/spreadsheets/d/18yE2KgDAx3sliD9Wihq2pwW6EPm5d-Q0MXObOSm0om4/edit',
>>>>>>> Stashed changes
        SMS_API_URL: 'https://docs.google.com/spreadsheets/d/17yPemlid9FVMdzVDX8Eg8Tu1W-zOg_prNtQeUeEidAg/edit',
        DUMP_SITE_URL: 'https://docs.google.com/spreadsheets/d/1rPiSSJlbfLDKjqofRxqh1DR1lP5VYS67SFBp99to5Iw/edit',

        // ===== Contacts =====
        ALERT_EMAIL: 'bakulinglings@gmail.com',
        ADMIN_EMAIL: 'mbs2edith@gmail.com',
        ALERT_SMS: '+639151272800'
    },

    DEV: {
        // ===== Inventory Spreadsheets =====
        // TODO: Replace with DEV spreadsheet URLs
        INVENTORY_URL_RF: 'https://docs.google.com/spreadsheets/d/1Q2hoJFrAvF3ts20ZCj_d1iiEEopVRwPOggJrb/edit',
        INVENTORY_URL_PCGH: 'https://docs.google.com/spreadsheets/d/DEV_INVENTORY_PCGH_PLACEHOLDER/edit',

        // ===== Archive Spreadsheets =====
        // TODO: Replace with DEV spreadsheet URLs
        ARCHIVE_URL_RF: 'https://docs.google.com/spreadsheets/d/1GYGlRNyca5v8rqGKIDV-U7dTAuRk5AKwwC9_-Xqk8ts/edit',
        ARCHIVE_URL_PCGH: 'https://docs.google.com/spreadsheets/d/DEV_ARCHIVE_PCGH_PLACEHOLDER/edit',

        // ===== PO & Cash Management =====
        // TODO: Replace with DEV spreadsheet URLs
        PO_URL: 'https://docs.google.com/spreadsheets/d/14qqh7FU6QABvDn2XZzmKE1E1BW7V2NRokczVwxTnAmA/edit',
<<<<<<< Updated upstream
        ARCHIVE_PO_URL: 'https://docs.google.com/spreadsheets/d/18yE2KgDAx3sliD9Wihq2pwW6EPm5d-Q0MXObOSm0om4/edit',
=======
        ARCHIVE_PO_URL: 'https://docs.google.com/spreadsheets/d/1XcJc4mK3EQ7Ag7wV9K25fqTkX-uWhlJrpRxkLyzMJEY/edit',
>>>>>>> Stashed changes
        SMS_API_URL: 'https://docs.google.com/spreadsheets/d/DEV_SMS_API_PLACEHOLDER/edit',
        DUMP_SITE_URL: 'https://docs.google.com/spreadsheets/d/DEV_DUMP_SITE_PLACEHOLDER/edit',

        // ===== Contacts (same as PRD for now) =====
        ALERT_EMAIL: 'bakulinglings@gmail.com',
        ADMIN_EMAIL: 'bakulinglings@gmail.com',
        ALERT_SMS: '+639151272800'
    }
};

// ===== Store Code Mappings =====
const STORE_CODES = {
    RF: '3252',
    PCGH: '3361'
};

/**
 * Returns configuration object for the specified environment.
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {Object} Configuration object for the environment.
 */
function getConfig(env = 'PRD') {
    if (!ENV[env]) {
        console.warn(`[CONFIG] Unknown environment '${env}', falling back to PRD`);
        return ENV.PRD;
    }
    return ENV[env];
}

/**
 * Returns the inventory spreadsheet URL for a given store code.
 * @param {string} storeCode - Store code ('3252' for RF, '3361' for PCGH)
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} Inventory spreadsheet URL
 */
function getInventoryUrlByConfig(storeCode, env = 'PRD') {
    const config = getConfig(env);
    if (storeCode == '3252' || storeCode == STORE_CODES.RF) {
        return config.INVENTORY_URL_RF;
    } else if (storeCode == '3361' || storeCode == STORE_CODES.PCGH) {
        return config.INVENTORY_URL_PCGH;
    }
    throw new Error(`[CONFIG] Unknown store code: ${storeCode}`);
}

/**
 * Returns the archive spreadsheet URL for a given store code.
 * @param {string} storeCode - Store code ('3252' for RF, '3361' for PCGH)
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} Archive spreadsheet URL
 */
function getArchiveUrlByConfig(storeCode, env = 'PRD') {
    const config = getConfig(env);
    if (storeCode == '3252' || storeCode == STORE_CODES.RF) {
        return config.ARCHIVE_URL_RF;
    } else if (storeCode == '3361' || storeCode == STORE_CODES.PCGH) {
        return config.ARCHIVE_URL_PCGH;
    }
    throw new Error(`[CONFIG] Unknown store code: ${storeCode}`);
}

/**
 * Returns the PO spreadsheet URL.
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} PO spreadsheet URL
 */
function getPoUrlByConfig(env = 'PRD') {
    return getConfig(env).PO_URL;
}

/**
 * Returns the Archive PO spreadsheet URL.
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} Archive PO spreadsheet URL
 */
function getArchivePoUrlByConfig(env = 'PRD') {
    return getConfig(env).ARCHIVE_PO_URL;
}

/**
 * Returns the SMS API spreadsheet URL.
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} SMS API spreadsheet URL
 */
function getSmsApiUrlByConfig(env = 'PRD') {
    return getConfig(env).SMS_API_URL;
}

/**
 * Returns the dump site spreadsheet URL.
 * @param {string} env - Environment name ('PRD' or 'DEV'). Defaults to 'PRD'.
 * @returns {string} Dump site spreadsheet URL
 */
function getDumpSiteUrlByConfig(env = 'PRD') {
    return getConfig(env).DUMP_SITE_URL;
}
