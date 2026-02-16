import * as XLSX from 'xlsx';

/**
 * Format address from 6-digit code to full address string.
 * @param {string|number} code - The 6-digit code (e.g., 530902)
 * @returns {string} - Formatted address (e.g., "新北市新店區中央路153號9樓之2")
 */
export const formatAddress = (code) => {
    const codeStr = String(code).padStart(6, '0'); // Ensure 6 digits

    const ab = codeStr.substring(0, 2);
    const c = codeStr.substring(2, 3);
    const d = codeStr.substring(3, 4);
    const e = codeStr.substring(4, 5);
    const f = codeStr.substring(5, 6);

    // Address Prefix
    const road = "新北市新店區中央路";
    const number = `1${ab}`; // Prefix '1' to 'ab'

    // Floor Logic
    // If c is '0', omit it. Only showing d.
    // Actually, standard logic usually implies 09 -> 9. 
    // Requirement: "如果'c'或'e'為'0'，則直接不顯示"
    // So if cd is '09', it becomes '9樓'. If '12', '12樓'.
    let floor = "";
    if (c === '0') {
        floor = d;
    } else {
        floor = `${c}${d}`;
    }

    // Unit Logic (之ef)
    // If e is '0', omit it.
    // If ef is '00', omit the whole part.
    let unit = "";
    if (e === '0' && f === '0') {
        unit = ""; // Rule 2: "00" recovers to standard address without unit
    } else {
        if (e === '0') {
            unit = `之${f}`;
        } else {
            unit = `之${e}${f}`;
        }
    }

    let fullAddress = `${road}${number}號${floor}樓${unit}`;

    // Special case adjustment based on requirement rule 2: 
    // "如果'ef'為'00'，恢復地址格式改為：'新北市新店區中央路1ab號cd樓'" 
    // (matches logic above where unit is empty)

    return fullAddress;
};

/**
 * Format date range from "yyyy/mm/dd~yyyy/mm/dd" to ROC year month.
 * @param {string} dateRange - e.g. "2026/01/31~2026/01/31"
 * @returns {string} - e.g. "115年1月"
 */
export const formatDateRange = (dateRange) => {
    if (!dateRange) return "";

    const convertParams = (dateStr) => {
        // dateStr expected: yyyy/mm/dd
        const parts = dateStr.split('/');
        if (parts.length < 2) return { year: 0, month: 0 };

        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        return { rocYear: year - 1911, month };
    };

    const [startStr, endStr] = dateRange.split('~');

    if (!startStr) return dateRange; // Fallback

    const start = convertParams(startStr.trim());

    let result = `${start.rocYear}年${start.month}月`;

    if (endStr) {
        const end = convertParams(endStr.trim());
        // Required: "如果前後月份相同，則僅顯示'115年1月'"
        if (start.rocYear === end.rocYear && start.month === end.month) {
            return result;
        }
        result += `至${end.rocYear}年${end.month}月`;
    }

    return result;
};

/**
 * Calculate Arrears Item based on amount.
 * @param {number} amount 
 * @returns {string} "車位清潔費" or "房屋管理費"
 */
export const calculateArrearsItem = (amount) => {
    // Ensure amount is a number
    const val = Number(amount);
    if (isNaN(val)) return "房屋管理費"; // Default fallback

    if (val % 500 === 0 && val !== 0) {
        return "車位清潔費";
    }
    return "房屋管理費";
};

/**
 * Format phone number.
 * Rule 1: Starts with "09" -> "09xx-xxx-xxx"
 * Rule 2: Others -> "xx-xxxx-xxxx"
 * @param {string|number} phone 
 * @returns {string} Formatted phone or original if invalid
 */
export const formatPhoneNumber = (phone) => {
    if (!phone) return "";

    // Convert to string and strip non-digits
    let str = String(phone).replace(/\D/g, '');

    // Rule: If starts with "9", prepend "0"
    if (str.startsWith("9")) {
        str = "0" + str;
    }

    // Check if it looks like a standard 10-digit number
    if (str.startsWith("09")) {
        if (str.length >= 10) {
            return `${str.substring(0, 4)}-${str.substring(4, 7)}-${str.substring(7)}`;
        }
    } else {
        // "xx-xxxx-xxxx"
        if (str.length >= 10) {
            return `${str.substring(0, 2)}-${str.substring(2, 6)}-${str.substring(6)}`;
        }
    }

    return str; // Return cleaned number if it doesn't match rules or length
};

/**
 * Main processing function
 */
export const processExcelFiles = async (arrearsFile, residentsFile) => {
    // Helper to read file
    const readFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve(workbook);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    };

    try {
        const [wbArrears, wbResidents] = await Promise.all([
            readFile(arrearsFile),
            readFile(residentsFile)
        ]);

        // 1. Process Arrears File (File 1)
        // Assume data is in first sheet
        const sheetArrears = wbArrears.Sheets[wbArrears.SheetNames[0]];
        const dataArrears = XLSX.utils.sheet_to_json(sheetArrears, { header: "A", defval: "" }); // Use A, B, C... notation for column access

        // Skip header if necessary. Usually row 1 is header. 
        // We'll assume row 1 is header and actual data starts from row 2.
        // However, sheet_to_json with header:"A" treats row 1 as data with key 'A'...
        // Let's refine: Use default header first to see structure? 
        // No, requirements specify "C欄", "H欄", "K欄". So treating as A-based columns is safer provided we skip the top rows if they are titles.
        // Let's assume row 1 is title/header. We process from row 2.

        // 2. Process Residents File (File 2)
        // Required sheet: "新店機廠捷17.18.19"
        const targetSheetName = "新店機廠捷17.18.19";
        if (!wbResidents.SheetNames.includes(targetSheetName)) {
            throw new Error(`找不到工作表: "${targetSheetName}"`);
        }
        const sheetResidents = wbResidents.Sheets[targetSheetName];
        const dataResidents = XLSX.utils.sheet_to_json(sheetResidents, { header: "A", defval: "" });

        // Build a lookup map from Residents file
        // Key: Address (Col C) -> Value: { Name: Col H, Phone: Col I }
        const residentMap = new Map();
        dataResidents.slice(1).forEach(row => {
            const address = row['C']?.toString().trim();
            if (address) {
                residentMap.set(address, {
                    name: row['H'],
                    phone: row['I']
                });
            }
        });

        // 3. Merge and Create Output Data
        const outputData = [];

        // Define headers for output
        outputData.push(["地址", "住戶姓名", "費用欠繳期間", "欠繳名目", "欠費總金額", "連絡方式"]);

        dataArrears.slice(1).forEach(row => { // Adjust slice based on actual header content
            // Validate if this is a valid data row (e.g. check if Col C exists)
            const rawAddressCode = row['C'];
            if (!rawAddressCode) return; // Skip empty rows

            // Apply formatting
            const formattedAddress = formatAddress(rawAddressCode);
            const formattedDate = formatDateRange(row['H']);
            const rawAmount = row['K'];
            // Remove commas if any for calculation
            const amountNum = typeof rawAmount === 'string' ? parseFloat(rawAmount.replace(/,/g, '')) : rawAmount;

            const arrearsItem = calculateArrearsItem(amountNum);

            // Format Amount for display (with thousands separator)
            const displayAmount = amountNum.toLocaleString('en-US');

            // Match with Resident Map
            const residentInfo = residentMap.get(formattedAddress) || { name: "", phone: "" };

            const formattedPhone = formatPhoneNumber(residentInfo.phone);

            outputData.push([
                formattedAddress,
                residentInfo.name,
                formattedDate,
                arrearsItem,
                displayAmount,
                formattedPhone
            ]);
        });

        // 4. Generate New Excel
        const newWs = XLSX.utils.aoa_to_sheet(outputData);
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "處理結果");

        // Write file
        const excelBuffer = XLSX.write(newWb, { bookType: 'xlsx', type: 'array' });
        return new Blob([excelBuffer], { type: 'application/octet-stream' });

    } catch (error) {
        console.error("Error processing files:", error);
        throw error;
    }
};
