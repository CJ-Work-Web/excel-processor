import XLSX from 'xlsx-js-style';

// ============================================================
// Formatting Helpers
// ============================================================

export const formatAddress = (code) => {
    const codeStr = String(code).padStart(6, '0');
    const ab = codeStr.substring(0, 2);
    const c = codeStr.substring(2, 3);
    const d = codeStr.substring(3, 4);
    const e = codeStr.substring(4, 5);
    const f = codeStr.substring(5, 6);

    const road = "新北市新店區中央路";
    const number = `1${ab}`;

    let floor = c === '0' ? d : `${c}${d}`;

    let unit = "";
    if (e === '0' && f === '0') {
        unit = "";
    } else if (e === '0') {
        unit = `之${f}`;
    } else {
        unit = `之${e}${f}`;
    }

    return `${road}${number}號${floor}樓${unit}`;
};

export const formatDateRange = (dateRange) => {
    if (!dateRange) return "";

    const convertParams = (dateStr) => {
        const parts = dateStr.split('/');
        if (parts.length < 2) return { rocYear: 0, month: 0 };
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        return { rocYear: year - 1911, month };
    };

    const [startStr, endStr] = dateRange.split('~');
    if (!startStr) return dateRange;

    const start = convertParams(startStr.trim());
    const startMonth = String(start.month).padStart(2, '0');
    let result = `${start.rocYear}年${startMonth}月`;

    if (endStr) {
        const end = convertParams(endStr.trim());
        if (start.rocYear === end.rocYear && start.month === end.month) {
            return result;
        }
        const endMonth = String(end.month).padStart(2, '0');
        result = `${start.rocYear}年${startMonth}月至${end.rocYear}年${endMonth}月`;
    }

    return result;
};

export const calculateArrearsItem = (amount) => {
    const val = Number(amount);
    if (isNaN(val)) return "房屋管理費";
    if (val % 500 === 0 && val !== 0) return "車位清潔費";
    return "房屋管理費";
};

export const formatPhoneNumber = (phone) => {
    if (!phone) return "";
    let str = String(phone).replace(/\D/g, '');
    if (str.startsWith("9")) str = "0" + str;
    if (str.startsWith("09") && str.length >= 10) {
        return `${str.substring(0, 4)}-${str.substring(4, 7)}-${str.substring(7)}`;
    } else if (str.length >= 10) {
        return `${str.substring(0, 2)}-${str.substring(2, 6)}-${str.substring(6)}`;
    }
    return str;
};

// ============================================================
// Style Definitions
// ============================================================

const thin = { style: 'thin' };
const medium = { style: 'medium' };

const TITLE_FONT = { name: '微軟正黑體', sz: 16, bold: true };
const BASE_FONT = { name: '微軟正黑體', sz: 14 };
const BOLD_FONT = { name: '微軟正黑體', sz: 14, bold: true };

const CENTER = { horizontal: 'center', vertical: 'center' };
const CENTER_WRAP = { horizontal: 'center', vertical: 'center', wrapText: true };
const LEFT_CENTER = { vertical: 'center' };
const LEFT_CENTER_WRAP = { vertical: 'center', wrapText: true };
const RIGHT_CENTER = { horizontal: 'right', vertical: 'center' };
const RIGHT_CENTER_WRAP = { horizontal: 'right', vertical: 'center', wrapText: true };
const LEFT_ALIGN_CENTER = { horizontal: 'left', vertical: 'center' };
const LEFT_ALIGN_CENTER_WRAP = { horizontal: 'left', vertical: 'center', wrapText: true };

// Fill colors
// Header / stats: Accent 2 (#C0504D) with tint 0.4 → #D99694
const HEADER_FILL = { fgColor: { rgb: 'FFD99694' }, patternType: 'solid' };

// Border presets for header row
const HDR_LEFT = { left: medium, right: thin, top: medium, bottom: thin };
const HDR_MID = { left: thin, right: thin, top: medium, bottom: thin };
const HDR_RIGHT = { left: thin, right: medium, top: medium, bottom: thin };

// Border presets for data rows
const DATA_LEFT = { left: medium, right: thin, top: thin, bottom: thin };
const DATA_MID = { left: thin, right: thin, top: thin, bottom: thin };
const DATA_RIGHT = { left: thin, right: medium, top: thin, bottom: thin };

// ============================================================
// Cell creation helpers
// ============================================================

const ec = (r, c) => XLSX.utils.encode_cell({ r, c });

const mkCell = (value, type, style) => ({ v: value, t: type, s: style });

const mkFormula = (formula, style, numFmt) => {
    const cell = { f: formula, s: style };
    if (numFmt) cell.s = { ...style, numFmt };
    return cell;
};

// ============================================================
// Main Processing
// ============================================================

export const processExcelFiles = async (arrearsFile, residentsFile, rocYear, month) => {
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

        // --- Parse arrears data ---
        const sheetArrears = wbArrears.Sheets[wbArrears.SheetNames[0]];
        const dataArrears = XLSX.utils.sheet_to_json(sheetArrears, { header: "A", defval: "" });

        // --- Parse residents data ---
        const targetSheetName = "新店機廠捷17.18.19";
        if (!wbResidents.SheetNames.includes(targetSheetName)) {
            throw new Error(`找不到工作表: "${targetSheetName}"`);
        }
        const sheetResidents = wbResidents.Sheets[targetSheetName];
        const dataResidents = XLSX.utils.sheet_to_json(sheetResidents, { header: "A", defval: "" });

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

        // --- Build processed rows ---
        const rows = [];
        const SKIP_B_VALUES = new Set(['小計', '總計']);

        dataArrears.slice(1).forEach(row => {
            // Rule 1: skip 小計 / 總計 rows in column B
            const colB = row['B']?.toString().trim();
            if (SKIP_B_VALUES.has(colB)) return;

            const rawAddressCode = row['C'];
            if (!rawAddressCode) return;

            const rawAmount = row['K'];
            const amountNum = typeof rawAmount === 'string'
                ? parseFloat(rawAmount.replace(/,/g, ''))
                : Number(rawAmount);
            if (isNaN(amountNum)) return;

            const arrearsItem = calculateArrearsItem(amountNum);
            const formattedDate = formatDateRange(row['H']);
            const formattedAddress = formatAddress(rawAddressCode);
            const residentInfo = residentMap.get(formattedAddress);

            if (arrearsItem === '車位清潔費') {
                // Rule 3: 車位清潔費 → include row, leave address blank, keep name & phone
                rows.push({
                    address: '',
                    name: residentInfo ? (residentInfo.name || '') : '',
                    phone: residentInfo ? formatPhoneNumber(residentInfo.phone) : '',
                    period: formattedDate,
                    itemType: arrearsItem,
                    amount: amountNum
                });
            } else {
                // Rule 2: 房屋管理費 → skip if address not in resident map
                if (!residentInfo) return;

                rows.push({
                    address: formattedAddress,
                    name: residentInfo.name || '',
                    phone: formatPhoneNumber(residentInfo.phone),
                    period: formattedDate,
                    itemType: arrearsItem,
                    amount: amountNum
                });
            }
        });

        // --- Layout constants ---
        const BUFFER_ROWS = 3;
        const DATA_START = 4;                              // Excel row 4 (0-indexed: 3)
        const lastDataR = DATA_START + rows.length - 1;    // last data Excel row
        const formulaEndR = lastDataR + BUFFER_ROWS;       // formula range end row
        const statsR1 = formulaEndR + 1;                   // stats row 1
        const statsR2 = formulaEndR + 2;                   // stats row 2
        const sumR1 = formulaEndR + 3;                     // summary row 1
        const sumR2 = formulaEndR + 4;                     // summary row 2
        const sumR3 = formulaEndR + 5;                     // summary row 3
        const totalRows = sumR3;

        const ceYear = rocYear + 1911;
        const lastDay = new Date(ceYear, month, 0).getDate();

        // --- Build worksheet ---
        const ws = {};

        // ========== ROW 1: Title ==========
        const titleText = `${rocYear}年${month}月社區管理費未繳明細`;
        for (let c = 0; c <= 10; c++) {
            ws[ec(0, c)] = mkCell(c === 0 ? titleText : '', 's', {
                font: TITLE_FONT, alignment: CENTER
            });
        }

        // ========== ROW 2: Date ==========
        const dateText = `${rocYear}.${month}.${lastDay}截止資料`;
        for (let c = 6; c <= 10; c++) {
            ws[ec(1, c)] = mkCell(c === 6 ? dateText : '', 's', {
                font: BASE_FONT, alignment: RIGHT_CENTER
            });
        }

        // ========== ROW 3: Headers ==========
        const headers = ['序號', '基地', '社區', '地址', '承租人', '欠繳月份', '欠繳名目', '金額', '簡訊/LINE催款', '電話催款', '催款備註'];
        headers.forEach((h, c) => {
            const border = c === 0 ? HDR_LEFT : c === 10 ? HDR_RIGHT : HDR_MID;
            ws[ec(2, c)] = mkCell(h, 's', { font: BOLD_FONT, alignment: CENTER, border, fill: HEADER_FILL });
        });
        // L3: 連絡電話 (outside formal document range, no border/fill)
        ws[ec(2, 11)] = mkCell('連絡電話', 's', { font: BOLD_FONT, alignment: CENTER });

        // ========== DATA ROWS ==========
        const writeDataRow = (r, rowData) => {
            // r = 0-indexed row number
            const colDefs = [
                // [value, type, alignment, numFmt]
                [rowData.seq, 'n', CENTER, undefined],
                ['新店機廠', 's', CENTER, undefined],
                ['美河市', 's', CENTER, undefined],
                [rowData.address, 's', LEFT_CENTER_WRAP, undefined],
                [rowData.name, 's', LEFT_CENTER_WRAP, undefined],
                [rowData.period, 's', CENTER_WRAP, undefined],
                [rowData.itemType, 's', CENTER, '@'],
                [rowData.amount, 'n', RIGHT_CENTER_WRAP, '#,##0'],
                ['', 's', CENTER, undefined],
                ['', 's', CENTER, undefined],
                ['', 's', CENTER, undefined],
            ];
            colDefs.forEach(([v, t, alignment, numFmt], c) => {
                const border = c === 0 ? DATA_LEFT : c === 10 ? DATA_RIGHT : DATA_MID;
                const style = { font: BASE_FONT, alignment, border };
                if (numFmt) style.numFmt = numFmt;
                ws[ec(r, c)] = mkCell(v, t, style);
            });
            // L column: 連絡電話 (outside formal range, no border)
            ws[ec(r, 11)] = mkCell(rowData.phone || '', 's', { font: BASE_FONT, alignment: CENTER });
        };

        rows.forEach((row, idx) => {
            const r = DATA_START - 1 + idx; // 0-indexed
            writeDataRow(r, { seq: idx + 1, ...row });
        });

        // ========== BUFFER ROWS (empty, with borders) ==========
        for (let b = 0; b < BUFFER_ROWS; b++) {
            const r = DATA_START - 1 + rows.length + b;
            for (let c = 0; c <= 10; c++) {
                const border = c === 0 ? DATA_LEFT : c === 10 ? DATA_RIGHT : DATA_MID;
                ws[ec(r, c)] = mkCell('', 's', { font: BASE_FONT, border });
            }
        }

        // ========== STATS ROWS ==========
        const sr1 = statsR1 - 1; // 0-indexed
        const sr2 = statsR2 - 1;
        const fStart = DATA_START;
        const fEnd = formulaEndR;

        // Row: stats 1 - 房屋管理費
        // A:D merged (empty)
        for (let c = 0; c <= 3; c++) {
            ws[ec(sr1, c)] = mkCell('', 's', { font: BASE_FONT });
        }
        ws[ec(sr1, 4)] = mkCell('房屋管理費欠款戶數', 's', {
            font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL,
            border: { left: medium, right: thin, top: medium, bottom: thin }
        });
        ws[ec(sr1, 5)] = mkFormula(
            `COUNTIF(G${fStart}:G${fEnd},"房屋管理費")`,
            { font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL, border: { left: thin, right: thin, top: medium, bottom: thin } }
        );
        ws[ec(sr1, 6)] = mkCell('欠款總計', 's', {
            font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL,
            border: { left: thin, right: thin, top: medium, bottom: thin }
        });
        ws[ec(sr1, 7)] = mkFormula(
            `SUMIF(G${fStart}:G${fEnd},"房屋管理費",H${fStart}:H${fEnd})`,
            { font: BOLD_FONT, alignment: LEFT_CENTER, fill: HEADER_FILL, border: { left: thin, right: medium, top: medium, bottom: thin } },
            '#,##0_ '
        );
        // I:K merged (empty)
        for (let c = 8; c <= 10; c++) {
            ws[ec(sr1, c)] = mkCell('', 's', { font: BASE_FONT });
        }

        // Row: stats 2 - 車位清潔費
        for (let c = 0; c <= 3; c++) {
            ws[ec(sr2, c)] = mkCell('', 's', { font: BASE_FONT });
        }
        ws[ec(sr2, 4)] = mkCell('車位清潔費欠款戶數', 's', {
            font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL,
            border: { left: medium, right: thin, top: thin, bottom: medium }
        });
        ws[ec(sr2, 5)] = mkFormula(
            `COUNTIF(G${fStart}:G${fEnd},"車位清潔費")`,
            { font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL, border: { left: thin, right: thin, bottom: medium } }
        );
        ws[ec(sr2, 6)] = mkCell('欠款總計', 's', {
            font: BOLD_FONT, alignment: CENTER, fill: HEADER_FILL,
            border: { left: thin, right: thin, top: thin, bottom: medium }
        });
        ws[ec(sr2, 7)] = mkFormula(
            `SUMIF(G${fStart}:G${fEnd},"車位清潔費",H${fStart}:H${fEnd})`,
            { font: BOLD_FONT, alignment: LEFT_CENTER, fill: HEADER_FILL, border: { left: thin, right: medium, bottom: medium } },
            '#,##0_ '
        );
        for (let c = 8; c <= 10; c++) {
            ws[ec(sr2, c)] = mkCell('', 's', { font: BASE_FONT });
        }

        // ========== SUMMARY ROWS ==========
        const smr1 = sumR1 - 1; // 0-indexed
        const smr2 = sumR2 - 1;
        const smr3 = sumR3 - 1;

        // A column: "小結" (merged A:B across 3 rows)
        ws[ec(smr1, 0)] = mkCell('小結', 's', {
            font: BASE_FONT, alignment: CENTER,
            border: { left: medium, top: medium }
        });
        ws[ec(smr1, 1)] = mkCell('', 's', { font: BASE_FONT, border: { top: medium } });
        ws[ec(smr2, 0)] = mkCell('', 's', { font: BASE_FONT, border: { left: medium } });
        ws[ec(smr2, 1)] = mkCell('', 's', { font: BASE_FONT });
        ws[ec(smr3, 0)] = mkCell('', 's', { font: BASE_FONT, border: { left: medium, bottom: medium } });
        ws[ec(smr3, 1)] = mkCell('', 's', { font: BASE_FONT, border: { bottom: medium } });

        // C column: 一、二、三、
        ws[ec(smr1, 2)] = mkCell('一、', 's', {
            font: BASE_FONT, alignment: CENTER_WRAP, border: { top: medium }
        });
        ws[ec(smr2, 2)] = mkCell('二、', 's', {
            font: BASE_FONT, alignment: CENTER
        });
        ws[ec(smr3, 2)] = mkCell('三、', 's', {
            font: BASE_FONT, alignment: CENTER, border: { bottom: medium }
        });

        // D column: summary text (merged D:K)
        // Row 1: formula-based summary
        // Use COUNT (not COUNTA) so buffer rows with empty strings are not counted
        const summaryFormula =
            `"截至${rocYear}/${month}/${lastDay}，房屋管理費、車位清潔費欠繳者共"` +
            `&COUNT(A${fStart}:A${fEnd})` +
            `&"戶，欠款金額"` +
            `&TEXT(SUM(H${fStart}:H${fEnd}),"#,##0")` +
            `&"元，其中欠款達2個月，共計"` +
            `&COUNTIF(F${fStart}:F${fEnd},"*至*")` +
            `&"戶，已陸續通知承租人。"`;

        ws[ec(smr1, 3)] = mkFormula(summaryFormula, {
            font: BASE_FONT, alignment: LEFT_ALIGN_CENTER_WRAP, border: { top: medium }
        });
        // Fill merge cells for D:K in row 1
        for (let c = 4; c <= 9; c++) {
            ws[ec(smr1, c)] = mkCell('', 's', { font: BASE_FONT, border: { top: medium } });
        }
        ws[ec(smr1, 10)] = mkCell('', 's', { font: BASE_FONT, border: { right: medium, top: medium } });

        // Row 2: fixed text
        const fixedText2 = '美河市住辦區固定每月9-10號銷帳，住宅區不定期銷帳，如承租人在銷帳時間以後匯款，將無法顯示最新結果於催收表中。';
        ws[ec(smr2, 3)] = mkCell(fixedText2, 's', {
            font: BASE_FONT, alignment: LEFT_ALIGN_CENTER
        });
        for (let c = 4; c <= 9; c++) {
            ws[ec(smr2, c)] = mkCell('', 's', { font: BASE_FONT });
        }
        ws[ec(smr2, 10)] = mkCell('', 's', { font: BASE_FONT, border: { right: medium } });

        // Row 3: fixed text
        const fixedText3 = '本月份催收的手段有加強電話通知的頻率，除了電話、簡訊、貼單通知外，如欠款時間過久另會登門拜訪1-2次，雖未有住戶開門回應，但此舉可能有助提高住戶繳款之意願。';
        ws[ec(smr3, 3)] = mkCell(fixedText3, 's', {
            font: BASE_FONT, alignment: LEFT_ALIGN_CENTER_WRAP, border: { bottom: medium }
        });
        for (let c = 4; c <= 9; c++) {
            ws[ec(smr3, c)] = mkCell('', 's', { font: BASE_FONT, border: { bottom: medium } });
        }
        ws[ec(smr3, 10)] = mkCell('', 's', { font: BASE_FONT, border: { right: medium, bottom: medium } });

        // ========== MERGES ==========
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },       // A1:K1 title
            { s: { r: 1, c: 6 }, e: { r: 1, c: 10 } },        // G2:K2 date
            { s: { r: sr1, c: 0 }, e: { r: sr2, c: 3 } },     // A:D stats merged
            { s: { r: sr1, c: 8 }, e: { r: sr2, c: 10 } },    // I:K stats merged
            { s: { r: smr1, c: 0 }, e: { r: smr3, c: 1 } },   // A:B summary "小結"
            { s: { r: smr1, c: 3 }, e: { r: smr1, c: 10 } },  // D:K summary row 1
            { s: { r: smr2, c: 3 }, e: { r: smr2, c: 10 } },  // D:K summary row 2
            { s: { r: smr3, c: 3 }, e: { r: smr3, c: 10 } },  // D:K summary row 3
        ];

        // ========== COLUMN WIDTHS ==========
        ws['!cols'] = [
            { wch: 10 },     // A 序號
            { wch: 15.4 },   // B 基地
            { wch: 17.6 },   // C 社區
            { wch: 57.7 },   // D 地址
            { wch: 46.2 },   // E 承租人
            { wch: 34.6 },   // F 欠繳月份
            { wch: 19.6 },   // G 欠繳名目
            { wch: 14 },     // H 金額
            { wch: 19.7 },   // I 簡訊/LINE催款
            { wch: 15.6 },   // J 電話催款
            { wch: 34.9 },   // K 催款備註
            { wch: 18 },     // L 連絡電話 (正式文件範圍外)
        ];

        // ========== ROW HEIGHTS ==========
        const rowArr = [];
        rowArr[0] = { hpt: 20.25 };  // Row 1
        rowArr[1] = { hpt: 18.5 };   // Row 2
        rowArr[2] = { hpt: 18.5 };   // Row 3
        // Data + buffer rows
        for (let i = 3; i < 3 + rows.length + BUFFER_ROWS; i++) {
            rowArr[i] = { hpt: 24 };
        }
        // Stats rows
        rowArr[sr1] = { hpt: 22 };
        rowArr[sr2] = { hpt: 22 };
        // Summary rows
        rowArr[smr1] = { hpt: 35 };
        rowArr[smr2] = { hpt: 35 };
        rowArr[smr3] = { hpt: 35 };
        ws['!rows'] = rowArr;

        // ========== RANGE ==========
        ws['!ref'] = `A1:L${totalRows}`;

        // ========== OUTPUT ==========
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '管理費未收明細');

        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([excelBuffer], { type: 'application/octet-stream' });

    } catch (error) {
        console.error("Error processing files:", error);
        throw error;
    }
};
