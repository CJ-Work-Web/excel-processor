import { formatAddress, formatDateRange, calculateArrearsItem, formatPhoneNumber } from './src/utils/excelProcessor.js';

console.log("Starting Verification...");

// Test 1: Address Formatting
const testAddress = (input, expected) => {
    const result = formatAddress(input);
    if (result === expected) {
        console.log(`[PASS] Address ${input} -> ${result}`);
    } else {
        console.error(`[FAIL] Address ${input} -> ${result} (Expected: ${expected})`);
    }
};

testAddress(530902, "新北市新店區中央路153號9樓之2");
testAddress("530900", "新北市新店區中央路153號9樓");

// Test 2: Date Formatting
const testDate = (input, expected) => {
    const result = formatDateRange(input);
    if (result === expected) {
        console.log(`[PASS] Date ${input} -> ${result}`);
    } else {
        console.error(`[FAIL] Date ${input} -> ${result} (Expected: ${expected})`);
    }
};

testDate("2026/01/31~2026/01/31", "115年1月");
testDate("2026/01/01~2026/02/01", "115年1月至115年2月");

// Test 3: Arrears Item
const testArrears = (input, expected) => {
    const result = calculateArrearsItem(input);
    if (result === expected) {
        console.log(`[PASS] Amount ${input} -> ${result}`);
    } else {
        console.error(`[FAIL] Amount ${input} -> ${result} (Expected: ${expected})`);
    }
};

testArrears(500, "車位清潔費");
testArrears(1000, "車位清潔費");
testArrears(501, "房屋管理費");

// Test 4: Phone Formatting
const testPhone = (input, expected) => {
    const result = formatPhoneNumber(input);
    if (result === expected) {
        console.log(`[PASS] Phone ${input} -> ${result}`);
    } else {
        console.error(`[FAIL] Phone ${input} -> ${result} (Expected: ${expected})`);
    }
};

testPhone("0912345678", "0912-345-678");
testPhone("0911222333", "0911-222-333");
testPhone("0223456789", "02-2345-6789");
testPhone("0412345678", "04-1234-5678");
testPhone("1234567890", "12-3456-7890"); // Fallback rule for non-09 but 10 digits
testPhone("912345678", "0912-345-678"); // Test prepending 0
testPhone("123", "123"); // Too short, return as is (cleaned)
testPhone("0912-345-678", "0912-345-678"); // Already formatted input (will be cleaned and re-formatted)

console.log("Verification Complete.");
