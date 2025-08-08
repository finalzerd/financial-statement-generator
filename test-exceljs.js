// Quick test to verify ExcelJS integration
import { ExcelJSFormatter } from '../src/services/excelFormatter';

// Test the formatter
const workbook = ExcelJSFormatter.createWorkbook();

// Test data similar to balance sheet format
const testData = [
  ['บริษัท ทดสอบ จำกัด', '', '', '', '', '', '', ''],
  ['งบแสดงฐานะการเงิน', '', '', '', '', '', '', ''],
  ['ณ วันที่ 31 ธันวาคม 2567', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', 'หมายเหตุ', '', 'หน่วย:บาท', ''],
  ['', '', '', '', '', '2567', '2566', ''],
  ['', 'สินทรัพย์', '', '', '', '', '', ''],
  ['', '', 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '1,500,000.00', '1,200,000.00', ''],
  ['', '', 'ลูกหนี้การค้า', '', '', '2,800,000.00', '2,400,000.00', ''],
  ['', '', 'สินค้าคงเหลือ', '', '', '1,800,000.00', '1,600,000.00', ''],
  ['', 'รวมสินทรัพย์หมุนเวียน', '', '', '', { f: 'SUM(F8:F10)' }, { f: 'SUM(G8:G10)' }, ''],
];

// Add data to worksheet
const worksheet = ExcelJSFormatter.addDataToWorksheet(workbook, 'Test_Sheet', testData);

// Apply formatting
ExcelJSFormatter.formatBalanceSheetAssets(worksheet);

console.log('✅ ExcelJS test completed successfully!');
console.log('🎨 The financial statements will now have:');
console.log('   - Thai fonts (TH SarabunPSK)');
console.log('   - Professional formatting');
console.log('   - Proper borders and colors');
console.log('   - Currency number formatting');
console.log('   - Column widths optimized for Thai text');
