// Test script to demonstrate financial statement output format
import { FinancialStatementGenerator } from './src/services/financialStatementGenerator';
import type { TrialBalanceEntry, CompanyInfo } from './src/types/financial';

// Sample trial balance data
const sampleTrialBalance: TrialBalanceEntry[] = [
  { accountCode: '1010', accountName: 'เงินสดในมือ', debitAmount: 50000, creditAmount: 0, balance: 50000 },
  { accountCode: '1015', accountName: 'เงินฝากธนาคาร', debitAmount: 200000, creditAmount: 0, balance: 200000 },
  { accountCode: '1140', accountName: 'ลูกหนี้การค้า', debitAmount: 80000, creditAmount: 0, balance: 80000 },
  { accountCode: '1300', accountName: 'สินค้าคงเหลือ', debitAmount: 100000, creditAmount: 0, balance: 100000 },
  { accountCode: '1610', accountName: 'ที่ดิน', debitAmount: 500000, creditAmount: 0, balance: 500000 },
  { accountCode: '1615', accountName: 'อาคาร', debitAmount: 300000, creditAmount: 100000, balance: 200000 },
  { accountCode: '2010', accountName: 'เจ้าหนี้การค้า', debitAmount: 0, creditAmount: 60000, balance: -60000 },
  { accountCode: '2110', accountName: 'เงินกู้ยืมระยะยาว', debitAmount: 0, creditAmount: 150000, balance: -150000 },
  { accountCode: '3010', accountName: 'ทุนจดทะเบียน', debitAmount: 0, creditAmount: 500000, balance: -500000 },
  { accountCode: '3020', accountName: 'กำไรสะสม', debitAmount: 0, creditAmount: 220000, balance: -220000 },
  { accountCode: '4010', accountName: 'รายได้จากการขาย', debitAmount: 0, creditAmount: 800000, balance: -800000 },
  { accountCode: '4210', accountName: 'รายได้อื่น', debitAmount: 0, creditAmount: 50000, balance: -50000 },
  { accountCode: '5010', accountName: 'ต้นทุนขาย', debitAmount: 400000, creditAmount: 0, balance: 400000 },
  { accountCode: '5310', accountName: 'เงินเดือน', debitAmount: 120000, creditAmount: 0, balance: 120000 },
  { accountCode: '5320', accountName: 'ค่าเช่า', debitAmount: 60000, creditAmount: 0, balance: 60000 },
  { accountCode: '5910', accountName: 'ภาษีเงินได้', debitAmount: 20000, creditAmount: 0, balance: 20000 }
];

const sampleCompanyInfo: CompanyInfo = {
  name: 'บริษัท ตัวอย่างการทำงบการเงิน จำกัด',
  type: 'บริษัทจำกัด',
  reportingPeriod: '1 มกราคม 2024 ถึง 31 ธันวาคม 2024',
  reportingYear: 2024,
  registrationNumber: '0000000000000'
};

const generator = new FinancialStatementGenerator();

// Test single-year generation
const statements = generator.generateFinancialStatements(
  sampleTrialBalance,
  sampleCompanyInfo,
  'single-year'
);

// Display sample output structure
console.log('=== Balance Sheet Assets (Sample Output) ===');
console.log('Row 1:', statements.balanceSheet.assets[0]);
console.log('Row 2:', statements.balanceSheet.assets[1]);
console.log('Row 3:', statements.balanceSheet.assets[2]);
console.log('Row 4:', statements.balanceSheet.assets[3]);
console.log('Row 5:', statements.balanceSheet.assets[4]);
console.log('Row 6:', statements.balanceSheet.assets[5]);
console.log('Row 7:', statements.balanceSheet.assets[6]);
console.log('Row 8:', statements.balanceSheet.assets[7]);

console.log('\n=== Profit & Loss Statement (Sample Output) ===');
console.log('Row 1:', statements.profitLossStatement[0]);
console.log('Row 2:', statements.profitLossStatement[1]);
console.log('Row 3:', statements.profitLossStatement[2]);
console.log('Row 6:', statements.profitLossStatement[5]);
console.log('Row 7:', statements.profitLossStatement[6]);
console.log('Row 8:', statements.profitLossStatement[7]);

export { statements, sampleTrialBalance, sampleCompanyInfo };
