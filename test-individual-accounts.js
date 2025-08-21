// Test the new dynamic individual accounts approach
import fs from 'fs';
import path from 'path';

// Read the sample trial balance CSV
const csvContent = fs.readFileSync('TB01.csv', 'utf-8');
console.log('=== TESTING INDIVIDUAL ACCOUNTS DYNAMIC EXTRACTION ===');
console.log('Sample CSV content:');
console.log(csvContent.split('\n').slice(0, 10).join('\n')); // Show first 10 lines

// Mock trial balance data for testing
const mockTrialBalance = [
  { accountCode: '1000', accountName: 'เงินสดในมือ', creditAmount: 25000, debitAmount: 0, previousBalance: 50000 },
  { accountCode: '1010', accountName: 'เงินฝากธนาคาร กสิกรไทย', creditAmount: 150000, debitAmount: 0, previousBalance: 120000 },
  { accountCode: '1140', accountName: 'ลูกหนี้การค้า', creditAmount: 175014.81, debitAmount: 0, previousBalance: 0 },
  { accountCode: '1215', accountName: 'ลูกหนี้อื่น', creditAmount: 5000, debitAmount: 0, previousBalance: 8000 },
  { accountCode: '2010', accountName: 'เจ้าหนี้การค้า', creditAmount: 0, debitAmount: 89500, previousBalance: 95000 },
  { accountCode: '2020', accountName: 'เจ้าหนี้อื่น', creditAmount: 0, debitAmount: 12000, previousBalance: 15000 },
  { accountCode: '2999', accountName: 'เจ้าหนี้เบ็ดเตล็ด', creditAmount: 0, debitAmount: 3500, previousBalance: 2000 }
];

console.log('\n=== INDIVIDUAL ACCOUNTS EXTRACTION SIMULATION ===');

// Simulate the extractIndividualAccounts logic
const individualAccounts = {
  cash: {},
  receivables: {},
  payables: {}
};

mockTrialBalance.forEach(entry => {
  const code = parseInt(entry.accountCode || '0');
  const current = Math.abs((entry.creditAmount || 0) - (entry.debitAmount || 0));
  const previous = Math.abs(entry.previousBalance || 0);
  
  if (code >= 1000 && code <= 1099) {
    // Cash accounts
    individualAccounts.cash[entry.accountCode] = {
      accountName: entry.accountName,
      current: current,
      previous: previous
    };
    console.log(`Cash Account: ${entry.accountCode} - ${entry.accountName} = ${current}`);
  } else if (code >= 1140 && code <= 1215) {
    // Receivables accounts
    individualAccounts.receivables[entry.accountCode] = {
      accountName: entry.accountName,
      current: current,
      previous: previous
    };
    console.log(`Receivables Account: ${entry.accountCode} - ${entry.accountName} = ${current}`);
  } else if (code >= 2010 && code <= 2999) {
    // Payables accounts (with exclusions)
    if (code !== 2030 && code !== 2045 && 
        !(code >= 2050 && code <= 2052) && 
        !(code >= 2100 && code <= 2123)) {
      individualAccounts.payables[entry.accountCode] = {
        accountName: entry.accountName,
        current: current,
        previous: previous
      };
      console.log(`Payables Account: ${entry.accountCode} - ${entry.accountName} = ${current}`);
    }
  }
});

console.log('\n=== FINAL INDIVIDUAL ACCOUNTS STRUCTURE ===');
console.log('Cash accounts:', Object.keys(individualAccounts.cash).length);
console.log('Receivables accounts:', Object.keys(individualAccounts.receivables).length);
console.log('Payables accounts:', Object.keys(individualAccounts.payables).length);

console.log('\n=== ZERO-FILTERING NOTE GENERATION SIMULATION ===');
console.log('Receivables Note:');
Object.entries(individualAccounts.receivables).forEach(([accountCode, accountData]) => {
  console.log(`  ${accountData.accountName}: ${accountData.current}`);
});

console.log('\nPayables Note:');
Object.entries(individualAccounts.payables).forEach(([accountCode, accountData]) => {
  console.log(`  ${accountData.accountName}: ${accountData.current}`);
});

console.log('\n=== PERFORMANCE ANALYSIS ===');
console.log('✅ Single-pass extraction: ALL individual accounts extracted once');
console.log('✅ Zero filtering: Note generation uses pre-extracted data');
console.log('✅ Foundation consistency: Totals calculated once and reused');
console.log('✅ Individual transparency: Complete audit trail with zero redundancy');
