/**
 * RESERVED UNUSED METHODS FROM financialStatementGenerator.ts
 * 
 * These methods were moved here during code cleanup to preserve them for potential future use.
 * They represent alternative implementations and enhanced features that were not being used
 * in the current production version.
 * 
 * Date: September 5, 2025
 * Original file: financialStatementGenerator.ts
 * Reason for moving: TypeScript compiler reported these methods as unused (never read)
 */

import type { TrialBalanceEntry, CompanyInfo } from '../../types/financial';

// Types needed for the unused methods
interface NoteRowTracker {
  currentRow: number;
  noteStartRow: number;
  headerRows: number[];
  yearHeaderRows: number[];
  detailRows: number[];
  totalRows: number[];
  unitRows: number[];
}

interface DetailedFinancialData {
  noteCalculations: {
    cash: {
      cash: { current: number; previous: number };
      bankDeposits: { current: number; previous: number };
      total: { current: number; previous: number };
    };
    receivables: {
      total: { current: number; previous: number };
    };
    payables: {
      total: { current: number; previous: number };
    };
  };
  balanceSheetTotals: {
    liabilities: {
      bankOverdraftsAndShortTermLoans: { current: number; previous: number };
      shortTermBorrowings: { current: number; previous: number };
    };
  };
  individualAccounts: {
    receivables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
      };
    };
    payables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
      };
    };
  };
}

/**
 * UNUSED METHOD: Enhanced PPE Note with Row Tracking
 * Original location: Line 1382 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation with enhanced tracking
 */
export function addPPENoteWithRowTrackingEnhanced(
  notes: any[][],
  trialBalanceData: TrialBalanceEntry[],
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  _trialBalancePrevious?: TrialBalanceEntry[],
  noteNumber: number = 6
): NoteRowTracker {
  const tracker: NoteRowTracker = {
    currentRow: notes.length + 1,
    noteStartRow: notes.length + 1,
    headerRows: [],        // Note headers and section headers
    yearHeaderRows: [],    // Year/column headers
    detailRows: [],        // Individual asset/depreciation accounts
    totalRows: [],         // Total/รวม rows and net book value
    unitRows: []           // หน่วย:บาท rows
  };

  // Get all PPE asset accounts (1610-1659 without decimal points)
  const assetAccounts = trialBalanceData.filter(entry => {
    const code = parseInt(entry.accountCode);
    return code >= 1610 && code <= 1659 && !entry.accountCode.includes('.');
  });

  // Get all accumulated depreciation accounts (1610-1659 with decimal points)
  const depreciationAccounts = trialBalanceData.filter(entry => {
    const code = parseInt(entry.accountCode);
    return code >= 1610 && code <= 1659 && entry.accountCode.includes('.');
  });

  // Only create note if there are any PPE accounts with balances
  if (assetAccounts.length === 0 && depreciationAccounts.length === 0) {
    return tracker;
  }

  // Check if any accounts have non-zero balances
  const hasAssetBalances = assetAccounts.some(acc => 
    Math.abs(acc.balance) !== 0 || Math.abs(acc.previousBalance || 0) !== 0
  );
  const hasDepreciationBalances = depreciationAccounts.some(acc => 
    Math.abs(acc.balance) !== 0 || Math.abs(acc.previousBalance || 0) !== 0
  );

  if (!hasAssetBalances && !hasDepreciationBalances) {
    return tracker;
  }

  console.log(`=== PPE NOTE ENHANCED ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

  // 1. Note Header Row
  notes.push([noteNumber.toString(), 'ที่ดิน อาคารและอุปกรณ์', '', '', '', '', '', '', 'หน่วย:บาท']);
  tracker.headerRows.push(tracker.currentRow);
  tracker.unitRows.push(tracker.currentRow);
  tracker.currentRow++;
  
  // 2. Column Headers (Year Headers) - Same structure for both single and multi-year
  if (processingType === 'multi-year') {
    notes.push(['', '', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear - 1}`, '', 'ซื้อเพิ่ม', 'จำหน่ายออก', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear}`]);
  } else {
    // Single-year: Same structure but use same year for both columns
    notes.push(['', '', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear}`, '', 'ซื้อเพิ่ม', 'จำหน่ายออก', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear}`]);
  }
  tracker.yearHeaderRows.push(tracker.currentRow);
  tracker.currentRow++;
  
  // 3. Asset Cost Section Header
  notes.push(['', '', 'ราคาทุนเดิม', '', '', '', '', '', '']);
  tracker.headerRows.push(tracker.currentRow); // Section header
  tracker.currentRow++;
  
  let assetTotalCurrent = 0;
  let assetTotalPrevious = 0;
  let assetTotalPurchases = 0;
  let assetTotalDisposals = 0;
  const assetStartRow = tracker.currentRow; // Track start of asset details (our internal tracking)

  // 4. Individual Asset Accounts (Detail Rows)
  assetAccounts.forEach(account => {
    const currentAmount = Math.abs(account.balance);
    const previousAmount = Math.abs(account.previousBalance || 0);
    const purchases = Math.max(0, currentAmount - previousAmount); // Only positive purchases
    const disposals = Math.max(0, previousAmount - currentAmount); // Only positive disposals
    
    if (currentAmount !== 0 || previousAmount !== 0) {
      // Use same structure for both single and multi-year processing
      const purchaseValue = purchases > 0 ? purchases : 0;
      const disposalValue = disposals > 0 ? disposals : 0;
      
      notes.push(['', '', account.accountName, 
        processingType === 'multi-year' ? previousAmount : '', '', // Column D: Previous amount (blank for single-year)
        purchaseValue, // Column F: Purchases (show 0 if no purchases)
        disposalValue, // Column G: Disposals (show 0 if no disposals)
        '', currentAmount]); // Column I: Current amount
      
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
      
      assetTotalCurrent += currentAmount;
      assetTotalPrevious += previousAmount;
      assetTotalPurchases += purchases;
      assetTotalDisposals += disposals;
    }
  });

  // 5. Asset Totals with Excel formulas (Total Row) - Same structure for both processing types
  const assetEndRow = tracker.currentRow - 1; // Last row of asset details (Excel 1-indexed)
  notes.push(['', '', 'รวม', 
    processingType === 'multi-year' ? { f: `SUM(D${assetStartRow}:D${assetEndRow})` } : '', '', // Column D: Previous total formula (blank for single-year)
    { f: `SUM(F${assetStartRow}:F${assetEndRow})` }, // Column F: Total purchases formula
    { f: `SUM(G${assetStartRow}:G${assetEndRow})` }, // Column G: Total disposals formula
    '', { f: `SUM(I${assetStartRow}:I${assetEndRow})` }]); // Column I: Current total formula
  tracker.totalRows.push(tracker.currentRow);
  tracker.currentRow++;
  
  // 7. Accumulated Depreciation Section Header
  notes.push(['', '', 'ค่าเสื่อมราคาสะสม', '', '', '', '', '', '']);
  tracker.headerRows.push(tracker.currentRow); // Section header
  tracker.currentRow++;
  
  let depreciationTotalCurrent = 0;
  let depreciationTotalPrevious = 0;
  let depreciationExpense = 0;
  let depreciationDisposal = 0;
  const depreciationStartRow = tracker.currentRow; // Track start of depreciation details

  // 8. Individual Depreciation Accounts (Detail Rows)
  depreciationAccounts.forEach(account => {
    const currentAmount = Math.abs(account.balance); // Convert to positive
    const previousAmount = Math.abs(account.previousBalance || 0); // Convert to positive
    const expenseAmount = Math.max(0, currentAmount - previousAmount); // Depreciation expense for the year
    const disposalAmount = Math.max(0, previousAmount - currentAmount); // Depreciation disposal for the year
    
    if (currentAmount !== 0 || previousAmount !== 0) {
      // Use same structure for both single and multi-year processing
      const expenseValue = expenseAmount > 0 ? expenseAmount : 0;
      const disposalAmountValue = disposalAmount > 0 ? disposalAmount : 0;
      
      notes.push(['', '', account.accountName, 
        processingType === 'multi-year' ? previousAmount : '', '', // Column D: Previous depreciation (blank for single-year)
        expenseValue, // Column F: Depreciation expense (show 0 if no expense)
        disposalAmountValue, // Column G: Depreciation disposal (show 0 if no disposal)
        '', currentAmount]); // Column I: Current depreciation
      
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
      
      depreciationTotalCurrent += currentAmount;
      depreciationTotalPrevious += previousAmount;
      depreciationExpense += expenseAmount;
      depreciationDisposal += disposalAmount;
    }
  });

  // 9. Depreciation Totals with Excel formulas (Total Row) - Same structure for both processing types
  const depreciationEndRow = tracker.currentRow - 1; // Last row of depreciation details
  
  notes.push(['', '', 'รวม', 
    processingType === 'multi-year' ? { f: `SUM(D${depreciationStartRow}:D${depreciationEndRow})` } : '', '', // Column D: Previous depreciation total formula (blank for single-year)
    { f: `SUM(F${depreciationStartRow}:F${depreciationEndRow})` }, // Column F: Total depreciation expense formula
    { f: `SUM(G${depreciationStartRow}:G${depreciationEndRow})` }, // Column G: Total depreciation disposal formula
    '', { f: `SUM(I${depreciationStartRow}:I${depreciationEndRow})` }]); // Column I: Current depreciation total formula
  tracker.totalRows.push(tracker.currentRow);
  tracker.currentRow++;

  // 10. Net book value with formulas referencing the totals above (Total Row)
  const assetTotalRowIndex = assetEndRow + 1; // Row number of asset totals
  const depreciationTotalRowIndex = tracker.currentRow - 1; // Row number of depreciation totals (just added above)
  
  // Net book value - Same structure for both processing types
  notes.push(['', '', 'มูลค่าสุทธิ', 
    processingType === 'multi-year' ? { f: `D${assetTotalRowIndex}-D${depreciationTotalRowIndex}` } : '', '', '', '', // Column D: Previous net value formula (blank for single-year)
    '', { f: `I${assetTotalRowIndex}-I${depreciationTotalRowIndex}` }]); // Column I: Current net value formula
  tracker.totalRows.push(tracker.currentRow);
  tracker.currentRow++;

  // 11. Depreciation expense summary - reference the depreciation expense total
  if (depreciationExpense > 0) {
    notes.push(['', '', 'ค่าเสื่อมราคา', '', '', '', '', '', 
      { f: `F${depreciationTotalRowIndex}` }]); // Column I (index 8): Reference depreciation expense total
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;
  }
  
  // 12. Final spacer
  notes.push(['', '', '', '', '', '', '', '', '']);
  tracker.currentRow++;

  console.log(`PPE Enhanced Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
  return tracker;
}

/**
 * UNUSED METHOD: Cash Note using Global Data
 * Original location: Line 1578 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation using global data architecture
 */
export function addCashNoteWithGlobalData(
  notes: any[][],
  globalData: DetailedFinancialData,
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 3
): void {
  // *** FOUNDATION-FIRST: Use note calculations as the source of truth ***
  const totalAmount = globalData.noteCalculations.cash.total.current;
  const prevTotalAmount = globalData.noteCalculations.cash.total.previous;

  // Only add note if there are actual balances
  if (totalAmount !== 0 || prevTotalAmount !== 0) {
    notes.push([noteNumber.toString(), 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    
    // *** FOUNDATION-FIRST with GROUPED BREAKDOWN: Show grouped categories from foundation layer ***
    console.log('=== FOUNDATION-FIRST CASH NOTE WITH GROUPED BREAKDOWN ===');
    console.log('Total (from global):', totalAmount, 'vs Previous:', prevTotalAmount);
    
    // Use foundation layer grouped calculations (cash + bank deposits)
    const cashAmount = globalData.noteCalculations.cash.cash.current;
    const bankAmount = globalData.noteCalculations.cash.bankDeposits.current;
    const prevCashAmount = globalData.noteCalculations.cash.cash.previous;
    const prevBankAmount = globalData.noteCalculations.cash.bankDeposits.previous;
    
    console.log('Cash (grouped from 1000):', cashAmount, 'vs Previous:', prevCashAmount);
    console.log('Bank Deposits (grouped from 1010-1099):', bankAmount, 'vs Previous:', prevBankAmount);
    
    // Show grouped breakdown from foundation layer
    if (cashAmount !== 0 || prevCashAmount !== 0) {
      notes.push(['', '', 'เงินสดในมือ', '', '', '', 
        cashAmount, '', 
        processingType === 'multi-year' ? prevCashAmount : '']);
    }
    
    if (bankAmount !== 0 || prevBankAmount !== 0) {
      notes.push(['', '', 'เงินฝากธนาคาร', '', '', '', 
        bankAmount, '', 
        processingType === 'multi-year' ? prevBankAmount : '']);
    }
    
    // Use foundation layer total (perfect consistency with Balance Sheet)
    if (processingType === 'multi-year') {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', prevTotalAmount]);
    } else {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', '']);
    }
    notes.push(['', '', '', '', '', '', '', '', '']);
  }
}

/**
 * UNUSED METHOD: Bank Overdrafts Note using Global Data
 * Original location: Line 1635 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation using global data architecture
 */
export function addBankOverdraftsNoteWithGlobalData(
  notes: any[][],
  globalData: DetailedFinancialData,
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 9
): void {
  const totalAmount = globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.current;
  const prevTotalAmount = globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.previous;

  // Only add note if there are actual balances
  if (totalAmount !== 0 || prevTotalAmount !== 0) {
    notes.push([noteNumber.toString(), 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      notes.push(['', '', 'เงินเบิกเกินบัญชีธนาคาร', '', '', '', totalAmount, '', prevTotalAmount]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      notes.push(['', '', 'เงินเบิกเกินบัญชีธนาคาร', '', '', '', totalAmount, '', '']);
    }
    notes.push(['', '', '', '', '', '', '', '', '']);
  }
}

/**
 * UNUSED METHOD: Short Term Borrowings Note using Global Data
 * Original location: Line 1660 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation using global data architecture
 */
export function addShortTermBorrowingsNoteWithGlobalData(
  notes: any[][],
  globalData: DetailedFinancialData,
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 13
): void {
  const totalAmount = globalData.balanceSheetTotals.liabilities.shortTermBorrowings.current;
  const prevTotalAmount = globalData.balanceSheetTotals.liabilities.shortTermBorrowings.previous;

  // Only add note if there are actual balances
  if (totalAmount !== 0 || prevTotalAmount !== 0) {
    notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      notes.push(['', '', 'เงินกู้ยืมจากสถาบันการเงิน', '', '', '', totalAmount, '', prevTotalAmount]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      notes.push(['', '', 'เงินกู้ยืมจากสถาบันการเงิน', '', '', '', totalAmount, '', '']);
    }
    notes.push(['', '', '', '', '', '', '', '', '']);
  }
}

/**
 * UNUSED METHOD: Trade Receivables Note using Global Data
 * Original location: Line 1685 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation with partial optimization
 */
export function addTradeReceivablesNoteWithGlobalData(
  notes: any[][],
  globalData: DetailedFinancialData,
  trialBalanceData: TrialBalanceEntry[],
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 4
): void {
  // *** FOUNDATION-FIRST: Use note calculations as the source of truth ***
  const totalAmount = globalData.noteCalculations.receivables.total.current;
  const prevTotalAmount = globalData.noteCalculations.receivables.total.previous;

  // Only add note if there are actual balances
  if (totalAmount !== 0 || prevTotalAmount !== 0) {
    notes.push([noteNumber.toString(), 'ลูกหนี้การค้าและลูกหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    
    // *** FOUNDATION-FIRST with DETAILED BREAKDOWN: Show ALL individual receivable accounts from trial balance ***
    console.log('=== DETAILED RECEIVABLES BREAKDOWN FROM TRIAL BALANCE ===');
    
    // Show ALL individual receivable accounts from trial balance (1140-1215)
    const receivableAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= 1140 && code <= 1215;
    });
    
    for (const account of receivableAccounts) {
      const currentAmount = Math.abs(account.balance || 0);
      const previousAmount = processingType === 'multi-year' ? 
        Math.abs(account.previousBalance || 0) : 0;
        
      // Only show accounts with non-zero balances
      if (currentAmount !== 0 || previousAmount !== 0) {
        notes.push(['', '', account.accountName || `บัญชี ${account.accountCode}`, '', '', '', 
          currentAmount, '', 
          processingType === 'multi-year' ? previousAmount : '']);
      }
    }
    
    if (processingType === 'multi-year') {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', prevTotalAmount]);
    } else {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', '']);
    }
    notes.push(['', '', '', '', '', '', '', '', '']);
  }
}

/**
 * UNUSED METHOD: Trade Payables Note using Global Data
 * Original location: Line 1738 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation with partial optimization
 */
export function addTradePayablesNoteWithGlobalData(
  notes: any[][],
  globalData: DetailedFinancialData,
  trialBalanceData: TrialBalanceEntry[],
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 12
): void {
  // *** FOUNDATION-FIRST: Use note calculations as the source of truth ***
  const totalAmount = globalData.noteCalculations.payables.total.current;
  const prevTotalAmount = globalData.noteCalculations.payables.total.previous;

  // Only add note if there are actual balances
  if (totalAmount !== 0 || prevTotalAmount !== 0) {
    notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    
    // *** FOUNDATION-FIRST with DETAILED BREAKDOWN: Show ALL individual payable accounts from trial balance ***
    console.log('=== DETAILED PAYABLES BREAKDOWN FROM TRIAL BALANCE ===');
    
    // Show ALL individual payable accounts from trial balance (2010-2999, excluding specific accounts)
    const payableAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      // Include 2010-2999 but exclude specific accounts that are handled separately
      return code >= 2010 && code <= 2999 && 
             code !== 2030 && // Short-term borrowings
             code !== 2045 && // Income tax payable
             !(code >= 2050 && code <= 2052) && // Other long-term loans
             !(code >= 2100 && code <= 2123); // Long-term loans from FI
    });
    
    for (const account of payableAccounts) {
      const currentAmount = Math.abs(account.balance || 0);
      const previousAmount = processingType === 'multi-year' ? 
        Math.abs(account.previousBalance || 0) : 0;
        
      // Only show accounts with non-zero balances
      if (currentAmount !== 0 || previousAmount !== 0) {
        notes.push(['', '', account.accountName || `บัญชี ${account.accountCode}`, '', '', '', 
          currentAmount, '', 
          processingType === 'multi-year' ? previousAmount : '']);
      }
    }
    
    if (processingType === 'multi-year') {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', prevTotalAmount]);
    } else {
      notes.push(['', '', 'รวม', '', '', '', totalAmount, '', '']);
    }
    notes.push(['', '', '', '', '', '', '', '', '']);
  }
}

/**
 * UNUSED METHOD: Trade Receivables Note using Individual Accounts
 * Original location: Line 1796 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation eliminating ALL filtering
 */
export function addTradeReceivablesNoteWithIndividualAccounts(
  notes: any[][],
  globalData: DetailedFinancialData,
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 4
): void {
  const receivableAccounts = globalData.individualAccounts.receivables;
  const totalAmount = globalData.noteCalculations.receivables.total.current;
  const prevTotalAmount = globalData.noteCalculations.receivables.total.previous;

  // Only add note if there are actual balances
  if (totalAmount === 0 && prevTotalAmount === 0 && Object.keys(receivableAccounts).length === 0) {
    return;
  }

  console.log('=== INDIVIDUAL RECEIVABLES NOTE (ZERO FILTERING!) ===');
  console.log('Individual accounts found:', Object.keys(receivableAccounts).length);
  console.log('Total from foundation:', totalAmount);

  // Add headers
  notes.push([noteNumber.toString(), 'ลูกหนี้การค้าและลูกหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
  if (processingType === 'multi-year') {
    notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
  } else {
    notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
  }

  // *** ZERO PROCESSING! Just iterate through pre-extracted individual accounts ***
  Object.entries(receivableAccounts).forEach(([accountCode, accountData]) => {
    console.log(`Showing individual account: ${accountCode} - ${accountData.accountName} = ${accountData.current}`);
    notes.push(['', '', accountData.accountName, '', '', '', 
      accountData.current, '', 
      processingType === 'multi-year' ? accountData.previous : '']);
  });

  // Total from foundation layer (guaranteed consistency with Balance Sheet)
  if (processingType === 'multi-year') {
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', prevTotalAmount]);
  } else {
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', '']);
  }
  notes.push(['', '', '', '', '', '', '', '', '']);
  console.log('=== END INDIVIDUAL RECEIVABLES NOTE ===');
}

/**
 * UNUSED METHOD: Trade Payables Note using Individual Accounts
 * Original location: Line 1843 in financialStatementGenerator.ts
 * Reason: Never used - alternative implementation eliminating ALL filtering
 */
export function addTradePayablesNoteWithIndividualAccounts(
  notes: any[][],
  globalData: DetailedFinancialData,
  companyInfo: CompanyInfo,
  processingType: 'single-year' | 'multi-year',
  noteNumber: number = 12
): void {
  const payableAccounts = globalData.individualAccounts.payables;
  const totalAmount = globalData.noteCalculations.payables.total.current;
  const prevTotalAmount = globalData.noteCalculations.payables.total.previous;

  // Only add note if there are actual balances
  if (totalAmount === 0 && prevTotalAmount === 0 && Object.keys(payableAccounts).length === 0) {
    return;
  }

  console.log('=== INDIVIDUAL PAYABLES NOTE (ZERO FILTERING!) ===');
  console.log('Individual accounts found:', Object.keys(payableAccounts).length);
  console.log('Total from foundation:', totalAmount);

  // Add headers
  notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
  if (processingType === 'multi-year') {
    notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
  } else {
    notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
  }

  // *** ZERO PROCESSING! Just iterate through pre-extracted individual accounts ***
  Object.entries(payableAccounts).forEach(([accountCode, accountData]) => {
    console.log(`Showing individual account: ${accountCode} - ${accountData.accountName} = ${accountData.current}`);
    notes.push(['', '', accountData.accountName, '', '', '', 
      accountData.current, '', 
      processingType === 'multi-year' ? accountData.previous : '']);
  });

  // Total from foundation layer (guaranteed consistency with Balance Sheet)
  if (processingType === 'multi-year') {
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', prevTotalAmount]);
  } else {
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', '']);
  }
  notes.push(['', '', '', '', '', '', '', '', '']);
  console.log('=== END INDIVIDUAL PAYABLES NOTE ===');
}
