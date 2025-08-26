// ============================================================================
// PPE NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates Property, Plant and Equipment note (Note 10) with complex structure
 * Handles cost/depreciation breakdown with movement tracking
 */
export class PPENoteGenerator {
  
  /**
   * Generate PPE note with enhanced row tracking and Excel formulas
   * Supports both single-year and multi-year processing
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
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
}
