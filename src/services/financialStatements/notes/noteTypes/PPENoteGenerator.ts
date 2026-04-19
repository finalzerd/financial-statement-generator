// ============================================================================
// PPE NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';
import type { SelectionFirstResult, ClassifiedAccount } from '../../selection/SelectionFirstClassifier';

/**
 * Generates Property, Plant and Equipment note (Note 10) with complex structure
 * Handles cost/depreciation breakdown with movement tracking
 */
export class PPENoteGenerator {
  
  /**
   * Abbreviate Thai month names for compact date display
   * Example: "31 ธันวาคม" → "31 ธ.ค."
   */
  private static abbreviateThaiMonth(dateString: string): string {
    const monthAbbreviations: { [key: string]: string } = {
      'มกราคม': 'ม.ค.',
      'กุมภาพันธ์': 'ก.พ.',
      'มีนาคม': 'มี.ค.',
      'เมษายน': 'เม.ย.',
      'พฤษภาคม': 'พ.ค.',
      'มิถุนายน': 'มิ.ย.',
      'กรกฎาคม': 'ก.ค.',
      'สิงหาคม': 'ส.ค.',
      'กันยายน': 'ก.ย.',
      'ตุลาคม': 'ต.ค.',
      'พฤศจิกายน': 'พ.ย.',
      'ธันวาคม': 'ธ.ค.'
    };
    
    let result = dateString;
    for (const [fullMonth, abbrev] of Object.entries(monthAbbreviations)) {
      result = result.replace(fullMonth, abbrev);
    }
    return result;
  }
  
  /**
   * Generate PPE note with enhanced row tracking and Excel formulas
   * Supports both single-year and multi-year processing
   */
  static generateWithRowTracking(
    notes: any[][], 
    _trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    _trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 6,
    selection?: SelectionFirstResult
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

    type PPEAccount = {
      accountCode?: string;
      accountName: string;
      current: number;
      previous: number;
    };

    const normalizeSelection = (accounts?: ClassifiedAccount[]): PPEAccount[] => {
      if (!accounts) return [];
      return accounts.map((acc) => ({
        accountCode: acc.accountCode,
        accountName: acc.accountName,
        current: acc.current || 0,
        previous: acc.previous || 0
      }));
    };

    const isDecimalAccount = (code?: string) => Boolean(code && code.includes('.'));
    const selectionCostSource = normalizeSelection(selection?.byCategory?.ppe_cost);
    const selectionDeprSource = normalizeSelection(selection?.byCategory?.ppe_accum_depr);

    console.log(`[PPE Note] Selection data received - Cost accounts: ${selectionCostSource.length}, Depr accounts: ${selectionDeprSource.length}`);
    if (selectionCostSource.length > 0) {
      console.log(`[PPE Note] Cost selection codes: ${selectionCostSource.map(a => a.accountCode).slice(0, 10).join(', ')}${selectionCostSource.length > 10 ? '...' : ''}`);
    }
    if (selectionDeprSource.length > 0) {
      console.log(`[PPE Note] Depr selection codes: ${selectionDeprSource.map(a => a.accountCode).slice(0, 10).join(', ')}${selectionDeprSource.length > 10 ? '...' : ''}`);
    }

    const selectionCostAccounts = selectionCostSource.filter(acc => !isDecimalAccount(acc.accountCode));
    const spilloverDecimalAccounts = selectionCostSource.filter(acc => isDecimalAccount(acc.accountCode));

    const seenCodes = new Set<string>();
    const addUnique = (list: PPEAccount[], target: PPEAccount[]) => {
      for (const acc of list) {
        const key = acc.accountCode ?? acc.accountName;
        if (seenCodes.has(key)) continue;
        seenCodes.add(key);
        target.push(acc);
      }
    };

    const selectionDeprAccounts: PPEAccount[] = [];
    addUnique(selectionDeprSource, selectionDeprAccounts);
    addUnique(spilloverDecimalAccounts, selectionDeprAccounts);

    // Debug: Check if 1646 is in any list
    const check1646InCost = selectionCostAccounts.find(a => a.accountCode?.startsWith('1646'));
    const check1646InDepr = selectionDeprAccounts.find(a => a.accountCode?.startsWith('1646'));
    if (check1646InCost) {
      console.log(`[PPE Note] DEBUG: Found 1646 in COST accounts: ${JSON.stringify(check1646InCost)}`);
    }
    if (check1646InDepr) {
      console.log(`[PPE Note] DEBUG: Found 1646 in DEPRECIATION accounts: ${JSON.stringify(check1646InDepr)}`);
    }

    // Use ONLY selection-first data - no fallback to ensure database mappings are respected
    const assetAccounts = selectionCostAccounts
      .filter(acc => acc.current !== 0 || acc.previous !== 0);
    const depreciationAccounts = selectionDeprAccounts
      .filter(acc => acc.current !== 0 || acc.previous !== 0);

    console.log(`[PPE Note] Using selection-first data -> assets: ${assetAccounts.length}, depreciation: ${depreciationAccounts.length}`);
    console.log(`[PPE Note] Asset account codes: ${assetAccounts.map(a => a.accountCode).join(', ')}`);
    console.log(`[PPE Note] Depreciation account codes: ${depreciationAccounts.map(a => a.accountCode).join(', ')}`);

    if (assetAccounts.length === 0 && depreciationAccounts.length === 0) {
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
      const prevDateFull = `ณ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear - 1}`;
      const currDateFull = `ณ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`;
      notes.push(['', '', '', this.abbreviateThaiMonth(prevDateFull), '', 'ซื้อเพิ่ม', 'จำหน่ายออก', '', this.abbreviateThaiMonth(currDateFull)]);
    } else {
      // Single-year: Same structure but use same year for both columns
      const currDateFull = `ณ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`;
      notes.push(['', '', '', this.abbreviateThaiMonth(currDateFull), '', 'ซื้อเพิ่ม', 'จำหน่ายออก', '', this.abbreviateThaiMonth(currDateFull)]);
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
      const currentAmount = account.current;
      const previousAmount = account.previous;
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
      const currentAmount = account.current;
      const previousAmount = account.previous;
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
