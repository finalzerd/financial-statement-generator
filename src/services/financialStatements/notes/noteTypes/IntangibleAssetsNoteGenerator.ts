// ============================================================================
// INTANGIBLE ASSETS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';
import type { SelectionFirstResult, ClassifiedAccount } from '../../selection/SelectionFirstClassifier';

/**
 * Generates Intangible Assets note with complex structure
 * Handles cost/amortization breakdown with movement tracking
 * Mirrors PPE structure but for intangible assets (computer software, licenses, etc.)
 */
export class IntangibleAssetsNoteGenerator {
  
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
   * Generate Intangible Assets note with enhanced row tracking and Excel formulas
   * Supports both single-year and multi-year processing
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    _trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 12,
    selection?: SelectionFirstResult
  ): NoteRowTracker {
    const tracker: NoteRowTracker = {
      currentRow: notes.length + 1,
      noteStartRow: notes.length + 1,
      headerRows: [],        // Note headers and section headers
      yearHeaderRows: [],    // Year/column headers
      detailRows: [],        // Individual asset/amortization accounts
      totalRows: [],         // Total/รวม rows and net book value
      unitRows: []           // หน่วย:บาท rows
    };

    type IntangibleAssetAccount = {
      accountCode?: string;
      accountName: string;
      current: number;
      previous: number;
    };

    const normalizeSelection = (accounts?: ClassifiedAccount[]): IntangibleAssetAccount[] => {
      if (!accounts) return [];
      return accounts.map((acc) => ({
        accountCode: acc.accountCode,
        accountName: acc.accountName,
        current: acc.current || 0,
        previous: acc.previous || 0
      }));
    };

    const isDecimalAccount = (code?: string) => Boolean(code && code.includes('.'));
    const selectionCostSource = normalizeSelection(selection?.byCategory?.intangible_assets_cost);
    const selectionAmortSource = normalizeSelection(selection?.byCategory?.intangible_assets_accum_amort);

    const selectionCostAccounts = selectionCostSource.filter(acc => !isDecimalAccount(acc.accountCode));
    const spilloverDecimalAccounts = selectionCostSource.filter(acc => isDecimalAccount(acc.accountCode));

    const seenCodes = new Set<string>();
    const addUnique = (list: IntangibleAssetAccount[], target: IntangibleAssetAccount[]) => {
      for (const acc of list) {
        const key = acc.accountCode ?? acc.accountName;
        if (seenCodes.has(key)) continue;
        seenCodes.add(key);
        target.push(acc);
      }
    };

    const selectionAmortAccounts: IntangibleAssetAccount[] = [];
    addUnique(selectionAmortSource, selectionAmortAccounts);
    addUnique(spilloverDecimalAccounts, selectionAmortAccounts);

    const normalizeTrialBalance = (entries: TrialBalanceEntry[], predicate: (entry: TrialBalanceEntry) => boolean): IntangibleAssetAccount[] => {
      return entries
        .filter(predicate)
        .map(entry => {
          const balance = (entry.balance ?? entry.currentBalance ?? 0) as number;
          const previous = (entry.previousBalance ?? 0) as number;
          return {
            accountCode: entry.accountCode,
            accountName: entry.accountName ?? entry.accountCode ?? 'ไม่ระบุ',
            current: Math.abs(balance),
            previous: Math.abs(previous)
          };
        });
    };

    const assetFallback = normalizeTrialBalance(trialBalanceData, entry => {
      const codeStr = entry.accountCode ?? '';
      if (codeStr.includes('.')) return false;
      const code = Number.parseInt(codeStr, 10);
      return Number.isFinite(code) && code >= 1800 && code <= 1859;
    });

    const amortizationFallback = normalizeTrialBalance(trialBalanceData, entry => {
      const codeStr = entry.accountCode ?? '';
      if (!codeStr.includes('.')) return false;
      const base = Math.floor(Number.parseFloat(codeStr));
      return Number.isFinite(base) && base >= 1800 && base <= 1859;
    });

  const usingSelection = selectionCostAccounts.length > 0 || selectionAmortAccounts.length > 0;

    const assetAccounts = (selectionCostAccounts.length > 0 ? selectionCostAccounts : assetFallback)
      .filter(acc => acc.current !== 0 || acc.previous !== 0);
    const amortizationAccounts = (selectionAmortAccounts.length > 0 ? selectionAmortAccounts : amortizationFallback)
      .filter(acc => acc.current !== 0 || acc.previous !== 0);

    console.log(`[Intangible Assets Note] Using ${usingSelection ? 'selection-first' : 'fallback'} data -> assets: ${assetAccounts.length}, amortization: ${amortizationAccounts.length}`);

    if (assetAccounts.length === 0 && amortizationAccounts.length === 0) {
      return tracker;
    }

    console.log(`=== INTANGIBLE ASSETS NOTE ENHANCED ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'สินทรัพย์ไม่มีตัวตน', '', '', '', '', '', '', 'หน่วย:บาท']);
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
    
    // 7. Accumulated Amortization Section Header
    notes.push(['', '', 'ค่าตัดจำหน่ายสะสม', '', '', '', '', '', '']);
    tracker.headerRows.push(tracker.currentRow); // Section header
    tracker.currentRow++;
    
    let amortizationTotalCurrent = 0;
    let amortizationTotalPrevious = 0;
    let amortizationExpense = 0;
    let amortizationDisposal = 0;
    const amortizationStartRow = tracker.currentRow; // Track start of amortization details

    // 8. Individual Amortization Accounts (Detail Rows)
    amortizationAccounts.forEach(account => {
      const currentAmount = account.current;
      const previousAmount = account.previous;
      const expenseAmount = Math.max(0, currentAmount - previousAmount); // Amortization expense for the year
      const disposalAmount = Math.max(0, previousAmount - currentAmount); // Amortization disposal for the year
      
      if (currentAmount !== 0 || previousAmount !== 0) {
        // Use same structure for both single and multi-year processing
        const expenseValue = expenseAmount > 0 ? expenseAmount : 0;
        const disposalAmountValue = disposalAmount > 0 ? disposalAmount : 0;
        
        notes.push(['', '', account.accountName, 
          processingType === 'multi-year' ? previousAmount : '', '', // Column D: Previous amortization (blank for single-year)
          expenseValue, // Column F: Amortization expense (show 0 if no expense)
          disposalAmountValue, // Column G: Amortization disposal (show 0 if no disposal)
          '', currentAmount]); // Column I: Current amortization
        
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
        
        amortizationTotalCurrent += currentAmount;
        amortizationTotalPrevious += previousAmount;
        amortizationExpense += expenseAmount;
        amortizationDisposal += disposalAmount;
      }
    });

    // 9. Amortization Totals with Excel formulas (Total Row) - Same structure for both processing types
    const amortizationEndRow = tracker.currentRow - 1; // Last row of amortization details
    
    notes.push(['', '', 'รวม', 
      processingType === 'multi-year' ? { f: `SUM(D${amortizationStartRow}:D${amortizationEndRow})` } : '', '', // Column D: Previous amortization total formula (blank for single-year)
      { f: `SUM(F${amortizationStartRow}:F${amortizationEndRow})` }, // Column F: Total amortization expense formula
      { f: `SUM(G${amortizationStartRow}:G${amortizationEndRow})` }, // Column G: Total amortization disposal formula
      '', { f: `SUM(I${amortizationStartRow}:I${amortizationEndRow})` }]); // Column I: Current amortization total formula
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 10. Net book value with formulas referencing the totals above (Total Row)
    const assetTotalRowIndex = assetEndRow + 1; // Row number of asset totals
    const amortizationTotalRowIndex = tracker.currentRow - 1; // Row number of amortization totals (just added above)
    
    // Net book value - Same structure for both processing types
    notes.push(['', '', 'มูลค่าสุทธิ', 
      processingType === 'multi-year' ? { f: `D${assetTotalRowIndex}-D${amortizationTotalRowIndex}` } : '', '', '', '', // Column D: Previous net value formula (blank for single-year)
      '', { f: `I${assetTotalRowIndex}-I${amortizationTotalRowIndex}` }]); // Column I: Current net value formula
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 11. Amortization expense summary - reference the amortization expense total
    if (amortizationExpense > 0) {
      notes.push(['', '', 'ค่าตัดจำหน่าย', '', '', '', '', '', 
        { f: `F${amortizationTotalRowIndex}` }]); // Column I (index 8): Reference amortization expense total
      tracker.totalRows.push(tracker.currentRow);
      tracker.currentRow++;
    }
    
    // 12. Final spacer
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Intangible Assets Enhanced Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
