// ============================================================================
// LONG TERM LOANS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates long term loans note with row tracking
 * Covers loans from financial institutions (accounts 2120-2123, excluding 2121)
 */
export class LongTermLoansNoteGenerator {
  
  /**
   * Calculate sum of accounts in numeric range (helper method)
   * TODO: Extract to shared utility when refactoring is complete
   */
  private static sumAccountsByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    return trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0');
        return code >= startCode && code <= endCode;
      })
      .reduce((sum, entry) => sum + (entry.balance || entry.currentBalance || 0), 0);
  }

  /**
   * Calculate sum of previous balances in numeric range (helper method)
   * TODO: Extract to shared utility when refactoring is complete
   */
  private static sumPreviousBalanceByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    return trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0');
        return code >= startCode && code <= endCode;
      })
      .reduce((sum, entry) => sum + (entry.previousBalance || 0), 0);
  }

  /**
   * Generate long term loans note with row tracking
   * Uses accounts 2120-2123 but excludes 2121 (special calculation)
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 11
  ): NoteRowTracker {
    const tracker: NoteRowTracker = {
      currentRow: notes.length + 1,
      noteStartRow: notes.length + 1,
      headerRows: [],
      yearHeaderRows: [],
      detailRows: [],
      totalRows: [],
      unitRows: []
    };

    // Special calculation: 2120-2123 minus 2121
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) - 
                       Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2121, 2121));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2120, 2123)) : 0;

    if (totalAmount === 0 && prevTotalAmount === 0) {
      return tracker;
    }

    console.log(`=== LONG TERM LOANS NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน', '', '', '', '', '', '', 'หน่วย:บาท']);
    tracker.headerRows.push(tracker.currentRow);
    tracker.unitRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 2. Year Header Row
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    tracker.yearHeaderRows.push(tracker.currentRow);
    tracker.currentRow++;
    
    // 3. Main Detail Row
    notes.push(['', '', 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน', '', '', '', totalAmount, '', 
      processingType === 'multi-year' ? prevTotalAmount : '']);
    tracker.detailRows.push(tracker.currentRow);
    const mainDataRowIndex = tracker.currentRow; // Store for formula reference
    tracker.currentRow++;
    
    // 4. Total Row with Excel formulas
    notes.push(['', '', 'รวม', '', '', '', 
      { f: `G${mainDataRowIndex}` }, '', // Reference main detail row current amount
      processingType === 'multi-year' ? { f: `I${mainDataRowIndex}` } : '']); // Reference main detail row previous amount
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Current Portion Detail - Calculate 10% as placeholder
    const currentPortion = Math.round(totalAmount * 0.1);
    notes.push(['', '', 'หัก ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี', '', '', '', currentPortion, '', '']);
    tracker.detailRows.push(tracker.currentRow);
    tracker.currentRow++;
    
    // 6. Net Long-term Loans Row
    notes.push(['', '', 'เงินกู้ยืมระยะยาวสุทธิจากส่วนที่ถึงกำหนดชำระคืนภายในหนึ่งปี', '', '', '', totalAmount - currentPortion, '', '']);
    tracker.detailRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 7. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Long Term Loans Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
