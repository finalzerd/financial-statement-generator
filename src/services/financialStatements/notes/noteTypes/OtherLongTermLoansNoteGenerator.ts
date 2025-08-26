// ============================================================================
// OTHER LONG TERM LOANS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates other long term loans note with row tracking
 * Covers non-financial institution long-term loans (accounts 2050-2052)
 */
export class OtherLongTermLoansNoteGenerator {
  
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
   * Generate other long term loans note with row tracking
   * Uses accounts 2050-2052 for general long-term borrowings
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 12
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

    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2050, 2052)) : 0;

    if (totalAmount === 0 && prevTotalAmount === 0) {
      return tracker;
    }

    console.log(`=== OTHER LONG TERM LOANS NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาว', '', '', '', '', '', '', 'หน่วย:บาท']);
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
    
    // 3. Detail Row
    notes.push(['', '', 'เงินกู้ยืมระยะยาว', '', '', '', totalAmount, '', 
      processingType === 'multi-year' ? prevTotalAmount : '']);
    tracker.detailRows.push(tracker.currentRow);
    const dataRowIndex = tracker.currentRow; // Store for formula reference
    tracker.currentRow++;
    
    // 4. Total Row with Excel formulas
    notes.push(['', '', 'รวม', '', '', '', 
      { f: `G${dataRowIndex}` }, '', // Reference detail row current amount
      processingType === 'multi-year' ? { f: `I${dataRowIndex}` } : '']); // Reference detail row previous amount
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Other Long Term Loans Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
