// ============================================================================
// OTHER INCOME NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates other income note (Note 14) with row tracking
 * for precise Excel formatting
 */
export class OtherIncomeNoteGenerator {
  
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
      .reduce((sum, entry) => {
        const balance = entry.creditAmount - entry.debitAmount; // P&L accounts: credit - debit
        return sum + balance;
      }, 0);
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
   * Generate other income note with row tracking architecture
   * Uses account range 4110-4999 for other income classification
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 14
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

    // Calculate other income amounts (account range 4110-4999)
    const currentAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4110, 4999));
    const previousAmount = Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 4110, 4999));

    if (currentAmount === 0 && previousAmount === 0) {
      return tracker;
    }

    console.log(`=== OTHER INCOME NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'รายได้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
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

    // 3. Detail Rows - Get individual other income accounts
    const otherIncomeAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= 4110 && code <= 4999 && (Math.abs(entry.debitAmount - entry.creditAmount) !== 0 || Math.abs(entry.previousBalance || 0) !== 0);
    });

    otherIncomeAccounts.forEach(account => {
      const current = Math.abs(account.creditAmount - account.debitAmount);
      const previous = Math.abs(account.previousBalance || 0);
      
      if (current !== 0 || previous !== 0) {
        notes.push(['', '', account.accountName, '', '', '', current, '', 
          processingType === 'multi-year' ? previous : '']);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      }
    });

    // 4. Total Row (if more than one item)
    if (otherIncomeAccounts.length > 1) {
      notes.push(['', '', 'รวม', '', '', '', currentAmount, '', 
        processingType === 'multi-year' ? previousAmount : '']);
      tracker.totalRows.push(tracker.currentRow);
      tracker.currentRow++;
    }

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Other Income Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
