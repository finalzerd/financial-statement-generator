// ============================================================================
// BANK OVERDRAFTS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates bank overdrafts and short-term borrowings from financial institutions note
 * Covers bank overdrafts and short-term loans from financial institutions (accounts 2001-2009)
 */
export class BankOverdraftsNoteGenerator {
  
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
   * Generate bank overdrafts note with row tracking
   * Uses accounts 2001-2009 for bank overdrafts and short-term borrowings from financial institutions
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 15,
    selection?: SelectionFirstResult
  ): NoteRowTracker {
    const isZeroLike = (value: unknown) => (
      value === null ||
      value === undefined ||
      (typeof value === 'number' && value === 0) ||
      (typeof value === 'string' && value.trim() === '')
    );
    const tracker: NoteRowTracker = {
      currentRow: notes.length + 1,
      noteStartRow: notes.length + 1,
      headerRows: [],
      yearHeaderRows: [],
      detailRows: [],
      totalRows: [],
      unitRows: []
    };

    // Try selection-first approach
    const selectionRows = selection?.byCategory?.bank_overdrafts ?? [];
    if (selectionRows.length > 0) {
      let suppressed = 0;
      const detailRows = selectionRows.filter(account => {
        const hide = isZeroLike(account.current) && (processingType === 'single-year' ? true : isZeroLike(account.previous));
        if (hide) suppressed++;
        return !hide;
      });

      const totalCurrent = detailRows.reduce((sum, account) => sum + (account.current ?? 0), 0);
      const totalPrevious = detailRows.reduce((sum, account) => sum + (account.previous ?? 0), 0);
      const totalsAreZero = processingType === 'multi-year'
        ? totalCurrent === 0 && totalPrevious === 0
        : totalCurrent === 0;

      if (detailRows.length === 0 || totalsAreZero) {
        console.log('[SelectionFirst] Bank overdrafts: skipping note - totals zero after filtering or no detail rows.');
        return tracker;
      }

      console.log(`[SelectionFirst] Bank overdrafts: using selection-first details (${detailRows.length} accounts after suppressing ${suppressed}). Total current=${totalCurrent}, previous=${totalPrevious}`);

      // Note header
      notes.push([noteNumber.toString(), 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน', '', '', '', '', '', '', 'หน่วย:บาท']);
      tracker.headerRows.push(tracker.currentRow);
      tracker.unitRows.push(tracker.currentRow);
      tracker.currentRow++;

      // Year headers
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      tracker.yearHeaderRows.push(tracker.currentRow);
      tracker.currentRow++;

      // Detail rows
      detailRows.forEach(account => {
        notes.push(['', '', account.accountName, '', '', '', account.current ?? 0, '', processingType === 'multi-year' ? account.previous ?? 0 : '']);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });

      // Total row with SUM formulas
      const firstDetailRow = tracker.detailRows[0];
      const lastDetailRow = tracker.detailRows[tracker.detailRows.length - 1];
      notes.push(['', '', 'รวม', '', '', '', { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` } as any, '', processingType === 'multi-year' ? ({ f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } as any) : '']);
      tracker.totalRows.push(tracker.currentRow);
      tracker.currentRow++;

      // Spacer
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;

      return tracker;
    }

    // No selection data - skip note to preserve classifier integrity
    console.log('[BankOverdrafts] No selection data; skipping note (no fallback to preserve classifier integrity)');
    return tracker;
  }
}