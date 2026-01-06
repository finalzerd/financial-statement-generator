// ============================================================================
// OTHER LONG TERM LOANS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
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
    noteNumber: number = 12,
    selection?: SelectionFirstResult
  ): NoteRowTracker {
    const isZeroLike = (v: any) => (
      v === null || v === undefined ||
      (typeof v === 'number' && v === 0) ||
      (typeof v === 'string' && v.trim() === '')
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

    const selRows = selection?.byCategory?.long_term_loans_other ?? [];
    if (selRows.length > 0) {
      let suppressed = 0;
      const detailRows = selRows.filter((a) => {
        const curr = a.current;
        const prev = a.previous;
        const hide = isZeroLike(curr) && (processingType === 'single-year' ? true : isZeroLike(prev));
        if (hide) { suppressed++; }
        return !hide;
      });

      const totalAmount = detailRows.reduce((s, a) => s + (a.current || 0), 0);
      const prevTotalAmount = detailRows.reduce((s, a) => s + (a.previous || 0), 0);
      const totalsAreZero = processingType === 'multi-year'
        ? totalAmount === 0 && prevTotalAmount === 0
        : totalAmount === 0;

      if (detailRows.length === 0 || totalsAreZero) {
        console.log('[SelectionFirst] LT loans (Other): skipping note - totals zero after filtering or no detail rows.');
        return tracker;
      }

      console.log(`[SelectionFirst] LT loans (Other): using selection-first details (${detailRows.length} accounts after suppressing ${suppressed}). Total current=${totalAmount}, previous=${prevTotalAmount}`);
      // Header
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาว', '', '', '', '', '', '', 'หน่วย:บาท']);
      tracker.headerRows.push(tracker.currentRow); tracker.unitRows.push(tracker.currentRow); tracker.currentRow++;
      // Year header
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      tracker.yearHeaderRows.push(tracker.currentRow); tracker.currentRow++;
      // Details
      detailRows.forEach(a => {
        notes.push(['', '', a.accountName, '', '', '', a.current, '', processingType === 'multi-year' ? a.previous : '']);
        tracker.detailRows.push(tracker.currentRow); tracker.currentRow++;
      });
      // Total via SUM
      const first = tracker.detailRows[0];
      const last = tracker.detailRows[tracker.detailRows.length - 1];
      notes.push(['', '', 'รวม', '', '', '', { f: `SUM(G${first}:G${last})` } as any, '', processingType === 'multi-year' ? ({ f: `SUM(I${first}:I${last})` } as any) : '']);
      tracker.totalRows.push(tracker.currentRow); tracker.currentRow++;
      // Spacer
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;
      return tracker;
    }

    // No selection data and no fallback - skip note entirely
    // DO NOT filter trial balance directly as it bypasses classifier rules
    console.log('[SelectionFirst] LT loans (Other): no selection details; skipping note (no fallback to preserve classifier integrity).');
    return tracker;
  }
}
