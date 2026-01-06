// ============================================================================
// LIABILITY SHORT TERM LOANS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates liability-side short term loans note with row tracking
 * Default numeric coverage: account 2030 (เงินกู้ยืมระยะสั้น - borrowed short-term loans)
 */
export class LiabilityShortTermLoansNoteGenerator {
  
  /**
   * Generate liability short term loans note with row tracking
   * Uses category 'short_term_loans' for loans BORROWED (liability-side)
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

    // Use selection-first data for liability-side short-term loans borrowed
    // Category key must match SelectionFirstClassifier (short_term_loans)
    const selRows = selection?.byCategory?.short_term_loans ?? [];
    if (selRows.length > 0) {
      let suppressed = 0;
      const detailRows = selRows.filter((a) => {
        const curr = a.current;
        const prev = a.previous;
        const hide = isZeroLike(curr) && (processingType === 'single-year' ? true : isZeroLike(prev));
        if (hide) { suppressed++; }
        return !hide;
      });

      const totalCurrent = detailRows.reduce((s, a) => s + (a.current || 0), 0);
      const totalPrevious = detailRows.reduce((s, a) => s + (a.previous || 0), 0);
      const totalsAreZero = processingType === 'multi-year'
        ? totalCurrent === 0 && totalPrevious === 0
        : totalCurrent === 0;

      if (detailRows.length === 0 || totalsAreZero) {
        console.log('[SelectionFirst] Liability short-term loans: skipping note - totals zero after filtering or no detail rows.');
        return tracker;
      }

      console.log(`[SelectionFirst] Liability short-term loans: using selection-first details (${detailRows.length} accounts after suppressing ${suppressed}). Total current=${totalCurrent}, previous=${totalPrevious}`);
      
      // 1. Header
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
      tracker.headerRows.push(tracker.currentRow); 
      tracker.unitRows.push(tracker.currentRow); 
      tracker.currentRow++;
      
      // 2. Year header
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      tracker.yearHeaderRows.push(tracker.currentRow); 
      tracker.currentRow++;
      
      // 3. Details
      detailRows.forEach(a => {
        notes.push(['', '', a.accountName, '', '', '', a.current, '', processingType === 'multi-year' ? a.previous : '']);
        tracker.detailRows.push(tracker.currentRow); 
        tracker.currentRow++;
      });
      
      // 4. Total via SUM formulas
      const first = tracker.detailRows[0];
      const last = tracker.detailRows[tracker.detailRows.length - 1];
      notes.push(['', '', 'รวม', '', '', '', 
        { f: `SUM(G${first}:G${last})` } as any, '', 
        processingType === 'multi-year' ? ({ f: `SUM(I${first}:I${last})` } as any) : '']);
      tracker.totalRows.push(tracker.currentRow); 
      tracker.currentRow++;
      
      // 5. Spacer
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;
      
      return tracker;
    }

    // No selection data - skip note to preserve classifier integrity
    console.log('[LiabilityShortTermLoans] No selection data; skipping note (no fallback to preserve classifier integrity)');
    return tracker;
  }
}
