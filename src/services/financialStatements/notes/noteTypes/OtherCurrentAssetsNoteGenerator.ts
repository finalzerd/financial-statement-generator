// ============================================================================
// OTHER CURRENT ASSETS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates other current assets note with row tracking
 */
export class OtherCurrentAssetsNoteGenerator {
  
  /**
   * Generate other current assets note with row tracking
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 11,
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

    const selRows = selection?.byCategory?.other_current_assets ?? [];
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
        console.log('[SelectionFirst] Other current assets: skipping note - totals zero after filtering or no detail rows.');
        return tracker;
      }

      console.log(`[SelectionFirst] Other current assets: using selection-first details (${detailRows.length} accounts after suppressing ${suppressed}). Total current=${totalCurrent}, previous=${totalPrevious}`);

      // 1. Note Header Row
      notes.push([noteNumber.toString(), 'สินทรัพย์หมุนเวียนอื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      tracker.headerRows.push(tracker.currentRow);
      tracker.unitRows.push(tracker.currentRow);
      tracker.currentRow++;

      // 2. Year Header Row
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', companyInfo.reportingYear.toString(), '', (companyInfo.reportingYear - 1).toString()]);
        tracker.yearHeaderRows.push(tracker.currentRow);
        tracker.currentRow++;
      } else {
        notes.push(['', '', '', '', '', '', companyInfo.reportingYear.toString(), '', '']);
        tracker.yearHeaderRows.push(tracker.currentRow);
        tracker.currentRow++;
      }

      // 3. Detail Rows
      const startDetailRow = tracker.currentRow;
      detailRows.forEach(account => {
        if (processingType === 'multi-year') {
          notes.push(['', '', account.accountName, '', '', '', account.current, '', account.previous]);
        } else {
          notes.push(['', '', account.accountName, '', '', '', account.current, '', '']);
        }
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });
      const endDetailRow = tracker.currentRow - 1;

      // 4. Total Row
      const totalFormulaCurrent = `SUM(G${startDetailRow}:G${endDetailRow})`;
      const totalFormulaPrevious = `SUM(I${startDetailRow}:I${endDetailRow})`;

      if (processingType === 'multi-year') {
        notes.push(['', '', 'รวม', '', '', '', { f: totalFormulaCurrent }, '', { f: totalFormulaPrevious }]);
      } else {
        notes.push(['', '', 'รวม', '', '', '', { f: totalFormulaCurrent }, '', '']);
      }
      tracker.totalRows.push(tracker.currentRow);
      tracker.currentRow++;

      // 5. Spacer Row
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;

      return tracker;
    }

    // Fallback logic removed as per "default account code range should not be set yet"
    // If no selection, we return empty tracker (note skipped)
    console.log('[SelectionFirst] Other current assets: no selection found, skipping note.');
    return tracker;
  }
}
