// ============================================================================
// TRADE PAYABLES NOTE GENERATOR
// ============================================================================

import type { DetailedFinancialData, NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { CompanyInfo } from '../../../../types/financial';

/**
 * Generates trade and other payables note (Note 12) with row tracking
 * for precise Excel formatting
 */
export class TradePayablesNoteGenerator {
  
  /**
   * Generate trade payables note with row tracking architecture
   * Uses dynamic individual accounts (no artificial grouping)
   */
  static generateWithRowTracking(
    notes: any[][], 
    globalData: DetailedFinancialData,
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
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

    const selectionRows = selection?.byCategory?.payables ?? [];
    const hasSelectionDetails = selectionRows.length > 0;

    const payableAccounts = globalData.individualAccounts.payables;
    const totalAmount = globalData.noteCalculations.payables.total.current;
    const prevTotalAmount = globalData.noteCalculations.payables.total.previous;

    const hasMappedAccounts = Object.keys(payableAccounts).length > 0;
    const totalsAreZero = processingType === 'multi-year'
      ? totalAmount === 0 && prevTotalAmount === 0
      : totalAmount === 0;

    if (totalsAreZero || (!hasSelectionDetails && !hasMappedAccounts)) {
      console.log('[SelectionFirst] Payables: skipping note - totals zero or no mapped accounts/selection.');
      return tracker;
    }

    console.log(`=== PAYABLES NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
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

    // 3. Detail Rows - Prefer selection rows; otherwise use globalData individual accounts
    if (hasSelectionDetails) {
      console.log(`[SelectionFirst] Payables: using selection-first details (${selectionRows.length} accounts). Total current=${selection?.totals?.payables?.current ?? 'n/a'}, previous=${selection?.totals?.payables?.previous ?? 'n/a'}`);
      let suppressed = 0;
      selectionRows.forEach((a) => {
        const curr = a.current;
        const prev = a.previous;
        const hide = isZeroLike(curr) && (processingType === 'single-year' ? true : isZeroLike(prev));
        if (hide) { suppressed++; return; }
        notes.push(['', '', a.accountName, '', '', '', 
          a.current, '', 
          processingType === 'multi-year' ? a.previous : '']);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });
      if (suppressed > 0) {
        console.log(`[SelectionFirst] Payables: suppressed ${suppressed} zero/blank detail rows.`);
      }
    } else {
      console.log(`[SelectionFirst] Payables: no selection details; falling back to globalData.individualAccounts (${Object.keys(payableAccounts).length} accounts).`);
      let suppressed = 0;
      Object.entries(payableAccounts).forEach(([_, accountData]) => {
        const curr = accountData.current;
        const prev = accountData.previous;
        const hide = isZeroLike(curr) && (processingType === 'single-year' ? true : isZeroLike(prev));
        if (hide) { suppressed++; return; }
        notes.push(['', '', accountData.accountName, '', '', '', 
          accountData.current, '', 
          processingType === 'multi-year' ? accountData.previous : '']);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });
      if (suppressed > 0) {
        console.log(`[Legacy] Payables: suppressed ${suppressed} zero/blank detail rows from globalData.`);
      }
    }

    // 4. Total Row
    // Use Excel SUM formulas over detail rows when available; fallback to numeric totals otherwise
    const hasDetails = tracker.detailRows.length > 0;
    const firstDetailRow = hasDetails ? tracker.detailRows[0] : null;
    const lastDetailRow = hasDetails ? tracker.detailRows[tracker.detailRows.length - 1] : null;

    const currentTotalCell = hasDetails
      ? { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` }
      : totalAmount;
    const previousTotalCell = processingType === 'multi-year'
      ? (hasDetails ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } : prevTotalAmount)
      : '';

    notes.push(['', '', 'รวม', '', '', '', currentTotalCell as any, '', previousTotalCell as any]);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Payables Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
