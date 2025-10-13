// ============================================================================
// ASSET LONG TERM LOANS GIVEN NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates asset-side long-term loans given note with row tracking
 * Default mapping uses account 1710; supports selection-first details.
 */
export class AssetLongTermLoansNoteGenerator {
  private static sumAccountsByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    return trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0', 10);
        return code >= startCode && code <= endCode;
      })
      .reduce((sum, entry) => sum + (entry.balance || entry.currentBalance || 0), 0);
  }

  private static sumPreviousBalanceByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    return trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0', 10);
        return code >= startCode && code <= endCode;
      })
      .reduce((sum, entry) => sum + (entry.previousBalance || 0), 0);
  }

  static generateWithRowTracking(
    notes: any[][],
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    trialBalancePrevious?: TrialBalanceEntry[],
    noteNumber: number = 0,
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

    const selRows = selection?.byCategory?.asset_long_term_loans ?? [];
    if (selRows.length > 0) {
      const detailRows = selRows.filter(a => !(isZeroLike(a.current) && (processingType === 'single-year' ? true : isZeroLike(a.previous))));
      const totalCurrent = detailRows.reduce((s, a) => s + (a.current || 0), 0);
      const totalPrevious = detailRows.reduce((s, a) => s + (a.previous || 0), 0);
      const totalsZero = processingType === 'multi-year' ? (totalCurrent === 0 && totalPrevious === 0) : (totalCurrent === 0);
      if (detailRows.length === 0 || totalsZero) return tracker;

      notes.push([noteNumber ? noteNumber.toString() : '', 'เงินให้กู้ยืมระยะยาว', '', '', '', '', '', '', 'หน่วย:บาท']);
      tracker.headerRows.push(tracker.currentRow); tracker.unitRows.push(tracker.currentRow); tracker.currentRow++;
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      tracker.yearHeaderRows.push(tracker.currentRow); tracker.currentRow++;
      detailRows.forEach(a => {
        notes.push(['', '', a.accountName, '', '', '', a.current, '', processingType === 'multi-year' ? a.previous : '']);
        tracker.detailRows.push(tracker.currentRow); tracker.currentRow++;
      });
      const first = tracker.detailRows[0];
      const last = tracker.detailRows[tracker.detailRows.length - 1];
      notes.push(['', '', 'รวม', '', '', '', { f: `SUM(G${first}:G${last})` } as any, '', processingType === 'multi-year' ? ({ f: `SUM(I${first}:I${last})` } as any) : '']);
      tracker.totalRows.push(tracker.currentRow); tracker.currentRow++;
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;
      return tracker;
    }

    // Legacy fallback: single-line by numeric account
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1710, 1710));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious
      ? Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 1710, 1710))
      : 0;
    const totalsZero = processingType === 'multi-year' ? (totalAmount === 0 && prevTotalAmount === 0) : (totalAmount === 0);
    if (totalsZero) return tracker;

    notes.push([noteNumber ? noteNumber.toString() : '', 'เงินให้กู้ยืมระยะยาว', '', '', '', '', '', '', 'หน่วย:บาท']);
    tracker.headerRows.push(tracker.currentRow); tracker.unitRows.push(tracker.currentRow); tracker.currentRow++;
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    tracker.yearHeaderRows.push(tracker.currentRow); tracker.currentRow++;
    notes.push(['', '', 'เงินให้กู้ยืมระยะยาว', '', '', '', totalAmount, '', processingType === 'multi-year' ? prevTotalAmount : '']);
    tracker.detailRows.push(tracker.currentRow); const dataRow = tracker.currentRow; tracker.currentRow++;
    notes.push(['', '', 'รวม', '', '', '', { f: `G${dataRow}` }, '', processingType === 'multi-year' ? { f: `I${dataRow}` } : '']);
    tracker.totalRows.push(tracker.currentRow); tracker.currentRow++;
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;
    return tracker;
  }
}
