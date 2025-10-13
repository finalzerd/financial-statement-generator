// ============================================================================
// OTHER INCOME NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';
import type { SelectionFirstResult, ClassifiedAccount } from '../../selection/SelectionFirstClassifier';

/**
 * Generates other income note (Note 14) with row tracking
 * for precise Excel formatting
 */
export class OtherIncomeNoteGenerator {
  
  /**
   * Generate other income note with row tracking architecture
   * Uses account range 4110-4999 for other income classification
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    _trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 14,
    selection?: SelectionFirstResult
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

    type OtherIncomeAccount = {
      accountCode?: string;
      accountName: string;
      current: number;
      previous: number;
    };

    const normalizeSelection = (accounts?: ClassifiedAccount[]): OtherIncomeAccount[] => {
      if (!accounts) return [];
      return accounts.map(acc => ({
        accountCode: acc.accountCode,
        accountName: acc.accountName,
        current: acc.current ?? 0,
        previous: acc.previous ?? 0
      }));
    };

    const selectionAccounts = normalizeSelection(selection?.byCategory?.other_income);

    const fallbackAccounts: OtherIncomeAccount[] = trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0', 10);
        return code >= 4110 && code <= 4999;
      })
      .map(entry => {
        const rawCurrent = (entry.balance ?? (entry.creditAmount ?? 0) - (entry.debitAmount ?? 0)) as number;
        const current = Math.abs(rawCurrent);
        const previous = Math.abs((entry.previousBalance ?? 0) as number);
        return {
          accountCode: entry.accountCode,
          accountName: entry.accountName ?? entry.accountCode ?? 'ไม่ระบุ',
          current,
          previous
        };
      })
      .filter(account => account.current !== 0 || account.previous !== 0);

    const accounts = (selectionAccounts.length > 0 ? selectionAccounts : fallbackAccounts)
      .filter(account => account.current !== 0 || account.previous !== 0);

    if (accounts.length === 0) {
      return tracker;
    }

    const usingSelection = selectionAccounts.length > 0;
    console.log(`[Other Income Note] Using ${usingSelection ? 'selection-first' : 'fallback'} data -> accounts: ${accounts.length}`);
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

    // 3. Detail Rows
    const detailStartRow = tracker.currentRow;
    accounts.forEach(account => {
      notes.push(['', '', account.accountName, '', '', '', account.current, '', 
        processingType === 'multi-year' ? account.previous : '']);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    });

    const detailEndRow = tracker.currentRow - 1;

    // 4. Total Row (if more than one item)
    if (accounts.length > 1) {
      notes.push(['', '', 'รวม', '', '', '', 
        { f: `SUM(G${detailStartRow}:G${detailEndRow})` }, '', 
        processingType === 'multi-year' ? { f: `SUM(I${detailStartRow}:I${detailEndRow})` } : ''
      ]);
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
