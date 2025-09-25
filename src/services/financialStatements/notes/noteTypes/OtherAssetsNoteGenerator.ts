// ============================================================================
// OTHER ASSETS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

/**
 * Generates other assets note with row tracking
 * Covers assets account range 1660-1700
 */
export class OtherAssetsNoteGenerator {
  
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
   * Generate other assets note with row tracking
   * Uses account range 1660-1700 for miscellaneous other assets
   */
  static generateWithRowTracking(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 7,
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

    const selRows = selection?.byCategory?.other_assets ?? [];
    if (selRows.length > 0) {
      console.log(`[SelectionFirst] Other assets: using selection-first details (${selRows.length} accounts). Total current=${selRows.reduce((s, a) => s + (a.current || 0), 0)}, previous=${selRows.reduce((s, a) => s + (a.previous || 0), 0)}`);

      // 1. Note Header Row
      notes.push([noteNumber.toString(), 'สินทรัพย์อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
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

      // 3. Detail Rows from selection
      selRows.forEach(a => {
        notes.push(['', '', a.accountName, '', '', '', a.current, '', processingType === 'multi-year' ? a.previous : '']);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });

      // 4. Total Row with SUM formulas
      const first = tracker.detailRows[0];
      const last = tracker.detailRows[tracker.detailRows.length - 1];
      notes.push(['', '', 'รวม', '', '', '', { f: `SUM(G${first}:G${last})` } as any, '', processingType === 'multi-year' ? ({ f: `SUM(I${first}:I${last})` } as any) : '']);
      tracker.totalRows.push(tracker.currentRow);
      tracker.currentRow++;

      // 5. Spacer Row
      notes.push(['', '', '', '', '', '', '', '', '']);
      tracker.currentRow++;
      return tracker;
    }

    // Fallback to legacy numeric-range-based generation
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1660, 1700));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 1660, 1700)) : 0;

    if (totalAmount === 0 && prevTotalAmount === 0) {
      return tracker;
    }
    console.log('[SelectionFirst] Other assets: no selection details; falling back to legacy range totals');

    console.log(`=== OTHER ASSETS NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'สินทรัพย์อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
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
    notes.push(['', '', 'สินทรัพย์อื่น', '', '', '', totalAmount, '', 
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

    console.log(`Other Assets Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
