// ============================================================================
// CASH NOTE GENERATOR
// ============================================================================

import type { DetailedFinancialData, NoteRowTracker } from '../../core/types';
import type { CompanyInfo } from '../../../../types/financial';

/**
 * Generates cash and cash equivalents note (Note 7) with row tracking
 * for precise Excel formatting
 */
export class CashNoteGenerator {
  
  /**
   * Generate cash note with row tracking architecture
   * This method tracks exact row positions for ExcelJS formatter
   */
  static generateWithRowTracking(
    notes: any[][], 
    globalData: DetailedFinancialData,
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 3
  ): NoteRowTracker {
    const tracker: NoteRowTracker = {
      currentRow: notes.length + 1, // Excel 1-indexed
      noteStartRow: notes.length + 1,
      headerRows: [],
      yearHeaderRows: [],
      detailRows: [],
      totalRows: [],
      unitRows: []
    };

    // Mapping-first: use individual selections for detail population
    const cashAccounts = globalData.individualAccounts.cash;
    const totalAmount = globalData.noteCalculations.cash.total.current;
    const prevTotalAmount = globalData.noteCalculations.cash.total.previous;

    // If both totals are zero AND no mapped accounts, skip the note
    if (totalAmount === 0 && prevTotalAmount === 0 && Object.keys(cashAccounts).length === 0) {
      return tracker; // No note generated
    }

    console.log(`=== CASH NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row (note number + title + "หน่วย:บาท")
    notes.push([noteNumber.toString(), 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', '', '', '', 'หน่วย:บาท']);
    tracker.headerRows.push(tracker.currentRow);
    tracker.unitRows.push(tracker.currentRow); // "หน่วย:บาท" is also in this row
    tracker.currentRow++;

    // 2. Year Header Row  
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    tracker.yearHeaderRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 3. Detail Rows - Mapping-based individual accounts (no artificial grouping)
    const sortedCashEntries = Object.entries(cashAccounts).sort((a, b) => {
      const aCode = parseInt(a[0] || '0', 10);
      const bCode = parseInt(b[0] || '0', 10);
      return aCode - bCode;
    });
    sortedCashEntries.forEach(([_, accountData]) => {
      notes.push(['', '', accountData.accountName, '', '', '',
        accountData.current, '',
        processingType === 'multi-year' ? accountData.previous : ''
      ]);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    });

    // 4. Total Row - Use SUM formulas over detail rows when available; fallback to numeric totals otherwise
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

    console.log(`Cash Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
