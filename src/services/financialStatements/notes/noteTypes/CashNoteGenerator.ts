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

    const totalAmount = globalData.noteCalculations.cash.total.current;
    const prevTotalAmount = globalData.noteCalculations.cash.total.previous;

    if (totalAmount === 0 && prevTotalAmount === 0) {
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

    // 3. Detail Rows
    const cashAmount = globalData.noteCalculations.cash.cash.current;
    const bankAmount = globalData.noteCalculations.cash.bankDeposits.current;
    const prevCashAmount = globalData.noteCalculations.cash.cash.previous;
    const prevBankAmount = globalData.noteCalculations.cash.bankDeposits.previous;

    if (cashAmount !== 0 || prevCashAmount !== 0) {
      notes.push(['', '', 'เงินสดในมือ', '', '', '', cashAmount, '', 
        processingType === 'multi-year' ? prevCashAmount : '']);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    }

    if (bankAmount !== 0 || prevBankAmount !== 0) {
      notes.push(['', '', 'เงินฝากธนาคาร', '', '', '', bankAmount, '', 
        processingType === 'multi-year' ? prevBankAmount : '']);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    }

    // 4. Total Row
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', 
      processingType === 'multi-year' ? prevTotalAmount : '']);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Cash Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
