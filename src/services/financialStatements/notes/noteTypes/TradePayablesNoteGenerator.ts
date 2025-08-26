// ============================================================================
// TRADE PAYABLES NOTE GENERATOR
// ============================================================================

import type { DetailedFinancialData, NoteRowTracker } from '../../core/types';
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
    noteNumber: number = 12
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

    const payableAccounts = globalData.individualAccounts.payables;
    const totalAmount = globalData.noteCalculations.payables.total.current;
    const prevTotalAmount = globalData.noteCalculations.payables.total.previous;

    if (totalAmount === 0 && prevTotalAmount === 0 && Object.keys(payableAccounts).length === 0) {
      return tracker;
    }

    console.log(`=== PAYABLES NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
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

    // 3. Detail Rows - Individual accounts (zero-filtering architecture)
    Object.entries(payableAccounts).forEach(([_, accountData]) => {
      notes.push(['', '', accountData.accountName, '', '', '', 
        accountData.current, '', 
        processingType === 'multi-year' ? accountData.previous : '']);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    });

    // 4. Total Row
    notes.push(['', '', 'รวม', '', '', '', totalAmount, '', 
      processingType === 'multi-year' ? prevTotalAmount : '']);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Payables Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
