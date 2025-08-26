// ============================================================================
// FINANCIAL APPROVAL NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { CompanyInfo } from '../../../../types/financial';

/**
 * Generates financial approval note with row tracking
 * This is a static note confirming board approval of financial statements
 */
export class FinancialApprovalNoteGenerator {
  
  /**
   * Generate financial approval note with row tracking
   * This note provides standard approval confirmation text
   */
  static generateWithRowTracking(
    notes: any[][], 
    companyInfo: CompanyInfo, 
    noteNumber: number = 16
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

    console.log(`=== FINANCIAL APPROVAL NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'การอนุมัติงบการเงิน', '', '', '', '', '', '', '']);
    tracker.headerRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 2. Approval Statement Row
    notes.push(['', '', 'งบการเงินนี้ได้การรับอนุมัติให้ออกงบการเงินโดยคณะกรรมการผู้มีอำนาจของบริษัทแล้ว', '', '', '', '', '', '']);
    tracker.detailRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 3. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Financial Approval Note: Header rows: ${tracker.headerRows}, Detail rows: ${tracker.detailRows}`);
    return tracker;
  }
}
