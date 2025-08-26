// ============================================================================
// EXPENSES BY NATURE NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { CompanyInfo } from '../../../../types/financial';

/**
 * Generates expenses by nature note with row tracking
 * This is a static template note with predefined expense categories
 */
export class ExpensesByNatureNoteGenerator {
  
  /**
   * Generate expenses by nature note with row tracking
   * This note provides a template structure for expense classification
   */
  static generateWithRowTracking(
    notes: any[][], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    noteNumber: number = 15
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

    console.log(`=== EXPENSES BY NATURE NOTE ROW TRACKING: Starting at row ${tracker.currentRow} ===`);

    // 1. Note Header Row
    notes.push([noteNumber.toString(), 'ค่าใช้จ่ายจำแนกตามธรรมชาติของค่าใช้จ่าย', '', '', '', '', '', '', 'หน่วย:บาท']);
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
    
    // 3. Expense Category Detail Rows (Template Structure)
    const expenseCategories = [
      'การเปลี่ยนแปลงในสินค้าสำเร็จรูปและงานระหว่างทำ',
      'งานที่ทำโดยกิจการและบันทึกเป็นรายจ่ายฝ่ายทุน',
      'วัตถุดิบและวัสดุสิ้นเปลืองใช้ไป',
      'ค่าใช้จ่ายผลประโยชน์พนักงาน',
      'ค่าเสื่อมราคาและค่าตัดจำหน่าย',
      'ค่าใช้จ่ายอื่น'
    ];

    expenseCategories.forEach(category => {
      notes.push(['', '', category, '', '', '', '', '', '']);
      tracker.detailRows.push(tracker.currentRow);
      tracker.currentRow++;
    });
    
    // 4. Total Row
    notes.push(['', '', 'รวม', '', '', '', '', '', '']);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // 5. Spacer Row
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    console.log(`Expenses By Nature Note: Header rows: ${tracker.headerRows}, Year rows: ${tracker.yearHeaderRows}, Detail rows: ${tracker.detailRows}, Total rows: ${tracker.totalRows}`);
    return tracker;
  }
}
