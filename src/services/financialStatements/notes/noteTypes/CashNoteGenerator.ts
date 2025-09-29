// ============================================================================
// CASH NOTE GENERATOR
// ============================================================================

import type { DetailedFinancialData, NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
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
    noteNumber: number = 3,
    selection?: SelectionFirstResult
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

    // Prefer selection-first for detail population when available; fallback to legacy globalData
    const selectionRows = selection?.byCategory?.cash ?? [];
    const hasSelectionDetails = selectionRows.length > 0;
    const cashAccounts = globalData.individualAccounts.cash;
    const totalAmount = globalData.noteCalculations.cash.total.current;
    const prevTotalAmount = globalData.noteCalculations.cash.total.previous;

    // If both totals are zero AND no mapped accounts AND no selection, skip the note
    if (!hasSelectionDetails && totalAmount === 0 && prevTotalAmount === 0 && Object.keys(cashAccounts).length === 0) {
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

    // 3. Detail Rows / Grouped Rows
    const cashGrouping = selection?.subCategories?.cash;
    if (cashGrouping?.grouped) {
      // Grouped mode: emit two sub-category rows only
      console.log(`[SelectionFirst] Cash: GROUPED MODE using DB sub-category rules.`);
      const cashCurrent = cashGrouping.cash.current;
      const cashPrevious = cashGrouping.cash.previous;
      const bankCurrent = cashGrouping.bankDeposits.current;
      const bankPrevious = cashGrouping.bankDeposits.previous;

      // Row 1: เงินสดในมือ (Cash)
      notes.push(['', '', 'เงินสดในมือ', '', '', '',
        cashCurrent, '',
        processingType === 'multi-year' ? bankPrevious /* intentionally cashPrevious? correct below */ : ''
      ]);
      // Correction: previous should be cashPrevious not bankPrevious
      if (processingType === 'multi-year') {
        notes[notes.length - 1][8] = cashPrevious; // fix previous year value
      }
      tracker.detailRows.push(tracker.currentRow); tracker.currentRow++;

      // Row 2: เงินฝากธนาคาร (Bank Deposits)
      notes.push(['', '', 'เงินฝากธนาคาร', '', '', '',
        bankCurrent, '',
        processingType === 'multi-year' ? bankPrevious : ''
      ]);
      tracker.detailRows.push(tracker.currentRow); tracker.currentRow++;
    } else if (hasSelectionDetails) {
      console.log(`[SelectionFirst] Cash: using selection-first details (${selectionRows.length} accounts). Total current=${selection?.totals?.cash?.current ?? 'n/a'}, previous=${selection?.totals?.cash?.previous ?? 'n/a'}`);
      selectionRows
        .slice()
        .sort((a, b) => parseInt(a.accountCode || '0', 10) - parseInt(b.accountCode || '0', 10))
        .forEach((a) => {
          notes.push(['', '', a.accountName, '', '', '',
            a.current, '',
            processingType === 'multi-year' ? a.previous : ''
          ]);
          tracker.detailRows.push(tracker.currentRow);
          tracker.currentRow++;
        });
    } else {
      const sortedCashEntries = Object.entries(cashAccounts).sort((a, b) => {
        const aCode = parseInt(a[0] || '0', 10);
        const bCode = parseInt(b[0] || '0', 10);
        return aCode - bCode;
      });
      console.log(`[SelectionFirst] Cash: no selection details; falling back to globalData.individualAccounts (${sortedCashEntries.length} accounts).`);
      sortedCashEntries.forEach(([_, accountData]) => {
        notes.push(['', '', accountData.accountName, '', '', '',
          accountData.current, '',
          processingType === 'multi-year' ? accountData.previous : ''
        ]);
        tracker.detailRows.push(tracker.currentRow);
        tracker.currentRow++;
      });
    }

    // 4. Total Row - Use SUM formulas over detail rows when available; fallback to numeric totals otherwise
    const hasDetails = tracker.detailRows.length > 0;
    const firstDetailRow = hasDetails ? tracker.detailRows[0] : null;
    const lastDetailRow = hasDetails ? tracker.detailRows[tracker.detailRows.length - 1] : null;
    let currentTotalCell: any;
    let previousTotalCell: any;
    if (cashGrouping?.grouped && hasDetails) {
      // Use SUM over the two grouped rows (still works generically)
      currentTotalCell = { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` };
      previousTotalCell = processingType === 'multi-year' ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } : '';
    } else {
      currentTotalCell = hasDetails ? { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` } : totalAmount;
      previousTotalCell = processingType === 'multi-year'
        ? (hasDetails ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } : prevTotalAmount)
        : '';
    }

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
