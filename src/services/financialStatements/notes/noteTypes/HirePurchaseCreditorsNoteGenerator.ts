// ============================================================================
// HIRE PURCHASE CREDITORS NOTE GENERATOR
// ============================================================================

import type { NoteRowTracker } from '../../core/types';
import type { SelectionFirstResult } from '../../selection/SelectionFirstClassifier';
import type { TrialBalanceEntry, CompanyInfo } from '../../../../types/financial';

interface DetailDefinition {
  code: string;
  label: string;
  type: 'principal' | 'offset';
}

export class HirePurchaseCreditorsNoteGenerator {
  private static readonly DETAIL_DEFINITIONS: DetailDefinition[] = [
    { code: '2015', label: 'เจ้าหนี้ตามสัญญาเช่าซื้อ', type: 'principal' },
    { code: '1644.2', label: 'หัก ดอกผลเช่าซื้อรอตัดบัญชี', type: 'offset' },
    { code: '1644.1', label: 'หัก ภาษีซื้อรอตัดบัญชี', type: 'offset' }
  ];

  static generateWithRowTracking(
    notes: any[][],
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    _trialBalancePrevious?: TrialBalanceEntry[],
    noteNumber: number = 21,
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

    const normalize = (code: string) => code.trim().replace(/\s+/g, '');
    const selectionMap = new Map<string, { current: number; previous: number; rawCurrent?: number; rawPrevious?: number }>();

    (selection?.byCategory?.hire_purchase_creditors ?? []).forEach(acc => {
      const key = normalize(acc.accountCode);
      selectionMap.set(key, {
        current: acc.current,
        previous: acc.previous,
        rawCurrent: acc.rawCurrent,
        rawPrevious: acc.rawPrevious
      });
    });

    const detailRows: Array<{
      def: DetailDefinition;
      current: number;
      previous: number;
      rawCurrent: number;
      rawPrevious: number;
    }> = [];

    for (const def of this.DETAIL_DEFINITIONS) {
      const key = normalize(def.code);
      const selectionEntry = selectionMap.get(key);
      const rawCurrent = selectionEntry?.rawCurrent ?? this.sumRawByCode(trialBalanceData, key, 'current');
      const rawPrevious = selectionEntry?.rawPrevious ?? this.sumRawByCode(trialBalanceData, key, 'previous');
      const displayCurrent = Math.abs(rawCurrent);
      const displayPrevious = Math.abs(rawPrevious);

      if (displayCurrent === 0 && (processingType !== 'multi-year' || displayPrevious === 0)) {
        continue;
      }

      detailRows.push({
        def,
        current: displayCurrent,
        previous: displayPrevious,
        rawCurrent,
        rawPrevious
      });
    }

    if (detailRows.length === 0) {
      console.log('[HirePurchaseNote] Skipping note - all detail values are zero.');
      return tracker;
    }

    // Header
    notes.push([noteNumber.toString(), 'เจ้าหนี้ตามสัญญาเช่าซื้อ', '', '', '', '', '', '', 'หน่วย:บาท']);
    tracker.headerRows.push(tracker.currentRow);
    tracker.unitRows.push(tracker.currentRow);
    tracker.currentRow++;

    // Year header row
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }
    tracker.yearHeaderRows.push(tracker.currentRow);
    tracker.currentRow++;

    const principalRows: number[] = [];
    const offsetRows: number[] = [];

    detailRows.forEach(detail => {
      notes.push([
        '',
        '',
        detail.def.label,
        '',
        '',
        '',
        detail.current,
        '',
        processingType === 'multi-year' ? detail.previous : ''
      ]);
      tracker.detailRows.push(tracker.currentRow);
      if (detail.def.type === 'principal') {
        principalRows.push(tracker.currentRow);
      } else {
        offsetRows.push(tracker.currentRow);
      }
      tracker.currentRow++;
    });

    // Net subtotal (principal minus offsets)
    let netCurrentFormula = '0';
    let netPreviousFormula = '0';

    if (principalRows.length > 0) {
      netCurrentFormula = `G${principalRows[0]}`;
      netPreviousFormula = processingType === 'multi-year' ? `I${principalRows[0]}` : '';
    } else if (offsetRows.length > 0) {
      netCurrentFormula = `-(${offsetRows.map(row => `G${row}`).join('+')})`;
      if (processingType === 'multi-year') {
        netPreviousFormula = `-(${offsetRows.map(row => `I${row}`).join('+')})`;
      }
    }

    if (principalRows.length > 0 && offsetRows.length > 0) {
      offsetRows.forEach(row => {
        netCurrentFormula += `-G${row}`;
        if (processingType === 'multi-year') {
          netPreviousFormula += `-I${row}`;
        }
      });
    }

    const hasPrevious = processingType === 'multi-year';
    const netCurrentCell = netCurrentFormula === '0' ? 0 : ({ f: netCurrentFormula } as any);
    const netPreviousCell = hasPrevious
      ? (netPreviousFormula === '0' ? 0 : ({ f: netPreviousFormula } as any))
      : '';

    notes.push([
      '',
      '',
      'รวมเจ้าหนี้ตามสัญญาเช่าซื้อสุทธิ',
      '',
      '',
      '',
      netCurrentCell,
      '',
      netPreviousCell
    ]);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    const dueWithinLabel = 'เจ้าหนี้ตามสัญญาเช่าซื้อสุทธิส่วนที่ถึงกำหนดชำระภายในหนึ่งปี';
    const dueBeyondLabel = 'เจ้าหนี้ตามสัญญาเช่าซื้อสุทธิจากส่วนที่ถึงกำหนดชำระในหนึ่งปี';

    const dueWithinRow = tracker.currentRow;
    notes.push([
      '',
      '',
      dueWithinLabel,
      '',
      '',
      '',
      '',
      '',
      hasPrevious ? '' : ''
    ]);
    tracker.detailRows.push(dueWithinRow);
    tracker.currentRow++;

    const dueBeyondRow = tracker.currentRow;
    notes.push([
      '',
      '',
      dueBeyondLabel,
      '',
      '',
      '',
      '',
      '',
      hasPrevious ? '' : ''
    ]);
    tracker.detailRows.push(dueBeyondRow);
    tracker.currentRow++;

    const totalCurrentFormula = `SUM(G${dueWithinRow}:G${dueBeyondRow})`;
    const totalPreviousFormula = hasPrevious ? `SUM(I${dueWithinRow}:I${dueBeyondRow})` : '';

    notes.push([
      '',
      '',
      'รวม',
      '',
      '',
      '',
      { f: totalCurrentFormula } as any,
      '',
      hasPrevious ? ({ f: totalPreviousFormula } as any) : ''
    ]);
    tracker.totalRows.push(tracker.currentRow);
    tracker.currentRow++;

    // Spacer
    notes.push(['', '', '', '', '', '', '', '', '']);
    tracker.currentRow++;

    return tracker;
  }

  private static sumRawByCode(
    trialBalanceData: TrialBalanceEntry[],
    targetCode: string,
    which: 'current' | 'previous'
  ): number {
    const normalized = targetCode.replace(/\s+/g, '');
    return trialBalanceData.reduce((sum, entry) => {
      const entryCode = (entry.accountCode || '').trim().replace(/\s+/g, '');
      if (entryCode !== normalized) {
        return sum;
      }
      const value = which === 'previous' ? (entry.previousBalance ?? 0) : (entry.balance ?? entry.currentBalance ?? 0);
      return sum + (value || 0);
    }, 0);
  }
}
