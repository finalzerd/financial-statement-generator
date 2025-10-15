import type { TrialBalanceEntry } from '../../../types/financial';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';

export class DetailTwoGenerator {
  /**
   * Detail Two: Selling and Administrative Expenses
   * Categorizes expenses into selling, administrative, and other categories
   * Uses Excel formulas for dynamic totals calculation
   */
  static generateDetailTwo(
    trialBalanceData: TrialBalanceEntry[],
    selection?: SelectionFirstResult,
    startingRow: number = 1
  ): any[][] {
    const detailNotes: any[][] = [];
    let currentRow = startingRow;
    const pushRow = (row: any[]): number => {
      detailNotes.push(row);
      const insertedRow = currentRow;
      currentRow += 1;
      return insertedRow;
    };

    // Header for Detail 2
    pushRow(['รายละเอียดประกอบที่ 2', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    
    // Column headers row
    pushRow(['ค่าใช้จ่ายในการขายและบริหาร', '', '', '', '', '', 'ค่าใช้จ่ายในการขาย', 'ค่าใช้จ่ายในการบริหาร', 'ค่าใช้จ่ายอื่น']);
    
    // Track starting row for totals calculation (first detail row will be the next push)
    let firstDetailRow = 0;
    let lastDetailRow = 0;

    // Selection-first: read categorized accounts; fallback to numeric if selection missing
    const sellingAccs = selection?.byCategory?.selling_expenses ?? [];
    const adminAccs = selection?.byCategory?.admin_expenses ?? [];
    const otherAccs = selection?.byCategory?.other_expenses ?? [];
    const financialAccs = selection?.byCategory?.financial_costs ?? [];

    const emitAccount = (name: string, sell: number, admin: number, other: number) => {
      const r = pushRow([name, '', '', '', '', '', sell, admin, other]);
      if (firstDetailRow === 0) firstDetailRow = r;
      lastDetailRow = r;
    };

    if (selection) {
      // Emit rows grouped by category using selection-first data
      for (const a of sellingAccs) {
        if ((a.current ?? 0) !== 0) emitAccount(a.accountName, a.current, 0, 0);
      }
      for (const a of adminAccs) {
        if ((a.current ?? 0) !== 0) emitAccount(a.accountName, 0, a.current, 0);
      }
      for (const a of otherAccs) {
        if ((a.current ?? 0) !== 0) emitAccount(a.accountName, 0, 0, a.current);
      }
    } else {
      // Legacy fallback: filter by numeric ranges
      const expenseAccounts = trialBalanceData.filter((entry: TrialBalanceEntry) => {
        const code = parseInt(entry.accountCode || '0');
        return code >= 5300 && code <= 5999 && entry.accountCode !== '5910';
      });
      expenseAccounts.forEach((account: TrialBalanceEntry) => {
        const accountCode = account.accountCode || '';
        const accountName = account.accountName || '';
        const amount = Math.abs(account.balance || 0);
        if (amount === 0) return;
        const codeNum = parseInt(accountCode);
        if (codeNum >= 5300 && codeNum <= 5311) {
          emitAccount(accountName, amount, 0, 0);
        } else if (codeNum >= 5312 && codeNum <= 5350) {
          emitAccount(accountName, 0, amount, 0);
        } else if ((codeNum >= 5351 && codeNum <= 5999)) {
          emitAccount(accountName, 0, 0, amount);
        }
      });
    }
    
    // Add total row with formulas
    const totalsG = firstDetailRow > 0 ? { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` } : 0;
    const totalsH = firstDetailRow > 0 ? { f: `SUM(H${firstDetailRow}:H${lastDetailRow})` } : 0;
    const totalsI = firstDetailRow > 0 ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } : 0;
    pushRow(['รวม', '', '', '', '', '', totalsG, totalsH, totalsI]);
    
    // Add financial costs row if there are any (selection-first)
    const financialCostsTotal = selection?.totals?.financial_costs?.current ?? (
      financialAccs.reduce((s, a) => s + (a.current ?? 0), 0)
    );
    if ((financialCostsTotal ?? 0) > 0) {
      pushRow(['ค่าใช้จ่ายต้นทุนทางการเงิน', '', '', '', '', '', 0, 0, financialCostsTotal]);
    }
    
    return detailNotes;
  }
}
