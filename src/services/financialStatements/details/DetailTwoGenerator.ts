import type { TrialBalanceEntry } from '../../../types/financial';

export class DetailTwoGenerator {
  /**
   * Detail Two: Selling and Administrative Expenses
   * Categorizes expenses into selling, administrative, and other categories
   * Uses Excel formulas for dynamic totals calculation
   */
  static generateDetailTwo(
    trialBalanceData: TrialBalanceEntry[],
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

    // Initialize totals for each category
    let sellingExpensesTotal = 0;
    let adminExpensesTotal = 0; 
    let otherExpensesTotal = 0;
    let financialCostsTotal = 0;
    
    // Get all expense accounts from trial balance data
    const expenseAccounts = trialBalanceData.filter((entry: TrialBalanceEntry) => {
      const code = parseInt(entry.accountCode || '0');
      return code >= 5300 && code <= 5999 && entry.accountCode !== '5910';
    });
    
    // Process each expense account
    expenseAccounts.forEach((account: TrialBalanceEntry) => {
      const accountCode = account.accountCode || '';
      const accountName = account.accountName || '';
      const amount = Math.abs(account.balance || 0); // Expenses are positive
      
      if (amount === 0) return; // Skip zero amounts
      
      // Check account code ranges
      const codeNum = parseInt(accountCode);
      
      if (codeNum >= 5360 && codeNum <= 5364) {
        // Financial costs - handle separately
        financialCostsTotal += amount;
      } else {
        // Regular expense accounts
        let sellingAmount = 0;
        let adminAmount = 0;
        let otherAmount = 0;
        
        if (codeNum >= 5300 && codeNum <= 5311) {
          // Selling expenses
          sellingAmount = amount;
          sellingExpensesTotal += amount;
        } else if (codeNum >= 5312 && codeNum <= 5350) {
          // Admin expenses
          adminAmount = amount;
          adminExpensesTotal += amount;
        } else if ((codeNum >= 5351 && codeNum <= 5359) || (codeNum >= 5365 && codeNum <= 5999)) {
          // Other expenses
          otherAmount = amount;
          otherExpensesTotal += amount;
        }
        
        // Add account row (account name in columns A-F merged, amounts in G,H,I)
        const r = pushRow([accountName, '', '', '', '', '', sellingAmount, adminAmount, otherAmount]);
        if (firstDetailRow === 0) firstDetailRow = r;
        lastDetailRow = r;
      }
    });
    
    // Add total row with formulas
    const totalsG = firstDetailRow > 0 ? { f: `SUM(G${firstDetailRow}:G${lastDetailRow})` } : 0;
    const totalsH = firstDetailRow > 0 ? { f: `SUM(H${firstDetailRow}:H${lastDetailRow})` } : 0;
    const totalsI = firstDetailRow > 0 ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` } : 0;
    pushRow(['รวม', '', '', '', '', '', totalsG, totalsH, totalsI]);
    
    // Add financial costs row if there are any
    if (financialCostsTotal > 0) {
      pushRow(['ค่าใช้จ่ายต้นทุนทางการเงิน', '', '', '', '', '', 0, 0, financialCostsTotal]);
    }
    
    return detailNotes;
  }
}
