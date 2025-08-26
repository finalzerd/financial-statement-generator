import type { TrialBalanceEntry } from '../../../types/financial';

export class DetailTwoGenerator {
  /**
   * Detail Two: Selling and Administrative Expenses
   * Categorizes expenses into selling, administrative, and other categories
   * Uses Excel formulas for dynamic totals calculation
   */
  static generateDetailTwo(
    trialBalanceData: TrialBalanceEntry[]
  ): any[][] {
    const detailNotes: any[][] = [];
    
    // Header for Detail 2
    detailNotes.push(['รายละเอียดประกอบที่ 2', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    
    // Column headers row
    detailNotes.push(['ค่าใช้จ่ายในการขายและบริหาร', '', '', '', '', '', 'ค่าใช้จ่ายในการขาย', 'ค่าใช้จ่ายในการบริหาร', 'ค่าใช้จ่ายอื่น']);
    
    // Track starting row for totals calculation
    const dataStartRow = detailNotes.length + 1; // Next row number (1-indexed for Excel)
    
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
        detailNotes.push([accountName, '', '', '', '', '', sellingAmount, adminAmount, otherAmount]);
      }
    });
    
    // Calculate ending row for totals
    const dataEndRow = detailNotes.length; // Current row number (1-indexed for Excel)
    
    // Add total row with formulas
    detailNotes.push(['รวม', '', '', '', '', '', 
      { f: `SUM(G${dataStartRow}:G${dataEndRow})` },
      { f: `SUM(H${dataStartRow}:H${dataEndRow})` },
      { f: `SUM(I${dataStartRow}:I${dataEndRow})` }
    ]);
    
    // Add financial costs row if there are any
    if (financialCostsTotal > 0) {
      detailNotes.push(['ค่าใช้จ่ายต้นทุนทางการเงิน', '', '', '', '', '', 0, 0, financialCostsTotal]);
    }
    
    return detailNotes;
  }
}
