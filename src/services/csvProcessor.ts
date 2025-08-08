// CSV Processing Service for flexible CSV format

import type { TrialBalanceEntry, CompanyInfo } from '../types/financial';

interface CSVTrialBalanceEntry {
  ชื่อบัญชี: string;
  รหัสบัญชี: string;
  ยอดยกมาต้นงวด: number;
  ยอดยกมางวดนี้: number;
  เดบิต: number;
  เครดิต: number;
}

interface CSVColumnMapping {
  accountName: number;      // Column index for account name
  accountCode: number;      // Column index for account code
  previousBalance: number;  // Column index for previous balance
  currentBalance: number;   // Column index for current balance
  debit: number;           // Column index for debit
  credit: number;          // Column index for credit
}

interface CSVParseOptions {
  delimiter?: string;       // CSV delimiter (default: ',')
  skipHeaderRows?: number;  // Number of header rows to skip (default: 1)
  encoding?: string;        // Text encoding (default: 'utf-8')
}

export class CSVProcessor {
  
  /**
   * Detect column mapping from CSV header line
   */
  static detectColumnMapping(headerLine: string, delimiter: string = ','): CSVColumnMapping {
    const headers = headerLine.split(delimiter).map(h => h.trim());
    
    return {
      accountName: this.findColumnIndex(headers, ['ชื่อบัญชี', 'ชื่อ', 'Account Name', 'AccountName']),
      accountCode: this.findColumnIndex(headers, ['รหัสบัญชี', 'รหัส', 'Account Code', 'Code', 'AccountCode']),
      previousBalance: this.findColumnIndex(headers, ['ยอดยกมาต้นงวด', 'ยอดต้นงวด', 'Previous Balance', 'PrevBalance']),
      currentBalance: this.findColumnIndex(headers, ['ยอดยกมางวดนี้', 'ยอดงวดนี้', 'Current Balance', 'CurrBalance']),
      debit: this.findColumnIndex(headers, ['เดบิต', 'Debit', 'DR', 'เดบิท']),
      credit: this.findColumnIndex(headers, ['เครดิต', 'Credit', 'CR', 'เครดิท'])
    };
  }
  
  /**
   * Find column index from possible header variations
   */
  private static findColumnIndex(headers: string[], possibleNames: string[]): number {
    for (const name of possibleNames) {
      const index = headers.findIndex(h => 
        h.toLowerCase().includes(name.toLowerCase()) || 
        name.toLowerCase().includes(h.toLowerCase())
      );
      if (index !== -1) return index;
    }
    throw new Error(`Column not found. Available headers: ${headers.join(', ')}. Expected one of: ${possibleNames.join(', ')}`);
  }
  
  /**
   * Auto-detect CSV delimiter
   */
  private static detectDelimiter(sampleLine: string): string {
    const delimiters = [',', ';', '\t', '|'];
    let maxCount = 0;
    let bestDelimiter = ',';
    
    for (const delimiter of delimiters) {
      const count = sampleLine.split(delimiter).length;
      if (count > maxCount) {
        maxCount = count;
        bestDelimiter = delimiter;
      }
    }
    
    return bestDelimiter;
  }
  
  /**
   * Parse CSV content into trial balance entries with flexible column mapping
   */
  static parseCSV(csvContent: string, options: CSVParseOptions = {}): CSVTrialBalanceEntry[] {
    const lines = csvContent.split('\n').filter(line => line.trim());
    if (lines.length < 2) {
      throw new Error('CSV must have at least a header and one data row');
    }
    
    const entries: CSVTrialBalanceEntry[] = [];
    
    // Auto-detect delimiter if not provided
    const delimiter = options.delimiter || this.detectDelimiter(lines[0]);
    
    // Try to detect column mapping from header
    let mapping: CSVColumnMapping;
    try {
      mapping = this.detectColumnMapping(lines[0], delimiter);
    } catch (error) {
      // Fallback to default column positions if header detection fails
      console.warn('Header detection failed, using default column positions:', error);
      mapping = {
        accountName: 0,
        accountCode: 1,
        previousBalance: 2,
        currentBalance: 3,
        debit: 4,
        credit: 5
      };
    }
    
    const skipRows = options.skipHeaderRows || 1;
    
    for (let i = skipRows; i < lines.length; i++) {
      const values = lines[i].split(delimiter).map(v => v.trim());
      
      // Skip rows that don't have enough columns or empty account code
      if (values.length <= Math.max(...Object.values(mapping)) || 
          !values[mapping.accountCode]?.trim()) {
        continue;
      }
      
      entries.push({
        ชื่อบัญชี: values[mapping.accountName] || '',
        รหัสบัญชี: values[mapping.accountCode] || '',
        ยอดยกมาต้นงวด: parseFloat(values[mapping.previousBalance]) || 0,
        ยอดยกมางวดนี้: parseFloat(values[mapping.currentBalance]) || 0,
        เดบิต: parseFloat(values[mapping.debit]) || 0,
        เครดิต: parseFloat(values[mapping.credit]) || 0
      });
    }
    
    return entries;
  }
  
  /**
   * Calculate correct balance based on account type and statement usage
   */
  private static calculateCorrectBalance(entry: CSVTrialBalanceEntry): number {
    const accountCode = entry.รหัสบัญชี;
    
    if (accountCode.startsWith('4')) {
      // Revenue accounts: Credit - Debit (normal credit balance)
      // For P&L: use เครดิต - เดบิต
      return entry.เครดิต - entry.เดบิต;
    } else if (accountCode.startsWith('5')) {
      // Expense accounts: Debit - Credit (normal debit balance)  
      // For P&L: use เดบิต - เครดิต
      return entry.เดบิต - entry.เครดิต;
    } else {
      // Balance Sheet accounts (1xxx, 2xxx, 3xxx): use cumulative balance
      // Use ยอดยกมาต้นงวด + ยอดยกมางวดนี้ for current year
      return entry.ยอดยกมาต้นงวด + entry.ยอดยกมางวดนี้;
    }
  }

  /**
   * Convert CSV entries to current year TrialBalanceEntry format
   */
  static convertToCurrentYear(csvEntries: CSVTrialBalanceEntry[]): TrialBalanceEntry[] {
    return csvEntries.map(entry => ({
      accountCode: entry.รหัสบัญชี,
      accountName: entry.ชื่อบัญชี,
      balance: this.calculateCorrectBalance(entry),
      debitAmount: entry.เดบิต || 0,
      creditAmount: entry.เครดิต || 0,
      // Preserve original CSV data for PPE movement calculation
      previousBalance: entry.ยอดยกมาต้นงวด,
      currentBalance: entry.ยอดยกมางวดนี้
    }));
  }
  
  /**
   * Calculate previous year balance for comparative statements
   */
  private static calculatePreviousYearBalance(entry: CSVTrialBalanceEntry): number {
    const accountCode = entry.รหัสบัญชี;
    
    if (accountCode.startsWith('4') || accountCode.startsWith('5')) {
      // P&L accounts: previous year would typically be 0 (P&L resets each year)
      // But we use ยอดยกมาต้นงวด if available for comparative P&L
      return entry.ยอดยกมาต้นงวด;
    } else {
      // Balance Sheet accounts: use ยอดยกมาต้นงวด
      return entry.ยอดยกมาต้นงวด;
    }
  }

  /**
   * Convert CSV entries to previous year TrialBalanceEntry format
   */
  static convertToPreviousYear(csvEntries: CSVTrialBalanceEntry[]): TrialBalanceEntry[] {
    return csvEntries.map(entry => ({
      accountCode: entry.รหัสบัญชี,
      accountName: entry.ชื่อบัญชี,
      balance: this.calculatePreviousYearBalance(entry),
      debitAmount: entry.ยอดยกมาต้นงวด > 0 ? entry.ยอดยกมาต้นงวด : 0,
      creditAmount: entry.ยอดยกมาต้นงวด < 0 ? Math.abs(entry.ยอดยกมาต้นงวด) : 0
    }));
  }
  
  /**
   * Determine if CSV data contains multi-year information
   */
  static detectMultiYear(csvEntries: CSVTrialBalanceEntry[]): boolean {
    // Check if any previous year balances exist (non-zero ยอดยกมาต้นงวด)
    return csvEntries.some(entry => entry.ยอดยกมาต้นงวด !== 0);
  }
  
  /**
   * Separate balance sheet and P&L accounts
   */
  static separateAccounts(entries: TrialBalanceEntry[]) {
    const balanceSheetAccounts = entries.filter(entry => 
      entry.accountCode.startsWith('1') || // Assets
      entry.accountCode.startsWith('2') || // Liabilities  
      entry.accountCode.startsWith('3')    // Equity
    );
    
    const profitLossAccounts = entries.filter(entry =>
      entry.accountCode.startsWith('4') || // Revenue
      entry.accountCode.startsWith('5')    // Expenses
    );
    
    return {
      balanceSheet: balanceSheetAccounts,
      profitLoss: profitLossAccounts
    };
  }
  
  /**
   * Check for inventory accounts
   */
  static checkHasInventory(entries: TrialBalanceEntry[]): { hasInventory: boolean; hasPurchases: boolean; inventoryAccount?: TrialBalanceEntry; purchaseAccounts: TrialBalanceEntry[] } {
    const inventoryAccount = entries.find(entry => 
      entry.accountCode === '1510' || entry.accountCode === '1300'
    );
    const purchaseAccounts = entries.filter(entry => 
      entry.accountCode === '5010'
    );
    
    return { 
      hasInventory: !!inventoryAccount,
      hasPurchases: purchaseAccounts.length > 0,
      inventoryAccount,
      purchaseAccounts 
    };
  }
  
  /**
   * Classify accounts by expense categories
   */
  static classifyAccounts(entries: TrialBalanceEntry[]) {
    const adminExpenses = entries.filter(entry => {
      const code = entry.accountCode;
      return (code >= '5300' && code <= '5311') || 
             (code >= '5312' && code <= '5350') ||
             (code >= '5351' && code <= '5399');
    });
    
    const salesExpenses = entries.filter(entry => {
      const code = entry.accountCode;
      return code >= '5200' && code <= '5299';
    });
    
    const financialCosts = entries.filter(entry => {
      const code = entry.accountCode;
      return code >= '5360' && code <= '5369';
    });
    
    const otherExpenses = entries.filter(entry => {
      const code = entry.accountCode;
      return code >= '5400' && code <= '5999';
    });
    
    return {
      salesExpenses,
      adminExpenses,
      otherExpenses,
      financialCosts
    };
  }
  
  /**
   * Process complete CSV file with multi-year support
   */
  static processCsvFile(csvContent: string, companyInfo: CompanyInfo) {
    const csvEntries = this.parseCSV(csvContent);
    const isMultiYear = this.detectMultiYear(csvEntries);
    
    const currentYearEntries = this.convertToCurrentYear(csvEntries);
    const { balanceSheet: currentBalanceSheet, profitLoss: currentProfitLoss } = 
      this.separateAccounts(currentYearEntries);
    
    // Debug logging
    console.log('=== CSV PROCESSING DEBUG ===');
    console.log('Total CSV entries:', csvEntries.length);
    console.log('Current year entries:', currentYearEntries.length);
    console.log('Balance Sheet accounts:', currentBalanceSheet.length);
    console.log('P&L accounts:', currentProfitLoss.length);
    console.log('Sample P&L accounts:', currentProfitLoss.slice(0, 5).map(acc => ({
      code: acc.accountCode,
      name: acc.accountName,
      balance: acc.balance,
      debit: acc.debitAmount,
      credit: acc.creditAmount
    })));
    console.log('=== END CSV DEBUG ===');
    
    if (isMultiYear) {
      const previousYearEntries = this.convertToPreviousYear(csvEntries);
      const { balanceSheet: previousBalanceSheet, profitLoss: previousProfitLoss } = 
        this.separateAccounts(previousYearEntries);
      
      return {
        trialBalance: [...currentBalanceSheet, ...currentProfitLoss], // ← Combine all accounts
        profitLoss: currentProfitLoss,
        trialBalancePrevious: [...previousBalanceSheet, ...previousProfitLoss], // ← Combine all accounts
        profitLossPrevious: previousProfitLoss,
        companyInfo: companyInfo,
        processingType: 'multi-year' as const
      };
    } else {
      return {
        trialBalance: [...currentBalanceSheet, ...currentProfitLoss], // ← Combine all accounts
        profitLoss: currentProfitLoss,
        companyInfo: companyInfo,
        processingType: 'single-year' as const
      };
    }
  }
}
