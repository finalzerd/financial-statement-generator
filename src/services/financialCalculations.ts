import type { TrialBalanceEntry, CompanyInfo } from '../types/financial';

/**
 * Financial Calculation Utilities
 * 
 * Phase 1: Pure calculation helper methods extracted from FinancialStatementGenerator
 * These are static utility methods that can be safely reused across the application
 * without breaking the existing generator architecture.
 */
export class FinancialCalculations {
  
  // ============================================================================
  // ACCOUNT BALANCE CALCULATIONS
  // ============================================================================
  
  /**
   * Sum accounts by numeric range (most frequently used calculation)
   * Optimized to use pre-calculated balance field from TrialBalanceEntry
   */
  static sumAccountsByNumericRange(
    trialBalanceData: TrialBalanceEntry[], 
    startCode: number, 
    endCode: number
  ): number {
    const matchingEntries = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });
    
    return matchingEntries.reduce((sum, entry) => {
      const balance = entry.balance || 0;
      return sum + balance;
    }, 0);
  }

  /**
   * Sum previous balance by numeric range for comparative statements
   */
  static sumPreviousBalanceByNumericRange(
    trialBalanceData: TrialBalanceEntry[], 
    startCode: number, 
    endCode: number
  ): number {
    const matchingEntries = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });

    return matchingEntries.reduce((sum, entry) => {
      return sum + Math.abs(entry.previousBalance || 0);
    }, 0);
  }

  /**
   * Get single account balance by account code
   */
  static getSingleAccountBalance(
    trialBalanceData: TrialBalanceEntry[], 
    accountCode: string
  ): number {
    const account = trialBalanceData.find(entry => entry.accountCode === accountCode);
    if (!account) return 0;
    
    // For credit accounts (like capital), the balance field already contains the correct amount
    return Math.abs(account.balance || 0);
  }

  /**
   * Get account balance for multiple account codes
   */
  static getAccountBalance(
    trialBalanceData: TrialBalanceEntry[], 
    accountCodes: string[]
  ): number {
    return accountCodes.reduce((sum, code) => {
      return sum + this.getSingleAccountBalance(trialBalanceData, code);
    }, 0);
  }

  /**
   * Sum accounts by range with flexible account code matching
   */
  static sumAccountsByRange(
    trialBalanceData: TrialBalanceEntry[], 
    accountCodes: string[], 
    dataType: 'current' | 'previous'
  ): number {
    const matchingEntries = trialBalanceData.filter(entry => 
      accountCodes.some(code => entry.accountCode === code || entry.accountCode?.startsWith(code))
    );

    return Math.abs(matchingEntries.reduce((sum, entry) => {
      const balance = dataType === 'previous' 
        ? (entry.previousBalance || 0)
        : (entry.balance || 0);
      return sum + balance;
    }, 0));
  }

  // ============================================================================
  // SPECIAL ACCOUNT TYPE CALCULATIONS
  // ============================================================================

  /**
   * Calculate current year revenue (4xxx accounts: credit - debit)
   * Uses original debit-credit logic for P&L accuracy
   */
  static calculateCurrentYearRevenue(trialBalanceData: TrialBalanceEntry[]): number {
    const revenueAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('4'));
    return revenueAccounts
      .reduce((sum, entry) => sum + ((entry.creditAmount || 0) - (entry.debitAmount || 0)), 0);
  }

  /**
   * Calculate current year expenses (5xxx accounts: debit - credit)
   * Uses original debit-credit logic for P&L accuracy
   */
  static calculateCurrentYearExpenses(trialBalanceData: TrialBalanceEntry[]): number {
    const expenseAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('5'));
    return expenseAccounts
      .reduce((sum, entry) => sum + ((entry.debitAmount || 0) - (entry.creditAmount || 0)), 0);
  }

  /**
   * Calculate current year profit (revenue - expenses)
   */
  static calculateCurrentYearProfit(trialBalanceData: TrialBalanceEntry[]): number {
    const revenue = this.calculateCurrentYearRevenue(trialBalanceData);
    const expenses = this.calculateCurrentYearExpenses(trialBalanceData);
    return revenue - expenses;
  }

  /**
   * Get opening retained earnings from account 3020 (credit - debit)
   */
  static getOpeningRetainedEarnings(trialBalanceData: TrialBalanceEntry[]): number {
    const retainedEarningsAccount = trialBalanceData.find(entry => entry.accountCode === '3020');
    return retainedEarningsAccount ? 
      ((retainedEarningsAccount.creditAmount || 0) - (retainedEarningsAccount.debitAmount || 0)) : 0;
  }

  // ============================================================================
  // BUSINESS LOGIC HELPERS
  // ============================================================================

  /**
   * Check if business has inventory (account 1510)
   */
  static checkHasInventory(trialBalanceData: TrialBalanceEntry[]): boolean {
    const inventoryAccount = trialBalanceData.find(entry => entry.accountCode === '1510');
    return inventoryAccount ? Math.abs(inventoryAccount.balance || 0) > 0 : false;
  }

  /**
   * Check if business has purchases (5010 accounts)
   */
  static checkHasPurchases(trialBalanceData: TrialBalanceEntry[]): boolean {
    const purchaseAccounts = trialBalanceData.filter(entry => 
      entry.accountCode?.startsWith('5010') && 
      (entry.balance !== 0)
    );
    return purchaseAccounts.length > 0;
  }

  /**
   * Determine if business is service-based (no inventory)
   */
  static isServiceBusiness(trialBalanceData: TrialBalanceEntry[]): boolean {
    return !this.checkHasInventory(trialBalanceData);
  }

  /**
   * Check if company is limited partnership
   */
  static isLimitedPartnership(companyInfo: CompanyInfo): boolean {
    return companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
  }

  // ============================================================================
  // FORMATTING UTILITIES
  // ============================================================================

  /**
   * Format number with Thai locale formatting
   */
  static formatNumber(amount: number): string {
    return new Intl.NumberFormat('th-TH', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(amount);
  }

  /**
   * Format currency with Thai Baht symbol
   */
  static formatCurrency(amount: number): string {
    return `฿ ${this.formatNumber(amount)}`;
  }

  /**
   * Build Excel SUM formula from row numbers
   */
  static buildSumFormula(rows: number[], column: string): string {
    if (rows.length === 0) return '';
    if (rows.length === 1) return `${column}${rows[0]}`;
    
    // Group consecutive rows into ranges for efficient formulas
    const ranges: string[] = [];
    let start = rows[0];
    let end = rows[0];
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] === end + 1) {
        end = rows[i];
      } else {
        ranges.push(start === end ? `${column}${start}` : `${column}${start}:${column}${end}`);
        start = end = rows[i];
      }
    }
    ranges.push(start === end ? `${column}${start}` : `${column}${start}:${column}${end}`);
    
    return `SUM(${ranges.join(',')})`;
  }

  // ============================================================================
  // ACCOUNT FILTERING UTILITIES
  // ============================================================================

  /**
   * Filter accounts by numeric range
   */
  static filterAccountsByRange(
    trialBalanceData: TrialBalanceEntry[], 
    startCode: number, 
    endCode: number
  ): TrialBalanceEntry[] {
    return trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });
  }

  /**
   * Filter accounts with non-zero balances
   */
  static filterNonZeroAccounts(trialBalanceData: TrialBalanceEntry[]): TrialBalanceEntry[] {
    return trialBalanceData.filter(entry => 
      Math.abs(entry.balance || 0) !== 0 || 
      Math.abs(entry.previousBalance || 0) !== 0
    );
  }

  /**
   * Filter cash accounts (1000-1099) with categorization
   */
  static filterCashAccounts(trialBalanceData: TrialBalanceEntry[]) {
    const cashAccounts = this.filterAccountsByRange(trialBalanceData, 1000, 1099);
    
    return {
      cash: cashAccounts.filter(entry => {
        const code = parseInt(entry.accountCode || '0');
        return code >= 1000 && code <= 1019;
      }),
      bankDeposits: cashAccounts.filter(entry => {
        const code = parseInt(entry.accountCode || '0');
        return code >= 1020 && code <= 1099;
      }),
      all: cashAccounts
    };
  }

  /**
   * Filter receivable accounts (1140-1215)
   */
  static filterReceivableAccounts(trialBalanceData: TrialBalanceEntry[]): TrialBalanceEntry[] {
    return this.filterAccountsByRange(trialBalanceData, 1140, 1215);
  }

  /**
   * Filter payable accounts (2010-2999 with exclusions)
   */
  static filterPayableAccounts(trialBalanceData: TrialBalanceEntry[]): TrialBalanceEntry[] {
    return trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      if (code < 2010 || code > 2999) return false;
      
      // Exclude specific account ranges that are handled separately
      const isExcluded = (code >= 2030 && code <= 2030) || // Short-term borrowings
                        (code >= 2045 && code <= 2045) || // Income tax payable
                        (code >= 2050 && code <= 2052) || // Other long-term loans
                        (code >= 2100 && code <= 2123);   // Long-term loans
      
      return !isExcluded;
    });
  }

  // ============================================================================
  // VALIDATION UTILITIES
  // ============================================================================

  /**
   * Validate that trial balance is balanced (sum of all balances should approach zero)
   */
  static validateTrialBalance(trialBalanceData: TrialBalanceEntry[]): {
    isBalanced: boolean;
    difference: number;
    tolerance: number;
  } {
    const tolerance = 0.01; // Allow 1 cent difference due to rounding
    const totalBalance = trialBalanceData.reduce((sum, entry) => sum + (entry.balance || 0), 0);
    
    return {
      isBalanced: Math.abs(totalBalance) <= tolerance,
      difference: totalBalance,
      tolerance
    };
  }

  /**
   * Get summary statistics for trial balance
   */
  static getTrialBalanceStats(trialBalanceData: TrialBalanceEntry[]) {
    const assets = this.filterAccountsByRange(trialBalanceData, 1000, 1999);
    const liabilities = this.filterAccountsByRange(trialBalanceData, 2000, 2999);
    const equity = this.filterAccountsByRange(trialBalanceData, 3000, 3999);
    const revenue = trialBalanceData.filter(entry => entry.accountCode?.startsWith('4'));
    const expenses = trialBalanceData.filter(entry => entry.accountCode?.startsWith('5'));

    return {
      totalAccounts: trialBalanceData.length,
      assetAccounts: assets.length,
      liabilityAccounts: liabilities.length,
      equityAccounts: equity.length,
      revenueAccounts: revenue.length,
      expenseAccounts: expenses.length,
      totalAssets: this.sumAccountsByNumericRange(trialBalanceData, 1000, 1999),
      totalLiabilities: this.sumAccountsByNumericRange(trialBalanceData, 2000, 2999),
      totalEquity: this.sumAccountsByNumericRange(trialBalanceData, 3000, 3999),
      totalRevenue: this.calculateCurrentYearRevenue(trialBalanceData),
      totalExpenses: this.calculateCurrentYearExpenses(trialBalanceData),
      netProfit: this.calculateCurrentYearProfit(trialBalanceData)
    };
  }
}
