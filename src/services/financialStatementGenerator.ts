import { saveAs } from 'file-saver';
import { ExcelJSFormatter } from './excelFormatter';
import type { 
  TrialBalanceEntry, 
  CompanyInfo, 
  FinancialStatements
} from '../types/financial';

// ============================================================================
// TYPE DEFINITIONS
// ============================================================================

/**
 * Global extracted financial data - calculated once and reused across all statements
 */
interface ExtractedFinancialData {
  // ASSETS
  assets: {
    cashAndCashEquivalents: { current: number; previous: number };
    tradeReceivables: { current: number; previous: number };
    inventory: { current: number; previous: number };
    prepaidExpenses: { current: number; previous: number };
    propertyPlantEquipment: { current: number; previous: number };
    otherAssets: { current: number; previous: number };
  };
  
  // LIABILITIES
  liabilities: {
    bankOverdraftsAndShortTermLoans: { current: number; previous: number };
    tradeAndOtherPayables: { current: number; previous: number };
    shortTermBorrowings: { current: number; previous: number };
    incomeTaxPayable: { current: number; previous: number };
    longTermLoansFromFI: { current: number; previous: number };
    otherLongTermLoans: { current: number; previous: number };
  };
  
  // EQUITY - GLOBAL VALUES (no more duplicated calculations!)
  equity: {
    paidUpCapital: { current: number; previous: number };
    retainedEarnings: { current: number; previous: number };
    openingRetainedEarnings: number; // Opening balance from account 3020 (credit - debit)
    legalReserve: { current: number; previous: number };
  };
  
  // INCOME STATEMENT
  income: {
    revenue: { total: number; mainRevenue: number; otherIncome: number };
    expenses: { total: number; costOfServices: number; adminExpenses: number; otherExpenses: number; incomeTax: number; financialCosts: number };
    netProfit: number;
  };
  
  // BUSINESS LOGIC FLAGS
  flags: {
    hasInventory: boolean;
    isServiceBusiness: boolean;
    isLimitedPartnership: boolean;
  };
}

/**
 * Interface for tracking cell positions during balance sheet generation
 * Used to generate accurate formulas that reference the correct cell locations
 */
interface CellTracker {
  currentRow: number;
  currentLiabilitiesRows: number[];
  nonCurrentLiabilitiesRows: number[];
  equityDataRows: number[];
  currentLiabilitiesTotalRow: number;
  nonCurrentLiabilitiesTotalRow: number;
  totalLiabilitiesRow?: number; // Track total liabilities row for grand total calculation
}

// ============================================================================
// MAIN FINANCIAL STATEMENT GENERATOR CLASS
// ============================================================================

export class FinancialStatementGenerator {
  
  // ============================================================================
  // GLOBAL DATA EXTRACTION (Calculate Once, Use Everywhere)
  // ============================================================================
  
  private extractedData: ExtractedFinancialData | null = null;
  
  /**
   * MAIN DATA EXTRACTION METHOD - Call this first to avoid redundant calculations
   */
  private extractAllFinancialData(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo
  ): ExtractedFinancialData {
    
    if (this.extractedData) {
      return this.extractedData; // Return cached data
    }

    console.log('=== EXTRACTING ALL FINANCIAL DATA (ONCE) ===');

    // ASSETS CALCULATION
    const assets = {
      cashAndCashEquivalents: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1000, 1099)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099))
      },
      tradeReceivables: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1140, 1215)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215))
      },
      inventory: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1500, 1519)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519))
      },
      prepaidExpenses: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1400, 1439)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439))
      },
      propertyPlantEquipment: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1600, 1659)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1659))
      },
      otherAssets: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1660, 1700)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700))
      }
    };

    // LIABILITIES CALCULATION
    const liabilities = {
      bankOverdraftsAndShortTermLoans: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2001, 2009)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2001, 2009))
      },
      tradeAndOtherPayables: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2010, 2999)) - 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)) - 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)) - 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) - 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2100, 2123)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2010, 2999)) - 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)) - 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)) - 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) - 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2123))
      },
      shortTermBorrowings: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030))
      },
      incomeTaxPayable: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045))
      },
      longTermLoansFromFI: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) - 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2121, 2121)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2120, 2123)) - 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2121, 2121))
      },
      otherLongTermLoans: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) + 
                 Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2100, 2119)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) + 
                  Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2119))
      }
    };

    // EQUITY CALCULATION WITH CORRECTED RETAINED EARNINGS
    // Calculate current year profit properly: Revenue (credit-debit) - Expenses (debit-credit)
    const revenueAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('4'));
    const currentYearRevenue = revenueAccounts
      .reduce((sum, entry) => sum + ((entry.creditAmount || 0) - (entry.debitAmount || 0)), 0);
    
    const expenseAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('5'));
    const currentYearExpenses = expenseAccounts
      .reduce((sum, entry) => sum + ((entry.debitAmount || 0) - (entry.creditAmount || 0)), 0);
    
    const currentYearProfit = currentYearRevenue - currentYearExpenses;
    
    // CORRECTED: Get opening retained earnings from account 3020 using credit - debit
    const retainedEarningsAccount = trialBalanceData.find(entry => entry.accountCode === '3020');
    const openingRetainedEarnings = retainedEarningsAccount ? 
      ((retainedEarningsAccount.creditAmount || 0) - (retainedEarningsAccount.debitAmount || 0)) : 0;
    
    // Final retained earnings = opening + current year profit (VBA-compliant)
    const finalRetainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);
    
    console.log('=== GLOBAL RETAINED EARNINGS CALCULATION (CORRECTED) ===');
    console.log('Current Year Revenue:', currentYearRevenue);
    console.log('Current Year Expenses:', currentYearExpenses);
    console.log('Current Year Profit:', currentYearProfit);
    console.log('Account 3020 Credit:', retainedEarningsAccount?.creditAmount || 0);
    console.log('Account 3020 Debit:', retainedEarningsAccount?.debitAmount || 0);
    console.log('Opening Retained Earnings (Credit - Debit):', openingRetainedEarnings);
    console.log('Final Retained Earnings:', finalRetainedEarnings);
    console.log('=== END GLOBAL RETAINED EARNINGS CALCULATION ===');

    const equity = {
      paidUpCapital: {
        current: this.getSingleAccountBalance(trialBalanceData, '3010'),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 3010, 3010))
      },
      retainedEarnings: {
        current: finalRetainedEarnings, // Use corrected calculation
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 3020, 3020))
      },
      openingRetainedEarnings: openingRetainedEarnings, // Store opening balance separately
      legalReserve: {
        current: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3030, 3039)),
        previous: Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 3030, 3039))
      }
    };

    // INCOME STATEMENT CALCULATION
    const revenue = {
      total: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4000, 4999)),
      mainRevenue: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4000, 4099)),
      otherIncome: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4100, 4999))
    };

    const expenses = {
      total: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5000, 5999)),
      costOfServices: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5000, 5099)),
      adminExpenses: Math.abs(
        this.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
        this.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
        this.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
        this.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
      ),
      otherExpenses: Math.abs(
        this.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
        this.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
        this.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
        this.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
      ),
      incomeTax: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5910, 5910)),
      financialCosts: Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5920, 5929))
    };

    const netProfit = revenue.total - expenses.total;

    // BUSINESS LOGIC FLAGS
    const flags = {
      hasInventory: assets.inventory.current > 0,
      isServiceBusiness: assets.inventory.current === 0,
      isLimitedPartnership: companyInfo.type === 'ห้างหุ้นส่วนจำกัด'
    };

    this.extractedData = {
      assets,
      liabilities,
      equity,
      income: { revenue, expenses, netProfit },
      flags
    };

    console.log('=== GLOBAL EQUITY VALUES (SINGLE SOURCE) ===');
    console.log('Paid-up Capital:', equity.paidUpCapital);
    console.log('Retained Earnings:', equity.retainedEarnings);
    console.log('Net Profit:', netProfit);
    console.log('=== END EXTRACTION ===');

    return this.extractedData;
  }
  
  // ============================================================================
  // PUBLIC INTERFACE METHODS
  // ============================================================================
  
  generateFinancialStatements(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    trialBalancePrevious?: TrialBalanceEntry[]
  ): FinancialStatements {
    
    // *** EXTRACT ALL DATA ONCE ***
    const globalData = this.extractAllFinancialData(trialBalanceData, companyInfo);
    
    console.log('=== USING GLOBAL DATA FOR ALL STATEMENTS ===');
    console.log('Paid-up Capital (Global):', globalData.equity.paidUpCapital);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
    const balanceSheetAssets = this.generateBalanceSheetAssets(trialBalanceData, companyInfo, processingType);
    const balanceSheetLiabilities = this.generateBalanceSheetLiabilities(trialBalanceData, companyInfo, processingType);
    const profitLossStatement = this.generateProfitLossStatement(trialBalanceData, companyInfo, processingType);
    const statementOfChangesInEquity = this.generateStatementOfChangesInEquity(trialBalanceData, companyInfo, processingType);
    const notesToFinancialStatements = this.generateNotesToFinancialStatements(companyInfo, trialBalanceData, processingType, trialBalancePrevious);
    const accountingNotes = this.generateAccountingNotes(trialBalanceData, companyInfo, processingType, trialBalancePrevious);
    const detailNotes = this.generateDetailNotes(trialBalanceData, companyInfo);

    return {
      balanceSheet: {
        assets: balanceSheetAssets,
        liabilities: balanceSheetLiabilities
      },
      profitLossStatement,
      changesInEquity: statementOfChangesInEquity,
      notes: notesToFinancialStatements,
      accountingNotes,
      detailNotes: {
        detail1: detailNotes,
        detail2: undefined
      },
      companyInfo,
      processingType
    };
  }

  async downloadAsExcel(statements: FinancialStatements): Promise<void> {
    console.log('Starting Excel generation with ExcelJS...');
    
    const workbook = ExcelJSFormatter.createWorkbook();

    // Create Balance Sheet - Assets
    const assetsWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'BS_Assets', statements.balanceSheet.assets);
    ExcelJSFormatter.formatBalanceSheetAssets(assetsWs);

    // Create Balance Sheet - Liabilities
    const liabilitiesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'BS_Liabilities', statements.balanceSheet.liabilities);
    ExcelJSFormatter.formatBalanceSheetAssets(liabilitiesWs); // Use available formatter

    // Create Profit & Loss Statement
    const plWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'P&L', statements.profitLossStatement);
    ExcelJSFormatter.formatBalanceSheetAssets(plWs); // Use available formatter

    // Create Statement of Changes in Equity
    const equityWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Changes_in_Equity', statements.changesInEquity);
    ExcelJSFormatter.formatStatementOfChangesInEquity(equityWs); // Use correct SCE formatter

    // Create Notes to Financial Statements (Policy Notes)
    const notesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Policy', statements.notes);
    ExcelJSFormatter.formatNotesToFinancialStatements(notesWs);

    // Create Accounting Notes (Detailed Notes)
    const accountingNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Accounting', statements.accountingNotes);
    ExcelJSFormatter.formatNotesWithoutBackground(accountingNotesWs);

    // Create Detail Notes (if available)
    if (statements.detailNotes?.detail1) {
      const detailNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Detail', statements.detailNotes.detail1);
      ExcelJSFormatter.formatDetailNotes(detailNotesWs); // Use detail notes specific formatting
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const fileName = `${statements.companyInfo.name}_FinancialStatements_${statements.companyInfo.reportingYear}.xlsx`;
    saveAs(blob, fileName);
    
    console.log('Excel file generated successfully');
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================
  
  private getAccountBalance(trialBalanceData: TrialBalanceEntry[], accountCodes: string[]): number {
    console.log(`Getting balance for codes:`, accountCodes);
    
    const matchingEntries = trialBalanceData.filter(entry => 
      accountCodes.some(code => {
        return entry.accountCode === code || entry.accountCode?.startsWith(code);
      })
    );
    
    console.log(`Found ${matchingEntries.length} matching entries:`, matchingEntries);
    
    const total = matchingEntries.reduce((sum, entry) => {
      const balance = (entry.debitAmount || 0) - (entry.creditAmount || 0);
      console.log(`Account ${entry.accountCode}: debit=${entry.debitAmount}, credit=${entry.creditAmount}, balance=${balance}`);
      return sum + balance;
    }, 0);
    
    console.log(`Total balance: ${total}`);
    return total;
  }

  private formatNumber(amount: number): string {
    if (amount === 0) return '-';
    return new Intl.NumberFormat('th-TH', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(Math.abs(amount));
  }

  private formatCurrency(amount: number): string {
    if (amount === 0) return '-';
    return new Intl.NumberFormat('th-TH', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(Math.abs(amount));
  }

  private buildSumFormula(rows: number[], column: string): string {
    if (rows.length === 0) return '0';
    if (rows.length === 1) return `${column}${rows[0]}`;
    
    // Create ranges for consecutive rows, individual cells for non-consecutive
    const ranges: string[] = [];
    let rangeStart = rows[0];
    let rangeEnd = rows[0];
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] === rangeEnd + 1) {
        rangeEnd = rows[i];
      } else {
        if (rangeStart === rangeEnd) {
          ranges.push(`${column}${rangeStart}`);
        } else {
          ranges.push(`${column}${rangeStart}:${column}${rangeEnd}`);
        }
        rangeStart = rows[i];
        rangeEnd = rows[i];
      }
    }
    
    // Add the last range
    if (rangeStart === rangeEnd) {
      ranges.push(`${column}${rangeStart}`);
    } else {
      ranges.push(`${column}${rangeStart}:${column}${rangeEnd}`);
    }
    
    return `SUM(${ranges.join(',')})`;
  }

  // ============================================================================
  // ACCOUNT BALANCE CALCULATION METHODS
  // ============================================================================

  private sumAccountsByRange(trialBalanceData: TrialBalanceEntry[], accountCodes: string[], dataType: 'current' | 'previous', trialBalancePrevious?: TrialBalanceEntry[]): number {
    const sourceData = dataType === 'previous' ? trialBalancePrevious || [] : trialBalanceData;
    
    const matchingEntries = sourceData.filter(entry => 
      accountCodes.some(code => entry.accountCode === code || entry.accountCode?.startsWith(code))
    );
    
    return Math.abs(matchingEntries.reduce((sum, entry) => {
      const balance = dataType === 'previous' 
        ? (entry.previousBalance || 0)
        : ((entry.debitAmount || 0) - (entry.creditAmount || 0));
      return sum + balance;
    }, 0));
  }

  private sumAccountsByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    const matchingEntries = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });
    
    return matchingEntries.reduce((sum, entry) => {
      const balance = (entry.debitAmount || 0) - (entry.creditAmount || 0);
      return sum + balance;
    }, 0);
  }

  private sumPreviousBalanceByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    const matchingEntries = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });
    
    return Math.abs(matchingEntries.reduce((sum, entry) => {
      return sum + (entry.previousBalance || 0);
    }, 0));
  }

  // ============================================================================
  // BALANCE SHEET GENERATION METHODS
  // ============================================================================

  private generateBalanceSheetAssets(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year'
  ) {
    // Calculate current year asset balances using VBA-compliant ranges
    const cashAndCashEquivalents = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1000, 1099));
    const tradeReceivables = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1140, 1215));
    const inventory = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1500, 1519));
    const prepaidExpenses = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1400, 1439));
    const landBuildingsEquipment = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1600, 1659));
    const otherAssets = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1660, 1700));

    // Calculate previous year asset balances using previousBalance field
    const prevCashAndCashEquivalents = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099);
    const prevTradeReceivables = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215);
    const prevInventory = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519);
    const prevPrepaidExpenses = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439);
    const prevLandBuildingsEquipment = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1659);
    const prevOtherAssets = this.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700);

    // Initialize worksheet data with headers
    const worksheetData: (string | number | {f: string})[][] = [
      [companyInfo.name, '', '', '', '', '', '', '', '', ''],
      ['งบแสดงฐานะการเงิน', '', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', 'หมายเหตุ', '', '', 'หน่วย:บาท', ''], // Row 5: หมายเหตุ (bold + underline) and หน่วย:บาท (bold)
      ['', 'สินทรัพย์', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '', ''] // Row 6: สินทรัพย์ (bold), years (general format)
    ];

    // Track current row and data rows for formulas
    let currentRow = worksheetData.length + 1;
    const currentAssetRows: number[] = [];

    // Current Assets section
    worksheetData.push(['', 'สินทรัพย์หมุนเวียน', '', '', '', '', '', '', '', '']);
    currentRow++;

    // Add current assets
    if (cashAndCashEquivalents !== 0) {
      worksheetData.push(['', '', 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '7', cashAndCashEquivalents, '', processingType === 'multi-year' ? prevCashAndCashEquivalents : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (tradeReceivables !== 0) {
      worksheetData.push(['', '', 'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น', '', '', '8', tradeReceivables, '', processingType === 'multi-year' ? prevTradeReceivables : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (inventory !== 0) {
      worksheetData.push(['', '', 'สินค้าคงเหลือ', '', '', '9', inventory, '', processingType === 'multi-year' ? prevInventory : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (prepaidExpenses !== 0) {
      worksheetData.push(['', '', 'ค่าใช้จ่ายจ่ายล่วงหน้า', '', '', '10', prepaidExpenses, '', processingType === 'multi-year' ? prevPrepaidExpenses : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    // Current Assets Total
    const currentAssetsFormula = this.buildSumFormula(currentAssetRows, 'G');
    const currentAssetsFormulaPrev = processingType === 'multi-year' ? this.buildSumFormula(currentAssetRows, 'I') : '';
    
    worksheetData.push(['', 'รวมสินทรัพย์หมุนเวียน', '', '', '', '', 
      { f: currentAssetsFormula }, 
      '', 
      processingType === 'multi-year' ? { f: currentAssetsFormulaPrev } : '', '']);
    const currentAssetsTotalRow = currentRow;
    currentRow++;

    // Spacer
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    currentRow++;

    // Non-Current Assets section
    const nonCurrentAssetRows: number[] = [];
    
    worksheetData.push(['', 'สินทรัพย์ไม่หมุนเวียน', '', '', '', '', '', '', '', '']);
    currentRow++;

    if (landBuildingsEquipment !== 0) {
      worksheetData.push(['', '', 'ที่ดิน อาคาร และอุปกรณ์ (สุทธิ)', '', '', '11', landBuildingsEquipment, '', processingType === 'multi-year' ? prevLandBuildingsEquipment : '', '']);
      nonCurrentAssetRows.push(currentRow);
      currentRow++;
    }

    if (otherAssets !== 0) {
      worksheetData.push(['', '', 'สินทรัพย์อื่น', '', '', '12', otherAssets, '', processingType === 'multi-year' ? prevOtherAssets : '', '']);
      nonCurrentAssetRows.push(currentRow);
      currentRow++;
    }

    // Non-Current Assets Total
    const nonCurrentAssetsFormula = this.buildSumFormula(nonCurrentAssetRows, 'G');
    const nonCurrentAssetsFormulaPrev = processingType === 'multi-year' ? this.buildSumFormula(nonCurrentAssetRows, 'I') : '';
    
    worksheetData.push(['', 'รวมสินทรัพย์ไม่หมุนเวียน', '', '', '', '', 
      { f: nonCurrentAssetsFormula }, 
      '', 
      processingType === 'multi-year' ? { f: nonCurrentAssetsFormulaPrev } : '', '']);
    const nonCurrentAssetsTotalRow = currentRow;
    currentRow++;

    // Total Assets
    const totalAssetsFormula = `G${currentAssetsTotalRow}+G${nonCurrentAssetsTotalRow}`;
    const totalAssetsFormulaPrev = processingType === 'multi-year' ? `I${currentAssetsTotalRow}+I${nonCurrentAssetsTotalRow}` : '';
    
    worksheetData.push(['', 'รวมสินทรัพย์', '', '', '', '', 
      { f: totalAssetsFormula }, 
      '', 
      processingType === 'multi-year' ? { f: totalAssetsFormulaPrev } : '', '']);

    // Add footer
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    worksheetData.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '', '']);

    return worksheetData;
  }

  private generateBalanceSheetLiabilities(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year'
  ) {
    // Extract global financial data once for consistency across all statements
    const globalData = this.extractAllFinancialData(trialBalanceData, companyInfo);
    console.log('=== BALANCE SHEET LIABILITIES: Using Global Data Extraction ===');
    
    // Company type and terminology determination
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    const liabilityAndEquityTerm = isLimitedPartnership ? 
      'หนี้สินและส่วนของผู้เป็นหุ้นส่วน' : 'หนี้สินและส่วนของผู้ถือหุ้น';
    const equityTerm = isLimitedPartnership ? 
      'ส่วนของผู้เป็นหุ้นส่วน' : 'ส่วนของผู้ถือหุ้น';

    // Use global data for current year - eliminates redundant calculations
    const bankOverdraftsAndShortTermLoans = globalData.liabilities.bankOverdraftsAndShortTermLoans.current;
    const tradeAndOtherPayables = globalData.liabilities.tradeAndOtherPayables.current;
    const shortTermBorrowings = globalData.liabilities.shortTermBorrowings.current;
    const incomeTaxPayable = globalData.liabilities.incomeTaxPayable.current;
    const longTermLoansFromFI = globalData.liabilities.longTermLoansFromFI.current;
    const otherLongTermLoans = globalData.liabilities.otherLongTermLoans.current;
    const paidUpCapital = globalData.equity.paidUpCapital.current;
    const retainedEarnings = globalData.equity.retainedEarnings.current;
    const legalReserve = globalData.equity.legalReserve.current;

    // Use global data for previous year - eliminates redundant calculations
    const prevBankOverdraftsAndShortTermLoans = globalData.liabilities.bankOverdraftsAndShortTermLoans.previous;
    const prevTradeAndOtherPayables = globalData.liabilities.tradeAndOtherPayables.previous;
    const prevShortTermBorrowings = globalData.liabilities.shortTermBorrowings.previous;
    const prevIncomeTaxPayable = globalData.liabilities.incomeTaxPayable.previous;
    const prevLongTermLoansFromFI = globalData.liabilities.longTermLoansFromFI.previous;
    const prevOtherLongTermLoans = globalData.liabilities.otherLongTermLoans.previous;
    const prevPaidUpCapital = globalData.equity.paidUpCapital.previous;
    const prevRetainedEarnings = globalData.equity.retainedEarnings.previous;
    const prevLegalReserve = globalData.equity.legalReserve.previous;

    // Handle registered capital separately (not in global data yet)
    const registeredCapital = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3000, 3009));
    const prevRegisteredCapital = this.sumPreviousBalanceByNumericRange(trialBalanceData, 3000, 3009);

    console.log('=== BALANCE SHEET LIABILITIES: Values Summary ===');
    console.log('Current Year Values:', {
      bankOverdraftsAndShortTermLoans,
      tradeAndOtherPayables,
      shortTermBorrowings,
      incomeTaxPayable,
      longTermLoansFromFI,
      otherLongTermLoans,
      registeredCapital,
      paidUpCapital,
      retainedEarnings,
      legalReserve
    });
    console.log('Previous Year Values:', {
      prevBankOverdraftsAndShortTermLoans,
      prevTradeAndOtherPayables,
      prevShortTermBorrowings,
      prevIncomeTaxPayable,
      prevLongTermLoansFromFI,
      prevOtherLongTermLoans,
      prevRegisteredCapital,
      prevPaidUpCapital,
      prevRetainedEarnings,
      prevLegalReserve
    });

    // Initialize worksheet data with headers
    const worksheetData: (string | number | {f: string})[][] = [
      [companyInfo.name, '', '', '', '', '', '', '', '', ''],
      ['งบแสดงฐานะการเงิน (ต่อ)', '', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', liabilityAndEquityTerm, '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '', '']
    ];

    // Initialize cell tracker for VBA-compliant formula generation
    const cellTracker: CellTracker = {
      currentRow: worksheetData.length + 1,
      currentLiabilitiesRows: [],
      nonCurrentLiabilitiesRows: [],
      equityDataRows: [],
      currentLiabilitiesTotalRow: 0,
      nonCurrentLiabilitiesTotalRow: 0
    };

    // Build sections using the helper methods with cell tracking
    this.buildCurrentLiabilitiesSection(worksheetData, cellTracker, 
      bankOverdraftsAndShortTermLoans, tradeAndOtherPayables, shortTermBorrowings, incomeTaxPayable,
      prevBankOverdraftsAndShortTermLoans, prevTradeAndOtherPayables, prevShortTermBorrowings, prevIncomeTaxPayable, processingType);
    
    this.buildNonCurrentLiabilitiesSection(worksheetData, cellTracker, 
      longTermLoansFromFI, otherLongTermLoans, prevLongTermLoansFromFI, prevOtherLongTermLoans, processingType);
    
    this.buildEquitySection(worksheetData, cellTracker, 
      isLimitedPartnership, equityTerm, registeredCapital, paidUpCapital, retainedEarnings, legalReserve,
      prevPaidUpCapital, prevRetainedEarnings, companyInfo, processingType);

    // Add footer rows
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    worksheetData.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '', '']);

    return worksheetData;
  }

  // ============================================================================
  // BALANCE SHEET SECTION BUILDERS (WITH CELL TRACKING)
  // ============================================================================

  private buildCurrentLiabilitiesSection(
    worksheetData: any[][],
    cellTracker: CellTracker,
    bankOverdraftsAndShortTermLoans: number,
    tradeAndOtherPayables: number,
    shortTermBorrowings: number,
    incomeTaxPayable: number,
    prevBankOverdrafts: number,
    prevTradeAndOtherPayables: number,
    prevShortTermBorrowings: number,
    prevIncomeTaxPayable: number,
    processingType: 'single-year' | 'multi-year'
  ) {
    // Current Liabilities header
    worksheetData.push(['', 'หนี้สินหมุนเวียน', '', '', '', '', '', '', 'หน่วย:บาท', '']);
    cellTracker.currentRow++;

    // Add each current liability and track its row
    if (bankOverdraftsAndShortTermLoans !== 0) {
      worksheetData.push(['', '', 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน', '', '', '15', bankOverdraftsAndShortTermLoans, '', processingType === 'multi-year' ? prevBankOverdrafts : '', '']);
      cellTracker.currentLiabilitiesRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    // Always include trade payables - VBA always shows this
    worksheetData.push(['', '', 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '16', tradeAndOtherPayables, '', processingType === 'multi-year' ? prevTradeAndOtherPayables : '', '']);
    cellTracker.currentLiabilitiesRows.push(cellTracker.currentRow);
    cellTracker.currentRow++;

    if (shortTermBorrowings !== 0) {
      worksheetData.push(['', '', 'เงินกู้ยืมระยะสั้น', '', '', '17', shortTermBorrowings, '', processingType === 'multi-year' ? prevShortTermBorrowings : '', '']);
      cellTracker.currentLiabilitiesRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    if (incomeTaxPayable !== 0) {
      worksheetData.push(['', '', 'ภาษีเงินได้นิติบุคคลค้างจ่าย', '', '', '18', incomeTaxPayable, '', processingType === 'multi-year' ? prevIncomeTaxPayable : '', '']);
      cellTracker.currentLiabilitiesRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    // Current Liabilities Total using tracked rows
    const currentLiabilitiesFormula = this.buildSumFormula(cellTracker.currentLiabilitiesRows, 'G');
    const currentLiabilitiesFormulaPrev = processingType === 'multi-year' ? this.buildSumFormula(cellTracker.currentLiabilitiesRows, 'I') : '';
    
    worksheetData.push(['', 'รวมหนี้สินหมุนเวียน', '', '', '', '', 
      { f: currentLiabilitiesFormula }, 
      '', 
      processingType === 'multi-year' ? { f: currentLiabilitiesFormulaPrev } : '', '']);
    
    // Track the current liabilities total row
    cellTracker.currentLiabilitiesTotalRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    // No spacer - directly continue with non-current liabilities
  }

  private buildNonCurrentLiabilitiesSection(
    worksheetData: any[][],
    cellTracker: CellTracker,
    longTermLoansFromFI: number,
    otherLongTermLoans: number,
    prevLongTermLoansFromFI: number,
    prevOtherLongTermLoans: number,
    processingType: 'single-year' | 'multi-year'
  ) {
    // Non-Current Liabilities header
    worksheetData.push(['', 'หนี้สินไม่หมุนเวียน', '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;

    // Add each non-current liability and track its row
    if (longTermLoansFromFI !== 0) {
      worksheetData.push(['', '', 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน', '', '', '19', longTermLoansFromFI, '', processingType === 'multi-year' ? prevLongTermLoansFromFI : '', '']);
      cellTracker.nonCurrentLiabilitiesRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    if (otherLongTermLoans !== 0) {
      worksheetData.push(['', '', 'เงินกู้ยืมระยะยาวอื่น', '', '', '20', otherLongTermLoans, '', processingType === 'multi-year' ? prevOtherLongTermLoans : '', '']);
      cellTracker.nonCurrentLiabilitiesRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    // Non-Current Liabilities Total using tracked rows
    const nonCurrentLiabilitiesFormula = this.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'G');
    const nonCurrentLiabilitiesFormulaPrev = processingType === 'multi-year' ? this.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'I') : '';
    
    worksheetData.push(['', 'รวมหนี้สินไม่หมุนเวียน', '', '', '', '', 
      { f: nonCurrentLiabilitiesFormula }, 
      '', 
      processingType === 'multi-year' ? { f: nonCurrentLiabilitiesFormulaPrev } : '', '']);
    
    // Track the non-current liabilities total row
    cellTracker.nonCurrentLiabilitiesTotalRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    // Total Liabilities - VBA-compliant: simple addition of two subtotal rows
    const totalLiabilitiesFormula = cellTracker.currentLiabilitiesTotalRow && cellTracker.nonCurrentLiabilitiesTotalRow 
      ? `G${cellTracker.currentLiabilitiesTotalRow}+G${cellTracker.nonCurrentLiabilitiesTotalRow}`
      : '0';
    const totalLiabilitiesFormulaPrev = processingType === 'multi-year' && cellTracker.currentLiabilitiesTotalRow && cellTracker.nonCurrentLiabilitiesTotalRow
      ? `I${cellTracker.currentLiabilitiesTotalRow}+I${cellTracker.nonCurrentLiabilitiesTotalRow}`
      : '';
    
    worksheetData.push(['', 'รวมหนี้สิน', '', '', '', '', 
      { f: totalLiabilitiesFormula }, 
      '', 
      processingType === 'multi-year' ? { f: totalLiabilitiesFormulaPrev } : '', '']);
    cellTracker.totalLiabilitiesRow = cellTracker.currentRow; // Track total liabilities row
    cellTracker.currentRow++;

    // Spacer
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;
  }

  private buildEquitySection(
    worksheetData: any[][],
    cellTracker: CellTracker,
    isLimitedPartnership: boolean,
    equityTerm: string,
    registeredCapital: number,
    paidUpCapital: number,
    retainedEarnings: number,
    legalReserve: number,
    prevPaidUpCapital: number,
    prevRetainedEarnings: number,
    _companyInfo: any,
    processingType: 'single-year' | 'multi-year'
  ) {
    // Get share information from company info
    const numberOfShares = _companyInfo?.shares || registeredCapital; // fallback to registeredCapital
    const shareValue = _companyInfo?.shareValue || 1; // fallback to 1 baht per share
    const numberOfPaidShares = paidUpCapital / shareValue; // calculate based on share value
    
    // Equity Section header
    worksheetData.push(['', equityTerm, '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;

    if (isLimitedPartnership) {
      // Partnership equity structure
      worksheetData.push(['', '', 'เงินลงทุนของผู้เป็นหุ้นส่วน คนที่ 1', '', '', '21', paidUpCapital / 2, '', processingType === 'multi-year' ? prevPaidUpCapital / 2 : '', '']);
      cellTracker.equityDataRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;

      worksheetData.push(['', '', 'เงินลงทุนของผู้เป็นหุ้นส่วน คนที่ 2', '', '', '22', paidUpCapital / 2, '', processingType === 'multi-year' ? prevPaidUpCapital / 2 : '', '']);
      cellTracker.equityDataRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;

      worksheetData.push(['', '', 'กำไรสะสม', '', '', '23', retainedEarnings, '', processingType === 'multi-year' ? prevRetainedEarnings : '', '']);
      cellTracker.equityDataRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    } else {
      // Limited company equity structure - Remove ทุนเรือนหุ้น row

      worksheetData.push(['', '', 'ทุนจดทะเบียน', '', '', '', '', '', '', '']);
      worksheetData.push(['', '', '', `หุ้นสามัญ ${numberOfShares.toLocaleString()} หุ้น มูลค่าหุ้นละ ${shareValue} บาท`, '', '', registeredCapital, '', processingType === 'multi-year' ? registeredCapital : '', '']);
      cellTracker.currentRow += 2;

      worksheetData.push(['', '', 'ทุนที่ออกและชำระแล้ว', '', '', '', '', '', '', '']);
      worksheetData.push(['', '', '', `หุ้นสามัญ ${numberOfPaidShares.toLocaleString()} หุ้น มูลค่าหุ้นละ ${shareValue} บาท`, '', '24', paidUpCapital, '', processingType === 'multi-year' ? prevPaidUpCapital : '', '']);
      cellTracker.equityDataRows.push(cellTracker.currentRow + 1);
      cellTracker.currentRow += 2;

      if (legalReserve !== 0) {
        worksheetData.push(['', '', 'ทุนสำรองตามกฎหมาย', '', '', '25', legalReserve, '', processingType === 'multi-year' ? legalReserve : '', '']);
        cellTracker.equityDataRows.push(cellTracker.currentRow);
        cellTracker.currentRow++;
      }

      worksheetData.push(['', '', 'กำไรสะสม', '', '', '26', retainedEarnings, '', processingType === 'multi-year' ? prevRetainedEarnings : '', '']);
      cellTracker.equityDataRows.push(cellTracker.currentRow);
      cellTracker.currentRow++;
    }

    // Total Equity using tracked rows
    const totalEquityFormula = this.buildSumFormula(cellTracker.equityDataRows, 'G');
    const totalEquityFormulaPrev = processingType === 'multi-year' ? this.buildSumFormula(cellTracker.equityDataRows, 'I') : '';
    
    worksheetData.push(['', `รวม${equityTerm}`, '', '', '', '', 
      { f: totalEquityFormula }, 
      '', 
      processingType === 'multi-year' ? { f: totalEquityFormulaPrev } : '', '']);
    const totalEquityRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    // Grand Total (Liabilities + Equity) - VBA compliant: simple addition of totals
    // Use the tracked row numbers for accurate calculation
    const grandTotalFormula = cellTracker.totalLiabilitiesRow 
      ? `G${cellTracker.totalLiabilitiesRow}+G${totalEquityRow}`
      : `G${totalEquityRow - cellTracker.equityDataRows.length - 2}+G${totalEquityRow}`; // Fallback to old calculation
    const grandTotalFormulaPrev = processingType === 'multi-year' 
      ? (cellTracker.totalLiabilitiesRow 
          ? `I${cellTracker.totalLiabilitiesRow}+I${totalEquityRow}`
          : `I${totalEquityRow - cellTracker.equityDataRows.length - 2}+I${totalEquityRow}`)
      : '';
    
    console.log('=== GRAND TOTAL DEBUG ===');
    console.log('Total Liabilities Row:', cellTracker.totalLiabilitiesRow);
    console.log('Total Equity Row:', totalEquityRow);
    console.log('Grand Total Formula:', grandTotalFormula);
    
    worksheetData.push(['', `รวม${'หนี้สินและส่วนของผู้ถือหุ้น'}`, '', '', '', '', 
      { f: grandTotalFormula }, 
      '', 
      processingType === 'multi-year' ? { f: grandTotalFormulaPrev } : '', '']);
  }

  // ============================================================================
  // OTHER FINANCIAL STATEMENT GENERATION METHODS
  // ============================================================================
  
  private generateProfitLossStatement(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): any[][] {
    console.log('=== START P&L DEBUG ===');
    console.log('Trial Balance Data:', trialBalanceData);
    
    // Calculate revenue (4000-4099)
    const revenue = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4000, 4099));
    const previousRevenue = processingType === 'multi-year' ? 0 : 0; // Placeholder for previous year
    
    // Calculate other income (4100-4999)
    const otherIncome = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4100, 4999));
    const previousOtherIncome = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate cost of services/goods sold (5000-5099)
    const costOfServices = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5000, 5099));
    const previousCostOfServices = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate administrative expenses (5300-5350 + selected ranges)
    const adminExpenses = Math.abs(
      this.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
      this.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
      this.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
      this.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
    );
    const previousAdminExpenses = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate other expenses (5351-5354, 5358-5361, 5364, 5366-5999)
    const otherExpenses = Math.abs(
      this.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
      this.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
      this.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
      this.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
    );
    const previousOtherExpenses = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate income tax (5910)
    const incomeTax = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5910, 5910));
    const previousIncomeTax = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate financial costs (5920-5929)
    const financialCosts = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5920, 5929));
    const previousFinancialCosts = processingType === 'multi-year' ? 0 : 0;
    
    console.log('Calculated values:', {
      revenue,
      otherIncome,
      costOfServices,
      adminExpenses,
      otherExpenses,
      incomeTax
    });
    console.log('=== END P&L DEBUG ===');
    
    return [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['งบกำไรขาดทุน จำแนกค่าใช้จ่ายตามหน้าที่ - แบบขั้นเดียว', '', '', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ 1 มกราคม ${companyInfo.reportingYear} ถึงวันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', 'หมายเหตุ', '', '', 'หน่วย:บาท'],
      ['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : ''],
      ['', 'รายได้', '', '', '', '', '', '', ''],
      ['', '', 'รายได้จากการขายหรือการให้บริการ', '', '', '1', revenue, '', processingType === 'multi-year' ? previousRevenue : ''],
      ['', '', 'รายได้อื่น', '', '', '2', otherIncome, '', processingType === 'multi-year' ? previousOtherIncome : ''],
      ['', 'รวมรายได้', '', '', '', '', { f: 'SUM(G8:G9)' }, '', processingType === 'multi-year' ? { f: 'SUM(I8:I9)' } : ''],
      ['', '', '', '', '', '', '', '', ''],
      ['', 'ค่าใช้จ่าย', '', '', '', '', '', '', ''],
      ['', '', 'ต้นทุนขายหรือต้นทุนการให้บริการ', '', '', '3', costOfServices, '', processingType === 'multi-year' ? previousCostOfServices : ''],
      ['', '', 'ค่าใช้จ่ายในการบริหาร', '', '', '4', adminExpenses, '', processingType === 'multi-year' ? previousAdminExpenses : ''],
      ['', '', 'ค่าใช้จ่ายอื่น', '', '', '5', otherExpenses, '', processingType === 'multi-year' ? previousOtherExpenses : ''],
      ['', 'รวมค่าใช้จ่าย', '', '', '', '', { f: 'SUM(G13:G15)' }, '', processingType === 'multi-year' ? { f: 'SUM(I13:I15)' } : ''],
      ['', 'กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้', '', '', '', '', { f: 'G10-G16' }, '', processingType === 'multi-year' ? { f: 'I10-I16' } : ''],
      ['', 'ต้นทุนทางการเงิน', '', '', '', '7', financialCosts, '', processingType === 'multi-year' ? previousFinancialCosts : ''],
      ['', 'กำไรก่อนภาษีเงินได้', '', '', '', '', { f: 'G17-G18' }, '', processingType === 'multi-year' ? { f: 'I17-I18' } : ''],
      ['', 'ภาษีเงินได้', '', '', '', '6', incomeTax, '', processingType === 'multi-year' ? previousIncomeTax : ''],
      ['', 'กำไร(ขาดทุน)สุทธิ', '', '', '', '', { f: 'G19-G20' }, '', processingType === 'multi-year' ? { f: 'I19-I20' } : ''],
      ['', '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '']
    ];
  }

  private generateStatementOfChangesInEquity(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): any[][] {
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    
    if (isLimitedPartnership) {
      return this.generatePartnershipEquityStatement(trialBalanceData, companyInfo, processingType);
    } else {
      return this.generateCorporateEquityStatement(trialBalanceData, companyInfo, processingType);
    }
  }

  private generatePartnershipEquityStatement(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    _processingType: 'single-year' | 'multi-year'
  ): any[][] {
    // Calculate partnership equity values
    const totalCapital = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3010, 3019));
    
    // CORRECTED RETAINED EARNINGS CALCULATION for Partnership
    // 1. Get opening retained earnings from account 3020 only
    // Note: Retained earnings typically has credit balance in trial balance, so we flip the sign
    const openingRetainedEarningsRaw = this.sumAccountsByNumericRange(trialBalanceData, 3020, 3020);
    const openingRetainedEarnings = -openingRetainedEarningsRaw; // Flip sign for credit balance
    console.log('=== PARTNERSHIP RETAINED EARNINGS DEBUG ===');
    console.log('1. Opening Retained Earnings (3020):');
    console.log(`   Raw balance from trial balance: ${openingRetainedEarningsRaw}`);
    console.log(`   Adjusted for credit balance: ${openingRetainedEarnings}`);
    
    // 2. Calculate current year revenue (4xxxx accounts: credit - debit)
    const revenueAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('4'));
    console.log('2. Revenue accounts found:', revenueAccounts.map(acc => ({
      code: acc.accountCode,
      name: acc.accountName,
      debit: acc.debitAmount,
      credit: acc.creditAmount,
      calculation: (acc.creditAmount || 0) - (acc.debitAmount || 0)
    })));
    
    const currentYearRevenue = revenueAccounts
      .reduce((sum, entry) => sum + ((entry.creditAmount || 0) - (entry.debitAmount || 0)), 0);
    console.log('2. Total Current Year Revenue:', currentYearRevenue);
    
    // 3. Calculate current year expenses (5xxxx accounts: debit - credit)  
    const expenseAccounts = trialBalanceData.filter(entry => entry.accountCode?.startsWith('5'));
    console.log('3. Expense accounts found:', expenseAccounts.map(acc => ({
      code: acc.accountCode,
      name: acc.accountName,
      debit: acc.debitAmount,
      credit: acc.creditAmount,
      calculation: (acc.debitAmount || 0) - (acc.creditAmount || 0)
    })));
    
    const currentYearExpenses = expenseAccounts
      .reduce((sum, entry) => sum + ((entry.debitAmount || 0) - (entry.creditAmount || 0)), 0);
    console.log('3. Total Current Year Expenses:', currentYearExpenses);
    
    // 4. Calculate current year profit
    const currentYearProfit = currentYearRevenue - currentYearExpenses;
    console.log('4. Current Year Profit (Revenue - Expenses):', currentYearProfit);
    
    // 5. Final retained earnings = opening + current year profit
    const retainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);
    console.log('5. Final Retained Earnings (Opening + Profit):', openingRetainedEarnings, '+', currentYearProfit, '=', retainedEarnings);
    console.log('=== END PARTNERSHIP RETAINED EARNINGS DEBUG ===');
    
    // Split capital equally between partners
    const partner1Capital = totalCapital / 2;
    const partner2Capital = totalCapital / 2;
    
    return [
      [`${companyInfo.name}`, '', '', '', '', '', ''],
      ['งบแสดงการเปลี่ยนแปลงส่วนของผู้เป็นหุ้นส่วน', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี สิ้นสุด วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['', 'ผู้เป็นหุ้นส่วน คนที่ 1', 'ผู้เป็นหุ้นส่วน คนที่ 2', 'กำไรสะสม', 'รวม', '', ''],
      ['ยอดคงเหลือ ณ วันต้นปี', partner1Capital, partner2Capital, retainedEarnings, { f: 'B7+C7+D7' }, '', ''],
      ['กำไรสุทธิสำหรับปี', '', '', currentYearProfit, currentYearProfit, '', ''],
      ['ยอดคงเหลือ ณ วันสิ้นปี', { f: 'B7+B8' }, { f: 'C7+C8' }, { f: 'D7+D8' }, { f: 'B9+C9+D9' }, '', ''],
      ['', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '']
    ];
  }

  private generateCorporateEquityStatement(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): any[][] {
    const isMultiYear = processingType === 'multi-year';
    const currentYear = companyInfo.reportingYear;
    const previousYear = currentYear - 1;
    
    // *** USE GLOBAL DATA - NO MORE REDUNDANT CALCULATIONS! ***
    const globalData = this.extractedData!; // Already extracted in main method
    
    console.log('=== CHANGES IN EQUITY USING GLOBAL DATA ===');
    console.log('Paid-up Capital (Global):', globalData.equity.paidUpCapital);
    console.log('Retained Earnings (Global):', globalData.equity.retainedEarnings);
    console.log('Opening Retained Earnings (Global):', globalData.equity.openingRetainedEarnings);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
    // Use global values directly - consistent across all statements!
    const paidUpCapitalCurrent = globalData.equity.paidUpCapital.current;
    const paidUpCapitalPrevious = globalData.equity.paidUpCapital.previous;
    const retainedEarningsCurrent = globalData.equity.retainedEarnings.current;
    const openingRetainedEarnings = globalData.equity.openingRetainedEarnings;
    const currentYearProfit = globalData.income.netProfit;
    
    const result: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['งบแสดงการเปลี่ยนแปลงส่วนของผู้ถือหุ้น', '', '', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี สิ้นสุด วันที่ 31 ธันวาคม ${currentYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
      ['', '', 'ทุนเรือนหุ้นที่ออกและชำระแล้ว', '', '', 'กำไร(ขาดทุน)สะสม', '', '', 'รวม'],
      ['', '', '', '', '', '', '', '', '']
    ];

    let rowIndex = 7;

    if (isMultiYear) {
      // F8, F9 - Leave blank if no previous year data in trial balance
      const prevYearOpeningRetained = ''; // As requested - leave blank
      const prevYearProfit = ''; // As requested - leave blank
      const prevYearTotalEquity = paidUpCapitalPrevious > 0 ? paidUpCapitalPrevious + (openingRetainedEarnings || 0) : '';
      
      result.push([`ยอดคงเหลือ ณ วันที่ 1 มกราคม ${previousYear}`, '', paidUpCapitalPrevious || '', '', '', prevYearOpeningRetained, '', '', prevYearTotalEquity]);
      result.push([`กำไร (ขาดทุน) สุทธิ สำหรับปี ${previousYear}`, '', '', '', '', prevYearProfit, '', '', '']);
      result.push([`ยอดคงเหลือ ณ วันที่ 31 ธันวาคม ${previousYear}`, '', { f: 'C8+C9' }, '', '', { f: 'F8+F9' }, '', '', { f: 'C10+F10' }]);
      result.push(['', '', '', '', '', '', '', '', '']);
      result.push(['', '', '', '', '', '', '', '', '']);
      rowIndex = 12;
    }

    // Current year section
    // F10 should be openingRetainedEarnings (from account 3020 credit-debit)
    const openingTotalCurrent = paidUpCapitalCurrent + openingRetainedEarnings;
    
    result.push([`ยอดคงเหลือ ณ วันที่ 1 มกราคม ${currentYear}`, '', paidUpCapitalCurrent, '', '', openingRetainedEarnings, '', '', openingTotalCurrent]);
    result.push([`กำไร (ขาดทุน) สุทธิ สำหรับปี ${currentYear}`, '', '', '', '', currentYearProfit, '', '', currentYearProfit]);
    result.push([`ยอดคงเหลือ ณ วันที่ 31 ธันวาคม ${currentYear}`, '', paidUpCapitalCurrent, '', '', retainedEarningsCurrent, '', '', { f: `C${rowIndex + 3}+F${rowIndex + 3}` }]);
    
    result.push(['', '', '', '', '', '', '', '', '']);
    result.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '']);
    
    return result;
  }

  private getSingleAccountBalance(trialBalanceData: TrialBalanceEntry[], accountCode: string): number {
    const account = trialBalanceData.find(entry => entry.accountCode === accountCode);
    // Use balance field if currentBalance is 0, taking absolute value for equity accounts
    const value = account ? Math.abs(account.currentBalance || account.balance || 0) : 0;
    return value;
  }

  private generateNotesToFinancialStatements(
    companyInfo: CompanyInfo, 
    trialBalanceData?: TrialBalanceEntry[], 
    processingType?: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[]
  ): (string | number | {f: string})[][] {
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    const entityType = isLimitedPartnership ? 'ห้างหุ้นส่วนจำกัด' : 'บริษัทจำกัด';
    const shortEntityName = isLimitedPartnership ? 'ห้างฯ' : 'บริษัทฯ';
    
    const worksheetData: (string | number | {f: string})[][] = [
      // Header section
      [`${companyInfo.name}`, '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ 1 มกราคม ${companyInfo.reportingYear} ถึงวันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', ''],
      ['', '', '', '', '', ''],
      
      // Note 1: ข้อมูลทั่วไป
      ['1', 'ข้อมูลทั่วไป', '', '', '', ''],
      ['', '', '1.1', `สถานะทางกฎหมาย เป็น${entityType}จดทะเบียนในประเทศไทย`, '', ''],
      ['', '', '1.2', `สถานที่ตั้ง ${companyInfo.address || '6 ถนนลาซาล 4 ซ.สุรศักดิ์ ต.สีลม อ.บางรัก'}`, '', ''],
      ['', '', '1.3', `ลักษณะธุรกิจและการดำเนินงาน ${companyInfo.businessDescription || 'ประกอบธุรกิจเกี่ยวกับเทคโนโลยีและบริการด้วยซอฟต์แวร์'}`, '', ''],
      ['', '', '', '', '', ''],
      
      // Note 2: ฐานะการดำเนินงานของบริษัท
      ['2', 'ฐานะการดำเนินงานของบริษัท', '', '', '', ''],
      ['', '', `${companyInfo.name} ได้จดทะเบียนตามประมวลกฎหมายแพ่งและพาณิชย์เป็นนิติบุคคล ประเภท ${entityType} เมื่อวันที่ 17 ก.ค. 2563 ทะเบียนเลขที่ 0335563000688`, '', '', '', ''],
      ['', '', '', '', '', ''],
      
      // Note 3: เกณฑ์การจัดทำงบการเงินของกิจการ
      ['3', 'เกณฑ์การจัดทำงบการเงินของกิจการ', '', '', '', ''],
      ['', '', `${shortEntityName}จัดทำงบการเงินนี้ตามมาตรฐานการรายงานทางการเงินสำหรับกิจการที่ไม่มีส่วนได้เสียสาธารณะ ตามมาตรฐานการบัญชี ฉบับที่ 48/2565 ลงวันที่ 14 กรกฎาคม 2565 เรื่องมาตรฐานการรายงานทางการเงิน ป.ป. 2547 และประกาศสภาวิชาชีพบัญชี เรื่องการใช้มาตรฐานการรายงานทางการเงิน ลงวันที่ 27 ธันวาคม 2566 เป็นต้น`, '', '', '', ''],
      ['', '', 'งบการเงินได้จัดทำขึ้น ตามหลักการคำนวณต้นทุนในการจัดทำขึ้นของกิจการของเงินกินค้างงบการเงิน ยกเว้นบางรายการที่เปิดเผยไว้ในนโยบายการบัญชีที่เกี่ยวข้อง', '', '', '', ''],
      ['', '', '', '', '', ''],
      
      // Note 4: นโยบายการบัญชีที่สำคัญ
      ['4', 'นโยบายการบัญชีที่สำคัญ', '', '', '', ''],
      ['4.1', 'การรับรู้รายได้และค่าใช้จ่าย', '', '', '', ''],
      ['', 'การรับรู้รายได้จากสินค้า', '', '', '', ''],
      ['', 'การรับรู้รายได้จากสินค้า รับรู้เมื่อกิจการได้โอนการควบคุมในสินค้าให้แก่ลูกค้าและสามารถวัดมูลค่าได้อย่างน่าเชื่อถือ กิจการต้องรับรู้รายได้จากการควบคุมในสินค้าให้แก่ลูกค้าในจุดเวลาใดจุดเวลาหนึ่ง โดยทั่วไปคือเมื่อได้จัดส่งสินค้าให้แก่ลูกค้าแล้ว', '', '', '', ''],
      ['', 'การรับรู้รายได้จากการให้บริการ', '', '', '', ''],
      ['', 'รายได้จากการให้บริการ รับรู้เมื่อได้ให้บริการผลตอบแทนไม่ชำระค่าความรับผิดชอบของกิจการต่อผู้รับบริการ', '', '', '', ''],
      ['', 'การรับรู้รายได้และค่าใช้จ่ายอื่น', '', '', '', ''],
      ['', 'รายได้และค่าใช้จ่ายอื่น รับรู้ตามหลักการคงค้าง', '', '', '', ''],
      ['4.2', 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', ''],
      ['', 'เงินสด ครอบคลุม เงินสดในมือและเงินฝากธนาคารที่ต้องการแคชชมน์ รายการเทียบเท่าเงินสด ครอบคลุม เงินลงทุนระยะสั้นที่มีสภาพคล่องสูง ซึ่งถึงกำหนดไม่เกิน 3 เดือนนับจากวันที่ได้มาและไม่มีข้อจำกัดในการเบิกใช้เงินสดในจำนวนที่สำคัญ', '', '', '', ''],
      ['4.3', 'ลูกหนี้การค้าและลูกหนี้อื่น ๆ', '', '', '', ''],
      ['', 'ลูกหนี้การค้าและลูกหนี้อื่น ๆ ครอบคลุม ลูกหนี้ที่เกิดขึ้นจากการขายสินค้าหรือการให้บริการในสายงานปกติของกิจการ และรอคอยชำระหนี้ ลูกหนี้อาจรวมถึงลูกหนี้การค้าและลูกหนี้อื่น ๆ', '', '', '', '']
    ];

    // Add all accounting policies
    worksheetData.push(
      ['4.4', 'สินค้าคงเหลือ ( Inventories )', '', '', '', ''],
      ['', 'หมายถึง สินค้าสำเร็จรูป งานหรือสินค้าระหว่างทำ วัตถุดิบและวัสดุที่ใช้ในการผลิตเพื่อขาย หรือให้บริการตามปกติของกิจการ ซึ่งรวมถึง ที่ดินรอการพัฒนาและต้นทุนระหว่างพัฒนาอสังหาริมทรัพย์และงานระหว่างก่อสร้าง ทั้งนี้ ให้เป็นไปตามมาตรฐาน การรายงานทางการเงินที่เกี่ยวข้อง', '', '', '', ''],
      
      ['4.5', 'ที่ดิน อาคารและอุปกรณ์', '', '', '', ''],
      ['', 'กิจการวัดมูลค่าของรายการที่ดิน อาคารและอุปกรณ์ที่เข้าเงื่อนไขโดยใช้ราคาทุน กิจการวัดมูลค่าภายหลังการรับรู้รายการ โดยใช้วิธีราคาทุน โดยแสดงรายการที่ดิน อาคารและอุปกรณ์ ด้วยราคาทุนหักค่าเสื่อมราคาสะสม และค่าเผื่อการลดลงของมูลค่า (ถ้ามี) กิจการคิดค่าเสื่อมราคา ด้วยวิธีเส้นตรง และ ด้วยวิธียอดคงเหลือลดลง', '', '', '', ''],
      
      ['4.6', 'สัญญาเช่าดำเนินงาน', '', '', '', ''],
      ['', 'สัญญาเช่าซึ่งความเสี่ยงและประโยชน์ส่วนใหญ่ จากการเป็นเจ้าของทรัพย์สินยังคงอยู่กับผู้ให้เช่า บันทึกเป็นสัญญาเช่า ดำเนินงาน ค่าเช่าที่เกิดขึ้นจากสัญญาเช่าดังกล่าว รับรู้เป็นค่าใช้จ่ายในงบกำไรขาดทุน ตามอายุของสัญญาเช่า', '', '', '', ''],
      
      ['4.7', 'ประมาณการทางบัญชี', '', '', '', ''],
      ['', 'ในการจัดทำงบการเงินตามมาตรฐานการรายงานทางการเงินสำหรับกิจการที่ไม่มีส่วนได้เสียสาธารณะ ฝ่ายบริหารต้องใช้ ดุลยพินิจในการกำหนดนโยบายบัญชี การประมาณการและตั้งข้อสมมติฐานหลายประการ ซึ่งจะมีผลกระทบต่อจำนวนเงิน ผลที่เกิดขึ้นจริงอาจมีความแตกต่างไปจากจำนวนที่ประมาณไว้', '', '', '', ''],
      
      ['4.8', 'ผลประโยชน์พนักงาน', '', '', '', ''],
      ['', 'ผลประโยชน์พนักงานจะรับรู้เมื่อห้าง มีภาระหนี้สินตามกฏหมายที่เกิดขึ้นในปัจจุบันอันเป็นผลมาจากเหตุการณ์ในอดีต และมีความเป็นไปได้ค่อนข้างแน่ว่า ประโยชน์เชิงเศรษฐกิจ จะต้องถูกจ่ายไปเพื่อชำระภาระหนี้สินดังกล่าว', '', '', '', ''],
      
      ['4.9', 'ภาษีเงินได้', '', '', '', ''],
      ['', 'กิจการคำนวณภาษีเงินได้ตามเกณฑ์ที่กำหนดไว้ในประมวลรัษฎากร และบันทึกภาษีเงินได้ตามเกณฑ์คงค้าง ห้างฯมิได้บันทึก ภาษีเงินได้ค้างจ่าย หรือค้างรับ สำหรับผลต่างทางบัญชีกับทางภาษี ที่เป็นผลต่างชั่วคราวในอนาคต', '', '', '', ''],
      
      ['5.1', 'การจัดประเภทรายการงบการเงิน', '', '', '', ''],
      ['', 'กิจการได้มีการจัดประเภทรายการทางการเงินใหม่เพื่อให้สอดคล้องกับมาตรฐานการบัญชีของสภาวิชาชีพบัญชี และเพื่อเป็นประโยชน์ต่อการตัดสินใจเชิงเศรษฐกิจของผู้ใช้งบการเงิน', '', '', '', '']
    );

    return worksheetData;
  }

  private generateAccountingNotes(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[]
  ): any[][] {
    // Extract global financial data once for consistency across all notes
    const globalData = this.extractAllFinancialData(trialBalanceData, companyInfo);
    console.log('=== NOTES_ACCOUNTING: Using Global Data Extraction ===');
    
    const notes: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน (ต่อ)', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
    ];

    let noteNumber = 3;
    
    // Generate specific notes using global data where possible
    this.addCashNoteWithGlobalData(notes, globalData, trialBalanceData, companyInfo, processingType, noteNumber++);
    this.addTradeReceivablesNoteWithGlobalData(notes, globalData, trialBalanceData, companyInfo, processingType, noteNumber++);
    this.addShortTermLoansNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addPPENote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addOtherAssetsNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addBankOverdraftsNoteWithGlobalData(notes, globalData, companyInfo, processingType, noteNumber++);
    this.addTradePayablesNoteWithGlobalData(notes, globalData, trialBalanceData, companyInfo, processingType, noteNumber++);
    this.addShortTermBorrowingsNoteWithGlobalData(notes, globalData, companyInfo, processingType, noteNumber++);
    this.addLongTermLoansNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addOtherLongTermLoansNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addRelatedPartyLoansNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addOtherIncomeNote(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    this.addExpensesByNatureNote(notes, companyInfo, processingType, noteNumber++);
    if (companyInfo.type === 'บริษัทจำกัด') {
      this.addFinancialApprovalNote(notes, companyInfo, noteNumber++);
    }

    return notes;
  }

  private generateDetailNotes(trialBalanceData: TrialBalanceEntry[], companyInfo: CompanyInfo): any[][] {
    // Extract global financial data once for consistency
    const globalData = this.extractAllFinancialData(trialBalanceData, companyInfo);
    console.log('=== NOTES_DETAIL: Using Global Data Extraction ===');
    
    const detailNotes: any[][] = [];
    
    // Add header
    detailNotes.push([`${companyInfo.name}`, '', '', '', '', '', '', '', '']);
    detailNotes.push(['รายละเอียดประกอบหมายเหตุประกอบงบการเงิน', '', '', '', '', '', '', '', '']);
    detailNotes.push([`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', '']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add DT1 - Cost of goods sold / Service costs
    this.addDetailOneWithGlobalData(detailNotes, trialBalanceData, globalData);
    
    // Add some spacing
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add DT2 - Selling and administrative expenses  
    this.addDetailTwoWithGlobalData(detailNotes, globalData);
    
    return detailNotes;
  }

  private addDetailOne(detailNotes: any[][], trialBalanceData: TrialBalanceEntry[]): void {
    const hasInventory = this.checkHasInventory(trialBalanceData);
    
    detailNotes.push(['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    if (!hasInventory) {
      // Service business
      detailNotes.push(['ต้นทุนการให้บริการ', '', '', '', '', '', '', '', '']);
      detailNotes.push(['', 'ค่าใช้จ่ายอื่นๆ ในการให้บริการ', '', '', '', '', '', '', '...']);
      detailNotes.push(['', 'รวม', '', '', '', '', '', '', { f: 'I' + (detailNotes.length) }]);
      return;
    }

    // Inventory-based business
    detailNotes.push(['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', '']);
    
    // Get beginning inventory
    const inventoryAccount = trialBalanceData.find(entry => entry.accountCode === '1510');
    const beginningInventory = inventoryAccount ? Math.abs(inventoryAccount.previousBalance || 0) : 0;
    
    detailNotes.push(['', 'สินค้าคงเหลือต้นงวด', '', '', '', '', '', '', beginningInventory]);
    
    // Process purchases and related accounts
    let totalPurchases = 0;
    
    // Main purchases (5010)
    const purchaseAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5010, 5010));
    if (purchaseAmount > 0) {
      detailNotes.push(['', 'บวก', 'ซื้อสินค้า', '', '', '', '', '', purchaseAmount]);
      totalPurchases += purchaseAmount;
    }
    
    // Purchase returns (5010.1) - debit, shown as negative
    const returnAmount = Math.abs(this.getAccountBalance(trialBalanceData, ['5010.1']));
    if (returnAmount > 0) {
      detailNotes.push(['', 'หัก', 'ส่งคืนสินค้า', '', '', '', '', '', returnAmount]);
      totalPurchases -= returnAmount;
    }
    
    // Purchase discounts (5010.2) - debit, shown as negative  
    const discountAmount = Math.abs(this.getAccountBalance(trialBalanceData, ['5010.2']));
    if (discountAmount > 0) {
      detailNotes.push(['', 'หัก', 'ส่วนลดรับ', '', '', '', '', '', discountAmount]);
      totalPurchases -= discountAmount;
    }
    
    // Transportation costs (5010.3) - debit, shown as positive
    const transportAmount = Math.abs(this.getAccountBalance(trialBalanceData, ['5010.3']));
    if (transportAmount > 0) {
      detailNotes.push(['', 'บวก', 'ค่าขนส่งเข้า', '', '', '', '', '', transportAmount]);
      totalPurchases += transportAmount;
    }
    
    // Direct labor (5010.4) - debit, shown as positive
    const laborAmount = Math.abs(this.getAccountBalance(trialBalanceData, ['5010.4']));
    if (laborAmount > 0) {
      detailNotes.push(['', '', 'ค่าแรงงานทางตรง', '', '', '', '', '', laborAmount]);
      totalPurchases += laborAmount;
    }
    
    // Net purchases total
    detailNotes.push(['', '', 'รวมซื้อสุทธิ', '', '', '', '', '', totalPurchases]);
    
    // Goods available for sale
    const goodsAvailable = beginningInventory + totalPurchases;
    detailNotes.push(['', 'สินค้ามีไว้เพื่อขาย', '', '', '', '', '', '', goodsAvailable]);
    
    // Ending inventory
    const endingInventory = inventoryAccount ? Math.abs(inventoryAccount.debitAmount - inventoryAccount.creditAmount) : 0;
    detailNotes.push(['', 'หัก', 'สินค้าคงเหลือปลายงวด', '', '', '', '', '', endingInventory]);
    
    // Cost of goods sold
    const costOfGoodsSold = goodsAvailable - endingInventory;
    detailNotes.push(['', 'ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', costOfGoodsSold]);
  }

  private addDetailTwo(detailNotes: any[][], trialBalanceData: TrialBalanceEntry[]): void {
    detailNotes.push(['รายละเอียดประกอบที่ 2', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Headers for the three columns
    detailNotes.push(['ค่าใช้จ่ายในการขายและบริหาร', '', '', '', '', '', 'ค่าใช้จ่ายในการขาย', 'ค่าใช้จ่ายในการบริหาร', 'ค่าใช้จ่ายอื่น']);
    
    const dataStartRow = detailNotes.length + 1; // +1 because Excel is 1-indexed
    let financialCosts = 0;
    
    // Process accounts 5300-5999 excluding 5910 and 5360-5364
    for (const entry of trialBalanceData) {
      const accountCode = entry.accountCode;
      const accountName = entry.accountName;
      const amount = Math.abs(entry.debitAmount || entry.creditAmount || 0);
      
      if (accountCode && accountCode >= '5300' && accountCode <= '5999' && 
          accountCode !== '5910' && amount > 0) {
        
        // Check if it's financial costs (5360-5364) - handle separately
        if (accountCode >= '5360' && accountCode <= '5364') {
          financialCosts += amount;
          continue;
        }
        
        // Determine which column the expense goes to
        let sellingCost = 0;
        let adminCost = 0;
        let otherCost = 0;
        
        const codePrefix = accountCode.substring(0, 4);
        if (codePrefix >= '5300' && codePrefix <= '5311') {
          sellingCost = amount;
        } else if (codePrefix >= '5312' && codePrefix <= '5350') {
          adminCost = amount;
        } else if ((codePrefix >= '5351' && codePrefix <= '5359') || 
                   (codePrefix >= '5365' && codePrefix <= '5999')) {
          otherCost = amount;
        }
        
        detailNotes.push([accountName, '', '', '', '', '', sellingCost, adminCost, otherCost]);
      }
    }
    
    // Add total row with formulas
    const totalRowIndex = detailNotes.length + 1;
    detailNotes.push(['รวม', '', '', '', '', '', 
      { f: `SUM(G${dataStartRow}:G${totalRowIndex - 1})` },
      { f: `SUM(H${dataStartRow}:H${totalRowIndex - 1})` },
      { f: `SUM(I${dataStartRow}:I${totalRowIndex - 1})` }
    ]);
    
    // Add financial costs row if any
    if (financialCosts > 0) {
      detailNotes.push(['ค่าใช้จ่ายต้นทุนทางการเงิน', '', '', '', '', '', 0, 0, financialCosts]);
    }
  }

  private checkHasInventory(trialBalanceData: TrialBalanceEntry[]): boolean {
    // Check for inventory account 1510
    const inventoryAccount = trialBalanceData.find(entry => entry.accountCode === '1510');
    if (inventoryAccount && (inventoryAccount.debitAmount !== 0 || inventoryAccount.creditAmount !== 0)) {
      return true;
    }

    // Check for purchase accounts 5010
    const purchaseAccounts = trialBalanceData.filter(entry => 
      entry.accountCode?.startsWith('5010') && 
      (entry.debitAmount !== 0 || entry.creditAmount !== 0)
    );
    return purchaseAccounts.length > 0;
  }

  // ============================================================================
  // NOTES_ACCOUNTING HELPER METHODS (Enhanced with Global Data)
  // ============================================================================

  // NEW: Cash Note using Global Data - eliminates redundant calculations
  private addCashNoteWithGlobalData(
    notes: any[][], 
    globalData: ExtractedFinancialData,
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 3
  ): void {
    const totalAmount = globalData.assets.cashAndCashEquivalents.current;
    const prevTotalAmount = globalData.assets.cashAndCashEquivalents.previous;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      
      // Calculate detailed breakdown for audit trail (individual calculations are fine)
      const cashAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1010, 1019));
      const bankAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1020, 1099));
      const prevCashAmount = processingType === 'multi-year' ? 
        Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1010, 1019)) : 0;
      const prevBankAmount = processingType === 'multi-year' ? 
        Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1020, 1099)) : 0;
      
      // Show individual breakdown with actual amounts
      if (cashAmount !== 0 || prevCashAmount !== 0) {
        notes.push(['', 'เงินสดในมือ', '', '', '', '', 
          cashAmount, '', 
          processingType === 'multi-year' ? prevCashAmount : '']);
      }
      
      if (bankAmount !== 0 || prevBankAmount !== 0) {
        notes.push(['', 'เงินฝากธนาคาร', '', '', '', '', 
          bankAmount, '', 
          processingType === 'multi-year' ? prevBankAmount : '']);
      }
      
      // Use global data for total (eliminates redundant calculation at summary level)
      if (processingType === 'multi-year') {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', prevTotalAmount]);
      } else {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  // NEW: Bank Overdrafts Note using Global Data - eliminates redundant calculations
  private addBankOverdraftsNoteWithGlobalData(
    notes: any[][], 
    globalData: ExtractedFinancialData,
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 9
  ): void {
    const totalAmount = globalData.liabilities.bankOverdraftsAndShortTermLoans.current;
    const prevTotalAmount = globalData.liabilities.bankOverdraftsAndShortTermLoans.previous;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
        notes.push(['', 'เงินเบิกเกินบัญชีธนาคาร', '', '', '', '', totalAmount, '', prevTotalAmount]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
        notes.push(['', 'เงินเบิกเกินบัญชีธนาคาร', '', '', '', '', totalAmount, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  // NEW: Short Term Borrowings Note using Global Data - eliminates redundant calculations
  private addShortTermBorrowingsNoteWithGlobalData(
    notes: any[][], 
    globalData: ExtractedFinancialData,
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 13
  ): void {
    const totalAmount = globalData.liabilities.shortTermBorrowings.current;
    const prevTotalAmount = globalData.liabilities.shortTermBorrowings.previous;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
        notes.push(['', 'เงินกู้ยืมจากสถาบันการเงิน', '', '', '', '', totalAmount, '', prevTotalAmount]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
        notes.push(['', 'เงินกู้ยืมจากสถาบันการเงิน', '', '', '', '', totalAmount, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  // NEW: Trade Receivables Note using Global Data - partial optimization
  private addTradeReceivablesNoteWithGlobalData(
    notes: any[][], 
    globalData: ExtractedFinancialData,
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 4
  ): void {
    const totalAmount = globalData.assets.tradeReceivables.current;
    const prevTotalAmount = globalData.assets.tradeReceivables.previous;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'ลูกหนี้การค้าและลูกหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      
      // For detailed breakdown, still use individual account logic (acceptable)
      notes.push(['', 'ลูกหนี้การค้า', '', '', '', '', '...', '', '...']);
      notes.push(['', 'ลูกหนี้อื่น', '', '', '', '', '...', '', '...']);
      
      if (processingType === 'multi-year') {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', prevTotalAmount]);
      } else {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  // NEW: Trade Payables Note using Global Data - partial optimization  
  private addTradePayablesNoteWithGlobalData(
    notes: any[][], 
    globalData: ExtractedFinancialData,
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year',
    noteNumber: number = 12
  ): void {
    const totalAmount = globalData.liabilities.tradeAndOtherPayables.current;
    const prevTotalAmount = globalData.liabilities.tradeAndOtherPayables.previous;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
      }
      
      // For detailed breakdown, still use individual account logic (acceptable)
      notes.push(['', 'เจ้าหนี้การค้า', '', '', '', '', '...', '', '...']);
      notes.push(['', 'เจ้าหนี้อื่น', '', '', '', '', '...', '', '...']);
      
      if (processingType === 'multi-year') {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', prevTotalAmount]);
      } else {
        notes.push(['', 'รวม', '', '', '', '', totalAmount, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  // ============================================================================
  // LEGACY NOTES METHODS (VBA-COMPLIANT) - Kept for non-optimized notes
  // ============================================================================

  private addCashNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 3
  ): void {
    // Calculate cash amounts (1010-1019) and bank deposits (1020-1099)
    const cashAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1010, 1019));
    const bankAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1020, 1099));
    const totalAmount = cashAmount + bankAmount;
    
    // Use previousBalance field from trialBalanceData for previous year amounts
    const prevCashAmount = processingType === 'multi-year' ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1010, 1019)) : 0;
    const prevBankAmount = processingType === 'multi-year' ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalanceData, 1020, 1099)) : 0;
    const prevTotalAmount = prevCashAmount + prevBankAmount;

    // Only add note if there are actual balances
    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      // Cash line - only show if has balance
      if (cashAmount !== 0 || prevCashAmount !== 0) {
        notes.push(['', '', 'เงินสด', '', '', '', 
          cashAmount, '', 
          processingType === 'multi-year' ? prevCashAmount : '']);
      }
      
      // Bank deposits line - only show if has balance
      if (bankAmount !== 0 || prevBankAmount !== 0) {
        notes.push(['', '', 'เงินฝากธนาคาร', '', '', '', 
          bankAmount, '', 
          processingType === 'multi-year' ? prevBankAmount : '']);
      }
      
      // Total line - use Excel formulas instead of calculated values
      const currentRowIndex = notes.length + 1; // Excel is 1-indexed
      const startRowIndex = currentRowIndex - (cashAmount !== 0 || prevCashAmount !== 0 ? 1 : 0) - (bankAmount !== 0 || prevBankAmount !== 0 ? 1 : 0);
      
      if (processingType === 'multi-year') {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${startRowIndex}:G${currentRowIndex - 1})` }, '', 
          { f: `SUM(I${startRowIndex}:I${currentRowIndex - 1})` }]);
      } else {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${startRowIndex}:G${currentRowIndex - 1})` }, '', '']);
      }
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addTradeReceivablesNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 4
  ): void {
    // Get all individual trade receivable accounts (1140-1215)
    const tradeReceivableAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= 1140 && code <= 1215;
    });

    // Check if any accounts have balances
    const hasBalances = tradeReceivableAccounts.some(account => 
      Math.abs((account.debitAmount || 0) - (account.creditAmount || 0)) !== 0 ||
      Math.abs(account.previousBalance || 0) !== 0
    );

    if (!hasBalances) {
      return; // Don't create note if no balances
    }

    // Add note header
    notes.push([noteNumber.toString(), 'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
    
    // Add column headers based on processing type
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }

    const dataStartRow = notes.length + 1; // Track start of individual account details
    let hasDetailLines = false;

    // Process each individual account that has balances
    tradeReceivableAccounts.forEach(account => {
      const currentBalance = Math.abs((account.debitAmount || 0) - (account.creditAmount || 0));
      const previousBalance = Math.abs(account.previousBalance || 0);
      
      // Only show accounts with non-zero balances
      if (currentBalance !== 0 || previousBalance !== 0) {
        notes.push(['', '', account.accountName, '', '', '', 
          currentBalance, '', 
          processingType === 'multi-year' ? previousBalance : '']);
        hasDetailLines = true;
      }
    });

    // Add total row with Excel formulas if there are detail lines
    if (hasDetailLines) {
      const currentRowIndex = notes.length + 1; // Excel 1-indexed
      const endRowIndex = currentRowIndex - 1; // Last detail row
      
      if (processingType === 'multi-year') {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${dataStartRow}:G${endRowIndex})` }, '', 
          { f: `SUM(I${dataStartRow}:I${endRowIndex})` }]);
      } else {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${dataStartRow}:G${endRowIndex})` }, '', '']);
      }
    }
    
    // Add spacing after the note
    notes.push(['', '', '', '', '', '', '', '', '']);
  }

  private addShortTermLoansNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 5
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1141, 1141));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 1141, 1141)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินให้กู้ยืมระยะสั้น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินให้กู้ยืมระยะสั้น', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added (0-based array length = 1-based Excel row - 1)
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addPPENote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 6
  ): void {
    // Get all PPE asset accounts (1610-1659 without decimal points)
    const assetAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode);
      return code >= 1610 && code <= 1659 && !entry.accountCode.includes('.');
    });

    // Get all accumulated depreciation accounts (1610-1659 with decimal points)
    const depreciationAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode);
      return code >= 1610 && code <= 1659 && entry.accountCode.includes('.');
    });

    // Only create note if there are any PPE accounts with balances
    if (assetAccounts.length === 0 && depreciationAccounts.length === 0) {
      return;
    }

    // Check if any accounts have non-zero balances
    const hasAssetBalances = assetAccounts.some(acc => 
      Math.abs(acc.balance) !== 0 || Math.abs(acc.previousBalance || 0) !== 0
    );
    const hasDepreciationBalances = depreciationAccounts.some(acc => 
      Math.abs(acc.balance) !== 0 || Math.abs(acc.previousBalance || 0) !== 0
    );

    if (!hasAssetBalances && !hasDepreciationBalances) {
      return;
    }

    // Create note header
    notes.push([noteNumber.toString(), 'ที่ดิน อาคารและอุปกรณ์', '', '', '', '', '', '', 'หน่วย:บาท']);
    
    // Add column headers
    if (processingType === 'multi-year') {
      notes.push(['', '', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear - 1}`, '', 'ซื้อเพิ่ม', 'จำหน่ายออก', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear}`]);
    } else {
      notes.push(['', '', '', '', '', '', '', '', `ณ 31 ธ.ค. ${companyInfo.reportingYear}`]);
    }
    
    // Add asset cost section
    notes.push(['', '', 'ราคาทุนเดิม', '', '', '', '', '', '']);
    
    let assetTotalCurrent = 0;
    let assetTotalPrevious = 0;
    let assetTotalPurchases = 0;
    const assetStartRow = notes.length + 1; // Track start of asset details (Excel 1-indexed)

    // Process individual asset accounts
    assetAccounts.forEach(account => {
      const currentAmount = Math.abs(account.balance);
      const previousAmount = Math.abs(account.previousBalance || 0);
      const purchases = Math.max(0, currentAmount - previousAmount); // Only positive purchases
      
      if (currentAmount !== 0 || previousAmount !== 0) {
        if (processingType === 'multi-year') {
          notes.push(['', '', account.accountName, 
            previousAmount, '', // Column D (index 3): Previous amount
            purchases > 0 ? purchases : '', '', // Column F (index 5): Purchases
            '', currentAmount]); // Column I (index 8): Current amount
        } else {
          notes.push(['', '', account.accountName, '', '', '', '', '', currentAmount]); // Column I (index 8): Current amount only
        }
        
        assetTotalCurrent += currentAmount;
        assetTotalPrevious += previousAmount;
        assetTotalPurchases += purchases;
      }
    });

    // Add asset totals with Excel formulas
    const assetEndRow = notes.length; // Last row of asset details (Excel 1-indexed)
    if (processingType === 'multi-year') {
      notes.push(['', '', 'รวม', 
        { f: `SUM(D${assetStartRow}:D${assetEndRow})` }, '', // Column D (index 3): Previous total formula
        { f: `SUM(F${assetStartRow}:F${assetEndRow})` }, '', // Column F (index 5): Total purchases formula
        '', { f: `SUM(I${assetStartRow}:I${assetEndRow})` }]); // Column I (index 8): Current total formula
    } else {
      notes.push(['', '', 'รวม', '', '', '', '', '', 
        { f: `SUM(I${assetStartRow}:I${assetEndRow})` }]); // Column I (index 8): Current total formula only
    }
    
    // Add space before depreciation section
    notes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add accumulated depreciation section
    notes.push(['', '', 'ค่าเสื่อมราคาสะสม', '', '', '', '', '', '']);
    
    let depreciationTotalCurrent = 0;
    let depreciationTotalPrevious = 0;
    let depreciationExpense = 0;
    const depreciationStartRow = notes.length + 1; // Track start of depreciation details (Excel 1-indexed)

    // Process individual depreciation accounts
    depreciationAccounts.forEach(account => {
      const currentAmount = Math.abs(account.balance); // Convert to positive
      const previousAmount = Math.abs(account.previousBalance || 0); // Convert to positive
      const expenseAmount = Math.max(0, currentAmount - previousAmount); // Depreciation expense for the year
      
      if (currentAmount !== 0 || previousAmount !== 0) {
        if (processingType === 'multi-year') {
          notes.push(['', '', account.accountName, 
            previousAmount, '', // Column D (index 3): Previous depreciation
            expenseAmount > 0 ? expenseAmount : '', '', // Column F (index 5): Depreciation expense
            '', currentAmount]); // Column I (index 8): Current depreciation
        } else {
          notes.push(['', '', account.accountName, '', '', '', '', '', currentAmount]); // Column I (index 8): Current depreciation only
        }
        
        depreciationTotalCurrent += currentAmount;
        depreciationTotalPrevious += previousAmount;
        depreciationExpense += expenseAmount;
      }
    });

    // Add depreciation totals with Excel formulas
    const depreciationEndRow = notes.length; // Last row of depreciation details (Excel 1-indexed)
    
    if (processingType === 'multi-year') {
      notes.push(['', '', 'รวม', 
        { f: `SUM(D${depreciationStartRow}:D${depreciationEndRow})` }, '', // Column D (index 3): Previous depreciation total formula
        { f: `SUM(F${depreciationStartRow}:F${depreciationEndRow})` }, '', // Column F (index 5): Total depreciation expense formula
        '', { f: `SUM(I${depreciationStartRow}:I${depreciationEndRow})` }]); // Column I (index 8): Current depreciation total formula
    } else {
      notes.push(['', '', 'รวม', '', '', '', '', '', 
        { f: `SUM(I${depreciationStartRow}:I${depreciationEndRow})` }]); // Column I (index 8): Current depreciation total formula only
    }

    // Add net book value with formulas referencing the totals above
    notes.push(['', '', '', '', '', '', '', '', '']);
    const assetTotalRowIndex = assetEndRow + 1; // Row number of asset totals
    const depreciationTotalRowIndex = notes.length - 1; // Row number of depreciation totals (just added above)
    
    if (processingType === 'multi-year') {
      notes.push(['', '', 'มูลค่าสุทธิ', 
        { f: `D${assetTotalRowIndex}-D${depreciationTotalRowIndex}` }, '', '', '', // Column D (index 3): Previous net value formula
        '', { f: `I${assetTotalRowIndex}-I${depreciationTotalRowIndex}` }]); // Column I (index 8): Current net value formula
    } else {
      notes.push(['', '', 'มูลค่าสุทธิ', '', '', '', '', '', 
        { f: `I${assetTotalRowIndex}-I${depreciationTotalRowIndex}` }]); // Column I (index 8): Current net value formula only
    }

    // Add depreciation expense summary - reference the depreciation expense total
    if (depreciationExpense > 0) {
      notes.push(['', '', 'ค่าเสื่อมราคา', '', '', '', '', '', 
        { f: `F${depreciationTotalRowIndex}` }]); // Column I (index 8): Reference depreciation expense total
    }
    
    notes.push(['', '', '', '', '', '', '', '', '']);
  }

  private addOtherAssetsNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 7
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1660, 1700));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 1660, 1700)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'สินทรัพย์อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'สินทรัพย์อื่น', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addBankOverdraftsNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 8
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2001, 2009));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2001, 2009)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินเบิกเกินบัญชี', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addTradePayablesNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 9
  ): void {
    // Get all accounts in the trade payables range (2010-2999)
    const allPayableAccounts = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= 2010 && code <= 2999;
    });

    // Define excluded account codes that appear as separate line items
    const excludedCodes = ['2030', '2045', '2050', '2051', '2052'];
    const excludedRangeCodes = Array.from({length: 24}, (_, i) => (2100 + i).toString()); // 2100-2123

    // Filter to get only trade payables accounts (excluding the specific items)
    const tradePayableAccounts = allPayableAccounts.filter(account => 
      !excludedCodes.includes(account.accountCode) && 
      !excludedRangeCodes.includes(account.accountCode)
    );

    // Check if any trade payable accounts have balances
    const hasBalances = tradePayableAccounts.some(account => 
      Math.abs((account.debitAmount || 0) - (account.creditAmount || 0)) !== 0 ||
      Math.abs(account.previousBalance || 0) !== 0
    );

    if (!hasBalances) {
      return; // Don't create note if no balances
    }

    // Add note header
    notes.push([noteNumber.toString(), 'เจ้าหนี้การค้าและเจ้าหนี้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
    
    // Add column headers based on processing type
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', '']);
    }

    const dataStartRow = notes.length + 1; // Track start of individual account details
    let hasDetailLines = false;

    // Process each individual trade payable account that has balances
    tradePayableAccounts.forEach(account => {
      const currentBalance = Math.abs((account.debitAmount || 0) - (account.creditAmount || 0));
      const previousBalance = Math.abs(account.previousBalance || 0);
      
      // Only show accounts with non-zero balances
      if (currentBalance !== 0 || previousBalance !== 0) {
        notes.push(['', '', account.accountName, '', '', '', 
          currentBalance, '', 
          processingType === 'multi-year' ? previousBalance : '']);
        hasDetailLines = true;
      }
    });

    // Add total row with Excel formulas if there are detail lines
    if (hasDetailLines) {
      const currentRowIndex = notes.length + 1; // Excel 1-indexed
      const endRowIndex = currentRowIndex - 1; // Last detail row
      
      if (processingType === 'multi-year') {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${dataStartRow}:G${endRowIndex})` }, '', 
          { f: `SUM(I${dataStartRow}:I${endRowIndex})` }]);
      } else {
        notes.push(['', '', 'รวม', '', '', '', 
          { f: `SUM(G${dataStartRow}:G${endRowIndex})` }, '', '']);
      }
    }
    
    // Add spacing after the note
    notes.push(['', '', '', '', '', '', '', '', '']);
  }

  private addShortTermBorrowingsNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 10
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2030, 2030));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2030, 2030)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะสั้นจากบุคคลหรือกิจการที่เกี่ยวข้องกัน', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินกู้ยืมระยะสั้น', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addLongTermLoansNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 11
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) - 
                       Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2121, 2121));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2120, 2123)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      
      // Add current portion detail - this would normally be user input in VBA
      const currentPortion = Math.round(totalAmount * 0.1); // Placeholder: 10% as current portion
      notes.push(['', '', 'หัก ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี', '', '', '', currentPortion, '', '']);
      notes.push(['', '', 'เงินกู้ยืมระยะยาวสุทธิจากส่วนที่ถึงกำหนดชำระคืนภายในหนึ่งปี', '', '', '', totalAmount - currentPortion, '', '']);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addOtherLongTermLoansNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 12
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2050, 2052)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาว', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินกู้ยืมระยะยาว', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addRelatedPartyLoansNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 13
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2100, 2100));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 2100, 2100)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'เงินกู้ยืมระยะยาวจากบุคคลหรือกิจการที่เกี่ยวข้องกัน', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'เงินกู้ยืมระยะยาวจากบุคคลที่เกี่ยวข้อง', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addOtherIncomeNote(
    notes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[], 
    noteNumber: number = 14
  ): void {
    const totalAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4020, 4999));
    const prevTotalAmount = processingType === 'multi-year' && trialBalancePrevious ? 
      Math.abs(this.sumPreviousBalanceByNumericRange(trialBalancePrevious, 4020, 4999)) : 0;

    if (totalAmount !== 0 || prevTotalAmount !== 0) {
      notes.push([noteNumber.toString(), 'รายได้อื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      if (processingType === 'multi-year') {
        notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
      } else {
        notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
      }
      
      notes.push(['', '', 'รายได้อื่น', '', '', '', 
        processingType === 'multi-year' ? totalAmount : '', '', 
        totalAmount !== 0 ? totalAmount : (processingType === 'multi-year' ? prevTotalAmount : totalAmount)]);
      
      const dataRowIndex = notes.length; // The data row we just added
      notes.push(['', '', 'รวม', '', '', '', 
        processingType === 'multi-year' ? {f: `G${dataRowIndex}`} : '', '', 
        {f: `I${dataRowIndex}`}]);
      notes.push(['', '', '', '', '', '', '', '', '']);
    }
  }

  private addExpensesByNatureNote(
    notes: any[][], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    noteNumber: number = 15
  ): void {
    notes.push([noteNumber.toString(), 'ค่าใช้จ่ายจำแนกตามธรรมชาติของค่าใช้จ่าย', '', '', '', '', '', '', 'หน่วย:บาท']);
    if (processingType === 'multi-year') {
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', `${companyInfo.reportingYear - 1}`]);
    } else {
      notes.push(['', '', '', '', '', '', '', '', `${companyInfo.reportingYear}`]);
    }
    
    notes.push(['', '', 'การเปลี่ยนแปลงในสินค้าสำเร็จรูปและงานระหว่างทำ', '', '', '', '', '', '']);
    notes.push(['', '', 'งานที่ทำโดยกิจการและบันทึกเป็นรายจ่ายฝ่ายทุน', '', '', '', '', '', '']);
    notes.push(['', '', 'วัตถุดิบและวัสดุสิ้นเปลืองใช้ไป', '', '', '', '', '', '']);
    notes.push(['', '', 'ค่าใช้จ่ายผลประโยชน์พนักงาน', '', '', '', '', '', '']);
    notes.push(['', '', 'ค่าเสื่อมราคาและค่าตัดจำหน่าย', '', '', '', '', '', '']);
    notes.push(['', '', 'ค่าใช้จ่ายอื่น', '', '', '', '', '', '']);
    notes.push(['', '', 'รวม', '', '', '', '', '', '']);
    notes.push(['', '', '', '', '', '', '', '', '']);
  }

  private addFinancialApprovalNote(
    notes: any[][], 
    companyInfo: CompanyInfo, 
    noteNumber: number = 16
  ): void {
    notes.push([noteNumber.toString(), 'การอนุมัติงบการเงิน', '', '', '', '', '', '', '']);
    notes.push(['', '', 'งบการเงินนี้ได้การรับอนุมัติให้ออกงบการเงินโดยคณะกรรมการผู้มีอำนาจของบริษัทแล้ว', '', '', '', '', '', '']);
    notes.push(['', '', '', '', '', '', '', '', '']);
  }

  // ============================================================================
  // NEW: Detail Notes methods using Global Data - eliminates redundant calculations
  // ============================================================================

  private addDetailOneWithGlobalData(
    detailNotes: any[][], 
    trialBalanceData: TrialBalanceEntry[], 
    globalData: ExtractedFinancialData
  ): void {
    const hasInventory = this.checkHasInventory(trialBalanceData);
    
    detailNotes.push(['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    if (!hasInventory) {
      // Service business - use global data for cost of services
      detailNotes.push(['ต้นทุนการให้บริการ', '', '', '', '', '', '', '', '']);
      detailNotes.push(['', 'ค่าใช้จ่ายอื่นๆ ในการให้บริการ', '', '', '', '', '', '', '...']);
      detailNotes.push(['', 'รวม', '', '', '', '', '', '', { f: 'I' + (detailNotes.length) }]);
      return;
    }

    // Inventory-based business - optimize with global data where possible
    detailNotes.push(['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', '']);
    
    // Use global data for inventory amounts
    const currentInventory = globalData.assets.inventory.current;
    const previousInventory = globalData.assets.inventory.previous;
    
    detailNotes.push(['', 'สินค้าคงเหลือต้นงวด', '', '', '', '', '', '', previousInventory]);
    
    // Process purchases (still need individual account details)
    let totalPurchases = 0;
    const purchaseAmount = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5010, 5010));
    if (purchaseAmount > 0) {
      detailNotes.push(['', 'บวก', 'ซื้อสินค้า', '', '', '', '', '', purchaseAmount]);
      totalPurchases += purchaseAmount;
    }
    
    // Purchase returns - still need individual calculation
    const returnAmount = Math.abs(this.getAccountBalance(trialBalanceData, ['5010.1']));
    if (returnAmount > 0) {
      detailNotes.push(['', 'หัก', 'ส่งคืนสินค้า', '', '', '', '', '', returnAmount]);
      totalPurchases -= returnAmount;
    }
    
    const availableForSale = previousInventory + totalPurchases;
    detailNotes.push(['', '', 'สินค้าที่มีไว้เพื่อขาย', '', '', '', '', '', availableForSale]);
    detailNotes.push(['', 'หัก', 'สินค้าคงเหลือปลายงวด', '', '', '', '', '', currentInventory]);
    
    const costOfGoodsSold = availableForSale - currentInventory;
    detailNotes.push(['', '', 'ต้นทุนสินค้าที่ขาย', '', '', '', '', '', costOfGoodsSold]);
  }

  private addDetailTwoWithGlobalData(
    detailNotes: any[][], 
    globalData: ExtractedFinancialData
  ): void {
    detailNotes.push(['รายละเอียดประกอบที่ 2', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    detailNotes.push(['ค่าใช้จ่ายในการขายและบริหาร', '', '', '', '', '', '', '', '']);
    
    // Use global data for total expenses - eliminates redundant calculations
    const totalExpenses = globalData.income.expenses.total;
    
    // Add expense categories with global totals
    detailNotes.push(['', 'ค่าใช้จ่ายในการขาย', '', '', '', '', '', '', globalData.income.expenses.adminExpenses]);
    detailNotes.push(['', 'ค่าใช้จ่ายในการบริหาร', '', '', '', '', '', '', globalData.income.expenses.otherExpenses]);
    detailNotes.push(['', '', 'รวม', '', '', '', '', '', totalExpenses]);
  }
}
