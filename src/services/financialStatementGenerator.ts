import { saveAs } from 'file-saver';
import { ExcelJSFormatter } from './excelFormatter';
import { FinancialCalculations } from './financialCalculations';
import type { 
  TrialBalanceEntry, 
  CompanyInfo, 
  FinancialStatements
} from '../types/financial';
import { 
  CashNoteGenerator, 
  TradeReceivablesNoteGenerator, 
  TradePayablesNoteGenerator,
  PPENoteGenerator,
  OtherIncomeNoteGenerator,
  ShortTermLoansNoteGenerator,
  OtherAssetsNoteGenerator,
  LongTermLoansNoteGenerator,
  OtherLongTermLoansNoteGenerator,
  RelatedPartyLoansNoteGenerator,
  ExpensesByNatureNoteGenerator,
  FinancialApprovalNoteGenerator
} from './financialStatements/notes/noteTypes';
import { 
  DetailOneGenerator,
  DetailTwoGenerator
} from './financialStatements/details';
import type { 
  NoteFormatter, 
  DetailedFinancialData, 
  CellTracker 
} from './financialStatements/core/types';

// ============================================================================
// MAIN FINANCIAL STATEMENT GENERATOR CLASS
// ============================================================================

export class FinancialStatementGenerator {
  
  // ============================================================================
  // GLOBAL DATA EXTRACTION (Calculate Once, Use Everywhere)
  // ============================================================================
  
  private extractedData: DetailedFinancialData | null = null;
  
  /**
   * MAIN DATA EXTRACTION METHOD - Call this first to avoid redundant calculations
   */
  private extractAllFinancialData(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo
  ): DetailedFinancialData {
    
    if (this.extractedData) {
      return this.extractedData; // Return cached data
    }

    console.log('=== FOUNDATION-FIRST DATA EXTRACTION (NOTES → BALANCE SHEET) ===');

    // ============================================================================
    // FOUNDATION LAYER: Note calculations (calculated once, used everywhere)
    // ============================================================================
    
    // Note 7: Cash and cash equivalents breakdown
    const cashNote = {
      cash: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1019)), // เงินสดในมือ (includes 1010)
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1019))
      },
      bankDeposits: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1020, 1099)), // เงินฝากธนาคาร
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1020, 1099))
      },
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099)), // Total for Balance Sheet
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099))
      }
    };

    // Note 8: Trade and other receivables breakdown (FOUNDATION-FIRST with DYNAMIC ACCOUNTS)
    const receivablesNote = {
      // Calculate total from individual accounts - guarantees consistency
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215)), // Total for Balance Sheet
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215))
      }
      // Individual accounts will be extracted later and provide the detailed breakdown
    };

    // Note 9: Inventories (if applicable)
    const inventoryNote = {
      inventory: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1510, 1510)), // สินค้าคงเหลือ
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1510, 1510))
      },
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519)), // Total for Balance Sheet
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519))
      }
    };

    // Note 10: Property, plant and equipment
    const ppeNote = {
      cost: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1629)), // ราคาทุน
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1629))
      },
      accumulatedDepreciation: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1630, 1659)), // ค่าเสื่อมราคาสะสม
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1630, 1659))
      },
      netBookValue: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1629)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1630, 1659)), // Net for Balance Sheet
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1629)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1630, 1659))
      }
    };

    // Note 12: Trade and other payables breakdown (FOUNDATION-FIRST with DYNAMIC ACCOUNTS)
    const payablesNote = {
      // Calculate total from individual accounts - guarantees consistency
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2010, 2999)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2123)), // Total for Balance Sheet
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2010, 2999)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2123))
      }
      // Individual accounts will provide the detailed breakdown
    };

    // ============================================================================
    // BALANCE SHEET TOTALS: Derived from note calculations + individual items
    // ============================================================================
    
    // Assets (mix of note-derived and individual calculations)
    const balanceSheetAssets = {
      cashAndCashEquivalents: cashNote.total,           // From Note 7
      tradeReceivables: receivablesNote.total,          // From Note 8
      inventory: inventoryNote.total,                   // From Note 9
      prepaidExpenses: {                                // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1400, 1439)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439))
      },
      propertyPlantEquipment: ppeNote.netBookValue,     // From Note 10
      otherAssets: {                                    // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700))
      }
    };

    // Liabilities (mix of note-derived and individual calculations)
    const balanceSheetLiabilities = {
      bankOverdraftsAndShortTermLoans: {                // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2001, 2009)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2001, 2009))
      },
      tradeAndOtherPayables: payablesNote.total,        // From Note 12
      shortTermBorrowings: {                            // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030))
      },
      incomeTaxPayable: {                               // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045))
      },
      longTermLoansFromFI: {                            // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) - 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2121, 2121)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2120, 2123)) - 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2121, 2121))
      },
      otherLongTermLoans: {                             // Individual calculation
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) + 
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2119)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) + 
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2119))
      }
    };

    // EQUITY CALCULATION WITH CORRECTED RETAINED EARNINGS
    // Calculate current year profit properly: Revenue (credit-debit) - Expenses (debit-credit)
    const currentYearProfit = FinancialCalculations.calculateCurrentYearProfit(trialBalanceData);
    
    // CORRECTED: Get opening retained earnings from account 3020 using credit - debit
    const openingRetainedEarnings = FinancialCalculations.getOpeningRetainedEarnings(trialBalanceData);
    
    // Final retained earnings = opening + current year profit (VBA-compliant)
    const finalRetainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);
    
    const balanceSheetEquity = {
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
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3030, 3039)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3030, 3039))
      }
    };

    // INCOME STATEMENT CALCULATION
    const revenue = {
      total: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4999)),
      mainRevenue: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4099)),
      otherIncome: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4100, 4999))
    };

    const expenses = {
      total: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5999)),
      costOfServices: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5099)),
      adminExpenses: Math.abs(
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
      ),
      otherExpenses: Math.abs(
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
      ),
      incomeTax: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5910, 5910)),
      financialCosts: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5920, 5929))
    };

    const netProfit = revenue.total - expenses.total;

    // ============================================================================
    // INDIVIDUAL ACCOUNTS EXTRACTION: Dynamic structure for note breakdowns
    // ============================================================================
    const individualAccounts = this.extractIndividualAccounts(trialBalanceData);

    // BUSINESS LOGIC FLAGS
    const flags = {
      hasInventory: balanceSheetAssets.inventory.current > 0,
      isServiceBusiness: balanceSheetAssets.inventory.current === 0,
      isLimitedPartnership: companyInfo.type === 'ห้างหุ้นส่วนจำกัด'
    };

    // ============================================================================
    // COMPLETE FOUNDATION-FIRST DATA STRUCTURE
    // ============================================================================
    this.extractedData = {
      // FOUNDATION LAYER: Note calculations
      noteCalculations: {
        cash: cashNote,
        receivables: receivablesNote,
        inventory: inventoryNote,
        ppe: ppeNote,
        payables: payablesNote
      },
      
      // INDIVIDUAL ACCOUNT DETAILS: Dynamic structure for note breakdowns
      individualAccounts,
      
      // BALANCE SHEET TOTALS: Derived from notes + individual calculations
      balanceSheetTotals: {
        assets: balanceSheetAssets,
        liabilities: balanceSheetLiabilities,
        equity: balanceSheetEquity
      },
      
      // INCOME STATEMENT
      income: { revenue, expenses, netProfit },
      
      // BUSINESS LOGIC FLAGS
      flags
    };

    console.log('=== FOUNDATION-FIRST ARCHITECTURE COMPLETE ===');
    console.log('Cash Note Total:', cashNote.total);
    console.log('Balance Sheet Cash:', balanceSheetAssets.cashAndCashEquivalents);
    console.log('Paid-up Capital:', balanceSheetEquity.paidUpCapital);
    console.log('Retained Earnings:', balanceSheetEquity.retainedEarnings);
    console.log('Net Profit:', netProfit);
    console.log('=== NOTES → BALANCE SHEET ARCHITECTURE READY ===');

    return this.extractedData;
  }

  /**
   * Extract individual account details for note breakdowns
   * This eliminates the need for filtering in note generation methods
   */
  private extractIndividualAccounts(trialBalanceData: TrialBalanceEntry[]): any {
    const individualAccounts = {
      cash: {},
      receivables: {},
      payables: {}
    };

    console.log('=== EXTRACTING INDIVIDUAL ACCOUNTS (DYNAMIC) ===');

    // SINGLE PASS through trial balance - store individual accounts directly
    for (const entry of trialBalanceData) {
      const code = parseInt(entry.accountCode || '0');
      const currentAmount = Math.abs(entry.balance || 0);
      const previousAmount = Math.abs(entry.previousBalance || 0);
      
      // Skip accounts with no balance
      if (currentAmount === 0 && previousAmount === 0) continue;
      
      // Store individual cash accounts (1000-1099)
      if (code >= 1000 && code <= 1099) {
        (individualAccounts.cash as any)[entry.accountCode || ''] = {
          accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
          current: currentAmount,
          previous: previousAmount,
          category: code <= 1019 ? 'cash' : 'bankDeposits' // For display grouping only
        };
        console.log(`Cash Account ${entry.accountCode}: ${entry.accountName} = ${currentAmount}`);
      }
      
      // Store ALL individual receivable accounts (1140-1215) - NO GROUPING
      else if (code >= 1140 && code <= 1215) {
        (individualAccounts.receivables as any)[entry.accountCode || ''] = {
          accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
          current: currentAmount,
          previous: previousAmount
          // No artificial categories - just store the raw data
        };
        console.log(`Receivable Account ${entry.accountCode}: ${entry.accountName} = ${currentAmount}`);
      }
      
      // Store ALL individual payable accounts (2010-2999, with exclusions) - NO GROUPING
      else if (code >= 2010 && code <= 2999) {
        // Apply exclusion logic but don't create artificial categories
        const isExcluded = code === 2030 || code === 2045 || 
                          (code >= 2050 && code <= 2052) || 
                          (code >= 2100 && code <= 2123);
        
        if (!isExcluded) {
          (individualAccounts.payables as any)[entry.accountCode || ''] = {
            accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
            current: currentAmount,
            previous: previousAmount
            // No categories - just individual account data
          };
          console.log(`Payable Account ${entry.accountCode}: ${entry.accountName} = ${currentAmount}`);
        }
      }
    }
    
    console.log('Individual Cash Accounts:', Object.keys(individualAccounts.cash).length);
    console.log('Individual Receivable Accounts:', Object.keys(individualAccounts.receivables).length);
    console.log('Individual Payable Accounts:', Object.keys(individualAccounts.payables).length);
    console.log('=== END INDIVIDUAL ACCOUNTS EXTRACTION ===');
    
    return individualAccounts;
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
    console.log('Paid-up Capital (Global):', globalData.balanceSheetTotals.equity.paidUpCapital);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
    const balanceSheetAssets = this.generateBalanceSheetAssets(trialBalanceData, companyInfo, processingType);
    const balanceSheetLiabilities = this.generateBalanceSheetLiabilities(trialBalanceData, companyInfo, processingType);
    const profitLossStatement = this.generateProfitLossStatement(trialBalanceData, companyInfo, processingType);
    const statementOfChangesInEquity = this.generateStatementOfChangesInEquity(trialBalanceData, companyInfo, processingType);
    const notesToFinancialStatements = this.generateNotesToFinancialStatements(companyInfo, trialBalanceData, processingType, trialBalancePrevious);
    const accountingNotesResult = this.generateAccountingNotes(trialBalanceData, companyInfo, processingType, trialBalancePrevious);
    const detailNotes = this.generateDetailNotes(trialBalanceData, companyInfo);

    return {
      balanceSheet: {
        assets: balanceSheetAssets,
        liabilities: balanceSheetLiabilities
      },
      profitLossStatement,
      changesInEquity: statementOfChangesInEquity,
      notes: notesToFinancialStatements,
      accountingNotes: accountingNotesResult.notes,
      accountingNotesFormatters: accountingNotesResult.formatters,
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

    // Create Accounting Notes (Detailed Notes) with Specific Formatting
    const accountingNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Accounting', statements.accountingNotes);
    
    // Use new specific formatting approach if formatters are available
    if (statements.accountingNotesFormatters && statements.accountingNotesFormatters.length > 0) {
      console.log('Using specific row-tracked formatting for Notes_Accounting');
      ExcelJSFormatter.formatNotesWithSpecificFormatting(accountingNotesWs, statements.accountingNotesFormatters);
    } else {
      console.log('Using fallback pattern-based formatting for Notes_Accounting');
      ExcelJSFormatter.formatNotesWithoutBackground(accountingNotesWs);
    }

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
  // ============================================================================
  // ACCOUNT BALANCE CALCULATION METHODS
  // ============================================================================

  private sumAccountsByNumericRange(trialBalanceData: TrialBalanceEntry[], startCode: number, endCode: number): number {
    const matchingEntries = trialBalanceData.filter(entry => {
      const code = parseInt(entry.accountCode || '0');
      return code >= startCode && code <= endCode;
    });
    
    return matchingEntries.reduce((sum, entry) => {
      const balance = entry.balance || 0;
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
    const cashAndCashEquivalents = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099));
    const tradeReceivables = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215));
    const inventory = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519));
    const prepaidExpenses = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1400, 1439));
    const landBuildingsEquipment = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1659));
    const otherAssets = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700));

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
    const currentAssetsFormula = FinancialCalculations.buildSumFormula(currentAssetRows, 'G');
    const currentAssetsFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(currentAssetRows, 'I') : '';
    
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
    const nonCurrentAssetsFormula = FinancialCalculations.buildSumFormula(nonCurrentAssetRows, 'G');
    const nonCurrentAssetsFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(nonCurrentAssetRows, 'I') : '';
    
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
    const bankOverdraftsAndShortTermLoans = globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.current;
    const tradeAndOtherPayables = globalData.balanceSheetTotals.liabilities.tradeAndOtherPayables.current;
    const shortTermBorrowings = globalData.balanceSheetTotals.liabilities.shortTermBorrowings.current;
    const incomeTaxPayable = globalData.balanceSheetTotals.liabilities.incomeTaxPayable.current;
    const longTermLoansFromFI = globalData.balanceSheetTotals.liabilities.longTermLoansFromFI.current;
    const otherLongTermLoans = globalData.balanceSheetTotals.liabilities.otherLongTermLoans.current;
    const paidUpCapital = globalData.balanceSheetTotals.equity.paidUpCapital.current;
    const retainedEarnings = globalData.balanceSheetTotals.equity.retainedEarnings.current;
    const legalReserve = globalData.balanceSheetTotals.equity.legalReserve.current;

    // Use global data for previous year - eliminates redundant calculations
    const prevBankOverdraftsAndShortTermLoans = globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.previous;
    const prevTradeAndOtherPayables = globalData.balanceSheetTotals.liabilities.tradeAndOtherPayables.previous;
    const prevShortTermBorrowings = globalData.balanceSheetTotals.liabilities.shortTermBorrowings.previous;
    const prevIncomeTaxPayable = globalData.balanceSheetTotals.liabilities.incomeTaxPayable.previous;
    const prevLongTermLoansFromFI = globalData.balanceSheetTotals.liabilities.longTermLoansFromFI.previous;
    const prevOtherLongTermLoans = globalData.balanceSheetTotals.liabilities.otherLongTermLoans.previous;
    const prevPaidUpCapital = globalData.balanceSheetTotals.equity.paidUpCapital.previous;
    const prevRetainedEarnings = globalData.balanceSheetTotals.equity.retainedEarnings.previous;
    const prevLegalReserve = globalData.balanceSheetTotals.equity.legalReserve.previous;

    // Handle registered capital separately (not in global data yet)
    const registeredCapital = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3000, 3009));
    const prevRegisteredCapital = FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3000, 3009);

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
    const currentLiabilitiesFormula = FinancialCalculations.buildSumFormula(cellTracker.currentLiabilitiesRows, 'G');
    const currentLiabilitiesFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.currentLiabilitiesRows, 'I') : '';
    
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
    const nonCurrentLiabilitiesFormula = FinancialCalculations.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'G');
    const nonCurrentLiabilitiesFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'I') : '';
    
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
    const totalEquityFormula = FinancialCalculations.buildSumFormula(cellTracker.equityDataRows, 'G');
    const totalEquityFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.equityDataRows, 'I') : '';
    
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
    const revenue = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4099));
    const previousRevenue = processingType === 'multi-year' ? 0 : 0; // Placeholder for previous year
    
    // Calculate other income (4100-4999)
    const otherIncome = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4100, 4999));
    const previousOtherIncome = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate cost of services/goods sold (5000-5099)
    const costOfServices = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5099));
    const previousCostOfServices = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate administrative expenses (5300-5350 + selected ranges)
    const adminExpenses = Math.abs(
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
    );
    const previousAdminExpenses = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate other expenses (5351-5354, 5358-5361, 5364, 5366-5999)
    const otherExpenses = Math.abs(
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
    );
    const previousOtherExpenses = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate income tax (5910)
    const incomeTax = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5910, 5910));
    const previousIncomeTax = processingType === 'multi-year' ? 0 : 0;
    
    // Calculate financial costs (5920-5929)
    const financialCosts = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5920, 5929));
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
    console.log('Paid-up Capital (Global):', globalData.balanceSheetTotals.equity.paidUpCapital);
    console.log('Retained Earnings (Global):', globalData.balanceSheetTotals.equity.retainedEarnings);
    console.log('Opening Retained Earnings (Global):', globalData.balanceSheetTotals.equity.openingRetainedEarnings);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
    // Use global values directly - consistent across all statements!
    const paidUpCapitalCurrent = globalData.balanceSheetTotals.equity.paidUpCapital.current;
    const paidUpCapitalPrevious = globalData.balanceSheetTotals.equity.paidUpCapital.previous;
    const retainedEarningsCurrent = globalData.balanceSheetTotals.equity.retainedEarnings.current;
    const openingRetainedEarnings = globalData.balanceSheetTotals.equity.openingRetainedEarnings;
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
  ): { notes: any[][], formatters: NoteFormatter[] } {
    // Extract global financial data once for consistency across all notes
    const globalData = this.extractAllFinancialData(trialBalanceData, companyInfo);
    console.log('=== NOTES_ACCOUNTING: Using Global Data Extraction with Row Tracking ===');
    
    const notes: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน (ต่อ)', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
    ];

    const formatters: NoteFormatter[] = [];
    let noteNumber = 3;
    
    // Generate specific notes using global data with row tracking
    const cashTracker = CashNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++);
    if (cashTracker.headerRows.length > 0) {
      formatters.push({ type: 'cash', tracker: cashTracker });
    }
    
    const receivablesTracker = TradeReceivablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++);
    if (receivablesTracker.headerRows.length > 0) {
      formatters.push({ type: 'receivables', tracker: receivablesTracker });
    }
    
    const payablesTracker = TradePayablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++);
    if (payablesTracker.headerRows.length > 0) {
      formatters.push({ type: 'payables', tracker: payablesTracker });
    }
    
    // Property, Plant & Equipment Note (PPE) with Row Tracking - Enhanced formatting
    const ppeTracker = PPENoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (ppeTracker.headerRows.length > 0) {
      formatters.push({ type: 'ppe', tracker: ppeTracker });
    }
    
    const shortTermLoansTracker = ShortTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (shortTermLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'shortTermLoans', tracker: shortTermLoansTracker });
    }
    
    const otherAssetsTracker = OtherAssetsNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (otherAssetsTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: otherAssetsTracker });
    }
    
    const longTermLoansTracker = LongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (longTermLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: longTermLoansTracker });
    }
    
    const otherLongTermLoansTracker = OtherLongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (otherLongTermLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: otherLongTermLoansTracker });
    }
    
    const relatedPartyLoansTracker = RelatedPartyLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (relatedPartyLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: relatedPartyLoansTracker });
    }
    
    const otherIncomeTracker = OtherIncomeNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (otherIncomeTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: otherIncomeTracker });
    }
    
    const expensesByNatureTracker = ExpensesByNatureNoteGenerator.generateWithRowTracking(notes, companyInfo, processingType, noteNumber++);
    if (expensesByNatureTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: expensesByNatureTracker });
    }
    
    if (companyInfo.type === 'บริษัทจำกัด') {
      const financialApprovalTracker = FinancialApprovalNoteGenerator.generateWithRowTracking(notes, companyInfo, noteNumber++);
      if (financialApprovalTracker.headerRows.length > 0) {
        formatters.push({ type: 'general', tracker: financialApprovalTracker });
      }
    }

    console.log(`Generated ${formatters.length} tracked notes with specific formatting`);
    return { notes, formatters };
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
    const detailOneData = DetailOneGenerator.generateDetailOne(trialBalanceData, globalData);
    detailNotes.push(...detailOneData);
    
    // Add some spacing
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add DT2 - Selling and administrative expenses  
    const detailTwoData = DetailTwoGenerator.generateDetailTwo(trialBalanceData);
    detailNotes.push(...detailTwoData);
    
    return detailNotes;
  }

  // ============================================================================
  // NOTES_ACCOUNTING HELPER METHODS (Enhanced with Global Data) - Methods moved to reserved/unusedMethods.ts
  // ============================================================================

  // Cash Note method moved to reserved/unusedMethods.ts

  // Bank Overdrafts Note method moved to reserved/unusedMethods.ts

  // Short Term Borrowings Note method moved to reserved/unusedMethods.ts

  // Trade Receivables Note with Global Data method moved to reserved/unusedMethods.ts

  // Trade Payables Note with Global Data method moved to reserved/unusedMethods.ts

  // Trade Receivables Note with Individual Accounts method moved to reserved/unusedMethods.ts

  // Trade Payables Note with Individual Accounts method moved to reserved/unusedMethods.ts
}
