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
import { AssetsBuilder } from './financialStatements/balanceSheet/AssetsBuilder';
import { LiabilitiesBuilder } from './financialStatements/balanceSheet/LiabilitiesBuilder';
import type { 
  NoteFormatter, 
  DetailedFinancialData 
} from './financialStatements/core/types';
import { ProfitLossBuilder } from './financialStatements/profitLoss/ProfitLossBuilder';
import { EquityBuilder } from './financialStatements/equity/EquityBuilder';
import { GlobalDataExtractor } from './financialStatements/core/GlobalDataExtractor';
import { NotesPolicyBuilder } from './financialStatements/notes/NotesPolicyBuilder';

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
  previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3010, 3010))
      },
      retainedEarnings: {
        current: finalRetainedEarnings, // Use corrected calculation
  previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3020, 3020))
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
  const globalData = this.extractedData || GlobalDataExtractor.extract(trialBalanceData, companyInfo);
  this.extractedData = globalData;
    
    console.log('=== USING GLOBAL DATA FOR ALL STATEMENTS ===');
    console.log('Paid-up Capital (Global):', globalData.balanceSheetTotals.equity.paidUpCapital);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
  const balanceSheetAssets = AssetsBuilder.build(trialBalanceData, companyInfo, processingType);
  const balanceSheetLiabilities = LiabilitiesBuilder.build(trialBalanceData, companyInfo, processingType);
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

  // ============================================================================
  // BALANCE SHEET GENERATION METHODS
  // ============================================================================

  // generateBalanceSheetAssets moved to AssetsBuilder

  // Liabilities builder moved to LiabilitiesBuilder

  // ============================================================================
  // OTHER FINANCIAL STATEMENT GENERATION METHODS
  // ============================================================================
  
  private generateProfitLossStatement(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): any[][] {
    // Delegated to ProfitLossBuilder (extracted verbatim)
    return ProfitLossBuilder.build(trialBalanceData, companyInfo, processingType);
  }

  private generateStatementOfChangesInEquity(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): any[][] {
    // Delegated to EquityBuilder
    const globalData = this.extractedData!; // already populated
    return EquityBuilder.build(trialBalanceData, companyInfo, processingType, globalData);
  }

  private getSingleAccountBalance(trialBalanceData: TrialBalanceEntry[], accountCode: string): number {
    const account = trialBalanceData.find(entry => entry.accountCode === accountCode);
    // Use balance field if currentBalance is 0, taking absolute value for equity accounts
    const value = account ? Math.abs(account.currentBalance || account.balance || 0) : 0;
    return value;
  }

  private generateNotesToFinancialStatements(
    companyInfo: CompanyInfo, 
    _trialBalanceData?: TrialBalanceEntry[], 
    _processingType?: 'single-year' | 'multi-year', 
    _trialBalancePrevious?: TrialBalanceEntry[]
  ): (string | number | {f: string})[][] {
    return NotesPolicyBuilder.build(companyInfo, _processingType);
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
