import { saveAs } from 'file-saver';
import { ExcelJSFormatter } from './excelFormatter';
// (moved) FinancialCalculations used in builders/extractor
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
import type { IAccountMappingProvider } from './financialStatements/mapping/IAccountMappingProvider';
import { SelectionFirstClassifier } from './financialStatements/selection/SelectionFirstClassifier';

// ============================================================================
// MAIN FINANCIAL STATEMENT GENERATOR CLASS
// ============================================================================

export class FinancialStatementGenerator {
  
  // ============================================================================
  // GLOBAL DATA EXTRACTION (Calculate Once, Use Everywhere)
  // ============================================================================
  
  private extractedData: DetailedFinancialData | null = null;
  private mappingProvider?: IAccountMappingProvider;
  
  /**
   * MAIN DATA EXTRACTION METHOD - Call this first to avoid redundant calculations
   */
  // (removed) extractAllFinancialData - use GlobalDataExtractor.extract in callers

  /**
   * Extract individual account details for note breakdowns
   * This eliminates the need for filtering in note generation methods
   */
  // (removed) extractIndividualAccounts - handled inside GlobalDataExtractor
  
  // ============================================================================
  // PUBLIC INTERFACE METHODS
  // ============================================================================
  
  generateFinancialStatements(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    trialBalancePrevious?: TrialBalanceEntry[],
    provider?: IAccountMappingProvider
  ): FinancialStatements {
    
    // *** EXTRACT ALL DATA ONCE ***
  // Set/update provider if supplied
  if (provider) {
    this.mappingProvider = provider;
  }
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
  this.extractedData = globalData;
    
    console.log('=== USING GLOBAL DATA FOR ALL STATEMENTS ===');
    console.log('Paid-up Capital (Global):', globalData.balanceSheetTotals.equity.paidUpCapital);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
  // Compute selection-first classification once for BS linkage as well
  const selectionForBS = SelectionFirstClassifier.classify(trialBalanceData, companyInfo, this.mappingProvider);
  const balanceSheetAssets = AssetsBuilder.build(trialBalanceData, companyInfo, processingType, globalData, selectionForBS);
  const balanceSheetLiabilities = LiabilitiesBuilder.build(trialBalanceData, companyInfo, processingType, globalData, selectionForBS);
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

  // (removed) getSingleAccountBalance - logic centralized in GlobalDataExtractor

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
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
    this.extractedData = globalData;
    console.log('=== NOTES_ACCOUNTING: Using Global Data Extraction with Row Tracking ===');
    // Compute selection-first classification once (mapping-first architecture)
  const selection = SelectionFirstClassifier.classify(trialBalanceData, companyInfo, this.mappingProvider);
  const totalClassified = Object.keys(selection.byAccount).length;
  const unmatchedCount = selection.unmatched?.length ?? 0;
  console.log(`[SelectionFirst] Classified ${totalClassified} accounts; unmatched=${unmatchedCount}. Receivables selected=${selection.byCategory?.receivables?.length ?? 0}`);
    
    const notes: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน (ต่อ)', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
    ];

    const formatters: NoteFormatter[] = [];
    let noteNumber = 3;
    
    // Generate specific notes using global data with row tracking
    const cashTracker = CashNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++, selection);
    if (cashTracker.headerRows.length > 0) {
      formatters.push({ type: 'cash', tracker: cashTracker });
    }
    
    const receivablesTracker = TradeReceivablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++, selection);
    if (receivablesTracker.headerRows.length > 0) {
      formatters.push({ type: 'receivables', tracker: receivablesTracker });
    }

    // Property, Plant & Equipment Note (PPE) with Row Tracking - Enhanced formatting (should come before Payables)
    const ppeTracker = PPENoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++);
    if (ppeTracker.headerRows.length > 0) {
      formatters.push({ type: 'ppe', tracker: ppeTracker });
    }
    
    const payablesTracker = TradePayablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber++, selection);
    if (payablesTracker.headerRows.length > 0) {
      formatters.push({ type: 'payables', tracker: payablesTracker });
    }
    
    const shortTermLoansTracker = ShortTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++, selection);
    if (shortTermLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'shortTermLoans', tracker: shortTermLoansTracker });
    }
    
    const otherAssetsTracker = OtherAssetsNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++, selection);
    if (otherAssetsTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: otherAssetsTracker });
    }
    
    const longTermLoansTracker = LongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++, selection);
    if (longTermLoansTracker.headerRows.length > 0) {
      formatters.push({ type: 'general', tracker: longTermLoansTracker });
    }
    
    const otherLongTermLoansTracker = OtherLongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber++, selection);
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
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
    this.extractedData = globalData;
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
