import { saveAs } from 'file-saver';
import { ExcelJSFormatter } from './excelFormatter';
// (moved) FinancialCalculations used in builders/extractor
import type { 
  TrialBalanceEntry, 
  CompanyInfo, 
  FinancialStatements,
  StatementResult
} from '../types/financial';
import type { DetailSettings } from '../types/detailSettings';
import { 
  CashNoteGenerator, 
  TradeReceivablesNoteGenerator, 
  TradePayablesNoteGenerator,
  PPENoteGenerator,
  OtherIncomeNoteGenerator,
  ShortTermLoansNoteGenerator,
  LiabilityShortTermLoansNoteGenerator,
  BankOverdraftsNoteGenerator,
  OtherAssetsNoteGenerator,
  OtherCurrentAssetsNoteGenerator,
  OtherCurrentLiabilitiesNoteGenerator,
  OtherNonCurrentLiabilitiesNoteGenerator,
  AssetLongTermLoansNoteGenerator,
  LongTermLoansNoteGenerator,
  HirePurchaseCreditorsNoteGenerator,
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
import type { SelectionFirstResult } from './financialStatements/selection/SelectionFirstClassifier';

// ============================================================================
// MAIN FINANCIAL STATEMENT GENERATOR CLASS
// ============================================================================

export class FinancialStatementGenerator {
  
  // ============================================================================
  // GLOBAL DATA EXTRACTION (Calculate Once, Use Everywhere)
  // ============================================================================
  
  private extractedData: DetailedFinancialData | null = null;
  private mappingProvider?: IAccountMappingProvider;
  private selectionData: SelectionFirstResult | null = null;
  private detailSettings: DetailSettings | null = null;
  
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
    provider?: IAccountMappingProvider,
    detailSettings?: DetailSettings
  ): FinancialStatements {
    
    // *** EXTRACT ALL DATA ONCE ***
  // Set/update provider if supplied
  if (provider) {
    this.mappingProvider = provider;
  }
  this.detailSettings = detailSettings ?? null;
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
  this.extractedData = globalData;
    
    console.log('=== USING GLOBAL DATA FOR ALL STATEMENTS ===');
    console.log('Paid-up Capital (Global):', globalData.balanceSheetTotals.equity.paidUpCapital);
    console.log('Net Profit (Global):', globalData.income.netProfit);
    
    // Compute selection-first classification once for BS linkage as well
    const selectionForBS = SelectionFirstClassifier.classify(trialBalanceData, companyInfo, this.mappingProvider);
    this.selectionData = selectionForBS;
    
    // Generate accounting notes first to get note registry
    const accountingNotesResult = this.generateAccountingNotes(trialBalanceData, companyInfo, processingType, trialBalancePrevious);
    
    // Now build Balance Sheets with note registry for dynamic note references
    const balanceSheetAssetsResult = AssetsBuilder.build(trialBalanceData, companyInfo, processingType, globalData, selectionForBS, accountingNotesResult.noteRegistry);
    const balanceSheetLiabilitiesResult = LiabilitiesBuilder.build(trialBalanceData, companyInfo, processingType, globalData, selectionForBS, accountingNotesResult.noteRegistry);
    
    // Generate P&L with note registry for Other Income and Expenses by Nature references
    const profitLossResult = this.generateProfitLossStatement(trialBalanceData, companyInfo, processingType, accountingNotesResult.noteRegistry);
    const changesInEquityResult = this.generateStatementOfChangesInEquity(trialBalanceData, companyInfo, processingType);
    const notesToFinancialStatements = this.generateNotesToFinancialStatements(companyInfo, trialBalanceData, processingType, trialBalancePrevious);
    const detailNotes = this.generateDetailNotes(trialBalanceData, companyInfo);

    return {
      balanceSheet: {
        assets: balanceSheetAssetsResult.data,
        liabilities: balanceSheetLiabilitiesResult.data,
        assetsSignatureRows: balanceSheetAssetsResult.signatureRows,
        liabilitiesSignatureRows: balanceSheetLiabilitiesResult.signatureRows
      },
      profitLossStatement: profitLossResult.data,
      profitLossSignatureRows: profitLossResult.signatureRows,
      changesInEquity: changesInEquityResult.data,
      changesInEquitySignatureRows: changesInEquityResult.signatureRows,
      notes: notesToFinancialStatements,
      accountingNotes: accountingNotesResult.notes,
      accountingNotesFormatters: accountingNotesResult.formatters,
      accountingNotesSignatureRows: accountingNotesResult.signatureRows,
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

    // Create cover sheet as the first sheet
    ExcelJSFormatter.createCoverSheet(
      workbook, 
      statements.companyInfo.name,
      statements.companyInfo.reportingPeriodEndDate || '31 ธันวาคม',
      statements.companyInfo.reportingYear
    );

    // Create Balance Sheet - Assets
    const assetsWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'BS-A', statements.balanceSheet.assets);
    ExcelJSFormatter.formatBalanceSheetAssets(assetsWs, statements.balanceSheet.assetsSignatureRows);

    // Create Balance Sheet - Liabilities
    const liabilitiesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'BS-L', statements.balanceSheet.liabilities);
    ExcelJSFormatter.formatBalanceSheetAssets(liabilitiesWs, statements.balanceSheet.liabilitiesSignatureRows); // Pass signature rows

    // Create Profit & Loss Statement
    const plWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'P&L', statements.profitLossStatement);
    ExcelJSFormatter.formatBalanceSheetAssets(plWs, statements.profitLossSignatureRows); // Pass signature rows

    // Create Statement of Changes in Equity
    const equityWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'SCE', statements.changesInEquity);
    ExcelJSFormatter.formatStatementOfChangesInEquity(equityWs, statements.changesInEquitySignatureRows); // Pass signature rows

    // Create Notes to Financial Statements (Policy Notes)
    const notesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'N-Policy', statements.notes);
    ExcelJSFormatter.formatNotesToFinancialStatements(notesWs);

    // Create Accounting Notes (Detailed Notes) with Specific Formatting
    const accountingNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'N-Acct', statements.accountingNotes);
    
    // Use new specific formatting approach if formatters are available
    if (statements.accountingNotesFormatters && statements.accountingNotesFormatters.length > 0) {
      console.log('Using specific row-tracked formatting for Notes_Accounting');
      ExcelJSFormatter.formatNotesWithSpecificFormatting(accountingNotesWs, statements.accountingNotesFormatters, statements.accountingNotesSignatureRows);
    } else {
      console.log('Using fallback pattern-based formatting for Notes_Accounting');
      ExcelJSFormatter.formatNotesWithoutBackground(accountingNotesWs, statements.accountingNotesSignatureRows);
    }

    // Create Detail Notes (if available)
    if (statements.detailNotes?.detail1) {
      const detailNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'N-Detail', statements.detailNotes.detail1);
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
    processingType: 'single-year' | 'multi-year',
    noteRegistry?: import('./financialStatements/core/types').NoteRegistry
  ): StatementResult {
    // Use strict selection (no numeric fallback) for all P&L buckets so UI rules drive the result
    const selectionStrict = SelectionFirstClassifier.classify(
      trialBalanceData,
      companyInfo,
      this.mappingProvider,
      { disableFallbackFor: [
        'revenue',
        'other_income',
        'detail_service_costs',
        'selling_expenses',
        'admin_expenses',
        'other_expenses',
        'financial_costs',
        'income_tax_expense'
      ] as any }
    );
    console.log('[P&L] Using strict selection (no fallback) for P&L categories');
    return ProfitLossBuilder.build(trialBalanceData, companyInfo, processingType, selectionStrict, noteRegistry);
  }

  private generateStatementOfChangesInEquity(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year'
  ): StatementResult {
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
  ): { notes: any[][], formatters: NoteFormatter[], noteRegistry: import('./financialStatements/core/types').NoteRegistry, signatureRows?: number[] } {
    // Extract global financial data once for consistency across all notes
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
    this.extractedData = globalData;
    console.log('=== NOTES_ACCOUNTING: Using Global Data Extraction with Row Tracking ===');
    // Compute selection-first classification once (mapping-first architecture)
  const selection = SelectionFirstClassifier.classify(trialBalanceData, companyInfo, this.mappingProvider);
  this.selectionData = selection;
  const totalClassified = Object.keys(selection.byAccount).length;
  const unmatchedCount = selection.unmatched?.length ?? 0;
  console.log(`[SelectionFirst] Classified ${totalClassified} accounts; unmatched=${unmatchedCount}. Receivables selected=${selection.byCategory?.receivables?.length ?? 0}`);
    
    const notes: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน (ต่อ)', '', '', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ ${companyInfo.reportingPeriodStartDate || '1 มกราคม'} ${companyInfo.reportingYear} ถึงวันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
    ];

    const formatters: NoteFormatter[] = [];
    const noteRegistry: import('./financialStatements/core/types').NoteRegistry = {};
    let noteNumber = 5;  // Start at note 5 (notes 1-4 are in Notes_Policy)
    
    // Generate specific notes using global data with row tracking
    // Only increment noteNumber if note actually has content
    const cashTracker = CashNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber, selection);
    if (cashTracker.headerRows.length > 0) {
      noteRegistry.cash = noteNumber++;
      formatters.push({ type: 'cash', tracker: cashTracker });
    }
    
    const receivablesTracker = TradeReceivablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber, selection);
    if (receivablesTracker.headerRows.length > 0) {
      noteRegistry.receivables = noteNumber++;
      formatters.push({ type: 'receivables', tracker: receivablesTracker });
    }

    // Move Short-term Loans (asset-side) to appear right after Trade Receivables
    const shortTermLoansTracker = ShortTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (shortTermLoansTracker.headerRows.length > 0) {
      noteRegistry.assetShortTermLoans = noteNumber++;
      formatters.push({ type: 'assetShortTermLoans', tracker: shortTermLoansTracker });
    }

    const otherCurrentAssetsTracker = OtherCurrentAssetsNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherCurrentAssetsTracker.headerRows.length > 0) {
      noteRegistry.otherCurrentAssets = noteNumber++;
      formatters.push({ type: 'otherCurrentAssets', tracker: otherCurrentAssetsTracker });
    }

    // Long-term Loans - Asset (เงินให้กู้ยืมระยะยาว)
    const assetLongTermLoansTracker = AssetLongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (assetLongTermLoansTracker.headerRows.length > 0) {
      noteRegistry.assetLongTermLoans = noteNumber++;
      formatters.push({ type: 'assetLongTermLoans', tracker: assetLongTermLoansTracker });
    }

    // Property, Plant & Equipment Note (PPE) with Row Tracking - Enhanced formatting
    const ppeTracker = PPENoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (ppeTracker.headerRows.length > 0) {
      noteRegistry.ppe = noteNumber++;
      formatters.push({ type: 'ppe', tracker: ppeTracker });
    }
    
    // Other Non-Current Assets (สินทรัพย์ไม่หมุนเวียนอื่น)
    const otherAssetsTracker = OtherAssetsNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherAssetsTracker.headerRows.length > 0) {
      noteRegistry.otherAssets = noteNumber++;
      formatters.push({ type: 'general', tracker: otherAssetsTracker });
    }
    
    // Bank Overdrafts and Short-term Borrowings from Financial Institutions Note (before payables)
    const bankOverdraftsTracker = BankOverdraftsNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (bankOverdraftsTracker.headerRows.length > 0) {
      noteRegistry.bankOverdrafts = noteNumber++;
      formatters.push({ type: 'bankOverdrafts', tracker: bankOverdraftsTracker });
    }
    
    const payablesTracker = TradePayablesNoteGenerator.generateWithRowTracking(notes, globalData, companyInfo, processingType, noteNumber, selection);
    if (payablesTracker.headerRows.length > 0) {
      noteRegistry.payables = noteNumber++;
      formatters.push({ type: 'payables', tracker: payablesTracker });
    }

    // Liability Short-term Loans (เงินกู้ยืมระยะสั้น - borrowed) - After trade payables
    const liabilityShortTermLoansTracker = LiabilityShortTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (liabilityShortTermLoansTracker.headerRows.length > 0) {
      noteRegistry.liabilityShortTermLoans = noteNumber++;
      formatters.push({ type: 'liabilityShortTermLoans', tracker: liabilityShortTermLoansTracker });
    }

    const otherCurrentLiabilitiesTracker = OtherCurrentLiabilitiesNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherCurrentLiabilitiesTracker.headerRows.length > 0) {
      noteRegistry.otherCurrentLiabilities = noteNumber++;
      formatters.push({ type: 'otherCurrentLiabilities', tracker: otherCurrentLiabilitiesTracker });
    }
    
    const hirePurchaseTracker = HirePurchaseCreditorsNoteGenerator.generateWithRowTracking(
      notes,
      trialBalanceData,
      companyInfo,
      processingType,
      trialBalancePrevious,
      noteNumber,
      selection,
      this.mappingProvider?.getSubCategoryRules('hire_purchase_creditors') || null
    );
    if (hirePurchaseTracker.headerRows.length > 0) {
      noteRegistry.hirePurchaseCreditors = noteNumber++;
      formatters.push({ type: 'hirePurchaseCreditors', tracker: hirePurchaseTracker });
    }

    // Long-term Loans from Financial Institutions (เงินกู้ยืมระยะยาวจากสถาบันการเงิน)
    const longTermLoansFromFITracker = LongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (longTermLoansFromFITracker.headerRows.length > 0) {
      noteRegistry.longTermLoansFromFI = noteNumber++;
      formatters.push({ type: 'longTermLoansFromFI', tracker: longTermLoansFromFITracker });
    }

    const otherLongTermLoansTracker = OtherLongTermLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherLongTermLoansTracker.headerRows.length > 0) {
      noteRegistry.otherLongTermLoans = noteNumber++;
      formatters.push({ type: 'general', tracker: otherLongTermLoansTracker });
    }

    const otherNonCurrentLiabilitiesTracker = OtherNonCurrentLiabilitiesNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherNonCurrentLiabilitiesTracker.headerRows.length > 0) {
      noteRegistry.otherNonCurrentLiabilities = noteNumber++;
      formatters.push({ type: 'otherNonCurrentLiabilities', tracker: otherNonCurrentLiabilitiesTracker });
    }
    
    const relatedPartyLoansTracker = RelatedPartyLoansNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (relatedPartyLoansTracker.headerRows.length > 0) {
      noteNumber++;  // Increment but don't store in registry (no Balance Sheet reference needed)
      formatters.push({ type: 'general', tracker: relatedPartyLoansTracker });
    }
    
    const otherIncomeTracker = OtherIncomeNoteGenerator.generateWithRowTracking(notes, trialBalanceData, companyInfo, processingType, trialBalancePrevious, noteNumber, selection);
    if (otherIncomeTracker.headerRows.length > 0) {
      noteRegistry.otherIncome = noteNumber++;  // Store in registry for P&L reference
      formatters.push({ type: 'general', tracker: otherIncomeTracker });
    }
    
    const expensesByNatureTracker = ExpensesByNatureNoteGenerator.generateWithRowTracking(notes, companyInfo, processingType, noteNumber);
    if (expensesByNatureTracker.headerRows.length > 0) {
      noteRegistry.expensesByNature = noteNumber++;  // Store in registry for P&L reference
      formatters.push({ type: 'expensesByNature', tracker: expensesByNatureTracker });
    }
    
    if (companyInfo.type === 'บริษัทจำกัด' || companyInfo.type === 'ห้างหุ้นส่วนจำกัด') {
      const financialApprovalTracker = FinancialApprovalNoteGenerator.generateWithRowTracking(notes, companyInfo, noteNumber);
      if (financialApprovalTracker.headerRows.length > 0) {
        noteNumber++;  // Increment but don't store in registry (no Balance Sheet reference needed)
        formatters.push({ type: 'general', tracker: financialApprovalTracker });
      }
    }

    // Add signatory block at the end of Notes_Accounting
    notes.push(['', '', '', '', '', '', '', '', '']);
    notes.push(['ขอรับรองว่าเป็นรายการอันถูกต้องและเป็นความจริง', '', '', '', '', '', '', '', '']);
    notes.push(['', '', '', '', '', '', '', '', '']);
    notes.push(['', '', '', '', '', '', '', '', '']);
    
    const signatureRowIndex = notes.length + 1; // 1-based for Excel
    const signatureTitle = companyInfo.type === 'ห้างหุ้นส่วนจำกัด' ? 'หุ้นส่วนผู้จัดการ' : 'กรรมการตามอำนาจ';
    notes.push([`ลงชื่อ ……………………..................................... ${signatureTitle}`, '', '', '', '', '', '', '', '']);
    
    const directorNameRowIndex = notes.length + 1; // 1-based for Excel
    const directorName = companyInfo.directorName || '...........................';
    notes.push([`(${directorName})`, '', '', '', '', '', '', '', '']);

    console.log(`Generated ${formatters.length} tracked notes with specific formatting`);
    return { notes, formatters, noteRegistry, signatureRows: [signatureRowIndex, directorNameRowIndex] };
  }

  private generateDetailNotes(trialBalanceData: TrialBalanceEntry[], companyInfo: CompanyInfo): any[][] {
    // Extract global financial data once for consistency
  const globalData = GlobalDataExtractor.extract(trialBalanceData, companyInfo, this.mappingProvider);
    this.extractedData = globalData;
    console.log('=== NOTES_DETAIL: Using Global Data Extraction ===');

  const selection = this.selectionData ?? SelectionFirstClassifier.classify(trialBalanceData, companyInfo, this.mappingProvider);
  this.selectionData = selection;

    // Diagnostics: show rule snapshots used for Detail Two categories
    try {
      const rulesSelling = this.mappingProvider?.getRules('selling_expenses') || null;
      const rulesAdmin = this.mappingProvider?.getRules('admin_expenses') || null;
      const rulesOther = this.mappingProvider?.getRules('other_expenses') || null;
      console.log('[DetailTwo][Rules] selling_expenses:', rulesSelling);
      console.log('[DetailTwo][Rules] admin_expenses:', rulesAdmin);
      console.log('[DetailTwo][Rules] other_expenses:', rulesOther);
      console.log('[DetailTwo][Selection sizes] selling:', selection.byCategory?.selling_expenses?.length ?? 0,
        'admin:', selection.byCategory?.admin_expenses?.length ?? 0,
        'other:', selection.byCategory?.other_expenses?.length ?? 0);
    } catch {}
    
    const detailNotes: any[][] = [];
    
    // Add header
    detailNotes.push([`${companyInfo.name}`, '', '', '', '', '', '', '', '']);
    detailNotes.push(['รายละเอียดประกอบหมายเหตุประกอบงบการเงิน', '', '', '', '', '', '', '', '']);
    detailNotes.push([`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ ${companyInfo.reportingPeriodStartDate || '1 มกราคม'} ${companyInfo.reportingYear} ถึงวันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', '']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add DT1 - Cost of goods sold / Service costs
    const detailOneMode = this.detailSettings?.detailOneMode ?? 'auto';
    const detailOneStartRow = detailNotes.length + 1;
    const detailOneData = DetailOneGenerator.generateDetailOne(
      trialBalanceData,
      globalData,
      selection,
      detailOneMode,
      this.mappingProvider,
      detailOneStartRow
    );
    detailNotes.push(...detailOneData);
    
    // Add some spacing
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    // Add DT2 - Selling and administrative expenses
    // Use strict selection (no numeric fallback) for expense buckets so UI rules drive the result
    const selectionStrict = SelectionFirstClassifier.classify(
      trialBalanceData,
      companyInfo,
      this.mappingProvider,
      { disableFallbackFor: ['selling_expenses','admin_expenses','other_expenses'] as any }
    );
    console.log('[DetailTwo] Using strict selection (no fallback) for selling/admin/other expenses');
    const detailTwoStart = detailNotes.length + 1;
    const detailTwoData = DetailTwoGenerator.generateDetailTwo(trialBalanceData, selectionStrict, detailTwoStart);
    detailNotes.push(...detailTwoData);
    console.log('[DetailTwo] Generated rows:', detailTwoData.length, 'starting at row', detailTwoStart);
    
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
