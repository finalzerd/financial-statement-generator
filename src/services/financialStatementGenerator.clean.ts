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
}

// ============================================================================
// MAIN FINANCIAL STATEMENT GENERATOR CLASS
// ============================================================================

export class FinancialStatementGenerator {
  
  // ============================================================================
  // PUBLIC INTERFACE METHODS
  // ============================================================================
  
  generateFinancialStatements(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    trialBalancePrevious?: TrialBalanceEntry[]
  ): FinancialStatements {
    
    const balanceSheetAssets = this.generateBalanceSheetAssets(trialBalanceData, companyInfo, processingType);
    const balanceSheetLiabilities = this.generateBalanceSheetLiabilities(trialBalanceData, companyInfo, processingType);
    const profitLossStatement = this.generateProfitLossStatement(trialBalanceData, companyInfo, processingType);
    const statementOfChangesInEquity = this.generateStatementOfChangesInEquity(trialBalanceData, companyInfo, processingType);
    const notesToFinancialStatements = this.generateNotesToFinancialStatements(companyInfo);
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
    ExcelJSFormatter.formatBalanceSheetAssets(equityWs); // Use available formatter

    // Create Notes to Financial Statements (Policy Notes)
    const notesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Policy', statements.notes);
    ExcelJSFormatter.formatBalanceSheetAssets(notesWs); // Use available formatter

    // Create Accounting Notes (Detailed Notes)
    const accountingNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Accounting', statements.accountingNotes);
    ExcelJSFormatter.formatBalanceSheetAssets(accountingNotesWs); // Use available formatter

    // Create Detail Notes (if available)
    if (statements.detailNotes?.detail1) {
      const detailNotesWs = ExcelJSFormatter.addDataToWorksheet(workbook, 'Notes_Detail', statements.detailNotes.detail1);
      ExcelJSFormatter.formatBalanceSheetAssets(detailNotesWs); // Use available formatter
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
    // Calculate asset balances using VBA-compliant ranges
    const cashAndCashEquivalents = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1000, 1099));
    const tradeReceivables = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1140, 1215));
    const inventory = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1500, 1519));
    const prepaidExpenses = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1400, 1439));
    const landBuildingsEquipment = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1600, 1659));
    const otherAssets = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 1660, 1700));

    // Initialize worksheet data with headers
    const worksheetData: (string | number | {f: string})[][] = [
      [''],
      ['', companyInfo.name],
      ['', 'งบแสดงฐานะการเงิน'],
      ['', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      [''],
      [''],
      ['', 'สินทรัพย์', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '', '']
    ];

    // Track current row and data rows for formulas
    let currentRow = worksheetData.length + 1;
    const currentAssetRows: number[] = [];

    // Current Assets section
    worksheetData.push(['', 'สินทรัพย์หมุนเวียน', '', '', '', '', '', '', 'หน่วย:บาท', '']);
    currentRow++;

    // Add current assets
    if (cashAndCashEquivalents !== 0) {
      worksheetData.push(['', '', 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '7', cashAndCashEquivalents, '', processingType === 'multi-year' ? 0 : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (tradeReceivables !== 0) {
      worksheetData.push(['', '', 'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น', '', '', '8', tradeReceivables, '', processingType === 'multi-year' ? 0 : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (inventory !== 0) {
      worksheetData.push(['', '', 'สินค้าคงเหลือ', '', '', '9', inventory, '', processingType === 'multi-year' ? 0 : '', '']);
      currentAssetRows.push(currentRow);
      currentRow++;
    }

    if (prepaidExpenses !== 0) {
      worksheetData.push(['', '', 'ค่าใช้จ่ายจ่ายล่วงหน้า', '', '', '10', prepaidExpenses, '', processingType === 'multi-year' ? 0 : '', '']);
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
      worksheetData.push(['', '', 'ที่ดิน อาคาร และอุปกรณ์ (สุทธิ)', '', '', '11', landBuildingsEquipment, '', processingType === 'multi-year' ? 0 : '', '']);
      nonCurrentAssetRows.push(currentRow);
      currentRow++;
    }

    if (otherAssets !== 0) {
      worksheetData.push(['', '', 'สินทรัพย์อื่น', '', '', '12', otherAssets, '', processingType === 'multi-year' ? 0 : '', '']);
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
    // Company type and terminology determination
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    const liabilityAndEquityTerm = isLimitedPartnership ? 
      'หนี้สินและส่วนของผู้เป็นหุ้นส่วน' : 'หนี้สินและส่วนของผู้ถือหุ้น';
    const equityTerm = isLimitedPartnership ? 
      'ส่วนของผู้เป็นหุ้นส่วน' : 'ส่วนของผู้ถือหุ้น';

    // Account balance calculations using VBA-compliant ranges
    const bankOverdraftsAndShortTermLoans = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2001, 2009));
    const totalCurrentPayablesRange = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2010, 2999));
    const excludeShortTermBorrowings = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2030, 2030));
    const excludeIncomeTax = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2045, 2045));
    const excludeLongTermCurrent = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052));
    const excludeLongTermFI = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2100, 2123));
    
    const tradeAndOtherPayables = totalCurrentPayablesRange - excludeShortTermBorrowings - 
                                 excludeIncomeTax - excludeLongTermCurrent - excludeLongTermFI;
    const shortTermBorrowings = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2030, 2030));
    const incomeTaxPayable = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2045, 2045));
    const longTermLoansFromFI = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) 
                               - Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2121, 2121));
    const otherLongTermLoans = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 2050, 2052));
    const registeredCapital = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3000, 3009));
    const paidUpCapital = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3010, 3019));
    const retainedEarnings = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3020, 3029));
    const legalReserve = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3030, 3039));

    // Initialize worksheet data with headers
    const worksheetData: (string | number | {f: string})[][] = [
      [''],
      ['', companyInfo.name],
      ['', 'งบแสดงฐานะการเงิน (ต่อ)'],
      ['', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      [''],
      [''],
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
      0, 0, 0, 0, processingType);
    
    this.buildNonCurrentLiabilitiesSection(worksheetData, cellTracker, 
      longTermLoansFromFI, otherLongTermLoans, 0, 0, processingType);
    
    this.buildEquitySection(worksheetData, cellTracker, 
      isLimitedPartnership, equityTerm, registeredCapital, paidUpCapital, retainedEarnings, legalReserve,
      0, 0, companyInfo, processingType);

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

    // Spacer
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;
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
      // Limited company equity structure
      worksheetData.push(['', '', 'ทุนเรือนหุ้น', '', '', '', '', '', '', '']);
      cellTracker.currentRow++;

      worksheetData.push(['', '', '', 'ทุนจดทะเบียน', '', '', '', '', '', '']);
      worksheetData.push(['', '', '', '', `หุ้นสามัญ ${registeredCapital.toLocaleString()} หุ้น มูลค่าหุ้นละ 1 บาท`, '', registeredCapital, '', processingType === 'multi-year' ? registeredCapital : '', '']);
      cellTracker.currentRow += 2;

      worksheetData.push(['', '', '', 'ทุนที่ออกและชำระแล้ว', '', '', '', '', '', '']);
      worksheetData.push(['', '', '', '', `หุ้นสามัญ ${paidUpCapital.toLocaleString()} หุ้น มูลค่าหุ้นละ 1 บาท`, '24', paidUpCapital, '', processingType === 'multi-year' ? prevPaidUpCapital : '', '']);
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
    const totalLiabilitiesRow = totalEquityRow - 1; // Total liabilities is right before equity section
    const grandTotalFormula = `G${totalLiabilitiesRow - cellTracker.equityDataRows.length - 2}+G${totalEquityRow}`;
    const grandTotalFormulaPrev = processingType === 'multi-year' ? `I${totalLiabilitiesRow - cellTracker.equityDataRows.length - 2}+I${totalEquityRow}` : '';
    
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
      ['รายได้', '', '', '', '', '', '', '', ''],
      ['', '', 'รายได้จากการขายหรือการให้บริการ', '', '', '1', revenue, '', processingType === 'multi-year' ? previousRevenue : ''],
      ['', '', 'รายได้อื่น', '', '', '2', otherIncome, '', processingType === 'multi-year' ? previousOtherIncome : ''],
      ['', 'รวมรายได้', '', '', '', '', { f: 'SUM(G8:G9)' }, '', processingType === 'multi-year' ? { f: 'SUM(I8:I9)' } : ''],
      ['', '', '', '', '', '', '', '', ''],
      ['ค่าใช้จ่าย', '', '', '', '', '', '', '', ''],
      ['', '', 'ต้นทุนขายหรือต้นทุนการให้บริการ', '', '', '3', costOfServices, '', processingType === 'multi-year' ? previousCostOfServices : ''],
      ['', '', 'ค่าใช้จ่ายในการบริหาร', '', '', '4', adminExpenses, '', processingType === 'multi-year' ? previousAdminExpenses : ''],
      ['', '', 'ค่าใช้จ่ายอื่น', '', '', '5', otherExpenses, '', processingType === 'multi-year' ? previousOtherExpenses : ''],
      ['', 'รวมค่าใช้จ่าย', '', '', '', '', { f: 'SUM(G13:G15)' }, '', processingType === 'multi-year' ? { f: 'SUM(I13:I15)' } : ''],
      ['', 'กำไรก่อนภาษีเงินได้', '', '', '', '', { f: 'G10-G16' }, '', processingType === 'multi-year' ? { f: 'I10-I16' } : ''],
      ['', '', 'ภาษีเงินได้', '', '', '6', incomeTax, '', processingType === 'multi-year' ? previousIncomeTax : ''],
      ['', 'กำไร(ขาดทุน)สุทธิ', '', '', '', '', { f: 'G17-G18' }, '', processingType === 'multi-year' ? { f: 'I17-I18' } : ''],
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
    const retainedEarnings = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3020, 3029));
    const currentYearProfit = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4000, 4999)) - 
                             Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5000, 5999));
    
    // Split capital equally between partners
    const partner1Capital = totalCapital / 2;
    const partner2Capital = totalCapital / 2;
    
    return [
      [`${companyInfo.name}`, '', '', '', '', '', ''],
      ['งบแสดงการเปลี่ยนแปลงส่วนของผู้เป็นหุ้นส่วน', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี สิ้นสุด วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['', '', '', '', '', '', 'หน่วย:บาท'],
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
    _processingType: 'single-year' | 'multi-year'
  ): any[][] {
    // Calculate corporate equity values
    const paidUpCapital = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3010, 3019));
    const legalReserve = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3030, 3039));
    const retainedEarnings = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 3020, 3029));
    const currentYearProfit = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 4000, 4999)) - 
                             Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5000, 5999));
    
    return [
      [`${companyInfo.name}`, '', '', '', '', '', ''],
      ['งบแสดงการเปลี่ยนแปลงส่วนของผู้ถือหุ้น', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี สิ้นสุด วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['', '', '', '', '', '', 'หน่วย:บาท'],
      ['', 'ทุนที่ออกและชำระแล้ว', 'ทุนสำรองตามกฎหมาย', 'กำไรสะสม', 'รวม', '', ''],
      ['ยอดคงเหลือ ณ วันต้นปี', paidUpCapital, legalReserve, retainedEarnings, { f: 'B7+C7+D7' }, '', ''],
      ['กำไรสุทธิสำหรับปี', '', '', currentYearProfit, currentYearProfit, '', ''],
      ['การจัดสรรกำไร', '', '', '', '', '', ''],
      ['  ทุนสำรองตามกฎหมาย', '', '', '', '', '', ''],
      ['ยอดคงเหลือ ณ วันสิ้นปี', { f: 'B7+B8' }, { f: 'C7+C8' }, { f: 'D7+D8' }, { f: 'B11+C11+D11' }, '', ''],
      ['', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '']
    ];
  }

  private generateNotesToFinancialStatements(companyInfo: CompanyInfo): any[][] {
    return [
      [`${companyInfo.name}`, '', ''],
      ['หมายเหตุประกอบงบการเงิน', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', ''],
      ['', '', ''],
      ['1. ข้อมูลทั่วไป', '', ''],
      ['', '', ''],
      ['1.1 ลักษณะของกิจการ', '', ''],
      [`${companyInfo.name} จดทะเบียนเป็น${companyInfo.type} เมื่อวันที่ ... `, '', ''],
      ['ประกอบกิจการ ...', '', ''],
      ['', '', ''],
      ['1.2 สถานที่ตั้ง', '', ''],
      ['สำนักงานใหญ่ตั้งอยู่เลขที่ ... ', '', ''],
      ['', '', ''],
      ['1.3 การอนุมัติงบการเงิน', '', ''],
      [`งบการเงินนี้ได้รับอนุมัติจากผู้มีอำนาจของ${companyInfo.name} เมื่อวันที่ ...`, '', ''],
      ['', '', ''],
      ['2. เกณฑ์การจัดทำงบการเงิน', '', ''],
      ['', '', ''],
      ['2.1 เกณฑ์การวัดมูลค่า', '', ''],
      ['งบการเงินนี้จัดทำขึ้นโดยใช้เกณฑ์ราคาทุนเดิม', '', ''],
      ['', '', ''],
      ['2.2 สกุลเงินที่ใช้ในการดำเนินงานและการนำเสนอ', '', ''],
      ['งบการเงินนี้จัดทำและแสดงหน่วยเงินตราเป็นบาท ซึ่งเป็นสกุลเงินที่ใช้ในการดำเนินงานของกิจการ', '', '']
    ];
  }

  private generateAccountingNotes(
    trialBalanceData: TrialBalanceEntry[], 
    companyInfo: CompanyInfo, 
    processingType: 'single-year' | 'multi-year', 
    trialBalancePrevious?: TrialBalanceEntry[]
  ): any[][] {
    const notes: any[][] = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงิน (ต่อ)', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
    ];

    // Add individual account notes based on VBA pattern
    let noteNumber = 3;
    
    // Cash and cash equivalents
    const cashBalance = this.sumAccountsByRange(trialBalanceData, ['1000', '1010', '1020'], 'current', trialBalancePrevious);
    if (cashBalance !== 0) {
      notes.push([noteNumber.toString(), 'เงินสดและรายการเทียบเท่าเงินสด', '', '', '', '', '', '', 'หน่วย:บาท']);
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '']);
      notes.push(['', 'เงินสดในมือ', '', '', '', '', '...', '', '...']);
      notes.push(['', 'เงินฝากธนาคาร', '', '', '', '', '...', '', '...']);
      notes.push(['', 'รวม', '', '', '', '', cashBalance, '', processingType === 'multi-year' ? '...' : '']);
      notes.push(['', '', '', '', '', '', '', '', '']);
      noteNumber++;
    }

    // Trade receivables
    const receivableBalance = this.sumAccountsByRange(trialBalanceData, ['1140', '1200', '1210', '1215'], 'current', trialBalancePrevious);
    if (receivableBalance !== 0) {
      notes.push([noteNumber.toString(), 'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น', '', '', '', '', '', '', 'หน่วย:บาท']);
      notes.push(['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '']);
      notes.push(['', 'ลูกหนี้การค้า', '', '', '', '', '...', '', '...']);
      notes.push(['', 'ลูกหนี้อื่น', '', '', '', '', '...', '', '...']);
      notes.push(['', 'รวม', '', '', '', '', receivableBalance, '', processingType === 'multi-year' ? '...' : '']);
      notes.push(['', '', '', '', '', '', '', '', '']);
      noteNumber++;
    }

    return notes;
  }

  private generateDetailNotes(trialBalanceData: TrialBalanceEntry[], companyInfo: CompanyInfo): any[][] {
    const hasInventory = this.checkHasInventory(trialBalanceData);
    
    if (!hasInventory) {
      return [
        [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
        ['รายละเอียดประกอบหมายเหตุประกอบงบการเงิน', '', '', '', '', '', '', '', ''],
        [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท'],
        ['', '', '', '', '', '', '', '', ''],
        ['ต้นทุนการให้บริการ', '', '', '', '', '', '', '', ''],
        ['', 'ค่าใช้จ่ายอื่นๆ ในการให้บริการ', '', '', '', '', '...', '', ''],
        ['', 'รวม', '', '', '', '', { f: 'G8' }, '', ''],
        ['', '', '', '', '', '', '', '', '']
      ];
    }

    // For inventory-based businesses
    const inventoryAccount = trialBalanceData.find(entry => entry.accountCode === '1510');
    const beginningInventory = inventoryAccount ? Math.abs(inventoryAccount.previousBalance || 0) : 0;
    const endingInventory = inventoryAccount ? Math.abs(inventoryAccount.debitAmount - inventoryAccount.creditAmount) : 0;
    const purchases = Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5010, 5019));

    return [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['รายละเอียดประกอบหมายเหตุประกอบงบการเงิน', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
      ['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท'],
      ['', '', '', '', '', '', '', '', ''],
      ['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', ''],
      ['', 'สินค้าคงเหลือต้นปี', '', '', '', '', beginningInventory, '', ''],
      ['', 'ซื้อสินค้า', '', '', '', '', purchases, '', ''],
      ['', 'สินค้าพร้อมขาย', '', '', '', '', { f: 'G8+G9' }, '', ''],
      ['', 'หัก สินค้าคงเหลือปลายปี', '', '', '', '', endingInventory, '', ''],
      ['', 'รวม', '', '', '', '', { f: 'G10-G11' }, '', ''],
      ['', '', '', '', '', '', '', '', '']
    ];
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
}
