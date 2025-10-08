import { FinancialCalculations } from '../../financialCalculations';
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { CellTracker, DetailedFinancialData } from '../core/types';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';
import { BalanceSheetLinkMap, NOTE_FIRST_MODE } from '../core/linking/balanceSheetLinkMap';

// Builds the Balance Sheet (Liabilities & Equity) worksheet data.
// This mirrors the original logic from FinancialStatementGenerator.generateBalanceSheetLiabilities
// without changing behavior or output shape.
export class LiabilitiesBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData?: DetailedFinancialData,
    selection?: SelectionFirstResult
  ): (string | number | { f: string })[][] {
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    const liabilityAndEquityTerm = isLimitedPartnership ? 'หนี้สินและส่วนของผู้เป็นหุ้นส่วน' : 'หนี้สินและส่วนของผู้ถือหุ้น';
    const equityTerm = isLimitedPartnership ? 'ส่วนของผู้เป็นหุ้นส่วน' : 'ส่วนของผู้ถือหุ้น';

    // Current year values (match ranges used in original extractor)
    const n = globalData?.noteCalculations;
    const sel = selection?.totals;
    const bankOverdraftsAndShortTermLoans = sel?.bank_overdrafts?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.bankOverdraftsAndShortTermLoans(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2001, 2009))));
    const tradeAndOtherPayables = sel?.payables?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.tradeAndOtherPayables(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.tradeAndOtherPayables.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2010, 2999)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2123))));
    const shortTermBorrowings = sel?.short_term_loans?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.shortTermBorrowings(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.shortTermBorrowings.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030))));
    const incomeTaxPayable = sel?.income_tax_payable?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.incomeTaxPayable(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.incomeTaxPayable.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045))));
    const longTermLoansFromFI = sel?.long_term_loans_fi?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.longTermLoansFromFI(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.longTermLoansFromFI.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2121, 2121))));
    const otherLongTermLoans = sel?.long_term_loans_other?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.otherLongTermLoans(n).current
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.otherLongTermLoans.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) +
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2119))));

    // Equity related values (current)
    const registeredCapital = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3000, 3009));
    const paidUpCapital = FinancialCalculations.getSingleAccountBalance(trialBalanceData, '3010');
    const openingRetainedEarnings = FinancialCalculations.getOpeningRetainedEarnings(trialBalanceData);
    const currentYearProfit = FinancialCalculations.calculateCurrentYearProfit(trialBalanceData);
    const retainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);
    const legalReserve = Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3030, 3039));

    // Previous year values
    const prevBankOverdraftsAndShortTermLoans = sel?.bank_overdrafts?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.bankOverdraftsAndShortTermLoans(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.bankOverdraftsAndShortTermLoans.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2001, 2009)));
    const prevTradeAndOtherPayables = sel?.payables?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.tradeAndOtherPayables(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.tradeAndOtherPayables.previous
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2010, 2999)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2123))));
    const prevShortTermBorrowings = sel?.short_term_loans?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.shortTermBorrowings(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.shortTermBorrowings.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)));
    const prevIncomeTaxPayable = sel?.income_tax_payable?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.incomeTaxPayable(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.incomeTaxPayable.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)));
    const prevLongTermLoansFromFI = sel?.long_term_loans_fi?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.longTermLoansFromFI(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.longTermLoansFromFI.previous
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2120, 2123)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2121, 2121))));
    const prevOtherLongTermLoans = sel?.long_term_loans_other?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.liabilities.otherLongTermLoans(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.liabilities.otherLongTermLoans.previous
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) +
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2119))));
    const prevPaidUpCapital = globalData
      ? globalData.balanceSheetTotals.equity.paidUpCapital.previous
      : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3010, 3010));
    const prevRetainedEarnings = globalData
      ? globalData.balanceSheetTotals.equity.retainedEarnings.previous
      : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3020, 3020));

    // Initialize worksheet data with headers
    const worksheetData: (string | number | { f: string })[][] = [
      [companyInfo.name, '', '', '', '', '', '', '', '', ''],
      ['งบแสดงฐานะการเงิน (ต่อ)', '', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', liabilityAndEquityTerm, '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '', '']
    ];

    // Initialize cell tracker for formula generation
    const cellTracker: CellTracker = {
      currentRow: worksheetData.length + 1,
      currentLiabilitiesRows: [],
      nonCurrentLiabilitiesRows: [],
      equityDataRows: [],
      currentLiabilitiesTotalRow: 0,
      nonCurrentLiabilitiesTotalRow: 0
    };

    // Sections
    this.buildCurrentLiabilitiesSection(
      worksheetData,
      cellTracker,
      bankOverdraftsAndShortTermLoans,
      tradeAndOtherPayables,
      shortTermBorrowings,
      incomeTaxPayable,
      prevBankOverdraftsAndShortTermLoans,
      prevTradeAndOtherPayables,
      prevShortTermBorrowings,
      prevIncomeTaxPayable,
      processingType
    );

    this.buildNonCurrentLiabilitiesSection(
      worksheetData,
      cellTracker,
      longTermLoansFromFI,
      otherLongTermLoans,
      prevLongTermLoansFromFI,
      prevOtherLongTermLoans,
      processingType
    );

    this.buildEquitySection(
      worksheetData,
      cellTracker,
      isLimitedPartnership,
      equityTerm,
      registeredCapital,
      paidUpCapital,
      retainedEarnings,
      legalReserve,
      prevPaidUpCapital,
      prevRetainedEarnings,
      companyInfo,
      processingType
    );

    // Footer
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    worksheetData.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '', '']);

    return worksheetData;
  }

  private static buildCurrentLiabilitiesSection(
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
    worksheetData.push(['', 'หนี้สินหมุนเวียน', '', '', '', '', '', '', 'หน่วย:บาท', '']);
    cellTracker.currentRow++;

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

    const currentLiabilitiesFormula = FinancialCalculations.buildSumFormula(cellTracker.currentLiabilitiesRows, 'G');
    const currentLiabilitiesFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.currentLiabilitiesRows, 'I') : '';

    worksheetData.push(['', 'รวมหนี้สินหมุนเวียน', '', '', '', '', { f: currentLiabilitiesFormula }, '', processingType === 'multi-year' ? { f: currentLiabilitiesFormulaPrev } : '', '']);
    cellTracker.currentLiabilitiesTotalRow = cellTracker.currentRow;
    cellTracker.currentRow++;
  }

  private static buildNonCurrentLiabilitiesSection(
    worksheetData: any[][],
    cellTracker: CellTracker,
    longTermLoansFromFI: number,
    otherLongTermLoans: number,
    prevLongTermLoansFromFI: number,
    prevOtherLongTermLoans: number,
    processingType: 'single-year' | 'multi-year'
  ) {
    worksheetData.push(['', 'หนี้สินไม่หมุนเวียน', '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;

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

    const nonCurrentLiabilitiesFormula = FinancialCalculations.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'G');
    const nonCurrentLiabilitiesFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.nonCurrentLiabilitiesRows, 'I') : '';

    worksheetData.push(['', 'รวมหนี้สินไม่หมุนเวียน', '', '', '', '', { f: nonCurrentLiabilitiesFormula }, '', processingType === 'multi-year' ? { f: nonCurrentLiabilitiesFormulaPrev } : '', '']);

    cellTracker.nonCurrentLiabilitiesTotalRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    const totalLiabilitiesFormula = cellTracker.currentLiabilitiesTotalRow && cellTracker.nonCurrentLiabilitiesTotalRow
      ? `G${cellTracker.currentLiabilitiesTotalRow}+G${cellTracker.nonCurrentLiabilitiesTotalRow}`
      : '0';
    const totalLiabilitiesFormulaPrev = processingType === 'multi-year' && cellTracker.currentLiabilitiesTotalRow && cellTracker.nonCurrentLiabilitiesTotalRow
      ? `I${cellTracker.currentLiabilitiesTotalRow}+I${cellTracker.nonCurrentLiabilitiesTotalRow}`
      : '';

    worksheetData.push(['', 'รวมหนี้สิน', '', '', '', '', { f: totalLiabilitiesFormula }, '', processingType === 'multi-year' ? { f: totalLiabilitiesFormulaPrev } : '', '']);
    cellTracker.totalLiabilitiesRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;
  }

  private static buildEquitySection(
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
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year'
  ) {
    const numberOfShares = (companyInfo as any)?.shares || registeredCapital;
    const shareValue = (companyInfo as any)?.shareValue || 1;
    const numberOfPaidShares = paidUpCapital / shareValue;

    worksheetData.push(['', equityTerm, '', '', '', '', '', '', '', '']);
    cellTracker.currentRow++;

    if (isLimitedPartnership) {
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

    const totalEquityFormula = FinancialCalculations.buildSumFormula(cellTracker.equityDataRows, 'G');
    const totalEquityFormulaPrev = processingType === 'multi-year' ? FinancialCalculations.buildSumFormula(cellTracker.equityDataRows, 'I') : '';

    worksheetData.push(['', `รวม${equityTerm}`, '', '', '', '', { f: totalEquityFormula }, '', processingType === 'multi-year' ? { f: totalEquityFormulaPrev } : '', '']);
    const totalEquityRow = cellTracker.currentRow;
    cellTracker.currentRow++;

    const grandTotalFormula = cellTracker.totalLiabilitiesRow
      ? `G${cellTracker.totalLiabilitiesRow}+G${totalEquityRow}`
      : `G${totalEquityRow - cellTracker.equityDataRows.length - 2}+G${totalEquityRow}`;
    const grandTotalFormulaPrev = processingType === 'multi-year'
      ? (cellTracker.totalLiabilitiesRow
          ? `I${cellTracker.totalLiabilitiesRow}+I${totalEquityRow}`
          : `I${totalEquityRow - cellTracker.equityDataRows.length - 2}+I${totalEquityRow}`)
      : '';

    worksheetData.push(['', `รวม${'หนี้สินและส่วนของผู้ถือหุ้น'}`, '', '', '', '', { f: grandTotalFormula }, '', processingType === 'multi-year' ? { f: grandTotalFormulaPrev } : '', '']);
  }
}
