// ============================================================================
// CORPORATE EQUITY STATEMENT BUILDER (EXTRACTED)
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';

export class CorporateEquityBuilder {
  static build(
    _trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData: DetailedFinancialData
  ): any[][] {
    const isMultiYear = processingType === 'multi-year';
    const currentYear = companyInfo.reportingYear;
    const previousYear = currentYear - 1;

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
      const prevYearOpeningRetained = '';
      const prevYearProfit = '';
      const prevYearTotalEquity = paidUpCapitalPrevious > 0 ? paidUpCapitalPrevious + (openingRetainedEarnings || 0) : '';

      result.push([`ยอดคงเหลือ ณ วันที่ 1 มกราคม ${previousYear}`, '', paidUpCapitalPrevious || '', '', '', prevYearOpeningRetained, '', '', prevYearTotalEquity]);
      result.push([`กำไร (ขาดทุน) สุทธิ สำหรับปี ${previousYear}`, '', '', '', '', prevYearProfit, '', '', '']);
  // Year-end (previous year) with formulas
  result.push([`ยอดคงเหลือ ณ วันที่ 31 ธันวาคม ${previousYear}`, '', { f: 'C8+C9' }, '', '', { f: 'F8+F9' }, '', '', { f: 'C10+F10' }]);
      result.push(['', '', '', '', '', '', '', '', '']);
      result.push(['', '', '', '', '', '', '', '', '']);
      rowIndex = 12;
    }

    const openingTotalCurrent = paidUpCapitalCurrent + openingRetainedEarnings;

  result.push([`ยอดคงเหลือ ณ วันที่ 1 มกราคม ${currentYear}`, '', paidUpCapitalCurrent, '', '', openingRetainedEarnings, '', '', openingTotalCurrent]);
    result.push([`กำไร (ขาดทุน) สุทธิ สำหรับปี ${currentYear}`, '', '', '', '', currentYearProfit, '', '', currentYearProfit]);
  // Year-end (current year) with formulas
  result.push([`ยอดคงเหลือ ณ วันที่ 31 ธันวาคม ${currentYear}`, '', paidUpCapitalCurrent, '', '', retainedEarningsCurrent, '', '', { f: `C${rowIndex + 3}+F${rowIndex + 3}` }]);

    result.push(['', '', '', '', '', '', '', '', '']);
    result.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '']);

    return result;
  }
}
