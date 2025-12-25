// ============================================================================
// PARTNERSHIP EQUITY STATEMENT BUILDER (EXTRACTED)
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';

export class PartnershipEquityBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    _processingType: 'single-year' | 'multi-year'
  ): any[][] {
    const sumAccountsByRange = (start: number, end: number) => {
      const matching = trialBalanceData.filter(e => {
        const code = parseInt(e.accountCode || '0');
        return code >= start && code <= end;
      });
      return matching.reduce((s, e) => s + (e.balance || 0), 0);
    };

    const totalCapital = Math.abs(sumAccountsByRange(3010, 3019));

    const openingRetainedEarningsRaw = sumAccountsByRange(3020, 3020);
    const openingRetainedEarnings = -openingRetainedEarningsRaw; // flip credit

    const revenueAccounts = trialBalanceData.filter(e => e.accountCode?.startsWith('4'));
    const currentYearRevenue = revenueAccounts.reduce((sum, e) => sum + ((e.creditAmount || 0) - (e.debitAmount || 0)), 0);

    const expenseAccounts = trialBalanceData.filter(e => e.accountCode?.startsWith('5'));
    const currentYearExpenses = expenseAccounts.reduce((sum, e) => sum + ((e.debitAmount || 0) - (e.creditAmount || 0)), 0);

    const currentYearProfit = currentYearRevenue - currentYearExpenses;
    const retainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);

    const partner1Capital = totalCapital / 2;
    const partner2Capital = totalCapital / 2;

    return [
      [`${companyInfo.name}`, '', '', '', '', '', ''],
      ['งบแสดงการเปลี่ยนแปลงส่วนของผู้เป็นหุ้นส่วน', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ ${companyInfo.reportingPeriodStartDate || '1 มกราคม'} ${companyInfo.reportingYear} ถึงวันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`, '', '', '', '', '', ''],
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
}
