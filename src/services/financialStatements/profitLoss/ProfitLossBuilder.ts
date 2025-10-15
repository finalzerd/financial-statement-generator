// ============================================================================
// PROFIT & LOSS STATEMENT BUILDER (EXTRACTED)
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import { FinancialCalculations } from '../../financialCalculations';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';

/**
 * Builds the Profit & Loss (Single-step) statement table.
 * Extracted verbatim from FinancialStatementGenerator.generateProfitLossStatement.
 * IMPORTANT: Preserve exact row order, labels, formulas, and column structure.
 */
export class ProfitLossBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    selection?: SelectionFirstResult
  ): any[][] {
    console.log('=== PROFITLOSS BUILDER START ===');

    const sel = selection?.totals;
    const revenue = sel?.revenue?.current ?? Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4099));
    const previousRevenue = processingType === 'multi-year' ? (sel?.revenue?.previous ?? 0) : 0;

    const otherIncome = sel?.other_income?.current ?? Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4100, 4999));
    const previousOtherIncome = processingType === 'multi-year' ? (sel?.other_income?.previous ?? 0) : 0;

    const costOfServices = sel?.detail_service_costs?.current ?? Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5099));
    const previousCostOfServices = processingType === 'multi-year' ? (sel?.detail_service_costs?.previous ?? 0) : 0;

    const adminExpenses = sel?.admin_expenses?.current ?? Math.abs(
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
    );
    const previousAdminExpenses = processingType === 'multi-year' ? (sel?.admin_expenses?.previous ?? 0) : 0;

    const otherExpenses = sel?.other_expenses?.current ?? Math.abs(
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
      FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
    );
    const previousOtherExpenses = processingType === 'multi-year' ? (sel?.other_expenses?.previous ?? 0) : 0;

    const incomeTax = sel?.income_tax_expense?.current ?? Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5910, 5910));
    const previousIncomeTax = processingType === 'multi-year' ? (sel?.income_tax_expense?.previous ?? 0) : 0;

    const financialCosts = sel?.financial_costs?.current ?? Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5920, 5929));
    const previousFinancialCosts = processingType === 'multi-year' ? (sel?.financial_costs?.previous ?? 0) : 0;

    console.log('P&L components calculated:', { revenue, otherIncome, costOfServices, adminExpenses, otherExpenses, incomeTax, financialCosts });

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
}
