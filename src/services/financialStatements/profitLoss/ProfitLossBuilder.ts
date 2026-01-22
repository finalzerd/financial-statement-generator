// ============================================================================
// PROFIT & LOSS STATEMENT BUILDER (EXTRACTED)
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo, StatementResult } from '../../../types/financial';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';
import type { NoteRegistry } from '../core/types';

/**
 * Builds the Profit & Loss (Single-step) statement table.
 * Extracted verbatim from FinancialStatementGenerator.generateProfitLossStatement.
 * IMPORTANT: Preserve exact row order, labels, formulas, and column structure.
 */
export class ProfitLossBuilder {
  static build(
    _trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    selection?: SelectionFirstResult,
    noteRegistry?: NoteRegistry
  ): StatementResult {
    console.log('=== PROFITLOSS BUILDER START ===');
  const sel = selection?.totals as Record<string, { current: number; previous: number }> | undefined;
    // STRICT SELECTION-FIRST: Do not fall back to numeric ranges. If a bucket is undefined -> treat as 0.
  const get = (cat: string) => sel?.[cat]?.current ?? 0;
  const getPrev = (cat: string) => sel?.[cat]?.previous ?? 0;

    const revenue = get('revenue');
    const previousRevenue = processingType === 'multi-year' ? getPrev('revenue') : 0;

    const otherIncome = get('other_income');
    const previousOtherIncome = processingType === 'multi-year' ? getPrev('other_income') : 0;

    const costOfServices = get('detail_service_costs');
    const previousCostOfServices = processingType === 'multi-year' ? getPrev('detail_service_costs') : 0;

    const sellingExpenses = get('selling_expenses');
    const previousSellingExpenses = processingType === 'multi-year' ? getPrev('selling_expenses') : 0;

    const adminExpenses = get('admin_expenses');
    const previousAdminExpenses = processingType === 'multi-year' ? getPrev('admin_expenses') : 0;

    const otherExpenses = get('other_expenses');
    const previousOtherExpenses = processingType === 'multi-year' ? getPrev('other_expenses') : 0;

    const incomeTax = get('income_tax_expense');
    const previousIncomeTax = processingType === 'multi-year' ? getPrev('income_tax_expense') : 0;

    const financialCosts = get('financial_costs');
    const previousFinancialCosts = processingType === 'multi-year' ? getPrev('financial_costs') : 0;

    console.log('[P&L Strict] Components (no numeric fallback):', { revenue, otherIncome, costOfServices, sellingExpenses, adminExpenses, otherExpenses, incomeTax, financialCosts });

    // Get note references from registry
    const otherIncomeNote = noteRegistry?.otherIncome?.toString() || '';
    const expensesByNatureNote = noteRegistry?.expensesByNature?.toString() || '';

    const plData = [
      [`${companyInfo.name}`, '', '', '', '', '', '', '', ''],
      ['งบกำไรขาดทุน จำแนกค่าใช้จ่ายตามหน้าที่ - แบบขั้นเดียว', '', '', '', '', '', '', '', ''],
      [`สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ ${companyInfo.reportingPeriodStartDate || '1 มกราคม'} ${companyInfo.reportingYear} ถึงวันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', 'หมายเหตุ', '', '', 'หน่วย:บาท'],
      ['', '', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : ''],
      ['', 'รายได้', '', '', '', '', '', '', ''],
      ['', '', 'รายได้จากการขายหรือการให้บริการ', '', '', '', revenue, '', processingType === 'multi-year' ? previousRevenue : ''],
      ['', '', 'รายได้อื่น', '', '', otherIncomeNote, otherIncome, '', processingType === 'multi-year' ? previousOtherIncome : ''],
      ['', 'รวมรายได้', '', '', '', '', { f: 'SUM(G8:G9)' }, '', processingType === 'multi-year' ? { f: 'SUM(I8:I9)' } : ''],
      ['', '', '', '', '', '', '', '', ''],
  ['', 'ค่าใช้จ่าย', '', '', '', '', '', '', ''],
  ['', '', 'ต้นทุนขายหรือต้นทุนการให้บริการ', '', '', '', costOfServices, '', processingType === 'multi-year' ? previousCostOfServices : ''],
  ['', '', 'ค่าใช้จ่ายในการขาย', '', '', '', sellingExpenses, '', processingType === 'multi-year' ? previousSellingExpenses : ''],
  ['', '', 'ค่าใช้จ่ายในการบริหาร', '', '', '', adminExpenses, '', processingType === 'multi-year' ? previousAdminExpenses : ''],
  ['', '', 'ค่าใช้จ่ายอื่น', '', '', '', otherExpenses, '', processingType === 'multi-year' ? previousOtherExpenses : ''],
  ['', 'รวมค่าใช้จ่าย', '', '', '', expensesByNatureNote, { f: 'SUM(G13:G16)' }, '', processingType === 'multi-year' ? { f: 'SUM(I13:I16)' } : ''],
  ['', 'กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้', '', '', '', '', { f: 'G10-G17' }, '', processingType === 'multi-year' ? { f: 'I10-I17' } : ''],
  ['', 'ต้นทุนทางการเงิน', '', '', '', '', financialCosts, '', processingType === 'multi-year' ? previousFinancialCosts : ''],
  ['', 'กำไรก่อนภาษีเงินได้', '', '', '', '', { f: 'G18-G19' }, '', processingType === 'multi-year' ? { f: 'I18-I19' } : ''],
  ['', 'ภาษีเงินได้', '', '', '', '', incomeTax, '', processingType === 'multi-year' ? previousIncomeTax : ''],
  ['', 'กำไร(ขาดทุน)สุทธิ', '', '', '', '', { f: 'G20-G21' }, '', processingType === 'multi-year' ? { f: 'I20-I21' } : ''],
      ['', 'กำไร (ขาดทุน) สะสมยังไม่แบ่งต้นงวด', '', '', '', '', processingType === 'multi-year' ? { f: 'I24' } : '', '', processingType === 'multi-year' ? '' : ''],
      ['', 'กำไร (ขาดทุน) สะสมยังไม่แบ่งปลายงวด', '', '', '', '', { f: 'G22+G23' }, '', processingType === 'multi-year' ? { f: 'I22+I23' } : ''],
      ['', '', '', '', '', '', '', '', ''],
      ['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '']
    ];
    
    // Add director signature block
    const meetingNumber = companyInfo.approvalMeetingNumber || '....';
    const meetingDate = companyInfo.approvalMeetingDate || '....';
    plData.push([`งบการเงินนี้ได้รับการอนุมัติจากที่ประชุมสามัญผู้ถือหุ้นครั้งที่ ${meetingNumber} เมื่อวันที่ ${meetingDate}`, '', '', '', '', '', '', '', '']);
    plData.push(['ขอรับรองว่าเป็นรายการอันถูกต้องและเป็นความจริง', '', '', '', '', '', '', '', '']);
    plData.push(['', '', '', '', '', '', '', '', '']);
    plData.push(['', '', '', '', '', '', '', '', '']);
    
    const signatureRowIndex = plData.length + 1;
    const signatureTitle = companyInfo.type === 'ห้างหุ้นส่วนจำกัด' ? 'หุ้นส่วนผู้จัดการ' : 'กรรมการตามอำนาจ';
    plData.push([`ลงชื่อ ……………………..................................... ${signatureTitle}`, '', '', '', '', '', '', '', '']);
    
    const directorNameRowIndex = plData.length + 1;
    const directorName = companyInfo.directorName || '...........................';
    plData.push([`(${directorName})`, '', '', '', '', '', '', '', '']);

    return {
      data: plData,
      signatureRows: [signatureRowIndex, directorNameRowIndex]
    };
  }
}
