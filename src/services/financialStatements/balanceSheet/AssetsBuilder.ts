import { FinancialCalculations } from '../../financialCalculations';
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';
import { BalanceSheetLinkMap, NOTE_FIRST_MODE } from '../core/linking/balanceSheetLinkMap';

// Builds the Balance Sheet (Assets) worksheet data.
// This mirrors the original logic from FinancialStatementGenerator.generateBalanceSheetAssets
// without changing behavior or output shape.
export class AssetsBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData?: DetailedFinancialData
  ): (string | number | { f: string })[][] {
    // Calculate current year asset balances using note-first link map when enabled
    const n = globalData?.noteCalculations;
    const cashAndCashEquivalents = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.cashAndCashEquivalents(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.cashAndCashEquivalents.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099)));
    const tradeReceivables = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.tradeReceivables(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.tradeReceivables.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215)));
    const inventory = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.inventory(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.inventory.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519)));
    const prepaidExpenses = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.prepaidExpenses(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.prepaidExpenses.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1400, 1439)));
    const landBuildingsEquipment = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.propertyPlantEquipment(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.propertyPlantEquipment.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1659)));
    const otherAssets = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.otherAssets(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.otherAssets.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700)));

    // Calculate previous year asset balances using previousBalance field
    const prevCashAndCashEquivalents = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.cashAndCashEquivalents(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.cashAndCashEquivalents.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099));
    const prevTradeReceivables = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.tradeReceivables(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.tradeReceivables.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215));
    const prevInventory = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.inventory(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.inventory.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519));
    const prevPrepaidExpenses = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.prepaidExpenses(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.prepaidExpenses.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439));
    const prevLandBuildingsEquipment = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.propertyPlantEquipment(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.propertyPlantEquipment.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1659));
    const prevOtherAssets = (NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.otherAssets(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.otherAssets.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700));

    // Initialize worksheet data with headers
    const worksheetData: (string | number | { f: string })[][] = [
      [companyInfo.name, '', '', '', '', '', '', '', '', ''],
      ['งบแสดงฐานะการเงิน', '', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', `ณ วันที่ 31 ธันวาคม ${companyInfo.reportingYear - 1}`, ''],
      ['', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', 'หมายเหตุ', '', '', 'หน่วย:บาท', ''],
      ['', 'สินทรัพย์', '', '', '', '', `${companyInfo.reportingYear}`, '', processingType === 'multi-year' ? `${companyInfo.reportingYear - 1}` : '', '']
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
}
