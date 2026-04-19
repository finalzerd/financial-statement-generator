import { FinancialCalculations } from '../../financialCalculations';
import type { TrialBalanceEntry, CompanyInfo, BalanceSheetResult } from '../../../types/financial';
import type { DetailedFinancialData, NoteRegistry } from '../core/types';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';
import { BalanceSheetLinkMap, NOTE_FIRST_MODE } from '../core/linking/balanceSheetLinkMap';

// Builds the Balance Sheet (Assets) worksheet data.
// This mirrors the original logic from FinancialStatementGenerator.generateBalanceSheetAssets
// without changing behavior or output shape.
export class AssetsBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData?: DetailedFinancialData,
    selection?: SelectionFirstResult,
    noteRegistry?: NoteRegistry
  ): BalanceSheetResult {
    // Calculate current year asset balances, preferring Selection-First totals when available
    const n = globalData?.noteCalculations;
    const sel = selection?.totals;
    const cashAndCashEquivalents = sel?.cash?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.cashAndCashEquivalents(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.cashAndCashEquivalents.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099))));
    const tradeReceivables = sel?.receivables?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.tradeReceivables(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.tradeReceivables.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215))));
    const assetShortTermLoans = sel?.asset_short_term_loans?.current ?? (
      globalData?.noteCalculations?.assetShortTermLoans?.current ??
      Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1141, 1141))
    );
    const inventory = sel?.inventory?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.inventory(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.inventory.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519))));
    const otherCurrentAssets = sel?.other_current_assets?.current ?? 0;
    const prevOtherCurrentAssets = sel?.other_current_assets?.previous ?? 0;

    // Prepaid expenses removed from presentation
    // Investment Property (net) = cost - accum depreciation
    const investmentPropertyCostCurrent = sel?.investment_property_cost?.current ?? 0;
    const investmentPropertyAccumCurrent = sel?.investment_property_accum_depr?.current ?? 0;
    const investmentProperty = (sel && (sel.investment_property_cost || sel.investment_property_accum_depr))
      ? (investmentPropertyCostCurrent - investmentPropertyAccumCurrent)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.investmentProperty(n).current
          : (globalData
              ? globalData.noteCalculations.investmentProperty.netBookValue.current
              : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1700, 1759))));
    
    // Intangible Assets (net) = cost - accum amortization
    const intangibleAssetsCostCurrent = sel?.intangible_assets_cost?.current ?? 0;
    const intangibleAssetsAccumCurrent = sel?.intangible_assets_accum_amort?.current ?? 0;
    const intangibleAssets = (sel && (sel.intangible_assets_cost || sel.intangible_assets_accum_amort))
      ? (intangibleAssetsCostCurrent - intangibleAssetsAccumCurrent)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.intangibleAssets(n).current
          : (globalData
              ? globalData.noteCalculations.intangibleAssets.netBookValue.current
              : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1800, 1859))));
    
    // PPE (net) = cost - accum depreciation; if selection not present, fallback existing
    const ppeCostCurrent = sel?.ppe_cost?.current ?? 0;
    const ppeAccumCurrent = sel?.ppe_accum_depr?.current ?? 0;
    const landBuildingsEquipment = (sel && (sel.ppe_cost || sel.ppe_accum_depr))
      ? (ppeCostCurrent - ppeAccumCurrent)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.propertyPlantEquipment(n).current
          : (globalData
              ? globalData.balanceSheetTotals.assets.propertyPlantEquipment.current
              : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1659))));
    const otherAssets = sel?.other_assets?.current ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.otherAssets(n).current
      : (globalData
          ? globalData.balanceSheetTotals.assets.otherAssets.current
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700))));

    // Calculate previous year asset balances using previousBalance field
    const prevCashAndCashEquivalents = sel?.cash?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.cashAndCashEquivalents(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.cashAndCashEquivalents.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099)));
    const prevTradeReceivables = sel?.receivables?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.tradeReceivables(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.tradeReceivables.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215)));
    const prevAssetShortTermLoans = sel?.asset_short_term_loans?.previous ?? (
      globalData?.noteCalculations?.assetShortTermLoans?.previous ??
      FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1141, 1141)
    );
    const prevInventory = sel?.inventory?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.inventory(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.inventory.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519)));
    // Previous prepaid expenses removed from presentation
    const investmentPropertyCostPrev = sel?.investment_property_cost?.previous ?? 0;
    const investmentPropertyAccumPrev = sel?.investment_property_accum_depr?.previous ?? 0;
    const prevInvestmentProperty = (sel && (sel.investment_property_cost || sel.investment_property_accum_depr))
      ? (investmentPropertyCostPrev - investmentPropertyAccumPrev)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.investmentProperty(n).previous
          : (globalData
              ? globalData.noteCalculations.investmentProperty.netBookValue.previous
              : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1700, 1759)));
    
    const intangibleAssetsCostPrev = sel?.intangible_assets_cost?.previous ?? 0;
    const intangibleAssetsAccumPrev = sel?.intangible_assets_accum_amort?.previous ?? 0;
    const prevIntangibleAssets = (sel && (sel.intangible_assets_cost || sel.intangible_assets_accum_amort))
      ? (intangibleAssetsCostPrev - intangibleAssetsAccumPrev)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.intangibleAssets(n).previous
          : (globalData
              ? globalData.noteCalculations.intangibleAssets.netBookValue.previous
              : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1800, 1859)));
    
    const ppeCostPrev = sel?.ppe_cost?.previous ?? 0;
    const ppeAccumPrev = sel?.ppe_accum_depr?.previous ?? 0;
    const prevLandBuildingsEquipment = (sel && (sel.ppe_cost || sel.ppe_accum_depr))
      ? (ppeCostPrev - ppeAccumPrev)
      : ((NOTE_FIRST_MODE && n)
          ? BalanceSheetLinkMap.assets.propertyPlantEquipment(n).previous
          : (globalData
              ? globalData.balanceSheetTotals.assets.propertyPlantEquipment.previous
              : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1659)));
    const prevOtherAssets = sel?.other_assets?.previous ?? ((NOTE_FIRST_MODE && n)
      ? BalanceSheetLinkMap.assets.otherAssets(n).previous
      : (globalData
          ? globalData.balanceSheetTotals.assets.otherAssets.previous
          : FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700)));

    // Initialize worksheet data with headers
    const worksheetData: (string | number | { f: string })[][] = [
      [companyInfo.name, '', '', '', '', '', '', '', '', ''],
      ['งบฐานะการเงิน', '', '', '', '', '', '', '', '', ''],
      [`ณ วันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear}`, '', '', '', '', '', '', '', `ณ วันที่ ${companyInfo.reportingPeriodEndDate || '31 ธันวาคม'} ${companyInfo.reportingYear - 1}`, ''],
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
    // Always show assets rows even when amounts are zero
    worksheetData.push(['', '', 'เงินสดและรายการเทียบเท่าเงินสด', '', '', noteRegistry?.cash?.toString() || '', cashAndCashEquivalents, '', processingType === 'multi-year' ? prevCashAndCashEquivalents : '', '']);
    currentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น', '', '', noteRegistry?.receivables?.toString() || '', tradeReceivables, '', processingType === 'multi-year' ? prevTradeReceivables : '', '']);
    currentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'เงินให้กู้ยืมระยะสั้น', '', '', noteRegistry?.assetShortTermLoans?.toString() || '', assetShortTermLoans, '', processingType === 'multi-year' ? prevAssetShortTermLoans : '', '']);
    currentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'สินค้าคงเหลือ', '', '', '', inventory, '', processingType === 'multi-year' ? prevInventory : '', '']);
    currentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'สินทรัพย์หมุนเวียนอื่น', '', '', noteRegistry?.otherCurrentAssets?.toString() || '', otherCurrentAssets, '', processingType === 'multi-year' ? prevOtherCurrentAssets : '', '']);
    currentAssetRows.push(currentRow);
    currentRow++;

    // Removed 'ค่าใช้จ่ายจ่ายล่วงหน้า' as it is not an accounting subject

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

    worksheetData.push(['', '', 'อสังหาริมทรัพย์เพื่อการลงทุน (สุทธิ)', '', '', noteRegistry?.investmentProperty?.toString() || '', investmentProperty, '', processingType === 'multi-year' ? prevInvestmentProperty : '', '']);
    nonCurrentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'ที่ดิน อาคาร และอุปกรณ์ (สุทธิ)', '', '', noteRegistry?.ppe?.toString() || '', landBuildingsEquipment, '', processingType === 'multi-year' ? prevLandBuildingsEquipment : '', '']);
    nonCurrentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'สินทรัพย์ไม่มีตัวตน (สุทธิ)', '', '', noteRegistry?.intangibleAssets?.toString() || '', intangibleAssets, '', processingType === 'multi-year' ? prevIntangibleAssets : '', '']);
    nonCurrentAssetRows.push(currentRow);
    currentRow++;

    // Long-term loans given (asset)
    const assetLongTermLoans = sel?.asset_long_term_loans?.current ?? (
      globalData?.noteCalculations?.assetLongTermLoans?.current ??
      Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1710, 1710))
    );
    const prevAssetLongTermLoans = sel?.asset_long_term_loans?.previous ?? (
      globalData?.noteCalculations?.assetLongTermLoans?.previous ??
      FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1710, 1710)
    );

    worksheetData.push(['', '', 'เงินให้กู้ยืมระยะยาว', '', '', noteRegistry?.assetLongTermLoans?.toString() || '', assetLongTermLoans, '', processingType === 'multi-year' ? prevAssetLongTermLoans : '', '']);
    nonCurrentAssetRows.push(currentRow);
    currentRow++;

    worksheetData.push(['', '', 'สินทรัพย์ไม่หมุนเวียนอื่น', '', '', noteRegistry?.otherAssets?.toString() || '', otherAssets, '', processingType === 'multi-year' ? prevOtherAssets : '', '']);
    nonCurrentAssetRows.push(currentRow);
    currentRow++;

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

    // Add footer with director signature block
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']); // Blank spacer
    worksheetData.push(['หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', '', '', '', '', '', '', '', '', '']);
    
    // Approval meeting line
    const meetingNumber = companyInfo.approvalMeetingNumber || '....';
    const meetingDate = companyInfo.approvalMeetingDate || '....';
    const approvalText = companyInfo.type === 'ห้างหุ้นส่วนจำกัด' 
      ? `งบการเงินนี้ได้รับการอนุมัติจากที่ประชุมของผู้เป็นหุ้นส่วนครั้งที่ ${meetingNumber} เมื่อวันที่ ${meetingDate}`
      : `งบการเงินนี้ได้รับการอนุมัติจากที่ประชุมสามัญผู้ถือหุ้นครั้งที่ ${meetingNumber} เมื่อวันที่ ${meetingDate}`;
    worksheetData.push([approvalText, '', '', '', '', '', '', '', '', '']);
    
    // Certification line
    worksheetData.push(['ขอรับรองว่าเป็นรายการอันถูกต้องและเป็นความจริง', '', '', '', '', '', '', '', '', '']);
    
    // Two blank rows
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    worksheetData.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Signature line (row to be center-aligned)
    const signatureRowIndex = worksheetData.length + 1; // 1-based
    const signatureTitle = companyInfo.type === 'ห้างหุ้นส่วนจำกัด' ? 'หุ้นส่วนผู้จัดการ' : 'กรรมการตามอำนาจ';
    worksheetData.push([`ลงชื่อ ……………………..................................... ${signatureTitle}`, '', '', '', '', '', '', '', '', '']);
    
    // Director name line (row to be center-aligned)
    const directorNameRowIndex = worksheetData.length + 1; // 1-based
    const directorName = companyInfo.directorName || '...........................';
    worksheetData.push([`(${directorName})`, '', '', '', '', '', '', '', '', '']);

    return {
      data: worksheetData,
      signatureRows: [signatureRowIndex, directorNameRowIndex]
    };
  }
}
