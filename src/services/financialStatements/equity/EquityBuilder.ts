// ============================================================================
// EQUITY BUILDER DISPATCHER
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';
import { CorporateEquityBuilder } from './CorporateEquityBuilder';

export class EquityBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData: DetailedFinancialData
  ): any[][] {
    return CorporateEquityBuilder.build(trialBalanceData, companyInfo, processingType, globalData);
  }
}
