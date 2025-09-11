// ============================================================================
// EQUITY BUILDER DISPATCHER
// ============================================================================
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';
import { PartnershipEquityBuilder } from './PartnershipEquityBuilder';
import { CorporateEquityBuilder } from './CorporateEquityBuilder';

export class EquityBuilder {
  static build(
    trialBalanceData: TrialBalanceEntry[],
    companyInfo: CompanyInfo,
    processingType: 'single-year' | 'multi-year',
    globalData: DetailedFinancialData
  ): any[][] {
    const isLimitedPartnership = companyInfo.type === 'ห้างหุ้นส่วนจำกัด';
    if (isLimitedPartnership) {
      return PartnershipEquityBuilder.build(trialBalanceData, companyInfo, processingType);
    }
    return CorporateEquityBuilder.build(trialBalanceData, companyInfo, processingType, globalData);
  }
}
