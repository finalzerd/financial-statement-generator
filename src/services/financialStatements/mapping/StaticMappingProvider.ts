import type { IAccountMappingProvider } from './IAccountMappingProvider';
import type { AccountMappingRules, SubCategoryRuleContainer } from '../../../types/accountMapping';

// Default rules mirroring current hard-coded numeric ranges
const DEFAULT_RULES: Record<string, AccountMappingRules> = {
  cash: { ranges: [{ from: 1000, to: 1099 }] },
  receivables: { ranges: [{ from: 1140, to: 1215 }] },
  inventory: { ranges: [{ from: 1500, to: 1519 }], includes: [1510] },
  // Detail note mappings for cost of goods sold
  inventory_purchases: { includes: [5010] },
  inventory_purchase_returns: { includes: [5010.1] },
  inventory_purchase_discounts: { includes: [5010.2] },
  prepaid_expenses: { ranges: [{ from: 1400, to: 1439 }] },
  ppe_cost: { ranges: [{ from: 1600, to: 1629 }] },
  ppe_accum_depr: { ranges: [{ from: 1630, to: 1659 }] },
  other_assets: { ranges: [{ from: 1660, to: 1700 }] },
  bank_overdrafts: { ranges: [{ from: 2001, to: 2009 }] },
  payables: { ranges: [{ from: 2010, to: 2999 }], excludes: [2030, 2045, 2050, 2051, 2052, 2100, 2101, 2102, 2103, 2120, 2121, 2122, 2123] },
  short_term_loans: { includes: [2030] },
  income_tax_payable: { includes: [2045] },
  long_term_loans_fi: { ranges: [{ from: 2120, to: 2123 }], excludes: [2121] },
  long_term_loans_other: { includes: [2050, 2051, 2052, 2100, 2101, 2102, 2103] },
  other_income: { ranges: [{ from: 4110, to: 4999 }] },
  // Main revenue from sales/services
  revenue: { ranges: [{ from: 4000, to: 4099 }] },
  // Other expenses (follow classifier fallback):
  // (5351-5354), (5358-5361), 5364, (5366-5999)
  other_expenses: { 
    ranges: [
      { from: 5351, to: 5354 },
      { from: 5358, to: 5361 },
      { from: 5366, to: 5999 }
    ],
    includes: [5364]
  },
  paid_up_capital: { includes: [3010] },
  legal_reserve: { ranges: [{ from: 3030, to: 3039 }] },
};

export class StaticMappingProvider implements IAccountMappingProvider {
  getRules(noteType: string): AccountMappingRules | null {
    return DEFAULT_RULES[noteType] || null;
  }
  getSubCategoryRules(_noteType: string): SubCategoryRuleContainer | null {
    return null; // Static provider has no sub-category rules
  }
  isActive(_noteType: string): boolean {
    return true;
  }
}
