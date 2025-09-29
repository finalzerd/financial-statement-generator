import type { AccountMappingRules, SubCategoryRuleContainer } from '../../../types/accountMapping';

/**
 * Provider interface for company-specific account mapping rules.
 * Implementations can return dynamic rules from DB or static defaults.
 */
export interface IAccountMappingProvider {
  /** Return mapping rules for a given note type key, or null if not available. */
  getRules(noteType: string): AccountMappingRules | null;
  /** Return sub-category rules (currently only implemented for 'cash'), or null if none. */
  getSubCategoryRules(noteType: string): SubCategoryRuleContainer | null;
  /** Whether a mapping for a given note type is active. If unknown, assume true. */
  isActive(noteType: string): boolean;
}
