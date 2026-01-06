import type { IAccountMappingProvider } from './IAccountMappingProvider';
import type { AccountMappingRules, SubCategoryRuleContainer } from '../../../types/accountMapping';

export interface CompanyMappingRecord {
  noteType: string;
  accountRanges: AccountMappingRules;
  subCategoryRules?: SubCategoryRuleContainer | null; // JSON from DB (column: sub_category_rules)
  isActive?: boolean;
}

export class DynamicMappingProvider implements IAccountMappingProvider {
  private rules: Map<string, AccountMappingRules> = new Map();
  private active: Map<string, boolean> = new Map();
  private subCategoryRules: Map<string, SubCategoryRuleContainer> = new Map();

  constructor(mappings: CompanyMappingRecord[] = []) {
    console.log(`[DynamicMappingProvider] Initializing with ${mappings.length} mappings`);
    for (const m of mappings) {
      if (!m || !m.noteType || !m.accountRanges) continue;
      this.rules.set(m.noteType, m.accountRanges);
      this.active.set(m.noteType, m.isActive !== false);
      if (m.subCategoryRules && typeof m.subCategoryRules === 'object') {
        this.subCategoryRules.set(m.noteType, m.subCategoryRules);
      }
      console.log(`[DynamicMappingProvider] Loaded mapping: ${m.noteType}`, m.accountRanges);
    }
  }

  getRules(noteType: string): AccountMappingRules | null {
    return this.rules.get(noteType) || null;
  }

  getSubCategoryRules(noteType: string): SubCategoryRuleContainer | null {
    return this.subCategoryRules.get(noteType) || null;
  }

  isActive(noteType: string): boolean {
    return this.active.get(noteType) ?? true;
  }
}
