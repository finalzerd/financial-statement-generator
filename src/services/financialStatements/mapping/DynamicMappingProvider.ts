import type { IAccountMappingProvider } from './IAccountMappingProvider';
import type { AccountMappingRules } from '../../../types/accountMapping';

export interface CompanyMappingRecord {
  noteType: string;
  accountRanges: AccountMappingRules;
  isActive?: boolean;
}

export class DynamicMappingProvider implements IAccountMappingProvider {
  private rules: Map<string, AccountMappingRules> = new Map();
  private active: Map<string, boolean> = new Map();

  constructor(mappings: CompanyMappingRecord[] = []) {
    for (const m of mappings) {
      if (m && m.noteType && m.accountRanges) {
        this.rules.set(m.noteType, m.accountRanges);
        this.active.set(m.noteType, m.isActive !== false);
      }
    }
  }

  getRules(noteType: string): AccountMappingRules | null {
    return this.rules.get(noteType) || null;
  }

  isActive(noteType: string): boolean {
    return this.active.get(noteType) ?? true;
  }
}
