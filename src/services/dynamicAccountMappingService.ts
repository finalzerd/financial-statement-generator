// ============================================================================
// DYNAMIC ACCOUNT MAPPING SERVICE
// ============================================================================

import type { TrialBalanceEntry } from '../types/financial';
import type { CompanyAccountMapping } from '../types/accountMapping';
import { AccountMappingUtils } from '../types/accountMapping';
import { ApiService } from './apiService';

/**
 * Service for handling dynamic account mappings in financial statement generation
 * This replaces the hardcoded account ranges with flexible company-specific mappings
 */
export class DynamicAccountMappingService {
  private static companyMappingsCache: Map<number, Map<string, CompanyAccountMapping>> = new Map();

  /**
   * Load and cache company account mappings
   */
  static async loadCompanyMappings(companyId: number): Promise<Map<string, CompanyAccountMapping>> {
    // Check cache first
    if (this.companyMappingsCache.has(companyId)) {
      return this.companyMappingsCache.get(companyId)!;
    }

    try {
      const mappings: CompanyAccountMapping[] = await ApiService.getCompanyAccountMappings(companyId.toString());
      const mappingMap = new Map(
        mappings
          .filter((m: CompanyAccountMapping) => m.isActive)
          .map((m: CompanyAccountMapping) => [m.noteType, m])
      );
      
      // Cache the mappings
      this.companyMappingsCache.set(companyId, mappingMap);
      return mappingMap;
    } catch (error) {
      console.error('Failed to load company mappings, using fallback:', error);
      // Return empty map if loading fails
      return new Map();
    }
  }

  /**
   * Clear cached mappings (call when mappings are updated)
   */
  static clearCache(companyId?: number) {
    if (companyId) {
      this.companyMappingsCache.delete(companyId);
    } else {
      this.companyMappingsCache.clear();
    }
  }

  /**
   * Filter trial balance entries based on company-specific mapping rules
   */
  static filterAccountsByMapping(
    trialBalanceData: TrialBalanceEntry[], 
    mapping: CompanyAccountMapping
  ): TrialBalanceEntry[] {
    if (!mapping || !mapping.isActive) {
      return [];
    }

    return trialBalanceData.filter(entry => 
      AccountMappingUtils.doesAccountMatch(entry.accountCode || '0', mapping.accountRanges)
    );
  }

  /**
   * Get accounts for a specific note type using dynamic mappings
   */
  static async getAccountsForNoteType(
    companyId: number,
    noteType: string,
    trialBalanceData: TrialBalanceEntry[]
  ): Promise<TrialBalanceEntry[]> {
    const mappings = await this.loadCompanyMappings(companyId);
    const mapping = mappings.get(noteType);
    
    if (!mapping) {
      console.warn(`No mapping found for note type: ${noteType}`);
      return [];
    }

    return this.filterAccountsByMapping(trialBalanceData, mapping);
  }

  /**
   * Calculate total balance for a note type using dynamic mappings
   */
  static async calculateNoteTotal(
    companyId: number,
    noteType: string,
    trialBalanceData: TrialBalanceEntry[],
    balanceField: 'balance' | 'previousBalance' = 'balance'
  ): Promise<number> {
    const accounts = await this.getAccountsForNoteType(companyId, noteType, trialBalanceData);
    
    return Math.abs(accounts.reduce((sum, entry) => {
      const value = balanceField === 'balance' 
        ? (entry.balance || entry.currentBalance || 0)
        : (entry.previousBalance || 0);
      return sum + value;
    }, 0));
  }

  /**
   * Get all note mappings with their calculated totals
   */
  static async getAllNoteTotals(
    companyId: number,
    trialBalanceData: TrialBalanceEntry[]
  ): Promise<Record<string, { current: number; previous: number; accounts: TrialBalanceEntry[] }>> {
    const mappings = await this.loadCompanyMappings(companyId);
    const results: Record<string, { current: number; previous: number; accounts: TrialBalanceEntry[] }> = {};

    for (const [noteType, mapping] of mappings) {
      const accounts = this.filterAccountsByMapping(trialBalanceData, mapping);
      
      const current = Math.abs(accounts.reduce((sum, entry) => 
        sum + (entry.balance || entry.currentBalance || 0), 0
      ));
      
      const previous = Math.abs(accounts.reduce((sum, entry) => 
        sum + (entry.previousBalance || 0), 0
      ));

      results[noteType] = { current, previous, accounts };
    }

    return results;
  }

  /**
   * Validate that all trial balance accounts are mapped to notes
   */
  static async validateAccountCoverage(
    companyId: number,
    trialBalanceData: TrialBalanceEntry[]
  ): Promise<{
    totalAccounts: number;
    mappedAccounts: number;
    unmappedAccounts: TrialBalanceEntry[];
    conflictingAccounts: Array<{ account: TrialBalanceEntry; noteTypes: string[] }>;
    coveragePercentage: number;
  }> {
    const mappings = await this.loadCompanyMappings(companyId);
    const mappedAccountCodes = new Set<string>();
    const accountToNotes = new Map<string, string[]>();

    // Track which accounts are mapped to which notes
    for (const [noteType, mapping] of mappings) {
      const accounts = this.filterAccountsByMapping(trialBalanceData, mapping);
      
      accounts.forEach(account => {
        const code = account.accountCode || '0';
        mappedAccountCodes.add(code);
        
        if (!accountToNotes.has(code)) {
          accountToNotes.set(code, []);
        }
        accountToNotes.get(code)!.push(noteType);
      });
    }

    // Find unmapped accounts
    const unmappedAccounts = trialBalanceData.filter(account => 
      !mappedAccountCodes.has(account.accountCode || '0')
    );

    // Find conflicting accounts (mapped to multiple notes)
    const conflictingAccounts = Array.from(accountToNotes.entries())
      .filter(([_, noteTypes]) => noteTypes.length > 1)
      .map(([accountCode, noteTypes]) => ({
        account: trialBalanceData.find(a => a.accountCode === accountCode)!,
        noteTypes
      }));

    const totalAccounts = trialBalanceData.length;
    const mappedAccounts = mappedAccountCodes.size;
    const coveragePercentage = totalAccounts > 0 ? (mappedAccounts / totalAccounts) * 100 : 0;

    return {
      totalAccounts,
      mappedAccounts,
      unmappedAccounts,
      conflictingAccounts,
      coveragePercentage
    };
  }

  /**
   * Generate mapping suggestions based on account names and codes
   */
  static generateMappingSuggestions(
    trialBalanceData: TrialBalanceEntry[]
  ): Array<{ accountCode: string; accountName: string; suggestedNoteType: string; confidence: number }> {
    const suggestions: Array<{ accountCode: string; accountName: string; suggestedNoteType: string; confidence: number }> = [];

    for (const entry of trialBalanceData) {
      const code = parseFloat(entry.accountCode || '0');
      const name = (entry.accountName || '').toLowerCase();
      
      let suggestedNoteType = '';
      let confidence = 0;

      // Rule-based suggestions based on account codes and names
      if (code >= 1000 && code <= 1099) {
        suggestedNoteType = 'cash';
        confidence = name.includes('เงินสด') || name.includes('ธนาคาร') ? 0.95 : 0.8;
      } else if (code >= 1140 && code <= 1215) {
        suggestedNoteType = 'receivables';
        confidence = name.includes('ลูกหนี้') ? 0.95 : 0.8;
      } else if (code === 1510) {
        suggestedNoteType = 'inventory';
        confidence = name.includes('สินค้า') ? 0.95 : 0.8;
      } else if (code >= 1610 && code <= 1659) {
        suggestedNoteType = 'ppe_cost';
        confidence = name.includes('ที่ดิน') || name.includes('อาคาร') || name.includes('เครื่องจักร') ? 0.95 : 0.8;
      } else if (code >= 2010 && code <= 2999) {
        suggestedNoteType = 'payables';
        confidence = name.includes('เจ้าหนี้') ? 0.95 : 0.7;
        
        // Check for exclusions
        if ([2030, 2045, 2050, 2051, 2052, 2100, 2101, 2102, 2103, 2120, 2121, 2122, 2123].includes(code)) {
          if (code === 2030) suggestedNoteType = 'short_term_loans';
          else if (code === 2045) suggestedNoteType = 'income_tax_payable';
          else if (code >= 2120 && code <= 2123) suggestedNoteType = 'long_term_loans_fi';
          else suggestedNoteType = 'long_term_loans_other';
          confidence = 0.9;
        }
      }

      if (suggestedNoteType && confidence > 0) {
        suggestions.push({
          accountCode: entry.accountCode || '0',
          accountName: entry.accountName || '',
          suggestedNoteType,
          confidence
        });
      }
    }

    return suggestions.sort((a, b) => b.confidence - a.confidence);
  }

  /**
   * Create default mappings for a new company
   */
  static async createDefaultMappingsForCompany(companyId: number): Promise<void> {
    // The default mappings are already created by the database migration
    // This method can be used to refresh or recreate them if needed
    try {
      await ApiService.resetAccountMappingsToDefault(companyId.toString());
      this.clearCache(companyId); // Clear cache to force reload
    } catch (error) {
      console.error('Failed to create default mappings:', error);
      throw error;
    }
  }
}
