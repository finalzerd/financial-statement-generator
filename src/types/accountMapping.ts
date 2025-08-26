// ============================================================================
// COMPANY ACCOUNT MAPPING INTERFACES
// ============================================================================

/**
 * Account range specification for flexible mapping
 */
export interface AccountRange {
  from: number;
  to: number;
}

/**
 * Complete account mapping rules for a note type
 */
export interface AccountMappingRules {
  ranges?: AccountRange[];           // Account ranges (e.g., 1000-1099)
  includes?: number[];               // Specific account codes to include
  excludes?: number[];               // Specific account codes to exclude from ranges
}

/**
 * Company-specific account mapping for a note type
 */
export interface CompanyAccountMapping {
  id: number;
  companyId: number;
  noteType: string;                  // 'cash', 'receivables', 'payables', etc.
  noteNumber?: number;               // Note number in financial statements
  noteTitle?: string;                // Custom note title
  accountRanges: AccountMappingRules; // Flexible mapping rules
  isActive: boolean;                 // Whether this mapping is active
  createdAt: string;
  updatedAt: string;
}

/**
 * Request interface for creating/updating account mappings
 */
export interface AccountMappingRequest {
  noteType: string;
  noteNumber?: number;
  noteTitle?: string;
  accountRanges: AccountMappingRules;
  isActive?: boolean;
}

/**
 * Response interface for account mapping validation
 */
export interface AccountMappingValidation {
  isValid: boolean;
  warnings: string[];               // Non-critical issues
  errors: string[];                 // Critical issues that prevent processing
  unmappedAccounts: Array<{         // Accounts not covered by any mapping
    accountCode: string;
    accountName: string;
    balance: number;
  }>;
  conflictingAccounts: Array<{      // Accounts mapped to multiple notes
    accountCode: string;
    accountName: string;
    noteTypes: string[];
  }>;
  mappingCoverage: {                // Summary of mapping coverage
    totalAccounts: number;
    mappedAccounts: number;
    unmappedAccounts: number;
    coveragePercentage: number;
  };
}

/**
 * Predefined note types with their standard configurations
 */
export const STANDARD_NOTE_TYPES = {
  // Balance Sheet Assets
  cash: {
    noteNumber: 7,
    noteTitle: 'เงินสดและรายการเทียบเท่าเงินสด',
    balanceSheetSection: 'current_assets'
  },
  receivables: {
    noteNumber: 8,
    noteTitle: 'ลูกหนี้การค้าและลูกหนี้อื่น',
    balanceSheetSection: 'current_assets'
  },
  inventory: {
    noteNumber: 9,
    noteTitle: 'สินค้าคงเหลือ',
    balanceSheetSection: 'current_assets'
  },
  prepaid_expenses: {
    noteNumber: 10,
    noteTitle: 'ค่าใช้จ่ายจ่ายล่วงหน้า',
    balanceSheetSection: 'current_assets'
  },
  ppe_cost: {
    noteNumber: 11,
    noteTitle: 'ที่ดิน อาคาร และอุปกรณ์',
    balanceSheetSection: 'non_current_assets'
  },
  other_assets: {
    noteNumber: 12,
    noteTitle: 'สินทรัพย์อื่น',
    balanceSheetSection: 'non_current_assets'
  },
  
  // Balance Sheet Liabilities
  bank_overdrafts: {
    noteNumber: 15,
    noteTitle: 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน',
    balanceSheetSection: 'current_liabilities'
  },
  payables: {
    noteNumber: 16,
    noteTitle: 'เจ้าหนี้การค้าและเจ้าหนี้อื่น',
    balanceSheetSection: 'current_liabilities'
  },
  short_term_loans: {
    noteNumber: 17,
    noteTitle: 'เงินกู้ยืมระยะสั้น',
    balanceSheetSection: 'current_liabilities'
  },
  income_tax_payable: {
    noteNumber: 18,
    noteTitle: 'ภาษีเงินได้นิติบุคคลค้างจ่าย',
    balanceSheetSection: 'current_liabilities'
  },
  long_term_loans_fi: {
    noteNumber: 19,
    noteTitle: 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน',
    balanceSheetSection: 'non_current_liabilities'
  },
  long_term_loans_other: {
    noteNumber: 20,
    noteTitle: 'เงินกู้ยืมระยะยาวอื่น',
    balanceSheetSection: 'non_current_liabilities'
  }
} as const;

/**
 * Helper type for note type keys
 */
export type NoteTypeKey = keyof typeof STANDARD_NOTE_TYPES;

/**
 * Account mapping utility functions
 */
export class AccountMappingUtils {
  /**
   * Check if an account code matches the mapping rules
   */
  static doesAccountMatch(accountCode: string, rules: AccountMappingRules): boolean {
    const code = parseFloat(accountCode);
    if (isNaN(code)) return false;

    // Check if explicitly excluded
    if (rules.excludes?.includes(code)) {
      return false;
    }

    // Check if explicitly included
    if (rules.includes?.includes(code)) {
      return true;
    }

    // Check if within any range
    return rules.ranges?.some(range => code >= range.from && code <= range.to) || false;
  }

  /**
   * Get all account codes that match the mapping rules
   */
  static getMatchingAccounts(
    trialBalanceData: any[], 
    rules: AccountMappingRules
  ): any[] {
    return trialBalanceData.filter(entry => 
      this.doesAccountMatch(entry.accountCode || '0', rules)
    );
  }

  /**
   * Validate mapping rules for logical consistency
   */
  static validateMappingRules(rules: AccountMappingRules): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    // Check for overlapping ranges
    if (rules.ranges && rules.ranges.length > 1) {
      for (let i = 0; i < rules.ranges.length; i++) {
        for (let j = i + 1; j < rules.ranges.length; j++) {
          const range1 = rules.ranges[i];
          const range2 = rules.ranges[j];
          
          if ((range1.from <= range2.to && range1.to >= range2.from)) {
            errors.push(`Overlapping ranges: ${range1.from}-${range1.to} and ${range2.from}-${range2.to}`);
          }
        }
      }
    }

    // Check for invalid ranges
    rules.ranges?.forEach((range, index) => {
      if (range.from > range.to) {
        errors.push(`Invalid range ${index + 1}: 'from' (${range.from}) cannot be greater than 'to' (${range.to})`);
      }
      if (range.from < 0 || range.to < 0) {
        errors.push(`Invalid range ${index + 1}: Account codes cannot be negative`);
      }
    });

    // Check for excluded accounts that are also included
    if (rules.includes && rules.excludes) {
      const conflicting = rules.includes.filter(code => rules.excludes!.includes(code));
      if (conflicting.length > 0) {
        errors.push(`Accounts cannot be both included and excluded: ${conflicting.join(', ')}`);
      }
    }

    return { isValid: errors.length === 0, errors };
  }

  /**
   * Generate human-readable description of mapping rules
   */
  static describeMappingRules(rules: AccountMappingRules): string {
    const parts: string[] = [];

    if (rules.ranges?.length) {
      const rangeDescriptions = rules.ranges.map(r => `${r.from}-${r.to}`);
      parts.push(`Ranges: ${rangeDescriptions.join(', ')}`);
    }

    if (rules.includes?.length) {
      parts.push(`Includes: ${rules.includes.join(', ')}`);
    }

    if (rules.excludes?.length) {
      parts.push(`Excludes: ${rules.excludes.join(', ')}`);
    }

    return parts.join(' | ') || 'No rules defined';
  }
}
