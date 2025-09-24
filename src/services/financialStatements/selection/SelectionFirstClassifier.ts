// ============================================================================
// SELECTION-FIRST CLASSIFIER
// Maps every Trial Balance entry to a NoteCategory up-front, as a single source
// of truth for note building and balance sheet linking.
// This keeps data selection separate from presentation (notes/BS builders).
// ============================================================================

import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { NoteCategory } from '../core/types';
import type { IAccountMappingProvider } from '../mapping/IAccountMappingProvider';

export interface ClassifiedAccount {
  accountCode: string;
  accountName: string;
  current: number;
  previous: number;
  category: NoteCategory | 'unmatched';
}

export interface SelectionFirstResult {
  // Flat list by account code
  byAccount: Record<string, ClassifiedAccount>;
  // Grouped by category
  byCategory: Record<NoteCategory | 'unmatched', ClassifiedAccount[]>;
  // Category totals (sums of current/previous)
  totals: Record<NoteCategory, { current: number; previous: number }>;
  // Accounts that didn’t match any mapping (for UI diagnostics)
  unmatched: ClassifiedAccount[];
  // Meta
  company: Pick<CompanyInfo, 'name' | 'reportingYear' | 'type'>;
}

// Priority order prevents double counting when ranges overlap
const CATEGORY_PRIORITY: NoteCategory[] = [
  'cash',
  'receivables',
  'inventory',
  'prepaid',
  'ppe_cost',
  'ppe_accum_depr',
  'other_assets',
  'bank_overdrafts',
  'short_term_loans',
  'income_tax_payable',
  'long_term_loans_fi',
  'long_term_loans_other',
  'payables'
];

export class SelectionFirstClassifier {
  static classify(
    trialBalanceData: TrialBalanceEntry[],
    company: CompanyInfo,
    provider?: IAccountMappingProvider
  ): SelectionFirstResult {
    const byAccount: Record<string, ClassifiedAccount> = {};
    const byCategory: Record<NoteCategory | 'unmatched', ClassifiedAccount[]> = Object.create(null);
    const totals: Record<NoteCategory, { current: number; previous: number }> = Object.create(null);

    // Initialize buckets
    [...CATEGORY_PRIORITY, 'unmatched' as const].forEach((cat: any) => {
      (byCategory as any)[cat] = [];
      if (cat !== 'unmatched') {
        (totals as any)[cat] = { current: 0, previous: 0 };
      }
    });

    // Build rule resolvers for speed
    const ruleCache = new Map<NoteCategory, ReturnType<typeof SelectionFirstClassifier.buildResolver>>();
    const getResolver = (cat: NoteCategory) => {
      if (ruleCache.has(cat)) return ruleCache.get(cat)!;
      const resolver = SelectionFirstClassifier.buildResolver(cat, provider);
      ruleCache.set(cat, resolver);
      return resolver;
    };

    // Assign each account to the first matching category by priority
    for (const e of trialBalanceData) {
      const codeStr = e.accountCode || '';
      const accountName = e.accountName || `บัญชี ${codeStr}`;
      const current = Math.abs((e.currentBalance ?? e.balance ?? 0) as number);
      const previous = Math.abs((e.previousBalance ?? 0) as number);

      let matched: NoteCategory | null = null;
      for (const cat of CATEGORY_PRIORITY) {
        const resolver = getResolver(cat);
        if (resolver(codeStr)) { matched = cat; break; }
      }

      const finalCat = matched ?? 'unmatched';
      const rec: ClassifiedAccount = {
        accountCode: codeStr,
        accountName,
        current,
        previous,
        category: finalCat as any
      };

      byAccount[codeStr] = rec;
      (byCategory as any)[finalCat].push(rec);
      if (finalCat !== 'unmatched') {
        (totals as any)[finalCat].current += current;
        (totals as any)[finalCat].previous += previous;
      }
    }

    return {
      byAccount,
      byCategory: byCategory as any,
      totals: totals as any,
      unmatched: (byCategory as any)['unmatched'],
      company: { name: company.name, reportingYear: company.reportingYear, type: company.type }
    };
  }

  // Build a predicate using provider rules; fall back to numeric ranges
  private static buildResolver(
    cat: NoteCategory,
    provider?: IAccountMappingProvider
  ) {
    const rules = provider?.getRules(cat);
    const includes = new Set<string>((rules?.includes ?? []).map(String));
    const excludes = new Set<string>((rules?.excludes ?? []).map(String));
    const ranges: Array<{ from: number; to: number }> = rules?.ranges ?? [];

    // Fallback numeric ranges per legacy logic
    const fallback = (codeNum: number) => {
      switch (cat) {
        case 'cash': return codeNum >= 1000 && codeNum <= 1099;
        case 'receivables': return codeNum >= 1140 && codeNum <= 1215;
        case 'inventory': return codeNum >= 1500 && codeNum <= 1519;
        case 'prepaid': return codeNum >= 1400 && codeNum <= 1439;
        case 'ppe_cost': return codeNum >= 1600 && codeNum <= 1629;
        case 'ppe_accum_depr': return codeNum >= 1630 && codeNum <= 1659;
        case 'other_assets': return codeNum >= 1660 && codeNum <= 1700;
        case 'bank_overdrafts': return codeNum >= 2001 && codeNum <= 2009;
        case 'short_term_loans': return codeNum === 2030;
        case 'income_tax_payable': return codeNum === 2045;
        case 'long_term_loans_fi': return codeNum >= 2120 && codeNum <= 2123 && codeNum !== 2121;
        case 'long_term_loans_other': return (codeNum >= 2050 && codeNum <= 2052) || (codeNum >= 2100 && codeNum <= 2119);
        case 'payables': {
          const excluded = codeNum === 2030 || codeNum === 2045 || (codeNum >= 2050 && codeNum <= 2052) || (codeNum >= 2100 && codeNum <= 2123);
          return codeNum >= 2010 && codeNum <= 2999 && !excluded;
        }
      }
    };

    return (codeStr: string): boolean => {
      const codeNum = Number.parseInt(codeStr || '0', 10);
      const inRanges = (lst?: Array<{ from: number; to: number }>) => Array.isArray(lst) && lst.some(r => codeNum >= r.from && codeNum <= r.to);

      // Provider-based rules
      const byInclude = includes.size > 0 && includes.has(codeStr);
      const byRange = ranges.length > 0 && Number.isFinite(codeNum) && inRanges(ranges);
  const excluded = excludes.has(codeStr);
      if ((byInclude || byRange) && !excluded) return true;

      // Fallback numeric mapping
      return Number.isFinite(codeNum) && fallback(codeNum) === true;
    };
  }
}
