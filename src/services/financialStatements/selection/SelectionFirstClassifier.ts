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
  // Optional sub-category (currently only used for cash: 'cash' | 'bankDeposits')
  subCategory?: string;
}

export interface SelectionFirstResult {
  // Flat list by account code
  byAccount: Record<string, ClassifiedAccount>;
  // Grouped by category
  byCategory: Record<NoteCategory | 'unmatched', ClassifiedAccount[]>;
  // Optional grouping for cash sub-categories when rules are present
  subCategories?: {
    cash?: {
      cash: { accounts: ClassifiedAccount[]; current: number; previous: number };
      bankDeposits: { accounts: ClassifiedAccount[]; current: number; previous: number };
      totals: { current: number; previous: number };
      grouped: boolean; // true when DB rules used
    }
  };
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
        const res = resolver(codeStr);
        if (res.matched) { matched = cat; break; }
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

    // ------------------------------------------------------------------
    // SUB-CATEGORY PARTITIONING (currently only cash)
    // If provider supplies subCategoryRules for 'cash', we group cash accounts
    // into two buckets: cash (เงินสด) and bankDeposits (เงินฝากธนาคาร).
    // Rules: We use the sub-category rule sets to test membership based on
    // includes/ranges first (excludes honored). Accounts not matching either
    // rule fall back to legacy heuristic (code range split: 1000-1019 cash, 1020-1099 bank).
    // If NO subCategoryRules provided, we keep individual listing (grouped=false)
    // and do not mutate existing note behavior.
    // ------------------------------------------------------------------
    let subCategories: SelectionFirstResult['subCategories'] | undefined;
    try {
      const cashRules = provider?.getSubCategoryRules('cash');
      const cashAccounts = (byCategory as any)['cash'] as ClassifiedAccount[];
      if (cashAccounts && cashAccounts.length > 0) {
        const ruleSet = cashRules?.cash; // Cash sub-category container { cash?, bankDeposits? }
        if (ruleSet && (ruleSet.cash || ruleSet.bankDeposits)) {
          // Helper for sub-category match
            const buildMatcher = (r: any) => {
              if (!r) return null;
              const includes = new Set<string>((r.includes || []).map((n: number) => String(n)));
              const excludes = new Set<string>((r.excludes || []).map((n: number) => String(n)));
              const ranges: Array<{ from: number; to: number }> = r.ranges || [];
              return (codeStr: string) => {
                if (!r) return false;
                if (excludes.has(codeStr)) return false;
                if (includes.size > 0 && includes.has(codeStr)) return true;
                const codeNum = Number.parseInt(codeStr || '0', 10);
                if (Number.isFinite(codeNum) && ranges.length > 0) {
                  return ranges.some(range => codeNum >= range.from && codeNum <= range.to);
                }
                return false;
              };
            };
          const isCashMatch = buildMatcher(ruleSet.cash);
          const isBankMatch = buildMatcher(ruleSet.bankDeposits);

          const cashBucket: ClassifiedAccount[] = [];
          const bankBucket: ClassifiedAccount[] = [];

          for (const acc of cashAccounts) {
            let assigned: 'cash' | 'bankDeposits' | null = null;
            if (isCashMatch && isCashMatch(acc.accountCode)) assigned = 'cash';
            else if (isBankMatch && isBankMatch(acc.accountCode)) assigned = 'bankDeposits';
            else {
              // Legacy heuristic fallback split by code range if not matched by explicit rules
              const codeNum = Number.parseInt(acc.accountCode || '0', 10);
              if (codeNum >= 1000 && codeNum <= 1019) assigned = 'cash';
              else if (codeNum >= 1020 && codeNum <= 1099) assigned = 'bankDeposits';
            }
            if (!assigned) {
              // If still not assigned, default to cash to retain total consistency
              assigned = 'cash';
            }
            acc.subCategory = assigned;
            if (assigned === 'cash') cashBucket.push(acc); else bankBucket.push(acc);
          }

          const sumBucket = (lst: ClassifiedAccount[]) => lst.reduce((s, a) => {
            s.current += a.current; s.previous += a.previous; return s;
          }, { current: 0, previous: 0 });
          const cashTotals = sumBucket(cashBucket);
          const bankTotals = sumBucket(bankBucket);
          subCategories = {
            cash: {
              cash: { accounts: cashBucket, current: cashTotals.current, previous: cashTotals.previous },
              bankDeposits: { accounts: bankBucket, current: bankTotals.current, previous: bankTotals.previous },
              totals: { current: cashTotals.current + bankTotals.current, previous: cashTotals.previous + bankTotals.previous },
              grouped: true
            }
          };
        } else {
          // No rule set provided -> keep individual listing; still compute raw totals
          const sumBucket = (lst: ClassifiedAccount[]) => lst.reduce((s, a) => {
            s.current += a.current; s.previous += a.previous; return s;
          }, { current: 0, previous: 0 });
          const cashTotals = sumBucket(cashAccounts.filter(a => {
            const codeNum = Number.parseInt(a.accountCode || '0', 10);
            return codeNum >= 1000 && codeNum <= 1019;
          }));
          const bankTotals = sumBucket(cashAccounts.filter(a => {
            const codeNum = Number.parseInt(a.accountCode || '0', 10);
            return codeNum >= 1020 && codeNum <= 1099;
          }));
          subCategories = {
            cash: {
              cash: { accounts: cashAccounts, current: cashTotals.current, previous: cashTotals.previous },
              bankDeposits: { accounts: cashAccounts, current: bankTotals.current, previous: bankTotals.previous },
              totals: { current: totals.cash.current, previous: totals.cash.previous },
              grouped: false
            }
          };
        }
      }
    } catch (err) {
      console.warn('Cash sub-category partition failed, continuing without grouping:', err);
    }

    return {
      byAccount,
      byCategory: byCategory as any,
      subCategories,
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

    return (codeStr: string): { matched: boolean } => {
      const codeNum = Number.parseInt(codeStr || '0', 10);
      const inRanges = (lst?: Array<{ from: number; to: number }>) => Array.isArray(lst) && lst.some(r => codeNum >= r.from && codeNum <= r.to);

      // Provider-based rules (top-level)
      const byInclude = includes.size > 0 && includes.has(codeStr);
      const byRange = ranges.length > 0 && Number.isFinite(codeNum) && inRanges(ranges);
      const excluded = excludes.has(codeStr);
      if ((byInclude || byRange) && !excluded) {
        return { matched: true };
      }

      // Fallback numeric mapping
      return { matched: Number.isFinite(codeNum) && fallback(codeNum) === true };
    };
  }
}
