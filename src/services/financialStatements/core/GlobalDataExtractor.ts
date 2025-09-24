// ============================================================================
// GLOBAL DATA EXTRACTOR
// Centralizes foundation-first extraction & individual account discovery.
// ============================================================================
import { FinancialCalculations } from '../../financialCalculations';
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from './types';
import type { IAccountMappingProvider } from '../mapping/IAccountMappingProvider';
import { sumByRules } from '../mapping/ruleSum';

export class GlobalDataExtractor {
  private static cacheKey(tb: TrialBalanceEntry[], company: CompanyInfo): string {
    return `${company.name}_${tb.length}_${company.reportingYear}`;
  }
  private static cache = new Map<string, DetailedFinancialData>();

  static extract(trialBalanceData: TrialBalanceEntry[], companyInfo: CompanyInfo, provider?: IAccountMappingProvider): DetailedFinancialData {
    const key = this.cacheKey(trialBalanceData, companyInfo);
    const useCache = !provider; // if provider is present (dynamic mappings), bypass cache to reflect changes immediately
    if (useCache && this.cache.has(key)) return this.cache.get(key)!;

    // Build note sections (foundation layer)
    const cashNote = this.buildCashNote(trialBalanceData, provider);
    const receivablesNote = this.buildReceivablesNote(trialBalanceData, provider);
    const inventoryNote = this.buildInventoryNote(trialBalanceData, provider);
    const ppeNote = this.buildPPENote(trialBalanceData, provider);
    const prepaidNote = this.buildPrepaidNote(trialBalanceData, provider);
    const otherAssetsNote = this.buildOtherAssetsNote(trialBalanceData, provider);
    const bankOverdraftsNote = this.buildBankOverdraftsNote(trialBalanceData, provider);
    const shortTermLoansNote = this.buildShortTermLoansNote(trialBalanceData, provider);
    const incomeTaxPayableNote = this.buildIncomeTaxPayableNote(trialBalanceData, provider);
    const longTermLoansFiNote = this.buildLongTermLoansFiNote(trialBalanceData, provider);
    const longTermLoansOtherNote = this.buildLongTermLoansOtherNote(trialBalanceData, provider);
    const payablesNote = this.buildPayablesNote(trialBalanceData, provider);

    // Balance sheet totals built from notes + specific items
    const balanceSheetAssets = this.buildBalanceSheetAssets(trialBalanceData, cashNote, receivablesNote, inventoryNote, ppeNote);
    const balanceSheetLiabilities = this.buildBalanceSheetLiabilities(trialBalanceData, provider, payablesNote);

    // Equity (retained earnings computed per VBA rules)
    const balanceSheetEquity = this.buildEquitySection(trialBalanceData, provider);

    // Income statement base
    const { revenue, expenses, netProfit } = this.buildIncomeStatementBase(trialBalanceData);

    // Individual accounts (detail layer)
    const individualAccounts = this.extractIndividualAccounts(trialBalanceData, provider);

    // Flags
    const flags = this.buildFlags(balanceSheetAssets, companyInfo);

    const extracted: DetailedFinancialData = {
      noteCalculations: {
        cash: cashNote,
        receivables: receivablesNote,
        inventory: inventoryNote,
        ppe: ppeNote,
        payables: payablesNote,
        prepaid: prepaidNote,
        otherAssets: otherAssetsNote,
        bankOverdrafts: bankOverdraftsNote,
        shortTermLoans: shortTermLoansNote,
        incomeTaxPayable: incomeTaxPayableNote,
        longTermLoansFi: longTermLoansFiNote,
        longTermLoansOther: longTermLoansOtherNote
      },
      individualAccounts,
      balanceSheetTotals: { assets: balanceSheetAssets, liabilities: balanceSheetLiabilities, equity: balanceSheetEquity },
      income: { revenue, expenses, netProfit },
      flags
    };

    if (useCache) {
      this.cache.set(key, extracted);
    }
    return extracted;
  }

  // ---- Helpers: Notes (foundation layer) ---------------------------------
  private static buildCashNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      cash: {
        current: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, { ranges: [{ from: 1000, to: 1019 }] }, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1019)),
        previous: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, { ranges: [{ from: 1000, to: 1019 }] }, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1019))
      },
      bankDeposits: {
        current: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, { ranges: [{ from: 1020, to: 1099 }] }, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1020, 1099)),
        previous: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, { ranges: [{ from: 1020, to: 1099 }] }, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1020, 1099))
      },
      total: {
        current: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, provider.getRules('cash')!, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099)),
        previous: provider?.getRules('cash')
          ? sumByRules(trialBalanceData, provider.getRules('cash')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099))
      }
    };
  }

  private static buildReceivablesNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      total: {
        current: provider?.getRules('receivables')
          ? sumByRules(trialBalanceData, provider.getRules('receivables')!, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215)),
        previous: provider?.getRules('receivables')
          ? sumByRules(trialBalanceData, provider.getRules('receivables')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215))
      }
    };
  }

  private static buildInventoryNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      inventory: {
        current: provider?.getRules('inventory')
          ? sumByRules(trialBalanceData, { includes: [1510] }, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1510, 1510)),
        previous: provider?.getRules('inventory')
          ? sumByRules(trialBalanceData, { includes: [1510] }, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1510, 1510))
      },
      total: {
        current: provider?.getRules('inventory')
          ? sumByRules(trialBalanceData, provider.getRules('inventory')!, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519)),
        previous: provider?.getRules('inventory')
          ? sumByRules(trialBalanceData, provider.getRules('inventory')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519))
      }
    };
  }

  private static buildPPENote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    const costCurrent = provider?.getRules('ppe_cost')
      ? sumByRules(trialBalanceData, provider.getRules('ppe_cost')!, 'current')
      : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1629));
    const costPrevious = provider?.getRules('ppe_cost')
      ? sumByRules(trialBalanceData, provider.getRules('ppe_cost')!, 'previous')
      : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1629));

    const accCurrent = provider?.getRules('ppe_accum_depr')
      ? sumByRules(trialBalanceData, provider.getRules('ppe_accum_depr')!, 'current')
      : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1630, 1659));
    const accPrevious = provider?.getRules('ppe_accum_depr')
      ? sumByRules(trialBalanceData, provider.getRules('ppe_accum_depr')!, 'previous')
      : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1630, 1659));

    return {
      cost: { current: costCurrent, previous: costPrevious },
      accumulatedDepreciation: { current: accCurrent, previous: accPrevious },
      netBookValue: { current: costCurrent - accCurrent, previous: costPrevious - accPrevious }
    };
  }

  private static buildPrepaidNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('prepaid')
        ? sumByRules(trialBalanceData, provider.getRules('prepaid')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1400, 1439)),
      previous: provider?.getRules('prepaid')
        ? sumByRules(trialBalanceData, provider.getRules('prepaid')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439))
    };
  }

  private static buildOtherAssetsNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('other_assets')
        ? sumByRules(trialBalanceData, provider.getRules('other_assets')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700)),
      previous: provider?.getRules('other_assets')
        ? sumByRules(trialBalanceData, provider.getRules('other_assets')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700))
    };
  }

  private static buildBankOverdraftsNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('bank_overdrafts')
        ? sumByRules(trialBalanceData, provider.getRules('bank_overdrafts')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2001, 2009)),
      previous: provider?.getRules('bank_overdrafts')
        ? sumByRules(trialBalanceData, provider.getRules('bank_overdrafts')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2001, 2009))
    };
  }

  private static buildShortTermLoansNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('short_term_loans')
        ? sumByRules(trialBalanceData, provider.getRules('short_term_loans')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)),
      previous: provider?.getRules('short_term_loans')
        ? sumByRules(trialBalanceData, provider.getRules('short_term_loans')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030))
    };
  }

  private static buildIncomeTaxPayableNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('income_tax_payable')
        ? sumByRules(trialBalanceData, provider.getRules('income_tax_payable')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)),
      previous: provider?.getRules('income_tax_payable')
        ? sumByRules(trialBalanceData, provider.getRules('income_tax_payable')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045))
    };
  }

  private static buildLongTermLoansFiNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('long_term_loans_fi')
        ? sumByRules(trialBalanceData, provider.getRules('long_term_loans_fi')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) -
          Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2121, 2121)),
      previous: provider?.getRules('long_term_loans_fi')
        ? sumByRules(trialBalanceData, provider.getRules('long_term_loans_fi')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2120, 2123)) -
          Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2121, 2121))
    };
  }

  private static buildLongTermLoansOtherNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      current: provider?.getRules('long_term_loans_other')
        ? sumByRules(trialBalanceData, provider.getRules('long_term_loans_other')!, 'current')
        : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) +
          Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2119)),
      previous: provider?.getRules('long_term_loans_other')
        ? sumByRules(trialBalanceData, provider.getRules('long_term_loans_other')!, 'previous')
        : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) +
          Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2119))
    };
  }

  private static buildPayablesNote(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    return {
      total: {
        current: provider?.getRules('payables')
          ? sumByRules(trialBalanceData, provider.getRules('payables')!, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2010, 2999)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) -
            Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2123)),
        previous: provider?.getRules('payables')
          ? sumByRules(trialBalanceData, provider.getRules('payables')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2010, 2999)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) -
            Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2123))
      }
    };
  }

  // ---- Helpers: Balance Sheet totals --------------------------------------
  private static buildBalanceSheetAssets(
    trialBalanceData: TrialBalanceEntry[],
    cashNote: any,
    receivablesNote: any,
    inventoryNote: any,
    ppeNote: any
  ) {
    return {
      cashAndCashEquivalents: cashNote.total,
      tradeReceivables: receivablesNote.total,
      inventory: inventoryNote.total,
      prepaidExpenses: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1400, 1439)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1400, 1439))
      },
      propertyPlantEquipment: ppeNote.netBookValue,
      otherAssets: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1660, 1700)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1660, 1700))
      }
    };
  }

  private static buildBalanceSheetLiabilities(
    trialBalanceData: TrialBalanceEntry[],
    provider: IAccountMappingProvider | undefined,
    payablesNote: any
  ) {
    return {
      bankOverdraftsAndShortTermLoans: this.buildBankOverdraftsNote(trialBalanceData, provider),
      tradeAndOtherPayables: payablesNote.total,
      shortTermBorrowings: this.buildShortTermLoansNote(trialBalanceData, provider),
      incomeTaxPayable: this.buildIncomeTaxPayableNote(trialBalanceData, provider),
      longTermLoansFromFI: this.buildLongTermLoansFiNote(trialBalanceData, provider),
      otherLongTermLoans: this.buildLongTermLoansOtherNote(trialBalanceData, provider)
    };
  }

  // ---- Helpers: Equity & Income -------------------------------------------
  private static buildEquitySection(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
    const currentYearProfit = FinancialCalculations.calculateCurrentYearProfit(trialBalanceData);
    const openingRetainedEarnings = FinancialCalculations.getOpeningRetainedEarnings(trialBalanceData);
    const finalRetainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);

    return {
      paidUpCapital: {
        current: provider?.getRules('paid_up_capital')
          ? sumByRules(trialBalanceData, provider.getRules('paid_up_capital')!, 'current')
          : this.getSingleAccountBalance(trialBalanceData, '3010'),
        previous: provider?.getRules('paid_up_capital')
          ? sumByRules(trialBalanceData, provider.getRules('paid_up_capital')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3010, 3010))
      },
      retainedEarnings: {
        current: finalRetainedEarnings,
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3020, 3020))
      },
      openingRetainedEarnings: openingRetainedEarnings,
      legalReserve: {
        current: provider?.getRules('legal_reserve')
          ? sumByRules(trialBalanceData, provider.getRules('legal_reserve')!, 'current')
          : Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3030, 3039)),
        previous: provider?.getRules('legal_reserve')
          ? sumByRules(trialBalanceData, provider.getRules('legal_reserve')!, 'previous')
          : Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3030, 3039))
      }
    };
  }

  private static buildIncomeStatementBase(trialBalanceData: TrialBalanceEntry[]) {
    const revenue = {
      total: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4999)),
      mainRevenue: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4000, 4099)),
      otherIncome: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 4100, 4999))
    };
    const expenses = {
      total: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5999)),
      costOfServices: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5000, 5099)),
      adminExpenses: Math.abs(
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5300, 5350) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5355, 5357) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5362, 5363) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5365, 5365)
      ),
      otherExpenses: Math.abs(
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5351, 5354) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5358, 5361) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5364, 5364) +
        FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5366, 5999)
      ),
      incomeTax: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5910, 5910)),
      financialCosts: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 5920, 5929))
    };
    const netProfit = revenue.total - expenses.total;
    return { revenue, expenses, netProfit };
  }

  // ---- Helpers: Flags ------------------------------------------------------
  private static buildFlags(assets: any, companyInfo: CompanyInfo) {
    return {
      hasInventory: assets.inventory.current > 0,
      isServiceBusiness: assets.inventory.current === 0,
      isLimitedPartnership: companyInfo.type === 'ห้างหุ้นส่วนจำกัด'
    };
  }

  private static getSingleAccountBalance(trialBalanceData: TrialBalanceEntry[], accountCode: string): number {
    const account = trialBalanceData.find(e => e.accountCode === accountCode);
    return account ? Math.abs(account.currentBalance || account.balance || 0) : 0;
  }

  private static extractIndividualAccounts(trialBalanceData: TrialBalanceEntry[], provider?: IAccountMappingProvider) {
  const individualAccounts: any = { cash: {}, receivables: {}, payables: {} };

    // Helper to get entries by provider rules or fallback numeric filters
    const getEntriesByRulesOrRange = (
      key: string,
      fallback: (e: TrialBalanceEntry) => boolean
    ): TrialBalanceEntry[] => {
      const rules = provider?.getRules(key);
      if (rules) {
        return this.selectByRules(trialBalanceData, rules);
      }
      return trialBalanceData.filter(fallback);
    };

    // Cash: 1000-1099 with category split
    const cashEntries = getEntriesByRulesOrRange('cash', e => {
      const code = parseInt(e.accountCode || '0', 10);
      return code >= 1000 && code <= 1099;
    });
    for (const entry of cashEntries) {
      const code = parseInt(entry.accountCode || '0', 10);
      const currentAmount = Math.abs((entry.currentBalance ?? entry.balance ?? 0) as number);
      const previousAmount = Math.abs((entry.previousBalance ?? 0) as number);
      if (currentAmount === 0 && previousAmount === 0) continue;
      const category = code <= 1019 ? 'cash' : 'bankDeposits';
      individualAccounts.cash[entry.accountCode || ''] = {
        accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
        current: currentAmount,
        previous: previousAmount,
        noteCategory: 'cash',
        category
      };
    }

    // Receivables: provider rules or 1140-1215
    const receivableEntries = getEntriesByRulesOrRange('receivables', e => {
      const code = parseInt(e.accountCode || '0', 10);
      return code >= 1140 && code <= 1215;
    });
    for (const entry of receivableEntries) {
      const currentAmount = Math.abs((entry.currentBalance ?? entry.balance ?? 0) as number);
      const previousAmount = Math.abs((entry.previousBalance ?? 0) as number);
      if (currentAmount === 0 && previousAmount === 0) continue;
      individualAccounts.receivables[entry.accountCode || ''] = {
        accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
        current: currentAmount,
        previous: previousAmount,
        noteCategory: 'receivables'
      };
    }

    // Payables: provider rules or 2010-2999 minus exclusions
    const payablesFallback = (e: TrialBalanceEntry) => {
      const code = parseInt(e.accountCode || '0', 10);
      const isExcluded = code === 2030 || code === 2045 || (code >= 2050 && code <= 2052) || (code >= 2100 && code <= 2123);
      return code >= 2010 && code <= 2999 && !isExcluded;
    };
    const payableEntries = getEntriesByRulesOrRange('payables', payablesFallback);
    for (const entry of payableEntries) {
      const currentAmount = Math.abs((entry.currentBalance ?? entry.balance ?? 0) as number);
      const previousAmount = Math.abs((entry.previousBalance ?? 0) as number);
      if (currentAmount === 0 && previousAmount === 0) continue;
      individualAccounts.payables[entry.accountCode || ''] = {
        accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
        current: currentAmount,
        previous: previousAmount,
        noteCategory: 'payables'
      };
    }

    return individualAccounts;
  }

  // Minimal rule-aware selector (kept private, in-file)
  private static selectByRules(tb: TrialBalanceEntry[], rules: any): TrialBalanceEntry[] {
    const toNum = (v: any) => Number.parseInt(String(v), 10);
    const inRanges = (codeNum: number, ranges?: Array<{ from: number; to: number }>) =>
      Array.isArray(ranges) && ranges.some(r => codeNum >= r.from && codeNum <= r.to);

    const includeSet = new Set<string>((rules?.includes ?? []).map((v: any) => String(v)));
    const excludeSet = new Set<string>((rules?.excludes ?? []).map((v: any) => String(v)));
    const hasRanges = Array.isArray(rules?.ranges) && rules.ranges.length > 0;
    const hasInclude = includeSet.size > 0;

    const out: TrialBalanceEntry[] = [];
    const seen = new Set<string>();
    for (const e of tb) {
      const codeStr = e.accountCode || '';
      const codeNum = toNum(codeStr);
      const byRange = hasRanges && Number.isFinite(codeNum) && inRanges(codeNum, rules.ranges);
      const byInclude = hasInclude && includeSet.has(codeStr);
      if (!(byRange || byInclude)) continue;
      if (excludeSet.has(codeStr)) continue;
      if (Number.isFinite(codeNum) && inRanges(codeNum, rules?.excludeRanges)) continue;
      if (!seen.has(codeStr)) {
        seen.add(codeStr);
        out.push(e);
      }
    }
    return out;
  }
}
