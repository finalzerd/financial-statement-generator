// ============================================================================
// GLOBAL DATA EXTRACTOR
// Centralizes foundation-first extraction & individual account discovery.
// ============================================================================
import { FinancialCalculations } from '../../financialCalculations';
import type { TrialBalanceEntry, CompanyInfo } from '../../../types/financial';
import type { DetailedFinancialData } from './types';

export class GlobalDataExtractor {
  private static cacheKey(tb: TrialBalanceEntry[], company: CompanyInfo): string {
    return `${company.name}_${tb.length}_${company.reportingYear}`;
  }
  private static cache = new Map<string, DetailedFinancialData>();

  static extract(trialBalanceData: TrialBalanceEntry[], companyInfo: CompanyInfo): DetailedFinancialData {
    const key = this.cacheKey(trialBalanceData, companyInfo);
    if (this.cache.has(key)) return this.cache.get(key)!;

    // NOTE 7
    const cashNote = {
      cash: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1019)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1019))
      },
      bankDeposits: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1020, 1099)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1020, 1099))
      },
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1000, 1099)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1000, 1099))
      }
    };

    // NOTE 8
    const receivablesNote = {
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1140, 1215)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1140, 1215))
      }
    };

    // NOTE 9
    const inventoryNote = {
      inventory: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1510, 1510)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1510, 1510))
      },
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1500, 1519)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1500, 1519))
      }
    };

    // NOTE 10
    const ppeNote = {
      cost: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1629)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1629))
      },
      accumulatedDepreciation: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1630, 1659)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1630, 1659))
      },
      netBookValue: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1600, 1629)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 1630, 1659)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1600, 1629)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 1630, 1659))
      }
    };

    // NOTE 12
    const payablesNote = {
      total: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2010, 2999)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2123)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2010, 2999)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2123))
      }
    };

    // Balance sheet totals
    const balanceSheetAssets = {
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

    const balanceSheetLiabilities = {
      bankOverdraftsAndShortTermLoans: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2001, 2009)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2001, 2009))
      },
      tradeAndOtherPayables: payablesNote.total,
      shortTermBorrowings: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2030, 2030)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2030, 2030))
      },
      incomeTaxPayable: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2045, 2045)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2045, 2045))
      },
      longTermLoansFromFI: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2120, 2123)) -
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2121, 2121)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2120, 2123)) -
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2121, 2121))
      },
      otherLongTermLoans: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2050, 2052)) +
                 Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 2100, 2119)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2050, 2052)) +
                  Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 2100, 2119))
      }
    };

    // Equity
    const currentYearProfit = FinancialCalculations.calculateCurrentYearProfit(trialBalanceData);
    const openingRetainedEarnings = FinancialCalculations.getOpeningRetainedEarnings(trialBalanceData);
    const finalRetainedEarnings = Math.abs(openingRetainedEarnings + currentYearProfit);

    const balanceSheetEquity = {
      paidUpCapital: {
        current: this.getSingleAccountBalance(trialBalanceData, '3010'),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3010, 3010))
      },
      retainedEarnings: {
        current: finalRetainedEarnings,
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3020, 3020))
      },
      openingRetainedEarnings: openingRetainedEarnings,
      legalReserve: {
        current: Math.abs(FinancialCalculations.sumAccountsByNumericRange(trialBalanceData, 3030, 3039)),
        previous: Math.abs(FinancialCalculations.sumPreviousBalanceByNumericRange(trialBalanceData, 3030, 3039))
      }
    };

    // Income statement base
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

    const individualAccounts = this.extractIndividualAccounts(trialBalanceData);

    const flags = {
      hasInventory: balanceSheetAssets.inventory.current > 0,
      isServiceBusiness: balanceSheetAssets.inventory.current === 0,
      isLimitedPartnership: companyInfo.type === 'ห้างหุ้นส่วนจำกัด'
    };

    const extracted: DetailedFinancialData = {
      noteCalculations: { cash: cashNote, receivables: receivablesNote, inventory: inventoryNote, ppe: ppeNote, payables: payablesNote },
      individualAccounts,
      balanceSheetTotals: { assets: balanceSheetAssets, liabilities: balanceSheetLiabilities, equity: balanceSheetEquity },
      income: { revenue, expenses, netProfit },
      flags
    };

    this.cache.set(key, extracted);
    return extracted;
  }

  private static getSingleAccountBalance(trialBalanceData: TrialBalanceEntry[], accountCode: string): number {
    const account = trialBalanceData.find(e => e.accountCode === accountCode);
    return account ? Math.abs(account.currentBalance || account.balance || 0) : 0;
  }

  private static extractIndividualAccounts(trialBalanceData: TrialBalanceEntry[]) {
    const individualAccounts: any = { cash: {}, receivables: {}, payables: {} };
    for (const entry of trialBalanceData) {
      const code = parseInt(entry.accountCode || '0');
      const currentAmount = Math.abs(entry.balance || 0);
      const previousAmount = Math.abs(entry.previousBalance || 0);
      if (currentAmount === 0 && previousAmount === 0) continue;
      if (code >= 1000 && code <= 1099) {
        individualAccounts.cash[entry.accountCode || ''] = { accountName: entry.accountName || `บัญชี ${entry.accountCode}`, current: currentAmount, previous: previousAmount, category: code <= 1019 ? 'cash' : 'bankDeposits' };
      } else if (code >= 1140 && code <= 1215) {
        individualAccounts.receivables[entry.accountCode || ''] = { accountName: entry.accountName || `บัญชี ${entry.accountCode}`, current: currentAmount, previous: previousAmount };
      } else if (code >= 2010 && code <= 2999) {
        const isExcluded = code === 2030 || code === 2045 || (code >= 2050 && code <= 2052) || (code >= 2100 && code <= 2123);
        if (!isExcluded) {
          individualAccounts.payables[entry.accountCode || ''] = { accountName: entry.accountName || `บัญชี ${entry.accountCode}`, current: currentAmount, previous: previousAmount };
        }
      }
    }
    return individualAccounts;
  }
}
