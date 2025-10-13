// Note-first link map: Balance Sheet reads all leaf values from note totals.
import type { DetailedFinancialData } from '../types';

export type NoteTotal = { current: number; previous: number };
export type NoteResults = DetailedFinancialData['noteCalculations'];

export interface BalanceSheetLinkResolver {
  assets: {
    cashAndCashEquivalents: (n: NoteResults) => NoteTotal;
    tradeReceivables: (n: NoteResults) => NoteTotal;
    inventory: (n: NoteResults) => NoteTotal;
    prepaidExpenses: (n: NoteResults) => NoteTotal;
    propertyPlantEquipment: (n: NoteResults) => NoteTotal;
    otherAssets: (n: NoteResults) => NoteTotal;
  };
  liabilities: {
    bankOverdraftsAndShortTermLoans: (n: NoteResults) => NoteTotal;
    tradeAndOtherPayables: (n: NoteResults) => NoteTotal;
    shortTermBorrowings: (n: NoteResults) => NoteTotal;
    incomeTaxPayable: (n: NoteResults) => NoteTotal;
    longTermLoansFromFI: (n: NoteResults) => NoteTotal;
    otherLongTermLoans: (n: NoteResults) => NoteTotal;
    hirePurchaseCreditors?: (n: NoteResults) => NoteTotal;
  };
}

export const BalanceSheetLinkMap: BalanceSheetLinkResolver = {
  assets: {
    cashAndCashEquivalents: (n) => n.cash.total,
    tradeReceivables: (n) => n.receivables.total,
    inventory: (n) => n.inventory.total,
    prepaidExpenses: (n) => n.prepaid, // populated in extractor extension
    propertyPlantEquipment: (n) => n.ppe.netBookValue,
    otherAssets: (n) => n.otherAssets // populated in extractor extension
  },
  liabilities: {
    bankOverdraftsAndShortTermLoans: (n) => n.bankOverdrafts, // populated in extractor extension
    tradeAndOtherPayables: (n) => n.payables.total,
    shortTermBorrowings: (n) => n.shortTermLoans, // populated in extractor extension
    incomeTaxPayable: (n) => n.incomeTaxPayable, // populated in extractor extension
    longTermLoansFromFI: (n) => n.longTermLoansFi, // populated in extractor extension
    otherLongTermLoans: (n) => n.longTermLoansOther, // populated in extractor extension
    hirePurchaseCreditors: (n) => n.hirePurchaseCreditors.total
  }
};

export const NOTE_FIRST_MODE = true; // Feature flag: set true to enable note-first for BS
