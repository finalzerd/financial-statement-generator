import type { TrialBalanceEntry } from '../../../types/financial';
import type { IAccountMappingProvider } from './IAccountMappingProvider';
import type { AccountMappingRules } from '../../../types/accountMapping';
import { AccountMappingUtils } from '../../../types/accountMapping';

export type NoteKey =
  | 'cash'
  | 'receivables'
  | 'inventory'
  | 'prepaid_expenses'
  | 'ppe_cost'
  | 'ppe_accum_depr'
  | 'other_assets'
  | 'bank_overdrafts'
  | 'payables'
  | 'short_term_loans'
  | 'income_tax_payable'
  | 'long_term_loans_fi'
  | 'long_term_loans_other';

export interface MatchedGroup {
  noteType: NoteKey;
  title: string;
  rules: AccountMappingRules | null;
  accounts: TrialBalanceEntry[];
  totalCurrent: number;
  totalPrevious: number;
}

export interface MappingPreviewResult {
  groups: MatchedGroup[];
  unmatched: TrialBalanceEntry[];
}

const NOTE_TITLES: Record<NoteKey, string> = {
  cash: 'เงินสดและรายการเทียบเท่าเงินสด',
  receivables: 'ลูกหนี้การค้าและลูกหนี้อื่น',
  inventory: 'สินค้าคงเหลือ',
  prepaid_expenses: 'ค่าใช้จ่ายจ่ายล่วงหน้า',
  ppe_cost: 'ที่ดิน อาคาร และอุปกรณ์ (ต้นทุน)',
  ppe_accum_depr: 'ค่าเสื่อมราคาสะสม',
  other_assets: 'สินทรัพย์อื่น',
  bank_overdrafts: 'เงินเบิกเกินบัญชี',
  payables: 'เจ้าหนี้การค้าและเจ้าหนี้อื่น',
  short_term_loans: 'เงินกู้ยืมระยะสั้น',
  income_tax_payable: 'ภาษีเงินได้นิติบุคคลค้างจ่าย',
  long_term_loans_fi: 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน',
  long_term_loans_other: 'เงินกู้ยืมระยะยาวอื่น'
};

export function buildMappingPreview(
  tb: TrialBalanceEntry[],
  provider: IAccountMappingProvider,
  includeNotes?: NoteKey[]
): MappingPreviewResult {
  const usedCodes = new Set<string>();
  const noteOrder: NoteKey[] = includeNotes ?? [
    'cash',
    'receivables',
    'inventory',
    'prepaid_expenses',
    'ppe_cost',
    'ppe_accum_depr',
    'other_assets',
    'bank_overdrafts',
    'payables',
    'short_term_loans',
    'income_tax_payable',
    'long_term_loans_fi',
    'long_term_loans_other'
  ];

  const groups: MatchedGroup[] = noteOrder.map(noteType => {
    const rules = provider.getRules(noteType);
    const accounts = rules
      ? AccountMappingUtils.getMatchingAccounts(tb, rules)
      : [];
    accounts.forEach(a => {
      if (a.accountCode) usedCodes.add(a.accountCode);
    });
    const totalCurrent = accounts.reduce((s, e) => s + Math.abs((e.currentBalance ?? e.balance ?? 0) as number), 0);
    const totalPrevious = accounts.reduce((s, e) => s + Math.abs((e.previousBalance ?? 0) as number), 0);
    return {
      noteType,
      title: NOTE_TITLES[noteType],
      rules: rules ?? null,
      accounts,
      totalCurrent,
      totalPrevious
    };
  });

  const unmatched = tb.filter(e => e.accountCode && !usedCodes.has(e.accountCode));

  return { groups, unmatched };
}
