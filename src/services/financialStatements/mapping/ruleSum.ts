import type { TrialBalanceEntry } from '../../../types/financial';
import type { AccountMappingRules } from '../../../types/accountMapping';

export function sumByRules(tb: TrialBalanceEntry[], rules: AccountMappingRules, which: 'current' | 'previous' = 'current'): number {
  const includeCodes = new Set((rules.includes || []).map(n => Number(n)));
  const excludeCodes = new Set((rules.excludes || []).map(n => Number(n)));
  const ranges = rules.ranges || [];

  let sum = 0;
  for (const e of tb) {
    const code = Number(e.accountCode || '0');
    if (Number.isNaN(code)) continue;

    let matched = includeCodes.has(code);
    if (!matched && ranges.length > 0) {
      matched = ranges.some(r => code >= r.from && code <= r.to);
    }
    if (!matched) continue;
    if (excludeCodes.has(code)) continue;

    const val = which === 'previous' ? (e.previousBalance ?? 0) : (e.balance ?? e.currentBalance ?? 0);
    sum += val;
  }

  return Math.abs(sum);
}
