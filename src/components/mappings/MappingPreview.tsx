import React from 'react';
import type { TrialBalanceEntry, CompanyInfo } from '../../types/financial';
import type { IAccountMappingProvider } from '../../services/financialStatements/mapping/IAccountMappingProvider';
import { buildMappingPreview, type NoteKey } from '../../services/financialStatements/mapping/mappingPreview';

type Props = {
  trialBalance: TrialBalanceEntry[];
  company: CompanyInfo;
  provider: IAccountMappingProvider;
  processingType: 'single-year' | 'multi-year';
};

const cell = { padding: '6px 8px', borderBottom: '1px solid #2a2a2a' } as const;
const th = { ...cell, fontWeight: 600, borderBottom: '1px solid #444' } as const;

export const MappingPreview: React.FC<Props> = ({ trialBalance, company, provider, processingType }) => {
  const { groups, unmatched } = React.useMemo(
    () => buildMappingPreview(trialBalance, provider),
    [trialBalance, provider]
  );

  const renderTable = (rows: TrialBalanceEntry[]) => (
    <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 8 }}>
      <thead>
        <tr>
          <th style={th}>รหัสบัญชี</th>
          <th style={th}>ชื่อบัญชี</th>
          <th style={{ ...th, textAlign: 'right' }}>{company.reportingYear}</th>
          {processingType === 'multi-year' && (
            <th style={{ ...th, textAlign: 'right' }}>{company.reportingYear - 1}</th>
          )}
        </tr>
      </thead>
      <tbody>
        {rows.map((e) => (
          <tr key={e.accountCode}>
            <td style={cell}>{e.accountCode}</td>
            <td style={cell}>{e.accountName}</td>
            <td style={{ ...cell, textAlign: 'right' }}>{
              Math.abs((e.currentBalance ?? e.balance ?? 0) as number).toLocaleString()
            }</td>
            {processingType === 'multi-year' && (
              <td style={{ ...cell, textAlign: 'right' }}>{
                Math.abs(e.previousBalance ?? 0).toLocaleString()
              }</td>
            )}
          </tr>
        ))}
        {rows.length === 0 && (
          <tr>
            <td style={cell} colSpan={processingType === 'multi-year' ? 4 : 3}>ไม่มีรายการ</td>
          </tr>
        )}
      </tbody>
    </table>
  );

  const order: NoteKey[] = [
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
    'long_term_loans_other',
  ];

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      {order.map((key) => {
        const g = groups.find((x) => x.noteType === key)!;
        return (
          <section key={key} style={{ padding: 12, border: '1px solid #2a2a2a', borderRadius: 6 }}>
            <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
              <h3 style={{ margin: 0 }}>{g.title}</h3>
              <small style={{ opacity: 0.8 }}>
                รวมปัจจุบัน: {g.totalCurrent.toLocaleString()}
                {processingType === 'multi-year' ? ` / ก่อนหน้า: ${g.totalPrevious.toLocaleString()}` : ''}
              </small>
            </header>
            {renderTable(g.accounts)}
          </section>
        );
      })}

      <section style={{ padding: 12, border: '1px solid #4a3', borderRadius: 6 }}>
        <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
          <h3 style={{ margin: 0 }}>บัญชีที่ไม่ถูกจัดเข้าหมายเหตุ</h3>
          <small style={{ opacity: 0.8 }}>จำนวน: {unmatched.length}</small>
        </header>
        {renderTable(unmatched)}
      </section>
    </div>
  );
};

export default MappingPreview;
