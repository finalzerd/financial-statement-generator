import React from 'react';
import type { FinancialStatements } from '../types/financial';

interface FinancialStatementsDisplayProps {
  statements: FinancialStatements;
  onDownload: () => void;
}

const FinancialStatementsDisplay: React.FC<FinancialStatementsDisplayProps> = ({ 
  statements, 
  onDownload 
}) => {
  const formatTableData = (data: string[][]) => {
    return data.map((row, rowIndex) => (
      <tr key={rowIndex}>
        {row.map((cell, cellIndex) => (
          <td key={cellIndex} style={{ 
            padding: '4px 8px', 
            borderBottom: '1px solid #eee',
            whiteSpace: 'pre-wrap',
            fontFamily: 'monospace'
          }}>
            {cell}
          </td>
        ))}
      </tr>
    ));
  };

  return (
    <div className="financial-statements-display" style={{ padding: '20px' }}>
      <div className="statements-header" style={{ marginBottom: '30px', textAlign: 'center' }}>
        <h2>งบการเงิน</h2>
        <p><strong>บริษัท:</strong> {statements.companyInfo.name}</p>
        <p><strong>ปีบัญชี:</strong> {statements.companyInfo.reportingYear}</p>
        <p><strong>ประเภท:</strong> {statements.processingType === 'single-year' ? 'งบการเงินปีเดียว' : 'งบการเงินเปรียบเทียบ'}</p>
        
        <button 
          onClick={onDownload}
          style={{
            backgroundColor: '#4CAF50',
            color: 'white',
            padding: '10px 20px',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            fontSize: '16px',
            marginTop: '15px'
          }}
        >
          📥 ดาวน์โหลดเป็น Excel
        </button>
      </div>

      {/* Balance Sheet - Assets */}
      <div className="statement-section" style={{ marginBottom: '40px' }}>
        <h3 style={{ backgroundColor: '#f5f5f5', padding: '10px', margin: '0 0 15px 0' }}>
          งบฐานะการเงิน - สินทรัพย์
        </h3>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #ddd' }}>
            <tbody>
              {formatTableData(statements.balanceSheet.assets)}
            </tbody>
          </table>
        </div>
      </div>

      {/* Balance Sheet - Liabilities */}
      <div className="statement-section" style={{ marginBottom: '40px' }}>
        <h3 style={{ backgroundColor: '#f5f5f5', padding: '10px', margin: '0 0 15px 0' }}>
          งบฐานะการเงิน - หนี้สินและส่วนของผู้ถือหุ้น
        </h3>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #ddd' }}>
            <tbody>
              {formatTableData(statements.balanceSheet.liabilities)}
            </tbody>
          </table>
        </div>
      </div>

      {/* Profit & Loss Statement */}
      <div className="statement-section" style={{ marginBottom: '40px' }}>
        <h3 style={{ backgroundColor: '#f5f5f5', padding: '10px', margin: '0 0 15px 0' }}>
          งบกำไรขาดทุน
        </h3>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #ddd' }}>
            <tbody>
              {formatTableData(statements.profitLossStatement)}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ textAlign: 'center', marginTop: '30px' }}>
        <p style={{ color: '#666', fontSize: '14px' }}>
          งบการเงินถูกสร้างจากข้อมูล Trial Balance แบบ {statements.processingType === 'single-year' ? 'ปีเดียว' : 'เปรียบเทียบ'}
        </p>
      </div>
    </div>
  );
};

export default FinancialStatementsDisplay;
