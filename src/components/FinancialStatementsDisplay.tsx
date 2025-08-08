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
  
  const handleDownload = async () => {
    try {
      console.log('Download button clicked');
      await onDownload();
    } catch (error) {
      console.error('Download error:', error);
      alert('เกิดข้อผิดพลาดในการดาวน์โหลด กรุณาลองใหม่อีกครั้ง');
    }
  };

  const formatFinancialStatement = (data: (string | number | {f: string})[][]) => {
    return data.map((row, rowIndex) => {
      // Skip empty rows with consistent spacing, special handling for row 4 (spacer row like in BS Assets)
      if (row.every(cell => !cell || (typeof cell === 'string' && cell.trim() === ''))) {
        return (
          <tr key={rowIndex} style={{ height: rowIndex === 3 ? '15pt' : 'auto' }}>
            <td colSpan={10} style={{ 
              borderBottom: 'none',
              backgroundColor: 'transparent',
              padding: '0'
            }}>&nbsp;</td>
          </tr>
        );
      }

      return (
        <tr key={rowIndex}>
          {row.map((cell, cellIndex) => {
            // Handle different cell types
            let cellText: string;
            if (typeof cell === 'object' && cell !== null && 'f' in cell) {
              cellText = `=${cell.f}`;
            } else if (typeof cell === 'number') {
              // Format numbers with commas and 2 decimal places
              cellText = new Intl.NumberFormat('th-TH', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
              }).format(Math.abs(cell));
            } else if (cell === null || cell === undefined) {
              cellText = '';
            } else {
              cellText = String(cell);
            }
            
            let cellStyle: React.CSSProperties = {
              padding: '8px 12px',
              fontSize: '14px',
              fontFamily: 'TH Sarabun New, Arial, sans-serif',
              borderBottom: '1px solid #e5e7eb',
              borderRight: '1px solid #f3f4f6',
              lineHeight: '1.4',
              verticalAlign: 'middle'
            };

            // Format based on content - consistent across all statements
            if (cellText.includes(statements.companyInfo.name) && rowIndex < 5) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                fontSize: '18px',
                textAlign: 'center',
                backgroundColor: '#1e3a8a',
                color: 'white',
                padding: '12px'
              };
            } else if (cellText.includes('งบ') && rowIndex < 5) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                fontSize: '16px',
                textAlign: 'center',
                backgroundColor: '#3b82f6',
                color: 'white',
                padding: '10px'
              };
            } else if (cellText.includes('ณ วันที่') || cellText.includes('สำหรับรอบระยะเวลา')) {
              cellStyle = {
                ...cellStyle,
                textAlign: 'center',
                backgroundColor: '#60a5fa',
                color: 'white',
                fontWeight: '500',
                padding: '8px'
              };
            } else if (cellText.includes('หมายเหตุ') || cellText.includes('หน่วย:บาท')) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                textAlign: 'center',
                backgroundColor: '#374151',
                color: 'white',
                fontSize: '13px'
              };
            } else if (cellText.includes('20') && cellText.length === 4) {
              // Year headers
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                textAlign: 'center',
                backgroundColor: '#6b7280',
                color: 'white',
                fontSize: '15px'
              };
            } else if (cellText.startsWith('สินทรัพย์') || cellText.startsWith('หนี้สิน') || cellText.startsWith('รายได้') || cellText.startsWith('ค่าใช้จ่าย') || cellText.startsWith('ส่วนของ')) {
              // Main category headers
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                backgroundColor: '#f59e0b',
                color: 'white',
                borderTop: '2px solid #d97706',
                fontSize: '15px',
                paddingLeft: '16px'
              };
            } else if (cellText.startsWith('รวม')) {
              // Total lines
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                backgroundColor: '#10b981',
                color: 'white',
                borderTop: '2px solid #059669',
                borderBottom: '2px solid #059669',
                fontSize: '15px'
              };
            } else if (typeof cell === 'number' || (cellText.includes('.') && /^\d/.test(cellText.replace(/,/g, '')))) {
              // Number cells - consistent formatting
              cellStyle = {
                ...cellStyle,
                textAlign: 'right',
                fontFamily: 'TH Sarabun New, monospace',
                backgroundColor: '#fefefe',
                fontWeight: '600',
                color: '#1f2937',
                border: '1px solid #e5e7eb',
                fontSize: '14px',
                paddingRight: '16px'
              };
            } else if (cellIndex === 2 || cellIndex === 3) {
              // Account name columns - consistent indentation
              let indentLevel = 0;
              let displayText = cellText;
              
              if (cellText.startsWith('    ')) {
                indentLevel = 2;
                displayText = cellText.substring(4).trim();
              } else if (cellText.startsWith('  ')) {
                indentLevel = 1;
                displayText = cellText.substring(2).trim();
              }
              
              cellStyle = {
                ...cellStyle,
                backgroundColor: indentLevel === 2 ? '#f8fafc' : indentLevel === 1 ? '#f1f5f9' : '#ffffff',
                paddingLeft: `${16 + indentLevel * 20}px`,
                color: '#374151',
                fontWeight: indentLevel === 0 ? '600' : '500',
                borderLeft: indentLevel > 0 ? `3px solid ${indentLevel === 2 ? '#e5e7eb' : '#3b82f6'}` : 'none'
              };
              
              cellText = displayText;
            } else {
              // Default text styling
              cellStyle = {
                ...cellStyle,
                color: '#374151',
                fontWeight: '500'
              };
            }

            return (
              <td key={cellIndex} style={cellStyle}>
                {cellText}
              </td>
            );
          })}
        </tr>
      );
    });
  };

  return (
    <div style={{ 
      padding: '20px', 
      backgroundColor: '#f8fafc',
      minHeight: '100vh',
      maxWidth: '1400px',
      margin: '0 auto'
    }}>
      {/* Header */}
      <div style={{ 
        textAlign: 'center', 
        marginBottom: '30px',
        padding: '25px',
        background: 'linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%)',
        color: 'white',
        borderRadius: '12px',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
      }}>
        <h1 style={{ margin: '0 0 15px 0', fontSize: '24px' }}>งบการเงิน</h1>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <p style={{ margin: '5px 0' }}><strong>บริษัท:</strong> {statements.companyInfo.name}</p>
            <p style={{ margin: '5px 0' }}><strong>ปีบัญชี:</strong> {statements.companyInfo.reportingYear}</p>
            <p style={{ margin: '5px 0' }}><strong>ประเภท:</strong> {statements.processingType === 'single-year' ? 'งบการเงินปีเดียว' : 'งบการเงินเปรียบเทียบ'}</p>
          </div>
          <button 
            onClick={handleDownload}
            style={{
              backgroundColor: '#27ae60',
              color: 'white',
              padding: '12px 24px',
              border: 'none',
              borderRadius: '6px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              transition: 'background-color 0.3s'
            }}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#229954'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = '#27ae60'}
          >
            📥 ดาวน์โหลดเป็น Excel
          </button>
        </div>
      </div>

      {/* Balance Sheet - Assets */}
      <div style={{ marginBottom: '40px' }}>
        <h3 style={{ 
          backgroundColor: '#1e40af', 
          color: 'white',
          padding: '15px', 
          margin: '0 0 0 0',
          borderRadius: '8px 8px 0 0',
          fontSize: '18px',
          fontWeight: 'bold'
        }}>
          งบฐานะการเงิน - สินทรัพย์
        </h3>
        <div style={{ 
          overflowX: 'auto',
          border: '2px solid #1e40af',
          borderTop: 'none',
          borderRadius: '0 0 8px 8px',
          backgroundColor: '#ffffff',
          boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
        }}>
          <table style={{ 
            width: '100%', 
            borderCollapse: 'collapse',
            backgroundColor: 'white'
          }}>
            <tbody>
              {formatFinancialStatement(statements.balanceSheet.assets)}
            </tbody>
          </table>
        </div>
      </div>

      {/* Balance Sheet - Liabilities */}
      <div style={{ marginBottom: '40px' }}>
        <h3 style={{ 
          backgroundColor: '#dc2626', 
          color: 'white',
          padding: '15px', 
          margin: '0 0 0 0',
          borderRadius: '8px 8px 0 0',
          fontSize: '18px',
          fontWeight: 'bold'
        }}>
          งบฐานะการเงิน - หนี้สินและส่วนของผู้ถือหุ้น
        </h3>
        <div style={{ 
          overflowX: 'auto',
          border: '2px solid #dc2626',
          borderTop: 'none',
          borderRadius: '0 0 8px 8px',
          backgroundColor: '#ffffff',
          boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
        }}>
          <table style={{ 
            width: '100%', 
            borderCollapse: 'collapse',
            backgroundColor: 'white'
          }}>
            <tbody>
              {formatFinancialStatement(statements.balanceSheet.liabilities)}
            </tbody>
          </table>
        </div>
      </div>

      {/* Profit & Loss Statement */}
      <div style={{ marginBottom: '40px' }}>
        <h3 style={{ 
          backgroundColor: '#059669', 
          color: 'white',
          padding: '15px', 
          margin: '0 0 0 0',
          borderRadius: '8px 8px 0 0',
          fontSize: '18px',
          fontWeight: 'bold'
        }}>
          งบกำไรขาดทุน
        </h3>
        <div style={{ 
          overflowX: 'auto',
          border: '2px solid #059669',
          borderTop: 'none',
          borderRadius: '0 0 8px 8px',
          backgroundColor: '#ffffff',
          boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
        }}>
          <table style={{ 
            width: '100%', 
            borderCollapse: 'collapse',
            backgroundColor: 'white'
          }}>
            <tbody>
              {formatFinancialStatement(statements.profitLossStatement)}
            </tbody>
          </table>
        </div>
      </div>

      {/* Statement of Changes in Equity */}
      {statements.changesInEquity && (
        <div style={{ marginBottom: '40px' }}>
          <h3 style={{ 
            backgroundColor: '#7c3aed', 
            color: 'white',
            padding: '15px', 
            margin: '0 0 0 0',
            borderRadius: '8px 8px 0 0',
            fontSize: '18px',
            fontWeight: 'bold'
          }}>
            งบการเปลี่ยนแปลงส่วนของผู้ถือหุ้น
          </h3>
          <div style={{ 
            overflowX: 'auto',
            border: '2px solid #7c3aed',
            borderTop: 'none',
            borderRadius: '0 0 8px 8px',
            backgroundColor: '#ffffff',
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
          }}>
            <table style={{ 
              width: '100%', 
              borderCollapse: 'collapse',
              backgroundColor: 'white'
            }}>
              <tbody>
                {formatFinancialStatement(statements.changesInEquity)}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Notes to Financial Statements */}
      {statements.notes && statements.notes.length > 0 && (
        <div style={{ marginBottom: '40px' }}>
          <h3 style={{ 
            backgroundColor: '#6f42c1', 
            color: 'white',
            padding: '15px', 
            margin: '0 0 0 0',
            borderRadius: '8px 8px 0 0',
            fontSize: '18px',
            fontWeight: 'bold'
          }}>
            หมายเหตุประกอบงบการเงิน
          </h3>
          <div style={{ 
            overflowX: 'auto',
            border: '2px solid #6f42c1',
            borderTop: 'none',
            borderRadius: '0 0 8px 8px',
            backgroundColor: '#ffffff',
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
          }}>
            <table style={{ 
              width: '100%', 
              borderCollapse: 'collapse',
              backgroundColor: 'white'
            }}>
              <tbody>
                {formatFinancialStatement(statements.notes)}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Footer */}
      <div style={{ 
        textAlign: 'center', 
        marginTop: '30px',
        padding: '15px',
        backgroundColor: '#f8f9fa',
        borderRadius: '8px',
        border: '1px solid #dee2e6'
      }}>
        <p style={{ 
          color: '#6c757d', 
          fontSize: '14px',
          margin: '0'
        }}>
          งบการเงินถูกสร้างจากข้อมูล Trial Balance แบบ {statements.processingType === 'single-year' ? 'ปีเดียว' : 'เปรียบเทียบ'}
        </p>
      </div>
    </div>
  );
};

export default FinancialStatementsDisplay;
