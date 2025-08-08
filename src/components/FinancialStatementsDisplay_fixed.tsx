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

  const formatFinancialStatement = (data: string[][]) => {
    return data.map((row, rowIndex) => {
      // Skip empty rows
      if (row.every(cell => !cell || cell.trim() === '')) {
        return (
          <tr key={rowIndex} style={{ height: '10px' }}>
            <td colSpan={8}>&nbsp;</td>
          </tr>
        );
      }

      return (
        <tr key={rowIndex}>
          {row.map((cell, cellIndex) => {
            let cellStyle: React.CSSProperties = {
              padding: '4px 8px',
              fontSize: '14px',
              fontFamily: 'Arial, sans-serif'
            };

            // Format based on content
            if (cell.includes(statements.companyInfo.name) && rowIndex < 5) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                fontSize: '16px',
                textAlign: 'center',
                backgroundColor: '#f8f9fa'
              };
            } else if (cell.includes('งบ') && rowIndex < 5) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                textAlign: 'center',
                backgroundColor: '#f8f9fa'
              };
            } else if (cell.includes('ณ วันที่') || cell.includes('สำหรับรอบระยะเวลา')) {
              cellStyle = {
                ...cellStyle,
                textAlign: 'center',
                backgroundColor: '#f8f9fa'
              };
            } else if (cell.includes('หมายเหตุ') || cell.includes('หน่วย:บาท')) {
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                textAlign: 'center',
                backgroundColor: '#e9ecef'
              };
            } else if (cell.includes('20') && cell.length === 4) {
              // Year headers
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                textAlign: 'center',
                backgroundColor: '#e9ecef'
              };
            } else if (cell.startsWith(',สินทรัพย์') || cell.startsWith(',หนี้สิน') || cell.startsWith(',รายได้') || cell.startsWith(',ค่าใช้จ่าย')) {
              // Main category headers
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                backgroundColor: '#f1f3f4',
                borderTop: '2px solid #ddd'
              };
            } else if (cell.startsWith(',รวม')) {
              // Total lines
              cellStyle = {
                ...cellStyle,
                fontWeight: 'bold',
                borderTop: '1px solid #333',
                borderBottom: '1px solid #333'
              };
            } else if (cell.includes('"') && (cell.includes(',') || cell.includes('.'))) {
              // Number cells
              cellStyle = {
                ...cellStyle,
                textAlign: 'right',
                fontFamily: 'monospace'
              };
            }

            // Remove commas and quotes for display
            let displayText = cell;
            if (cell.startsWith(',')) {
              displayText = cell.substring(1);
            }
            displayText = displayText.replace(/"/g, '');

            return (
              <td key={cellIndex} style={cellStyle}>
                {displayText}
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
      backgroundColor: '#fff',
      maxWidth: '1200px',
      margin: '0 auto'
    }}>
      {/* Header */}
      <div style={{ 
        textAlign: 'center', 
        marginBottom: '30px',
        padding: '20px',
        backgroundColor: '#2c3e50',
        color: 'white',
        borderRadius: '8px'
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
        <div style={{ 
          overflowX: 'auto',
          border: '1px solid #ddd',
          borderRadius: '8px',
          backgroundColor: '#fff'
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
        <div style={{ 
          overflowX: 'auto',
          border: '1px solid #ddd',
          borderRadius: '8px',
          backgroundColor: '#fff'
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
        <div style={{ 
          overflowX: 'auto',
          border: '1px solid #ddd',
          borderRadius: '8px',
          backgroundColor: '#fff'
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
