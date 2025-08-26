import { useState } from 'react';
import { AccountMappingManager } from './AccountMappingManager';
import type { Company } from '../types/database';
import './CompanyMappingPage.css';

interface CompanyMappingPageProps {
  selectedCompany: Company | null;
  trialBalanceData?: any[];
}

export function CompanyMappingPage({ selectedCompany, trialBalanceData = [] }: CompanyMappingPageProps) {
  const [showMappingManager, setShowMappingManager] = useState(false);

  if (!selectedCompany) {
    return (
      <div className="company-mapping-page">
        <div className="no-company-selected">
          <h2>Account Mapping Configuration</h2>
          <p>Please select a company to configure account mappings.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="company-mapping-page">
      <div className="page-header">
        <h1>Account Mapping Configuration</h1>
        <p>Company: <strong>{selectedCompany.name}</strong></p>
        {!showMappingManager && (
          <div className="intro-section">
            <p>
              Customize how your trial balance account codes map to financial statement notes. 
              This allows you to adapt the system to your company's specific chart of accounts.
            </p>
            <div className="features-list">
              <h3>Features:</h3>
              <ul>
                <li>✅ Flexible account ranges (e.g., 1000-1099 for cash accounts)</li>
                <li>✅ Include specific non-sequential accounts</li>
                <li>✅ Exclude accounts from broader ranges</li>
                <li>✅ Real-time validation with your trial balance data</li>
                <li>✅ Coverage reporting to ensure all accounts are mapped</li>
              </ul>
            </div>
            <button 
              onClick={() => setShowMappingManager(true)} 
              className="btn-primary"
            >
              Configure Account Mappings
            </button>
          </div>
        )}
      </div>

      {showMappingManager && (
        <>
          <div className="mapping-controls">
            <button 
              onClick={() => setShowMappingManager(false)} 
              className="btn-secondary"
            >
              ← Back to Overview
            </button>
          </div>

          <AccountMappingManager
            companyId={selectedCompany.id}
            trialBalanceData={trialBalanceData}
            onMappingsChanged={() => {
              console.log('Account mappings have been updated');
              // Here you could refresh any cached data or show success message
            }}
          />
        </>
      )}
    </div>
  );
}
