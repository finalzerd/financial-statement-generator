import { useState, useEffect } from 'react'
import FileUploadWithDatabase from './components/FileUploadWithDatabase'
import FinancialStatementsDisplay from './components/FinancialStatementsDisplay'
import SampleFileDownloader from './components/SampleFileDownloader'
import CompanyInfoForm from './components/CompanyInfoForm'
import CompanySelector from './components/CompanySelector'
import { AccountMappingManager } from './components/AccountMappingManager'
import type { CompanyInfoFormData } from './components/CompanyInfoForm'
import { CSVProcessor } from './services/csvProcessor'
import { FinancialStatementGenerator } from './services/financialStatementGenerator'
import { ApiService } from './services/apiService'
import type { FinancialStatements, CompanyInfo } from './types/financial'
import type { Company } from './types/database'
import { DynamicMappingProvider } from './services/financialStatements/mapping/DynamicMappingProvider'
import { StaticMappingProvider } from './services/financialStatements/mapping/StaticMappingProvider'
import type { IAccountMappingProvider } from './services/financialStatements/mapping/IAccountMappingProvider'
import MappingPreview from './components/mappings/MappingPreview'
import './App.css'

function App() {
  // Processing state
  const [processing, setProcessing] = useState(false)
  const [error, setError] = useState<string | undefined>()
  const [financialStatements, setFinancialStatements] = useState<FinancialStatements | null>(null)
  
  // Company management state
  const [selectedCompany, setSelectedCompany] = useState<Company | null>(null)
  const [showCompanySelector, setShowCompanySelector] = useState(true)
  
  // Navigation state
  const [activeTab, setActiveTab] = useState<'generate' | 'mappings'>('generate')
  
  // Legacy form state (for companies not yet in database)
  const [showCompanyForm, setShowCompanyForm] = useState(false)
  const [pendingFile, setPendingFile] = useState<File | null>(null)
  
  // Store trial balance data for mapping validation
  const [currentTrialBalanceData, setCurrentTrialBalanceData] = useState<any[]>([])
  const [currentCompanyInfo, setCurrentCompanyInfo] = useState<CompanyInfo | null>(null)
  const [currentProcessingType, setCurrentProcessingType] = useState<'single-year' | 'multi-year'>('single-year')
  const [currentMappingProvider, setCurrentMappingProvider] = useState<IAccountMappingProvider | null>(null)

  const statementGenerator = new FinancialStatementGenerator()

  // Check API health on app start
  useEffect(() => {
    ApiService.checkHealth()
      .then(() => console.log('API connection established'))
      .catch(err => console.warn('API connection failed:', err))
  }, [])

  // Enhanced file processing with database integration
  const handleFileProcessed = async (trialBalanceSetId: string, csvData: any, companyInfo: CompanyInfo) => {
    setProcessing(true)
    setError(undefined)
    setFinancialStatements(null)

    try {
      console.log('Processing file with database integration:', { trialBalanceSetId, csvData, companyInfo });
      
  // Store trial balance data and metadata for mapping preview/validation
  setCurrentTrialBalanceData(csvData.trialBalance || []);
  setCurrentCompanyInfo(companyInfo);
  setCurrentProcessingType(csvData.processingType);
      
      // Build mapping provider (dynamic from DB if available, otherwise static defaults)
      let provider = undefined as undefined | DynamicMappingProvider | StaticMappingProvider
      try {
        if (selectedCompany?.id) {
          const mappings = await ApiService.getCompanyAccountMappings(selectedCompany.id)
          if (Array.isArray(mappings) && mappings.length > 0) {
            provider = new DynamicMappingProvider(mappings)
          } else {
            provider = new StaticMappingProvider()
          }
        }
      } catch (e) {
        console.warn('Failed to load account mappings, using static defaults:', e)
        provider = new StaticMappingProvider()
      }

      // Persist provider for Mapping Preview tab
      if (provider) setCurrentMappingProvider(provider)

      // Generate financial statements using provider-aware logic
      const statements = statementGenerator.generateFinancialStatements(
        csvData.trialBalance,
        companyInfo,
        csvData.processingType,
        csvData.trialBalancePrevious,
        provider
      )
      console.log('Financial statements generated:', statements)
      
      // Save generated statements to database
      try {
        // For now, just save the statement data without Excel buffer
        // TODO: Implement getExcelBuffer method in FinancialStatementGenerator
        await ApiService.saveGeneratedStatements(trialBalanceSetId, statements, new ArrayBuffer(0));
        console.log('Statements saved to database');
      } catch (saveError) {
        console.warn('Failed to save statements to database:', saveError);
        // Continue anyway - user can still see the statements
      }
      
      setFinancialStatements(statements)
      console.log('Financial statements set in state')
      
    } catch (error) {
      console.error('Error processing file:', error)
      setError(error instanceof Error ? error.message : 'An unexpected error occurred')
    } finally {
      setProcessing(false)
    }
  }

  const handleCompanySelected = (company: Company) => {
    setSelectedCompany(company);
    setShowCompanySelector(false);
  }

  const handleCompanyCreated = (company: Company) => {
    setSelectedCompany(company);
    setShowCompanySelector(false);
  }

  const handleBackToCompanySelector = () => {
    setSelectedCompany(null);
    setShowCompanySelector(true);
    setFinancialStatements(null);
    setError(undefined);
    setCurrentTrialBalanceData([]);
    setActiveTab('generate');
  }

  const processCsvFile = async (file: File, companyInfo: CompanyInfo) => {
    setProcessing(true)
    setError(undefined)
    setFinancialStatements(null)

    try {
      // Read CSV file content
      const csvContent = await file.text()
      console.log('CSV Content loaded, length:', csvContent.length)
      
      // Process CSV data with multi-year support
      const csvData = CSVProcessor.processCsvFile(csvContent, companyInfo)
      console.log('CSV Data processed:', csvData)
      
  // Store trial balance data and metadata for mapping preview/validation
  setCurrentTrialBalanceData(csvData.trialBalance || []);
  setCurrentCompanyInfo(companyInfo);
  setCurrentProcessingType(csvData.processingType);
      
      // Build mapping provider (dynamic from DB if available, otherwise static defaults)
      let provider = undefined as undefined | DynamicMappingProvider | StaticMappingProvider
      try {
        if (selectedCompany?.id) {
          const mappings = await ApiService.getCompanyAccountMappings(selectedCompany.id)
          if (Array.isArray(mappings) && mappings.length > 0) {
            provider = new DynamicMappingProvider(mappings)
          } else {
            provider = new StaticMappingProvider()
          }
        }
      } catch (e) {
        console.warn('Failed to load account mappings, using static defaults:', e)
        provider = new StaticMappingProvider()
      }

      // Persist provider for Mapping Preview tab
      if (provider) setCurrentMappingProvider(provider)

      // Generate financial statements (the generator will handle multi-year logic internally)
      const statements = statementGenerator.generateFinancialStatements(
        csvData.trialBalance,
        companyInfo,
        csvData.processingType,
        csvData.trialBalancePrevious,
        provider
      )
      console.log('Financial statements generated:', statements)
      
      setFinancialStatements(statements)
      console.log('Financial statements set in state')
      
    } catch (error) {
      console.error('Error processing file:', error)
      setError(error instanceof Error ? error.message : 'An unexpected error occurred')
    } finally {
      setProcessing(false)
    }
  }

  const handleCompanyInfoSubmit = (formData: CompanyInfoFormData) => {
    console.log('Company info submitted:', formData)
    if (pendingFile) {
      const companyInfo: CompanyInfo = {
        name: formData.companyName,
        type: formData.companyType,
        reportingPeriod: formData.reportingPeriod,
        reportingYear: parseInt(formData.reportingPeriod) || new Date().getFullYear(),
        registrationNumber: formData.registrationNumber,
        address: formData.address,
        businessDescription: formData.businessDescription,
        shares: formData.numberOfShares,
        shareValue: formData.shareValue
      }
      console.log('Processing file with company info:', companyInfo)
      
      processCsvFile(pendingFile, companyInfo)
    }
    
    setShowCompanyForm(false)
    setPendingFile(null)
  }

  const handleCompanyInfoCancel = () => {
    setShowCompanyForm(false)
    setPendingFile(null)
  }

  const handleDownload = async () => {
    console.log('handleDownload called');
    if (financialStatements) {
      console.log('Financial statements exist, calling downloadAsExcel');
      try {
        await statementGenerator.downloadAsExcel(financialStatements);
        console.log('Download completed successfully');
      } catch (error) {
        console.error('Download failed:', error);
        alert('เกิดข้อผิดพลาดในการดาวน์โหลด: ' + (error instanceof Error ? error.message : 'Unknown error'));
      }
    } else {
      console.log('No financial statements available');
      alert('ไม่มีข้อมูลงบการเงินสำหรับดาวน์โหลด');
    }
  }

  return (
    <div className="app">
      <header className="app-header">
        <h1>📊 ระบบสร้างงบการเงิน</h1>
        <p>อัพโหลดไฟล์ CSV Trial Balance เพื่อสร้างงบการเงิน</p>
        {selectedCompany && (
          <div className="selected-company-info">
            <p>กำลังทำงานกับ: <strong>{selectedCompany.name}</strong></p>
            <button 
              onClick={handleBackToCompanySelector}
              className="btn btn-secondary btn-sm"
            >
              เปลี่ยนบริษัท
            </button>
          </div>
        )}
        
        {/* Navigation Tabs */}
        {selectedCompany && !showCompanySelector && (
          <div className="navigation-tabs">
            <button 
              className={`tab-button ${activeTab === 'generate' ? 'active' : ''}`}
              onClick={() => setActiveTab('generate')}
            >
              📊 สร้างงบการเงิน
            </button>
            <button 
              className={`tab-button ${activeTab === 'mappings' ? 'active' : ''}`}
              onClick={() => setActiveTab('mappings')}
            >
              ⚙️ จัดการรหัสบัญชี
            </button>
          </div>
        )}
      </header>
      
      <main className="app-main">
        {showCompanySelector ? (
          <CompanySelector
            onCompanySelected={handleCompanySelected}
            onCompanyCreated={handleCompanyCreated}
          />
        ) : (
          <>
            {/* Content based on active tab */}
            {activeTab === 'generate' ? (
              <>
                <div className="upload-section">
                  <FileUploadWithDatabase
                    selectedCompany={selectedCompany}
                    onFileProcessed={handleFileProcessed}
                    processing={processing}
                    error={error}
                  />
                </div>
                
                <SampleFileDownloader />
                
                {financialStatements && (
                  <div className="results-section">
                    <FinancialStatementsDisplay
                      statements={financialStatements}
                      onDownload={handleDownload}
                    />
                  </div>
                )}
              </>
            ) : (
              <div className="mappings-section">
                <div className="mappings-header">
                  <h2>⚙️ จัดการรหัสบัญชี</h2>
                  <p>ปรับแต่งการจับคู่รหัสบัญชีกับหมายเหตุประกอบงบการเงิน</p>
                  {currentTrialBalanceData.length > 0 && (
                    <div className="trial-balance-info">
                      <span className="info-badge">
                        📋 Trial Balance: {currentTrialBalanceData.length} รายการ
                      </span>
                    </div>
                  )}
                </div>
                
                <AccountMappingManager
                  companyId={selectedCompany?.id || '1'}
                  trialBalanceData={currentTrialBalanceData}
                  onMappingsChanged={() => {
                    console.log('Account mappings updated - you may want to regenerate statements');
                    // Optionally clear financial statements to force regeneration
                    // setFinancialStatements(null);
                  }}
                />

                {/* Mapping Preview: show matched/unmatched accounts per note */}
                {currentTrialBalanceData.length > 0 && currentCompanyInfo && currentMappingProvider && (
                  <div style={{ marginTop: 24 }}>
                    <h3>🔎 ตัวอย่างการจับคู่บัญชีตามหมายเหตุ</h3>
                    <MappingPreview
                      trialBalance={currentTrialBalanceData}
                      company={currentCompanyInfo}
                      provider={currentMappingProvider}
                      processingType={currentProcessingType}
                    />
                  </div>
                )}
              </div>
            )}
          </>
        )}
        
        {showCompanyForm && (
          <div className="modal-overlay">
            <div className="modal-content">
              <CompanyInfoForm
                onSubmit={handleCompanyInfoSubmit}
                onCancel={handleCompanyInfoCancel}
              />
            </div>
          </div>
        )}
      </main>
    </div>
  )
}

export default App
