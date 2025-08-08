import { useState, useEffect } from 'react'
import FileUploadWithDatabase from './components/FileUploadWithDatabase'
import FinancialStatementsDisplay from './components/FinancialStatementsDisplay'
import SampleFileDownloader from './components/SampleFileDownloader'
import CompanyInfoForm from './components/CompanyInfoForm'
import CompanySelector from './components/CompanySelector'
import type { CompanyInfoFormData } from './components/CompanyInfoForm'
import { CSVProcessor } from './services/csvProcessor'
import { FinancialStatementGenerator } from './services/financialStatementGenerator'
import { ApiService } from './services/apiService'
import type { FinancialStatements, CompanyInfo } from './types/financial'
import type { Company } from './types/database'
import './App.css'

function App() {
  // Processing state
  const [processing, setProcessing] = useState(false)
  const [error, setError] = useState<string | undefined>()
  const [financialStatements, setFinancialStatements] = useState<FinancialStatements | null>(null)
  
  // Company management state
  const [selectedCompany, setSelectedCompany] = useState<Company | null>(null)
  const [showCompanySelector, setShowCompanySelector] = useState(true)
  
  // Legacy form state (for companies not yet in database)
  const [showCompanyForm, setShowCompanyForm] = useState(false)
  const [pendingFile, setPendingFile] = useState<File | null>(null)

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
      
      // Generate financial statements using existing logic
      const statements = statementGenerator.generateFinancialStatements(
        csvData.trialBalance,
        companyInfo,
        csvData.processingType,
        csvData.trialBalancePrevious
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
      
      // Generate financial statements (the generator will handle multi-year logic internally)
      const statements = statementGenerator.generateFinancialStatements(
        csvData.trialBalance,
        companyInfo,
        csvData.processingType,
        csvData.trialBalancePrevious
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
        businessDescription: formData.businessDescription
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
      </header>
      
      <main className="app-main">
        {showCompanySelector ? (
          <CompanySelector
            onCompanySelected={handleCompanySelected}
            onCompanyCreated={handleCompanyCreated}
          />
        ) : (
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
