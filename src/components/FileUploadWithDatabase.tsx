import React, { useState, useRef } from 'react';
import type { UploadedFile, CompanyInfo } from '../types/financial';
import type { Company } from '../types/database';
import { CSVProcessor } from '../services/csvProcessor';
import { ApiService } from '../services/apiService';

interface FileUploadProps {
  selectedCompany: Company | null;
  onFileProcessed: (trialBalanceSetId: string, csvData: any, companyInfo: CompanyInfo) => void;
  processing: boolean;
  error?: string;
}

const FileUploadWithDatabase: React.FC<FileUploadProps> = ({ 
  selectedCompany, 
  onFileProcessed, 
  processing, 
  error 
}) => {
  const [dragActive, setDragActive] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<UploadedFile | null>(null);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [localProcessing, setLocalProcessing] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      handleFileSelection(file);
    }
  };

  const determineProcessingType = (csvData: any[]): 'single-year' | 'multi-year' => {
    // Check if there are previous balance columns with actual data
    const hasPreviousData = csvData.some(entry => 
      entry.previousBalance && entry.previousBalance !== 0
    );
    return hasPreviousData ? 'multi-year' : 'single-year';
  };

  const extractAccountCodeRange = (csvData: any[]): { min: string; max: string } => {
    const codes = csvData
      .map(entry => entry.accountCode)
      .filter(code => code && typeof code === 'string')
      .sort();
    
    return {
      min: codes[0] || '0000',
      max: codes[codes.length - 1] || '9999'
    };
  };

  const checkForInventory = (csvData: any[]): boolean => {
    return csvData.some(entry => 
      entry.accountCode === '1510' || 
      entry.accountName?.includes('สินค้าคงเหลือ') ||
      entry.accountName?.includes('Inventory')
    );
  };

  const checkForPPE = (csvData: any[]): boolean => {
    return csvData.some(entry => 
      entry.accountCode?.startsWith('13') || 
      entry.accountName?.includes('ที่ดิน') ||
      entry.accountName?.includes('อาคาร') ||
      entry.accountName?.includes('เครื่องจักร')
    );
  };

  const handleFileSelection = async (file: File) => {
    if (!selectedCompany) {
      alert('กรุณาเลือกบริษัทก่อนอัพโหลดไฟล์');
      return;
    }

    const isCSV = file.type === 'text/csv' || file.name.endsWith('.csv');
    
    if (!isCSV) {
      alert('กรุณาเลือกไฟล์ CSV (.csv) เท่านั้น');
      return;
    }

    const uploadedFile: UploadedFile = {
      file,
      name: file.name,
      size: file.size,
      type: file.type
    };

    setUploadedFile(uploadedFile);
    setLocalProcessing(true);
    setUploadProgress(0);

    try {
      // Step 1: Process CSV content
      const csvContent = await file.text();
      console.log('CSV Content loaded, length:', csvContent.length);
      
      // Create company info from selected company
      const companyInfo: CompanyInfo = {
        name: selectedCompany.name,
        type: selectedCompany.type,
        reportingPeriod: selectedCompany.defaultReportingYear?.toString() || new Date().getFullYear().toString(),
        reportingYear: selectedCompany.defaultReportingYear || new Date().getFullYear(),
        registrationNumber: selectedCompany.registrationNumber,
        address: selectedCompany.address,
        businessDescription: selectedCompany.businessDescription
      };

      // Step 2: Process CSV data
      const csvData = CSVProcessor.processCsvFile(csvContent, companyInfo);
      console.log('CSV Data processed:', csvData);
      
      setUploadProgress(30);

      // Step 3: Prepare metadata for database
      const processingType = determineProcessingType(csvData.trialBalance);
      const accountCodeRange = extractAccountCodeRange(csvData.trialBalance);
      
      const metadata = {
        fileName: file.name,
        processingType,
        reportingYear: companyInfo.reportingYear,
        reportingPeriod: companyInfo.reportingPeriod,
        totalEntries: csvData.trialBalance.length,
        hasInventory: checkForInventory(csvData.trialBalance),
        hasPPE: checkForPPE(csvData.trialBalance),
        accountCodeRange
      };

      setUploadProgress(60);

      // Step 4: Save to database
      console.log('Saving trial balance to database...');
      const result = await ApiService.saveTrialBalance(
        selectedCompany.id,
        csvData.trialBalance,
        metadata
      );

      console.log('Trial balance saved:', result);
      setUploadProgress(100);

      // Step 5: Notify parent component with database ID
      onFileProcessed(result.trialBalanceSetId, csvData, companyInfo);

    } catch (error) {
      console.error('Error processing and saving file:', error);
      alert(`เกิดข้อผิดพลาด: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setLocalProcessing(false);
      setUploadProgress(0);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelection(e.target.files[0]);
    }
  };

  const handleClick = () => {
    if (!selectedCompany) {
      alert('กรุณาเลือกบริษัทก่อนอัพโหลดไฟล์');
      return;
    }
    fileInputRef.current?.click();
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const isProcessing = processing || localProcessing;

  return (
    <div className="file-upload-container">
      <div
        className={`file-upload-area ${dragActive ? 'drag-active' : ''} ${isProcessing ? 'processing' : ''}`}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
        onClick={handleClick}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".csv"
          onChange={handleFileInput}
          style={{ display: 'none' }}
          disabled={isProcessing || !selectedCompany}
        />
        
        <div className="file-upload-content">
          {!selectedCompany ? (
            <div className="upload-disabled">
              <div className="disabled-icon">🏢</div>
              <p className="disabled-text">
                กรุณาเลือกบริษัทก่อนอัพโหลดไฟล์
              </p>
            </div>
          ) : isProcessing ? (
            <div className="processing-indicator">
              <div className="spinner"></div>
              <p>กำลังประมวลผลและบันทึกข้อมูล...</p>
              {uploadProgress > 0 && (
                <div className="progress-bar">
                  <div 
                    className="progress-fill" 
                    style={{ width: `${uploadProgress}%` }}
                  ></div>
                  <span className="progress-text">{uploadProgress}%</span>
                </div>
              )}
            </div>
          ) : uploadedFile ? (
            <div className="file-info">
              <div className="file-icon">📊</div>
              <div className="file-details">
                <p className="file-name">{uploadedFile.name}</p>
                <p className="file-size">{formatFileSize(uploadedFile.size)}</p>
                <p className="company-info">บริษัท: {selectedCompany.name}</p>
              </div>
              <button 
                className="change-file-btn"
                onClick={(e) => {
                  e.stopPropagation();
                  setUploadedFile(null);
                  if (fileInputRef.current) fileInputRef.current.value = '';
                }}
              >
                เปลี่ยนไฟล์
              </button>
            </div>
          ) : (
            <div className="upload-prompt">
              <div className="upload-icon">📤</div>
              <p className="upload-text">
                <strong>คลิกเพื่ออัพโหลด</strong> หรือลากไฟล์มาวาง
              </p>
              <p className="upload-subtext">
                ไฟล์ CSV (.csv) เท่านั้น - บริษัท: {selectedCompany.name}
              </p>
            </div>
          )}
        </div>
      </div>
      
      {error && (
        <div className="error-message">
          ⚠️ {error}
        </div>
      )}
    </div>
  );
};

export default FileUploadWithDatabase;
