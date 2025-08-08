import React, { useState, useRef } from 'react';
import type { UploadedFile } from '../types/financial';

interface FileUploadProps {
  onFileUpload: (file: File) => void;
  processing: boolean;
  error?: string;
}

const FileUpload: React.FC<FileUploadProps> = ({ onFileUpload, processing, error }) => {
  const [dragActive, setDragActive] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<UploadedFile | null>(null);
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

  const handleFileSelection = (file: File) => {
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
    onFileUpload(file);
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelection(e.target.files[0]);
    }
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="file-upload-container">
      <div
        className={`file-upload-area ${dragActive ? 'drag-active' : ''} ${processing ? 'processing' : ''}`}
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
          disabled={processing}
        />
        
        <div className="file-upload-content">
          {processing ? (
            <div className="processing-indicator">
              <div className="spinner"></div>
              <p>กำลังประมวลผล...</p>
            </div>
          ) : uploadedFile ? (
            <div className="file-info">
              <div className="file-icon">📊</div>
              <div className="file-details">
                <p className="file-name">{uploadedFile.name}</p>
                <p className="file-size">{formatFileSize(uploadedFile.size)}</p>
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
                ไฟล์ CSV (.csv) เท่านั้น
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

export default FileUpload;
