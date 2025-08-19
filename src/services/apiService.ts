export class ApiService {
  private static readonly BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001/api';

  // ============== HEALTH CHECK ==============
  
  static async checkHealth() {
    const response = await fetch(`${this.BASE_URL}/health`);
    const result = await response.json();
    
    if (!response.ok) {
      throw new Error('API health check failed');
    }
    
    return result;
  }

  // ============== COMPANY OPERATIONS ==============

  static async getCompanies() {
    const response = await fetch(`${this.BASE_URL}/companies`);
    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.message || 'Failed to fetch companies');
    }
    
    // Handle different API response formats
    return {
      companies: result.data || result.companies || []
    };
  }

  static async createCompany(companyData: {
    name: string;
    type: string;
    registrationNumber?: string;
    address?: string;
    businessDescription?: string;
    taxId?: string;
    defaultReportingYear?: number;
    numberOfShares?: number;
    shareValue?: number;
  }) {
    const response = await fetch(`${this.BASE_URL}/companies`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(companyData)
    });

    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.details || result.error || 'Failed to create company');
    }

    return result;
  }

  static async updateCompany(companyId: string, companyData: {
    name: string;
    type: string;
    registrationNumber?: string;
    address?: string;
    businessDescription?: string;
    taxId?: string;
    defaultReportingYear?: number;
    numberOfShares?: number;
    shareValue?: number;
  }) {
    const response = await fetch(`${this.BASE_URL}/companies/${companyId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(companyData)
    });

    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.details || result.error || 'Failed to update company');
    }

    return result;
  }

  static async deleteCompany(companyId: string) {
    const response = await fetch(`${this.BASE_URL}/companies/${companyId}`, {
      method: 'DELETE'
    });

    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.details || result.error || 'Failed to delete company');
    }

    return result;
  }

  // ============== FILE UPLOAD OPERATIONS ==============

  static async uploadFile(
    file: File,
    companyId: string,
    onProgress?: (progress: number) => void
  ) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('companyId', companyId);

    const xhr = new XMLHttpRequest();
    
    return new Promise((resolve, reject) => {
      xhr.upload.addEventListener('progress', (e) => {
        if (e.lengthComputable && onProgress) {
          onProgress((e.loaded / e.total) * 100);
        }
      });

      xhr.addEventListener('load', () => {
        if (xhr.status === 200) {
          try {
            const result = JSON.parse(xhr.responseText);
            resolve(result);
          } catch (error) {
            reject(new Error('Invalid response format'));
          }
        } else {
          try {
            const error = JSON.parse(xhr.responseText);
            reject(new Error(error.message || 'Upload failed'));
          } catch {
            reject(new Error('Upload failed'));
          }
        }
      });

      xhr.addEventListener('error', () => {
        reject(new Error('Network error during upload'));
      });

      xhr.open('POST', `${this.BASE_URL}/upload`);
      xhr.send(formData);
    });
  }

  // ============== TRIAL BALANCE OPERATIONS ==============

  static async saveTrialBalance(
    companyId: string,
    trialBalanceEntries: any[],
    metadata: {
      fileName: string;
      processingType: 'single-year' | 'multi-year';
      reportingYear: number;
      reportingPeriod: string;
      totalEntries: number;
      hasInventory: boolean;
      hasPPE: boolean;
      accountCodeRange: { min: string; max: string };
    }
  ) {
    const response = await fetch(`${this.BASE_URL}/trial-balance`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        companyId,
        trialBalanceEntries,
        metadata
      })
    });

    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.details || result.error || 'Failed to save trial balance');
    }

    return result;
  }

  static async getTrialBalanceSets(companyId: string) {
    const response = await fetch(`${this.BASE_URL}/trial-balance/company/${companyId}`);
    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.message || 'Failed to fetch trial balance sets');
    }
    
    return result;
  }

  static async getTrialBalanceEntries(trialBalanceSetId: string) {
    const response = await fetch(`${this.BASE_URL}/trial-balance/${trialBalanceSetId}/entries`);
    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.message || 'Failed to fetch trial balance entries');
    }
    
    return result;
  }

  // ============== FINANCIAL STATEMENT OPERATIONS ==============

  static async saveGeneratedStatements(
    trialBalanceSetId: string,
    statements: any,
    excelBuffer: ArrayBuffer
  ) {
    const formData = new FormData();
    formData.append('trialBalanceSetId', trialBalanceSetId);
    formData.append('statements', JSON.stringify(statements));
    formData.append('excelFile', new Blob([excelBuffer]), 'statements.xlsx');

    const response = await fetch(`${this.BASE_URL}/statements`, {
      method: 'POST',
      body: formData
    });

    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.details || result.error || 'Failed to save statements');
    }

    return result;
  }

  static async getStatements(trialBalanceSetId: string) {
    const response = await fetch(`${this.BASE_URL}/statements/${trialBalanceSetId}`);
    const result = await response.json();
    
    if (!response.ok) {
      throw new Error(result.message || 'Failed to fetch statements');
    }
    
    return result;
  }

  static async downloadStatement(statementId: string) {
    const response = await fetch(`${this.BASE_URL}/statements/${statementId}/download`);
    
    if (!response.ok) {
      throw new Error('Failed to download statement');
    }
    
    return response.blob();
  }
}
