// Database types for Financial Statement Generator
// Based on simplified schema without audit trails and monitoring

export interface Company {
  id: string;
  name: string;
  type: 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด';
  registrationNumber?: string;
  address?: string;
  businessDescription?: string; // Business type/description
  taxId?: string;
  defaultReportingYear: number;
  createdAt: Date;
  updatedAt: Date;
}

export interface TrialBalanceSet {
  id: string;
  companyId: string;
  reportingYear: number;
  reportingPeriod: string;
  processingType: 'single-year' | 'multi-year';
  fileName: string;
  fileHash: string;
  uploadedAt: Date;
  status: 'processed' | 'archived';
  metadata: {
    totalEntries: number;
    hasInventory: boolean;
    hasPPE: boolean;
    accountCodeRange: { min: string; max: string };
  };
}

export interface TrialBalanceEntry {
  id: string;
  trialBalanceSetId: string;
  accountCode: string;
  accountName: string;
  previousBalance: number;
  currentBalance: number;
  debitAmount: number;
  creditAmount: number;
  calculatedBalance: number;
  accountCategory: 'asset' | 'liability' | 'equity' | 'revenue' | 'expense';
}

export interface GeneratedStatement {
  id: string;
  trialBalanceSetId: string;
  statementType: string;
  data: any;
  excelBuffer?: Buffer;
  generatedAt: Date;
  
  // ExcelJSFormatter configuration - single source of truth
  fontName: string;
  fontSize: number;
  columnWidths: number[];
  hasGreenBackground: boolean;
}

export interface StatementTemplate {
  id: string;
  name: string;
  type: 'balance-sheet' | 'profit-loss' | 'equity' | 'notes';
  companyType: 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด' | 'both';
  isDefault: boolean;
  createdAt: Date;
}

export interface AccountMapping {
  id: string;
  templateId: string;
  accountCodeStart: string;
  accountCodeEnd: string;
  excludeCodes?: string[];
  statementLine: string;
  displayOrder: number;
  formula?: string;
  isTotalLine: boolean;
  indentLevel: number;
}

export interface CompanySetting {
  id: string;
  companyId: string;
  settingKey: string;
  settingValue: string;
  createdAt: Date;
}

// Re-export existing types from the Financial Statement Generator
export type { CompanyInfo, FinancialStatements } from './financial';

// CSV types from csvProcessor service
export interface CSVTrialBalanceEntry {
  ชื่อบัญชี: string;
  รหัสบัญชี: string;
  ยอดยกมาต้นงวด: number;
  ยอดยกมางวดนี้: number;
  เดบิต: number;
  เครดิต: number;
}
