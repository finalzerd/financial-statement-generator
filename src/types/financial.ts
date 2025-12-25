// Types for financial statement generation

export interface TrialBalanceEntry {
  accountCode: string;
  accountName: string;
  debitAmount: number;
  creditAmount: number;
  balance: number;
  // Additional fields for PPE movement calculation
  previousBalance?: number;  // ยอดยกมาต้นงวด from CSV
  currentBalance?: number;   // ยอดยกมางวดนี้ from CSV
}

export interface CompanyInfo {
  name: string;
  type: 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด'; // Limited Partnership or Limited Company
  registrationNumber?: string;
  registrationDate?: string; // Registration date for Notes_Policy
  address?: string;
  businessDescription?: string; // Business type/description for Notes_Policy
  reportingPeriod: string;
  reportingYear: number;
  reportingPeriodStartDate?: string; // Accounting period start date (e.g., "1 มกราคม")
  reportingPeriodEndDate?: string; // Accounting period end date (e.g., "31 ธันวาคม")
  shares?: number; // Number of shares for Limited Company
  shareValue?: number; // Par value per share for Limited Company
  // Director signature fields
  directorName?: string;
  approvalMeetingNumber?: string;
  approvalMeetingDate?: string;
}

export interface SheetValidation {
  trialBalanceCount: number;
  trialPLCount: number;
  isValid: boolean;
  processingType: 'single-year' | 'multi-year' | 'invalid';
}

export interface ProcessingResult {
  success: boolean;
  message: string;
  data?: any;
  errors?: string[];
}

export interface AccountClassification {
  salesExpenses: TrialBalanceEntry[];
  adminExpenses: TrialBalanceEntry[];
  otherExpenses: TrialBalanceEntry[];
  financialCosts: TrialBalanceEntry[];
}

export interface InventoryInfo {
  hasInventory: boolean;
  inventoryAccount?: TrialBalanceEntry;
  purchaseAccounts: TrialBalanceEntry[];
}

export interface FinancialStatements {
  balanceSheet: any;
  profitLossStatement: any;
  profitLossSignatureRows?: number[];
  notes: (string | number | {f: string})[][];
  accountingNotes: (string | number | {f: string})[][];
  accountingNotesFormatters?: any[]; // Row tracking formatters for specific note formatting
  changesInEquity?: any;
  changesInEquitySignatureRows?: number[];
  detailNotes?: {
    detail1?: (string | number | {f: string})[][];
    detail2?: (string | number | {f: string})[][];
  };
  companyInfo: CompanyInfo;
  processingType: 'single-year' | 'multi-year';
}

export interface UploadedFile {
  file: File;
  name: string;
  size: number;
  type: string;
}

// Generic statement result with signature row metadata for formatting
export interface StatementResult {
  data: (string | number | { f: string })[][];
  signatureRows: number[]; // Row indices (1-based) for center-across-selection formatting
}

// Alias for backward compatibility
export type BalanceSheetResult = StatementResult;
