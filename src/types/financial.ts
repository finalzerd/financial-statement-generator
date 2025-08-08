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
  address?: string;
  businessDescription?: string; // Business type/description for Notes_Policy
  reportingPeriod: string;
  reportingYear: number;
  shares?: number; // Number of shares for Limited Company
  shareValue?: number; // Par value per share for Limited Company
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
  notes: (string | number | {f: string})[][];
  accountingNotes: (string | number | {f: string})[][];
  changesInEquity?: any;
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
