// ============================================================================
// FINANCIAL STATEMENTS TYPE DEFINITIONS
// ============================================================================

/**
 * Interface for tracking row positions during note generation
 * Used to communicate exact formatting requirements to ExcelJS formatter
 */
export interface NoteRowTracker {
  currentRow: number;
  noteStartRow: number;
  headerRows: number[];        // Rows with note headers (should be bold)
  yearHeaderRows: number[];    // Rows with year headers (center, underline, general format)
  detailRows: number[];        // Rows with account details (normal formatting)
  totalRows: number[];         // Rows with "รวม" text (bold text, normal amounts)
  unitRows: number[];          // Rows with "หน่วย:บาท" (should be bold)
}

/**
 * Interface for note formatting information
 */
export interface NoteFormatter {
  type: 'cash' | 'receivables' | 'payables' | 'ppe' | 'inventory' | 'general';
  tracker: NoteRowTracker;
}

/**
 * Foundation-first architecture: Note calculations drive Balance Sheet values
 * This ensures perfect consistency between Notes and Balance Sheet
 */
export interface DetailedFinancialData {
  // FOUNDATION LAYER: Note calculations (calculated once, used everywhere)
  noteCalculations: {
    // Note 7: Cash and cash equivalents
    cash: {
      cash: { current: number; previous: number };          // เงินสดในมือ (1000)
      bankDeposits: { current: number; previous: number };  // เงินฝากธนาคาร (1010-1099)
      total: { current: number; previous: number };         // Total for Balance Sheet
    };
    
    // Note 8: Trade and other receivables (DYNAMIC - no artificial grouping)
    receivables: {
      total: { current: number; previous: number };                 // Total for Balance Sheet
      // Individual accounts provide the detailed breakdown (replaces tradeReceivables + otherReceivables)
    };
    
    // Note 9: Inventories (if applicable)
    inventory: {
      inventory: { current: number; previous: number };     // สินค้าคงเหลือ (1510)
      total: { current: number; previous: number };         // Total for Balance Sheet
    };
    
    // Note 10: Property, plant and equipment
    ppe: {
      cost: { current: number; previous: number };              // ราคาทุน
      accumulatedDepreciation: { current: number; previous: number }; // ค่าเสื่อมราคาสะสม
      netBookValue: { current: number; previous: number };      // มูลค่าตามบัญชี (for Balance Sheet)
    };
    
    // Note 12: Trade and other payables (DYNAMIC - no artificial grouping)
    payables: {
      total: { current: number; previous: number };             // Total for Balance Sheet
      // Individual accounts provide the detailed breakdown (replaces tradePayables + otherPayables)
    };
  };
  
  // INDIVIDUAL ACCOUNT DETAILS: Dynamic structure for note breakdowns
  individualAccounts: {
    // Cash accounts - automatically categorized for display
    cash: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
        category: 'cash' | 'bankDeposits'; // Auto-categorized based on code range
      };
    };
    
    // ALL individual receivable accounts (no artificial grouping)
    receivables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
      };
    };
    
    // ALL individual payable accounts (no artificial grouping)
    payables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
      };
    };
  };
  
  // BALANCE SHEET TOTALS: Derived from note calculations + individual items
  balanceSheetTotals: {
    assets: {
      cashAndCashEquivalents: { current: number; previous: number };    // From noteCalculations.cash.total
      tradeReceivables: { current: number; previous: number };          // From noteCalculations.receivables.total
      inventory: { current: number; previous: number };                 // From noteCalculations.inventory.total
      prepaidExpenses: { current: number; previous: number };           // Individual calculation (1300-1399)
      propertyPlantEquipment: { current: number; previous: number };    // From noteCalculations.ppe.netBookValue
      otherAssets: { current: number; previous: number };               // Individual calculation (1900-1999)
    };
    
    liabilities: {
      bankOverdraftsAndShortTermLoans: { current: number; previous: number }; // Individual (2001)
      tradeAndOtherPayables: { current: number; previous: number };           // From noteCalculations.payables.total
      shortTermBorrowings: { current: number; previous: number };             // Individual (2110)
      incomeTaxPayable: { current: number; previous: number };                // Individual (2120)
      longTermLoansFromFI: { current: number; previous: number };             // Individual (2410)
      otherLongTermLoans: { current: number; previous: number };              // Individual (2490)
    };
    
    equity: {
      paidUpCapital: { current: number; previous: number };
      retainedEarnings: { current: number; previous: number };
      openingRetainedEarnings: number; // Opening balance from account 3020 (credit - debit)
      legalReserve: { current: number; previous: number };
    };
  };
  
  // INCOME STATEMENT
  income: {
    revenue: { total: number; mainRevenue: number; otherIncome: number };
    expenses: { total: number; costOfServices: number; adminExpenses: number; otherExpenses: number; incomeTax: number; financialCosts: number };
    netProfit: number;
  };
  
  // BUSINESS LOGIC FLAGS
  flags: {
    hasInventory: boolean;
    isServiceBusiness: boolean;
    isLimitedPartnership: boolean;
  };
}

/**
 * Interface for tracking cell positions during balance sheet generation
 * Used to generate accurate formulas that reference the correct cell locations
 */
export interface CellTracker {
  currentRow: number;
  currentLiabilitiesRows: number[];
  nonCurrentLiabilitiesRows: number[];
  equityDataRows: number[];
  currentLiabilitiesTotalRow: number;
  nonCurrentLiabilitiesTotalRow: number;
  totalLiabilitiesRow?: number; // Track total liabilities row for grand total calculation
}
