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
  type:
    | 'cash'
    | 'receivables'
    | 'payables'
    | 'ppe'
    | 'inventory'
    | 'general'
    | 'shortTermLoans'
    | 'assetShortTermLoans'
    | 'assetLongTermLoans'
    | 'hirePurchaseCreditors'
    | 'bankOverdrafts'
    | 'otherCurrentAssets'
    | 'otherCurrentLiabilities'
    | 'otherNonCurrentLiabilities';
  tracker: NoteRowTracker;
}

/**
 * Registry of actual note numbers assigned during generation
 * Used to link Balance Sheet line items to their corresponding notes
 */
export interface NoteRegistry {
  cash?: number;
  receivables?: number;
  assetShortTermLoans?: number;
  otherCurrentAssets?: number;
  ppe?: number;
  bankOverdrafts?: number;
  payables?: number;
  liabilityShortTermLoans?: number;
  otherCurrentLiabilities?: number;
  otherAssets?: number;
  assetLongTermLoans?: number;
  longTermLoansFromFI?: number;
  hirePurchaseCreditors?: number;
  otherLongTermLoans?: number;
  otherNonCurrentLiabilities?: number;
}

// Note category classification for selection-first architecture
export type NoteCategory =
  | 'cash'
  | 'asset_short_term_loans'
  | 'asset_long_term_loans'
  | 'hire_purchase_creditors'
  | 'receivables'
  | 'inventory'
  | 'inventory_purchases'
  | 'inventory_purchase_returns'
  | 'inventory_purchase_discounts'
  | 'prepaid'
  | 'other_current_assets'
  | 'ppe_cost'
  | 'ppe_accum_depr'
  | 'other_assets'
  | 'bank_overdrafts'
  | 'payables'
  | 'short_term_loans'
  | 'income_tax_payable'
  | 'other_current_liabilities'
  | 'long_term_loans_fi'
  | 'long_term_loans_other'
  | 'other_non_current_liabilities'
  | 'other_income'
  | 'detail_service_costs'
  // P&L categories (selection-first for Profit & Loss)
  | 'revenue'
  | 'admin_expenses'
  | 'other_expenses'
  | 'income_tax_expense'
  | 'financial_costs'
  | 'selling_expenses';

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

    // Asset loans given
    assetShortTermLoans: { current: number; previous: number }; // เงินให้กู้ยืมระยะสั้น (e.g., 1141)
    assetLongTermLoans: { current: number; previous: number };  // เงินให้กู้ยืมระยะยาว (e.g., 1710)
    
    hirePurchaseCreditors: {
      principal: { current: number; previous: number };      // เจ้าหนี้ตามสัญญาเช่าซื้อ (2015)
      deferredCharges: { current: number; previous: number }; // ดอกผลเช่าซื้อรอตัดบัญชี (1644.2)
      taxCredit: { current: number; previous: number };       // ภาษีซื้อรอตัดบัญชี (1644.1)
      total: { current: number; previous: number };           // Total for Balance Sheet
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

    // Additional note-derived totals that feed Balance Sheet directly
    prepaid: { current: number; previous: number };             // ค่าใช้จ่ายจ่ายล่วงหน้า (1400-1439)
    otherAssets: { current: number; previous: number };         // สินทรัพย์อื่น (1660-1700)
    bankOverdrafts: { current: number; previous: number };      // เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน (2001-2009)
    shortTermLoans: { current: number; previous: number };      // เงินกู้ยืมระยะสั้น (2030)
    incomeTaxPayable: { current: number; previous: number };    // ภาษีเงินได้นิติบุคคลค้างจ่าย (2045)
    longTermLoansFi: { current: number; previous: number };     // เงินกู้ยืมระยะยาวจากสถาบันการเงิน (2120-2123 ยกเว้น 2121)
    longTermLoansOther: { current: number; previous: number };  // เงินกู้ยืมระยะยาวอื่น (2050-2052,2100-2119)
  };
  
  // INDIVIDUAL ACCOUNT DETAILS: Dynamic structure for note breakdowns
  individualAccounts: {
    // Cash accounts - automatically categorized for display
    cash: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
        // Selection-first: which note this account belongs to
        noteCategory?: NoteCategory; // 'cash'
        category: 'cash' | 'bankDeposits'; // Auto-categorized based on code range
      };
    };
    
    // ALL individual receivable accounts (no artificial grouping)
    receivables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
        // Selection-first: which note this account belongs to
        noteCategory?: NoteCategory; // 'receivables'
      };
    };
    
    // ALL individual payable accounts (no artificial grouping)
    payables: {
      [accountCode: string]: {
        accountName: string;
        current: number;
        previous: number;
        // Selection-first: which note this account belongs to
        noteCategory?: NoteCategory; // 'payables'
      };
    };
  };
  
  // BALANCE SHEET TOTALS: Derived from note calculations + individual items
  balanceSheetTotals: {
    assets: {
      cashAndCashEquivalents: { current: number; previous: number };    // From noteCalculations.cash.total
      shortTermLoansGiven?: { current: number; previous: number };       // From noteCalculations.assetShortTermLoans
      tradeReceivables: { current: number; previous: number };          // From noteCalculations.receivables.total
      inventory: { current: number; previous: number };                 // From noteCalculations.inventory.total
      prepaidExpenses: { current: number; previous: number };           // Individual calculation (1300-1399)
      propertyPlantEquipment: { current: number; previous: number };    // From noteCalculations.ppe.netBookValue
      longTermLoansGiven?: { current: number; previous: number };        // From noteCalculations.assetLongTermLoans
      otherAssets: { current: number; previous: number };               // Individual calculation (1900-1999)
    };
    
    liabilities: {
      bankOverdraftsAndShortTermLoans: { current: number; previous: number }; // Individual (2001)
      tradeAndOtherPayables: { current: number; previous: number };           // From noteCalculations.payables.total
      shortTermBorrowings: { current: number; previous: number };             // Individual (2110)
      incomeTaxPayable: { current: number; previous: number };                // Individual (2120)
      longTermLoansFromFI: { current: number; previous: number };             // Individual (2410)
      otherLongTermLoans: { current: number; previous: number };              // Individual (2490)
      hirePurchaseCreditors?: { current: number; previous: number };          // From noteCalculations.hirePurchaseCreditors.total
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
