<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Financial Statement Generator - Production Ready System

This is a comprehensive React TypeScript web application that processes Excel/CSV files containing trial balance data and generates professional financial statements following Thai accounting standards and VBA-compliant business logic.

## Project Structure

- **Frontend**: React with TypeScript, Vite build tool, responsive dark theme UI
- **Excel Processing**: ExcelJS library for reading/writing Excel files with professional formatting
- **CSV Processing**: Custom CSV processor with flexible column mapping and auto-detection
- **File Processing**: Supports both Excel (.xlsx) and CSV formats with validation
- **Downloads**: Generates professionally formatted Excel workbooks with multiple worksheets

## Comprehensive Features

### 1. **File Processing**
- **Multi-format Support**: Excel (.xlsx) and CSV file processing
- **Flexible CSV Processing**: Auto-detects delimiters and column mappings
- **Data Validation**: Trial balance validation with proper error handling
- **Previous Year Integration**: Handles multi-year comparatives from previousBalance field

### 2. **Financial Statement Generation**
- **Balance Sheet Assets (BS_Assets)**: VBA-compliant with formula-based totals
- **Balance Sheet Liabilities (BS_Liabilities)**: Comprehensive liability and equity sections
- **Profit & Loss Statement**: Revenue and expense categorization with proper classifications
- **Statement of Changes in Equity**: Both Limited Company and Partnership formats
- **Notes to Financial Statements**: Policy, Accounting details, and supplementary information

### 3. **Advanced Business Logic**
- **VBA Compliance**: Follows original Excel VBA system logic exactly
- **Account Code Ranges**: Precise mapping (1000-1099 for cash, 2010-2999 for payables, etc.)
- **Multi-Year Processing**: Comparative financial statements with previous year data
- **Company Type Detection**: Automatic detection and appropriate statement generation
- **Inventory vs Service Business**: Different cost structures based on account 1510 presence

## Technical Architecture

### **Core Services**
- **FinancialStatementGenerator**: Main engine with 2000+ lines of VBA-compliant logic
- **ExcelJSFormatter**: Professional Excel formatting with Thai accounting standards
- **CSVProcessor**: Flexible CSV parsing with auto-detection capabilities
- **ExcelProcessor**: Legacy Excel file processing support

### **Advanced Features**
- **Formula Integration**: Excel formulas for dynamic calculations (SUM, cell references)
- **Professional Formatting**: Bold headers, underlines, number formatting, column widths
- **Multi-Worksheet Output**: Generates 6+ worksheets per financial statement package
- **Error Recovery**: Comprehensive error handling with fallback mechanisms

## Excel Formatting Architecture

### **Critical Pattern: Separation of Data and Formatting**
- **Data Generation**: FinancialStatementGenerator creates plain string/number data arrays
- **Formatting Application**: ExcelJSFormatter handles ALL visual formatting via pattern recognition
- **Bold Text Implementation**: Use `formatTotalLinesProfessional()` with text pattern matching

### **Bold Formatting Pattern**
```typescript
// In formatTotalLinesProfessional() - ADD specific Thai text patterns:
else if (value === 'กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้' || 
         value === 'กำไรก่อนภาษีเงินได้' || 
         value === 'กำไร(ขาดทุน)สุทธิ') {
  this.formatKeyProfitLine(worksheet, row);
}

// Create corresponding formatting function:
private static formatKeyProfitLine(worksheet: ExcelJS.Worksheet, row: number): void {
  // Apply bold font formatting to entire row
}
```

### **Why This Pattern Works**
- **Maintainability**: All formatting logic centralized in ExcelJSFormatter
- **Flexibility**: Easy to add new bold patterns without touching data generation
- **Performance**: Single pass formatting after data insertion
- **Separation of Concerns**: Data logic separate from presentation logic

**IMPORTANT**: Never add `{text: string, bold: true}` objects in FinancialStatementGenerator. Always use plain strings and handle formatting in ExcelJSFormatter pattern recognition.

## Financial Statement Components

### **Balance Sheet (BS_Assets & BS_Liabilities)**
- **Current Assets**: Cash equivalents, trade receivables, inventory, prepaid expenses
- **Non-Current Assets**: PPE with depreciation calculations, other assets
- **Current Liabilities**: Trade payables, short-term borrowings, accrued expenses
- **Non-Current Liabilities**: Long-term loans, related party loans
- **Equity**: Registered capital, retained earnings with current year profit integration

### **Notes System**
- **Notes_Policy**: Accounting policies and basis of preparation
- **Notes_Accounting**: Detailed breakdowns (Cash, Receivables, PPE, Payables, etc.)
- **Notes_Detail**: Cost of goods sold details (DT1) and expense categorization (DT2)

### **Advanced Calculations**
- **Retained Earnings**: Opening balance + (4xxx revenue - 5xxx expenses)
- **PPE Depreciation**: Individual asset tracking with net book value calculations
- **Cost of Goods Sold**: Beginning inventory + purchases - ending inventory
- **Expense Classification**: Selling (5300-5311), Admin (5312-5350), Other (5351+)

## Data Processing Capabilities

### **CSV Format Support**
```csv
ชื่อบัญชี,รหัสบัญชี,ยอดยกมาต้นงวด,ยอดยกมางวดนี้,เดบิต,เครดิต
เงินสดในมือ,1000,50000,25000,0,0
ลูกหนี้การค้า,1140,0,175014.81,0,0
```

### **Account Code Mapping**
- **1000-1099**: Cash and cash equivalents
- **1140-1215**: Trade receivables and other current receivables
- **1500-1519**: Inventory accounts
- **2010-2999**: Current payables (with smart exclusions)
- **3020**: Retained earnings (with profit integration)
- **4xxx**: Revenue accounts (credit - debit)
- **5xxx**: Expense accounts (debit - credit)

## Quality Assurance

### **VBA Compliance Testing**
- Account code ranges match original VBA system exactly
- Formula calculations replicate VBA logic
- Multi-year processing follows original decision tree
- Company type detection uses VBA-compliant patterns

### **Professional Standards**
- Thai accounting standards compliance
- Proper financial statement formatting
- Audit-ready detailed notes
- Excel formula transparency for verification

## Development Standards

- **TypeScript**: Strict type safety with comprehensive interfaces
- **Modular Architecture**: Separation of concerns with dedicated service classes
- **Error Handling**: Comprehensive try-catch with user-friendly messages
- **Performance**: Efficient data processing for large trial balance files
- **Maintainability**: Well-documented code with clear business logic separation

## Production Readiness

- **Backup System**: Git version control with GitHub repository
- **Debug Capabilities**: Comprehensive logging for troubleshooting
- **Formula Verification**: Excel formulas for audit trail and transparency
- **Multi-Environment**: Supports development and production builds
- **User Experience**: Professional UI with progress indicators and error feedback

This system represents a complete migration from Excel VBA to modern web technology while maintaining 100% compatibility with the original business logic and financial statement requirements.