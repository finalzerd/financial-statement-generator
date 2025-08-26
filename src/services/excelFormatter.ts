import ExcelJS from 'exceljs';

export class ExcelJSFormatter {
  
  /**
   * Thai font name - using widely available Thai font
   */
  static readonly THAI_FONT_NAME = 'TH Sarabun New';
  static readonly THAI_FONT_FALLBACK = 'Angsana New';
  
  /**
   * Standard colors for Thai financial statements - matching the professional format
   */
  static readonly COLORS = {
    HEADER_GREEN: '92D050',     // Green background for company header
    HEADER_BG: 'E7E6E6',       // Light gray for column headers
    BORDER: '000000',          // Black borders
    TEXT: '000000',            // Black text
    TOTAL_BG: 'D9D9D9',       // Gray background for totals
    COMPANY_NAME: '000000',    // Black company name text
    WHITE: 'FFFFFF'            // White background
  };
  
  /**
   * Create a new workbook with Thai formatting
   */
  static createWorkbook(): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    
    // Set workbook properties
    workbook.creator = 'Financial Statement Generator';
    workbook.lastModifiedBy = 'Financial Statement Generator';
    workbook.created = new Date();
    workbook.modified = new Date();
    
    return workbook;
  }
  
  /**
   * Add data to worksheet and return the worksheet
   */
  static addDataToWorksheet(workbook: ExcelJS.Workbook, worksheetName: string, data: (string | number | {f: string})[][]): ExcelJS.Worksheet {
    const worksheet = workbook.addWorksheet(worksheetName);
    
    // Add data row by row
    data.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        const cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
        
        if (typeof cellValue === 'object' && 'f' in cellValue && cellValue.f) {
          // Handle Excel formulas
          cell.value = { formula: cellValue.f };
        } else if (typeof cellValue === 'number') {
          // Direct numeric value - no conversion needed
          cell.value = cellValue;
        } else if (typeof cellValue === 'string') {
          // Check if string contains a formatted number
          const numericValue = this.parseFormattedNumber(cellValue);
          if (numericValue !== null) {
            // It's a formatted number string - convert to actual number
            cell.value = numericValue;
          } else {
            // It's actual text content - preserve zeros as numbers, empty as empty
            cell.value = cellValue !== undefined && cellValue !== null ? cellValue : '';
          }
        } else {
          // Handle any other case
          cell.value = '';
        }
      });
    });
    
    return worksheet;
  }
  
  /**
   * Parse formatted number strings back to numeric values
   */
  private static parseFormattedNumber(value: string): number | null {
    if (!value || typeof value !== 'string') return null;
    
    // Remove common formatting characters but preserve negative indicators
    const cleanValue = value.trim()
      .replace(/,/g, '')           // Remove commas
      .replace(/\s/g, '')          // Remove spaces
      .replace(/฿/g, '')           // Remove Thai Baht symbol
      .replace(/บาท/g, '')         // Remove "บาท" text
      .replace(/^\(/, '-')         // Convert (123) to -123
      .replace(/\)$/, '');         // Remove closing parenthesis
    
    // Check if it's a valid number after cleaning
    if (cleanValue === '' || cleanValue === '-') return null;
    
    const parsed = parseFloat(cleanValue);
    return isNaN(parsed) ? null : parsed;
  }
  
  /**
   * Format Balance Sheet Assets worksheet - Professional Thai format
   */
  /**
   * Set default worksheet font to ensure consistent width calculations
   */
  private static setWorksheetDefaultFont(worksheet: ExcelJS.Worksheet): void {
    // Set default font for the entire worksheet - ensure ALL cells use TH Sarabun New
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        // Always ensure the font name is TH Sarabun New, preserve other properties
        if (!cell.font) {
          cell.font = {
            name: this.THAI_FONT_NAME,
            size: 14,
            color: { argb: 'FF000000' }
          };
        } else {
          // Update existing font to ensure it uses TH Sarabun New
          cell.font = {
            ...cell.font,
            name: this.THAI_FONT_NAME
          };
        }
      });
    });
  }

  /**
   * Calculate width based on font size adjustment
   */
  private static adjustWidthForFont(baseWidth: number, fontSize: number = 14): number {
    // Base calculation assumes 11pt font, adjust for our 14pt font
    const fontSizeRatio = fontSize / 11;
    const adjustedWidth = Math.round(baseWidth * fontSizeRatio * 100) / 100; // Round to 2 decimal places
    console.log(`Width: ${baseWidth} -> ${adjustedWidth} (ratio: ${fontSizeRatio.toFixed(2)})`);
    return adjustedWidth;
  }

  /**
   * Debug height calculations and compare with expected values
   */
  private static debugHeightCalculations(worksheet: ExcelJS.Worksheet): void {
    console.log('=== HEIGHT DEBUGGING ===');
    console.log('Font:', this.THAI_FONT_NAME, '- Size: 14pt');
    console.log('Strategy: Autofit for all rows except row 4 (fixed spacer)');
    console.log('');
    console.log('Row | ExcelJS | Type       | Description');
    console.log('----|---------|------------|------------');
    
    const descriptions = ['Company Name', 'Statement Title', 'Date', 'Spacer (Fixed)', 'Column Headers'];
    
    // Check first 5 rows specifically
    for (let i = 1; i <= 5; i++) {
      const row = worksheet.getRow(i);
      const actualHeight = row.height;
      const heightType = actualHeight ? `Fixed: ${actualHeight}` : 'Autofit';
      const desc = descriptions[i - 1] || '';
      
      console.log(`  ${i}  |  ${(actualHeight || 'auto').toString().padStart(6)}  | ${heightType.padStart(10)} | ${desc}`);
    }
    
    // Check a few data rows
    console.log('');
    console.log('Sample Data Rows (should all be autofit):');
    for (let i = 6; i <= 10; i++) {
      const row = worksheet.getRow(i);
      const actualHeight = row.height;
      const heightType = actualHeight ? `Fixed: ${actualHeight}` : 'Autofit';
      
      console.log(`  ${i}  |  ${(actualHeight || 'auto').toString().padStart(6)}  | ${heightType.padStart(10)} | Data Row ${i}`);
    }
    
    console.log('=== END HEIGHT DEBUG ===');
    console.log('Note: Autofit rows will adjust to content and font automatically in Excel');
  }

  /**
   * Debug width calculations for Balance Sheet/P&L format
   */
  private static debugWidthCalculationsBS(worksheet: ExcelJS.Worksheet): void {
    console.log('=== BS/PL WIDTH DEBUGGING ===');
    console.log('Font:', this.THAI_FONT_NAME, '- Size: 14pt');
    console.log('Target: BS/PL format - Original VBA widths with font adjustment');
    console.log('');
    console.log('Column | ExcelJS | VBA Target | Difference');
    console.log('-------|---------|------------|----------');
    
    const targetBSWidths = [5, 7, 8, 7, 28, 7, 14, 2, 14]; // Original BS/PL VBA widths
    
    worksheet.columns.forEach((column, index) => {
      const letter = String.fromCharCode(65 + index);
      const actualWidth = column.width || 0;
      const targetWidth = targetBSWidths[index] || 0;
      const difference = Math.abs(actualWidth - targetWidth);
      
      console.log(`   ${letter}   |  ${actualWidth.toFixed(2).padStart(6)} |     ${targetWidth.toString().padStart(6)}     |   ${difference.toFixed(2).padStart(6)}`);
    });
    
    console.log('=== END BS/PL WIDTH DEBUG ===');
  }

  /**
   * Debug width calculations for Statement of Changes in Equity format
   */
  private static debugWidthCalculationsSCE(worksheet: ExcelJS.Worksheet): void {
    console.log('=== SCE WIDTH DEBUGGING ===');
    console.log('Font:', this.THAI_FONT_NAME, '- Size: 14pt');
    console.log('Target: SCE format - A=30, B=2, C=14, D=2, E=2, F=14, G=2, H=2, I=14');
    console.log('');
    console.log('Column | ExcelJS | VBA Target | Difference');
    console.log('-------|---------|------------|----------');
    
    const targetSCEWidths = [30, 2, 14, 2, 2, 14, 2, 2, 14]; // SCE VBA widths
    
    worksheet.columns.forEach((column, index) => {
      const letter = String.fromCharCode(65 + index);
      const actualWidth = column.width || 0;
      const targetWidth = targetSCEWidths[index] || 0;
      const difference = Math.abs(actualWidth - targetWidth);
      
      console.log(`   ${letter}   |  ${actualWidth.toFixed(2).padStart(6)} |     ${targetWidth.toString().padStart(6)}     |   ${difference.toFixed(2).padStart(6)}`);
    });
    
    console.log('=== END SCE WIDTH DEBUG ===');
  }

  /**
   * Format Balance Sheet Assets (no green background)
   */
  static formatBalanceSheetAssets(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying professional Thai Balance Sheet formatting (no green background)');
    
    // Set column widths with font size adjustment for Balance Sheet format
    // Original VBA widths: [5, 7, 8, 7, 21, 7, 14, 2, 14]
    // Adjusted down to compensate for 14pt font scaling (÷1.27)
    const baseWidths = [3.94, 5.51, 6.30, 5.51, 16.54, 5.51, 11.02, 1.57, 11.02];
    console.log('BS/PL Base widths (compensated for font scaling):', baseWidths);
    
    const adjustedWidths = baseWidths.map(width => this.adjustWidthForFont(width, 14));
    console.log('Final BS/PL calculated widths:', adjustedWidths);
    
    worksheet.columns = [
      { width: adjustedWidths[0] },    // A
      { width: adjustedWidths[1] },    // B
      { width: adjustedWidths[2] },    // C
      { width: adjustedWidths[3] },    // D
      { width: adjustedWidths[4] },    // E
      { width: adjustedWidths[5] },    // F
      { width: adjustedWidths[6] },    // G
      { width: adjustedWidths[7] },    // H
      { width: adjustedWidths[8] }     // I
    ];
    
    // Debug width calculations for BS/PL
    this.debugWidthCalculationsBS(worksheet);
    
    // Apply all formatting sections
    this.formatCompanyHeaderNoBackground(worksheet);
    this.formatColumnHeadersProfessional(worksheet);
    this.formatAccountLinesProfessional(worksheet);
    this.formatTotalLinesProfessional(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    // Debug height calculations after setting heights
    this.debugHeightCalculations(worksheet);
    
    // Set consistent font across worksheet
    this.setWorksheetDefaultFont(worksheet);
    
    // Clear all empty cells to prevent text cutoff issues
    this.clearEmptyCells(worksheet);
    
    console.log('Professional Thai formatting completed (BS/PL format with no green background)');
  }

  /**
   * Format Statement of Changes in Equity (no green background, SCE-specific layout)
   */
  static formatStatementOfChangesInEquity(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying SCE-specific formatting (no green background)');
    
    // Set column widths based on VBA specification for Statement of Changes in Equity
    // Adjusted from original VBA: A=30, B=2, C=14, D:E=2, F=14, G:H=2, I=14
    // Current custom sizes: A=28, B=2, C=14, D:E=2, F=14, G:H=2, I=14
    worksheet.columns = [
      { width: 28 },  // A = 28 (slightly reduced from 30)
      { width: 2 },   // B = 2  
      { width: 14 },  // C = 14 (kept original)
      { width: 2 },   // D = 2
      { width: 2 },   // E = 2
      { width: 14 },  // F = 14 (kept original)
      { width: 2 },   // G = 2
      { width: 2 },   // H = 2
      { width: 14 }   // I = 14 (kept original)
    ];
    
    console.log('Applied SCE Column Widths (Custom): A=28, B=2, C=14, D=2, E=2, F=14, G=2, H=2, I=14');
    
    // Merge row 5 from column C to I
    worksheet.mergeCells('C5:I5');
    console.log('Merged cells C5:I5');
    
    // Apply all formatting sections
    this.formatCompanyHeaderNoBackground(worksheet);
    this.formatColumnHeadersProfessional(worksheet);
    this.formatAccountLinesProfessional(worksheet);
    this.formatTotalLinesProfessional(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    // Special formatting for Statement of Changes in Equity
    this.formatStatementOfChangesInEquitySpecific(worksheet);
    
    // Set consistent font across worksheet
    this.setWorksheetDefaultFont(worksheet);
    
    // Clear all empty cells to prevent text cutoff issues
    this.clearEmptyCells(worksheet);
    
    console.log('Professional Thai formatting completed (SCE format with specific formatting)');
  }

  /**
   * Format other financial statements (with green background)
   */
  static formatOtherStatements(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying professional Thai formatting (with green background)');
    
    // Set TH Sarabun New font for all cells
    this.setWorksheetDefaultFont(worksheet);
    
    // Set column widths to match the professional format
    worksheet.columns = [
      { width: 3 },   // A - Empty/spacer
      { width: 45 },  // B - Account names (extra wide for Thai text with indentation)
      { width: 3 },   // C - Spacer
      { width: 12 },  // D - Note references (หมายเหตุ)
      { width: 3 },   // E - Spacer
      { width: 18 },  // F - Current year amounts (2567)
      { width: 18 }   // G - Previous year amounts (2566)
    ];
    
    // Apply all formatting sections - use header with green background for other sheets
    this.formatCompanyHeaderProfessional(worksheet);
    this.formatColumnHeadersProfessional(worksheet);
    this.formatAccountLinesProfessional(worksheet);
    this.formatTotalLinesProfessional(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    console.log('Professional Thai formatting completed (with green background)');
  }
  
  /**
   * Format company header section - Professional Thai style (no green background for BS_Assets)
   */
  private static formatCompanyHeaderNoBackground(worksheet: ExcelJS.Worksheet): void {
    // Company name (A1) - Merge cells A1:I1 with no background or borders
    worksheet.mergeCells('A1:I1');
    const companyCell = worksheet.getCell('A1');
    companyCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    companyCell.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // Statement title (A2) - Merge cells A2:I2
    worksheet.mergeCells('A2:I2');
    const titleCell = worksheet.getCell('A2');
    titleCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // Date (A3) - Merge cells A3:I3 with thick bottom border
    worksheet.mergeCells('A3:I3');
    const dateCell = worksheet.getCell('A3');
    dateCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      color: { argb: 'FF000000' }
    };
    dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // Add thick bottom border to entire row 3 (A3:I3)
    for (let col = 1; col <= 9; col++) {
      const cell = worksheet.getCell(3, col);
      cell.border = {
        bottom: { style: 'thick', color: { argb: 'FF000000' } }
      };
    }
  }

  /**
   * Format company header section - Professional Thai style
   */
  private static formatCompanyHeaderProfessional(worksheet: ExcelJS.Worksheet): void {
    // Company name (B1) - Merge cells B1:G1 and add green background
    worksheet.mergeCells('B1:G1');
    const companyCell = worksheet.getCell('B1');
    companyCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    companyCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF92D050' }  // Green background like in image
    };
    companyCell.alignment = { horizontal: 'center', vertical: 'middle' };
    companyCell.border = this.getAllBorders();
    
    // Statement title (B2) - Merge cells B2:G2
    worksheet.mergeCells('B2:G2');
    const titleCell = worksheet.getCell('B2');
    titleCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.border = this.getAllBorders();
    
    // Date (B3) - Merge cells B3:G3
    worksheet.mergeCells('B3:G3');
    const dateCell = worksheet.getCell('B3');
    dateCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      color: { argb: 'FF000000' }
    };
    dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
    dateCell.border = this.getAllBorders();
  }
  
  /**
   * Format column headers - Professional Thai style matching the image
   */
  private static formatColumnHeadersProfessional(worksheet: ExcelJS.Worksheet): void {
    // Row 5: Format หมายเหตุ in F5 (bold + underline) and หน่วย:บาท in I5 (bold)
    const headerCells: { cell: string; text: string; bold: boolean; underline?: boolean }[] = [
      { cell: 'F5', text: 'หมายเหตุ', bold: true, underline: true },
      { cell: 'I5', text: 'หน่วย:บาท', bold: true }
    ];
    
    headerCells.forEach(({ cell, text, bold, underline }) => {
      const headerCell = worksheet.getCell(cell);
      headerCell.value = text;
      headerCell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: bold,
        underline: underline || false,
        color: { argb: 'FF000000' }
      };
      headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Row 6: Format สินทรัพย์ in B6 (bold) and year columns in general format
    const b6Cell = worksheet.getCell('B6');
    b6Cell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    b6Cell.alignment = { horizontal: 'left', vertical: 'middle' };

    // Year columns in G6 and I6 - MUST be general format (not bold) and center aligned
    const yearCells = ['G6', 'I6'];
    yearCells.forEach(cellAddress => {
      const cell = worksheet.getCell(cellAddress);
      // Format regardless of content to ensure General format is applied
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: false, // General format - not bold
        color: { argb: 'FF000000' }
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      // Force General format - this is critical for proper display
      cell.numFmt = 'General';
      
      console.log(`Applied General format to ${cellAddress}: numFmt = ${cell.numFmt}`);
    });
    
    // Clear old positions that might have been set incorrectly
    const clearCells = ['D5', 'E5', 'G5'];
    clearCells.forEach(cellAddress => {
      const clearCell = worksheet.getCell(cellAddress);
      clearCell.value = '';
    });
  }
  
  /**
   * Format account detail lines - Professional Thai style with proper indentation
   */
  private static formatAccountLinesProfessional(worksheet: ExcelJS.Worksheet): void {
    // Apply base formatting to all data rows (7-40) - Skip row 6 to preserve header formatting
    for (let row = 7; row <= 40; row++) {
      // Account names (Columns B through E)
      ['B', 'C', 'D', 'E'].forEach(col => {
        const nameCell = worksheet.getCell(`${col}${row}`);
        nameCell.font = { 
          name: this.THAI_FONT_NAME, 
          size: 14,
          color: { argb: 'FF000000' }
        };
        nameCell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
      
      // Note references (Column F - moved from D)
      const noteCell = worksheet.getCell(`F${row}`);
      noteCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
      noteCell.alignment = { horizontal: 'center', vertical: 'middle' };
      
      // Amount columns (G and I) - Apply number format but exclude row 6 (headers)
      ['G', 'I'].forEach(col => {
        const amountCell = worksheet.getCell(`${col}${row}`);
        amountCell.font = { 
          name: this.THAI_FONT_NAME, 
          size: 14,
          color: { argb: 'FF000000' }
        };
        amountCell.alignment = { horizontal: 'right', vertical: 'middle' };
        amountCell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      });
      
      // Column H (spacer)
      const spacerCell = worksheet.getCell(`H${row}`);
      spacerCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
    }
  }
  
  /**
   * Format total lines and section headers - Professional Thai style
   */
  private static formatTotalLinesProfessional(worksheet: ExcelJS.Worksheet): void {
    // We'll dynamically detect and format total lines and section headers
    for (let row = 6; row <= 60; row++) {
      const cell = worksheet.getCell(`B${row}`);
      if (cell.value && typeof cell.value === 'string') {
        const value = cell.value.toString().trim();
        
        // Main section headers (สินทรัพย์, หนี้สินและส่วนของผู้ถือหุ้น, ส่วนของผู้ถือหุ้น, ส่วนของผู้เป็นหุ้นส่วน, รายได้, ค่าใช้จ่าย)
        if (value === 'สินทรัพย์' || value === 'หนี้สินและส่วนของผู้ถือหุ้น' || 
            value === 'ส่วนของผู้ถือหุ้น' || value === 'ส่วนของผู้เป็นหุ้นส่วน' ||
            value === 'รายได้' || value === 'ค่าใช้จ่าย') {
          this.formatMainSectionHeader(worksheet, row);
        }
        // Key profit line items (should be bold)
        else if (value === 'กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้' || 
                 value === 'กำไรก่อนภาษีเงินได้' || 
                 value === 'กำไร(ขาดทุน)สุทธิ') {
          this.formatKeyProfitLine(worksheet, row);
        }
        // Sub-section headers (สินทรัพย์หมุนเวียน, สินทรัพย์ไม่หมุนเวียน)
        else if (value.includes('หมุนเวียน') || value.includes('ไม่หมุนเวียน')) {
          this.formatSubSectionHeader(worksheet, row);
        }
        // Total lines (รวม...)
        else if (value.includes('รวม')) {
          this.formatTotalLine(worksheet, row);
        }
      }
    }
  }
  
  /**
   * Format main section headers (สินทรัพย์)
   */
  private static formatMainSectionHeader(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      
      // Special handling for year columns in row 6 (G6 and I6) - keep center alignment
      if (row === 6 && (col === 'G' || col === 'I')) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.numFmt = 'General'; // Ensure General format for year columns
        cell.font.bold = false; // Year columns should not be bold
        console.log(`Preserved center alignment and General format for ${col}${row}`);
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }
  
  /**
   * Format sub-section headers (สินทรัพย์หมุนเวียน)
   */
  private static formatSubSectionHeader(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      if (col === 'B') {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }
  
  /**
   * Format total lines (รวม...)
   */
  private static formatTotalLine(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      
      if (col === 'B') {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }
  
  /**
   * Format key profit line items (should be bold)
   */
  private static formatKeyProfitLine(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      
      if (col === 'B') {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }
  
  /**
   * Set professional row heights using autofit (except row 4)
   */
  private static setRowHeightsProfessional(worksheet: ExcelJS.Worksheet): void {
    console.log('=== SETTING ROW HEIGHTS WITH AUTOFIT ===');
    
    // Only set row 4 as a fixed spacer, let all other rows autofit
    worksheet.getRow(4).height = 15; // Keep row 4 as spacer
    console.log('Row 4 height set to fixed: 15 (spacer)');
    console.log('All other rows will use autofit (no height set)');
    
    console.log('=== ROW HEIGHTS SET COMPLETE (AUTOFIT ENABLED) ===');
  }
  
  /**
   * Special formatting for Statement of Changes in Equity
   */
  private static formatStatementOfChangesInEquitySpecific(worksheet: ExcelJS.Worksheet): void {
    // Format C6, F6, I6 with wrapped text and bottom border
    const columnsToFormat = ['C', 'F', 'I'];
    const row = 6;
    
    columnsToFormat.forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      
      // Set wrapped text
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };
      
      // Set bottom border
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FF000000' } }
      };
      
      // Ensure font formatting
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
    });
    
    console.log('Applied SCE-specific formatting: C6, F6, I6 with wrapped text and bottom border');
  }
  
  /**
   * Clear all empty cells to prevent text cutoff issues
   */
  private static clearEmptyCells(worksheet: ExcelJS.Worksheet): void {
    // Clear all cells that don't have meaningful content - Extended range for Notes_Accounting
    // Cover up to row 100 and column K (11) to be thorough for notes
    for (let row = 1; row <= 100; row++) {
      for (let col = 1; col <= 11; col++) { // A to K columns
        const cell = worksheet.getCell(row, col);
        
        // Check if cell is effectively empty or contains only whitespace
        // Protect columns F(6) and G(7) from having zeros cleared (PPE movement columns)
        const shouldClearZeros = (col !== 6 && col !== 7);
        const isProtectedColumn = (col === 6 || col === 7);
        
        if ((cell.value === undefined || cell.value === null) ||
            (typeof cell.value === 'string' && cell.value.trim() === '') ||
            (typeof cell.value === 'string' && /^\s*$/.test(cell.value)) ||
            (typeof cell.value === 'number' && cell.value === 0 && row > 10 && shouldClearZeros && !isProtectedColumn)) {
          
          
          // Clear the cell value completely
          cell.value = null;
          
          // Also clear any formatting that might cause display issues
          cell.style = {};
        }
      }
    }
    
    console.log('Cleared all empty cells and whitespace-only cells (extended range for Notes_Accounting)');
  }
  
  /**
   * Get border style for all sides
   */
  private static getAllBorders(): Partial<ExcelJS.Borders> {
    const borderStyle: Partial<ExcelJS.Border> = {
      style: 'thin',
      color: { argb: 'FF000000' }
    };
    
    return {
      top: borderStyle,
      bottom: borderStyle,
      left: borderStyle,
      right: borderStyle
    };
  }
  
  /**
   * Format P&L Statement
   */
  static formatProfitLossStatement(worksheet: ExcelJS.Worksheet): void {
    // Use formatting with green background for P&L Statement
    this.formatOtherStatements(worksheet);
  }
  
  /**
   * Format Statement of Changes in Equity
   */
  static formatEquityStatement(worksheet: ExcelJS.Worksheet): void {
    // Use formatting with green background for Equity Statement
    this.formatOtherStatements(worksheet);
  }
  
  /**
   * Clear all blank spaces and empty strings from worksheet
   */
  static clearBlankSpaces(worksheet: ExcelJS.Worksheet): void {
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        // Clear empty strings, spaces, and null values
        if (cell.value === '' || cell.value === ' ' || cell.value === null || 
            (typeof cell.value === 'string' && cell.value.trim() === '')) {
          cell.value = null;
        }
      });
    });
  }

  /**
   * Enhanced blank space clearing specifically for Notes_Accounting
   */
  static clearNotesAccountingBlanks(worksheet: ExcelJS.Worksheet): void {
    console.log('Performing enhanced blank space clearing for Notes_Accounting');
    
    // Step 1: Use the standard clearBlankSpaces
    this.clearBlankSpaces(worksheet);
    
    // Step 2: More aggressive clearing for Notes_Accounting
    for (let row = 1; row <= 150; row++) { // Extended range for notes
      for (let col = 1; col <= 12; col++) { // A to L columns
        const cell = worksheet.getCell(row, col);
        
        // Clear various empty states
        if (cell.value === undefined || 
            cell.value === null ||
            cell.value === '' ||
            (typeof cell.value === 'string' && /^\s*$/.test(cell.value)) ||
            (typeof cell.value === 'string' && cell.value === ' ') ||
            (typeof cell.value === 'string' && cell.value === '  ') ||
            (typeof cell.value === 'string' && cell.value === '\t') ||
            (typeof cell.value === 'string' && cell.value === '\n')) {
          
          // Completely clear the cell
          cell.value = null;
          
          // Reset style to prevent formatting artifacts
          cell.style = {
            font: { name: this.THAI_FONT_NAME, size: 14 },
            alignment: { horizontal: 'left', vertical: 'middle' }
          };
        }
      }
    }
    
    console.log('Enhanced Notes_Accounting blank space clearing completed');
  }

  /**
   * Format Notes to Financial Statements (Note_Policy) with specific column widths
   */
  static formatNotesToFinancialStatements(worksheet: ExcelJS.Worksheet): void {
    // Clear blank spaces first
    this.clearBlankSpaces(worksheet);
    // Apply Note_Policy specific formatting
    this.formatNotesPolicy(worksheet);
  }

  /**
   * Format Note_Policy with specific column widths and special formatting
   */
  static formatNotesPolicy(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying Note_Policy specific formatting');
    
    // Clear blank spaces first
    this.clearBlankSpaces(worksheet);
    
    // Set TH Sarabun New font for all cells
    this.setWorksheetDefaultFont(worksheet);
    
    // Set specific column widths for Note_Policy
    worksheet.columns = [
      { width: 5 },     // A = 4.3
      { width: 5.3 },   // B = 4.6
      { width: 8.33 },  // C = 7.6
      { width: 10 },    // D = 9.3
      { width: 14.11 }, // E = 13.4
      { width: 10 },    // F = 9.3
      { width: 16.44 }, // G = 15.7
      { width: 13 },    // H = 13
      { width: 2.33 }   // I = 1.6
    ];
    
    // Apply general formatting sections (without column headers that add หมายเหตุ)
    this.formatCompanyHeaderNoBackground(worksheet);
    // Skip formatColumnHeadersProfessional to avoid adding หมายเหตุ to F5
    this.formatAccountLinesProfessional(worksheet);
    this.formatTotalLinesProfessional(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    // Apply special Notes_Policy formatting 
    this.formatNotesPolicySpecial(worksheet);
    
    console.log('Note_Policy formatting completed');
  }

  /**
   * Apply special formatting for Notes_Policy
   */
  private static formatNotesPolicySpecial(worksheet: ExcelJS.Worksheet): void {
    // 1. Delete หมายเหตุ in F5 - clear F5 specifically
    const f5Cell = worksheet.getCell('F5');
    f5Cell.value = '';
    
    // 2. B4 and B5 set to TH Sarabun New font
    const b4Cell = worksheet.getCell('B4');
    const b5Cell = worksheet.getCell('B5');
    
    [b4Cell, b5Cell].forEach(cell => {
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        color: { argb: 'FF000000' }
      };
    });
    
    // 3. C6:C8 align center
    ['C6', 'C7', 'C8'].forEach(cellRef => {
      const cell = worksheet.getCell(cellRef);
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    
    // 4. Bold D6 to D8
    ['D6', 'D7', 'D8'].forEach(cellRef => {
      const cell = worksheet.getCell(cellRef);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
    });
    
    // 5. Merge D6:I6, D7:I7, D8:I8 and align left (merge and across) with wrapped text
    ['D6:I6', 'D7:I7', 'D8:I8'].forEach((range) => {
      worksheet.mergeCells(range);
      const firstCell = worksheet.getCell(range.split(':')[0]);
      firstCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      
      // Use autofit for rows 6, 7, 8 - no fixed height
      // Let Excel automatically adjust height based on content
    });
    
    // 6. Merge C11:I11 for Note 2 (ฐานะการดำเนินงานของบริษัท) content with wrapped text and align left
    worksheet.mergeCells('C11:I11');
    const c11Cell = worksheet.getCell('C11');
    c11Cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    
    // 7. B17:B18 (4.1 and 4.2) => Center align and middle align
    ['B17', 'B18'].forEach(cellRef => {
      const cell = worksheet.getCell(cellRef);
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    
    // 8. Merge C14:I14, C15:I15 for เกณฑ์การจัดทำงบการเงินของกิจการ section with wrapped text and align left
    ['C14:I14', 'C15:I15'].forEach(range => {
      worksheet.mergeCells(range);
      const firstCell = worksheet.getCell(range.split(':')[0]);
      firstCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    });

    // 8. Merge B to I for policy section rows: 20,22,24,26,28,30,32,34,36,38,40,42 with wrapped text and align left
    // Use autofit for rows 22 and 24 - no fixed height
    const policyMergeRows = [20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42];
    policyMergeRows.forEach(rowNum => {
      const range = `B${rowNum}:I${rowNum}`;
      worksheet.mergeCells(range);
      const firstCell = worksheet.getCell(`B${rowNum}`);
      firstCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    });

    // 9. Merge C44:I44, C46:I46, C48:I48 with wrapped text
    // Note: Removed rows 14,15,20,22,26 (no merge/wrap - autofit), removed rows 34,36,38,40,42 (conflict with B-to-I merging in step 8)
    const mergeRows = [44, 46, 48];
    mergeRows.forEach(rowNum => {
      const range = `C${rowNum}:I${rowNum}`;
      worksheet.mergeCells(range);
      const firstCell = worksheet.getCell(`C${rowNum}`);
      firstCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    });
    
    // 11. Set specific row heights for merged content rows based on text content
    // Rows 6,7,8,22,24 use autofit (no fixed height)
    // Rows 17,18,23,25,29 use autofit and no merging/wrapping
    const rowHeights = {
      11: 42,  // Note 2 content (ฐานะการดำเนินงานของบริษัท) - 2 lines height (increased by 10%)
      14: 49,  // เกณฑ์การจัดทำงบการเงินของกิจการ section text (increased by ~10% for 14pt font)
      15: 75,  // Longer description text (increased by ~10% for 14pt font)
      // 17: removed - autofit, no merge, no wrap
      // 18: removed - autofit, no merge, no wrap
      20: 75,  // Policy section content (increased by ~10% for 14pt font)
      // 22: removed - autofit
      // 23: removed - autofit, no merge, no wrap
      // 24: removed - autofit
      // 25: removed - autofit, no merge, no wrap
      26: 75,  // Policy section content (increased by ~10% for 14pt font)
      28: 75,  // Policy section content (increased by ~10% for 14pt font)
      // 29: removed - autofit, no merge, no wrap
      30: 75,  // Policy section content (increased by ~10% for 14pt font)
      32: 75,  // Policy section content (increased by ~10% for 14pt font)
      34: 49,  // 2 lines of text (increased by ~10% for 14pt font)
      36: 75,  // 3 lines of text (increased by ~10% for 14pt font)
      38: 75,  // 3 lines of text (increased by ~10% for 14pt font)
      40: 49,  // 2 lines of text (increased by ~10% for 14pt font)
      42: 75,  // 3 lines of text (increased by ~10% for 14pt font)
      44: 75,  // 3 lines of text (increased by ~10% for 14pt font)
      46: 75,  // 3 lines of text (increased by ~10% for 14pt font)
      48: 49   // 2 lines of text (increased by ~10% for 14pt font)
    };
    
    Object.entries(rowHeights).forEach(([rowNum, height]) => {
      worksheet.getRow(parseInt(rowNum)).height = height;
    });
  }
  
  /**
   * Format notes with specific note-by-note formatting using row trackers
   */
  static formatNotesWithSpecificFormatting(worksheet: ExcelJS.Worksheet, formatters: any[]): void {
    console.log('Applying specific note-by-note formatting with row trackers');
    
    // STEP 1: Enhanced blank space clearing for Notes_Accounting
    this.clearNotesAccountingBlanks(worksheet);
    
    // STEP 2: Set basic worksheet formatting
    this.setWorksheetDefaultFont(worksheet);
    this.setAccountingNotesColumnWidths(worksheet);
    this.formatCompanyHeaderNoBackground(worksheet);
    
    // STEP 3: Apply specific formatting for each note
    formatters.forEach(formatter => {
      console.log(`Formatting ${formatter.type} note:`, formatter.tracker);
      switch (formatter.type) {
        case 'cash':
          this.formatCashNote(worksheet, formatter.tracker);
          break;
        case 'receivables':
          this.formatReceivablesNote(worksheet, formatter.tracker);
          break;
        case 'payables':
          this.formatPayablesNote(worksheet, formatter.tracker);
          break;
        case 'ppe':
          this.formatPPENote(worksheet, formatter.tracker);
          break;
        case 'shortTermLoans':
          this.formatShortTermLoansNote(worksheet, formatter.tracker);
          break;
        default:
          this.formatGeneralNote(worksheet, formatter.tracker);
      }
    });
    
    // STEP 4: Final cleanup - clear any remaining empty cells but preserve PPE movement columns
    this.clearEmptyCells(worksheet);
    
    console.log(`Specific note formatting completed for ${formatters.length} notes with enhanced blank cell cleanup`);
  }

  /**
   * Set column widths for accounting notes
   */
  private static setAccountingNotesColumnWidths(worksheet: ExcelJS.Worksheet): void {
    worksheet.columns = [
      { width: 4 },   // A - note number
      { width: 4 },   // B  
      { width: 25 },  // C - account names
      { width: 12 },  // D
      { width: 2 },   // E
      { width: 12 },  // F
      { width: 12 },  // G - current year amounts
      { width: 2 },   // H
      { width: 12 }   // I - previous year amounts
    ];
  }

  /**
   * Format Cash Note specifically
   */
  private static formatCashNote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting Cash Note, rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // 1. Format Note Header (bold in column B, "หน่วย:บาท" bold in column I)
    tracker.headerRows.forEach((row: number) => {
      const noteNumberCell = worksheet.getCell(`A${row}`);
      const noteTitleCell = worksheet.getCell(`B${row}`);
      const unitCell = worksheet.getCell(`I${row}`);
      
      // Note number
      noteNumberCell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      noteNumberCell.alignment = { horizontal: 'left', vertical: 'middle' };
      
      // Note title - BOLD
      noteTitleCell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      noteTitleCell.alignment = { horizontal: 'left', vertical: 'middle' };
      
      // "หน่วย:บาท" - BOLD
      unitCell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      unitCell.alignment = { horizontal: 'right', vertical: 'middle' };
    });

    // 2. Format Year Headers (center, underline, general format)
    tracker.yearHeaderRows.forEach((row: number) => {
      const currentYearCell = worksheet.getCell(`G${row}`);
      const previousYearCell = worksheet.getCell(`I${row}`);
      
      [currentYearCell, previousYearCell].forEach(cell => {
        if (cell.value) {
          cell.font = { name: this.THAI_FONT_NAME, size: 14, underline: true };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.numFmt = '@'; // General format, not number format
        }
      });
    });

    // 3. Format Detail Rows (normal text in column C, amounts right-aligned)
    tracker.detailRows.forEach((row: number) => {
      const detailNameCell = worksheet.getCell(`C${row}`);
      const currentAmountCell = worksheet.getCell(`G${row}`);
      const previousAmountCell = worksheet.getCell(`I${row}`);
      
      // Detail name - normal
      detailNameCell.font = { name: this.THAI_FONT_NAME, size: 14 };
      detailNameCell.alignment = { horizontal: 'left', vertical: 'middle' };
      
      // Amounts - right aligned with number format
      [currentAmountCell, previousAmountCell].forEach(cell => {
        if (cell.value) {
          cell.font = { name: this.THAI_FONT_NAME, size: 14 };
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
        }
      });
    });

    // 4. Format Total Row ("รวม" bold in column C, amounts not bold)
    tracker.totalRows.forEach((row: number) => {
      const totalTextCell = worksheet.getCell(`C${row}`);
      const currentTotalCell = worksheet.getCell(`G${row}`);
      const previousTotalCell = worksheet.getCell(`I${row}`);
      
      // "รวม" text - BOLD
      totalTextCell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      totalTextCell.alignment = { horizontal: 'left', vertical: 'middle' };
      
      // Total amounts - NOT bold
      [currentTotalCell, previousTotalCell].forEach(cell => {
        if (cell.value) {
          cell.font = { name: this.THAI_FONT_NAME, size: 14, bold: false };
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
        }
      });
    });
  }

  /**
   * Format Receivables Note specifically
   */
  private static formatReceivablesNote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting Receivables Note, rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // Same structure as Cash Note
    this.formatCashNote(worksheet, tracker);
  }

  /**
   * Format Payables Note specifically
   */
  private static formatPayablesNote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting Payables Note, rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // Same structure as Cash Note
    this.formatCashNote(worksheet, tracker);
  }

  /**
   * Format PPE Note specifically
   */
  /**
   * Format PPE Note specifically - Enhanced for complex structure
   */
  private static formatPPENote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting PPE Note (Enhanced), rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // PPE notes have complex structure with section headers, so use enhanced formatting
    
    // 1. Header Rows (Note headers, section headers like "ราคาทุนเดิม", "ค่าเสื่อมราคาสะสม")
    tracker.headerRows.forEach((row: number) => {
      const rowObj = worksheet.getRow(row);
      rowObj.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      
      // Also apply to individual cells for safety
      for (let col = 1; col <= 9; col++) {
        const cell = worksheet.getCell(row, col);
        if (cell.value) {
          cell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
        }
      }
    });

    // 2. Unit Rows (หน่วย:บาท)
    tracker.unitRows.forEach((row: number) => {
      const cell = worksheet.getCell(row, 9); // Column I
      if (cell.value) {
        cell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      }
    });

    // 3. Year Header Rows (column headers)
    tracker.yearHeaderRows.forEach((row: number) => {
      const rowObj = worksheet.getRow(row);
      rowObj.font = { name: this.THAI_FONT_NAME, size: 14, bold: false };
      rowObj.alignment = { horizontal: 'center', vertical: 'middle' };
      
      // Apply underline to year headers
      for (let col = 1; col <= 9; col++) {
        const cell = worksheet.getCell(row, col);
        if (cell.value) {
          cell.font = { 
            name: this.THAI_FONT_NAME, 
            size: 14, 
            bold: false,
            underline: true 
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }
    });

    // 4. Detail Rows (individual accounts)
    tracker.detailRows.forEach((row: number) => {
      // Apply number formatting to amount columns
      for (let col = 4; col <= 9; col++) { // Columns D through I
        const cell = worksheet.getCell(row, col);
        if (typeof cell.value === 'number') { // Remove the && cell.value condition to include zeros
          cell.numFmt = '#,##0.00';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.font = { name: this.THAI_FONT_NAME, size: 14, bold: false };
        }
      }
    });

    // 5. Total Rows (รวม, มูลค่าสุทธิ, ค่าเสื่อมราคา)
    tracker.totalRows.forEach((row: number) => {
      // Apply bold to text in column C
      const textCell = worksheet.getCell(row, 3);
      if (textCell.value) {
        textCell.font = { name: this.THAI_FONT_NAME, size: 14, bold: true };
      }
      
      // Apply number formatting to amount columns but keep them normal weight
      for (let col = 4; col <= 9; col++) { // Columns D through I
        const cell = worksheet.getCell(row, col);
        if (typeof cell.value === 'number' || (typeof cell.value === 'object' && cell.value && 'formula' in cell.value)) {
          cell.numFmt = '#,##0.00';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.font = { name: this.THAI_FONT_NAME, size: 14, bold: false };
        }
      }
    });

    console.log(`PPE Note formatting completed: ${tracker.headerRows.length} headers, ${tracker.detailRows.length} details, ${tracker.totalRows.length} totals`);
  }

  /**
   * Format Short Term Loans Note specifically
   */
  private static formatShortTermLoansNote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting Short Term Loans Note, rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // Use same structure as Cash Note - simple note format
    this.formatCashNote(worksheet, tracker);
  }

  /**
   * Format General Note specifically
   */
  private static formatGeneralNote(worksheet: ExcelJS.Worksheet, tracker: any): void {
    console.log('Formatting General Note, rows:', tracker.noteStartRow, 'to', tracker.currentRow - 1);
    
    // Same structure as Cash Note
    this.formatCashNote(worksheet, tracker);
  }

  /**
   * Format notes with no green background - clean professional look (Notes_Accounting)
   */
  static formatNotesWithoutBackground(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying Notes_Accounting specific formatting (fallback method)');
    
    // Enhanced blank space clearing first
    this.clearNotesAccountingBlanks(worksheet);
    
    // Set TH Sarabun New font for all cells
    this.setWorksheetDefaultFont(worksheet);
    
    // Set specific column widths for Notes_Accounting
    worksheet.columns = [
      { width: 4 },   // A
      { width: 4 },   // B
      { width: 25 },  // C
      { width: 12 },  // D
      { width: 2 },   // E
      { width: 12 },  // F
      { width: 12 },  // G
      { width: 2 },   // H
      { width: 12 }   // I
    ];
    
    // Apply formatting sections without green background
    this.formatCompanyHeaderNoBackground(worksheet);
    this.formatColumnHeadersAccountingNotes(worksheet);
    this.formatAccountLinesAccountingNotes(worksheet);
    this.formatTotalLinesAccountingNotes(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    console.log('Notes_Accounting formatting completed');
  }

  /**
   * Format column headers for accounting notes - No หมายเหตุ in F5
   */
  private static formatColumnHeadersAccountingNotes(worksheet: ExcelJS.Worksheet): void {
    // Row 5: Clear all header cells (no หมายเหตุ for accounting notes)
    ['D', 'E', 'F'].forEach(col => {
      const headerCell = worksheet.getCell(`${col}5`);
      headerCell.value = '';
    });
  }
  
  /**
   * Format account detail lines for accounting notes - Current amount in G, Previous amount in I
   */
  private static formatAccountLinesAccountingNotes(worksheet: ExcelJS.Worksheet): void {
    // Apply base formatting to all data rows (6-100) - Extended range to cover all notes
    for (let row = 6; row <= 100; row++) {
      // Account names (Columns B through E)
      ['B', 'C', 'D', 'E'].forEach(col => {
        const nameCell = worksheet.getCell(`${col}${row}`);
        nameCell.font = { 
          name: this.THAI_FONT_NAME, 
          size: 14,
          color: { argb: 'FF000000' }
        };
        nameCell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
      
      // Note references (Column F)
      const noteCell = worksheet.getCell(`F${row}`);
      noteCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
      noteCell.alignment = { horizontal: 'center', vertical: 'middle' };
      
      // Current period amounts (Column G - moved from E)
      const currentAmountCell = worksheet.getCell(`G${row}`);
      currentAmountCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
      currentAmountCell.alignment = { horizontal: 'right', vertical: 'middle' };
      currentAmountCell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      
      // Previous period amounts (Column I - moved from F)
      const previousAmountCell = worksheet.getCell(`I${row}`);
      previousAmountCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
      previousAmountCell.alignment = { horizontal: 'right', vertical: 'middle' };
      previousAmountCell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      
      // Column H (spacer)
      const spacerCell = worksheet.getCell(`H${row}`);
      spacerCell.font = { 
        name: this.THAI_FONT_NAME, 
        size: 14,
        color: { argb: 'FF000000' }
      };
    }
  }
  
  /**
   * Format total lines for accounting notes - Current amount in G, Previous amount in I
   */
  private static formatTotalLinesAccountingNotes(worksheet: ExcelJS.Worksheet): void {
    // We'll dynamically detect and format total lines, section headers, note headers, and year headers
    // Extended range to cover all possible note rows (up to 100 rows to be safe)
    for (let row = 6; row <= 100; row++) {
      const cellA = worksheet.getCell(`A${row}`);
      const cellB = worksheet.getCell(`B${row}`);
      const cellG = worksheet.getCell(`G${row}`);
      const cellI = worksheet.getCell(`I${row}`);
      
      // Check for year headers (strings like "2567", "2566")
      if (cellG.value && typeof cellG.value === 'string' && /^\d{4}$/.test(cellG.value.trim())) {
        this.formatYearHeaderAccountingNotes(worksheet, row, 'G');
      }
      if (cellI.value && typeof cellI.value === 'string' && /^\d{4}$/.test(cellI.value.trim())) {
        this.formatYearHeaderAccountingNotes(worksheet, row, 'I');
      }
      
      // Check for "หน่วย:บาท" in column I (should be bold)
      if (cellI.value && typeof cellI.value === 'string' && cellI.value.includes('หน่วย:บาท')) {
        cellI.font = {
          name: this.THAI_FONT_NAME,
          size: 14,
          bold: true,
          color: { argb: 'FF000000' }
        };
        cellI.alignment = { horizontal: 'right', vertical: 'middle' };
      }
      
      // Check for note headers (note number in column A, note title in column B)
      const cellAValue = cellA.value && typeof cellA.value === 'string' ? cellA.value.toString().trim() : '';
      const cellBValue = cellB.value && typeof cellB.value === 'string' ? cellB.value.toString().trim() : '';
      
      // Note headers: number in A and descriptive text in B
      if (/^\d+$/.test(cellAValue) && cellBValue && this.isNoteHeader(cellBValue)) {
        this.formatNoteHeaderAccountingNotes(worksheet, row);
      }
      // Also check for note headers that might be only in column B
      else if (cellBValue && this.isNoteHeader(cellBValue)) {
        this.formatNoteHeaderAccountingNotes(worksheet, row);
      }
      // Main section headers (สินทรัพย์, หนี้สินและส่วนของผู้ถือหุ้น, ส่วนของผู้ถือหุ้น, ส่วนของผู้เป็นหุ้นส่วน)
      else if (cellBValue === 'สินทรัพย์' || cellBValue === 'หนี้สินและส่วนของผู้ถือหุ้น' || 
          cellBValue === 'ส่วนของผู้ถือหุ้น' || cellBValue === 'ส่วนของผู้เป็นหุ้นส่วน') {
        this.formatMainSectionHeaderAccountingNotes(worksheet, row);
      }
      // Sub-section headers (สินทรัพย์หมุนเวียน, สินทรัพย์ไม่หมุนเวียน)
      else if (cellBValue.includes('หมุนเวียน') || cellBValue.includes('ไม่หมุนเวียน')) {
        this.formatSubSectionHeaderAccountingNotes(worksheet, row);
      }
      // Total lines (รวม...)
      else if (cellBValue.includes('รวม')) {
        this.formatTotalLineAccountingNotes(worksheet, row);
      }
    }
  }
  
  /**
   * Format main section headers for accounting notes (สินทรัพย์)
   */
  private static formatMainSectionHeaderAccountingNotes(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
  }
  
  /**
   * Format sub-section headers for accounting notes (สินทรัพย์หมุนเวียน)
   */
  private static formatSubSectionHeaderAccountingNotes(worksheet: ExcelJS.Worksheet, row: number): void {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: true,
        color: { argb: 'FF000000' }
      };
      if (col === 'B') {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }
  
  /**
   * Format total lines for accounting notes (รวม...)
   */
  private static formatTotalLineAccountingNotes(worksheet: ExcelJS.Worksheet, row: number): void {
    // Make only the text "รวม" bold (column B), not the amounts
    const textCell = worksheet.getCell(`B${row}`);
    textCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    textCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Format other columns normally (not bold)
    ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: false, // Keep amounts not bold
        color: { argb: 'FF000000' }
      };
      
      if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }

  /**
   * Check if a value is a note header based on common patterns
   */
  private static isNoteHeader(value: string): boolean {
    // More comprehensive patterns - check if the value starts with digit (note number)
    const startsWithNoteNumber = /^\d+\s/.test(value.trim());
    
    // Specific note header patterns (including variations found in the code)
    const noteHeaderPatterns = [
      'เงินสดและรายการเทียบเท่าเงินสด',
      'ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น',
      'ลูกหนี้การค้าและลูกหนี้อื่น',
      'เงินกู้ยืมระยะสั้น',
      'ที่ดิน อาคารและอุปกรณ์',
      'สินทรัพย์อื่น',
      'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน',
      'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้น',
      'เจ้าหนี้การค้าและเจ้าหนี้อื่น',
      'เงินกู้ยืมระยะยาวจากสถาบันการเงิน',
      'เงินกู้ยืมระยะยาวอื่น',
      'เงินกู้ยืมจากกิจการที่เกี่ยวข้องกัน',
      'รายได้อื่น',
      'ค่าใช้จ่ายตามลักษณะ',
      'การอนุมัติงบการเงิน'
    ];
    
    // Check if it's a note header by pattern or starts with note number
    const matchesPattern = noteHeaderPatterns.some(pattern => value.includes(pattern));
    const isNumberedHeader = startsWithNoteNumber && (
      value.includes('เงินสด') || 
      value.includes('ลูกหนี้') || 
      value.includes('เจ้าหนี้') || 
      value.includes('เงินกู้') || 
      value.includes('เงินเบิก') ||
      value.includes('ที่ดิน') || 
      value.includes('สินทรัพย์') || 
      value.includes('รายได้') ||
      value.includes('ค่าใช้จ่าย') ||
      value.includes('การอนุมัติ')
    );
    
    return matchesPattern || isNumberedHeader;
  }

  /**
   * Format note headers for accounting notes (เงินสด, ลูกหนี้การค้า, etc.)
   */
  private static formatNoteHeaderAccountingNotes(worksheet: ExcelJS.Worksheet, row: number): void {
    // Make only the note header text (column B) bold, not the amounts
    const headerCell = worksheet.getCell(`B${row}`);
    headerCell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: true,
      color: { argb: 'FF000000' }
    };
    headerCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Format other columns normally (not bold)
    ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.font = {
        name: this.THAI_FONT_NAME,
        size: 14,
        bold: false, // Keep amounts and other content not bold
        color: { argb: 'FF000000' }
      };
      
      if (col === 'F') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (col === 'G' || col === 'I') {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }

  /**
   * Format year headers for accounting notes (2567, 2566)
   */
  private static formatYearHeaderAccountingNotes(worksheet: ExcelJS.Worksheet, row: number, col: string): void {
    const cell = worksheet.getCell(`${col}${row}`);
    cell.font = {
      name: this.THAI_FONT_NAME,
      size: 14,
      bold: false,
      underline: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.numFmt = '@'; // Text format to display as general format, not number format
  }

  /**
   * Format Detail Notes (DT1 and DT2) with specific column widths
   */
  static formatDetailNotes(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying Detail Notes (DT1 & DT2) specific formatting');
    
    // Clear blank spaces first
    this.clearBlankSpaces(worksheet);
    
    // Set TH Sarabun New font for all cells
    this.setWorksheetDefaultFont(worksheet);
    
    // Set specific column widths for Detail Notes (optimized for both DT1 and DT2)
    worksheet.columns = [
      { width: 13 },   // A
      { width: 8 },    // B
      { width: 3.5 },  // C
      { width: 3.5 },  // D
      { width: 3.5 },  // E
      { width: 3.5 },  // F
      { width: 18 },   // G - Selling expenses
      { width: 20 },   // H - Administrative expenses  
      { width: 18 }    // I - Other expenses / DT1 amounts
    ];
    
    // Apply formatting sections without green background
    this.formatCompanyHeaderNoBackground(worksheet);
    this.formatColumnHeadersAccountingNotes(worksheet); // Use accounting notes headers (no หมายเหตุ in F5)
    this.formatDetailNotesAccountLines(worksheet);
    this.formatDetailTwoMerging(worksheet); // Add specific merging for Detail 2
    this.formatTotalLinesProfessional(worksheet);
    this.setRowHeightsProfessional(worksheet);
    
    console.log('Detail Notes formatting completed');
  }

  /**
   * Format account detail lines specifically for Detail Notes - amounts in column I
   */
  private static formatDetailNotesAccountLines(worksheet: ExcelJS.Worksheet): void {
    // Apply base formatting to all data rows (6-50)
    for (let row = 6; row <= 50; row++) {
      // Account names (Columns A through F)
      ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
        const nameCell = worksheet.getCell(`${col}${row}`);
        nameCell.font = { 
          name: this.THAI_FONT_NAME, 
          size: 14,
          color: { argb: 'FF000000' }
        };
        nameCell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
      
      // Amount columns (G, H, I) - all formatted for numbers
      ['G', 'H', 'I'].forEach(col => {
        const amountCell = worksheet.getCell(`${col}${row}`);
        amountCell.font = { 
          name: this.THAI_FONT_NAME, 
          size: 14,
          color: { argb: 'FF000000' }
        };
        amountCell.alignment = { horizontal: 'right', vertical: 'middle' };
        amountCell.numFmt = '#,##0.00_);[Red](#,##0.00)';
      });
    }
  }
  
  /**
   * Format Detail 2 specific merging - merge columns A-F for expense account rows
   */
  private static formatDetailTwoMerging(worksheet: ExcelJS.Worksheet): void {
    // Find Detail 2 section by looking for "รายละเอียดประกอบที่ 2"
    let detailTwoStartRow = -1;
    let detailTwoHeaderRow = -1;
    
    // Find the start of Detail 2
    for (let row = 1; row <= 100; row++) {
      const cellValue = worksheet.getCell(row, 1).value;
      if (cellValue && cellValue.toString().includes('รายละเอียดประกอบที่ 2')) {
        detailTwoStartRow = row;
        break;
      }
    }
    
    // Find the header row (ค่าใช้จ่ายในการขายและบริหาร)
    if (detailTwoStartRow > 0) {
      for (let row = detailTwoStartRow; row <= detailTwoStartRow + 5; row++) {
        const cellValue = worksheet.getCell(row, 1).value;
        if (cellValue && cellValue.toString().includes('ค่าใช้จ่ายในการขายและบริหาร')) {
          detailTwoHeaderRow = row;
          break;
        }
      }
    }
    
    if (detailTwoHeaderRow > 0) {
      // Merge the header row A-F
      worksheet.mergeCells(`A${detailTwoHeaderRow}:F${detailTwoHeaderRow}`);
      const headerCell = worksheet.getCell(`A${detailTwoHeaderRow}`);
      headerCell.alignment = { horizontal: 'left', vertical: 'middle' };
      
      // Add borders to header row
      for (let col = 1; col <= 9; col++) {
        const cell = worksheet.getCell(detailTwoHeaderRow, col);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
      
      // Process expense account rows (data rows after header)
      const dataStartRow = detailTwoHeaderRow + 1;
      let lastDataRow = dataStartRow;
      
      for (let row = dataStartRow; row <= dataStartRow + 50; row++) {
        const cellA = worksheet.getCell(row, 1);
        const cellG = worksheet.getCell(row, 7);
        const cellH = worksheet.getCell(row, 8);
        const cellI = worksheet.getCell(row, 9);
        
        // Check if this row has expense data (non-zero values in G, H, or I, or is a total/financial cost row)
        const hasExpenseData = (cellG.value && cellG.value !== 0) || 
                              (cellH.value && cellH.value !== 0) || 
                              (cellI.value && cellI.value !== 0) ||
                              (cellA.value && (cellA.value.toString().includes('รวม') || 
                                             cellA.value.toString().includes('ค่าใช้จ่ายต้นทุนทางการเงิน')));
        
        // Check if row has account name in column A
        const hasAccountName = cellA.value && cellA.value.toString().trim() !== '';
        
        if (hasAccountName && hasExpenseData) {
          // Merge columns A-F for this expense account row
          try {
            worksheet.mergeCells(`A${row}:F${row}`);
            cellA.alignment = { horizontal: 'left', vertical: 'middle' };
          } catch (error) {
            // Skip if merge fails (already merged or other issue)
            console.log(`Merge failed for row ${row}: ${error}`);
          }
          
          // Add borders to all columns for this data row (A-I)
          for (let col = 1; col <= 9; col++) {
            const cell = worksheet.getCell(row, col);
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
          }
          
          lastDataRow = row;
        }
        
        // Stop processing if we reach an empty section (end of Detail 2)
        if (!hasAccountName && !hasExpenseData) {
          // Check next few rows to see if we've reached the end
          let emptyRowCount = 0;
          for (let checkRow = row; checkRow <= row + 3; checkRow++) {
            const checkCell = worksheet.getCell(checkRow, 1);
            if (!checkCell.value || checkCell.value.toString().trim() === '') {
              emptyRowCount++;
            }
          }
          if (emptyRowCount >= 3) break; // End of Detail 2 section
        }
      }
      
      // Add bottom border to complete the table
      if (lastDataRow > dataStartRow) {
        for (let col = 1; col <= 9; col++) {
          const cell = worksheet.getCell(lastDataRow, col);
          cell.border = {
            ...cell.border,
            bottom: { style: 'medium' }
          };
        }
      }
    }
    
    console.log('Detail 2 cell merging completed');
  }
  
  /**
   * Save workbook as buffer for download
   */
  static async saveWorkbook(workbook: ExcelJS.Workbook): Promise<Buffer> {
    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }
}
