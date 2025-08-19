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
        
        if (typeof cellValue === 'object' && cellValue.f) {
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
            // It's actual text content
            cell.value = cellValue || '';
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
    // Set default font for the entire worksheet
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (!cell.font) {
          cell.font = {
            name: this.THAI_FONT_NAME,
            size: 14,
            color: { argb: 'FF000000' }
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
    // VBA: A=30, B=2, C=14, D=2, E=2, F=14, G=2, H=2, I=14
    const vbaWidths = [30, 2, 14, 2, 2, 14, 2, 2, 14];
    console.log('SCE VBA Target Widths:', vbaWidths);
    
    // Apply the exact VBA widths without font adjustment for SCE format
    worksheet.columns = [
      { width: vbaWidths[0] },    // A = 30
      { width: vbaWidths[1] },    // B = 2  
      { width: vbaWidths[2] },    // C = 14
      { width: vbaWidths[3] },    // D = 2
      { width: vbaWidths[4] },    // E = 2
      { width: vbaWidths[5] },    // F = 14
      { width: vbaWidths[6] },    // G = 2
      { width: vbaWidths[7] },    // H = 2
      { width: vbaWidths[8] }     // I = 14
    ];
    
    console.log('Applied SCE Column Widths:', vbaWidths);
    
    // Debug width calculations with SCE targets
    this.debugWidthCalculationsSCE(worksheet);
    
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
    
    console.log('Professional Thai formatting completed (SCE format with no green background)');
  }

  /**
   * Format other financial statements (with green background)
   */
  static formatOtherStatements(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying professional Thai formatting (with green background)');
    
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
    for (let row = 6; row <= 40; row++) {
      const cell = worksheet.getCell(`B${row}`);
      if (cell.value && typeof cell.value === 'string') {
        const value = cell.value.toString().trim();
        
        // Main section headers (สินทรัพย์, หนี้สินและส่วนของผู้ถือหุ้น, ส่วนของผู้ถือหุ้น, ส่วนของผู้เป็นหุ้นส่วน)
        if (value === 'สินทรัพย์' || value === 'หนี้สินและส่วนของผู้ถือหุ้น' || 
            value === 'ส่วนของผู้ถือหุ้น' || value === 'ส่วนของผู้เป็นหุ้นส่วน') {
          this.formatMainSectionHeader(worksheet, row);
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
   * Clear all empty cells to prevent text cutoff issues
   */
  private static clearEmptyCells(worksheet: ExcelJS.Worksheet): void {
    // Clear all cells that don't have meaningful content (from A1 to I50 to be thorough)
    for (let row = 1; row <= 50; row++) {
      for (let col = 1; col <= 9; col++) { // A to I columns
        const cell = worksheet.getCell(row, col);
        
        // Check if cell is effectively empty
        if (!cell.value || 
            (typeof cell.value === 'string' && cell.value.trim() === '') ||
            (typeof cell.value === 'number' && cell.value === 0)) {
          
          // Clear the cell value only
          cell.value = null;
        }
      }
    }
    
    console.log('Cleared all empty cells to prevent text cutoff');
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
        size: 12,
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
   * Format notes with no green background - clean professional look (Notes_Accounting)
   */
  static formatNotesWithoutBackground(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying Notes_Accounting specific formatting');
    
    // Clear blank spaces first
    this.clearBlankSpaces(worksheet);
    
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
    // Apply base formatting to all data rows (6-40)
    for (let row = 6; row <= 40; row++) {
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
    // We'll dynamically detect and format total lines and section headers
    for (let row = 6; row <= 40; row++) {
      const cell = worksheet.getCell(`B${row}`);
      if (cell.value && typeof cell.value === 'string') {
        const value = cell.value.toString().trim();
        
        // Main section headers (สินทรัพย์, หนี้สินและส่วนของผู้ถือหุ้น, ส่วนของผู้ถือหุ้น, ส่วนของผู้เป็นหุ้นส่วน)
        if (value === 'สินทรัพย์' || value === 'หนี้สินและส่วนของผู้ถือหุ้น' || 
            value === 'ส่วนของผู้ถือหุ้น' || value === 'ส่วนของผู้เป็นหุ้นส่วน') {
          this.formatMainSectionHeaderAccountingNotes(worksheet, row);
        }
        // Sub-section headers (สินทรัพย์หมุนเวียน, สินทรัพย์ไม่หมุนเวียน)
        else if (value.includes('หมุนเวียน') || value.includes('ไม่หมุนเวียน')) {
          this.formatSubSectionHeaderAccountingNotes(worksheet, row);
        }
        // Total lines (รวม...)
        else if (value.includes('รวม')) {
          this.formatTotalLineAccountingNotes(worksheet, row);
        }
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
   * Format Detail Notes (DT1 and DT2) with specific column widths
   */
  static formatDetailNotes(worksheet: ExcelJS.Worksheet): void {
    console.log('Applying Detail Notes (DT1 & DT2) specific formatting');
    
    // Clear blank spaces first
    this.clearBlankSpaces(worksheet);
    
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
    this.formatColumnHeadersProfessional(worksheet);
    this.formatDetailNotesAccountLines(worksheet);
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
   * Save workbook as buffer for download
   */
  static async saveWorkbook(workbook: ExcelJS.Workbook): Promise<Buffer> {
    return await workbook.xlsx.writeBuffer() as Buffer;
  }
}
