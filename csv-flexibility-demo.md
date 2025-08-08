# Flexible CSV Processor - Usage Examples

The enhanced `CSVProcessor` now supports multiple CSV formats with automatic column detection and flexible parsing options.

## Key Features

### 1. **Automatic Column Detection**
The processor can now handle different header names:

```typescript
// Original format
"ชื่อบัญชี,รหัสบัญชี,ยอดยกมาต้นงวด,ยอดยกมางวดนี้,เดบิต,เครดิต"

// Shortened Thai names
"ชื่อ,รหัส,ยอดต้นงวด,ยอดงวดนี้,เดบิต,เครดิต"

// English headers
"Account Name,Account Code,Previous Balance,Current Balance,Debit,Credit"

// Mixed format
"ชื่อ,Code,ยอดต้นงวด,Current,DR,CR"
```

### 2. **Multiple Delimiter Support**
Auto-detects common CSV delimiters:
- Comma (`,`) - default
- Semicolon (`;`) 
- Tab (`\t`)
- Pipe (`|`)

### 3. **Flexible Parsing Options**

```typescript
// Basic usage (auto-detects everything)
const entries = CSVProcessor.parseCSV(csvContent);

// With custom options
const entries = CSVProcessor.parseCSV(csvContent, {
  delimiter: ';',           // Force semicolon delimiter
  skipHeaderRows: 2,        // Skip first 2 rows
  encoding: 'utf-8'         // Specify encoding
});
```

### 4. **Robust Error Handling**

The processor includes fallback mechanisms:
- If header detection fails → uses default column positions (0,1,2,3,4,5)
- If delimiter detection fails → defaults to comma
- Invalid data rows are skipped automatically
- Detailed error messages for troubleshooting

### 5. **Supported Header Variations**

| Thai Column | Alternative Names | English Names |
|-------------|-------------------|---------------|
| ชื่อบัญชี | ชื่อ | Account Name, AccountName |
| รหัสบัญชี | รหัส | Account Code, Code, AccountCode |
| ยอดยกมาต้นงวด | ยอดต้นงวด | Previous Balance, PrevBalance |
| ยอดยกมางวดนี้ | ยอดงวดนี้ | Current Balance, CurrBalance |
| เดบิต | เดบิท | Debit, DR |
| เครดิต | เครดิท | Credit, CR |

## Example Usage

```typescript
// CSV with different headers
const csvContent = `
ชื่อ,รหัส,ยอดต้นงวด,ยอดงวดนี้,เดบิต,เครดิต
เงินสด,1010,100000,0,150000,0
ลูกหนี้การค้า,1210,50000,0,75000,0
`;

// Process automatically
const data = CSVProcessor.processCsvFile(csvContent, companyInfo);

// The processor will:
// 1. Detect that headers use shortened names
// 2. Map columns correctly 
// 3. Parse data into standard TrialBalanceEntry format
// 4. Generate financial statements normally
```

## Benefits

1. **Backward Compatible**: Existing CSV files still work
2. **User-Friendly**: Accepts common header variations
3. **International**: Supports both Thai and English headers
4. **Robust**: Graceful fallback when detection fails
5. **Extensible**: Easy to add new header variations

This flexibility eliminates the need for users to format their CSV files to exact specifications, making the application more user-friendly and adaptable to different accounting software exports.
