# PPE Note Layout - Enhanced Row Tracking with Zero Display

## PPE Note Structure (Notes_Accounting Worksheet)

```
Column:    A    B                           C             D                E    F         G           H    I
Row:
15         6.   ที่ดิน อาคารและอุปกรณ์                                                                        หน่วย:บาท
16              
17                                         ณ 31 ธ.ค. 2566     ซื้อเพิ่ม    จำหน่ายออก         ณ 31 ธ.ค. 2567
18              ราคาทุนเดิม
19                   ที่ดิน                   500,000         0           0                 500,000
20                   อาคารและสิ่งปลูกสร้าง    1,800,000      200,000      0               2,000,000
21                   เครื่องจักรและอุปกรณ์     700,000       100,000      0                 800,000
22                   รวม                      =SUM(D19:D21)  =SUM(F19:F21) =SUM(G19:G21)   =SUM(I19:I21)
23              ค่าเสื่อมราคาสะสม
24                   อาคารและสิ่งปลูกสร้าง    (250,000)      (50,000)     0               (300,000)
25                   เครื่องจักรและอุปกรณ์    (150,000)      (50,000)     0               (200,000)
26                   รวม                      =SUM(D24:D25)  =SUM(F24:F25) =SUM(G24:G25)   =SUM(I24:I25)
27              มูลค่าสุทธิ                   =D22-D26       =F22-F26     =G22-G26         =I22-I26
```

## Key Features of the Enhanced Layout:

### **1. Header Structure (Rows 15-17)**
- **Row 15**: Note number (6.) + Title + Unit designation (หน่วย:บาท)
- **Row 16**: Blank spacer row
- **Row 17**: Column headers with year labels

### **2. Asset Cost Section (Rows 18-22)**
- **Row 18**: Section header "ราคาทุนเดิม"
- **Rows 19-21**: Individual asset accounts with amounts
- **Row 22**: Total row with SUM formulas

### **3. Depreciation Section (Rows 23-26)**
- **Row 23**: Section header "ค่าเสื่อมราคาสะสม"
- **Rows 24-25**: Depreciation accounts (negative amounts)
- **Row 26**: Total row with SUM formulas

### **4. Net Book Value Section (Row 27)**
- **Row 27**: Net book value calculation (มูลค่าสุทธิ)

## Critical Zero Display Fix:

### **Problem (FIXED)**
```typescript
// OLD CODE - Cleared zeros:
if (!cell.value || ...) {  // !0 = true, so zeros were cleared!
```

### **Solution (IMPLEMENTED)**
```typescript
// NEW CODE - Preserves zeros:
if ((cell.value === undefined || cell.value === null) || 
    (typeof cell.value === 'string' && cell.value.trim() === '') ||
    (typeof cell.value === 'number' && cell.value === 0 && row > 10 && shouldClearZeros && !isProtectedColumn)) {
```

### **Protected Columns**
- **Column F (6)**: ซื้อเพิ่ม (Purchases) - Zeros now display as "0"
- **Column G (7)**: จำหน่ายออก (Disposals) - Zeros now display as "0"

## Row Tracking Implementation:

### **NoteRowTracker Structure**
```typescript
tracker = {
  currentRow: 15,           // Starting row
  noteStartRow: 15,         // Note beginning
  headerRows: [15, 18, 23],          // Bold headers (note title, cost section, depreciation section)
  yearHeaderRows: [17],              // Underlined year headers
  detailRows: [19,20,21,24,25],      // Detail account rows
  totalRows: [22, 26, 27],           // Bold total rows (cost total, depreciation total, net value)
  unitRows: [15]                     // Unit designation rows
}
```

## Formatting Applied:

### **Bold Text**
- Note title and section headers (headerRows)
- Total rows with "รวม" (totalRows)
- Unit designation "หน่วย:บาท" (unitRows)

### **Underlined Text**
- Year column headers (yearHeaderRows)

### **Number Formatting**
- Thai locale with comma separators
- Negative numbers in parentheses
- Zero values displayed as "0" (not blank)

### **Excel Formulas**
- SUM formulas for automatic calculation
- Cell references for net book value calculations
- Dynamic updates when source data changes

This layout ensures professional financial statement presentation with proper zero display and comprehensive Excel formula integration!
